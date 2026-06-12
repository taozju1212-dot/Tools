#include "sys.h"
#include "delay.h"
#include "bsp.h"
#include "led.h"
#include "app_config.h"
#include "app_flash.h"
#include "app_motor.h"
#include "app_oled.h"
#include "app_calib.h"
#include "app_scan.h"
#include "app_print.h"
#include "app_osc.h"

/* ============================================================================
 *  全自动工装板 main.c - 完整状态机
 *  MCU: STM32F103C8T6  时钟: HSI->PLL 64MHz
 * ============================================================================ */

/* ---------- 前向声明：被 Handle_SpeedEdit/CurEdit 早于定义点调用 ---------- */
static void Blink_Reset_On(void);

/* ---------- 辅助：读编码器格数 ---------- */
static int8_t Enc_GetDelta(void)
{
    int16_t cur;
    int16_t diff;
    int8_t clicks;
    cur    = Bsp_Encoder1_Get();
    diff   = (int16_t)(cur - g_enc_last);
    clicks = (int8_t)(diff / 2);          /* EC11: 2 counts/detent in 4x mode */
    /* 只前进已消费的整数格，保留半格余数，避免被下一次轮询吃掉 */
    g_enc_last = (int16_t)(g_enc_last + (int16_t)clicks * 2);
    return (int8_t)(-clicks);             /* 翻转方向：CW/CCW 与界面体感一致 */
}

/* ---------- 辅助：读 ADC 5 次均值 ---------- */
static uint16_t ADC_Read_Avg5(void)
{
    uint32_t sum = 0;
    uint8_t i;
    for (i = 0; i < 5; i++)
        sum += Bsp_ADC_Read_Thermistor();
    return (uint16_t)(sum / 5);
}

/* ---------- 状态辅助：草稿写回 g_cfg + 硬件 + Flash ----------
 * LED 反馈协议：
 *   Flash 写失败 → 3 次 150ms 慢闪
 *   TMC 写失败  → 5 次 80ms 快闪（最常见硬件链路问题）
 *   两者同失败  → 先慢闪后快闪 */
static void Save_Draft(void)
{
    uint16_t arr;
    uint8_t flash_ok;
    uint8_t tmc_ok = 0;
    uint8_t i, k;

    if (g_edit_motor == 1) {
        g_cfg.m1_rpm  = g_draft_rpm;
        g_cfg.m1_irun = g_draft_irun;
        arr = Rpm_To_Arr(g_draft_rpm);
        TIM_SetAutoreload(TIM3, arr);
    } else {
        g_cfg.m2_rpm  = g_draft_rpm;
        g_cfg.m2_irun = g_draft_irun;
    }

    /* TMC IRUN 重试 3 次（半双工 UART 可能偶发 echo 丢字节/CRC 错） */
    for (k = 0; k < 3 && !tmc_ok; k++) {
        tmc_ok = TMC_Write_IRUN(g_edit_motor, g_draft_irun);
        if (!tmc_ok) delay_ms(5);
    }

    flash_ok = Flash_Save_Config();

    if (!flash_ok) {
        for (i = 0; i < 3; i++) { LED_ON(); delay_ms(150); LED_OFF(); delay_ms(150); }
    }
    if (!tmc_ok) {
        for (i = 0; i < 5; i++) { LED_ON(); delay_ms(80); LED_OFF(); delay_ms(80); }
    }
}

/* ---------- 状态辅助：保存并返回 HOME（长按使用） ---------- */
static void Save_And_Go_Home(void)
{
    Save_Draft();
    g_state     = STATE_HOME;
    g_idle_ms   = 0;
    g_dirty_all = 1;
}

/* ---------- 状态辅助：静默返回 HOME ---------- */
static void Go_Home(void)
{
    g_state     = STATE_HOME;
    g_idle_ms   = 0;
    g_dirty_all = 1;
}

/* ============================================================================
 *  状态处理函数
 * ============================================================================ */

/* ---- HOME ---- */
static void Handle_Home(int8_t delta)
{
    if (delta != 0) {
        if (g_home_focus == FOCUS_COUNT) {
            /* 首次旋转：顺转先选 M1，逆转先选 M2 */
            g_home_focus = (delta > 0) ? FOCUS_M1 : FOCUS_M2;
        } else {
            /* 已有焦点：在 M1 ↔ M2 间切换 */
            g_home_focus = (g_home_focus == FOCUS_M1) ? FOCUS_M2 : FOCUS_M1;
        }
        g_dirty_all = 1;
    }

    if (g_key_short) {
        g_key_short = 0;
        if (g_home_focus == FOCUS_M1 || g_home_focus == FOCUS_M2) {
            g_edit_motor  = (g_home_focus == FOCUS_M1) ? 1 : 2;
            g_motor_focus = MOTOR_FOCUS_SPEED;
            g_draft_rpm   = (g_edit_motor == 1) ? g_cfg.m1_rpm  : g_cfg.m2_rpm;
            g_draft_irun  = (g_edit_motor == 1) ? g_cfg.m1_irun : g_cfg.m2_irun;
            g_state = STATE_MOTOR_MENU;
            g_dirty_all = 1;
        }
    }
    /* HOME 下长按无效 */
    if (g_key_long) g_key_long = 0;
}

/* ---- MOTOR_MENU ---- */
static void Handle_MotorMenu(int8_t delta)
{
    if (delta != 0) {
        uint8_t f = (g_motor_focus == MOTOR_FOCUS_SPEED) ? 1 : 0;
        g_motor_focus = f ? MOTOR_FOCUS_CUR : MOTOR_FOCUS_SPEED;
        g_dirty_all = 1;
    }

    if (g_key_short) {
        g_key_short = 0;
        g_state = (g_motor_focus == MOTOR_FOCUS_SPEED) ? STATE_SPEED_EDIT : STATE_CUR_EDIT;
        g_dirty_all = 1;
    }

    if (g_key_long) {
        g_key_long = 0;
        Go_Home();
    }
}

/* ---- SPEED_EDIT ---- */
static void Handle_SpeedEdit(int8_t delta)
{
    if (delta != 0) {
        int32_t rpm = (int32_t)g_draft_rpm + (int32_t)delta * RPM_STEP;
        if (rpm < RPM_MIN) rpm = RPM_MIN;
        if (rpm > RPM_MAX) rpm = RPM_MAX;
        g_draft_rpm = (uint16_t)rpm;
        g_dirty_all = 1;
        Blink_Reset_On();           /* 旋转后立即点亮新值 */
    }

    if (g_key_short) {
        g_key_short = 0;
        Save_Draft();                 /* 短按 = 确认保存当前修改 */
        g_state = STATE_MOTOR_MENU;
        g_idle_ms = 0;
        g_dirty_all = 1;
    }

    if (g_key_long) {
        g_key_long = 0;
        Save_And_Go_Home();
    }
}

/* ---- CUR_EDIT ---- */
static void Handle_CurEdit(int8_t delta)
{
    if (delta != 0) {
        int8_t irun = (int8_t)g_draft_irun + (int8_t)(delta * IRUN_STEP);
        if (irun < IRUN_MIN) irun = IRUN_MIN;
        if (irun > IRUN_MAX) irun = IRUN_MAX;
        g_draft_irun = (uint8_t)irun;
        g_dirty_all = 1;
        Blink_Reset_On();           /* 旋转后立即点亮新值 */
    }

    if (g_key_short) {
        g_key_short = 0;
        Save_Draft();                 /* 短按 = 确认保存当前修改 */
        g_state = STATE_MOTOR_MENU;
        g_idle_ms = 0;
        g_dirty_all = 1;
    }

    if (g_key_long) {
        g_key_long = 0;
        Save_And_Go_Home();
    }
}

/* ---- PRINT_RUN ---- */
static void Handle_PrintRun(void)
{
    Print_Run();
    delay_ms(3000);              /* PRINTING → PRINT OK 间隔 3s，给打印机留时间 */
    UI_Render_Print_Done();
    delay_ms(1000);
    Go_Home();
}

/* ---- SCAN_RUN ---- */
#define SCAN_TIMEOUT_MS   3000
#define SCAN_IDLE_RETURN  5000    /* 成功结果显示 5s */
#define SCAN_FAIL_RETURN  2500    /* 失败结果显示 2.5s（成功的一半）*/
#define SCAN_FRAME_GAP_MS 200

static void Handle_ScanRun(void)
{
    static uint32_t scan_start_ms = 0;
    static uint32_t last_rx_ms    = 0;
    static uint8_t  scan_done     = 0;
    static char     scan_result[SCAN_RESULT_MAX];
    static uint32_t result_ms     = 0;

    /* 进入状态时初始化 */
    if (g_dirty_all) {
        scan_start_ms = g_tick_ms;
        last_rx_ms    = g_tick_ms;
        scan_done     = 0;
        scan_result[0]= '\0';
        result_ms     = 0;
        g_dirty_all   = 0;
        UI_Render_Scan(NULL);
    }

    if (!scan_done) {
        uint8_t frame_done;
        uint8_t timed_out;

        /* 收到新字节 */
        if (g_scan_rx_flag) {
            g_scan_rx_flag = 0;
            last_rx_ms     = g_tick_ms;
        }

        /* 帧间隔判定 */
        frame_done = (g_scan_buf.len > 0 &&
                      TICK_ELAPSED(last_rx_ms) > SCAN_FRAME_GAP_MS);
        /* 超时判定 */
        timed_out  = (g_scan_buf.len == 0 &&
                      TICK_ELAPSED(scan_start_ms) > SCAN_TIMEOUT_MS);

        if (frame_done) {
            Scan_Process(scan_result);
            scan_done  = 1;
            result_ms  = g_tick_ms;
            UI_Render_Scan(scan_result);
        } else if (timed_out) {
            scan_done = 1;
            scan_result[0] = '\0';
            result_ms = g_tick_ms;
            UI_Render_Scan_Fail();
        }
    } else {
        /* 已有结果：短按重扫 */
        if (g_key_short) {
            g_key_short = 0;
            Scan_Start();
            scan_start_ms = g_tick_ms;
            last_rx_ms    = g_tick_ms;
            scan_done     = 0;
            scan_result[0]= '\0';
            g_dirty_all   = 1;
        }
        /* 成功 5s / 失败 2.5s 无操作自动返回 */
        {
            uint32_t timeout = (scan_result[0] == '\0') ? SCAN_FAIL_RETURN : SCAN_IDLE_RETURN;
            if (TICK_ELAPSED(result_ms) > timeout) {
                Go_Home();
            }
        }
    }
}

/* ---- CAL_STEP1 ---- */
#define CAL_STEP1_FAIL_SHOW_MS  2000

static void Handle_CalStep1(void)
{
    static uint8_t  result_shown = 0;
    static uint32_t fail_ms      = 0;

    if (g_dirty_all) {
        result_shown = 0;
        fail_ms      = 0;
        g_dirty_all  = 0;
    }

    if (!result_shown) {
        uint8_t ok = Calib_Step1();
        result_shown = 1;
        if (ok) {
            UI_Render_Cal_Step1_OK();
        } else {
            /* 当前 Calib_Step1 失败 = PB4 无响应 = 通讯失败
               UI_Render_Cal_Step1_Fail（校准失败 靠近传感器）保留供后续区分使用 */
            fail_ms = g_tick_ms;
            UI_Render_Cal_Comm_Fail();
        }
    }

    /* 失败：等 2s 后返回 HOME */
    if (fail_ms && TICK_ELAPSED(fail_ms) > CAL_STEP1_FAIL_SHOW_MS) {
        Go_Home();
        return;
    }

    /* 成功等待：短按→STEP2；长按→取消回 HOME */
    if (g_key_short) {
        g_key_short = 0;
        g_state = STATE_CAL_STEP2;
        g_dirty_all = 1;
    }
    if (g_key_long) {
        g_key_long = 0;
        Go_Home();
    }
}

/* ---- CAL_STEP2 ---- */
#define CAL_STEP2_DONE_MS   3000
#define CAL_STEP2_FAIL_MS   2000
#define CAL_STEP2_MAX_RETRY 3

static void Handle_CalStep2(void)
{
    static uint8_t  done_shown  = 0;
    static uint32_t result_ms   = 0;
    static uint8_t  retry_cnt   = 0;
    static uint8_t  step2_ok    = 0;
    static uint8_t  agc_val     = 0;
    static uint16_t led_val     = 0;
    static uint8_t  ref_val     = 0;

    if (g_dirty_all) {
        done_shown  = 0;
        result_ms   = 0;
        retry_cnt   = 0;
        step2_ok    = 0;
        g_dirty_all = 0;
    }

    if (!done_shown) {
        uint8_t agc = 0; uint16_t led = 0; uint8_t ref = 0;
        uint8_t ok = Calib_Step2(&agc, &led, &ref);
        done_shown = 1;
        result_ms  = g_tick_ms;
        if (ok) {
            step2_ok = 1;
            agc_val = agc; led_val = led; ref_val = ref;
            UI_Render_Cal_Step2_OK(agc, led, ref);
        } else {
            retry_cnt++;
            UI_Render_Cal_Step2_Fail();
        }
    }

    if (step2_ok) {
        /* 成功：显示 3s 后回 HOME */
        if (TICK_ELAPSED(result_ms) > CAL_STEP2_DONE_MS) Go_Home();
        (void)agc_val; (void)led_val; (void)ref_val;
    } else {
        /* 失败：等 2s 后继续等待短按重试，或超过 3 次回 HOME */
        if (TICK_ELAPSED(result_ms) > CAL_STEP2_FAIL_MS) {
            if (retry_cnt >= CAL_STEP2_MAX_RETRY) {
                Go_Home();
            } else {
                /* 等待 KEY1 短按重试 */
                if (g_key_short) {
                    g_key_short = 0;
                    done_shown  = 0;
                }
            }
        }
    }
}

/* ============================================================================
 *  编辑态值闪烁节拍：500ms 切换。仅在 SPEED_EDIT / CUR_EDIT 状态有效。
 *  外部可调 Blink_Reset() 强制点亮（如旋转编码器后立即显示新值）。
 * ============================================================================ */
static uint8_t  g_blink_phase = 1;
static uint32_t g_blink_tick  = 0;

static void Blink_Reset_On(void)
{
    g_blink_phase = 1;
    g_blink_tick  = g_tick_ms;
    if (g_state == STATE_SPEED_EDIT) UI_BlinkValue_Speed(1);
    else if (g_state == STATE_CUR_EDIT) UI_BlinkValue_Cur(1);
}

static void Blink_500ms(void)
{
    if (g_state != STATE_SPEED_EDIT && g_state != STATE_CUR_EDIT) {
        g_blink_phase = 1;            /* 离开编辑态，重置为"亮" */
        g_blink_tick  = g_tick_ms;
        return;
    }
    if (TICK_ELAPSED(g_blink_tick) < 500) return;
    g_blink_tick = g_tick_ms;
    g_blink_phase ^= 1;
    if (g_state == STATE_SPEED_EDIT) UI_BlinkValue_Speed(g_blink_phase);
    else                              UI_BlinkValue_Cur(g_blink_phase);
}

/* ============================================================================
 *  1Hz 温度采样
 * ============================================================================ */
static void Temp_1Hz(void)
{
    static uint32_t temp_tick = 0;
    if (TICK_ELAPSED(temp_tick) >= 1000) {
        temp_tick   = g_tick_ms;
        g_temp_raw  = ADC_Read_Avg5();
        UI_Render_Overlay_Temp();
    }
}

/* ============================================================================
 *  启动诊断：扫描 4 个 TMC 从机地址（0/1/2/3），每个发一次读 GCONF
 *  显示每个地址的 RX 字节数和首响应字节，便于确认硬件 MS1/MS2 配置
 *
 *  实测硬件（PCB MS0/MS1 命名与 TMC 数据手册反向）：
 *    S0 RX:0C B:05  ← M1 在地址 0
 *    S1 RX:04 B:00  ← 此地址无芯片
 *    S2 RX:0C B:05  ← M2 在地址 2（PCB 接 MS0=GND, MS1=VCC）
 *    S3 RX:04 B:00  ← 此地址无芯片
 *  代码中已用 slave = (motor==1) ? 0 : 2 对应此实际接线
 * ============================================================================ */
static void TMC_Diag_Show(void)
{
    static const char HEX[] = "0123456789ABCDEF";
    uint8_t s, rx, b;
    char buf[16];

    OLED_Clear();

    for (s = 0; s < 4; s++) {
        TMC_Probe_Slave(s, &rx, &b);
        delay_ms(20);

        /* "S0 RX:NN B:XX" = 13 chars × 8 = 104px，居中 x=12 */
        buf[0]='S'; buf[1]=(char)('0'+s); buf[2]=' ';
        buf[3]='R'; buf[4]='X'; buf[5]=':';
        buf[6]=HEX[(rx>>4)&0x0F]; buf[7]=HEX[rx&0x0F];
        buf[8]=' '; buf[9]='B'; buf[10]=':';
        buf[11]=HEX[(b>>4)&0x0F]; buf[12]=HEX[b&0x0F];
        buf[13]='\0';
        FB_DrawStr8x16(12, (uint8_t)(s * 16), buf);
    }

    OLED_MarkDirty();
    OLED_Flush();
    delay_ms(6000);   /* 6 秒方便看 4 行 */
}

/* ============================================================================
 *  主函数
 * ============================================================================ */
int main(void)
{
    Bsp_Init();     /* 时钟 + GPIO + EXTI + TIM + UART + ADC + I2C + TIM4 tick */
    delay_init();   /* SysTick delay 基准（非 OS 模式，只需 fac_us/ms 计算） */

    /* 上电同步：PA11 自锁键实际电平 → g_loop_mode + PC14 LED
       LED_Init() 默认熄灭，这里覆盖为实际状态，避免上电瞬间显示错误 */
    g_loop_mode = KEY_F2();
    LED1 = g_loop_mode;            /* 高电平点亮 */

    /* ---------- 上电初始化序列（PRD §3 要求顺序） ---------- */
    Flash_Load_Config();          /* 读 Flash 参数（失败用默认值） */

    g_temp_raw = ADC_Read_Avg5(); /* 首次 ADC 采样 */

    /* 初始化光开关状态（主动读 GPIO 电平） */
    g_opt_sw1 = (OPT_SW1() == 0) ? 1 : 0;
    g_opt_sw2 = (OPT_SW2() == 0) ? 1 : 0;

    OLED_Init();                  /* OLED 硬件初始化 */

    Motor_Init_From_Config();     /* 写 TIM3 ARR + TMC2226 IRUN */

    Osc_Init();                   /* 振荡测试模式初始化 */

    /* ---------- 扫描模块上电稳定（首次 SCN 按下时自动探测波特率） ----------
       MZR813H 上电后约需 2 秒完成内部启动 */
    delay_ms(2000);
    Scan_Init();                  /* 当前为空函数，保留以便将来扩展 */

    /* ---------- 初始编码器基准 ---------- */
    g_enc_last = Bsp_Encoder1_Get();

    /* ---------- 渲染首页 ---------- */
    UI_Render_Home();

    /* ---------- 主循环 ---------- */
    while (1) {
        int8_t delta;
        AppState prev;

        /* 1. 读编码器格数 */
        delta = Enc_GetDelta();
        if (delta) g_idle_ms = 0;

        /* 2. 长按检测（主循环轮询，≥2s 置位 g_key_long；armed 标志防重复触发） */
        if (g_key_down && !g_key_long_armed) {
            if (TICK_ELAPSED(g_key_press_ms) >= 1333) {  /* 原 2000ms 的 2/3 */
                g_key_long       = 1;
                g_key_long_armed = 1;
                g_idle_ms        = 0;
            }
        }

        /* 3. 30s 超时自动回 HOME */
        if (g_idle_ms >= IDLE_TIMEOUT_MS && g_state != STATE_HOME) {
            g_key_short = 0;
            g_key_long  = 0;
            Go_Home();
        }

        /* 4. 点动评估（含振荡模式按键路由） */
        Motor_Jog_Eval();

        /* 4.5 振荡模式状态机更新（OFF/ARMED 状态直接 return） */
        Osc_Update();

        /* 4.6 加减速斜波推进（每 5ms 一步，把当前 RPM 朝目标平滑变化） */
        Motor_Ramp_Update();

        /* 5. 状态机调度 */
        prev = g_state;
        switch (g_state) {
            case STATE_HOME:
                Handle_Home(delta);
                if (g_dirty_all && g_state == STATE_HOME) {
                    UI_Render_Home();
                    g_dirty_all = 0;
                }
                break;

            case STATE_MOTOR_MENU:
                Handle_MotorMenu(delta);
                if (g_dirty_all || g_state != prev) {
                    if (g_state == STATE_MOTOR_MENU) {
                        UI_Render_MotorMenu();
                    } else if (g_state == STATE_HOME) {
                        UI_Render_Home();
                    }
                    g_dirty_all = 0;
                }
                break;

            case STATE_SPEED_EDIT:
                Handle_SpeedEdit(delta);
                if (g_dirty_all || g_state != prev) {
                    if (g_state == STATE_SPEED_EDIT)   UI_Render_SpeedEdit();
                    else if (g_state == STATE_MOTOR_MENU) UI_Render_MotorMenu();
                    else if (g_state == STATE_HOME)    UI_Render_Home();
                    g_dirty_all = 0;
                }
                break;

            case STATE_CUR_EDIT:
                Handle_CurEdit(delta);
                if (g_dirty_all || g_state != prev) {
                    if (g_state == STATE_CUR_EDIT)    UI_Render_CurEdit();
                    else if (g_state == STATE_MOTOR_MENU) UI_Render_MotorMenu();
                    else if (g_state == STATE_HOME)   UI_Render_Home();
                    g_dirty_all = 0;
                }
                break;

            case STATE_PRINT_RUN:
                if (g_dirty_all) {
                    UI_Render_Print_Running();
                    g_dirty_all = 0;
                    Handle_PrintRun();  /* 阻塞执行，完成后状态已切回 HOME */
                    UI_Render_Home();
                }
                break;

            case STATE_SCAN_RUN:
                Handle_ScanRun();
                if (g_state == STATE_HOME) {
                    UI_Render_Home();
                    g_dirty_all = 0;
                }
                break;

            case STATE_CAL_STEP1:
                Handle_CalStep1();
                if (g_dirty_all || g_state != prev) {
                    if (g_state == STATE_HOME) {
                        UI_Render_Home();
                    }
                    g_dirty_all = 0;
                }
                break;

            case STATE_CAL_STEP2:
                Handle_CalStep2();
                if (g_state == STATE_HOME) {
                    UI_Render_Home();
                    g_dirty_all = 0;
                }
                break;

            default: break;
        }

        /* 6. 温度 1Hz 刷新（顶栏覆盖，不影响子界面中间区） */
        Temp_1Hz();

        /* 7. 光开关/DO1 状态条刷新（脏时） */
        if (g_dirty_status) {
            g_dirty_status = 0;
            UI_Render_Overlay_Status();
        }

        /* 8. 编辑态值闪烁（500ms） */
        Blink_500ms();

        /* 9. PC14 由 PA11 ISR 直接控制，此处无需额外 LED 处理 */
    }
}
