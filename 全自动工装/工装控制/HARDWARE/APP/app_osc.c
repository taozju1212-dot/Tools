#include "app_osc.h"
#include "app_motor.h"
#include "bsp.h"

/* ============================================================================
 *  PA11 振荡测试模式 —— 纯振荡：起点 ⇄ GPI 触发位置
 *  按 TIM1 硬件计数的步数往复（不再按时间），保证每次往返位置精确匹配。
 *
 *  状态机：
 *    OFF → PA11 → ARMED → M1_U/M1_D → GO ⇄ BACK 循环
 *
 *  GO :  初始方向 → GPI1 触发 (EXTI ISR 内硬停)，记录步数 osc_steps_go
 *  BACK: 反向，TIM1 步数计到 osc_steps_go → 硬停 + 反回初始方向，进入 GO
 *
 *  GPI 触发瞬间通过 EXTI1 ISR 调 Motor_FastStop_M1_FromIsr()，机械过冲 < 2 微步
 * ============================================================================ */

volatile OscState g_osc_state = OSC_OFF;

/* 反向阶段需要走的步数（GO 阶段实测的步数，ISR/主循环都访问 → volatile） */
static volatile uint16_t osc_steps_go = 0;

/* 初始方向：1=正转, 0=反转 */
static volatile uint8_t  osc_dir_initial = 1;

/* GPI1 边沿检测（仅主循环安全网用） */
static uint8_t  osc_last_sw1 = 0;

void Osc_Init(void)
{
    g_osc_state = OSC_OFF;
    osc_last_sw1 = g_opt_sw1;
}

/* PA11 电平驱动：en=1 按住锁定→ARMED，en=0 弹出→立即硬停并退出 */
void Osc_SetEnabled(uint8_t en)
{
    if (en) {
        if (g_osc_state == OSC_OFF)
            g_osc_state = OSC_ARMED;
    } else {
        if (g_osc_state != OSC_OFF) {
            Motor_FastStop_M1_FromIsr();   /* ISR 内调用，直接清 CCR，无阻塞 */
            g_osc_state = OSC_OFF;
        }
    }
}

/* ARMED 状态下被 Motor_Jog_Eval 调用，设定初始方向并启动 */
void Osc_Start_With_Dir(uint8_t forward)
{
    if (g_osc_state != OSC_ARMED) return;
    osc_dir_initial = forward ? 1 : 0;
    osc_last_sw1 = g_opt_sw1;
    osc_steps_go = 0;
    Bsp_M1_StepCount_Reset();          /* GO 阶段开始计数 */
    Motor_Set_Target(1, g_cfg.m1_rpm, osc_dir_initial);
    g_osc_state = OSC_GO;
}

/* ============================================================================
 *  EXTI1 ISR 调用：GPI1 边沿瞬间硬停电机 + 记录 GO 阶段步数 + 切到 BACK
 *  步数 = TIM1 计数（每个 STEP 脉冲 +1），比时间更精确
 * ============================================================================ */
void Osc_OnGpi1Edge_FromIsr(void)
{
    if (g_osc_state != OSC_GO) return;

    osc_steps_go = Bsp_M1_StepCount_Get();   /* 记录 GO 实际步数 */
    Motor_FastStop_M1_FromIsr();              /* TIM3 CCR=0 → STEP 停 → TIM1 也停 */
    Bsp_M1_StepCount_Reset();                 /* BACK 阶段重新从 0 计数 */
    Motor_Set_Target(1, g_cfg.m1_rpm, osc_dir_initial ? 0 : 1);
    g_osc_state = OSC_BACK;
}

/* 主循环每次调用 */
void Osc_Update(void)
{
    uint8_t cur_sw1;
    uint8_t gpi_trig;

    if (g_osc_state == OSC_OFF || g_osc_state == OSC_ARMED) {
        osc_last_sw1 = g_opt_sw1;
        return;
    }

    cur_sw1 = g_opt_sw1;
    gpi_trig = (cur_sw1 != osc_last_sw1);
    osc_last_sw1 = cur_sw1;

    switch (g_osc_state) {
    case OSC_GO:
        /* 通常 ISR 已处理；这里仅作安全网 */
        if (gpi_trig) {
            Osc_OnGpi1Edge_FromIsr();
        }
        break;

    case OSC_BACK:
        /* 反向步数累计到 GO 阶段步数 → 切回 GO 方向重新开始 */
        if (Bsp_M1_StepCount_Get() >= osc_steps_go) {
            Motor_FastStop_M1_FromIsr();     /* 精确停在起点位置 */
            Bsp_M1_StepCount_Reset();        /* 下一轮 GO 重新计数 */
            Motor_Set_Target(1, g_cfg.m1_rpm, osc_dir_initial);
            g_osc_state = OSC_GO;
        }
        break;

    default:
        break;
    }
}
