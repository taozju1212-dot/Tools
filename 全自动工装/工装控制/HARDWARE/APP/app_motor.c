#include "app_motor.h"
#include "app_osc.h"
#include "bsp.h"
#include "delay.h"

/* ============================================================================
 *  TMC2226 UART 协议 (半双工单线, 115200 8N1)
 *  PA9 = AF_PP（发送）/ IN_FLOATING（接收）手动切换
 *
 *  从机地址（实测确认，硬件 MS pin 命名与 TMC 数据手册相反）：
 *    M1 = slave 0  (PCB 标号 MS0=GND, MS1=GND → TMC MS1=0, MS2=0)
 *    M2 = slave 2  (PCB 标号 MS0=GND, MS1=VCC → TMC MS1=0, MS2=1)
 *
 *  关键设计（按参考工程方式）：
 *  - 发送时 PA9 = AF_PP，USART 主动驱动总线
 *  - 读寄存器时，发完 4 字节请求后切到 IN_FLOATING，让 TMC 驱动总线返回 8 字节
 *  - 读取的 RX 缓冲含 12 字节：前 4 为本机发出的 echo，后 8 为 TMC 响应
 *  - CRC 按 TMC 数据手册：每字节 LSB first 处理，多项式 0x07
 * ============================================================================ */

/* 关键寄存器地址 */
#define TMC_GCONF        0x00
#define TMC_IFCNT        0x02
#define TMC_IHOLD_IRUN   0x10
#define TMC_TPWMTHRS     0x13
#define TMC_CHOPCONF     0x6C

/* 参考工程的初始化值（VSENSE=0 保留 1.35A 上限）*/
#define TMC_GCONF_VAL     0x000000C1UL    /* I_scale_analog=1, pdn_disable=1, mstep_reg_select=1 */
#define TMC_CHOPCONF_VAL  0x05000153UL    /* MRES=5(1/8), VSENSE=0, INTPOL=0, TOFF=3, HSTRT=5, HEND=2 */
#define TMC_TPWMTHRS_VAL  0x00000001UL    /* StealthChop 切换阈值 */
#define TMC_IHOLDDELAY    6               /* IHOLD_IRUN bits[19:16] = 6（参考工程值） */

/* ============================================================================
 *  PA9 方向切换（AF_PP 发送 ↔ IN_FLOATING 接收）
 * ============================================================================ */
static void pa9_set_tx(void)
{
    GPIO_InitTypeDef gi;
    gi.GPIO_Mode  = GPIO_Mode_AF_PP;
    gi.GPIO_Speed = GPIO_Speed_50MHz;
    gi.GPIO_Pin   = GPIO_Pin_9;
    GPIO_Init(GPIOA, &gi);
}

static void pa9_set_rx(void)
{
    GPIO_InitTypeDef gi;
    gi.GPIO_Mode  = GPIO_Mode_IN_FLOATING;
    gi.GPIO_Speed = GPIO_Speed_50MHz;
    gi.GPIO_Pin   = GPIO_Pin_9;
    GPIO_Init(GPIOA, &gi);
}

/* TMC UART CRC8（多项式 0x07）—— 按 TMC2226 数据手册伪代码：每字节 LSB first */
static uint8_t tmc_crc8(const uint8_t *buf, uint8_t len)
{
    uint8_t crc = 0;
    uint8_t b;
    uint8_t i;
    while (len--) {
        b = *buf++;
        for (i = 0; i < 8; i++) {
            if ((crc >> 7) ^ (b & 0x01)) {
                crc = (uint8_t)((crc << 1) ^ 0x07);
            } else {
                crc = (uint8_t)(crc << 1);
            }
            b >>= 1;
        }
    }
    return crc;
}

/* 发送单字节（参考工程方式：先写 DR，后等 TXE） */
static void tmc_send_byte(uint8_t b)
{
    USART_ClearITPendingBit(USART1, USART_IT_TC);
    USART1->DR = b;
    while (USART_GetFlagStatus(USART1, USART_FLAG_TXE) == RESET);
}

/* 发完一组字节后等 TC（传输完成） */
static void tmc_wait_tc(void)
{
    while (USART_GetFlagStatus(USART1, USART_FLAG_TC) == RESET);
}

/* 清 RX 计数器和 DR（不动 RXNE 中断） */
static void tmc_clear_rx(void)
{
    __disable_irq();
    /* 清掉 DR 中可能残留的字节 */
    (void)USART1->SR;
    (void)USART1->DR;
    g_tmc_rx_cnt = 0;
    __enable_irq();
}

/* ============================================================================
 *  写寄存器（8 字节帧，不需要读响应）
 *  发完不切 GPIO，RX 缓冲会自动收到 8 字节 echo，下次 clear 时丢弃
 * ============================================================================ */
static void tmc_write_reg(uint8_t slave, uint8_t addr, uint32_t data)
{
    uint8_t frame[8];
    uint8_t i;

    frame[0] = 0x05;
    frame[1] = slave;
    frame[2] = addr | 0x80;
    frame[3] = (uint8_t)(data >> 24);
    frame[4] = (uint8_t)(data >> 16);
    frame[5] = (uint8_t)(data >>  8);
    frame[6] = (uint8_t)(data);
    frame[7] = tmc_crc8(frame, 7);

    tmc_clear_rx();
    for (i = 0; i < 8; i++) tmc_send_byte(frame[i]);
    tmc_wait_tc();
    delay_ms(10);   /* 给 TMC 时间处理 */
}

/* ============================================================================
 *  读寄存器（4 字节请求 → 切到 RX → 等 12 字节 → 跳过 4 byte echo）
 *  返回：1=成功, 0=超时/CRC/sync/slave/addr 校验失败
 * ============================================================================ */
static uint8_t tmc_read_reg(uint8_t slave, uint8_t addr, uint32_t *out)
{
    uint8_t frame[4];
    uint8_t resp[8];
    uint8_t i;
    uint32_t t0;

    frame[0] = 0x05;
    frame[1] = slave;
    frame[2] = addr;
    frame[3] = tmc_crc8(frame, 3);

    tmc_clear_rx();
    delay_ms(2);

    /* 发送 4 字节请求 */
    for (i = 0; i < 4; i++) tmc_send_byte(frame[i]);
    tmc_wait_tc();

    /* 切到接收态：PA9 IN_FLOATING，TMC 驱动总线发响应 */
    pa9_set_rx();

    /* 等总共 12 字节（4 echo + 8 response），10ms 超时 */
    t0 = g_tick_ms;
    while (g_tmc_rx_cnt < 12) {
        if (TICK_ELAPSED(t0) > 10) {
            pa9_set_tx();
            return 0;
        }
    }

    /* 切回发送态 */
    pa9_set_tx();
    delay_ms(10);

    /* 跳过前 4 字节 echo，取后 8 字节响应 */
    for (i = 0; i < 8; i++) resp[i] = g_tmc_rx_buf[4 + i];

    /* 严格校验：sync(0x05) + slave(0xFF 表示主→从) + addr 回显 + CRC */
    if (resp[7] != tmc_crc8(resp, 7)) return 0;
    if (resp[0] != 0x05)              return 0;
    if (resp[1] != 0xFF)              return 0;
    if (resp[2] != addr)              return 0;

    *out = ((uint32_t)resp[3] << 24) |
           ((uint32_t)resp[4] << 16) |
           ((uint32_t)resp[5] <<  8) |
            (uint32_t)resp[6];
    return 1;
}

/* 通用寄存器读（诊断/外部使用） */
uint8_t TMC_Read_Reg(uint8_t motor, uint8_t addr, uint32_t *out)
{
    uint8_t slave = (motor == 1) ? 0 : 2;
    return tmc_read_reg(slave, addr, out);
}

/* ============================================================================
 *  诊断：探测指定从机地址（0-3），统计 RX 缓冲实际收到了多少字节
 *  不做 CRC/sync 校验，只把"原始事实"返回出来
 * ============================================================================ */
void TMC_Probe_Slave(uint8_t slave, uint8_t *bytes_received, uint8_t *first_resp_byte)
{
    uint8_t frame[4];
    uint8_t i;
    uint32_t t0;

    if (slave > 3) {
        *bytes_received = 0;
        *first_resp_byte = 0;
        return;
    }

    frame[0] = 0x05;
    frame[1] = slave;
    frame[2] = TMC_GCONF;
    frame[3] = tmc_crc8(frame, 3);

    pa9_set_tx();
    tmc_clear_rx();
    delay_ms(2);

    for (i = 0; i < 4; i++) tmc_send_byte(frame[i]);
    tmc_wait_tc();

    pa9_set_rx();

    /* 诊断用更宽容的 50ms 超时 */
    t0 = g_tick_ms;
    while (g_tmc_rx_cnt < 12) {
        if (TICK_ELAPSED(t0) > 50) break;
    }

    pa9_set_tx();
    delay_ms(10);

    *bytes_received  = g_tmc_rx_cnt;
    *first_resp_byte = (g_tmc_rx_cnt >= 5) ? g_tmc_rx_buf[4] : 0;
}

/* 老接口：按 motor 编号探测（M1=slave0, M2=slave1） */
void TMC_Diag_RawRead(uint8_t motor, uint8_t *bytes_received, uint8_t *first_resp_byte)
{
    TMC_Probe_Slave((motor == 1) ? 0 : 2, bytes_received, first_resp_byte);
}

/* ============================================================================
 *  GCONF 写入与回读验证
 *  必须看到 I_scale_analog(b0) + pdn_disable(b6) + mstep_reg_select(b7) = 1
 * ============================================================================ */
uint8_t TMC_Write_GCONF(uint8_t motor)
{
    uint8_t slave = (motor == 1) ? 0 : 2;
    uint32_t readback = 0;
    tmc_write_reg(slave, TMC_GCONF, TMC_GCONF_VAL);
    delay_ms(2);
    if (!tmc_read_reg(slave, TMC_GCONF, &readback)) return 0;
    return ((readback & 0xC1) == 0xC1) ? 1 : 0;
}

/* ============================================================================
 *  CHOPCONF 写入与回读验证（VSENSE=0, MRES=5）
 * ============================================================================ */
uint8_t TMC_Write_CHOPCONF(uint8_t motor)
{
    uint8_t slave = (motor == 1) ? 0 : 2;
    uint32_t readback = 0;
    tmc_write_reg(slave, TMC_CHOPCONF, TMC_CHOPCONF_VAL);
    delay_ms(2);
    if (!tmc_read_reg(slave, TMC_CHOPCONF, &readback)) return 0;
    /* 校验 MRES=5 且 VSENSE=0 */
    if (((readback >> 24) & 0x0F) != 0x05) return 0;
    if ((readback & (1UL << 16)) != 0)     return 0;   /* VSENSE (bit16) 必须为 0 */
    return 1;
}

/* ============================================================================
 *  IHOLD_IRUN 写入（参考工程：IHOLDDELAY=6，IHOLD=irun/2，IRUN=irun）
 *  寄存器只写，用 IFCNT 验证
 * ============================================================================ */
uint8_t TMC_Write_IRUN(uint8_t motor, uint8_t irun)
{
    uint8_t slave = (motor == 1) ? 0 : 2;
    uint8_t ihold = irun / 2;
    uint32_t cnt_before = 0;
    uint32_t cnt_after = 0;
    uint32_t val;

    if (!tmc_read_reg(slave, TMC_IFCNT, &cnt_before)) return 0;

    val = ((uint32_t)TMC_IHOLDDELAY << 16) | ((uint32_t)irun << 8) | ihold;
    tmc_write_reg(slave, TMC_IHOLD_IRUN, val);
    delay_ms(2);

    if (!tmc_read_reg(slave, TMC_IFCNT, &cnt_after)) return 0;
    return ((cnt_after & 0xFF) == ((cnt_before + 1) & 0xFF)) ? 1 : 0;
}

/* ============================================================================
 *  TPWMTHRS 写入（参考工程：写 1，激活 StealthChop 切换阈值）
 *  寄存器只写，用 IFCNT 验证
 * ============================================================================ */
static uint8_t tmc_write_tpwmthrs(uint8_t motor)
{
    uint8_t slave = (motor == 1) ? 0 : 2;
    uint32_t cnt_before = 0;
    uint32_t cnt_after = 0;
    if (!tmc_read_reg(slave, TMC_IFCNT, &cnt_before)) return 0;
    tmc_write_reg(slave, TMC_TPWMTHRS, TMC_TPWMTHRS_VAL);
    delay_ms(2);
    if (!tmc_read_reg(slave, TMC_IFCNT, &cnt_after)) return 0;
    return ((cnt_after & 0xFF) == ((cnt_before + 1) & 0xFF)) ? 1 : 0;
}

/* ============================================================================
 *  按参考工程顺序初始化单片 TMC：GCONF → IHOLD_IRUN → CHOPCONF → TPWMTHRS
 *  每步最多重试 3 次
 * ============================================================================ */
static void tmc_init_one(uint8_t motor, uint8_t irun)
{
    uint8_t r;
    for (r = 0; r < 3; r++) { if (TMC_Write_GCONF(motor))      break; delay_ms(5); }
    for (r = 0; r < 3; r++) { if (TMC_Write_IRUN(motor, irun)) break; delay_ms(5); }
    for (r = 0; r < 3; r++) { if (TMC_Write_CHOPCONF(motor))   break; delay_ms(5); }
    for (r = 0; r < 3; r++) { if (tmc_write_tpwmthrs(motor))   break; delay_ms(5); }
}

void Motor_Init_From_Config(void)
{
    uint16_t arr = Rpm_To_Arr(g_cfg.m1_rpm);
    TIM_SetAutoreload(TIM3, arr);
    TIM_SetCompare1(TIM3, 0);
    TIM_SetCompare2(TIM3, 0);

    /* 确保 PA9 处于发送态 */
    pa9_set_tx();

    tmc_init_one(1, g_cfg.m1_irun);
    tmc_init_one(2, g_cfg.m2_irun);
}

/* ============================================================================
 *  加减速斜波控制
 *  状态机决策树：
 *    1. target_rpm==0          → 减速到 0，停转
 *    2. dir 需要切换 && 在转   → 先减速到 0
 *    3. 静止 && target>0       → 跳到 RAMP_START_RPM（或 target，若更小）
 *    4. cur < target           → 加速
 *    5. cur > target           → 减速到 target
 *    6. cur == target          → 稳态
 * ============================================================================ */

typedef struct {
    uint16_t target_rpm;     /* 目标 RPM；0 = 停 */
    uint8_t  target_dir;     /* 目标方向 */
    uint16_t cur_rpm;        /* 当前 RPM */
    uint8_t  cur_dir;        /* 当前方向 */
    uint8_t  running;        /* 驱动器是否使能 */
    uint32_t last_ms;        /* 上次更新时刻 */
} MotorRamp;

static MotorRamp g_ramp[2];  /* [0]=M1, [1]=M2 */

void Motor_Set_Target(uint8_t motor, uint16_t rpm, uint8_t dir)
{
    if (motor < 1 || motor > 2) return;
    g_ramp[motor - 1].target_rpm = rpm;
    g_ramp[motor - 1].target_dir = dir ? 1 : 0;
}

/* ISR 内紧急停 M1：脉冲立即消失，驱动器保持使能锁定位置；ramp 状态重置 */
void Motor_FastStop_M1_FromIsr(void)
{
    TIM_SetCompare1(TIM3, 0);   /* TIM3 CH1 CCR=0 → STEP 脚保持低电平 */
    g_ramp[0].cur_rpm = 0;       /* 下次 Motor_Ramp_Update 走 case 3 重新启动 */
    /* M1_EN 保持 0（使能），由 TMC2226 保持转矩锁定当前微步位置 */
}

void Motor_Ramp_Update(void)
{
    uint8_t  i;
    uint8_t  motor;
    uint32_t dt;
    uint16_t step;
    MotorRamp *r;

    for (i = 0; i < 2; i++) {
        r = &g_ramp[i];
        motor = (uint8_t)(i + 1);

        dt = g_tick_ms - r->last_ms;
        if (dt < RAMP_UPDATE_INTV_MS) continue;
        r->last_ms = g_tick_ms;

        /* 每 tick 的 RPM 步长（向上取整保证至少 1）*/
        step = (uint16_t)((RAMP_ACCEL_RPM_S * dt) / 1000);
        if (step == 0) step = 1;

        /* --- 1. 目标 = 停转 --- */
        if (r->target_rpm == 0) {
            if (r->cur_rpm > step) {
                r->cur_rpm = (uint16_t)(r->cur_rpm - step);
                Bsp_Motor_Start(motor, r->cur_rpm, r->cur_dir);
            } else if (r->cur_rpm > 0) {
                r->cur_rpm = 0;
                Bsp_Motor_Stop(motor);
                r->running = 0;
            } else if (r->running) {
                Bsp_Motor_Stop(motor);
                r->running = 0;
            }
            continue;
        }

        /* --- 2. 方向不同且仍在转：先减速到 0 才换向 --- */
        if (r->cur_dir != r->target_dir && r->cur_rpm > 0) {
            if (r->cur_rpm > step) {
                r->cur_rpm = (uint16_t)(r->cur_rpm - step);
                Bsp_Motor_Start(motor, r->cur_rpm, r->cur_dir);  /* 旧方向继续 */
            } else {
                r->cur_rpm = 0;
                r->cur_dir = r->target_dir;                      /* 安全换向 */
                /* 不停机；下一轮 update 会从 RAMP_START_RPM 起加速 */
            }
            continue;
        }

        /* --- 3. 静止启动：跳到 RAMP_START_RPM (上限 target) --- */
        if (r->cur_rpm == 0) {
            uint16_t start = RAMP_START_RPM;
            if (start > r->target_rpm) start = r->target_rpm;
            r->cur_rpm = start;
            r->cur_dir = r->target_dir;
            Bsp_Motor_Start(motor, r->cur_rpm, r->cur_dir);
            r->running = 1;
            continue;
        }

        /* --- 4. 加速 --- */
        if (r->cur_rpm < r->target_rpm) {
            uint32_t next = (uint32_t)r->cur_rpm + step;
            if (next >= r->target_rpm) r->cur_rpm = r->target_rpm;
            else                       r->cur_rpm = (uint16_t)next;
            Bsp_Motor_Start(motor, r->cur_rpm, r->cur_dir);
            continue;
        }

        /* --- 5. 减速到非 0 目标 --- */
        if (r->cur_rpm > r->target_rpm) {
            if (r->cur_rpm < r->target_rpm + step) r->cur_rpm = r->target_rpm;
            else                                    r->cur_rpm = (uint16_t)(r->cur_rpm - step);
            Bsp_Motor_Start(motor, r->cur_rpm, r->cur_dir);
            continue;
        }

        /* --- 6. 稳态：cur_rpm == target_rpm，无需操作 --- */
    }
}

/* ---------- 点动评估 (主循环调用) ---------- */
void Motor_Jog_Eval(void)
{
    /* 振荡模式 RUN 中：所有手动按键忽略，且清状态防积压 */
    if (g_osc_state >= OSC_GO) {
        g_jog_m1u = 0; g_jog_m1d = 0;
        g_jog_m2u = 0; g_jog_m2d = 0;
        g_jog_dirty = 0;
        return;
    }

    /* 振荡模式 ARMED：等 M1_U 或 M1_D 按下设定初始方向，然后启动振荡 */
    if (g_osc_state == OSC_ARMED) {
        if (g_jog_dirty) {
            if (g_jog_m1u)      Osc_Start_With_Dir(1);
            else if (g_jog_m1d) Osc_Start_With_Dir(0);
            /* 不论按了哪个键，都清掉避免下一轮影响 */
            g_jog_m1u = 0; g_jog_m1d = 0;
            g_jog_m2u = 0; g_jog_m2d = 0;
            g_jog_dirty = 0;
        }
        return;
    }

    /* OSC_OFF：正常点动逻辑（走斜波）*/
    if (!g_jog_dirty) return;
    g_jog_dirty = 0;

    if (g_jog_m1u)      Motor_Set_Target(1, g_cfg.m1_rpm, 1);
    else if (g_jog_m1d) Motor_Set_Target(1, g_cfg.m1_rpm, 0);
    else                Motor_Set_Target(1, 0, 0);

    if (g_jog_m2u)      Motor_Set_Target(2, g_cfg.m2_rpm, 1);
    else if (g_jog_m2d) Motor_Set_Target(2, g_cfg.m2_rpm, 0);
    else                Motor_Set_Target(2, 0, 0);
}
