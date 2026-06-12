#ifndef __APP_MOTOR_H
#define __APP_MOTOR_H

#include "app_config.h"

/* ---------- 初始化 ---------- */
void Motor_Init_From_Config(void);  /* 从 g_cfg 写 TIM3 ARR + TMC2226 IRUN */

/* ---------- TMC2226 UART ---------- */
uint8_t TMC_Write_IRUN(uint8_t motor, uint8_t irun);   /* 1=OK 0=fail */
uint8_t TMC_Write_GCONF(uint8_t motor);                 /* pdn_dis=1, mstep_reg_select=1 */
uint8_t TMC_Write_CHOPCONF(uint8_t motor);              /* MRES=5 (1/8 细分) */
uint8_t TMC_Read_Reg(uint8_t motor, uint8_t addr, uint32_t *out);  /* 通用读寄存器（诊断用）*/

/* 启动诊断：尝试读 GCONF，返回收到的总字节数 + 响应首字节
   *bytes_received: 0=完全无响应, 4=只有 echo, 12=完整收到 (echo+resp)
   *first_resp_byte: 收到 ≥5 字节时为 rx_buf[4]（应是 sync 0x05），否则 0 */
void TMC_Diag_RawRead(uint8_t motor, uint8_t *bytes_received, uint8_t *first_resp_byte);

/* 探测任意从机地址 0-3 的响应（用于确认硬件 MS1/MS2 实际配置） */
void TMC_Probe_Slave(uint8_t slave, uint8_t *bytes_received, uint8_t *first_resp_byte);

/* ---------- 点动评估（主循环调用） ---------- */
void Motor_Jog_Eval(void);          /* 根据 g_jog_xxx 更新电机状态 */

/* ---------- 加减速斜波控制 ----------
 * 替代直接调 Bsp_Motor_Start/Stop。
 *   Motor_Set_Target: 设定目标 RPM 和方向，立即返回；rpm=0 表示停转
 *   Motor_Ramp_Update: 主循环周期性调用，每 RAMP_UPDATE_INTV_MS 推进一步
 * 方向反转时会先减速到 0 再换向，避免堵转。 */
void Motor_Set_Target(uint8_t motor, uint16_t rpm, uint8_t dir);
void Motor_Ramp_Update(void);

/* ISR 内紧急停 M1 STEP 脉冲（保持驱动器使能 → 保持转矩锁定位置）
 * 同时把 ramp.cur_rpm 重置为 0，下次 Motor_Ramp_Update 会从 RAMP_START_RPM 起
 * 按 target_dir 加速。用于振荡模式 GPI 触发后最快反转。 */
void Motor_FastStop_M1_FromIsr(void);

/* ---------- RPM <-> ARR 换算 ---------- */
/* ARR+1 = 37500 / rpm  (TIM3 1MHz base, 8-microstep, 200steps/rev) */
static __inline uint16_t Rpm_To_Arr(uint16_t rpm)
{
    uint32_t arr;
    if (rpm == 0) return 1499;
    arr = 37500UL / rpm;
    if (arr < 1) arr = 1;
    return (uint16_t)(arr - 1);
}

/* 电流显示：Rsense=150mΩ, I = 1.35 * (IRUN+1) / 32 (A)；返回 0.1A 单位整数（截断）
   IRUN=31: 1.35A → 13 (显示 "1.3A")
   IRUN=28: 1.22A → 12 (显示 "1.2A")
   注意：用截断而非四舍五入，与用户规格一致 */
static __inline uint16_t Irun_To_mA10(uint8_t irun)
{
    return (uint16_t)((1350UL * (irun + 1)) / 3200);
}

#endif /* __APP_MOTOR_H */
