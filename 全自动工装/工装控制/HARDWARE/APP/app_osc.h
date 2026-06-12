#ifndef __APP_OSC_H
#define __APP_OSC_H

#include "app_config.h"

/* ============================================================================
 *  PA11 M1 振荡测试模式
 *
 *  状态机：
 *    OFF → PA11 → ARMED → M1_U/M1_D → GO ⇄ BACK 循环
 *    任意 RUN 状态 → PA11 → OFF (强制退出，M1 停)
 *
 *  GO :  初始方向运行 → GPI1 触发 → 记录耗时 t_go，反向，进入 BACK
 *  BACK: 反向运行 t_go 时长 → 视为回到起点，反回初始方向，进入 GO
 *  来回振荡，直到 PA11 再次按下
 * ============================================================================ */

typedef enum {
    OSC_OFF = 0,
    OSC_ARMED,
    OSC_GO,           /* 初始方向 → 等 GPI 触发 */
    OSC_BACK,         /* 反向 → 回起点（t_go 时长） */
} OscState;

extern volatile OscState g_osc_state;

void Osc_Init(void);
void Osc_SetEnabled(uint8_t en);             /* PA11 ISR 调用：1=按下启动，0=弹出停止 */
void Osc_Start_With_Dir(uint8_t forward);    /* ARMED 下由 Motor_Jog_Eval 调用 */
void Osc_Update(void);                       /* 主循环每次调用 */
void Osc_OnGpi1Edge_FromIsr(void);           /* EXTI1 ISR 调用：GO 状态下立即反转 */

#endif /* __APP_OSC_H */
