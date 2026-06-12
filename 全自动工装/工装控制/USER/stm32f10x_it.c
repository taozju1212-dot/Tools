/* STM32F10x 中断处理 - 全自动工装板 */
#include "stm32f10x_it.h"
#include "stm32f10x.h"
#include "bsp.h"
#include "led.h"
#include "app_config.h"
#include "app_osc.h"

/* ============================================================================
 *  Cortex-M 系统异常
 * ============================================================================ */
void NMI_Handler(void)        {}
void HardFault_Handler(void)  { while (1); }
void MemManage_Handler(void)  { while (1); }
void BusFault_Handler(void)   { while (1); }
void UsageFault_Handler(void) { while (1); }
void SVC_Handler(void)        {}
void DebugMon_Handler(void)   {}
void PendSV_Handler(void)     {}
/* SysTick_Handler 由 delay.c 提供（非 OS 模式为空/不存在）*/

/* ============================================================================
 *  TIM4 更新中断 - 1ms 系统时基
 * ============================================================================ */
void TIM4_IRQHandler(void)
{
    if (TIM_GetITStatus(TIM4, TIM_IT_Update) != RESET) {
        TIM_ClearITPendingBit(TIM4, TIM_IT_Update);
        g_tick_ms++;
        g_idle_ms++;
    }
}

/* ============================================================================
 *  EXTI1 - PB1 光开关 1（双沿）
 * ============================================================================ */
void EXTI1_IRQHandler(void)
{
    if (EXTI_GetITStatus(EXTI_Line1) != RESET) {
        g_opt_sw1 = (OPT_SW1() == 0) ? 1 : 0;   /* 低电平 = 遮挡 = ✓ */
        g_dirty_status = 1;
        Osc_OnGpi1Edge_FromIsr();               /* 振荡 GO 状态下立即硬停反转 */
        EXTI_ClearITPendingBit(EXTI_Line1);
    }
}

/* ============================================================================
 *  EXTI9_5 - PB9 光开关 2 (EXTI9) + PA8 编码器按键 KEY1 (EXTI8)
 * ============================================================================ */
void EXTI9_5_IRQHandler(void)
{
    /* PB9 光开关 2 */
    if (EXTI_GetITStatus(EXTI_Line9) != RESET) {
        g_opt_sw2 = (OPT_SW2() == 0) ? 1 : 0;
        g_dirty_status = 1;
        EXTI_ClearITPendingBit(EXTI_Line9);
    }

    /* PA8 KEY1 编码器按键（双沿） */
    if (EXTI_GetITStatus(EXTI_Line8) != RESET) {
        if (KEY_ENC1() == 0) {
            /* 下降沿：按下 */
            g_key_down     = 1;
            g_key_press_ms = g_tick_ms;
        } else {
            /* 上升沿：抬起 */
            if (g_key_down) {
                uint32_t held = (uint32_t)(g_tick_ms - g_key_press_ms);
                g_key_down       = 0;
                g_key_long_armed = 0;   /* 释放后允许下次长按重新触发 */
                /* 短按：抬起时长 < 1.33s 且长按未已触发 */
                if (held < 1333 && !g_key_long) {
                    g_key_short = 1;
                }
            }
        }
        g_idle_ms = 0;
        EXTI_ClearITPendingBit(EXTI_Line8);
    }
}

/* ============================================================================
 *  EXTI3 - PB3 M2 反转 M2_D（自锁止双沿，参考 PC13 DO1 风格）
 *  按下锁定=上升沿→toggle；再按弹出=下降沿→toggle；按下取消另一方向
 * ============================================================================ */
void EXTI3_IRQHandler(void)
{
    if (EXTI_GetITStatus(EXTI_Line3) != RESET) {
        g_jog_m2d ^= 1;
        if (g_jog_m2d) g_jog_m2u = 0;   /* 互斥：本路启动时取消另一路 */
        g_jog_dirty = 1;
        g_idle_ms   = 0;
        EXTI_ClearITPendingBit(EXTI_Line3);
    }
}

/* ============================================================================
 *  EXTI4 - PA4 M1 正转 M1_U（自锁止双沿）
 * ============================================================================ */
void EXTI4_IRQHandler(void)
{
    if (EXTI_GetITStatus(EXTI_Line4) != RESET) {
        g_jog_m1u ^= 1;
        if (g_jog_m1u) g_jog_m1d = 0;
        g_jog_dirty = 1;
        g_idle_ms   = 0;
        EXTI_ClearITPendingBit(EXTI_Line4);
    }
}

/* ============================================================================
 *  EXTI15_10 - PA12 M1_D / PA15 M2_U / PA11 KEY_F2 / PC13 DO1
 * ============================================================================ */
void EXTI15_10_IRQHandler(void)
{
    /* PA12 M1 反转 M1_D（自锁止双沿） */
    if (EXTI_GetITStatus(EXTI_Line12) != RESET) {
        g_jog_m1d ^= 1;
        if (g_jog_m1d) g_jog_m1u = 0;
        g_jog_dirty = 1;
        g_idle_ms   = 0;
        EXTI_ClearITPendingBit(EXTI_Line12);
    }

    /* PA15 M2 正转 M2_U（自锁止双沿） */
    if (EXTI_GetITStatus(EXTI_Line15) != RESET) {
        g_jog_m2u ^= 1;
        if (g_jog_m2u) g_jog_m2d = 0;
        g_jog_dirty = 1;
        g_idle_ms   = 0;
        EXTI_ClearITPendingBit(EXTI_Line15);
    }

    /* PA11 KEY_F2 自锁键：双沿，电平驱动循环模式，PC14 同步指示 */
    if (EXTI_GetITStatus(EXTI_Line11) != RESET) {
        g_loop_mode = KEY_F2();          /* 1=按住锁定，0=弹出 */
        LED1 = g_loop_mode;             /* 高电平点亮：按住时 LED 亮 */
        Osc_SetEnabled(g_loop_mode);
        g_dirty_all = 1;                /* 刷新主界面 R 指示符 */
        g_idle_ms = 0;
        EXTI_ClearITPendingBit(EXTI_Line11);
    }

    /* PC13 DO1 触发键（上升沿 → toggle PB0，按下=高电平） */
    if (EXTI_GetITStatus(EXTI_Line13) != RESET) {
        g_do1_state ^= 1;
        DO1(g_do1_state);
        g_idle_ms = 0;
        EXTI_ClearITPendingBit(EXTI_Line13);
    }
}

/* ============================================================================
 *  USART1 - TMC2226 半双工 UART 接收
 * ============================================================================ */
void USART1_IRQHandler(void)
{
    if (USART_GetITStatus(USART1, USART_IT_RXNE) != RESET) {
        uint8_t b = (uint8_t)USART_ReceiveData(USART1);
        if (g_tmc_rx_cnt < TMC_BUF_SIZE)
            g_tmc_rx_buf[g_tmc_rx_cnt++] = b;
    }
}

/* ============================================================================
 *  USART2 - 条码模块接收（环形缓冲）
 * ============================================================================ */
void USART2_IRQHandler(void)
{
    if (USART_GetITStatus(USART2, USART_IT_RXNE) != RESET) {
        uint8_t b = (uint8_t)USART_ReceiveData(USART2);
        if (g_scan_buf.len < SCAN_BUF_SIZE) {
            g_scan_buf.buf[g_scan_buf.tail] = b;
            g_scan_buf.tail = (g_scan_buf.tail + 1) % SCAN_BUF_SIZE;
            g_scan_buf.len++;
        }
        g_scan_rx_flag = 1;
    }
}

/* ============================================================================
 *  USART3 - 热敏打印机（通常单向，接收丢弃）
 * ============================================================================ */
void USART3_IRQHandler(void)
{
    if (USART_GetITStatus(USART3, USART_IT_RXNE) != RESET)
        (void)USART_ReceiveData(USART3);
}
