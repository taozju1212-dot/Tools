#ifndef __BSP_H
#define __BSP_H

#include "sys.h"

/* ============================================================================
 * 全自动工装板 - 板级支持包 (BSP)
 * MCU: STM32F103C8T6 (LQFP48)  时钟: HSI->PLL 64MHz
 * 引脚分配与 docs/管脚配置.md 0530 版严格对应
 * ----------------------------------------------------------------------------
 *   PA0  TIM2_CH1  编码器 A 相
 *   PA1  TIM2_CH2  编码器 B 相
 *   PA2  USART2_TX 条码模块 TX
 *   PA3  USART2_RX 条码模块 RX
 *   PA4  EXTI4     M1 正转键 M1_U
 *   PA5  ADC1_IN5  热敏电阻
 *   PA6  TIM3_CH1  步进 1 STEP PWM
 *   PA7  TIM3_CH2  步进 2 STEP PWM
 *   PA8  EXTI8     编码器按键 KEY1
 *   PA9  USART1_TX TMC2226 半双工 UART
 *   PA10 Analog    未用（防浮空）
 *   PA11 EXTI11    功能预留键 KEY_F2
 *   PA12 EXTI12    M1 反转键 M1_D
 *   PA13 SWDIO
 *   PA14 SWCLK
 *   PA15 EXTI15    M2 正转键 M2_U（需 JTAG-disable）
 *
 *   PB0  GPO       DO1 MOS 控制
 *   PB1  EXTI1     光开关 1 O1
 *   PB2  -         BOOT1 硬件接地，不作 GPIO
 *   PB3  EXTI3     M2 反转键 M2_D（需 JTAG-disable）
 *   PB4  GPI       OPB9000 OUT（需 JTAG-disable）
 *   PB5  GPO       OPB9000 CAL
 *   PB6  I2C1_SCL  OLED I2C（默认引脚，无重映射）
 *   PB7  I2C1_SDA  OLED I2C
 *   PB8  -         未用（浮空，不作按键）
 *   PB9  EXTI9     光开关 2 O2
 *   PB10 USART3_TX 热敏打印机 TX
 *   PB11 USART3_RX 热敏打印机 RX
 *   PB12 GPO       步进 1 DIR
 *   PB13 GPO       步进 2 DIR
 *   PB14 GPO       步进 1 EN（低有效）
 *   PB15 GPO       步进 2 EN（低有效）
 *
 *   PC13 EXTI13    DO1 触发键
 *   PC14 GPO       LED1（低电平点亮）
 * ============================================================================ */

/* ---------- 输出宏 ---------- */
#define M1_DIR(x)       (PBout(12) = (x))
#define M2_DIR(x)       (PBout(13) = (x))
#define M1_EN(x)        (PBout(14) = (x))   /* 0=使能，1=禁用 */
#define M2_EN(x)        (PBout(15) = (x))
#define DO1(x)          (PBout(0)  = (x))
#define OPB9000_CAL(x)  (PBout(5)  = (x))

/* ---------- 输入宏 ---------- */
#define OPB9000_OUT()   (PBin(4))
#define KEY_ENC1()      (PAin(8))
#define KEY_F2()        (PAin(11))
#define KEY_M1_U()      (PAin(4))
#define KEY_M1_D()      (PAin(12))
#define KEY_M2_U()      (PAin(15))
#define KEY_M2_D()      (PBin(3))
#define KEY_DO1()       (PCin(13))
#define OPT_SW1()       (PBin(1))
#define OPT_SW2()       (PBin(9))

/* ---------- BSP 初始化 ---------- */
void Bsp_Init(void);
void Bsp_Clock_Init(void);
void Bsp_GPIO_Init(void);
void Bsp_EXTI_Init(void);
void Bsp_TIM3_PWM_Init(uint16_t arr, uint16_t psc);
void Bsp_TIM2_Encoder_Init(void);
void Bsp_USART1_Init(uint32_t baud);
void Bsp_USART2_Init(uint32_t baud);
void Bsp_USART3_Init(uint32_t baud);
void Bsp_ADC1_Init(void);
void Bsp_I2C1_Init(void);
void Bsp_TIM4_Tick_Init(void);  /* 1ms 系统时基 (TIM4 UPD 中断) */

/* ---------- 应用辅助 ---------- */
uint16_t Bsp_ADC_Read_Thermistor(void);
int16_t  Bsp_Encoder1_Get(void);
void     Bsp_TIM3_Set_Step1_Freq(uint32_t hz); /* 0=停 */
void     Bsp_TIM3_Set_Step2_Freq(uint32_t hz);
void     Bsp_Motor_Start(uint8_t motor, uint16_t rpm, uint8_t dir);
void     Bsp_Motor_Stop(uint8_t motor);

/* M1 步数计数器（TIM1 硬件计数，时钟源 = TIM3.OC1REF）
   每发出 1 个 STEP 脉冲，TIM1 +1。CCR1=0 时 OC1REF 保持低，TIM1 停。
   16-bit 限制：单次最多计到 65535 步（≈1.6s @1500 RPM 1/8 细分） */
void     Bsp_TIM1_StepCounter_Init(void);
static __inline void     Bsp_M1_StepCount_Reset(void) { TIM1->CNT = 0; }
static __inline uint16_t Bsp_M1_StepCount_Get(void)   { return (uint16_t)TIM1->CNT; }

#endif /* __BSP_H */
