#include "bsp.h"
#include "led.h"

/* ============================================================================
 *  Bsp_Clock_Init  -  HSI 8MHz -> PLL x16 / 2 = 64MHz
 *  原理图无 HSE 晶振 (PD0/PD1 悬空)，全程使用内部 RC。
 *  HCLK   = 64MHz
 *  PCLK1  = 32MHz (APB1 必须 <=36MHz)，TIM2/3/4 时钟 = PCLK1 x2 = 64MHz
 *  PCLK2  = 64MHz，TIM1 / ADC / GPIO 时钟域
 *  ADCCLK = PCLK2/6 = 10.67MHz (必须 <=14MHz)
 *  同时完成 AFIO 时钟使能、JTAG-only-SWD 重映射。
 * ============================================================================ */
void Bsp_Clock_Init(void)
{
    /* 复位 RCC */
    RCC_DeInit();

    /* 启动 HSI */
    RCC_HSICmd(ENABLE);
    while (RCC_GetFlagStatus(RCC_FLAG_HSIRDY) == RESET);

    /* Flash: 64MHz 需要 2 wait states，使能预取 */
    FLASH_PrefetchBufferCmd(FLASH_PrefetchBuffer_Enable);
    FLASH_SetLatency(FLASH_Latency_2);

    /* 总线分频 */
    RCC_HCLKConfig(RCC_SYSCLK_Div1);   /* AHB  = SYSCLK = 64MHz */
    RCC_PCLK2Config(RCC_HCLK_Div1);    /* APB2 = HCLK    = 64MHz */
    RCC_PCLK1Config(RCC_HCLK_Div2);    /* APB1 = HCLK/2  = 32MHz */
    RCC_ADCCLKConfig(RCC_PCLK2_Div6);  /* ADC  = 10.67MHz */

    /* PLL = HSI/2 * 16 = 4MHz * 16 = 64MHz (HSI PLL 路径强制除 2) */
    RCC_PLLConfig(RCC_PLLSource_HSI_Div2, RCC_PLLMul_16);
    RCC_PLLCmd(ENABLE);
    while (RCC_GetFlagStatus(RCC_FLAG_PLLRDY) == RESET);

    /* 切换 SYSCLK = PLL */
    RCC_SYSCLKConfig(RCC_SYSCLKSource_PLLCLK);
    while (RCC_GetSYSCLKSource() != 0x08);

    /* 更新 SystemCoreClock 变量，使 delay_init 算系数正确 */
    SystemCoreClock = 64000000;

    /* AFIO 时钟必须先开，才能改重映射 */
    RCC_APB2PeriphClockCmd(RCC_APB2Periph_AFIO, ENABLE);

    /* 关闭 JTAG 仅保留 SWD —— 释放 PA15/PB3/PB4 作 GPIO */
    GPIO_PinRemapConfig(GPIO_Remap_SWJ_JTAGDisable, ENABLE);

}

/* ============================================================================
 *  Bsp_GPIO_Init  -  所有数字 GPIO 输入/输出
 *  说明: TIM/USART/I2C/ADC 等复用功能脚的初始化放在各自的 Bsp_XXX_Init() 内。
 *  这里只处理纯 GPIO 用途的引脚。
 * ============================================================================ */
void Bsp_GPIO_Init(void)
{
    GPIO_InitTypeDef gi;

    RCC_APB2PeriphClockCmd(RCC_APB2Periph_GPIOA |
                           RCC_APB2Periph_GPIOB |
                           RCC_APB2Periph_GPIOC, ENABLE);

    /* ---------- 推挽输出 ---------- */
    /* PB0  DO1 MOS 控制 */
    /* PB5  OPB9000 CAL */
    /* PB12 DIR1, PB13 DIR2, PB14 EN1, PB15 EN2 */
    gi.GPIO_Mode  = GPIO_Mode_Out_PP;
    gi.GPIO_Speed = GPIO_Speed_50MHz;
    gi.GPIO_Pin   = GPIO_Pin_0 | GPIO_Pin_5
                  | GPIO_Pin_12 | GPIO_Pin_13 | GPIO_Pin_14 | GPIO_Pin_15;
    GPIO_Init(GPIOB, &gi);

    /* 初始电平：MOS 关、CAL 低、DIR 默认 0、EN 高 (TMC2226 失能) */
    DO1(0);
    OPB9000_CAL(0);
    M1_DIR(0);
    M2_DIR(0);
    M1_EN(1);
    M2_EN(1);

    /* PC14 LED (单独由 LED_Init() 初始化，2MHz) */
    LED_Init();

    /* ---------- 上拉输入：编码器按键 KEY1（按下接 GND，活动低电平）---------- */
    gi.GPIO_Mode = GPIO_Mode_IPU;
    gi.GPIO_Pin  = GPIO_Pin_8;
    GPIO_Init(GPIOA, &gi);

    /* ---------- 下拉输入：电机/DO 按键（按下接 3.3V，活动高电平）---------- */
    /* PA4 M1_U, PA11 KEY_F2, PA12 M1_D, PA15 M2_U */
    gi.GPIO_Mode = GPIO_Mode_IPD;
    gi.GPIO_Pin  = GPIO_Pin_4 | GPIO_Pin_11 | GPIO_Pin_12 | GPIO_Pin_15;
    GPIO_Init(GPIOA, &gi);

    /* PB3 M2_D */
    gi.GPIO_Pin = GPIO_Pin_3;
    GPIO_Init(GPIOB, &gi);

    /* PC13 DO1 触发按键 */
    gi.GPIO_Pin = GPIO_Pin_13;
    GPIO_Init(GPIOC, &gi);

    /* ---------- 浮空输入：U5 缓冲后的光开关 (U5 有明确驱动) + OPB9000 OUT + PB8 未用 ---------- */
    /* PB1 OPT_SW1, PB4 OPB9000_OUT, PB8 未用, PB9 OPT_SW2 */
    gi.GPIO_Mode = GPIO_Mode_IN_FLOATING;
    gi.GPIO_Pin  = GPIO_Pin_1 | GPIO_Pin_4 | GPIO_Pin_8 | GPIO_Pin_9;
    GPIO_Init(GPIOB, &gi);

    /* ---------- 模拟模式：PA10 未用，防止 CMOS 悬空耗电 ---------- */
    gi.GPIO_Mode = GPIO_Mode_AIN;
    gi.GPIO_Pin  = GPIO_Pin_10;
    GPIO_Init(GPIOA, &gi);
}

/* ============================================================================
 *  Bsp_EXTI_Init  -  9 路外部中断
 *  EXTI1=PB1, EXTI3=PB3, EXTI4=PA4, EXTI8=PA8, EXTI9=PB9,
 *  EXTI11=PA11, EXTI12=PA12, EXTI13=PC13, EXTI15=PA15
 *  光开关、电机点动、编码器按键 -> 双沿（松开/遮挡均需响应）
 *  DO1 toggle、KEY_F2 -> 仅下降沿
 * ============================================================================ */
void Bsp_EXTI_Init(void)
{
    EXTI_InitTypeDef ei;
    NVIC_InitTypeDef ni;

    NVIC_PriorityGroupConfig(NVIC_PriorityGroup_2);

    /* ---- 端口 -> EXTI 线 映射 ---- */
    GPIO_EXTILineConfig(GPIO_PortSourceGPIOB, GPIO_PinSource1);   /* PB1  OPT_SW1  */
    GPIO_EXTILineConfig(GPIO_PortSourceGPIOB, GPIO_PinSource9);   /* PB9  OPT_SW2  */
    GPIO_EXTILineConfig(GPIO_PortSourceGPIOB, GPIO_PinSource3);   /* PB3  M2_D     */
    GPIO_EXTILineConfig(GPIO_PortSourceGPIOA, GPIO_PinSource4);   /* PA4  M1_U     */
    GPIO_EXTILineConfig(GPIO_PortSourceGPIOA, GPIO_PinSource8);   /* PA8  KEY1     */
    GPIO_EXTILineConfig(GPIO_PortSourceGPIOA, GPIO_PinSource11);  /* PA11 KEY_F2   */
    GPIO_EXTILineConfig(GPIO_PortSourceGPIOA, GPIO_PinSource12);  /* PA12 M1_D     */
    GPIO_EXTILineConfig(GPIO_PortSourceGPIOC, GPIO_PinSource13);  /* PC13 DO1      */
    GPIO_EXTILineConfig(GPIO_PortSourceGPIOA, GPIO_PinSource15);  /* PA15 M2_U     */

    /* ---- 双沿：光开关、编码器按键、所有自锁止按键 ----
       PC13 DO1、PA4 M1_U、PA12 M1_D、PA15 M2_U、PB3 M2_D 均为自锁止按键：
       按下=锁定（HIGH），再按=弹出（LOW），每次切换都是一次边沿。
       ISR 直接镜像 GPIO 电平到状态变量，不再使用 toggle */
    ei.EXTI_Mode    = EXTI_Mode_Interrupt;
    ei.EXTI_Trigger = EXTI_Trigger_Rising_Falling;
    ei.EXTI_LineCmd = ENABLE;
    ei.EXTI_Line    = EXTI_Line1 | EXTI_Line9  | EXTI_Line8  | EXTI_Line13
                    | EXTI_Line3 | EXTI_Line4  | EXTI_Line12 | EXTI_Line15
                    | EXTI_Line11;   /* PA11 KEY_F2 自锁键：双沿镜像电平 */
    EXTI_Init(&ei);

    /* ---- NVIC ---- */
    ni.NVIC_IRQChannelCmd = ENABLE;

    ni.NVIC_IRQChannel = EXTI1_IRQn;    /* 光开关 1 (PB1) */
    ni.NVIC_IRQChannelPreemptionPriority = 1;
    ni.NVIC_IRQChannelSubPriority        = 0;
    NVIC_Init(&ni);

    /* EXTI9_5 同时覆盖 PB9(OPT_SW2) 和 PA8(KEY1)，优先级取光开关级别 */
    ni.NVIC_IRQChannel = EXTI9_5_IRQn;
    ni.NVIC_IRQChannelPreemptionPriority = 1;
    ni.NVIC_IRQChannelSubPriority        = 1;
    NVIC_Init(&ni);

    ni.NVIC_IRQChannel = EXTI3_IRQn;    /* PB3 M2_D */
    ni.NVIC_IRQChannelPreemptionPriority = 2;
    ni.NVIC_IRQChannelSubPriority        = 0;
    NVIC_Init(&ni);

    ni.NVIC_IRQChannel = EXTI4_IRQn;    /* PA4 M1_U */
    ni.NVIC_IRQChannelPreemptionPriority = 2;
    ni.NVIC_IRQChannelSubPriority        = 1;
    NVIC_Init(&ni);

    ni.NVIC_IRQChannel = EXTI15_10_IRQn;/* PA12 M1_D / PA15 M2_U / PA11 KEY_F2 / PC13 DO1 */
    ni.NVIC_IRQChannelPreemptionPriority = 2;
    ni.NVIC_IRQChannelSubPriority        = 2;
    NVIC_Init(&ni);
}

/* ============================================================================
 *  Bsp_TIM3_PWM_Init  -  TIM3 CH1/CH2 = PA6/PA7  STEP1/STEP2
 *  PWM 边沿对齐, 50% 占空 (步进脉冲只需要边沿, 占空非关键)
 *  TIM3 时钟 = APB1*2 = 64MHz
 *    psc=63 -> 计数频率 1MHz
 *    arr 决定 PWM 周期 -> STEP 频率 = 1MHz / (arr+1)
 *  默认两路均停 (CCR=0)
 * ============================================================================ */
void Bsp_TIM3_PWM_Init(uint16_t arr, uint16_t psc)
{
    GPIO_InitTypeDef gi;
    TIM_TimeBaseInitTypeDef tb;
    TIM_OCInitTypeDef oc;

    RCC_APB2PeriphClockCmd(RCC_APB2Periph_GPIOA, ENABLE);
    RCC_APB1PeriphClockCmd(RCC_APB1Periph_TIM3,  ENABLE);

    /* PA6/PA7 复用推挽输出 */
    gi.GPIO_Mode  = GPIO_Mode_AF_PP;
    gi.GPIO_Speed = GPIO_Speed_50MHz;
    gi.GPIO_Pin   = GPIO_Pin_6 | GPIO_Pin_7;
    GPIO_Init(GPIOA, &gi);

    /* TIM3 时基 */
    tb.TIM_Period        = arr;
    tb.TIM_Prescaler     = psc;
    tb.TIM_ClockDivision = TIM_CKD_DIV1;
    tb.TIM_CounterMode   = TIM_CounterMode_Up;
    TIM_TimeBaseInit(TIM3, &tb);

    /* CH1 (STEP1) PWM1, 默认占空 0 */
    oc.TIM_OCMode      = TIM_OCMode_PWM1;
    oc.TIM_OutputState = TIM_OutputState_Enable;
    oc.TIM_Pulse       = 0;
    oc.TIM_OCPolarity  = TIM_OCPolarity_High;
    TIM_OC1Init(TIM3, &oc);
    TIM_OC1PreloadConfig(TIM3, TIM_OCPreload_Enable);

    /* CH2 (STEP2) */
    TIM_OC2Init(TIM3, &oc);
    TIM_OC2PreloadConfig(TIM3, TIM_OCPreload_Enable);

    TIM_ARRPreloadConfig(TIM3, ENABLE);

    /* TRGO = OC1REF：把 TIM3 CH1 的 PWM 内部信号（每个 STEP 脉冲一次上升沿）
       通过 TIM1 ITR2 给到 TIM1 当外部时钟，做硬件步数计数 */
    TIM_SelectOutputTrigger(TIM3, TIM_TRGOSource_OC1Ref);

    TIM_Cmd(TIM3, ENABLE);
}

/* ============================================================================
 *  Bsp_TIM1_StepCounter_Init - TIM1 作为 M1 步数计数器
 *  时钟源 = TIM3.OC1REF（每发 1 个 STEP 上升沿计数 +1）
 *  TIM1 在 STM32F1 内部触发表中 ITR2 = TIM3，External Clock Mode 1 接管计数
 *  必须在 Bsp_TIM3_PWM_Init 之后调用
 * ============================================================================ */
void Bsp_TIM1_StepCounter_Init(void)
{
    TIM_TimeBaseInitTypeDef tb;

    RCC_APB2PeriphClockCmd(RCC_APB2Periph_TIM1, ENABLE);

    tb.TIM_Period            = 0xFFFF;
    tb.TIM_Prescaler         = 0;
    tb.TIM_ClockDivision     = TIM_CKD_DIV1;
    tb.TIM_CounterMode       = TIM_CounterMode_Up;
    tb.TIM_RepetitionCounter = 0;
    TIM_TimeBaseInit(TIM1, &tb);

    /* TS = ITR2（TIM3），SMS = External Clock Mode 1 */
    TIM_SelectInputTrigger(TIM1, TIM_TS_ITR2);
    TIM_SelectSlaveMode(TIM1, TIM_SlaveMode_External1);

    TIM_SetCounter(TIM1, 0);
    TIM_Cmd(TIM1, ENABLE);
}

void Bsp_TIM3_Set_Step1_Freq(uint32_t hz)
{
    uint32_t arr;
    if (hz == 0) { TIM_SetCompare1(TIM3, 0); return; }
    arr = 1000000UL / hz;
    if (arr == 0) arr = 1;
    TIM_SetAutoreload(TIM3, arr - 1);
    TIM_SetCompare1(TIM3, arr / 2);
}

void Bsp_TIM3_Set_Step2_Freq(uint32_t hz)
{
    uint32_t arr;
    if (hz == 0) { TIM_SetCompare2(TIM3, 0); return; }
    arr = 1000000UL / hz;
    if (arr == 0) arr = 1;
    TIM_SetAutoreload(TIM3, arr - 1);
    TIM_SetCompare2(TIM3, arr / 2);
}

/* ============================================================================
 *  Bsp_TIM2_Encoder_Init  -  TIM2 编码器模式 TI1+TI2 (4x 解码)
 *  PA0 = TIM2_CH1, PA1 = TIM2_CH2
 *  编码器外部已有 10K 上拉到 3.3V (R30/R37)，10R 串联滤波 (R7/R8)
 *  -> MCU 侧浮空输入即可
 *  CNT 16-bit 自动累加，DIR 位反映方向
 * ============================================================================ */
void Bsp_TIM2_Encoder_Init(void)
{
    GPIO_InitTypeDef gi;
    TIM_TimeBaseInitTypeDef tb;
    TIM_ICInitTypeDef ic;

    RCC_APB2PeriphClockCmd(RCC_APB2Periph_GPIOA, ENABLE);
    RCC_APB1PeriphClockCmd(RCC_APB1Periph_TIM2,  ENABLE);

    gi.GPIO_Mode = GPIO_Mode_IN_FLOATING;
    gi.GPIO_Pin  = GPIO_Pin_0 | GPIO_Pin_1;
    GPIO_Init(GPIOA, &gi);

    tb.TIM_Period        = 0xFFFF;
    tb.TIM_Prescaler     = 0;
    tb.TIM_ClockDivision = TIM_CKD_DIV1;
    tb.TIM_CounterMode   = TIM_CounterMode_Up;
    TIM_TimeBaseInit(TIM2, &tb);

    /* 编码器模式 3: TI1 + TI2 双边沿 (4x) */
    TIM_EncoderInterfaceConfig(TIM2, TIM_EncoderMode_TI12,
                               TIM_ICPolarity_Rising, TIM_ICPolarity_Rising);

    /* 输入滤波：减少机械编码器抖动 */
    TIM_ICStructInit(&ic);
    ic.TIM_Channel   = TIM_Channel_1;
    ic.TIM_ICFilter  = 0xA;     /* fSAMPLING=fDTS/32, N=8 */
    TIM_ICInit(TIM2, &ic);
    ic.TIM_Channel   = TIM_Channel_2;
    TIM_ICInit(TIM2, &ic);

    TIM_SetCounter(TIM2, 0);
    TIM_Cmd(TIM2, ENABLE);
}

int16_t Bsp_Encoder1_Get(void) { return (int16_t)TIM_GetCounter(TIM2); }

/* ============================================================================
 *  Bsp_USART1_Init  -  TMC2226 共享 UART (半双工单线)
 *  参考工程方式：
 *    PA9 = AF_PP（推挽，发送时主动驱动高/低）
 *    收发切换：发完手动把 PA9 切到 IN_FLOATING 让 TMC 驱动总线
 *    半双工通过手动设 CR3.HDSEL=1 启用，并显式清 LIN/SmartCard/IrDA 相关位
 *  PA10 默认浮空输入即可（半双工不使用）
 * ============================================================================ */
void Bsp_USART1_Init(uint32_t baud)
{
    GPIO_InitTypeDef gi;
    USART_InitTypeDef ui;
    NVIC_InitTypeDef  ni;

    RCC_APB2PeriphClockCmd(RCC_APB2Periph_GPIOA | RCC_APB2Periph_USART1, ENABLE);

    /* PA9 复用推挽（发送态）。读 TMC 寄存器时由 app_motor 临时切到 IN_FLOATING */
    gi.GPIO_Mode  = GPIO_Mode_AF_PP;
    gi.GPIO_Speed = GPIO_Speed_50MHz;
    gi.GPIO_Pin   = GPIO_Pin_9;
    GPIO_Init(GPIOA, &gi);

    ui.USART_BaudRate            = baud;
    ui.USART_WordLength          = USART_WordLength_8b;
    ui.USART_StopBits            = USART_StopBits_1;
    ui.USART_Parity              = USART_Parity_No;
    ui.USART_HardwareFlowControl = USART_HardwareFlowControl_None;
    ui.USART_Mode                = USART_Mode_Rx | USART_Mode_Tx;
    USART_Init(USART1, &ui);

    /* 半双工：手动设 HDSEL，显式清可能冲突的位（参考工程方式） */
    USART1->CR3 |=  (1 << 3);    /* HDSEL=1 (Half-Duplex) */
    USART1->CR3 &= ~(1 << 1);    /* IREN=0  (IrDA 关) */
    USART1->CR3 &= ~(1 << 5);    /* NACK=0  (SmartCard NACK 关) */
    USART1->CR2 &= ~(1 << 14);   /* LINEN=0 (LIN 关) */
    USART1->CR2 &= ~(1 << 11);   /* CLKEN=0 (Sync 时钟关) */

    USART_ITConfig(USART1, USART_IT_RXNE, ENABLE);

    ni.NVIC_IRQChannel = USART1_IRQn;
    ni.NVIC_IRQChannelPreemptionPriority = 3;
    ni.NVIC_IRQChannelSubPriority        = 0;
    ni.NVIC_IRQChannelCmd = ENABLE;
    NVIC_Init(&ni);

    USART_Cmd(USART1, ENABLE);
}

/* ============================================================================
 *  Bsp_USART2_Init  -  条码模块 (PA2 TX / PA3 RX)
 * ============================================================================ */
void Bsp_USART2_Init(uint32_t baud)
{
    GPIO_InitTypeDef gi;
    USART_InitTypeDef ui;
    NVIC_InitTypeDef  ni;

    RCC_APB2PeriphClockCmd(RCC_APB2Periph_GPIOA, ENABLE);
    RCC_APB1PeriphClockCmd(RCC_APB1Periph_USART2, ENABLE);

    /* PA2 复用推挽 */
    gi.GPIO_Mode  = GPIO_Mode_AF_PP;
    gi.GPIO_Speed = GPIO_Speed_50MHz;
    gi.GPIO_Pin   = GPIO_Pin_2;
    GPIO_Init(GPIOA, &gi);

    /* PA3 浮空输入（与参考工程 STM_Fully_Test 一致；扫码模块 TX 本身有驱动） */
    gi.GPIO_Mode  = GPIO_Mode_IN_FLOATING;
    gi.GPIO_Pin   = GPIO_Pin_3;
    GPIO_Init(GPIOA, &gi);

    ui.USART_BaudRate            = baud;
    ui.USART_WordLength          = USART_WordLength_8b;
    ui.USART_StopBits            = USART_StopBits_1;
    ui.USART_Parity              = USART_Parity_No;
    ui.USART_HardwareFlowControl = USART_HardwareFlowControl_None;
    ui.USART_Mode                = USART_Mode_Rx | USART_Mode_Tx;
    USART_Init(USART2, &ui);

    USART_ITConfig(USART2, USART_IT_RXNE, ENABLE);

    ni.NVIC_IRQChannel = USART2_IRQn;
    ni.NVIC_IRQChannelPreemptionPriority = 3;
    ni.NVIC_IRQChannelSubPriority        = 1;
    ni.NVIC_IRQChannelCmd = ENABLE;
    NVIC_Init(&ni);

    USART_Cmd(USART2, ENABLE);
}

/* ============================================================================
 *  Bsp_USART3_Init  -  热敏打印机 (PB10 TX / PB11 RX)
 * ============================================================================ */
void Bsp_USART3_Init(uint32_t baud)
{
    GPIO_InitTypeDef gi;
    USART_InitTypeDef ui;
    NVIC_InitTypeDef  ni;

    RCC_APB2PeriphClockCmd(RCC_APB2Periph_GPIOB, ENABLE);
    RCC_APB1PeriphClockCmd(RCC_APB1Periph_USART3, ENABLE);

    /* PB10 复用推挽 */
    gi.GPIO_Mode  = GPIO_Mode_AF_PP;
    gi.GPIO_Speed = GPIO_Speed_50MHz;
    gi.GPIO_Pin   = GPIO_Pin_10;
    GPIO_Init(GPIOB, &gi);

    /* PB11 上拉输入 */
    gi.GPIO_Mode  = GPIO_Mode_IPU;
    gi.GPIO_Pin   = GPIO_Pin_11;
    GPIO_Init(GPIOB, &gi);

    ui.USART_BaudRate            = baud;
    ui.USART_WordLength          = USART_WordLength_8b;
    ui.USART_StopBits            = USART_StopBits_1;
    ui.USART_Parity              = USART_Parity_No;
    ui.USART_HardwareFlowControl = USART_HardwareFlowControl_None;
    ui.USART_Mode                = USART_Mode_Rx | USART_Mode_Tx;
    USART_Init(USART3, &ui);

    USART_ITConfig(USART3, USART_IT_RXNE, ENABLE);

    ni.NVIC_IRQChannel = USART3_IRQn;
    ni.NVIC_IRQChannelPreemptionPriority = 3;
    ni.NVIC_IRQChannelSubPriority        = 2;
    ni.NVIC_IRQChannelCmd = ENABLE;
    NVIC_Init(&ni);

    USART_Cmd(USART3, ENABLE);
}

/* ============================================================================
 *  Bsp_ADC1_Init  -  热敏电阻 PA5 = ADC1_IN5
 *  单次软件触发, 12-bit, ADCCLK=10.67MHz, 采样时间 55.5 cycles
 * ============================================================================ */
void Bsp_ADC1_Init(void)
{
    GPIO_InitTypeDef gi;
    ADC_InitTypeDef  ai;

    RCC_APB2PeriphClockCmd(RCC_APB2Periph_GPIOA | RCC_APB2Periph_ADC1, ENABLE);

    /* PA5 模拟输入 */
    gi.GPIO_Mode = GPIO_Mode_AIN;
    gi.GPIO_Pin  = GPIO_Pin_5;
    GPIO_Init(GPIOA, &gi);

    ADC_DeInit(ADC1);

    ai.ADC_Mode               = ADC_Mode_Independent;
    ai.ADC_ScanConvMode       = DISABLE;
    ai.ADC_ContinuousConvMode = DISABLE;
    ai.ADC_ExternalTrigConv   = ADC_ExternalTrigConv_None;
    ai.ADC_DataAlign          = ADC_DataAlign_Right;
    ai.ADC_NbrOfChannel       = 1;
    ADC_Init(ADC1, &ai);

    ADC_Cmd(ADC1, ENABLE);

    /* 复位校准 */
    ADC_ResetCalibration(ADC1);
    while (ADC_GetResetCalibrationStatus(ADC1));
    ADC_StartCalibration(ADC1);
    while (ADC_GetCalibrationStatus(ADC1));
}

uint16_t Bsp_ADC_Read_Thermistor(void)
{
    ADC_RegularChannelConfig(ADC1, ADC_Channel_5, 1, ADC_SampleTime_55Cycles5);
    ADC_SoftwareStartConvCmd(ADC1, ENABLE);
    while (!ADC_GetFlagStatus(ADC1, ADC_FLAG_EOC));
    return ADC_GetConversionValue(ADC1);
}

/* ============================================================================
 *  Bsp_I2C1_Init  -  OLED (默认引脚 PB6 SCL / PB7 SDA，无需重映射)
 *  100kHz 标准模式
 * ============================================================================ */
void Bsp_I2C1_Init(void)
{
    GPIO_InitTypeDef gi;
    I2C_InitTypeDef  ii;

    RCC_APB2PeriphClockCmd(RCC_APB2Periph_GPIOB, ENABLE);
    RCC_APB1PeriphClockCmd(RCC_APB1Periph_I2C1, ENABLE);

    /* PB6/PB7 复用开漏 */
    gi.GPIO_Mode  = GPIO_Mode_AF_OD;
    gi.GPIO_Speed = GPIO_Speed_50MHz;
    gi.GPIO_Pin   = GPIO_Pin_6 | GPIO_Pin_7;
    GPIO_Init(GPIOB, &gi);

    ii.I2C_Mode                = I2C_Mode_I2C;
    ii.I2C_DutyCycle           = I2C_DutyCycle_2;
    ii.I2C_OwnAddress1         = 0x00;
    ii.I2C_Ack                 = I2C_Ack_Enable;
    ii.I2C_AcknowledgedAddress = I2C_AcknowledgedAddress_7bit;
    ii.I2C_ClockSpeed          = 100000;
    I2C_Init(I2C1, &ii);

    I2C_Cmd(I2C1, ENABLE);
}

/* ============================================================================
 *  Bsp_TIM4_Tick_Init  -  1ms 系统时基
 *  TIM4 时钟 = APB1*2 = 64MHz，psc=63 -> 1MHz，arr=999 -> 1kHz 溢出中断
 * ============================================================================ */
void Bsp_TIM4_Tick_Init(void)
{
    TIM_TimeBaseInitTypeDef tb;
    NVIC_InitTypeDef        ni;

    RCC_APB1PeriphClockCmd(RCC_APB1Periph_TIM4, ENABLE);

    tb.TIM_Period        = 999;
    tb.TIM_Prescaler     = 63;
    tb.TIM_ClockDivision = TIM_CKD_DIV1;
    tb.TIM_CounterMode   = TIM_CounterMode_Up;
    TIM_TimeBaseInit(TIM4, &tb);

    TIM_ITConfig(TIM4, TIM_IT_Update, ENABLE);

    ni.NVIC_IRQChannel                   = TIM4_IRQn;
    ni.NVIC_IRQChannelPreemptionPriority  = 0;
    ni.NVIC_IRQChannelSubPriority         = 0;
    ni.NVIC_IRQChannelCmd                 = ENABLE;
    NVIC_Init(&ni);

    TIM_Cmd(TIM4, ENABLE);
}

/* ============================================================================
 *  Bsp_Motor_Start / Bsp_Motor_Stop
 *  motor=1→M1(TIM3_CH1,PB12 DIR,PB14 EN), motor=2→M2(TIM3_CH2,PB13 DIR,PB15 EN)
 *  ARR+1 = 37500/rpm (1MHz base, 8-microstep, 200 steps/rev)
 * ============================================================================ */
void Bsp_Motor_Start(uint8_t motor, uint16_t rpm, uint8_t dir)
{
    uint32_t arr = (rpm > 0) ? (37500UL / rpm) : 1500;
    if (arr < 1) arr = 1;
    TIM_SetAutoreload(TIM3, (uint16_t)(arr - 1));

    if (motor == 1) {
        M1_DIR(dir);
        TIM_SetCompare1(TIM3, (uint16_t)(arr / 2));
        M1_EN(0);
    } else {
        M2_DIR(dir);
        TIM_SetCompare2(TIM3, (uint16_t)(arr / 2));
        M2_EN(0);
    }
}

void Bsp_Motor_Stop(uint8_t motor)
{
    if (motor == 1) {
        TIM_SetCompare1(TIM3, 0);
        M1_EN(1);
    } else {
        TIM_SetCompare2(TIM3, 0);
        M2_EN(1);
    }
}

/* ============================================================================
 *  Bsp_Init  -  顶层一键初始化
 *  调用顺序很重要：
 *    1. 时钟与重映射（先于任何外设/GPIO）
 *    2. GPIO 基础配置
 *    3. EXTI
 *    4. 外设 (TIM/UART/ADC/I2C)
 * ============================================================================ */
void Bsp_Init(void)
{
    Bsp_Clock_Init();

    Bsp_GPIO_Init();
    Bsp_EXTI_Init();

    /* TIM3 默认 1kHz PWM 频率, psc=63 -> 1MHz, arr=999 -> 1kHz */
    Bsp_TIM3_PWM_Init(999, 63);

    /* TIM1 步数计数器（必须在 TIM3 之后） */
    Bsp_TIM1_StepCounter_Init();

    Bsp_TIM2_Encoder_Init();

    Bsp_USART1_Init(115200);    /* TMC2226 默认 */
    Bsp_USART2_Init(115200);    /* MZR813H 条码（出厂默认 115200，手册 P13） */
    Bsp_USART3_Init(9600);      /* 热敏打印机 */

    Bsp_ADC1_Init();
    Bsp_I2C1_Init();
    Bsp_TIM4_Tick_Init();
}
