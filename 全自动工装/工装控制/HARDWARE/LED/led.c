#include "led.h"

/* 工装板单 LED 指示灯初始化
 * 引脚：PC14 (VBAT 域)
 * 模式：通用推挽输出
 * 速度：2MHz (VBAT 域引脚最大频率 2MHz)
 * 复位状态：熄灭 (LED 为高电平点亮的反向逻辑——阳极接 VCC，PC14 拉低点亮)
 */
void LED_Init(void)
{
    GPIO_InitTypeDef GPIO_InitStructure;

    /* PC14 在 APB2 的 GPIOC 域 */
    RCC_APB2PeriphClockCmd(RCC_APB2Periph_GPIOC, ENABLE);

    GPIO_InitStructure.GPIO_Pin   = LED_PIN;
    GPIO_InitStructure.GPIO_Mode  = GPIO_Mode_Out_PP;
    GPIO_InitStructure.GPIO_Speed = GPIO_Speed_2MHz;
    GPIO_Init(LED_PORT, &GPIO_InitStructure);

    LED_OFF();
}
