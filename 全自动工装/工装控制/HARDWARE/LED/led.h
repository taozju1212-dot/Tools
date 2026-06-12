#ifndef __LED_H
#define __LED_H
#include "sys.h"

/* 工装板单 LED 指示灯 - PC14 (VBAT 域，最大驱动 3mA / 2MHz) */
/* 电路：PC14 -> R21 -> LED3 阴极, 阳极接 VCC，低电平点亮 */
#define LED_PIN         GPIO_Pin_14
#define LED_PORT        GPIOC
#define LED1            PCout(14)

#define LED_ON()        (LED1 = 0)
#define LED_OFF()       (LED1 = 1)
#define LED_TOGGLE()    (LED1 = !LED1)

void LED_Init(void);

#endif
