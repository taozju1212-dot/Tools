#ifndef __APP_CALIB_H
#define __APP_CALIB_H

#include "app_config.h"

/* ============================================================================
 *  OPB9000 光纤传感器校准
 *  PB5 = CAL 输出 (Manchester bit-bang)
 *  PB4 = OUT 输入 (反馈电平)
 *  位周期约 2ms（半位 1ms）
 * ============================================================================ */

/* 校准 STEP1：写 Bank2 REF/OP，检查 PB4 翻转
   返回 1=成功, 0=失败 */
uint8_t Calib_Step1(void);

/* 校准 STEP2：发 RR 命令读回，解析 AGC / LED / REF
   返回 1=成功, 0=失败; 成功时写 *agc, *led, *ref */
uint8_t Calib_Step2(uint8_t *agc, uint16_t *led, uint8_t *ref);

#endif /* __APP_CALIB_H */
