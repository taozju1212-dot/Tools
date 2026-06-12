#ifndef __APP_FLASH_H
#define __APP_FLASH_H

#include "app_config.h"

void     Flash_Load_Config(void);
uint8_t  Flash_Save_Config(void);
uint32_t CRC32_Calc(const uint8_t *data, uint32_t len);

#endif /* __APP_FLASH_H */
