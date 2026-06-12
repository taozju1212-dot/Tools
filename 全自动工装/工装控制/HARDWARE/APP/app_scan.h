#ifndef __APP_SCAN_H
#define __APP_SCAN_H

#include "app_config.h"

void Scan_Init(void);                     /* 上电配置：发送电平模式选择命令（一次性） */
void Scan_Start(void);                    /* 发送触发命令 */
void Scan_Process(char *result_buf);      /* 轮询处理：result_buf 传 NULL 时仅处理；
                                             填充最多 SCAN_RESULT_MAX 字节结果 */
#define SCAN_RESULT_MAX  64

#endif /* __APP_SCAN_H */
