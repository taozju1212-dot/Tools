#include "app_print.h"
#include "bsp.h"
#include "delay.h"

/* ============================================================================
 *  热敏打印机：USART3 9600 8N1
 *  输出 5 行阶梯字母
 * ============================================================================ */

static void usart3_send(uint8_t b)
{
    while (USART_GetFlagStatus(USART3, USART_FLAG_TXE) == RESET);
    USART_SendData(USART3, b);
}

static void usart3_send_str(const char *s)
{
    while (*s) usart3_send((uint8_t)*s++);
}

void Print_Run(void)
{
    static const char * const lines[] = {
        "A\r\n",
        "BB\r\n",
        "CCC\r\n",
        "DDDD\r\n",
        "EEEEE\r\n",
    };
    {
        uint8_t i;
        for (i = 0; i < 5; i++) {
            usart3_send_str(lines[i]);
            delay_ms(50);
        }
    }
}
