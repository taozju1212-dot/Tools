#include "app_scan.h"
#include "bsp.h"
#include "delay.h"

/* ============================================================================
 *  MZR813H 条码模块：USART2 自动波特率探测（115200 / 9600）
 *  - 不发送任何配置命令，不修改模块波特率
 *  - 首次按 SCN 时分别用 115200 和 9600 发触发命令，看哪个有回应
 *  - 锁定有回应的波特率，后续直接使用
 *
 *  开始扫描指令（手册 P8，指令模式触发）：
 *    02 00 01 05 00 04 11 00 00 01 10 03
 * ============================================================================ */

static const uint8_t SCAN_CMD[12] = {
    0x02, 0x00, 0x01, 0x05, 0x00, 0x04,
    0x11, 0x00, 0x00, 0x01, 0x10, 0x03
};

/* 已锁定的波特率；0=未探测 */
static volatile uint32_t g_learned_baud = 0;

/* ---------- 工具：动态设置 USART2 波特率（不改其他参数） ---------- */
static void usart2_set_baud(uint32_t baud)
{
    USART_InitTypeDef ui;
    USART_Cmd(USART2, DISABLE);
    ui.USART_BaudRate            = baud;
    ui.USART_WordLength          = USART_WordLength_8b;
    ui.USART_StopBits            = USART_StopBits_1;
    ui.USART_Parity              = USART_Parity_No;
    ui.USART_HardwareFlowControl = USART_HardwareFlowControl_None;
    ui.USART_Mode                = USART_Mode_Rx | USART_Mode_Tx;
    USART_Init(USART2, &ui);
    USART_Cmd(USART2, ENABLE);
}

static void scan_buf_reset(void)
{
    g_scan_buf.head = 0;
    g_scan_buf.tail = 0;
    g_scan_buf.len  = 0;
    g_scan_rx_flag  = 0;
}

static void usart2_send(uint8_t b)
{
    while (USART_GetFlagStatus(USART2, USART_FLAG_TXE) == RESET);
    USART_SendData(USART2, b);
}

static void usart2_send_n(const uint8_t *buf, uint8_t n)
{
    uint8_t i;
    for (i = 0; i < n; i++) usart2_send(buf[i]);
    while (USART_GetFlagStatus(USART2, USART_FLAG_TC) == RESET);
}

/* ---------- 在指定波特率下发触发命令，等 300ms 看回应 ---------- */
/* 返回 1 = 有回应（≥1 字节进入缓冲），0 = 超时无回应 */
static uint8_t scan_probe(uint32_t baud)
{
    uint32_t t0;

    usart2_set_baud(baud);
    delay_ms(20);                    /* 让总线稳定 */
    scan_buf_reset();
    usart2_send_n(SCAN_CMD, 12);

    /* 不同波特率下 12 字节发完时间不同：115200 ≈1ms, 9600 ≈12.5ms
       300ms 总等待，留足模块响应时间 */
    t0 = g_tick_ms;
    while (TICK_ELAPSED(t0) < 300) {
        if (g_scan_buf.len > 0) return 1;
    }
    return 0;
}

/* ---------- 初始化：不再发任何模块配置（避免修改模块状态） ---------- */
void Scan_Init(void)
{
    /* 保留接口但不动模块。波特率探测延后到首次 Scan_Start */
}

/* ---------- 触发扫码 ----------
 *  首次：探测 115200 → 9600，锁定有回应的那一个
 *  后续：直接用锁定的波特率发命令
 *  两个都失败：不锁定（下次 SCN 再试），当前停在 9600（最后试的波特率）
 */
void Scan_Start(void)
{
    if (g_learned_baud != 0) {
        /* 已锁定 — 直接发命令。USART 波特率已是 g_learned_baud */
        scan_buf_reset();
        usart2_send_n(SCAN_CMD, 12);
        return;
    }

    /* 第一阶段：试 115200（手册默认） */
    if (scan_probe(115200)) {
        g_learned_baud = 115200;
        return;
    }

    /* 第二阶段：试 9600 */
    if (scan_probe(9600)) {
        g_learned_baud = 9600;
        return;
    }

    /* 两个都没回应 — 不锁定，下次 SCN 再试一次
       主循环的 SCAN 超时会显示 SCAN FAIL */
}

void Scan_Process(char *result_buf)
{
    uint16_t len;
    uint16_t i;
    if (!result_buf) return;
    len = g_scan_buf.len;
    if (len >= SCAN_RESULT_MAX) len = SCAN_RESULT_MAX - 1;
    for (i = 0; i < len; i++) {
        uint16_t idx = (g_scan_buf.head + i) % SCAN_BUF_SIZE;
        result_buf[i] = (char)g_scan_buf.buf[idx];
    }
    result_buf[len] = '\0';
}
