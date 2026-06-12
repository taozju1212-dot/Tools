#include "app_calib.h"
#include "bsp.h"
#include "delay.h"

/* ============================================================================
 *  OPB9000 Manchester bit-bang
 *  Manchester (IEEE 802.3): 0 = H→L; 1 = L→H（位中点跳变）
 *  半位周期 = 1ms（delay_ms(1)）
 *  CAL 线：PB5（推挽输出）
 *  OUT 线：PB4（浮空输入，外部上拉）
 * ============================================================================ */

/* 输出一个 Manchester 编码位（发送到 PB5） */
static void man_send_bit(uint8_t bit)
{
    if (bit) {
        OPB9000_CAL(0);  /* 前半低 */
        delay_ms(1);
        OPB9000_CAL(1);  /* 后半高 */
        delay_ms(1);
    } else {
        OPB9000_CAL(1);  /* 前半高 */
        delay_ms(1);
        OPB9000_CAL(0);  /* 后半低 */
        delay_ms(1);
    }
}

/* 发送 N 位（MSB first） */
static void man_send_bits(uint32_t data, uint8_t bits)
{
    int8_t i;
    for (i = bits - 1; i >= 0; i--)
        man_send_bit((data >> i) & 1);
}

/* 接收一位（在位中点采样 PB4） */
static uint8_t man_recv_bit(void)
{
    uint8_t v;
    delay_ms(1);         /* 等过前半 */
    v = OPB9000_OUT();
    delay_ms(1);         /* 等过后半 */
    return v;
}

/* 接收 N 位，MSB first */
static uint32_t man_recv_bits(uint8_t bits)
{
    uint32_t result = 0;
    uint8_t i;
    for (i = 0; i < bits; i++)
        result = (result << 1) | man_recv_bit();
    return result;
}

/* ============================================================================
 *  OPB9000 命令帧格式（参考原型工程）
 *  CS (Write) 命令帧: PREAMBLE(8) | CMD(4) | BANK(2) | REF(4) | DS(1) | OP(1) | ...
 *  简化版：发 SYNC+CMD+BANK+DATA 共 12 bit，等待 ACKOP=1
 *
 *  本实现参考 OPB9000 产品手册和原型 main.c：
 *  - PREAMBLE: 8 个 1（同步头）
 *  - CMD:  1100 = CS-Write,  0001 = RR-Read
 *  - BANK: 10 = Bank2
 *  - DATA: REF[3:0] DS OP
 *
 *  RR 响应（37 bit）:
 *  StartBit(1) | Bank1: CA(1) AGC(2) LED(10) | Bank2: REF(4) DS(1) OP(1) | Bank3: StartBit(1) EF(1) | 填充(15)
 * ============================================================================ */

/* 发送同步头（8 个 1） */
static void send_preamble(void)
{
    uint8_t i;
    for (i = 0; i < 8; i++) man_send_bit(1);
}

/* 校准 STEP1：写 Bank2 */
static uint8_t write_bank2(uint8_t ref, uint8_t op)
{
    uint32_t t0;
    uint8_t prev;

    send_preamble();
    /* CMD = 0xC (1100b, CS-Write), BANK = 0x2 (10b, Bank2) */
    man_send_bits(0x0C, 4);
    man_send_bits(0x02, 2);
    /* DATA: REF[3:0] | DS=0 | OP */
    man_send_bits(ref & 0x0F, 4);
    man_send_bit(0);      /* DS = 0 */
    man_send_bit(op & 1);
    OPB9000_CAL(0);       /* 空闲低 */

    /* 等待 PB4 OUT 响应（最多 100ms） */
    t0 = g_tick_ms;
    prev = OPB9000_OUT();
    while (TICK_ELAPSED(t0) < 100) {
        uint8_t cur = OPB9000_OUT();
        if (cur != prev) return 1;   /* 检测到翻转 */
        delay_ms(2);
    }
    return 0;
}

uint8_t Calib_Step1(void)
{
    OPB9000_CAL(0);
    delay_ms(5);

    /* 写 REF=15, OP=1（最高灵敏度，反向输出） */
    if (!write_bank2(15, 1)) return 0;
    delay_ms(10);

    /* 写 REF=4, OP=0（正常极性） */
    if (!write_bank2(4, 0)) return 0;

    return 1;
}

/* ============================================================================
 *  校准 STEP2：读回 RR 响应，解析 DA/GA
 * ============================================================================ */
uint8_t Calib_Step2(uint8_t *agc, uint16_t *led, uint8_t *ref)
{
    uint32_t resp_hi;
    uint8_t  resp_lo;
    uint8_t start1, agc_v, ref_v, op_rb, ef;
    uint16_t led_v;

    send_preamble();
    /* CMD = 0x1 (0001b, RR-Read), BANK 不需要指定（全读） */
    man_send_bits(0x01, 4);
    OPB9000_CAL(0);
    delay_ms(5);      /* 等待传感器准备响应 */

    /* 接收 37 位响应 */
    /* StartBit(1) | CA(1) | AGC(2) | LED(10) | REF(4) | DS(1) | OP(1) | StartBit(1) | EF(1) | 填充(15) */
    resp_hi = man_recv_bits(32);         /* 前 32 位 */
    resp_lo = (uint8_t)man_recv_bits(5); /* 后 5 位 */

    start1 = (resp_hi >> 31) & 1;
    agc_v  = (uint8_t)((resp_hi >> 28) & 0x03);
    led_v  = (uint16_t)((resp_hi >> 18) & 0x3FF);
    ref_v  = (uint8_t)((resp_hi >> 14) & 0x0F);
    op_rb  = (uint8_t)((resp_hi >> 12) & 0x01);
    ef     = (uint8_t)((resp_hi >> 10) & 0x01);
    (void)resp_lo; (void)start1; (void)op_rb;

    /* EF=1 or verification failure */
    if (ef) return 0;

    *agc = agc_v;
    *led = led_v;
    *ref = ref_v;
    return 1;
}
