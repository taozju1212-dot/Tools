#include "app_flash.h"
#include "stm32f10x_flash.h"

/* ============================================================================
 *  内部 Flash 持久化
 *  地址: 0x0800FC00 (64KB 末尾 1KB 页，页大小 1KB)
 * ============================================================================ */

#define FLASH_PAGE_ADDR  TUNECFG_FLASH_ADDR
#define FLASH_PAGE_SIZE  0x400U   /* 1 KB */

/* CRC32 (IEEE 802.3 / CRC-32/ISO-HDLC, poly 0xEDB88320) */
uint32_t CRC32_Calc(const uint8_t *data, uint32_t len)
{
    uint32_t crc = 0xFFFFFFFFUL;
    uint8_t i;
    while (len--) {
        crc ^= *data++;
        for (i = 0; i < 8; i++)
            crc = (crc >> 1) ^ (0xEDB88320UL & -(crc & 1));
    }
    return crc ^ 0xFFFFFFFFUL;
}

/* ---------- 从 Flash 读取配置，校验失败则填默认值 ---------- */
void Flash_Load_Config(void)
{
    const TuneCfg *p = (const TuneCfg *)FLASH_PAGE_ADDR;
    uint32_t crc_expected = CRC32_Calc((const uint8_t *)p,
                                        sizeof(TuneCfg) - sizeof(uint32_t));
    if (p->magic == TUNECFG_MAGIC && p->crc32 == crc_expected) {
        g_cfg = *p;
    } else {
        g_cfg.magic      = TUNECFG_MAGIC;
        g_cfg.m1_rpm     = M1_RPM_DEFAULT;
        g_cfg.m1_irun    = M1_IRUN_DEFAULT;
        g_cfg.m1_reserved= 0;
        g_cfg.m2_rpm     = M2_RPM_DEFAULT;
        g_cfg.m2_irun    = M2_IRUN_DEFAULT;
        g_cfg.m2_reserved= 0;
        g_cfg.crc32      = CRC32_Calc((const uint8_t *)&g_cfg,
                                       sizeof(TuneCfg) - sizeof(uint32_t));
    }
}

/* ---------- 擦页 + 写 TuneCfg，返回 1=成功 0=失败 ---------- */
uint8_t Flash_Save_Config(void)
{
    FLASH_Status st;
    const uint32_t *src;
    uint32_t addr;
    uint32_t i;
    const TuneCfg *p;
    uint32_t crc_check;

    g_cfg.magic = TUNECFG_MAGIC;
    g_cfg.crc32 = CRC32_Calc((const uint8_t *)&g_cfg,
                               sizeof(TuneCfg) - sizeof(uint32_t));

    FLASH_Unlock();

    st = FLASH_ErasePage(FLASH_PAGE_ADDR);
    if (st != FLASH_COMPLETE) { FLASH_Lock(); return 0; }

    src  = (const uint32_t *)&g_cfg;
    addr = FLASH_PAGE_ADDR;
    for (i = 0; i < sizeof(TuneCfg) / 4; i++) {
        st = FLASH_ProgramWord(addr, src[i]);
        if (st != FLASH_COMPLETE) { FLASH_Lock(); return 0; }
        addr += 4;
    }

    FLASH_Lock();

    p         = (const TuneCfg *)FLASH_PAGE_ADDR;
    crc_check = CRC32_Calc((const uint8_t *)p, sizeof(TuneCfg) - sizeof(uint32_t));
    return (p->magic == TUNECFG_MAGIC && p->crc32 == crc_check) ? 1 : 0;
}
