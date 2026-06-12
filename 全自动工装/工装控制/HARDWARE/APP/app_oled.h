#ifndef __APP_OLED_H
#define __APP_OLED_H

#include "app_config.h"
#include "app_motor.h"

/* ============================================================================
 *  SSD1306/SH1106 128×64 OLED I2C 驱动 + 全自动工装 UI
 * ============================================================================ */

#define OLED_W     128
#define OLED_H     64
#define OLED_PAGES 8    /* 64 / 8 */
#define OLED_ADDR  0x78 /* 7-bit 0x3C → 8-bit write 0x78 */

/* SH1106 列偏移（RAM 从列 2 开始）；SSD1306 设为 0 */
#define OLED_COL_OFFSET  2

/* ---------- 驱动层 ---------- */
void OLED_Init(void);
void OLED_Flush(void);          /* 把脏页写到屏幕 */
void OLED_Clear(void);          /* 清帧缓存并刷屏 */
void OLED_MarkDirty(void);      /* 标记所有页脏 */

/* ---------- 绘图原语 ---------- */
void FB_SetPixel(uint8_t x, uint8_t y);
void FB_ClrPixel(uint8_t x, uint8_t y);
void FB_FillRect(uint8_t x, uint8_t y, uint8_t w, uint8_t h, uint8_t fill);
void FB_DrawRect(uint8_t x, uint8_t y, uint8_t w, uint8_t h, uint8_t thick);

/* ---------- 文字（ASCII 8×16, 数字 8×16） ---------- */
void FB_DrawChar8x16(uint8_t x, uint8_t y, char c);
void FB_DrawStr8x16(uint8_t x, uint8_t y, const char *s);

/* ---------- 汉字 16×16（索引 CJKIDX_xxx） ---------- */
void FB_DrawCJK(uint8_t x, uint8_t y, uint8_t idx);
void FB_DrawCJKStr(uint8_t x, uint8_t y, const uint8_t *idx_arr, uint8_t cnt);

/* ---------- CJK 字符索引表 ---------- */
#define CJKIDX_DA    0   /* 打 */
#define CJKIDX_YIN   1   /* 印 */
#define CJKIDX_JIAO  2   /* 校 */
#define CJKIDX_ZHUN  3   /* 准 */
#define CJKIDX_SAOQ  4   /* 扫 */
#define CJKIDX_MA    5   /* 码 */
#define CJKIDX_ZHUAN 6   /* 转 */
#define CJKIDX_SU    7   /* 速 */
#define CJKIDX_DIAN  8   /* 电 */
#define CJKIDX_LIU   9   /* 流 */
#define CJKIDX_ZHONG 10  /* 中 */
#define CJKIDX_WAN   11  /* 完 */
#define CJKIDX_CHENG 12  /* 成 */
#define CJKIDX_GONG  13  /* 功 */
#define CJKIDX_TIAO  14  /* 条 */
#define CJKIDX_NEI   15  /* 内 */
#define CJKIDX_RONG  16  /* 容 */
#define CJKIDX_SHI   17  /* 失 */
#define CJKIDX_BAI   18  /* 败 */
#define CJKIDX_YI    19  /* 移 */
#define CJKIDX_KAI   20  /* 开 */
#define CJKIDX_GUAN  21  /* 管 */
#define CJKIDX_HOU   22  /* 后 */
#define CJKIDX_ZAI   23  /* 再 */
#define CJKIDX_CI    24  /* 次 */
#define CJKIDX_AN    25  /* 按 */
#define CJKIDX_JIAN  26  /* 键 */
#define CJKIDX_QING  27  /* 请 */
#define CJKIDX_CHA   28  /* 查 */
#define CJKIDX_JU    29  /* 距 */
#define CJKIDX_LI    30  /* 离 */
#define CJKIDX_DU    31  /* 读 */
#define CJKIDX_QU    32  /* 取 */
#define CJKIDX_YANG  33  /* 样 */
#define CJKIDX_BEN   34  /* 本 */
#define CJKIDX_SHI2  35  /* 试 */
#define CJKIDX_JIAN2 36  /* 检 */
#define CJKIDX_DA2   37  /* 达 (DA label) */
#define CJKIDX_GUANG 38  /* 光 */
#define CJKIDX_KAI2  39  /* 开 (duplicated for clarity) */
/* 新增（全屏 UI 使用，字模需 PCtoLCD2002 提取替换占位）*/
#define CJKIDX_DU2   40  /* 度 */
#define CJKIDX_TONG  41  /* 通 */
#define CJKIDX_XUN   42  /* 讯 */
#define CJKIDX_LIAN  43  /* 连 */
#define CJKIDX_JIE   44  /* 接 */
#define CJKIDX_CHUAN 45  /* 传 */
#define CJKIDX_GAN   46  /* 感 */
#define CJKIDX_QI    47  /* 器 */
#define CJKIDX_KAO   48  /* 靠 */
#define CJKIDX_JIN   49  /* 近 */
#define CJKIDX_XIA   50  /* 下 */
#define CJKIDX_HUI   51  /* 回 */
#define CJK_CHAR_COUNT 52

/* ---------- UI 渲染（每个状态各一个函数） ---------- */
void UI_Render_Home(void);
void UI_Render_MotorMenu(void);
void UI_Render_SpeedEdit(void);
void UI_Render_CurEdit(void);
void UI_Render_Print_Running(void);
void UI_Render_Print_Done(void);
void UI_Render_Scan(const char *result);  /* result=NULL → 等待中 */
void UI_Render_Scan_Fail(void);
void UI_Render_Cal_Step1_OK(void);                              /* 校准成功 再次按下回读 */
void UI_Render_Cal_Step1_Fail(void);                            /* 校准失败 靠近传感器 */
void UI_Render_Cal_Comm_Fail(void);                             /* 通讯失败 连接传感器 */
void UI_Render_Cal_Step2_OK(uint8_t agc, uint16_t led, uint8_t ref);
void UI_Render_Cal_Step2_Fail(void);
void UI_Render_Overlay_Temp(void);        /* 仅刷顶栏温度 */
void UI_Render_Overlay_Status(void);      /* 仅刷 O1/O2 状态条 */
void UI_BlinkValue_Speed(uint8_t on);     /* 编辑态：仅刷新左面板 RPM 值 */
void UI_BlinkValue_Cur(uint8_t on);       /* 编辑态：仅刷新右面板 AMP 值 */

#endif /* __APP_OLED_H */
