#ifndef __APP_CONFIG_H
#define __APP_CONFIG_H

#include "stm32f10x.h"

#ifndef NULL
#define NULL ((void*)0)
#endif

/* ============================================================================
 *  全自动工装板 - 应用层共享配置
 *  所有状态机类型、全局变量声明、常量定义
 * ============================================================================ */

/* ---------- 状态机 ---------- */
typedef enum {
    STATE_HOME = 0,
    STATE_MOTOR_MENU,
    STATE_SPEED_EDIT,
    STATE_CUR_EDIT,
    STATE_PRINT_RUN,
    STATE_SCAN_RUN,
    STATE_CAL_STEP1,
    STATE_CAL_STEP2,
} AppState;

typedef enum {
    FOCUS_M1 = 0,
    FOCUS_M2,
    FOCUS_PRINT,
    FOCUS_CALIB,
    FOCUS_SCAN,
    FOCUS_COUNT
} HomeFocus;

typedef enum {
    MOTOR_FOCUS_SPEED = 0,
    MOTOR_FOCUS_CUR,
    MOTOR_FOCUS_COUNT
} MotorFocus;

/* ---------- 持久化配置结构 ---------- */
#define TUNECFG_MAGIC       0xA55A1234UL
#define TUNECFG_FLASH_ADDR  0x0800FC00UL   /* 64KB Flash 末尾 1KB 页 */

#define RPM_MIN             25
#define RPM_MAX             1500
#define RPM_STEP            25
#define IRUN_MIN            0
#define IRUN_MAX            31
#define IRUN_STEP           3              /* 编码器每格 ±3 (约 0.13A/格) */

/* ---------- 步进电机加减速斜波 ---------- */
#define RAMP_START_RPM      30             /* 启动 RPM (克服静摩擦) */
#define RAMP_ACCEL_RPM_S    1000           /* 加/减速率 RPM/秒 (1500 RPM 约 1.5 秒达到) */
#define RAMP_UPDATE_INTV_MS 5              /* 主循环斜波刷新间隔 */
#define M1_RPM_DEFAULT      100
#define M2_RPM_DEFAULT      100
#define M1_IRUN_DEFAULT     27             /* ~1.18A，目标 1.2A */
#define M2_IRUN_DEFAULT     9              /* ~0.42A，目标 0.4A */

typedef struct {
    uint32_t magic;
    uint16_t m1_rpm;
    uint8_t  m1_irun;
    uint8_t  m1_reserved;
    uint16_t m2_rpm;
    uint8_t  m2_irun;
    uint8_t  m2_reserved;
    uint32_t crc32;
} TuneCfg;

/* ---------- 扫码环形缓冲 ---------- */
#define SCAN_BUF_SIZE   128
typedef struct {
    uint8_t  buf[SCAN_BUF_SIZE];
    uint16_t head;
    uint16_t tail;
    uint16_t len;
} RingBuf;

/* ---------- 全局状态 ---------- */
extern volatile AppState  g_state;
extern volatile HomeFocus g_home_focus;
extern volatile MotorFocus g_motor_focus;
extern volatile uint8_t   g_edit_motor;   /* 1=M1, 2=M2 */
extern volatile uint16_t  g_draft_rpm;
extern volatile uint8_t   g_draft_irun;
extern TuneCfg             g_cfg;

/* ---------- 编码器 ---------- */
extern volatile int16_t   g_enc_last;     /* 上次读取的 TIM2 计数 */
extern volatile int8_t    g_enc_delta;    /* 本轮格数变化 (+/-N 格) */

/* ---------- 按键事件 (ISR 置位, 主循环清) ---------- */
extern volatile uint8_t   g_key_short;    /* KEY1 短按事件 */
extern volatile uint8_t   g_key_long;     /* KEY1 长按事件 */
extern volatile uint8_t   g_key_down;     /* KEY1 当前是否按下 */
extern volatile uint8_t   g_key_long_armed; /* 本次按键已触发长按，防重复 */
extern volatile uint32_t  g_key_press_ms; /* 按下时刻 (ms) */

/* ---------- 光开关状态 ---------- */
extern volatile uint8_t   g_opt_sw1;     /* 1=遮挡(✓), 0=通透(✗) */
extern volatile uint8_t   g_opt_sw2;

/* ---------- DO1 ---------- */
extern volatile uint8_t   g_do1_state;   /* PB0 当前电平 */

/* ---------- 循环模式（PA11 自锁键，按住=1，弹出=0，PC14 同步） ---------- */
extern volatile uint8_t   g_loop_mode;

/* ---------- 点动标志 (ISR 置位, 主循环消费) ---------- */
extern volatile uint8_t   g_jog_m1u;
extern volatile uint8_t   g_jog_m1d;
extern volatile uint8_t   g_jog_m2u;
extern volatile uint8_t   g_jog_m2d;
extern volatile uint8_t   g_jog_dirty;   /* 需要重新评估电机状态 */

/* ---------- 温度 ---------- */
extern volatile uint16_t  g_temp_raw;    /* ADC1_IN5 12-bit 原始值 */

/* ---------- 扫码缓冲 ---------- */
extern RingBuf             g_scan_buf;
extern volatile uint8_t   g_scan_rx_flag;/* 有新字节到达 */

/* ---------- TMC2226 回复缓冲 ---------- */
#define TMC_BUF_SIZE  16
extern uint8_t             g_tmc_rx_buf[TMC_BUF_SIZE];
extern volatile uint8_t    g_tmc_rx_cnt;

/* ---------- 系统时基 ---------- */
extern volatile uint32_t   g_tick_ms;    /* TIM4 每 1ms 自增 */

/* ---------- 空闲超时 (ms，任意按键/旋钮清零) ---------- */
extern volatile uint32_t   g_idle_ms;
#define IDLE_TIMEOUT_MS     30000

/* ---------- UI 脏区标志 ---------- */
extern volatile uint8_t    g_dirty_all;  /* 整屏刷新 */
extern volatile uint8_t    g_dirty_m1;
extern volatile uint8_t    g_dirty_m2;
extern volatile uint8_t    g_dirty_center;
extern volatile uint8_t    g_dirty_right;
extern volatile uint8_t    g_dirty_status;

/* ---------- 辅助宏 ---------- */
#define CLAMP(v, lo, hi)   ((v) < (lo) ? (lo) : ((v) > (hi) ? (hi) : (v)))
#define TICK_ELAPSED(t0)   ((uint32_t)(g_tick_ms - (t0)))

#endif /* __APP_CONFIG_H */
