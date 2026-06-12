#include "app_config.h"

/* ============================================================================
 *  全局变量定义
 * ============================================================================ */

volatile AppState   g_state       = STATE_HOME;
volatile HomeFocus  g_home_focus  = FOCUS_COUNT;  /* 无默认选中；首次转编码器后激活 */
volatile MotorFocus g_motor_focus = MOTOR_FOCUS_SPEED;
volatile uint8_t    g_edit_motor  = 1;
volatile uint16_t   g_draft_rpm   = M1_RPM_DEFAULT;
volatile uint8_t    g_draft_irun  = M1_IRUN_DEFAULT;
TuneCfg             g_cfg;

volatile int16_t    g_enc_last    = 0;
volatile int8_t     g_enc_delta   = 0;

volatile uint8_t    g_key_short      = 0;
volatile uint8_t    g_key_long       = 0;
volatile uint8_t    g_key_down       = 0;
volatile uint8_t    g_key_long_armed = 0;
volatile uint32_t   g_key_press_ms   = 0;

volatile uint8_t    g_opt_sw1     = 0;
volatile uint8_t    g_opt_sw2     = 0;

volatile uint8_t    g_do1_state   = 0;
volatile uint8_t    g_loop_mode   = 0;

volatile uint8_t    g_jog_m1u     = 0;
volatile uint8_t    g_jog_m1d     = 0;
volatile uint8_t    g_jog_m2u     = 0;
volatile uint8_t    g_jog_m2d     = 0;
volatile uint8_t    g_jog_dirty   = 0;

volatile uint16_t   g_temp_raw    = 0;

RingBuf             g_scan_buf;
volatile uint8_t    g_scan_rx_flag= 0;

uint8_t             g_tmc_rx_buf[TMC_BUF_SIZE];
volatile uint8_t    g_tmc_rx_cnt  = 0;

volatile uint32_t   g_tick_ms     = 0;
volatile uint32_t   g_idle_ms     = 0;

volatile uint8_t    g_dirty_all   = 1;
volatile uint8_t    g_dirty_m1    = 0;
volatile uint8_t    g_dirty_m2    = 0;
volatile uint8_t    g_dirty_center= 0;
volatile uint8_t    g_dirty_right = 0;
volatile uint8_t    g_dirty_status= 0;
