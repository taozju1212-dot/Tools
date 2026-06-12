#ifndef __OPB9000_H__
#define __OPB9000_H__

#include "stm32f1xx_hal.h"

//#define __CAL_SET(x)    do { \
//                          GPIOB->BSRR = (uint32_t)GPIO_PIN_8 << 16u; \
//                          OPB9000_Delay(); \
//                          GPIOB->BSRR = (uint32_t)GPIO_PIN_8; \
//                          OPB9000_Delay(); \
//                        } while(0U)

//#define __CAL_RESET(x)  do { \
//                          GPIOB->BSRR = (uint32_t)GPIO_PIN_8; \
//                          OPB9000_Delay(); \
//                          GPIOB->BSRR = (uint32_t)GPIO_PIN_8 << 16u; \
//                          OPB9000_Delay(); \
//                        } while(0U)

//#define __CAL_IN()      do { \
//                          GPIOB->CRH&=0XFFFFFFF0; \
//                          GPIOB->CRH|=(uint32_t)3<<0; \
//                        } while(0U)

//#define __CAL_OUT()     do { \
//                          GPIOB->CRH&=0XFFFFFFF0; \
//                          GPIOB->CRH|=(uint32_t)8<<0; \
//                        } while(0U)

//#define __CAL_GET()     (GPIOB->IDR & GPIO_PIN_8)

//#define __OUT_GET()     (GPIOB->IDR & GPIO_PIN_9)

typedef enum{
  OPB9000_Com_Status_Idle               = 0x00U,
  OPB9000_Com_Status_TransmitWaiting    = 0x01U,
  OPB9000_Com_Status_Transmitting       = 0x02U,
  OPB9000_Com_Status_TransmitDone       = 0x04U,
  OPB9000_Com_Status_ReceiveWaiting     = 0x08U,
  OPB9000_Com_Status_Receiving          = 0x10U,
  OPB9000_Com_Status_ReceiveDone        = 0x20U,
  OPB9000_Com_Status_Done               = 0x40U
}OPB9000_Com_Status_t;  //OPB9000 Communication Status

typedef enum{
  OPB9000_Cmd_Status_Idle               = 0x00U,
  OPB9000_Cmd_Status_Reserved           = 0x01U,
  OPB9000_Cmd_Status_ReadRequest        = 0x02U,
  OPB9000_Cmd_Status_WriteBank2bits     = 0x04U,
  OPB9000_Cmd_Status_CalibrateRequest   = 0x08U
}OPB9000_Cmd_Status_t;  //OPB9000 Command Status

typedef struct
{
  union {
    uint16_t Data;
    struct {
      uint16_t StartBit :1;     //起始位
      uint16_t ErrorFlag:1;     //错误标志
      uint16_t Reserved :9;
    }Bank3Bit;
  }Bank3Data;
  
  union {
    uint16_t Data;
    struct {
      uint16_t CA       :1;     //校准成功标志
      uint16_t AGC      :2;     //自动增益控制？
      uint16_t LED      :10;    //LED驱动计数
    }Bank1Bit;
  }Bank1Data;
  
  union {
    uint16_t Data;
    struct {
      uint16_t REF      :4;     //比较器参考级别
      uint16_t DS       :1;     //漏极选择  0 推挽 1 开漏
      uint16_t OP       :1;     //输出类型
      uint16_t Reserved :7;
    }Bank2Bit;
  }Bank2Data;
  
}OPB9000_Bank_Bit_Data_t;

typedef enum{
  OPB9000_Cal_Status_Undef              = 0x00U,  //用于强制改变电平
  OPB9000_Cal_Status_Prepare            = 0x01U,
  OPB9000_Cal_Status_Turn               = 0x02U,
  OPB9000_Status_Status_Prepare         = 0x04U,
  OPB9000_Status_Status_Turn            = 0x08U,
  OPB9000_Status_Status_Respond         = 0x10U
}OPB9000_Cal_Status_t;  //OPB9000 Cal\Status Status

typedef enum{
  OPB9000_Out_Status_Reset              = 0x00U,
  OPB9000_Out_Status_Set                = 0x01U,
  OPB9000_Out_Status_Start              = 0x02U,
  OPB9000_Out_Status_Prepare            = 0x04U,
  OPB9000_Out_Status_Turn               = 0x08U
}OPB9000_Out_Status_t;  //OPB9000 Output Status

typedef enum{
  OPB9000_Cal_Result_Calibrating			  = 0x00U,
  OPB9000_Cal_Result_Unsuccessful       = 0x01U,
  OPB9000_Cal_Result_Successful         = 0x02U
}OPB9000_Cal_Result_t;


extern OPB9000_Com_Status_t     OPB9000_Com_Status;
extern OPB9000_Cmd_Status_t     OPB9000_Cmd_Status;
extern OPB9000_Bank_Bit_Data_t  OPB9000_Bank_Data_Value;
extern OPB9000_Cal_Status_t     OPB9000_Cal_Status;
extern OPB9000_Out_Status_t     OPB9000_Out_Status;
extern OPB9000_Cal_Result_t     OPB9000_Cal_Result;

#define ManchesterDataSize  2000
extern uint16_t ManchesterData[ManchesterDataSize];

uint64_t getOPB9000Data(void);
uint32_t getOPB9000Data2(uint8_t index);
void setOPB9000Bank2(uint16_t REF, uint16_t DS, uint16_t OP);
void OPB9000_Read_Request(void);
void OPB9000_WriteBank2bits(void);
void OPB9000_Calibrate_Request(void);
void OPB9000_Calibrate_Stand(void);

#endif
