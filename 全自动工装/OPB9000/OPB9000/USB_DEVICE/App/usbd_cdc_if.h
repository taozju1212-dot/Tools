/* USER CODE BEGIN Header */
/**
  ******************************************************************************
  * @file           : usbd_cdc_if.h
  * @version        : v2.0_Cube
  * @brief          : Header for usbd_cdc_if.c file.
  ******************************************************************************
  * @attention
  *
  * Copyright (c) 2024 STMicroelectronics.
  * All rights reserved.
  *
  * This software is licensed under terms that can be found in the LICENSE file
  * in the root directory of this software component.
  * If no LICENSE file comes with this software, it is provided AS-IS.
  *
  ******************************************************************************
  */
/* USER CODE END Header */

/* Define to prevent recursive inclusion -------------------------------------*/
#ifndef __USBD_CDC_IF_H__
#define __USBD_CDC_IF_H__

#ifdef __cplusplus
 extern "C" {
#endif

/* Includes ------------------------------------------------------------------*/
#include "usbd_cdc.h"

/* USER CODE BEGIN INCLUDE */

/* USER CODE END INCLUDE */

/** @addtogroup STM32_USB_OTG_DEVICE_LIBRARY
  * @brief For Usb device.
  * @{
  */

/** @defgroup USBD_CDC_IF USBD_CDC_IF
  * @brief Usb VCP device module
  * @{
  */

/** @defgroup USBD_CDC_IF_Exported_Defines USBD_CDC_IF_Exported_Defines
  * @brief Defines.
  * @{
  */
/* Define size for the receive and transmit buffer over CDC */
#define APP_RX_DATA_SIZE  1024
#define APP_TX_DATA_SIZE  1024
/* USER CODE BEGIN EXPORTED_DEFINES */

/* USER CODE END EXPORTED_DEFINES */

/**
  * @}
  */

/** @defgroup USBD_CDC_IF_Exported_Types USBD_CDC_IF_Exported_Types
  * @brief Types.
  * @{
  */

/* USER CODE BEGIN EXPORTED_TYPES */
typedef enum
{
  USB_CDC_Receive_Idle = 0x00U,
  USB_CDC_Receive_Busy = 0x01U,
  USB_CDC_Receive_Complete = 0x02U
}USB_CDC_ReceiveStateTypeDef;
/* USER CODE END EXPORTED_TYPES */

/**
  * @}
  */

/** @defgroup USBD_CDC_IF_Exported_Macros USBD_CDC_IF_Exported_Macros
  * @brief Aliases.
  * @{
  */

/* USER CODE BEGIN EXPORTED_MACRO */
#define USB_CDC_TRANSMIT_GROOVE_SIZE  20      //usb发送的最大条数
#define USB_CDC_TRANSMIT_BUFFER_SIZE  64      //usb发送的最大字节数
#define USB_CDC_RECEIVE_GROOVE_SIZE   10      //usb接收的最大条数
#define USB_CDC_RECEIVE_BUFFER_SIZE   64      //usb接收的最大字节数

#define USB_CDC_TrainsmitDataWriteDone(  ) \
        u8UsbCdcWriteBufferIndex = ( (u8UsbCdcWriteBufferIndex+1) == USB_CDC_TRANSMIT_GROOVE_SIZE ) ? 0 : (u8UsbCdcWriteBufferIndex+1);\
        u8UsbCdcTransmitMessageNumber++
        
#define USB_CDC_TrainsmitDataDone(  ) \
        u8UsbCdcTransmitGrooveIndex = ( (u8UsbCdcTransmitGrooveIndex+1) == USB_CDC_TRANSMIT_GROOVE_SIZE ) ? 0 : (u8UsbCdcTransmitGrooveIndex+1);\
        u8UsbCdcTransmitMessageNumber--
        
#define USB_CDC_ReceiveDataDone(  ) \
        u8UsbCdcReceiveGrooveIndex = ( (u8UsbCdcReceiveGrooveIndex+1) == USB_CDC_RECEIVE_GROOVE_SIZE ) ? 0 : (u8UsbCdcReceiveGrooveIndex+1);\
        u8UsbCdcReceiveMessageNumber++
        
#define USB_CDC_ReceiveDataReadDone(  ) \
        u8UsbCdcReadBufferIndex = ( (u8UsbCdcReadBufferIndex+1) == USB_CDC_RECEIVE_GROOVE_SIZE ) ? 0 : (u8UsbCdcReadBufferIndex+1);\
        u8UsbCdcReceiveMessageNumber--
        
/* USER CODE END EXPORTED_MACRO */

/**
  * @}
  */

/** @defgroup USBD_CDC_IF_Exported_Variables USBD_CDC_IF_Exported_Variables
  * @brief Public variables.
  * @{
  */

/** CDC Interface callback. */
extern USBD_CDC_ItfTypeDef USBD_Interface_fops_FS;

/* USER CODE BEGIN EXPORTED_VARIABLES */
extern uint8_t u8UsbCdcTransmitBuffer[USB_CDC_TRANSMIT_GROOVE_SIZE][USB_CDC_TRANSMIT_BUFFER_SIZE]; //发送数据缓冲区
extern uint8_t u8UsbCdcTransmitLength[USB_CDC_TRANSMIT_GROOVE_SIZE];    //发送缓冲区长度
extern uint8_t u8UsbCdcTransmitGrooveIndex;    //发送的第几个缓冲
extern uint8_t u8UsbCdcWriteBufferIndex;       //写入第几个缓冲
extern uint8_t u8UsbCdcTransmitMessageNumber;  //需要传输的消息数

extern uint8_t u8UsbCdcReceiveBuffer[USB_CDC_RECEIVE_GROOVE_SIZE][USB_CDC_RECEIVE_BUFFER_SIZE];  //接收数据缓冲区
extern uint8_t u8UsbCdcReceiveLength[USB_CDC_RECEIVE_GROOVE_SIZE];   //接收缓冲区长度
extern uint8_t u8UsbCdcReceiveGrooveIndex;   //接收的第几个缓冲
extern uint8_t u8UsbCdcReadBufferIndex;      //读取第几个缓冲
extern uint8_t u8UsbCdcReceiveMessageNumber; //需要处理的消息数

extern USB_CDC_ReceiveStateTypeDef UsbCdcReceiveState;
/* USER CODE END EXPORTED_VARIABLES */

/**
  * @}
  */

/** @defgroup USBD_CDC_IF_Exported_FunctionsPrototype USBD_CDC_IF_Exported_FunctionsPrototype
  * @brief Public functions declaration.
  * @{
  */

uint8_t CDC_Transmit_FS(uint8_t* Buf, uint16_t Len);

/* USER CODE BEGIN EXPORTED_FUNCTIONS */
const uint8_t *getUSB_CDC_ReceiveBuffer(void);
uint8_t getUSB_CDC_ReceiveLength(void);
void USB_CDC_SendData(uint8_t *Buf, uint32_t Len);
int USB_CDC_FillBuffer(char *fmt, ...);
int USB_CDC_Ptintf(char *fmt, ...);
/* USER CODE END EXPORTED_FUNCTIONS */

/**
  * @}
  */

/**
  * @}
  */

/**
  * @}
  */

#ifdef __cplusplus
}
#endif

#endif /* __USBD_CDC_IF_H__ */

