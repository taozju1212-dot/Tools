/* USER CODE BEGIN Header */
/**
  ******************************************************************************
  * @file           : usbd_cdc_if.c
  * @version        : v2.0_Cube
  * @brief          : Usb device for Virtual Com Port.
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

/* Includes ------------------------------------------------------------------*/
#include "usbd_cdc_if.h"

/* USER CODE BEGIN INCLUDE */
#include <stdio.h>
#include <string.h>
#include <stdarg.h>

/* USER CODE END INCLUDE */

/* Private typedef -----------------------------------------------------------*/
/* Private define ------------------------------------------------------------*/
/* Private macro -------------------------------------------------------------*/

/* USER CODE BEGIN PV */
/* Private variables ---------------------------------------------------------*/
uint8_t u8UsbCdcTransmitBuffer[USB_CDC_TRANSMIT_GROOVE_SIZE][USB_CDC_TRANSMIT_BUFFER_SIZE]; //发送数据缓冲区
uint8_t u8UsbCdcTransmitLength[USB_CDC_TRANSMIT_GROOVE_SIZE];    //发送缓冲区长度
uint8_t u8UsbCdcTransmitGrooveIndex;    //发送的第几个缓冲
uint8_t u8UsbCdcWriteBufferIndex;       //写入第几个缓冲
uint8_t u8UsbCdcTransmitMessageNumber;  //需要传输的消息数

uint8_t u8UsbCdcReceiveBuffer[USB_CDC_RECEIVE_GROOVE_SIZE][USB_CDC_RECEIVE_BUFFER_SIZE];  //接收数据缓冲区
uint8_t u8UsbCdcReceiveLength[USB_CDC_RECEIVE_GROOVE_SIZE];   //接收缓冲区长度
uint8_t u8UsbCdcReceiveGrooveIndex;   //接收的第几个缓冲
uint8_t u8UsbCdcReadBufferIndex;      //读取第几个缓冲
uint8_t u8UsbCdcReceiveMessageNumber; //需要处理的消息数

USB_CDC_ReceiveStateTypeDef UsbCdcReceiveState;
/* USER CODE END PV */

/** @addtogroup STM32_USB_OTG_DEVICE_LIBRARY
  * @brief Usb device library.
  * @{
  */

/** @addtogroup USBD_CDC_IF
  * @{
  */

/** @defgroup USBD_CDC_IF_Private_TypesDefinitions USBD_CDC_IF_Private_TypesDefinitions
  * @brief Private types.
  * @{
  */

/* USER CODE BEGIN PRIVATE_TYPES */

/* USER CODE END PRIVATE_TYPES */

/**
  * @}
  */

/** @defgroup USBD_CDC_IF_Private_Defines USBD_CDC_IF_Private_Defines
  * @brief Private defines.
  * @{
  */

/* USER CODE BEGIN PRIVATE_DEFINES */
/* USER CODE END PRIVATE_DEFINES */

/**
  * @}
  */

/** @defgroup USBD_CDC_IF_Private_Macros USBD_CDC_IF_Private_Macros
  * @brief Private macros.
  * @{
  */

/* USER CODE BEGIN PRIVATE_MACRO */

/* USER CODE END PRIVATE_MACRO */

/**
  * @}
  */

/** @defgroup USBD_CDC_IF_Private_Variables USBD_CDC_IF_Private_Variables
  * @brief Private variables.
  * @{
  */
/* Create buffer for reception and transmission           */
/* It's up to user to redefine and/or remove those define */
/** Received data over USB are stored in this buffer      */
uint8_t UserRxBufferFS[APP_RX_DATA_SIZE];

/** Data to send over USB CDC are stored in this buffer   */
uint8_t UserTxBufferFS[APP_TX_DATA_SIZE];

/* USER CODE BEGIN PRIVATE_VARIABLES */
USBD_CDC_LineCodingTypeDef USBD_CDC_LineCoding =
{
  115200,      // 默认波特率
  0X00,        // 1位停止位
  0X00,        // 无奇偶校
  0X08,        // 无流控，8bit数据位
};
/* USER CODE END PRIVATE_VARIABLES */

/**
  * @}
  */

/** @defgroup USBD_CDC_IF_Exported_Variables USBD_CDC_IF_Exported_Variables
  * @brief Public variables.
  * @{
  */

extern USBD_HandleTypeDef hUsbDeviceFS;

/* USER CODE BEGIN EXPORTED_VARIABLES */

/* USER CODE END EXPORTED_VARIABLES */

/**
  * @}
  */

/** @defgroup USBD_CDC_IF_Private_FunctionPrototypes USBD_CDC_IF_Private_FunctionPrototypes
  * @brief Private functions declaration.
  * @{
  */

static int8_t CDC_Init_FS(void);
static int8_t CDC_DeInit_FS(void);
static int8_t CDC_Control_FS(uint8_t cmd, uint8_t* pbuf, uint16_t length);
static int8_t CDC_Receive_FS(uint8_t* pbuf, uint32_t *Len);

/* USER CODE BEGIN PRIVATE_FUNCTIONS_DECLARATION */

/* USER CODE END PRIVATE_FUNCTIONS_DECLARATION */

/**
  * @}
  */

USBD_CDC_ItfTypeDef USBD_Interface_fops_FS =
{
  CDC_Init_FS,
  CDC_DeInit_FS,
  CDC_Control_FS,
  CDC_Receive_FS
};

/* Private functions ---------------------------------------------------------*/
/**
  * @brief  Initializes the CDC media low layer over the FS USB IP
  * @retval USBD_OK if all operations are OK else USBD_FAIL
  */
static int8_t CDC_Init_FS(void)
{
  /* USER CODE BEGIN 3 */
  /* Set Application Buffers */
  USBD_CDC_SetTxBuffer(&hUsbDeviceFS, UserTxBufferFS, 0);
  USBD_CDC_SetRxBuffer(&hUsbDeviceFS, UserRxBufferFS);
  return (USBD_OK);
  /* USER CODE END 3 */
}

/**
  * @brief  DeInitializes the CDC media low layer
  * @retval USBD_OK if all operations are OK else USBD_FAIL
  */
static int8_t CDC_DeInit_FS(void)
{
  /* USER CODE BEGIN 4 */
  return (USBD_OK);
  /* USER CODE END 4 */
}

/**
  * @brief  Manage the CDC class requests
  * @param  cmd: Command code
  * @param  pbuf: Buffer containing command data (request parameters)
  * @param  length: Number of data to be sent (in bytes)
  * @retval Result of the operation: USBD_OK if all operations are OK else USBD_FAIL
  */
static int8_t CDC_Control_FS(uint8_t cmd, uint8_t* pbuf, uint16_t length)
{
  /* USER CODE BEGIN 5 */
  switch(cmd)
  {
    case CDC_SEND_ENCAPSULATED_COMMAND:

    break;

    case CDC_GET_ENCAPSULATED_RESPONSE:

    break;

    case CDC_SET_COMM_FEATURE:

    break;

    case CDC_GET_COMM_FEATURE:

    break;

    case CDC_CLEAR_COMM_FEATURE:

    break;

  /*******************************************************************************/
  /* Line Coding Structure                                                       */
  /*-----------------------------------------------------------------------------*/
  /* Offset | Field       | Size | Value  | Description                          */
  /* 0      | dwDTERate   |   4  | Number |Data terminal rate, in bits per second*/
  /* 4      | bCharFormat |   1  | Number | Stop bits                            */
  /*                                        0 - 1 Stop bit                       */
  /*                                        1 - 1.5 Stop bits                    */
  /*                                        2 - 2 Stop bits                      */
  /* 5      | bParityType |  1   | Number | Parity                               */
  /*                                        0 - None                             */
  /*                                        1 - Odd                              */
  /*                                        2 - Even                             */
  /*                                        3 - Mark                             */
  /*                                        4 - Space                            */
  /* 6      | bDataBits  |   1   | Number Data bits (5, 6, 7, 8 or 16).          */
  /*******************************************************************************/
    case CDC_SET_LINE_CODING:
      USBD_CDC_LineCoding.bitrate = (pbuf[3] << 24) | (pbuf[2] << 16) | (pbuf[1] << 8) | pbuf[0];
      USBD_CDC_LineCoding.format = pbuf[4];
      USBD_CDC_LineCoding.paritytype = pbuf[5];
      USBD_CDC_LineCoding.datatype = pbuf[6];
    break;

    case CDC_GET_LINE_CODING:
      pbuf[0] = (uint8_t)(USBD_CDC_LineCoding.bitrate);
      pbuf[1] = (uint8_t)(USBD_CDC_LineCoding.bitrate >> 8);
      pbuf[2] = (uint8_t)(USBD_CDC_LineCoding.bitrate >> 16);
      pbuf[3] = (uint8_t)(USBD_CDC_LineCoding.bitrate >> 24);
      pbuf[4] = USBD_CDC_LineCoding.format;
      pbuf[5] = USBD_CDC_LineCoding.paritytype;
      pbuf[6] = USBD_CDC_LineCoding.datatype;
    break;

    case CDC_SET_CONTROL_LINE_STATE:

    break;

    case CDC_SEND_BREAK:

    break;

  default:
    break;
  }

  return (USBD_OK);
  /* USER CODE END 5 */
}

/**
  * @brief  Data received over USB OUT endpoint are sent over CDC interface
  *         through this function.
  *
  *         @note
  *         This function will issue a NAK packet on any OUT packet received on
  *         USB endpoint until exiting this function. If you exit this function
  *         before transfer is complete on CDC interface (ie. using DMA controller)
  *         it will result in receiving more data while previous ones are still
  *         not sent.
  *
  * @param  Buf: Buffer of data to be received
  * @param  Len: Number of data received (in bytes)
  * @retval Result of the operation: USBD_OK if all operations are OK else USBD_FAIL
  */
static int8_t CDC_Receive_FS(uint8_t* Buf, uint32_t *Len)
{
  /* USER CODE BEGIN 6 */
  USBD_CDC_SetRxBuffer(&hUsbDeviceFS, &Buf[0]);
  USBD_CDC_ReceivePacket(&hUsbDeviceFS);
  
  u8UsbCdcReceiveLength[u8UsbCdcReceiveGrooveIndex] = *Len;
  memset(u8UsbCdcReceiveBuffer[u8UsbCdcReceiveGrooveIndex], 0, USB_CDC_RECEIVE_BUFFER_SIZE);
  memcpy(u8UsbCdcReceiveBuffer[u8UsbCdcReceiveGrooveIndex], Buf, *Len);
  USB_CDC_ReceiveDataDone(  );
  
  return (USBD_OK);
  /* USER CODE END 6 */
}

/**
  * @brief  CDC_Transmit_FS
  *         Data to send over USB IN endpoint are sent over CDC interface
  *         through this function.
  *         @note
  *
  *
  * @param  Buf: Buffer of data to be sent
  * @param  Len: Number of data to be sent (in bytes)
  * @retval USBD_OK if all operations are OK else USBD_FAIL or USBD_BUSY
  */
uint8_t CDC_Transmit_FS(uint8_t* Buf, uint16_t Len)
{
  uint8_t result = USBD_OK;
  /* USER CODE BEGIN 7 */
  USBD_CDC_HandleTypeDef *hcdc = (USBD_CDC_HandleTypeDef*)hUsbDeviceFS.pClassData;
  if (hcdc->TxState != 0){
    return USBD_BUSY;
  }
  USBD_CDC_SetTxBuffer(&hUsbDeviceFS, Buf, Len);
  result = USBD_CDC_TransmitPacket(&hUsbDeviceFS);
  /* USER CODE END 7 */
  return result;
}

/* USER CODE BEGIN PRIVATE_FUNCTIONS_IMPLEMENTATION */
const uint8_t *getUSB_CDC_ReceiveBuffer(void)
{
  return (const uint8_t *)u8UsbCdcReceiveBuffer[u8UsbCdcReadBufferIndex];
}

uint8_t getUSB_CDC_ReceiveLength(void)
{
  return (uint8_t)u8UsbCdcReceiveLength[u8UsbCdcReadBufferIndex];
}

void setUSB_CDC_TransmitBuffer(uint8_t *Buf)
{
  memcpy(u8UsbCdcTransmitBuffer[u8UsbCdcWriteBufferIndex], Buf, u8UsbCdcTransmitLength[u8UsbCdcWriteBufferIndex]);
}

void setUSB_CDC_TransmitLength(uint8_t Len)
{
  u8UsbCdcTransmitLength[u8UsbCdcWriteBufferIndex] = Len;
}

void USB_CDC_SendData(uint8_t *Buf, uint32_t Len)
{
  uint32_t i;
  uint8_t result;
  uint8_t *p=NULL;
  uint32_t packetNum, packetLength, lastpacketLength;
  
  packetNum = Len / CDC_DATA_FS_MAX_PACKET_SIZE;
  packetLength = CDC_DATA_FS_MAX_PACKET_SIZE;
  lastpacketLength = Len % CDC_DATA_FS_MAX_PACKET_SIZE;
  if (lastpacketLength)
    ++packetNum;
  
  for(i=1;i<=packetNum;++i)
  {
    if (i != packetNum)
      packetLength = CDC_DATA_FS_MAX_PACKET_SIZE;
    else
      packetLength = lastpacketLength;
    p = &Buf[(i-1)*CDC_DATA_FS_MAX_PACKET_SIZE];
    do{
      result = CDC_Transmit_FS(p, packetLength);
    }while(result != USBD_OK);//等待数据发送完毕再发送下一个包
  }
}

int USB_CDC_FillBuffer(char *fmt, ...)
{
  va_list args;
  char pstr[CDC_DATA_FS_MAX_PACKET_SIZE];
  int n;
  
  va_start(args,fmt);
  memset(pstr,0,CDC_DATA_FS_MAX_PACKET_SIZE);
  n=vsnprintf(pstr,CDC_DATA_FS_MAX_PACKET_SIZE,fmt,args);
  setUSB_CDC_TransmitLength(n);
  setUSB_CDC_TransmitBuffer((uint8_t*)pstr);
  USB_CDC_TrainsmitDataWriteDone(  );
  HAL_Delay(10);
  va_end(args);
  return n;
}

int USB_CDC_Ptintf(char *fmt, ...)
{
  va_list args;
  char pstr[CDC_DATA_FS_MAX_PACKET_SIZE];
  int n;
  
  va_start(args,fmt);
  memset(pstr,0,CDC_DATA_FS_MAX_PACKET_SIZE);
  n=vsnprintf(pstr,CDC_DATA_FS_MAX_PACKET_SIZE,fmt,args);
  CDC_Transmit_FS((uint8_t*)pstr,n);
  HAL_Delay(2);
  va_end(args);
  return n;
}
/* USER CODE END PRIVATE_FUNCTIONS_IMPLEMENTATION */

/**
  * @}
  */

/**
  * @}
  */
