#include "OPB9000.h"

#include <string.h>

#include "gpio.h"
#include "tim.h"

#include "usbd_cdc_if.h"

OPB9000_Com_Status_t      OPB9000_Com_Status;
OPB9000_Cmd_Status_t      OPB9000_Cmd_Status;
OPB9000_Bank_Bit_Data_t   OPB9000_Bank_Data_Value;
OPB9000_Out_Status_t      OPB9000_Out_Status;
OPB9000_Cal_Status_t      OPB9000_Cal_Status;
OPB9000_Cal_Result_t      OPB9000_Cal_Result;
uint32_t                  OPB9000_ReceiveTime;

static uint64_t OPB9000_Command_Data;
static uint32_t OPB9000_Command_Data2[2];
static uint32_t OPB9000_Command_Data3;

/* 发送/接收传输个数 */
static uint32_t           OPB9000_TransmitCount;
static uint32_t           OPB9000_TransmitLen;
static uint32_t           OPB9000_ReceiveCount;
static uint32_t           OPB9000_ReceiveLen;

static void OPB9000_Delay(void)//延时
{
  __nop();__nop();__nop();__nop();__nop();
  __nop();__nop();__nop();__nop();__nop();
}

static void __CAL_SET(OPB9000_Cal_Status_t status)//Pb8 引脚设置
{
  if (status == OPB9000_Cal_Status_Prepare)
    HAL_GPIO_WritePin(GPIOB, GPIO_PIN_8, GPIO_PIN_RESET);
  else if (status == OPB9000_Cal_Status_Turn)
    HAL_GPIO_WritePin(GPIOB, GPIO_PIN_8, GPIO_PIN_SET);
  else if (status == OPB9000_Cal_Status_Undef)
    HAL_GPIO_WritePin(GPIOB, GPIO_PIN_8, GPIO_PIN_SET);
}

static void __CAL_RESET(OPB9000_Cal_Status_t status)//Pb8 引脚设置
{
  if (status == OPB9000_Cal_Status_Prepare)
    HAL_GPIO_WritePin(GPIOB, GPIO_PIN_8, GPIO_PIN_SET);
  else if (status == OPB9000_Cal_Status_Turn)
    HAL_GPIO_WritePin(GPIOB, GPIO_PIN_8, GPIO_PIN_RESET);
  else if (status == OPB9000_Cal_Status_Undef)
    HAL_GPIO_WritePin(GPIOB, GPIO_PIN_8, GPIO_PIN_RESET);
}

static void __CAL_IN(void)//Pb8 引脚配置输入
{
  GPIOB_Pin8_In_Init();
}

static void __CAL_OUT(void)//Pb8 引脚配置输出
{
  GPIOB_Pin8_Out_Init();
}

static GPIO_PinState __CAL_GET(void)//Pb8 引脚读取
{
  return HAL_GPIO_ReadPin(GPIOB, GPIO_PIN_8);
}

static GPIO_PinState __OUT_GET(void)//Pb9 引脚读取
{
  return HAL_GPIO_ReadPin(GPIOB, GPIO_PIN_9);
}

uint64_t getOPB9000Data(void)//获取OPB9000_Command_Data	
{
  return OPB9000_Command_Data;
}

uint32_t getOPB9000Data2(uint8_t index)//获取OPB9000_Command_Data2[index]
{
  return OPB9000_Command_Data2[index];
}

void setOPB9000Bank2(uint16_t REF, uint16_t DS, uint16_t OP)
{
  OPB9000_Command_Data3 = 0;
  OPB9000_Command_Data3 |= ((REF << 8) & 0xF00);
  OPB9000_Command_Data3 |= (( DS << 7) & 0x080);
  OPB9000_Command_Data3 |= (( OP << 6) & 0x040);
}

void OPB9000_Read_Request(void)
{
  /* 1100-01 Read Request */
  /* 把CAL口置为输出 */
  __CAL_OUT();
  OPB9000_Cal_Status = OPB9000_Cal_Status_Undef;
  __CAL_RESET(OPB9000_Cal_Status);
  
  OPB9000_Command_Data = 0x0023;//  0x0010 0011   1100-01 读取请求 先发高位
  
  /* 准备发送数据 */
  OPB9000_Com_Status = OPB9000_Com_Status_TransmitWaiting;
  /* 读取请求指令 */
  OPB9000_Cmd_Status = OPB9000_Cmd_Status_ReadRequest;
  /* Cal脚进入准备态 */
  OPB9000_Cal_Status = OPB9000_Cal_Status_Prepare;
  OPB9000_TransmitCount = 0;
  OPB9000_TransmitLen = 6;
  OPB9000_ReceiveCount = 0;
  OPB9000_ReceiveLen = 37;
  TIM2_Enable();
}

void OPB9000_WriteBank2bits(void)
{
  /* 1100-10-bbbbbb Write Bank 2 bits */
  /* 把CAL口置为输出 */
  __CAL_OUT();
  OPB9000_Cal_Status = OPB9000_Cal_Status_Undef;
  __CAL_RESET(OPB9000_Cal_Status);
  
  OPB9000_Command_Data = 0x13;//0001 0011   1100-10-bbbbbb //写bank2
//  OPB9000_Command_Data = 0x8D3;
  OPB9000_Command_Data |= OPB9000_Command_Data3; //  1D3  0001 11  01  0011 
  
  /* 准备发送数据 */
  OPB9000_Com_Status = OPB9000_Com_Status_TransmitWaiting;
  /* Bank写入指令 */
  OPB9000_Cmd_Status = OPB9000_Cmd_Status_WriteBank2bits;
  /* Cal脚进入准备态 */
  OPB9000_Cal_Status = OPB9000_Cal_Status_Prepare;
  OPB9000_TransmitCount = 0;
  OPB9000_TransmitLen = 12;
  OPB9000_ReceiveCount = 0;
  OPB9000_ReceiveLen = 0;
  TIM2_Enable();
}
void OPB9000_Calibrate_Stand(void)
{
	//第一步  设置灵敏度为15 
	setOPB9000Bank2(15,0,0);
	OPB9000_WriteBank2bits();
	HAL_Delay(100);
	while(OPB9000_Com_Status != OPB9000_Com_Status_Done)
		;
//	USB_CDC_Ptintf("WriteBank2bits\r\n");
	USB_CDC_Ptintf("写灵敏度为15  极性反向\r\n");
	OPB9000_Read_Request();
	HAL_Delay(100);
	if (HAL_GPIO_ReadPin(GPIOB, GPIO_PIN_9) != GPIO_PIN_RESET)
		HAL_GPIO_WritePin(GPIOC, GPIO_PIN_13, GPIO_PIN_SET);
	else
		HAL_GPIO_WritePin(GPIOC, GPIO_PIN_13, GPIO_PIN_RESET);
	
	if(!HAL_GPIO_ReadPin(GPIOC, GPIO_PIN_13))
	{
		setOPB9000Bank2(15,0,1);
		OPB9000_WriteBank2bits();
		HAL_Delay(100);
		while(OPB9000_Com_Status != OPB9000_Com_Status_Done)
		;
//		USB_CDC_Ptintf("WriteBank2bits\r\n");
		USB_CDC_Ptintf("写灵敏度为15  极性正向\r\n");
		OPB9000_Read_Request();
		HAL_Delay(100);
		if (HAL_GPIO_ReadPin(GPIOB, GPIO_PIN_9) != GPIO_PIN_RESET)
			HAL_GPIO_WritePin(GPIOC, GPIO_PIN_13, GPIO_PIN_SET);
		else
			HAL_GPIO_WritePin(GPIOC, GPIO_PIN_13, GPIO_PIN_RESET);
	}
	
	if(HAL_GPIO_ReadPin(GPIOC, GPIO_PIN_13))
	{
		OPB9000_Read_Request();
		HAL_Delay(100);
		while(OPB9000_Com_Status != OPB9000_Com_Status_Done)
			;
		
		if(OPB9000_Bank_Data_Value.Bank2Data.Bank2Bit.REF != 15 || OPB9000_Bank_Data_Value.Bank2Data.Bank2Bit.DS != 0 || OPB9000_Bank_Data_Value.Bank3Data.Bank3Bit.ErrorFlag == 1)
		{
			//设置失败
			OPB9000_Com_Status = OPB9000_Com_Status_Idle;
			OPB9000_Cmd_Status = OPB9000_Cmd_Status_Idle;
//			USB_CDC_Ptintf("--Write Fail\r\n");
			USB_CDC_Ptintf("写灵敏度失败 请确认在蓝灯灭时读取\r\n");
			return;
		}
		else
			USB_CDC_Ptintf("写灵敏度成功\r\n");
//			USB_CDC_Ptintf("--Write Ok\r\n");
	}
//	else
//	{
//		OPB9000_Com_Status = OPB9000_Com_Status_Idle;
//		OPB9000_Cmd_Status = OPB9000_Cmd_Status_Idle;
//		USB_CDC_Ptintf("--Write Fail\r\n");
//		return;
//	}
	
	//第三步  校准
	OPB9000_Calibrate_Request();
	
	while(OPB9000_Cal_Result == OPB9000_Cal_Result_Calibrating)
		;
	if(OPB9000_Cal_Result == OPB9000_Cal_Result_Unsuccessful)
	{
		OPB9000_Com_Status = OPB9000_Com_Status_Idle;
		OPB9000_Cmd_Status = OPB9000_Cmd_Status_Idle;
//		USB_CDC_Ptintf("--Calibration Unsuccessful\r\n");
		USB_CDC_Ptintf("校准失败 检查试管位置是否过远\r\n");
		return;
	}
	else if(OPB9000_Cal_Result == OPB9000_Cal_Result_Successful)
//		USB_CDC_Ptintf("--Calibration Successful\r\n");
		USB_CDC_Ptintf("校准成功\r\n");
	
	HAL_Delay(100);
	//第四步  设置灵敏度为5
	setOPB9000Bank2(2,0,1);
	OPB9000_WriteBank2bits();
	HAL_Delay(100);
	while(OPB9000_Com_Status != OPB9000_Com_Status_Done)
		;
	USB_CDC_Ptintf("写灵敏度为2 极性正向\r\n");
//	USB_CDC_Ptintf("WriteBank2bits\r\n");
	OPB9000_Read_Request();
//	USB_CDC_Ptintf("read and move\r\n");
	USB_CDC_Ptintf("移开试管 蓝色指示灯灭掉后 发送读取指令RR\r\n");

	OPB9000_Com_Status = OPB9000_Com_Status_Idle;
	OPB9000_Cmd_Status = OPB9000_Cmd_Status_Idle;
}
void OPB9000_Calibrate_Request(void)
{
  /* 1100-11 Calibrate Request */
  /* 把CAL口置为输出 */
  __CAL_OUT();
	OPB9000_Cal_Result = OPB9000_Cal_Result_Calibrating;
  OPB9000_Cal_Status = OPB9000_Cal_Status_Undef;
  __CAL_RESET(OPB9000_Cal_Status);
  
  OPB9000_Command_Data = 0x0033;	//0x0011 0011  1100-11 校准请求
  
  /* 准备发送数据 */
  OPB9000_Com_Status = OPB9000_Com_Status_TransmitWaiting;
  /* 校准请求指令 */
  OPB9000_Cmd_Status = OPB9000_Cmd_Status_CalibrateRequest;
  /* Cal脚进入准备态 */
  OPB9000_Cal_Status = OPB9000_Cal_Status_Prepare;
  OPB9000_TransmitCount = 0;
  OPB9000_TransmitLen = 6;
  OPB9000_ReceiveCount = 0;
  OPB9000_ReceiveLen = 13;
  TIM2_Enable();
}

uint16_t ManchesterData[ManchesterDataSize];
void HAL_GPIO_EXTI_Callback(uint16_t GPIO_Pin)
{
  GPIO_PinState bitstatus;
  
  switch (GPIO_Pin)
  {
    case GPIO_PIN_8:
    {
      bitstatus = __CAL_GET();
      
      if (OPB9000_Com_Status == OPB9000_Com_Status_ReceiveWaiting)
      {
        OPB9000_Cal_Status = OPB9000_Status_Status_Respond;
      }
    }break;
    
    case GPIO_PIN_9:
    {
      bitstatus = __OUT_GET();
      
      if (OPB9000_Cmd_Status == OPB9000_Cmd_Status_Idle)
      {
        if (bitstatus == GPIO_PIN_SET)
          OPB9000_Out_Status = OPB9000_Out_Status_Set;
        else
          OPB9000_Out_Status = OPB9000_Out_Status_Reset;
      }
    }break;
  }
}

uint32_t reverseBitsInRange(uint32_t num, int x, int y)
{
  while (x < y)
  {
    // 获取第 x 和 y 位的值
    int bitX = (num >> x) & 1;
    int bitY = (num >> y) & 1;

    // 交换第 x 和 y 位的值
    if (bitX != bitY) 
    {
        num ^= (1 << x);
        num ^= (1 << y);
    }
    
    x++;
    y--;
  }
  
  return num;
}
extern uint8_t bIslevelShow;
#define MaxGap 20
void findData(uint16_t *Data)
{
	int i,j,k,l;
	int nextJumpBit = 0;
	int JumpGap = 0;
	int cntSameBit = 0;
	uint8_t st_findData = 0;
	uint8_t bitMove = 0;
	for(i = 1;i < 100;i++)
	{
		//找到第一个跳变位置
		if(Data[i] != Data[i - 1])
		{
			for(j = 1;j < 100;j++)
			{
				//找到第二个跳变位置 从这个位置开始计算 
				if(Data[j+i] != Data[j+i - 1])
				{
					nextJumpBit = j+i;
					JumpGap = j;
					break;
				}
			}
			break;
		}
	}
	for(k = 0;k < 37 * 2;k++)
	{
		for(l = 1;l < 100;l++)
		{
			if(st_findData == 0)//第一步 找到这次跳变的值
			{
				if(Data[l+nextJumpBit] != Data[l+nextJumpBit + 1])
				{
					if (Data[l+nextJumpBit + 1] == 0)
					{
						OPB9000_Command_Data &= ~(1<<bitMove);
						Data[l+nextJumpBit + 1] = 2;
					}
					else
					{
						OPB9000_Command_Data |= (1<<bitMove);
						Data[l+nextJumpBit + 1] = 3;
					}
					bitMove++;
					nextJumpBit = l+nextJumpBit + 1;
					st_findData = 1;
					break;
				}
			}
			else if(st_findData == 1)//第二步 找到下次跳变的开始
			{
				if(Data[l+nextJumpBit] == Data[l+nextJumpBit + 1])
				{
					cntSameBit++;
					if(cntSameBit >= MaxGap)
					{
						nextJumpBit = nextJumpBit + JumpGap;
						cntSameBit = 0;
						st_findData = 0;
						break;
					}
					else
					{
							if(Data[l+nextJumpBit] != Data[l+nextJumpBit + 1])
							{
								nextJumpBit = l+nextJumpBit + 1;
								cntSameBit = 0;
								st_findData = 0;
								break;
							}
					}
				}
				
			}
		
		}
	}
	
	if((OPB9000_Command_Data & 0xFFFFFF) == 0)
	{
		bIslevelShow = 1;
	}
}

void HAL_TIM_PeriodElapsedCallback(TIM_HandleTypeDef *htim)
{
  static uint32_t tick = 0;
  GPIO_PinState bitstatus;
 
	 switch (OPB9000_Cmd_Status)
	{
		case OPB9000_Cmd_Status_Idle:
		{
			TIM2_Disable();
		}break;
			
		case OPB9000_Cmd_Status_Reserved:
		{
			TIM2_Disable();
		}break;
		
		case OPB9000_Cmd_Status_ReadRequest:
		{
			/* Idle */
			if (OPB9000_Com_Status == OPB9000_Com_Status_Idle)
			{
				TIM2_Disable();
			}
			/* TransmitWaiting */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_TransmitWaiting)
			{
				OPB9000_Com_Status = OPB9000_Com_Status_Transmitting;
			}
			/* Transmitting */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_Transmitting)
			{
				if (OPB9000_Cal_Status == OPB9000_Cal_Status_Prepare)
				{
					if (OPB9000_Command_Data & (1<<OPB9000_TransmitCount))
						__CAL_SET(OPB9000_Cal_Status);
					else
						__CAL_RESET(OPB9000_Cal_Status);
					OPB9000_Cal_Status = OPB9000_Cal_Status_Turn;
				}
				else if (OPB9000_Cal_Status == OPB9000_Cal_Status_Turn)
				{
					if (OPB9000_Command_Data & (1<<OPB9000_TransmitCount))
						__CAL_SET(OPB9000_Cal_Status);
					else
						__CAL_RESET(OPB9000_Cal_Status);
					OPB9000_Cal_Status = OPB9000_Cal_Status_Prepare;
					++OPB9000_TransmitCount;
				}
				if (OPB9000_TransmitCount == OPB9000_TransmitLen)
				{
//          memset(ManchesterData, 0, sizeof(ManchesterData));
					__disable_irq();
					for (int i = 0; i<ManchesterDataSize; ++i)
						ManchesterData[i] = (GPIOB->IDR & GPIO_PIN_9) ? 1: 0;
					__enable_irq();
					OPB9000_Com_Status = OPB9000_Com_Status_TransmitDone;
					OPB9000_Cal_Status = OPB9000_Cal_Status_Undef;
				}
			}
			/* TransmitDone */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_TransmitDone)
			{
				__CAL_SET(OPB9000_Cal_Status);
				__CAL_IN();
				OPB9000_Command_Data = 0U;
				OPB9000_Com_Status = OPB9000_Com_Status_ReceiveWaiting;
			}
			/* ReceiveWaiting */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_ReceiveWaiting)
			{
//        if (OPB9000_Cal_Status == OPB9000_Status_Status_Respond)
				{
					OPB9000_Com_Status = OPB9000_Com_Status_Receiving;
					OPB9000_Cal_Status = OPB9000_Status_Status_Prepare;
				}
			}
			/* Receiving */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_Receiving)
			{
				findData(ManchesterData);
				
				OPB9000_Com_Status = OPB9000_Com_Status_ReceiveDone;
				OPB9000_Cal_Status = OPB9000_Cal_Status_Undef;
				
			}
			/* ReceiveDone */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_ReceiveDone)
			{
				__CAL_OUT();
				__CAL_SET(OPB9000_Cal_Status);
				uint32_t tempOPB9000_Command_Data;
				tempOPB9000_Command_Data = OPB9000_Command_Data;
				tempOPB9000_Command_Data = reverseBitsInRange(OPB9000_Command_Data, 13-1, 14-1);
				tempOPB9000_Command_Data = reverseBitsInRange(OPB9000_Command_Data, 15-1, 24-1);
				tempOPB9000_Command_Data = reverseBitsInRange(OPB9000_Command_Data, 25-1, 28-1);
				OPB9000_Bank_Data_Value.Bank3Data.Data = ((tempOPB9000_Command_Data & 0x000007FF) >>  0);
				OPB9000_Bank_Data_Value.Bank1Data.Data = ((tempOPB9000_Command_Data & 0x00FFF800) >> 11);
				OPB9000_Bank_Data_Value.Bank2Data.Data = ((tempOPB9000_Command_Data & 0xFF000000) >> 24);
//        OPB9000_Command_Data = 0;
				OPB9000_Com_Status = OPB9000_Com_Status_Done;
				if (bitstatus == GPIO_PIN_SET)
					OPB9000_Out_Status = OPB9000_Out_Status_Set;
				else
					OPB9000_Out_Status = OPB9000_Out_Status_Reset;
				TIM2_Disable();
			}
			/* Done */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_Done)
			{
				
			}
		}break;
		
		case OPB9000_Cmd_Status_WriteBank2bits:
		{
			/* Idle */
			if (OPB9000_Com_Status == OPB9000_Com_Status_Idle)
			{
				TIM2_Disable();
			}
			/* TransmitWaiting */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_TransmitWaiting)
			{
				OPB9000_Com_Status = OPB9000_Com_Status_Transmitting;
			}
			/* Transmitting */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_Transmitting)
			{
				if (OPB9000_Cal_Status == OPB9000_Cal_Status_Prepare)
				{
					if (OPB9000_Command_Data & (1<<OPB9000_TransmitCount))
						__CAL_SET(OPB9000_Cal_Status);//DI
					else
						__CAL_RESET(OPB9000_Cal_Status);//GAO
					OPB9000_Cal_Status = OPB9000_Cal_Status_Turn;
				}
				else if (OPB9000_Cal_Status == OPB9000_Cal_Status_Turn)
				{
					if (OPB9000_Command_Data & (1<<OPB9000_TransmitCount))
						__CAL_SET(OPB9000_Cal_Status);//GAO
					else
						__CAL_RESET(OPB9000_Cal_Status);//DI
					OPB9000_Cal_Status = OPB9000_Cal_Status_Prepare;
					++OPB9000_TransmitCount;
				}
				if (OPB9000_TransmitCount == OPB9000_TransmitLen)
				{
					OPB9000_Com_Status = OPB9000_Com_Status_TransmitDone;
					OPB9000_Cal_Status = OPB9000_Cal_Status_Undef;
				}
			}
			/* TransmitDone */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_TransmitDone)
			{
				__CAL_SET(OPB9000_Cal_Status);
				__CAL_IN();
				OPB9000_Command_Data = 0U;
				OPB9000_Com_Status = OPB9000_Com_Status_ReceiveWaiting;
			}
			/* ReceiveWaiting */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_ReceiveWaiting)
			{
				OPB9000_Com_Status = OPB9000_Com_Status_Done;
				TIM2_Disable();
			}
			/* Done */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_Done)
			{
				
			}
		}break;
		
		case OPB9000_Cmd_Status_CalibrateRequest:
		{
			/* Idle */
			if (OPB9000_Com_Status == OPB9000_Com_Status_Idle)
			{
				TIM2_Disable();
			}
			/* TransmitWaiting */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_TransmitWaiting)
			{
				OPB9000_Com_Status = OPB9000_Com_Status_Transmitting;
			}
			/* Transmitting */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_Transmitting)
			{
				if (OPB9000_Cal_Status == OPB9000_Cal_Status_Prepare)
				{
					if (OPB9000_Command_Data & (1<<OPB9000_TransmitCount))
						__CAL_SET(OPB9000_Cal_Status);
					else
						__CAL_RESET(OPB9000_Cal_Status);
					OPB9000_Cal_Status = OPB9000_Cal_Status_Turn;
				}
				else if (OPB9000_Cal_Status == OPB9000_Cal_Status_Turn)
				{
					if (OPB9000_Command_Data & (1<<OPB9000_TransmitCount))
						__CAL_SET(OPB9000_Cal_Status);
					else
						__CAL_RESET(OPB9000_Cal_Status);
					OPB9000_Cal_Status = OPB9000_Cal_Status_Prepare;
					++OPB9000_TransmitCount;
				}
				if (OPB9000_TransmitCount == OPB9000_TransmitLen)
				{
					OPB9000_Com_Status = OPB9000_Com_Status_TransmitDone;
					OPB9000_Cal_Status = OPB9000_Cal_Status_Undef;
				}
			}
			/* TransmitDone */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_TransmitDone)
			{
				__CAL_SET(OPB9000_Cal_Status);
				__CAL_IN();
				OPB9000_Command_Data = 0U;
				OPB9000_Com_Status = OPB9000_Com_Status_ReceiveWaiting;
				tick = HAL_GetTick();
			}
			/* ReceiveWaiting */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_ReceiveWaiting)
			{
				if (OPB9000_Cal_Status == OPB9000_Status_Status_Respond)
				{
					OPB9000_Com_Status = OPB9000_Com_Status_Done;
					OPB9000_Cal_Result = OPB9000_Cal_Result_Successful;
					TIM2_Disable();
				}
				if (HAL_GetTick() - tick > 100)
				{
					OPB9000_Com_Status = OPB9000_Com_Status_Done;
					OPB9000_Cal_Result = OPB9000_Cal_Result_Unsuccessful;
					TIM2_Disable();
				}
			}
			/* Done */
			else if (OPB9000_Com_Status == OPB9000_Com_Status_Done)
			{
				
			}
		}break;
		
		default:
		{
			TIM2_Disable();
		}break;
	}



}
