/* USER CODE BEGIN Header */
/**
  ******************************************************************************
  * @file           : main.c
  * @brief          : Main program body
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
#include "main.h"
#include "tim.h"
#include "usb_device.h"
#include "gpio.h"

/* Private includes ----------------------------------------------------------*/
/* USER CODE BEGIN Includes */
#include "usbd_cdc_if.h"
#include "OPB9000.h"
/* USER CODE END Includes */

/* Private typedef -----------------------------------------------------------*/
/* USER CODE BEGIN PTD */

/* USER CODE END PTD */

/* Private define ------------------------------------------------------------*/
/* USER CODE BEGIN PD */

/* USER CODE END PD */

/* Private macro -------------------------------------------------------------*/
/* USER CODE BEGIN PM */

/* USER CODE END PM */

/* Private variables ---------------------------------------------------------*/

/* USER CODE BEGIN PV */

/* USER CODE END PV */

/* Private function prototypes -----------------------------------------------*/
void SystemClock_Config(void);
/* USER CODE BEGIN PFP */

/* USER CODE END PFP */

/* Private user code ---------------------------------------------------------*/
/* USER CODE BEGIN 0 */

/* USER CODE END 0 */

/**
  * @brief  The application entry point.
  * @retval int
  */
uint8_t bIslevelShow = 0;
int main(void)
{

  /* USER CODE BEGIN 1 */
  uint32_t tick = 0;
  
  /* USER CODE END 1 */

  /* MCU Configuration--------------------------------------------------------*/

  /* Reset of all peripherals, Initializes the Flash interface and the Systick. */
  HAL_Init();

  /* USER CODE BEGIN Init */

  /* USER CODE END Init */

  /* Configure the system clock*/
  SystemClock_Config();
//	SystemCoreClockUpdate();
  /* USER CODE BEGIN SysInit */

  /* USER CODE END SysInit */

  /* Initialize all configured peripherals */
  MX_GPIO_Init();
  MX_USB_DEVICE_Init();
  MX_TIM2_Init();
//	MX_TIM3_Init();
  /* USER CODE BEGIN 2 */
  
  HAL_GPIO_TogglePin(GPIOC, GPIO_PIN_13);
  HAL_Delay(250);
  HAL_GPIO_TogglePin(GPIOC, GPIO_PIN_13);
  HAL_Delay(250);
  HAL_GPIO_TogglePin(GPIOC, GPIO_PIN_13);
  HAL_Delay(250);
  HAL_GPIO_TogglePin(GPIOC, GPIO_PIN_13);
  HAL_Delay(250);
  
  /* USER CODE END 2 */

  /* Infinite loop */
  /* USER CODE BEGIN WHILE */
  while (1)
  {
    /* USER CODE END WHILE */

    /* USER CODE BEGIN 3 */
    
    if (OPB9000_Com_Status == OPB9000_Com_Status_Done)
    {
      if (OPB9000_Cmd_Status == OPB9000_Cmd_Status_ReadRequest)
      {
        USB_CDC_Ptintf("ReadRequest\r\n");
        USB_CDC_Ptintf("--Bank1   CA:%4d  AGC:%4d  LED:%4d\r\n", \
        OPB9000_Bank_Data_Value.Bank1Data.Bank1Bit.CA, \
        OPB9000_Bank_Data_Value.Bank1Data.Bank1Bit.AGC, \
        OPB9000_Bank_Data_Value.Bank1Data.Bank1Bit.LED \
        );
        USB_CDC_Ptintf("--Bank2  REF:%4d   DS:%4d   OP:%4d\r\n", \
        OPB9000_Bank_Data_Value.Bank2Data.Bank2Bit.REF, \
        OPB9000_Bank_Data_Value.Bank2Data.Bank2Bit.DS, \
        OPB9000_Bank_Data_Value.Bank2Data.Bank2Bit.OP \
        );
        USB_CDC_Ptintf("--Bank3   SB:%4d   EF:%4d\r\n", \
        OPB9000_Bank_Data_Value.Bank3Data.Bank3Bit.StartBit, \
        OPB9000_Bank_Data_Value.Bank3Data.Bank3Bit.ErrorFlag \
        );
				if(OPB9000_Bank_Data_Value.Bank3Data.Bank3Bit.ErrorFlag)
				{
					 USB_CDC_Ptintf("读取失败，请移开试剂 使蓝灯熄灭后再读\r\n");					
				}
				else
				{
					USB_CDC_Ptintf("读取成功，灵敏度为：%4d    ", OPB9000_Bank_Data_Value.Bank2Data.Bank2Bit.REF);
					if(OPB9000_Bank_Data_Value.Bank2Data.Bank2Bit.OP)
						USB_CDC_Ptintf("极性正向\r\n");			
					else
						USB_CDC_Ptintf("极性反向\r\n");		
				}
				if(bIslevelShow == 1)
				{
					for (int i = 0; i<ManchesterDataSize; ++i)
					{
						if (ManchesterData[i] == 0)
							USB_CDC_Ptintf("_");
						else if (ManchesterData[i] == 1)
							USB_CDC_Ptintf("-");
						else if (ManchesterData[i] == 2)
							USB_CDC_Ptintf("0");
						else if (ManchesterData[i] == 3)
							USB_CDC_Ptintf("1");
					}
				}
/*//        USB_CDC_Ptintf("\r\n");
//        uint64_t tempData;
//        tempData= getOPB9000Data();
//        for (int i = 0; i<37; ++i)
//        {
//          if (tempData & (1<<i))
//            USB_CDC_Ptintf("1");
//          else
//            USB_CDC_Ptintf("0");
//          if (i == 10 || i == 23)
//            USB_CDC_Ptintf("\r\n");
//        }
//        uint32_t tempData;
//        tempData= getOPB9000Data2(0);
//        for (int i = 0; i<32; ++i)
//        {
//          if (tempData & (1<<i))
//            USB_CDC_Ptintf("1");
//          else
//            USB_CDC_Ptintf("0");
//        }
//        tempData= getOPB9000Data2(1);
//        for (int i = 32; i<37; ++i)
//        {
//          if (tempData & (1<<(i-32)))
//            USB_CDC_Ptintf("1");
//          else
//            USB_CDC_Ptintf("0");
//        }
//        USB_CDC_Ptintf("\r\n");*/
      }
      else if (OPB9000_Cmd_Status == OPB9000_Cmd_Status_WriteBank2bits)
      {
        USB_CDC_Ptintf("WriteBank2bits\r\n");
        USB_CDC_Ptintf("--Complete\r\n");
        OPB9000_Read_Request();
      }
      else if (OPB9000_Cmd_Status == OPB9000_Cmd_Status_CalibrateRequest)
      {
        USB_CDC_Ptintf("CalibrateRequest\r\n");
        if (OPB9000_Cal_Result == OPB9000_Cal_Result_Successful)
          USB_CDC_Ptintf("--Successful\r\n");
        else
          USB_CDC_Ptintf("--Unsuccessful\r\n");
        OPB9000_Read_Request();
      }
      OPB9000_Com_Status = OPB9000_Com_Status_Idle;
      OPB9000_Cmd_Status = OPB9000_Cmd_Status_Idle;
    }
    
		/* USB 处理 */
		if (u8UsbCdcReceiveMessageNumber)
		{
			const char *string = NULL;
			uint8_t len;
			char CMD[5];
			uint32_t p1,p2,p3;
			int parameterNum;
			
			/* 处理消息 */
			string = (const char *)getUSB_CDC_ReceiveBuffer();
			len = getUSB_CDC_ReceiveLength();
			USB_CDC_Ptintf("receive data(%2d) -- %s\r\n", len, string);
			
			/*
			协议格式 
			## / CMD / p1 / p2 / p3 / ##
			*/
			parameterNum = sscanf(string, "## / %s / %u / %u / %u / ##", CMD,&p1,&p2,&p3);
			
			if (parameterNum != 4)
			{
				/* 未知命令 --help */
				USB_CDC_Ptintf("********************  help  ********************\r\n");
				USB_CDC_Ptintf("*   协议格式                                   *\r\n");
				USB_CDC_Ptintf("*   ## / CMD / P1 / P2 / P3 / ##               *\r\n");
				USB_CDC_Ptintf("*    --CMD                                     *\r\n");
				USB_CDC_Ptintf("*      -- RR (ReadRequest)                     *\r\n");
				USB_CDC_Ptintf("*        -- P1 (no use)                        *\r\n");
				USB_CDC_Ptintf("*        -- P2 (no use)                        *\r\n");
				USB_CDC_Ptintf("*        -- P3 (no use)                        *\r\n");
				USB_CDC_Ptintf("*      -- WB (WriteBank2bits)                  *\r\n");
				USB_CDC_Ptintf("*        -- P1 (REF)                           *\r\n");
				USB_CDC_Ptintf("*        -- P2 (DS)                            *\r\n");
				USB_CDC_Ptintf("*        -- P3 (OP)                            *\r\n");
				USB_CDC_Ptintf("*      -- CR (CalibrateRequest)                *\r\n");
				USB_CDC_Ptintf("*        -- P1 (no use)                        *\r\n");
				USB_CDC_Ptintf("*        -- P2 (no use)                        *\r\n");
				USB_CDC_Ptintf("*        -- P3 (no use)                        *\r\n");
				USB_CDC_Ptintf("*      -- CS (CalibrateStand)                  *\r\n");
				USB_CDC_Ptintf("*        -- P1 (no use)                        *\r\n");
				USB_CDC_Ptintf("*        -- P2 (no use)                        *\r\n");
				USB_CDC_Ptintf("*        -- P3 (no use)                        *\r\n");
				USB_CDC_Ptintf("*      -- LS (LevelShow)                       *\r\n");
				USB_CDC_Ptintf("*        -- P1 (0 off 1 on)                    *\r\n");
				USB_CDC_Ptintf("*        -- P2 (no use)                        *\r\n");
				USB_CDC_Ptintf("*        -- P3 (no use)                        *\r\n");
				USB_CDC_Ptintf("********************  help  ********************\r\n");
			}
			else
			{
				/* 读取Bank请求 */
				if (strcmp("RR",CMD) == 0)
				{
					OPB9000_Read_Request();
				}
				/* 写入Bank2 */
				else if (strcmp("WB",CMD) == 0)
				{
					setOPB9000Bank2(p1,p2,p3);
					OPB9000_WriteBank2bits();
				}
				/* 校准请求 */
				else if (strcmp("CR",CMD) == 0)
				{
					OPB9000_Calibrate_Request();
				}
				else if(strcmp("LS",CMD) == 0)
				{
					if(p1 == 0)
					{
						bIslevelShow = 0;
						USB_CDC_Ptintf("LevelShow off\r\n");
					}
					else if(p1 == 1)
					{
						bIslevelShow = 1;
						USB_CDC_Ptintf("LevelShow open\r\n");
					}
					
				}
				else if(strcmp("CS",CMD) == 0)
				{
					OPB9000_Calibrate_Stand();
				}
			}
			USB_CDC_ReceiveDataReadDone(  );
		}
		if (u8UsbCdcTransmitMessageNumber)
		{
			USB_CDC_SendData(u8UsbCdcTransmitBuffer[u8UsbCdcTransmitGrooveIndex],u8UsbCdcTransmitLength[u8UsbCdcTransmitGrooveIndex]);
			USB_CDC_TrainsmitDataDone(  );
		}
		
		/* LED输出 */
		if (OPB9000_Cmd_Status == OPB9000_Cmd_Status_Idle)
		{
			if (HAL_GPIO_ReadPin(GPIOB, GPIO_PIN_9) != GPIO_PIN_RESET)
				HAL_GPIO_WritePin(GPIOC, GPIO_PIN_13, GPIO_PIN_SET);
			else
				HAL_GPIO_WritePin(GPIOC, GPIO_PIN_13, GPIO_PIN_RESET);
		}
  }
  /* USER CODE END 3 */
}

/**
  * @brief System Clock Configuration
  * @retval None
  */
void SystemClock_Config(void)
{
  RCC_OscInitTypeDef RCC_OscInitStruct = {0};
  RCC_ClkInitTypeDef RCC_ClkInitStruct = {0};
  RCC_PeriphCLKInitTypeDef PeriphClkInit = {0};

  /** Initializes the RCC Oscillators according to the specified parameters
  * in the RCC_OscInitTypeDef structure.
  */
  RCC_OscInitStruct.OscillatorType = RCC_OSCILLATORTYPE_HSE;
  RCC_OscInitStruct.HSEState = RCC_HSE_ON;
  RCC_OscInitStruct.HSEPredivValue = RCC_HSE_PREDIV_DIV1;
  RCC_OscInitStruct.HSIState = RCC_HSI_ON;
  RCC_OscInitStruct.PLL.PLLState = RCC_PLL_ON;
  RCC_OscInitStruct.PLL.PLLSource = RCC_PLLSOURCE_HSE;
  RCC_OscInitStruct.PLL.PLLMUL = RCC_PLL_MUL9;
  if (HAL_RCC_OscConfig(&RCC_OscInitStruct) != HAL_OK)
  {
    Error_Handler();
  }

  /** Initializes the CPU, AHB and APB buses clocks
  */
  RCC_ClkInitStruct.ClockType = RCC_CLOCKTYPE_HCLK|RCC_CLOCKTYPE_SYSCLK
                              |RCC_CLOCKTYPE_PCLK1|RCC_CLOCKTYPE_PCLK2;
  RCC_ClkInitStruct.SYSCLKSource = RCC_SYSCLKSOURCE_PLLCLK;
  RCC_ClkInitStruct.AHBCLKDivider = RCC_SYSCLK_DIV1;
  RCC_ClkInitStruct.APB1CLKDivider = RCC_HCLK_DIV2;
  RCC_ClkInitStruct.APB2CLKDivider = RCC_HCLK_DIV1;

  if (HAL_RCC_ClockConfig(&RCC_ClkInitStruct, FLASH_LATENCY_2) != HAL_OK)
  {
    Error_Handler();
  }
  PeriphClkInit.PeriphClockSelection = RCC_PERIPHCLK_USB;
  PeriphClkInit.UsbClockSelection = RCC_USBCLKSOURCE_PLL_DIV1_5;
  if (HAL_RCCEx_PeriphCLKConfig(&PeriphClkInit) != HAL_OK)
  {
    Error_Handler();
  }
}

/* USER CODE BEGIN 4 */

/* USER CODE END 4 */

/**
  * @brief  This function is executed in case of error occurrence.
  * @retval None
  */
void Error_Handler(void)
{
  /* USER CODE BEGIN Error_Handler_Debug */
  /* User can add his own implementation to report the HAL error return state */
  __disable_irq();
  while (1)
  {
  }
  /* USER CODE END Error_Handler_Debug */
}

#ifdef  USE_FULL_ASSERT
/**
  * @brief  Reports the name of the source file and the source line number
  *         where the assert_param error has occurred.
  * @param  file: pointer to the source file name
  * @param  line: assert_param error line source number
  * @retval None
  */
void assert_failed(uint8_t *file, uint32_t line)
{
  /* USER CODE BEGIN 6 */
  /* User can add his own implementation to report the file name and line number,
     ex: printf("Wrong parameters value: file %s on line %d\r\n", file, line) */
  /* USER CODE END 6 */
}
#endif /* USE_FULL_ASSERT */
