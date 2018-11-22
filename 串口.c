#include <reg52.h>
typedef unsigned int uint;
typedef unsigned char uchar;
void UsartInit()
{
  SCON=0X50;
  TMOD=0X20;
  PCON=0X80;
  TH1=0XF3;
  TL1=0XF3;
  ES=1;
  EA=1;
  TR1=1;
  P0=0X00;
}

void main()
{
  UsartInit();
  while(1);
}

void Usart() interrupt 4
{
  uchar receiveData;
  receiveData=SBUF;
  P0=receiveData;
  RI=0;
  SBUF=receiveData;
  while(!TI);
  TI=0;
}