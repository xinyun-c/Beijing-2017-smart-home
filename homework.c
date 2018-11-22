#include<reg52.h>
#include<intrins.h>
typedef unsigned int uint;
typedef unsigned char uchar;
uint indi;
uchar trans;
uint status;
uint i;

void delay(uint i)
{
	while(i--);
}
void defa()
{
	TMOD=0x20;  
	TH1=0xfd;
	TL1=0xfd;
	SCON=0x50;
	PCON &=0x7f;
	TR1=1;
	EA=1;
	ES=1;
}
void main()
{
	defa();
	while(indi==1)
	{
	 P0=trans;
	}
}

void exchange() interrupt 4
{
	 uchar agent;
	 P1=0xff;	 
	 delay(200);	 
	 P1=~P1;
	 delay(200);
	 P1&=P1;
	 delay(200);	 
	 indi=1;
	 agent=SBUF;
	 RI=0;
	 P0=agent;
	 trans=agent;
}