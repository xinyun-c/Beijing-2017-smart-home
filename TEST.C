#include <reg52.h>
unsigned int flag;
void init()
{
	 TMOD=0x20;
     TH1=0xfd;
     TL1=0xfd;
	 SCON = 0x50;			// 设定串行口工作方式
	 PCON = PCON & 0x7f;			// 波特率不倍增
	 TR1=1;
     EA=1;
     ES=1;
}

main()
{
	init();
	
	while(1)
	{
    init();

		if(flag==1)
		  { 
			P0 = 0x00;
		  }
	}
}

void read_serial() interrupt 4 
{ 
  if(RI==1) 
   {  
	   RI = 0;
	   flag = 1;
   }
}