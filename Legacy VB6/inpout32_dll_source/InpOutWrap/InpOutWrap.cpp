
#include "stdafx.h"
#include "windows.h"
#include "stdio.h"

typedef void	(__stdcall *lpOut32)(short, short);
typedef short	(__stdcall *lpInp32)(short);
typedef BOOL	(__stdcall *lpIsInpOutDriverOpen)(void);
typedef BOOL	(__stdcall *lpIsXP64Bit)(void);


#define DllExport __declspec( dllexport )

extern short Inp(const char *Port);

extern void Out(const char *Port, const char *Info);


//Some global function pointers (messy but fine for an example)
lpOut32 gfpOut32;
lpInp32 gfpInp32;
lpIsInpOutDriverOpen gfpIsInpOutDriverOpen;
lpIsXP64Bit gfpIsXP64Bit;


extern short Inp(const char *Port)
{
	//Dynamically load the DLL at runtime (not linked at compile time)
	HINSTANCE hInpOutDll ;
	hInpOutDll = LoadLibrary ( "InpOut32.DLL" ) ;	//The 32bit DLL. If we are building x64 C++ 
													//applicaiton then use InpOutx64.dll
	if ( hInpOutDll != NULL )
	{
		gfpOut32 = (lpOut32)GetProcAddress(hInpOutDll, "Out32");
		gfpInp32 = (lpInp32)GetProcAddress(hInpOutDll, "Inp32");
		gfpIsInpOutDriverOpen = (lpIsInpOutDriverOpen)GetProcAddress(hInpOutDll, "IsInpOutDriverOpen");
		gfpIsXP64Bit = (lpIsXP64Bit)GetProcAddress(hInpOutDll, "IsXP64Bit");

		WORD wData = 0;
		if (gfpIsInpOutDriverOpen())
		{
			short iPort = atoi(Port);
			wData = gfpInp32(iPort);	//Read the port
		}

		FreeLibrary ( hInpOutDll ) ;
		return wData;
	}
}

extern void Out(const char *Port, const char *Info)
{	
	//Dynamically load the DLL at runtime (not linked at compile time)
	HINSTANCE hInpOutDll ;
	hInpOutDll = LoadLibrary ( "InpOut32.DLL" ) ;	//The 32bit DLL. If we are building x64 C++ 
													//applicaiton then use InpOutx64.dll
	if ( hInpOutDll != NULL )
	{
		gfpOut32 = (lpOut32)GetProcAddress(hInpOutDll, "Out32");
		gfpInp32 = (lpInp32)GetProcAddress(hInpOutDll, "Inp32");
		gfpIsInpOutDriverOpen = (lpIsInpOutDriverOpen)GetProcAddress(hInpOutDll, "IsInpOutDriverOpen");
		gfpIsXP64Bit = (lpIsXP64Bit)GetProcAddress(hInpOutDll, "IsXP64Bit");

		WORD wData = 0;
		if (gfpIsInpOutDriverOpen())
		{

			short iPort = atoi(Port);
			wData = atoi(Info);
			gfpOut32(iPort, wData);

		}
		//All done
		FreeLibrary ( hInpOutDll ) ;
	}
}


