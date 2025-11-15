/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sat Oct 11 11:10:00 2025
 */
/* Compiler settings for C:\Documents and Settings\Nickels\Desktop\CMDexe\dll.source\Redirect.idl:
    Os (OptLev=s), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

const IID IID_IApplication = {0x983145CD,0xC8DC,0x11D3,{0x9D,0xE4,0x40,0x00,0x0A,0x4A,0x15,0x41}};


const IID LIBID_RedirectLib = {0x983145C1,0xC8DC,0x11D3,{0x9D,0xE4,0x40,0x00,0x0A,0x4A,0x15,0x41}};


const IID DIID__IApplicationEvents = {0x983145CF,0xC8DC,0x11D3,{0x9D,0xE4,0x40,0x00,0x0A,0x4A,0x15,0x41}};


const CLSID CLSID_Application = {0x983145CE,0xC8DC,0x11D3,{0x9D,0xE4,0x40,0x00,0x0A,0x4A,0x15,0x41}};


#ifdef __cplusplus
}
#endif

