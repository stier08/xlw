/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Feb 20 18:13:23 2003
 */
/* Compiler settings for D:\Projects\xlw\xllExample\xllExample.idl:
    Os (OptLev=s), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __xllExample_h__
#define __xllExample_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

/* interface __MIDL_itf_xllExample_0000 */
/* [local] */ 

#ifndef INC_xldata32_H
#define INC_xldata32_H
typedef struct  xlref
    {
    WORD rwFirst;
    WORD rwLast;
    BYTE colFirst;
    BYTE colLast;
    }	XLREF;

typedef struct xlref __RPC_FAR *LPXLREF;

typedef struct  xlmref
    {
    WORD count;
    XLREF reftbl[ 1 ];
    }	XLMREF;

typedef struct xlmref __RPC_FAR *LPXLMREF;

typedef struct  xloper
    {
    union 
        {
        double num;
        LPSTR str;
        WORD bool_;
        WORD err;
        short w;
        struct  
            {
            WORD count;
            XLREF ref;
            }	sref;
        struct  
            {
            XLMREF __RPC_FAR *lpmref;
            DWORD idSheet;
            }	mref;
        struct  
            {
            struct xloper __RPC_FAR *lparray;
            WORD rows;
            WORD columns;
            }	array;
        struct  
            {
            union 
                {
                short level;
                short tbctrl;
                DWORD idSheet;
                }	valflow;
            WORD rw;
            BYTE col;
            BYTE xlflow;
            }	flow;
        struct  
            {
            union 
                {
                BYTE __RPC_FAR *lpbData;
                HANDLE hdata;
                }	h;
            long cbData;
            }	bigdata;
        }	val;
    WORD xltype;
    }	XLOPER;

typedef struct xloper __RPC_FAR *LPXLOPER;

typedef struct  _FP
    {
    unsigned short rows;
    unsigned short columns;
    double array[ 1 ];
    }	FP;

typedef struct _FP __RPC_FAR *LPFP;

#endif // INC_xldata32_H
#include <xlw/xll.h>


extern RPC_IF_HANDLE __MIDL_itf_xllExample_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_xllExample_0000_v0_0_s_ifspec;


#ifndef __XLL_EXAMPLE_LIBRARY_DEFINED__
#define __XLL_EXAMPLE_LIBRARY_DEFINED__

/* library XLL_EXAMPLE */
/* [helpstring][version][uuid] */ 

typedef /* [helpstring][uuid][v1_enum] */ 
enum QlOptionType
    {	qlOptionCall	= 1,
	qlOptionPut	= 2,
	qlOptionStraddle	= 3
    }	QlOptionType;


EXTERN_C const IID LIBID_XLL_EXAMPLE;


#ifndef __XLL_Example_MODULE_DEFINED__
#define __XLL_Example_MODULE_DEFINED__


/* module XLL_Example */
/* [dllname][custom][helpstring][uuid] */ 

/* [helpstring][entry] */ void __stdcall xllSayHello( void);

/* [helpstring][entry] */ VARIANT_BOOL __stdcall xllIsNaN( 
    /* [custom][in] */ double number);

/* [helpstring][entry] */ double __stdcall xllCirc( 
    /* [defaultvalue][custom][custom][in] */ double diameter);

/* [custom][helpstring][entry] */ double __stdcall xllStdEuroOptValue( 
    /* [custom][custom][in] */ QlOptionType optionType,
    /* [custom][in] */ double underlying,
    /* [custom][in] */ double strike,
    /* [custom][in] */ double dividendYield,
    /* [custom][custom][in] */ double riskFreeRate,
    /* [custom][in] */ double startTime,
    /* [custom][in] */ double endTime,
    /* [custom][in] */ double volatility);

/* [hidden][helpstring][entry] */ void __stdcall xllMINV( 
    /* [custom][out][in] */ FP __RPC_FAR *Matrix);

#endif /* __XLL_Example_MODULE_DEFINED__ */
#endif /* __XLL_EXAMPLE_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
