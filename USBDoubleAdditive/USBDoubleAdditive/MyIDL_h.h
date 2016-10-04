

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Tue Oct 04 11:40:45 2016
 */
/* Compiler settings for MyIDL.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __MyIDL_h_h__
#define __MyIDL_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IUSBDoubleAdditive_FWD_DEFINED__
#define __IUSBDoubleAdditive_FWD_DEFINED__
typedef interface IUSBDoubleAdditive IUSBDoubleAdditive;

#endif 	/* __IUSBDoubleAdditive_FWD_DEFINED__ */


#ifndef ___USBDoubleAdditive_FWD_DEFINED__
#define ___USBDoubleAdditive_FWD_DEFINED__
typedef interface _USBDoubleAdditive _USBDoubleAdditive;

#endif 	/* ___USBDoubleAdditive_FWD_DEFINED__ */


#ifndef __USBDoubleAdditive_FWD_DEFINED__
#define __USBDoubleAdditive_FWD_DEFINED__

#ifdef __cplusplus
typedef class USBDoubleAdditive USBDoubleAdditive;
#else
typedef struct USBDoubleAdditive USBDoubleAdditive;
#endif /* __cplusplus */

#endif 	/* __USBDoubleAdditive_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_MyIDL_0000_0000 */
/* [local] */ 

#pragma once


extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_s_ifspec;


#ifndef __USBDoubleAdditive_LIBRARY_DEFINED__
#define __USBDoubleAdditive_LIBRARY_DEFINED__

/* library USBDoubleAdditive */
/* [version][helpstring][uuid] */ 


EXTERN_C const IID LIBID_USBDoubleAdditive;

#ifndef __IUSBDoubleAdditive_DISPINTERFACE_DEFINED__
#define __IUSBDoubleAdditive_DISPINTERFACE_DEFINED__

/* dispinterface IUSBDoubleAdditive */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IUSBDoubleAdditive;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("023B9FE8-6191-4aca-85C1-25D361D65F51")
    IUSBDoubleAdditive : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IUSBDoubleAdditiveVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IUSBDoubleAdditive * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IUSBDoubleAdditive * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IUSBDoubleAdditive * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IUSBDoubleAdditive * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IUSBDoubleAdditive * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IUSBDoubleAdditive * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IUSBDoubleAdditive * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IUSBDoubleAdditiveVtbl;

    interface IUSBDoubleAdditive
    {
        CONST_VTBL struct IUSBDoubleAdditiveVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IUSBDoubleAdditive_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IUSBDoubleAdditive_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IUSBDoubleAdditive_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IUSBDoubleAdditive_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IUSBDoubleAdditive_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IUSBDoubleAdditive_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IUSBDoubleAdditive_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IUSBDoubleAdditive_DISPINTERFACE_DEFINED__ */


#ifndef ___USBDoubleAdditive_DISPINTERFACE_DEFINED__
#define ___USBDoubleAdditive_DISPINTERFACE_DEFINED__

/* dispinterface _USBDoubleAdditive */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__USBDoubleAdditive;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("89C2E949-342C-4c0f-AD32-26FE966EC70A")
    _USBDoubleAdditive : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _USBDoubleAdditiveVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _USBDoubleAdditive * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _USBDoubleAdditive * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _USBDoubleAdditive * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _USBDoubleAdditive * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _USBDoubleAdditive * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _USBDoubleAdditive * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _USBDoubleAdditive * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _USBDoubleAdditiveVtbl;

    interface _USBDoubleAdditive
    {
        CONST_VTBL struct _USBDoubleAdditiveVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _USBDoubleAdditive_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _USBDoubleAdditive_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _USBDoubleAdditive_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _USBDoubleAdditive_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _USBDoubleAdditive_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _USBDoubleAdditive_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _USBDoubleAdditive_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___USBDoubleAdditive_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_USBDoubleAdditive;

#ifdef __cplusplus

class DECLSPEC_UUID("E68C56D2-0AC6-433e-8B26-255813847301")
USBDoubleAdditive;
#endif
#endif /* __USBDoubleAdditive_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


