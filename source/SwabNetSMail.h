/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Jun 05 16:38:59 2008
 */
/* Compiler settings for D:\development\SwabNet.Security.Mail\SwabNetSMail.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __SwabNetSMail_h__
#define __SwabNetSMail_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __ISMailPlug_FWD_DEFINED__
#define __ISMailPlug_FWD_DEFINED__
typedef interface ISMailPlug ISMailPlug;
#endif 	/* __ISMailPlug_FWD_DEFINED__ */


#ifndef __IPropSMail_FWD_DEFINED__
#define __IPropSMail_FWD_DEFINED__
typedef interface IPropSMail IPropSMail;
#endif 	/* __IPropSMail_FWD_DEFINED__ */


#ifndef __ILiteSPage_FWD_DEFINED__
#define __ILiteSPage_FWD_DEFINED__
typedef interface ILiteSPage ILiteSPage;
#endif 	/* __ILiteSPage_FWD_DEFINED__ */


#ifndef __SMailPlug_FWD_DEFINED__
#define __SMailPlug_FWD_DEFINED__

#ifdef __cplusplus
typedef class SMailPlug SMailPlug;
#else
typedef struct SMailPlug SMailPlug;
#endif /* __cplusplus */

#endif 	/* __SMailPlug_FWD_DEFINED__ */


#ifndef __PropSMail_FWD_DEFINED__
#define __PropSMail_FWD_DEFINED__

#ifdef __cplusplus
typedef class PropSMail PropSMail;
#else
typedef struct PropSMail PropSMail;
#endif /* __cplusplus */

#endif 	/* __PropSMail_FWD_DEFINED__ */


#ifndef __LiteSPage_FWD_DEFINED__
#define __LiteSPage_FWD_DEFINED__

#ifdef __cplusplus
typedef class LiteSPage LiteSPage;
#else
typedef struct LiteSPage LiteSPage;
#endif /* __cplusplus */

#endif 	/* __LiteSPage_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __ISMailPlug_INTERFACE_DEFINED__
#define __ISMailPlug_INTERFACE_DEFINED__

/* interface ISMailPlug */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ISMailPlug;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("F8608145-BB5A-4A5A-AE7E-244C6E028093")
    ISMailPlug : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct ISMailPlugVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ISMailPlug __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ISMailPlug __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ISMailPlug __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ISMailPlug __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ISMailPlug __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ISMailPlug __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ISMailPlug __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } ISMailPlugVtbl;

    interface ISMailPlug
    {
        CONST_VTBL struct ISMailPlugVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISMailPlug_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ISMailPlug_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ISMailPlug_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ISMailPlug_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ISMailPlug_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ISMailPlug_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ISMailPlug_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ISMailPlug_INTERFACE_DEFINED__ */


#ifndef __IPropSMail_INTERFACE_DEFINED__
#define __IPropSMail_INTERFACE_DEFINED__

/* interface IPropSMail */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IPropSMail;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("C76C5E4D-FDFF-4756-B925-B29817CE6288")
    IPropSMail : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct IPropSMailVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IPropSMail __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IPropSMail __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IPropSMail __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IPropSMail __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IPropSMail __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IPropSMail __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IPropSMail __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } IPropSMailVtbl;

    interface IPropSMail
    {
        CONST_VTBL struct IPropSMailVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IPropSMail_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IPropSMail_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IPropSMail_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IPropSMail_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IPropSMail_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IPropSMail_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IPropSMail_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IPropSMail_INTERFACE_DEFINED__ */


#ifndef __ILiteSPage_INTERFACE_DEFINED__
#define __ILiteSPage_INTERFACE_DEFINED__

/* interface ILiteSPage */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ILiteSPage;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("067CA4E2-556E-47DD-9EFA-ADAABB2C0D95")
    ILiteSPage : public IDispatch
    {
    public:
    };
    
#else 	/* C style interface */

    typedef struct ILiteSPageVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ILiteSPage __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ILiteSPage __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ILiteSPage __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            ILiteSPage __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            ILiteSPage __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            ILiteSPage __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            ILiteSPage __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } ILiteSPageVtbl;

    interface ILiteSPage
    {
        CONST_VTBL struct ILiteSPageVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ILiteSPage_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ILiteSPage_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ILiteSPage_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ILiteSPage_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define ILiteSPage_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define ILiteSPage_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define ILiteSPage_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ILiteSPage_INTERFACE_DEFINED__ */



#ifndef __OUTLOOKADDINLib_LIBRARY_DEFINED__
#define __OUTLOOKADDINLib_LIBRARY_DEFINED__

/* library OUTLOOKADDINLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_OUTLOOKADDINLib;

EXTERN_C const CLSID CLSID_SMailPlug;

#ifdef __cplusplus

class DECLSPEC_UUID("C704648D-6030-47E9-ADBA-1E13B6A784AE")
SMailPlug;
#endif

EXTERN_C const CLSID CLSID_PropSMail;

#ifdef __cplusplus

class DECLSPEC_UUID("BE1A7148-9F02-40F9-AB5A-5AD4E7D11E62")
PropSMail;
#endif

EXTERN_C const CLSID CLSID_LiteSPage;

#ifdef __cplusplus

class DECLSPEC_UUID("21881BCB-9F1E-42CE-BD5B-ED0E13379A54")
LiteSPage;
#endif
#endif /* __OUTLOOKADDINLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
