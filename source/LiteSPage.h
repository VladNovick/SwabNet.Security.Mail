
////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
//
////////////////////////////////////////////////////////////////////////////

// LiteSPage.h : Declaration of the CLiteSPage

#ifndef __LITESPAGE_H_
#define __LITESPAGE_H_

#include "resource.h"       // main symbols
#include <atlctl.h>
#import "C:\Program Files\Microsoft Office 2003\OFFICE11\msoutl.olb" raw_interfaces_only, raw_native_types, no_namespace, named_guids 


/////////////////////////////////////////////////////////////////////////////
// CLiteSPage
class ATL_NO_VTABLE CLiteSPage : 
public CComObjectRootEx<CComSingleThreadModel>,
public IDispatchImpl<ILiteSPage, &IID_ILiteSPage, &LIBID_OUTLOOKADDINLib>,
public CComCompositeControl<CLiteSPage>,
public IPersistStreamInitImpl<CLiteSPage>,
public IPersistPropertyBagImpl<CLiteSPage>,
public IOleControlImpl<CLiteSPage>,
public IOleObjectImpl<CLiteSPage>,
public IOleInPlaceActiveObjectImpl<CLiteSPage>,
public IViewObjectExImpl<CLiteSPage>,
public IOleInPlaceObjectWindowlessImpl<CLiteSPage>,
public ISupportErrorInfo,
public CComCoClass<CLiteSPage, &CLSID_LiteSPage>,
public IDispatchImpl<PropertyPage,&__uuidof(PropertyPage),&LIBID_OUTLOOKADDINLib>
{
public:
	CLiteSPage()
	{
		m_bWindowOnly = TRUE;
		CalcExtent(m_sizeExtent);
	}
	
	DECLARE_REGISTRY_RESOURCEID(IDR_LITEPAGE)
		
		DECLARE_PROTECT_FINAL_CONSTRUCT()
		
		BEGIN_COM_MAP(CLiteSPage)
		COM_INTERFACE_ENTRY(ILiteSPage)
		COM_INTERFACE_ENTRY2(IDispatch,ILiteSPage)
		COM_INTERFACE_ENTRY(IViewObjectEx)
		COM_INTERFACE_ENTRY(IViewObject2)
		COM_INTERFACE_ENTRY(IViewObject)
		COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
		COM_INTERFACE_ENTRY(IOleInPlaceObject)
		COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
		COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
		COM_INTERFACE_ENTRY(IOleControl)
		COM_INTERFACE_ENTRY(IOleObject)
		COM_INTERFACE_ENTRY(IPersistStreamInit)
		COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
		COM_INTERFACE_ENTRY(ISupportErrorInfo)
		COM_INTERFACE_ENTRY(PropertyPage)
		END_COM_MAP()
		
		BEGIN_PROP_MAP(CLiteSPage)
		PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
		PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
		// Example entries
		// PROP_ENTRY("Property Description", dispid, clsid)
		// PROP_PAGE(CLSID_StockColorPage)
		END_PROP_MAP()
		
		BEGIN_MSG_MAP(CLiteSPage)
		CHAIN_MSG_MAP(CComCompositeControl<CLiteSPage>)
		END_MSG_MAP()
		// Handler prototypes:
		//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
		//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
		//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
		
		BEGIN_SINK_MAP(CLiteSPage)
		//Make sure the Event Handlers have __stdcall calling convention
		END_SINK_MAP()
		STDMETHOD(SetClientSite)(IOleClientSite *pSite);
	STDMETHOD(GetControlInfo)(LPCONTROLINFO pCI);
	
	STDMETHOD(OnAmbientPropertyChange)(DISPID dispid)
	{
		if (dispid == DISPID_AMBIENT_BACKCOLOR)
		{
			SetBackgroundColorFromAmbient();
			FireViewChange();
		}
		return IOleControlImpl<CLiteSPage>::OnAmbientPropertyChange(dispid);
	}
	
	
	
	// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid)
	{
		static const IID* arr[] = 
		{
			&IID_ILiteSPage,
		};
		for (int i=0; i<sizeof(arr)/sizeof(arr[0]); i++)
		{
			if (InlineIsEqualGUID(*arr[i], riid))
				return S_OK;
		}
		return S_FALSE;
	}
	
	// IViewObjectEx
	DECLARE_VIEW_STATUS(0)
		
		// ILiteSPage
public:
	
	enum { IDD = IDD_LITEPAGE };
	// PropertyPage
	STDMETHOD(GetPageInfo)(BSTR * HelpFile, LONG * HelpContext)
	{
		ATLTRACE("GetPageInfo");
		if (HelpFile == NULL)
			return E_POINTER;
		
		if (HelpContext == NULL)
			return E_POINTER;
		
		return E_NOTIMPL;
	}
	STDMETHOD(get_Dirty)(VARIANT_BOOL * Dirty)
	{
		ATLTRACE("GetDirty");
		if (Dirty == NULL)
			return E_POINTER;
		
		return E_NOTIMPL;
	}
	STDMETHOD(Apply)()
	{
		ATLTRACE("Apply");
		return E_NOTIMPL;
	}
};

#endif //__LITESPAGE_H_
