
////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
//
////////////////////////////////////////////////////////////////////////////

// PropSMail.h : Declaration of the CPropSMail

#ifndef __PROPSMAIL_H_
#define __PROPSMAIL_H_

#include "resource.h"       // main symbols
#include <atlctl.h>
#include "RegisterSave.h"
#include "PictureExWnd.h"
#import "C:\Program Files\Microsoft Office 2003\OFFICE11\msoutl.olb" raw_interfaces_only
#include "atlcontrols.h"
using namespace ATLControls;
/////////////////////////////////////////////////////////////////////////////
// CPropSMail
class ATL_NO_VTABLE CPropSMail : 
public CComObjectRootEx<CComSingleThreadModel>,
public IDispatchImpl<IPropSMail, &IID_IPropSMail, &LIBID_OUTLOOKADDINLib>,
public CComCompositeControl<CPropSMail>,
public IPersistStreamInitImpl<CPropSMail>,
public IOleControlImpl<CPropSMail>,
public IOleObjectImpl<CPropSMail>,
public IOleInPlaceActiveObjectImpl<CPropSMail>,
public IViewObjectExImpl<CPropSMail>,
public IOleInPlaceObjectWindowlessImpl<CPropSMail>,
public IPersistStorageImpl<CPropSMail>,
public ISpecifyPropertyPagesImpl<CPropSMail>,
public IQuickActivateImpl<CPropSMail>,
public IDataObjectImpl<CPropSMail>,
public IProvideClassInfo2Impl<&CLSID_PropSMail, NULL, &LIBID_OUTLOOKADDINLib>,
public CComCoClass<CPropSMail, &CLSID_PropSMail>,
public IPersistPropertyBagImpl<CPropSMail>,
public IDispatchImpl<PropertyPage,&__uuidof(PropertyPage),&LIBID_OUTLOOKADDINLib>
{
public:
	CPropSMail()
	{
		m_bWindowOnly = TRUE;
		CalcExtent(m_sizeExtent);
	}
	
	DECLARE_REGISTRY_RESOURCEID(IDR_PROPPAGE)
		
		DECLARE_PROTECT_FINAL_CONSTRUCT()
		
		BEGIN_COM_MAP(CPropSMail)
		COM_INTERFACE_ENTRY(IPropSMail)
		COM_INTERFACE_ENTRY2(IDispatch,IPropSMail)
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
		COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
		COM_INTERFACE_ENTRY(IQuickActivate)
		COM_INTERFACE_ENTRY(IPersistStorage)
		COM_INTERFACE_ENTRY(IDataObject)
		COM_INTERFACE_ENTRY(IProvideClassInfo)
		COM_INTERFACE_ENTRY(IProvideClassInfo2)
		COM_INTERFACE_ENTRY(IPersistPropertyBag) 
		COM_INTERFACE_ENTRY(Outlook::PropertyPage) 
		END_COM_MAP()
		
		BEGIN_PROP_MAP(CPropSMail)
		PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
		PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
		// Example entries
		// PROP_ENTRY("Property Description", dispid, clsid)
		// PROP_PAGE(CLSID_StockColorPage)
		END_PROP_MAP()
		
		BEGIN_MSG_MAP(CPropSMail)
		CHAIN_MSG_MAP(CComCompositeControl<CPropSMail>)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		COMMAND_HANDLER(IDC_BUTTON_ABOUT_SWABNET, BN_CLICKED, OnClickedAboutSwabNet)
		COMMAND_HANDLER(IDC_EDIT_SERVER_AUTHONTICATION, EN_KILLFOCUS, OnKillfocusEdit_server_authontication)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		END_MSG_MAP()
		// Handler prototypes:
		//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
		//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
		//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
		
		BEGIN_SINK_MAP(CPropSMail)
		//Make sure the Event Handlers have __stdcall calling convention
		END_SINK_MAP()
		
		STDMETHOD(GetControlInfo)(LPCONTROLINFO pCI);
	STDMETHOD(SetClientSite)(IOleClientSite *pSite);
	
	STDMETHOD(OnAmbientPropertyChange)(DISPID dispid)
	{
		if (dispid == DISPID_AMBIENT_BACKCOLOR)
		{
			SetBackgroundColorFromAmbient();
			FireViewChange();
		}
		return IOleControlImpl<CPropSMail>::OnAmbientPropertyChange(dispid);
	}
	
	
	
	// IViewObjectEx
	DECLARE_VIEW_STATUS(0)
		
		// IPropSMail
public:
	
	enum { IDD = IDD_PROPPAGE };
	// PropertyPage
	STDMETHOD(GetPageInfo)(BSTR * HelpFile, LONG * HelpContext);
	STDMETHOD(get_Dirty)(VARIANT_BOOL * Dirty);
	STDMETHOD(Apply)();
	
	
	
private:
	void DrawBitmap(HDC hDC, long ID_Bitmap, long x, long y);
	CWindow m_Control;  
	CPictureExWnd m_wndBanner;
	CComVariant m_fDirty;
	CComPtr<Outlook::PropertyPageSite> m_pPropSMailSite; 
	CRegisterSave *reg_store ;
	
	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	
	
	
	LRESULT OnClickedAboutSwabNet(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	
	
	LRESULT OnKillfocusEdit_server_authontication(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
	
	LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	
	
};

#endif //__PROPSMAIL_H_
