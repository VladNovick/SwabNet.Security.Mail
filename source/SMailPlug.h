
////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
// 
////////////////////////////////////////////////////////////////////////////

// SMailPlug.h : Declaration of the CSMailPlug

#ifndef __SMAILPLUG_H_
#define __SMAILPLUG_H_
#include "stdafx.h"
#include "resource.h"
#include "Utils.h"

     
#import "C:\Program Files\Common Files\Designer\MSADDNDR.DLL" raw_interfaces_only, raw_native_types, no_namespace, named_guids 

/////////////////////////////////////////////////////////////////////////////
// CSMailPlug
extern _ATL_FUNC_INFO OnClickButtonInfo;
extern _ATL_FUNC_INFO OnOptionsAddPagesInfo;
extern _ATL_FUNC_INFO OnInspectorInfo;
extern _ATL_FUNC_INFO OnInsertItemCalendarInfo;
extern _ATL_FUNC_INFO OnChangeItemCalendarInfo;
extern _ATL_FUNC_INFO OnRemoveItemCalendarInfo;




class ATL_NO_VTABLE CSMailPlug : 
public CComObjectRootEx<CComSingleThreadModel>,
public CComCoClass<CSMailPlug, &CLSID_SMailPlug>,
public ISupportErrorInfo,
public IDispatchImpl<ISMailPlug, &IID_ISMailPlug, &LIBID_OUTLOOKADDINLib>,
public IDispatchImpl<_IDTExtensibility2, &IID__IDTExtensibility2, &LIBID_AddInDesignerObjects>,
public IDispEventSimpleImpl<1,CSMailPlug,&__uuidof(Office::_CommandBarButtonEvents)>,
public IDispEventSimpleImpl<2,CSMailPlug,&__uuidof(Office::_CommandBarButtonEvents)>,
public IDispEventSimpleImpl<3,CSMailPlug,&__uuidof(Outlook::ApplicationEvents)>,
public IDispEventSimpleImpl<4,CSMailPlug,&__uuidof(Outlook::InspectorsEvents )>,
public IDispEventSimpleImpl<5,CSMailPlug,&__uuidof(Outlook::ItemsEvents)>,
public IDispEventSimpleImpl<6,CSMailPlug,&__uuidof(Outlook::ItemsEvents)>,
public IDispEventSimpleImpl<7,CSMailPlug,&__uuidof(Outlook::ItemsEvents)>


{
private:

	CComPtr<Outlook::_Items> itemsFolderCalendar ;
public:

	typedef IDispEventSimpleImpl</*nID =*/ 1,CSMailPlug, &__uuidof(Office::_CommandBarButtonEvents)> CommandButton1Events;
	typedef IDispEventSimpleImpl</*nID =*/ 2,CSMailPlug, &__uuidof(Office::_CommandBarButtonEvents)> CommandMenuEvents;
	typedef IDispEventSimpleImpl</*nID =*/ 3,CSMailPlug, &__uuidof(Outlook::ApplicationEvents)> AppEvents;
    typedef IDispEventSimpleImpl</*nID =*/4,CSMailPlug, &__uuidof(Outlook::InspectorsEvents)> MyInspectorsEvents;
    typedef IDispEventSimpleImpl</*nID =*/5,CSMailPlug, &__uuidof(Outlook::ItemsEvents)> ItemCalendarEvents;
    typedef IDispEventSimpleImpl</*nID =*/6,CSMailPlug, &__uuidof(Outlook::ItemsEvents)> ItemCalendarChangeEvents;
    typedef IDispEventSimpleImpl</*nID =*/7,CSMailPlug, &__uuidof(Outlook::ItemsEvents)> ItemCalendarRemoveEvents;




	CSMailPlug()
	{
	}
	
	DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)
		
		DECLARE_PROTECT_FINAL_CONSTRUCT()
		
		BEGIN_COM_MAP(CSMailPlug)
		COM_INTERFACE_ENTRY(ISMailPlug)
		//COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(ISupportErrorInfo)
		COM_INTERFACE_ENTRY2(IDispatch, ISMailPlug)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
		END_COM_MAP()
		
		BEGIN_SINK_MAP(CSMailPlug)
		SINK_ENTRY_INFO(1, __uuidof(Office::_CommandBarButtonEvents),/*dispid*/ 0x01, OnClickButton, &OnClickButtonInfo)
		SINK_ENTRY_INFO(2, __uuidof(Office::_CommandBarButtonEvents),/*dispid*/ 0x01, OnClickMenu, &OnClickButtonInfo)
		SINK_ENTRY_INFO(3,__uuidof(Outlook::ApplicationEvents),/*dispinterface*/,0xf005,OnOptionsAddPages,&OnOptionsAddPagesInfo)
		SINK_ENTRY_INFO(4, __uuidof(Outlook::InspectorsEvents),/*dispid*/ 0xf001, NewInspector, &OnInspectorInfo)
		SINK_ENTRY_INFO(5, __uuidof(Outlook::ItemsEvents),/*dispid*/ 0xf001, OnInsertCalendarItem, &OnInsertItemCalendarInfo)
		SINK_ENTRY_INFO(6, __uuidof(Outlook::ItemsEvents),/*dispid*/ 0xf002, OnChangeCalendarItem, &OnChangeItemCalendarInfo)
		SINK_ENTRY_INFO(7, __uuidof(Outlook::ItemsEvents),/*dispid*/ 0xf003, OnRemoveCalendarItem, &OnRemoveItemCalendarInfo)




		END_SINK_MAP()
		
		
		
		// ISupportsErrorInfo
		STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);
	void __stdcall OnClickButton(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
	void __stdcall OnClickMenu(IDispatch * /*Office::_CommandBarButton**/ Ctrl,VARIANT_BOOL * CancelDefault);
	void __stdcall OnOptionsAddPages(IDispatch* /*PropertyPages**/ Ctrl);
    void __stdcall NewInspector(IDispatch* pdispInspector);


	// ISMailPlug
public:


	// _IDTExtensibility2

	STDMETHOD(OnConnection)(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch * AddInInst, SAFEARRAY * * custom)
	{





	CUtils::initialization();



		CComPtr < Office::_CommandBars> spCmdBars; 
		CComPtr < Office::CommandBar> spCmdBar;
		CComPtr<Outlook::_Inspectors> spInspectors; 

		HRESULT hr;

		
		// QI() for _Application
		CComQIPtr <Outlook::_Application> spApp(Application); 
		ATLASSERT(spApp);
		// get the CommandBars interface that represents Outlook's 		//toolbars & menu 
		//items	
		m_spApp = spApp;

//------------------
/*
		{


			CComPtr<Outlook::Recipient> pCurUser;
	CComBSTR UserEmail;
	CComBSTR name;
	CComPtr<Outlook::_NameSpace> nameSpace;
	CComPtr<Outlook::_NameSpace> session;
	spApp->get_Session(&session);
    OlDefaultFolders folderType = olFolderContacts;
	CComPtr<MAPIFolder> mapiFolder;
	session->GetDefaultFolder(folderType,&mapiFolder);


	
//	spApp->GetNamespace(L"MAPI",&nameSpace);
	session->get_CurrentUser(&pCurUser);
	CComPtr<AddressEntry> addressEntry;
	
	pCurUser->get_AddressEntry(&addressEntry);

	_bstr_t user_mail = CUtils::getSMTPMail(addressEntry);



//	addressEntry->
//	pCurUser->get_Address(&UserEmail);
//	pCurUser->get_Name(&name);
		}

*/
//------------------

		CComPtr<Outlook::_NameSpace> nameSpace;
		CComPtr<Outlook::MAPIFolder> folderCalendar;


		spApp->GetNamespace(L"MAPI",&nameSpace);



		nameSpace->GetDefaultFolder(olFolderCalendar,&folderCalendar);







	   hr = folderCalendar->get_Items(&itemsFolderCalendar);

	   


	   hr = ItemCalendarEvents::DispEventAdvise((IDispatch*)itemsFolderCalendar,&__uuidof(Outlook::ItemsEvents));

	   hr = ItemCalendarChangeEvents::DispEventAdvise((IDispatch*)itemsFolderCalendar,&__uuidof(Outlook::ItemsEvents));

	   hr = ItemCalendarRemoveEvents::DispEventAdvise((IDispatch*)itemsFolderCalendar,&__uuidof(Outlook::ItemsEvents));
		
		
		CComPtr<Outlook::_Explorer> spExplorer; 	
		spApp->ActiveExplorer(&spExplorer);
		spApp->get_Inspectors(&spInspectors);

		m_spInspectors=spInspectors;



		hr = spExplorer->get_CommandBars(&spCmdBars);
		if(FAILED(hr))
			return hr;
		
		ATLASSERT(spCmdBars);
		// now we add a new toolband to Outlook
		
		CComVariant vName("SwabNet Mail Agent");
		CComPtr <Office::CommandBar> spNewCmdBar;
		
		// position it below all toolbands
		//MsoBarPosition::msoBarTop = 1;
		CComVariant vPos(1); 
		
		CComVariant vTemp(VARIANT_TRUE); // menu is temporary		
		CComVariant vEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR);			
		//Add a new toolband through Add method
		// vMenuTemp holds an unspecified parameter
		//spNewCmdBar points to the newly created toolband
		spNewCmdBar = spCmdBars->Add(vName, vPos, vEmpty, vTemp);
		
		//now get the toolband's CommandBarControls
		CComPtr < Office::CommandBarControls> spBarControls;
		spBarControls = spNewCmdBar->GetControls();
		ATLASSERT(spBarControls);
		
		//MsoControlType::msoControlButton = 1
		CComVariant vToolBarType(1);
		//show the toolbar?
		CComVariant vShow(VARIANT_TRUE);
		
		CComPtr < Office::CommandBarControl> spNewBar; 
		
		// add button
		spNewBar = spBarControls->Add(vToolBarType, vEmpty, vEmpty, vEmpty, vShow); 
		ATLASSERT(spNewBar);
		
		
		_bstr_t bstrNewCaption(OLESTR("SGCombo.com"));
		_bstr_t bstrTipText(OLESTR("About SwabNet Mail Agent"));
		
		// get CommandBarButton interface for each toolbar button
		// so we can specify button styles and stuff
		// each button displays a bitmap and caption next to it
		CComQIPtr < Office::_CommandBarButton> spCmdButton(spNewBar);
		
		
		ATLASSERT(spCmdButton);
		
		
		// to set a bitmap to a button, load a 32x32 bitmap
		// and copy it to clipboard. Call CommandBarButton's PasteFace()
		// to copy the bitmap to the button face. to use
		// Outlook's set of predefined bitmap, set button's FaceId to 	//the
		// button whose bitmap you want to use
		HBITMAP hBmp =(HBITMAP)::LoadImage(_Module.GetResourceInstance(),
			MAKEINTRESOURCE(IDB_BITMAP1),IMAGE_BITMAP,0,0,LR_LOADMAP3DCOLORS);
		
		// put bitmap into Clipboard
		::OpenClipboard(NULL);
		::EmptyClipboard();
		::SetClipboardData(CF_BITMAP, (HANDLE)hBmp);
		::CloseClipboard();
		::DeleteObject(hBmp);		
		// set style before setting bitmap
		spCmdButton->PutStyle(Office::msoButtonIconAndCaption);
		
		
		//specify predefined bitmap
		// spCmdButton->PutFaceId(1758); 
		
		hr = spCmdButton->PasteFace();
		if (FAILED(hr))
			return hr;
		
		spCmdButton->PutVisible(VARIANT_TRUE); 
		spCmdButton->PutCaption(OLESTR("SwabNet")); 
		spCmdButton->PutEnabled(VARIANT_TRUE);
		spCmdButton->PutTooltipText(OLESTR("About SwabNet Mail Agent ")); 
		spCmdButton->PutTag(OLESTR("About SwabNet Mail Agent")); 
		
		//show the toolband
		spNewCmdBar->PutVisible(VARIANT_FALSE); 
		
		
		
		
		m_spButton = spCmdButton;
		
		
		_bstr_t bstrNewMenuText(OLESTR("About SwabNet Mail Agent"));
		CComPtr < Office::CommandBarControls> spCmdCtrls;
		CComPtr < Office::CommandBarControls> spCmdBarCtrls; 
		CComPtr < Office::CommandBarPopup> spCmdPopup;
		CComPtr < Office::CommandBarControl> spCmdCtrl;
		
		// get CommandBar that is Outlook's main menu
		hr = spCmdBars->get_ActiveMenuBar(&spCmdBar); 
		if (FAILED(hr))
			return hr;
		// get menu as CommandBarControls 
		spCmdCtrls = spCmdBar->GetControls(); 
		ATLASSERT(spCmdCtrls);
		
		// we want to add a menu entry to Outlook's 6th(Tools) menu 	//item
		CComVariant vItem(5);
		spCmdCtrl= spCmdCtrls->GetItem(vItem);
		ATLASSERT(spCmdCtrl);
		
		IDispatchPtr spDisp;
		spDisp = spCmdCtrl->GetControl(); 
		
		// a CommandBarPopup interface is the actual menu item
		CComQIPtr < Office::CommandBarPopup> ppCmdPopup(spDisp);  
		ATLASSERT(ppCmdPopup);
		
		spCmdBarCtrls = ppCmdPopup->GetControls();
		ATLASSERT(spCmdBarCtrls);
		
		CComVariant vMenuType(1); // type of control - menu
		CComVariant vMenuPos(6);  
		CComVariant vMenuEmpty(DISP_E_PARAMNOTFOUND, VT_ERROR);
		CComVariant vMenuShow(VARIANT_TRUE); // menu should be visible
		CComVariant vMenuTemp(VARIANT_TRUE); // menu is temporary		
		
		
		CComPtr < Office::CommandBarControl> spNewMenu;
		// now create the actual menu item and add it
		spNewMenu = spCmdBarCtrls->Add(vMenuType, vMenuEmpty, vMenuEmpty, vMenuEmpty, vMenuTemp); 
		ATLASSERT(spNewMenu);
		
		spNewMenu->PutCaption(bstrNewMenuText);
		spNewMenu->PutEnabled(VARIANT_TRUE);
		spNewMenu->PutVisible(VARIANT_TRUE); 
		
		
		//we'd like our new menu item to look cool and display
		// an icon. Get menu item  as a CommandBarButton
		CComQIPtr < Office::_CommandBarButton> spCmdMenuButton(spNewMenu);
		ATLASSERT(spCmdMenuButton);
		spCmdMenuButton->PutStyle(Office::msoButtonIconAndCaption);
		
		// we want to use the same toolbar bitmap for menuitem too.
		// we grab the CommandBarButton interface so we can add
		// a bitmap to it through PasteFace().
		spCmdMenuButton->PasteFace(); 
		// show the menu		
		spNewMenu->PutVisible(VARIANT_TRUE); 
		m_spMenu = spCmdMenuButton;
		
		hr = CommandButton1Events::DispEventAdvise((IDispatch*)m_spButton);
		if(FAILED(hr))
			return hr;
		
		
		hr = CommandMenuEvents::DispEventAdvise((IDispatch*)m_spMenu);
		if(FAILED(hr))
			return hr;
		
		hr = AppEvents::DispEventAdvise((IDispatch*)m_spApp,&__uuidof(Outlook::ApplicationEvents));
		if(FAILED(hr))
		{
			ATLTRACE("Failed advising to ApplicationEvents");
			return hr;
		}
		
		hr=MyInspectorsEvents::DispEventAdvise((IDispatch*)m_spInspectors);

		if(FAILED(hr))
			return hr;


		
		bConnected = true;
		
		return S_OK;
		
	}
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
	{
		if(bConnected)
		{
			HRESULT hr = CommandButton1Events::DispEventUnadvise((IDispatch*)m_spButton);
			if(FAILED(hr))
				return hr;
			
			
			hr = CommandMenuEvents::DispEventUnadvise((IDispatch*)m_spMenu);
			if(FAILED(hr))
				return hr;
			hr = AppEvents::DispEventUnadvise((IDispatch*)m_spApp);
			if(FAILED(hr))
				return hr;

			hr = AppEvents::DispEventUnadvise((IDispatch*)m_spApp);
			if(FAILED(hr))

				return hr;
			hr = MyInspectorsEvents::DispEventUnadvise((IDispatch*)m_spInspectors);
			if(FAILED(hr))
				return hr;

			
			
			bConnected = false;
		}
		return S_OK;
		
	}
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom)
	{
		return E_NOTIMPL;
	}
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom)
	{

				RecipientPtr currentUser;

		


		return E_NOTIMPL;
	}
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom)
	{
		return E_NOTIMPL;
	}
	

    void __stdcall OnInsertCalendarItem(IDispatch *Item);
	void __stdcall OnChangeCalendarItem(IDispatch *Item);
	void __stdcall OnRemoveCalendarItem();

	public:
		CComPtr<Outlook::_Application> m_spApp;

	private:
		CComPtr<Office::_CommandBarButton> m_spButton; 
		CComPtr<Office::_CommandBarButton> m_spMenu; 

		CComPtr<Outlook::_Inspectors> m_spInspectors; 
		CComPtr<Outlook::_Inspector>m_spInspector; 
		bool bConnected;
};

#endif //__SMAILPLUG_H_
