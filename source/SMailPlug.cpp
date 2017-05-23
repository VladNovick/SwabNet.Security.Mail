
////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
//  
////////////////////////////////////////////////////////////////////////////


// SMailPlug.cpp : Implementation of CSMailPlug
#include "stdafx.h"
#include "SwabNetSMail.h"
#include "SMailPlug.h"
#include "SetdInternetLine.h"
#include "ProcessorAppointmentItem.h"
#include "About.h"
#include "AppointmentInfo.h"

/////////////////////////////////////////////////////////////////////////////
// CSMailPlug
_ATL_FUNC_INFO OnClickButtonInfo ={CC_STDCALL,VT_EMPTY,2,{VT_DISPATCH,VT_BYREF | VT_BOOL}};
_ATL_FUNC_INFO OnOptionsAddPagesInfo={CC_STDCALL,VT_EMPTY,1,{VT_DISPATCH}};
_ATL_FUNC_INFO OnInspectorInfo ={CC_STDCALL,VT_EMPTY,1,{VT_DISPATCH}};

_ATL_FUNC_INFO  OnInsertItemCalendarInfo = {CC_STDCALL,VT_EMPTY,1,{VT_DISPATCH}};
_ATL_FUNC_INFO OnChangeItemCalendarInfo = {CC_STDCALL,VT_EMPTY,1,{VT_DISPATCH}};
_ATL_FUNC_INFO OnRemoveItemCalendarInfo = {CC_STDCALL,VT_EMPTY,0,{VT_EMPTY}};


void __stdcall CSMailPlug::OnInsertCalendarItem(IDispatch *Item){
	
	
	CComQIPtr<Outlook::_AppointmentItem> t_spAppointmentItem;
	Item->QueryInterface(&t_spAppointmentItem);
	
	CAppointmentInfo *dat = new CAppointmentInfo();
	dat->save("new",t_spAppointmentItem);
	delete dat;
	
	return ;
}




void __stdcall CSMailPlug::OnChangeCalendarItem(IDispatch *Item){
	CAppointmentInfo *dat = new CAppointmentInfo();
	dat->save("modify",Item);
	delete dat;
	return ;
}

void __stdcall CSMailPlug::OnRemoveCalendarItem(){
	
	
	CComPtr< Outlook::_Explorer > activeExplorer;
	m_spApp->ActiveExplorer(&activeExplorer);
	
	
	
	if (activeExplorer)	{  
		CComPtr<Outlook::_NameSpace> session;
		activeExplorer->get_Session(&session);
		
		if (session){
			
			CComPtr<Outlook::MAPIFolder> mapiDeletedFolder;
			session->GetDefaultFolder(olFolderDeletedItems,&mapiDeletedFolder);
			
			
			CComPtr<Outlook::_Items> deletedItems ;
			mapiDeletedFolder->get_Items(&deletedItems);
			
			if (deletedItems){
				
				long count = 0;
				deletedItems->get_Count(&count); 
				
				if (count > 0 ){
					VARIANT index;
					index.vt = VT_I4;
					index.ulVal = count;
					CComPtr<IDispatch> item;
					deletedItems->Item(index,&item);
					
					if (item) {
						
						CComQIPtr<Outlook::_AppointmentItem> t_spAppointmentItem(item);
						
						if (t_spAppointmentItem ){
							CAppointmentInfo *dat = new CAppointmentInfo();
							dat->save("deleted",t_spAppointmentItem);
							delete dat;		
						}
					}
				}
			}
		}
		
	}  
	
	
	return ;
}	



void __stdcall CSMailPlug::NewInspector(IDispatch* pInspector)
{
	
	static bool go = true; 
	
	HRESULT hr; 
	
	
	
	CComQIPtr<Outlook::_MailItem> spMailItem ; 
	CComQIPtr<Outlook::_MeetingItem> spMeetingItem ; 
	CComQIPtr<Outlook::_ContactItem> spContactItem ; 
	CComQIPtr<Outlook::_TaskItem> spTaskItem ;
	
	
	
	
	CComPtr<IDispatch> spCurrItem; 
	
	CComPtr<Outlook::_Inspector> spInspector; 
    hr = pInspector->QueryInterface(&spInspector); 
    if(FAILED(hr)) return; 
	
	
	
	
	
    // get the current item 
	
    hr = spInspector->get_CurrentItem(&spCurrItem); 
    if(FAILED(hr)) return; 
	
	
	
	
	OlEditorType editType;
	spInspector->get_EditorType(&editType);
	
	
	spCurrItem->QueryInterface(&spMailItem);
    // This is a mail item 
	if (spMailItem != NULL) {
		//
		return; 
	}
	
	
	
	spCurrItem->QueryInterface(Outlook::IID__MeetingItem,(void**)&spMeetingItem);
    // This is a Meeting item 
	if (spMeetingItem != NULL) {
		//
		return; 
	}
	
	
	spCurrItem->QueryInterface(&spContactItem);
    // This is a Contact item 
	if (spContactItem != NULL) {
		//
		return; 
	}
	
	
	spCurrItem->QueryInterface(&spTaskItem);
    // This is a Task item 
	if (spTaskItem != NULL) {
		//
		return; 
	}
	
	
    CProcessorAppointmentItem *pr = NULL;
	pr = CProcessorAppointmentItem::Factory(spInspector);
	if ( pr != NULL ){
		return;
	}
	
	
	
	
}


STDMETHODIMP CSMailPlug::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ISMailPlug
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

void __stdcall CSMailPlug::OnClickButton(IDispatch* /*Office::_CommandBarButton* */ Ctrl,VARIANT_BOOL * CancelDefault)
{
	USES_CONVERSION;
	CComQIPtr<Office::_CommandBarButton> pCommandBarButton(Ctrl);
	
	
}



void __stdcall CSMailPlug::OnClickMenu(IDispatch* /*Office::_CommandBarButton* */ Ctrl,VARIANT_BOOL * CancelDefault)
{
	USES_CONVERSION;
	CComQIPtr<Office::_CommandBarButton> pCommandBarButton(Ctrl);
	
	CAbout::ShowDialog();	
}

void __stdcall CSMailPlug::OnOptionsAddPages(IDispatch *Ctrl)
{
	CComQIPtr<Outlook::PropertyPages> spPages(Ctrl);
	ATLASSERT(spPages);
	CComVariant varProgId(OLESTR("SwabNetSMail.PropSMail"));
	CComBSTR bstrTitle(OLESTR("SwabNet Options"));
	CComVariant varLiteProgId(OLESTR("SwabNetSMail.LiteSPage"));
	CComBSTR bstrLiteTitle(OLESTR("SwabNet History"));
	
	HRESULT hr = spPages->Add((_variant_t)varProgId,(_bstr_t)bstrTitle);
	if(FAILED(hr))
		ATLTRACE("\nFailed addin propertypage");
	hr = spPages->Add((_variant_t)varLiteProgId,(_bstr_t)bstrLiteTitle);
	if(FAILED(hr))
		ATLTRACE("\nFailed adding lite PropertyPage");
	
	
}

