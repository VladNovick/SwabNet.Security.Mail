////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
// 
////////////////////////////////////////////////////////////////////////////
//
// ProcessorAppointmentItem.cpp: implementation of the CProcessorAppointmentItem class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ProcessorAppointmentItem.h"

_ATL_FUNC_INFO OnDestroyInspectorInfo ={CC_STDCALL,VT_EMPTY,1,{VT_DISPATCH}};

_ATL_FUNC_INFO  OnDeleteItemInfo = {CC_STDCALL,VT_EMPTY,3,{VT_DISPATCH,VT_BYREF | VT_BOOL}};


void __stdcall CProcessorAppointmentItem::OnDestroy(struct IDispatch *) {
	DestroyInspectorEvents::DispEventUnadvise((IDispatch*) this->spAppointmentItem);
	delete this;
}

CProcessorAppointmentItem::CProcessorAppointmentItem()
{

}

CProcessorAppointmentItem::~CProcessorAppointmentItem()
{

}


CProcessorAppointmentItem::CProcessorAppointmentItem(CComPtr<Outlook::_Inspector> spInspector,CComQIPtr<Outlook::_AppointmentItem> spAppointmentItem){
	this->spInspector = spInspector;
	this->spAppointmentItem = spAppointmentItem;


	CComPtr<Office::_CommandBars> spCmdBars;
	this->spInspector->get_CommandBars(&spCmdBars);

 

}
	
	

	CommandBarControlPtr CProcessorAppointmentItem::findCommandBar(CComPtr<Office::_CommandBars> spCmdBars,int ID ){
	int count = spCmdBars->GetCount();
	
	for (int i=1;i<=count;i++){
		CommandBarPtr bar;
		VARIANT ip;
		ip.vt = VT_I4;
		ip.ulVal = i;
		spCmdBars->get_Item(ip,&bar);
		
		//now get the toolband's CommandBarControls
		CComPtr < Office::CommandBarControls> spBarControls;
		spBarControls = bar->GetControls();


		int countBarControls = spBarControls->GetCount();

		int j = 1;
		while (j <= countBarControls ){

		
		

		CommandBarControlPtr commandBarControl;
				VARIANT ip1;
		ip1.vt = VT_I4;
		ip1.ulVal = j;
		spBarControls->get_Item(ip1,&commandBarControl);


		int id= 0;
		BSTR caption;
		commandBarControl->get_Id(&id);
		commandBarControl->get_Caption(&caption);

			_bstr_t value(caption);

			if (id==ID){ // "&Save and Close"
				 return commandBarControl;
			}

		j++;
		}
	}
		
	
	
	
	
	//    CComQIPtr < Office::_CommandBarButton> spCmdButton(spNewBar);
	
	
	
	
	return NULL;
}


void CProcessorAppointmentItem::DispEventAdvise()
{
	DestroyInspectorEvents::DispEventAdvise((IDispatch*) this->spAppointmentItem);
	DeleteItemEvents::DispEventAdvise((IDispatch*) this->spAppointmentItem);
}


CProcessorAppointmentItem *CProcessorAppointmentItem::Factory(CComPtr<Outlook::_Inspector> spInspector){
	CComQIPtr<Outlook::_AppointmentItem> spAppointmentItem ;
	CComPtr<IDispatch> spCurrItem;
	HRESULT hr; 
	hr = spInspector->get_CurrentItem(&spCurrItem);
	if(FAILED(hr)) return NULL;
	spCurrItem->QueryInterface(&spAppointmentItem);
	// This is an Appointment item 
	if (spAppointmentItem != NULL) {
		CProcessorAppointmentItem *p =  new CProcessorAppointmentItem(spInspector,spAppointmentItem); 
		p->DispEventAdvise();
		return p;
	} else
		return NULL;
}