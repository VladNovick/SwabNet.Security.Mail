////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
// 
////////////////////////////////////////////////////////////////////////////
//
//
// ProcessorAppointmentItem.h: interface for the CProcessorAppointmentItem class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_PROCESSORAPPOINTMENTITEM_H__C2C2BD0A_26B1_4D30_987E_D339E804283B__INCLUDED_)
#define AFX_PROCESSORAPPOINTMENTITEM_H__C2C2BD0A_26B1_4D30_987E_D339E804283B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000



#include "AppointmentInfo.h"

extern _ATL_FUNC_INFO OnDestroyInspectorInfo;
extern _ATL_FUNC_INFO  OnClickSendInfo;
extern  _ATL_FUNC_INFO  OnDeleteItemInfo;



class CProcessorAppointmentItem :
public IDispEventSimpleImpl<1,CProcessorAppointmentItem,&__uuidof(Outlook::ItemEvents )>,
public IDispEventSimpleImpl<3,CProcessorAppointmentItem,&__uuidof(Outlook::ItemEvents )>


{
public:
	
	
	typedef IDispEventSimpleImpl</*nID =*/ 1,CProcessorAppointmentItem, &__uuidof(Outlook::ItemEvents)> DestroyInspectorEvents;
	typedef IDispEventSimpleImpl</*nID =*/ 3,CProcessorAppointmentItem, &__uuidof(Outlook::ItemEvents)> DeleteItemEvents;
	
	
	
	virtual ~CProcessorAppointmentItem();
	static CProcessorAppointmentItem * Factory(CComPtr<Outlook::_Inspector> spInspector);
	
	BEGIN_SINK_MAP(CProcessorAppointmentItem)
		SINK_ENTRY_INFO( 1,__uuidof(Outlook::ItemEvents),/*dispid*/0xf004,OnDestroy,&OnDestroyInspectorInfo)
		SINK_ENTRY_INFO(3, __uuidof(Outlook:: ItemEvents),/*dispid*/ 0xf001,OnDeleteItemEvent, &OnDeleteItemInfo)
		
		END_SINK_MAP()
protected:
	CProcessorAppointmentItem(CComPtr<Outlook::_Inspector> spInspector,CComQIPtr<Outlook::_AppointmentItem> spAppointmentItem);
	CProcessorAppointmentItem();
	
private:
	CommandBarControlPtr commandBarSaveAndClose;
	CComQIPtr<Office::_CommandBarButton> commandBarButtonPtr;
	CComQIPtr<Outlook::_AppointmentItem> spAppointmentItem ;
	CComPtr<Outlook::_Inspector> spInspector; 

public:
	
	
	void __stdcall OnDeleteItemEvent(IDispatch*  
	item ){

	}


	
	
	
	
	void __stdcall OnDestroy(IDispatch * inspector);
	void DispEventAdvise();
	CommandBarControlPtr findCommandBar(CComPtr<Office::_CommandBars> spCmdBars,int ID );
	
	
};

#endif // !defined(AFX_PROCESSORAPPOINTMENTITEM_H__C2C2BD0A_26B1_4D30_987E_D339E804283B__INCLUDED_)
