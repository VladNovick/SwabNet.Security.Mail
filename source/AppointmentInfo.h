////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//	 Email :vlad.novick@gmail.com
//
////////////////////////////////////////////////////////////////////////////
// AppointmentInfo.h: interface for the CAppointmentInfo class.
//
//////////////////////////////////////////////////////////////////////



#if !defined(AFX_APPOINTMENTINFO_H__261D9C50_6653_4649_B086_934941CA813C__INCLUDED_)
#define AFX_APPOINTMENTINFO_H__261D9C50_6653_4649_B086_934941CA813C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


#include "comdate.h"
#include "HTTPTransport.h"

#import "C:\Program Files\Common Files\Microsoft Shared\OFFICE11\MSXML5.DLL"
using namespace MSXML2;




class CAppointmentInfo  
{
	
private:
	
	
	
	void addRootDelimiter(CComPtr<MSXML2::IXMLDOMElement> pRoot,CComPtr<MSXML2::IXMLDOMDocument> pDoc){
		
		CComPtr<MSXML2::IXMLDOMText> delimiter = pDoc->createTextNode(L"\n\t");
		pRoot->appendChild(delimiter);
	}
	
	
	void addElement(MSXML2::IXMLDOMElementPtr pRoot, MSXML2::IXMLDOMDocumentPtr pDoc,char *elementName,DATE date){
		char valueTime[30];
		char valueDate[30];
		
		CComDATE *d = new CComDATE(date);
		d->FormatTime(valueTime, NULL, 0, LOCALE_NEUTRAL);
		d->FormatDate(valueDate, NULL, 0, LOCALE_NEUTRAL);
		_bstr_t value(valueDate);
		value += " ";
		value += valueTime;
		addElement(pRoot,pDoc,elementName,value);
		delete d;
	}
	
	void addElement(MSXML2::IXMLDOMElementPtr pRoot, MSXML2::IXMLDOMDocumentPtr pDoc,char *elementName,VARIANT_BOOL b){
		if (b == VARIANT_TRUE){
			addElement(pRoot,pDoc,elementName, "true");
		} else {
			addElement(pRoot,pDoc,elementName, "false");
		}
		
	}
	


/*
	const char * getTextRecurrenceState(OlRecurrenceState &oRecurrenceState){
		switch (oRecurrenceState){
		case olNormal:
			return "Normal";
		case olPersonal:
			return "Personal";
		case olPrivate:
			return "Private";
		case olConfidential:
			return "Confidential";
		}
		return "undefined";
	}	
*/	
	const char * getTextSensitivity(OlSensitivity &oSensitivity){
		switch (oSensitivity){
		case olNormal:
			return "Normal";
		case olPersonal:
			return "Personal";
		case olPrivate:
			return "Private";
		case olConfidential:
			return "Confidential";
		}
		return "undefined";
	}
	
	
	const char * getTextBusyStatus(OlBusyStatus &busyStatus){
		switch (busyStatus){
		case olFree:
			return "Free";
		case olTentative:
			return "Tentative";
		case olBusy:
			return "Busy";
		case olOutOfOffice:
			return "OutOfOffice";
		}
		return "undefined";
	}
	
	

	
	void addAttribute(CComPtr<MSXML2::IXMLDOMElement> pRoot, CComPtr<MSXML2::IXMLDOMDocument> pDoc,char *attributeName,CComVariant value){
		IXMLDOMAttributePtr pa;
		pa = pDoc->createAttribute(attributeName);
		if (pa != NULL) 
		{
			pa->value = value;
			pRoot->setAttributeNode(pa);
			pa.Release();
		}
	}
	
	void addAttribute(CComPtr<MSXML2::IXMLDOMElement> pRoot, CComPtr<MSXML2::IXMLDOMDocument> pDoc,char *attributeName,CComBSTR &value){
		IXMLDOMAttributePtr pa;
		pa = pDoc->createAttribute(attributeName);
		if (pa != NULL) 
		{
			pa->value = CComVariant(value);
			pRoot->setAttributeNode(pa);
			pa.Release();
		}
	}
	
	
	void addElement(MSXML2::IXMLDOMElementPtr pRoot, MSXML2::IXMLDOMDocumentPtr pDoc,char *elementName,BSTR bstr){
		
		_bstr_t value = _bstr_t(bstr,true);
		addElement(pRoot,pDoc,elementName,value);
	}
	
	void addElement(MSXML2::IXMLDOMElementPtr pRoot, MSXML2::IXMLDOMDocumentPtr pDoc,char *elementName,_bstr_t value){
		IXMLDOMElementPtr pe;
		pe=pDoc->createElement(elementName);
		if (pe != NULL)
		{
			pe->text = value;
			pRoot->appendChild(pe);
			pe.Release();
		}
	}
	
	
public:
	CAppointmentInfo();
	virtual ~CAppointmentInfo();
	
	
	void CAppointmentInfo::save(char * dataOperation,CComQIPtr<Outlook::_AppointmentItem> m_spAppointmentItem);	
	
	
};

#endif // !defined(AFX_APPOINTMENTINFO_H__261D9C50_6653_4649_B086_934941CA813C__INCLUDED_)
