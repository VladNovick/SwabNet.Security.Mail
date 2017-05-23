////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
// 
////////////////////////////////////////////////////////////////////////////

// AppointmentInfo.cpp: implementation of the CAppointmentInfo class.
// 
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "AppointmentInfo.h"
#include "Utils.h"



//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CAppointmentInfo::CAppointmentInfo()
{
	
}

CAppointmentInfo::~CAppointmentInfo()
{
	
	
}






void CAppointmentInfo::save(char * dataOperation,CComQIPtr<Outlook::_AppointmentItem> m_spAppointmentItem){
	
	
	CComBSTR comBstr;
	CComPtr<MSXML2::IXMLDOMElement> pRoot;
	CComPtr<MSXML2::IXMLDOMDocument> pDoc;

	BOOL sendToServer = FALSE;
	
	CComBSTR  body;
	CComBSTR  location;
	CComBSTR  categories;
	VARIANT_BOOL allDayEvent;
	DATE end;
	DATE creationDate;
	DATE start;
    VARIANT_BOOL  is_saved;
    CComBSTR  subject;
	CComBSTR  entry_ID;
	OlMeetingStatus meetingStatus;
	RecipientsPtr recipients;
	Outlook::OlImportance importance;
	
	try {	
	m_spAppointmentItem->get_Body(&body);
	m_spAppointmentItem->get_AllDayEvent(&allDayEvent);
	m_spAppointmentItem->get_End(&end);
	m_spAppointmentItem->get_MeetingStatus(&meetingStatus);

		m_spAppointmentItem->get_Recipients(&recipients);

	m_spAppointmentItem->get_CreationTime(&creationDate);
	m_spAppointmentItem->get_Start(&start);
	m_spAppointmentItem->get_Saved(&is_saved);
	m_spAppointmentItem->get_Subject(&subject);
	m_spAppointmentItem->get_EntryID(&entry_ID);
	m_spAppointmentItem->get_Location(&location);
	m_spAppointmentItem->get_Categories(&categories);
	
	
	
	
	
	HRESULT hResult;
	
	
	
	CComPtr<MSXML2::IXMLDOMElement> pe;
	CComPtr<MSXML2::IXMLDOMProcessingInstruction> pProcessingInstruction;   
	
	
	hResult=pDoc.CoCreateInstance (__uuidof(MSXML2::DOMDocument30));
	pDoc->put_async(VARIANT_FALSE);
	
	pDoc->validateOnParse = VARIANT_FALSE;
	
	
	pProcessingInstruction=pDoc->createProcessingInstruction(L"xml",   
		L"version='1.0'   encoding='windows-1252'");   
	pDoc->appendChild(pProcessingInstruction); 
	
	
	
	
	
	
	pRoot=pDoc->createElement(L"Appointment"); 
	
	addAttribute(pRoot,pDoc,"ID",CComBSTR(entry_ID));
	
	
	addAttribute(pRoot,pDoc,"Action",CComBSTR(dataOperation));
	
	CComPtr<Outlook::_Application> my_Application;
	m_spAppointmentItem->get_Application(&my_Application);
	


//	CComQIPtr<Outlook::_AppointmentItem> m_spAppointmentItem2;
//	m_spAppointmentItem2->


	CComPtr<Outlook::Recipient> pCurUser;
	CComBSTR UserEmail;
	CComBSTR name;
	CComPtr<Outlook::_NameSpace> nameSpace;
	my_Application->GetNamespace(L"MAPI",&nameSpace);
	nameSpace->get_CurrentUser(&pCurUser);

	CComPtr<AddressEntry> addressEntry;
	pCurUser->get_AddressEntry(&addressEntry);

	_bstr_t user_mail = CUtils::getSMTPMail(addressEntry);

	
	addAttribute(pRoot,pDoc,"user",name);
	addAttribute(pRoot,pDoc,"mail",CComVariant((char *)user_mail));
	
	
	
	
	
	pDoc->appendChild(pRoot);   
	
	
	   const char *strUserName = CUtils::getActiveUserName();
	   const char *strDomainName = CUtils::getComputerDomainName();
	   const char *strComputerName = CUtils::getComputerName();
	   
	   
	   CComBSTR user(strUserName);
	   CComBSTR domain(strDomainName);
	   CComBSTR computer(strComputerName);
	   
	   addRootDelimiter(pRoot,pDoc);
	   pe=pDoc->createElement("registration");
	   
	   
	   pe->appendChild(pDoc->createTextNode("\n\t\t"));
	   addElement(pe,pDoc,"user",user);
	   pe->appendChild(pDoc->createTextNode("\n\t\t"));
	   addElement(pe,pDoc,"workstation",computer);
	   pe->appendChild(pDoc->createTextNode("\n\t\t"));
	   addElement(pe,pDoc,"domain",domain);
	   pRoot->appendChild(pe);
	   
	   
	   
	   
	   
	   
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"subject",subject);
	   
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"location",location);
	   
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"categories",categories);
	   
	   
	   
	   BSTR companies;
	   m_spAppointmentItem->get_Companies(&companies);
	   addElement(pRoot,pDoc,"companies",CComBSTR(companies));
	   
	   //---------------------------
	   
	   
	   
	   CComQIPtr<Outlook::_AppointmentItem> m_spAppointmentItem1;




	   //------------------

		OlSensitivity sensitivity;
	   m_spAppointmentItem->get_Sensitivity(&sensitivity);
	   CComBSTR val(getTextSensitivity(sensitivity));
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"Sensitivity",val);

	   BSTR organaizer;
	   m_spAppointmentItem->get_Organizer(&organaizer);
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"organaizer",CComBSTR(organaizer));
	   
	   
	   addRootDelimiter(pRoot,pDoc);
	   m_spAppointmentItem->get_Importance(&importance);
	   switch (importance){
	   case olImportanceLow:
		   {
			   comBstr = "Low";
			   break;
		   }
	   case olImportanceNormal:
		   {
			   comBstr = "Normal";
			   break;
		   }
	   case olImportanceHigh:
		   {
			   comBstr = "High";
			   break;
		   }
	   }
	   addElement(pRoot,pDoc,"importance",comBstr);
	   
	   //---------------------
	   
	   
	   OlBusyStatus busyStatus;
	   comBstr = getTextBusyStatus(busyStatus);
	   
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"BusyStatus",comBstr);
	   
	   //---------------------	
	   
	   m_spAppointmentItem->get_MeetingStatus(&meetingStatus);
	   
	   comBstr = CComBSTR("");
	   
	   addRootDelimiter(pRoot,pDoc);
	   switch (meetingStatus){
	   case olNonMeeting:
		   {
			   comBstr = "None";
			   break;
		   }
	   case olMeeting:
		   {
			   comBstr = "Meeting";
			   break;
		   }
	   case olMeetingReceived:
		   {
			   comBstr = "Received";
			   break;
		   }
	   case olMeetingCanceled:
		   {
			   comBstr = "Canceled";
			   break;
		   }
	   }
	   addElement(pRoot,pDoc,"MeetingStatus",comBstr);	
	   
	   
	   
	   
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"start",start);
	   
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"end",end);
	   
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"creation",creationDate);
	   
	   
	   
	   DATE lastModify;
	   m_spAppointmentItem->get_LastModificationTime(&lastModify);
	   
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"LastModification",lastModify);
	   
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"allDayEvent",allDayEvent);
	   
	   
	
	   long duration;
	   m_spAppointmentItem->get_Duration(&duration);
	   char buf[255];
	   sprintf(buf,"%d",duration);
	   CComBSTR valDuration(buf);
	   addRootDelimiter(pRoot,pDoc);
	   addElement(pRoot,pDoc,"Duration",valDuration);	   


	   pe=pDoc->createElement("body");
	   
	   addRootDelimiter(pRoot,pDoc);
	   
	   IXMLDOMCDATASectionPtr pcd;
	   pcd = pDoc->createCDATASection(_bstr_t(body,true));
	   if (pcd != NULL) {
		   pe->appendChild(pcd);
		   pcd.Release();
	   }
	   pRoot->appendChild(pe);
	   pe.Release();
	   
	   
	   
	   
	   
	   RecipientPtr r;
	   long count;
	   recipients->get_Count(&count);
	   
	   
	   
	   
	   
	   pe=pDoc->createElement("recipients");
	   if (pe != NULL)
	   {
		   addRootDelimiter(pRoot,pDoc);
		   
		   
		   //---------
		   long i= 1;
		   while ( i <= count){ 
			   
			   
			   IXMLDOMDocumentFragmentPtr pdf;
			   pdf = pDoc->createDocumentFragment();
			   IXMLDOMElementPtr recipientDoc;
			   
			   VARIANT ip;
			   ip.vt = VT_I4;
			   ip.ulVal = i;
			   recipients->Item(ip,&r);
			   pdf->appendChild(pDoc->createTextNode("\n\t\t"));
			   recipientDoc = pDoc->createElement("recipient");
			   
			   BSTR bstrName;
			   
			   r->get_Name(&bstrName);
			   recipientDoc->appendChild(pDoc->createTextNode("\n\t\t\t"));
			   addElement(recipientDoc,pDoc,"name",bstrName);
			   

			   _bstr_t str(bstrName);
			    char *rez = strstr(str,"CNF Bridge");
					if (rez != NULL ){
					   sendToServer = TRUE;
					}
			   
				CComPtr<AddressEntry> addressEntry;
				r->get_AddressEntry(&addressEntry);

				_bstr_t user_mail = CUtils::getSMTPMail(addressEntry);


			   recipientDoc->appendChild(pDoc->createTextNode("\n\t\t\t"));
			   addElement(recipientDoc,pDoc,"address",user_mail);
			   
			   
			   Outlook::OlResponseStatus meetingResponseStatus;
			   r->get_MeetingResponseStatus(&meetingResponseStatus);
			   
			   
			   
			   comBstr = "";
			   
			   
			   
			   
			   switch (meetingResponseStatus){
			   case Outlook::olResponseNone:
				   {
					   comBstr = "None";
					   break;
				   }
			   case Outlook::olResponseOrganized:
				   {
					   comBstr = "Organized";
					   break;
				   }
			   case Outlook::olResponseTentative:
				   {
					   comBstr = "Tentative";
					   break;
				   }
			   case Outlook::olResponseAccepted:
				   {
					   comBstr = "Accepted";
					   break;
				   }
			   case Outlook::olResponseDeclined:
				   {
					   comBstr = "Declined";
					   break;
				   }
			   case Outlook::olResponseNotResponded:
				   {
					   comBstr = "NotResponded";
					   break;
				   }
			   }
			   
			   recipientDoc->appendChild(pDoc->createTextNode("\n\t\t\t"));
			   addElement(recipientDoc,pDoc,"MeetingResponseStatus",comBstr);
			   
			   //--------------------
			   
			   
			   long type;
			   r->get_Type(&type);
			   switch (type){
			   case Outlook::olOrganizer:
				   {
					   comBstr = "Organizer";
					   break;
				   }
			   case Outlook::olRequired:
				   {
					   comBstr = "Required";
					   break;
				   }
			   case Outlook::olOptional:
				   {
					   comBstr = "Optional";
					   break;
				   }
			   default:
				   {
					   char buf[30];
					   sprintf(buf,"Status: %d",type);
					   comBstr = CComBSTR(buf);
				   }
			   }
			   
			   
			   recipientDoc->appendChild(pDoc->createTextNode("\n\t\t\t"));
			   addElement(recipientDoc,pDoc,"MeetingRecipientType",comBstr);
			   
			   recipientDoc->appendChild(pDoc->createTextNode("\n\t\t"));
			   pdf->appendChild(recipientDoc);
			   pdf->appendChild(pDoc->createTextNode("\n\t"));	
			   pe->appendChild(pdf);
			   //		    pdf.Release();			
			   
			   i++;  
			   
		}
		
		
		
		
		//---------
		
		
		
		pRoot->appendChild(pe);
		pe.Release();
		pRoot->appendChild(pDoc->createTextNode("\n"));
		
	}
	
	
	CComBSTR xml;
	pDoc->get_xml(&xml); 
	
	
    _bstr_t bstrt(xml);

	if (sendToServer){
	   CHTTPTransport::HTTPPost(bstrt);
	}
	
	//	pDoc->save("c:\\1.xml");
	
	pDoc.Release();

	} catch (void * ){
		//
	}
	
	
	}
	
	
	
	
	
	
	
	
	
	
