////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
// 
////////////////////////////////////////////////////////////////////////////

// Utils.h: interface for the CUtils class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_UTILS_H__AC239F1F_C81D_4870_A088_D61B4C9227C0__INCLUDED_)
#define AFX_UTILS_H__AC239F1F_C81D_4870_A088_D61B4C9227C0__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define __BUFFER_SIZE 255

#include "MAPIDEFS.H"

DEFINE_OLEGUID(IID_IMailUser,       0x0002030A, 0, 0);
#define PR_SMTP_ADDRESS 0x39FE001E


#import "C:\Program Files\Common Files\Designer\MSADDNDR.DLL" raw_interfaces_only, raw_native_types, no_namespace, named_guids 


#include "comdate.h"

class CUtils  
{
	
private:
	
    static void GetActiveUser(char *userName,unsigned int lenUserNameBuffer,char *domain,unsigned int lDomainNameBuffer);
	static char szUser[255];
	static char szDomain[255];
	static char strCumputerName[__BUFFER_SIZE];
public:
	CUtils();
	virtual ~CUtils();
	static const char * getActiveUserName(){
		return szUser;
	}
	static const char * getComputerDomainName(){
		return szDomain;
	}
	
	static const char * getComputerName(){
		return strCumputerName;
	}
	
	
	static _bstr_t  getSMTPMail(CComPtr<AddressEntry> addressEntry){
		
		HRESULT hr;
		
		CComPtr<IUnknown> pUnk2= NULL;
		if (addressEntry == NULL )
			return _bstr_t("undefined");
		hr = addressEntry->get_MAPIOBJECT(&pUnk2);
		
		
		CComPtr<IMailUser> pMailUser = NULL;
		hr = pUnk2->QueryInterface(IID_IMailUser, (void**)&pMailUser);
		
		if(pMailUser)
		{

			SPropTagArray props;
			props.aulPropTag[0] = PR_SMTP_ADDRESS;
			props.cValues = 1;
			SPropValue* vals = NULL;
			unsigned long valCount = 0;
			hr = pMailUser->GetProps(&props, 0,
				&valCount, &vals);
			char * bstrAddress = vals->Value.lpszA;
			
			return _bstr_t(bstrAddress);
			
			
		}
		
		return _bstr_t("unknown");
		
		
	}
	
	
	static void initialization(){
		CUtils::GetActiveUser(szUser,sizeof(szUser),  szDomain,sizeof(szDomain));
		unsigned long len = __BUFFER_SIZE;
		::GetComputerName(strCumputerName, &len);
		
	}
	
	
	static _bstr_t getCurrentTime(){
		char valueTime[30];
		char valueDate[30];
		SYSTEMTIME st;
		::GetSystemTime(&st);
		CComDATE *d = new CComDATE(st);
		d->FormatTime(valueTime, NULL, 0, LOCALE_NEUTRAL);
		d->FormatDate(valueDate, NULL, 0, LOCALE_NEUTRAL);
		_bstr_t value(valueDate);
		value += " ";
		value += valueTime;
		delete d;
		return value;
	   }
	
	
	
	
};

#endif // !defined(AFX_UTILS_H__AC239F1F_C81D_4870_A088_D61B4C9227C0__INCLUDED_)
