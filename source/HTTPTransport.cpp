////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
// 
////////////////////////////////////////////////////////////////////////////

// HTTPTransport.cpp: implementation of the CHTTPTransport class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "HTTPTransport.h"
#include "HTTPClient.h"
#include "Utils.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CHTTPTransport::CHTTPTransport(void)
{

}

CHTTPTransport::~CHTTPTransport()
{
}



void CHTTPTransport::HTTPPost(_bstr_t &data){
	
	DWORD TestthreadID;
	HANDLE threadDD;


	char *str = (char *)data;

	unsigned int len= strlen(str)+1;
	char *s = (char *)malloc(len);
    memcpy(s,str,len);


	threadDD = CreateThread(0, 0,(LPTHREAD_START_ROUTINE) TransportThreadFunc,
		(LPVOID)s, 0, &TestthreadID);
	
}


void CHTTPTransport::TransportThreadFunc(LPVOID pArg)
{ 
	char  *t = (char *)pArg;

	CHTTPTransport *transport = new CHTTPTransport();
	try {
	  transport->Process(t);
	} catch (void *){
		//
	}
	free(t);
	delete transport;
}






void CHTTPTransport::Process(char * data){
	   
	   CRegisterSave *mReg;
	   char strHttpAddress[255];


	   
	   mReg = new CRegisterSave;
	   strHttpAddress[0]=0;
	   mReg->UserKeyStore("HTTPServer",strHttpAddress,255,CRegisterSave::READ);
	   delete mReg;
	   
	   
	   CHTTPClient *pClient = new CHTTPClient();
	   

	   _bstr_t currenTime = CUtils::getCurrentTime();

	   const char *strUserName = CUtils::getActiveUserName();
	   const char *strDomainName = CUtils::getComputerDomainName();


		_bstr_t user(strUserName);
		_bstr_t domain(strDomainName);

	   
	   pClient->InitilizePostArguments();
	   pClient->AddPostArguments(_T(__TAG_HTTP_TIME),currenTime);
	   pClient->AddPostArguments(_T(__TAG_HTTP_USER_NAME),user );
	   pClient->AddPostArguments(_T(__TAG_HTTP_DAMAIN),domain);
	   pClient->AddPostArguments(_T(__TAG_HTTP_XML),data,strlen(data) );
	   LPCTSTR szResult;
	   
	   if(pClient->Request(strHttpAddress, 
		   CHTTPClient::RequestPostMethod)){        
		   szResult=pClient->QueryHTTPResponse();
	   }else{
		   
		   int p=0;
		   //
		   
		   
	   }
	   delete pClient;
	   
}
