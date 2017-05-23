////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
// 
////////////////////////////////////////////////////////////////////////////

// HTTPTransport.h: interface for the CHTTPTransport class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_HTTPTRANSPORT_H__A09DCC60_3A6F_4050_B0D4_BB0B5C6D13AA__INCLUDED_)
#define AFX_HTTPTRANSPORT_H__A09DCC60_3A6F_4050_B0D4_BB0B5C6D13AA__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "RegisterSave.h"
#include "wininet.h"
#include "comdate.h"


#define __TAG_HTTP_XML "xml"
#define __TAG_HTTP_TIME "time"
#define __TAG_HTTP_USER_NAME  "user"
#define __TAG_HTTP_DAMAIN  "domain"
#define __USER_NAME_MAX_SIZE 125;

class CHTTPTransport  
{
public:
	CHTTPTransport() ;
	virtual ~CHTTPTransport();
	static void HTTPPost(_bstr_t &data);

private:
	void  Process(char *);
	static void TransportThreadFunc(LPVOID pArg);

};

#endif // !defined(AFX_HTTPTRANSPORT_H__A09DCC60_3A6F_4050_B0D4_BB0B5C6D13AA__INCLUDED_)
