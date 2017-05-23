////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
// 
////////////////////////////////////////////////////////////////////////////

// Utils.cpp: implementation of the CUtils class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Utils.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CUtils::CUtils()
{

}

CUtils::~CUtils()
{

}

  char CUtils::szUser[255];
  char CUtils::szDomain[255];
  char CUtils::strCumputerName[__BUFFER_SIZE];


void CUtils::GetActiveUser(char *userName,unsigned int lenUserNameBuffer,char *domain,unsigned int lDomainNameBuffer)
{

	PSID pSid = NULL;
	BYTE bySidBuffer[256];
	DWORD dwSidSize = sizeof(bySidBuffer);
	DWORD dwDomainNameSize = lDomainNameBuffer;
	DWORD dwUsernameSize = lenUserNameBuffer;
	SID_NAME_USE sidType;   


	pSid = (PSID)bySidBuffer;
	dwSidSize = sizeof(bySidBuffer);

	GetUserName(userName, &dwUsernameSize);

	LookupAccountName(NULL, userName, (PSID) pSid, &dwSidSize,
		domain, &dwDomainNameSize,	(PSID_NAME_USE)&sidType);

		

}



