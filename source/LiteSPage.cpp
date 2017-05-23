////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
//
////////////////////////////////////////////////////////////////////////////

// LiteSPage.cpp : Implementation of CLiteSPage

#include "stdafx.h"
#include "SwabNetSMail.h"
#include "LiteSPage.h"

/////////////////////////////////////////////////////////////////////////////
// CLiteSPage

STDMETHODIMP CLiteSPage::SetClientSite(IOleClientSite *pSite)
{
	ATLTRACE("SetClientSite");
	if(pSite!=NULL)
	IOleObjectImpl<CLiteSPage>::SetClientSite(pSite);
	
	return S_OK;

	
}
STDMETHODIMP CLiteSPage::GetControlInfo(LPCONTROLINFO pCI)
{
	return S_OK;	
}