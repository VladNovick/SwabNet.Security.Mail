
////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
//
////////////////////////////////////////////////////////////////////////////

// PropSMail.cpp : Implementation of CPropSMail

#include "stdafx.h"
#include "SwabNetSMail.h"
#include "PropSMail.h"
#include "SetdInternetLine.h"
#include "RegisterSave.h"


#include "About.h"

/////////////////////////////////////////////////////////////////////////////
// CPropSMail

LRESULT CPropSMail::OnClickedAboutSwabNet(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	// TODO : Add Code for control notification handler.
	m_fDirty = true;
	CAbout::ShowDialog();
	return 0;
}




void CPropSMail::DrawBitmap(HDC hDC, long ID_Bitmap, long x, long y)
{
	HBITMAP	  hBitmap = ::LoadBitmap(_Module.GetModuleInstance(),MAKEINTRESOURCE(ID_Bitmap));
	
	BITMAP bm;
    // Create a memory DC.
    HDC hdcMem = ::CreateCompatibleDC(hDC);
    // Select the bitmap into the mem DC.
    ::GetObject(hBitmap, sizeof(BITMAP), (LPSTR)&bm);
	
	HBITMAP bmMemOld    = (HBITMAP)::SelectObject(hdcMem, hBitmap);
	
    // Blt the bits.
    ::BitBlt(hDC,
		x, y,
		bm.bmWidth , bm.bmHeight,
		hdcMem,
		0, 0,
		SRCCOPY);
    ::SelectObject(hdcMem, bmMemOld);
    ::DeleteDC(hdcMem); 
	::DeleteObject(hBitmap);
}



LRESULT CPropSMail::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	
	
	
	   this->reg_store = new CRegisterSave();
	   
	   m_Control = GetParent();  
	   
	   //			 HDC t_hDC = ::GetWindowDC(this->m_hWnd);
	   //            DrawBitmap(t_hDC,IDB_BITMAP_TITLE,0,0);
	   //		    ::ReleaseDC(this->m_hWnd,t_hDC);
	   
	   
	   
	   char buf[255];
	   
	   HWND the_hwnd = GetDlgItem(IDC_EDIT_SERVER_AUTHONTICATION);
	   
	   
	   RECT rectControl;
	   ::GetWindowRect(the_hwnd,&rectControl);
	   
	   
	   
	   RECT rc;  
	   
	   // Get tab control dimensions  
	   
	   m_Control.GetWindowRect(&rc);
	   
	   
	   
	   
	   
	   
	   // set control dimention
	   
	   ::MoveWindow(the_hwnd, rectControl.left-rc.left+10, rectControl.top-rc.top, 
		   (rc.right-rc.left-20),
		   (rectControl.bottom-rectControl.top), true);
	   
	   reg_store->UserKeyStore("HTTPServer",buf,254,CRegisterSave::READ);
	   ::SetWindowText( the_hwnd ,buf);
	   
	   CWindow pWnd ;
	   pWnd.Attach(GetDlgItem(IDC_BUTTON_ABOUT_SWABNET));
	   ATLASSERT(pWnd);
	   CButton pButton =pWnd.Detach();
	   HBITMAP hBitmap =(HBITMAP)::LoadImage(_Module.GetResourceInstance(),
		   MAKEINTRESOURCE(IDB_BITMAP_SWABNET_BUTTON),IMAGE_BITMAP,0,0,LR_LOADMAP3DCOLORS);
	   pButton.SetBitmap( hBitmap );
	   
	   RECT lrect;
	   
	   HWND hwndButton = GetDlgItem(IDC_BUTTON_ABOUT_SWABNET);
	   
	   ::GetWindowRect(hwndButton,&lrect);
	   
	   LONG ix = lrect.left;
	   LONG iy = lrect.top;
	   LONG iw = lrect.right - lrect.left;
	   LONG ih = lrect.bottom - lrect.top;
	   
	   BITMAP bm;
	   ::GetObject(hBitmap, sizeof(BITMAP), (LPSTR)&bm);
	   
	   
	   iw = bm.bmWidth ;
	   ih = bm.bmHeight ;
	   
	   ::MoveWindow(hwndButton, lrect.left-rc.left, lrect.top-rc.top, 
		   (iw+5),
		   (ih+5), true);
	   
	   
	   
/*
	   HWND hWnd = GetDlgItem(IDC_BANNER);
	   if (hWnd)
	   {
		   m_wndBanner.SubclassWindow(hWnd);
		   if (m_wndBanner.Load(MAKEINTRESOURCE(IDR_BANNER),_T("GIF")))
			   m_wndBanner.Draw();
	   };
	   
*/	   
	   
	   return 0;
}





LRESULT CPropSMail::OnKillfocusEdit_server_authontication(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
{
	
	   char buf[255];
	   
	   HWND the_hwnd = GetDlgItem(IDC_EDIT_SERVER_AUTHONTICATION);
	   ::GetWindowText( the_hwnd ,buf,254);
	   reg_store->UserKeyStore("HTTPServer",buf,strlen(buf),CRegisterSave::WRITE);
	   
	   
	   return 0;
}



LRESULT CPropSMail::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	delete reg_store;
	reg_store = NULL;
	return 0;
}


STDMETHODIMP  CPropSMail::GetControlInfo(LPCONTROLINFO pCI ) {
	return S_OK;
};

// IOleObject
STDMETHODIMP CPropSMail::SetClientSite(IOleClientSite *pClientSite){
	HRESULT result;
	// call default ATL implementation
	result = IOleObjectImpl<CPropSMail>::SetClientSite(pClientSite);
	if (result != S_OK) return result;
	// pClientSite may be NULL when container has being destructed
	if (pClientSite != NULL) {
		CComQIPtr<Outlook::PropertyPageSite> pPropertyPageSite(pClientSite);
		result = pPropertyPageSite.CopyTo(&m_pPropSMailSite);
	}
	m_fDirty = false;
	return result;
}

STDMETHODIMP CPropSMail::get_Dirty(VARIANT_BOOL *Dirty)
{
	*Dirty = m_fDirty.boolVal;
	return S_OK;
}
STDMETHODIMP CPropSMail::Apply()
{
	return S_OK;
}
STDMETHODIMP CPropSMail::GetPageInfo(BSTR *HelpFile,LONG *HelpContext)
{
	if (HelpFile == NULL)
		return E_POINTER;
	if (HelpContext == NULL)
		return E_POINTER;
	return E_NOTIMPL;	
}