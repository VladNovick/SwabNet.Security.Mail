////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
// 
////////////////////////////////////////////////////////////////////////////

// About.cpp : Implementation of CAbout
#include "stdafx.h"
#include "About.h"

/////////////////////////////////////////////////////////////////////////////
// CAbout

LRESULT CAbout::OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{

/*

  HWND hWnd = GetDlgItem(IDC_GIF_ANIMATION);
	   if (hWnd)
	   {
		   m_wndBanner.SubclassWindow(hWnd);
		   if (m_wndBanner.Load(MAKEINTRESOURCE(IDR_BANNER),_T("GIF")))
			   m_wndBanner.Draw();
	   };

*/
		return 1;  // Let the system set the focus
	}
