////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
//
////////////////////////////////////////////////////////////////////////////

// About.h : Declaration of the CAbout

#ifndef __ABOUT_H_
#define __ABOUT_H_

#include "resource.h"       // main symbols
#include <atlhost.h>
#include "PictureExWnd.h"

/////////////////////////////////////////////////////////////////////////////
// CAbout
class CAbout : 
	public CAxDialogImpl<CAbout>
{
private: 
	CPictureExWnd m_wndBanner;
public:
	CAbout()
	{
	}

	~CAbout()
	{
	}

	enum { IDD = IDD_ABOUT };

BEGIN_MSG_MAP(CAbout)
	MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
	COMMAND_ID_HANDLER(IDOK, OnOK)
	COMMAND_ID_HANDLER(IDCANCEL, OnCancel)
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnInitDialog(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);


	LRESULT OnOK(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		EndDialog(wID);
		return 0;
	}

	LRESULT OnCancel(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled)
	{
		EndDialog(wID);
		return 0;
	}


static VOID DialogThreadFunc(LPVOID pArg)
{ 
	static CAbout *about = NULL;
	if (about == NULL ){
		about = new CAbout();
		about->DoModal(::GetDesktopWindow());
		delete about;
		about=NULL;
	} 
	
}





static void ShowDialog(){
	DWORD TestthreadID;
	HANDLE threadDD;
	char *strBuf1 = NULL;
	threadDD = CreateThread(0, 0,(LPTHREAD_START_ROUTINE) DialogThreadFunc,
		(LPVOID)strBuf1, 0, &TestthreadID);
	
}


};

#endif //__ABOUT_H_
