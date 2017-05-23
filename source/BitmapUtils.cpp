////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
// 
////////////////////////////////////////////////////////////////////////////
// BitmapUtils.cpp: implementation of the CBitmapUtils class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "BitmapUtils.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBitmapUtils::CBitmapUtils()
{

}

CBitmapUtils::~CBitmapUtils()
{

}



void CBitmapUtils::getBitmapSize(HBITMAP  hBitmap, LONG &iw, LONG &ih)
{

	BITMAP bm;
    ::GetObject(hBitmap, sizeof(BITMAP), (LPSTR)&bm);
   

     iw = bm.bmWidth;
	 ih = bm.bmHeight;

}

void CBitmapUtils::DrawBitmap(HDC hDC, long ID_Bitmap, long x, long y)
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