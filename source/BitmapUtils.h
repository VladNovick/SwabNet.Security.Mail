////////////////////////////////////////////////////////////////////////////
//	Copyright 2008 : Vladimir Novick  (  https://www.linkedin.com/in/vladimirnovick/  )
//
//    NO WARRANTIES ARE EXTENDED. USE AT YOUR OWN RISK. 
//
// To contact the author with suggestions or comments, use  :vlad.novick@gmail.com
// 
////////////////////////////////////////////////////////////////////////////
// BitmapUtils.h: interface for the CBitmapUtils class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BITMAPUTILS_H__1E1FF1FA_73E3_4893_A5DD_C00FBA032225__INCLUDED_)
#define AFX_BITMAPUTILS_H__1E1FF1FA_73E3_4893_A5DD_C00FBA032225__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CBitmapUtils  
{
public:
	CBitmapUtils();
	virtual ~CBitmapUtils();
	void DrawBitmap(HDC hDC, long ID_Bitmap, long x, long y);
	void getBitmapSize(HBITMAP  hBitmap, LONG &iw, LONG &ih);

};

#endif // !defined(AFX_BITMAPUTILS_H__1E1FF1FA_73E3_4893_A5DD_C00FBA032225__INCLUDED_)
