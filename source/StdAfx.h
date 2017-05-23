// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#if !defined(AFX_STDAFX_H__6A4E951E_A471_4D04_95D2_73B98E912FF3__INCLUDED_)
#define AFX_STDAFX_H__6A4E951E_A471_4D04_95D2_73B98E912FF3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>
#include <atlhost.h>
#include <atlctl.h>
#include <ObjModel\addauto.h>
#include <ObjModel\appdefs.h>
#include <ObjModel\appauto.h>
#include <ObjModel\blddefs.h>
#include <ObjModel\bldauto.h>
#include <ObjModel\textdefs.h>
#include <ObjModel\textauto.h>
#include <ObjModel\dbgdefs.h>
#include <ObjModel\dbgauto.h>
#include <atlwin.h>

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
#import "C:\Program Files\Common Files\Microsoft Shared\OFFICE11\mso.dll" rename_namespace("Office"), named_guids
using namespace Office;

#import "C:\Program Files\Microsoft Office 2003\OFFICE11\MSOUTL.olb" rename_namespace("Outlook"), named_guids, raw_interfaces_only
using namespace Outlook;





#endif // !defined(AFX_STDAFX_H__6A4E951E_A471_4D04_95D2_73B98E912FF3__INCLUDED)
