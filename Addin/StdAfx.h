// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#if !defined(AFX_STDAFX_H__E8A43C44_80B7_4CAE_B7B9_E651F0F88666__INCLUDED_)
#define AFX_STDAFX_H__E8A43C44_80B7_4CAE_B7B9_E651F0F88666__INCLUDED_

#pragma once

#define STRICT

#ifdef STRICT

#define FOR_WIN2000

#ifdef FOR_WIN98
#define WINVER       0x400
#define _WIN32_WINNT 0x410
#define WIN32_WINNT  0x400
#endif // FOR_WIN98

#ifdef FOR_WINME
#define WINVER              0x45A
#define _WIN32_WINNT        0x45A
#define WIN32_WINNT         0x45A
#define STD_CALL
#define CONDITION_HANDLING  1
#define NU_UP               1
#define NT_INST             0
#define WIN32               100
#define _NT1X               100
#define WINNT               1 
#define WIN32_LEAN_AND_MEAN 1
#define DECLSPEC_NORETURN
#endif

#if defined( FOR_WIN2000 ) || defined( FOR_WINXP )
#define WINVER       0x500
#define _WIN32_WINNT 0x500
#define WIN32_WINNT  0x500
#define DECLSPEC_NORETURN
#endif

#if defined( FOR_WINNT )
#define WINVER       0x400
#define _WIN32_WINNT 0x400
#define WIN32_WINNT  0x400
#endif 

#endif

#define _ATL_APARTMENT_THREADED

// Proxy/stub stuff never required for addins 
//#define _MERGE_PROXYSTUB

#define MSO9_COMPATIBLE

#ifdef _DEBUG
#define _ATL_DEBUG_INTERFACES
#endif

#include <atlbase.h>

//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>
#include <atlhost.h>
#include <atlctl.h>
#include <atlwin.h>

#ifdef _DEBUG

#define CHECKNOTZERO(expr) ATLASSERT( ((expr)!=0) )
#define CHECKNOTNULL(expr) ATLASSERT( ((expr)!=NULL) )
#define CHECKTRUE(expr) ATLASSERT( (expr) )
#define CHECKHR(expr) ATLASSERT( (SUCCEEDED(expr)) )
#define CHECKREGOP(expr) ATLASSERT( ((expr)==ERROR_SUCCESS) )
#define TRACE(expr) ATLTRACE2( (expr) )

#else

#define _BREAK(expr, msg) \
		char buf[1024]; \
		sprintf(buf,"File: %s\nLine: %d\nExpression: %s\n\n Terminating Outlook", __FILE__, __LINE__,msg); \
		MessageBoxA(NULL,buf,"GroupWare ERROR", MB_OK); \
		CAddin::Quit(); //__asm { int 3 };

#define CHECKNOTZERO(expr) if ( (expr) == 0 ) { _BREAK((expr), #expr) } 
#define CHECKNOTNULL(expr) if ( (expr) == NULL ) { _BREAK((expr), #expr) } 
#define CHECKTRUE(expr) if ( (expr) != true ) { _BREAK((expr), #expr) } 
#define CHECKHR(expr) if ( FAILED(expr) ) { _BREAK((expr), #expr) } 
#define CHECKREGOP(expr) if ( (expr) != ERROR_SUCCESS ) { _BREAK((expr), #expr) } 
#define TRACE(expr) void(0)

#endif

#include "resource.h"       // main symbols

#import "MSADDNDR.DLL" raw_interfaces_only, raw_native_types, no_namespace, named_guids, rename("RGB","MsRGB")

// MSO9 import libraries can be used to check for MSO9 compatibility at compile time
//#import "mso9.dll" rename_namespace("Office"), named_guids, rename("RGB","OfRGB"), rename("DocumentProperties","OfDocumentProperties")
//#import "MSOUTL9.OLB" rename_namespace("Outlook"), raw_interfaces_only, named_guids, rename("CopyFile","OlCopyFile")
#import "mso.dll" rename_namespace("Office"), named_guids, rename("RGB","OfRGB"), rename("DocumentProperties","OfDocumentProperties")
#import "MSOUTL.OLB" rename_namespace("Outlook"), raw_interfaces_only, named_guids, rename("CopyFile","OlCopyFile")

#include <wininet.h>
#include <bits.h>

#include <string>
#include <list>
#include <memory>
#include <sstream>
#include <vector>
#include <set>
#include <algorithm>
#include <map>
#include <fstream>

#ifdef UNICODE
#define tstring wstring
#define CAtlRECharTraitsT CAtlRECharTraitsW 
#else
#define tstring string
#define CAtlRECharTraitsT CAtlRECharTraitsA 
#endif

#include "GroupWare.h"
#include "mapi.h"
#include "ldap.h"
#include "COMUtils.h"
#include "AutoUpgrade.h"
#include "wtypes.h"

class CAddin;
class CStore;
class CSyncDialog;
class CTextBox;
class CLogger;
#include "Error.h"
#include "Matrix.h"
#include "Display.h"
#include "Attribute.h"
#include "Entry.h"
#include "StoreArchive.h"
#include "Folder.h"
#include "Cfg.h"
#include "Store.h"
#include "Addin.h"
#include "Logger.h"
#include "TextBox.h"
#include "SyncDialog.h"
#include "CfgDialog.h"
#include "Contact.h"

//{{AFX_INSERT_LOCATION}}

#endif // !defined(AFX_STDAFX_H__E8A43C44_80B7_4CAE_B7B9_E651F0F88666__INCLUDED)
