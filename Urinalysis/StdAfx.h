// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__75DDABF1_250E_498C_AA84_F4D3C97E1F8D__INCLUDED_)
#define AFX_STDAFX_H__75DDABF1_250E_498C_AA84_F4D3C97E1F8D__INCLUDED_
#undef   WINVER  
#define   WINVER   0X500 
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define VC_EXTRALEAN		// Exclude rarely-used stuff from Windows headers

#include <afxwin.h>         // MFC core and standard components
#include <afxext.h>         // MFC extensions
#include <afxdisp.h>        // MFC Automation classes
#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Common Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT
//#include "SkinPPWTL.h"    //引入美化链接库

#include"SkinH.h"
#pragma comment(lib,"SkinH.lib")
#include <atlbase.h>      //引入atl 保证CRegKey的使用
#import "c:\Program Files\Common Files\System\ado\msado15.dll"no_namespace \
rename("EOF","adoEOF")rename("BOF","adoBOF")
	#define _ATL_APARTMENT_THREADED
#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
class CUrinalysisModule : public CComModule
{
public:
	LONG Unlock();
	LONG Lock();
	LPCTSTR FindOneOf(LPCTSTR p1, LPCTSTR p2);
	DWORD dwThreadID;
};
extern CUrinalysisModule _Module;
#include <atlcom.h>

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__75DDABF1_250E_498C_AA84_F4D3C97E1F8D__INCLUDED_)
