// Urinalysis.h : main header file for the URINALYSIS application
//

#if !defined(AFX_URINALYSIS_H__B10AFBF1_8B6F_4A72_897D_0175044E3F1D__INCLUDED_)
#define AFX_URINALYSIS_H__B10AFBF1_8B6F_4A72_897D_0175044E3F1D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif
//#include "SkinPPWTL.h"
#include "resource.h"		// main symbols
#include "Urinalysis_i.h"

/////////////////////////////////////////////////////////////////////////////
// CUrinalysisApp:
// See Urinalysis.cpp for the implementation of this class
//


class CUrinalysisApp : public CWinApp
{
public:
	CUrinalysisApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CUrinalysisApp)
	public:
		virtual BOOL InitInstance();

		virtual int ExitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CUrinalysisApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
private:
	BOOL m_bATLInited;
private:
	BOOL InitATL();
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_URINALYSIS_H__B10AFBF1_8B6F_4A72_897D_0175044E3F1D__INCLUDED_)
