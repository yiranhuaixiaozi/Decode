#if !defined(AFX_PRINTPREVIEW_H__B2F17B4C_BDB3_488C_82A0_6A42C2B635F3__INCLUDED_)
#define AFX_PRINTPREVIEW_H__B2F17B4C_BDB3_488C_82A0_6A42C2B635F3__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// PrintPreview.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CPrintPreview dialog

class CPrintPreview : public CDialog
{
// Construction
public:
	CPrintPreview(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CPrintPreview)
	enum { IDD = IDD_PREPRINT };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPrintPreview)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CPrintPreview)
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PRINTPREVIEW_H__B2F17B4C_BDB3_488C_82A0_6A42C2B635F3__INCLUDED_)
