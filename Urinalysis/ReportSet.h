#if !defined(AFX_REPORTSET_H__96173600_2E82_4B0A_BE9F_E3240ADBAC6E__INCLUDED_)
#define AFX_REPORTSET_H__96173600_2E82_4B0A_BE9F_E3240ADBAC6E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ReportSet.h : header file
//报告单设置对话框

/////////////////////////////////////////////////////////////////////////////
// CReportSet dialog

class CReportSet : public CDialog
{
// Construction
public:
	CReportSet(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CReportSet)
	enum { IDD = IDD_DIALOG1 };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CReportSet)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CReportSet)
	afx_msg void OnDestroy();
	virtual BOOL OnInitDialog();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_REPORTSET_H__96173600_2E82_4B0A_BE9F_E3240ADBAC6E__INCLUDED_)
