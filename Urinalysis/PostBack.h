#if !defined(AFX_POSTBACK_H__263EFBA7_A260_4C71_907D_C709FFE7342D__INCLUDED_)
#define AFX_POSTBACK_H__263EFBA7_A260_4C71_907D_C709FFE7342D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// PostBack.h : header file
//批量回传对话框

/////////////////////////////////////////////////////////////////////////////
// CPostBack dialog

class CPostBack : public CDialog
{
// Construction
public:
	CPostBack(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CPostBack)
	enum { IDD = IDD_POSTBACK };
	CTime	m_backTime;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPostBack)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CPostBack)
	afx_msg void OnDestroy();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_POSTBACK_H__263EFBA7_A260_4C71_907D_C709FFE7342D__INCLUDED_)
