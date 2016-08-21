#if !defined(AFX_SHOWINFORMATION_H__124EAE5A_F0BE_4E7B_9B46_A4DEAAF9AC11__INCLUDED_)
#define AFX_SHOWINFORMATION_H__124EAE5A_F0BE_4E7B_9B46_A4DEAAF9AC11__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// ShowInformation.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// CShowInformation dialog

class CShowInformation : public CDialog
{
// Construction
public:
	void LoadCheckCaption();
	void LoadListContent();
	CShowInformation(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CShowInformation)
	enum { IDD = IDD_SHOWINF };
	CDragListBox	m_listBox;
	int		m_language;
	BOOL	m_check1;
	BOOL	m_check2;
	BOOL	m_check3;
	BOOL	m_check4;
	BOOL	m_check5;
	BOOL	m_check6;
	BOOL	m_check7;
	BOOL	m_check8;
	BOOL	m_check9;
	BOOL	m_check10;
	BOOL	m_check11;
	//}}AFX_DATA

	CStringArray item;   //存储十一项是否选中标记的标志位 T 或 F
// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CShowInformation)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CShowInformation)
	virtual BOOL OnInitDialog();
	afx_msg void OnRadio1();
	afx_msg void OnRadio2();
	afx_msg void OnCheck1();
	afx_msg void OnCheck2();
	afx_msg void OnCheck3();
	afx_msg void OnCheck4();
	afx_msg void OnCheck5();
	afx_msg void OnCheck6();
	afx_msg void OnCheck7();
	afx_msg void OnCheck8();
	afx_msg void OnCheck9();
	afx_msg void OnCheck10();
	afx_msg void OnCheck11();
	virtual void OnOK();
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SHOWINFORMATION_H__124EAE5A_F0BE_4E7B_9B46_A4DEAAF9AC11__INCLUDED_)
