#ifndef SORTHEADERCTRL_H
#define SORTHEADERCTRL_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CSortHeaderCtrl : public CHeaderCtrl
{
// Construction
public:
	CSortHeaderCtrl();
	void SetFont(int A1,int A2,int A3,int A4,int A5,int A6,int A7,int A8,int A9,int A10,int A11,int A12,int A13);
    LOGFONT     m_Logfont;

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSortHeaderCtrl)
	public:
	virtual void Serialize(CArchive& ar);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CSortHeaderCtrl();

	void SetSortArrow( const int iColumn, const BOOL bAscending );
	COLORREF t_ColorBk;
	COLORREF t_ColorText;

	// Generated message map functions
protected:
	void DrawItem( LPDRAWITEMSTRUCT lpDrawItemStruct );

	int m_iSortColumn;
	BOOL m_bSortAscending;


	//{{AFX_MSG(CSortHeaderCtrl)
		// NOTE - the ClassWizard will add and remove member functions here.
	afx_msg void OnPaint();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // SORTHEADERCTRL_H
