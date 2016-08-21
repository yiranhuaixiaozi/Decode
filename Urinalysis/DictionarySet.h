#if !defined(AFX_DICTIONARYSET_H__A579BD77_8647_407F_8C65_6F29530D094B__INCLUDED_)
#define AFX_DICTIONARYSET_H__A579BD77_8647_407F_8C65_6F29530D094B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// DictionarySet.h : header file
//数据字典对话框

/////////////////////////////////////////////////////////////////////////////
// CDictionarySet dialog

class CDictionarySet : public CDialog
{
// Construction
public:
	CDictionarySet(CWnd* pParent = NULL);   // standard constructor

// Dialog Data
	//{{AFX_DATA(CDictionarySet)
	enum { IDD = IDD_DICTIONARY };
	CListBox	m_DataValue;
	CListBox	m_DataDictionary;
	CString	m_insertValue;
	CString curSelText;   
	//}}AFX_DATA
  void OnInitADOConn();
   void ExitConnect();
	//添加一个指向Connection对象的指针
	_ConnectionPtr m_pConnection;
	//添加一个指向Recordset对象的指针
	_RecordsetPtr m_pRecordset;
	
	//_variant_t sField[10];
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDictionarySet)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CDictionarySet)
	virtual BOOL OnInitDialog();
	afx_msg void OnDestroy();
	afx_msg void OnSelchangeList1();
	afx_msg void OnDictionaryadd();
	afx_msg void OnDictionarydel();
	afx_msg void OnSelchangeList2();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DICTIONARYSET_H__A579BD77_8647_407F_8C65_6F29530D094B__INCLUDED_)
