// OptionDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Urinalysis.h"
#include "OptionDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// COptionDlg dialog


COptionDlg::COptionDlg(CWnd* pParent /*=NULL*/)
	: CDialog(COptionDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(COptionDlg)
	//}}AFX_DATA_INIT
}


void COptionDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(COptionDlg)
	DDX_Control(pDX, IDC_COMBO_THEME, m_comboxTheme);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(COptionDlg, CDialog)
	//{{AFX_MSG_MAP(COptionDlg)
	ON_BN_CLICKED(IDC_BUTTON_THEME, OnButtonTheme)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// COptionDlg message handlers

void COptionDlg::OnButtonTheme() 
{
	// TODO: Add your control notification handler code here

		 	CString str;
			int nselect=m_comboxTheme.GetCurSel();							//获取当前组合框选项的索引
			m_comboxTheme.GetLBText(nselect,str);	
			CRegKey Rek;
			Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
			Rek.SetValue(str+".ssk","Theme");
			Rek.Close();
            MessageBox("已修改为"+str);

}

BOOL COptionDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	// TODO: Add extra initialization here
	m_comboxTheme.AddString(_T("AquaOs"));
	m_comboxTheme.AddString(_T("Longhorn"));
	m_comboxTheme.AddString(_T("Vista"));
	m_comboxTheme.AddString(_T("vladstudio"));
	m_comboxTheme.AddString(_T("Steel"));	// TODO: Add extra initialization here
	m_comboxTheme.SetCurSel(0);

	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}


