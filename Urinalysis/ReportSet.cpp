// ReportSet.cpp : implementation file
//

#include "stdafx.h"
#include "Urinalysis.h"
#include "ReportSet.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CReportSet dialog
extern CString Caption;

CReportSet::CReportSet(CWnd* pParent /*=NULL*/)
	: CDialog(CReportSet::IDD, pParent)
{
	//{{AFX_DATA_INIT(CReportSet)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CReportSet::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CReportSet)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CReportSet, CDialog)
	//{{AFX_MSG_MAP(CReportSet)
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CReportSet message handlers

void CReportSet::OnDestroy() 
{

	((CEdit*)GetDlgItem(IDC_REPORTCAPTION))->GetWindowText(Caption);//报告单设置对话框关闭前，将编辑框中设计的标题填入全局变量
	
	CDialog::OnDestroy();
	
	// TODO: Add your message handler code here
	
}

BOOL CReportSet::OnInitDialog() 
{
	CDialog::OnInitDialog();

	//将注册表中的标题值来初始化编辑框
	char caption[100];
    CRegKey Rek;
	DWORD cbA=200;
	Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
	Rek.QueryValue(caption,"ReportCaption",&cbA);
   	((CEdit*)GetDlgItem(IDC_REPORTCAPTION))->SetWindowText(caption);
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
