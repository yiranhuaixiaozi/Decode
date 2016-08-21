// PrintPreview.cpp : implementation file
//

#include "stdafx.h"
#include "Urinalysis.h"
#include "PrintPreview.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPrintPreview dialog


CPrintPreview::CPrintPreview(CWnd* pParent /*=NULL*/)
	: CDialog(CPrintPreview::IDD, pParent)
{
	//{{AFX_DATA_INIT(CPrintPreview)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
}


void CPrintPreview::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPrintPreview)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPrintPreview, CDialog)
	//{{AFX_MSG_MAP(CPrintPreview)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPrintPreview message handlers
extern BOOL Upload;
BOOL CPrintPreview::OnInitDialog() 
{
	CDialog::OnInitDialog();
	if (!Upload)
	{ 
		GetDlgItem(IDOK)->EnableWindow(FALSE);
	}
   
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}
