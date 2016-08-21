// PostBack.cpp : implementation file
//

#include "stdafx.h"
#include "urinalysis.h"
#include "PostBack.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CPostBack dialog
extern CString postbackTime;

CPostBack::CPostBack(CWnd* pParent /*=NULL*/)
	: CDialog(CPostBack::IDD, pParent)
{
	//{{AFX_DATA_INIT(CPostBack)
	m_backTime = CTime::GetCurrentTime();
	//}}AFX_DATA_INIT
}


void CPostBack::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CPostBack)
	DDX_MonthCalCtrl(pDX, IDC_MONTHCALENDAR1, m_backTime);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CPostBack, CDialog)
	//{{AFX_MSG_MAP(CPostBack)
	ON_WM_DESTROY()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CPostBack message handlers

void CPostBack::OnDestroy() 
{    //对话框销毁前postbackTime保存选定要回传的时间
	 SYSTEMTIME sysTime;
     CMonthCalCtrl * pMonthCalendar = (CMonthCalCtrl*)GetDlgItem(IDC_MONTHCALENDAR1);
     pMonthCalendar->GetCurSel(&sysTime);   // 获取选中项的日期
     CString strTime;
     strTime.Format("%4d%2d%2d", sysTime.wYear, sysTime.wMonth, sysTime.wDay);  // 如果数字3将会显示成' 3'，不足两位时，前面以空格填充
     strTime.Replace(' ', '0');   // 将字符串中的空格用字符'0'来代替
     postbackTime=strTime;
	 CDialog::OnDestroy();

	
}
