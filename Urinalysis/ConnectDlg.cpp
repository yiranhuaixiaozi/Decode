// ConnectDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Urinalysis.h"
#include "ConnectDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CConnectDlg dialog


CConnectDlg::CConnectDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CConnectDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CConnectDlg)
	m_port = -1;
	m_bandRate = -1;
	//}}AFX_DATA_INIT
}


void CConnectDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CConnectDlg)
	DDX_Radio(pDX, IDC_RADIO1, m_port);
	DDX_Radio(pDX, IDC_RADIO5, m_bandRate);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CConnectDlg, CDialog)
	//{{AFX_MSG_MAP(CConnectDlg)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CConnectDlg message handlers

BOOL CConnectDlg::OnInitDialog() 
{
	CDialog::OnInitDialog();
	// TODO: Add extra initialization here
	CRegKey Rek;
    DWORD port=200;
	DWORD rate=200;
	
    Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
	Rek.QueryValue(port,"Port");
	Rek.QueryValue(rate,"BandRate");
    m_port=(int)port-1;

	switch(int(rate)){
	case 9600: m_bandRate=0;break;
	case 19200: m_bandRate=1;break;
	case 38400: m_bandRate=2;break;
	case 57600:m_bandRate=3;break;
	}
    UpdateData(FALSE);
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CConnectDlg::OnOK() 
{
	// TODO: Add extra validation here
    UpdateData(TRUE);
	CRegKey Rek;
    DWORD port=200;
	DWORD rate=200;
	switch(m_bandRate){
	case 0: rate=9600;break;
	case 1: rate=19200;break;
	case 2: rate=38400;break;
	case 3: rate=57600;break;
	}
	port=DWORD(m_port+1);
    Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
	Rek.SetValue(port,"Port");
	Rek.SetValue(rate,"BandRate");
    Rek.Close();
	CDialog::OnOK();
}
