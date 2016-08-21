// UrinalysisDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Urinalysis.h"
#include "UrinalysisDlg.h"
#include "OptionDlg.h"
#include "ShowInformation.h"
#include "PrintPreview.h"
#include "DictionarySet.h"
#include "ReportSet.h"
#include "Login.h"
#include "Splash.h"
#include  "winuser.h"
#include "HyperLinkButton.h"
#include "ConnectDlg.h"
#include"PostBack.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About
CSize	m_Forsize;
double m_x,m_y;    //H �ؼ��Ŵ���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();
	
	// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	CHyperLinkButton	m_homePage;
	//}}AFX_DATA
	
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL
	
	// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	afx_msg void OnLinkbutton();
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnHelp();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	DDX_Control(pDX, IDC_LINKBUTTON, m_homePage);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
//{{AFX_MSG_MAP(CAboutDlg)
	ON_BN_CLICKED(IDC_LINKBUTTON, OnLinkbutton)
	ON_WM_KILLFOCUS()
	ON_BN_CLICKED(IDC_HELP, OnHelp)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CUrinalysisDlg dialog
CString Result[22];           ///ȫ�ֱ����������Ի�����Է��ʣ����汨�浥��11����Ŀ������Լ�ʮ�����Ϣ   ����0-10ʮһ��������11-20������Ϣ
BOOL  Upload=FALSE;           // ȫ�ֱ�������־λ����ʾ�Ƿ������ϴ�������
CString Caption;              // ȫ�ֱ��������⣬��ӡ����Ķ�����
BOOL isOpen=FALSE;            // ȫ�ֱ������Ƿ���������
BOOL isPostback=FALSE;        //�Ƿ����ش�
CString postbackTime;         //ѡ������ش�ʱ��
CUrinalysisDlg::CUrinalysisDlg(CWnd* pParent /*=NULL*/)
: CDialog(CUrinalysisDlg::IDD, pParent)
{ 
	//��ʼ����������
	
	m_themeChange=FALSE;   //�Ƿ�����ˢ�½���
	//{{AFX_DATA_INIT(CUrinalysisDlg)
	m_age = _T("");  
	m_caseNumbe = _T("");
	m_conclusion = _T("");
	m_department = _T("");
	m_doctorName = _T("");
	m_name = _T("");
	m_sampleNumber = _T("");
	m_sampleType = _T("");
	m_sex = -1;
	m_timeSerach = FALSE;
	m_nameSerach = FALSE;
	m_timeSer = CTime::GetCurrentTime();
	m_nameSer = _T("");
	m_item2 = _T("");
	m_item3 = _T("");
	m_item4 = _T("");
	m_item5 = _T("");
	m_item6 = _T("");
	m_item7 = _T("");
	m_item9 = _T("");
	m_item8 = _T("");
	m_item10 = _T("");
	m_item11 = _T("");
	m_item1 = _T("");
	portReceiveMessage="";
	m_time = _T("");
	m_time2 = _T("");
	monitorCount=0;     	//�������ں��յ������ݸ���
	postbackCount=0;      //�ش��������ܸ���
    postbackNumber="";    //�ش��������ܸ���(�ַ�������)

	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	for (int i=0;i<21;i++)            //��ʼ�������
	{
		Result[i]="";         //�޽����ʱ�򣬽����Ĭ������ֵ
	}

	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CUrinalysisDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CUrinalysisDlg)
	DDX_Control(pDX, IDC_TIMESHOW, m_timeShow);
	DDX_Control(pDX, IDC_LIST1, m_grid);
	DDX_Text(pDX, IDC_AGE, m_age);
	DDX_Text(pDX, IDC_CASENUMBER, m_caseNumbe);
	DDX_CBString(pDX, IDC_CONCLUSION, m_conclusion);
	DDX_CBString(pDX, IDC_DEPARTMENT, m_department);
	DDX_CBString(pDX, IDC_DOCTORNAME, m_doctorName);
	DDX_Text(pDX, IDC_NAME, m_name);
	DDX_Text(pDX, IDC_SAMPLENUMBER, m_sampleNumber);
	DDX_CBString(pDX, IDC_SAMPLETYPE, m_sampleType);
	DDX_Radio(pDX, IDC_SEX1, m_sex);
	DDX_Check(pDX, IDC_TIMESERACH, m_timeSerach);
	DDX_Check(pDX, IDC_NAMESERACH, m_nameSerach);
	DDX_DateTimeCtrl(pDX, IDC_TIMESERACH_EDIT, m_timeSer);
	DDX_Text(pDX, IDC_NAMESERACH_EDIT, m_nameSer);
	DDX_Text(pDX, IDC_ITEM2, m_item2);
	DDX_Text(pDX, IDC_ITEM3, m_item3);
	DDX_Text(pDX, IDC_ITEM4, m_item4);
	DDX_Text(pDX, IDC_ITEM5, m_item5);
	DDX_Text(pDX, IDC_ITEM6, m_item6);
	DDX_Text(pDX, IDC_ITEM7, m_item7);
	DDX_Text(pDX, IDC_ITEM9, m_item9);
	DDX_Text(pDX, IDC_ITEM8, m_item8);
	DDX_Text(pDX, IDC_ITEM10, m_item10);
	DDX_Text(pDX, IDC_ITEM11, m_item11);
	DDX_Text(pDX, IDC_ITEM1, m_item1);
	DDX_Control(pDX, IDC_MSCOMM1, m_Comm);
	DDX_Text(pDX, IDC_TIME, m_time);
	DDX_Text(pDX, IDC_TIME2, m_time2);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CUrinalysisDlg, CDialog)
//{{AFX_MSG_MAP(CUrinalysisDlg)
ON_WM_SYSCOMMAND()
ON_WM_PAINT()
ON_WM_QUERYDRAGICON()
ON_WM_CTLCOLOR()
ON_WM_CREATE()
ON_BN_CLICKED(IDC_OPTION, OnOption)
	ON_BN_CLICKED(IDC_SERACH, OnSerach)
	ON_NOTIFY(NM_CLICK, IDC_LIST1, OnClickList1)
	ON_BN_CLICKED(IDC_SHOWINFORMATION, OnShowinformation)
	ON_BN_CLICKED(IDC_SAVEINFORMATION, OnSaveinformation)
	ON_WM_DESTROY()
	ON_WM_TIMER()
	ON_BN_CLICKED(IDC_PRINT, OnPrint)
	ON_BN_CLICKED(IDC_DATADICTIONARY, OnDatadictionary)
	ON_BN_CLICKED(IDC_DELETEDATA, OnDeletedata)
	ON_BN_CLICKED(IDC_CLEAR, OnClear)
	ON_BN_CLICKED(IDC_REPORTSET, OnReportset)
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDC_BUTTON1, OnButton1)
	ON_BN_CLICKED(IDC_CONNECT, OnConnect)
	ON_BN_CLICKED(IDC_REFRESHDATA, OnRefreshdata)
	ON_NOTIFY(NM_SETFOCUS, IDC_LIST1, OnSetfocusList1)
	ON_NOTIFY(LVN_INSERTITEM, IDC_LIST1, OnInsertitemList1)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST1, OnItemchangedList1)
	ON_BN_CLICKED(IDC_POSTBACK, OnPostback)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CUrinalysisDlg message handlers

BOOL CUrinalysisDlg::OnInitDialog()
{   
	Sleep(1300);   //Ϊ�˻�ӭ������ʾ������������ͣ1600ms
	CDialog::OnInitDialog();
	// Add "About..." menu item to system menu.
	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);
	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}
	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
		OnInitADOConn();    //���ݿ��ʼ��
		RegisterActiveX("MSCOMM32.OCX");   //ִ��ע����ûЧ��
		RegisterActiveX("MSADODC.OCX");    //ע��activeX�ؼ�
        RegisterActiveX("MSDATGRD.OCX");   //ע��activeX�ؼ�
		RegisterActiveX("MSSTDFMT.DLL");   //ע��activeX�ؼ�
	          ///��ʼ�����ڣ����ò�������
		InitControlContent();         //�������ֵ�����ȡ��������ռ���
		GetDlgItem(IDC_PRINT)->EnableWindow(false);   //��ӡ��ť��ʼ��������
		GetDlgItem(IDC_REFRESHDATA)->EnableWindow(false);   //���°�ť��ʼ��������
		GetDlgItem(IDC_SAVEINFORMATION)->EnableWindow(false);  
		GetDlgItem(IDC_POSTBACK)->EnableWindow(FALSE);   //�ش���ť������
	  	CString szCurPath(""); 
		GetModuleFileName(NULL,szCurPath.GetBuffer(MAX_PATH),MAX_PATH);
		szCurPath.ReleaseBuffer();  
 		g_curPath = szCurPath.Left(szCurPath.ReverseFind('\\') + 1);
	//	g_languagePath=g_curPath+ "Data\\Language.ini";
		m_Forsize.cx = GetSystemMetrics(SM_CXSCREEN);             //H���ó�����
		m_Forsize.cy = GetSystemMetrics(SM_CYSCREEN);          //H���ó���߶�
// 		if(m_Forsize.cx<1024||m_Forsize.cy<738)
// 			MessageBox(g_LoadString("IDS_RESOLUTION_MESSAGE"));
		SizeWindow();                                         //H�����ؼ��ռ�Ĵ�С
	    MoveWindow(0,0,m_Forsize.cx,m_Forsize.cy);  //H���������Ϊȫ�������κηֱ����£�
        ///////////////////�����Ͽ�ؼ����Ŵ��������������
		CRect rect;
		GetDlgItem(IDC_SAMPLETYPE)->GetWindowRect(&rect);
		ScreenToClient(&rect);//���ؼ���Сת��Ϊ�ڶԻ����е���������
		GetDlgItem(IDC_SAMPLETYPE)->MoveWindow(rect.left,rect.top,rect.Width(),rect.bottom+300);//���ÿؼ���С
	    GetDlgItem(IDC_DOCTORNAME)->GetWindowRect(&rect);
		ScreenToClient(&rect);//���ؼ���Сת��Ϊ�ڶԻ����е���������
		GetDlgItem(IDC_DOCTORNAME)->MoveWindow(rect.left,rect.top,rect.Width(),rect.bottom+300);//���ÿؼ���С
	    GetDlgItem(IDC_DEPARTMENT)->GetWindowRect(&rect);
		ScreenToClient(&rect);//���ؼ���Сת��Ϊ�ڶԻ����е���������
		GetDlgItem(IDC_DEPARTMENT)->MoveWindow(rect.left,rect.top,rect.Width(),rect.bottom+300);//���ÿؼ���С
	    GetDlgItem(IDC_CONCLUSION)->GetWindowRect(&rect);
		ScreenToClient(&rect);//���ؼ���Сת��Ϊ�ڶԻ����е���������
		GetDlgItem(IDC_CONCLUSION)->MoveWindow(rect.left,rect.top,rect.Width(),rect.bottom+300);//���ÿؼ���С

        SetTimer(1,1000,NULL);   //ˢ��ʱ��Ķ�ʱ��
		SkinH_Attach();     //����Ƥ��
	   // AnimateWindow(GetSafeHwnd(),1000,AW_HOR_POSITIVE);   //����Ч���������ڴ�����չ��
		UpdateData(FALSE);
/*		m_grid.SetExtendedStyle(LVS_EX_FLATSB             //list�б�ĳ�ʼ�����ð�������������
			|LVS_EX_FULLROWSELECT
			|LVS_EX_HEADERDRAGDROP
			|LVS_EX_ONECLICKACTIVATE
			|LVS_EX_GRIDLINES);
		m_grid.InsertColumn(0,"�걾��",LVCFMT_LEFT,45,4);
		m_grid.InsertColumn(1,"��������",LVCFMT_LEFT,90,5);
        m_grid.InsertColumn(2,"ʱ��",LVCFMT_LEFT,50,5);
		m_grid.InsertColumn(3,"��������",LVCFMT_LEFT,60,0);  
		m_grid.InsertColumn(4,"�Ա�",LVCFMT_LEFT,40,1);
		m_grid.InsertColumn(5,"����",LVCFMT_LEFT,40,2);
		m_grid.InsertColumn(6,"������",LVCFMT_LEFT,80,3);
		m_grid.SetHeadcolor(RGB(0,0,0),RGB(219,216,215));*/
	m_grid.SetExtendedStyle(LVS_EX_FULLROWSELECT|LVS_EX_CHECKBOXES );   //���϶�ѡΪ������ɫ������Щ
	m_grid.SetHeadcolor(RGB(0,0,0),RGB(215,233,245));
	char t_Text[200];
	sprintf(t_Text,"%s,45;%s,75;%s,60;%s,80;%s,40;%s,40;%s,80","�걾��","��������","ʱ��","��������","�Ա�","����","������");
    m_grid.SetHeadings(_T(t_Text));         //����list�ؼ��ı���
	//��static�ؼ�����ָ������
		pItem[0]=&m_item1;
		pItem[1]=&m_item2;
		pItem[2]=&m_item3;
		pItem[3]=&m_item4;
		pItem[4]=&m_item5;
		pItem[5]=&m_item6;
		pItem[6]=&m_item7;
		pItem[7]=&m_item8;
		pItem[8]=&m_item9;
		pItem[9]=&m_item10;
		pItem[10]=&m_item11;
		InitItemField();   //��ʼ����Ϣ�Ի��򣬽���Ŀ��������
		UpdateData(FALSE);
        DateBaseAutoClear(60000);   //���ݿ�洢����Ϊ60000��
 

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CUrinalysisDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CUrinalysisDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting
		
		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);
		
		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;
		
		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
		
	}
	else
	{
// 		CPaintDC dcPic(GetDlgItem(IDC_PICLOGO));
// 		CRect rc;
// 		GetDlgItem(IDC_PICLOGO)->GetWindowRect(&rc);
// 		CDC dcMem;
// 		BITMAP bmp;
// 		dcMem.CreateCompatibleDC(&dcPic);
// 		dcMem.SelectObject(&logo);
// 		logo.GetBitmap(&bmp);
//        dcPic.StretchBlt(0,0,rc.Width(),rc.Height(),&dcMem,0,0,bmp.bmWidth,bmp.bmHeight,SRCCOPY);
   
	   
	   
	   //	CPaintDC dc(this);
    //	dcPic.BitBlt(0, 0, bmp.bmWidth, bmp.bmHeight, &dcMem, 0, 0, SRCCOPY);
// 		CFile m_file(g_curPath+"Data\\logo.jpg",CFile::modeRead );
// 		//��ȡ�ļ�����
// 		DWORD m_filelen = m_file.GetLength(); 
// 		//�ڶ��Ϸ���ռ�
// 		HGLOBAL m_hglobal = GlobalAlloc(GMEM_MOVEABLE,m_filelen);
// 		LPVOID pvdata = NULL;
// 		//�����ѿռ�,��ȡָ��ѿռ��ָ��
// 		pvdata = GlobalLock(m_hglobal);
// 		//���ļ����ݶ���������
// 		m_file.ReadHuge(pvdata,m_filelen);
// 		IStream*  m_stream;
// 		GlobalUnlock(m_hglobal);
// 		//�ڶ��д���������
// 		CreateStreamOnHGlobal(m_hglobal,TRUE,&m_stream);	
// 		//����������ͼ��
//  		OleLoadPicture(m_stream,m_filelen,TRUE,IID_IPicture,(LPVOID*)&m_picture);
//     	m_stream->Release();
// 		m_picture->get_Width(&m_width);// ��ߣ�MM_HIMETRIC ģʽ����λ��0.01����
// 		m_picture->get_Height(&m_height);
// 		m_IsShow = TRUE;
// 		m_file.Close();
// 		
// 		if (m_IsShow==TRUE){
// 			CRect rect;
// 			GetClientRect(rect);
// 			int nW, nH;
// 			nW = (int)(rect.Width());
// 			nH = (int)(rect.Height());
// 			m_picture->Render(dc,0,0,nW,nH,0,m_height,m_width,-m_height,&rect);
		//}
/*		if (m_themeChange)
		{	
			//////�����������
			char data[100];
			CRegKey Rek;
			DWORD cbA=200;
			Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
			Rek.QueryValue(data,"Theme",&cbA);
			Rek.Close();
			skinppLoadSkin(_T(data));//AquaOS.sskΪ��Ŀ�µ�Ƥ���ļ�	
			m_themeChange=FALSE;
		}*/
	
		CDialog::OnPaint();

		
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CUrinalysisDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}


void CUrinalysisDlg::SizeWindow()           //�����Ի������пؼ��Ĵ�С��ʹ���ȫ��
{
	CRect m_ClientRect;
	GetClientRect(&m_ClientRect);
	m_x = m_Forsize.cx/(double)m_ClientRect.Width() ;//��ȷ���Ŵ���
	m_y = m_Forsize.cy/(double)m_ClientRect.Height();//�߶ȷ���Ŵ���
	//�����ؼ��Ĵ�С
	CWnd *pWnd = NULL; 
	pWnd = GetWindow(GW_CHILD);
	CRect rect;   //��ȡ�ؼ��仯ǰ��С
	while(pWnd)//�ж��Ƿ�Ϊ�գ���Ϊ�Ի��򴴽�ʱ����ô˺���������ʱ�ؼ���δ����
	{
		pWnd->GetWindowRect(&rect);
		ScreenToClient(&rect);//���ؼ���Сת��Ϊ�ڶԻ����е���������
		int width = rect.Width();
		int height = rect.Height();
		//�����ؼ�λ�úʹ�С���仯
		rect.top = m_y * rect.top;
		rect.left = m_x * rect.left;
		rect.bottom = m_y * rect.bottom;
		rect.right = m_x * rect.right;
		pWnd->MoveWindow(&rect);//���ÿؼ���С
		pWnd = pWnd->GetWindow(GW_HWNDNEXT);}
	
	
}
CString CUrinalysisDlg::g_LoadString(CString szID)   //���ⲿ�ļ���ȡ�ַ��ĺ���
{
	CString szValue;
	DWORD dwSize = 1000;
	GetPrivateProfileString("String",szID,"Not found",
		szValue.GetBuffer(dwSize),dwSize,g_languagePath);
	szValue.ReleaseBuffer();
	szValue.Replace("\\n","\n");	//�滻�ػ��з���
	return szValue;
}

HBRUSH CUrinalysisDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)   //�˹���δ��Ч
{
HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
	
	// TODO: Change any attributes of the DC here
	switch (nCtlColor)
	{
		case CTLCOLOR_STATIC:
		{
			if(pWnd->GetDlgCtrlID() == IDC_STATIC1)//��̬�ؼ�
			{
			//	pDC->SetBkColor(RGB(0,255,255));//���ñ���ɫ
				pDC->SetTextColor(RGB(5,0,0));//�����ı���ɫ
				hbr=(HBRUSH)GetStockObject(LTGRAY_BRUSH);//�ؼ��������ɫ
			}
			else
			{
				hbr=CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
			}
		}
		break;
	default:
		break;
	}
	// TODO: Return a different brush if the default is not desired
	return hbr;
}

int CUrinalysisDlg::OnCreate(LPCREATESTRUCT lpCreateStruct)   //////�����ʼ�����漰ע�����ʵĶ���������
{

    CSplashWnd::ShowSplashScreen(this);//��ʾ��������
	
	if (CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;
	InitRegist();

	////////////////////////�����������
	char data[100];
    CRegKey Rek;
	DWORD cbA=200;
	Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
	Rek.QueryValue(data,"Theme",&cbA);
     //skinppLoadSkin(_T(data));//AquaOS.sskΪ��Ŀ�µ�Ƥ���ļ�	
	Rek.Close();
	 //skinppSetNoSkinHwnd(GetDlgItem(IDC_LIST1)->m_hWnd);//
	 //skinppLoadSkin("UMskin.ssk");
	return 0;
}



void CUrinalysisDlg::LoadContent(CStringArray& item,CString language)  //��������Ŀ���Ƽ���CStringArray������
{    
    patientResult.RemoveAll();	
	if (language=="zhongwen")
	{
		if(item.GetAt(0)=="T")patientResult.Add("��ԭ");
		if(item.GetAt(1)=="T")patientResult.Add("ǱѪ"); 
	    if(item.GetAt(2)=="T")patientResult.Add("������");  
		if(item.GetAt(3)=="T")patientResult.Add("ͪ��"); 
		if(item.GetAt(4)=="T")patientResult.Add("��ϸ��");  
		if(item.GetAt(5)=="T")patientResult.Add("������");   
		if(item.GetAt(6)=="T")patientResult.Add("������");  
		if(item.GetAt(7)=="T")patientResult.Add("����");  
		if(item.GetAt(8)=="T")patientResult.Add("��������");
		if(item.GetAt(9)=="T")patientResult.Add("����"); 
		if(item.GetAt(10)=="T")patientResult.Add("ά����C"); 
	}
	
	if (language=="yingwen")
	{
		if(item.GetAt(0)=="T")patientResult.Add("URO");
		if(item.GetAt(1)=="T")patientResult.Add("BLD");
		if(item.GetAt(2)=="T")patientResult.Add("BIL"); 
		if(item.GetAt(3)=="T")patientResult.Add("KET");
		if(item.GetAt(4)=="T")patientResult.Add("ERY"); 
		if(item.GetAt(5)=="T")patientResult.Add("GLU"); 
		if(item.GetAt(6)=="T")patientResult.Add("PRO");
		if(item.GetAt(7)=="T")patientResult.Add("PH"); 
		if(item.GetAt(8)=="T")patientResult.Add("NIT");
		if(item.GetAt(9)=="T")patientResult.Add("SG");
		if(item.GetAt(10)=="T")patientResult.Add("VC");
	}
	
}




//�����һ�δ򿪣�Ϊ��װע��� 
void CUrinalysisDlg::InitRegist()
{
    CRegKey Rek;
	if(Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting")!=ERROR_SUCCESS)   //H��ʧ��
	{
		Rek.Create(HKEY_CURRENT_USER,"HTUrime\\Setting"); 
		Rek.SetValue(1,"Port");
		Rek.SetValue(9600,"BandRate");
	//	Rek.SetValue("AquaOS.ssk","Theme");
        Rek.SetValue("zhongwen","ItemLanguage");
	    Rek.SetValue("HIGHTOP","ReportCaption");

		for(int num=1;num<=11;num++){
		CString s;
		s.Format("%d",num);
        Rek.SetValue("T","Item"+s);
		}
	
	   
	}



	
}
//�򿪸�����ť
void CUrinalysisDlg::OnOption()   //����ѡ�ť��Ӧ
{
	
	COptionDlg dlg;
	if (dlg.DoModal()==IDOK)
	{
		m_themeChange=TRUE;
		Invalidate();
	}

}
//list�б��ѯ��Ӧ����

void CUrinalysisDlg::OnSerach()    ////��ѯ��¼�İ�����Ӧ
{

	int count=0;  //��¼��Ѱ���Ľ������
//	UpdateData(FALSE);
    UpdateData(TRUE);
    m_grid.DeleteAllItems();  //�����������
	GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
	GetDlgItem(IDC_REFRESHDATA)->EnableWindow(FALSE);
	_bstr_t bstrSQL;
   	m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //���ؼ����������   
	m_sex=-1;    //���ؼ����������  
	InitItemField();    //��ս��ʮһ�����ʾ���
	
  UpdateData(FALSE);
if (m_nameSerach==FALSE&&m_timeSerach==FALSE)    
{
   MessageBox("������ѡ��һ�ֲ�ѯ��ʽ","HIGHTOP����", MB_ICONEXCLAMATION );  
   return;
}
if (m_nameSerach&&m_timeSerach)           //������ʱ��˫��ѯ
{    

	if (m_nameSer==""){ MessageBox("������Ҫ������֣�����Ϊ��","HIGHTOP����", MB_ICONEXCLAMATION );  return;}
    CString str;
    CString strTime=m_timeSer.Format("%Y-%m-%d");  //��ȡ����Ϊ���� �磺2010-03-05
    str=" where �������� like '%"+m_nameSer+"%' and ����ʱ��='"+strTime+"'"; 
    bstrSQL = "select*from tb_Sample1"+str+"order by ID asc";	
    
}

else if(m_nameSerach) {	              //��������ѯ
	CString str;
	if (m_nameSer.IsEmpty()) 
	{MessageBox("������Ҫ������֣�����Ϊ��","HIGHTOP����", MB_ICONEXCLAMATION ); return;}
	str=" where �������� like '%"+m_nameSer+"%'"; 
	bstrSQL = "select*from tb_Sample1"+str+"order by ID asc";

}
else if (m_timeSerach){               //��ʱ���ѯ
	CString str; 
	CString strTime=m_timeSer.Format("%Y-%m-%d");  //��ȡ����Ϊ���� �磺2010-03-05
	str=" where ����ʱ��='"+strTime; 
	bstrSQL = "select*from tb_Sample1"+str+"' order by ID asc";
}
//  _variant_t RecordsAffected; 
//CString sql="delete from tb_Sample1 where �걾��,��������,ʱ�� in ( select �걾��,��������,ʱ�� from table group by �걾��,�������� having count(�걾��,��������,ʱ��)>1)";
//m_pConnection->Execute((_bstr_t)sql,&RecordsAffected,adCmdText);
    CString a;
	a.Format("%d",count);
    GetDlgItem(IDC_STATIC1)->SetWindowText("������"+a+"������");	
	m_pRecordset.CreateInstance(__uuidof(Recordset));
	m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
	while(!m_pRecordset->adoEOF)
	{   

		CString s[7];
		s[0]=(char*)(_bstr_t)m_pRecordset->GetCollect("������");
		s[1]=(char*)(_bstr_t)m_pRecordset->GetCollect("����ʱ��");
        s[2]=(char*)(_bstr_t)m_pRecordset->GetCollect("ʱ��");
		s[3]=(char*)(_bstr_t)m_pRecordset->GetCollect("��������");
		s[4]=(char*)(_bstr_t)m_pRecordset->GetCollect("�Ա�");
		s[5]=(char*)(_bstr_t)m_pRecordset->GetCollect("����");
		s[6]=(char*)(_bstr_t)m_pRecordset->GetCollect("������");

		
		m_grid.AddItem(s[0],s[1],s[2],s[3],s[4],s[5],s[6],"");
        //��ÿ����Ϣ�����ɫ
        CString color=(char*)(_bstr_t)m_pRecordset->GetCollect("�Ƿ��쳣");
		if(color=="1")m_grid.SetItemColor(count,0,RGB(0,0,0),RGB(255,0,0));
        if(color=="0")m_grid.SetItemColor(count,0,RGB(0,0,0),RGB(0,255,0));
	//	m_grid.SetItemText(0,0,(char*)(_bstr_t)m_pRecordset->GetCollect("������"));
	//	m_grid.SetItemText(0,1,(char*)(_bstr_t)m_pRecordset->GetCollect("����ʱ��"));
	 //  m_grid.SetItemText(0,2,(char*)(_bstr_t)m_pRecordset->GetCollect("ʱ��"));
	//	m_grid.SetItemText(0,3,(char*)(_bstr_t)m_pRecordset->GetCollect("��������"));
	//	m_grid.SetItemText(0,4,(char*)(_bstr_t)m_pRecordset->GetCollect("�Ա�"));
	//	m_grid.SetItemText(0,5,(char*)(_bstr_t)m_pRecordset->GetCollect("����"));
	//	m_grid.SetItemText(0,6,(char*)(_bstr_t)m_pRecordset->GetCollect("������"));
		
	    count++;     //��������������
		CString a;
	    a.Format("%d",count);
        GetDlgItem(IDC_STATIC1)->SetWindowText("������"+a+"������");
		m_pRecordset->MoveNext();
	}
	
	m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //���ؼ����������
    m_sex=-1;    //���ؼ����������
      
}

void CUrinalysisDlg::OnDeletedata()              ///֧��ѡ�ж��ͬʱɾ��
{

int count=0;     //ɾ������������
CString s;
s.Format("%d",count);
GetDlgItem(IDC_STATIC1)->SetWindowText("�ɹ�ɾ��"+s+"������");   //��ʾ�ɹ�ɾ��0������
//���¿���ʵ�ֿ�ѡɾ��
int *pBuf = new int[m_grid.GetItemCount()];
int num = 0;
POSITION pos;
pos = m_grid.GetFirstSelectedItemPosition();
while(pos != NULL)
{
     pBuf[num++] = m_grid.GetNextSelectedItem(pos);	
	
}
for(int i=num; i>0; i--){
    	CString str1=m_grid.GetItemText(pBuf[i-1],0);   //������	
		CString str2=m_grid.GetItemText(pBuf[i-1],1);  //����ʱ��
		CString str3=m_grid.GetItemText(pBuf[i-1],2);//ʱ��
	    CString bstrSQL;
		_variant_t RecordsAffected;
		bstrSQL = "delete from tb_Sample1 where ������='"+str1+"' and ����ʱ��='"+str2+"' and ʱ��='"+str3+"'";
		m_pConnection->Execute((_bstr_t)bstrSQL,&RecordsAffected,adCmdText); 
		m_grid.DeleteItem(pBuf[i-1]);
		count++;          //ɾ������������
		CString s;
		s.Format("%d",count);
        GetDlgItem(IDC_STATIC1)->SetWindowText("�ɹ�ɾ��"+s+"������");
		
}
        delete [] pBuf;	// TODO: Add your control notification handler code here
		GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);              //��ӡ��ťʧЧ
	    GetDlgItem(IDC_REFRESHDATA)->EnableWindow(FALSE);        //���°�ťʧЧ
		
	
	
		m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //���ؼ����������
        m_sex=-1;    //���ؼ����������
        InitItemField();    //��ս��ʮһ�����ʾ���
		UpdateData(FALSE);   
/*
	for(int i=0; i<m_grid.GetItemCount(); i++){					//���������б���ͼ
		
		if(m_grid.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED ){
			CString str1=m_grid.GetItemText(i,0);   //������	
			CString str2=m_grid.GetItemText(i,1);  //����ʱ��
		    CString str3=m_grid.GetItemText(i,2);//ʱ��
			CString bstrSQL;
			_variant_t RecordsAffected;
	
			bstrSQL = "delete from tb_Sample1 where ������='"+str1+"' and ����ʱ��='"+str2+"' and ʱ��='"+str3+"'";
			m_pConnection->Execute((_bstr_t)bstrSQL,&RecordsAffected,adCmdText); 
	        m_grid.DeleteItem(i);
			
			GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
	        GetDlgItem(IDC_REFRESHDATA)->EnableWindow(FALSE);
            GetDlgItem(IDC_STATIC1)->SetWindowText("�ɹ�ɾ��1������");
			m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //���ؼ����������
            m_sex=-1;    //���ؼ����������
		
           	UpdateData(FALSE);
		}	//��ȡѡ����
				
	}  */

}


		

void CUrinalysisDlg::OnClear()        //��հ�ť��Ӧ
{
	// TODO: Add your control notification handler code here
	m_grid.DeleteAllItems();
	GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
	GetDlgItem(IDC_REFRESHDATA)->EnableWindow(FALSE);
	GetDlgItem(IDC_STATIC1)->SetWindowText("�б������");
	m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //���ؼ����������
    m_sex=-1;    //���ؼ����������
    InitItemField();    //��ս��ʮһ�����ʾ���
	UpdateData(FALSE);
}





void CUrinalysisDlg::OnInitADOConn()   //�������ݿ�����
{
	::CoInitialize(NULL);
	try
	{
		m_pConnection.CreateInstance("ADODB.Connection");
		_bstr_t strConnect="uid=;pwd=;DRIVER={Microsoft Access Driver (*.mdb)};DBQ=db_Data.mdb;";
		m_pConnection->Open(strConnect,"","",adModeUnknown);
	}
	catch(_com_error e)
	{
		//AfxMessageBox(e.Description());
		AfxMessageBox("���ݿⶪʧ");
	}
}

void CUrinalysisDlg::ExitConnect()     //�ر����ݿ�����
{
	if(m_pRecordset!=NULL)
		m_pRecordset->Close();
	m_pConnection->Close();
}

void CUrinalysisDlg::OnClickList1(NMHDR* pNMHDR, LRESULT* pResult) 
{   
	int count=0;   //�ж�ѡ���˼����һ������ʾ����2������������ϸ��ʾ
	// TODO: Add your control notification handler code here
	for(int i=0; i<m_grid.GetItemCount(); i++)					//���������б���ͼ
	{ 
		if(m_grid.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED )	//��ȡѡ����
		{   count++;
		    m_sex=-1;  //�Ƚ��Ա�ѡ��ť��Ϊ��ѡ��״̬
			GetDlgItem(IDC_PRINT)->EnableWindow(TRUE);
         	GetDlgItem(IDC_REFRESHDATA)->EnableWindow(TRUE);
			CString str1=m_grid.GetItemText(i,0);   //������	
			CString str2=m_grid.GetItemText(i,1);   //���ʱ��
			CString str3=m_grid.GetItemText(i,2);//ʱ��
			_bstr_t bstrSQL;
			bstrSQL = "select *from tb_Sample1 where ������='"+str1+"' and ����ʱ��='"+str2+"' and ʱ��='"+str3+"' order by ID desc";	
			m_pRecordset.CreateInstance(__uuidof(Recordset));
			m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
      
			//   while(!m_pRecordset->adoEOF)
	 //  { 
            if (m_pRecordset->adoEOF)     //��������ɾ�����б����Դ��ڣ�����ʾ�û�������Ч
            {
			  MessageBox("����������Ч���볢������������ֱ��ɾ��","HIGHTOP����", MB_ICONSTOP );  	
         
			}
            if (count>1)return;   //�ж�ѡ���˼����һ������ʾ����2��������ֱ��������������ִ��
			if(!m_pRecordset->adoEOF){        //��������
				
				m_name=(char*)(_bstr_t)m_pRecordset->GetCollect("��������");
				CString s1="��";
				CString s2="Ů";
                if ((char*)(_bstr_t)m_pRecordset->GetCollect("�Ա�")==s1)
                {
					m_sex=0;
                }
				if ((char*)(_bstr_t)m_pRecordset->GetCollect("�Ա�")==s2)
				{
					m_sex=1;
				
				}
	        	m_age=(char*)(_bstr_t)m_pRecordset->GetCollect("����");
				m_caseNumbe=(char*)(_bstr_t)m_pRecordset->GetCollect("������");
				m_sampleType=(char*)(_bstr_t)m_pRecordset->GetCollect("��������");
				m_sampleNumber=(char*)(_bstr_t)m_pRecordset->GetCollect("������");
				m_doctorName=(char*)(_bstr_t)m_pRecordset->GetCollect("ҽ��");
				m_department=(char*)(_bstr_t)m_pRecordset->GetCollect("����");
				///////////////��ʱ��ֵ�������ڿؼ���������
			//	CString s=(char*)(_bstr_t)m_pRecordset->GetCollect("����ʱ��");
// 				int  nYear,nMonth,nDate;   
// 				sscanf(s,"%d-%d-%d", &nYear, &nMonth, &nDate);   
// 				CTime  t(nYear,   nMonth,   nDate,   0,   0,   0);

				m_time=(char*)(_bstr_t)m_pRecordset->GetCollect("����ʱ��");
				m_conclusion=(char*)(_bstr_t)m_pRecordset->GetCollect("���");
				m_time2=(char*)(_bstr_t)m_pRecordset->GetCollect("ʱ��");
  					for (int p=0;p<21;p++)            //��21����Ϣȫ�����
					{
	                	Result[p]="";         
					}
		        //������ʮ����Ϣд��result������
				Result[11]=m_name;
				Result[12]=(char*)(_bstr_t)m_pRecordset->GetCollect("�Ա�");
                Result[13]=m_age;
				Result[14]=m_caseNumbe;
				Result[15]=m_sampleType;
				Result[16]=m_sampleNumber;
				Result[17]=m_doctorName;
				Result[18]=m_department;
			//	s.Left(12);
			//	Result[19]=s;
				Result[19]=m_time;
				Result[20]=m_conclusion;
				Result[21]=m_time2;
          ///////////��ʼ����Ϣ�� ����һ������ʾ����Ŀ��䣬�ڶ������������ַ��󲹼���Ŀ���
		    CString language=InitItemField();   //��ʮһ��������䵽�ı��ؼ���
			//��ʮһ������䵽�ı��ؼ���
 			for (int i=0;i<11;i++)
 		{ 
		    if (*pItem[i]=="��ԭ"||*pItem[i]=="URO")
				{
					if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{ *pItem[i]+=InsertSpace(12);}
	                *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("��ԭ");
					Result[0]=(char*)(_bstr_t)m_pRecordset->GetCollect("��ԭ");    //Result�����¼ѡ�в���������������Ϣ

				}
			if (*pItem[i]=="ǱѪ"||*pItem[i]=="BLD"){
                if(language=="zhongwen"){*pItem[i]+=InsertSpace(14);}else{ *pItem[i]+=InsertSpace(12);}
			   *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("ǱѪ");
			   Result[1]=(char*)(_bstr_t)m_pRecordset->GetCollect("ǱѪ");
			}		
			if (*pItem[i]=="������"||*pItem[i]=="BIL"){
                if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{*pItem[i]+=InsertSpace(12);}
                *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("������");
                Result[2]=(char*)(_bstr_t)m_pRecordset->GetCollect("������");
			}
 					
 			if (*pItem[i]=="ͪ��"||*pItem[i]=="KET"){ 		
				 if(language=="zhongwen"){*pItem[i]+=InsertSpace(14);}else{ *pItem[i]+=InsertSpace(12);}
 			     *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("ͪ��");
				 Result[3]=(char*)(_bstr_t)m_pRecordset->GetCollect("ͪ��");
 			}
					
			if (*pItem[i]=="��ϸ��"||*pItem[i]=="ERY"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{ *pItem[i]+=InsertSpace(12);}
                *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("��ϸ��");
				Result[4]=(char*)(_bstr_t)m_pRecordset->GetCollect("��ϸ��");
			}
					
			if (*pItem[i]=="������"||*pItem[i]=="GLU"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{ *pItem[i]+=InsertSpace(12);}
				*pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("������");
				Result[5]=(char*)(_bstr_t)m_pRecordset->GetCollect("������");

			}
			if (*pItem[i]=="������"||*pItem[i]=="PRO"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{ *pItem[i]+=InsertSpace(12);}
				*pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("������");
				Result[6]=(char*)(_bstr_t)m_pRecordset->GetCollect("������");
			}
				
			if (*pItem[i]=="����"||*pItem[i]=="PH"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{*pItem[i]+=InsertSpace(13);}
                *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("����");
				Result[7]=(char*)(_bstr_t)m_pRecordset->GetCollect("����");
			}
					
			if (*pItem[i]=="��������"||*pItem[i]=="NIT"){
                if(language=="zhongwen"){*pItem[i]+=InsertSpace(10);}else{*pItem[i]+=InsertSpace(12);}
                 *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("��������");
				 Result[8]=(char*)(_bstr_t)m_pRecordset->GetCollect("��������");
			}
				
			if (*pItem[i]=="����"||*pItem[i]=="SG"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(14);}else{*pItem[i]+=InsertSpace(13);}
				*pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("����");
				Result[9]=(char*)(_bstr_t)m_pRecordset->GetCollect("����");
			}
					
			if (*pItem[i]=="ά����C"||*pItem[i]=="VC"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(11);}else{ *pItem[i]+=InsertSpace(13);}
	            *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("ά����C");
				Result[10]=(char*)(_bstr_t)m_pRecordset->GetCollect("ά����C");
			}
 		}
	}
			//	m_pRecordset->MoveNext();
	
		
	//	}

			UpdateData(false);	
		}
	}
	//	GetDlgItem(IDC_PRINT)->EnableWindow(true);   //��ӡ��ť����ʹ��
    	*pResult = 0;
}

void CUrinalysisDlg::OnShowinformation()     //��ʾ��Ϣ�����İ�����Ӧ
{
	// TODO: Add your control notification handler code here
	CShowInformation dlg;
 
	if (dlg.DoModal()==IDOK)
	{
    GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
	InitItemField();

	}	
}

void CUrinalysisDlg::ClearItemContent()   //�����ʾ��Ŀstatic������������ֵ
{
	
	for (int j=0;j<11;j++)
	{   
		*pItem[j]="";
	}
}

CString CUrinalysisDlg::InitItemField()             //��ʮһ��������䵽�ı��ؼ���
{
char dataLanguage[100];
CRegKey Rek;
DWORD cbA=100;
Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
////////////////////��ȡ����
Rek.QueryValue(dataLanguage,"ItemLanguage",&cbA);
CString dataLanguage2(dataLanguage);  //��ȡע���������ֵ���ַ���
////////////////////////��ȡ��Щ��Ŀ��ѡ�еı�־λ
CStringArray item;

for (int i=0;i<11;i++)   //����Ŀ��ѡ��ѡ�б�־λд��item������
{
	CString s;
	char value[20]; 
	s.Format("%d",i+1);
	Rek.QueryValue(value,"Item"+s,&cbA);
	item.Add(value);         //���ʮһ���־λ�����飬ÿ��Ԫ��ֵΪ"T"��"F"
}

// 	/////////��������Ŀ���Ƽ���CStringArray������
LoadContent(item,dataLanguage2);    //��һ��������ʾ��ѡ���Ƿ�ѡ�б�־λ���ڶ����������Ա�־λ,������Ҫʵ��ΪpatientResult���Ҫ��ʾ����Ŀ
ClearItemContent();                 //�����ʾ��Ŀstatic������������ֵ
Rek.Close();
for (int j=0;j<patientResult.GetSize();j++)
{   
	*pItem[j]=patientResult.GetAt(j);
//	UpdateData(FALSE);
}
UpdateData(FALSE);
return dataLanguage2;
}

CString CUrinalysisDlg::InsertSpace(int num)   //����ո�ĺ���
{
  CString str;
  for (int i=0;i<num;i++)
  str+=" ";
  return str;
}

void CUrinalysisDlg::OnSaveinformation()       //��ģ����ʧЧ
{
	LoadSaveModule();

}

void CUrinalysisDlg::OnDestroy()              //
{
	CDialog::OnDestroy();
	ExitConnect();
    AnimateWindow(GetSafeHwnd(),500,AW_HIDE|AW_BLEND);   //���ʹ�ý���Ч���˳�
}

void CUrinalysisDlg::OnTimer(UINT nIDEvent) 
{
     if (nIDEvent==1)                         //ˢ��ʱ��
     {
		 CTime time=CTime::GetCurrentTime();
		 CString m_timeShow=time.Format("������: %Y��%m��%d��  %H:%M:%S");
		 
		 GetDlgItem(IDC_TIMESHOW)->SetWindowText(m_timeShow);
	     UpdateWindow();
     }
	 if (nIDEvent==2){                     //�豸���ӵȺ���
		 KillTimer(2);
		 if (!isOpen)
		 { 
			 
			 if (m_Comm.GetPortOpen())
			 {
				 m_Comm.SetPortOpen(FALSE);
			 }
			 GetDlgItem(IDC_STATIC1)->SetWindowText("�豸����ʧ��");
             MessageBox("�豸����ʧ��","HIGHTOP����", MB_ICONSTOP );        
			 GetDlgItem(IDC_CONNECT)->EnableWindow(TRUE);
	
		 }

	 }

	CDialog::OnTimer(nIDEvent);
}

void CUrinalysisDlg::OnPrint()          //��ӡ������ť��Ӧ
{ 	
//	CPrintPreview dlg;
//	if (dlg.DoModal()==IDOK)
//	{
     //   LoadSaveModule();
//	}else{
    // 	LoadPrintModule();
//	}
    UpdateData(TRUE);
	            Result[11]=m_name;
				CString sex;
				if (m_sex==0)  {sex="��";}
				if(m_sex==1){sex="Ů";}
				Result[12]=sex;
                Result[13]=m_age;
				Result[14]=m_caseNumbe;
				Result[15]=m_sampleType;
				Result[17]=m_doctorName;
				Result[18]=m_department;
				Result[20]=m_conclusion;

	_variant_t RecordsAffected; 
	CString strSQL; 
	strSQL="UPDATE  tb_Sample1 SET ��������='"+Result[11]+"',����='"+Result[13]+"',�Ա�='"+Result[12]+"',������='"+Result[14]+"',��������='"+Result[15]+"',ҽ��='"+Result[17]+"',����='"+Result[18]+"',���='"+Result[20]+"' WHERE ������='"+Result[16]+"' and ����ʱ��='"+Result[19]+"' and ʱ��='"+Result[21]+"'";
	m_pConnection->Execute((_bstr_t)strSQL,&RecordsAffected,adCmdText); 
	for(int i=0; i<m_grid.GetItemCount(); i++){
		if (m_grid.GetItemText(i,0)==Result[16]&&m_grid.GetItemText(i,1)==Result[19]&&m_grid.GetItemText(i,2)==Result[21])
		{
		
		m_grid.SetItemText(i,3,Result[11]);
		m_grid.SetItemText(i,4,Result[12]);
		m_grid.SetItemText(i,5,Result[13]);
		m_grid.SetItemText(i,6,Result[14]);
		}
	}
	LoadPrintModule();            //���ɱ�������

}



void CUrinalysisDlg::LoadPrintModule()            //���ɱ�������
 { 
    _Application app;	
	Documents doc;
	CComVariant a (_T("")),b(false),c(0),d(true),aa(1),bb(20);
	VARIANT varOptional,varTable;//��һ�������������ڶ���������ӱ��Ĳ���
	_Document doc1;
    _Font font;
    Paragraphs  pParas;
	Selection sele;
    _ParagraphFormat pf;
    Range wordRange;
	Shapes shp;
	COleVariant vOpt((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
 	VariantInit(&varOptional);
	varOptional.vt=VT_ERROR;
	varOptional.scode=DISP_E_PARAMNOTFOUND;
	VariantInit(&varTable);
	varTable.vt=VT_I4;//������Ӧ����������Ϊlong
	varTable.lVal=0;
	//��ʼ������
	app.CreateDispatch("word.Application");  //����Ӧ��ʵ��
	doc.AttachDispatch(app.GetDocuments());   //����ĵ�����ʵ��
	doc1.AttachDispatch(doc.Add(&a,&b,&c,&d));  //����ĵ�ʵ��
	
	sele.AttachDispatch(app.GetSelection());  //�������ʵ��
	pf=sele.GetParagraphFormat();          //��ʽ��ʵ��
	pParas=sele.GetParagraphs();           //����ʵ��
	font=sele.GetFont();
	shp=doc1.GetShapes();
 	wordRange==doc1.Range(&varOptional,&varOptional);   //��ȡwordRange
 	OnInitADOConn();         //�������ݿ�
	CTime time=CTime::GetCurrentTime();
	CString m_timeShow=time.Format("%H:%M:%S");
 	_bstr_t bstrSQL;
 
	COleVariant vStyle0((long)-2);
 	pf.SetStyle(&vStyle0);
	pParas.SetAlignment(1);   //���־���
	font.SetBold(true);      //����Ӵ�
////////////////////////////////////�����б��϶˵Ĺ̶�����
	
	//////��ע�������ȡ���浥����
	char caption[100];
    CRegKey Rek;
	DWORD cbA=200;
	Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
	Rek.QueryValue(caption,"ReportCaption",&cbA);
    sele.TypeText(caption);

    sele.TypeParagraph();          //����  
	sele.TypeText("������"+Result[11]+"\t�Ա�"+Result[12]+"\t���䣺"+Result[13]+"\t�����ţ�"+Result[14]+"\t�������ͣ�"+Result[15]);
	sele.TypeParagraph();          //���� 
	sele.TypeText("�����ţ�"+Result[16]+"\t\t�ͼ�ҽ����"+Result[17]+"\t\t���ң�"+Result[18]);
	sele.TypeParagraph();          //����
    sele.TypeText("����ʱ�䣺"+Result[19]+"   "+Result[21]+"     ��ϣ�"+Result[20]);
	sele.TypeParagraph();          //����
	sele.TypeParagraph();          //����
	sele.TypeText("��ӡʱ�䣺"+m_timeShow);
	font.SetBold(true);
	sele.TypeText("\t\t\t\t���ֻ�Ը���������");
		sele.TypeParagraph(); 
			sele.TypeParagraph(); 
	shp.AddLine(20,187,580,187,vOpt);       //�Ӻ���
	shp.AddLine(20,188,580,188,vOpt);

	shp.AddLine(20,600,580,600,vOpt);
	shp.AddLine(20,602,580,602,vOpt);


	////////////////////////////////////////////�ȴ�word������������
	 app.SetVisible(true);    //��word  
	

 	//�������Ĺ���
	bstrSQL = "select*from tb_Print";
	m_pRecordset.CreateInstance(__uuidof(Recordset));
	m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
	Fields* fields=NULL;
	long countl;
	BSTR bstr;
	m_pRecordset->get_Fields(&fields);
	countl = fields->Count;
	_variant_t sField[10];
    wordRange=sele.GetRange();
	Tables wordTables=sele.GetTables();
	Table wordTable=wordTables.Add(wordRange,14,countl+1,&varOptional,&varTable);

	wordTable.SetAllowAutoFit(false);         //�������Ӧ�п� 
	Columns tblColumns;
	tblColumns=wordTable.GetColumns();
	int colCount = tblColumns.GetCount();   //�����������ñ���п�
   for(int i=1;i<=colCount;i++)
   {
   	Column   wordColumn = tblColumns.Item(i);
    switch(i)
    {
  
	 case 1:                //��1��
     wordColumn.SetWidth(70.0,0);
     break;
	 case 2:                //��2��
     wordColumn.SetWidth(70.0,0);
     break;
	 case 3:                //��3��
     wordColumn.SetWidth(90.0,0);
     break;
	 case 4:                //��4��
     wordColumn.SetWidth(140.0,0);
     break;
	 case 5:                //��5��
     wordColumn.SetWidth(100.0,0);
     break;
 
    }
   }

	Borders tblBorders;//�߿����
	tblBorders=wordTable.GetBorders();
	tblBorders.SetEnable(0);//�ޱ߿�
    Cell tblCell;
	
/////////////////////////////////////�������
for(long num=1;num<=countl+1;num++)   //��Ҫ������һ�� ����+1
 {
	
		
 if (num<=3)
 {	
   tblCell=wordTable.Cell(1,num);
   
   wordRange=tblCell.GetRange();
   fields->Item[(long)(num-1)]->get_Name(&bstr);
   sField[num-1] = (_variant_t)bstr;
  // font.SetBold(true); 
   wordRange.InsertAfter((char*)(_bstr_t)bstr);
  
   continue;
 }
 	
 if (num==4)   //�ֹ�������һ��
 {	
 	tblCell=wordTable.Cell(1,num);
  	wordRange=tblCell.GetRange();
//	font.SetBold(true); 
  	wordRange.InsertAfter("    ���    ");
    continue;

}
if (num>4)         
{	
		tblCell=wordTable.Cell(1,num);
		wordRange=tblCell.GetRange();
		fields->Item[(long)(num-2)]->get_Name(&bstr);
		sField[num-2] = (_variant_t)bstr;
	//	font.SetBold(true); 
		wordRange.InsertAfter((char*)(_bstr_t)bstr);
 }
}

///////////////////////////���ݿ��ȡ��������䵽�����
long row=2;
while(!m_pRecordset->adoEOF)
{

for(long num=1;num<=countl+1;num++)
{
		
	tblCell=wordTable.Cell(row,num);
	wordRange=tblCell.GetRange();
	if (num<4)
	{
		wordRange.InsertAfter((char*)(_bstr_t)m_pRecordset->GetCollect(sField[num-1]));
		continue;
	}
	if (num==4)        //������������
	{
		wordRange.InsertAfter(Result[row-2]);
		
		
		continue;
	}
	if (num>4)
	{
		wordRange.InsertAfter((char*)(_bstr_t)m_pRecordset->GetCollect(sField[num-2]));
		continue;
	}

 }
	
 	 		m_pRecordset->MoveNext();
 			row++;
  }


    doc1.PrintPreview();
	wordTable.ReleaseDispatch();
	wordTables.ReleaseDispatch();
	sele.ReleaseDispatch();
	doc.ReleaseDispatch();
	doc1.ReleaseDispatch();
	app.ReleaseDispatch();
	
}

void CUrinalysisDlg::LoadSaveModule()     //����ʵ�ֱ���ı���ģ��
{
	_variant_t RecordsAffected; 
   	UpdateData(true);
	if(!Upload)
	{
		AfxMessageBox("����δ�ϴ�����Ч���ݣ��޷�����");
		return;
	}
	CString sex,time;
    if (m_sex==0)  {sex="��";}else{sex="Ů";}
    //time = m_time.Format("%Y-%m-%d");
	time=m_time;
	CString strSQL; 
	strSQL.Format("INSERT INTO tb_Sample1(��������,����,�Ա�,������,��������,������,����ʱ��,ҽ��,����,���,��ԭ,ǱѪ,������,ͪ��,��ϸ��,������,������,����,��������,����,ά����C)VALUES ('%s','%s','%s','%s','%s', '%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')", m_name,m_age,sex,m_caseNumbe,m_sampleType,m_sampleNumber,time,m_doctorName,m_department,
		m_conclusion,Result[0],Result[1],Result[2],Result[3],Result[4],Result[5],Result[6],Result[7],Result[8],Result[9],Result[10]); 
	m_pConnection->Execute((_bstr_t)strSQL,&RecordsAffected,adCmdText); //��Ӽ�¼
	////�ؼ���ֵ��Ϊ��ʼ��״̬
    
	
	
	_bstr_t bstrSQL = "select*from tb_Sample1 where ��������='"+m_name+"' and ������='"+m_sampleNumber+"' and ����ʱ��='"+time+"'";
	m_pRecordset.CreateInstance(__uuidof(Recordset));
	m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
	while(!m_pRecordset->adoEOF)
	{
			
		
	CString s[7];
		s[0]=(char*)(_bstr_t)m_pRecordset->GetCollect("������");
		s[1]=(char*)(_bstr_t)m_pRecordset->GetCollect("����ʱ��");
        s[2]=(char*)(_bstr_t)m_pRecordset->GetCollect("ʱ��");
		s[3]=(char*)(_bstr_t)m_pRecordset->GetCollect("��������");
		s[4]=(char*)(_bstr_t)m_pRecordset->GetCollect("�Ա�");
		s[5]=(char*)(_bstr_t)m_pRecordset->GetCollect("����");
		s[6]=(char*)(_bstr_t)m_pRecordset->GetCollect("������");
		m_grid.AddItem(s[0],s[1],s[2],s[3],s[4],s[5],s[6]);


		m_pRecordset->MoveNext();
	}
	
	m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion="";
	m_sex=0;
//	m_time=CTime::GetCurrentTime();
	m_time="";
	m_time2="";
	UpdateData(false);	
	MessageBox("����ɹ�");

}

void CUrinalysisDlg::OnDatadictionary()           //�����ֵ���Ӧ����
{

	// TODO: Add your control notification handler code here
	
CDictionarySet dlg;
if (dlg.DoModal()==IDOK)
{   
	WinExec("Urinalysis.EXE",SW_SHOWMAXIMIZED);   //���´򿪳���
	SkinH_Detach();                //ж��Ƥ��
	exit(0);                      //�رյ�ǰ����
}

}



void CUrinalysisDlg::InitControlContent()            //�������ֵ�����ȡ��������ռ���
{
	CString s[4]={"��������","�ͼ�ҽ��","����","���"};
    for (int i=0;i<4;i++)
    {
		_bstr_t bstrSQL;
		bstrSQL = "select VALUE1 from tb_Datadictionary where DATATYPE1='"+s[i]+"'";
		m_pRecordset.CreateInstance(__uuidof(Recordset));
		m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
		while(!m_pRecordset->adoEOF)
		{   
		if(i==0)((CComboBox*)GetDlgItem(IDC_SAMPLETYPE))->AddString((char*)(_bstr_t)m_pRecordset->GetCollect("VALUE1"));
		if(i==1)((CComboBox*)GetDlgItem(IDC_DOCTORNAME))->AddString((char*)(_bstr_t)m_pRecordset->GetCollect("VALUE1"));
		if(i==2)((CComboBox*)GetDlgItem(IDC_DEPARTMENT))->AddString((char*)(_bstr_t)m_pRecordset->GetCollect("VALUE1"));	
		if(i==3)((CComboBox*)GetDlgItem(IDC_CONCLUSION))->AddString((char*)(_bstr_t)m_pRecordset->GetCollect("VALUE1"));	
		m_pRecordset->MoveNext();
		}

    }


}

void CUrinalysisDlg::OnReportset()          //���浥���ð�ť
{
	// TODO: Add your control notification handler code here

	CReportSet dlg;

	if (dlg.DoModal()==IDOK)
	{   
		CRegKey Rek;
       	Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");  //H��ʧ��
	    Rek.SetValue(Caption,"ReportCaption");           //����ȫ�ֱ���caption��������������
		Rek.Close();
	}
	
}

void CUrinalysisDlg::RegisterActiveX(CString s)        //ע��ACTIVEX�ؼ�
{
		CString strFileName=g_curPath+s;
		if (strFileName.IsEmpty())
			return;
		
		//װ��ActiveX�ؼ�  
		HINSTANCE hInstance = LoadLibrary(strFileName); 
		if (hInstance == NULL)  
		{     
			AfxMessageBox("��������Dll/OCX�ļ�!");     
			return; 
		} 
		//ȡ��ע�ắ��DllRegisterServer��ַ 
		FARPROC lpFunc;  
		lpFunc = GetProcAddress(hInstance,_T("DllRegisterServer"));  
		//����ע�ắ��DllRegisterServer 
		if(lpFunc!=NULL)  
		{       
			if(FAILED((*lpFunc)()))   
			{       
				//AfxMessageBox("����DllRegisterServer ʧ��!");  
				FreeLibrary(hInstance);   //�ͷ���Դ       
				return;     
			}     
			//MessageBox(s+"ע��ɹ�");  
		}
		else  
		{
		//	AfxMessageBox("����DllRegisterServerʧ��!");
		}

}

void CUrinalysisDlg::OnClose()           
{
	// TODO: Add your message handler code here and/or call default
	CDialog::OnClose();
}

void CUrinalysisDlg::OnCancel() 
{
	// TODO: Add extra cleanup here
	AnimateWindow(GetSafeHwnd(),500,AW_HIDE|AW_BLEND);
	CDialog::OnCancel();
}

void CUrinalysisDlg::OnButton1()     //�����������Ӧ
{
	// TODO: Add your control notification handler code here
   GetDlgItem(IDC_STATIC1)->SetWindowText("��ӭ�����ʺ���");
   CAboutDlg dlg;
   dlg.DoModal();	
   
}

void CAboutDlg::OnLinkbutton()      //���ʺ��Ƶ���Ӧ
{
	// TODO: Add your control notification handler code here
	
	ShellExecute(NULL, 
		"open", 
		"http://www.hightopqd.com/", 
		NULL, 
		NULL, 
		SW_SHOWNORMAL);	
}

BEGIN_EVENTSINK_MAP(CUrinalysisDlg, CDialog)               //������Ӧӳ�䣨�Զ����ɣ�
    //{{AFX_EVENTSINK_MAP(CUrinalysisDlg)
	ON_EVENT(CUrinalysisDlg, IDC_MSCOMM1, 1 /* OnComm */, OnOnCommMscomm1, VTS_NONE)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

void CUrinalysisDlg::OnOnCommMscomm1()            //���ڴ�������
{   
	int k, nEvent;
	nEvent = m_Comm.GetCommEvent();
	switch(nEvent)
	{
	case 2:  //�յ�����RTHresshold���ַ�
		k = m_Comm.GetInBufferCount(); //���յ����ַ���Ŀ
		if(k==5)          //5���ַ�������
		{
			portReceiveMessage=receivePortData();    //ȡ�ش��ڽ��յ�������
			if (portReceiveMessage=="READY")
			{
				isOpen=TRUE;            //�����ӳɹ��ı�־λ����
                GetDlgItem(IDC_CONNECT)->EnableWindow(TRUE);
				GetDlgItem(IDC_CONNECT)->SetWindowText("�Ͽ�����");
				GetDlgItem(IDC_POSTBACK)->EnableWindow(TRUE);
				GetDlgItem(IDC_SERACH)->EnableWindow(FALSE);
				GetDlgItem(IDC_DELETEDATA)->EnableWindow(FALSE);
				GetDlgItem(IDC_CLEAR)->EnableWindow(FALSE);
                monitorCount=0;                 //���ο������ں��յ��Ľ�������ļ�����
				m_grid.DeleteAllItems();
				InitItemField();  //��ս�����ڵ�����
   				m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //���ؼ����������
				m_sex=-1;    //���ؼ����������
				UpdateData(FALSE);
			    GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
			    GetDlgItem(IDC_STATIC1)->SetWindowText("�豸���ӳɹ�");
     			MessageBox("���ӳɹ�","HIGHTOP����", MB_ICONINFORMATION );
			    m_Comm.SetRThreshold(32); //32���ַ���������
		
			}

		}   
		else if (k==32)      //32���ַ�����
		{	
               //���յ�����������
			portReceiveMessage=receivePortData();
			///////////////////////////////////////////////////ʵʱ���
            if (portReceiveMessage.Left(3)=="RES"&&portReceiveMessage.Right(3)=="END")   //��ʽ���ϵ�ָ��
            {
		    CString tempResult[22];
			CString exceptionFlag="0";      //�ж�ʮһ�������쳣�ģ���Ϊ1������쳣
			for(int a=0;a<22;a++)  tempResult[a]="";    
		  //  GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
			CString tempSampleNumber,tempTime,tempTime2; ///��ʱ�����洢������ ���� ʱ��
			 for (int i=3;i<6;i++)  tempSampleNumber+=portReceiveMessage.GetAt(i); 
			 tempResult[16]=tempSampleNumber;     //�걾��
			/////////////////////////////////////////
			 tempTime+="20";           //ȡ����
			 for (int j=6;j<12;j++)    
			 { 
			   tempTime+=portReceiveMessage.GetAt(j);  
			   if (j%2==1&&j<11) tempTime+="-";   
			 }
			 tempResult[19]=tempTime;
            ////////////////////////////////////////
			 for (int k=12;k<18;k++)                     //ȡʱ��
			 {
				 tempTime2+=portReceiveMessage.GetAt(k);
				 if (k%2==1&&k<17) tempTime2+=":";
			 }
             tempResult[21]=tempTime2;
			 ///////////////////////////////////////
			 for (int m=18;m<29;m++)                     //�ж�ʮһ����Ŀ�Ƿ�����쳣�ֻҪ��һ�����Ϊ1
			 {   
				 if (portReceiveMessage.GetAt(m)!='0')
				 {
					 exceptionFlag="1";
                   
				 }
				 tempResult[m-18]=decideResult(m-18,portReceiveMessage.GetAt(m));   //ȡʮһ������
	
			 }
	
             _variant_t RecordsAffected; 
			  CString strSQL; 
		//	CString	strSQL2 = "delete from tb_Sample1 where ������='"+tempResult[16]+"' and ����ʱ��='"+tempResult[19]+"' and ʱ��='"+tempResult[21]+"'";
        //    m_pConnection->Execute((_bstr_t)strSQL2,&RecordsAffected,adCmdText); //Ԥɾ�������ظ��ļ�¼			  
			 strSQL.Format("INSERT INTO tb_Sample1(��������,����,�Ա�,������,��������,������,����ʱ��,ҽ��,����,���,ʱ��,��ԭ,ǱѪ,������,ͪ��,��ϸ��,������,������,����,��������,����,ά����C,�Ƿ��쳣)VALUES ('%s','%s','%s','%s','%s', '%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')",tempResult[11],tempResult[13],tempResult[12],tempResult[14],tempResult[15],tempResult[16],tempResult[19],tempResult[17],tempResult[18],
			  tempResult[20],tempResult[21],tempResult[0],tempResult[1],tempResult[2],tempResult[3],tempResult[4],tempResult[5],tempResult[6],tempResult[7],tempResult[8],tempResult[9],tempResult[10],exceptionFlag); 
			 m_pConnection->Execute((_bstr_t)strSQL,&RecordsAffected,adCmdText); //��Ӽ�¼
             //if(exceptionFlag="1"){colorArray.Add(1);}else{colorArray.Add(0);}

			 ////�ؼ���ֵ��Ϊ��ʼ��״̬	
		CString s[7];
		s[0]=tempResult[16];
		s[1]=tempResult[19];
        s[2]=tempResult[21];
	    s[3]="";
		s[4]="";
		s[5]="";
		s[6]="";
	    m_grid.AddItem(s[0],s[1],s[2],s[3],s[4],s[5],s[6],"");
		if(exceptionFlag=="1")     //�����쳣��
		{
			m_grid.SetItemColor(monitorCount,0,RGB(0,0,0),RGB(255,0,0));   //�����쳣��Ľ���ź�ɫ
			m_grid.SetItemColor(monitorCount,1,RGB(0,0,0),RGB(255,0,0));   //�����쳣��Ľ���ź�ɫ
		//	m_grid.SetItemColor(monitorCount,2,RGB(0,0,0),RGB(255,0,0));   //�����쳣��Ľ���ź�ɫ

		}
		if(exceptionFlag=="0")    //�������쳣��
		{
			m_grid.SetItemColor(monitorCount,0,RGB(0,0,0),RGB(0,255,0));   //�������쳣��Ľ������ɫ
			m_grid.SetItemColor(monitorCount,1,RGB(0,0,0),RGB(0,255,0));   //�������쳣��Ľ������ɫ
		//	m_grid.SetItemColor(monitorCount,2,RGB(0,0,0),RGB(0,255,0));   //�������쳣��Ľ������ɫ

		}
      CString count1;
	  CString count2;
	  count1.Format("%d",monitorCount+1);       //�ܽ��յ�������������
      count2.Format("%d",postbackCount+1);      //�ش�����������
	  if(isPostback)         //�Ƿ��ڻش�״̬
	  {
		   GetDlgItem(IDC_STATIC1)->SetWindowText("�ش���"+count2+"������,��"+postbackNumber+"��");  
	       postbackCount++;    //�ش�����
	  }
	  else{
		    GetDlgItem(IDC_STATIC1)->SetWindowText("���յ�"+count1+"������");
	  }
	          //��Ļ���½���Ϣ
	     monitorCount++;   //�������Ӻ��ܹ��յ��ı걾����

  }
///////////////////////////////////////////////////////////�����ش�ͷ��ָ��
	  if (portReceiveMessage.Left(3)=="FIR"&&portReceiveMessage.Right(3)=="END")  {

		  GetDlgItem(IDC_LIST1)->EnableWindow(FALSE);    //���б�ؼ���������ֹ�����������ݶ�ʧ
		  isPostback=TRUE;    //�ش�ģʽ��־λ
	      postbackCount=0;    //��ͳ�ƻش����ı�����Ϊ0
	      CString temp1[3];
	//	  int temp2[3];
		   //��¼���λش�������Ŀ��
		  for(int i=3;i<6;i++){temp1[i-3]=portReceiveMessage.GetAt(i); /* temp2[i-3]=atoi(temp1[i-3]);*/}
		  if(temp1[0]=="0"&&temp1[1]=="0")
		  {
			  postbackNumber=temp1[2];
		  }
		  else if(temp1[0]=="0"&&temp1[1]!="0")
		  {   
			 
			  postbackNumber=temp1[1]+temp1[2];	
		  }
		  else
		  {
		      postbackNumber=temp1[0]+temp1[1]+temp1[2];
		  }	
	  }
////////////////////////////////////////////////////////�����ش�����ָ��
 if (portReceiveMessage.Left(3)=="LAS"&&portReceiveMessage.Right(3)=="END")  {
	 
	  isPostback=FALSE;
	  GetDlgItem(IDC_STATIC1)->SetWindowText("���ݻش����");
	   if (postbackCount==0)    //û�лش�����
	   {
		  MessageBox("û��������ش�","HIGHTOP����", MB_ICONINFORMATION );

	   }
	  if (postbackCount>0){
		  MessageBox("ȫ�����ݻش����","HIGHTOP����", MB_ICONINFORMATION );
	   }
        GetDlgItem(IDC_LIST1)->EnableWindow(TRUE);  //���������б��⿪
	    postbackCount=0;    //��ͳ�ƻش����ı�����Ϊ0
	  }
	}
        else{
		m_Comm.GetInput();     //������뻺����
		}
		
	}    
	
}

void CUrinalysisDlg::InitSerialPort()         //��ʼ�����ڣ�������connectָ��
{	//char data[100];
    CRegKey Rek;
    DWORD port=200;
	DWORD rate=200;
	CString bandRate;
    Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
	Rek.QueryValue(port,"Port");
	Rek.QueryValue(rate,"BandRate");
    bandRate.Format("%d",rate);
	m_Comm.SetCommPort(port); //ѡ��COM1
	m_Comm.SetInBufferSize(1024); //�������뻺�����Ĵ�С��Bytes
	m_Comm.SetOutBufferSize(512); //�������뻺�����Ĵ�С��Bytes
	try
	{

	if (!m_Comm.GetPortOpen())
	{
		m_Comm.SetPortOpen(TRUE);
	}
	m_Comm.SetInputMode(1); //�������뷽ʽΪ�����Ʒ�ʽ
	m_Comm.SetSettings(bandRate+",n,8,1"); //���ò����ʵȲ���
	m_Comm.SetRThreshold(5); //Ϊ1��ʾ��һ���ַ�����һ���¼���������
//	m_Comm.SetInputLen(0);
	sendPortData("CONNECT");   //����������������
	}
	catch (_com_error *e)
	{
		AfxMessageBox(e->ErrorMessage());
		
	}

}

void CUrinalysisDlg::OnConnect()        //���Ӱ�ť��Ӧ
{ 
	if (!isOpen)    //isOpen�������ӱ�־λ��δ����
	{
		CConnectDlg dlg;		
		switch(dlg.DoModal()){
		case IDOK:
            InitSerialPort();     //��ʼ�����ڣ�����connectָ��
			GetDlgItem(IDC_CONNECT)->EnableWindow(FALSE);
			SetTimer(2,1500,NULL);
		}
	}
    else{           //�Ѿ�������
		sendPortData("DISCONNECT");   //�Ͽ���������λ�����ͶϿ�ָ��
		if(m_Comm.GetPortOpen()) //�򿪴���   
		{
		  m_Comm.SetPortOpen(FALSE);

		}
 		isOpen=FALSE;    
	    GetDlgItem(IDC_STATIC1)->SetWindowText("�Ͽ��豸����");
		GetDlgItem(IDC_CONNECT)->SetWindowText("�����豸");
		GetDlgItem(IDC_SERACH)->EnableWindow(TRUE);
		GetDlgItem(IDC_DELETEDATA)->EnableWindow(TRUE);
		GetDlgItem(IDC_CLEAR)->EnableWindow(TRUE);
		GetDlgItem(IDC_POSTBACK)->EnableWindow(FALSE);   //�ش���ť������
		MessageBox("���ӶϿ�","HIGHTOP����", MB_ICONEXCLAMATION );
	}
	// TODO: Add your control notification handler code here
}


/////////////////////////////�򴮿ڷ�������
void CUrinalysisDlg::sendPortData(CString data)
{
	char TxData[100];
    int Count = data.GetLength();
	for(int i = 0; i < Count; i++)
		TxData[i] = data.GetAt(i);
    CByteArray array;     
    array.RemoveAll();
    array.SetSize(Count);
    for(i=0;i<Count;i++)
		array.SetAt(i, TxData[i]);
    m_Comm.SetOutput(COleVariant(array)); // ��������	
}



/////////////////////////////�Ӵ��ڽ�������
CString CUrinalysisDlg::receivePortData()
{
	VARIANT variant_inp;
    COleSafeArray safearray_inp;   //COleSafeArray�������ڴ����������ͺ�ά�����������**********
    long i = 0;
    int len;
	BYTE rxdata[1000];    
	CString receive;
	//��ȡ����������
	variant_inp = m_Comm.GetInput();
	//��VARIANT�ͱ���ֵ����ColeSafeArray���ͱ���
	safearray_inp = variant_inp;
	len = safearray_inp.GetOneDimSize(); //����һ��һά��COleSafeArray�����е�Ԫ�ظ��� 
	for(i = 0; i < len; i++)
	{
		safearray_inp.GetElement(&i, &rxdata[i]);  //�����ݱ��浽�ַ�������
	}
	rxdata[i] = '\0'; 
	receive= rxdata;  	
	return receive;
}
	




CString CUrinalysisDlg::decideResult(int location, CString result)   //����������жϺ���
{   CString res="";
    switch(location){

	case 0:       //��ԭ
		if (result=="0") {res="NORM 3.3 umol/l "; break;}
        if (result=="1") {res="1+   33 umol/l  ��"; break;}
		if (result=="2") {res="2+   66 umol/l  ��"; break;}
		if (result=="3") {res=">=3+ 131 umol/l ��"; break;}
		if (result=="9") {res="�����Ч"; break;}    // ��������Ч�������Ϊ9˵������������

	case 1:    //ǱѪ
		if (result=="0") {res="-    0 cells/ul";  break;}
		if (result=="1") {res="+-   10 cells/ul  ��"; break;}
		if (result=="2") {res="+    25 cells/ul  ��"; break;}
		if (result=="3") {res="2+   50 cells/ul  ��";  break;}
        if (result=="4") {res="3+   250 cells/ul ��"; break;}
		if (result=="9") {res="�����Ч"; break;}    // ��������Ч�������Ϊ9˵������������

        
	case 2:   //������
		if (result=="0") {res="-    0 umol/l"; break;}
        if (result=="1") {res="+    17 umol/l  ��"; break;}
		if (result=="2") {res="2+   50 umol/l  ��"; break;}
		if (result=="3") {res="3+   100 umol/l ��"; break;}
	    if (result=="9") {res="�����Ч"; break;}    // ��������Ч�������Ϊ9˵������������
	case 3:   //ͪ��
		if (result=="0") {res="-    0 mmol/l"; break;}
        if (result=="1") {res="+    1.5 mmol/l  ��"; break;}
		if (result=="2") {res="2+   4.0 umol/l  ��"; break;}
		if (result=="3") {res="3+   8.0 umol/l  ��"; break;}
	    if (result=="9") {res="�����Ч"; break;}    // ��������Ч�������Ϊ9˵������������
	case 4:   //��ϸ��
		if (result=="0") {res="-    0 cells/ul"; break;}
        if (result=="1") {res="+-   15 cells/ul ��"; break;}
		if (result=="2") {res="+    70 cells/ul ��"; break;}
		if (result=="3") {res="2+   125 cells/ul��"; break;}
        if (result=="4") {res="3+   500 cells/ul��"; break;}
        if (result=="9") {res="�����Ч"; break;}    // ��������Ч�������Ϊ9˵������������
    case 5:  //������
		if (result=="0") {res="-    0 mmol/l"; break;}
        if (result=="1") {res="+-   2.8 mmol/l ��"; break;}
		if (result=="2") {res="+    5.5 mmol/l ��"; break;}
		if (result=="3") {res="2+   14 mmol/l  ��"; break;}
        if (result=="4") {res="3+   28 mmol/l  ��"; break;}
		if (result=="5") {res="4+   55 mmol/l  ��"; break;}
	    if (result=="9") {res="�����Ч"; break;}    // ��������Ч�������Ϊ9˵������������
	case 6:  //������
		if (result=="0") {res="-    0 g/l"; break;}
        if (result=="1") {res="+-   0.15 g/l ��"; break;}
		if (result=="2") {res="+    0.3 g/l  ��"; break;}
		if (result=="3") {res="2+   1.0 g/l  ��"; break;}
		if (result=="3") {res=">=3+ 3 g/l    ��"; break;}
	    if (result=="9") {res="�����Ч"; break;}    // ��������Ч�������Ϊ9˵������������
	case 7:  //PH
       	if (result=="0") {res="5.0"; break;}
        if (result=="1") {res="6.0  ��"; break;}
		if (result=="2") {res="7.0  ��"; break;}
		if (result=="3") {res="8.0  ��"; break;}
		if (result=="4") {res="9.0  ��"; break;}
	    if (result=="9") {res="�����Ч"; break;}    // ��������Ч�������Ϊ9˵������������
	case 8:    //��������
		if (result=="0") {res="-    0 umol/l"; break;}
		if (result=="1") {res="+    18 umol/l  ��"; break;}
	    if (result=="9") {res="�����Ч"; break;}    // ��������Ч�������Ϊ9˵������������
		
	case 9:     //����
		if (result=="0") {res="<=1.005"; break;}
        if (result=="1") {res="1.010   ��";break;}
		if (result=="2") {res="1.015   ��"; break;}
		if (result=="3") {res="1.020   ��"; break;}
        if (result=="4") {res="1.025   ��"; break;}
		if (result=="5") {res="1.030   ��"; break;}
	    if (result=="9") {res="�����Ч"; break;}    // ��������Ч�������Ϊ9˵������������
	case 10:   //ά����c
		if (result=="0") {res="-    0 mmol/l"; break;}
        if (result=="1") {res="+-   0.6 mmol/l  ��"; break;}
		if (result=="2") {res="+    1.4 mmol/l  ��"; break;}
		if (result=="3") {res="2+   2.8 mmol/l  ��"; break;}
        if (result=="4") {res="3+   5.6 mmol/l  ��"; break;}
	    if (result=="9") {res="�����Ч"; break;}    // ��������Ч�������Ϊ9˵������������
	}
		return res;

}

void CUrinalysisDlg::OnRefreshdata()             //�������ݰ�ť
{
                UpdateData(TRUE);
	            Result[11]=m_name;
				CString sex;
				if (m_sex==0)  {sex="��";}
				if(m_sex==1){sex="Ů";}
				Result[12]=sex;
                Result[13]=m_age;
				Result[14]=m_caseNumbe;
				Result[15]=m_sampleType;
				Result[17]=m_doctorName;
				Result[18]=m_department;
				Result[20]=m_conclusion;

	_variant_t RecordsAffected; 
	CString strSQL; 
	strSQL="UPDATE  tb_Sample1 SET ��������='"+Result[11]+"',����='"+Result[13]+"',�Ա�='"+Result[12]+"',������='"+Result[14]+"',��������='"+Result[15]+"',ҽ��='"+Result[17]+"',����='"+Result[18]+"',���='"+Result[20]+"' WHERE ������='"+Result[16]+"' and ����ʱ��='"+Result[19]+"' and ʱ��='"+Result[21]+"'";
	m_pConnection->Execute((_bstr_t)strSQL,&RecordsAffected,adCmdText); 
	for(int i=0; i<m_grid.GetItemCount(); i++){
		if (m_grid.GetItemText(i,0)==Result[16]&&m_grid.GetItemText(i,1)==Result[19]&&m_grid.GetItemText(i,2)==Result[21])
		{
		
		m_grid.SetItemText(i,3,Result[11]);
		m_grid.SetItemText(i,4,Result[12]);
		m_grid.SetItemText(i,5,Result[13]);
		m_grid.SetItemText(i,6,Result[14]);
		}
	}
	 GetDlgItem(IDC_STATIC1)->SetWindowText("�ɹ���������");
     MessageBox("���³ɹ�","HIGHTOP����", MB_ICONINFORMATION );
}




void CAboutDlg::OnKillFocus(CWnd* pNewWnd) 
{
	CDialog::OnKillFocus(pNewWnd);
	
	// TODO: Add your message handler code here
	
}


void CUrinalysisDlg::OnSetfocusList1(NMHDR* pNMHDR, LRESULT* pResult) 
{

	*pResult = 0;
}


void CUrinalysisDlg::OnInsertitemList1(NMHDR* pNMHDR, LRESULT* pResult) 
{
	NM_LISTVIEW* pNMListView = (NM_LISTVIEW*)pNMHDR;
    
 
	*pResult = 0;
}

void CUrinalysisDlg::OnItemchangedList1(NMHDR* pNMHDR, LRESULT* pResult) 
{

	*pResult = 0;
}


void CUrinalysisDlg::OnPostback()     //�ش���ť��Ӧ
{
	CPostBack dlg;

	if(dlg.DoModal()==IDOK){

	//	if(isOpen){	
			GetDlgItem(IDC_STATIC1)->SetWindowText("�ȴ����ͣ����Ժ�....");
			CString request="REQ"+postbackTime.Right(6)+"END";
			sendPortData(request);        //��������ĳ�����ݵ�ָ��
	//		MessageBox(request);
	//	}else{
		//	MessageBox("δ�����豸�������Ӻ���");
	//	}

	}
}

void CAboutDlg::OnHelp()     //������ť����Ӧ
{

	ShellExecute(NULL,_T("open"),_T(".\\˵����.doc"),NULL,NULL,SW_SHOWNORMAL);
	// TODO: Add your control notification handler code here
	
}

void CUrinalysisDlg::OnOK() 
{
	// TODO: Add extra validation here
	
	//CDialog::OnOK();
}

BOOL CUrinalysisDlg::PreTranslateMessage(MSG* pMsg)   //����ESC �� ENTER��SPACE �����������������˳�����
{
if(pMsg->message==WM_KEYDOWN)
{  
	switch(pMsg->wParam)
	{
	case VK_RETURN:
	     return TRUE;
	case VK_ESCAPE:
		 return TRUE;
	case VK_SPACE:
		 return TRUE;
	}

}
	return CDialog::PreTranslateMessage(pMsg);
}



void CUrinalysisDlg::DateBaseAutoClear(int item)    //���ݿ�������������ָ�������󣬻��������ݿ�
{
	    int a=0;
		_bstr_t bstrSQL;
		 CString strSQL; 
		_variant_t RecordsAffected; 
		bstrSQL = "select ID from tb_Sample1 order by ID desc";	
		m_pRecordset.CreateInstance(__uuidof(Recordset));
		m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
		while(!m_pRecordset->adoEOF){
			a++;
			if(a>item)
			{	
			   long id=(long)m_pRecordset->GetCollect("ID");
	           strSQL.Format("delete from tb_Sample1 where ID=%d",id);
		       m_pConnection->Execute((_bstr_t)strSQL,&RecordsAffected,adCmdText); //��Ӽ�¼
			}
			m_pRecordset->MoveNext();	
		}
}

void CAboutDlg::OnOK()     
{
	// TODO: Add extra validation here
	GetParent()->Invalidate();   ///�ػ����Ի��򣬷�ֹ���ַ��������
	CDialog::OnOK();
}
