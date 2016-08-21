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
double m_x,m_y;    //H 控件放大倍数

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
CString Result[22];           ///全局变量，其他对话框可以访问，保存报告单上11个项目监测结果以及十项病人信息   索引0-10十一项检测结果，11-20病人信息
BOOL  Upload=FALSE;           // 全局变量，标志位，表示是否尿仪上传过数据
CString Caption;              // 全局变量，标题，打印报告的顶标题
BOOL isOpen=FALSE;            // 全局变量，是否连接尿仪
BOOL isPostback=FALSE;        //是否开启回传
CString postbackTime;         //选定的需回传时间
CUrinalysisDlg::CUrinalysisDlg(CWnd* pParent /*=NULL*/)
: CDialog(CUrinalysisDlg::IDD, pParent)
{ 
	//初始化各个变量
	
	m_themeChange=FALSE;   //是否重新刷新界面
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
	monitorCount=0;     	//开启串口后收到的数据个数
	postbackCount=0;      //回传的样本总个数
    postbackNumber="";    //回传的样本总个数(字符串类型)

	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	for (int i=0;i<21;i++)            //初始化结果栏
	{
		Result[i]="";         //无结果的时候，结果栏默认填充的值
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
	Sleep(1300);   //为了欢迎界面显示，整个进程悬停1600ms
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
		OnInitADOConn();    //数据库初始化
		RegisterActiveX("MSCOMM32.OCX");   //执行注册结果没效果
		RegisterActiveX("MSADODC.OCX");    //注册activeX控件
        RegisterActiveX("MSDATGRD.OCX");   //注册activeX控件
		RegisterActiveX("MSSTDFMT.DLL");   //注册activeX控件
	          ///初始化串口，设置参数并打开
		InitControlContent();         //从数据字典中提取数据填入空间中
		GetDlgItem(IDC_PRINT)->EnableWindow(false);   //打印按钮初始化不可用
		GetDlgItem(IDC_REFRESHDATA)->EnableWindow(false);   //更新按钮初始化不可用
		GetDlgItem(IDC_SAVEINFORMATION)->EnableWindow(false);  
		GetDlgItem(IDC_POSTBACK)->EnableWindow(FALSE);   //回传按钮不可用
	  	CString szCurPath(""); 
		GetModuleFileName(NULL,szCurPath.GetBuffer(MAX_PATH),MAX_PATH);
		szCurPath.ReleaseBuffer();  
 		g_curPath = szCurPath.Left(szCurPath.ReverseFind('\\') + 1);
	//	g_languagePath=g_curPath+ "Data\\Language.ini";
		m_Forsize.cx = GetSystemMetrics(SM_CXSCREEN);             //H设置程序宽度
		m_Forsize.cy = GetSystemMetrics(SM_CYSCREEN);          //H设置程序高度
// 		if(m_Forsize.cx<1024||m_Forsize.cy<738)
// 			MessageBox(g_LoadString("IDS_RESOLUTION_MESSAGE"));
		SizeWindow();                                         //H调整控件空间的大小
	    MoveWindow(0,0,m_Forsize.cx,m_Forsize.cy);  //H将程序调整为全屏（在任何分辨率下）
        ///////////////////解决组合框控件被放大后不能下拉的问题
		CRect rect;
		GetDlgItem(IDC_SAMPLETYPE)->GetWindowRect(&rect);
		ScreenToClient(&rect);//将控件大小转换为在对话框中的区域坐标
		GetDlgItem(IDC_SAMPLETYPE)->MoveWindow(rect.left,rect.top,rect.Width(),rect.bottom+300);//设置控件大小
	    GetDlgItem(IDC_DOCTORNAME)->GetWindowRect(&rect);
		ScreenToClient(&rect);//将控件大小转换为在对话框中的区域坐标
		GetDlgItem(IDC_DOCTORNAME)->MoveWindow(rect.left,rect.top,rect.Width(),rect.bottom+300);//设置控件大小
	    GetDlgItem(IDC_DEPARTMENT)->GetWindowRect(&rect);
		ScreenToClient(&rect);//将控件大小转换为在对话框中的区域坐标
		GetDlgItem(IDC_DEPARTMENT)->MoveWindow(rect.left,rect.top,rect.Width(),rect.bottom+300);//设置控件大小
	    GetDlgItem(IDC_CONCLUSION)->GetWindowRect(&rect);
		ScreenToClient(&rect);//将控件大小转换为在对话框中的区域坐标
		GetDlgItem(IDC_CONCLUSION)->MoveWindow(rect.left,rect.top,rect.Width(),rect.bottom+300);//设置控件大小

        SetTimer(1,1000,NULL);   //刷新时间的定时器
		SkinH_Attach();     //加载皮肤
	   // AnimateWindow(GetSafeHwnd(),1000,AW_HOR_POSITIVE);   //动画效果，讲窗口从中央展开
		UpdateData(FALSE);
/*		m_grid.SetExtendedStyle(LVS_EX_FLATSB             //list列表的初始化设置包括顶栏的内容
			|LVS_EX_FULLROWSELECT
			|LVS_EX_HEADERDRAGDROP
			|LVS_EX_ONECLICKACTIVATE
			|LVS_EX_GRIDLINES);
		m_grid.InsertColumn(0,"标本号",LVCFMT_LEFT,45,4);
		m_grid.InsertColumn(1,"检验日期",LVCFMT_LEFT,90,5);
        m_grid.InsertColumn(2,"时间",LVCFMT_LEFT,50,5);
		m_grid.InsertColumn(3,"病人姓名",LVCFMT_LEFT,60,0);  
		m_grid.InsertColumn(4,"性别",LVCFMT_LEFT,40,1);
		m_grid.InsertColumn(5,"年龄",LVCFMT_LEFT,40,2);
		m_grid.InsertColumn(6,"病历号",LVCFMT_LEFT,80,3);
		m_grid.SetHeadcolor(RGB(0,0,0),RGB(219,216,215));*/
	m_grid.SetExtendedStyle(LVS_EX_FULLROWSELECT|LVS_EX_CHECKBOXES );   //加上多选为了让颜色条更宽些
	m_grid.SetHeadcolor(RGB(0,0,0),RGB(215,233,245));
	char t_Text[200];
	sprintf(t_Text,"%s,45;%s,75;%s,60;%s,80;%s,40;%s,40;%s,80","标本号","检验日期","时间","病人姓名","性别","年龄","病历号");
    m_grid.SetHeadings(_T(t_Text));         //加载list控件的标题
	//将static控件赋给指针数组
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
		InitItemField();   //初始化信息对话框，将项目名称填入
		UpdateData(FALSE);
        DateBaseAutoClear(60000);   //数据库存储上限为60000条
 

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
// 		//获取文件长度
// 		DWORD m_filelen = m_file.GetLength(); 
// 		//在堆上分配空间
// 		HGLOBAL m_hglobal = GlobalAlloc(GMEM_MOVEABLE,m_filelen);
// 		LPVOID pvdata = NULL;
// 		//锁定堆空间,获取指向堆空间的指针
// 		pvdata = GlobalLock(m_hglobal);
// 		//将文件数据读区到堆中
// 		m_file.ReadHuge(pvdata,m_filelen);
// 		IStream*  m_stream;
// 		GlobalUnlock(m_hglobal);
// 		//在堆中创建流对象
// 		CreateStreamOnHGlobal(m_hglobal,TRUE,&m_stream);	
// 		//利用流加载图像
//  		OleLoadPicture(m_stream,m_filelen,TRUE,IID_IPicture,(LPVOID*)&m_picture);
//     	m_stream->Release();
// 		m_picture->get_Width(&m_width);// 宽高，MM_HIMETRIC 模式，单位是0.01毫米
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
			//////加载美化风格
			char data[100];
			CRegKey Rek;
			DWORD cbA=200;
			Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
			Rek.QueryValue(data,"Theme",&cbA);
			Rek.Close();
			skinppLoadSkin(_T(data));//AquaOS.ssk为项目下的皮肤文件	
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


void CUrinalysisDlg::SizeWindow()           //调整对话框及所有控件的大小，使软件全屏
{
	CRect m_ClientRect;
	GetClientRect(&m_ClientRect);
	m_x = m_Forsize.cx/(double)m_ClientRect.Width() ;//宽度方向放大倍数
	m_y = m_Forsize.cy/(double)m_ClientRect.Height();//高度方向放大倍数
	//调整控件的大小
	CWnd *pWnd = NULL; 
	pWnd = GetWindow(GW_CHILD);
	CRect rect;   //获取控件变化前大小
	while(pWnd)//判断是否为空，因为对话框创建时会调用此函数，而当时控件还未创建
	{
		pWnd->GetWindowRect(&rect);
		ScreenToClient(&rect);//将控件大小转换为在对话框中的区域坐标
		int width = rect.Width();
		int height = rect.Height();
		//其它控件位置和大小均变化
		rect.top = m_y * rect.top;
		rect.left = m_x * rect.left;
		rect.bottom = m_y * rect.bottom;
		rect.right = m_x * rect.right;
		pWnd->MoveWindow(&rect);//设置控件大小
		pWnd = pWnd->GetWindow(GW_HWNDNEXT);}
	
	
}
CString CUrinalysisDlg::g_LoadString(CString szID)   //从外部文件中取字符的函数
{
	CString szValue;
	DWORD dwSize = 1000;
	GetPrivateProfileString("String",szID,"Not found",
		szValue.GetBuffer(dwSize),dwSize,g_languagePath);
	szValue.ReleaseBuffer();
	szValue.Replace("\\n","\n");	//替换回换行符号
	return szValue;
}

HBRUSH CUrinalysisDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)   //此功能未生效
{
HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor);
	
	// TODO: Change any attributes of the DC here
	switch (nCtlColor)
	{
		case CTLCOLOR_STATIC:
		{
			if(pWnd->GetDlgCtrlID() == IDC_STATIC1)//静态控件
			{
			//	pDC->SetBkColor(RGB(0,255,255));//设置背景色
				pDC->SetTextColor(RGB(5,0,0));//设置文本颜色
				hbr=(HBRUSH)GetStockObject(LTGRAY_BRUSH);//控件的填充颜色
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

int CUrinalysisDlg::OnCreate(LPCREATESTRUCT lpCreateStruct)   //////程序初始化中涉及注册表访问的都在这里面
{

    CSplashWnd::ShowSplashScreen(this);//显示启动画面
	
	if (CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;
	InitRegist();

	////////////////////////加载美化风格
	char data[100];
    CRegKey Rek;
	DWORD cbA=200;
	Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
	Rek.QueryValue(data,"Theme",&cbA);
     //skinppLoadSkin(_T(data));//AquaOS.ssk为项目下的皮肤文件	
	Rek.Close();
	 //skinppSetNoSkinHwnd(GetDlgItem(IDC_LIST1)->m_hWnd);//
	 //skinppLoadSkin("UMskin.ssk");
	return 0;
}



void CUrinalysisDlg::LoadContent(CStringArray& item,CString language)  //将检验项目名称加入CStringArray数组中
{    
    patientResult.RemoveAll();	
	if (language=="zhongwen")
	{
		if(item.GetAt(0)=="T")patientResult.Add("尿胆原");
		if(item.GetAt(1)=="T")patientResult.Add("潜血"); 
	    if(item.GetAt(2)=="T")patientResult.Add("胆红素");  
		if(item.GetAt(3)=="T")patientResult.Add("酮体"); 
		if(item.GetAt(4)=="T")patientResult.Add("白细胞");  
		if(item.GetAt(5)=="T")patientResult.Add("葡萄糖");   
		if(item.GetAt(6)=="T")patientResult.Add("蛋白质");  
		if(item.GetAt(7)=="T")patientResult.Add("酸碱度");  
		if(item.GetAt(8)=="T")patientResult.Add("亚硝酸盐");
		if(item.GetAt(9)=="T")patientResult.Add("比重"); 
		if(item.GetAt(10)=="T")patientResult.Add("维生素C"); 
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




//程序第一次打开，为安装注册表 
void CUrinalysisDlg::InitRegist()
{
    CRegKey Rek;
	if(Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting")!=ERROR_SUCCESS)   //H打开失败
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
//打开辅助按钮
void CUrinalysisDlg::OnOption()   //辅助选项按钮响应
{
	
	COptionDlg dlg;
	if (dlg.DoModal()==IDOK)
	{
		m_themeChange=TRUE;
		Invalidate();
	}

}
//list列表查询响应函数

void CUrinalysisDlg::OnSerach()    ////查询记录的按键响应
{

	int count=0;  //记录搜寻到的结果数量
//	UpdateData(FALSE);
    UpdateData(TRUE);
    m_grid.DeleteAllItems();  //清除表中数据
	GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
	GetDlgItem(IDC_REFRESHDATA)->EnableWindow(FALSE);
	_bstr_t bstrSQL;
   	m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //将控件的内容清空   
	m_sex=-1;    //将控件的内容清空  
	InitItemField();    //清空结果十一项的显示结果
	
  UpdateData(FALSE);
if (m_nameSerach==FALSE&&m_timeSerach==FALSE)    
{
   MessageBox("请至少选择一种查询方式","HIGHTOP汉唐", MB_ICONEXCLAMATION );  
   return;
}
if (m_nameSerach&&m_timeSerach)           //按姓名时间双查询
{    

	if (m_nameSer==""){ MessageBox("请输入要查的名字，不能为空","HIGHTOP汉唐", MB_ICONEXCLAMATION );  return;}
    CString str;
    CString strTime=m_timeSer.Format("%Y-%m-%d");  //获取到的为日期 如：2010-03-05
    str=" where 病人姓名 like '%"+m_nameSer+"%' and 检验时间='"+strTime+"'"; 
    bstrSQL = "select*from tb_Sample1"+str+"order by ID asc";	
    
}

else if(m_nameSerach) {	              //按姓名查询
	CString str;
	if (m_nameSer.IsEmpty()) 
	{MessageBox("请输入要查的名字，不能为空","HIGHTOP汉唐", MB_ICONEXCLAMATION ); return;}
	str=" where 病人姓名 like '%"+m_nameSer+"%'"; 
	bstrSQL = "select*from tb_Sample1"+str+"order by ID asc";

}
else if (m_timeSerach){               //按时间查询
	CString str; 
	CString strTime=m_timeSer.Format("%Y-%m-%d");  //获取到的为日期 如：2010-03-05
	str=" where 检验时间='"+strTime; 
	bstrSQL = "select*from tb_Sample1"+str+"' order by ID asc";
}
//  _variant_t RecordsAffected; 
//CString sql="delete from tb_Sample1 where 标本号,检验日期,时间 in ( select 标本号,检验日期,时间 from table group by 标本号,检验日期 having count(标本号,检验日期,时间)>1)";
//m_pConnection->Execute((_bstr_t)sql,&RecordsAffected,adCmdText);
    CString a;
	a.Format("%d",count);
    GetDlgItem(IDC_STATIC1)->SetWindowText("搜索到"+a+"条数据");	
	m_pRecordset.CreateInstance(__uuidof(Recordset));
	m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
	while(!m_pRecordset->adoEOF)
	{   

		CString s[7];
		s[0]=(char*)(_bstr_t)m_pRecordset->GetCollect("样本号");
		s[1]=(char*)(_bstr_t)m_pRecordset->GetCollect("检验时间");
        s[2]=(char*)(_bstr_t)m_pRecordset->GetCollect("时间");
		s[3]=(char*)(_bstr_t)m_pRecordset->GetCollect("病人姓名");
		s[4]=(char*)(_bstr_t)m_pRecordset->GetCollect("性别");
		s[5]=(char*)(_bstr_t)m_pRecordset->GetCollect("年龄");
		s[6]=(char*)(_bstr_t)m_pRecordset->GetCollect("病历号");

		
		m_grid.AddItem(s[0],s[1],s[2],s[3],s[4],s[5],s[6],"");
        //给每条信息添加颜色
        CString color=(char*)(_bstr_t)m_pRecordset->GetCollect("是否异常");
		if(color=="1")m_grid.SetItemColor(count,0,RGB(0,0,0),RGB(255,0,0));
        if(color=="0")m_grid.SetItemColor(count,0,RGB(0,0,0),RGB(0,255,0));
	//	m_grid.SetItemText(0,0,(char*)(_bstr_t)m_pRecordset->GetCollect("样本号"));
	//	m_grid.SetItemText(0,1,(char*)(_bstr_t)m_pRecordset->GetCollect("检验时间"));
	 //  m_grid.SetItemText(0,2,(char*)(_bstr_t)m_pRecordset->GetCollect("时间"));
	//	m_grid.SetItemText(0,3,(char*)(_bstr_t)m_pRecordset->GetCollect("病人姓名"));
	//	m_grid.SetItemText(0,4,(char*)(_bstr_t)m_pRecordset->GetCollect("性别"));
	//	m_grid.SetItemText(0,5,(char*)(_bstr_t)m_pRecordset->GetCollect("年龄"));
	//	m_grid.SetItemText(0,6,(char*)(_bstr_t)m_pRecordset->GetCollect("病历号"));
		
	    count++;     //数据条数计数器
		CString a;
	    a.Format("%d",count);
        GetDlgItem(IDC_STATIC1)->SetWindowText("搜索到"+a+"条数据");
		m_pRecordset->MoveNext();
	}
	
	m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //将控件的内容清空
    m_sex=-1;    //将控件的内容清空
      
}

void CUrinalysisDlg::OnDeletedata()              ///支持选中多个同时删除
{

int count=0;     //删除条数计数器
CString s;
s.Format("%d",count);
GetDlgItem(IDC_STATIC1)->SetWindowText("成功删除"+s+"条数据");   //显示成功删除0条数据
//以下可以实现框选删除
int *pBuf = new int[m_grid.GetItemCount()];
int num = 0;
POSITION pos;
pos = m_grid.GetFirstSelectedItemPosition();
while(pos != NULL)
{
     pBuf[num++] = m_grid.GetNextSelectedItem(pos);	
	
}
for(int i=num; i>0; i--){
    	CString str1=m_grid.GetItemText(pBuf[i-1],0);   //样本号	
		CString str2=m_grid.GetItemText(pBuf[i-1],1);  //检验时间
		CString str3=m_grid.GetItemText(pBuf[i-1],2);//时间
	    CString bstrSQL;
		_variant_t RecordsAffected;
		bstrSQL = "delete from tb_Sample1 where 样本号='"+str1+"' and 检验时间='"+str2+"' and 时间='"+str3+"'";
		m_pConnection->Execute((_bstr_t)bstrSQL,&RecordsAffected,adCmdText); 
		m_grid.DeleteItem(pBuf[i-1]);
		count++;          //删除条数计数器
		CString s;
		s.Format("%d",count);
        GetDlgItem(IDC_STATIC1)->SetWindowText("成功删除"+s+"条数据");
		
}
        delete [] pBuf;	// TODO: Add your control notification handler code here
		GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);              //打印按钮失效
	    GetDlgItem(IDC_REFRESHDATA)->EnableWindow(FALSE);        //更新按钮失效
		
	
	
		m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //将控件的内容清空
        m_sex=-1;    //将控件的内容清空
        InitItemField();    //清空结果十一项的显示结果
		UpdateData(FALSE);   
/*
	for(int i=0; i<m_grid.GetItemCount(); i++){					//遍历整个列表视图
		
		if(m_grid.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED ){
			CString str1=m_grid.GetItemText(i,0);   //样本号	
			CString str2=m_grid.GetItemText(i,1);  //检验时间
		    CString str3=m_grid.GetItemText(i,2);//时间
			CString bstrSQL;
			_variant_t RecordsAffected;
	
			bstrSQL = "delete from tb_Sample1 where 样本号='"+str1+"' and 检验时间='"+str2+"' and 时间='"+str3+"'";
			m_pConnection->Execute((_bstr_t)bstrSQL,&RecordsAffected,adCmdText); 
	        m_grid.DeleteItem(i);
			
			GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
	        GetDlgItem(IDC_REFRESHDATA)->EnableWindow(FALSE);
            GetDlgItem(IDC_STATIC1)->SetWindowText("成功删除1条数据");
			m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //将控件的内容清空
            m_sex=-1;    //将控件的内容清空
		
           	UpdateData(FALSE);
		}	//获取选中行
				
	}  */

}


		

void CUrinalysisDlg::OnClear()        //清空按钮响应
{
	// TODO: Add your control notification handler code here
	m_grid.DeleteAllItems();
	GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
	GetDlgItem(IDC_REFRESHDATA)->EnableWindow(FALSE);
	GetDlgItem(IDC_STATIC1)->SetWindowText("列表已清空");
	m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //将控件的内容清空
    m_sex=-1;    //将控件的内容清空
    InitItemField();    //清空结果十一项的显示结果
	UpdateData(FALSE);
}





void CUrinalysisDlg::OnInitADOConn()   //开启数据库连接
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
		AfxMessageBox("数据库丢失");
	}
}

void CUrinalysisDlg::ExitConnect()     //关闭数据库连接
{
	if(m_pRecordset!=NULL)
		m_pRecordset->Close();
	m_pConnection->Close();
}

void CUrinalysisDlg::OnClickList1(NMHDR* pNMHDR, LRESULT* pResult) 
{   
	int count=0;   //判断选中了几项，若一项则显示，若2项以上则不再详细显示
	// TODO: Add your control notification handler code here
	for(int i=0; i<m_grid.GetItemCount(); i++)					//遍历整个列表视图
	{ 
		if(m_grid.GetItemState(i, LVIS_SELECTED) == LVIS_SELECTED )	//获取选中行
		{   count++;
		    m_sex=-1;  //先将性别单选按钮置为不选中状态
			GetDlgItem(IDC_PRINT)->EnableWindow(TRUE);
         	GetDlgItem(IDC_REFRESHDATA)->EnableWindow(TRUE);
			CString str1=m_grid.GetItemText(i,0);   //样本号	
			CString str2=m_grid.GetItemText(i,1);   //检测时间
			CString str3=m_grid.GetItemText(i,2);//时间
			_bstr_t bstrSQL;
			bstrSQL = "select *from tb_Sample1 where 样本号='"+str1+"' and 检验时间='"+str2+"' and 时间='"+str3+"' order by ID desc";	
			m_pRecordset.CreateInstance(__uuidof(Recordset));
			m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
      
			//   while(!m_pRecordset->adoEOF)
	 //  { 
            if (m_pRecordset->adoEOF)     //若数据已删除但列表中仍存在，就提示用户数据无效
            {
			  MessageBox("此数据已无效，请尝试重新搜索或直接删除","HIGHTOP汉唐", MB_ICONSTOP );  	
         
			}
            if (count>1)return;   //判断选中了几项，若一项则显示，若2项以上则直接跳出函数不再执行
			if(!m_pRecordset->adoEOF){        //存在数据
				
				m_name=(char*)(_bstr_t)m_pRecordset->GetCollect("病人姓名");
				CString s1="男";
				CString s2="女";
                if ((char*)(_bstr_t)m_pRecordset->GetCollect("性别")==s1)
                {
					m_sex=0;
                }
				if ((char*)(_bstr_t)m_pRecordset->GetCollect("性别")==s2)
				{
					m_sex=1;
				
				}
	        	m_age=(char*)(_bstr_t)m_pRecordset->GetCollect("年龄");
				m_caseNumbe=(char*)(_bstr_t)m_pRecordset->GetCollect("病历号");
				m_sampleType=(char*)(_bstr_t)m_pRecordset->GetCollect("样本类型");
				m_sampleNumber=(char*)(_bstr_t)m_pRecordset->GetCollect("样本号");
				m_doctorName=(char*)(_bstr_t)m_pRecordset->GetCollect("医生");
				m_department=(char*)(_bstr_t)m_pRecordset->GetCollect("科室");
				///////////////将时间值赋给日期控件关联变量
			//	CString s=(char*)(_bstr_t)m_pRecordset->GetCollect("检验时间");
// 				int  nYear,nMonth,nDate;   
// 				sscanf(s,"%d-%d-%d", &nYear, &nMonth, &nDate);   
// 				CTime  t(nYear,   nMonth,   nDate,   0,   0,   0);

				m_time=(char*)(_bstr_t)m_pRecordset->GetCollect("检验时间");
				m_conclusion=(char*)(_bstr_t)m_pRecordset->GetCollect("诊断");
				m_time2=(char*)(_bstr_t)m_pRecordset->GetCollect("时间");
  					for (int p=0;p<21;p++)            //将21项信息全部清空
					{
	                	Result[p]="";         
					}
		        //将病人十项信息写入result集合中
				Result[11]=m_name;
				Result[12]=(char*)(_bstr_t)m_pRecordset->GetCollect("性别");
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
          ///////////初始化信息栏 ，第一步将显示的项目填充，第二部，在填充的字符后补加项目结果
		    CString language=InitItemField();   //将十一项名称填充到文本控件中
			//将十一项结果填充到文本控件中
 			for (int i=0;i<11;i++)
 		{ 
		    if (*pItem[i]=="尿胆原"||*pItem[i]=="URO")
				{
					if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{ *pItem[i]+=InsertSpace(12);}
	                *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("尿胆原");
					Result[0]=(char*)(_bstr_t)m_pRecordset->GetCollect("尿胆原");    //Result数组记录选中病人样本的所有信息

				}
			if (*pItem[i]=="潜血"||*pItem[i]=="BLD"){
                if(language=="zhongwen"){*pItem[i]+=InsertSpace(14);}else{ *pItem[i]+=InsertSpace(12);}
			   *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("潜血");
			   Result[1]=(char*)(_bstr_t)m_pRecordset->GetCollect("潜血");
			}		
			if (*pItem[i]=="胆红素"||*pItem[i]=="BIL"){
                if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{*pItem[i]+=InsertSpace(12);}
                *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("胆红素");
                Result[2]=(char*)(_bstr_t)m_pRecordset->GetCollect("胆红素");
			}
 					
 			if (*pItem[i]=="酮体"||*pItem[i]=="KET"){ 		
				 if(language=="zhongwen"){*pItem[i]+=InsertSpace(14);}else{ *pItem[i]+=InsertSpace(12);}
 			     *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("酮体");
				 Result[3]=(char*)(_bstr_t)m_pRecordset->GetCollect("酮体");
 			}
					
			if (*pItem[i]=="白细胞"||*pItem[i]=="ERY"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{ *pItem[i]+=InsertSpace(12);}
                *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("白细胞");
				Result[4]=(char*)(_bstr_t)m_pRecordset->GetCollect("白细胞");
			}
					
			if (*pItem[i]=="葡萄糖"||*pItem[i]=="GLU"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{ *pItem[i]+=InsertSpace(12);}
				*pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("葡萄糖");
				Result[5]=(char*)(_bstr_t)m_pRecordset->GetCollect("葡萄糖");

			}
			if (*pItem[i]=="蛋白质"||*pItem[i]=="PRO"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{ *pItem[i]+=InsertSpace(12);}
				*pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("蛋白质");
				Result[6]=(char*)(_bstr_t)m_pRecordset->GetCollect("蛋白质");
			}
				
			if (*pItem[i]=="酸碱度"||*pItem[i]=="PH"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(12);}else{*pItem[i]+=InsertSpace(13);}
                *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("酸碱度");
				Result[7]=(char*)(_bstr_t)m_pRecordset->GetCollect("酸碱度");
			}
					
			if (*pItem[i]=="亚硝酸盐"||*pItem[i]=="NIT"){
                if(language=="zhongwen"){*pItem[i]+=InsertSpace(10);}else{*pItem[i]+=InsertSpace(12);}
                 *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("亚硝酸盐");
				 Result[8]=(char*)(_bstr_t)m_pRecordset->GetCollect("亚硝酸盐");
			}
				
			if (*pItem[i]=="比重"||*pItem[i]=="SG"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(14);}else{*pItem[i]+=InsertSpace(13);}
				*pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("比重");
				Result[9]=(char*)(_bstr_t)m_pRecordset->GetCollect("比重");
			}
					
			if (*pItem[i]=="维生素C"||*pItem[i]=="VC"){
				if(language=="zhongwen"){*pItem[i]+=InsertSpace(11);}else{ *pItem[i]+=InsertSpace(13);}
	            *pItem[i]+=(char*)(_bstr_t)m_pRecordset->GetCollect("维生素C");
				Result[10]=(char*)(_bstr_t)m_pRecordset->GetCollect("维生素C");
			}
 		}
	}
			//	m_pRecordset->MoveNext();
	
		
	//	}

			UpdateData(false);	
		}
	}
	//	GetDlgItem(IDC_PRINT)->EnableWindow(true);   //打印按钮可以使用
    	*pResult = 0;
}

void CUrinalysisDlg::OnShowinformation()     //显示信息调整的按键响应
{
	// TODO: Add your control notification handler code here
	CShowInformation dlg;
 
	if (dlg.DoModal()==IDOK)
	{
    GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
	InitItemField();

	}	
}

void CUrinalysisDlg::ClearItemContent()   //清空显示项目static关联变量的数值
{
	
	for (int j=0;j<11;j++)
	{   
		*pItem[j]="";
	}
}

CString CUrinalysisDlg::InitItemField()             //将十一项名称填充到文本控件中
{
char dataLanguage[100];
CRegKey Rek;
DWORD cbA=100;
Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
////////////////////获取语言
Rek.QueryValue(dataLanguage,"ItemLanguage",&cbA);
CString dataLanguage2(dataLanguage);  //获取注册表中语言值的字符串
////////////////////////获取那些项目被选中的标志位
CStringArray item;

for (int i=0;i<11;i++)   //将项目复选框选中标志位写入item数组中
{
	CString s;
	char value[20]; 
	s.Format("%d",i+1);
	Rek.QueryValue(value,"Item"+s,&cbA);
	item.Add(value);         //存放十一项标志位的数组，每个元素值为"T"或"F"
}

// 	/////////将检验项目名称加入CStringArray数组中
LoadContent(item,dataLanguage2);    //第一个参数表示复选框是否选中标志位，第二个代表语言标志位,函数主要实现为patientResult填冲要显示的项目
ClearItemContent();                 //清空显示项目static关联变量的数值
Rek.Close();
for (int j=0;j<patientResult.GetSize();j++)
{   
	*pItem[j]=patientResult.GetAt(j);
//	UpdateData(FALSE);
}
UpdateData(FALSE);
return dataLanguage2;
}

CString CUrinalysisDlg::InsertSpace(int num)   //插入空格的函数
{
  CString str;
  for (int i=0;i<num;i++)
  str+=" ";
  return str;
}

void CUrinalysisDlg::OnSaveinformation()       //此模块已失效
{
	LoadSaveModule();

}

void CUrinalysisDlg::OnDestroy()              //
{
	CDialog::OnDestroy();
	ExitConnect();
    AnimateWindow(GetSafeHwnd(),500,AW_HIDE|AW_BLEND);   //软件使用渐隐效果退出
}

void CUrinalysisDlg::OnTimer(UINT nIDEvent) 
{
     if (nIDEvent==1)                         //刷新时间
     {
		 CTime time=CTime::GetCurrentTime();
		 CString m_timeShow=time.Format("现在是: %Y年%m月%d日  %H:%M:%S");
		 
		 GetDlgItem(IDC_TIMESHOW)->SetWindowText(m_timeShow);
	     UpdateWindow();
     }
	 if (nIDEvent==2){                     //设备连接等候处理
		 KillTimer(2);
		 if (!isOpen)
		 { 
			 
			 if (m_Comm.GetPortOpen())
			 {
				 m_Comm.SetPortOpen(FALSE);
			 }
			 GetDlgItem(IDC_STATIC1)->SetWindowText("设备连接失败");
             MessageBox("设备连接失败","HIGHTOP汉唐", MB_ICONSTOP );        
			 GetDlgItem(IDC_CONNECT)->EnableWindow(TRUE);
	
		 }

	 }

	CDialog::OnTimer(nIDEvent);
}

void CUrinalysisDlg::OnPrint()          //打印报单按钮响应
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
				if (m_sex==0)  {sex="男";}
				if(m_sex==1){sex="女";}
				Result[12]=sex;
                Result[13]=m_age;
				Result[14]=m_caseNumbe;
				Result[15]=m_sampleType;
				Result[17]=m_doctorName;
				Result[18]=m_department;
				Result[20]=m_conclusion;

	_variant_t RecordsAffected; 
	CString strSQL; 
	strSQL="UPDATE  tb_Sample1 SET 病人姓名='"+Result[11]+"',年龄='"+Result[13]+"',性别='"+Result[12]+"',病历号='"+Result[14]+"',样本类型='"+Result[15]+"',医生='"+Result[17]+"',科室='"+Result[18]+"',诊断='"+Result[20]+"' WHERE 样本号='"+Result[16]+"' and 检验时间='"+Result[19]+"' and 时间='"+Result[21]+"'";
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
	LoadPrintModule();            //生成报单函数

}



void CUrinalysisDlg::LoadPrintModule()            //生成报单函数
 { 
    _Application app;	
	Documents doc;
	CComVariant a (_T("")),b(false),c(0),d(true),aa(1),bb(20);
	VARIANT varOptional,varTable;//第一个操作参数，第二个用于添加表格的参数
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
	varTable.vt=VT_I4;//变量对应的数据类型为long
	varTable.lVal=0;
	//初始化连接
	app.CreateDispatch("word.Application");  //创建应用实例
	doc.AttachDispatch(app.GetDocuments());   //添加文档集合实例
	doc1.AttachDispatch(doc.Add(&a,&b,&c,&d));  //添加文档实例
	
	sele.AttachDispatch(app.GetSelection());  //创建光标实例
	pf=sele.GetParagraphFormat();          //格式话实例
	pParas=sele.GetParagraphs();           //段落实例
	font=sele.GetFont();
	shp=doc1.GetShapes();
 	wordRange==doc1.Range(&varOptional,&varOptional);   //获取wordRange
 	OnInitADOConn();         //链接数据库
	CTime time=CTime::GetCurrentTime();
	CString m_timeShow=time.Format("%H:%M:%S");
 	_bstr_t bstrSQL;
 
	COleVariant vStyle0((long)-2);
 	pf.SetStyle(&vStyle0);
	pParas.SetAlignment(1);   //文字居中
	font.SetBold(true);      //字体加粗
////////////////////////////////////设置列表上端的固定文字
	
	//////从注册表中提取报告单标题
	char caption[100];
    CRegKey Rek;
	DWORD cbA=200;
	Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
	Rek.QueryValue(caption,"ReportCaption",&cbA);
    sele.TypeText(caption);

    sele.TypeParagraph();          //换行  
	sele.TypeText("姓名："+Result[11]+"\t性别："+Result[12]+"\t年龄："+Result[13]+"\t病历号："+Result[14]+"\t样本类型："+Result[15]);
	sele.TypeParagraph();          //换行 
	sele.TypeText("样本号："+Result[16]+"\t\t送检医生："+Result[17]+"\t\t科室："+Result[18]);
	sele.TypeParagraph();          //换行
    sele.TypeText("检验时间："+Result[19]+"   "+Result[21]+"     诊断："+Result[20]);
	sele.TypeParagraph();          //换行
	sele.TypeParagraph();          //换行
	sele.TypeText("打印时间："+m_timeShow);
	font.SetBold(true);
	sele.TypeText("\t\t\t\t结果只对该样本负责");
		sele.TypeParagraph(); 
			sele.TypeParagraph(); 
	shp.AddLine(20,187,580,187,vOpt);       //加横线
	shp.AddLine(20,188,580,188,vOpt);

	shp.AddLine(20,600,580,600,vOpt);
	shp.AddLine(20,602,580,602,vOpt);


	////////////////////////////////////////////先打开word让其慢慢加载
	 app.SetVisible(true);    //打开word  
	

 	//创建表格的过程
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

	wordTable.SetAllowAutoFit(false);         //表格不自适应列宽 
	Columns tblColumns;
	tblColumns=wordTable.GetColumns();
	int colCount = tblColumns.GetCount();   //随心所欲设置表格列宽
   for(int i=1;i<=colCount;i++)
   {
   	Column   wordColumn = tblColumns.Item(i);
    switch(i)
    {
  
	 case 1:                //第1列
     wordColumn.SetWidth(70.0,0);
     break;
	 case 2:                //第2列
     wordColumn.SetWidth(70.0,0);
     break;
	 case 3:                //第3列
     wordColumn.SetWidth(90.0,0);
     break;
	 case 4:                //第4列
     wordColumn.SetWidth(140.0,0);
     break;
	 case 5:                //第5列
     wordColumn.SetWidth(100.0,0);
     break;
 
    }
   }

	Borders tblBorders;//边框对象
	tblBorders=wordTable.GetBorders();
	tblBorders.SetEnable(0);//无边框
    Cell tblCell;
	
/////////////////////////////////////添加列名
for(long num=1;num<=countl+1;num++)   //还要插入结果一列 所以+1
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
 	
 if (num==4)   //手工加入结果一列
 {	
 	tblCell=wordTable.Cell(1,num);
  	wordRange=tblCell.GetRange();
//	font.SetBold(true); 
  	wordRange.InsertAfter("    结果    ");
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

///////////////////////////数据库获取的数据填充到表格中
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
	if (num==4)        //将结果填充至此
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

void CUrinalysisDlg::LoadSaveModule()     //用于实现保存的保存模块
{
	_variant_t RecordsAffected; 
   	UpdateData(true);
	if(!Upload)
	{
		AfxMessageBox("尿仪未上传过有效数据，无法保存");
		return;
	}
	CString sex,time;
    if (m_sex==0)  {sex="男";}else{sex="女";}
    //time = m_time.Format("%Y-%m-%d");
	time=m_time;
	CString strSQL; 
	strSQL.Format("INSERT INTO tb_Sample1(病人姓名,年龄,性别,病历号,样本类型,样本号,检验时间,医生,科室,诊断,尿胆原,潜血,胆红素,酮体,白细胞,葡萄糖,蛋白质,酸碱度,亚硝酸盐,比重,维生素C)VALUES ('%s','%s','%s','%s','%s', '%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')", m_name,m_age,sex,m_caseNumbe,m_sampleType,m_sampleNumber,time,m_doctorName,m_department,
		m_conclusion,Result[0],Result[1],Result[2],Result[3],Result[4],Result[5],Result[6],Result[7],Result[8],Result[9],Result[10]); 
	m_pConnection->Execute((_bstr_t)strSQL,&RecordsAffected,adCmdText); //添加记录
	////控件中值归为初始化状态
    
	
	
	_bstr_t bstrSQL = "select*from tb_Sample1 where 病人姓名='"+m_name+"' and 样本号='"+m_sampleNumber+"' and 检验时间='"+time+"'";
	m_pRecordset.CreateInstance(__uuidof(Recordset));
	m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
	while(!m_pRecordset->adoEOF)
	{
			
		
	CString s[7];
		s[0]=(char*)(_bstr_t)m_pRecordset->GetCollect("样本号");
		s[1]=(char*)(_bstr_t)m_pRecordset->GetCollect("检验时间");
        s[2]=(char*)(_bstr_t)m_pRecordset->GetCollect("时间");
		s[3]=(char*)(_bstr_t)m_pRecordset->GetCollect("病人姓名");
		s[4]=(char*)(_bstr_t)m_pRecordset->GetCollect("性别");
		s[5]=(char*)(_bstr_t)m_pRecordset->GetCollect("年龄");
		s[6]=(char*)(_bstr_t)m_pRecordset->GetCollect("病历号");
		m_grid.AddItem(s[0],s[1],s[2],s[3],s[4],s[5],s[6]);


		m_pRecordset->MoveNext();
	}
	
	m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion="";
	m_sex=0;
//	m_time=CTime::GetCurrentTime();
	m_time="";
	m_time2="";
	UpdateData(false);	
	MessageBox("保存成功");

}

void CUrinalysisDlg::OnDatadictionary()           //数据字典响应函数
{

	// TODO: Add your control notification handler code here
	
CDictionarySet dlg;
if (dlg.DoModal()==IDOK)
{   
	WinExec("Urinalysis.EXE",SW_SHOWMAXIMIZED);   //重新打开程序
	SkinH_Detach();                //卸载皮肤
	exit(0);                      //关闭当前程序
}

}



void CUrinalysisDlg::InitControlContent()            //从数据字典中提取数据填入空间中
{
	CString s[4]={"样本类型","送检医生","科室","诊断"};
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

void CUrinalysisDlg::OnReportset()          //报告单设置按钮
{
	// TODO: Add your control notification handler code here

	CReportSet dlg;

	if (dlg.DoModal()==IDOK)
	{   
		CRegKey Rek;
       	Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");  //H打开失败
	    Rek.SetValue(Caption,"ReportCaption");           //利用全局变量caption，将报告标题更改
		Rek.Close();
	}
	
}

void CUrinalysisDlg::RegisterActiveX(CString s)        //注册ACTIVEX控件
{
		CString strFileName=g_curPath+s;
		if (strFileName.IsEmpty())
			return;
		
		//装载ActiveX控件  
		HINSTANCE hInstance = LoadLibrary(strFileName); 
		if (hInstance == NULL)  
		{     
			AfxMessageBox("不能载入Dll/OCX文件!");     
			return; 
		} 
		//取得注册函数DllRegisterServer地址 
		FARPROC lpFunc;  
		lpFunc = GetProcAddress(hInstance,_T("DllRegisterServer"));  
		//调用注册函数DllRegisterServer 
		if(lpFunc!=NULL)  
		{       
			if(FAILED((*lpFunc)()))   
			{       
				//AfxMessageBox("调用DllRegisterServer 失败!");  
				FreeLibrary(hInstance);   //释放资源       
				return;     
			}     
			//MessageBox(s+"注册成功");  
		}
		else  
		{
		//	AfxMessageBox("调用DllRegisterServer失败!");
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

void CUrinalysisDlg::OnButton1()     //标题栏点击响应
{
	// TODO: Add your control notification handler code here
   GetDlgItem(IDC_STATIC1)->SetWindowText("欢迎您访问汉唐");
   CAboutDlg dlg;
   dlg.DoModal();	
   
}

void CAboutDlg::OnLinkbutton()      //访问汉唐的响应
{
	// TODO: Add your control notification handler code here
	
	ShellExecute(NULL, 
		"open", 
		"http://www.hightopqd.com/", 
		NULL, 
		NULL, 
		SW_SHOWNORMAL);	
}

BEGIN_EVENTSINK_MAP(CUrinalysisDlg, CDialog)               //串口响应映射（自动生成）
    //{{AFX_EVENTSINK_MAP(CUrinalysisDlg)
	ON_EVENT(CUrinalysisDlg, IDC_MSCOMM1, 1 /* OnComm */, OnOnCommMscomm1, VTS_NONE)
	//}}AFX_EVENTSINK_MAP
END_EVENTSINK_MAP()

void CUrinalysisDlg::OnOnCommMscomm1()            //串口触发函数
{   
	int k, nEvent;
	nEvent = m_Comm.GetCommEvent();
	switch(nEvent)
	{
	case 2:  //收到大于RTHresshold个字符
		k = m_Comm.GetInBufferCount(); //接收到的字符数目
		if(k==5)          //5个字符触处理
		{
			portReceiveMessage=receivePortData();    //取回串口接收到的数据
			if (portReceiveMessage=="READY")
			{
				isOpen=TRUE;            //将连接成功的标志位置真
                GetDlgItem(IDC_CONNECT)->EnableWindow(TRUE);
				GetDlgItem(IDC_CONNECT)->SetWindowText("断开连接");
				GetDlgItem(IDC_POSTBACK)->EnableWindow(TRUE);
				GetDlgItem(IDC_SERACH)->EnableWindow(FALSE);
				GetDlgItem(IDC_DELETEDATA)->EnableWindow(FALSE);
				GetDlgItem(IDC_CLEAR)->EnableWindow(FALSE);
                monitorCount=0;                 //本次开启串口后收到的结果条数的计数器
				m_grid.DeleteAllItems();
				InitItemField();  //清空结果栏内的内容
   				m_name=m_age=m_caseNumbe=m_sampleType=m_sampleNumber=m_doctorName=m_department=m_conclusion=m_time=m_time2="";   //将控件的内容清空
				m_sex=-1;    //将控件的内容清空
				UpdateData(FALSE);
			    GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
			    GetDlgItem(IDC_STATIC1)->SetWindowText("设备连接成功");
     			MessageBox("连接成功","HIGHTOP汉唐", MB_ICONINFORMATION );
			    m_Comm.SetRThreshold(32); //32个字符触发函数
		
			}

		}   
		else if (k==32)      //32个字符处理
		{	
               //已收到的样本个数
			portReceiveMessage=receivePortData();
			///////////////////////////////////////////////////实时监控
            if (portReceiveMessage.Left(3)=="RES"&&portReceiveMessage.Right(3)=="END")   //格式符合的指令
            {
		    CString tempResult[22];
			CString exceptionFlag="0";      //判断十一项结果有异常的，若为1则存在异常
			for(int a=0;a<22;a++)  tempResult[a]="";    
		  //  GetDlgItem(IDC_PRINT)->EnableWindow(FALSE);
			CString tempSampleNumber,tempTime,tempTime2; ///临时变量存储样本号 日期 时间
			 for (int i=3;i<6;i++)  tempSampleNumber+=portReceiveMessage.GetAt(i); 
			 tempResult[16]=tempSampleNumber;     //标本号
			/////////////////////////////////////////
			 tempTime+="20";           //取日期
			 for (int j=6;j<12;j++)    
			 { 
			   tempTime+=portReceiveMessage.GetAt(j);  
			   if (j%2==1&&j<11) tempTime+="-";   
			 }
			 tempResult[19]=tempTime;
            ////////////////////////////////////////
			 for (int k=12;k<18;k++)                     //取时间
			 {
				 tempTime2+=portReceiveMessage.GetAt(k);
				 if (k%2==1&&k<17) tempTime2+=":";
			 }
             tempResult[21]=tempTime2;
			 ///////////////////////////////////////
			 for (int m=18;m<29;m++)                     //判断十一个项目是否存在异常项，只要有一项就置为1
			 {   
				 if (portReceiveMessage.GetAt(m)!='0')
				 {
					 exceptionFlag="1";
                   
				 }
				 tempResult[m-18]=decideResult(m-18,portReceiveMessage.GetAt(m));   //取十一项检测结果
	
			 }
	
             _variant_t RecordsAffected; 
			  CString strSQL; 
		//	CString	strSQL2 = "delete from tb_Sample1 where 样本号='"+tempResult[16]+"' and 检验时间='"+tempResult[19]+"' and 时间='"+tempResult[21]+"'";
        //    m_pConnection->Execute((_bstr_t)strSQL2,&RecordsAffected,adCmdText); //预删除可能重复的纪录			  
			 strSQL.Format("INSERT INTO tb_Sample1(病人姓名,年龄,性别,病历号,样本类型,样本号,检验时间,医生,科室,诊断,时间,尿胆原,潜血,胆红素,酮体,白细胞,葡萄糖,蛋白质,酸碱度,亚硝酸盐,比重,维生素C,是否异常)VALUES ('%s','%s','%s','%s','%s', '%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')",tempResult[11],tempResult[13],tempResult[12],tempResult[14],tempResult[15],tempResult[16],tempResult[19],tempResult[17],tempResult[18],
			  tempResult[20],tempResult[21],tempResult[0],tempResult[1],tempResult[2],tempResult[3],tempResult[4],tempResult[5],tempResult[6],tempResult[7],tempResult[8],tempResult[9],tempResult[10],exceptionFlag); 
			 m_pConnection->Execute((_bstr_t)strSQL,&RecordsAffected,adCmdText); //添加记录
             //if(exceptionFlag="1"){colorArray.Add(1);}else{colorArray.Add(0);}

			 ////控件中值归为初始化状态	
		CString s[7];
		s[0]=tempResult[16];
		s[1]=tempResult[19];
        s[2]=tempResult[21];
	    s[3]="";
		s[4]="";
		s[5]="";
		s[6]="";
	    m_grid.AddItem(s[0],s[1],s[2],s[3],s[4],s[5],s[6],"");
		if(exceptionFlag=="1")     //存在异常项
		{
			m_grid.SetItemColor(monitorCount,0,RGB(0,0,0),RGB(255,0,0));   //存在异常项的结果着红色
			m_grid.SetItemColor(monitorCount,1,RGB(0,0,0),RGB(255,0,0));   //存在异常项的结果着红色
		//	m_grid.SetItemColor(monitorCount,2,RGB(0,0,0),RGB(255,0,0));   //存在异常项的结果着红色

		}
		if(exceptionFlag=="0")    //不存在异常项
		{
			m_grid.SetItemColor(monitorCount,0,RGB(0,0,0),RGB(0,255,0));   //不存在异常项的结果着绿色
			m_grid.SetItemColor(monitorCount,1,RGB(0,0,0),RGB(0,255,0));   //不存在异常项的结果着绿色
		//	m_grid.SetItemColor(monitorCount,2,RGB(0,0,0),RGB(0,255,0));   //不存在异常项的结果着绿色

		}
      CString count1;
	  CString count2;
	  count1.Format("%d",monitorCount+1);       //总接收到样本条数计数
      count2.Format("%d",postbackCount+1);      //回传样本数计数
	  if(isPostback)         //是否处于回传状态
	  {
		   GetDlgItem(IDC_STATIC1)->SetWindowText("回传第"+count2+"个样本,共"+postbackNumber+"个");  
	       postbackCount++;    //回传计数
	  }
	  else{
		    GetDlgItem(IDC_STATIC1)->SetWindowText("已收到"+count1+"个样本");
	  }
	          //屏幕左下角信息
	     monitorCount++;   //建立连接后总共收到的标本总数

  }
///////////////////////////////////////////////////////////批量回传头条指令
	  if (portReceiveMessage.Left(3)=="FIR"&&portReceiveMessage.Right(3)=="END")  {

		  GetDlgItem(IDC_LIST1)->EnableWindow(FALSE);    //将列表控件锁死，防止操作导致数据丢失
		  isPostback=TRUE;    //回传模式标志位
	      postbackCount=0;    //将统计回传数的变量置为0
	      CString temp1[3];
	//	  int temp2[3];
		   //记录本次回传数的条目数
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
////////////////////////////////////////////////////////批量回传结束指令
 if (portReceiveMessage.Left(3)=="LAS"&&portReceiveMessage.Right(3)=="END")  {
	 
	  isPostback=FALSE;
	  GetDlgItem(IDC_STATIC1)->SetWindowText("数据回传完毕");
	   if (postbackCount==0)    //没有回传数据
	   {
		  MessageBox("没有数据需回传","HIGHTOP汉唐", MB_ICONINFORMATION );

	   }
	  if (postbackCount>0){
		  MessageBox("全部数据回传完毕","HIGHTOP汉唐", MB_ICONINFORMATION );
	   }
        GetDlgItem(IDC_LIST1)->EnableWindow(TRUE);  //将锁死的列表框解开
	    postbackCount=0;    //将统计回传数的变量置为0
	  }
	}
        else{
		m_Comm.GetInput();     //清除输入缓冲区
		}
		
	}    
	
}

void CUrinalysisDlg::InitSerialPort()         //初始化串口，并发送connect指令
{	//char data[100];
    CRegKey Rek;
    DWORD port=200;
	DWORD rate=200;
	CString bandRate;
    Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
	Rek.QueryValue(port,"Port");
	Rek.QueryValue(rate,"BandRate");
    bandRate.Format("%d",rate);
	m_Comm.SetCommPort(port); //选择COM1
	m_Comm.SetInBufferSize(1024); //设置输入缓冲区的大小，Bytes
	m_Comm.SetOutBufferSize(512); //设置输入缓冲区的大小，Bytes
	try
	{

	if (!m_Comm.GetPortOpen())
	{
		m_Comm.SetPortOpen(TRUE);
	}
	m_Comm.SetInputMode(1); //设置输入方式为二进制方式
	m_Comm.SetSettings(bandRate+",n,8,1"); //设置波特率等参数
	m_Comm.SetRThreshold(5); //为1表示有一个字符引发一个事件　　　　
//	m_Comm.SetInputLen(0);
	sendPortData("CONNECT");   //发送连接请求命令
	}
	catch (_com_error *e)
	{
		AfxMessageBox(e->ErrorMessage());
		
	}

}

void CUrinalysisDlg::OnConnect()        //连接按钮响应
{ 
	if (!isOpen)    //isOpen建立连接标志位，未连接
	{
		CConnectDlg dlg;		
		switch(dlg.DoModal()){
		case IDOK:
            InitSerialPort();     //初始化串口，发送connect指令
			GetDlgItem(IDC_CONNECT)->EnableWindow(FALSE);
			SetTimer(2,1500,NULL);
		}
	}
    else{           //已经连接了
		sendPortData("DISCONNECT");   //断开连接向下位机发送断开指令
		if(m_Comm.GetPortOpen()) //打开串口   
		{
		  m_Comm.SetPortOpen(FALSE);

		}
 		isOpen=FALSE;    
	    GetDlgItem(IDC_STATIC1)->SetWindowText("断开设备连接");
		GetDlgItem(IDC_CONNECT)->SetWindowText("连接设备");
		GetDlgItem(IDC_SERACH)->EnableWindow(TRUE);
		GetDlgItem(IDC_DELETEDATA)->EnableWindow(TRUE);
		GetDlgItem(IDC_CLEAR)->EnableWindow(TRUE);
		GetDlgItem(IDC_POSTBACK)->EnableWindow(FALSE);   //回传按钮不可用
		MessageBox("连接断开","HIGHTOP汉唐", MB_ICONEXCLAMATION );
	}
	// TODO: Add your control notification handler code here
}


/////////////////////////////向串口发送数据
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
    m_Comm.SetOutput(COleVariant(array)); // 发送数据	
}



/////////////////////////////从串口接收数据
CString CUrinalysisDlg::receivePortData()
{
	VARIANT variant_inp;
    COleSafeArray safearray_inp;   //COleSafeArray类是用于处理任意类型和维数的数组的类**********
    long i = 0;
    int len;
	BYTE rxdata[1000];    
	CString receive;
	//读取缓冲区数据
	variant_inp = m_Comm.GetInput();
	//将VARIANT型变量值赋给ColeSafeArray类型变量
	safearray_inp = variant_inp;
	len = safearray_inp.GetOneDimSize(); //返回一个一维的COleSafeArray对象中的元素个数 
	for(i = 0; i < len; i++)
	{
		safearray_inp.GetElement(&i, &rxdata[i]);  //将数据保存到字符数组中
	}
	rxdata[i] = '\0'; 
	receive= rxdata;  	
	return receive;
}
	




CString CUrinalysisDlg::decideResult(int location, CString result)   //做出结果的判断函数
{   CString res="";
    switch(location){

	case 0:       //尿胆原
		if (result=="0") {res="NORM 3.3 umol/l "; break;}
        if (result=="1") {res="1+   33 umol/l  ↑"; break;}
		if (result=="2") {res="2+   66 umol/l  ↑"; break;}
		if (result=="3") {res=">=3+ 131 umol/l ↑"; break;}
		if (result=="9") {res="结果无效"; break;}    // 若机器无效结果，则为9说明此样本报废

	case 1:    //潜血
		if (result=="0") {res="-    0 cells/ul";  break;}
		if (result=="1") {res="+-   10 cells/ul  ↑"; break;}
		if (result=="2") {res="+    25 cells/ul  ↑"; break;}
		if (result=="3") {res="2+   50 cells/ul  ↑";  break;}
        if (result=="4") {res="3+   250 cells/ul ↑"; break;}
		if (result=="9") {res="结果无效"; break;}    // 若机器无效结果，则为9说明此样本报废

        
	case 2:   //胆红素
		if (result=="0") {res="-    0 umol/l"; break;}
        if (result=="1") {res="+    17 umol/l  ↑"; break;}
		if (result=="2") {res="2+   50 umol/l  ↑"; break;}
		if (result=="3") {res="3+   100 umol/l ↑"; break;}
	    if (result=="9") {res="结果无效"; break;}    // 若机器无效结果，则为9说明此样本报废
	case 3:   //酮体
		if (result=="0") {res="-    0 mmol/l"; break;}
        if (result=="1") {res="+    1.5 mmol/l  ↑"; break;}
		if (result=="2") {res="2+   4.0 umol/l  ↑"; break;}
		if (result=="3") {res="3+   8.0 umol/l  ↑"; break;}
	    if (result=="9") {res="结果无效"; break;}    // 若机器无效结果，则为9说明此样本报废
	case 4:   //白细胞
		if (result=="0") {res="-    0 cells/ul"; break;}
        if (result=="1") {res="+-   15 cells/ul ↑"; break;}
		if (result=="2") {res="+    70 cells/ul ↑"; break;}
		if (result=="3") {res="2+   125 cells/ul↑"; break;}
        if (result=="4") {res="3+   500 cells/ul↑"; break;}
        if (result=="9") {res="结果无效"; break;}    // 若机器无效结果，则为9说明此样本报废
    case 5:  //葡萄糖
		if (result=="0") {res="-    0 mmol/l"; break;}
        if (result=="1") {res="+-   2.8 mmol/l ↑"; break;}
		if (result=="2") {res="+    5.5 mmol/l ↑"; break;}
		if (result=="3") {res="2+   14 mmol/l  ↑"; break;}
        if (result=="4") {res="3+   28 mmol/l  ↑"; break;}
		if (result=="5") {res="4+   55 mmol/l  ↑"; break;}
	    if (result=="9") {res="结果无效"; break;}    // 若机器无效结果，则为9说明此样本报废
	case 6:  //蛋白质
		if (result=="0") {res="-    0 g/l"; break;}
        if (result=="1") {res="+-   0.15 g/l ↑"; break;}
		if (result=="2") {res="+    0.3 g/l  ↑"; break;}
		if (result=="3") {res="2+   1.0 g/l  ↑"; break;}
		if (result=="3") {res=">=3+ 3 g/l    ↑"; break;}
	    if (result=="9") {res="结果无效"; break;}    // 若机器无效结果，则为9说明此样本报废
	case 7:  //PH
       	if (result=="0") {res="5.0"; break;}
        if (result=="1") {res="6.0  ↑"; break;}
		if (result=="2") {res="7.0  ↑"; break;}
		if (result=="3") {res="8.0  ↑"; break;}
		if (result=="4") {res="9.0  ↑"; break;}
	    if (result=="9") {res="结果无效"; break;}    // 若机器无效结果，则为9说明此样本报废
	case 8:    //亚硝酸盐
		if (result=="0") {res="-    0 umol/l"; break;}
		if (result=="1") {res="+    18 umol/l  ↑"; break;}
	    if (result=="9") {res="结果无效"; break;}    // 若机器无效结果，则为9说明此样本报废
		
	case 9:     //比重
		if (result=="0") {res="<=1.005"; break;}
        if (result=="1") {res="1.010   ↑";break;}
		if (result=="2") {res="1.015   ↑"; break;}
		if (result=="3") {res="1.020   ↑"; break;}
        if (result=="4") {res="1.025   ↑"; break;}
		if (result=="5") {res="1.030   ↑"; break;}
	    if (result=="9") {res="结果无效"; break;}    // 若机器无效结果，则为9说明此样本报废
	case 10:   //维生素c
		if (result=="0") {res="-    0 mmol/l"; break;}
        if (result=="1") {res="+-   0.6 mmol/l  ↑"; break;}
		if (result=="2") {res="+    1.4 mmol/l  ↑"; break;}
		if (result=="3") {res="2+   2.8 mmol/l  ↑"; break;}
        if (result=="4") {res="3+   5.6 mmol/l  ↑"; break;}
	    if (result=="9") {res="结果无效"; break;}    // 若机器无效结果，则为9说明此样本报废
	}
		return res;

}

void CUrinalysisDlg::OnRefreshdata()             //更新数据按钮
{
                UpdateData(TRUE);
	            Result[11]=m_name;
				CString sex;
				if (m_sex==0)  {sex="男";}
				if(m_sex==1){sex="女";}
				Result[12]=sex;
                Result[13]=m_age;
				Result[14]=m_caseNumbe;
				Result[15]=m_sampleType;
				Result[17]=m_doctorName;
				Result[18]=m_department;
				Result[20]=m_conclusion;

	_variant_t RecordsAffected; 
	CString strSQL; 
	strSQL="UPDATE  tb_Sample1 SET 病人姓名='"+Result[11]+"',年龄='"+Result[13]+"',性别='"+Result[12]+"',病历号='"+Result[14]+"',样本类型='"+Result[15]+"',医生='"+Result[17]+"',科室='"+Result[18]+"',诊断='"+Result[20]+"' WHERE 样本号='"+Result[16]+"' and 检验时间='"+Result[19]+"' and 时间='"+Result[21]+"'";
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
	 GetDlgItem(IDC_STATIC1)->SetWindowText("成功更新样本");
     MessageBox("更新成功","HIGHTOP汉唐", MB_ICONINFORMATION );
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


void CUrinalysisDlg::OnPostback()     //回传按钮响应
{
	CPostBack dlg;

	if(dlg.DoModal()==IDOK){

	//	if(isOpen){	
			GetDlgItem(IDC_STATIC1)->SetWindowText("等待传送，请稍后....");
			CString request="REQ"+postbackTime.Right(6)+"END";
			sendPortData(request);        //发送请求某日数据的指令
	//		MessageBox(request);
	//	}else{
		//	MessageBox("未连接设备，请连接后尝试");
	//	}

	}
}

void CAboutDlg::OnHelp()     //帮助按钮的响应
{

	ShellExecute(NULL,_T("open"),_T(".\\说明书.doc"),NULL,NULL,SW_SHOWNORMAL);
	// TODO: Add your control notification handler code here
	
}

void CUrinalysisDlg::OnOK() 
{
	// TODO: Add extra validation here
	
	//CDialog::OnOK();
}

BOOL CUrinalysisDlg::PreTranslateMessage(MSG* pMsg)   //屏蔽ESC 和 ENTER、SPACE 键，不允许按此三键退出程序
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



void CUrinalysisDlg::DateBaseAutoClear(int item)    //数据库清理函数，超过指定条数后，会清理数据库
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
		       m_pConnection->Execute((_bstr_t)strSQL,&RecordsAffected,adCmdText); //添加记录
			}
			m_pRecordset->MoveNext();	
		}
}

void CAboutDlg::OnOK()     
{
	// TODO: Add extra validation here
	GetParent()->Invalidate();   ///重绘主对话框，防止出现发虚的现象
	CDialog::OnOK();
}
