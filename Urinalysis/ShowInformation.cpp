// ShowInformation.cpp : implementation file
//

#include "stdafx.h"
#include "Urinalysis.h"
#include "ShowInformation.h"
#include "UrinalysisDlg.h"
/*#include "UrinalysisDlg.cpp"*/
#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CShowInformation dialog


CShowInformation::CShowInformation(CWnd* pParent /*=NULL*/)
	: CDialog(CShowInformation::IDD, pParent)
{
	//{{AFX_DATA_INIT(CShowInformation)
	
	m_language = -1;
	m_check1 = FALSE;
	m_check2 = FALSE;
	m_check3 = FALSE;
	m_check4 = FALSE;
	m_check5 = FALSE;
	m_check6 = FALSE;
	m_check7 = FALSE;
	m_check8 = FALSE;
	m_check9 = FALSE;
	m_check10 = FALSE;
	m_check11 = FALSE;
	//}}AFX_DATA_INIT

}


void CShowInformation::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CShowInformation)
	DDX_Control(pDX, IDC_LIST1, m_listBox);
	DDX_Radio(pDX, IDC_RADIO1, m_language);
	DDX_Check(pDX, IDC_CHECK1, m_check1);
	DDX_Check(pDX, IDC_CHECK2, m_check2);
	DDX_Check(pDX, IDC_CHECK3, m_check3);
	DDX_Check(pDX, IDC_CHECK4, m_check4);
	DDX_Check(pDX, IDC_CHECK5, m_check5);
	DDX_Check(pDX, IDC_CHECK6, m_check6);
	DDX_Check(pDX, IDC_CHECK7, m_check7);
	DDX_Check(pDX, IDC_CHECK8, m_check8);
	DDX_Check(pDX, IDC_CHECK9, m_check9);
	DDX_Check(pDX, IDC_CHECK10, m_check10);
	DDX_Check(pDX, IDC_CHECK11, m_check11);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CShowInformation, CDialog)
	//{{AFX_MSG_MAP(CShowInformation)
	ON_BN_CLICKED(IDC_RADIO1, OnRadio1)
	ON_BN_CLICKED(IDC_RADIO2, OnRadio2)
	ON_BN_CLICKED(IDC_CHECK1, OnCheck1)
	ON_BN_CLICKED(IDC_CHECK2, OnCheck2)
	ON_BN_CLICKED(IDC_CHECK3, OnCheck3)
	ON_BN_CLICKED(IDC_CHECK4, OnCheck4)
	ON_BN_CLICKED(IDC_CHECK5, OnCheck5)
	ON_BN_CLICKED(IDC_CHECK6, OnCheck6)
	ON_BN_CLICKED(IDC_CHECK7, OnCheck7)
	ON_BN_CLICKED(IDC_CHECK8, OnCheck8)
	ON_BN_CLICKED(IDC_CHECK9, OnCheck9)
	ON_BN_CLICKED(IDC_CHECK10, OnCheck10)
	ON_BN_CLICKED(IDC_CHECK11, OnCheck11)
	ON_WM_MOUSEMOVE()
	ON_WM_CREATE()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CShowInformation message handlers
int CShowInformation::OnCreate(LPCREATESTRUCT lpCreateStruct) 
{
	if (CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;
	
	// TODO: Add your specialized creation code here
	char dataLanguage[100];
    CRegKey Rek;
	DWORD cbA=100;
	Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");
	////////////////////获取语言
	Rek.QueryValue(dataLanguage,"ItemLanguage",&cbA);
	CString dataLanguage2(dataLanguage);  //获取注册表中语言值得字符串
	if (dataLanguage2=="zhongwen")  m_language=0;
    if (dataLanguage2=="yingwen")   m_language=1;
	////////////////////获取复选框是否选中标志位
	for (int i=0;i<11;i++)
	{
        CString s;
        char value[20]; 
        s.Format("%d",i+1);
        Rek.QueryValue(value,"Item"+s,&cbA);
        item.Add(value);
	}
	if(item.GetAt(0)=="T"){m_check1=TRUE;}
    if(item.GetAt(1)=="T"){m_check2=TRUE;}
	if(item.GetAt(2)=="T"){m_check3=TRUE;}
    if(item.GetAt(3)=="T"){m_check4=TRUE;}
	if(item.GetAt(4)=="T"){m_check5=TRUE;}
    if(item.GetAt(5)=="T"){m_check6=TRUE;}
	if(item.GetAt(6)=="T"){m_check7=TRUE;}
    if(item.GetAt(7)=="T"){m_check8=TRUE;}
	if(item.GetAt(8)=="T"){m_check9=TRUE;}
    if(item.GetAt(9)=="T"){m_check10=TRUE;}
	if(item.GetAt(10)=="T"){m_check11=TRUE;}
	Rek.Close();
	return 0;
}


BOOL CShowInformation::OnInitDialog() 
{
	CDialog::OnInitDialog();
	
	LoadCheckCaption();
	LoadListContent();
	return TRUE;
}



void CShowInformation::OnRadio1() 
{   m_language=0;
	LoadCheckCaption();
	LoadListContent();

}
void CShowInformation::OnRadio2() 
{
   m_language=1;
   LoadCheckCaption();
   LoadListContent();

}



void CShowInformation::OnCheck1() 
{   m_check1=!m_check1;
	LoadListContent();

}

void CShowInformation::OnCheck2() 
{   
	m_check2=!m_check2;
	LoadListContent();

}

void CShowInformation::OnCheck3() 
{ 
	m_check3=!m_check3;
	LoadListContent();

}

void CShowInformation::OnCheck4() 
{   m_check4=!m_check4;
	LoadListContent();
	// TODO: Add your control notification handler code here
	
}

void CShowInformation::OnCheck5() 
{
	// TODO: Add your control notification handler code here
	m_check5=!m_check5;
	LoadListContent();
}

void CShowInformation::OnCheck6() 
{
	m_check6=!m_check6;
	LoadListContent();
	// TODO: Add your control notification handler code here
	
}

void CShowInformation::OnCheck7() 
{
	// TODO: Add your control notification handler code here
	m_check7=!m_check7;
	LoadListContent();
}

void CShowInformation::OnCheck8() 
{
	// TODO: Add your control notification handler code here
	m_check8=!m_check8;
	LoadListContent();	
}

void CShowInformation::OnCheck9() 
{
	// TODO: Add your control notification handler code here
	m_check9=!m_check9;
	LoadListContent();
}

void CShowInformation::OnCheck10() 
{
	// TODO: Add your control notification handler code here
	m_check10=!m_check10;
	LoadListContent();
}

void CShowInformation::OnCheck11() 
{
	// TODO: Add your control notification handler code here
	m_check11=!m_check11;
	LoadListContent();
}

void CShowInformation::OnOK() 
{
	// TODO: Add extra validation here
	 
    if (!m_check1&&!m_check2&&!m_check3&&!m_check4&&!m_check5&&                  //防止复选框一项都没选
		!m_check6&&!m_check7&&!m_check8&&!m_check9&&!m_check10&&!m_check11)
    {
	MessageBox("至少选择一项","HIGHTOP汉唐", MB_ICONEXCLAMATION );
	return;
    }
	CRegKey Rek;
	Rek.Open(HKEY_CURRENT_USER,"HTUrime\\Setting");

	/////////////////////保存语言设置
	if (m_language==0)
	{Rek.SetValue("zhongwen","ItemLanguage");
	}
	if (m_language==1)
	{Rek.SetValue("yingwen","ItemLanguage");
	}
     
	///////////////保存复选框打钩内容
	if (m_check1){Rek.SetValue("T","Item1");}else{Rek.SetValue("F","Item1");}
	if (m_check2){Rek.SetValue("T","Item2");}else{Rek.SetValue("F","Item2");}
	if (m_check3){Rek.SetValue("T","Item3");}else{Rek.SetValue("F","Item3");}
	if (m_check4){Rek.SetValue("T","Item4");}else{Rek.SetValue("F","Item4");}
	if (m_check5){Rek.SetValue("T","Item5");}else{Rek.SetValue("F","Item5");}
	if (m_check6){Rek.SetValue("T","Item6");}else{Rek.SetValue("F","Item6");}
	if (m_check7){Rek.SetValue("T","Item7");}else{Rek.SetValue("F","Item7");}
	if (m_check8){Rek.SetValue("T","Item8");}else{Rek.SetValue("F","Item8");}
	if (m_check9){Rek.SetValue("T","Item9");}else{Rek.SetValue("F","Item9");}
	if (m_check10){Rek.SetValue("T","Item10");}else{Rek.SetValue("F","Item10");}
	if (m_check11){Rek.SetValue("T","Item11");}else{Rek.SetValue("F","Item11");}

    Rek.Close();

	CDialog::OnOK();
}


void CShowInformation::LoadCheckCaption()
{
	
  if (m_language==0)
  {
	  GetDlgItem(IDC_CHECK1)->SetWindowText("尿胆原");
	  GetDlgItem(IDC_CHECK2)->SetWindowText("潜血");
	  GetDlgItem(IDC_CHECK3)->SetWindowText("胆红素");
	  GetDlgItem(IDC_CHECK4)->SetWindowText("酮体");
	  GetDlgItem(IDC_CHECK5)->SetWindowText("白细胞");
	  GetDlgItem(IDC_CHECK6)->SetWindowText("葡萄糖");
	  GetDlgItem(IDC_CHECK7)->SetWindowText("蛋白质");
	  GetDlgItem(IDC_CHECK8)->SetWindowText("酸碱度");
	  GetDlgItem(IDC_CHECK9)->SetWindowText("亚硝酸盐");
	  GetDlgItem(IDC_CHECK10)->SetWindowText("比重");
      GetDlgItem(IDC_CHECK11)->SetWindowText("维生素C");
	  
  }
  if (m_language==1)
  {
	  GetDlgItem(IDC_CHECK1)->SetWindowText("URO");
	  GetDlgItem(IDC_CHECK2)->SetWindowText("BLD");
	  GetDlgItem(IDC_CHECK3)->SetWindowText("BIL");
	  GetDlgItem(IDC_CHECK4)->SetWindowText("KET");
	  GetDlgItem(IDC_CHECK5)->SetWindowText("ERY");
	  GetDlgItem(IDC_CHECK6)->SetWindowText("GLU");
	  GetDlgItem(IDC_CHECK7)->SetWindowText("PRO");
	  GetDlgItem(IDC_CHECK8)->SetWindowText("PH");
	  GetDlgItem(IDC_CHECK9)->SetWindowText("NIT");
	  GetDlgItem(IDC_CHECK10)->SetWindowText("SG");
      GetDlgItem(IDC_CHECK11)->SetWindowText("VC");
	 
  }
}

void CShowInformation::LoadListContent()
{    
	m_listBox.ResetContent();
	if (m_language==0)
	{
		if(m_check1)m_listBox.AddString("尿胆原");
		if(m_check2)m_listBox.AddString("潜血"); 
		if(m_check3)m_listBox.AddString("胆红素");  
		if(m_check4)m_listBox.AddString("酮体"); 
		if(m_check5)m_listBox.AddString("白细胞");  
		if(m_check6)m_listBox.AddString("葡萄糖");   
		if(m_check7)m_listBox.AddString("蛋白质");  
		if(m_check8)m_listBox.AddString("酸碱度");  
		if(m_check9)m_listBox.AddString("亚硝酸盐");
		if(m_check10)m_listBox.AddString("比重"); 
		if(m_check11)m_listBox.AddString("维生素C"); 
		}
		
		if (m_language==1)
		{
			if(m_check1)m_listBox.AddString("URO");
			if(m_check2)m_listBox.AddString("BLD");
			if(m_check3)m_listBox.AddString("BIL"); 
			if(m_check4)m_listBox.AddString("KET");
			if(m_check5)m_listBox.AddString("ERY"); 
			if(m_check6)m_listBox.AddString("GLU"); 
			if(m_check7)m_listBox.AddString("PRO");
			if(m_check8)m_listBox.AddString("PH"); 
			if(m_check9)m_listBox.AddString("NIT");
			if(m_check10)m_listBox.AddString("SG");
			if(m_check11)m_listBox.AddString("VC");
		}
		
}
