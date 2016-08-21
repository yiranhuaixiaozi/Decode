// DictionarySet.cpp : implementation file
//

#include "stdafx.h"
#include "Urinalysis.h"
#include "DictionarySet.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CDictionarySet dialog


CDictionarySet::CDictionarySet(CWnd* pParent /*=NULL*/)
	: CDialog(CDictionarySet::IDD, pParent)
{
	//{{AFX_DATA_INIT(CDictionarySet)
	m_insertValue = _T("");
	//}}AFX_DATA_INIT
}


void CDictionarySet::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CDictionarySet)
	DDX_Control(pDX, IDC_LIST2, m_DataValue);
	DDX_Control(pDX, IDC_LIST1, m_DataDictionary);
	DDX_Text(pDX, IDC_EDIT1, m_insertValue);
	//}}AFX_DATA_MAP
}


BEGIN_MESSAGE_MAP(CDictionarySet, CDialog)
	//{{AFX_MSG_MAP(CDictionarySet)
	ON_WM_DESTROY()
	ON_LBN_SELCHANGE(IDC_LIST1, OnSelchangeList1)
	ON_BN_CLICKED(IDC_DICTIONARYADD, OnDictionaryadd)
	ON_BN_CLICKED(IDC_DICTIONARYDEL, OnDictionarydel)
	ON_LBN_SELCHANGE(IDC_LIST2, OnSelchangeList2)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


void CDictionarySet::OnInitADOConn()   //开启数据库连接
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



void CDictionarySet::ExitConnect()     //关闭数据库连接
{
	if(m_pRecordset!=NULL)
		m_pRecordset->Close();
	m_pConnection->Close();
}
/////////////////////////////////////////////////////////////////////////////
// CDictionarySet message handlers

BOOL CDictionarySet::OnInitDialog() 
{
	CDialog::OnInitDialog();
    OnInitADOConn(); 

//      _bstr_t bstrSQL;
//  	bstrSQL = "select distinct DATATYPE1 from tb_Datadictionary";
//  	m_pRecordset.CreateInstance(__uuidof(Recordset));
//  	m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
// 	while(!m_pRecordset->adoEOF)
// 	{   
// 		m_DataDictionary.AddString((char*)(_bstr_t)m_pRecordset->GetCollect("DATATYPE1"));
// 		
// 		m_pRecordset->MoveNext();
//  	}
	m_DataDictionary.AddString("样本类型");
	m_DataDictionary.AddString("送检医生");
	m_DataDictionary.AddString("科室");
	m_DataDictionary.AddString("诊断");


	GetDlgItem(IDC_DICTIONARYADD)->EnableWindow(FALSE);

    GetDlgItem(IDC_DICTIONARYDEL)->EnableWindow(FALSE);
	
	return TRUE;  // return TRUE unless you set the focus to a control
	              // EXCEPTION: OCX Property Pages should return FALSE
}

void CDictionarySet::OnDestroy() 
{

    ExitConnect();
	CDialog::OnDestroy();
	
	// TODO: Add your message handler code here
	
}
 
void CDictionarySet::OnSelchangeList1()      
{
	// TODO: Add your control notification handler code here
     m_DataValue.ResetContent();
     GetDlgItem(IDC_DICTIONARYADD)->EnableWindow(TRUE);
	m_DataDictionary.GetText(m_DataDictionary.GetCurSel(),curSelText);
	_bstr_t bstrSQL;
	bstrSQL = "select VALUE1 from tb_Datadictionary where DATATYPE1='"+curSelText+"'";
   	m_pRecordset.CreateInstance(__uuidof(Recordset));
 	m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,adLockOptimistic,adCmdText);
	while(!m_pRecordset->adoEOF)
	{   
	m_DataValue.AddString((char*)(_bstr_t)m_pRecordset->GetCollect("VALUE1"));
	m_pRecordset->MoveNext();
 	}
	
}

void CDictionarySet::OnDictionaryadd()    //添加按钮响应

{
	// TODO: Add your control notification handler code here
	 UpdateData(TRUE);
	if (m_insertValue=="")
	{MessageBox("输入不能为空");
	return;
	}
  
	_variant_t RecordsAffected;
	CString strSQL; 
	strSQL.Format("INSERT INTO tb_Datadictionary(DATATYPE1,VALUE1) VALUES ('%s','%s')",curSelText,m_insertValue); 
	m_pConnection->Execute((_bstr_t)strSQL,&RecordsAffected,adCmdText); //添加记录
	if (RecordsAffected)
	{   
	m_DataValue.AddString(m_insertValue);
	m_insertValue="";
	UpdateData(FALSE);
	}


}

void CDictionarySet::OnDictionarydel()   //删除按钮响应
{
    UpdateData(TRUE);
	_variant_t RecordsAffected;
	CString strSQL; 
	strSQL.Format("DELETE FROM tb_Datadictionary WHERE VALUE1='"+m_insertValue+"'"); 
	m_pConnection->Execute((_bstr_t)strSQL,&RecordsAffected,adCmdText); 
	m_DataValue.DeleteString(m_DataValue.GetCurSel());
	m_insertValue="";
    GetDlgItem(IDC_DICTIONARYDEL)->EnableWindow(FALSE);
		UpdateData(FALSE);

	
}

void CDictionarySet::OnSelchangeList2() 
{
	// TODO: Add your control notification handler code here
   GetDlgItem(IDC_DICTIONARYDEL)->EnableWindow(TRUE);
   m_DataValue.GetText(m_DataValue.GetCurSel(),m_insertValue);
   UpdateData(FALSE);
}
