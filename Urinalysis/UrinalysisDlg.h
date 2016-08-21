// UrinalysisDlg.h : header file
//
//{{AFX_INCLUDES()
#include "mscomm.h"1
//}}AFX_INCLUDES

#if !defined(AFX_URINALYSISDLG_H__AF885123_6686_4238_B687_B74D080F7206__INCLUDED_)
#define AFX_URINALYSISDLG_H__AF885123_6686_4238_B687_B74D080F7206__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
//#include "SkinPPWTL.h"  
#include "SortListCtrl.h"
#include "resource.h"
#include "msword9.h"
#include "atlbase.h"

//#include "ColorListCtrl.h"
/////////////////////////////////////////////////////////////////////////////
// CUrinalysisDlg dialog

class CUrinalysisDlg : public CDialog
{
// Construction
public:
	void DateBaseAutoClear(int item);


	CString decideResult(int location,CString result);

	CString receivePortData();        //将串口数据处理后返回给程序
	void sendPortData(CString data);  //将字符串处理后，发送到串口的函数
	void InitSerialPort();            //初始化串口，并发送connect指令
	void RegisterActiveX(CString s);  //注册ACTIVEX控件
	void InitControlContent();        //访问数据库将数据字典内容填充至各个控件中
	void LoadSaveModule();            //保存功能模块
	void LoadPrintModule();           //打印功能模块
	CString InsertSpace(int num);     //信息显示结果中插入指定数量的空格，调整板式
	CString InitItemField();          //填充文本控件，显示结果
	void ClearItemContent();          //清除十一项件的内容
	void InitRegist();                //注册表的初始化，程序首次开启建立目录等
	BOOL m_themeChange;               //此变量无用
	CUrinalysisDlg(CWnd* pParent = NULL);	// standard constructor
    void  SizeWindow();                //调整控件大小
	void OnInitADOConn();              //和数据建立连接的函数
	void ExitConnect();                //和数据库断开连接的函数
	void LoadContent(CStringArray& item,CString language);    //根据从注册表中获得的信息，建立项目显示方式
// Dialog Data
	//{{AFX_DATA(CUrinalysisDlg)
	enum { IDD = IDD_URINALYSIS_DIALOG };
	CStatic	m_timeShow;
	CStringArray patientResult;  //保存十一项项目名称的数组
	CSortListCtrl	m_grid;     //list控件的变量
	CString	m_age;              //年龄控件
	CString	m_caseNumbe;        //病历号控件
	CString	m_conclusion;       //诊断控件
	CString	m_department;       //科室
	CString	m_doctorName;       //送检医生
	CString	m_name;             //病人姓名
	CString	m_sampleNumber;     //样本号
	CString	m_sampleType;       //样本类型
	int		m_sex;              //性别
	BOOL	m_timeSerach;       //按时间查询复选框
	BOOL	m_nameSerach;       //按姓名查询复选框
	CTime	m_timeSer;          //按时间查询
	CString	m_nameSer;          //按姓名查询
	/////////十一项显示文本控件
	CString	m_item2; 
	CString	m_item3;
	CString	m_item4;
	CString	m_item5;
	CString	m_item6;
	CString	m_item7;
	CString	m_item9;
	CString	m_item8;
	CString	m_item10;
	CString	m_item11;
	CString	m_item1;
	CMSComm	m_Comm;              //串口类
	CString	m_time;              //日期
	CString	m_time2;             //时间
	//}}AFX_DATA
CString g_LoadString(CString szID);  //加载文字资源的函数
CString g_languagePath;              //语言ini路径
CString g_curPath;                   //程序主路径
CString* pItem[11];                  //十一个文本控件的指针

CString portReceiveMessage;          //坚守缓冲区字符串的变量
int monitorCount;                    //建立连接后总接收样本条数
int postbackCount;                   //一次回传的条数
CString postbackNumber;              //回传的条数字符串表示
 //添加一个指向Connection对象的指针
_ConnectionPtr m_pConnection;
//添加一个指向Recordset对象的指针
	_RecordsetPtr m_pRecordset;
////////////////////加载图片资源
IPicture *m_picture;
OLE_XSIZE_HIMETRIC m_width;
OLE_YSIZE_HIMETRIC m_height;
BOOL m_IsShow;

//CBitmap logo;

//初始化连接
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CUrinalysisDlg)
	public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CUrinalysisDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnOption();
	afx_msg void OnSerach();
	afx_msg void OnClickList1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnShowinformation();
	afx_msg void OnSaveinformation();
	afx_msg void OnDestroy();
	afx_msg void OnTimer(UINT nIDEvent);
	afx_msg void OnPrint();
	afx_msg void OnDatadictionary();
	afx_msg void OnDeletedata();
	afx_msg void OnClear();
	afx_msg void OnReportset();
	afx_msg void OnClose();
	virtual void OnCancel();
	afx_msg void OnButton1();
	afx_msg void OnOnCommMscomm1();
	afx_msg void OnConnect();
	afx_msg void OnRefreshdata();
	afx_msg void OnSetfocusList1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnInsertitemList1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnItemchangedList1(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnPostback();
	virtual void OnOK();
	DECLARE_EVENTSINK_MAP()
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_URINALYSISDLG_H__AF885123_6686_4238_B687_B74D080F7206__INCLUDED_)
