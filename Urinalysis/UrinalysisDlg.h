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

	CString receivePortData();        //���������ݴ���󷵻ظ�����
	void sendPortData(CString data);  //���ַ�������󣬷��͵����ڵĺ���
	void InitSerialPort();            //��ʼ�����ڣ�������connectָ��
	void RegisterActiveX(CString s);  //ע��ACTIVEX�ؼ�
	void InitControlContent();        //�������ݿ⽫�����ֵ���������������ؼ���
	void LoadSaveModule();            //���湦��ģ��
	void LoadPrintModule();           //��ӡ����ģ��
	CString InsertSpace(int num);     //��Ϣ��ʾ����в���ָ�������Ŀո񣬵�����ʽ
	CString InitItemField();          //����ı��ؼ�����ʾ���
	void ClearItemContent();          //���ʮһ���������
	void InitRegist();                //ע���ĳ�ʼ���������״ο�������Ŀ¼��
	BOOL m_themeChange;               //�˱�������
	CUrinalysisDlg(CWnd* pParent = NULL);	// standard constructor
    void  SizeWindow();                //�����ؼ���С
	void OnInitADOConn();              //�����ݽ������ӵĺ���
	void ExitConnect();                //�����ݿ�Ͽ����ӵĺ���
	void LoadContent(CStringArray& item,CString language);    //���ݴ�ע����л�õ���Ϣ��������Ŀ��ʾ��ʽ
// Dialog Data
	//{{AFX_DATA(CUrinalysisDlg)
	enum { IDD = IDD_URINALYSIS_DIALOG };
	CStatic	m_timeShow;
	CStringArray patientResult;  //����ʮһ����Ŀ���Ƶ�����
	CSortListCtrl	m_grid;     //list�ؼ��ı���
	CString	m_age;              //����ؼ�
	CString	m_caseNumbe;        //�����ſؼ�
	CString	m_conclusion;       //��Ͽؼ�
	CString	m_department;       //����
	CString	m_doctorName;       //�ͼ�ҽ��
	CString	m_name;             //��������
	CString	m_sampleNumber;     //������
	CString	m_sampleType;       //��������
	int		m_sex;              //�Ա�
	BOOL	m_timeSerach;       //��ʱ���ѯ��ѡ��
	BOOL	m_nameSerach;       //��������ѯ��ѡ��
	CTime	m_timeSer;          //��ʱ���ѯ
	CString	m_nameSer;          //��������ѯ
	/////////ʮһ����ʾ�ı��ؼ�
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
	CMSComm	m_Comm;              //������
	CString	m_time;              //����
	CString	m_time2;             //ʱ��
	//}}AFX_DATA
CString g_LoadString(CString szID);  //����������Դ�ĺ���
CString g_languagePath;              //����ini·��
CString g_curPath;                   //������·��
CString* pItem[11];                  //ʮһ���ı��ؼ���ָ��

CString portReceiveMessage;          //���ػ������ַ����ı���
int monitorCount;                    //�������Ӻ��ܽ�����������
int postbackCount;                   //һ�λش�������
CString postbackNumber;              //�ش��������ַ�����ʾ
 //���һ��ָ��Connection�����ָ��
_ConnectionPtr m_pConnection;
//���һ��ָ��Recordset�����ָ��
	_RecordsetPtr m_pRecordset;
////////////////////����ͼƬ��Դ
IPicture *m_picture;
OLE_XSIZE_HIMETRIC m_width;
OLE_YSIZE_HIMETRIC m_height;
BOOL m_IsShow;

//CBitmap logo;

//��ʼ������
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
