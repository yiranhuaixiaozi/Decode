// Urinalysis.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "Urinalysis.h"
#include "UrinalysisDlg.h"
#include "Splash.h"
#include <initguid.h>
#include "Urinalysis_i.c"
#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CUrinalysisApp

BEGIN_MESSAGE_MAP(CUrinalysisApp, CWinApp)
	//{{AFX_MSG_MAP(CUrinalysisApp)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG
	ON_COMMAND(ID_HELP, CWinApp::OnHelp)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CUrinalysisApp construction

CUrinalysisApp::CUrinalysisApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CUrinalysisApp object

CUrinalysisApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CUrinalysisApp initialization

BOOL CUrinalysisApp::InitInstance()
{

	CCommandLineInfo cmdInfo;
	ParseCommandLine(cmdInfo);

	if (cmdInfo.m_bRunEmbedded || cmdInfo.m_bRunAutomated)
	{
		return TRUE;
	}


	CSplashWnd::EnableSplashScreen(cmdInfo.m_bShowSplash);

	if (!InitATL())
		return FALSE;

	AfxEnableControlContainer();

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	//  of your final executable, you should remove from the following
	//  the specific initialization routines you do not need.

#ifdef _AFXDLL
	Enable3dControls();			// Call this when using MFC in a shared DLL
#else
	Enable3dControlsStatic();	// Call this when linking to MFC statically
#endif

	CUrinalysisDlg dlg;
	m_pMainWnd = &dlg;
	////////////////����ע���ȷ�������Ƿ񾭹���װ��δ����װ�ĳ���򲻿�
char data[100];
CRegKey Rek;
DWORD cbA=100;
Rek.Open(HKEY_CURRENT_USER,"HIGHTOP");
////////////////////��ȡ����
Rek.QueryValue(data,"Password",&cbA);
CString password(data);  //��ȡע����е�����
	if (password==_T("houruiqi"))
	{
		int nResponse = dlg.DoModal();
	}else{
       AfxMessageBox("��Ǹ�������谲װ��������");
	}
	
//	if (nResponse == IDOK)
//	{
	
	//}
//	else if (nResponse == IDCANCEL)
//	{
	
//	}

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.

   
	return FALSE;
}



	
CUrinalysisModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
END_OBJECT_MAP()

LONG CUrinalysisModule::Unlock()
{
	AfxOleUnlockApp();
	return 0;
}

LONG CUrinalysisModule::Lock()
{
	AfxOleLockApp();
	return 1;
}
LPCTSTR CUrinalysisModule::FindOneOf(LPCTSTR p1, LPCTSTR p2)
{
	while (*p1 != NULL)
	{
		LPCTSTR p = p2;
		while (*p != NULL)
		{
			if (*p1 == *p)
				return CharNext(p1);
			p = CharNext(p);
		}
		p1++;
	}
	return NULL;
}


int CUrinalysisApp::ExitInstance()
{
	if (m_bATLInited)
	{
		_Module.RevokeClassObjects();
		_Module.Term();
		CoUninitialize();
	}

	return CWinApp::ExitInstance();

}

BOOL CUrinalysisApp::InitATL()
{
	m_bATLInited = TRUE;

#if _WIN32_WINNT >= 0x0400
	HRESULT hRes = CoInitializeEx(NULL, COINIT_MULTITHREADED);
#else
	HRESULT hRes = CoInitialize(NULL);
#endif

	if (FAILED(hRes))
	{
		m_bATLInited = FALSE;
		return FALSE;
	}

	_Module.Init(ObjectMap, AfxGetInstanceHandle());
	_Module.dwThreadID = GetCurrentThreadId();

	LPTSTR lpCmdLine = GetCommandLine(); //this line necessary for _ATL_MIN_CRT
	TCHAR szTokens[] = _T("-/");

	BOOL bRun = TRUE;
	LPCTSTR lpszToken = _Module.FindOneOf(lpCmdLine, szTokens);
	while (lpszToken != NULL)
	{
		if (lstrcmpi(lpszToken, _T("UnregServer"))==0)
		{
			_Module.UpdateRegistryFromResource(IDR_URINALYSIS, FALSE);
			_Module.UnregisterServer(TRUE); //TRUE means typelib is unreg'd
			bRun = FALSE;
			break;
		}
		if (lstrcmpi(lpszToken, _T("RegServer"))==0)
		{
			_Module.UpdateRegistryFromResource(IDR_URINALYSIS, TRUE);
			_Module.RegisterServer(TRUE);
			bRun = FALSE;
			break;
		}
		lpszToken = _Module.FindOneOf(lpszToken, szTokens);
	}

	if (!bRun)
	{
		m_bATLInited = FALSE;
		_Module.Term();
		CoUninitialize();
		return FALSE;
	}

	hRes = _Module.RegisterClassObjects(CLSCTX_LOCAL_SERVER, 
		REGCLS_MULTIPLEUSE);
	if (FAILED(hRes))
	{
		m_bATLInited = FALSE;
		CoUninitialize();
		return FALSE;
	}	

	return TRUE;

}
