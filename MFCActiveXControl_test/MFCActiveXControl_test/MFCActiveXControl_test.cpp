// MFCActiveXControl_test.cpp : CMFCActiveXControl_testApp 和 DLL 注册的实现。

#include "stdafx.h"
#include "MFCActiveXControl_test.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


CMFCActiveXControl_testApp theApp;

const GUID CDECL _tlid = { 0xB850D2C0, 0x7AB5, 0x42CA, { 0x99, 0x71, 0x4C, 0xB2, 0xE2, 0x9C, 0xB7, 0x2A } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;



// CMFCActiveXControl_testApp::InitInstance - DLL 初始化

BOOL CMFCActiveXControl_testApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		// TODO:  在此添加您自己的模块初始化代码。
	}

	return bInit;
}



// CMFCActiveXControl_testApp::ExitInstance - DLL 终止

int CMFCActiveXControl_testApp::ExitInstance()
{
	// TODO:  在此添加您自己的模块终止代码。

	return COleControlModule::ExitInstance();
}



// DllRegisterServer - 将项添加到系统注册表

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}



// DllUnregisterServer - 将项从系统注册表中移除

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}
