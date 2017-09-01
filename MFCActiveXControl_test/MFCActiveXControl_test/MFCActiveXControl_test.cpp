// MFCActiveXControl_test.cpp : CMFCActiveXControl_testApp �� DLL ע���ʵ�֡�

#include "stdafx.h"
#include "MFCActiveXControl_test.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


CMFCActiveXControl_testApp theApp;

const GUID CDECL _tlid = { 0xB850D2C0, 0x7AB5, 0x42CA, { 0x99, 0x71, 0x4C, 0xB2, 0xE2, 0x9C, 0xB7, 0x2A } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;



// CMFCActiveXControl_testApp::InitInstance - DLL ��ʼ��

BOOL CMFCActiveXControl_testApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		// TODO:  �ڴ�������Լ���ģ���ʼ�����롣
	}

	return bInit;
}



// CMFCActiveXControl_testApp::ExitInstance - DLL ��ֹ

int CMFCActiveXControl_testApp::ExitInstance()
{
	// TODO:  �ڴ�������Լ���ģ����ֹ���롣

	return COleControlModule::ExitInstance();
}



// DllRegisterServer - ������ӵ�ϵͳע���

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}



// DllUnregisterServer - �����ϵͳע������Ƴ�

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}
