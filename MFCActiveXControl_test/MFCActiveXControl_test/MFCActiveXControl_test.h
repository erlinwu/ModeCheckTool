#pragma once

// MFCActiveXControl_test.h : MFCActiveXControl_test.DLL ����ͷ�ļ�

#if !defined( __AFXCTL_H__ )
#error "�ڰ������ļ�֮ǰ������afxctl.h��"
#endif

#include "resource.h"       // ������


// CMFCActiveXControl_testApp : �й�ʵ�ֵ���Ϣ������� MFCActiveXControl_test.cpp��

class CMFCActiveXControl_testApp : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

