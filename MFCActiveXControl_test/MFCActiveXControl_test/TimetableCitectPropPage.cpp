// TimetableCitectPropPage.cpp : CTimetableCitectPropPage ����ҳ���ʵ�֡�

#include "stdafx.h"
#include "MFCActiveXControl_test.h"
#include "TimetableCitectPropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CTimetableCitectPropPage, COlePropertyPage)

// ��Ϣӳ��

BEGIN_MESSAGE_MAP(CTimetableCitectPropPage, COlePropertyPage)
END_MESSAGE_MAP()

// ��ʼ���๤���� guid

IMPLEMENT_OLECREATE_EX(CTimetableCitectPropPage, "MFCACTIVEXCONT.TimetableCitecPropPage.1",
	0x96a740c4, 0x5f99, 0x4067, 0x96, 0x19, 0x7d, 0x2f, 0xc0, 0x6d, 0xb2, 0x83)

// CTimetableCitectPropPage::CTimetableCitectPropPageFactory::UpdateRegistry -
// ��ӻ��Ƴ� CTimetableCitectPropPage ��ϵͳע�����

BOOL CTimetableCitectPropPage::CTimetableCitectPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_MFCACTIVEXCONTROL_TEST_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}

// CTimetableCitectPropPage::CTimetableCitectPropPage - ���캯��

CTimetableCitectPropPage::CTimetableCitectPropPage() :
	COlePropertyPage(IDD, IDS_MFCACTIVEXCONTROL_TEST_PPG_CAPTION)
{
}

// CTimetableCitectPropPage::DoDataExchange - ��ҳ�����Լ��ƶ�����

void CTimetableCitectPropPage::DoDataExchange(CDataExchange* pDX)
{
	DDP_PostProcessing(pDX);
}

// CTimetableCitectPropPage ��Ϣ�������
