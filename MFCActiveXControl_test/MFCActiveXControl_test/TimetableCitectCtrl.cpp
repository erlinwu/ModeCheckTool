// TimetableCitectCtrl.cpp : CTimetableCitectCtrl ActiveX �ؼ����ʵ�֡�

#include "stdafx.h"
#include "MFCActiveXControl_test.h"
#include "TimetableCitectCtrl.h"
#include "TimetableCitectPropPage.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CTimetableCitectCtrl, COleControl)

// ��Ϣӳ��

BEGIN_MESSAGE_MAP(CTimetableCitectCtrl, COleControl)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()

// ����ӳ��

BEGIN_DISPATCH_MAP(CTimetableCitectCtrl, COleControl)
	DISP_FUNCTION_ID(CTimetableCitectCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

// �¼�ӳ��

BEGIN_EVENT_MAP(CTimetableCitectCtrl, COleControl)
END_EVENT_MAP()

// ����ҳ

// TODO: ������Ҫ��Ӹ�������ҳ�����ס���Ӽ���!
BEGIN_PROPPAGEIDS(CTimetableCitectCtrl, 1)
	PROPPAGEID(CTimetableCitectPropPage::guid)
END_PROPPAGEIDS(CTimetableCitectCtrl)

// ��ʼ���๤���� guid

IMPLEMENT_OLECREATE_EX(CTimetableCitectCtrl, "MFCACTIVEXCONTRO.TimetableCitectCtrl.1",
	0x334fab27, 0xfc50, 0x4fab, 0xb5, 0xc, 0x4d, 0x82, 0x72, 0x71, 0x72, 0x1f)

// ����� ID �Ͱ汾

IMPLEMENT_OLETYPELIB(CTimetableCitectCtrl, _tlid, _wVerMajor, _wVerMinor)

// �ӿ� ID

const IID IID_DMFCActiveXControl_test = { 0xA16C8269, 0x3DC4, 0x4C96, { 0x8A, 0x57, 0x21, 0xC3, 0x9D, 0x1F, 0xE9, 0x75 } };
const IID IID_DMFCActiveXControl_testEvents = { 0x4E5DD1CB, 0xC9D3, 0x4438, { 0x88, 0x59, 0xB6, 0xCA, 0x8, 0x32, 0x49, 0x69 } };

// �ؼ�������Ϣ

static const DWORD _dwMFCActiveXControl_testOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CTimetableCitectCtrl, IDS_MFCACTIVEXCONTROL_TEST, _dwMFCActiveXControl_testOleMisc)

// CTimetableCitectCtrl::CTimetableCitectCtrlFactory::UpdateRegistry -
// ��ӻ��Ƴ� CTimetableCitectCtrl ��ϵͳע�����

BOOL CTimetableCitectCtrl::CTimetableCitectCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO:  ��֤���Ŀؼ��Ƿ���ϵ�Ԫģ���̴߳������
	// �йظ�����Ϣ����ο� MFC ����˵�� 64��
	// ������Ŀؼ������ϵ�Ԫģ�͹�����
	// �����޸����´��룬��������������
	// afxRegApartmentThreading ��Ϊ 0��

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_MFCACTIVEXCONTROL_TEST,
			IDB_MFCACTIVEXCONTROL_TEST,
			afxRegApartmentThreading,
			_dwMFCActiveXControl_testOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


// CTimetableCitectCtrl::CTimetableCitectCtrl - ���캯��

CTimetableCitectCtrl::CTimetableCitectCtrl()
{
	InitializeIIDs(&IID_DMFCActiveXControl_test, &IID_DMFCActiveXControl_testEvents);
	// TODO:  �ڴ˳�ʼ���ؼ���ʵ�����ݡ�
}

// CTimetableCitectCtrl::~CTimetableCitectCtrl - ��������

CTimetableCitectCtrl::~CTimetableCitectCtrl()
{
	// TODO:  �ڴ�����ؼ���ʵ�����ݡ�
}

// CTimetableCitectCtrl::OnDraw - ��ͼ����

void CTimetableCitectCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& /* rcInvalid */)
{
	if (!pdc)
		return;

	// TODO:  �����Լ��Ļ�ͼ�����滻����Ĵ��롣
	pdc->FillRect(rcBounds, CBrush::FromHandle((HBRUSH)GetStockObject(WHITE_BRUSH)));
	pdc->Ellipse(rcBounds);
}

// CTimetableCitectCtrl::DoPropExchange - �־���֧��

void CTimetableCitectCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Ϊÿ���־õ��Զ������Ե��� PX_ ������
}


// CTimetableCitectCtrl::OnResetState - ���ؼ�����ΪĬ��״̬

void CTimetableCitectCtrl::OnResetState()
{
	COleControl::OnResetState();  // ���� DoPropExchange ���ҵ���Ĭ��ֵ

	// TODO:  �ڴ��������������ؼ�״̬��
}


// CTimetableCitectCtrl::AboutBox - ���û���ʾ�����ڡ���

void CTimetableCitectCtrl::AboutBox()
{
	CDialogEx dlgAbout(IDD_ABOUTBOX_MFCACTIVEXCONTROL_TEST);
	dlgAbout.DoModal();
}


// CTimetableCitectCtrl ��Ϣ�������
