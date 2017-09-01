// TimetableCitectCtrl.cpp : CTimetableCitectCtrl ActiveX 控件类的实现。

#include "stdafx.h"
#include "MFCActiveXControl_test.h"
#include "TimetableCitectCtrl.h"
#include "TimetableCitectPropPage.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CTimetableCitectCtrl, COleControl)

// 消息映射

BEGIN_MESSAGE_MAP(CTimetableCitectCtrl, COleControl)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()

// 调度映射

BEGIN_DISPATCH_MAP(CTimetableCitectCtrl, COleControl)
	DISP_FUNCTION_ID(CTimetableCitectCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

// 事件映射

BEGIN_EVENT_MAP(CTimetableCitectCtrl, COleControl)
END_EVENT_MAP()

// 属性页

// TODO: 根据需要添加更多属性页。请记住增加计数!
BEGIN_PROPPAGEIDS(CTimetableCitectCtrl, 1)
	PROPPAGEID(CTimetableCitectPropPage::guid)
END_PROPPAGEIDS(CTimetableCitectCtrl)

// 初始化类工厂和 guid

IMPLEMENT_OLECREATE_EX(CTimetableCitectCtrl, "MFCACTIVEXCONTRO.TimetableCitectCtrl.1",
	0x334fab27, 0xfc50, 0x4fab, 0xb5, 0xc, 0x4d, 0x82, 0x72, 0x71, 0x72, 0x1f)

// 键入库 ID 和版本

IMPLEMENT_OLETYPELIB(CTimetableCitectCtrl, _tlid, _wVerMajor, _wVerMinor)

// 接口 ID

const IID IID_DMFCActiveXControl_test = { 0xA16C8269, 0x3DC4, 0x4C96, { 0x8A, 0x57, 0x21, 0xC3, 0x9D, 0x1F, 0xE9, 0x75 } };
const IID IID_DMFCActiveXControl_testEvents = { 0x4E5DD1CB, 0xC9D3, 0x4438, { 0x88, 0x59, 0xB6, 0xCA, 0x8, 0x32, 0x49, 0x69 } };

// 控件类型信息

static const DWORD _dwMFCActiveXControl_testOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CTimetableCitectCtrl, IDS_MFCACTIVEXCONTROL_TEST, _dwMFCActiveXControl_testOleMisc)

// CTimetableCitectCtrl::CTimetableCitectCtrlFactory::UpdateRegistry -
// 添加或移除 CTimetableCitectCtrl 的系统注册表项

BOOL CTimetableCitectCtrl::CTimetableCitectCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO:  验证您的控件是否符合单元模型线程处理规则。
	// 有关更多信息，请参考 MFC 技术说明 64。
	// 如果您的控件不符合单元模型规则，则
	// 必须修改如下代码，将第六个参数从
	// afxRegApartmentThreading 改为 0。

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


// CTimetableCitectCtrl::CTimetableCitectCtrl - 构造函数

CTimetableCitectCtrl::CTimetableCitectCtrl()
{
	InitializeIIDs(&IID_DMFCActiveXControl_test, &IID_DMFCActiveXControl_testEvents);
	// TODO:  在此初始化控件的实例数据。
}

// CTimetableCitectCtrl::~CTimetableCitectCtrl - 析构函数

CTimetableCitectCtrl::~CTimetableCitectCtrl()
{
	// TODO:  在此清理控件的实例数据。
}

// CTimetableCitectCtrl::OnDraw - 绘图函数

void CTimetableCitectCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& /* rcInvalid */)
{
	if (!pdc)
		return;

	// TODO:  用您自己的绘图代码替换下面的代码。
	pdc->FillRect(rcBounds, CBrush::FromHandle((HBRUSH)GetStockObject(WHITE_BRUSH)));
	pdc->Ellipse(rcBounds);
}

// CTimetableCitectCtrl::DoPropExchange - 持久性支持

void CTimetableCitectCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: 为每个持久的自定义属性调用 PX_ 函数。
}


// CTimetableCitectCtrl::OnResetState - 将控件重置为默认状态

void CTimetableCitectCtrl::OnResetState()
{
	COleControl::OnResetState();  // 重置 DoPropExchange 中找到的默认值

	// TODO:  在此重置任意其他控件状态。
}


// CTimetableCitectCtrl::AboutBox - 向用户显示“关于”框

void CTimetableCitectCtrl::AboutBox()
{
	CDialogEx dlgAbout(IDD_ABOUTBOX_MFCACTIVEXCONTROL_TEST);
	dlgAbout.DoModal();
}


// CTimetableCitectCtrl 消息处理程序
