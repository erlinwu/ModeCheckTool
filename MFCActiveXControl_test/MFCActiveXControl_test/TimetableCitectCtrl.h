#pragma once

// TimetableCitectCtrl.h : CTimetableCitectCtrl ActiveX 控件类的声明。


// CTimetableCitectCtrl : 有关实现的信息，请参阅 TimetableCitectCtrl.cpp。

class CTimetableCitectCtrl : public COleControl
{
	DECLARE_DYNCREATE(CTimetableCitectCtrl)

// 构造函数
public:
	CTimetableCitectCtrl();

// 重写
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();

// 实现
protected:
	~CTimetableCitectCtrl();

	DECLARE_OLECREATE_EX(CTimetableCitectCtrl)    // 类工厂和 guid
	DECLARE_OLETYPELIB(CTimetableCitectCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CTimetableCitectCtrl)     // 属性页 ID
	DECLARE_OLECTLTYPE(CTimetableCitectCtrl)		// 类型名称和杂项状态

// 消息映射
	DECLARE_MESSAGE_MAP()

// 调度映射
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

// 事件映射
	DECLARE_EVENT_MAP()

// 调度和事件 ID
public:
	enum {
	};
};

