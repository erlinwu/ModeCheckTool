#pragma once

// TimetableCitectPropPage.h : CTimetableCitectPropPage 属性页类的声明。


// CTimetableCitectPropPage : 有关实现的信息，请参阅 TimetableCitectPropPage.cpp。

class CTimetableCitectPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CTimetableCitectPropPage)
	DECLARE_OLECREATE_EX(CTimetableCitectPropPage)

// 构造函数
public:
	CTimetableCitectPropPage();

// 对话框数据
	enum { IDD = IDD_PROPPAGE_MFCACTIVEXCONTROL_TEST };

// 实现
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 消息映射
protected:
	DECLARE_MESSAGE_MAP()
};

