#pragma once

// TimetableCitectPropPage.h : CTimetableCitectPropPage ����ҳ���������


// CTimetableCitectPropPage : �й�ʵ�ֵ���Ϣ������� TimetableCitectPropPage.cpp��

class CTimetableCitectPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CTimetableCitectPropPage)
	DECLARE_OLECREATE_EX(CTimetableCitectPropPage)

// ���캯��
public:
	CTimetableCitectPropPage();

// �Ի�������
	enum { IDD = IDD_PROPPAGE_MFCACTIVEXCONTROL_TEST };

// ʵ��
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ��Ϣӳ��
protected:
	DECLARE_MESSAGE_MAP()
};

