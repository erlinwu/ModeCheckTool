#pragma once

// TimetableCitectCtrl.h : CTimetableCitectCtrl ActiveX �ؼ����������


// CTimetableCitectCtrl : �й�ʵ�ֵ���Ϣ������� TimetableCitectCtrl.cpp��

class CTimetableCitectCtrl : public COleControl
{
	DECLARE_DYNCREATE(CTimetableCitectCtrl)

// ���캯��
public:
	CTimetableCitectCtrl();

// ��д
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();

// ʵ��
protected:
	~CTimetableCitectCtrl();

	DECLARE_OLECREATE_EX(CTimetableCitectCtrl)    // �๤���� guid
	DECLARE_OLETYPELIB(CTimetableCitectCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CTimetableCitectCtrl)     // ����ҳ ID
	DECLARE_OLECTLTYPE(CTimetableCitectCtrl)		// �������ƺ�����״̬

// ��Ϣӳ��
	DECLARE_MESSAGE_MAP()

// ����ӳ��
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

// �¼�ӳ��
	DECLARE_EVENT_MAP()

// ���Ⱥ��¼� ID
public:
	enum {
	};
};

