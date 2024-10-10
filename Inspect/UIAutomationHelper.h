#pragma once

#include <windows.h>
#include <UIAutomationClient.h>
#include <atlcomcli.h>
#include <vector>
#include <string>
#include <atltypes.h>


enum UIAUTOMATION_ERROR
{
	UIAE_INVALID_PARAM = -1,
	UIAE_INVALID_HWND = -2,
	UIAE_COM_INIT = -3,
	UIAE_IUIAutomation = -4,
	UIAE_GET_ROOT_ELEMENT = -5,
	UIAE_FIND_CONTROL_BY_PROP_NAME = -6,
	UIAE_CREATE_PROP_COND = -7,
	UIAE_FIND_CONTROL = -8,
	UIAE_GET_CONTROL = -9,
	UIAE_GET_NAME = -10,
	UIAE_GET_CLASS_NAME = -11,
};

class CUIElement
{
public:
	CUIElement* m_pParent = nullptr;
	CUIElement* m_pChild = nullptr;	// 第一个孩子
	CUIElement* m_pNext = nullptr;	// 下面第一个兄弟

	IUIAutomationElement* m_pBindElement = nullptr;

	std::wstring m_strAcceleratorKey;
	std::wstring m_strAccessKey;
	std::wstring m_strAutomationId;
	CRect m_rcBoundingRectangle;
	std::wstring m_strClassName;
	UINT m_ControlType = 0;
	std::wstring m_strFrameworkId;
	BOOL m_bHasKeyboardFocus = FALSE;
	std::wstring m_strHelpText;
	BOOL m_bContentElement = FALSE;
	BOOL m_bControlElement = FALSE;
	BOOL m_bEnabled = FALSE;
	BOOL m_bKeyboardFocusable = FALSE;
	BOOL m_bOffscreen = FALSE;
	BOOL m_bPassword = FALSE;
	BOOL m_bRequiredForForm = FALSE;
	std::wstring m_strItemStatus;
	std::wstring m_strItemType;
	UINT m_LabeledBy = 0;
	std::wstring m_strLocalizedControlType;
	std::wstring m_strName;
	HWND m_hWnd = nullptr;
	UINT m_Orientation;
	UINT m_uPid = 0;
};

// 不支持多线程，所以A线程中的IUIAutomationElement，不能在B线程中使用。
class CUIAutomationHelper
{
public:
	CUIAutomationHelper();
	virtual ~CUIAutomationHelper();

public:
	int Init(HWND hWnd);
	HWND GetHwnd();
	IUIAutomation* GetAutomation();

	int BuildControlTree();
	int BuildContentTree();
	int BuildRawTree();
	void ReleaseUITree();

	int ElementFromPoint(POINT pt, IUIAutomationElement** ppElement);
	int GetElement(LPCWSTR lpszAutomationID, IUIAutomationElement** ppElement);
	// lControlType: UIAutomationClient.h line:1318
	int GetElementByControlType(long lControlType, LPCWSTR lpszText, BOOL bEqual, IUIAutomationElement** ppElement);
	int GetElementsByControlType(long lControlType, LPCWSTR lpszText, BOOL bEqual, std::vector<IUIAutomationElement*> *pElements);

	int CreateElementProp(IUIAutomationElement* pElement, CUIElement** ppElement);

	CUIElement* GetRootElement();

	// 返回指定Item下方与指定Item最靠近的可见Item(按List顺序).
	CUIElement* GetNextElement(CUIElement* pCurElement);

protected:
	int WalkerElement(IUIAutomationElement* pElement, LPARAM lParam, CUIElement* pParent, CUIElement* pPreviousSibling, __out CUIElement** ppElement);
	int BuildTrueTreeRecursive(IUIAutomationTreeWalker* pWalker, IUIAutomationElement* pParentElement, LPARAM lParam, CUIElement* pParent, CUIElement* pPreviousSibling, __out CUIElement** ppNewElement);

protected:
	HWND m_hWndHost = nullptr;
	IUIAutomation *m_pClientUIA = nullptr;
	IUIAutomationElement *m_pRootElement = nullptr;

	CUIElement* m_pElement = nullptr;
};

HRESULT InvokeButton(IUIAutomationElement* pButtonElement);
