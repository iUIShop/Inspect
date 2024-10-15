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

// 响应UI树结构改变事件
class StructureChangedEventHandler : public IUIAutomationStructureChangedEventHandler
{
private:
	LONG _refCount;

public:
	int _eventCount;

	// Constructor.
	StructureChangedEventHandler() : _refCount(1), _eventCount(0)
	{
	}

	// IUnknown methods.
	ULONG STDMETHODCALLTYPE AddRef()
	{
		ULONG ret = InterlockedIncrement(&_refCount);
		return ret;
	}

	ULONG STDMETHODCALLTYPE Release()
	{
		ULONG ret = InterlockedDecrement(&_refCount);
		if (ret == 0)
		{
			delete this;
			return 0;
		}
		return ret;
	}

	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface)
	{
		if (riid == __uuidof(IUnknown))
			*ppInterface = static_cast<IUIAutomationStructureChangedEventHandler*>(this);
		else if (riid == __uuidof(IUIAutomationStructureChangedEventHandler))
			*ppInterface = static_cast<IUIAutomationStructureChangedEventHandler*>(this);
		else
		{
			*ppInterface = NULL;
			return E_NOINTERFACE;
		}
		this->AddRef();
		return S_OK;
	}

	// IUIAutomationStructureChangedEventHandler methods
	HRESULT STDMETHODCALLTYPE HandleStructureChangedEvent(IUIAutomationElement* pSender, StructureChangeType changeType, SAFEARRAY* pRuntimeID);
};

// 响应自定义事件，我们可以在自己的程序中，调用UiaRaiseNotificationEvent发送自定义事件，
// 这样，自动化测试程序就可以收到这个事件了。
class NotifyEventHandler : public IUIAutomationNotificationEventHandler
{
private:
	LONG _refCount;

public:
	int _eventCount;

	// Constructor.
	NotifyEventHandler() : _refCount(1), _eventCount(0)
	{
	}

	// IUnknown methods.
	ULONG STDMETHODCALLTYPE AddRef()
	{
		ULONG ret = InterlockedIncrement(&_refCount);
		return ret;
	}

	ULONG STDMETHODCALLTYPE Release()
	{
		ULONG ret = InterlockedDecrement(&_refCount);
		if (ret == 0)
		{
			delete this;
			return 0;
		}
		return ret;
	}

	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface)
	{
		if (riid == __uuidof(IUnknown))
			*ppInterface = static_cast<IUIAutomationNotificationEventHandler*>(this);
		else if (riid == __uuidof(IUIAutomationNotificationEventHandler))
			*ppInterface = static_cast<IUIAutomationNotificationEventHandler*>(this);
		else
		{
			*ppInterface = NULL;
			return E_NOINTERFACE;
		}
		this->AddRef();
		return S_OK;
	}

	// IUIAutomationStructureChangedEventHandler methods
	virtual HRESULT STDMETHODCALLTYPE HandleNotificationEvent(
		/* [in] */ __RPC__in_opt IUIAutomationElement* sender,
		enum NotificationKind notificationKind,
		enum NotificationProcessing notificationProcessing,
		/* [in] */ __RPC__in BSTR displayString,
		/* [in] */ __RPC__in BSTR activityId);
};

class CUINode
{
public:
	int InitProp();
	// 释放自己及所有后代
	int Release();

	// 返回指定Item下方与指定Item最靠近的可见Item(按List顺序).
	static CUINode* GetNextElement(CUINode* pCurElement);
	static int GetElementProp(IUIAutomationElement* pElement, __out CUINode* pNode);

public:
	CUINode* m_pParent = nullptr;
	CUINode* m_pChild = nullptr;	// 第一个孩子
	CUINode* m_pNext = nullptr;	// 下面第一个兄弟

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

	BOOL m_bInitProp = FALSE;
};

// 不支持多线程，所以A线程中的IUIAutomationElement，不能在B线程中使用。
class CUIAutomationHelper
{
public:
	CUIAutomationHelper();
	virtual ~CUIAutomationHelper();

public:
	// 指定要解析的窗口，如果传入nullptr，表示桌面
	int Init(HWND hWnd);
	HWND GetHwnd();

	// 基于要解析的窗口，构建UIAutomation树
	int BuildRawTree();

	// 释放
	void Release();

	int ElementFromPoint(POINT pt, IUIAutomationElement** ppElement);
	int GetUINode(LPCWSTR lpszAutomationID, CUINode ** ppUINode);
	int GetCacheUINode(LPCWSTR lpszAutomationID, CUINode** ppUINode);
	int GetUINodes(LPCWSTR lpszAutomationID, std::vector<CUINode*> *pvUINodes);
	int GetCacheUINodes(LPCWSTR lpszAutomationID, std::vector<CUINode*>* pvUINodes);

	// lControlType: UIAutomationClient.h line:1318
	int GetElementByControlType(long lControlType, LPCWSTR lpszText, BOOL bEqual, IUIAutomationElement** ppElement);
	int GetElementsByControlType(long lControlType, LPCWSTR lpszText, BOOL bEqual, std::vector<IUIAutomationElement*> *pElements);

	// 基于第一个参数pElement，创建出第二个参数ppUINode来。
	int CreateElementProp(IUIAutomationElement* pElement, CUINode** ppUINode);

	CUINode* GetRootUINode();

	// 注册事件，这样我们就可以收到事件了。
	int RegisterElementStructureChangedEvent(LPCWSTR lpszAutomationId);
	int RegisterNotifyEvent(LPCWSTR lpszAutomationId);

protected:
	// 遍历到pElement的回调。BuildTrueTreeRecursive在遍历到元素后，就会调用WalkerElement
	int WalkerElement(IUIAutomationElement* pElement, LPARAM lParam, CUINode* pParent, CUINode* pPreviousSibling, __out CUINode** ppUINode);
	// 遍历元素pParentElement
	int BuildTrueTreeRecursive(IUIAutomationTreeWalker* pWalker, IUIAutomationElement* pParentElement, LPARAM lParam, CUINode* pParent, CUINode* pPreviousSibling, __out CUINode** ppNewElement);

protected:
	HWND m_hWndHost = nullptr;
	IUIAutomation *m_pClientUIA = nullptr;
	IUIAutomationElement *m_pRootElement = nullptr;

	// UI Automation元素树型结构
	CUINode* m_pRootNode = nullptr;

	// 响应UI结构变化通知
	StructureChangedEventHandler* m_pStructureChangedHandler = nullptr;
	// 响应自定义通知
	NotifyEventHandler* m_pNotifyHandler = nullptr;
};

// 由于UI Automation不支持多线程，并且遍历整个桌面很慢。
// 所以，我们在遍历到IUIAutomationElement后，先不解析IUIAutomationElement
// 等用到IUIAutomationElement的时候再解析。
// 但由于遍历和解析必须在同一个线程中，线程在遍历完后，生成未解析属性的CUINode树
// 等使用CUINode的属性时，解析的那个线程可能早已退出。
// 所以，我们模拟Windows的消息循环，让一个线程，即可以遍历，又可以解析
// 我们必须创建一个消息队列
LRESULT SendMsg(UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT PostMsg(UINT uMsg, WPARAM wParam, LPARAM lParam);
BOOL GetMsg(__out LPMSG lpMsg, WNDPROC fnMsgHandler);

HRESULT InvokeButton(IUIAutomationElement* pButtonElement);
