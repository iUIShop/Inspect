﻿#include "UIAutomationHelper.h"
#include <atlstr.h>
#include <list>
#include <mutex>
#include <UIAutomation.h>
#include <UIAutomationClient.h>

#pragma comment(lib, "UIAutomationCore.lib")


class CMsgQueue
{
	bool IsEmpty() const
	{
		return m_queue.empty();
	}

public:
	CMsgQueue()
	{

	}

	void Put(const MSG& msg)
	{
		std::unique_lock<std::mutex> locker(m_mutex);

		m_queue.push_back(msg);
		m_notEmpty.notify_one(); // 再次拥有mutex
	} // unique_lock释放mutex

	void Take(MSG& msg)
	{
		// 如果队列为空，就不能取数据，线程将等待，等待插入数据的线程发出不为空的通知时，
		// 本线程被唤醒，数据被取走。
		// 注意，Take和Put几乎不在同一个线程中被调用。
		std::unique_lock<std::mutex> locker(m_mutex);
		m_notEmpty.wait(locker, [this] {return !IsEmpty(); });

		msg = m_queue.front();
		m_queue.pop_front();
	}

	bool Empty()
	{
		std::lock_guard<std::mutex> locker(m_mutex);
		return m_queue.empty();
	}

	size_t Size()
	{
		std::lock_guard<std::mutex> locker(m_mutex);
		return m_queue.size();
	}

private:
	// 一个消息队列
	std::list<MSG> m_queue;
	std::mutex m_mutex;	// 保护m_queue
	std::condition_variable m_notEmpty; // 队列不空的条件变量
};

CMsgQueue g_MsgQueue;
std::mutex g_mutexSendMsg;
std::condition_variable g_cvSendMsg; // 等待消息处理完成
bool g_bWaitSendMsg = false;
LRESULT g_lr = 0;	// SendMsg返回值

LRESULT SendMsg(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	MSG msg;
	msg.message = uMsg;
	msg.wParam = wParam;
	msg.lParam = lParam;

	g_MsgQueue.Put(msg);

	std::unique_lock<std::mutex> locker(g_mutexSendMsg);
	g_bWaitSendMsg = false;
	g_cvSendMsg.wait(locker, [] {return g_bWaitSendMsg; });

	return g_lr;
}

LRESULT PostMsg(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	MSG msg;
	msg.message = uMsg;
	msg.wParam = wParam;
	msg.lParam = lParam;

	g_MsgQueue.Put(msg);

	return 0;
}

BOOL GetMsg(__out LPMSG lpMsg, WNDPROC fnMsgHandler)
{
	MSG msg;
	g_MsgQueue.Take(msg);

	*lpMsg = msg;

	g_lr = fnMsgHandler(nullptr, msg.message, msg.wParam, msg.lParam);

	// 当hwnd有值时，模拟同步消息，没有值时，模拟异步消息
	g_bWaitSendMsg = true;
	g_cvSendMsg.notify_one();

	if (lpMsg->message == WM_QUIT)
	{
		return FALSE;
	}

	return TRUE;
}

// UIA_ButtonControlTypeId
// <div>aaa</div>: UIA_TextControlTypeId
// <img id="IDC_IMG_EMPTY" alt="" class="empty-pic">: UIA_PaneControlTypeId
HRESULT InvokeButton(IUIAutomationElement* pButtonElement)
{
	HRESULT hr = S_OK;

	// 声明指向InvokePattern接口的指针  
	IUIAutomationInvokePattern* pInvokePattern = nullptr;

	// 检查元素是否支持InvokePattern  
	hr = pButtonElement->GetCurrentPattern(UIA_InvokePatternId, reinterpret_cast<IUnknown**>(&pInvokePattern));
	if (FAILED(hr) || !pInvokePattern)
	{
		// 错误处理：元素不支持InvokePattern  
		if (pInvokePattern)
		{
			pInvokePattern->Release(); // 如果GetCurrentPattern成功但返回了空指针，则不需要释放  
		}
		return hr;
	}

	// 调用Invoke方法来模拟点击  
	hr = pInvokePattern->Invoke();
	if (FAILED(hr))
	{
		// 错误处理：Invoke失败  
	}

	// 释放资源  
	pInvokePattern->Release();

	return hr;
}

CUIAutomationHelper::CUIAutomationHelper()
{
}

CUIAutomationHelper::~CUIAutomationHelper()
{
}

int CUIAutomationHelper::Init(HWND hWndHost)
{
	m_hWndHost = hWndHost;

	int nRet = 0;
	HRESULT hr = S_OK;

	do
	{
		if (nullptr != m_hWndHost && !IsWindow(m_hWndHost))
		{
			nRet = UIAE_INVALID_HWND;
			break;
		}

		hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation),
			reinterpret_cast<void**>(&m_pClientUIA));
		if (FAILED(hr))
		{
			nRet = UIAE_IUIAutomation;
			break;
		}

		if (nullptr == m_hWndHost)
		{
			// 得到桌面元素
			hr = m_pClientUIA->GetRootElement(&m_pRootElement);
		}
		else
		{
			hr = m_pClientUIA->ElementFromHandle(m_hWndHost, &m_pRootElement);
		}
		if (FAILED(hr))
		{
			nRet = UIAE_GET_ROOT_ELEMENT;
			break;
		}

	} while (false);

	return nRet;
}

HWND CUIAutomationHelper::GetHwnd()
{
	return m_hWndHost;
}

int CUIAutomationHelper::WalkerElement(IUIAutomationElement* pElement, LPARAM lParam, CUINode* pParent, CUINode* pPreviousSibling, __out CUINode** ppUINode)
{
	CreateElementProp(pElement, ppUINode);

	(*ppUINode)->m_pParent = pParent;

	if (nullptr == pParent)
	{
		m_pRootNode = *ppUINode;
	}
	else
	{
		if (nullptr == pPreviousSibling)
		{
			pParent->m_pChild = *ppUINode;
		}
		else
		{
			pPreviousSibling->m_pNext = *ppUINode;
		}
	}

	return 0;
}

int CUIAutomationHelper::BuildTrueTreeRecursive(IUIAutomationTreeWalker* pWalker, IUIAutomationElement* pCurElement, LPARAM lParam,
	CUINode* pParent, CUINode* pPreviousSibling, __out CUINode** ppUINode)
{
	HRESULT hr = S_OK;
	int nRet = 0;

	do
	{
		// 自已
		CUINode* pUINode = nullptr;
		WalkerElement(pCurElement, lParam, pParent, pPreviousSibling, &pUINode);
		if (nullptr != ppUINode)
		{
			*ppUINode = pUINode;
		}

		// 所有儿子
		IUIAutomationElement* pChild = nullptr;
		hr = pWalker->GetFirstChildElement(pCurElement, &pChild);
		if (FAILED(hr))
		{
			nRet = -2;
			break;
		}

		CUINode* pChildPreviousSibling = nullptr;
		while (nullptr != pChild)
		{
			CUINode* pNewElement = nullptr;
			BuildTrueTreeRecursive(pWalker, pChild, lParam, pUINode, pChildPreviousSibling, &pNewElement);
			pChildPreviousSibling = pNewElement;

			IUIAutomationElement* pNext = nullptr;
			pWalker->GetNextSiblingElement(pChild, &pNext);
			pChild = pNext;
		}
	} while (false);

	return nRet;
}

int CUIAutomationHelper::BuildRawTree()
{
	HRESULT hr = S_OK;
	int nRet = 0;

	do
	{
		IUIAutomationTreeWalker* pWalker = nullptr;
		hr = m_pClientUIA->get_RawViewWalker(&pWalker);
		if (FAILED(hr))
		{
			nRet = -2;
			break;
		}

		CUINode* pRootElement = nullptr;
		nRet = BuildTrueTreeRecursive(pWalker, m_pRootElement, LPARAM(this), nullptr, nullptr, &pRootElement);
		_ASSERT(pRootElement == m_pRootNode);
	} while (false);

	return nRet;
}

void CUIAutomationHelper::Release()
{
	std::vector<CUINode*> vElements;

	CUINode* pUINode = m_pRootNode;
	while (nullptr != pUINode)
	{
		vElements.push_back(pUINode);

		pUINode = GetNextElement(pUINode);
	}

	for (auto& ele : vElements)
	{
		delete ele;
	}
	vElements = std::vector<CUINode*>();

	m_pRootNode = nullptr;
	m_hWndHost = nullptr;
	if (nullptr != m_pRootElement)
	{
		m_pRootElement->Release();
	}
	if (nullptr != m_pClientUIA)
	{
		m_pClientUIA->Release();
	}
}

int CUIAutomationHelper::ElementFromPoint(POINT pt, IUIAutomationElement** ppUINode)
{
	return -1;
}

// 由于UI上的元素有可能是动态生成和销毁的，所以我们不能使用保存的UI树来查询
// 而应该是每次查询前实时获取一次。
int CUIAutomationHelper::GetElement(LPCWSTR lpszElementID, IUIAutomationElement** ppUINode)
{
	Release();

	return -1;
}

// 从已有的UI树中查询
int CUIAutomationHelper::GetCacheElement(LPCWSTR lpszAutomationID, IUIAutomationElement** ppUINode)
{
	CUINode* pUINode = m_pRootNode;
	while (nullptr != pUINode)
	{
		if (pUINode->m_strAutomationId == lpszAutomationID)
		{
			if (*ppUINode != nullptr)
			{
				*ppUINode = pUINode->m_pBindElement;
			}
			break;
		}

		pUINode = GetNextElement(pUINode);
	}
	return 0;
}

int CUIAutomationHelper::GetElementByControlType(long lControlType, LPCWSTR lpszText, BOOL bEqual, IUIAutomationElement** ppUINode)
{

	return -1;
}

int CUIAutomationHelper::GetElementsByControlType(long lControlType, LPCWSTR lpszText, BOOL bEqual, std::vector<IUIAutomationElement*>* pElements)
{
	return -1;
}

int CUIAutomationHelper::CreateElementProp(IUIAutomationElement* pElement, CUINode** ppUINode)
{
	HRESULT hr = S_OK;
	int nRet = 0;

	do
	{
		if (nullptr == pElement || nullptr == ppUINode)
		{
			nRet = UIAE_INVALID_PARAM;
			break;
		}

		// 先只加载name和控件类型属性用于显示，其它属性等用到的时候再加载。
		// title
		BSTR bstrName;
		hr = pElement->get_CurrentName(&bstrName);
		if (FAILED(hr))
		{
			nRet = UIAE_GET_NAME;
			break;
		}

		BSTR bstrLocalizedControlType;
		hr = pElement->get_CurrentLocalizedControlType(&bstrLocalizedControlType);
		if (FAILED(hr))
		{
			nRet = UIAE_GET_NAME;
			break;
		}

		*ppUINode = new CUINode;
		(*ppUINode)->m_pBindElement = pElement;
		(*ppUINode)->m_strName = bstrName == nullptr ? L"" : bstrName;
		(*ppUINode)->m_strLocalizedControlType = bstrLocalizedControlType == nullptr ? L"" : bstrLocalizedControlType;

		if (0)
		{
			CONTROLTYPEID eType;
			pElement->get_CurrentControlType(&eType);
			if (eType == UIA_GroupControlTypeId)
			{
				int n = 0;
			}

			BSTR bstrAutomationId;
			hr = pElement->get_CurrentAutomationId(&bstrAutomationId);
			if (FAILED(hr))
			{
				nRet = UIAE_GET_NAME;
				break;
			}

			BSTR bstrItemType;
			hr = pElement->get_CurrentItemType(&bstrItemType);
			if (FAILED(hr))
			{
				nRet = UIAE_GET_NAME;
				break;
			}

			BSTR bstrFrameworkId;
			hr = pElement->get_CurrentFrameworkId(&bstrFrameworkId);
			if (FAILED(hr))
			{
				nRet = UIAE_GET_NAME;
				break;
			}

			BSTR bstrItemStatus;
			hr = pElement->get_CurrentItemStatus(&bstrItemStatus);
			if (FAILED(hr))
			{
				nRet = UIAE_GET_NAME;
				break;
			}

			BSTR bstrAriaProperties;
			hr = pElement->get_CurrentAriaProperties(&bstrAriaProperties);
			if (FAILED(hr))
			{
				nRet = UIAE_GET_NAME;
				break;
			}

			BSTR bstrProviderDescription;
			hr = pElement->get_CurrentProviderDescription(&bstrProviderDescription);
			if (FAILED(hr))
			{
				nRet = UIAE_GET_NAME;
				break;
			}

			BSTR bstrClassName;
			hr = pElement->get_CurrentClassName(&bstrClassName);
			if (FAILED(hr))
			{
				nRet = UIAE_GET_CLASS_NAME;
				break;
			}

			BSTR bstrHelpText;
			hr = pElement->get_CurrentHelpText(&bstrHelpText);
			if (FAILED(hr))
			{
				nRet = UIAE_GET_CLASS_NAME;
				break;
			}

			(*ppUINode)->m_strAutomationId = bstrAutomationId == nullptr ? L"" : bstrAutomationId;
			(*ppUINode)->m_ControlType = eType;
			(*ppUINode)->m_strFrameworkId = bstrFrameworkId == nullptr ? L"" : bstrFrameworkId;
			(*ppUINode)->m_strItemType = bstrItemType == nullptr ? L"" : bstrItemType;
			(*ppUINode)->m_strItemStatus = bstrItemStatus == nullptr ? L"" : bstrItemStatus;
			(*ppUINode)->m_strClassName = bstrClassName == nullptr ? L"" : bstrClassName;
			(*ppUINode)->m_strHelpText = bstrHelpText == nullptr ? L"" : bstrHelpText;
		}
	} while (false);

	return nRet;
}

CUINode* CUIAutomationHelper::GetRootUINode()
{
	return m_pRootNode;
}

CUINode* CUIAutomationHelper::GetNextElement(CUINode* pCurElement)
{
	// 如果Item有孩子，并且Item是展开的，则返回第一个孩子
	if (pCurElement->m_pChild != nullptr)
	{
		return pCurElement->m_pChild;
	}

checkNext:
	// 如果Item有下一个兄弟Item，则返回下一个兄弟Item
	if (pCurElement->m_pNext != nullptr)
	{
		return pCurElement->m_pNext;
	}

	// 向上沿线查找
	pCurElement = pCurElement->m_pParent;
	if (pCurElement != nullptr)
	{
		goto checkNext;
	}

	return nullptr;
}

int CUINode::InitProp()
{
	HRESULT hr = S_OK;
	int nRet = 0;

	do
	{
		if (nullptr == m_pBindElement)
		{
			nRet = UIAE_INVALID_PARAM;
			break;
		}

		CONTROLTYPEID eType;
		m_pBindElement->get_CurrentControlType(&eType);
		if (eType == UIA_GroupControlTypeId)
		{
			int n = 0;
		}

		BSTR bstrAutomationId;
		hr = m_pBindElement->get_CurrentAutomationId(&bstrAutomationId);
		if (FAILED(hr))
		{
			nRet = UIAE_GET_NAME;
			break;
		}

		BSTR bstrItemType;
		hr = m_pBindElement->get_CurrentItemType(&bstrItemType);
		if (FAILED(hr))
		{
			nRet = UIAE_GET_NAME;
			break;
		}

		BSTR bstrFrameworkId;
		hr = m_pBindElement->get_CurrentFrameworkId(&bstrFrameworkId);
		if (FAILED(hr))
		{
			nRet = UIAE_GET_NAME;
			break;
		}

		BSTR bstrItemStatus;
		hr = m_pBindElement->get_CurrentItemStatus(&bstrItemStatus);
		if (FAILED(hr))
		{
			nRet = UIAE_GET_NAME;
			break;
		}

		BSTR bstrAriaProperties;
		hr = m_pBindElement->get_CurrentAriaProperties(&bstrAriaProperties);
		if (FAILED(hr))
		{
			nRet = UIAE_GET_NAME;
			break;
		}

		BSTR bstrProviderDescription;
		hr = m_pBindElement->get_CurrentProviderDescription(&bstrProviderDescription);
		if (FAILED(hr))
		{
			nRet = UIAE_GET_NAME;
			break;
		}

		BSTR bstrClassName;
		hr = m_pBindElement->get_CurrentClassName(&bstrClassName);
		if (FAILED(hr))
		{
			nRet = UIAE_GET_CLASS_NAME;
			break;
		}

		BSTR bstrHelpText;
		hr = m_pBindElement->get_CurrentHelpText(&bstrHelpText);
		if (FAILED(hr))
		{
			nRet = UIAE_GET_CLASS_NAME;
			break;
		}

		m_strAutomationId = bstrAutomationId == nullptr ? L"" : bstrAutomationId;
		m_ControlType = eType;
		m_strFrameworkId = bstrFrameworkId == nullptr ? L"" : bstrFrameworkId;
		m_strItemType = bstrItemType == nullptr ? L"" : bstrItemType;
		m_strItemStatus = bstrItemStatus == nullptr ? L"" : bstrItemStatus;
		m_strClassName = bstrClassName == nullptr ? L"" : bstrClassName;
		m_strHelpText = bstrHelpText == nullptr ? L"" : bstrHelpText;

		m_bInitProp = TRUE;

	} while (false);

	return nRet;
}
