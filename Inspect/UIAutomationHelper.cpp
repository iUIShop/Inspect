#include "UIAutomationHelper.h"
#include <atlstr.h>
#include <UIAutomation.h>
#include <UIAutomationClient.h>
#pragma comment(lib, "UIAutomationCore.lib")

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

IUIAutomation* CUIAutomationHelper::GetAutomation()
{
	return m_pClientUIA;
}

int CUIAutomationHelper::BuildControlTree()
{
	return -1;
}

int CUIAutomationHelper::BuildContentTree()
{
	return -1;
}

int CUIAutomationHelper::WalkerElement(IUIAutomationElement* pElement, LPARAM lParam, CUIElement* pParent, CUIElement* pPreviousSibling, __out CUIElement **ppElement)
{
	CreateElementProp(pElement, ppElement);

	(*ppElement)->m_pParent = pParent;

	if (nullptr == pParent)
	{
		m_pElement = *ppElement;
	}
	else
	{
		if (nullptr == pPreviousSibling)
		{
			pParent->m_pChild = *ppElement;
		}
		else
		{
			pPreviousSibling->m_pNext = *ppElement;
		}
	}

	return 0;
}

int CUIAutomationHelper::BuildTrueTreeRecursive(IUIAutomationTreeWalker* pWalker, IUIAutomationElement* pCurElement, LPARAM lParam,
	CUIElement *pParent, CUIElement *pPreviousSibling, __out CUIElement **ppNewElement)
{
	HRESULT hr = S_OK;
	int nRet = 0;

	do
	{
		// 自已
		CUIElement* pElement = nullptr;
		WalkerElement(pCurElement, lParam, pParent, pPreviousSibling, &pElement);
		if (nullptr != ppNewElement)
		{
			*ppNewElement = pElement;
		}

		// 所有儿子
		IUIAutomationElement* pChild = nullptr;
		hr = pWalker->GetFirstChildElement(pCurElement, &pChild);
		if (FAILED(hr))
		{
			nRet = -2;
			break;
		}

		CUIElement* pChildPreviousSibling = nullptr;
		while (nullptr != pChild)
		{
			CUIElement* pNewElement = nullptr;
			BuildTrueTreeRecursive(pWalker, pChild, lParam, pElement, pChildPreviousSibling, &pNewElement);
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

		CUIElement* pRootElement = nullptr;
		nRet = BuildTrueTreeRecursive(pWalker, m_pRootElement, LPARAM(this), nullptr, nullptr, &pRootElement);
		_ASSERT(pRootElement == m_pElement);
	} while (false);

	return nRet;
}

void CUIAutomationHelper::ReleaseUITree()
{
	std::vector<CUIElement*> vElements;

	CUIElement* pElement = m_pElement;
	while (nullptr != pElement)
	{
		vElements.push_back(pElement);

		pElement = GetNextElement(pElement);
	}

	for (auto& ele : vElements)
	{
		delete ele;
	}
	vElements = std::vector<CUIElement*>();

	m_pElement = nullptr;
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

int CUIAutomationHelper::ElementFromPoint(POINT pt, IUIAutomationElement** ppElement)
{
	return -1;
}

int CUIAutomationHelper::GetElement(LPCWSTR lpszElementID, IUIAutomationElement** ppElement)
{
	return -1;
}

int CUIAutomationHelper::GetElementByControlType(long lControlType, LPCWSTR lpszText, BOOL bEqual, IUIAutomationElement** ppElement)
{

	return -1;
}

int CUIAutomationHelper::GetElementsByControlType(long lControlType, LPCWSTR lpszText, BOOL bEqual, std::vector<IUIAutomationElement*>* pElements)
{
	return -1;
}

int CUIAutomationHelper::CreateElementProp(IUIAutomationElement* pElement, CUIElement** ppElement)
{
	HRESULT hr = S_OK;
	int nRet = 0;

	do
	{
		if (nullptr == pElement || nullptr == ppElement)
		{
			nRet = UIAE_INVALID_PARAM;
			break;
		}

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

		*ppElement = new CUIElement;
		(*ppElement)->m_pBindElement = pElement;

		(*ppElement)->m_strName = bstrName == nullptr ? L"" : bstrName;
		(*ppElement)->m_strLocalizedControlType = bstrLocalizedControlType == nullptr ? L"" : bstrLocalizedControlType;
		(*ppElement)->m_strAutomationId = bstrAutomationId == nullptr ? L"" : bstrAutomationId;
		(*ppElement)->m_ControlType = eType;
		(*ppElement)->m_strFrameworkId = bstrFrameworkId == nullptr ? L"" : bstrFrameworkId;
		(*ppElement)->m_strItemType = bstrItemType == nullptr ? L"" : bstrItemType;
		(*ppElement)->m_strItemStatus = bstrItemStatus == nullptr ? L"" : bstrItemStatus;
		(*ppElement)->m_strClassName = bstrClassName == nullptr ? L"" : bstrClassName;
		(*ppElement)->m_strHelpText = bstrHelpText == nullptr ? L"" : bstrHelpText;

	} while (false);

	return nRet;
}

CUIElement* CUIAutomationHelper::GetRootElement()
{
	return m_pElement;
}

CUIElement* CUIAutomationHelper::GetNextElement(CUIElement* pCurElement)
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
