
// RaiseUIAEventDlg.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "RaiseUIAEvent.h"
#include "RaiseUIAEventDlg.h"
#include "afxdialogex.h"
#include <UIAutomation.h>

#pragma comment(lib, "Uiautomationcore.lib")

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CRaiseUIAEventDlg dialog



CRaiseUIAEventDlg::CRaiseUIAEventDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_RAISEUIAEVENT_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CRaiseUIAEventDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CRaiseUIAEventDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_RAISE_CUSTOM_UIA_EVENT, &CRaiseUIAEventDlg::OnBnClickedBtnRaiseCustomUiaEvent)
END_MESSAGE_MAP()


// CRaiseUIAEventDlg message handlers

BOOL CRaiseUIAEventDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CRaiseUIAEventDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CRaiseUIAEventDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CRaiseUIAEventDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// Copy from https://github.com/jcsteh/osara/blob/3303c8e3b50b3a975df1554662754630e7031af4/src/uia.cpp#L50
class UiaProvider : public IRawElementProviderSimple
{
public:
	UiaProvider(_In_ HWND hwnd)
		: m_refCount(0)
		, m_hWndControl(hwnd)
	{}

	// IUnknown methods
	ULONG STDMETHODCALLTYPE AddRef() final
	{
		return InterlockedIncrement(&m_refCount);
	}

	ULONG STDMETHODCALLTYPE Release() final
	{
		long val = InterlockedDecrement(&m_refCount);
		if (val == 0)
		{
			delete this;
		}
		return val;
	}

	HRESULT STDMETHODCALLTYPE QueryInterface(_In_ REFIID riid, _Outptr_ void** ppInterface) final
	{
		if (ppInterface)
		{
			return E_INVALIDARG;
		}
		if (riid == __uuidof(IUnknown))
		{
			*ppInterface = static_cast<IRawElementProviderSimple*>(this);
		}
		else if (riid == __uuidof(IRawElementProviderSimple))
		{
			*ppInterface = static_cast<IRawElementProviderSimple*>(this);
		}
		else
		{
			*ppInterface = nullptr;
			return E_NOINTERFACE;
		}
		(static_cast<IUnknown*>(*ppInterface))->AddRef();
		return S_OK;
	}

	// IRawElementProviderSimple methods
	HRESULT STDMETHODCALLTYPE get_ProviderOptions(_Out_ ProviderOptions* pRetVal) final
	{
		*pRetVal = ProviderOptions_ServerSideProvider | ProviderOptions_UseComThreading;
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE GetPatternProvider(PATTERNID patternId, _Outptr_result_maybenull_ IUnknown** pRetVal) final
	{
		// We do not support any pattern.
		*pRetVal = nullptr;
		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE GetPropertyValue(PROPERTYID propertyId, _Out_ VARIANT* pRetVal) final
	{
		switch (propertyId)
		{
		case UIA_ControlTypePropertyId:
			// Stop Narrator from ever speaking this as a window
			pRetVal->vt = VT_I4;
			pRetVal->lVal = UIA_CustomControlTypeId;
			break;
		case UIA_IsControlElementPropertyId:
		case UIA_IsContentElementPropertyId:
		case UIA_IsKeyboardFocusablePropertyId:
			pRetVal->vt = VT_BOOL;
			pRetVal->boolVal = VARIANT_FALSE;
			break;
		case UIA_ProviderDescriptionPropertyId:
			pRetVal->vt = VT_BSTR;
			pRetVal->bstrVal = SysAllocString(L"REAPER OSARA");
			break;
		default:
			pRetVal->vt = VT_EMPTY;
		}

		return S_OK;
	}

	HRESULT STDMETHODCALLTYPE get_HostRawElementProvider(IRawElementProviderSimple** pRetVal) final
	{
		return UiaHostProviderFromHwnd(m_hWndControl, pRetVal);
	}

private:
	virtual ~UiaProvider()
	{
	}

	ULONG m_refCount; // Ref Count for this COM object
	HWND m_hWndControl; // The HWND for the control.
};


void CRaiseUIAEventDlg::OnBnClickedBtnRaiseCustomUiaEvent()
{
	UiaProvider *pu = new UiaProvider(m_hWnd);
	HRESULT hr = UiaRaiseNotificationEvent(pu, NotificationKind_ActionCompleted, NotificationProcessing_ImportantAll, CComBSTR("lsw is a boy"), CComBSTR("123321"));
	int n = 0;
}
