
// InspectDlg.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "Inspect.h"
#include "InspectDlg.h"
#include "afxdialogex.h"
#include <thread>

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


// CInspectDlg dialog



CInspectDlg::CInspectDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_INSPECT_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CInspectDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TRE_UIA, m_treUIA);
}

BEGIN_MESSAGE_MAP(CInspectDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_BUILD_UI_TREE, &CInspectDlg::OnBnClickedBtnBuildUiTree)
END_MESSAGE_MAP()


// CInspectDlg message handlers

BOOL CInspectDlg::OnInitDialog()
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
	std::thread t(&CInspectDlg::BuildUITreeThread, this);
	t.detach();

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CInspectDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CInspectDlg::OnPaint()
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
HCURSOR CInspectDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

int CInspectDlg::WalkerUITree(const CUIElement* pElement,
	HTREEITEM hItemParent,
	HTREEITEM hItemPreviousSibling,
	__out HTREEITEM *phItem)
{
	if (nullptr == pElement)
	{
		return -1;
	}

	// втря
	CStringW strItem;
	strItem = L"\"";
	strItem += pElement->m_strName.c_str();
	strItem += L"\" ";
	strItem += pElement->m_strLocalizedControlType.c_str();

	if (nullptr == hItemParent)
	{
		*phItem = m_treUIA.InsertItem(strItem);
	}
	else
	{
		*phItem = m_treUIA.InsertItem(strItem, hItemParent, hItemPreviousSibling);
	}

	CUIElement* pChildElement = pElement->m_pChild;
	HTREEITEM hPreviousSibling = nullptr;
	while (nullptr != pChildElement)
	{
		HTREEITEM hNewItem = nullptr;
		WalkerUITree(pChildElement, *phItem, hPreviousSibling, &hNewItem);
		hPreviousSibling = hNewItem;

		pChildElement = pChildElement->m_pNext;
	}

	return 0;
}

int CInspectDlg::BuildUITreeThread()
{
	CoInitialize(nullptr);

	while (true)
	{
		std::unique_lock<std::mutex> locker(m_mutex);
		m_cvReloadUITree.wait(locker, [this] {return m_bReloadUITree; });
		m_bReloadUITree = false;

		CUIAutomationHelper helper;
		helper.Init(nullptr);	// Desktop
		helper.BuildRawTree();

		m_treUIA.DeleteAllItems();
		const CUIElement *pRootElement = helper.GetRootElement();
		HTREEITEM hItem = nullptr;
		WalkerUITree(pRootElement, nullptr, nullptr, &hItem);
	}

	CoUninitialize();

	return 0;
}


void CInspectDlg::OnBnClickedBtnBuildUiTree()
{
	m_bReloadUITree = true;
	m_cvReloadUITree.notify_one();
}
