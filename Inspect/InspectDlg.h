
// InspectDlg.h : header file
//

#pragma once

#include "UIAutomationHelper.h"
#include <mutex>


#define UIA_BUILD_UIA					WM_USER + 1
#define UIA_GET_ELEMENT_PROP			WM_USER + 2

// CInspectDlg dialog
class CInspectDlg : public CDialogEx
{
	friend LRESULT OnUIAThreadMsg(HWND, UINT uMsg, WPARAM wParam, LPARAM lParam);

// Construction
public:
	CInspectDlg(CWnd* pParent = nullptr);	// standard constructor

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_INSPECT_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

protected:
	int WalkerUITree(CUINode*pUINode,
		HTREEITEM hItemParent,
		HTREEITEM hItemPreviousSibling,
		__out HTREEITEM* phItem);
	int BuildUITreeThread();

public:
	CTreeCtrl m_treUIA;
	CListCtrl m_lstElementProp;
	CUIAutomationHelper m_UIAHelper;

public:
	afx_msg void OnBnClickedBtnBuildUiTree();
	afx_msg void OnTvnSelchangedTreUia(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnClose();
};
