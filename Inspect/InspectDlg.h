
// InspectDlg.h : header file
//

#pragma once

#include "UIAutomationHelper.h"
#include <mutex>

// CInspectDlg dialog
class CInspectDlg : public CDialogEx
{
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
	int WalkerUITree(CUIElement *pElement,
		HTREEITEM hItemParent,
		HTREEITEM hItemPreviousSibling,
		__out HTREEITEM* phItem);
	int BuildUITreeThread();
	std::mutex m_mutex;
	std::condition_variable m_cvReloadUITree;
	bool m_bReloadUITree = false;
	CTreeCtrl m_treUIA;
	CListCtrl m_lstElementProp;
	CUIAutomationHelper m_UIAHelper;
	std::atomic<bool> m_bRunning = true;

public:
	afx_msg void OnBnClickedBtnBuildUiTree();
	afx_msg void OnTvnSelchangedTreUia(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnClose();
};
