
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
	int WalkerUITree(const CUIElement *pElement,
		HTREEITEM hItemParent,
		HTREEITEM hItemPreviousSibling,
		__out HTREEITEM* phItem);
	int BuildUITreeThread();
	std::mutex m_mutex;
	std::condition_variable m_cvReloadUITree;
	bool m_bReloadUITree = false;
public:
	afx_msg void OnBnClickedBtnBuildUiTree();
	CTreeCtrl m_treUIA;
};
