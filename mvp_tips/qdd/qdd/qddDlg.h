// qddDlg.h : header file
//

#pragma once
#include "afxwin.h"


// CqddDlg dialog
class CqddDlg : public CDialog
{
// Construction
public:
        CqddDlg(CWnd* pParent = NULL);  // standard constructor

// Dialog Data
        enum { IDD = IDD_QDD_DIALOG };

        protected:
        virtual void DoDataExchange(CDataExchange* pDX);        // DDX/DDV support


// Implementation
protected:
        HICON m_hIcon;

        int CaptionGap;
        int EditGap;

        void LogError(DWORD err);
        BOOL DoQueryDosDevice(LPCTSTR name);
        BOOL AddExpansion(HTREEITEM item, LPCTSTR name);
        void FindFrom(HTREEITEM item, const CString & search);
        
        // Generated message map functions
        virtual BOOL OnInitDialog();
        virtual void OnOK();
        virtual void OnCancel();
        DECLARE_MESSAGE_MAP()
        CEdit c_Device;
        CEdit c_Result;
        CTreeCtrl c_DeviceList;
        CEdit c_RetVal;
        CStatic c_Frame;
        CEdit c_FindString;
        CStatic x_Finder;
protected:
        afx_msg void OnClose();
        afx_msg void OnEnChangeDevice();
        afx_msg void OnSize(UINT nType, int cx, int cy);
        afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
        afx_msg void OnPaint();
        afx_msg HCURSOR OnQueryDragIcon();
        afx_msg void OnNMDblclkDeviceList(NMHDR *pNMHDR, LRESULT *pResult);
        afx_msg void OnBnClickedClear();
        afx_msg void OnBnClickedExpandAll();
        afx_msg void OnGetMinMaxInfo(MINMAXINFO* lpMMI);
        afx_msg void OnBnClickedFindNext();
        afx_msg void OnEnChangeFindString();
        afx_msg void OnBnClickedFindPrev();
        afx_msg void OnBnClickedHome();
};
