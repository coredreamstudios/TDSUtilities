// qddDlg.cpp : implementation file
//

#include "stdafx.h"
#include "resource.h"
#include "qddDlg.h"
#include "ErrorString.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

/****************************************************************************
*                                About Dialog
****************************************************************************/

// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
 CAboutDlg();

// Dialog Data
        enum { IDD = IDD_ABOUTBOX };

        protected:
        virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
        DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
        CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CqddDlg dialog

/****************************************************************************
*                              CqddDlg::CqddDlg
* Inputs:
*       CWnd * parent:
* Effect: 
*       Constructor
****************************************************************************/

CqddDlg::CqddDlg(CWnd* pParent /*=NULL*/)
        : CDialog(CqddDlg::IDD, pParent)
{
        m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}


/****************************************************************************
*                           CqddDlg::DoDataExchange
* Inputs:
*       CDataExchange * pDX:
* Result: void
*       
* Effect: 
*       Binds controls to variables
****************************************************************************/

void CqddDlg::DoDataExchange(CDataExchange* pDX)
{
CDialog::DoDataExchange(pDX);
DDX_Control(pDX, IDC_DEVICE, c_Device);
DDX_Control(pDX, IDC_RESULT, c_Result);
DDX_Control(pDX, IDC_DEVICE_LIST, c_DeviceList);
DDX_Control(pDX, IDC_RETVAL, c_RetVal);
DDX_Control(pDX, IDC_FRAME, c_Frame);
DDX_Control(pDX, IDC_FIND_STRING, c_FindString);
DDX_Control(pDX, IDC_FIND_CAPTION, x_Finder);
    }

/****************************************************************************
*                                 Message Map
****************************************************************************/

BEGIN_MESSAGE_MAP(CqddDlg, CDialog)
        ON_WM_SYSCOMMAND()
        ON_WM_PAINT()
        ON_WM_QUERYDRAGICON()
        //}}AFX_MSG_MAP
        ON_EN_CHANGE(IDC_DEVICE, &CqddDlg::OnEnChangeDevice)
        ON_WM_SIZE()
        ON_WM_CLOSE()
        ON_NOTIFY(NM_DBLCLK, IDC_DEVICE_LIST, &CqddDlg::OnNMDblclkDeviceList)
        ON_BN_CLICKED(IDC_CLEAR, &CqddDlg::OnBnClickedClear)
        ON_BN_CLICKED(IDC_EXPAND_ALL, &CqddDlg::OnBnClickedExpandAll)
        ON_WM_GETMINMAXINFO()
        ON_BN_CLICKED(IDC_FIND_NEXT, &CqddDlg::OnBnClickedFindNext)
        ON_EN_CHANGE(IDC_FIND_STRING, &CqddDlg::OnEnChangeFindString)
        ON_BN_CLICKED(IDC_FIND_PREV, &CqddDlg::OnBnClickedFindPrev)
        ON_BN_CLICKED(IDC_HOME, &CqddDlg::OnBnClickedHome)
END_MESSAGE_MAP()


// CqddDlg message handlers


/****************************************************************************
*                            CqddDlg::OnInitDialog
* Result: BOOL
*       TRUE, always
* Effect: 
*       Initializes dialog
****************************************************************************/

BOOL CqddDlg::OnInitDialog()
{
        CDialog::OnInitDialog();

        // Add "About..." menu item to system menu.

        // IDM_ABOUTBOX must be in the system command range.
        ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
        ASSERT(IDM_ABOUTBOX < 0xF000);

        CMenu* pSysMenu = GetSystemMenu(FALSE);
        if (pSysMenu != NULL)
        {
                CString strAboutMenu;
                strAboutMenu.LoadString(IDS_ABOUTBOX);
                if (!strAboutMenu.IsEmpty())
                {
                        pSysMenu->AppendMenu(MF_SEPARATOR);
                        pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
                }
        }

        // Set the icon for this dialog.  The framework does this automatically
        //  when the application's main window is not a dialog
        SetIcon(m_hIcon, TRUE);                 // Set big icon
        SetIcon(m_hIcon, FALSE);                // Set small icon

        // TODO: Add extra initialization here

        c_Device.SetWindowText(_T(""));

        CRect client;
        GetClientRect(&client);
        
        CRect caption;
        x_Finder.GetWindowRect(&caption);
        ScreenToClient(&caption);

        CaptionGap = client.right - caption.right;

        CRect edit;
        c_FindString.GetWindowRect(&edit);
        ScreenToClient(&edit);

        EditGap = client.right - edit.right;

        return TRUE;  // return TRUE  unless you set the focus to a control
}


/****************************************************************************
*                            CqddDlg::OnSysCommand
* Inputs:
*       UINT nID: menu ID
*       LPARAM lParam: passed on to superclass
* Result: void
*       
* Effect: 
*       �
****************************************************************************/

void CqddDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
        if ((nID & 0xFFF0) == IDM_ABOUTBOX)
        {
                CAboutDlg dlgAbout;
                dlgAbout.DoModal();
        }
        else
        {
                CDialog::OnSysCommand(nID, lParam);
        }
}


/****************************************************************************
*                               CqddDlg::OnPaint
* Result: void
*       
* Effect: 
*       Useless vestige of MFC 16-bit
****************************************************************************/

void CqddDlg::OnPaint()
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
                CDialog::OnPaint();
        }
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CqddDlg::OnQueryDragIcon()
{
        return static_cast<HCURSOR>(m_hIcon);
}


/****************************************************************************
*                              CqddDlg::LogError
* Inputs:
*       DWORD err: Error code
* Result: void
*       
* Effect: 
*       Logs the error
****************************************************************************/

void CqddDlg::LogError(DWORD err)
    {
     CString s = ErrorString(err);
     c_Result.SetWindowText(s);
    } // CqddDlg::LogError

/****************************************************************************
*                            CqddDlg::AddExpansion
* Inputs:
*       HTREEITEM item: Item to add it under
*       LPCTSR name: Name to expand
* Result: BOOL
*       
* Effect: 
*       Expands one level down
****************************************************************************/

BOOL CqddDlg::AddExpansion(HTREEITEM item, LPCTSTR name)
    {
     DWORD len = 1024;
     LPTSTR buffer = (LPTSTR)malloc(len * sizeof(TCHAR));
     while(TRUE)
        { /* expand values */
         DWORD result = QueryDosDevice(name, buffer, len);
         if(result == 0)
            { /* failed */
             DWORD err = ::GetLastError();
             if(err == ERROR_INSUFFICIENT_BUFFER)
                { /* reallocate */
                 len *= 2;
                 buffer = (LPTSTR)realloc(buffer, len * sizeof(TCHAR));
                 continue;
                } /* reallocate */
             else
                { /* other failure */
                 free(buffer);
                 return FALSE;
                } /* other failure */
            } /* failed */
         break;
        } /* expand values */
     LPTSTR pos = buffer;
     while(_tcslen(pos) > 0)
        { /* add each */
         c_DeviceList.InsertItem(buffer, item);
         pos += _tcslen(pos) + 1;
        } /* add each */
     free(buffer);
     return TRUE;
    } // CqddDlg::AddExpansion

/****************************************************************************
*                          CqddDlg::DoQueryDosDevice
* Inputs:
*       LPCTSTR name: Name of DOS device, or NULL
* Result: BOOL
*       TRUE if successful
*       FALSE if error
* Effect: 
*       Executes QueryDosDevice
* Notes:
*       This uses malloc/realloc so I can simulate what it would look
*       like in C code
****************************************************************************/

BOOL CqddDlg::DoQueryDosDevice(LPCTSTR name)
    {
     LPTSTR devices = NULL;
     DWORD len = 1024;

     devices = (LPTSTR)malloc(len * sizeof(TCHAR));
     if(devices == NULL)
        { /* insufficient space */
         LogError(ERROR_NO_SYSTEM_RESOURCES);
         return FALSE;
        } /* insufficient space */

     while(TRUE)
        { /* read devices */
         
         DWORD result = ::QueryDosDevice(name, devices, len);
         if(result == 0)
            { /* failed */
             DWORD err = ::GetLastError();
             c_RetVal.SetWindowText(_T("0"));
             if(ERROR_INSUFFICIENT_BUFFER == err)
                { /* realloc and retry */
                 len *= 2;
                 devices = (LPTSTR)realloc(devices, len * sizeof(TCHAR));
                 continue;
                } /* realloc and retry */
             else
                { /* serious error */
                 LogError(err);
                 free(devices);
                 ::SetLastError(err);
                 return FALSE;
                } /* serious error */
            } /* failed */
         CString s;
         s.Format(_T("%u"), result);
         c_RetVal.SetWindowText(s);
         break;
        } /* read devices */

     
     LPTSTR pos = devices;
     LogError(ERROR_SUCCESS);
     /*******************************************************************************
       From the documentation:
       
       If lpDeviceName is non-NULL, the function retrieves information about
       the particular MS-DOS device specified by lpDeviceName. The first
       null-terminated string stored into the buffer is the current mapping for
       the device. The other null-terminated strings represent undeleted prior
       mappings for the device.
     *******************************************************************************/

     if(name == NULL)
        { /* show all */
         HTREEITEM item = TVI_ROOT;
         while(_tcslen(pos) != 0)
            { /* scan each */
             HTREEITEM sub = c_DeviceList.InsertItem(pos, item);
             AddExpansion(sub, pos);
             pos += _tcslen(pos) + 1;
            } /* scan each */
        } /* show all */
     else
        { /* show specific */
         HTREEITEM root = c_DeviceList.InsertItem(name, TVI_ROOT);
         while(_tcslen(pos) != 0)
            { /* show each */
             HTREEITEM item = c_DeviceList.InsertItem(pos, root);
             pos += _tcslen(pos) + 1;
            } /* show each */
         c_DeviceList.Expand(root, TVE_EXPAND);
        } /* show specific */

     free(devices);
     return TRUE;
    } // CqddDlg::DoQueryDosDevice

/****************************************************************************
*                          CqddDlg::OnEnChangeDevice
* Result: void
*       
* Effect: 
*       Handles a GetDosDevice request
****************************************************************************/

void CqddDlg::OnEnChangeDevice()
    {
     CString s;
     c_Device.GetWindowText(s);
     s.Trim();
     LPCTSTR p = s.IsEmpty() ? NULL : (LPCTSTR)s;
     c_DeviceList.SetRedraw(FALSE);

     c_DeviceList.DeleteAllItems();
     if(!DoQueryDosDevice(p))
        { /* failed */
        } /* failed */
     else
        { /* succeeded */
        } /* succeeded */
     c_DeviceList.SetRedraw(TRUE);
    }


/****************************************************************************
*                               CqddDlg::OnSize
* Inputs:
*       UINT nType:
*       int cx:
*       int cy:
* Result: void
*       
* Effect: 
*       Resizes the control
****************************************************************************/

void CqddDlg::OnSize(UINT nType, int cx, int cy)
    {
     CDialog::OnSize(nType, cx, cy);

     if(c_DeviceList.GetSafeHwnd() != NULL)
        { /* resize */
         CRect r;
         c_DeviceList.GetWindowRect(&r);
         ScreenToClient(&r);
         c_DeviceList.SetWindowPos(NULL, 0, 0, cx, cy - r.top,
                                   SWP_NOMOVE | SWP_NOZORDER);
        } /* resize */
     if(c_Device.GetSafeHwnd() != NULL)
        { /* resize edit box */
         CRect r;
         c_Device.GetWindowRect(&r);
         ScreenToClient(&r);
         c_Device.SetWindowPos(NULL, 0, 0, cx - r.left, r.Height(),
                               SWP_NOMOVE | SWP_NOZORDER);
        } /* resize edit box */

     if(c_Result.GetSafeHwnd() != NULL)
        { /* resize result */
         CRect r;
         c_Result.GetWindowRect(&r);
         ScreenToClient(&r);
         c_Result.SetWindowPos(NULL, 0, 0, cx - r.left, r.Height(),
                               SWP_NOMOVE | SWP_NOZORDER);
        } /* resize result */

     if(x_Finder.GetSafeHwnd() != NULL)
        { /* resize group box */
         CRect r;
         x_Finder.GetWindowRect(&r);
         ScreenToClient(&r);
         x_Finder.SetWindowPos(NULL, 0, 0, cx - r.left - CaptionGap, r.Height(),
                             SWP_NOMOVE | SWP_NOZORDER);
        } /* resize group box */

     if(c_FindString.GetSafeHwnd() != NULL)
        { /* resize edit */
         CRect r;
         c_FindString.GetWindowRect(&r);
         ScreenToClient(&r);
         c_FindString.SetWindowPos(NULL, 0, 0, cx - r.left - EditGap, r.Height(),
                                   SWP_NOMOVE | SWP_NOZORDER);
        } /* resize edit */
    }

/****************************************************************************
*                                CqddDlg::OnOK
* Result: void
*       
* Effect: 
*       Does nothing
****************************************************************************/

void CqddDlg::OnOK()
    {
     // Does nothing
    } // CqddDlg::OnOK

/****************************************************************************
*                              CqddDlg::OnCancel
* Result: void
*       
* Effect: 
*       Does nothing
****************************************************************************/

void CqddDlg::OnCancel()
    {
     // Does nothing
    } // CqddDlg::OnCancel

/****************************************************************************
*                              CqddDlg::OnClose
* Result: void
*       
* Effect: 
*       Closes the dialog
****************************************************************************/

void CqddDlg::OnClose()
    {
     CDialog::OnOK();
    } // CqddDlg::OnClose


/****************************************************************************
*                        CqddDlg::OnNMDblclkDeviceList
* Inputs:
*       NMHDR * pNMHDR:
*       LRESULT * pResult:
* Result: void
*       
* Effect: 
*       Selects the element specified
****************************************************************************/

void CqddDlg::OnNMDblclkDeviceList(NMHDR *pNMHDR, LRESULT *pResult)
    {
     HTREEITEM item = c_DeviceList.GetSelectedItem();
     if(item == NULL)
        return;
     
     if(c_DeviceList.GetParentItem(item) != NULL)
        return; // ignore subitems

     CString s = c_DeviceList.GetItemText(item);
     c_Device.SetWindowText(s);
     *pResult = 0;
    }

/****************************************************************************
*                          CqddDlg::OnBnClickedClear
* Result: void
*       
* Effect: 
*       Clears the device name (making the call use NULL)
****************************************************************************/

void CqddDlg::OnBnClickedClear()
    {
     c_Device.SetWindowTextW(_T(""));
    }


/****************************************************************************
*                        CqddDlg::OnBnClickedExpandAll
* Result: void
*       
* Effect: 
*       Expands all the elements
****************************************************************************/

void CqddDlg::OnBnClickedExpandAll()
    {
     HTREEITEM sel = c_DeviceList.GetSelectedItem();
     
     for(HTREEITEM item = c_DeviceList.GetRootItem(); item != NULL; item = c_DeviceList.GetNextSiblingItem(item))
         { /* expand all */
          c_DeviceList.Expand(item, TVE_EXPAND);
         } /* expand all */
     if(sel != NULL)
        { /* select item */
         c_DeviceList.SelectItem(sel);
         c_DeviceList.EnsureVisible(sel);
        } /* select item */
     else
        { /* select root */
         c_DeviceList.SelectItem(c_DeviceList.GetRootItem());
         c_DeviceList.EnsureVisible(c_DeviceList.GetRootItem());
        } /* select root */
    }


/****************************************************************************
*                           CqddDlg::OnGetMinMaxInfo
* Inputs:
*       MINMAXINFO * lpMMI:
* Result: void
*       
* Effect: 
*       Stops resizing from hiding controls
****************************************************************************/

void CqddDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
    {
     if(c_Frame.GetSafeHwnd() != NULL)
        { /* can limit */
         CRect r;
         c_Frame.GetWindowRect(&r);
         ScreenToClient(&r);
         CalcWindowRect(&r);
         lpMMI->ptMinTrackSize.x = r.Width();
         lpMMI->ptMinTrackSize.y = r.Height();
         return;
        } /* can limit */

     CDialog::OnGetMinMaxInfo(lpMMI);
    }


/****************************************************************************
*                              CqddDlg::FindFrom
* Inputs:
*       HTREEITEM item: Item to start from
*       const CString & search: Search string
* Result: �
*       �
* Effect: 
*       �
****************************************************************************/

void CqddDlg::FindFrom(HTREEITEM item, const CString & search)
    {
     for(HTREEITEM target = item; target != NULL; target = c_DeviceList.GetNextSiblingItem(target))
        { /* scan */
         CString ts = c_DeviceList.GetItemText(target);
         ts.MakeUpper();
         if(ts.Find(search) >= 0)
            { /* found it */
             c_DeviceList.SelectItem(target);
             c_DeviceList.EnsureVisible(target);
             return;
            } /* found it */
        } /* scan */
     
    } // CqddDlg::FindFrom

/****************************************************************************
*                         CqddDlg::OnBnClickedFindNext
* Result: void
*       
* Effect: 
*       Finds the next item
****************************************************************************/

void CqddDlg::OnBnClickedFindNext()
    {
     HTREEITEM item = c_DeviceList.GetSelectedItem();
     if(item == NULL)
        item = c_DeviceList.GetRootItem();
     if(item == NULL)
        return;

     CString search;
     c_FindString.GetWindowText(search);
     search.MakeUpper();
     
     if(search.IsEmpty())
        return;

     FindFrom(c_DeviceList.GetNextSiblingItem(item), search);
    }

/****************************************************************************
*                           CqddDlg::OnEnChangeFindString
* Result: void
*       
* Effect: 
*       Forces a find
****************************************************************************/

void CqddDlg::OnEnChangeFindString()
    {
     HTREEITEM item = c_DeviceList.GetSelectedItem();
     if(item == NULL)
        item = c_DeviceList.GetRootItem();
     if(item == NULL)
        return;

     CString search;
     c_FindString.GetWindowText(search);
     search.MakeUpper();

     if(search.IsEmpty())
        return;

     FindFrom(item, search);
    }

/****************************************************************************
*                         CqddDlg::OnBnClickedFindPrev
* Result: void
*       
* Effect: 
*       Finds previous item
****************************************************************************/

void CqddDlg::OnBnClickedFindPrev()
    {
     HTREEITEM item = c_DeviceList.GetSelectedItem();
     if(item == NULL)
        item = c_DeviceList.GetRootItem();
     if(item == NULL)
        return;

     CString search;
     c_FindString.GetWindowText(search);
     search.MakeUpper();

     if(search.IsEmpty())
        return;

     for(HTREEITEM target = c_DeviceList.GetPrevSiblingItem(item); target != NULL; target = c_DeviceList.GetPrevSiblingItem(target))
        { /* scan */
         CString ts = c_DeviceList.GetItemText(target);
         ts.MakeUpper();
         if(ts.Find(search) >= 0)
            { /* found it */
             c_DeviceList.SelectItem(target);
             c_DeviceList.EnsureVisible(target);
             return;
            } /* found it */
        } /* scan */
    }

/****************************************************************************
*                           CqddDlg::OnBnClickedHome
* Result: void
*       
* Effect: 
*       Goes to root
****************************************************************************/

void CqddDlg::OnBnClickedHome()
    {
     HTREEITEM item = c_DeviceList.GetRootItem();
     if(item == NULL)
        return;
     c_DeviceList.SelectItem(item);
     c_DeviceList.EnsureVisible(item);
    }
