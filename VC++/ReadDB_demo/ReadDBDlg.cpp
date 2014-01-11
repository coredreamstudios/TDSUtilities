// ReadDBDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ReadDB.h"
#include "ReadDBDlg.h"
#include "afxdb.h"
#include "odbcinst.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CReadDBDlg dialog

CReadDBDlg::CReadDBDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CReadDBDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CReadDBDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CReadDBDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CReadDBDlg)
	DDX_Control(pDX, IDC_ListControl, m_ListControl);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CReadDBDlg, CDialog)
	//{{AFX_MSG_MAP(CReadDBDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_READ, OnRead)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CReadDBDlg message handlers

BOOL CReadDBDlg::OnInitDialog()
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
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CReadDBDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CReadDBDlg::OnPaint() 
{

	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

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

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CReadDBDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CReadDBDlg::OnRead() 
{
	// TODO: Add your control notification handler code here
	CDatabase database;
	CString SqlString;
	CString sCatID, sCategory;
	CString sDriver = "MICROSOFT ACCESS DRIVER (*.mdb)";
	CString sDsn;
	CString sFile = "Test.mdb";
	// You must change above path if it's different
	int iRec = 0; 	
	
	// Build ODBC connection string
	sDsn.Format("ODBC;DRIVER={%s};DSN='';DBQ=%s", sDriver, sFile);

	TRY
	{
		// Open the database
		database.Open(NULL, false, false, sDsn);
		
		// Allocate the recordset
		CRecordset recset( &database );

		// Build the SQL statement
		SqlString =  "SELECT CatID, Category "
				"FROM Categories";

		// Execute the query
		recset.Open(CRecordset::forwardOnly, SqlString, CRecordset::readOnly);
		// Reset List control if there is any data
		ResetListControl();
		// populate Grids
		ListView_SetExtendedListViewStyle(m_ListControl, LVS_EX_GRIDLINES);
 
		// Column width and heading
		m_ListControl.InsertColumn(0, "Category Id", LVCFMT_LEFT, -1, 0);
		m_ListControl.InsertColumn(1, "Category", LVCFMT_LEFT, -1, 1);
		m_ListControl.SetColumnWidth(0, 120);
		m_ListControl.SetColumnWidth(1, 200);

		// Loop through each record
		while( !recset.IsEOF() )
		{
			// Copy each column into a variable
			recset.GetFieldValue("CatID", sCatID);
			recset.GetFieldValue("Category", sCategory);

			// Insert values into the list control
			iRec = m_ListControl.InsertItem(0, sCatID, 0);
			m_ListControl.SetItemText(0, 1, sCategory);

			// goto next record
			recset.MoveNext();
		}
		// Close the database
		database.Close();
	}
	CATCH(CDBException, e)
	{
		// If a database exception occured, show error msg
		AfxMessageBox("Database error: "+ e->m_strError);
	}
	END_CATCH;
}
	
// Reset List control
void CReadDBDlg::ResetListControl()
{
	m_ListControl.DeleteAllItems();
	int iNbrOfColumns;

	CHeaderCtrl* pHeader = (CHeaderCtrl*)m_ListControl.GetDlgItem(0);

	if (pHeader)
	{
		iNbrOfColumns = pHeader->GetItemCount();
	}
	for (int i = iNbrOfColumns; i >= 0; i--)
	{
		m_ListControl.DeleteColumn(i);
	}
}
