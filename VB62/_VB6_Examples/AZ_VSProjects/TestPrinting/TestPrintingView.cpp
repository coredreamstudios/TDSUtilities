// TestPrintingView.cpp : implementation of the CTestPrintingView class
//

#include "stdafx.h"
#include "TestPrinting.h"

#include "TestPrintingDoc.h"
#include "TestPrintingView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTestPrintingView

IMPLEMENT_DYNCREATE(CTestPrintingView, CView)

BEGIN_MESSAGE_MAP(CTestPrintingView, CView)
	//{{AFX_MSG_MAP(CTestPrintingView)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
	// Standard printing commands
	ON_COMMAND(ID_FILE_PRINT, CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, CView::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, CView::OnFilePrintPreview)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTestPrintingView construction/destruction

CTestPrintingView::CTestPrintingView()
{
	// TODO: add construction code here

}

CTestPrintingView::~CTestPrintingView()
{
}

BOOL CTestPrintingView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return CView::PreCreateWindow(cs);
}

/////////////////////////////////////////////////////////////////////////////
// CTestPrintingView drawing

void CTestPrintingView::OnDraw(CDC* pDC)
{
	CTestPrintingDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	// TODO: add draw code for native data here

	int x = 100;
	int y = 100;

	pDC->TextOut(x,y,"Hello World");
}

/////////////////////////////////////////////////////////////////////////////
// CTestPrintingView printing

BOOL CTestPrintingView::OnPreparePrinting(CPrintInfo* pInfo)
{
	// default preparation
	return DoPreparePrinting(pInfo);

}

void CTestPrintingView::OnBeginPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add extra initialization before printing

	
}

void CTestPrintingView::OnEndPrinting(CDC* /*pDC*/, CPrintInfo* /*pInfo*/)
{
	// TODO: add cleanup after printing
}

/////////////////////////////////////////////////////////////////////////////
// CTestPrintingView diagnostics

#ifdef _DEBUG
void CTestPrintingView::AssertValid() const
{
	CView::AssertValid();
}

void CTestPrintingView::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}

CTestPrintingDoc* CTestPrintingView::GetDocument() // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CTestPrintingDoc)));
	return (CTestPrintingDoc*)m_pDocument;
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CTestPrintingView message handlers
