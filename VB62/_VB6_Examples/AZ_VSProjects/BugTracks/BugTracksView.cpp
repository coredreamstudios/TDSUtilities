// BugTracksView.cpp : implementation of the CBugTracksView class
//

#include "stdafx.h"
#include "BugTracks.h"

#include "BugTracksDoc.h"
#include "BugTracksView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBugTracksView

IMPLEMENT_DYNCREATE(CBugTracksView, CView)

BEGIN_MESSAGE_MAP(CBugTracksView, CView)
	//{{AFX_MSG_MAP(CBugTracksView)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBugTracksView construction/destruction

CBugTracksView::CBugTracksView()
{
	// TODO: add construction code here

}

CBugTracksView::~CBugTracksView()
{
}

BOOL CBugTracksView::PreCreateWindow(CREATESTRUCT& cs)
{
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return CView::PreCreateWindow(cs);
}

/////////////////////////////////////////////////////////////////////////////
// CBugTracksView drawing

void CBugTracksView::OnDraw(CDC* pDC)
{
	CBugTracksDoc* pDoc = GetDocument();
	ASSERT_VALID(pDoc);
	// TODO: add draw code for native data here
	
	int x = 144.678;
	int y = 100;

	int N=pDoc->m_BugDataArray.GetSize();
	
	for (int i=0; i < N-2; i++)
	{
		pDC->MoveTo(pDoc->m_BugDataArray[i].x, pDoc->m_BugDataArray[i].y);
		pDC->LineTo(pDoc->m_BugDataArray[i+1].x, pDoc->m_BugDataArray[i+1].y);

		pDC->TextOut(x, y, "This is it!");
		pDC->TextOut(x + 50, y + 50, "This is the time to call for all hands!");

		x = x + 100;
		y = y + 100;
	}
}

/////////////////////////////////////////////////////////////////////////////
// CBugTracksView diagnostics

#ifdef _DEBUG
void CBugTracksView::AssertValid() const
{
	CView::AssertValid();
}

void CBugTracksView::Dump(CDumpContext& dc) const
{
	CView::Dump(dc);
}

CBugTracksDoc* CBugTracksView::GetDocument() // non-debug version is inline
{
	ASSERT(m_pDocument->IsKindOf(RUNTIME_CLASS(CBugTracksDoc)));
	return (CBugTracksDoc*)m_pDocument;
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CBugTracksView message handlers
