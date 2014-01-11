// TestPrintingDoc.cpp : implementation of the CTestPrintingDoc class
//

#include "stdafx.h"
#include "TestPrinting.h"

#include "TestPrintingDoc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CTestPrintingDoc

IMPLEMENT_DYNCREATE(CTestPrintingDoc, CDocument)

BEGIN_MESSAGE_MAP(CTestPrintingDoc, CDocument)
	//{{AFX_MSG_MAP(CTestPrintingDoc)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CTestPrintingDoc construction/destruction

CTestPrintingDoc::CTestPrintingDoc()
{
	// TODO: add one-time construction code here

}

CTestPrintingDoc::~CTestPrintingDoc()
{
}

BOOL CTestPrintingDoc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;

	// TODO: add reinitialization code here
	// (SDI documents will reuse this document)

	return TRUE;
}



/////////////////////////////////////////////////////////////////////////////
// CTestPrintingDoc serialization

void CTestPrintingDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}

/////////////////////////////////////////////////////////////////////////////
// CTestPrintingDoc diagnostics

#ifdef _DEBUG
void CTestPrintingDoc::AssertValid() const
{
	CDocument::AssertValid();
}

void CTestPrintingDoc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CTestPrintingDoc commands
