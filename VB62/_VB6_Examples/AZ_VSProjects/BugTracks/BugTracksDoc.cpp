// BugTracksDoc.cpp : implementation of the CBugTracksDoc class
//

#include "stdafx.h"
#include "BugTracks.h"

#include "BugTracksDoc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CBugTracksDoc

IMPLEMENT_DYNCREATE(CBugTracksDoc, CDocument)

BEGIN_MESSAGE_MAP(CBugTracksDoc, CDocument)
	//{{AFX_MSG_MAP(CBugTracksDoc)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CBugTracksDoc construction/destruction

CBugTracksDoc::CBugTracksDoc()
{
	// TODO: add one-time construction code here

}

CBugTracksDoc::~CBugTracksDoc()
{
}

BOOL CBugTracksDoc::OnNewDocument()
{
	if (!CDocument::OnNewDocument())
		return FALSE;

	m_BugDataArray.RemoveAll();
	m_ReedStr.RemoveAll();

	// TODO: add reinitialization code here
	// (SDI documents will reuse this document)

	return TRUE;
}



/////////////////////////////////////////////////////////////////////////////
// CBugTracksDoc serialization

void CBugTracksDoc::Serialize(CArchive& ar)
{
	if (!ar.IsStoring())
	{
		SBugData Data;
		CString strOneLine;

		while(ar.ReadString(strOneLine))
		{
			sscanf(strOneLine,"%g %g\n",&Data.x,&Data.y);
			m_BugDataArray.Add(Data);
		}
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}

/////////////////////////////////////////////////////////////////////////////
// CBugTracksDoc diagnostics

#ifdef _DEBUG
void CBugTracksDoc::AssertValid() const
{
	CDocument::AssertValid();
}

void CBugTracksDoc::Dump(CDumpContext& dc) const
{
	CDocument::Dump(dc);
}
#endif //_DEBUG

/////////////////////////////////////////////////////////////////////////////
// CBugTracksDoc commands
