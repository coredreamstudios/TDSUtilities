// BugTracksDoc.h : interface of the CBugTracksDoc class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_BUGTRACKSDOC_H__CF11269E_7033_418A_B2A9_CA808FEDC877__INCLUDED_)
#define AFX_BUGTRACKSDOC_H__CF11269E_7033_418A_B2A9_CA808FEDC877__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <afxtempl.h>

struct SBugData
{
	float x;
	float y;
};

struct SStr
{
	float rstr;
	float rstr2;
};

class CBugTracksDoc : public CDocument
{
protected: // create from serialization only
	CBugTracksDoc();
	DECLARE_DYNCREATE(CBugTracksDoc)

// Attributes
public:

	CArray <SBugData, SBugData> m_BugDataArray;
	CArray <SStr, SStr> m_ReedStr;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBugTracksDoc)
	public:
	virtual BOOL OnNewDocument();
	virtual void Serialize(CArchive& ar);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CBugTracksDoc();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	//{{AFX_MSG(CBugTracksDoc)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BUGTRACKSDOC_H__CF11269E_7033_418A_B2A9_CA808FEDC877__INCLUDED_)
