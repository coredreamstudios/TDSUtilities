// TestPrintingDoc.h : interface of the CTestPrintingDoc class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_TESTPRINTINGDOC_H__22395C22_AC0E_4677_89FC_2809FF595F6A__INCLUDED_)
#define AFX_TESTPRINTINGDOC_H__22395C22_AC0E_4677_89FC_2809FF595F6A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class CTestPrintingDoc : public CDocument
{
protected: // create from serialization only
	CTestPrintingDoc();
	DECLARE_DYNCREATE(CTestPrintingDoc)

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTestPrintingDoc)
	public:
	virtual BOOL OnNewDocument();
	virtual void Serialize(CArchive& ar);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CTestPrintingDoc();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	//{{AFX_MSG(CTestPrintingDoc)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TESTPRINTINGDOC_H__22395C22_AC0E_4677_89FC_2809FF595F6A__INCLUDED_)
