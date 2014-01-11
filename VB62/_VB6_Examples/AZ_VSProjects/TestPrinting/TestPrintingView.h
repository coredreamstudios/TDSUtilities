// TestPrintingView.h : interface of the CTestPrintingView class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_TESTPRINTINGVIEW_H__D8B9C8DD_766A_4EC0_B42F_5F2AE619E0C0__INCLUDED_)
#define AFX_TESTPRINTINGVIEW_H__D8B9C8DD_766A_4EC0_B42F_5F2AE619E0C0__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class CTestPrintingView : public CView
{
protected: // create from serialization only
	CTestPrintingView();
	DECLARE_DYNCREATE(CTestPrintingView)

// Attributes
public:
	CTestPrintingDoc* GetDocument();

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTestPrintingView)
	public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	protected:
	virtual BOOL OnPreparePrinting(CPrintInfo* pInfo);
	virtual void OnBeginPrinting(CDC* pDC, CPrintInfo* pInfo);
	virtual void OnEndPrinting(CDC* pDC, CPrintInfo* pInfo);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CTestPrintingView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	//{{AFX_MSG(CTestPrintingView)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

#ifndef _DEBUG  // debug version in TestPrintingView.cpp
inline CTestPrintingDoc* CTestPrintingView::GetDocument()
   { return (CTestPrintingDoc*)m_pDocument; }
#endif

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TESTPRINTINGVIEW_H__D8B9C8DD_766A_4EC0_B42F_5F2AE619E0C0__INCLUDED_)
