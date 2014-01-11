// BugTracksView.h : interface of the CBugTracksView class
//
/////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_BUGTRACKSVIEW_H__2BF6AAD9_FD65_4347_9547_FA553C02101B__INCLUDED_)
#define AFX_BUGTRACKSVIEW_H__2BF6AAD9_FD65_4347_9547_FA553C02101B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class CBugTracksView : public CView
{
protected: // create from serialization only
	CBugTracksView();
	DECLARE_DYNCREATE(CBugTracksView)

// Attributes
public:
	CBugTracksDoc* GetDocument();

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CBugTracksView)
	public:
	virtual void OnDraw(CDC* pDC);  // overridden to draw this view
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	protected:
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CBugTracksView();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

protected:

// Generated message map functions
protected:
	//{{AFX_MSG(CBugTracksView)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

#ifndef _DEBUG  // debug version in BugTracksView.cpp
inline CBugTracksDoc* CBugTracksView::GetDocument()
   { return (CBugTracksDoc*)m_pDocument; }
#endif

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_BUGTRACKSVIEW_H__2BF6AAD9_FD65_4347_9547_FA553C02101B__INCLUDED_)
