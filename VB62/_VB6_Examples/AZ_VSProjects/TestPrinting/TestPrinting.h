// TestPrinting.h : main header file for the TESTPRINTING application
//

#if !defined(AFX_TESTPRINTING_H__6424D768_ABF5_43F2_A68C_4C620D0F56AB__INCLUDED_)
#define AFX_TESTPRINTING_H__6424D768_ABF5_43F2_A68C_4C620D0F56AB__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CTestPrintingApp:
// See TestPrinting.cpp for the implementation of this class
//

class CTestPrintingApp : public CWinApp
{
public:
	CTestPrintingApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CTestPrintingApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation
	//{{AFX_MSG(CTestPrintingApp)
	afx_msg void OnAppAbout();
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_TESTPRINTING_H__6424D768_ABF5_43F2_A68C_4C620D0F56AB__INCLUDED_)
