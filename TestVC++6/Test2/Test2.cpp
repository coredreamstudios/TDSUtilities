//=================================================================================================================================================
// Test2.cpp : Defines the entry point for the application.
//
//
//
//
//=================================================================================================================================================

#include "stdafx.h"
#include "resource.h"
#include <iostream>

#import "C:\program files\common files\system\ado\msado15.dll" rename("EOF", "ADOEOF")

using namespace std;
using namespace ADODB;

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;													// current instance
TCHAR szTitle[MAX_LOADSTRING];										// The title bar text
TCHAR szWindowClass[MAX_LOADSTRING];								// The title bar text

// Foward declarations of functions included in this code module:
LRESULT CALLBACK	mainDlg(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK	About(HWND, UINT, WPARAM, LPARAM);

INT_PTR CALLBACK	DlgProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);

//=================================================================================================================================================
int APIENTRY WinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPSTR     lpCmdLine,
                     int       nCmdShow)
{
 	// TODO: Place code here.
	MSG msg;
	HACCEL hAccelTable;
	
	hInst = hInstance;

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_TEST2, szWindowClass, MAX_LOADSTRING);

	InitCommonControls();

	DialogBox(hInstance, (LPCTSTR)IDD_DIALOG1, NULL, (DLGPROC)mainDlg);

	hAccelTable = LoadAccelerators(hInstance, (LPCTSTR)IDC_TEST2);

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0)) 
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg)) 
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	return msg.wParam;
}

//=================================================================================================================================================
// Mesage handler for about box.
LRESULT CALLBACK mainDlg(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
		case WM_INITDIALOG:
				return TRUE;

		case WM_COMMAND:
			if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) 
			{
				EndDialog(hDlg, LOWORD(wParam));
				return TRUE;
			}
			switch(wParam)
			{
				case IDC_BUTTON1:
					MessageBox(hDlg, "Hello World", "Test Message Box", MB_OK);
					break;
				
				case IDC_BUTTON2:
					DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG2), hDlg, (DLGPROC)DlgProc);
					break;
			}
			break;

		case IDCANCEL:
			PostQuitMessage(0);
			break;
	}

    return FALSE;
}

//=================================================================================================================================================
// Mesage handler for about box.
LRESULT CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
		case WM_INITDIALOG:
				return TRUE;

		case WM_COMMAND:
			if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) 
			{
				EndDialog(hDlg, LOWORD(wParam));
				return TRUE;
			}
			break;
	}

    return FALSE;
}

//===============================================================================================================================
INT_PTR CALLBACK DlgProc(HWND hWndDlg, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	switch (Msg)
	{
		case WM_INITDIALOG:
				return TRUE;

		case WM_COMMAND:
			if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) 
			{
				EndDialog(hWndDlg, LOWORD(wParam));
				return TRUE;
			}
			break;
	}

    return FALSE;
}

//=================================================================================================================================================
BOOL OpenADORecordset(HWND hWndParent)
{
	/*HRESULT hr;
	CoInitialize(NULL);
	
	try
	{
		_ConnectionPtr conn;
		_RecordsetPtr rs;

		hr = conn.CreateInstance(__uuidof(Connection));
		hr = rs.CreateInstance(__uuidof(Recordset));

		conn->CursorLocation = adUseClient;
		
		_bstr_t strConn("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\InStore\\OMInStore.mdb;Persist Security Info=False");
		hr = conn->Open(strConn, "", "", adConnectUnspecified);

		rs->Open("SELECT * FROM Ingredient", conn.GetInterfacePtr(), adOpenForwardOnly, adLockReadOnly, adCmdText);

		while(!rs->ADOEOF)
		{
			_bstr_t strData(rs->Fields->GetItem(L"Description")->GetValue());

			cTxtLen = sizeof(strData);
			pszMem = (LPWSTR) VirtualAlloc((LPVOID)NULL, (DWORD)(cTxtLen + 1), MEM_COMMIT, PAGE_READWRITE);
			pszMem = strData;
			
			int index = SendDlgItemMessage(hWndParent, IDC_LIST1, LB_ADDSTRING, 0, (DWORD)((LPSTR)pszMem));
			int nTimes = rs->Fields->GetItem(L"ID")->GetValue();
			SendDlgItemMessage(hWndParent, IDC_LIST1, LB_SETITEMDATA, (WPARAM)index, (LPARAM)nTimes);
		   
			rs->MoveNext();
		}

		rs->Close();
	}
	catch(_com_error &e)
	{
		//dump_error(e);
		MessageBox(NULL, L"" + e.Error(), NULL, NULL);
	}
	catch(...)
	{
		cout << "Unhandled Exception";
	};*/
	
	return true;
}

//=================================================================================================================================================
// end of file
//=================================================================================================================================================
