// SnarlTestCPP.cpp : Defines the entry point for the application.
//

#include "stdafx.h"

#include "..\..\..\Depricated\SnarlInterface_v39.h"
#include "..\..\..\SnarlInterface_v41\SnarlInterface.h"
#include "..\..\..\SnarlInterface_v42\SnarlInterface.h"

#include "Main.h"
#include "SnarlTestHelper.h"
#include "SnarlV39Test.h"
#include "SnarlV41Test.h"
#include "SnarlV42Test.h"

// ----------------------------------------------------------------------------

static HWND hWnd       = NULL;
static HWND hWndEdit   = NULL;
static CSnarlTestHelper* pTestHelper = NULL;

static Snarl::V39::SnarlInterface* pV39Snarl = NULL;
static Snarl::V41::SnarlInterface* pV41Snarl = NULL;
static Snarl::V42::SnarlInterface* pV42Snarl = NULL;
static CSnarlV39Test* pV39SnarlTest = NULL;
static CSnarlV41Test* pV41SnarlTest = NULL;
static CSnarlV42Test* pV42SnarlTest = NULL;

// ----------------------------------------------------------------------------

#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;                                // current instance
TCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
TCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

UINT nSnarlGlobalMsg = 0;

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY _tWinMain(HINSTANCE hInstance,
                       HINSTANCE hPrevInstance,
                       LPTSTR    lpCmdLine,
                       int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	// TODO: Place code here.
	MSG msg;
	HACCEL hAccelTable;

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_SNARLTESTCPP, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}

	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_SNARLTESTCPP));

	// ------------------------------------------------------------------------
	// Create Snarl objects
	pTestHelper = new CSnarlTestHelper(hWnd, hWndEdit);
	
	pV39Snarl = new Snarl::V39::SnarlInterface();
	pV41Snarl = new Snarl::V41::SnarlInterface();
	pV42Snarl = new Snarl::V42::SnarlInterface();
	
	pV39SnarlTest = new CSnarlV39Test(pV39Snarl, pTestHelper);
	pV41SnarlTest = new CSnarlV41Test(pV41Snarl, pTestHelper);
	pV42SnarlTest = new CSnarlV42Test(pV42Snarl, pTestHelper);

	// Get the Snarl broadcast message
	nSnarlGlobalMsg = pV42Snarl->Broadcast();

	// ------------------------------------------------------------------------
	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	delete pV39Snarl;
	delete pV41SnarlTest;
	delete pTestHelper;
	delete pV39SnarlTest;

	return (int)msg.wParam;
}


//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
//  COMMENTS:
//
//    This function and its usage are only necessary if you want this code
//    to be compatible with Win32 systems prior to the 'RegisterClassEx'
//    function that was added to Windows 95. It is important to call this function
//    so that the application will get 'well formed' small icons associated
//    with it.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SNARLTESTCPP));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_SNARLTESTCPP);
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

	return RegisterClassEx(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	hInst = hInstance; // Store instance handle in our global variable

	hWnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
	  CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);

	if (!hWnd)
	  return FALSE;

	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);

	return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;

	WndProcV42(hWnd, message, wParam, lParam);

	// Test if Snarl broadcast message
	if (message == nSnarlGlobalMsg)
	{
		if (wParam == Snarl::V42::SnarlEnums::SnarlLaunched)
		{
			MessageBox(NULL, _T("Snarl Launched"), _T("Snarl C++ test app"), 0);
		}
		else if (wParam == Snarl::V42::SnarlEnums::SnarlQuit)
		{
			MessageBox(NULL, _T("Snarl Quit"), _T("Snarl C++ test app"), 0);
		}

		return 1;
	}

	// Normal message switch
	switch (message)
	{
	case WM_CREATE:
		// Create textbox
		hWndEdit = CreateWindow(_T("EDIT"), NULL,
			WS_CHILD | WS_VISIBLE | WS_VSCROLL | 
			ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
			0, 0, 0, 0, hWnd, NULL, hInst, NULL); //(HMENU) ID_EDITCHILD

		break;

	case WM_SIZE: 
		// Make the edit control the size of the window's client area. 
		MoveWindow(hWndEdit, 
				   0, 0,                  // starting x- and y-coordinates 
				   LOWORD(lParam),        // width of client area 
				   HIWORD(lParam),        // height of client area 
				   TRUE);                 // repaint window 
		return 0;


	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		
		// Parse the menu selections:
		switch (wmId)
		{
		case IDM_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			break;
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;

		// -----------------------------------------------------------------------------------------------------------

		case IDM_SNARLV38_TEST1:
			//Snarl_V38_Test1();
			pV39SnarlTest->V38Test1();
			break;
		case IDM_SNARLV38_TEST2:
			pV39SnarlTest->V38Test2();
			break;
		case IDM_SNARLV38_TEST3:
			pV39SnarlTest->V38Test3();
			break;

		case IDM_SNARLV38_HIDEMESSAGE:
			pV39SnarlTest->V38HideMessage();
			break;
		case IDM_SNARLV38_ISMESSAGEVISIBLE:
			pV39SnarlTest->V38IsMessageVisible();
			break;
		case IDM_SNARLV38_REGISTERCONFIG:
			pV39SnarlTest->V38RegisterConfig();
			break;
		case IDM_SNARLV38_REVOKECONFIG:
			pV39SnarlTest->V38RevokeConfig();
			break;
		case IDM_SNARLV38_SETTIMEOUT:
			pV39SnarlTest->V38SetTimeout();
			break;
		case IDM_SNARLV38_SHOWMESSAGE:
			pV39SnarlTest->V38ShowMessage();
			break;
		case IDM_SNARLV38_SHOWMESSAGEEX:
			pV39SnarlTest->V38ShowMessageEx();
			break;
		case IDM_SNARLV38_UPDATEMESSAGE:
			pV39SnarlTest->V38UpdateMessage();
			break;

		// -----------------------------------------------------------------------------------------------------------
		
		case IDM_SNARLV39_TEST1:
			pV39SnarlTest->V39Test1();
			break;
		case IDM_SNARLV39_TEST2:
			pV39SnarlTest->V39Test2();
			break;
			
		case IDM_SNARLV39_ADDCLASS:
			pV39SnarlTest->V39AddClass();
			break;
		case IDM_SNARLV39_CHANGEATTRIBUTE:
			pV39SnarlTest->V39ChangeAttribute();
			break;
		case IDM_SNARLV39_GETAPPMSG:
			pV39SnarlTest->V39GetAppMsg();
			break;
		case IDM_SNARLV39_GETREVISION:
			pV39SnarlTest->V39GetRevision();
			break;
		case IDM_SNARLV39_REGISTERAPP:
			pV39SnarlTest->V39RegisterApp();
			break;
		case IDM_SNARLV39_SETASSNARLAPP:
			pV39SnarlTest->V39SetAsSnarlApp();
			break;
		case IDM_SNARLV39_SETCLASSDEFAULT:
			pV39SnarlTest->V39SetClassDefault();
			break;
		case IDM_SNARLV39_SHOWNOTIFICATION:
			pV39SnarlTest->V39ShowNotification();
			break;
		case IDM_SNARLV39_UNREGISTERAPP:
			pV39SnarlTest->V39UnRegisterApp();
			break;
		
		// ----------------------------------------------------------------------------------------

		case IDM_SNARLV41_TEST1 :
			pV41SnarlTest->Test1();
			break;

		case IDM_SNARLV41_TEST2 :
			pV41SnarlTest->Test2();
			break;

		case IDM_SNARLV41_TEST3 :
			pV41SnarlTest->Test3();
			break;

		// ----------------------------------------------------------------------------------------

		case IDM_SNARLV42_TEST1 :
			pV42SnarlTest->Test1();
			break;

		case IDM_SNARLV42_TEST2 :
			pV42SnarlTest->Test2();
			break;

		case IDM_SNARLV42_TEST3 :
			pV42SnarlTest->Test3();
			break;

		case IDM_SNARLV42_ESCAPETEST :
			pV42SnarlTest->EscapeTest1();
			break;

		// -----------------------------------------------------------------------------------------------------------
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		
		// TODO: Add any drawing code here...
		
		EndPaint(hWnd, &ps);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}
