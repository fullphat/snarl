#include "StdAfx.h"
#include "SnarlV41Test.h"
#include "SnarlTestHelper.h"

// using namespace Snarl::V41;

static LPCTSTR APPNAME  = _T("Snarl C++ Test - v41");

static LPCTSTR CLASS1 = _T("Class1");
static LPCTSTR CLASS2 = _T("Class2");
static LPCTSTR CLASS_DESC1 = _T("Class 1");
static LPCTSTR CLASS_DESC2 = _T("Class 2");

static LPCTSTR TESTMSG1 = L"Test text\nSpecial characters: 完了しました != 完乾Eました and おはよう != おEよう";
static LPCTSTR TESTMSG2 = L"Test text 2\nSecond line";

static LONG32 DEFAULT_TIMEOUT = 10;

// ----------------------------------------------------------------------------

CSnarlV41Test::CSnarlV41Test(Snarl::V41::SnarlInterface* s, CSnarlTestHelper* pTestHelper)
	: pHelper(pTestHelper), snarl(s), hWnd(pHelper->hWndMain)
{
	
}

CSnarlV41Test::~CSnarlV41Test(void)
{
}

// ----------------------------------------------------------------------------

LPCTSTR CSnarlV41Test::GetIcon(int icon)
{
	LPTSTR szIcon = NULL;
	switch (icon)
	{
		case 0: szIcon = _T("snarl.png"); break;
		case 1: szIcon = _T("snarl-update.png"); break;
		case 2: szIcon = _T("display.png"); break;
		case 3: szIcon = _T("info.png"); break;
		case 4: szIcon = _T("default_style.png"); break;
		case 5: szIcon = _T("critical.png"); break;
		default: szIcon = _T("snarl.png"); break;
	}
	
	LPCTSTR szIconPath = snarl->GetIconsPath();
	size_t iconLen = _tcslen(szIcon);
	size_t fullLen = _tcslen(szIconPath) + iconLen + 1; // + NULL
	
	LPTSTR szRet = snarl->AllocateString(fullLen);
	_tcsncpy_s(szRet, fullLen, szIconPath, _TRUNCATE);
	_tcsncat_s(szRet, fullLen, szIcon, _TRUNCATE);
	snarl->FreeString(szIconPath);

	return szRet;
}


///////////////////////////////////////////////////////////////////////////////////////////////////
// Test a "normal" use case, which incl.
// Register config, register class, sending messages and cleanup
///////////////////////////////////////////////////////////////////////////////////////////////////
void CSnarlV41Test::Test1()
{
	pHelper->DisableMenu();

	LPCTSTR snarlIcon2 = GetIcon(2); // Free with snarl->FreeString()
	LPCTSTR snarlIcon3 = GetIcon(3); // Free with snarl->FreeString()

	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));

	pHelper->WriteLine(_T("RegisterApp: %d"), snarl->RegisterApp(APPNAME, _T("C++ test app"), _T(""), NULL, 0, Snarl::V41::SnarlEnums::AppHasAbout));

	pHelper->WriteLine(_T("RegisterApp: %d"), snarl->RegisterApp(APPNAME, _T("C++ test app"), _T(""), NULL, 0, Snarl::V41::SnarlEnums::AppHasAbout));
	
	pHelper->WriteLine(_T("UpdateApp: %d"), snarl->UpdateApp(_T("C++ test app updated"), _T("")));
	pHelper->WriteLine(_T("UpdateApp: %d"), snarl->UpdateApp(_T("C++ test 2"), snarlIcon2));

	pHelper->WriteLine(_T("AddClass: %d"), snarl->AddClass(CLASS1, CLASS_DESC1));
	pHelper->WriteLine(_T("RemoveClass: %d"), snarl->RemoveClass(CLASS1));

	pHelper->WriteLine(_T("AddClass: %d"), snarl->AddClass(CLASS1, CLASS_DESC1));
	pHelper->WriteLine(_T("AddClass: %d"), snarl->AddClass(CLASS2, CLASS_DESC2));

	// Test EZNotify
	pHelper->WriteLine(_T("EZNotify: %d"), snarl->EZNotify(CLASS1, _T("Message 1"), _T("Test text"), DEFAULT_TIMEOUT, snarlIcon3, 0, _T("ack"), _T("val")));
	pHelper->WriteLine(_T("EZNotify: %d"), snarl->EZNotify(CLASS1, _T("Message 2"), TESTMSG1, DEFAULT_TIMEOUT, snarlIcon3, 0, _T("ack"), _T("val")));
	pHelper->WriteLine(_T("EZNotify: %d"), snarl->EZNotify(CLASS1, _T("Message 3"), TESTMSG1, DEFAULT_TIMEOUT, NULL, 0, NULL, NULL));
	pHelper->WriteLine(_T("EZNotify: %d"), snarl->EZNotify(CLASS1, _T("Message 4"), TESTMSG2));

	// Test Notify
	TCHAR szNotify[512];
	_tcsncpy_s(szNotify, _T("title::Notify test#?text::Notify custom packet message\nTimeout: 10#?timeout::10#?icon::#?priority::0#?ack::#?value::"), _TRUNCATE);
	pHelper->WriteLine(_T("Notify: %d"), snarl->Notify(CLASS2, szNotify));
	
	
	// Clean up
	pHelper->WriteLine(_T("Will cleanup in 15 seconds..."));
	pHelper->Wait(15 * 1000);

	pHelper->WriteLine(_T("RemoveAllClasses: %d"), snarl->RemoveAllClasses());
	pHelper->WriteLine(_T("UnregisterApp: %d"), snarl->UnregisterApp());

	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));

	snarl->FreeString(snarlIcon2);
	snarl->FreeString(snarlIcon3);

	pHelper->EnableMenu();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Test of functionality
// Esp. Update functions
///////////////////////////////////////////////////////////////////////////////////////////////////
void CSnarlV41Test::Test2()
{
	pHelper->DisableMenu();

	LPCTSTR snarlIcon2 = GetIcon(2); // Free with snarl->FreeString()
	LPCTSTR snarlIcon3 = GetIcon(3); // Free with snarl->FreeString()

	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));

	pHelper->WriteLine(_T("RegisterApp: %d"), snarl->RegisterApp(APPNAME, _T("C++ test app"), snarlIcon2, NULL, 0, Snarl::V41::SnarlEnums::AppDefault));
	pHelper->WriteLine(_T("AddClass: %d"), snarl->AddClass(CLASS1, CLASS_DESC1));
	pHelper->WriteLine(_T("AddClass: %d"), snarl->AddClass(CLASS2, CLASS_DESC2));

	// Test EZNotify and Update
	pHelper->WriteLine(_T("EZNotify: %d"), snarl->EZNotify(CLASS1, _T("Message 4"), TESTMSG1, 0));
	pHelper->Wait(2000);
	pHelper->WriteLine(_T("EZUpdate: %d"), snarl->EZUpdate(snarl->GetLastMsgToken(), _T("New title"), _T("New text"), 0, snarlIcon3));
	pHelper->Wait(2000);
	pHelper->WriteLine(_T("EZUpdate: %d"), snarl->EZUpdate(snarl->GetLastMsgToken(), NULL, _T("Only updating text")));
	pHelper->Wait(2000);
	pHelper->WriteLine(_T("EZUpdate: %d"), snarl->EZUpdate(snarl->GetLastMsgToken(), NULL, _T("Updating text and icon"), -1, snarlIcon2));
	pHelper->Wait(2000);
	pHelper->WriteLine(_T("EZUpdate: %d"), snarl->EZUpdate(snarl->GetLastMsgToken(), _T("Updating timeout"), NULL, DEFAULT_TIMEOUT));

	// Test Notify
	const int NOTIFY_BUF_SIZE = 512;
	TCHAR szNotify[NOTIFY_BUF_SIZE];

	_tcsncpy_s(szNotify, _T("title::Notify test#?text::Notify custom packet message\nTimeout: 0#?timeout::0#?icon::#?priority::0#?ack::#?value::"), _TRUNCATE);
	pHelper->WriteLine(_T("Notify: %d"), snarl->Notify(CLASS2, szNotify));
	pHelper->Wait(2000);

	_tcsncpy_s(szNotify, _T("title::Only title update"), _TRUNCATE);
	pHelper->WriteLine(_T("Notify: %d"), snarl->Update(snarl->GetLastMsgToken(), szNotify));
	pHelper->Wait(2000);

	_tcsncpy_s(szNotify, _T("text::Text and timeout change\nTimeout: 5#?timeout::5"), _TRUNCATE);
	pHelper->WriteLine(_T("Notify: %d"), snarl->Update(snarl->GetLastMsgToken(), szNotify));
	
	
	// Clean up
	// pHelper->WriteLine(_T("Will cleanup in 15 seconds..."));
	// pHelper->Wait(15 * 1000);

	pHelper->WriteLine(_T("RemoveAllClasses: %d"), snarl->RemoveAllClasses());
	pHelper->WriteLine(_T("UnregisterApp: %d"), snarl->UnregisterApp());

	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));

	snarl->FreeString(snarlIcon2);
	snarl->FreeString(snarlIcon3);

	pHelper->EnableMenu();
}


///////////////////////////////////////////////////////////////////////////////////////////////////
// Test of misc functionality
///////////////////////////////////////////////////////////////////////////////////////////////////
void CSnarlV41Test::Test3()
{
	pHelper->DisableMenu();

	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));

	pHelper->WriteLine(_T("RegisterApp: %d"), snarl->RegisterApp(APPNAME, _T("C++ test app"), NULL, NULL, 0, Snarl::V41::SnarlEnums::AppDefault));
	pHelper->WriteLine(_T("GetLastError: %d"), snarl->GetLastError());
	pHelper->WriteLine(_T("AddClass: %d"), snarl->AddClass(CLASS1, CLASS_DESC1));
	pHelper->WriteLine(_T("GetLastError: %d"), snarl->GetLastError());

	pHelper->WriteLine(_T("IsSnarlRunning: %d"), snarl->IsSnarlRunning());
	pHelper->WriteLine(_T("AppMsg: %d"), snarl->AppMsg());
	pHelper->WriteLine(_T("Broadcast: %d"), snarl->Broadcast());
	pHelper->WriteLine(_T("GetVersion: %d"), snarl->GetVersion());
	
	// Test notification functions
	pHelper->WriteLine(_T("EZNotify: %d"), snarl->EZNotify(CLASS1, _T("Message 4"), TESTMSG1, 0));
	pHelper->Wait(2000);

	pHelper->WriteLine(_T("IsVisible: %d"), snarl->IsVisible(snarl->GetLastMsgToken()));
	pHelper->WriteLine(_T("Hide: %d"), snarl->Hide(snarl->GetLastMsgToken()));
	pHelper->WriteLine(_T("GetLastError: %d"), snarl->GetLastError());
	pHelper->WriteLine(_T("Hide: %d"), snarl->Hide(snarl->GetLastMsgToken()));
	pHelper->WriteLine(_T("GetLastError: %d"), snarl->GetLastError());

	pHelper->Wait(2000);
	pHelper->WriteLine(_T("UnregisterApp: %d"), snarl->UnregisterApp());
	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));

	pHelper->EnableMenu();
}