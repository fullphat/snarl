#include "StdAfx.h"
#include "SnarlV39Test.h"
#include "SnarlTestHelper.h"

using namespace Snarl::V39;


static LPCSTR STR_APPNAME  = "Snarl C++ Test";
static LPCSTR STR_CLASS1   = "Test class 1";


CSnarlV39Test::CSnarlV39Test(SnarlInterface* s, CSnarlTestHelper* pTestHelper)
	: pHelper(pTestHelper), snarl(s), hWnd(pHelper->hWndMain)
{
}

CSnarlV39Test::~CSnarlV39Test(void)
{
}

// ----------------------------------------------------------------------------

void CSnarlV39Test::V38Test1()
{
	LONG snGlobalMsg = snarl->GetGlobalMsg();
	HWND snGetSnarlWindow = snarl->GetSnarlWindow();
	
	WORD snMajor = 0, snMinor = 0;
	snarl->GetVersion(&snMajor, &snMinor);
	LONG32 snGetVersionEx = snarl->GetVersionEx();

	LPCTSTR snGetAppPath = snarl->GetAppPath();
	LPCTSTR snGetIconsPath = snarl->GetIconsPath();

	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));

	pHelper->WriteLine(_T("GlobalMsg: %x | snGetSnarlWindow: %x"), snGlobalMsg, snGetSnarlWindow);
	pHelper->WriteLine(_T("GetAppPath: %s"), snGetAppPath);
	pHelper->WriteLine(_T("GetIconsPath: %s"), snGetIconsPath);

	pHelper->WriteLine(_T("GetVersion: %d.%d"), snMajor, snMinor);
	pHelper->WriteLine(_T("GetVersionEx: %d"), snGetVersionEx);

	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));

	snarl->FreeString(snGetIconsPath);
	snarl->FreeString(snGetAppPath);
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Test a "normal" use case, which incl.
// Register config, register class, sending a message and cleanup
///////////////////////////////////////////////////////////////////////////////////////////////////
void CSnarlV39Test::V38Test2()
{
	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));
	pHelper->WriteLine(_T("This test registers with Snarl, so remember to call RevokeConfig after the test has run"));

	pHelper->WriteLine(_T("RegisterConfig2: %x"),       snarl->RegisterConfig2(hWnd, STR_APPNAME, 0, ""));
	pHelper->WriteLine(_T("RegisterAlert: %x"),         snarl->RegisterAlert(STR_APPNAME, STR_CLASS1));
	
	pHelper->WriteLine(_T("ShowMessageEx: %x"),         snarl->ShowMessageEx(STR_CLASS1, "Test Title", "Test text\n:)", 10, "", 0, 0, ""));

	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Test unicode send

void CSnarlV39Test::V38Test3()
{
	LPCWSTR szTestApp = L"Unicode app test";
	LPCWSTR szTestClass = L"TestClass2";
	
	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));
	pHelper->WriteLine(_T("This test should print legal unicode charactors"));

	pHelper->WriteLine(_T("RegisterConfig2: %x"),       snarl->RegisterConfig2(hWnd, szTestApp, 0, L""));
	pHelper->WriteLine(_T("RegisterAlert: %x"),         snarl->RegisterAlert(szTestApp, szTestClass));
	
	//pHelper->WriteLine(_T("ShowMessageEx: %x"),         snarl->ShowMessageEx(szTestClass, L"Test Title \u03a3", L"Test text\nSpecial characters: Ⱡ, Ǻ, Ș, 葉, \u03a0 Pi, \u03a3 Sigma, α Alpha, β Beta, γ Gamma", 10, L"", 0, 0, L""));
	pHelper->WriteLine(_T("ShowMessageEx: %x"),         snarl->ShowMessageEx(szTestClass, L"Test Title \u03a3", L"Test text\nSpecial characters: 完了しました != 完乾Eました and おはよう != おEよう", 10, L"", 0, 0, L""));

	pHelper->WriteLine(_T("RevokeConfig: %x"),          snarl->RevokeConfig(hWnd));

	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V38RevokeConfig()
{
	pHelper->WriteLine(_T("RevokeConfig: %x"),			snarl->RevokeConfig(hWnd));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V38HideMessage()
{
	pHelper->WriteLine(_T("HideMessage: %x"),			snarl->HideMessage());
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V38IsMessageVisible()
{
	pHelper->WriteLine(_T("IsMessageVisible: %x"),		snarl->IsMessageVisible());
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V38RegisterConfig()
{
	pHelper->WriteLine(_T("RegisterConfig: %x"),		snarl->RegisterConfig(hWnd, STR_APPNAME, 0));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V38SetTimeout()
{
	pHelper->WriteLine(_T("SetTimeout: %x"),			snarl->SetTimeout(8));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V38ShowMessage()
{
	pHelper->WriteLine(_T("ShowMessage: %x"),			snarl->ShowMessage("Test Message", "Text\nLine 2", 0, "", 0, 0));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V38ShowMessageEx()
{
	pHelper->WriteLine(_T("ShowMessageEx: %x"),			snarl->ShowMessageEx(STR_CLASS1, "ShowMessageEx", "Text message class 1", 10, "", 0, 0, ""));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V38UpdateMessage()
{
	pHelper->WriteLine(_T("UpdateMessage: %x"),		snarl->UpdateMessage("Updated message", "New text", ""));
}


///////////////////////////////////////////////////////////////////////////////////////////////////
// V39 API TEST
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V39Test1()
{
	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));
	
	V39GetAppMsg();
	V39GetRevision();
	
	V39RegisterApp();
	V39AddClass();
	V39SetClassDefault();
	V39ShowNotification();
	V39UnRegisterApp();
	
	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V39Test2()
{
	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));
	V39RegisterApp();
	
	SNARL_CLASS_FLAGS Flags = (SNARL_CLASS_FLAGS)(SNARL_CLASS_ENABLED | SNARL_CLASS_NO_DUPLICATES);
	pHelper->WriteLine(_T("AddClass: %x"), snarl->AddClass(L"STR_CLASS1", L"Description", Flags, L"Default title", L"", 10));
	
	pHelper->WriteLine(_T("ShowNotification: %x"), snarl->ShowNotification(L"STR_CLASS1", L"Test title", L"Test text\nSpecial characters: Ⱡ, Ǻ, Ș, 葉, \u03a0 Pi, \u03a3 Sigma, α Alpha, β Beta, γ Gamma"));
	
	V39UnRegisterApp();	
	pHelper->WriteLine(_T("--------------------------------------------------------------------------------------------------"));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V39AddClass()
{
	SNARL_CLASS_FLAGS Flags = (SNARL_CLASS_FLAGS)(SNARL_CLASS_ENABLED | SNARL_CLASS_NO_DUPLICATES);
	
	pHelper->WriteLine(_T("AddClass: %x"),
		snarl->AddClass(L"STR_CLASS1", L"Description", Flags, L"Default title", L"", 10));
		//snarl->AddClass(STR_CLASS1, "Description", Flags, "Default title", "", 10));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V39ChangeAttribute()
{
	pHelper->WriteLine(_T("ChangeAttribute: %x"),	snarl->ChangeAttribute(snarl->GetLastMessageId(), SNARL_ATTRIBUTE_TITLE, "New title"));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V39GetAppMsg()
{
	pHelper->WriteLine(_T("GetAppMsg: %x"),			snarl->GetAppMsg());
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V39GetRevision()
{
	pHelper->WriteLine(_T("GetRevision: %d"),		snarl->GetRevision());
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V39RegisterApp()
{
	pHelper->WriteLine(_T("RegisterApp: %x"),		snarl->RegisterApp(STR_APPNAME, "", "", hWnd));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V39SetAsSnarlApp()
{
	pHelper->WriteLine(_T("SetAsSnarlApp called - no return type"));
	snarl->SetAsSnarlApp(hWnd);
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V39SetClassDefault()
{
	pHelper->WriteLine(_T("SetClassDefault: %x : Timeout=20"),	snarl->SetClassDefault(STR_CLASS1, SNARL_ATTRIBUTE_TIMEOUT, "20"));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V39ShowNotification()
{
	pHelper->WriteLine(_T("ShowNotification: %x"), snarl->ShowNotification(STR_CLASS1, "Test title", "Test text"));
}

///////////////////////////////////////////////////////////////////////////////////////////////////

void CSnarlV39Test::V39UnRegisterApp()
{
	pHelper->WriteLine(_T("UnRegisterApp: %x"),			snarl->UnregisterApp());
}

