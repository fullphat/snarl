#include "StdAfx.h"
#include <strsafe.h>
#include "SnarlTestHelper.h"
#include "resource.h"


CSnarlTestHelper::CSnarlTestHelper(HWND hWndMain, HWND hWndEdit)
	: hWndMain(hWndMain), hWndEdit(hWndEdit)
{
}


CSnarlTestHelper::~CSnarlTestHelper(void)
{
}

// ----------------------------------------------------------------------------

void CSnarlTestHelper::WriteLine(LPCTSTR str, ...)
{
	LRESULT cchOldTextLen = 0, cchStrTextLen = 0; // Length in characters
	LRESULT cchNewTextLen = 0;
	size_t nTemp = 0;

	if (!str)
		return;

	va_list args;
	va_start(args, str);
	cchStrTextLen = _vsctprintf(str, args);
	if (cchStrTextLen <= 0)
		return;

	cchOldTextLen = SendMessage(hWndEdit, WM_GETTEXTLENGTH, 0, 0);
	cchNewTextLen = cchOldTextLen + cchStrTextLen;
	cchNewTextLen += 3; // + \r\n + null

	// Allocate for the text and copy text
	TCHAR* strNewText = new TCHAR[cchNewTextLen];

	SendMessage(hWndEdit, WM_GETTEXT, cchNewTextLen, (LPARAM)strNewText);
	nTemp = _vsntprintf_s(strNewText + cchOldTextLen, cchStrTextLen + 1, _TRUNCATE, str, args);
	StringCchCat(strNewText, cchNewTextLen, _T("\r\n"));

	SendMessage(hWndEdit, WM_SETTEXT, 0, (LPARAM)strNewText);

	va_end(args);
	delete [] strNewText;
}

void CSnarlTestHelper::Wait(DWORD dwMilliseconds)
{
	InvalidateRect(hWndEdit, NULL, TRUE);
	UpdateWindow(hWndEdit);

	Sleep(dwMilliseconds);
}

void CSnarlTestHelper::EnableMenu()
{
	//MenuHelper(MF_ENABLED);
}

void CSnarlTestHelper::DisableMenu()
{
	//MenuHelper(MF_GRAYED);
}

void CSnarlTestHelper::MenuHelper(UINT flag)
{
	HMENU hMenu = GetMenu(hWndMain);
	if (hMenu == NULL)
		return;

	int itemCount = GetMenuItemCount(hMenu);
	if (itemCount == -1)
		return;

	for (int i = 0; i < itemCount; ++i)
	{
		EnableMenuItem(hMenu, i, MF_BYPOSITION | flag);
	}

	InvalidateRect(hWndMain, NULL, TRUE);
	UpdateWindow(hWndMain);
	Sleep(0);
}
