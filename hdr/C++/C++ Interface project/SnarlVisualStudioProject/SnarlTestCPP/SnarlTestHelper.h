#pragma once


class CSnarlTestHelper
{
public:
	CSnarlTestHelper(HWND hWndMain, HWND hWndEdit);
	~CSnarlTestHelper(void);

	void WriteLine(LPCTSTR str, ...);
	void Wait(DWORD dwMilliseconds);
	void EnableMenu();
	void DisableMenu();

public:
	void MenuHelper(UINT flag);

	const HWND hWndMain;
	const HWND hWndEdit;
};
