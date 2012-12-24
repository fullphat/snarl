#ifndef SNARL_V39_TEST_HEADER
#define SNARL_V39_TEST_HEADER

#pragma once

#include "SnarlTestHelper.h"
#include "..\..\..\Depricated\SnarlInterface_v39.h"

class CSnarlV39Test
{
public:
	// V38
	void V38Test1();
	void V38Test2();
	void V38Test3();

	void V38HideMessage();
	void V38IsMessageVisible();
	void V38RegisterConfig();
	void V38RevokeConfig();
	void V38SetTimeout();
	void V38ShowMessage();
	void V38ShowMessageEx();
	void V38UpdateMessage();

	// V39
	void V39Test1();
	void V39Test2();
	void V39AddClass();
	void V39ChangeAttribute();
	void V39GetAppMsg();
	void V39GetRevision();
	void V39RegisterApp();
	void V39SetAsSnarlApp();
	void V39SetClassDefault();
	void V39ShowNotification();
	void V39UnRegisterApp();	


	CSnarlV39Test(Snarl::V39::SnarlInterface* snarl, CSnarlTestHelper* pTestHelper);
	~CSnarlV39Test(void);

private:
	CSnarlTestHelper* pHelper;
	Snarl::V39::SnarlInterface* snarl;
	HWND hWnd;
};

#endif // SNARL_V39_TEST_HEADER