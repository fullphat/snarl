#ifndef SNARL_V41_TEST_HEADER
#define SNARL_V41_TEST_HEADER

#pragma once

#include "SnarlTestHelper.h"
#include "..\..\..\SnarlInterface_v41\SnarlInterface.h"

class CSnarlV41Test
{
public:
	void Test1();
	void Test2();
	void Test3();


	CSnarlV41Test(Snarl::V41::SnarlInterface* snarl, CSnarlTestHelper* pTestHelper);
	~CSnarlV41Test(void);

private:
	LPCTSTR GetIcon(int i);

	CSnarlTestHelper* pHelper;
	Snarl::V41::SnarlInterface* snarl;
	HWND hWnd;
};

#endif // SNARL_V41_TEST_HEADER