// SimpleTest.cpp
// This is quick test and example code for using the Snarl C++ Interface
// Includes basic Win32 example and STL example

#include "stdafx.h"

#include "..\..\..\SnarlInterface_v42\SnarlInterface.h"
using namespace Snarl::V42;

#ifdef UNICODE
#define tout std::wcout
#else
#define tout std::cout
#endif

void Example1();
void Example2();
void Example3();


int _tmain(int argc, _TCHAR* argv[])
{
	if (!SnarlInterface::IsSnarlRunning()) {
		tout << _T("Snarl is not running") << std::endl;
		return 1;
	}

	Example1();
	Example2();
	Example3();

	tout << _T("Hit a key to quit") << std::endl;
	_getch();

	return 0;
}

void Example1()
{
	const LPCTSTR APP_ID = _T("CppTest"); 

	SnarlInterface snarl;

	// Manuel test
	LONG32 ret = snarl.DoRequest(L"hello");
	if (ret < 0)
		tout << _T("Last call returned an error. Errorcode: ") << abs(ret) << std::endl;

	// Simple test
	snarl.Register(APP_ID, _T("C++ test app"), NULL);
	snarl.AddClass(_T("Class1"), _T("Class 1"));

	tout << _T("Ready for action. Will post some messages...") << std::endl;

	snarl.Notify(_T("Class1"), _T("C++ example 1"), _T("Some text"), 10);

	tout << _T("Hit a key to unregister") << std::endl;
	_getch();
	snarl.Unregister(APP_ID);
}

// Strict example from SnarlInterface.cpp
void Example2()
{
	const LPCTSTR APP_ID = _T("CppTest");

	SnarlInterface snarl;
	snarl.Register(APP_ID, _T("C++ test app"), NULL);

	snarl.AddClass(_T("Class1"), _T("Some class description"));
	snarl.Notify(_T("Class1"), _T("C++ example 1"), _T("Some text"), 10);

	tout << _T("Hit a key to unregister") << std::endl;
	_getch();
	snarl.Unregister(APP_ID);
}

void Example3()
{
	SnarlInterface snarl;

	const LPCTSTR APP_ID = _T("CppTest");
	snarl.Register(APP_ID, _T("C++ test app"), NULL);
	snarl.AddClass(_T("Class1"), _T("Class 1"));

	tout << _T("Ready for action. Will post some messages...") << std::endl;
	
	snarl.Notify(_T("Class1"), _T("C++ example 1"), _T("Some text"), 10);

	std::basic_stringstream<TCHAR> sstr1;
	sstr1 << _T("Size of TCHAR = ") << sizeof(TCHAR) << std::endl;
	sstr1 << _T("Snarl version = ") << snarl.GetVersion() << std::endl;
	sstr1 << _T("Snarl windows = ") << snarl.GetSnarlWindow() << std::endl;

	snarl.Notify(_T("Class1"), _T("Runtime info"), sstr1.str().c_str(), 10);
	sstr1 = std::basic_stringstream<TCHAR>();

	// -------------------------------------------------------------------

	// DON'T DO THIS
	// sstr1 << _T("Snarl icons path = ") << snarl.GetIconsPath() << std::endl;
	// We need to free the string!
		
	LPCTSTR tmp = snarl.GetIconsPath(); // Release with FreeString
	if (tmp != NULL) {
		sstr1 << tmp << _T("info.png");
		snarl.Notify(_T("Class1"), _T("Icon test"), _T("Some text and an icon"), 10, sstr1.str().c_str());
		snarl.FreeString(tmp);		
	}

	tout << _T("Hit a key to unregister") << std::endl;
	_getch();
	snarl.Unregister(APP_ID);
}
