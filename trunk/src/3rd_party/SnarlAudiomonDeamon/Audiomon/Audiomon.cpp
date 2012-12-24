// --------------------------------------------------------------------------------------------
// Snarl Audiomon deamon
//
// By Jonus Conrad and Toke Noer
//
// This code is release as open source under the MIT license. Please see License.txt.
// --------------------------------------------------------------------------------------------

#include "stdafx.h"
#include <mmdeviceapi.h>
#include <endpointvolume.h>

#include "Common.h"
#include "VolumeNotification.h"
#include "Audiomon.h"


static const LPCTSTR WINDOW_CLASS_NAME = _T("snarl-audiomon-class");
static const LPCTSTR WINDOW_TITLE 	   = _T("snarl-audiomon");

// Global Variables:
HINSTANCE              hInst               = NULL;
HWND                   hWnd                = NULL;
CVolumeNotification*   volumeNotification  = NULL;
IAudioEndpointVolume*  endpointVolume 	   = NULL;


int APIENTRY _tWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	HRESULT hr;

	hInst = hInstance;

	HWND hWnd = FindWindow(WINDOW_CLASS_NAME, WINDOW_TITLE);
	if (IsWindow(hWnd))
	{
		// Check if -quit was supplied
		if (_tcsstr(lpCmdLine, _T("-quit")) != NULL)
		{
			SendMessage(hWnd, WM_CLOSE, 0, 0);
		}

		// exit this instance
		return 0;
	}

	// Create WindowClass
	WNDCLASSEX wcex = {0};
	wcex.cbSize         = sizeof(WNDCLASSEX);
	wcex.lpfnWndProc    = WndProc;
	wcex.hInstance      = hInstance;
	wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszClassName  = WINDOW_CLASS_NAME;

	ATOM classAtom = RegisterClassEx(&wcex);

	// Do initialization
	hWnd = CreateWindow(WINDOW_CLASS_NAME, WINDOW_TITLE, NULL, CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);

	if (hWnd == NULL)
		return -1;

	hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	if (hr == S_FALSE) {
		CoUninitialize();
		return -1;
	}

	// Create audio notification class
	if (!InitVolumeNotification())
		return -1;

	// Notify Snarl on startup
	float fVol = 0.0f;
	endpointVolume->GetMasterVolumeLevelScalar(&fVol);
	int iVol = CVolumeNotification::Double2Int(fVol);
	
	HWND hWndSnarl = GetAudiomonWindow();
	if (hWndSnarl != NULL)
	{
		SendMessage(hWndSnarl, SNARL_MSG, SNARL_MSG_VOLUME, iVol);
		if (volumeNotification->GetMuted())
			SendMessage(hWndSnarl, SNARL_MSG, SNARL_MSG_MUTED, 1);
		else
			SendMessage(hWndSnarl, SNARL_MSG, SNARL_MSG_MUTED, 0);
	}

	// ShowWindow(hWnd, nCmdShow);
	// UpdateWindow(hWnd);

	// Main message loop
	MSG msg;
	BOOL bRet;

	while( (bRet = GetMessage(&msg, hWnd, 0, 0)) != 0)
	{ 
		if (bRet == -1)
		{
			// handle the error and possibly exit
		}
		else
		{
			TranslateMessage(&msg); 
			DispatchMessage(&msg); 
		}
	}

	UnregisterVolumeNotifiction();
	CoUninitialize();

	return (int) msg.wParam;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

BOOL InitVolumeNotification()
{
	HRESULT hr;
	IMMDeviceEnumerator *deviceEnumerator = NULL;;
	IMMDevice *defaultDevice = NULL;

	// 
	//    Instantiate an endpoint volume object.
	//
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_INPROC_SERVER, __uuidof(IMMDeviceEnumerator), (LPVOID *)&deviceEnumerator);
	if (hr != S_OK)
		return FALSE;

	hr = deviceEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &defaultDevice);
	deviceEnumerator->Release();
	deviceEnumerator = NULL;
	if (hr != S_OK)
		return FALSE;

	hr = defaultDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_INPROC_SERVER, NULL, (LPVOID*)&endpointVolume);
	if (hr != S_OK || endpointVolume == NULL)
	{
		defaultDevice->Release();
		return FALSE;
	}

	defaultDevice->Release();
	defaultDevice = NULL;

	volumeNotification = new CVolumeNotification();
	
	BOOL muted = false;
	hr = endpointVolume->GetMute(&muted);
	volumeNotification->SetMuted(muted);

	hr = endpointVolume->RegisterControlChangeNotify(volumeNotification);
	if (hr != S_OK)
	{
		endpointVolume->Release(); 
		volumeNotification->Release();
		return FALSE;
	}

	return TRUE;
}

void UnregisterVolumeNotifiction()
{
	endpointVolume->UnregisterControlChangeNotify(volumeNotification); 
	endpointVolume->Release(); 
	volumeNotification->Release(); // Last Release will delete
	volumeNotification = NULL;
}

HWND GetAudiomonWindow()
{
	HWND hWnd = FindWindow(AUDIOMON_CLASSNAME, NULL);
	if (IsWindow(hWnd))
		return hWnd;
	
	return NULL;
}
