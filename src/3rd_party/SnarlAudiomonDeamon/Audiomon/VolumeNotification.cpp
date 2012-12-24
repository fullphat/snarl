#include "stdafx.h"

#include "Audiomon.h"
#include "VolumeNotification.h"


CVolumeNotification::CVolumeNotification(void)
	: m_refCount(1), m_muted(false)
{
}

STDMETHODIMP_(ULONG) CVolumeNotification::AddRef()
{
	return InterlockedIncrement(&m_refCount);
}

STDMETHODIMP_(ULONG) CVolumeNotification::Release()
{ 
	LONG ref = InterlockedDecrement(&m_refCount);
	if (ref == 0)
		delete this;
	return ref;
}

STDMETHODIMP CVolumeNotification::QueryInterface(REFIID IID, void **ReturnValue)
{ 
	if (IID == IID_IUnknown || IID== __uuidof(IAudioEndpointVolumeCallback))
	{
		*ReturnValue = static_cast<IUnknown*>(this);
		AddRef();
		return S_OK;
	} 
	*ReturnValue = NULL;
	return E_NOINTERFACE; 
} 

STDMETHODIMP CVolumeNotification::OnNotify(PAUDIO_VOLUME_NOTIFICATION_DATA NotificationData) 
{ 
	HWND hSnarlWindow = GetAudiomonWindow();

	if (NotificationData->bMuted != m_muted) //(Un)muted
	{
		m_muted = NotificationData->bMuted;
		DebugOutput(m_muted ? "muted" : "unmuted");

		if (hSnarlWindow != NULL)
		{
			if (m_muted)
				SendMessage(hSnarlWindow, SNARL_MSG, SNARL_MSG_MUTED, 1);
			else
				SendMessage(hSnarlWindow, SNARL_MSG, SNARL_MSG_MUTED, 0);
		}
	}
	else // Volume changed
	{
		int iVol = Double2Int(NotificationData->fMasterVolume);
		DebugOutput(iVol);

		if (hSnarlWindow != NULL)
		{
			SendMessage(hSnarlWindow, SNARL_MSG, SNARL_MSG_VOLUME, iVol);
		}
	}
	return S_OK; 
}

BOOL CVolumeNotification::GetMuted() const
{
	return m_muted;
}

void CVolumeNotification::SetMuted(BOOL muted)
{
	m_muted = muted;
}

void CVolumeNotification::DebugOutput(int n)
{
#ifdef _DEBUG
	std::stringstream ss;
	ss << "Volume=" << n << std::endl;
	OutputDebugStringA(ss.str().c_str());
#endif
}

void CVolumeNotification::DebugOutput(LPCSTR str)
{
#ifdef _DEBUG
	std::stringstream ss;
	ss << str << std::endl;
	OutputDebugStringA(ss.str().c_str());
#endif
}

// static
int CVolumeNotification::Double2Int(const double d)
{
	int ret = static_cast<int>((d * 200.0 + 1.0) / 2);
	return ret;
}
