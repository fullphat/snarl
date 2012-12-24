#pragma once

#include <endpointvolume.h>


class CVolumeNotification : public IAudioEndpointVolumeCallback 
{
public: 
	CVolumeNotification(void);

	STDMETHODIMP_(ULONG)AddRef();
	STDMETHODIMP_(ULONG)Release();
	STDMETHODIMP QueryInterface(REFIID IID, void **ReturnValue);
	STDMETHODIMP OnNotify(PAUDIO_VOLUME_NOTIFICATION_DATA NotificationData);

	BOOL GetMuted() const;
	void SetMuted(BOOL muted);

	static int Double2Int(const double d);

private:
	~CVolumeNotification(void) {};

	void DebugOutput(int n);
	void DebugOutput(LPCSTR str);

	LONG m_refCount;
	BOOL m_muted;
};
