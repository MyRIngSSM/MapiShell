#pragma once
#include <objbase.h>
#include <initguid.h>

#define INITGUID
#define USES_IID_IMAPIAdviseSink

#include <windows.h>
//#include <cemapi.h>
#include <mapiutil.h>

class CMAPIAdviseSink :public IMAPIAdviseSink {
private:
    long m_lRef;
public:
    CMAPIAdviseSink();
    ~CMAPIAdviseSink();

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, LPVOID* ppv);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // IMAPIAdviseSink
    STDMETHODIMP_(ULONG) OnNotify(ULONG cNotif,
        LPNOTIFICATION lpNotifications);
};