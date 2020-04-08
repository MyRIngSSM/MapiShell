#include "MapiNotify.h"

// CMAPIAdviseSink object
CMAPIAdviseSink::CMAPIAdviseSink() {
    m_lRef = 1;
    return;
}

CMAPIAdviseSink::~CMAPIAdviseSink() {
    return;
}

// IMAPIAdviseSink's IUnknown interface
STDMETHODIMP CMAPIAdviseSink::QueryInterface(REFIID riid,
    LPVOID* ppv) {
    if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid,
        IID_IMAPIAdviseSink)) {
        *ppv = (IMAPIAdviseSink*)this;
        AddRef();
        return NO_ERROR;
    }

    *ppv = NULL;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CMAPIAdviseSink::AddRef() {
    return (ULONG)InterlockedIncrement(&m_lRef);
}

STDMETHODIMP_(ULONG) CMAPIAdviseSink::Release() {
    ULONG ulCount = (ULONG)InterlockedDecrement(&m_lRef);
    if (ulCount == 0)
        delete this;

    return ulCount;
}

STDMETHODIMP_(ULONG) CMAPIAdviseSink::OnNotify(ULONG cNotif,
    LPNOTIFICATION lpNotifications)
{
    //MessageBox(0, "a", "a", 0);
    switch (lpNotifications->ulEventType) {
    case fnevCriticalError:
        OutputDebugString(TEXT("Critical Error\r\n"));
        break;
    case fnevNewMail:
        OutputDebugString(TEXT("New Mail\r\n"));
        break;
    case fnevObjectCreated:
        OutputDebugString(TEXT("Object Created\r\n"));
        break;
    case fnevObjectDeleted:
        OutputDebugString(TEXT("Object Deleted\r\n"));
        break;
    case fnevObjectModified:
        OutputDebugString(TEXT("Object Modified\r\n"));
        break;
    case fnevObjectMoved:
        OutputDebugString(TEXT("Object Moved\r\n"));
        break;
    case fnevObjectCopied:
        OutputDebugString(TEXT("Object Copied\r\n"));
        break;
    case fnevTableModified:
        OutputDebugString(TEXT("Table Modified\r\n"));
        break;
    }
    return NO_ERROR;
}