#pragma once
#include "MAPIShell.h"
#include <atlstr.h>
#include <vector>

STDMETHODIMP OpenDefaultMessageStore(LPMAPISESSION lpMAPISession, LPMDB* lpMDB);
STDMETHODIMP SendMail(LPMAPISESSION lpMAPISession, CString szSubject, CString szBody, std::vector<CString> lpRecipients, CString szSenderName);
STDMETHODIMP OpenInbox(LPMDB lpMDB, LPMAPIFOLDER* lpInboxFolder);
STDMETHODIMP OpenFolder(LPMDB lpMDB, LPMAPIFOLDER* lpFolder, ULONG entryId);
STDMETHODIMP BuildEmail(LPMAPISESSION lpMAPISession, LPMDB lpMDB, LPMAPIPROP lpMessage, CString szSubject, CString szBody, std::vector<CString> lpRecipients, CString szSenderName);
STDMETHODIMP SetPropertyString(LPMAPIPROP lpProp, ULONG ulProperty, CString szProperty);
STDMETHODIMP SetPropertyStream(LPMAPIPROP lpProp, ULONG ulProperty, CString szProperty);
STDMETHODIMP DeleteAfterSubmit(LPMAPIPROP lpMessage);
STDMETHODIMP AddRecipient(LPMAPISESSION lpMAPISession, LPMAPIPROP lpMessage, CString szRecipient);
STDMETHODIMP ListMessages(LPMDB lpMDB, LPMAPIFOLDER lpFolder, CString szSubject);