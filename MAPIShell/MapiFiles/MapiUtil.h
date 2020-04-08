#pragma once
#include "MAPIShell.h"
#include <atlstr.h>
#include <vector>

STDMETHODIMP OpenDefaultMessageStore(LPMAPISESSION lpMAPISession, LPMDB* lpMDB, BOOL bOnline);
STDMETHODIMP SendMail(LPMAPISESSION lpMAPISession, LPMDB lpMDB, CString szSubject, CString szBody, std::vector<CString> lpRecipients, CString szSenderName);
STDMETHODIMP OpenInbox(LPMDB lpMDB, LPMAPIFOLDER* lpInboxFolder, LPSTR lpMessageClass);
STDMETHODIMP OpenFolder(LPMDB lpMDB, LPMAPIFOLDER* lpFolder, ULONG entryId);
STDMETHODIMP BuildEmail(LPMAPISESSION lpMAPISession, LPMDB lpMDB, LPMAPIPROP lpMessage, CString szSubject, CString szBody, std::vector<CString> lpRecipients, CString szSenderName);
STDMETHODIMP SetPropertyString(LPMAPIPROP lpProp, ULONG ulProperty, CString szProperty);
STDMETHODIMP SetPropertyStream(LPMAPIPROP lpProp, ULONG ulProperty, CString szProperty);
STDMETHODIMP DeleteAfterSubmit(LPMAPIPROP lpMessage);
STDMETHODIMP AddRecipient(LPMAPISESSION lpMAPISession, LPMAPIPROP lpMessage, CString szRecipient);
STDMETHODIMP ListMessages(LPMDB lpMDB, LPMAPIFOLDER lpFolder, CString szSubject);
STDMETHODIMP RegisterNewMessage(LPMDB lpMDB, LPMAPIFOLDER lpFolder);
STDMETHODIMP SetReceiveFolder(LPMAPISESSION lpMAPISession);