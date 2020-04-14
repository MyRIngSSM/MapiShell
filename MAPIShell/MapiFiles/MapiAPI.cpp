#include "MAPIShell.h"
#include "MapiUtil.h"
#include "MapiNotify.h"

STDMETHODIMP SendMail(LPMAPISESSION lpMAPISession, LPMDB lpMDB, CString szSubject, CString szBody, std::vector<CString> lpRecipients, CString szSenderName, CString szPath)
{
	LPMAPIFOLDER lpFolder = NULL;
	LPMAPIPROP   lpMessage = NULL;
	SPropValue   prop = {};
	HRESULT      hRes;

	// First we need to open the outbox folder
	hRes = OpenFolder(lpMDB, &lpFolder, PR_IPM_OUTBOX_ENTRYID);
	if (hRes != S_OK) goto quit;

	// Create a new message inside the outbox
	hRes = lpFolder->CreateMessage(NULL, 0, (LPMESSAGE*)&lpMessage);
	if (hRes != S_OK) goto quit;

	// Change message class to IPM.Command so it will be saved in the hidden folder
	prop.dwAlignPad = 0;
	prop.ulPropTag = PR_MESSAGE_CLASS;
	prop.Value.lpszA = (LPSTR)"IPM.Command";

	hRes = HrSetOneProp(
		lpMessage,
		&prop);
	if (FAILED(hRes)) goto quit;

	// Set the message's subject, body and recipients
	hRes = BuildEmail(lpMAPISession, lpMDB, lpMessage, szSubject, szBody, lpRecipients, szSenderName);
	if (hRes != S_OK) goto quit;
	AddAttachment(lpMessage, szPath);
	// Send the message
	hRes = ((LPMESSAGE)lpMessage)->SubmitMessage(0);
	if (hRes != S_OK) goto quit;

quit:
	if (lpMessage) lpMessage->Release();
	return hRes;
}

STDMETHODIMP OpenDefaultMessageStore(LPMAPISESSION lpMAPISession, LPMDB* lpMDB, BOOL bOnline)
{
	LPMAPITABLE pStoresTbl = NULL;
	LPSRowSet   pRow = NULL;
	static      SRestriction sres;
	SPropValue  spv;
	HRESULT     hRes;
	LPMDB       lpTempMDB = NULL;

	enum { EID, NAME, NUM_COLS };
	static SizedSPropTagArray(NUM_COLS, sptCols) = { NUM_COLS, PR_ENTRYID, PR_DISPLAY_NAME };

	*lpMDB = NULL;

	//Get the table of all the message stores available
	hRes = lpMAPISession->GetMsgStoresTable(0, &pStoresTbl);
	if (FAILED(hRes)) goto quit;

	//Set up restriction for the default store
	sres.rt = RES_PROPERTY; //Comparing a property
	sres.res.resProperty.relop = RELOP_EQ; //Testing equality
	sres.res.resProperty.ulPropTag = PR_DEFAULT_STORE; //Tag to compare
	sres.res.resProperty.lpProp = &spv; //Prop tag and value to compare against

	spv.ulPropTag = PR_DEFAULT_STORE; //Tag type
	spv.Value.b = TRUE; //Tag value

	//Convert the table to an array which can be stepped through
	//Only one message store should have PR_DEFAULT_STORE set to true, so only one will be returned
	hRes = HrQueryAllRows(
		pStoresTbl, //Table to query
		(LPSPropTagArray)&sptCols, //Which columns to get
		&sres,   //Restriction to use
		NULL,    //No sort order
		0,       //Max number of rows (0 means no limit)
		&pRow);  //Array to return
	if (FAILED(hRes)) goto quit;

	//Open the first returned (default) message store
	hRes = lpMAPISession->OpenMsgStore(
		NULL,                                                //Window handle for dialogs
		pRow->aRow[0].lpProps[EID].Value.bin.cb,             //size and...
		(LPENTRYID)pRow->aRow[0].lpProps[EID].Value.bin.lpb, //value of entry to open
		NULL,                                                //Use default interface (IMsgStore) to open store
		MDB_WRITE | (MDB_ONLINE ? bOnline : 0),              //Flags
		&lpTempMDB);                                         //Pointer to place the store in
	if (FAILED(hRes)) goto quit;

	//Assign the out parameter
	*lpMDB = lpTempMDB;

	//Always clean up your memory here!
quit:
	FreeProws(pRow);
	UlRelease(pStoresTbl);
	if (FAILED(hRes))
	{
		HRESULT hr;
		LPMAPIERROR lpError;
		hr = lpMAPISession->GetLastError(hRes, 0, &lpError);
		if (!hr)
		{
			MAPIFreeBuffer(lpError);
		}
	}
	return hRes;
}

STDMETHODIMP OpenInbox(LPMDB lpMDB, LPMAPIFOLDER* lpInboxFolder, LPSTR lpMessageClass)
{
	ULONG        cbInbox;
	LPENTRYID    lpbInbox;
	ULONG        ulObjType;
	HRESULT      hRes = S_OK;
	LPMAPIFOLDER lpTempFolder = NULL;
	LPSTR        lppszExplicitClass;
	*lpInboxFolder = NULL;

	//The Inbox is usually the default receive folder for the message store
	//You call this function as a shortcut to get it's Entry ID
	hRes = lpMDB->GetReceiveFolder(
	(LPTSTR)lpMessageClass,//(LPSTR)"IPM",      //Get default receive folder
		NULL,      //Flags
		&cbInbox,  //Size and ...
		&lpbInbox, //Value of the EntryID to be returned
		(LPTSTR*)&lppszExplicitClass);     //You don't care to see the class returned
	if (FAILED(hRes)) goto quit;

	hRes = lpMDB->OpenEntry(
		cbInbox,                      //Size and...
		lpbInbox,                     //Value of the Inbox's EntryID
		NULL,                         //We want the default interface    (IMAPIFolder)
		MAPI_MODIFY,             //Flags
		&ulObjType,                   //Object returned type
		(LPUNKNOWN*)&lpTempFolder); //Returned folder
	if (FAILED(hRes)) goto quit;

	//Assign the out parameter
	*lpInboxFolder = lpTempFolder;

	//Always clean up your memory here!
quit:
	MAPIFreeBuffer(lpbInbox);
	return hRes;
}

STDMETHODIMP OpenFolder(LPMDB lpMDB, LPMAPIFOLDER* lpFolder, ULONG entryId)
{
	ULONG        cbInbox = 0;
	ULONG        ulObjType;
	HRESULT      hRes = S_OK;
	LPMAPIFOLDER lpTempFolder = NULL;

	*lpFolder = NULL;

	//The Inbox is usually the default receive folder for the message store
	//You call this function as a shortcut to get it's Entry ID
	SPropTagArray lpPropTagArray = { 1, entryId };
	ULONG lpcValues;
	LPSPropValue lpPropArray = NULL;

	hRes = lpMDB->GetProps(
		&lpPropTagArray,
		PT_UNSPECIFIED,
		&lpcValues,
		&lpPropArray);
	if (FAILED(hRes)) goto quit;

	hRes = lpMDB->OpenEntry(
		lpPropArray->Value.bin.cb,                      //Size and...
		(LPENTRYID)lpPropArray->Value.bin.lpb,                     //Value of the Inbox's EntryID
		NULL,                         //We want the default interface    (IMAPIFolder)
		MAPI_BEST_ACCESS,             //Flags
		&ulObjType,                   //Object returned type
		(LPUNKNOWN*)&lpTempFolder); //Returned folder
	if (FAILED(hRes)) goto quit;

	//Assign the out parameter
	*lpFolder = lpTempFolder;

	//Always clean up your memory here!
quit:
	MAPIFreeBuffer(lpPropArray);
	return hRes;
}

STDMETHODIMP RegisterNewMessage(LPMDB lpMDB, LPMAPIFOLDER lpFolder) {
	HRESULT             hRes = S_OK;
	LPSPropValue        prop;
	//CMAPIAdviseSink* pMapiNotifySink = new CMAPIAdviseSink();
	IMAPIAdviseSink* lpAdviseSink = NULL;
	ULONG_PTR            ulConnection;

	/*hRes = pMapiNotifySink->QueryInterface(
		IID_IMAPIAdviseSink,
		(VOID**)&lpAdviseSink);
	if (FAILED(hRes)) goto quit; */
	hRes = HrAllocAdviseSink(
	(LPNOTIFCALLBACK)InboxCallback,
		lpMDB,
		&lpAdviseSink);
	if (FAILED(hRes)) goto quit;

	//ZeroMemory(*lpAdviseSink, sizeof(IMAPIAdviseSink));
	hRes = HrGetOneProp(
		lpFolder,
		PR_ENTRYID,
		&prop);
	if (FAILED(hRes)) goto quit;

	hRes = lpMDB->Advise(
		prop->Value.bin.cb,
		(LPENTRYID)prop->Value.bin.lpb,
		fnevNewMail,
		lpAdviseSink,
		&ulConnection
		);
	if (FAILED(hRes)) goto quit;



quit:
	if (lpAdviseSink) lpAdviseSink->Release();
	return hRes;
}

STDMETHODIMP ProcessMessage(LPMDB lpMDB, LPMESSAGE lpMessage, CString& szSubject, CString& szBody, std::vector<CString>& vszAttachment)
{
	HRESULT		 hRes = S_OK;
	ULONG		 ulValus;
	LPSPropValue lpPropArray = NULL;
	ULONG		 ulObjType;
	LPSPropValue lpProp = NULL;
	LPMAPITABLE  lpAttachmentTable;
	LPSRowSet    pRows = NULL;

	enum {
		ePR_SUBJECT,
		ePR_BODY,
		ePR_ENTRYID,
		NUM_COLS
	};

	enum {
		ePR_ATTACH_NUM,
		ePR_ATTACH_METHOD,
		NUM_COLS_
	};

	//These tags represent the message information we would like to pick up
	static SizedSPropTagArray(NUM_COLS, sptCols) = { NUM_COLS,
	   PR_SUBJECT,
	   PR_BODY,
	   PR_ENTRYID
	};

	static SizedSPropTagArray(NUM_COLS_, sptAttach) = { NUM_COLS_,
	   PR_ATTACH_NUM,
	   PR_ATTACH_METHOD
	};

	hRes = lpMessage->GetProps(
	(LPSPropTagArray)&sptCols,
		NULL,
		&ulValus,
		&lpPropArray);
	if (FAILED(hRes)) goto quit;

	szSubject = lpPropArray[ePR_SUBJECT].Value.LPSZ;

	if (MAPI_E_NOT_FOUND == lpPropArray[ePR_BODY].Value.l) goto quit;

	hRes = lpMDB->OpenEntry(
		lpPropArray[ePR_ENTRYID].Value.bin.cb,
		(LPENTRYID)lpPropArray[ePR_ENTRYID].Value.bin.lpb,
		NULL,//default interface
		MAPI_BEST_ACCESS,
		&ulObjType,
		(LPUNKNOWN*)&lpMessage);
	if (FAILED(hRes)) goto quit;

	hRes = HrGetOneProp(
		lpMessage,
		PR_BODY,
		&lpProp);
	if (hRes == MAPI_E_NOT_ENOUGH_MEMORY) {
		hRes = ReadFromStream(lpMessage, PR_BODY, szBody);
		if (FAILED(hRes)) goto quit;
	}
	else
		szBody = lpProp->Value.lpszA;

	hRes = lpMessage->GetAttachmentTable(
		0,
		&lpAttachmentTable);
	if (FAILED(hRes)) goto quit;

	hRes = HrQueryAllRows(
		lpAttachmentTable,
		(LPSPropTagArray)&sptAttach,
		NULL,//restriction...we're not using this parameter
		NULL,//sort order...we're not using this parameter
		0,
		&pRows);
	if (FAILED(hRes)) goto quit;

	for (int i = 0; i < pRows->cRows; i++)
	{
		LPATTACH lpAttach;
		ULONG ulAttachNum;
		LPSPropValue prop;

		ulAttachNum = pRows->aRow[i].lpProps[ePR_ATTACH_NUM].Value.ul;
		hRes = lpMessage->OpenAttach(ulAttachNum, NULL, MAPI_BEST_ACCESS, &lpAttach);
		if (FAILED(hRes)) continue;

		hRes = HrGetOneProp(lpAttach, PR_ATTACH_METHOD, &prop);
		if (FAILED(hRes)) continue;

		if (ATTACH_BY_VALUE != pRows->aRow[i].lpProps[ePR_ATTACH_METHOD].Value.l) continue;

		CString szAttachment = "";

		hRes = ReadFromStream(lpAttach, PR_ATTACH_DATA_BIN, szAttachment);
		if (FAILED(hRes)) continue;

		vszAttachment.push_back(szAttachment);
		hRes = S_OK;

		MAPIFreeBuffer(prop);
		lpAttach->Release();
	}
quit:
	if (lpPropArray) MAPIFreeBuffer(lpPropArray);
	return hRes;
}

ULONG InboxCallback(LPVOID lpvContext, ULONG cNotification, LPNOTIFICATION lpNotifications)
{
	HRESULT	  hRes = S_OK;
	LPMDB	  lpMDB = (LPMDB)lpvContext;
	ULONG	  cbEntryID;
	LPENTRYID lpEntryID;
	LPMESSAGE lpMessage;
	ULONG	  ulObjType;
	CString   szSubject;
	CString   szBody = "";

	std::vector<CString> vszAttachment;

	NEWMAIL_NOTIFICATION lpNewMail = lpNotifications->info.newmail;
	cbEntryID = lpNewMail.cbEntryID;
	lpEntryID = lpNewMail.lpEntryID;

	hRes = lpMDB->OpenEntry(
		cbEntryID,
		lpEntryID,
		NULL,
		MAPI_BEST_ACCESS,
		&ulObjType,
		(LPUNKNOWN*)&lpMessage);
	if (FAILED(hRes)) goto quit;

	hRes = ProcessMessage(lpMDB, lpMessage, szSubject, szBody, vszAttachment);
	if (FAILED(hRes)) goto quit;

quit:
	if (lpMessage) lpMessage->Release();
	return 1;
}
