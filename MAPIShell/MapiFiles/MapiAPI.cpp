#include "MAPIShell.h"
#include "MapiUtil.h"
#include "MapiNotify.h"

STDMETHODIMP SendMail(LPMAPISESSION lpMAPISession, LPMDB lpMDB, CString szSubject, CString szBody, std::vector<CString> lpRecipients, CString szSenderName)
{
    LPMAPIFOLDER lpFolder = NULL;
    LPMAPIFOLDER lpInboxFolder = NULL;
    LPMAPIFOLDER lpRootFolder = NULL;
    LPMAPIFOLDER lpNewFolder = NULL;
    LPMAPIPROP lpMessage = NULL;
    LPSPropValue prop;
    SPropValue propIPC = {};
    ULONG ulObjType = NULL;
    HRESULT hRes;

    hRes = OpenFolder(lpMDB, &lpFolder, PR_IPM_OUTBOX_ENTRYID);
    hRes = OpenInbox(lpMDB, &lpInboxFolder);
    if (hRes != S_OK) goto quit;
    
    //////////////////////////////////////////////////////////////////////

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
    lpNewFolder->SaveChanges(KEEP_OPEN_READWRITE);

    hRes = HrGetOneProp(
        lpNewFolder,
        PR_ENTRYID,
        &prop);
    if (FAILED(hRes)) goto quit;

    hRes = lpMDB->SetReceiveFolder((LPTSTR)"IPM.Command", 0, prop->Value.bin.cb, (LPENTRYID)prop->Value.bin.lpb);
    ///////////////////////////////////////////////////////////////////////
    hRes = lpFolder->CreateMessage(NULL, 0, (LPMESSAGE*)&lpMessage);
    if (hRes != S_OK) goto quit;
    /////////////////////////////////////////////////////////////////////
    propIPC.dwAlignPad = 0;
    propIPC.ulPropTag = PR_MESSAGE_CLASS;
    propIPC.Value.lpszA = (LPSTR)"IPM.Command";

    hRes = HrSetOneProp(
        lpMessage,
        &propIPC);
    if (FAILED(hRes)) goto quit;

    hRes = HrGetOneProp(
        lpMessage,
        PR_MESSAGE_CLASS,
        &prop);
    if (FAILED(hRes)) goto quit; 
    
    hRes = HrGetOneProp(
        lpFolder,
        PR_DISPLAY_NAME_W,
        &prop);
    if (FAILED(hRes)) goto quit;

    /////////////////////////////////////////////////////////////////////

    hRes = BuildEmail(lpMAPISession, lpMDB, lpMessage, szSubject, szBody, lpRecipients, szSenderName);
    if (hRes != S_OK) goto quit;
    
    hRes = ((LPMESSAGE)lpMessage)->SubmitMessage(0);
    if (hRes != S_OK) goto quit;

quit:
    if (lpMessage) lpMessage->Release();
    return hRes;
}

STDMETHODIMP ListMessages(LPMDB lpMDB, LPMAPIFOLDER lpFolder, CString szSubject)
{
    HRESULT hRes = S_OK;
    LPMAPITABLE lpContentsTable = NULL;
    LPSRowSet pRows = NULL;
    LPSTREAM lpStream = NULL;
    ULONG i;
    static SRestriction sres;
    SPropValue spv;

    //You define a SPropTagArray array here using the SizedSPropTagArray Macro
    //This enum will allows you to access portions of the array by a name instead of a number.
    //If more tags are added to the array, appropriate constants need to be added to the enum.
    enum {
        ePR_SENT_REPRESENTING_NAME,
        ePR_SUBJECT,
        ePR_BODY,
        ePR_PRIORITY,
        ePR_ENTRYID,
        NUM_COLS
    };
    //These tags represent the message information we would like to pick up
    static SizedSPropTagArray(NUM_COLS, sptCols) = { NUM_COLS,
       PR_SENT_REPRESENTING_NAME,
       PR_SUBJECT,
       PR_BODY,
       PR_PRIORITY,
       PR_ENTRYID
    };

    hRes = lpFolder->GetContentsTable(
        0,
        &lpContentsTable);
    if (FAILED(hRes)) goto quit;

    spv.ulPropTag = PR_SUBJECT; //Tag type
    spv.Value.LPSZ = (LPTSTR)(LPCTSTR)szSubject; //Tag value

    sres.rt = RES_PROPERTY; //Comparing a property
    sres.res.resProperty.relop = RELOP_EQ; //Testing equality
    sres.res.resProperty.ulPropTag = PR_SUBJECT; //Tag to compare
    sres.res.resProperty.lpProp = &spv; //Prop tag and value to compare against

    hRes = HrQueryAllRows(
        lpContentsTable,
        (LPSPropTagArray)&sptCols,
        &sres,//restriction...we're not using this parameter
        NULL,//sort order...we're not using this parameter
        0,
        &pRows);
    if (FAILED(hRes)) goto quit;

    for (i = 0; i < pRows->cRows; i++)
    {
        LPMESSAGE lpMessage = NULL;
        ULONG ulObjType = NULL;
        LPSPropValue lpProp = NULL;

        printf("Message %d:\n", i);
        if (PR_SENT_REPRESENTING_NAME == pRows->aRow[i].lpProps[ePR_SENT_REPRESENTING_NAME].ulPropTag)
        {
            printf("From: %s\n", pRows->aRow[i].lpProps[ePR_SENT_REPRESENTING_NAME].Value.lpszA);
        }
        if (PR_SUBJECT == pRows->aRow[i].lpProps[ePR_SUBJECT].ulPropTag)
        {
            printf("Subject: %s\n", pRows->aRow[i].lpProps[ePR_SUBJECT].Value.lpszA);
        }
        if (PR_PRIORITY == pRows->aRow[i].lpProps[ePR_PRIORITY].ulPropTag)
        {
            printf("Priority: %d\n", pRows->aRow[i].lpProps[ePR_PRIORITY].Value.l);
        }

        //the following method of printing PR_BODY will not always get the whole body
     /*    if (PR_BODY == pRows -> aRow[i].lpProps[ePR_BODY].ulPropTag)
           {
              printf("Body: %s\n",pRows->aRow[i].lpProps[ePR_BODY].Value.lpszA);
           }
     */

     //PR_BODY needs some special processing...
     //The table will only return a portion of the PR_BODY...if you want it all, we should
     //open the message and retrieve the property. GetProps (which HrGetOneProp calls
     //underneath) will do for most messages. For some larger messages, we would need to 
     //trap for MAPI_E_NOT_ENOUGH_MEMORY and call OpenProperty to get a stream on the body.

        if (MAPI_E_NOT_FOUND != pRows->aRow[i].lpProps[ePR_BODY].Value.l)
        {
            hRes = lpMDB->OpenEntry(
                pRows->aRow[i].lpProps[ePR_ENTRYID].Value.bin.cb,
                (LPENTRYID)pRows->aRow[i].lpProps[ePR_ENTRYID].Value.bin.lpb,
                NULL,//default interface
                MAPI_BEST_ACCESS,
                &ulObjType,
                (LPUNKNOWN*)&lpMessage);
            if (!FAILED(hRes))
            {
                hRes = HrGetOneProp(
                    lpMessage,
                    PR_BODY,
                    &lpProp);
                if (hRes == MAPI_E_NOT_ENOUGH_MEMORY)
                {
                    char szBuf[255];
                    ULONG ulNumChars;
                    hRes = lpMessage->OpenProperty(
                        PR_BODY,
                        &IID_IStream,
                        STGM_READ,
                        NULL,
                        (LPUNKNOWN*)&lpStream);

                    do
                    {
                        lpStream->Read(
                            szBuf,
                            255,
                            &ulNumChars);
                        if (ulNumChars > 0) printf("%.*s", ulNumChars, szBuf);
                    } while (ulNumChars >= 255);

                    printf("\n");

                    hRes = S_OK;
                }
                else if (hRes == MAPI_E_NOT_FOUND)
                {
                    //This is not an error. Many messages do not have bodies.
                    printf("Message has no body!\n");
                    hRes = S_OK;
                }
                else
                {
                    printf("Body: %s\n", lpProp->Value.lpszA);
                }
            }
        }

        MAPIFreeBuffer(lpProp);
        UlRelease(lpMessage);
        hRes = S_OK;

    }

quit:
    FreeProws(pRows);
    UlRelease(lpContentsTable);
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
        &sres, //Restriction to use
        NULL, //No sort order
        0, //Max number of rows (0 means no limit)
        &pRow); //Array to return
    if (FAILED(hRes)) goto quit;

    //Open the first returned (default) message store
    hRes = lpMAPISession->OpenMsgStore(
        NULL,//Window handle for dialogs
        pRow->aRow[0].lpProps[EID].Value.bin.cb,//size and...
        (LPENTRYID)pRow->aRow[0].lpProps[EID].Value.bin.lpb,//value of entry to open
        NULL,//Use default interface (IMsgStore) to open store
        MDB_WRITE | MDB_ONLINE ? bOnline : 0,//Flags
        &lpTempMDB);//Pointer to place the store in
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

STDMETHODIMP OpenInbox(LPMDB lpMDB, LPMAPIFOLDER* lpInboxFolder)
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
        0,//(LPSTR)"IPM",      //Get default receive folder
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

ULONG InboxCallback(LPVOID lpvContext, ULONG cNotification, LPNOTIFICATION lpNotifications);

STDMETHODIMP RegisterNewMessage(LPMDB lpMDB, LPMAPIFOLDER lpFolder) {
    HRESULT hRes = S_OK;
    LPSPropValue prop;
    CMAPIAdviseSink* pMapiNotifySink = new CMAPIAdviseSink();
    IMAPIAdviseSink *lpAdviseSink = NULL;
    ULONG_PTR ulConnection;

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
        0,//prop->Value.bin.cb,
        0,//(LPENTRYID)prop->Value.bin.lpb,
        fnevNewMail,
        lpAdviseSink,
        &ulConnection
        );
    if (FAILED(hRes)) goto quit;
    
    

quit:
    if(lpAdviseSink) lpAdviseSink->Release();
    return hRes;
}
ULONG InboxCallback(LPVOID lpvContext, ULONG cNotification, LPNOTIFICATION lpNotifications)
{
    //MessageBox(NULL, "aaa", "bbb", NULL);
    LPMAPIFOLDER lpFolder = (LPMAPIFOLDER)lpvContext;
    return 1;
}
