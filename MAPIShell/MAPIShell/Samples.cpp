#include "../MapiFiles/MAPIShell.h"
#include "../MapiFiles/MapiUtil.h"
#include "../MapiFiles/MapiNotify.h"

STDMETHODIMP ListMessages(LPMDB lpMDB, LPMAPIFOLDER lpFolder, CString szSubject)
{
	HRESULT              hRes = S_OK;
	LPMAPITABLE          lpContentsTable = NULL;
	LPSRowSet            pRows = NULL;
	LPSTREAM             lpStream = NULL;
	ULONG                i;
	static SRestriction  sres;
	SPropValue           spv;

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