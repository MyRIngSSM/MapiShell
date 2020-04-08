#include <Windows.h>
#include "../MapiFiles/MAPIShell.h"
#include "../MapiFiles/MapiUtil.h"
#include <vector>
#include <atlstr.h>

LPMAPIINITIALIZE pfnMAPIInitialize = NULL;
LPMAPIUNINITIALIZE pfnMAPIUninitialize = NULL;

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	LPMAPISESSION FAR lpMAPISession = NULL;
	//MessageBox(0, L"123", L"456", 0);
	BOOL ret = InitializeSession(lpMAPISession);
	if (ret == FALSE)
		return FALSE;

	LPMDB lpMDB = NULL;
	LPMAPIFOLDER lpFolder = NULL;
	
	
	
	
	HRESULT hRes = OpenDefaultMessageStore(lpMAPISession, &lpMDB, FALSE);

	
	hRes = OpenInbox(lpMDB, &lpFolder);
	RegisterNewMessage(lpMDB, lpFolder);
	//ListMessages(lpMDB, lpFolder, "Test Subject");
	std::vector<CString> lpRecipients{ "buhni99@walla.com", "sahar15926@outlook.com" };
    SendMail(lpMAPISession, lpMDB, "Test Subject", "Test Body", lpRecipients, "Sahar");
	while (1) {
		Sleep(1000);
	}
    lpMAPISession->Release();

	MAPIUninitialize();

	return 1;
}
