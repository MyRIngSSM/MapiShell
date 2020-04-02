#include <Windows.h>
#include "MapiUtil.h"
#include "MAPIShell.h"
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

	HRESULT hRes = OpenDefaultMessageStore(lpMAPISession, &lpMDB);


	hRes = OpenInbox(lpMDB, &lpFolder);


	ListMessages(lpMDB, lpFolder, "Test Subject");

	std::vector<CString> lpRecipients{ "buhni99@walla.com", "sahar15926@gmail.com" };
    SendMail(lpMAPISession, "Test Subject", "Test Body", lpRecipients, "Sahar");

    lpMAPISession->Release();

	MAPIUninitialize();

	return 1;
}
