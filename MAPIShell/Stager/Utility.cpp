#include "Utility.h"

int GetProcessID(CString szProcessName) {
	HANDLE snaptool = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

	if (snaptool != INVALID_HANDLE_VALUE) {
		PROCESSENTRY32 pEntry;
		pEntry.dwSize = sizeof(PROCESSENTRY32);

		if (Process32First(snaptool, &pEntry)) {
			do {
				if (pEntry.szExeFile == szProcessName) {
					CloseHandle(snaptool);
					return pEntry.th32ProcessID;
				}
			} while (Process32Next(snaptool, &pEntry));
		}
	}

	CloseHandle(snaptool);
	return -1;
}

DWORD GetResource(DWORD dwResourceId, LPVOID &lpResourceData) {
	HRSRC hRsrc = FindResource(NULL, MAKEINTRESOURCE(dwResourceId), RT_RCDATA);
	if (hRsrc == NULL)
		return -1;

	HGLOBAL hGlobal = LoadResource(NULL, hRsrc);
	if (hGlobal == NULL)
		return -1;

	lpResourceData = LockResource(hGlobal);
	if (lpResourceData == NULL)
		return -1;

	return SizeofResource(NULL, hRsrc);
}

BOOL DeployResource(LPVOID lpResourceData, DWORD dwSizeofData) {
	HANDLE hFile = CreateFile(_T("C:\\temp\\MAPIShell.exe"), GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
		return FALSE;

	DWORD dwNumberOfBytesWritten;
	BOOL ret = WriteFile(hFile, lpResourceData, dwSizeofData, &dwNumberOfBytesWritten, NULL);
	if (ret == FALSE) {
		int a = GetLastError();
		CloseHandle(hFile);
		return FALSE;
	}

	CloseHandle(hFile);
	return TRUE;
}

BOOL IsWin64bit() {
	SYSTEM_INFO lpSystemInfo;

	GetNativeSystemInfo(&lpSystemInfo);

	if (lpSystemInfo.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64) {
		return 1;
	}
	else {
		return 0;
	}
}