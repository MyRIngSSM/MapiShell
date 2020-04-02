#pragma once
#include <Windows.h>
#include <atlstr.h>
#include <TlHelp32.h>

int GetProcessID(CString szProcessName);
DWORD GetResource(DWORD dwResourceId, LPVOID &lpResourceData);
BOOL DeployResource(LPVOID lpResourceData, DWORD dwSizeofData);
BOOL IsWin64bit();