#pragma once

#include <MAPIX.h>
#include <MAPIDefS.h>
#include <MAPIForm.h>
#include <MAPIUtil.h>
#include <MAPIAux.h>
#include <MAPI.h>
#include <MAPICode.h>
#include <MAPIDbg.h>
#include <MAPIGuid.h>
#include <MAPIHook.h>
#include <MAPINls.h>
#include <MAPIOID.h>
#include <MAPISPI.h>
#include <MAPITags.h>
#include <MAPIUtil.h>
#include <MAPIVal.h>
#include <MAPIWin.h>
#include <MAPIWz.h>
#include <MSPST.h>

extern LPMAPIINITIALIZE pfnMAPIInitialize;
extern LPMAPIUNINITIALIZE pfnMAPIUninitialize;

typedef BOOL(STDAPICALLTYPE FGETCOMPONENTPATH)
(LPSTR szComponent,
	LPSTR szQualifier,
	LPSTR szDllPath,
	DWORD cchBufferSize,
	BOOL fInstall);
typedef FGETCOMPONENTPATH FAR* LPFGETCOMPONENTPATH;

BOOL InitializeSession(LPMAPISESSION FAR &lpMAPISession);