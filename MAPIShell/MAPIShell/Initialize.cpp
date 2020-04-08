#include "../MapiFiles/MAPIShell.h"

BOOL InitializeSession(LPMAPISESSION FAR & lpMAPISession)
{
	HRESULT hr;
	MAPIINIT_0 mapiInit = { MAPI_INIT_VERSION, NULL };
	hr = MAPIInitialize(&mapiInit);

	if (hr != S_OK)
		return FALSE;

	hr = MAPILogonEx(0, NULL, NULL, MAPI_ALLOW_OTHERS | MAPI_BG_SESSION | MAPI_EXTENDED | MAPI_UNICODE | MAPI_USE_DEFAULT, &lpMAPISession);
	if (hr != S_OK)
		return FALSE;

	return TRUE;
}
