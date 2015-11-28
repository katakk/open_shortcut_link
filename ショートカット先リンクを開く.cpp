#pragma once

#define WIN32_LEAN_AND_MEAN		// Windows ヘッダーから使用されていない部分を除外します。

#include <windows.h>
#include <stdlib.h>
#include <malloc.h>
#include <memory.h>
#include <tchar.h>

#include <winnetwk.h>
#include <shlobj.h>
#include <shellapi.h>

#define GUID_DEFS_ONLY
#include <shlguid.h>

BOOL ResolveShortcut(LPCTSTR lpszFileIn,
					 LPTSTR lpszFileOut, int cchPath)
{
	int Result = FALSE;
	SHFILEINFO info;
	WCHAR wcszFileIn[MAX_PATH];

	IShellLink   *psl = NULL;
	IPersistFile *ppf = NULL;

	*lpszFileOut = 0;

#ifndef UNICODE
	MultiByteToWideChar( CP_ACP, 0, lpszFileIn, -1, wcszFileIn, MAX_PATH );
#else
	memcpy(wcszFileIn, lpszFileIn, MAX_PATH );
#endif

	if ((SHGetFileInfo(lpszFileIn, 0, &info, sizeof(info),
		SHGFI_ATTRIBUTES) == 0) || !(info.dwAttributes & SFGAO_LINK))
	{
		return FALSE;
	}

	CoInitialize( NULL );
	if ( FAILED(CoCreateInstance( CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER,
		IID_IShellLink, (void **)&psl )) || psl == NULL)
	{
		return FALSE;
	}

	if (SUCCEEDED(psl->QueryInterface(IID_IPersistFile, (LPVOID*)&ppf)))
	{
		if (ppf != NULL && SUCCEEDED(ppf->Load(wcszFileIn, STGM_READ)))
		{
			if (SUCCEEDED(psl->Resolve(GetForegroundWindow(), SLR_ANY_MATCH)))
			{
				psl->GetPath(lpszFileOut, cchPath, NULL, SLGP_UNCPRIORITY);
				Result = TRUE;
			}
		}
	}

	if (ppf != NULL) ppf->Release();
	if (psl != NULL) psl->Release();
	CoUninitialize();
	return FALSE;
}


int APIENTRY _tWinMain(HINSTANCE hInstance,
					   HINSTANCE hPrevInstance,
					   LPTSTR    lpCmdLine,
					   int       nCmdShow)
{
	if( __argc >= 2 ) {
		LPCTSTR lpszFileIn = __targv[1];
		TCHAR szFileOut[MAX_PATH];
		TCHAR szParameters[MAX_PATH + 64];
		if (SUCCEEDED(ResolveShortcut(lpszFileIn, szFileOut, MAX_PATH))) {
			_stprintf(szParameters, _T("/select,\"%s\""), szFileOut );
			ShellExecute(GetForegroundWindow(),NULL, _T("explorer"),szParameters, NULL, SW_SHOWNORMAL);
		}
	}
	return FALSE;
}
