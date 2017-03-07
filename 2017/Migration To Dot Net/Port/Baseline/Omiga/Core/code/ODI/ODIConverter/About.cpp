///////////////////////////////////////////////////////////////////////////////
//	FILE:			ABOUT.CPP
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS		24/09/01    First version
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "about.h"

BOOL DisplayAboutInfo()
{
	BOOL bResult = TRUE;
	static _TCHAR szModuleFileName[MAX_PATH]	= _T("\0");
	static const unsigned int uBufLen			= 256;
	static const _TCHAR* pszBlockHeaderKey		= _T("\\StringFileInfo\\080904b0\\");

	if (GetModuleFileName(NULL, szModuleFileName, _countof(szModuleFileName)) != 0)
	{
		DWORD dwHandle = 0;
		DWORD dwVerInfoSize = 0;

		if ((dwVerInfoSize = GetFileVersionInfoSize(szModuleFileName, &dwHandle)) != 0)
		{
			LPTSTR lpszVffInfo = NULL;		// Pointer to block to hold info
			HANDLE hFile = INVALID_HANDLE_VALUE;
			WIN32_FIND_DATA FindData;

			// Get a block big enough to hold version info
			lpszVffInfo	= (LPTSTR)new _TCHAR[dwVerInfoSize];

			_ASSERTE(lpszVffInfo);

			if (lpszVffInfo)
			{
				if (GetFileVersionInfo((szModuleFileName), 0L, dwVerInfoSize, lpszVffInfo))
				{
					BOOL bResult			= FALSE;
					LPVOID lpvVersion		= NULL;
					UINT uVersionLen		= 0;		// for call to VerQueryValue
					_TCHAR szBuffer[uBufLen]= _T("\0");
					_TCHAR szType[uBufLen]	= _T("\0");

					lpvVersion = NULL;
					uVersionLen = 0;
					_tcscpy(szBuffer, pszBlockHeaderKey);
					_tcscat(szBuffer, _T("ProductName"));
					bResult = 
						VerQueryValue(
							lpszVffInfo, 
							(LPTSTR)szBuffer,
							&lpvVersion,
							&uVersionLen);
					if (bResult && uVersionLen && lpvVersion)
					{
						_Module.LogDebug(_T("Product: %s\n"), (LPCTSTR)lpvVersion);
					}

					lpvVersion = NULL;
					uVersionLen = 0;
					_tcscpy(szBuffer, _T("\\"));
					bResult = 
						VerQueryValue(
							lpszVffInfo, 
							(LPTSTR)szBuffer,
							&lpvVersion,
							&uVersionLen);
					if (bResult && uVersionLen && lpvVersion)
					{
						VS_FIXEDFILEINFO* lpFixedFileInfo = (VS_FIXEDFILEINFO*)lpvVersion;
						DWORD dwNum = lpFixedFileInfo->dwFileFlags;

						// display in the file version
						_Module.LogDebug(
							_T("Version: %02d.%02d.%02d.%02d\n"),
							HIWORD(lpFixedFileInfo->dwFileVersionMS),
							LOWORD(lpFixedFileInfo->dwFileVersionMS),
							HIWORD(lpFixedFileInfo->dwFileVersionLS),
							LOWORD(lpFixedFileInfo->dwFileVersionLS));

						// File flags are bitwise or'ed so there can be more than one.
						// dwNum is used to make this easier to read
						_stprintf(
							szType,
							_T("%s%s%s%s%s"),
							(LPTSTR)(VS_FF_PRERELEASE		& dwNum ? _T("Internal ") 	: _T("")),
							(LPTSTR)(VS_FF_DEBUG			& dwNum ? _T("Debug ")   	: _T("Release ")),
							(LPTSTR)(VS_FF_PATCHED			& dwNum ? _T("Patched ") 	: _T("")),
							(LPTSTR)(VS_FF_INFOINFERRED		& dwNum ? _T("Info ")    	: _T("")),
							(LPTSTR)(0xFFFFFF00L			& dwNum ? _T("Unknown ") 	: _T("")));
					
						lpvVersion = NULL;
						uVersionLen = 0;
						_tcscpy(szBuffer, pszBlockHeaderKey);
						_tcscat(szBuffer, _T("PrivateBuild"));
						bResult = 
							VerQueryValue(
								lpszVffInfo, 
								(LPTSTR)szBuffer,
								&lpvVersion,
								&uVersionLen);
						if (bResult && uVersionLen && lpvVersion)
						{
							_tcscat(szType, _T("Private("));
							_tcscat(szType, (LPCTSTR)lpvVersion);
							_tcscat(szType, _T(") "));
						}

						lpvVersion = NULL;
						uVersionLen = 0;
						_tcscpy(szBuffer, pszBlockHeaderKey);
						_tcscat(szBuffer, _T("SpecialBuild"));
						bResult = 
							VerQueryValue(
								lpszVffInfo, 
								(LPTSTR)szBuffer,
								&lpvVersion,
								&uVersionLen);
						if (bResult && uVersionLen && lpvVersion)
						{
							_tcscat(szType, _T("Special("));
							_tcscat(szType, (LPCTSTR)lpvVersion);
							_tcscat(szType, _T(") "));
						}

						if (lstrlen(szType) == 0)
						{
							_tcscpy(szType, _T("Full"));
						}

						_Module.LogDebug(_T("Type: %s\n"), szType);
					}

					if ((hFile = FindFirstFile(szModuleFileName, &FindData)) != INVALID_HANDLE_VALUE)
					{
						FILETIME LocalTime;
						SYSTEMTIME SystemTime;

						if (
								FileTimeToLocalFileTime(&(FindData.ftCreationTime), &LocalTime) != 0 && 
								FileTimeToSystemTime(&LocalTime, &SystemTime) != 0)
						{
							_Module.LogDebug(
								_T("Date: %02d/%02d/%04d %02d:%02d:%02d\n"),
								SystemTime.wDay,
								SystemTime.wMonth,
								SystemTime.wYear,
								SystemTime.wHour,
								SystemTime.wMinute,
								SystemTime.wSecond);
						}
					}
				
					lpvVersion = NULL;
					uVersionLen = 0;
					_tcscpy(szBuffer, pszBlockHeaderKey);
					_tcscat(szBuffer, _T("LegalCopyright"));
					bResult = 
						VerQueryValue(
							lpszVffInfo, 
							(LPTSTR)szBuffer,
							&lpvVersion,
							&uVersionLen);
					if (bResult && uVersionLen && lpvVersion)
					{
						_Module.LogDebug(_T("Copyright: %s\n"), (LPCTSTR)lpvVersion);
					}
	    		}   

				delete []lpszVffInfo;
			}

			OSVERSIONINFO osvi;

			// Initialize the OSVERSIONINFO structure.
			::ZeroMemory(&osvi, sizeof(OSVERSIONINFO));
			osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);

			if (::GetVersionEx(&osvi))
			{
				const _TCHAR* pszWin	= _T("");
				DWORD dwMajorVersion	= 0;
				DWORD dwMinorVersion	= 0;
				DWORD dwBuildNumber		= 0;
				if (osvi.dwPlatformId == VER_PLATFORM_WIN32_NT)
				{
					pszWin = _T("NT");
					dwMajorVersion	= osvi.dwMajorVersion;
					dwMinorVersion	= osvi.dwMinorVersion;
					dwBuildNumber	= osvi.dwBuildNumber;
				}
				else if (osvi.dwPlatformId == VER_PLATFORM_WIN32_WINDOWS)
				{
					if (osvi.dwMajorVersion > 4 || (osvi.dwMajorVersion == 4 && osvi.dwMinorVersion > 0))
					{
						pszWin = _T("98");
					}
					else
					{
						pszWin = _T("95");
					}
					dwMajorVersion	= HIBYTE(HIWORD(osvi.dwBuildNumber));
					dwMinorVersion	= LOBYTE(HIWORD(osvi.dwBuildNumber));
					dwBuildNumber	= LOWORD(osvi.dwBuildNumber);
				}
				else if (osvi.dwPlatformId == VER_PLATFORM_WIN32s)
				{
					pszWin = _T("Win32s");
					_ASSERTE(FALSE);
				}
				_Module.LogDebug(_T("Microsoft Windows %s %d.%d Build %d %s\n\n"), 
					pszWin, dwMajorVersion, dwMinorVersion, dwBuildNumber, osvi.szCSDVersion);
			}
		}
	}
	
	return bResult;
}

