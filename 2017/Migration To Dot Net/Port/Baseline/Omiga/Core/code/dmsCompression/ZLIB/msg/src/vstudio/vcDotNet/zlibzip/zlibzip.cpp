///////////////////////////////////////////////////////////////////////////////
//	FILE:			ZLIBZIP.CPP
//	DESCRIPTION: 	
//		Provides a Win32 DLL wrapper for minizip and miniunz, which support
//		zipping and unzipping files using pkware/Winzip compression.
//		The supplied source code for minizip and miniunz generates a console 
//		application. With minimal changes to the source code, this has been
//		changed to also support a Win32 DLL; see ZLIBZIP in minizip.c and miniunz.c.
//					
//	SYSTEM:	    	ZLIBZIP
//	COPYRIGHT:		(c) 2005, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		10/10/2005	First version.
////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <stdlib.h>
#include "zlibzip.h"
#include "zlibHelper.h"

DWORD g_dwtlsIndex = TLS_OUT_OF_INDEXES;


BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
			g_dwtlsIndex = TlsAlloc();
			break;
		case DLL_THREAD_ATTACH:
			break;
		case DLL_THREAD_DETACH:
			if (g_dwtlsIndex != TLS_OUT_OF_INDEXES)
			{
				char* error = static_cast<char*>(TlsGetValue(g_dwtlsIndex));
				if (error != NULL)
				{
					delete []error;
				}
			}
			break;
		case DLL_PROCESS_DETACH:
			if (g_dwtlsIndex != TLS_OUT_OF_INDEXES)
			{
				TlsFree(g_dwtlsIndex);
			}
			break;
    }
    return TRUE;
}

char* GetTlsErrorBuffer()
{
	char* error = NULL;

	if (g_dwtlsIndex != TLS_OUT_OF_INDEXES)
	{
		error = static_cast<char*>(TlsGetValue(g_dwtlsIndex));
		if (error == NULL)
		{
			error = new char[MAXLEN_ERROR];
			if (error != NULL)
			{
				memset(error, 0, MAXLEN_ERROR * sizeof(error[0]));
				TlsSetValue(g_dwtlsIndex, error);
			}
		}
	}

	return error;
}


// main is defined in minizip.c.
extern "C" int main_MINIZIP(int, char*[]);
extern "C" int main_MINIUNZ(int, char*[]);
extern "C" int main_MINIUNZEX(int, char*[],	char*, long*, long*);

static const int MAXLEN_ARGS = 7;
static const int MAXLEN_ARG = _MAX_PATH;


///////////////////////////////////////////////////////////////////////////////
//	ZipFile:
//		Zip a file in pkware/Winzip format.
///////////////////////////////////////////////////////////////////////////////
ZLIBZIP_API int minizipFile(
	const wchar_t* zipFileName,	// Zip file.
	const wchar_t* fileName,	// File to zip.
	BOOL overwrite,				// Overwrite existing zip file?
	BOOL append,				// Append to existing zip file?
	int compressLevel,			// -1 = default, 0 = store only, 1 = faster, 9 = better.
	const wchar_t* password,	// Password may be NULL.
	BOOL nopaths,				// Do not include file path name in zip file?
	wchar_t* error,				// Receives any error.
	int maxlenError)
{
	char* argv[MAXLEN_ARGS];

	USES_CONVERSION;

	memset(error, 0, maxlenError * sizeof(error[0]));
	
	int argc = 0;
	static char* argProgramName = "zlibzip.dll";
	static char* argOverwrite = "-o";
	static char* argAppend = "-a";
	static char* argCompressLevel = "-1";
	static char* argPassword = "-p";
	static char* argNoPaths = "-n";

	// Must overwrite or append if file exists, so default to overwrite. 
	// This prevents scanf being called in minizip.c, which would require user interaction.
	if (!overwrite && !append) overwrite = TRUE;

	argv[argc++] = argProgramName;
	if (overwrite) argv[argc++] = argOverwrite;
	if (append) argv[argc++] = argAppend;
	if (compressLevel >= 0 && compressLevel <= 9)
	{
		argCompressLevel[1] = '0' + compressLevel;
		argv[argc++] = argCompressLevel;
	}
	if (password != NULL)
	{
		argv[argc++] = argPassword;
		argv[argc++] = W2A(password);
	}
	if (nopaths) argv[argc++] = argNoPaths;

	argv[argc++] = W2A(zipFileName);
	argv[argc++] = W2A(fileName);

	int result = main_MINIZIP(argc, argv);

	if (g_dwtlsIndex != TLS_OUT_OF_INDEXES)
	{
		char* tls_Error = static_cast<char*>(TlsGetValue(g_dwtlsIndex));
		if (tls_Error != NULL)
		{
			wcsncpy(error, A2W(tls_Error), maxlenError);
		}
	}
		
	return result;
}


///////////////////////////////////////////////////////////////////////////////
//	UnzipFile:
//		Unzip a file in pkware/Winzip format.
///////////////////////////////////////////////////////////////////////////////
ZLIBZIP_API int miniunzFile(
	const wchar_t* zipFileName,	// Zip file.
	const wchar_t* fileName,	// File to extract; may be NULL for all files.
	BOOL extractWithout,		// Extract without pathname (junk paths)?
	BOOL extractWith,			// Extract with pathname?
	const wchar_t* password,	// Password may be NULL.
	const wchar_t* unzipDir,	// Location in which to unzip files.
	char* pUnzippedFiles,		// Output buffer to hold unzipped files.
	long* plUnzippedFilesMaxSize,// Maximum size of output buffer.
	long* plUnzippedFilesCount,	// Number of unzipped files.
	wchar_t* error,				// Receives any error.
	int maxlenError)
{
	char* argv[MAXLEN_ARGS];

	USES_CONVERSION;

	memset(error, 0, maxlenError * sizeof(error[0]));
	
	int argc = 0;
	static char* argProgramName = "zlibzip.dll";
	static char* argExtractWithout = "-e";
	static char* argExtractWith = "-x";
	static char* argOverwrite = "-o";
	static char* argPassword = "-p";
	static char* argUnzipDir = "-d";

	argv[argc++] = argProgramName;
	// Always overwrite; prevents the user from being prompted to confirm overwrite.
	argv[argc++] = argOverwrite;	
	if (extractWithout) argv[argc++] = argExtractWithout;
	if (extractWith) argv[argc++] = argExtractWith;
	if (unzipDir != NULL && wcslen(unzipDir) > 0)
	{
		argv[argc++] = argUnzipDir;
		argv[argc++] = W2A(unzipDir);
	}
	if (password != NULL && wcslen(password) > 0)
	{
		argv[argc++] = argPassword;
		argv[argc++] = W2A(password);
	}

	argv[argc++] = W2A(zipFileName);
	if (fileName != NULL && wcslen(fileName) > 0) argv[argc++] = W2A(fileName);

	int result = main_MINIUNZEX(argc, argv, pUnzippedFiles, plUnzippedFilesMaxSize,	plUnzippedFilesCount);

	if (g_dwtlsIndex != TLS_OUT_OF_INDEXES)
	{
		char* tls_Error = static_cast<char*>(TlsGetValue(g_dwtlsIndex));
		if (tls_Error != NULL)
		{
			wcsncpy(error, A2W(tls_Error), maxlenError);
		}
	}
		
	return result;
}
