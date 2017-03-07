// zlibzipTest.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include <crtdbg.h>
#include <stdio.h>
#include <stdlib.h>
#include "..\zlibzip\zlibzip.h"
#include "..\zlibwapi\unzip.h"

int main(int argc, char* argv[])
{
	const int maxlenError = 1024;
	wchar_t error[maxlenError] = L""; 

	int result = 
		minizipFile(
			L"C:\\projects\\BaselineCode\\Core\\dmsCompression\\ZLIB\\msg\\src\\vstudio\\vc6\\zlibzipTest\\Debug\\hyper.zip", 
			L"C:\\projects\\BaselineCode\\Core\\dmsCompression\\ZLIB\\msg\\src\\vstudio\\vc6\\zlibzipTest\\Debug\\hyper.doc", 
			FALSE,				// Overwrite existing zip file?
			FALSE,				// Append to existing zip file?
			-1,					// 0 = store only, 1 = faster, 9 = better.
			NULL,
			TRUE,				// No paths in zip file.
			error,
			maxlenError);

	wprintf(error);
	wprintf(L"\n");
	/*

	result = 
		miniunzFile(
			L"C:\\projects\\BaselineCode\\Core\\dmsCompression\\ZLIB\\msg\\src\\vstudio\\vc6\\zlibzipTest\\Debug\\FilePrintComparison.zip", 
			NULL, 
			FALSE,				// Extract without pathname (junk paths)?
			FALSE,				// Extract with pathname?
			NULL,				// Password may be NULL.
			L"C:\\projects\\BaselineCode\\Core\\dmsCompression\\ZLIB\\msg\\src\\vstudio\\vc6\\zlibzipTest\\Debug\\Temp",
			NULL,
			NULL,
			NULL,
			error,
			maxlenError);

	wprintf(error);
	wprintf(L"\n");

	char* pUnzippedFiles = NULL;
	long lUnzippedFilesMaxSize = 0xFFFF;
	long lUnzippedFilesCount = 0;

	do
	{
		lUnzippedFilesCount = 0;

		if (pUnzippedFiles != NULL)
		{
			_ASSERTE(_CrtIsValidHeapPointer(pUnzippedFiles));
			free(pUnzippedFiles);
		}

		pUnzippedFiles = (char*)malloc(lUnzippedFilesMaxSize);
        
		_ASSERTE(_CrtIsValidHeapPointer(pUnzippedFiles));

		result = 
			miniunzFile(
				//L"C:\\projects\\BaselineCode\\Core\\dmsCompression\\ZLIB\\msg\\src\\vstudio\\vc6\\zlibzipTest\\Debug\\Temp.zip", 
				L"C:\\projects\\BaselineCode\\Core\\dmsCompression\\ZLIB\\msg\\src\\vstudio\\vc6\\zlibzipTest\\Debug\\FilePrintComparison.zip", 
				NULL, 
				FALSE,					// Extract without pathname (junk paths)?
				FALSE,					// Extract with pathname?
				NULL,					// Password may be NULL.
				NULL,					// Location in which to unzip files.
				pUnzippedFiles,			// Output buffer to hold unzipped files.
				&lUnzippedFilesMaxSize,	// Maximum size of output buffer.
				&lUnzippedFilesCount,	// Number of unzipped files.
				error,
				maxlenError);
        _ASSERTE(_CrtIsValidHeapPointer(pUnzippedFiles));

	} while (result == UNZ_MOREMEM);

	_ASSERTE(_CrtCheckMemory());

	char* pUnzippedFilesTemp = pUnzippedFiles;
	for (int unzipped = 0; unzipped < lUnzippedFilesCount; unzipped++)
	{
		char pathName[_MAX_PATH];
		char fileName[_MAX_PATH];

		strcpy(fileName, pUnzippedFilesTemp);
		char* pfileName = NULL;
		if ((pfileName = strrchr(fileName, '\\')) != NULL)
		{
			pfileName++;
		}
		else if ((pfileName = strrchr(fileName, '/')) != NULL)
		{
			pfileName++;
		}
		else
		{
			pfileName = fileName;
		}

		wsprintf(pathName, "C:\\projects\\BaselineCode\\Core\\dmsCompression\\ZLIB\\msg\\src\\vstudio\\vc6\\zlibzipTest\\Debug\\Temp\\%s", pfileName);
		pUnzippedFilesTemp += strlen(pUnzippedFilesTemp) + 1;
		long* pnUnzippedFileSize = (long*)pUnzippedFilesTemp;
		pUnzippedFilesTemp += sizeof(long);
		FILE* fp = fopen(pathName, "wb");
		fwrite(pUnzippedFilesTemp, *pnUnzippedFileSize, 1, fp);
		fclose(fp);
		pUnzippedFilesTemp += *pnUnzippedFileSize;
	}

	free(pUnzippedFiles);

	wprintf(error);
	wprintf(L"\n");
*/

	return result;
}
