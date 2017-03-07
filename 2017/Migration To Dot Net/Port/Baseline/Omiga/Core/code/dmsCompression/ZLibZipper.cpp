///////////////////////////////////////////////////////////////////////////////
//	FILE:			Zipper.cpp
//	DESCRIPTION: 	ZLib wrapper compression class.
//	COPYRIGHT:		(c) 2004, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		12/05/04	First version
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ZLibZipper.h"
#define ZLIB_DLL
#define ZLIB_WINAPI
#include "zlib\msg\src\vstudio\vc6\zlibwapi\zlib.h"
#include "zlib\msg\src\vstudio\vc6\zlibzip\zlibzip.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CZLibZipper::CZLibZipper()
{
}


CZLibZipper::~CZLibZipper()
{
}


///////////////////////////////////////////////////////////////////////////////
//	Wrapper for third party compression routine.
///////////////////////////////////////////////////////////////////////////////
BOOL CZLibZipper::Compress()
{
	BOOL bSuccess = TRUE;
	BOOL bOwnSrcBuffer = FALSE;

//	::DebugBreak();

	if (m_pSrcBuffer == NULL)
	{
		// Source buffer has not already been initialised, 
		// e.g., the source is a file, so need to read data into source buffer. 
		m_pSrcBuffer = new char[m_ulSrcMaxLen];

		if (m_pSrcBuffer)
		{
			bOwnSrcBuffer = TRUE;
			// Read will calculate m_ulSrcLen as a side effect.
			DWORD dwBytesRead = Read(m_pSrcBuffer, m_ulSrcMaxLen);
			if (dwBytesRead == m_ulSrcMaxLen)
			{
				bSuccess = TRUE;
			}
			else
			{
				SetLastError(ERROR_NOT_ENOUGH_MEMORY);
			}
		}				
		else
		{
			SetLastError(ERROR_NOT_ENOUGH_MEMORY);
		}
	}
	else
	{
		// Source length has already been provided.
		m_ulSrcLen = m_ulSrcMaxLen;
	}

	if (bSuccess)
	{
		// Use zlib compressBound to calculate minimum target buffer size.
		// The actual compressed size may be less, once the compression has been done.
		DWORD dwTgtCompressSize = ::compressBound(m_ulSrcLen);

		// Allow space for source size in compressed buffer.
		DWORD dwTgtBufferSize = dwTgtCompressSize + sizeof(m_ulSrcLen);
		m_pTgtBuffer = new char[dwTgtBufferSize];

		if (m_pTgtBuffer)
		{
			// First DWORD of target buffer is the source buffer size.
			memcpy(m_pTgtBuffer, &m_ulSrcLen, sizeof(m_ulSrcLen));
			char* pTgtCompressBuffer = m_pTgtBuffer + sizeof(m_ulSrcLen);

			int nResult = 
				::compress(
					reinterpret_cast<Bytef*>(pTgtCompressBuffer), 
					&dwTgtCompressSize, 
					reinterpret_cast<Bytef*>(m_pSrcBuffer), 
					m_ulSrcLen);
			if (nResult == Z_OK)
			{
				// Write will calculate m_ulTgtLen as a side effect.
				// Recalculate the actual target buffer size, now that the compressed size has been
				// updated by doing the compression.
				dwTgtBufferSize = dwTgtCompressSize + sizeof(m_ulSrcLen);
				DWORD dwBytesWritten = Write(m_pTgtBuffer, dwTgtBufferSize);

				if (dwBytesWritten == m_ulTgtLen)
				{
					bSuccess = TRUE;
				}
			}
			else
			{
				SetLastError(nResult);
			}

			delete []m_pTgtBuffer;
		}
	}
	
	if (bOwnSrcBuffer)
	{
		delete []m_pSrcBuffer;
	}

	return bSuccess;
}

///////////////////////////////////////////////////////////////////////////////
//	Wrapper for third party decompression routine.
///////////////////////////////////////////////////////////////////////////////
BOOL CZLibZipper::Decompress()
{
	BOOL bSuccess = TRUE;
	BOOL bOwnSrcBuffer = FALSE;

	if (m_pSrcBuffer == NULL)
	{
		// Source buffer has not already been initialised, 
		// e.g., the source is a file, so need to read data into source buffer. 
		m_pSrcBuffer = new char[m_ulSrcMaxLen];

		if (m_pSrcBuffer)
		{
			bOwnSrcBuffer = TRUE;

			// Read will calculate m_ulSrcLen as a side effect.
			DWORD dwBytesRead = Read(m_pSrcBuffer, m_ulSrcMaxLen);
			if (dwBytesRead == m_ulSrcMaxLen)
			{
				bSuccess = TRUE;
			}
			else
			{
				SetLastError(ERROR_NOT_ENOUGH_MEMORY);
			}
		}
		else
		{
			SetLastError(ERROR_NOT_ENOUGH_MEMORY);
		}
	}
	else
	{
		// Source length has already been provided.
		m_ulSrcLen = m_ulSrcMaxLen;
	}

	if (bSuccess)
	{
		DWORD dwTgtBufferSize = 0;
		memcpy(&dwTgtBufferSize, m_pSrcBuffer, sizeof(m_ulTgtLen));

		m_pTgtBuffer = new char[dwTgtBufferSize];

		if (m_pTgtBuffer)
		{
			char* pSrcCompressBuffer = m_pSrcBuffer + sizeof(m_ulTgtLen);
			
			int nResult =
				::uncompress(
					reinterpret_cast<Bytef*>(m_pTgtBuffer), 
					&dwTgtBufferSize, 
					reinterpret_cast<Bytef*>(pSrcCompressBuffer), 
					m_ulSrcLen);

			if (nResult == Z_OK)
			{
				// Write will calculate m_ulTgtLen as a side effect.
				DWORD dwBytesWritten = Write(m_pTgtBuffer, dwTgtBufferSize);

				if (dwBytesWritten == m_ulTgtLen)
				{
					bSuccess = TRUE;
				}
			}
			else
			{
				SetLastError(nResult);
			}

			delete []m_pTgtBuffer;
		}
		else
		{
			SetLastError(ERROR_NOT_ENOUGH_MEMORY);
		}
	}

	if (bOwnSrcBuffer)
	{
		delete []m_pSrcBuffer;
	}

	return bSuccess;
}

///////////////////////////////////////////////////////////////////////////////
//	ZipFile:
//		Zip a file in pkware/Winzip format.
///////////////////////////////////////////////////////////////////////////////
BOOL CZLibZipper::ZipFile(const wchar_t* pszTgtFileName, const wchar_t* pszSrcFileName, BOOL noPaths)
{
	const int maxlenError = 1024;
	wchar_t error[maxlenError] = L""; 

	// minizipFile is defined in zlibzip.dll.
	int result = 
		minizipFile(
			pszTgtFileName,		// Zip file.
			pszSrcFileName,		// File to zip.
			TRUE,				// Overwrite existing zip file?
			FALSE,				// Append to existing zip file?
			-1,					// -1 = default, 0 = store only, 1 = faster, 9 = better.
			NULL,				// Password may be NULL.
			noPaths,			// Do not include file path name in zip file?
			error,				// Receives any error.
			maxlenError);

	if (result != 0)
	{
		SetLastErrorText(error);
	}

	return result == 0;
}

///////////////////////////////////////////////////////////////////////////////
//	UnzipFile:
//		Unzip a file in pkware/Winzip format.
///////////////////////////////////////////////////////////////////////////////
BOOL CZLibZipper::UnzipFile(
	const wchar_t* pszTgtFileName, 
	const wchar_t* pszSrcFileName, 
	const wchar_t* pszUnzipDir)
{
	const int maxlenError = 1024;
	wchar_t error[maxlenError] = L""; 

	// miniunzFile is defined in zlibzip.dll.
	int result = 
		miniunzFile(
			pszTgtFileName,			// Zip file.
			pszSrcFileName,			// File to unzip.
			FALSE,					// Extract without pathname (junk paths)?
			FALSE,					// Extract with pathname?
			NULL,					// Password may be NULL.
			pszUnzipDir,			// Location to which to unzip files.
			NULL,					// Output buffer to hold unzipped files.
			NULL,					// Maximum size of output buffer.
			NULL,					// Number of unzipped files.
			error,					// Receives any error.
			maxlenError);

	if (result != 0)
	{
		SetLastErrorText(error);
	}

	return result == 0;
}


///////////////////////////////////////////////////////////////////////////////
//	UnzipFileToBuffer:
//		Unzip a file in pkware/Winzip format to a buffer.
///////////////////////////////////////////////////////////////////////////////
int CZLibZipper::UnzipFileToBuffer(
	const wchar_t* pszTgtFileName, 
	const wchar_t* pszSrcFileName, 
	char* pUnzippedFiles,
	long* plUnzippedFilesMaxSize,
	long* plUnzippedFilesCount)
{
	const int maxlenError = 1024;
	wchar_t error[maxlenError] = L""; 

	// miniunzFile is defined in zlibzip.dll.
	int result = 
		miniunzFile(
			pszTgtFileName,			// Zip file.
			pszSrcFileName,			// File to unzip.
			FALSE,					// Extract without pathname (junk paths)?
			FALSE,					// Extract with pathname?
			NULL,					// Password may be NULL.
			NULL,					// Location to which to unzip files.
			pUnzippedFiles,			// Output buffer to hold unzipped files.
			plUnzippedFilesMaxSize,	// Maximum size of output buffer.
			plUnzippedFilesCount,	// Number of unzipped files.
			error,					// Receives any error.
			maxlenError);

	if (result != 0)
	{
		SetLastErrorText(error);
	}

	return result;
}
