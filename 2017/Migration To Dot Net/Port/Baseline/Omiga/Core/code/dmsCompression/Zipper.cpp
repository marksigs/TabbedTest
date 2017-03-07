///////////////////////////////////////////////////////////////////////////////
//	FILE:			Zipper.cpp
//	DESCRIPTION: 	Base compression class.
//	COPYRIGHT:		(c) 2004, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		12/05/04	First version
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <errno.h>
#include "Zipper.h"
#include "CompAPIZipper.h"
#include "ZLibZipper.h"
#include "crc.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CZipper::CZipper() :
	m_pSrcBuffer(NULL),
	m_pTgtBuffer(NULL),
	m_fpSrc(INVALID_HANDLE_VALUE),
	m_fpTgt(INVALID_HANDLE_VALUE),
	m_ulSrcLen(0),
	m_ulTgtLen(0),
	m_ulSrcMaxLen(0),
	m_ulTgtMaxLen(0),
	m_pAutoBuffer(NULL),
	m_pReader(NULL),
	m_pWriter(NULL),
	m_nCrc(0),
	m_nLastError(0)
{
	ZeroMemory(m_sLastErrorText, MAXLEN_ERROR * sizeof(m_sLastErrorText[0]));
}

CZipper::~CZipper()
{
}


CZipper* CZipper::CreateZipper(EZipperType ZipperType)
{
	CZipper* pZipper = NULL;

	switch (ZipperType)
	{
	case typeZLib:
		pZipper = new CZLibZipper();
		break;
	default:
		pZipper = new CCompAPIZipper();
		break;
	}

	_ASSERTE(pZipper != NULL);

	return pZipper;
}

CZipper* CZipper::CreateZipper(LPCWSTR szZipperType)
{
	return CreateZipper(ResolveZipperType(szZipperType));
}

CZipper::EZipperType CZipper::ResolveZipperType(LPCWSTR szZipperType)
{
	EZipperType ZipperType = typeCompAPI;

	if (_wcsicmp(szZipperType, L"COMPAPI") == 0)
	{
		ZipperType = typeCompAPI;
	}
	else if (_wcsicmp(szZipperType, L"ZLIB") == 0)
	{
		ZipperType = typeZLib;
	}
	else
	{
		_ASSERTE(FALSE);
	}

	return ZipperType;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//	Compress a file into another file.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::ZipFileToFile(
	const wchar_t* pszTgtFileName, 
	const wchar_t* pszSrcFileName, 
	unsigned short* pnCrc)
{
	return FileToFile(opCompress, pszTgtFileName, pszSrcFileName, pnCrc);
}


///////////////////////////////////////////////////////////////////////////////
//	Decompress a file into another file.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::UnzipFileToFile(
	const wchar_t* pszTgtFileName, 
	const wchar_t* pszSrcFileName, 
	unsigned short* pnCrc)
{
	return FileToFile(opDecompress, pszTgtFileName, pszSrcFileName, pnCrc);
}


///////////////////////////////////////////////////////////////////////////////
//	Compress/decompress a file into another file.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::FileToFile(
	EOp Op,
	const wchar_t* pszTgtFileName, 
	const wchar_t* pszSrcFileName, 
	unsigned short* pnCrc)
{
	BOOL bSuccess = FALSE;

	m_ulTgtLen = 0;
	if (pszTgtFileName && wcslen(pszTgtFileName) > 0 && pszSrcFileName && wcslen(pszSrcFileName) > 0)
	{
		m_nCrc			= 0;
		m_ulSrcLen		= 0;
		m_pSrcBuffer	= NULL;
		m_ulTgtLen		= 0;
		m_pTgtBuffer	= NULL;
		m_pReader		= GetFileReader();
		m_pWriter		= GetFileWriter();

		if	(
				(m_fpTgt = ::CreateFile(pszTgtFileName, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, NULL, NULL)) != INVALID_HANDLE_VALUE &&
				(m_fpSrc = ::CreateFile(pszSrcFileName, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, NULL, NULL)) != INVALID_HANDLE_VALUE
			)
		{
			DWORD dwSrcFileSizeHigh = 0;
			DWORD dwSrcFileSize = ::GetFileSize(m_fpSrc, &dwSrcFileSizeHigh);
			if (dwSrcFileSize != INVALID_FILE_SIZE)
			{
				m_ulSrcMaxLen = dwSrcFileSize;

				if (Op == opCompress)
				{
					bSuccess = Compress();
				}
				else if (Op == opDecompress)
				{
					bSuccess = Decompress();
				}
			}
			else
			{
				SetLastError(EBADF);
			}

			if (m_fpSrc)
			{
				::CloseHandle(m_fpSrc);
				m_fpSrc = INVALID_HANDLE_VALUE;
			}

			if (m_fpTgt)
			{
				::CloseHandle(m_fpTgt);
				m_fpTgt = INVALID_HANDLE_VALUE;
			}

			if (pnCrc)
			{
				*pnCrc = m_nCrc;
			}
		}
		else
		{
			SetLastError(EBADF);
		}
	}

	if (!bSuccess)
	{
		m_ulTgtLen = 0;
	}

	return bSuccess;
}               


/////////////////////////////////////////
//	Compress a file into a buffer.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::ZipFileToBuffer(
	char* pTgtBuffer, 
	unsigned long ulTgtMaxLen, 
	const wchar_t* pszSrcFileName, 
	unsigned short* pnCrc)
{
	return pTgtBuffer ? FileToBuffer(opCompress, pTgtBuffer, ulTgtMaxLen, NULL, pszSrcFileName, pnCrc) : 0;
}

BOOL CZipper::ZipFileToBuffer(
	CAutoBuffer* pTgtAutoBuffer, 
	const wchar_t* pszSrcFileName, 
	unsigned short* pnCrc)
{
	return pTgtAutoBuffer ? FileToBuffer(opCompress, NULL, 0, pTgtAutoBuffer, pszSrcFileName, pnCrc) : 0;
}


///////////////////////////////////////////////////////////////////////////////
//	Decompress a file into a buffer.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::UnzipFileToBuffer(
	char* pTgtBuffer, 
	unsigned long ulTgtMaxLen, 
	const wchar_t* pszSrcFileName, 
	unsigned short* pnCrc)
{
	return pTgtBuffer ? FileToBuffer(opDecompress, pTgtBuffer, ulTgtMaxLen, NULL, pszSrcFileName, pnCrc) : 0;
}

BOOL CZipper::UnzipFileToBuffer(
	CAutoBuffer* pTgtAutoBuffer, 
	const wchar_t* pszSrcFileName, 
	unsigned short* pnCrc)
{
	return pTgtAutoBuffer ? FileToBuffer(opDecompress, NULL, 0, pTgtAutoBuffer, pszSrcFileName, pnCrc) : 0;
}


///////////////////////////////////////////////////////////////////////////////
//	Compress/Decompress a file into a buffer.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::FileToBuffer(
	EOp Op,
	char* pTgtBuffer, 
	unsigned long ulTgtMaxLen, 
	CAutoBuffer* pTgtAutoBuffer, 
	const wchar_t* pszSrcFileName, 
	unsigned short* pnCrc)
{
	BOOL bSuccess = FALSE;

	m_ulTgtLen = 0;
	if (pTgtBuffer && pszSrcFileName && wcslen(pszSrcFileName) > 0)
	{
		m_nCrc			= 0;
		m_ulSrcLen		= 0;
		m_pSrcBuffer	= NULL;
		m_ulTgtLen		= 0;
		m_pTgtBuffer	= pTgtBuffer;
		m_ulTgtMaxLen	= ulTgtMaxLen;
		m_pAutoBuffer	= pTgtAutoBuffer;
		m_pReader		= GetFileReader();
		m_pWriter		= GetBufferWriter();

		if ((m_fpSrc = ::CreateFile(pszSrcFileName, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, NULL, NULL)) != INVALID_HANDLE_VALUE)
		{
			DWORD dwSrcFileSizeHigh = 0;
			DWORD dwSrcFileSize = ::GetFileSize(m_fpSrc, &dwSrcFileSizeHigh);
			if (dwSrcFileSize != INVALID_FILE_SIZE)
			{
				m_ulSrcMaxLen = dwSrcFileSize;

				if (Op == opCompress)
				{
					bSuccess = Compress();
				}
				else if (Op == opDecompress)
				{
					bSuccess = Decompress();
				}
			}
			else
			{
				SetLastError(EBADF);
			}

			if (m_fpSrc)
			{
				::CloseHandle(m_fpSrc);
				m_fpSrc = INVALID_HANDLE_VALUE;
			}

			if (pnCrc)
			{
				*pnCrc = m_nCrc;
			}
		}
		else
		{
			SetLastError(EBADF);
		}
	}

	if (!bSuccess)
	{
		m_ulTgtLen = 0;
	}

	return bSuccess;
}               


///////////////////////////////////////////////////////////////////////////////
//	Compress a buffer into a file.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::ZipBufferToFile(
	const wchar_t* pszTgtFileName, 
	char* pSrcBuffer, 
	unsigned long ulSrcMaxLen, 
	unsigned short* pnCrc)
{
	return BufferToFile(opCompress, pszTgtFileName, pSrcBuffer, ulSrcMaxLen, pnCrc);
}

BOOL CZipper::ZipBufferToFile(
	const wchar_t* pszTgtFileName, 
	CAutoBuffer* pSrcAutoBuffer, 
	unsigned short* pnCrc)
{
	return pSrcAutoBuffer ? BufferToFile(opCompress, pszTgtFileName, pSrcAutoBuffer->GetData(), pSrcAutoBuffer->GetDataSize(), pnCrc) : 0;
}

///////////////////////////////////////////////////////////////////////////////
//	Decompress a buffer into a file.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::UnzipBufferToFile(
	const wchar_t* pszTgtFileName, 
	char* pSrcBuffer, 
	unsigned long ulSrcMaxLen, 
	unsigned short* pnCrc)
{
	return BufferToFile(opDecompress, pszTgtFileName, pSrcBuffer, ulSrcMaxLen, pnCrc);
}

BOOL CZipper::UnzipBufferToFile(
	const wchar_t* pszTgtFileName, 
	CAutoBuffer* pSrcAutoBuffer, 
	unsigned short* pnCrc)
{
	return pSrcAutoBuffer ? BufferToFile(opDecompress, pszTgtFileName, pSrcAutoBuffer->GetData(), pSrcAutoBuffer->GetDataSize(), pnCrc) : 0;
}

///////////////////////////////////////////////////////////////////////////////
//	Compress/Decompress a buffer into a file.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::BufferToFile(
	EOp Op,
	const wchar_t* pszTgtFileName, 
	char* pSrcBuffer, 
	unsigned long ulSrcMaxLen, 
	unsigned short* pnCrc)
{
	BOOL bSuccess = FALSE;

	m_ulTgtLen = 0;
	if (pszTgtFileName && wcslen(pszTgtFileName) > 0 && pSrcBuffer)
	{
		m_nCrc			= 0;

		m_ulSrcLen		= 0;
		m_pSrcBuffer	= pSrcBuffer;
		m_ulSrcMaxLen	= ulSrcMaxLen;

		m_ulTgtLen		= 0;
		m_pTgtBuffer	= NULL;
		m_ulTgtMaxLen	= 0;

		m_pAutoBuffer	= NULL;

		m_pReader		= GetBufferReader();
		m_pWriter		= GetFileWriter();

		if ((m_fpTgt = ::CreateFile(pszTgtFileName, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, NULL, NULL)) != INVALID_HANDLE_VALUE)
		{
			if (Op == opCompress)
			{
				bSuccess = Compress();
			}
			else if (Op == opDecompress)
			{
				bSuccess = Decompress();
			}
		}
		else
		{
			SetLastError(EBADF);
		}

		if (m_fpTgt)
		{
			::CloseHandle(m_fpTgt);
			m_fpTgt = INVALID_HANDLE_VALUE;
		}

		if (pnCrc)
		{
			*pnCrc = m_nCrc;
		}
	}

	if (!bSuccess)
	{
		m_ulTgtLen = 0;
	}

	return bSuccess;
}               


///////////////////////////////////////////////////////////////////////////////
//	Compress a buffer into another buffer.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::ZipBufferToBuffer(
	char* pTgtBuffer, 
	unsigned long ulTgtMaxLen, 
	char* pSrcBuffer, 
	unsigned ulSrcMaxLen,
	unsigned short* pnCrc)
{
	return BufferToBuffer(opCompress, pTgtBuffer, ulTgtMaxLen, NULL, pSrcBuffer, ulSrcMaxLen, pnCrc);
}

///////////////////////////////////////////////////////////////////////////////
//	Compress a buffer into an auto buffer that will grow as necessary.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::ZipBufferToBuffer(
	CAutoBuffer* pTgtAutoBuffer, 
	char* pSrcBuffer, 
	unsigned ulSrcMaxLen,
	unsigned short* pnCrc)
{
	return BufferToBuffer(opCompress, NULL, 0, pTgtAutoBuffer, pSrcBuffer, ulSrcMaxLen, pnCrc);
}

///////////////////////////////////////////////////////////////////////////////
//	Decompress a buffer into another buffer.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::UnzipBufferToBuffer(
	char* pTgtBuffer, 
	unsigned long ulTgtMaxLen, 
	char* pSrcBuffer, 
	unsigned ulSrcMaxLen,
	unsigned short* pnCrc)
{
	return BufferToBuffer(opDecompress, pTgtBuffer, ulTgtMaxLen, NULL, pSrcBuffer, ulSrcMaxLen, pnCrc);
}

///////////////////////////////////////////////////////////////////////////////
//	Decompress a buffer into an auto buffer that will grow as necessary.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::UnzipBufferToBuffer(
	CAutoBuffer* pTgtAutoBuffer, 
	char* pSrcBuffer, 
	unsigned ulSrcMaxLen,
	unsigned short* pnCrc)
{
	return BufferToBuffer(opDecompress, NULL, 0, pTgtAutoBuffer, pSrcBuffer, ulSrcMaxLen, pnCrc);
}

///////////////////////////////////////////////////////////////////////////////
//	Compress/decompress a buffer into another buffer.
///////////////////////////////////////////////////////////////////////////////
BOOL CZipper::BufferToBuffer(
	EOp Op,
	char* pTgtBuffer, 
	unsigned long ulTgtMaxLen, 
	CAutoBuffer* pTgtAutoBuffer,
	char* pSrcBuffer, 
	unsigned ulSrcMaxLen,
	unsigned short* pnCrc)
{
	BOOL bSuccess = FALSE;

	m_ulTgtLen = 0;

	if ((pTgtBuffer || pTgtAutoBuffer) && pSrcBuffer)
	{
		m_nCrc			= 0;

		m_ulSrcLen		= 0;
		m_pSrcBuffer	= pSrcBuffer;
		m_ulSrcMaxLen	= ulSrcMaxLen;

		m_ulTgtLen		= 0;
		m_pTgtBuffer	= pTgtBuffer;
		m_ulTgtMaxLen	= ulTgtMaxLen;

		m_pAutoBuffer	= pTgtAutoBuffer;

		m_pReader		= GetBufferReader();
		m_pWriter		= GetBufferWriter();

		if (Op == opCompress)
		{
			bSuccess = Compress();
		}
		else if (Op == opDecompress)
		{
			bSuccess = Decompress();
		}
	}

	if (pnCrc)
	{
		*pnCrc = m_nCrc;
	}

	if (!bSuccess)
	{
		m_ulTgtLen = 0;
	}

	return bSuccess;
}


///////////////////////////////////////////////////////////////////////////////
//	Generic read function that uses polymorphic reader.
///////////////////////////////////////////////////////////////////////////////
unsigned long CZipper::Read(char* pBuffer, unsigned long ulBytesToRead)
{
	unsigned long ulBytesRead = 0;

	if (m_pReader)
	{
		ulBytesRead = m_pReader->Read(this, pBuffer, ulBytesToRead);
	}

	return ulBytesRead;
}

////////////////////////////////////////////////////////////// /////////////////
//	Generic write function that uses polymorphic writer.
///////////////////////////////////////////////////////////////////////////////
unsigned long CZipper::Write(char* pBuffer, unsigned long ulBytesToWrite)
{
	unsigned long ulBytesWritten = 0;

	if (m_pWriter)
	{
		ulBytesWritten = m_pWriter->Write(this, pBuffer, ulBytesToWrite);
	}

	return ulBytesWritten;
}

///////////////////////////////////////////////////////////////////////////////
//	Read single byte from stream.
///////////////////////////////////////////////////////////////////////////////
int CZipper::GetChar()
{
	int ch = 0;
	
	if (Read((char*)&ch, 1) == 0)
	{
		ch = EOF;
	}

	return ch;
}


//////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////
//	Polymorphic Reader/Writer classes.
//////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
//	Read from a buffer.
//////////////////////////////////////////////////////////////////////
unsigned long CBufferReader::Read(CZipper* pZipper, char* pBuffer, unsigned long ulBytesToRead)
{
	unsigned long ulBytesRead = 0;

	if (pZipper && pBuffer)
	{
		while (ulBytesToRead && pZipper->m_ulSrcLen < pZipper->m_ulSrcMaxLen)
		{
			*pBuffer++ = *pZipper->m_pSrcBuffer++;
			ulBytesRead++;
			ulBytesToRead--;
			pZipper->m_ulSrcLen++;
		}
	}

	return ulBytesRead;
}


//////////////////////////////////////////////////////////////////////
//	Write to a buffer.
//////////////////////////////////////////////////////////////////////
unsigned long CBufferWriter::Write(CZipper* pZipper, char* pBuffer, unsigned long ulBytesToWrite)
{
	unsigned long ulBytesWritten = 0;

	::internal_crc((unsigned char*)pBuffer, &pZipper->m_nCrc, ulBytesToWrite);

	if (pZipper && pBuffer)
	{
		if (pZipper->m_pAutoBuffer)
		{
			// There is an auto buffer, so use this (buffer will grow automatically).
			if (pZipper->m_pAutoBuffer->Append(pBuffer, ulBytesToWrite))
			{
				ulBytesWritten		= ulBytesToWrite;
				pZipper->m_ulTgtLen += ulBytesWritten;
			}
			else
			{
				// Emulate file stream "Not enough memory" error.
				pZipper->SetLastError(ENOMEM);
			}
		}
		else
		{
			// Writing to a fixed sized buffer.
			while (ulBytesToWrite && pZipper->m_ulTgtLen < pZipper->m_ulTgtMaxLen)
			{
				*pZipper->m_pTgtBuffer++ = *pBuffer++;
				ulBytesWritten++;
				ulBytesToWrite--;
				pZipper->m_ulTgtLen++;
			}
		}
	}

	return ulBytesWritten;
}


//////////////////////////////////////////////////////////////////////
//	Read from a file into a buffer.
//////////////////////////////////////////////////////////////////////
unsigned long CFileReader::Read(CZipper* pZipper, char* pBuffer, unsigned long ulBytesToRead)
{
	unsigned long ulBytesRead = 0;

	if (pZipper && pZipper->m_fpSrc != INVALID_HANDLE_VALUE && pBuffer)
	{
		if (::ReadFile(pZipper->m_fpSrc, pBuffer, ulBytesToRead, &ulBytesRead, NULL))
		{
			pZipper->m_ulSrcLen += ulBytesRead;
		}
		else
		{
			pZipper->SetLastError(::GetLastError());
		}
	}

	return ulBytesRead;
}

//////////////////////////////////////////////////////////////////////
//	Write to a file from a buffer.
//////////////////////////////////////////////////////////////////////
unsigned long CFileWriter::Write(CZipper* pZipper, char* pBuffer, unsigned long ulBytesToWrite)
{
	unsigned long ulBytesWritten = 0;

	if (pZipper && pZipper->m_fpTgt != INVALID_HANDLE_VALUE && pBuffer)
	{
		if (::WriteFile(pZipper->m_fpTgt, pBuffer, ulBytesToWrite, &ulBytesWritten, NULL))
		{
			pZipper->m_ulTgtLen += ulBytesWritten;

			if (ulBytesWritten > 0)
			{
				::internal_crc((unsigned char*)pBuffer, &pZipper->m_nCrc, ulBytesToWrite);
			}
		}
		else
		{
			pZipper->SetLastError(::GetLastError());
		}
	}

	return ulBytesWritten;
}


BOOL CZipper::ZipFile(const wchar_t* pszTgtFileName, const wchar_t* pszSrcFileName, BOOL noPaths)
{
	SetLastError(ERROR_CALL_NOT_IMPLEMENTED);
	return FALSE;
}

BOOL CZipper::UnzipFile(const wchar_t* pszTgtFileName, const wchar_t* pszSrcFileName, const wchar_t* pszUnzipDir)
{
	SetLastError(ERROR_CALL_NOT_IMPLEMENTED);
	return FALSE;
}
