///////////////////////////////////////////////////////////////////////////////
//	FILE:			CompAPIZipper.h
//	DESCRIPTION: 	Base compression class.
//	COPYRIGHT:		(c) 2004, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		12/05/04	First version
///////////////////////////////////////////////////////////////////////////////

#ifndef ZIPPER_H
#define ZIPPER_H

#include <stdio.h>
#include "ZipDef.h"
#include "AutoBuff.h"

// warning C4996: 'function' was declared deprecated
#pragma warning(disable: 4996)

class CZipper;

class CReader
{
public:
	virtual unsigned long Read(CZipper* pZipper, char* pBuffer, unsigned long ulBytesToRead) = 0;
};

class CWriter
{
public:
	virtual unsigned long Write(CZipper* pZipper, char* pBuffer, unsigned long ulBytesToWrite) = 0;
};

class COMPAPI CBufferReader : public CReader
{
public:
	virtual unsigned long Read(CZipper* pZipper, char* pBuffer, unsigned long ulBytesToRead);
};

class COMPAPI CBufferWriter : public CWriter
{
public:
	virtual unsigned long Write(CZipper* pZipper, char* pBuffer, unsigned long ulBytesToWrite);
};

class COMPAPI CFileReader : public CReader
{
public:
	virtual unsigned long Read(CZipper* pZipper, char* pBuffer, unsigned long ulBytesToRead);
};

class COMPAPI CFileWriter : public CWriter
{
public:
	virtual unsigned long Write(CZipper* pZipper, char* pBuffer, unsigned long ulBytesToWrite);
};

class COMPAPI CZipper
{
friend class CBufferReader;
friend class CBufferWriter;
friend class CFileReader;
friend class CFileWriter;

	enum EZipperType
	{
		typeCompAPI,
		typeZLib
	};

protected:
	char*					m_pSrcBuffer;
	char*					m_pTgtBuffer;
	HANDLE					m_fpSrc;
	HANDLE					m_fpTgt;
	unsigned long			m_ulSrcLen;
	unsigned long			m_ulTgtLen;
	unsigned long			m_ulSrcMaxLen;
	unsigned long			m_ulTgtMaxLen;

	CAutoBuffer*			m_pAutoBuffer;
	unsigned short			m_nCrc;
	CReader*				m_pReader;
	CWriter*				m_pWriter;

	CBufferReader			m_BufferReader;
	CBufferWriter			m_BufferWriter;
	CFileReader				m_FileReader;
	CFileWriter				m_FileWriter;

	int						m_nLastError;
	enum { MAXLEN_ERROR = 1024 };
	wchar_t					m_sLastErrorText[MAXLEN_ERROR];

	enum EOp
	{
		opCompress,
		opDecompress
	};

public:
	CZipper();
	virtual ~CZipper();
	static CZipper* CreateZipper(EZipperType ZipperType);
	static CZipper* CreateZipper(LPCWSTR szZipperType);

	unsigned long GetTgtLen() const { return m_ulTgtLen; }
	unsigned long GetSrcLen() const { return m_ulSrcLen; }

	virtual BOOL ZipBufferToBuffer(
		char* pTgtBuffer, 
		unsigned long ulTgtMaxLen, 
		char* pSrcBuffer, 
		unsigned ulSrcMaxLen,
		unsigned short* pnCrc = NULL);
	virtual BOOL ZipBufferToBuffer(
		CAutoBuffer* pTgtAutoBuffer, 
		char* pSrcBuffer, 
		unsigned ulSrcMaxLen,
		unsigned short* pnCrc = NULL);

	virtual BOOL UnzipBufferToBuffer(
		CAutoBuffer* pTgtAutoBuffer, 
		char* pSrcBuffer, 
		unsigned ulSrcMaxLen,
		unsigned short* pnCrc = NULL);
	virtual BOOL UnzipBufferToBuffer(
		char* pTgtBuffer, 
		unsigned long ulTgtMaxLen, 
		char* pSrcBuffer, 
		unsigned ulSrcMaxLen,
		unsigned short* pnCrc = NULL);

	virtual BOOL ZipFileToFile(
		const wchar_t* pszTgtFileName, 
		const wchar_t* pszSrcFileName, 
		unsigned short* pnCrc = NULL);

	virtual BOOL UnzipFileToFile(
		const wchar_t* pszTgtFileName, 
		const wchar_t* pszSrcFileName, 
		unsigned short* pnCrc = NULL);

	virtual BOOL ZipFileToBuffer(
		char* pTgtBuffer, 
		unsigned long ulTgtMaxLen, 
		const wchar_t* pszSrcFileName, 
		unsigned short* pnCrc = NULL);
	virtual BOOL ZipFileToBuffer(
		CAutoBuffer* pTgtAutoBuffer, 
		const wchar_t* pszSrcFileName, 
		unsigned short* pnCrc = NULL);

	virtual BOOL UnzipFileToBuffer(
		char* pTgtBuffer, 
		unsigned long ulTgtMaxLen, 
		const wchar_t* pszSrcFileName, 
		unsigned short* pnCrc = NULL);
	virtual BOOL UnzipFileToBuffer(
		CAutoBuffer* pTgtAutoBuffer, 
		const wchar_t* pszSrcFileName, 
		unsigned short* pnCrc = NULL);

	virtual BOOL ZipBufferToFile(
		const wchar_t* pszTgtFileName, 
		char* pSrcBuffer, 
		unsigned long ulSrcMaxLen, 
		unsigned short* pnCrc = NULL);
	virtual BOOL ZipBufferToFile(
		const wchar_t* pszTgtFileName, 
		CAutoBuffer* pSrcAutoBuffer, 
		unsigned short* pnCrc = NULL);

	virtual BOOL UnzipBufferToFile(
		const wchar_t* pszTgtFileName, 
		char* pSrcBuffer, 
		unsigned long ulSrcMaxLen, 
		unsigned short* pnCrc = NULL);
	virtual BOOL UnzipBufferToFile(
		const wchar_t* pszTgtFileName, 
		CAutoBuffer* pSrcAutoBuffer, 
		unsigned short* pnCrc = NULL);

	virtual BOOL ZipFile(
		const wchar_t* pszTgtFileName, 
		const wchar_t* pszSrcFileName,
		BOOL noPaths);

	virtual BOOL UnzipFile(
		const wchar_t* pszTgtFileName, 
		const wchar_t* pszSrcFileName,
		const wchar_t* pszUnzipDir);

	unsigned long Read(char* pBuffer, unsigned long ulBytesToRead);
	unsigned long Write(char* pBuffer, unsigned long ulBytesToWrite);
	virtual int GetChar();
	virtual void SetLastError(int nLastError) { m_nLastError = nLastError; }
	virtual int GetLastError() const { return m_nLastError; }
	virtual void SetLastErrorText(const wchar_t* pszErrorText) { wcsncpy(m_sLastErrorText, pszErrorText, MAXLEN_ERROR); }
	virtual const wchar_t* GetLastErrorText() const { return m_sLastErrorText; }

protected:
	static EZipperType ResolveZipperType(LPCWSTR szZipperType);

	virtual BOOL BufferToBuffer(
		EOp Op,
		char* pTgtBuffer, 
		unsigned long ulTgtMaxLen, 
		CAutoBuffer* pTgtAutoBuffer,
		char* pSrcBuffer, 
		unsigned ulSrcMaxLen,
		unsigned short* pnCrc = NULL);
	virtual BOOL FileToBuffer(
		EOp Op,
		char* pTgtBuffer, 
		unsigned long ulTgtMaxLen, 
		CAutoBuffer* pTgtAutoBuffer, 
		const wchar_t* pszSrcFileName, 
		unsigned short* pnCrc = NULL);
	virtual BOOL BufferToFile(
		EOp Op,
		const wchar_t* pszTgtFileName, 
		char* pSrcBuffer, 
		unsigned long ulSrcMaxLen, 
		unsigned short* pnCrc = NULL);
	virtual BOOL FileToFile(
		EOp Op,
		const wchar_t* pszTgtFileName, 
		const wchar_t* pszSrcFileName, 
		unsigned short* pnCrc = NULL);
	virtual BOOL Compress() = 0;
	virtual BOOL Decompress() = 0;
	virtual CReader* GetBufferReader() { return &m_BufferReader; }
	virtual CWriter* GetBufferWriter() { return &m_BufferWriter; }
	virtual CReader* GetFileReader() { return &m_FileReader; }
	virtual CWriter* GetFileWriter() { return &m_FileWriter; }
};


#endif
