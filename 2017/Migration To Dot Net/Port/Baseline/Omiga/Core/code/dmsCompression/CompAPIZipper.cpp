///////////////////////////////////////////////////////////////////////////////
//	FILE:			Zipper.cpp
//	DESCRIPTION: 	New Dac compression class.
//	SYSTEM:	    	Data Access Layer
//	COPYRIGHT:		(c) 2000, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		04/01/00	First version
//	LD		23/04/02	Modified CCompAPIZipper::BufferToBuffer to accept
//						either a target buffer or a target auto buffer
//////////////////////////////////////////////////////////////////////

#include "stdhdr.h"
#include <stdio.h>
#include <errno.h>
#include "CompAPIZipper.h"
#include "CompApi.h"


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCompAPIZipper::CCompAPIZipper()
{
	offset				= 0;
	in_count			= 0;
	bytes_out			= 0;
	prefxcode			= 0;
	nextfree			= 0;
	highcode			= 0;
	maxcode				= 0;
	hashsize			= 0;
	bits				= 0;
	sfx					= NULLPTR(char);
#if (SPLIT_PFX)
	pfx[0]				= NULLPTR(CODE);
	pfx[1]				= NULLPTR(CODE);
#else
	pfx					= NULLPTR(CODE);
#endif
#if (SPLIT_HT)
	ht[0]				= NULLPTR(CODE);
	ht[1]				= NULLPTR(CODE);
#else
	ht					= NULLPTR(CODE);
#endif
	oldmaxcode			= 0;
	oldhashsize			= 0;
	oldbits				= 0;
	memset(&outbuf, 0, sizeof(outbuf));
	prevbits			= 0;
	size				= 0;
	memset(&inbuf, 0, sizeof(MAXBITS));
	token				= 0;
	maxtoklen			= MAXTOKLEN;
	maxbits				= DFLTBITS;
	exit_stat			= 0;
	block_compress		= BLOCK_MASK;

	memset(&m_ZipData, 0, sizeof(m_ZipData));
}

CCompAPIZipper::~CCompAPIZipper()
{
}

///////////////////////////////////////////////////////////////////////////////
//	Compress a data pool into a buffer.
///////////////////////////////////////////////////////////////////////////////
#if defined(DALZIP)
BOOL CCompAPIZipper::ZipPoolToBuffer(char* pTgtBuffer, unsigned long ulTgtMaxLen, HPool hSrcPool)
{
	return pTgtBuffer ? ZipPoolToBuffer(pTgtBuffer, ulTgtMaxLen, NULL, hSrcPool) : 0;
}
#endif

///////////////////////////////////////////////////////////////////////////////
//	Compress a data pool into a auto buffer, which will grow as necessary
///////////////////////////////////////////////////////////////////////////////
#if defined(DALZIP)
BOOL CCompAPIZipper::ZipPoolToBuffer(CAutoBuffer* pTgtAutoBuffer, HPool hSrcPool)
{
	return pTgtAutoBuffer ? ZipPoolToBuffer(NULL, 0, pTgtAutoBuffer, hSrcPool) : 0;
}
#endif

#if defined(DALZIP)
BOOL CCompAPIZipper::ZipPoolToBuffer(
	char* pTgtBuffer, 
	unsigned long ulTgtMaxLen, 
	CAutoBuffer* pTgtAutoBuffer, 
	HPool hSrcPool)
{
	BOOL bZipped = FALSE;

	m_ulTgtLen = 0;
	if (InitFlatRead(hSrcPool))
	{
		m_nCrc			= 0;

		m_ulSrcLen		= 0;
		m_pSrcBuffer	= NULL;
		m_ulSrcMaxLen	= 0;

		m_ulTgtLen		= 0;
		m_pTgtBuffer	= pTgtBuffer;
		m_ulTgtMaxLen	= ulTgtMaxLen;

		m_pAutoBuffer	= pTgtAutoBuffer;

		m_pReader		= GetPoolReader();
		m_pWriter		= GetBufferWriter();

		bZipped			= Compress();
	}

	if (!bZipped)
	{
		m_ulTgtLen = 0;
	}
	
	return bZipped;
}               
#endif

///////////////////////////////////////////////////////////////////////////////
//	Decompress an auto buffer into a data pool.
///////////////////////////////////////////////////////////////////////////////
#if defined(DALZIP)
BOOL CCompAPIZipper::UnzipBufferToPool(HPool hTgtPool, CAutoBuffer* pSrcAutoBuffer)
{
	return pSrcAutoBuffer ? UnzipBufferToPool(hTgtPool, pSrcAutoBuffer->GetData(), pSrcAutoBuffer->GetDataSize()) : 0;
}
#endif

///////////////////////////////////////////////////////////////////////////////
//	Decompress a buffer into a data pool.
///////////////////////////////////////////////////////////////////////////////
#if defined(DALZIP)
BOOL CCompAPIZipper::UnzipBufferToPool(HPool hTgtPool, char* pSrcBuffer, unsigned long ulSrcMaxLen)
{
	BOOL bUnzipped = FALSE;

	m_ulSrcLen = 0;
	if (InitFlatWrite(hTgtPool))
	{
		m_nCrc			= 0;

		m_ulSrcLen		= 0;
		m_pSrcBuffer	= pSrcBuffer;
		m_ulSrcMaxLen	= ulSrcMaxLen;

		m_ulTgtLen		= 0;
		m_pTgtBuffer	= NULL;
		m_ulTgtMaxLen	= 0;

		m_pReader		= GetBufferReader();
		m_pWriter		= GetPoolWriter();

		bUnzipped		= Decompress();
	}

	if (!bUnzipped)
	{
		m_ulSrcLen = 0;
	}

	return bUnzipped;
}               
#endif


///////////////////////////////////////////////////////////////////////////////
//	Wrapper for third party compression routine.
///////////////////////////////////////////////////////////////////////////////
BOOL CCompAPIZipper::Compress()
{
	::compress(this);
	return GetLastError() == 0;
}

///////////////////////////////////////////////////////////////////////////////
//	Wrapper for third party decompression routine.
///////////////////////////////////////////////////////////////////////////////
BOOL CCompAPIZipper::Decompress()
{
	::decompress(this);
	return GetLastError() == 0;
}


#if defined(DALZIP)
// AS 06/05/04 Pool code should be extracted into sub class.

//////////////////////////////////////////////////////////////////////
//	Read from a data pool into a buffer.
//////////////////////////////////////////////////////////////////////
unsigned long CPoolReader::Read(CCompAPIZipper* pZipper, char* pBuffer, unsigned long ulBytesToRead)
{
	unsigned long ulBytesRead = 0;

	if (pZipper)
	{
		ulBytesRead = pZipper->FlatRead(pBuffer, ulBytesToRead);
	}

	return ulBytesRead;
}

//////////////////////////////////////////////////////////////////////
//	Write to a data pool from a buffer.
//////////////////////////////////////////////////////////////////////
unsigned long CPoolWriter::Write(CCompAPIZipper* pZipper, char* pBuffer, unsigned long ulBytesToWrite)
{
	unsigned long ulBytesWritten = 0;

	if (pZipper)
	{
		ulBytesWritten = pZipper->FlatWrite(pBuffer, ulBytesToWrite);
	}

	return ulBytesWritten;
}
#endif

