///////////////////////////////////////////////////////////////////////////////
//	FILE:			CompAPIZipper.h
//	DESCRIPTION: 	New Dac compression class.
//	SYSTEM:	    	Data Access Layer
//	COPYRIGHT:		(c) 2000, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		04/01/00	First version
///////////////////////////////////////////////////////////////////////////////

#ifndef COMPAPIZIPPER_H
#define COMPAPIZIPPER_H

#include "Zipper.h"
#if defined(DALZIP)
#include "dal.h"
#endif
#include "comp.h"

class CCompAPIZipper;

class COMPAPI CPoolReader : public CReader
{
public:
	virtual unsigned long Read(CCompAPIZipper* pZipper, char* pBuffer, unsigned long ulBytesToRead);
};

class COMPAPI CPoolWriter : public CWriter
{
public:
	virtual unsigned long Write(CCompAPIZipper* pZipper, char* pBuffer, unsigned long ulBytesToWrite);
};

struct ZIPDATA
{
};

class COMPAPI CCompAPIZipper : public CZipper
{
friend class CPoolReader;
friend class CPoolWriter;

public:
	// Third party compression code attributes.
	int						offset;
	long int				in_count;		// length of input
	long int				bytes_out;		// length of compressed output
	INTCODE					prefxcode;
	INTCODE					nextfree;
	INTCODE					highcode;
	INTCODE					maxcode;
	HASH					hashsize;
	int						bits;
	char FAR*				sfx;			// = NULLPTR(char);
#if (SPLIT_PFX)
	CODE FAR*				pfx[2];			// = NULLPTR(CODE);
#else
	CODE FAR*				pfx;			// = NULLPTR(CODE);
#endif
#if (SPLIT_HT)
	CODE FAR*				ht[2];			// = {NULLPTR(CODE),NULLPTR(CODE)};
#else
	CODE FAR*				ht;				// = NULLPTR(CODE);
#endif
	INTCODE					oldmaxcode;
	HASH					oldhashsize;
	int						oldbits;
	UCHAR					outbuf[MAXBITS];
	int						prevbits;
	int						size;
	UCHAR					inbuf[MAXBITS];
	char*					token;			// String buffer to build token
	int						maxtoklen;		// = MAXTOKLEN
	int						maxbits;
	int						exit_stat;
	int						block_compress;

protected:
#if defined(DALZIP)
	// AS 06/05/04 Pool code should be extracted into sub class.
	CPoolReader				m_PoolReader;
	CPoolWriter				m_PoolWriter;
#endif

	ZIPDATA					m_ZipData;

public:
	CCompAPIZipper();
	virtual ~CCompAPIZipper();

	virtual ZIPDATA* GetZipData() { return &m_ZipData; }

#if defined(DALZIP)
	virtual BOOL ZipPoolToBuffer(char* pTgtBuffer, unsigned long ulTgtMaxLen, HPool hSrcPool);
	virtual BOOL ZipPoolToBuffer(CAutoBuffer* pTgtAutoBuffer, HPool hSrcPool);
	virtual BOOL UnzipBufferToPool(HPool hTgtPool, char* pSrcBuffer, unsigned long ulSrcMaxLen);
	virtual BOOL UnzipBufferToPool(HPool hTgtPool, CAutoBuffer* pSrcAutoBuffer);
#endif

protected:
	BOOL Compress();
	BOOL Decompress();

#if defined(DALZIP)
	virtual BOOL ZipPoolToBuffer(
		char* pTgtBuffer, 
		unsigned long ulTgtMaxLen, 
		CAutoBuffer* pTgtAutoBuffer, 
		HPool hSrcPool);
	virtual BOOL InitFlatRead(HPool hSrcPool) = 0;
	virtual BOOL InitFlatWrite(HPool hTgtPool) = 0;
	virtual unsigned long FlatRead(char* pBuffer, unsigned long ulBytesToRead) = 0;
	virtual unsigned long FlatWrite(char* pBuffer, unsigned long ulBytesToWrite) = 0;
	virtual CReader* GetPoolReader() { return &m_PoolReader; }
	virtual CWriter* GetPoolWriter() { return &m_PoolWriter; }
#endif
};

#endif 
