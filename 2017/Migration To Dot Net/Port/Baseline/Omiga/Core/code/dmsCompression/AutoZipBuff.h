///////////////////////////////////////////////////////////////////////////////
//	FILE:			AutoZipBuff.h
//	DESCRIPTION: 	Interface for smart buffers, which can zip and crc.
//	SYSTEM:	    	EServer
//	COPYRIGHT:		(c) 1998, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		10/12/99	First version
////////////////////////////////////////////////////////////////////////////////

#ifndef AUTOZIPBUFF_H
#define AUTOZIPBUFF_H

#include "ZipDef.h"
#if defined(DAL)
#include "dal.h"
#endif
#include "autobuff.h"
#include "CompAPIZipper.h"

struct ZIPDATA;

class COMPAPI CAutoZipBuffer : public CAutoBuffer
{
protected:
	BOOL				m_bZipped;
	CCompAPIZipper*			m_pZipper;
	BOOL				m_bOwnsZipper;
	unsigned long		m_ulZippedSize;
	unsigned long		m_ulUnzippedSize;
	unsigned short		m_uCrc;

#ifdef ESERVER
	Session*			m_pSession;
#endif

	enum { m_nZipThreshold = 16 };

public:
	CAutoZipBuffer(char* pData = 0, unsigned long ulSize = 0, BOOL bZipped = FALSE, BOOL bOwnsData = FALSE);
	CAutoZipBuffer(unsigned long ulSize, BOOL bZipped = FALSE);
	CAutoZipBuffer(const CAutoZipBuffer& Buffer);
	CAutoZipBuffer& operator=(const CAutoZipBuffer& Buffer);
	virtual ~CAutoZipBuffer();
	virtual char* Release() const;
	CCompAPIZipper* GetZipper() const { return m_pZipper; }
	CCompAPIZipper* ReleaseZipper() const; 
	BOOL SetData(char* pData, unsigned long ulSize, BOOL bOwnsData = FALSE) 
	{
		return SetZippedData(pData, ulSize, m_bZipped, bOwnsData);
	}
	BOOL SetZippedData(unsigned long ulSize, BOOL bZipped = FALSE);
	BOOL SetZippedData(char* pData, unsigned long ulSize, BOOL bZipped = FALSE, BOOL bOwnsData = FALSE);
	void SetZippedSize(unsigned long ulZippedSize) { m_ulZippedSize = ulZippedSize; }
	void SetUnzippedSize(unsigned long ulUnzippedSize) { m_ulUnzippedSize = ulUnzippedSize; }
	void SetZipped(BOOL bZipped) { m_bZipped = bZipped; }
	unsigned long GetZippedSize() const { return m_ulZippedSize; }
	unsigned long GetUnzippedSize() const { return m_ulUnzippedSize; }
	unsigned long GetSize() const { return m_bZipped ? m_ulZippedSize : m_ulUnzippedSize; }
	unsigned short GetCrc() const { return m_uCrc; }
	unsigned short UpdateCrc();
	BOOL GetZipped() const { return m_bZipped; }
	BOOL Zip();
#if defined(DALZIP)
	BOOL Zip(HPool hSrcPool);
#endif
	BOOL Unzip();
#if defined(DALZIP)
	BOOL Unzip(HPool hTgtPool);
#endif
	virtual BOOL Grow(unsigned long ulSize);
	virtual ZIPDATA* GetZipData();

protected:
	virtual BOOL AllocData(unsigned long ulSize) { return AllocData(ulSize, m_bZipped); }
	BOOL AllocData(unsigned long ulSize, BOOL bZipped = FALSE);
	virtual void Free();
	virtual BOOL AllocZipper() = 0;
	void FreeZipper();
	BOOL SetZipper(CCompAPIZipper* pZipper);
	void CopyObject(const CAutoZipBuffer& Buffer, BOOL bRelease = TRUE);
};

#endif
