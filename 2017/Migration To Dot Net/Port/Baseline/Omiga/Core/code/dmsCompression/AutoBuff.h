///////////////////////////////////////////////////////////////////////////////
//	FILE:			AUTOBUFF.H
//	DESCRIPTION: 	Interface for smart buffers.
//	SYSTEM:	    	EServer
//	COPYRIGHT:		(c) 1999, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		10/12/99	First version
////////////////////////////////////////////////////////////////////////////////

#ifndef AUTOBUFF_H
#define AUTOBUFF_H

#include "ZipDef.h"

class COMPAPI CAutoBuffer
{
protected:
	char*					m_pData;			// Start of data buffer.
	BOOL					m_bOwnsData;		// If TRUE, delete buffer memory on destruction.
	unsigned long			m_ulSize;			// Size of buffer.
	unsigned long			m_ulDataSize;		// Size of data contained in buffer.
	unsigned long			m_ulGrowBy;			// How much to grow the buffer (in bytes).
	static unsigned long	s_ulDefaultGrowBy;
	enum { eDefaultGrowBy = 1024 };				// Buffer grows 1K at a time by default.

	unsigned long		m_ulMaxSize;

public:
	enum { eMaxSizeUnlimited = -1 };

public:
	CAutoBuffer(char* pData = 0, unsigned long ulSize = 0, BOOL bOwnsData = FALSE);
	CAutoBuffer(unsigned long ulSize);
	CAutoBuffer(const CAutoBuffer& Buffer);
	CAutoBuffer& operator=(const CAutoBuffer& Buffer);
	virtual ~CAutoBuffer();
#if defined(_WIN32)
	// won't compile 16 bit
	char& operator*() const { return *GetData(); }
#pragma warning(disable:4284)
	// disable warning:
	// return type for 'CAutoBuffer::operator ->' is not a UDT or reference to a UDT.  
	// Will produce errors if applied using infix notation
	char* operator->() const { return GetData(); }
#pragma warning(default:4284)
#endif
	char* GetData() const { return m_pData; }
	BOOL SetData(unsigned long ulSize);
	BOOL SetData(char* pData, unsigned long ulSize, BOOL bOwnsData = FALSE);
	virtual char* Release() const; 
	virtual BOOL Grow(unsigned long ulSize);
	BOOL Prefix(const char* pszData);
	BOOL Prefix(const char* pData, unsigned long ulSize);
	BOOL Append(const char* pszData);
	BOOL Append(const char* pData, unsigned long ulSize);
	BOOL AppendV(const char* pszFormat, ...);
	BOOL AppendV(const char* pszFormat, va_list argList);
	unsigned long GetSize() const { return m_ulSize; }
	unsigned long GetDataSize() const { return min(m_ulSize, m_ulDataSize); }
	void SetDataSize(unsigned long ulDataSize) { m_ulDataSize = ulDataSize; }
	static unsigned long GetDefaultGrowBy() { return s_ulDefaultGrowBy; }
	static void SetDefaultGrowBy(unsigned long ulDefaultGrowBy) { s_ulDefaultGrowBy = ulDefaultGrowBy; }
	unsigned long GetGrowBy() const { return m_ulGrowBy; }
	void SetGrowBy(unsigned long ulGrowBy) { m_ulGrowBy = ulGrowBy; }
	virtual BOOL IsValid() const;
	unsigned long GetMaxSize() const { return m_ulMaxSize; }
	void SetMaxSize(unsigned long ulMaxSize) { m_ulMaxSize = ulMaxSize; }

protected:
	virtual BOOL AllocData(unsigned long ulSize);
	virtual void Free();
	void CopyObject(const CAutoBuffer& Buffer, BOOL bRelease = TRUE);
};

#endif
