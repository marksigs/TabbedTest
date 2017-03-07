///////////////////////////////////////////////////////////////////////////////
//	FILE:			AUTOBUFF.CPP
//	DESCRIPTION: 	
//
//		This class implements a smart pointer (i.e., automatically deletes any 
//		memory it allocates). It will work with both MSVC 16 and 32 bit.
//
//		Most of the smart pointer code is based on the C++ STL auto_ptr class.
//
//	SYSTEM:	    	DAL
//	COPYRIGHT:		(c) 1999, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		10/12/99	First version
//  LD      14/05/02	Grow in steps (m_ulGrowBy) rather than in increments of one
//						byte after the first allocation
////////////////////////////////////////////////////////////////////////////////

#include "stdhdr.h"
#include <stdio.h>
#include "autobuff.h"

unsigned long CAutoBuffer::s_ulDefaultGrowBy = eDefaultGrowBy;

CAutoBuffer::CAutoBuffer(char* pData, unsigned long ulSize, BOOL bOwnsData) :
	m_pData(pData),
	m_bOwnsData(pData != 0 && bOwnsData),
	m_ulSize(ulSize),
	m_ulDataSize(ulSize),
	m_ulGrowBy(s_ulDefaultGrowBy),
#ifdef _WIN32
	// Maximum size for 32 bit buffer is unlimited.
	m_ulMaxSize(eMaxSizeUnlimited)
#else
	// Maximum size for 16 bit buffer = 64K.
	m_ulMaxSize(0xFFFF)
#endif
{
}

CAutoBuffer::CAutoBuffer(unsigned long ulSize) :
	m_pData(0),
	m_bOwnsData(FALSE),
	m_ulSize(ulSize),
	m_ulDataSize(0),
	m_ulGrowBy(eDefaultGrowBy),
#ifdef _WIN32
	// Maximum size for 32 bit buffer is unlimited.
	m_ulMaxSize(eMaxSizeUnlimited)
#else
	// Maximum size for 16 bit buffer = 64K.
	m_ulMaxSize(0xFFFF)
#endif
{
	AllocData(ulSize);
}

CAutoBuffer::CAutoBuffer(const CAutoBuffer& Buffer)
{
	CopyObject(Buffer);
}

CAutoBuffer& CAutoBuffer::operator=(const CAutoBuffer& Buffer) 
{
	if (m_pData != Buffer.m_pData)
	{
		Free();
		CopyObject(Buffer);
	}
	else 
	{
		if (Buffer.m_bOwnsData)
		{
			m_bOwnsData = TRUE;
		}
	}

	return *this; 
}

void CAutoBuffer::CopyObject(const CAutoBuffer& Buffer, BOOL bRelease)
{
	if (bRelease)
	{
		// This object owns the data.
		m_bOwnsData			= Buffer.m_bOwnsData;
		m_pData				= Buffer.Release();
	}
	else
	{
		// This object does not own the data, so take a copy.
		if (SetData(Buffer.GetSize()))
		{
			memcpy(m_pData, Buffer.m_pData, Buffer.m_ulSize);
		}
	}
	m_ulDataSize		= Buffer.m_ulDataSize;
	m_ulSize			= Buffer.m_ulSize;
}

CAutoBuffer::~CAutoBuffer() 
{ 
	Free();
}

BOOL CAutoBuffer::AllocData(unsigned long ulSize)
{
	BOOL bSuccess = FALSE;

	if (m_ulMaxSize == eMaxSizeUnlimited || ulSize < m_ulMaxSize)
	{
		char* pData = new char[ulSize];

		_ASSERTE(pData != 0);

		if (pData != 0)
		{
			memset(pData, 0, (size_t)ulSize);
		}

		bSuccess		= SetData(pData, ulSize, TRUE);
		m_ulDataSize	= 0;	// Always 0 when allocating.
	}

	return bSuccess;
}

BOOL CAutoBuffer::SetData(unsigned long ulSize)
{
	return AllocData(ulSize);
}

BOOL CAutoBuffer::SetData(char* pData, unsigned long ulSize, BOOL bOwnsData) 
{ 
	if (m_pData != pData)
	{
		Free();
	}

	m_pData				= pData;
	m_ulSize			= ulSize;
	m_ulDataSize		= m_ulSize;
	m_bOwnsData			= pData != 0 && bOwnsData;

	return m_bOwnsData;
}

void CAutoBuffer::Free()
{
	if (m_bOwnsData) 
	{
		delete []m_pData; 
		m_pData				= 0;
		m_ulDataSize		= 0;
	}
	Release();
}

char* CAutoBuffer::Release() const
{
	((CAutoBuffer*)this)->m_bOwnsData = FALSE;

	return m_pData; 
}

BOOL CAutoBuffer::Grow(unsigned long ulSize)
{
	BOOL bSuccess = TRUE;

	// Can't grow a buffer we don't own, or buffer does not already exist.
	// Don't grow buffer if it is already big enough.
	if ((m_bOwnsData || m_pData == 0) && m_ulSize < ulSize)
	{
		CAutoBuffer AutoBuffer(max(m_ulSize + m_ulGrowBy, ulSize)); // always grow in steps

		if (AutoBuffer.m_pData != 0)
		{
			if (m_pData != 0 && m_ulSize > 0)
			{
				memcpy(AutoBuffer.m_pData, m_pData, m_ulSize);
			}
			AutoBuffer.m_ulDataSize		= m_ulDataSize;
			*this						= AutoBuffer;
		}
		else
		{
			bSuccess = FALSE;
		}
	}

	return bSuccess;
}

BOOL CAutoBuffer::Prefix(const char* pszData)
{
	return Prefix(pszData, strlen(pszData) + 1);
}

BOOL CAutoBuffer::Prefix(const char* pData, unsigned long ulSize)
{
	BOOL bSuccess = FALSE;

	if (Grow(m_ulDataSize + ulSize))
	{
		memmove(m_pData + ulSize, m_pData, m_ulDataSize);
		memcpy(m_pData, pData, ulSize);
		m_ulDataSize += ulSize;
		bSuccess = TRUE;
	}

	return bSuccess;
}

BOOL CAutoBuffer::Append(const char* pszData)
{
	return Append(pszData, strlen(pszData) + 1);
}

BOOL CAutoBuffer::Append(const char* pData, unsigned long ulSize)
{
	BOOL bSuccess = FALSE;

	if (Grow(m_ulDataSize + ulSize))
	{
		memcpy(m_pData + m_ulDataSize, pData, ulSize);
		m_ulDataSize += ulSize;
		bSuccess = TRUE;
	}

	return bSuccess;
}

BOOL CAutoBuffer::AppendV(const char* pszFormat, ...)
{
	BOOL bSuccess = FALSE;

	va_list argList;
	va_start(argList, pszFormat);

	bSuccess = AppendV(pszFormat, argList);

	va_end(argList);

	return bSuccess;
}

BOOL CAutoBuffer::AppendV(const char* pszFormat, va_list argList)
{
	BOOL bSuccess = FALSE;

	// Grow the buffer so it is at least big enough to contain the format string.
	bSuccess = Grow(m_ulDataSize + strlen(pszFormat) + 1);

	if (bSuccess)
	{
		int nWritten = 0;
		do 
		{
			if (bSuccess)
			{
				nWritten = 
					_vsnprintf(
						m_pData + m_ulDataSize, 
						m_ulSize - m_ulDataSize, 
						pszFormat, 
						argList);

				if (nWritten >= 0)
				{
					// Successfully written to buffer.
					m_ulDataSize += nWritten + 1;	// Extra byte for terminating null character.
				}
				else
				{
					// Buffer too small, thus increase its size.
					bSuccess = Grow(m_ulSize + m_ulGrowBy);
				}
			}
		}
		while (bSuccess && nWritten < 0);
	}

	return bSuccess;
}

BOOL CAutoBuffer::IsValid() const
{
	BOOL bValid = FALSE;
	
	if (m_bOwnsData)
	{
		bValid = m_pData != NULL;
		
#ifdef ESERVER
		if (bValid)
		{
			bValid = MemChecker.IsValidPtr(m_pData, sizeof(char) * m_ulSize, __FILE__, __LINE__);
		}
#endif
	}
	else
	{
		bValid = TRUE;
	}

	return bValid;
}
