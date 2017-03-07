///////////////////////////////////////////////////////////////////////////////
//	FILE:			AutoZipBuff.cpp
//	DESCRIPTION: 	
//
//		This class implements a smart pointer (i.e., automatically deletes any 
//		memory it allocates). It will work with both MSVC 16 and 32 bit.  It 
//		also is able to zip and unzip its data, using the DAL compression
//		routines.  
//
//		The ESERVER compile time constant is necessary because the 
//		implementations of the EServer and the client DAL compression routines
//		are (unnecessarily) different.
//
//		Most of the smart pointer code is based on the C++ STL auto_ptr class.
//
//	SYSTEM:	    	DAL
//	COPYRIGHT:		(c) 1998, Marlborough Stirling. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//	AS		15/04/98	First version
////////////////////////////////////////////////////////////////////////////////

#include "stdhdr.h"
#if !defined(_WIN32)
#include <memory.h>
#endif
#include "autozipbuff.h"
#include "Crc.h"


CAutoZipBuffer::CAutoZipBuffer(char* pData, unsigned long ulSize, BOOL bZipped, BOOL bOwnsData) : 
	CAutoBuffer(pData, ulSize, bOwnsData),
	m_bZipped(bZipped),
	m_pZipper(0),
	m_bOwnsZipper(FALSE),
	m_ulZippedSize(bZipped ? ulSize : 0),
	m_ulUnzippedSize(bZipped ? 0 : ulSize),
	m_uCrc(0)
{
}

CAutoZipBuffer::CAutoZipBuffer(unsigned long ulSize, BOOL bZipped) :
	CAutoBuffer(ulSize),
	m_bZipped(bZipped),
	m_pZipper(0),
	m_bOwnsZipper(FALSE),
	m_ulZippedSize(bZipped ? ulSize : 0),
	m_ulUnzippedSize(bZipped ? 0 : ulSize),
	m_uCrc(0)
{
	AllocData(ulSize, bZipped);
}

CAutoZipBuffer::CAutoZipBuffer(const CAutoZipBuffer& Buffer)
{
	CopyObject(Buffer);
}

CAutoZipBuffer& CAutoZipBuffer::operator=(const CAutoZipBuffer& Buffer) 
{
	BOOL bSame = m_pData == Buffer.m_pData;

	CAutoBuffer::operator=(Buffer);

	if (bSame)
	{
		if (Buffer.m_bOwnsZipper)
		{
			m_bOwnsZipper = TRUE;
		}
	}

	return *this; 
}

void CAutoZipBuffer::CopyObject(const CAutoZipBuffer& Buffer, BOOL bRelease)
{
	CAutoBuffer::CopyObject(Buffer, bRelease);

	// m_pZipper is always shared, even if releasing.
	m_bOwnsZipper		= Buffer.m_bOwnsZipper;
	m_pZipper			= Buffer.ReleaseZipper();

	m_ulZippedSize		= Buffer.m_ulZippedSize;
	m_ulUnzippedSize	= Buffer.m_ulUnzippedSize;
	m_uCrc				= Buffer.m_uCrc;
	m_bZipped			= Buffer.m_bZipped;
}

CAutoZipBuffer::~CAutoZipBuffer() 
{ 
	Free();
}

BOOL CAutoZipBuffer::AllocData(unsigned long ulSize, BOOL bZipped)
{
	m_bZipped = bZipped;
	return CAutoBuffer::AllocData(ulSize);
}

BOOL CAutoZipBuffer::SetZippedData(unsigned long ulSize, BOOL bZipped)
{
	return AllocData(ulSize, bZipped);
}

BOOL CAutoZipBuffer::SetZippedData(char* pData, unsigned long ulSize, BOOL bZipped, BOOL bOwnsData) 
{ 
	m_bOwnsData			= CAutoBuffer::SetData(pData, ulSize, bOwnsData);

	m_bZipped			= bZipped;
	m_ulZippedSize		= bZipped ? ulSize : m_ulZippedSize;
	m_ulUnzippedSize	= bZipped ? m_ulUnzippedSize : ulSize;

	return m_bOwnsData;
}

void CAutoZipBuffer::Free()
{
	CAutoBuffer::Free();

	FreeZipper();
}

char* CAutoZipBuffer::Release() const
{
	char* pData = CAutoBuffer::Release();

	ReleaseZipper();

	return pData;
}

BOOL CAutoZipBuffer::Grow(unsigned long ulSize)
{
	BOOL bSuccess = CAutoBuffer::Grow(ulSize);

	if (bSuccess)
	{
		m_ulZippedSize		= m_bZipped ? ulSize : m_ulZippedSize;
		m_ulUnzippedSize	= m_bZipped ? m_ulUnzippedSize : ulSize;
	}

	return bSuccess;
}

BOOL CAutoZipBuffer::SetZipper(CCompAPIZipper* pZipper) 
{ 
	FreeZipper();

	m_pZipper		= pZipper;
	m_bOwnsZipper	= pZipper != 0;

	return m_bOwnsZipper;
}

void CAutoZipBuffer::FreeZipper()
{
	if (m_bOwnsZipper) 
	{
		delete m_pZipper;
		m_pZipper = 0;
	}

	ReleaseZipper();
}

CCompAPIZipper* CAutoZipBuffer::ReleaseZipper() const
{
	((CAutoZipBuffer*)this)->m_bOwnsZipper = FALSE;

	return m_pZipper; 
}

ZIPDATA* CAutoZipBuffer::GetZipData()
{ 
	ZIPDATA* pZIPDATA = 0;

	if (m_pZipper == 0)
	{
		AllocZipper();
	}
	
	if (m_pZipper)
	{
		pZIPDATA = m_pZipper->GetZipData();
	}

	return pZIPDATA;
}

unsigned short CAutoZipBuffer::UpdateCrc() 
{ 
	if (IsValid())
	{
		m_uCrc = ::crc((unsigned char*)m_pData, m_ulDataSize ? m_ulDataSize : m_ulSize); 
	}

	return m_uCrc;
}

#if defined(DALZIP)
BOOL CAutoZipBuffer::Zip(HPool hSrcPool)
{
	BOOL bSuccess = FALSE;

	_ASSERTE(!m_bZipped);

	if (IsValid() && !m_bZipped && AllocZipper())
	{
		bSuccess = m_pZipper->ZipPoolToBuffer(this, hSrcPool);
	}

	return bSuccess;
}
#endif

BOOL CAutoZipBuffer::Zip()
{
	BOOL bSuccess = FALSE;

	_ASSERTE(!m_bZipped);

	if (IsValid() && !m_bZipped && AllocZipper())
	{
		// assumes zipped data size will never be greater than size of unzipped data * 2
		// (although it is possible for zipped data size to be greater than unzipped data size)
		char* pZippedData = new char[m_ulUnzippedSize * 2];

		if (pZippedData != 0)
		{
			unsigned long ulZippedSize = 0;
			

			if (m_ulUnzippedSize <= m_nZipThreshold)
			{
				// don't zip if very small
				memcpy(pZippedData, m_pData, (size_t)m_ulUnzippedSize);
				ulZippedSize = m_ulUnzippedSize;
				bSuccess = TRUE;
			}
			else
			{
				memset(pZippedData, 0, (size_t)m_ulUnzippedSize);

				bSuccess = 
					m_pZipper->ZipBufferToBuffer(
						pZippedData, 
						m_ulUnzippedSize * 2, 
						m_pData, 
						m_ulUnzippedSize,
						&m_uCrc);
				ulZippedSize = m_pZipper->GetTgtLen();
			}

			// AS 18/05/1998 possible for compressed data to be larger
			// _ASSERTE(ulZippedSize <= m_ulUnzippedSize);

			if (bSuccess)
			{
				SetZippedData(pZippedData, ulZippedSize, TRUE, TRUE);
			}
		}
	}

	return bSuccess;
}

#if defined(DALZIP)
BOOL CAutoZipBuffer::Unzip(HPool hTgtPool)
{
	BOOL bSuccess = FALSE;

	_ASSERTE(m_bZipped);

	if (IsValid() && m_bZipped && AllocZipper())
	{
		bSuccess = m_pZipper->UnzipBufferToPool(hTgtPool, this);
	}

	return bSuccess;
}
#endif

BOOL CAutoZipBuffer::Unzip()
{
	BOOL bSuccess = FALSE;

	_ASSERTE(m_bZipped);

	if (IsValid() && m_bZipped && AllocZipper())
	{
		unsigned long ulUnzippedSize = 0;

		if (m_ulUnzippedSize > 0)
		{
			// Unzipped size is already known.
			char* pUnzippedData = new char[m_ulUnzippedSize];

			if (pUnzippedData)
			{
				if (m_ulUnzippedSize > 0 && m_ulUnzippedSize <= m_nZipThreshold)
				{
					// Don't unzip if very small.
					memcpy(pUnzippedData, m_pData, (size_t)m_ulUnzippedSize);
					ulUnzippedSize = m_ulUnzippedSize;
					bSuccess = TRUE;
				}
				else
				{
					bSuccess = 
						m_pZipper->UnzipBufferToBuffer(
							pUnzippedData, 
							m_ulUnzippedSize, 
							m_pData, 
							m_ulZippedSize);
					ulUnzippedSize = m_pZipper->GetTgtLen();
				}

				if (bSuccess)
				{
					SetZippedData(pUnzippedData, ulUnzippedSize, FALSE, TRUE);
				}
			}
		}
		else
		{
			CAutoBuffer UnzippedData;

			// Unzipped size is not already known, so automatically grow buffer to accommodate it.
			bSuccess = 
				m_pZipper->UnzipBufferToBuffer(
					&UnzippedData, 
					m_pData, 
					m_ulZippedSize);
			ulUnzippedSize = m_pZipper->GetTgtLen();

			if (bSuccess)
			{
				SetZippedData(UnzippedData.GetData(), ulUnzippedSize, FALSE, TRUE);
			}
		}

		// AS 18/05/1998 possible for compressed data to be larger
		// _ASSERTE(ulUnzippedSize >= m_ulZippedSize);
	}

	return bSuccess;
}


