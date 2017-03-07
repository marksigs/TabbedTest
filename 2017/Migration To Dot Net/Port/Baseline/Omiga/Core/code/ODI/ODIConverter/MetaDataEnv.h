///////////////////////////////////////////////////////////////////////////////
//	FILE:			MetaDataEnv.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  AS		28/09/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_METADATAENV_H__73FA5685_46E4_40F4_8139_F7E1E2AC9C32__INCLUDED_)
#define AFX_METADATAENV_H__73FA5685_46E4_40F4_8139_F7E1E2AC9C32__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class CMetaDataEnv  
{
public:
	CMetaDataEnv(LPCWSTR szType, LPCWSTR szName);
	virtual ~CMetaDataEnv();
	static CMetaDataEnv* Create(LPCWSTR szType, LPCWSTR szName);
	virtual bool Init(const MSXML::IXMLDOMNodePtr ptrXMLDOMNodeParent) = 0;
	virtual void Save(const MSXML::IXMLDOMNodePtr ptrXMLDOMNodeParent) = 0;
	bool GetInitialised() { return m_bInitialised; }
	_bstr_t GetName() const { return m_bstrName; }
	_bstr_t GetType() const { return m_bstrType; }
	void SetType(LPCWSTR szType) { m_bstrType = szType; }
	namespaceMutex::CSharedLock& GetSharedLock() { return m_sharedlock; }

protected:
	void SetInitialised(bool bInitialised) { m_bInitialised = bInitialised; }

private:
	namespaceMutex::CSharedLockSmallOperations m_sharedlock;
	bool				m_bInitialised;				// Initialisation flag.
	_bstr_t				m_bstrName;					// The unique name for this object.
	_bstr_t				m_bstrType;					// The type of this object (corresponds to derived class).
};


#endif // !defined(AFX_METADATAENV_H__73FA5685_46E4_40F4_8139_F7E1E2AC9C32__INCLUDED_)

