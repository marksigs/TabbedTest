/******************************************************************************
Module name: Optex.h
Written by:  Jeffrey Richter
Purpose:     Defines the COptex (optimized mutex) synchronization object
******************************************************************************/


#pragma once


///////////////////////////////////////////////////////////////////////////////

LPSECURITY_ATTRIBUTES SetSecurityDescriptorNullDacl(LPSECURITY_ATTRIBUTES lpsa, PSECURITY_DESCRIPTOR psd);

class COptex 
{
public:
	COptex(LPCSTR pszName,  DWORD dwSpinCount = 4000);
	COptex(LPCWSTR pszName, DWORD dwSpinCount = 4000);
	~COptex();
	void SetSpinCount(DWORD dwSpinCount);
	void Enter();
	BOOL TryEnter();
	void Leave();
	BOOL IsValid() { return m_hevt != NULL && m_hfm != NULL && m_pSharedInfo != NULL; }

private:
	typedef struct 
	{
		DWORD m_dwSpinCount;
		long  m_lLockCount;
		DWORD m_dwThreadId;
		long  m_lRecurseCount;
	} SHAREDINFO, *PSHAREDINFO;

	BOOL        m_fUniprocessorHost;
	HANDLE      m_hevt;
	HANDLE      m_hfm;
	PSHAREDINFO m_pSharedInfo;

private:
	BOOL CommonConstructor(PVOID pszName, BOOL fUnicode, DWORD dwSpinCount);
};


///////////////////////////////////////////////////////////////////////////////


inline COptex::COptex(LPCSTR pszName, DWORD dwSpinCount) 
{
	CommonConstructor((PVOID) pszName, FALSE, dwSpinCount);
}


///////////////////////////////////////////////////////////////////////////////


inline COptex::COptex(LPCWSTR pszName, DWORD dwSpinCount) 
{
	CommonConstructor((PVOID) pszName, TRUE, dwSpinCount);
}


///////////////////////////////// End of File /////////////////////////////////
