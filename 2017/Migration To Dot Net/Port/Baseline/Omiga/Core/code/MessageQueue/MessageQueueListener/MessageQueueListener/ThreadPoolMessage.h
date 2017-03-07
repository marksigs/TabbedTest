///////////////////////////////////////////////////////////////////////////////
//	FILE:			ThreadPoolMessage.h
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2000, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      30/08/00    Initial version
///////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_THREADPOOLMESSAGE_H__A94D0FA6_4B60_11D4_8237_005004E8D1A7__INCLUDED_)
#define AFX_THREADPOOLMESSAGE_H__A94D0FA6_4B60_11D4_8237_005004E8D1A7__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

///////////////////////////////////////////////////////////////////////////////

class CThreadPoolMessage : public CObject
{
public:
    typedef void* funcparam;
    typedef int (*funcptr)(funcparam);  

    CThreadPoolMessage() :
        m_pFunction(NULL), 
		m_pFunctionParameter(NULL)
    {
	}
    CThreadPoolMessage(funcptr pFunction, funcparam pFunctionParameter) : 
        m_pFunction(pFunction), 
		m_pFunctionParameter(pFunctionParameter)
	{
	}
	virtual ~CThreadPoolMessage()
	{
	}

public:
    void ExecuteAndDelete() 
	{
		if (m_pFunction) 
		{
			(*m_pFunction)(m_pFunctionParameter);
		} 
		m_pFunction = NULL; 
		m_pFunctionParameter = NULL; 
		delete this;
	}

private:
    funcptr m_pFunction;
    funcparam m_pFunctionParameter;
};

///////////////////////////////////////////////////////////////////////////////

#endif // !defined(AFX_THREADPOOLMESSAGE_H__A94D0FA6_4B60_11D4_8237_005004E8D1A7__INCLUDED_)
