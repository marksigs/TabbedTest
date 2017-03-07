///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueListenerWMI.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2001, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  LD      12/01/01    Initial version
///////////////////////////////////////////////////////////////////////////////

#include <fwcommon.h>  // This must be the first include.
#include "MessageQueueListenerWMI.h"

#import "..\MessageQueueListener\MessageQueueListener.tlb" no_namespace

CMessageQueueListenerWMI MyMessageQueueListenerWMISet (PROVIDER_NAME_MESSAGEQUEUELISTENERWMI, L"root\\cimv2\\Marlborough_Stirling") ;

///////////////////////////////////////////////////////////////////////////////

CMessageQueueListenerWMI::CMessageQueueListenerWMI (LPCWSTR lpwszName, LPCWSTR lpwszNameSpace ) :
    Provider(lpwszName, lpwszNameSpace)
{
}

CMessageQueueListenerWMI::~CMessageQueueListenerWMI ()
{
}


/*****************************************************************************
*
*  FUNCTION    :    CMessageQueueListenerWMI::ExecMethod
*
*  DESCRIPTION :    Override this function to provide support for methods.  
*                   A method is an entry point for the user of your provider 
*                   to request your class perform some function above and 
*                   beyond a change of state.  (A change of state should be 
*                   handled by PutInstance() )
*
*  INPUTS      :    A pointer to a CInstance containing the instance the method was executed against.
*                   A string containing the method name
*                   A pointer to the CInstance which contains the IN parameters.
*                   A pointer to the CInstance to contain the OUT parameters.
*                   A set of specialized method flags
*
*  RETURNS     :    WBEM_E_PROVIDER_NOT_CAPABLE if not implemented for this class
*                   WBEM_S_NO_ERROR if method executes successfully
*                   WBEM_E_FAILED if error occurs executing method 
*
*  COMMENTS    : 
*
*****************************************************************************/
HRESULT CMessageQueueListenerWMI::ExecMethod ( const CInstance& Instance,
                        const BSTR bstrMethodName,
                        CInstance *pInParams,
                        CInstance *pOutParams,
                        long lFlags)
{
	
	HRESULT hr = WBEM_E_PROVIDER_NOT_CAPABLE;

	if (wcscmp(bstrMethodName, L"Configure") == 0)
	{
		// process the Configure method
		enum
		{
			eNULL,
			eGetXMLIn,
			eCreateObject,
			eCallConfigure,
			eSetXMLOut,
			eDone
		} eAction = eNULL;

		try
		{
			// forward the XML in to the MessageQueueListener component and return XML out
			eAction = eGetXMLIn;
			CHString strXMLIn;
			if (pInParams->GetCHString(L"strXMLIn", strXMLIn))
			{
				eAction = eCreateObject;
				IMessageQueueListener1Ptr MessageQueueListener1Ptr(__uuidof(MessageQueueListener1));
				if (MessageQueueListener1Ptr)
				{
					eAction = eCallConfigure;
					_bstr_t bstrXMLOut = MessageQueueListener1Ptr->Configure(_bstr_t(strXMLIn));
					
					eAction = eSetXMLOut;
					if (pOutParams->SetCHString(L"strXMLOut", (LPWSTR)bstrXMLOut))
					{
						eAction = eDone;		
					}	
				}
			}	
		}
		catch (...)
		{
		}
	
		
		// set the return value
		enum
		{
			CONFIGURE_SUCCESS_CALL = 0,
			CONFIGURE_FAIL_GETXMLIN = 1,
			CONFIGURE_FAIL_CREATEOBJECT = 2,
			CONFIGURE_FAIL_CALLCONFIGURE = 3,
			CONFIGURE_FAIL_SETXMLOUT = 4,
			CONFIGURE_FAIL_UNKNOWN = 5
		} dwReturnValue;
		switch (eAction)
		{
			case eGetXMLIn:
				dwReturnValue = CONFIGURE_FAIL_GETXMLIN;
				break;
			case eCreateObject:
				dwReturnValue = CONFIGURE_FAIL_CREATEOBJECT;
				break;
			case eCallConfigure:
				dwReturnValue = CONFIGURE_FAIL_CALLCONFIGURE;
				break;
			case eSetXMLOut:
				dwReturnValue = CONFIGURE_FAIL_SETXMLOUT;
				break;
			case eDone:
				dwReturnValue = CONFIGURE_SUCCESS_CALL;
				break;
			default :
				dwReturnValue = CONFIGURE_FAIL_UNKNOWN;
				break;
		}
		if (pOutParams->SetWORD(L"returnValue", dwReturnValue))
		{
			hr = WBEM_S_NO_ERROR;
		}
		else
		{
			hr = WBEM_E_FAILED;
		}
	}
	return hr;
}
