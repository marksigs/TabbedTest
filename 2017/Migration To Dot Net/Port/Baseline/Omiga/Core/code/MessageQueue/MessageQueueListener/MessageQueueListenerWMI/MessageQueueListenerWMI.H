///////////////////////////////////////////////////////////////////////////////
//	FILE:			MessageQueueListenerWMI.h
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

#ifndef _MessageQueueListenerWMI_H_
#define _MessageQueueListenerWMI_H_

#define PROVIDER_NAME_MESSAGEQUEUELISTENERWMI L"MessageQueueListenerWMI"

///////////////////////////////////////////////////////////////////////////////

class CMessageQueueListenerWMI : public Provider 
{
    public:
        // Constructor/destructor
        //=======================

        CMessageQueueListenerWMI(LPCWSTR lpwszClassName, LPCWSTR lpwszNameSpace);
        virtual ~CMessageQueueListenerWMI();

    protected:
        virtual HRESULT ExecMethod( const CInstance& Instance,
                        const BSTR bstrMethodName,
                        CInstance *pInParams,
                        CInstance *pOutParams,
                        long lFlags = 0L );
} ;

///////////////////////////////////////////////////////////////////////////////

#endif
