///////////////////////////////////////////////////////////////////////////////
//	FILE:			DbHelperOMMQ.cpp
//	DESCRIPTION:	
//	SYSTEM: 		Omiga
//	COPYRIGHT:		(c) 2004, Marlborough Stirling Group. All rights reserved 
//
//	HISTORY
//	=======
//
//	Prog	Date		Description
//  RF		24/02/04	BMIDS727 - Different queues can use different tables
///////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <comdef.h>		// for bstr_t
#include "DbHelperOMMQ.h"

bstr_t CDbHelperOMMQ::GetSQLOLEDBCommandText(
	BSTR bstrTableSuffix,
	CDbHelperOMMQ::eDataAction eSprocType)
{
	bstr_t bstrSprocName;

	try
	{
		switch (eSprocType)
		{
			case CDbHelperOMMQ::eGetMessage:
				bstrSprocName = bstr_t("USP_MQLOMMQGETMESSAGEORNTEXT");
				break;
			case CDbHelperOMMQ::eLockMessage:
				bstrSprocName = bstr_t("USP_MQLOMMQLOCKMESSAGE");
				break;
			case CDbHelperOMMQ::eMoveMessage:
				bstrSprocName = bstr_t("USP_MQLOMMQMOVEMESSAGE");
				break;
			case CDbHelperOMMQ::eResetQueue:
				bstrSprocName = bstr_t("USP_MQLOMMQRESETQUEUE");
				break;
			case CDbHelperOMMQ::eSendMessage:
				bstrSprocName = bstr_t("USP_MQLOMMQSENDMESSAGENTEXT");
				break;
			case CDbHelperOMMQ::eSendMessageOrNtext:
				bstrSprocName = bstr_t("USP_MQLOMMQSENDMESSAGENTEXT");
				break;
			case CDbHelperOMMQ::eUnLockMessage:
				bstrSprocName = bstr_t("USP_MQLOMMQUNLOCKMESSAGE");
				break;
			case CDbHelperOMMQ::eUnlockMessages:
				bstrSprocName = bstr_t("USP_MQLOMMQUNLOCKMESSAGES");
				break;
			default :
				_ASSERTE(0); // should not reach here 
				throw L"Unrecognised stored procedure type";
				break;
		}

		// add table suffix
		bstrSprocName += bstrTableSuffix;

		return bstrSprocName.copy();
	}
	catch(...)
	{
		_ASSERTE(0); // should not reach here 
		return bstr_t("");
	}
};

