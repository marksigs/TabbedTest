///////////////////////////////////////////////////////////////////////////////
//	FILE:			DbHelperOMMQ.h
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

#ifndef DBHELPEROMMQ_H
#define DBHELPEROMMQ_H

class CDbHelperOMMQ
{
public:
	CDbHelperOMMQ() {};
	~CDbHelperOMMQ() {};

	static enum eDataAction
	{
		eGetMessage,
		eLockMessage,
		eMoveMessage,
		eResetQueue,
		eSendMessage,
		eSendMessageOrNtext,
		eUnLockMessage,
		eUnlockMessages
	};

	bstr_t GetSQLOLEDBCommandText(
		BSTR bstrTableSuffix,
		CDbHelperOMMQ::eDataAction eSprocType);
};

#endif // DBHELPEROMMQ_H