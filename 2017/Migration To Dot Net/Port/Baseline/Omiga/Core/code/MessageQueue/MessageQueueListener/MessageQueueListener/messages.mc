;///////////////////////////////////////////////////////////////////////////////
;//	FILE:			MESSAGES.MC/MESSAGES.H generated from MESSAGES.MC
;//	DESCRIPTION: 	
;//		Message Compile (MC) file for messages written to the event log.
;//		Compile as follows:
;//			mc -c messages.mc
;//		The -c option marks all messages as user-defined by setting the Customer 
;//		code flag (the 29th bit of the 32-bit message). This bit can be used to 
;//		determine if a message has come from the system or from an application.
;//		The resulting messages.rc and messages.h are included into the 
;//		project, so that the messages are available.
;//
;//	SYSTEM:	    	EServer
;//	COPYRIGHT:		(c) 1998, Marlborough Stirling. All rights reserved 
;//
;//	HISTORY
;//	=======
;//
;//	Prog	Date		Description
;//	LD      30/08/00    Initial version
;////////////////////////////////////////////////////////////////////////////////

MessageIdTypedef=DWORD

MessageId=1024
Severity=Success
Facility=Application
SymbolicName=MESSAGEQUEUELISTENER_SUCCESS_TYPE
Language=English
%1!s!.
.

MessageId=
Severity=Informational
Facility=Application
SymbolicName=MESSAGEQUEUELISTENER_INFORMATION_TYPE
Language=English
%1!s!.
.

MessageId=
Severity=Warning
Facility=Application
SymbolicName=MESSAGEQUEUELISTENER_WARNING_TYPE
Language=English
%1!s!.
.

MessageId=
Severity=Error
Facility=Application
SymbolicName=MESSAGEQUEUELISTENER_ERROR_TYPE
Language=English
%1!s!.
.

