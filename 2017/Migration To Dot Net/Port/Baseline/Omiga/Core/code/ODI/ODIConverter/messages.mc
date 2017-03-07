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
;//	SYSTEM:	    	ODI
;//	COPYRIGHT:		(c) 2001, Marlborough Stirling. All rights reserved 
;//
;//	HISTORY
;//	=======
;//
;//	Prog	Date		Description
;//	AS      06/08/01    Initial version
;////////////////////////////////////////////////////////////////////////////////

MessageIdTypedef=DWORD
;//SeverityNames=(
;//    Success=0x0
;//    Informational=0x1
;//    Warning=0x2
;//    Error=0x3
;//    )
;//FacilityNames=(
;//    System=0x0FF
;//    Application=0xFFF
;//    )
;//LanguageNames=(English=1:MSG00001) 

MessageId=1024
Severity=Success
Facility=Application
SymbolicName=ODICONVERTER_SUCCESS_TYPE
Language=English
%1!s!
.

MessageId=
Severity=Informational
Facility=Application
SymbolicName=ODICONVERTER_INFORMATION_TYPE
Language=English
%1!s!
.

MessageId=
Severity=Warning
Facility=Application
SymbolicName=ODICONVERTER_WARNING_TYPE
Language=English
%1!s!
.

MessageId=
Severity=Error
Facility=Application
SymbolicName=ODICONVERTER_ERROR_TYPE
Language=English
%1!s!
.
