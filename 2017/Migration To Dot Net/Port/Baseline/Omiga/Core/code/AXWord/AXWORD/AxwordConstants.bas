Attribute VB_Name = "AxwordConstants"
'header ----------------------------------------------------------------------------------
'Workfile:      errAxwordConstants.bas
'Copyright:     Copyright © 2002 Marlborough Stirling
'
'Description:   Header file for globals within just Axword
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'History:
'
'Prog   Date        Description
'DR     05/03/02    Initial creation
'DJB    08/05/02    Changed to general global file.
'------------------------------------------------------------------------------------------

Option Explicit

' Enumerations.
Public Enum AXWORDERROR
    
    axUnspecifiedError = vbObjectError + 512 + 999
    
    'XML related errors
    axXMLParserError = vbObjectError + 512 + 1
    axXMLZeroLengthString = vbObjectError + 512 + 2 'No XML supplied
    axXMLMissingMandatoryNode = vbObjectError + 512 + 3
    
    'Word Related Errors
    axWINWORDProcessError = vbObjectError + 512 + 10 'Word is already running on the system
    axWINWORDOpenError = vbObjectError + 512 + 11 'Word couldnt open a file
    
    'OLE Related Errors
    axOLECreateEmbedError = vbObjectError + 512 + 20 'Cant create the embedded object, permission error? registry error?
    
End Enum

Public Enum DELIVERYTYPE
    DELIVERYTYPE_UNK = 0
    DELIVERYTYPE_DOC = 10
    DELIVERYTYPE_PDF = 20
    DELIVERYTYPE_RTF = 30
    DELIVERYTYPE_XML = 40
End Enum

