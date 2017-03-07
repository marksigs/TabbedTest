Attribute VB_Name = "errConstants"
'header ----------------------------------------------------------------------------------
'Workfile:      errConstants.bas
'Copyright:     Copyright © 2001 Marlborough Stirling
'
'Description:   Header file for error numbers
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'History:
'
'Prog    Date        Description
'IK      01/11/00    Initial creation
'ASm     10/01/01    SYS1817: Rationalisation of methods following meeting with PC and IK
'DJP     12/03/01    SYS1839 added oeObjectNotCreatable
'DM      19/09/01    SYS2716 added oeInvalidMessageQueueType
'JR      27/05/02    SYS4639 Changed oeObjectNotCreatable message# from 563 to 576
'AC      14/04/04    BBG196 - added new constant for PLOM user verification
'------------------------------------------------------------------------------------------
Option Explicit
' This enumeration is SYSTEM-WIDE errors only!!
' This is not the place to specify Omiga4 Application Component Errors
' Please can you use consecutitve error numbering from 500+
Public Enum OMIGAERROR
    oeUnspecifiedError = vbObjectError + 512 + 999
    oeNoAcceptedQuote = 308
    oeRecordNotFound = 500
    oeXMLParserError = 502
    oeInvalidKeyString = 503
    oeNoAfterImagePresent = 504
    oeCommandFailed = 505
    oeMissingPrimaryTag = 506
    oeInvalidParameter = 507
    oeDuplicateKey = 508
    oeNoDataForCreate = 509
    oeNoDataForUpdate = 510
    oeArrayLimitExceeded = 511
    oeNoRowsAffected = 512
    oeInvalidNoOfRows = 513
    oeInternalError = 514
    oeMissingFieldDesc = 515
    oeMissingTableDesc = 516
    oeMissingTypeDesc = 517
    oeMissingElementValue = 518
    oeMissingElement = 519
    oeMissingKeyDesc = 520
    oeMissingUpdateSet = 521
    oeMissingTableName = 522
    oeMissingXMLTableName = 523
    oeMissingKey = 524
    oeNoRowsAffectedByDeleteAll = 525
    oeInValidKeyValue = 526
    oeInValidDataTypeValue = 527
    oeInValidKey = 528
    oeNoFieldsFound = 529
    oeNoFieldItemFound = 530
    oeNoFieldItemName = 531
    oeNoComboTagValue = 532
    oeInvalidDateTimeFormat = 533
    oeNoBeforeImagePresent = 534
    oeChildRecordsFound = 535
    oeMTSNotFound = 555
    oeUnableToConnect = 556
    oeMissingParameter = 557
    oeXMLMissingElement = oeMissingElement
    oeXMLMissingElementText = oeMissingElementValue
    oeXMLMissingAttribute = 558
    oeXMLInvalidAttributeValue = 559
    oeXMLMissingPrimaryKey = 560
    oeXMLInvalidRequestNode = 561
    oeXMLAttributeAlreadyExists = 562
    oeObjectNotCreatable = 576
    oeSchemaNotLoaded = 800
    oeSchemaParseError = 801
    oeNotImplemented = 900
    oeInvalidMessageNo = 901
    oeInvalidMessageQueueType = 902
    oeInvalidPLOMUserLevel = 10002
End Enum
Public Enum OMIGAERRORTYPE
    omWARNING
    omNO_ERR
End Enum
Public Enum MESSAGE
    omMESSAGE_TEXT
    omMESSAGE_TYPE
End Enum
Public Const clngMAX_ERROR_NO = 10000
Public Const clngADO_START_ERROR_NO = 3000
Public Const clngADO_END_ERROR_NO = 3999
Public Enum ERRORS
    eNOTMIPLEMENTED = vbObjectError + 512
    eXMLMISSINGELEMENT
    eXMLMISSINGATTRIBUTE
    eXMLINVALIDATTTRIBUTE
    eXMLINVALIDATTTRIBUTEVALUE
    eRECORDNOTFOUND
    eUNSPECIFIEDERROR = vbObjectError + 512 + 999
    eFAILEDTOLOADREQUEST = vbObjectError + 512 + 1
End Enum
