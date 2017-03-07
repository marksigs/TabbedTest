Attribute VB_Name = "errConstants"
Option Explicit

' This enumeration is SYSTEM-WIDE errors only!!
' This is not the place to specify Omiga4 Application Component Errors
' Please can you use consecutitve error numbering from 500+
Public Enum OMIGAERROR
    oeUnspecifiedError = vbObjectError + 512 + 999
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
    oeCreditCheckError = 548
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
    oeObjectNotCreatable = 563
    oeSchemaNotLoaded = 800
    oeSchemaParseError = 801
    oeNotImplemented = 900
    oeInvalidMessageNo = 901
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

' Data Ingestion Rules component, start error number
' The error number range 7625 - 7649 has been reserved for the Data Ingestion Rules component
Private Const direStartNumber = 7625

' Custom Error Constants
Public Enum DATAINGESTIONRULESERROR
    direOmigaStartErrorNumber = vbObjectError + 512
    direMandatoryXMLNodeMissing = direOmigaStartErrorNumber + direStartNumber
    direMandatoryXMLAttributeMissing = direOmigaStartErrorNumber + direStartNumber + 1
    direValidationFailed = direOmigaStartErrorNumber + direStartNumber + 2
    direFailedToLoadXML = direOmigaStartErrorNumber + 502 ' standard Omiga Error
    direMortgageProductDoesntExist = direOmigaStartErrorNumber + direStartNumber + 3
    direLoanComponentWasNotCreated = direOmigaStartErrorNumber + direStartNumber + 4
    direFeeValuationTypeNotKnown = direOmigaStartErrorNumber + direStartNumber + 5
    direFeeControlStringMissing = direOmigaStartErrorNumber + direStartNumber + 6
    direApplicationFeeTypeMissing = direOmigaStartErrorNumber + direStartNumber + 7
End Enum
