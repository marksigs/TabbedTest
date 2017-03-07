using System;
namespace omLockManager
{

    // header ----------------------------------------------------------------------------------
    // Workfile:      errConstants.bas
    // Copyright:     Copyright © 2001 Marlborough Stirling
    // 
    // Description:   Header file for error numbers
    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    // History:
    // 
    // Prog    Date        Description
    // IK      01/11/00    Initial creation
    // ASm     10/01/01    SYS1817: Rationalisation of methods following meeting with PC and IK
    // DJP     12/03/01    SYS1839 added oeObjectNotCreatable
    // DM      19/09/01    SYS2716 added oeInvalidMessageQueueType
    // JR      27/05/02    SYS4639 Changed oeObjectNotCreatable message# from 563 to 576
    // AC      14/04/04    BBG196 - added new constant for PLOM user verification
    // ------------------------------------------------------------------------------------------
    // This enumeration is SYSTEM-WIDE errors only!!
    // This is not the place to specify Omiga4 Application Component Errors
    // Please can you use consecutitve error numbering from 500+
    public enum OMIGAERROR {
        //[DS]- May 1st 2007
        //oeUnspecifiedError = 0x80040000 + 512 + 999,
        oeUnspecifiedError = Microsoft.VisualBasic.Constants.vbObjectError + 512 + 999,
        oeNoAcceptedQuote = 308,
        oeRecordNotFound = 500,
        oeXMLParserError = 502,
        oeInvalidKeyString = 503,
        oeNoAfterImagePresent = 504,
        oeCommandFailed = 505,
        oeMissingPrimaryTag = 506,
        oeInvalidParameter = 507,
        oeDuplicateKey = 508,
        oeNoDataForCreate = 509,
        oeNoDataForUpdate = 510,
        oeArrayLimitExceeded = 511,
        oeNoRowsAffected = 512,
        oeInvalidNoOfRows = 513,
        oeInternalError = 514,
        oeMissingFieldDesc = 515,
        oeMissingTableDesc = 516,
        oeMissingTypeDesc = 517,
        oeMissingElementValue = 518,
        oeMissingElement = 519,
        oeMissingKeyDesc = 520,
        oeMissingUpdateSet = 521,
        oeMissingTableName = 522,
        oeMissingXMLTableName = 523,
        oeMissingKey = 524,
        oeNoRowsAffectedByDeleteAll = 525,
        oeInValidKeyValue = 526,
        oeInValidDataTypeValue = 527,
        oeInValidKey = 528,
        oeNoFieldsFound = 529,
        oeNoFieldItemFound = 530,
        oeNoFieldItemName = 531,
        oeNoComboTagValue = 532,
        oeInvalidDateTimeFormat = 533,
        oeNoBeforeImagePresent = 534,
        oeChildRecordsFound = 535,
        oeMTSNotFound = 555,
        oeUnableToConnect = 556,
        oeMissingParameter = 557,
        oeXMLMissingElement = OMIGAERROR.oeMissingElement,
        oeXMLMissingElementText = OMIGAERROR.oeMissingElementValue,
        oeXMLMissingAttribute = 558,
        oeXMLInvalidAttributeValue = 559,
        oeXMLMissingPrimaryKey = 560,
        oeXMLInvalidRequestNode = 561,
        oeXMLAttributeAlreadyExists = 562,
        oeObjectNotCreatable = 576,
        oeSchemaNotLoaded = 800,
        oeSchemaParseError = 801,
        oeNotImplemented = 900,
        oeInvalidMessageNo = 901,
        oeInvalidMessageQueueType = 902,
        oeInvalidPLOMUserLevel = 10002,
    }


    public enum OMIGAERRORTYPE {
        omWARNING,
        omNO_ERR,
    }


    public enum MESSAGE {
        omMESSAGE_TEXT,
        omMESSAGE_TYPE,
    }


    public enum ERRORS {
        eNOTMIPLEMENTED = Microsoft.VisualBasic.Constants.vbObjectError + 512,
        eXMLMISSINGELEMENT,
        eXMLMISSINGATTRIBUTE,
        eXMLINVALIDATTTRIBUTE,
        eXMLINVALIDATTTRIBUTEVALUE,
        eRECORDNOTFOUND,
        eUNSPECIFIEDERROR = Microsoft.VisualBasic.Constants.vbObjectError + 512 + 999,
        eFAILEDTOLOADREQUEST = Microsoft.VisualBasic.Constants.vbObjectError + 512 + 1,
    }

    public class errConstants
    {

        public const int clngMAX_ERROR_NO = 10000;
        public const int clngADO_START_ERROR_NO = 3000;
        public const int clngADO_END_ERROR_NO = 3999;

    }

}
