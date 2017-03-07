using System;
namespace omLockManager
{
    public class StdData
    {

        // Workfile:      StdData.cls
        // Copyright:     Copyright © 1999 Marlborough Stirling
        // Description:   Standard data
        // Dependencies:
        // Issues:        This module should be common to all Omiga4 projects but not
        // shared between them, i.e. each component should have its own
        // version.
        // ------------------------------------------------------------------------------------------
        // History:
        // 
        // Prog  Date     Description
        // MCS    08/11/99 Created
        // PSC    05/01/01 Added Administration System Interface, omAdmin
        // PSC    07/03/01 SYS1879 Add Application Processing, omAppProc
        // PSC    17/04/01 SYS2188 Add Omiga Task Management, omTM and
        // Batch Scheduler, omBatch
        // MDC    05/06/01 Add Admin Rules (omAdminRules) constant
        // SR     06/09/01 New constants for 'ReportOnTitle' and 'RateChange'
        // LD     03/10/01 New constant gstrFVS_COMPONENT
        // DM     23/11/01 New constant gstrMESSAGE_QUEUE_COMPONENT
        // DR     27/02/02 New constant gstrPRINTMANAGER_COMPONENT
        // ------------------------------------------------------------------------------------------
        // BBG History
        // 
        // Prog   Date        AQR         Description
        // BS     09/04/04    New constant gstrTRANSACT_COMPONENT
        // MV     04/06/2004  BBG48 - 21/01/2004  BMIDS646    Added Compensating Resource Manager.
        // ------------------------------------------------------------------------------------------
        // old refs, do not use...here for backwards compatability.
        public const string gstrCUSTOMEREMPLOYMENT_COMPONENT = "omCE";
        public const string gstrCUSTOMER_FINANCIAL_COMPONENT = "omCF";
        public const string gstrCOMPONENT_NAME = "omCR";
        public const string gstrSUBMISSION_COMPONENT = "omSub";
        public const string gstrCREDITCHECK_COMPONENT = "omCC";
        public const string gstrRISKASSESSMENT_COMPONENT = "omRA";
        public const string gstrDOWNLOAD_COMPONENT = "om4To3";
        // new refs, programmers use these...
        public const string gstrOMIGA3_MANAGER_COMPONENT = "Om3Manager";
        public const string gstrOMIGA4ToOMIGA3DOWNLOAD = "om4to3";
        public const string gstrAPPLICATION_COMPONENT = "omApp";
        public const string cstrTABLE_NAME = "ADDRESS";
        public const string gstrAPPLICATIONQUOTE = "omAQ";
        public const string gstrAUDIT_COMPONENT = "omAU";
        public const string gstrBASE_COMPONENT = "omBase";
        public const string gstrBUILDINGSANDCONTENTS = "omBC";
        public const string gstrCREDIT_CHECK = "omCC";
        public const string gstrCOST_MODEL_COMPONENT = "omCM";
        public const string gstrCOMPLETENESS_RULES_COMP = "omCompRules";
        public const string gstrCUSTOMER_EMPLOYMENT = "omCE";
        public const string gstrCUSTOMER_FINANCIAL = "omCF";
        public const string gstrCUSTOMER_COMPONENT = "omCust";
        public const string gstrDPS_COMPONENT = "omDPS";
        public const string gstrEXPERIAN_COMPONENT = "omExp";
        public const string gstrIMPORT_COMPONENT = "OmImp";
        public const string gstrLIFECOVER_COMPONENT = "omLC";
        public const string gstrMORTGAGEPRODUCT = "omMP";
        public const string gstrORGANISATION_COMPONENT = "omOrg";
        public const string gstrPP_COMPONENT = "omPP";
        public const string gstrQUICK_QUOTE = "omQQ";
        public const string gstrRISK_ASSESSMENT = "omRA";
        public const string gstrRISK_ASSESSMENT_RULES_COMP = "omRARules";
        public const string gstrSubmission = "omSub";
        public const string gstrREQUEST_BROKER_COMPONENT = "omRB";
        public const string gstrTHIRDPARTY_COMPONENT = "omTP";
        public const string gstrADMIN_INTERFACE = "omAdmin";
        public const string gstrPAYMENTPROCESSING = "omPayProc";
        public const string gstrAPPLICATIONPROCESSING = "omAppProc";
        public const string gstrMsgTm_COMPONENT = "MsgTm";
        public const string gstrPRINT_COMPONENT = "omPrint";
        public const string gstrTASKMANAGEMENT_COMPONENT = "omTM";
        public const string gstrBATCH_SCHEDULER_COMPONENT = "omBatch";
        public const string gstrADMINRULES = "omAdminRules";
        public const string gstrREPORTONTITLE_COMPONENT = "omROT";
        public const string gstrRATECHANGE = "omRC";
        public const string gstrMESSAGE_QUEUE_COMPONENT = "omMessageQueue";
        public const string gstrPRINTMANAGER_COMPONENT = "omPM";
        public const string gstrHUNTERINTERFACE_COMPONENT = "omHI";
        // new refs added for use with Visual Studio Analyser
        public const string gstrAIP_COMPONENT = "omAIP";
        public const string gstrCOMP_CHECK_COMPONENT = "omComp";
        public const string gstrDECISION_MANAGER_COMPONENT = "omDM";
        public const string gstrHOME_USE_COMPONENT = "omHU";
        public const string gstrINTERMEDIARY = "omIM";
        public const string gstrPAF = "omPAF";
        public const string gstrFVS_COMPONENT = "omFVS";
        public const string gstrMSGSPM = "MTxSpm.SharedPropertyGroupManager.1";
        public const string gstrTRANSACT_COMPONENT = "omTransact";
        public const string gstrCRM_COMPONENT = "omCRMLock";
#if TIMINGS
        public const string TIMEFORMAT = "000.000";
#endif 

    }

}
