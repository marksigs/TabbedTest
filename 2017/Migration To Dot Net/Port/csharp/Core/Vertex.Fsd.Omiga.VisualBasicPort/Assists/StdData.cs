/*
--------------------------------------------------------------------------------------------
Workfile:			StdData.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Standard data. 
					This module should be common to all Omiga4 projects but not shared between
					them, i.e. each component should have its own version.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
MCS		08/11/1999	Created
PSC		05/01/2001	Added Administration System Interface, omAdmin
PSC		07/03/2001	SYS1879 Add Application Processing, omAppProc
PSC		17/04/2001	SYS2188 Add Omiga Task Management, omTM and Batch Scheduler, omBatch
MDC		05/06/2001	Add Admin Rules (omAdminRules) constant
SR		06/09/2001	New constants for 'ReportOnTitle' and 'RateChange'
LD		03/10/2001	New constant gstrFVS_COMPONENT
DM		23/11/2001	New constant gstrMESSAGE_QUEUE_COMPONENT
DR		27/02/2002	New constant gstrPRINTMANAGER_COMPONENT
------------------------------------------------------------------------------------------
BBG History

Prog   Date			Description
BS	 09/04/2004	New constant gstrTRANSACT_COMPONENT
MV	 04/06/2004	BBG48 - 21/01/2004  BMIDS646	Added Compensating Resource Manager.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from StdData.bas.
--------------------------------------------------------------------------------------------
*/
using System;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Defines	constant strings for the names of some Omiga components.
	/// </summary>
	public static class StdData
	{
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
		// PSC 03/04/2006 MAR1573
		public const string gstrCRUD_COMPONENT = "omCRUD";
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
		// JD MAR41 new object hometrack
		public const string gstrHOMETRACK_COMPONENT = "omHomeTrack";
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
