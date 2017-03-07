<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP010.asp
Copyright:     Copyright © 2006 Marlborough Stirling

Description:   Application/Approval summary menu
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
LH		14/09/06	Created

*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Forms Here */ %>
<form id="frmToMN070" method="post" action="MN070.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP040" method="post" action="AP040.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" >
<div id="divSummaryDetails" style="TOP: 60px; LEFT: 10px; HEIGHT: 250px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
	Summary Details:
</span>
<span style="TOP:70px; LEFT:200px; POSITION:ABSOLUTE">
	<input id="btnAppDetails" value="Application Details" type="button" style="HEIGHT: 40px;WIDTH:250px" class="msgButton">
</span>
</div>
<div id="divAppStatus" style="TOP: 320px; LEFT: 10px; HEIGHT: 70px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
	Current Application Status:
</span>
<span style="TOP:30px; LEFT:54px; POSITION:ABSOLUTE">
	<input id="btnDecline" value="Decline" type="button" style="HEIGHT: 29px; WIDTH:100px" class="msgButton">
	<img id="imgDeclineTickInd" src="images/chk_tick.gif" style=" LEFT: 105px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE" WIDTH="29" HEIGHT="29">
	<img id="imgDeclineBlankInd" src="images/chk_cross.gif" style=" LEFT: 105px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: VISIBLE; POSITION: ABSOLUTE" WIDTH="29" HEIGHT="29">
</span>
<span style="TOP:30px; LEFT:220px; POSITION:ABSOLUTE">
	<input id="btnRecommend" value="Recommend" type="button" style="HEIGHT: 29px; WIDTH:100px" class="msgButton">
	<img id="imgRecommendTickInd" src="images/chk_tick.gif" style=" LEFT: 105px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE" WIDTH="29" HEIGHT="29">
	<img id="imgRecommendBlankInd" src="images/chk_cross.gif" style=" LEFT: 105px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: VISIBLE; POSITION: ABSOLUTE" WIDTH="29" HEIGHT="29">
</span>
<span style="TOP:30px; LEFT:386px; POSITION:ABSOLUTE">
	<input id="btnApprove" value="Approve" type="button" style="HEIGHT: 29px; WIDTH:100px" class="msgButton">
	<img id="imgApproveTickInd" src="images/chk_tick.gif" style=" LEFT: 105px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: HIDDEN; POSITION: ABSOLUTE" WIDTH="29" HEIGHT="29">
	<img id="imgApproveBlankInd" src="images/chk_cross.gif" style=" LEFT: 105px; HEIGHT: 29px; WIDTH: 29px; VISIBILITY: VISIBLE; POSITION: ABSOLUTE" WIDTH="29" HEIGHT="29">
</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 395px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/AP010Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_sReadOnly = "";
var m_sTaskXML = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sStageId = "";
var m_sActivityId = "";
var m_sUserId = "";
var m_sUnitId = "";
var m_sUnitName = "";
var m_sAuthorityLevel = "";
var m_sAppPriority = "";
var m_taskXML = null;
var scScreenFunctions;
var m_bButtonsDisabled = false;
var m_sRecommendedUserId = "";
var m_blnReadOnly = false;
var m_sCancelStageId = "";
var m_sDeclineStageId = "";

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Application/Approval Summary","AP010",scScreenFunctions);

	RetrieveContextData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	Initialise();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP010");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sStageId = scScreenFunctions.GetContextParameter(window,"idStageId",null);
	m_sActivityId = scScreenFunctions.GetContextParameter(window,"idActivityId",null);
	m_sUserId = scScreenFunctions.GetContextParameter(window,"idUserId",null);
	m_sUnitId = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	m_sUnitName = scScreenFunctions.GetContextParameter(window,"idUnitName",null);
	m_sAuthorityLevel = scScreenFunctions.GetContextParameter(window,"idRole",null);
	m_sAppPriority = scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null);
	<% /* MO - 26/11/2002 - BMIDS01089 */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_sDeclineStageId = XML.GetGlobalParameterAmount(document, "DeclinedStageValue");
	m_sCancelStageId = XML.GetGlobalParameterAmount(document, "CancelledStageValue");
	XML = null;
	
}
function Initialise()
{
		
	//SYS2173 - GD 20/04/01
	if (m_sAuthorityLevel < 50) //Underwriter
	{
		DisableButtons();
	}
	
	if(m_sTaskXML.length > 0 && m_taskXML == null)
	{
		m_taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_taskXML.LoadXML(m_sTaskXML);
	}
	var AppXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AppXML.CreateRequestTag(window, null)
	AppXML.CreateActiveTag("APPLICATION");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	// 	AppXML.RunASP(document,"GetApplicationData.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			AppXML.RunASP(document,"GetApplicationData.asp");
			break;
		default: // Error
			AppXML.SetErrorResponse();
		}

	if(AppXML.IsResponseOK())
	{
		AppXML.SelectTag(null, "APPLICATIONFACTFIND");
		m_sRecommendedUserId = AppXML.GetTagText("APPLICATIONRECOMMENDEDUSERID")
		if(AppXML.GetTagText("APPLICATIONRECOMMENDEDDATE")!= "")
		{
			ChangeTickStatus(frmScreen.imgRecommendBlankInd,frmScreen.imgRecommendTickInd);
			frmScreen.btnRecommend.disabled = true;
		}
		else
		{
			ChangeTickStatus(frmScreen.imgRecommendTickInd,frmScreen.imgRecommendBlankInd);
			var XML	= new top.frames[1].document.all.scXMLFunctions.XMLObject();
			if(XML.GetGlobalParameterBoolean(document, "RecommendApplicationInd"))
				frmScreen.btnApprove.disabled = true;
			XML = null;
		}
		if(AppXML.GetTagText("APPLICATIONAPPROVALDATE") != "")
		{
			ChangeTickStatus(frmScreen.imgApproveBlankInd,frmScreen.imgApproveTickInd);
			DisableButtons();
		}
		else ChangeTickStatus(frmScreen.imgApproveTickInd,frmScreen.imgApproveBlankInd);
	}
	
	if(m_sStageId == m_sDeclineStageId) //Decline
	{
		ChangeTickStatus(frmScreen.imgDeclineBlankInd,frmScreen.imgDeclineTickInd);
		DisableButtons();
	}
	if( m_sTaskXML.length == 0 || m_sReadOnly == "1")
		DisableButtons();
		
	if(m_bButtonsDisabled == false)
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window , "GetCurrentStage");
		XML.CreateActiveTag("CASEACTIVITY");
		XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		XML.SetAttribute("CASEID", m_sApplicationNumber);
		XML.SetAttribute("ACTIVITYID", m_sActivityId);
		XML.SetAttribute("ACTIVITYINSTANCE", "1");
		// 		XML.RunASP(document, "MsgTMBO.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "MsgTMBO.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		if(XML.IsResponseOK())
		{
			var sTMRemodelMortgage = "";
			var sTMRiskAssess3 = "";
			var globalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			sTMRemodelMortgage = globalXML.GetGlobalParameterString(document, "TMRemodelMortgage");
			sTMRiskAssess3 = globalXML.GetGlobalParameterString(document, "TMRiskAssess");
			globalXML = null;
				
			XML.SelectTag(null,"RESPONSE");
			XML.CreateTagList("CASETASK");
			var bExit = false;
			for(var nTask = 0; nTask < XML.ActiveTagList.length && bExit == false; nTask++)
			{
				XML.SelectTagListItem(nTask);
				if(XML.GetAttribute("TASKSTATUS") == "10" || XML.GetAttribute("TASKSTATUS") == "20")
				{
					if(XML.GetAttribute("TASKID") == sTMRemodelMortgage)
					{
						alert("Application requires re-modelling. Data may be inconsistent therefore Recommendation and Approval cannot be given.");
						frmScreen.btnRecommend.disabled = true;
						frmScreen.btnApprove.disabled = true;
						bExit = true;
					}
					else if (XML.GetAttribute("TASKID") == sTMRiskAssess3)
					{
						alert("Application requires Risk Assessment3 to be performed. Recommendation and Approval cannot be given.");
						frmScreen.btnRecommend.disabled = true;
						frmScreen.btnApprove.disabled = true;
						bExit = true;
					}
				}
			}
		}
		<% /* MV BM0221 */ %>
		if (AppXML.GetTagText("ACCEPTEDQUOTENUMBER") != null && bExit == false )
		EnableApproveButton();
		<% /* MV BM0221 End */ %>
	}
}
function ChangeTickStatus(ImgToHide, ImgToShow)
{
	ImgToHide.style.visibility="hidden";
	ImgToShow.style.visibility="visible";
}
function DisableButtons()
{
	frmScreen.btnApprove.disabled = true;
	frmScreen.btnDecline.disabled = true;
	frmScreen.btnRecommend.disabled = true;
	m_bButtonsDisabled = true;
}
function frmScreen.btnAppDetails.onclick()
{
	frmToAP040.submit();
}

function AddReason(sReasonType, sAuthorisedUserId)
{
	var sReturn = null;
	var bReturn = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ArrayArguments = new Array();
	ArrayArguments[0] = XML.CreateRequestAttributeArray(window);
	ArrayArguments[1] = sReasonType;
	ArrayArguments[2] = m_sApplicationNumber;
	ArrayArguments[3] = m_sApplicationFactFindNumber;
	ArrayArguments[4] = m_sUserId;
	ArrayArguments[5] = sAuthorisedUserId;
	ArrayArguments[6] = m_sUnitId;
	ArrayArguments[7] = m_sTaskXML;
	ArrayArguments[8] = m_sAuthorityLevel;
	ArrayArguments[9] = m_sAppPriority;
	sReturn = scScreenFunctions.DisplayPopup(window, document, "AP014.asp", ArrayArguments, 540, 240);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		bReturn = true;
	}
	XML = null;
	return bReturn;
}
function CheckPassword()
{
	var sAuthorisedUserId = "";
	var sReturn = null;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ArrayArguments = new Array();
	ArrayArguments[0] = XML.CreateRequestAttributeArray(window);
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = m_sUnitId;
	ArrayArguments[3] = m_sUnitName;
	ArrayArguments[4] = m_sApplicationNumber;
	ArrayArguments[5] = m_sApplicationFactFindNumber;
	sReturn = scScreenFunctions.DisplayPopup(window, document, "AP012.asp", ArrayArguments, 335, 275);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		sAuthorisedUserId = sReturn[1];
	}
	XML = null;
	
	return(sAuthorisedUserId);
}
function frmScreen.btnApprove.onclick()
{
	var sAuthorisedUserId = CheckPassword();
	if(sAuthorisedUserId != "")
	{
		if(sAuthorisedUserId == m_sRecommendedUserId)
			alert("The same user who recommended the application cannot approve it.");
		else
		{
			if(AddReason("Approve", sAuthorisedUserId))
				Initialise();
		}
	}
}
function frmScreen.btnRecommend.onclick()
{
	var sAuthorisedUserId = CheckPassword();
	if(sAuthorisedUserId != "")
	{
		if(AddReason("Recommend", sAuthorisedUserId))
			Initialise();
	}
}
function frmScreen.btnDecline.onclick()
{
	if(confirm("Are you sure you want to decline the application?"))
	{
		var sAuthorisedUserId = CheckPassword();
		if(sAuthorisedUserId != "")
		{
			var sReason = "";
			var ArrayArguments = new Array();
			sReturn = scScreenFunctions.DisplayPopup(window, document, "TM038.asp", ArrayArguments, 550, 200);
			if (sReturn != null)
			{
				FlagChange(sReturn[0]);
				sReason = sReturn[1];
			}
			
			if(sReason != "")
			{
				<%/* BMIDS729 GHun 17/03/2004 */%>
				var sLocalPrinters = GetLocalPrinters();
				var LocalPrintersXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				LocalPrintersXML.LoadXML(sLocalPrinters);
				LocalPrintersXML.SelectTag(null, "RESPONSE");	
				var sPrinter = LocalPrintersXML.GetTagText("PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
				<%/* BMIDS729 End */%>
		
				m_taskXML.SelectTag(null, "CASETASK");
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				var tagReq = XML.CreateRequestTag(window, "DeclineApplication");
				XML.CreateActiveTag("APPLICATION");
				XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
				XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
				XML.SetAttribute("APPLICATIONPRIORITY", m_sAppPriority);
				XML.ActiveTag = tagReq;
				XML.CreateActiveTag("CASESTAGE");
				XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
				XML.SetAttribute("CASEID",m_sApplicationNumber);
				XML.SetAttribute("ACTIVITYID",m_taskXML.GetAttribute("ACTIVITYID"));
				XML.SetAttribute("ACTIVITYINSTANCE",m_taskXML.GetAttribute("ACTIVITYINSTANCE"));
				XML.SetAttribute("CASESTAGESEQUENCENO",m_taskXML.GetAttribute("CASESTAGESEQUENCENO"));
				XML.SetAttribute("STAGEID", m_sDeclineStageId); // Declined
				XML.SetAttribute("EXCEPTIONREASON",sReason);
				XML.ActiveTag = tagReq;
				XML.CreateActiveTag("CASETASK");
				XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
				XML.SetAttribute("CASEID",m_sApplicationNumber);
				XML.SetAttribute("ACTIVITYID",m_taskXML.GetAttribute("ACTIVITYID"));
				XML.SetAttribute("ACTIVITYINSTANCE",m_taskXML.GetAttribute("ACTIVITYINSTANCE"));
				XML.SetAttribute("CASESTAGESEQUENCENO",m_taskXML.GetAttribute("CASESTAGESEQUENCENO"));
				XML.SetAttribute("TASKID",m_taskXML.GetAttribute("TASKID"));
				XML.SetAttribute("TASKINSTANCE",m_taskXML.GetAttribute("TASKINSTANCE"));
				XML.SetAttribute("STAGEID",m_sStageId); // current stageId
				
				<%/* BMIDS729 GHun 17/03/2004 */%>
				if (sPrinter != "")
				{
					XML.ActiveTag = tagReq;
					XML.CreateActiveTag("PRINTER");
					XML.SetAttribute("PRINTERNAME", sPrinter);
					XML.SetAttribute("DEFAULTIND", "1");
				}
				<%/* BMIDS729 End */%>
				
				// 				XML.RunASP(document, "OmigaTMBO.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									<% /* BMIDS697 GHun 04/02/2004 
									XML.RunASP(document, "OmigaTMBO.asp");*/ %>
									XML.RunASP(document, "omTmNoTxBO.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
					}

				if(XML.IsResponseOK())
				{
				//BMIDS723 - DRC Don't set to read only
					// scScreenFunctions.SetContextParameter(window,"idReadOnly","1");
					SetStageInfo();
					//if declined set the flag
					if((scScreenFunctions.GetContextParameter(window,"idStageId",0) == scScreenFunctions.getCancelledStageValue(window))||(scScreenFunctions.GetContextParameter(window,"idStageId",0) == scScreenFunctions.getDeclinedStageValue(window)))
					{
						scScreenFunctions.SetContextParameter(window,"idCancelDeclineFreezeDataIndicator","1")
					}
					
					RetrieveContextData();
					Initialise();
				}
			}
		}
	}
}
function SetStageInfo()
{
	var stageXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	stageXML.CreateRequestTag(window , "GetCurrentStage");
	stageXML.CreateActiveTag("CASEACTIVITY");
	stageXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	stageXML.SetAttribute("CASEID", m_sApplicationNumber);
	stageXML.SetAttribute("ACTIVITYID", m_taskXML.GetAttribute("ACTIVITYID"));
	stageXML.SetAttribute("ACTIVITYINSTANCE", "1");

	// 	stageXML.RunASP(document, "MsgTMBO.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			stageXML.RunASP(document, "MsgTMBO.asp");
			break;
		default: // Error
			stageXML.SetErrorResponse();
		}

	if(stageXML.IsResponseOK())
	{
		stageXML.SelectTag(null, "CASESTAGE");
		var saOrigStageInfo = FW030GetStageInfo(stageXML.GetAttribute("STAGEID"));
		scScreenFunctions.SetContextParameter(window,"idStageId", stageXML.GetAttribute("STAGEID"));
		scScreenFunctions.SetContextParameter(window,"idStageName", saOrigStageInfo[0]);
		scScreenFunctions.SetContextParameter(window,"idCurrentStageOrigSeqNo", saOrigStageInfo[1]);
		// Having reset the stage, update the screen
		FW030SetTitles("Application/Approval Summary","AP010",scScreenFunctions);
	}
	stageXML = null;
}
function btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
	{
		if(m_sTaskXML.length == 0)frmToMN070.submit();
		else frmToTM030.submit();
	}
}

function EnableApproveButton()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var tagReq = XML.CreateRequestTag(window, "VALIDATEUSERLOGON");
	XML.CreateActiveTag("APPLICATION");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	
	switch (ScreenRules())
	{
		case 1: 
		case 0:
				XML.RunASP(document, "ValidateUserMandateLevel.asp");
				break;
		default:
				XML.SetErrorResponse();
	}

	if(XML.IsResponseOK())
	{
		XML.SelectTag(null, "USER");
		if(XML.GetAttribute("VALIDUSER") == "1")
		<% /* BS BM0187 14/02/03*/ %>
		{
			//frmScreen.btnApprove.disabled = false;
			GetValuerInstructions();
		}
		<% /* BS BM0187 End 14/02/03*/ %>
		else
			frmScreen.btnApprove.disabled = true;
	}
}
<% /* BS BM0187 14/02/03*/ %>
function GetValuerInstructions()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = XML.CreateRequestTag(window , "GetValuerInstructions");

	XML.CreateActiveTag("VALUERINSTRUCTIONS");
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	//XML.SetAttribute("INSTRUCTIONSEQUENCENO",m_sInsSeqNo);

	XML.RunASP(document, "omAppProc.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);

	if(ErrorReturn[0] == true)
	{
		XML.SelectTag(null, "VALUERINSTRUCTIONS");
		var sValuationStatus;
		var sValuationType;
		var sNPValuationType;
		var ValidationList = new Array(1);
		ValidationList[0] = "N";
		
		<% /* If the New Property valuation type has a ValidationType = 'N' then a 
		valuation report isn't needed, but if one has been requested then it must be 
		complete. If the ValidationType is not 'N' then the most recent valuation
		must be complete, if it isn't the disable the approve button. */ %>
		sNPValuationType = XML.GetAttribute("NEWPROPERTYVALUATIONTYPE");
		sValuationStatus = XML.GetAttribute("VALUATIONSTATUS");
		sValuationType = XML.GetAttribute("VALUATIONTYPE");
		
		if( XML.IsInComboValidationList(document,"ValuationType", sNPValuationType, ValidationList))
		{
			<% /* ValidationType is 'N' - if valuation requested then must be complete*/ %>
			if (sValuationType != "" && sValuationStatus != "30")
			{
				<% /* A valuer has been instructed but its not complete */ %>
				alert("A completed valuation report must be input before the case can be approved");
				frmScreen.btnApprove.disabled = true;
			}
		}
		else
		{
			<% /* ValidationType not 'N' */ %>
			if (sValuationStatus != "30")
			{
				<% /* Valuation not complete */ %>
				alert("A completed valuation report must be input before the case can be approved");
				frmScreen.btnApprove.disabled = true;
			}
		}
	}
	else if(ErrorReturn[1] == ErrorTypes[0])
	{
		//No valuer instructions found which is ok if the NewProperty valuation type 
		//validationid is 'N', else error

		var NPXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		NPXML.CreateRequestTag(window,null);
		NPXML.CreateActiveTag("NEWPROPERTY");
		NPXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		NPXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	
		switch (ScreenRules())
		{
			case 1: 
			case 0:
					NPXML.RunASP(document, "GetValuationTypeAndLocation.asp");
					break;
			default:
					NPXML.SetErrorResponse();
		}

		if(NPXML.IsResponseOK())
		{
			var sNPValuationType;
			var ValidationList = new Array(1);
			ValidationList[0] = "N";
		
			<% /* If the New Property valuation type has a ValidationType = 'N' then a 
			valuation report isn't needed, but if one has been requested then it must be 
			complete.  */ %>
			NPXML.SelectTag(null, "NEWPROPERTY");
			sNPValuationType = NPXML.GetTagText("VALUATIONTYPE");
		
			if(!NPXML.IsInComboValidationList(document,"ValuationType", sNPValuationType, ValidationList))
			{
				alert("A completed valuation report must be input before the case can be approved");
				frmScreen.btnApprove.disabled = true;
			}
		}
		else
		{
			alert("A completed valuation report must be input before the case can be approved");
			frmScreen.btnApprove.disabled = true;
		}
	}
}
<% /* BS BM0187 End 14/02/03*/ %>

/*
function window.onunload() 
{
	scScreenFunctions.SetContextParameter(window,"idTaskXML",null);
}
*/	

<%/* BMIDS729 GHun 17/03/2004 */%>
function GetLocalPrinters()
{
	var strOut = "";
	var objOmPC = new ActiveXObject("omPC.PCAttributesBO");
	if (objOmPC != null)
	{
		var strXML = "<?xml version='1.0'?><REQUEST ACTION='CREATE'></REQUEST>";
		strOut = objOmPC.FindLocalPrinterList(strXML);
		objOmPC = null;
	}
	return strOut;
}
<%/* BMIDS729 End */%>

-->
</script>
</body>
</html>
