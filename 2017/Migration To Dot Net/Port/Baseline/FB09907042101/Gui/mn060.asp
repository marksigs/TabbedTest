<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      mn060.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		14/02/00	msgButtons removed
AY		01/03/00	SYS0060 - Application Processing changed to Cost Modelling
					(although graphics still outstanding)
					ApplicationMode and MetaAction context fields set up
					Routings modified
AY		10/03/00	SYS0242 - Introduction of stage checks
					SYS0060 - Cost Modelling button images included
					Button sizes were not consistent
AY		17/03/00	MetaAction settings removed
MC		23/03/00	SYS0428 - Enable tabbing between image buttons
AY		03/04/00	New top menu/scScreenFunctions change
AY		06/04/00	SYS0569 - New context field for stage information - names are
					no longer hard coded
					
GD		20/12/00	Added attribute XML to CheckStage(), incl. call to OmigaTMBO.asp
JLD		10/01/01	SYS1805: trap error 5003 and give friendly message.
JLD		15/01/01	SYS1808 changed GUI context settings for stage information.
AP		15/02/01	SYS1955 changed button images and reset positioning.	
CL		05/03/01	SYS1920 Read only functionality added
BG		26/04/01	SYS2096	Added extra attributes to MovetoStage call for printing
DRC     24/07/01    SYS2032 Change to CheckStage to prevent advancing beyond next stage
JLD		08/10/01	SYS2736 Don't download omPC again
SG		06/12/01	SYS3357 Make screen read only if application is Cancelled or Declined.
DRC     03/05/02    SYS4530 Check whether a new AdminSys account number was added in an Automatic task
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		06/06/02	SYS4822 Error from SYS4727
DRC		27/06/02    SYS4968 - changed greyout of Quick Quote from 40 (CostModel) to 30 (AiP)
SG		05/07/02	SYS4442 Remove hardcoding of stageids
SG		19/07/02	SYS5176 Change SYS4442 forces user to do Quick Quote - make this optional
SG		23/07/02	SYS5213 Bug from SYS5176
SG		24/07/02	SYS5225 Bug from SYS5176
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MO		01/10/2002	BMIDS00548	Took a new version of the file from Core and integrated it into BMIDS Code
DPF		25/10/2002	BMIDS00589	The images called from this screen have been amended
SA		29/10/2002	BMIDS00686	Make screen editable if application stage is Cancelled/Declined.
MO		25/11/2002	BMIDS01076	Capture the automatic task failed error in MoveToStage, display the error 
								and then move on as normal
GHun	30/01/2004	BMIDS697	Call the correct version of MoveToStage
HMA     23/03/2004  BMIDS731    Refresh Stage Name correctly after MoveToStage
MC		30/04/2004	BM0468		Cancel and Decline stage freeze screens functionality
SR		20/05/2004  BMIDS772    Cost Modelling routes to CM010 instead of DC330.asp
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		ID			Description
MF		22/07/2005				Process flow changes IA_WP01 DIP first screen is App Verification		
MF		28/07/2005				Quick quote (renamed as KFI) starts with MQ010 - Basic Data Capture
AM		26/09/2005				KFI is enabled in all stages. 
PE		14/12/05	MAR868		If the global paramater "NewPropertySummary" is 0, the screen flow is changed.
PJO     08/03/2006  MAR1378     Not all screens frozen properly for declined cases
DRC		30/03/2006  MAR1550     Enable the KFI only before stage 30 (reverses change of 26/09/2005)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		ID			Description

SAB		03/04/2006	EP8			Made AiP button routing conditional on HideApplicationVerification global param
PB		23/05/2006	EP603		MAR1794 Disable KFI based on global parameter
SR		16/02/2007	EP2_1099	modified logic for building the message to be displayed when buttons are clicked.
AShaw	19/03/2007	EP2_1592	Different routing from KFI stage for PSW case.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
<link href="stylesheet.css" rel="STYLESHEET" type="text/css"><title></title>
</head>
<body><!-- Form Manager Object - Controls Soft Coded Field Attributes -->

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>
<!-- Specify Forms Here -->
<form id="frmQuickQuote" method="post" action="mq010.asp" STYLE="DISPLAY: none"></form>
<% //MF 22/07/2005 IA_WP01 process flow changes: start at dc012.asp %>
<form id="frmAiP" method="post" action="dc012.asp" STYLE="DISPLAY: none"></form>
<% // SAB 03/04/2006 EP8 Process flow changes: start at dc010.asp %>
<form id="frmAiPAlt" method="post" action="dc010.asp" STYLE="DISPLAY: none"></form>

<form id="frmCompleteChecks" method="post" action="dc300.asp" STYLE="DISPLAY: none"></form>
<form id="frmApplicationCosts" method="post" action="mn070.asp" STYLE="DISPLAY: none"></form>
<% /*APS 01/06/00 - Change action of frmCostModelling to route to dc330.asp because gn300.asp has
 not been developed yet!! */ %>
 <% /* SR 20/05/2004 : BMIDS772 - Cost Modelling routes to CM010 instead of DC330.asp  */ %>
<form id="frmCostModelling" method="post" action="CM010.asp" STYLE="DISPLAY: none"></form>
<% /* SR 20/05/2004 : BMIDS772 - End */ %>
<form id="frmMortApp" method="post" action="dc210.asp" STYLE="DISPLAY: none"></form>
<form id="frmDC200" method="post" action="dc200.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" validate ="onchange" mark>
<div id="divBackground" style="HEIGHT: 440px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 0px; POSITION: absolute; TOP: 40px" class="msgLabel">				
		<span id="spnQuickQuote" style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">						
			<input id="imgQuickQuote1" type="button" style="BACKGROUND-IMAGE: url(images/kfi1.jpg); HEIGHT: 100px; LEFT: 100px; POSITION: absolute; TOP: 10px; WIDTH: 200px; Z-INDEX: 2" class="msgMainButton">
			<IMG alt="" border =0 height=100 id=imgQuickQuote2 src="images/kfi2.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 1" width=200 >
		</span>
		<span id="spnAgreementInPrinciple" style="LEFT: 220px; POSITION: absolute; TOP: 10px" class="msgLabel">						
			<input id="imgAgreementInPrinciple1" type="button" style="BACKGROUND-IMAGE: url(images/AgreementinPrinciple_2.jpg); HEIGHT: 100px; LEFT: 100px; POSITION: absolute; TOP: 10px; WIDTH: 200px; Z-INDEX: 2" class="msgMainButton">
			<IMG alt="" border =0 height=100 id=imgAgreementInPrinciple2 src="images/AgreementinPrinciple_3.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 1" width=200 >
		</span>
		<span id="spnCostModelling" style="LEFT: 10px; POSITION: absolute; TOP: 120px" class="msgLabel">						
			<input id="imgCostModelling1" type="button" style="BACKGROUND-IMAGE: url(images/cost_model_1.jpg); HEIGHT: 100px; LEFT: 100px; POSITION: absolute; TOP: 10px; WIDTH: 200px; Z-INDEX: 2" class="msgMainButton">
			<IMG alt="" border =0 height=100 id=imgCostModelling2 src="images/cost_model_2.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 1" width=200 >
		</span>
		<span id="spnMortgageApplication" style="LEFT: 220px; POSITION: absolute; TOP: 120px" class="msgLabel">						
			<input id="imgMortgageApplication1" type="button" style="BACKGROUND-IMAGE: url(images/mortgage_app_1.jpg); HEIGHT: 100px; LEFT: 100px; POSITION: absolute; TOP: 10px; WIDTH: 200px; Z-INDEX: 2" class="msgMainButton">
			<IMG alt="" border =0 height=100 id=imgMortgageApplication2 src="images/mortgage_app_2.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 1" width=200 >
		</span>
		<span id="spnApplicationCosts" style="LEFT: 115px; POSITION: absolute; TOP: 230px" class="msgLabel">						
			<input id="imgApplicationCosts1" type="button" style="BACKGROUND-IMAGE: url(images/app_proc_1.jpg); HEIGHT: 100px; LEFT: 100px; POSITION: absolute; TOP: 10px; WIDTH: 200px; Z-INDEX: 2" class="msgMainButton">
			<IMG alt="" border =0 height=100 id=imgApplicationCosts2 src="images/app_proc_2.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 1" width=200 >
		</span>

	</span>
</div>				
</form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- Specify Code Here -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/mn060Attribs.asp" -->
<script language="JScript">
<!--
var scScreenFunctions;
var m_blnReadOnly = false;
var m_blnNoQuickQuote = false;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";

<%/* SG 19/07/02 SYS5176 START */%>
<%/* Variables moved from CheckStage so can be referenced in MoveToStage */%>
<%/* Did not bother to rename "m_" - sorry */%>
var sStageName = "";
var sParamStageID = "";
var sParamStageSeq = "";
var sNewStageSeq = "";
var nCurrentStageSeq = 0;
var nNewStageSeq = 0;
var nNewStageID = 0;
<%/* SG 19/07/02 SYS5176 END */%>

<%/* SG 23/07/02 SYS5213 Moved declaration here */%>
var sNewStageId;

<% /* EP2_1592 Identify whether this is a PSW case*/ %>
var sAppType = "";


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	FW030SetTitles("Application Menu","MN060",scScreenFunctions);
	// AQR SYS4968 DRC 27/06/02 - changed greyout of Quick Quote from 40 (CostModel) to 30 (AiP)
	
	<% /*	DRC MAR1550 enable the KFI only before stage 30   */%>
	<% /* PB 23/05/2006 EP603 / MAR1794 */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var AIPStage = XML.GetGlobalParameterString(document, "TMAIPStageId");
	
	if(parseInt(scScreenFunctions.GetContextParameter(window,"idStageId",0)) >= parseInt(AIPStage))
	<% /* EP603 End */ %>
	{	
		m_blnNoQuickQuote = true;
	
		frmScreen.imgQuickQuote1.style.backgroundImage="url(images/kfi_disabled.jpg)";
		frmScreen.imgQuickQuote2.src="images/kfi_disabled.jpg";
	} 
	<% /*	DRC MAR1550 end  */%>
	<% /* PJO 08/03/2006 MAR1378 - don't use hard coded values
	if((scScreenFunctions.GetContextParameter(window,"idStageId",0) == 910)||(scScreenFunctions.GetContextParameter(window,"idStageId",0) == 920)) */ %>
	var iStageID ;
	iStageID = scScreenFunctions.GetContextParameter(window,"idStageId",0) ;
	if((iStageID == scScreenFunctions.getCancelledStageValue(window))||
	   (iStageID == scScreenFunctions.getDeclinedStageValue(window)))
	{
		scScreenFunctions.SetContextParameter(window,"idCancelDeclineFreezeDataIndicator","1")
	}
	
	<%/* Added by automated update TW 09 Oct 2002 SYS5115 */%>
	SetMasks();
	Validation_Init();
	scScreenFunctions.SetContextParameter(window, "idApplicationMode", "");
	scScreenFunctions.SetContextParameter(window, "idMetaAction", "");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "MN060");
	
	//SG 5/12/01 SYS3357
	<%/* BMIDS00686 Make screen editable if cancelled/declined
	if (m_blnReadOnly == false)
	{
		//If the Screen is not read only, check to see what the status of the Application is.
		//Cancelled or Declined applications should be read only.
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var CancelledStageValue = XML.GetGlobalParameterAmount(document, "CancelledStageValue");
		var DeclinedStageValue = XML.GetGlobalParameterAmount(document, "DeclinedStageValue");
		var IdStageID = scScreenFunctions.GetContextParameter(window,"idStageID",null);

		if (IdStageID == CancelledStageValue)
			{
				m_blnReadOnly = true;
				scScreenFunctions.SetScreenToReadOnly(frmScreen);
				scScreenFunctions.SetContextParameter(window,"idReadOnly","1");		
			}
			
		if (IdStageID == DeclinedStageValue)
			{
				m_blnReadOnly = true;
				scScreenFunctions.SetScreenToReadOnly(frmScreen);
				scScreenFunctions.SetContextParameter(window,"idReadOnly","1");		
			}
	}
	//END SG SYS3357
	SA end BMIDS00686 */%>
	
	// EP2_1592 - Get the Application Type.
	sAppType = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue",null);
	
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
function SwapImage(imgImage)
{
<% /* brings the image to the front */
%>	imgImage.style.cursor="hand";
	imgImage.style.zIndex="3";
}
function RestoreImage(imgImage)
{
<% /* moves the image to the back */
%>	imgImage.style.zIndex="1";
}	
// SYS0428 MDC 23/03/2000					
function frmScreen.imgQuickQuote1.onfocus()
{
	ResetImages();
	SwapImage(frmScreen.imgQuickQuote2);
}
function frmScreen.imgAgreementInPrinciple1.onfocus()
{
	ResetImages();
	SwapImage(frmScreen.imgAgreementInPrinciple2);
}
function frmScreen.imgCostModelling1.onfocus()
{
	ResetImages();
	SwapImage(frmScreen.imgCostModelling2);
}
function frmScreen.imgMortgageApplication1.onfocus()
{
	ResetImages();
	SwapImage(frmScreen.imgMortgageApplication2);
}
function frmScreen.imgApplicationCosts1.onfocus()
{
	ResetImages();
	SwapImage(frmScreen.imgApplicationCosts2);
}

function frmScreen.imgQuickQuote1.onclick()
{
	frmScreen.imgQuickQuote2.onclick();
}
function frmScreen.imgAgreementInPrinciple1.onclick()
{
	frmScreen.imgAgreementInPrinciple2.onclick();
}
function frmScreen.imgCostModelling1.onclick()
{
	frmScreen.imgCostModelling2.onclick();
}
function frmScreen.imgMortgageApplication1.onclick()
{
	frmScreen.imgMortgageApplication2.onclick();
}
function frmScreen.imgApplicationCosts1.onclick()
{
	frmScreen.imgApplicationCosts2.onclick();
}

function ResetImages()
{
	RestoreImage(frmScreen.imgQuickQuote2);
	RestoreImage(frmScreen.imgAgreementInPrinciple2);
	RestoreImage(frmScreen.imgCostModelling2);
	RestoreImage(frmScreen.imgMortgageApplication2);
	RestoreImage(frmScreen.imgApplicationCosts2);
}

function frmScreen.imgQuickQuote2.onclick()
{
	if(m_blnNoQuickQuote)
		return;
	if(CheckStage("QQ"))
	{
		scScreenFunctions.SetContextParameter(window, "idApplicationMode", "Quick Quote");
		frmQuickQuote.submit();
	}
}
function frmScreen.imgApplicationCosts2.onclick()
{
		scScreenFunctions.SetContextParameter(window, "idApplicationMode", "Application Costs");
		frmApplicationCosts.submit();
}

function frmScreen.imgQuickQuote1.onmouseover()
{
	ResetImages();
	SwapImage(frmScreen.imgQuickQuote2);
	frmScreen.imgQuickQuote1.focus();
}

function frmScreen.imgApplicationCosts1.onmouseover()
{
	ResetImages();
	SwapImage(frmScreen.imgApplicationCosts2);
	frmScreen.imgApplicationCosts1.focus();
}

function frmScreen.imgQuickQuote2.onmouseout()
{
	RestoreImage(frmScreen.imgQuickQuote2);
}

function frmScreen.imgApplicationCosts2.onmouseout()
{
	RestoreImage(frmScreen.imgApplicationCosts2);
}

function frmScreen.imgAgreementInPrinciple2.onclick()
{
	if(CheckStage("AIP"))
	{
		scScreenFunctions.SetContextParameter(window, "idApplicationMode", "AiP");
		<% /* EP8 - Change Routing */ %>
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if (XML.GetGlobalParameterBoolean(document, "HideApplicationVerification"))
			frmAiPAlt.submit();
		else
			frmAiP.submit();
	}
}
function frmScreen.imgAgreementInPrinciple1.onmouseover()
{
	ResetImages();
	SwapImage(frmScreen.imgAgreementInPrinciple2);
	frmScreen.imgAgreementInPrinciple1.focus();
}
function frmScreen.imgAgreementInPrinciple2.onmouseout()
{
	RestoreImage(frmScreen.imgAgreementInPrinciple2);
}
function frmScreen.imgCostModelling2.onclick()
{
	if(CheckStage("CM"))
	{
		scScreenFunctions.SetContextParameter(window, "idApplicationMode", "Cost Modelling");
		frmCostModelling.submit();
	}
}
function frmScreen.imgCostModelling1.onmouseover()
{
	ResetImages();
	SwapImage(frmScreen.imgCostModelling2);
	frmScreen.imgCostModelling1.focus();
}
function frmScreen.imgCostModelling2.onmouseout()
{
	RestoreImage(frmScreen.imgCostModelling2);
}
function frmScreen.imgMortgageApplication2.onclick()		
{
	if(CheckStage("MA"))
	{
		// MAR868
		// Peter Edney - 14/12/2005
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var bNewPropertySummary = XML.GetGlobalParameterBoolean(document, "NewPropertySummary");
					
		if (bNewPropertySummary)	
		{
			scScreenFunctions.SetContextParameter(window, "idApplicationMode", "MortgageApp");
			frmMortApp.submit();
		}
		else
		{
			frmDC200.submit();
		}
	}
}
function frmScreen.imgMortgageApplication1.onmouseover()
{
	ResetImages();
	SwapImage(frmScreen.imgMortgageApplication2);
	frmScreen.imgMortgageApplication1.focus();
}
function frmScreen.imgMortgageApplication2.onmouseout()
{
	RestoreImage(frmScreen.imgMortgageApplication2);
}

function GetStageInfo(sStageId)
{
<%	/* returns the stage name for the stage id sent in */
%>	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var saReturnArray = new Array();
	XML.CreateRequestTag(window, "GetStageDetail");
	XML.CreateActiveTag("STAGE");
	XML.SetAttribute("STAGEID", sStageId);
	XML.RunASP(document, "MsgTMBO.asp");
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null, "STAGE");
		sStageName = XML.GetAttribute("STAGENAME");
	}
	XML = null;
	return(sStageName);
}

function CheckStage(sArea)
{
   if (sAppType == "60")  // PSW
		switch(sArea) 
		{
			case "QQ":
				sParamStageID = "TMQQStageIDPSW";
				sParamStageSeq = "TMQuickQuoteStagePSW";
			break;
			case "AIP":
				sParamStageID = "TMAIPStageIDPSW";
				sParamStageSeq = "TMAIPStagePSW";
			break;
			case "CM":
				sParamStageID = "TMCMStageIDPSW";
				sParamStageSeq = "TMCostModellingStagePSW";
			break;
			case "MA":
				sParamStageID = "TMMortAppStageIDPSW";
				sParamStageSeq = "TMMortAppStagePSW";
			break;
			default:
				alert("Invalid area string for CheckStage");
				return false;
			break;
		}
    else

		switch(sArea)
		{
			case "QQ":
				sParamStageID = "TMQQStageID";
				sParamStageSeq = "TMQuickQuoteStage";
			break;
			case "AIP":
				sParamStageID = "TMAIPStageID";
				sParamStageSeq = "TMAIPStage";
			break;
			case "CM":
				sParamStageID = "TMCMStageID";
				sParamStageSeq = "TMCostModellingStage";
			break;
			case "MA":
				sParamStageID = "TMMortAppStageID";
				sParamStageSeq = "TMMortAppStage";
			break;
			default:
				alert("Invalid area string for CheckStage");
				return false;
			break;
		}
	
	
	
	<%/* SG 05/07/02 SYS4442 START */%>
	var GlobalParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	sNewStageId = GlobalParamXML.GetGlobalParameterString(document,sParamStageID);
	sStageName = GetStageInfo(sNewStageId);
	
	//Errors?
	if (GlobalParamXML.IsResponseOK == false)
		return false;
			
	if (sNewStageId == "")
	{
		//Yes - route to screen
		return true;		
	}
	else			
	{
		//No - get the stage sequence no
		sNewStageSeq = GlobalParamXML.GetGlobalParameterAmount(document,sParamStageSeq);
		
		//Error?
		if (GlobalParamXML.IsResponseOK == false)
			return false;
		
		//Blank parameter?
		if (sNewStageSeq == "")
		{
			alert("The parameter " + sParamStageSeq + " had not been defined within Supervisor. Please refer to your System Administrator");
			return false;
		}
	}	
	
	<%/* SG 19/07/02 SYS5176 Moved declaration to module level */%>
	//Get current stage from context
	nCurrentStageSeq = parseInt(scScreenFunctions.GetContextParameter(window,"idCurrentStageOrigSeqNo",null));
	nNewStageSeq = parseInt(sNewStageSeq);
	nNewStageID = parseInt(sNewStageId);
	<%/* SG 19/07/02 SYS5176 End */%>
	
	if (nNewStageSeq == nCurrentStageSeq + 1)
	{
		<%/* SG 24/07/02 SYS5225 Pass back result of MoveToStage() */%>
		if (MoveToStage())
			return true;
		else
			return false;
	}
	// PSW KFI -> PSW Offer Stage 
	else if ((sAppType == "60") && (nNewStageSeq == 4) && (nNewStageSeq == nCurrentStageSeq + 2))
	{
		if (MoveToStage())
			return true;
		else
			return false;
	}
	
	else if (nNewStageSeq > nCurrentStageSeq + 1)  // SR 16/02/2007 : EP2_1099
	{
		// Find the next stage name
		var iNext = nCurrentStageSeq;
		var bIsNextStageQQ = true;
		var sNextStageId;
		var sNextStageName = "";
		var saStageInfo;  

		sQQStageSeq = GlobalParamXML.GetGlobalParameterAmount(document,"TMQuickQuoteStage");	  
		while (bIsNextStageQQ && ( iNext < nNewStageSeq ))
		{
			iNext += 1;
			saStageInfo = GetStageBySeqNo(iNext);
			if(saStageInfo[1] != sQQStageSeq)
			{	// next stage is not QQ 
				bIsNextStageQQ = false; 
				if(sNextStageName == "") sNextStageName = saStageInfo[0];
			}
			else sNextStageName = saStageInfo[0];
		}  // SR 16/02/2007 : EP2_1099 - End
		
	 	if (bIsNextStageQQ)
	 	{
		 	 sNextStageName = "next stage";
	 	}  
	 
	 	alert("Sorry, you must do " + sNextStageName + " first");
		return false;
	}
	else return true;
}
<%/* SG	19/07/02 SYS5176 Existing code moved to this function */%>
function MoveToStage()
{	
	if (m_blnReadOnly == true) 
	{
		alert("You cannot advance to the next stage of a Read Only application.");
		return false;
	}
			
	<%/* SG 05/07/02 SYS4442 amend string concatenation */%>	
	if (!confirm("Do you wish to progress the application to the " + sStageName + " stage?"))
		return false;

	var sLocalPrinters = GetLocalPrinters();
	LocalPrintersXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	LocalPrintersXML.LoadXML(sLocalPrinters);
	LocalPrintersXML.SelectTag(null, "RESPONSE");	
		
	var sPrinter = LocalPrintersXML.GetTagText("PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
								
	if(sPrinter == "")
	{					
		alert("You do not have a default printer set on your PC.");	
		return false;		
	}		
		
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
	ReqTag = XML.CreateRequestTag(window, "MoveToStage");//added 20.12.2000
	XML.CreateActiveTag("CASESTAGE");//added 20.12.2000
	XML.SetAttribute("STAGEID", sNewStageId);//added 20.12.2000
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");//added 20.12.2000
	XML.SetAttribute("CASEID", scScreenFunctions.GetContextParameter(window,"idApplicationNumber",""));//added 20.12.2000
	XML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));//added 20.12.2000
	XML.SetAttribute("ACTIVITYINSTANCE", "1");//added 20.12.2000
	XML.SetAttribute("CASESTAGESEQUENCENUMBER", scScreenFunctions.GetContextParameter(window,"idCurrentStageOrigSeqNo",""));
			
	XML.ActiveTag = ReqTag;//added 20.12.2000
	XML.CreateActiveTag("APPLICATION");//added 20.12.2000
	var sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","")
	XML.SetAttribute("APPLICATIONNUMBER", sApplicationNumber);//added 20.12.2000
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",""));//added 20.12.2000
	XML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window,"idApplicationPriority",""));
		
	var iCount = 0;
	var sCustomerVersionNumber = "";
	var sCustomerNumber = "";
	for (iCount = 1; iCount <= 5; iCount++)
	{								
		sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
		if (sCustomerNumber != "")
		{	
			sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);	
			XML.SelectTag(null, "APPLICATION");
			XML.CreateActiveTag("CUSTOMER");				
			XML.SetAttribute("CUSTOMERNUMBER", sCustomerNumber);
			XML.SetAttribute("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);																		
		}
	}
			
	XML.ActiveTag = ReqTag;
	XML.CreateActiveTag("PRINTER");
	XML.SetAttribute("PRINTERNAME", sPrinter);
	XML.SetAttribute("DEFAULTIND", "1");	
		
	// 	XML.RunASP(document, "OmigaTMBO.asp");//added 20.12.2000
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			<% /* BMIDS697 GHun 30/01/2004 Call omTmNoTxBO.MoveToStage
			XML.RunASP(document, "OmigaTMBO.asp");//added 20.12.2000 */ %>
			XML.RunASP(document, "omTmNoTxBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}
	
	<% /* MO - 25/11/2002 - BMIDS01076 - Capture Automatted task errors */ %>
	var ErrorTypes = new Array("MANDTASKSOUTSTANDING", "AUTOTASKERROR");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if(ErrorReturn[1] == ErrorTypes[0])
		alert("There are mandatory tasks still outstanding on the current stage.  You must complete these before moving to another stage");
	else if((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[1]))
	{
		<% /* MO - 25/11/2002 - BMIDS01076 - An automatted task errored */ %>
		if (ErrorReturn[1] == ErrorTypes[1]) {
			<% /* Show the error message */ %>
			alert(ErrorReturn[2]);
		}

		ResetStage();
		
		scScreenFunctions.SetContextParameter(window,"idStageName",sStageName);
		scScreenFunctions.SetContextParameter(window,"idCurrentStageOrigSeqNo",sNewStageSeq);
		scScreenFunctions.SetContextParameter(window,"idStageId",sNewStageId);			
		scScreenFunctions.SetOtherSystemACNoInContext(sApplicationNumber);
			
		return true;
	}
	return false;
}

<% /* BMIDS697 Function converted from VBScript to JScript to improve performance */ %>
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
<% /* BMIDS697 End */ %>
<% /* BMIDS731 Start */ %>
function ResetStage()
{
	stageXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	stageXML.CreateRequestTag(window , "GetCurrentStage");
	stageXML.CreateActiveTag("CASEACTIVITY");
	stageXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	stageXML.SetAttribute("CASEID", scScreenFunctions.GetContextParameter(window,"idApplicationNumber",""));
	stageXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));
	stageXML.SetAttribute("ACTIVITYINSTANCE", "1");

	stageXML.RunASP(document, "MsgTMBO.asp");

	if(stageXML.IsResponseOK())
	{
		stageXML.SelectTag(null, "CASESTAGE");
		
		sNewStageId = stageXML.GetAttribute("STAGEID");
		
		var saOrigStageInfo = FW030GetStageInfo(sNewStageId);

		sStageName = saOrigStageInfo[0];
		sNewStageSeq = saOrigStageInfo[1]
	}
	return;
}
<% /* BMIDS731 End */ %>
// SR EP2_1099 - new function GetStageBySeqNo
function GetStageBySeqNo(sStageSeqNo)
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var saReturnArray = new Array();
	
	XML.CreateRequestTag(window, "GetStageList");
	XML.CreateActiveTag("STAGE");
	XML.SetAttribute("STAGESEQUENCENO", sStageSeqNo);
	XML.RunASP(document, "MsgTMBO.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);

	if(ErrorReturn[0] == true)
	{
		XML.SelectTag(null, "STAGE");
		saReturnArray[0] = XML.GetAttribute("STAGENAME");
		saReturnArray[1] = XML.GetAttribute("STAGESEQUENCENO");
	}
	XML = null;
	return(saReturnArray);
}
-->
</script>
</body>
</html>