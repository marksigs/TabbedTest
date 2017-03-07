<%@  Language=JavaScript %>
<% 
	/* PSC 16/10/2005 MAR57 - Start */
	var sStartScreen = Request.QueryString("Screen");
	/* PSC 16/10/2005 MAR57 - End */
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/* 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      fw020.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Screen navigation menu
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		25/02/00	Initial revision
AY		25/02/00	More options enabled
AY		29/02/00	Cost modelling options added
AY		01/03/00	SYS0400 - Group Connections added and New Loan and Property routing incorrect
AY		03/03/00	SYS0400 - Addition of Customer Registration section
					Inclusion of context variable settings
AY		13/03/00	Introduction of bookmarks
AY		15/03/00	SYS0242 - Introduction of Stage Checks
AY		17/03/00	MetaAction setting removed
IW		17/03/00	Added DC012 - Application Verification option
AY		20/03/00	Scrolling to top of screen prevented on Stage Check
IW		22/03/00	Added DC305 - Application Credit Card
IW		22/03/00	Added Navigation ToolTips
MC		23/03/00	Added General Heading with Credit Check Summary link
AY		04/03/00	APPLICATION MENU title moved to modal_OmigaMenu
AY		06/04/00	SYS0569 - New context field for stage information - names are
					no longer hard coded
					scScreenFunctions now used to set context fields
MC		10/04/2000	Remove General heading and Credit Check link as now accessed from 
					top menu/title bar.
IW		14/04/00	Added DC155 - Regular Outgoings		
IW		20/04/00	Added DC255 - New Property Details			
MC		27/04/00	Added DC295 - Other Insurance Company
BG		17/05/00	SYS0752 Removed Tooltips
MH      31/05/00    Scroll bar issues see below
GD		05/12/00	Added Application Processing-->Application Stage menu
GD		15/12/00	Added attribute XML to CheckStage(), incl. call to OmigaTMBO.asp
DJP		19/12/00	Added Application Processing-->Application Transfer menu
JLD		15/01/01	SYS1808 changed GUI context settings for stage information.
JLD		15/01/01	SYS1805 trap error 4802 and give friendly message.
ADP		13/02/01	SYS1944 Amend Application Processing menu options
JLD		05/03/01	SYS1879 Enable ApplicationSummary screen
MC		07/03/01	SYS2016 Enable Payee and Disbursement links
APS		12/03/01	SYS1920 Read Only Processing
IK		21/03/01	SYS1924 clear idTaskXML on move to stage
BG		12/04/01	SYS2096 Added additional attributes to call to Move to Stage
DRC     24/07/01    SYS2032 Change to CheckStage to prevent advancing beyond next stage
MC		03/10/01	SYS2767 Client specific customisation.
JLD		08/10/01	SYS2736 Don't download omPC again
DR		21/11/01	SYS3095 Add id to menu item to allow customisation for SYS2805/SYS2874
GD		04/12/01	SYS3335 Enable Expenses Link From Application Processing(Client MCAP - idNavLink9)
DR		12/12/01	SYS3464	Add further ids to menu item to allow MCAP customication for SYS3462
DR		03/01/02	SYS3561	Add another id to menu item to allow MCAP customication for SYS2874
MC		22/02/02	SYS4137 Allow client customisation of Application Verification menu item.
SG		05/07/02	SYS4442 Remove hardcoding of stageids
JLD		16/07/02	SYS5138 Add a checkstage case for "CR"
SG		19/07/02	SYS5176 Change SYS4442 forces user to do Quick Quote - make this optional
SG		23/07/02	SYS5213 Bug from SYS5176
SG		24/07/02	SYS5225 Bug from SYS5176
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		16/05/2002	BMIDS0008	Modified Division divAIPOptions for AIP
MV		17/05/2002	BMIDS0008	Modified Division divAIPOptions for AIP
MV		17/05/2002	BMIDS0008	Modified Division divAIPOptions for AIP Movd CCJ History 
GD		19/06/2002	BMIDS00077  Upgrade to Core 7.0.2
MO		02/07/2002	BMIDS00090	Changed 'Customer Registration' to 'Customer Search'
DPF		04/07/2002	BMIDS00138	SYS4968 greyout of QuickQuote changed from 40 (CostModel) to 30 (AiP)
MV		08/07/2002	BMIDS00188	Core Code Upgrade Error -  Modified Window.OnLoad()
PSC		05/08/2002	BMIDS00006	Change Mortgage Accounts to Existing Accounts
MO		01/10/2002	BMIDS00548	Intigrated core AQR's (SYS4442, SYS5138, SYS5176, SYS5213, SYS5225) 
								 which deal reading stage id's into BMIDS code, 
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115								 
MV		09/10/2002	BMIDS00556	Removed Validation_Init()
ASu		11/10/2002	BMIDS00610	Remove DC100.asp from the index (Not used for BMIDS)
DPF		23/10/2002	BMIDS00589	Removed references to Agreement In Principle & Mortgage Aplication with
								Data Capture 1 & 2.
MO		13/11/2002	BMIDS00723	Made changes after the font sizes were changed
MO		25/11/2002	BMIDS01076	Capture the automatic task failed error in MoveToStage, display the error 
								and then move on as normal
GHun	30/01/2004	BMIDS697	Call the correct version of MoveToStage
SR		20/05/2004  BMIDS772    Attitude to Borrowing (DC330.asp)is not required 
SR		19/06/2004  BMIDS772    Remove Loans/Liabilities, Arrears History, Bankruptcy Histor, CCJ History,
								Mortgages Declined and Bank/Credit cards
MC		22/06/2004	BMIDS772	CostModeling Submenu removed.
HMA     22/11/2004  BMIDS948    Enhancements.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
MF		28/07/2005	MAR19		Change nav bar text and routing.
MF		03/08/2005	MAR20		Parameterised routing in many places for Global parameters
								ThirdPartySummary and NewPropertySummary
AM		23/08/2005	MAR019		Calling the Decision Screen
PSC		12/10/2005	MAR57		Decrease size of divNavigation to 145px
MV		21/10/2005	MAR256		Amended the Quote Menu Node
PE		14/12/2005	MAR868		If the global parameter "NewPropertySummary" is 0, the screen flow is changed.
PJO     20/12/2005  MAR825      Show / hide other income on global parameter
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
SAB		03/04/2006	EP8			Base display of 'Application Verification' and 'Decision' 
								on values held in global variables
PB		23/05/2006	EP603		MAR1794 Disable KFI based on global parameter
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
</head>

<body class="navBarBackground" style="BACKGROUND-IMAGE: none">
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" viewastext tabindex="-1"></object>
<% /* XML Functions Object - see comments within scXMLFunctions.asp for details */ %>
<object data="scScreenFunctions.asp" height="1" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex="-1" type="text/x-scriptlet" width="1" viewastext></object>
<object data="scXMLFunctions.asp" height="1" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex="-1" type="text/x-scriptlet" width="1" viewastext></object>
<object data="scMathFunctions.asp" height="1" id="scMathFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex="-1" type="text/x-scriptlet" width="1" viewastext></object>
<% /* This div has been indented at 2 instead of 4 deliberately. The modal frameset 
   containing this needs to have a split of 152,648 since the RHS needs 648 exactly, not 647.
   With an indent of 4 then a minimum of 154 is required. This then does not fit on an 800
   screen - which is a design requirement. */ %> 

<div id="divNavigation" class="navOutline" style="LEFT: 2px; POSITION: absolute; WIDTH: 145px; TOP: 0px; VISIBILITY: hidden">
	<div class="navTitle" style="HEIGHT: 18px; LEFT: 4px; TEXT-INDENT: 4px; MARGIN-TOP: 2px">APPLICATION MENU</div>
	<input id="btnStageTrigger" type="button" style="display:none" onclick="CheckOptions()">
<%	/*	When adding any dropdown divs ensure that the id contains the word "Options"
	*/
%>	<a name="ToCR"></a>
	<div class="navHead" style="HEIGHT: 20px; WIDTH: 130px"><a href="#ToCR" onclick="DisplayOptions('divCROptions')" tabIndex=-1><img id="arrow_cr" src="images/arrow_right.gif" width="14" height="9" border="0">Customer Search</a>
		<div id="divCROptions" style="DISPLAY: none; LEFT: 5px; POSITION: relative">
			<div class="navLink" style="HEIGHT: 20px"><a href="#ToCR" onclick="Navigate('CR','CR030.ASP')" tabIndex=-1>Customer Application</a></div>
		</div>
	</div>

	<a name="ToQQ"></a>
	<div id="divQQHead" class="navHead" style="HEIGHT: 20px; WIDTH: 130px"><a href="#ToQQ" onclick="DisplayOptions('divQQOptions')" tabIndex=-1><img id="arrow_qq" src="images/arrow_right.gif" width="14" height="9" border="0">KFI</a>
		<div id="divQQOptions" style="DISPLAY: none; LEFT: 5px; POSITION: relative">
			<div class="navLink"><a href="#ToQQ" onclick="Navigate('QQ','MQ010.ASP')" tabIndex=-1>Basic Data Capture</a></div>
			<% /*<div class="navLink"><a href="#ToQQ" onclick="Navigate('QQ','DC330.ASP')" tabIndex=-1>Attitude To Borrowing</a></div> */ %>
			<div class="navLink" style="HEIGHT: 20px"><a href="#ToQQ" onclick="Navigate('QQ','CM065.ASP')" tabIndex=-1>Stored Quotes</a></div> 
			<% /* <div class="navLink" style="HEIGHT: 20px"><a href="#ToQQ" onclick="Navigate('QQ','CM010.ASP')" tabIndex=-1>Cost Modelling</a></div> */ %>
		</div>
	</div>
<!-- DPF 23/10/02:  Changed Agreement In Principle to Data Capture 1 -->
	<a name="ToAIP"></a>
	<div class="navHead" style="HEIGHT: 20px; WIDTH: 130px"><a href="#ToAIP" onclick="DisplayOptions('divAIPOptions')" tabIndex=-1><img id="arrow_aip" src="images/arrow_right.gif" width="14" height="9" border="0">Agreement in Principle</a>
		<div id="divAIPOptions" style="DISPLAY: none; LEFT: 5px; POSITION: relative">
			<div class="navLink" id="idNavLink12"><a href="#ToAIP" onclick="Navigate('AIP','DC012.ASP')" tabIndex=-1>Application Verification</a></div>
			<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC010.ASP')" tabIndex=-1>Application Source</a></div>
			<% /* <div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC020.ASP')" tabIndex=-1>Group Connections</a></div>*/ %>
			<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC030.ASP')" tabIndex=-1>Personal Details</a></div>
			<div class="navLink" id="idNavLink6"></div> <!-- Dependants (Core) -->
			<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC060.ASP')" tabIndex=-1>Address</a></div>
			<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC160.ASP')" tabIndex=-1>Employment</a></div>
			<% // PJO 20/12/2005 MAR825 give this div an ID %>
			<div id="OthInc" class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC195.ASP')" tabIndex=-1>Other Income</a></div>
			<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC085.ASP')" tabIndex=-1>Financial Summary</a></div>
			<% /* MF MAR19 New Decision screen */ %>
			<div id="idDecision" class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC370.ASP')" tabIndex=-1>Decision</a></div>
			<% /* PSC 05/08/2002 BMIDS00006 */ %>
			<% /* SR 21/05/2004 : BMIDS772 - remove Existing Accounts and Reular Outoings from menu. Add Financial Summay and Credit History instead */ %>
			<% /*<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC070.ASP')" tabIndex=-1>Existing Accounts</a></div>
			<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC155.ASP')" tabIndex=-1>Regular Outgoings</a></div> */ %>
			
			<% /* MF removed<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC110.ASP')" tabIndex=-1>Credit History</a></div>*/ %>
			<% /* SR 21/05/2004 : BMIDS772 - End*/ %>
			
			<% /* MF Removed for WP01 MAR019 <div class="navLink" style="HEIGHT: 20px"><a href="#ToAIP" onclick="Navigate('AIP','DC200.ASP')" tabIndex="-1">New Loan &amp; Property</a></div> */%>
			<div class="navLink" id="idNavLink11"></div>	
			<%/* SR 19/06/2004 : BMIDS772 - remove Loans/Liabilities,Arrears History,Bankruptcy Histor,CCJ History,Mortgages Declined
			<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC090.ASP')" tabIndex=-1>Loans/Liabilities</a></div>
			<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC130.ASP')" tabIndex=-1>Arrears History</a></div>
			<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC140.ASP')" tabIndex=-1>Bankruptcy History</a></div>
			<div class="navLink" id="idNavLink7"></div> <!-- CCJ History (Core) -->
			<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC120.ASP')" tabIndex=-1>Mortgages Declined</a></div> 
			*/%>
			<% /* ASu BMIDS00610 - Start 
			<div class="navLink"><a href="#ToAIP" onclick="Navigate('AIP','DC100.ASP')" tabIndex=-1>Mortgage Life Products</a></div>
			ASu BMIDS00610 - End */ %>			
			<% /* SR 19/06/2004 : BMIDS772 - remove Bank/Credit Cards
			<div class="navLink" style="HEIGHT: 20px" id="idNavLink8"><a href="#ToAIP" onclick="Navigate('AIP','DC080.ASP')" tabIndex=-1>Bank/Credit Cards</a></div>
			*/ %>
		</div>
	</div>

	<a name="ToCM"></a>
	<%/*=[MC]
		CostModeling SubDIV elements removed. all sub menu items have been removed.
		Header navigation option will route to CM010.
	*/%>
	<% /* MF MAR019 quote moved into Application
	<div class="navHead" style="HEIGHT: 20px; LEFT: 14px; POSITION: relative; WIDTH: 130px"><a href="#ToCM" onclick="Navigate('CM','CM010.ASP')" tabIndex=-1>Quote</a>
	<!--<div class="navHead" style="HEIGHT: 20px; WIDTH: 130px"><a href="#ToCM" onclick="DisplayOptions('divCMOptions')" tabIndex=-1>Cost Modelling</a>-->
		<div id="divCMOptions" style="DISPLAY: none; LEFT: 5px; POSITION: relative">
			<!-- <div class="navDisabledLink">Credit Check</div> -->
*/ %>			
			<% /* SR 20/05/2004 : BMIDS772 - Attitude to Borrowing (DC330.asp)is not required  */ %> 
<% /* MF MAR019 quote moved into Application
		<!--<div class="navLink"><a href="#ToCM" onclick="Navigate('CM','DC330.ASP')" tabIndex=-1 >Attitude To Borrowing</a></div> -->
		<!--<div class="navLink" style="HEIGHT: 20px"><a href="#ToCM" onclick="Navigate('CM','CM010.ASP')" tabIndex=-1 >Cost Modelling</a></div>-->
		</div>
	</div>
	*/ %>
	
<!-- DPF 23/10/02:  Changed Mortgage Application to Data Capture 2 -->
	<a name="ToMA"></a>
	<div class="navHead" style="HEIGHT: 20px; WIDTH: 130px"><a href="#ToMA" onclick="DisplayOptions('divMAOptions')" tabIndex=-1><img id="arrow_ma" src="images/arrow_right.gif" WIDTH="14" HEIGHT="9" BORDER="0">Application</a>
		<div id="divMAOptions" style="DISPLAY: none; LEFT: 5px; POSITION: relative">
			<div class="navLink" id="idNavLink200"></div> <!-- New Loan & Property (MARS) -->
			<div class="navLink"><a href="#ToMA" onclick="Navigate('MA','DC210.ASP')" tabIndex=-1>Property Address</a></div>
			<div class="navLink" id="idNavLink10"></div> <!-- Legal Address (MCAP) -->							
			<div class="navLink" id="idNavLink220"></div> <!-- New Property Description (Core) -->
			<div class="navLink" id="idNavLink201"></div> <!-- New Property Details (MARS) -->
			<div class="navLink" id="idNavLink225"></div> <!-- New Property Details (Core) -->
			<div class="navLink" id="idNavLink230"></div> <!-- Other Residents Summary (Core) -->
			<div class="navLink" id="idNavLink240"></div> <!-- Third Party Summary (Core) -->
			<div class="navLink" id="idNavLink280"></div> <!-- Legal Representative (MARS) -->
			<div class="navLink" id="idNavLink270"></div> <!-- Bank Details (MARS) -->
			<div class="navLink" id="idNavLink295"></div> <!-- Other Insurance Co (Core) -->
			<div class="navLink" id="idNavLink300"></div> <!-- B & C Declaration (Core) -->			
		</div>
	</div>
	
	<a name="ToCM"></a>
	<div class="navHead" style="HEIGHT: 20px; WIDTH: 130px"><A href="#ToCM" onclick="Navigate('CM','CM010.ASP')" tabIndex=-1><img id="arrow_ap" src="images/arrow_right.gif" WIDTH="14" HEIGHT="9" BORDER="0">Quote</A>
	</div>

	<a name="ToAP"></a>
	<div class="navHead" style="HEIGHT: 20px; WIDTH: 144px"><a href="#ToAP" onclick="DisplayOptions('divAPOptions')" tabIndex=-1><img id="arrow_ap" src="images/arrow_right.gif" WIDTH="14" HEIGHT="9" BORDER="0">Application Processing</a>
		<div id="divAPOptions" style="DISPLAY: none; LEFT: 5px; POSITION: relative">
			<div class="navLink"><a href="#ToAP" onclick=NavigateAS() tabIndex=-1>Application Summary</a></div>	
			<div class="navLink"><a href="#ToAP" onclick=NavigateTM() tabIndex=-1>Application Stage</a></div>
			<div class="navLink"><a href="#ToAP" onclick=NavigateAT() tabIndex=-1>Application Transfer</a></div>
			<div class="navLink"><a href="#ToAP" onclick=NavigateAC() tabIndex=-1>Application Costs</a></div>
			<div class="navLink" id="idNavLink9"></div>
			<div class="navLink"><a href="#ToAP" onclick=NavigatePY() tabIndex=-1>Payees</a></div>
			<div class="navLink"><a href="#ToAP" onclick=NavigateDB() tabIndex=-1>Disbursements</a></div>
		</div>
	</div>
</div>

<!--  #include FILE="Customise/FW020Customise.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/fw020Attribs.asp" -->

<script language="JScript" type="text/javascript">
var scScreenFunctions;

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


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Customise();	//SYS2767

	scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	<!-- this window is used to cache objects which other frames use. The other frames are now ready to load -->
	<% /* MV - 08/07/2002 - BMIDS00188 - Core Upgrade Code Error */ %>
	<% /* PSC 16/10/2005 MAR57 */ %>
	top.frames[0].navigate("modal_omigamenu.asp?Screen=<%=sStartScreen%>");
	
	<% // PJO 20/12/2005 MAR825 - Show / hide other income on global parameter %>
	var GlobalParamXML = new scXMLFunctions.XMLObject()
	if (! GlobalParamXML.GetGlobalParameterBoolean(document,"OtherIncomeSummary"))
	{
		OthInc.style.display = "none";
	}

	<% /* SAB - 03/04/2006 - EP8 - Start */ %>
	if (GlobalParamXML.GetGlobalParameterBoolean(document,"HideApplicationVerification"))
	{
		idNavLink12.style.display = "none";
	}

	if (GlobalParamXML.GetGlobalParameterBoolean(document,"HideDecision"))
	{
		idDecision.style.display = "none";
	}
	<% /* SAB - 03/04/2006 - EP8 - End */ %>
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function DisplayOptions(sDiv)
{
	<% /* BMIDS948  Enhancements 
	if(divNavigation.all(sDiv).style.display != "none") divNavigation.all(sDiv).style.display = "none";
	else divNavigation.all(sDiv).style.display = "block"; */ %>
	
	var sBefore = divNavigation.all(sDiv).style.display;
	
	divCROptions.style.display = "none";
	divQQOptions.style.display = "none";
	divAIPOptions.style.display = "none";
	divAPOptions.style.display = "none";
	divMAOptions.style.display = "none";

	arrow_aip.src = "images/arrow_right.gif"
	arrow_cr.src = "images/arrow_right.gif"
	arrow_qq.src = "images/arrow_right.gif"
	arrow_ma.src = "images/arrow_right.gif"
	arrow_ap.src = "images/arrow_right.gif"
	
	if (sBefore == "none")
	{
		divNavigation.all(sDiv).style.display = "block";
		switch(sDiv)
		{	
			case "divCROptions":
				arrow_cr.src = "images/arrow_down.gif";
			break;
			case "divAIPOptions":
				arrow_aip.src = "images/arrow_down.gif";
			break;
			case "divMAOptions":
				arrow_ma.src = "images/arrow_down.gif";
			break;
			case "divAPOptions":
				arrow_ap.src = "images/arrow_down.gif";
			break;
			case "divQQOptions":
				arrow_qq.src = "images/arrow_down.gif";
			break;
		}						
	}	
}

function NavigateAT()
{
	parent.main.location = "AT010.asp";
}

function NavigatePP()
{
	parent.main.location = "MN070.asp";
}

function NavigateAC()
{
	parent.main.location = "PP010.asp";
}

function NavigateAS()
{
	scScreenFunctions.SetContextParameter(window,"idXML","");
	parent.main.location = "AP010.asp";
}

function NavigateTM()
{
	parent.main.location = "TM030.asp";
}

function NavigatePY()
{
	parent.main.location = "PP070.asp";
}

function NavigateDB()
{
	parent.main.location = "PP050.asp";
}

function Navigate(sArea,sScreen)
{
	var sNextScreen = sScreen;
	if(CheckStage(sArea))
	{
		<% /* PSC 12/08/2002 BMIDS00006 */ %>
		scScreenFunctions.SetContextParameter(window,"idAccountGuid","");
		scScreenFunctions.SetContextParameter(window,"idApplicationMode","");
		scScreenFunctions.SetContextParameter(window,"idMetaAction","");
		scScreenFunctions.SetContextParameter(window,"idXML","");
		scScreenFunctions.SetContextParameter(window,"idXML2","");
		scScreenFunctions.SetContextParameter(window,"idTaskXML","");
		scScreenFunctions.SetContextParameter(window,"idCustomerIndex","");

		switch(sArea)
		{
			case "QQ":
				scScreenFunctions.SetContextParameter(window,"idApplicationMode","Quick Quote");
			break;
			case "AIP":
				scScreenFunctions.SetContextParameter(window,"idApplicationMode","AiP");
			break;
			case "CM":
				scScreenFunctions.SetContextParameter(window,"idApplicationMode","Cost Modelling");
			break;
			case "MA":
				scScreenFunctions.SetContextParameter(window,"idApplicationMode","MortgageApp");
			break;
			default:
			break;
		}
		parent.main.location = sNextScreen;
	}
}

function divNavigation.onpropertychange()
{
<%	/*	If the visibility property on divNavigation has been changed to hidden, find all drop down
		divs and 'collapse' them
	*/
%>	if(divNavigation.style.visibility == "hidden")
		for(var i0 = 0;i0 < divNavigation.all.length;i0++)
		{
			var sId = divNavigation.all(i0).id;
			if(sId.indexOf("Options") != -1) 
			{
				divNavigation.all(i0).style.display = "none";
				switch(sId)
				{	
					case "divCROptions":
						arrow_cr.src = "images/arrow_right.gif";
					break;
					case "divAIPOptions":
						arrow_aip.src = "images/arrow_right.gif";
					break;
					case "divMAOptions":
						arrow_ma.src = "images/arrow_right.gif";
					break;
					case "divAPOptions":
						arrow_ap.src = "images/arrow_right.gif";
					break;
					case "divQQOptions":
						arrow_qq.src = "images/arrow_right.gif";
					break;
				}						
			}
		}
}
function GetStageInfo(sStageId)
{
<%	/* Returns an array holding the stagename in [0] and the stagesequencenumber in [1]
		from the STAGE table for the given stageid */
%>	var XML = new scXMLFunctions.XMLObject();
	var saReturnArray = new Array();
	XML.CreateRequestTag(window, "GetStageDetail");
	XML.CreateActiveTag("STAGE");
	XML.SetAttribute("STAGEID", sStageId);
	XML.RunASP(document, "MsgTMBO.asp");
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null, "STAGE");
		saReturnArray[0] = XML.GetAttribute("STAGENAME");
		saReturnArray[1] = XML.GetAttribute("STAGESEQUENCENO");
	}
	XML = null;
	return(saReturnArray);
}

function GetStageName(sStageId)
{
<%	/* returns the stage name for the stage id sent in */
%>	var XML = new scXMLFunctions.XMLObject();
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
	<%/* SG 19/07/02 SYS5176 Moved declaration to module level */%>
	<%/* SG 05/07/02 SYS4442 */%>
	<%/*var sStageName = "";
	var sParamStageID = "";
	var sParamStageSeq = "";
	var sNewStageSeq = "";*/%>
	<%/* SG 05/07/02 SYS4442 */%>
	<%/* SG 19/07/02 SYS5176 End */%>

	<%/* SG 23/07/02 SYS5213 
	var sNewStageId; */%>
	
	switch(sArea)
	{
		case "CR":
			<% /* We do not need to change stage so route directly to area  JLD SYS5138 */%>
			return true;
		break;
		case "QQ":
			<%/* SG 05/07/02 SYS4442 */%>
			//sNewStageId = "20";
			//sStageName = "Quick Quote";
			sParamStageID = "TMQQStageID";
			sParamStageSeq = "TMQuickQuoteStage";
		break;
		case "AIP":
			<%/* SG 05/07/02 SYS4442 */%>
			//sNewStageId = "30";
			//sStageName = "Agreement In Principle";
			sParamStageID = "TMAIPStageID";
			sParamStageSeq = "TMAIPStage";
		break;
		case "CM":
			<%/* SG 05/07/02 SYS4442 */%>
			//sNewStageId = "40";
			//sStageName = "Cost Modelling";
			sParamStageID = "TMCMStageID";
			sParamStageSeq = "TMCostModellingStage";
		break;
		case "MA":
			<%/* SG 05/07/02 SYS4442 */%>
			//sNewStageId = "50";
			//sStageName = "Mortgage Application";
			sParamStageID = "TMMortAppStageID";
			sParamStageSeq = "TMMortAppStage";
		break;
		default:
			alert("Invalid area string for CheckStage");
			return false;
		break;
	}
	
	<%/* SG 05/07/02 SYS4442 START */%>
	var GlobalParamXML = new scXMLFunctions.XMLObject()
	sNewStageId = GlobalParamXML.GetGlobalParameterString(document,sParamStageID);
	sStageName = GetStageName(sNewStageId);

	//Errors?
	if (GlobalParamXML.IsResponseOK == false)
		return false;
			
	//Blank parameter?
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
	
	<%	//	SG 05/07/02 SYS4442
		//	Note:	Spec uses a boolean variable MOVETONEXTSTAGE to determine
		//			whether we pass/fail the IF statements above.
		//			Not needed - if we're here we've passed all tests.
	%>
	
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
	<%/* SG 05/07/02 SYS4442 amend elseif
	//SYS2032 DRC
	else if (nCurrentStageSeq + 1 < nNewStageSeq )*/%>
	else if (nNewStageSeq > nCurrentStageSeq + 1)
	{
		//User is attempting to miss at least 1 stage
	
		// Find the next stage name
		var iNext = nCurrentStageSeq;
		var bNotFound = true;
		var sNextStageId;
		var sNextStageName;
		var saStageInfo;
	  
		while  (bNotFound && ( iNext < nNewStageSeq ))
		{
			iNext += 1;
			sNextStageId = iNext + "0";
			saStageInfo = GetStageInfo(sNextStageId);
			if (iNext = parseInt(saStageInfo[1]))
			{
				<%/* SG 19/07/02 SYS5176 Don't set bNotFound if next stage is Quick Quote.*/%>
				sNextStageName = saStageInfo[0];
				if (sNextStageName != "Quick Quote")
					bNotFound = false;
			}																
		}			

		<%/* SG 19/07/02 SYS5176 Start */%>					
		if (iNext == nNewStageSeq)
		{
			//User selected the next stage after Quick Quote - let them move there.
			<%/* SG 24/07/02 SYS5225 Pass back result of MoveToStage() */%>
			if (MoveToStage())
				return true;
			else
				return false;
		}
		<%/* SG 19/07/02 SYS5176 End */%>

	 	if (bNotFound)
	 	{
		 	 sNextStageName = "next stage";
	 	} 

	 
	 	alert("Sorry, you must do Quick Quote or " + sNextStageName + " first");
		return false;
	}
    //SYS2032 DRC
    <%/* SG 05/07/02 SYS4442 Added Else */%>
	else
		return true;

}

function CheckOptions()
{
	<% /* PB 23/05/2006 EP603 */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var AIPStage = XML.GetGlobalParameterString(document, "TMAIPStageId");
	<% /* EP603 - End */ %>

	// BMIDS00138 DPF 04/07/02 changed from 40 (CostModel) to 30 (AiP)
	<% /* PB 23/05/2006 EP603
	if(scScreenFunctions.GetContextParameter(window,"idStageId",0) < 30)
	*/ %>
	if(parseInt(scScreenFunctions.GetContextParameter(window,"idStageId",0)) < parseInt(AIPStage))
	<% /* EP603 End */ %>
		divNavigation.all("divQQHead").style.display = "block";
	else
		divNavigation.all("divQQHead").style.display = "none";
}

<%/* SG	19/07/02 SYS5176 Existing code moved to this function */%>
function MoveToStage()
{
<%  // SG 05/07/02 SYS4442 No longer neeed this code				
	//	//var sName = scScreenFunctions.GetNameForStageNumber(window,sNumber);
	//	var saStageInfo = FW030GetStageInfo(sNewStageId);
	//	var nCurrentStageSeq = parseInt(scScreenFunctions.GetContextParameter(window,"idCurrentStageOrigSeqNo",null));
	//	//SYS2032 DRC - cannot progress beyond next stage
	//	var nNewStageSeq = parseInt(saStageInfo[1]);
	//	
	//	if ((nCurrentStageSeq + 1 == nNewStageSeq) ||
	//	    ((sArea == "AIP") && (nCurrentStageSeq == 1)) // can jump over QQ
	//	   )
	//	//SYS2032 DRC
	//	{
%>
<%/* SG 05/07/02 SYS4442 END */%>	
	
	if (scScreenFunctions.IsMainScreenReadOnly(window, "") == true) 
	{
		alert("You cannot advance to the next stage of a Read Only application.");
		return false;
	}

	<%/* SG 05/07/02 SYS4442 amend string concatenation */%>	
	if (!confirm("Do you wish to progress the application to the " + sStageName + " stage?"))
		return false;

	var sLocalPrinters = GetLocalPrinters();
	LocalPrintersXML = new scXMLFunctions.XMLObject();
	LocalPrintersXML.LoadXML(sLocalPrinters);
	LocalPrintersXML.SelectTag(null, "RESPONSE");	
		
	var sPrinter = LocalPrintersXML.GetTagText("PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
								
	if(sPrinter == "")
	{					
		alert("You do not have a default printer set on your PC.");	
		return false;		
	}		
		
	var XML = new scXMLFunctions.XMLObject()
		
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
	
	<% /* BMIDS697 GHun 30/01/2004 Call omTmNoTxBO.MoveToStage
	XML.RunASP(document, "OmigaTMBO.asp");//added 20.12.2000 */ %>
	XML.RunASP(document, "omTmNoTxBO.asp");
	
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
		//var sStage = new Array(sNumber,null);

		<%/* SG 05/07/02 SYS4442
		scScreenFunctions.SetContextParameter(window,"idStageName",saStageInfo[0]);
		scScreenFunctions.SetContextParameter(window,"idCurrentStageOrigSeqNo",saStageInfo[1]);*/%>
			
		scScreenFunctions.SetContextParameter(window,"idStageName",sStageName);
		scScreenFunctions.SetContextParameter(window,"idCurrentStageOrigSeqNo",sNewStageSeq);
		scScreenFunctions.SetContextParameter(window,"idStageId",sNewStageId);
			
		// AQR SYS4530 - Check whether a new AdminSys account number was added in an Automatic task
		<%/* SG 06/06/02 SYS4822 */%>
		<%/* scScrFunctions.SetOtherSystemACNoInContext(sApplicationNumber); */%>
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

</script>
</body>
</html>
