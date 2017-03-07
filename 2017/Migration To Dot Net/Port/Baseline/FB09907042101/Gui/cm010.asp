<%@ LANGUAGE="JSCRIPT" %>
<% /*
Workfile:		cm010.asp
Copyright:		Copyright © 2000 Marlborough Stirling
Description:	Cost Modelling Menu Screen.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:
Prog	Date		Description
RF		28/01/00	Reworked for IE5 and performance.
AY		14/02/00	Change to msgButtons button types
AY		22/02/00	Initialisation and CreateNewQuotation functionality
AY		23/02/00	Affordability gauge changes, buttons reenabled on CreateNewQuote
AY		17/03/00	MetaAction changed to ApplicationMode
AY		22/03/00	Introduction of CM060/065
AY		29/03/00	New top menu/scScreenFunctions change
AY		11/04/00	SYS0328 - Dynamic currency display
AY		12/04/00	SYS0328 - £ signs missed in first change
APS		09/05/00	Calculate processing added
APS		30/05/00	SYS0753 - Disabled/Enable processing
APS		05/06/00	SYS0812 - Db operation failed on Calculate
APS		07/06/00	SYS0833 - Allow BC and PP Sub Quote only on valid Mortgage Sub Quote
DLM		10/07/00	SYS0949 - Not flag as an error a mortgage subquote with no loan components
DLM		08/08/00	SYS0887 - Added a method called on exit of the screen to update the required flags
BG		18/08/00	SYS0918 - Disable the majority of the screen if there is an invalid quote.
BG		05/09/00	SYS0964 - Show statement if payment frequency in BC quote is Annual.
SR		17/10/00	SYS0883	- Add BCSubQuoteNumber, PPSubQuoteNumber to the request for Affordability calculations
BG		08/12/00	SYS1637 - Made various changes to fix problems 1, 3 and 5.
CL		05/03/01	SYS1920 - Read only functionality added
CL		22/03/01	SYS1994	- Added facility to update case task
IK		22/03/01	SYS2145 - Add idNoCompleteCheck
APS		22/03/01	SYS1994	- Removed functionality for Demo
CL		23/03/01	SYS1994 - Changed UpdateStageSequence to call MsgTMBO.asp
CL		29/03/01	SYS2192 - Change to UpdateStageSequence
DC      10/08/01    SYS0977 - Change to PopulateScreen  
DC      09/11/01    SYS0977 commenetd out changes of 10/08/01    SYS0997 - Change to PopulateScreen
JLD		10/12/01	SYS2806 routing for completeness checking
PSC		13/12/01	SYS3425	Update Remodel Mortgage task on accept quote if present
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		28/05/2002  BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
GD		19/06/2002	BMIDS00077 - Upgrade to Core 7.0.2
DPF		04/07/2002	BMIDS00084 - Removal of Life Cover button and any mention of Life Cover from code
MO		11/07/2002	BMIDS00199 - Disabled Building and Contents and Payment Protection buttons for delivery 1, 
								 these will be re-enabled for CMWP4, to reenable the buttons remove the noted lines 
								 of code in functions EnableSubQuoteButton, and EnableCheckBox
								 
GD		16/07/2002	BMIDS00045 - Re-enable  Building and Contents and Payment Protection buttons for CMWP4.
								 Show different page title depending on context ie. 'Quick Quote Cost Modelling' or
								 'Application Cost Modelling'
GD		13/08/2002	BMIDS00320	 Default buttons to X's.
GD		15/08/2002	BMIDS00341	 Remove LifeCover references. Default all flags to FALSE, and set all values to zero.
GD		21/08/2002	BMIDS00312
GD		13/08/2002	BMIDS00320	 Check for "0" and ""	
TW      09/10/2002  Modified to incorporate client validation - SYS5115
MDC		31/10/2002	BMIDS00268	 Hide Affordability button and gauge as requested.
MDC		04/11/2002	BMIDS00654	 BM088 ICWP Income Calculations
MV		05/11/2002	BMIDS00797	Removed CalculateAffordability functionality 
GD		15/11/2002	BMIDS00376	Disable appropriate buttons if readonly
AW		18/11/2002	BMIDS00971	Reset accepted quote num
DPF		21/11/2002  BMIDS01047  In addition to changes for BMIDS00376 have made other buttons disabled
MDC		20/12/2002	BM0176		Use correct Mortgage Sub Quote Number in Quote Summary to get One Off Costs
MV		29/01/2003	BM0300	-	Modified PopulateScreen();
MV		03/03/2003	BM0404		Amended CompletenessCheckRouting()
DB		19/03/2003	BM0443		Stop CompletenessCheck running when screen is ReadOnly or Cancel clicked in CM100.
SR		01/06/2004	BMIDS772	
MC		22/06/2004	BMIDS772	DEFECT FIXES FOR BMIDS772 - ROUTING ETC (DC110,DC200,CM010)
JD		28/07/2004	BMIDS749	check for changed rates on entry
SR		31/08/2004  BMIDS815  
HMA     30/09/2004  BMIDS903    When LTV has changed, perform validation on Accept instead of at initialisation.
HMA     17/03/2005  BMIDS977    CC089  Route to PP010 on Accept Quote to display application fees.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR		Description
HMA     07/07/2005  MAR11   Remove redundant characters displayed on load
MV		19/09/2005	MAR44	Amended btnAccept.onclick()
MV		30/09/2005	MAR44  Code Review fixes
Maha T	21/11/2005	MAR424  On successfull CompleteInitialisation, set  the context paramter "idApplicationMode" to "Cost Modelling"
JD		27/01/2006	MAR1100	Removed max borrowing button
JD		10/02/2006	MAR1248 Changed comment marks from MAR1100 as mortgage button events were being commented out too
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR		Description
PB		12/05/2006	EP529	Merged MAR1704 - Always run Completeness check on page load.
PB		12/05/2006	EP529	Merged MAR1705 - added check for a modelled quote before checking if rates have changed
AShaw	05/02/2007	EP2_1164 Reset m_sApplicationMode after Completeness tests.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/ %>
<html>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% // Form Manager Object - Controls Soft Coded Field Attributes %>
<script src="validation.js" language="JScript"></script>
<!-- GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scMathFunctions.asp height=1 id=scMath style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>-->
<% // Specify Forms Here %>
<form id="frmApplicationMenu" method="post" action="mn060.asp" STYLE="DISPLAY: none"></form>
<form id="frmPaymentProtectionDetails" method="post" action="cm040.asp" STYLE="DISPLAY: none"></form>
<form id="frmBuildingAndContentsDetails" method="post" action="cm020.asp" STYLE="DISPLAY: none"></form>
<form id="frmMaximumBorrowing" method="post" action="cm055.asp" STYLE="DISPLAY: none"></form>
<form id="frmAffordability" method="post" action="cm050.asp" STYLE="DISPLAY: none"></form>
<form id="frmLoanComponents" method="post" action="cm100.asp" STYLE="DISPLAY: none"></form>
<% /* SR 21/05/2004 : BMIDS772 - remove routing to screen Attitude To Borrowing. instead route to MN060 */ %>
<form id="frmMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form> 
<form id="frmToDC200" method="post" action="DC200.asp" STYLE="DISPLAY: none"></form>
<%/* <form id="frmAttitudeToBorrowing" method="post" action="dc330.asp" STYLE="DISPLAY: none"></form> */ %>
<% /* SR 21/05/2004 : BMIDS772 - End */ %>
<form id="frmQuotationSummary" method="post" action="cm060.asp" STYLE="DISPLAY: none"></form>
<form id="frmStoredQuotes" method="post" action="cm065.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>
<% /* BMIDS977  Add routing to PP010 */ %>
<form id="frmToPP010" method="post" action="PP010.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% // Specify Screen Layout Here %>
<div id="divStatus" style="TOP: 60px; LEFT: 10px; WIDTH: 614px; HEIGHT: 100px; POSITION: absolute" class="msgGroup">
	<table id="tblStatus" width="604px" height="100px" border="0" cellspacing="0" cellpadding="0">
		<!--GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2--<tr align="center"><td id="colStatus" align="center" class="msgLabel" style="FONT-SIZE: 20px">Please Wait...</td></tr>-->
		<tr align="center"><td id="colStatus" align="center" class="msgLabelWait">Please Wait...</td></tr>
	</table>
</div>

<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 450px; WIDTH: 604px; POSITION: ABSOLUTE; VISIBILITY: HIDDEN" class="msgGroup">
<form id="frmScreen" mark validate="onchange" year4 style="VISIBILITY: hidden">
<%/* JD MAR1100 <span style="TOP: 5px; LEFT: 10px; POSITION: ABSOLUTE">
<input id="btnMaximumBorrowing" type="button" style="WIDTH: 65px; HEIGHT: 53px" class="msgMaxBorrowing">
</span>*/ %>
<span style="TOP: 5px; LEFT: 100px; POSITION: ABSOLUTE">
<% /* BMIDS00268 MDC 31/10/2002 - Hide Affordability button as requested 
<input id="btnAffordability" type="button" style="WIDTH: 65px; HEIGHT: 53px" class="msgAffordability"> */ %>
<input id="btnAffordability" type="button" style="WIDTH: 65px; HEIGHT: 53px; VISIBILITY: hidden" class="msgAffordability">
</span>
<span id="spnAffordabilityGaugeText" style="LEFT: 300px; POSITION: absolute; TOP: 17px; VISIBILITY: hidden" class="msgLabel">
<strong>Affordability <br>Gauge</strong>
</span>
<span id="spnAffordabilityGauge" style="LEFT: 380px; POSITION: absolute; TOP: 8px" class="msgLabel">
<span id="spnMininumAffordability" style="LEFT: 0px; POSITION: absolute; TOP: 0px; VISIBILITY:hidden" class="msgLabel"></span>
<span id="spnAffordabilityIncomeMarker" style="LEFT: 110px; POSITION: absolute; TOP: 0px; VISIBILITY:hidden" class="msgLabel"></span>
<span id="spnMaximumAffordability" style="LEFT: 190px; POSITION: absolute; TOP: 0px; VISIBILITY:hidden" class="msgLabel"></span>
<div id="divAffordabilityGauge" style="BACKGROUND: white; HEIGHT: 20px; LEFT: 0px; POSITION: absolute; TOP: 15px; WIDTH: 200px; VISIBILITY:hidden" class="msgBorder"></div>
<div id="divOutsideAffordability" style="BACKGROUND: red; HEIGHT: 10px; LEFT: 0px; POSITION: absolute; TOP: 17px; WIDTH: 0px; VISIBILITY:hidden" class="msgWhiteBorder"></div>
<div id="divInsideAffordability" style="BACKGROUND: blue; HEIGHT: 10px; LEFT: 0px; POSITION: absolute; TOP: 17px; WIDTH: 0px; VISIBILITY:hidden" class="msgWhiteBorder"></div>
<div id="divAffordabilityIncomeMarker"style="HEIGHT: 20px; LEFT: 0px; POSITION: absolute; TOP: 15px; WIDTH: 120px; VISIBILITY:hidden" class="msgBorder"></div>
</span>
</div>
<div style="HEIGHT: 300px; LEFT: 10px; POSITION: absolute; TOP: 125px; WIDTH: 475px" class="msgGroup">
<span id="spnMortgage" style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
<input id="btnMortgage" type="button" style="LEFT: 0px; TOP: 10px; HEIGHT: 52px; WIDTH: 254px; POSITION: absolute;" class="msgMortgage">
<span style="LEFT: 320px; POSITION: absolute; TOP: 27px">
<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
<input id="txtMonthlyMortgageCost" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 70px" class="msgTxt">
<span style="LEFT: 80px; POSITION: absolute; TOP: 0px; WIDTH: 60px" class="msgLabel">per month</span>
</span>
</span>

<% /* BMIDS00084 - DPF 04/07/02 - removal of Life Cover
<span id="spnLife" style="LEFT: 10px; POSITION: absolute; TOP: 70px" class="msgLabel">
<input id="btnLife" type="button" style="LEFT: 0px; TOP: 10px; HEIGHT: 52px; WIDTH: 254px; POSITION: absolute;" class="msgLife">
<!--GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2<input id="btnLifeRequired" type="button" style="LEFT: 265px; TOP: 20px; HEIGHT: 29px; WIDTH: 29px; POSITION: absolute;" >-->
<input id="btnLifeRequired" type="button" style="LEFT: 265px; TOP: 20px; HEIGHT: 29px; WIDTH: 29px; POSITION: absolute;" class="msgTick">
<span style="LEFT: 320px; POSITION: absolute; TOP: 27px">
<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
<input id="txtMonthlyLifeCost" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 70px" class="msgTxt">
<span style="LEFT: 80px; POSITION: absolute; TOP: 0px; WIDTH: 60px" class="msgLabel">per month</span>
</span>
</span>
*/ %>
<% //BMIDS00084 - have moved up screen by 60px because of Life Cover removal. (DPF 04/07/02) %>
<span id="spnBuildingsAndContents" style="LEFT: 10px; POSITION: absolute; TOP: 70px" class="msgLabel">

<!--BMIDS00320
<input id="btnBuildingsAndContents" type="button" style="LEFT: 0px; TOP: 10px; HEIGHT: 52px; WIDTH: 254px; POSITION: absolute;" class="msgBuildingsContents">
<input id="btnBuildingsAndContentsRequired" type="button" style="LEFT: 265px; TOP: 20px; HEIGHT: 29px; WIDTH: 29px; POSITION: absolute;" class="msgTick">
-->
<input id="btnBuildingsAndContents" type="button" style="LEFT: 0px; TOP: 10px; HEIGHT: 52px; WIDTH: 254px; POSITION: absolute;" class="msgBuildingsContentsDisabled" disabled>
<input id="btnBuildingsAndContentsRequired" type="button" style="LEFT: 265px; TOP: 20px; HEIGHT: 29px; WIDTH: 29px; POSITION: absolute;" class="msgCross">


<span style="LEFT: 320px; POSITION: absolute; TOP: 27px">
<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
<input id="txtMonthlyBuildingsContentsCost" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 70px" class="msgTxt">
<span style="LEFT: 80px; POSITION: absolute; TOP: 0px; WIDTH: 60px" class="msgLabel">per month</span>
</span>
</span>
<% //BMIDS00084 - have moved up screen by 60px because of Life Cover removal. (DPF 04/07/02) %>
<span id="spnPaymentProtection" style="LEFT: 10px; POSITION: absolute; TOP: 130px" class="msgLabel">

<!--BMIDS00320
<input id="btnPaymentProtection" type="button" style="LEFT: 0px; TOP: 10px; HEIGHT: 52px; WIDTH: 254px; POSITION: absolute;" class="msgPaymentProtection">
<input id="btnPaymentProtectionRequired" type="button" style="LEFT: 265px; TOP: 20px; HEIGHT: 29px; WIDTH: 29px; POSITION: absolute;" class="msgTick">
-->
<input id="btnPaymentProtection" type="button" style="LEFT: 0px; TOP: 10px; HEIGHT: 52px; WIDTH: 254px; POSITION: absolute;" class="msgPaymentProtectionDisabled" disabled>
<input id="btnPaymentProtectionRequired" type="button" style="LEFT: 265px; TOP: 20px; HEIGHT: 29px; WIDTH: 29px; POSITION: absolute;" class="msgCross">



<span style="LEFT: 320px; POSITION: absolute; TOP: 27px">
<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
<input id="txtMonthlyPaymentProtectionCost" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 70px" class="msgTxt">
<span style="LEFT: 80px; POSITION: absolute; TOP: 0px; WIDTH: 60px" class="msgLabel">per month</span>
</span>
</span>
<% //BMIDS00084 - have moved up screen by 60px because of Life Cover removal. (DPF 04/07/02) %>
<span style="LEFT: 200px; POSITION: absolute; TOP: 215px">
<span style="LEFT: 0px; POSITION: absolute; TOP: 0px" class="msgLabel">
Total Monthly Cost
<span style="LEFT: 130px; POSITION: absolute; TOP: 0px">
<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
<input id="txtTotalMonthlyCost" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 70px" class="msgTxt">
<span style="LEFT: 75px; POSITION: absolute; TOP: -5px">
<input id="btnCalculate" value="Calculate" type="button" style="WIDTH: 60px" class="msgButton">
</span></span></span></span>
</div>
<div style="HEIGHT: 300px; LEFT: 490px; POSITION: absolute; TOP: 125px; WIDTH: 124px" class="msgGroup">
<span style="LEFT: 10px; POSITION: absolute; TOP: 40px">
<input id="btnRecommend" value="Recommend Quote" type="button" style="WIDTH: 100px" class="msgButton">
</span>
<span style="LEFT: 10px; POSITION: absolute; TOP: 70px">
<input id="btnAccept" value="Accept Quote" type="button" style="WIDTH: 100px" class="msgButton">
</span>
<% //BMIDS00084 - have moved each button up screen by 50px because of Life Cover removal. (DPF 04/07/02) %>
<span style="LEFT: 10px; POSITION: absolute; TOP: 100px">
<input id="btnSummary" value="Quote Summary" type="button" style="WIDTH: 100px" class="msgButton">
</span>
<span style="LEFT: 10px; POSITION: absolute; TOP: 130px">
<input id="btnCreateNewQuote" value="Create Quote" type="button" style="WIDTH: 100px" class="msgButton">
</span>
<span style="LEFT: 10px; POSITION: absolute; TOP: 160px">
<input id="btnStoredQuotes" value="Stored Quotes" type="button" style="WIDTH: 100px" class="msgButton">
</span>
</div>
<div style="HEIGHT: 30px; LEFT: 10px; POSITION: absolute; TOP: 430px; WIDTH: 604px" class="msgGroup">
<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
<!--GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2<font face="MS Sans Serif" size="1">-->
<strong>
<label id="lblMonthlyCost" style="TOP: 0px; LEFT: 0px; WIDTH: 600px; POSITION: absolute"></label>
</strong>
<!--GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2</font>-->
</span>
</form>
</div>
<% // Main Buttons %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 465px; WIDTH: 612px">
<!-- #include FILE="msgButtons.asp" -->
</div>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>
<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cm010Attribs.asp" -->
<% // Specify Code Here %>
<script language="JScript">
<!--
var m_sApplicationMode	= "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var iBandCValue = "";
var iLifeValue = ""; 
var iPayProt = "";
var m_sCurrencyText;
var scScreenFunctions;
var bAnnualPayment = false;
var sBCMonthlyCost = "";
var bLifeRequired = true;
var m_bBusy=false;
var m_blnReadOnly = false;
var m_TaskID = "";
var m_CMStageID = "";
var m_bNoCompleteCheck = false;
var m_bCalledFromAcceptQuote = true;    // BMIDS977

<% /* SR 27/08/2004 : BMIDS815 */ %>
var m_sCallingScreen = "" ;
var m_sLTVChanged = "" ;
var m_sActiveStageNumber = "" ;
<% /* SR 27/08/2004 : BMIDS815  - End */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
var AppXML= null;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2 scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	%>

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	<%//GD		16/07/2002	BMIDS00045 %>
	var sTitle = "Application Cost Modelling";
	RetrieveContextData();
	if (m_sApplicationMode == "Quick Quote") sTitle = "Quick Quote Cost Modelling";
	FW030SetTitles(sTitle,"CM010",scScreenFunctions);
	
	<% /* DB BM0443 - Need to check if screen is read-only. If it is then do not run Completeness Check */ %>
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CM010");
	
	if(m_bNoCompleteCheck || m_blnReadOnly) <% /* DB BM0433 - Also added a read-only check. */ %>
	{
		scScreenFunctions.SetContextParameter(window,"idNoCompleteCheck","0");
		CompleteInitialisation();
	}
	else
	{
		colStatus.innerText = "Please wait ... Running Completeness Check";
		window.setTimeout("CompleteInitialisation()",50);
	}
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}	
	
function CompleteInitialisation()	
{
	if(!m_bNoCompleteCheck)
	{
		if (!CostModelCompleteCheck())
			return;
	}
	
	<% /* MAR424 - Maha T */ %>
	scScreenFunctions.SetContextParameter(window,"idApplicationMode","Cost Modelling");
	// EP2_1164 - Need to reset variable.
	m_sApplicationMode = scScreenFunctions.GetContextParameter(window,"idApplicationMode","Quick Quote");
	
	var sButtonList = new Array("Submit", "Cancel");	// APS 07/03/00 - SYS0332
	ShowMainButtons(sButtonList);
	m_sCurrencyText = scScreenFunctions.SetCurrency(window,frmScreen);

// DEBUG - Set context variablles
/*
m_sApplicationNumber = "B00003239";
m_sApplicationFactFindNumber = "1";

scScreenFunctions.SetContextParameter(window, "idCustomerNumber1", "00070599");
scScreenFunctions.SetContextParameter(window, "idCustomerVersionNumber1", "1");
scScreenFunctions.SetContextParameter(window, "idCustomerRoleType1", "1");
scScreenFunctions.SetContextParameter(window, "idCustomerNumber2", "00070637");
scScreenFunctions.SetContextParameter(window, "idCustomerVersionNumber2", "1");
scScreenFunctions.SetContextParameter(window, "idCustomerRoleType2", "1");
scScreenFunctions.SetContextParameter(window, "idMortgageApplicationValue", "10");

scScreenFunctions.SetContextParameter(window, "idMortgageSubQuoteNumber", "1");
scScreenFunctions.SetContextParameter(window, "idLifeSubQuoteNumber", "1");
scScreenFunctions.SetContextParameter(window, "idQuotationNumber", "1");
scScreenFunctions.SetContextParameter(window, "idPPSubQuoteNumber", "1");
scScreenFunctions.SetContextParameter(window, "idBCSubQuoteNumber", "1"); */
// END DEBUG
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	PopulateScreen();
	SetScreenToReadOnly();

	divStatus.style.visibility = "hidden";
	divBackground.style.visibility = "visible";
	frmScreen.style.visibility = "visible";
	scScreenFunctions.SetFocusToFirstField(frmScreen);

	// BG 05/09/00 SYS0964 Show statement if payment frequency in BC quote is Annual.
	if (bAnnualPayment == true)
		ShowBCText(sBCMonthlyCost);

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CM010");	
	
	 <% // GD BMIDS000376 START %>
	 if (m_blnReadOnly == true)
	 {
		frmScreen.btnAccept.disabled = true;
		frmScreen.btnCreateNewQuote.disabled = true;
		frmScreen.btnRecommend.disabled = true;
		//DPF - BMIDS01047 - more of the items on screen need to be made read only
		frmScreen.btnCalculate.disabled = true;
		frmScreen.btnMortgage.disabled = true;
		frmScreen.btnBuildingsAndContents.disabled = true;
		frmScreen.btnPaymentProtection.disabled = true;
		frmScreen.btnBuildingsAndContentsRequired.disabled = true;
		frmScreen.btnPaymentProtectionRequired.disabled = true;
		//END DPF
	 }
	 <% // GD BMIDS000376 END %>
	 // JD BMIDS749 check if rates have changed
	 <% // PB EP529 / MAR1705 12/05/2006 Only check if rates have changed when a quote has been modelled otherwise rates are null %>
	 if (frmScreen.txtMonthlyMortgageCost.value != "") {
		CheckForChangedRates(m_sApplicationNumber, m_sApplicationFactFindNumber);
	 }
	 <% /* EP529 / MAR1705 End */ %>
	 <% /* SR 27/08/2004 : BMIDS815 : If LTV changed in calculations of CM100, do Risk Assessment */ %>
	 <% /* BMIDS903 If LTV has changed, validation is now done on Accept but a warning is displayed here */ %>
	 if (m_sCallingScreen == 'CM100' && m_sLTVChanged == '1')
	 {
		alert('The LTV has changed. Risk Assessment will be performed when the quote is accepted.');	 
	 }
	 <% /* SR 27/08/2004 : BMIDS815 - End */ %>
}

function UpdateQuote()
{
	<%//var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	// AQR2007 - No point in updating if there's no quote number
	var QuoteNo;
	QuoteNo = scScreenFunctions.GetContextParameter(window,"idQuotationNumber");
	if (QuoteNo != "")
	{
		XML.CreateRequestTag(window, "UPDATE");
		XML.CreateActiveTag("QUOTATION");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("QUOTATIONNUMBER", QuoteNo);
	
		<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null
		if (IsChecked(frmScreen.btnLifeRequired))
			XML.CreateTag("LIFESUBQUOTEREQUIRED", "1");
		else 
			XML.CreateTag("LIFESUBQUOTEREQUIRED", "0");
		*/%> 
		<%//GD BMIDS00341 15/08/2002 Default Life flags to NULL%>
		XML.CreateTag("LIFESUBQUOTEREQUIRED", "0");
		if (IsChecked(frmScreen.btnBuildingsAndContentsRequired))
			XML.CreateTag("BCSUBQUOTEREQUIRED", "1");
		else 
			XML.CreateTag("BCSUBQUOTEREQUIRED", "0");
			
		if (IsChecked(frmScreen.btnPaymentProtectionRequired))
			XML.CreateTag("PPSUBQUOTEREQUIRED", "1");
		else 
			XML.CreateTag("PPSUBQUOTEREQUIRED", "0");

		//alert("a");
		// 		XML.RunASP(document, "UpdateQuotation.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "UpdateQuotation.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		//alert("b");
	
		if(XML.IsResponseOK())
			return true;
		else
			{
				alert("Error updating the quotation.");
				return false;	
			}
	// AQR2007 - No point in updating if there's no quote number
	}	
	else
	   return true;	
	// AQR2007 
}


//function window.onunload()
//{
//	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
//	XML.CreateRequestTag(window, "UPDATE");
//	XML.CreateActiveTag("QUOTATION");
//	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
//	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
//	XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idQuotationNumber"));
//	if (IsChecked(frmScreen.btnLifeRequired))XML.CreateTag("LIFESUBQUOTEREQUIRED", "1");
//	else XML.CreateTag("LIFESUBQUOTEREQUIRED", "0");
//	if (IsChecked(frmScreen.btnBuildingsAndContentsRequired))XML.CreateTag("BCSUBQUOTEREQUIRED", "1");
//	else XML.CreateTag("BCSUBQUOTEREQUIRED", "0");
//	if (IsChecked(frmScreen.btnPaymentProtectionRequired))XML.CreateTag("PPSUBQUOTEREQUIRED", "1");
//	else XML.CreateTag("PPSUBQUOTEREQUIRED", "0");

//	XML.RunASP(document, "UpdateQuotation.asp");
//	XML.IsResponseOK();
//}

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
m_sApplicationMode = scScreenFunctions.GetContextParameter(window,"idApplicationMode","Quick Quote");
m_sReadOnly	= scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B00003239");
m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
<% /* EP529 / MAR1704 - This is not required, as completeness check should ALWAYS be done
if(scScreenFunctions.GetContextParameter(window,"idNoCompleteCheck","0") == "1")
	m_bNoCompleteCheck = true;
*/ %>
<% /* SR 27/08/2004 : BMIDS815 */ %>
m_sCallingScreen = scScreenFunctions.GetContextParameter(window,"idCallingScreenID");   
scScreenFunctions.SetContextParameter(window,"idReturnScreenId",""); 
m_sLTVChanged = scScreenFunctions.GetContextParameter(window,"idLTVChanged");
<% /* BMIDS903 Clear LTVChanged context parameter after Risk Assessment has been done 
scScreenFunctions.SetContextParameter(window,"idLTVChanged","");  */ %>
m_sActiveStageNumber = scScreenFunctions.GetContextParameter(window,"idStageId");
<% /* SR 27/08/2004 : BMIDS815 - End */ %>
}
function ShowBCText(sBCMonthlyCost)
{
if (sBCMonthlyCost != "")
{
	var iAnnualBCCost = parseFloat(sBCMonthlyCost) * 12;
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var sBCAnnualCost = top.frames[1].document.all.scMathFunctions.RoundValue(iAnnualBCCost, 2);
	%>
	var sBCAnnualCost = top.frames[1].document.all.scMathFunctions.RoundValue(iAnnualBCCost, 2);
	lblMonthlyCost.innerHTML = "Total Monthly Cost includes an Annual Buildings and Contents Premium of " + m_sCurrencyText + sBCAnnualCost + " paid monthly";
}
}
function SetScreenToReadOnly()
{
scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.txtMonthlyMortgageCost.id);
//- BMIDS00084 - DPF 04/07/02 - Removal of life cover
//scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.txtMonthlyLifeCost.id);
scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.txtMonthlyBuildingsContentsCost.id);
scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.txtMonthlyPaymentProtectionCost.id);
scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.txtTotalMonthlyCost.id);
if (m_sReadOnly == "1" || m_sApplicationMode == "Quick Quote")
{
	frmScreen.btnAccept.disabled = true;
	frmScreen.btnRecommend.disabled	= true;
	DisableSubQuoteButton(frmScreen.btnAffordability, "msgAffordabilityDisabled");
	if (m_sReadOnly == "1")
	{
		frmScreen.btnCreateNewQuote.disabled = true;
		frmScreen.btnCalculate.disabled	= true;
		//DisableCheckBox(frmScreen.btnLifeRequired);	- BMIDS00084 - DPF 04/07/02 - Removal of life cover
		DisableCheckBox(frmScreen.btnBuildingsAndContentsRequired);
		DisableCheckBox(frmScreen.btnPaymentProtectionRequired);
	}
}
<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null */ %>
else
{
	//DisableCheckBox(frmScreen.btnLifeRequired);	- BMIDS00084 - DPF 04/07/02 - Removal of life cover
	//DisableSubQuoteButton(frmScreen.btnLife, "msgLifeDisabled");	- BMIDS00084 - DPF 04/07/02 - Removal of life cover
	
}
}
function PopulateScreen()
{
<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var XML = new scXMLFunctions.XMLObject();
%>

var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
XML.CreateRequestTag(window,"SEARCH");
XML.CreateActiveTag("BASICQUOTATIONDETAILS");
XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
// DC SYS0977 put in Customer list for LTV 
// var tagLIST = XML.CreateActiveTag("CUSTOMERLIST");
// for (var iLoop=1; iLoop<6; iLoop++)
// {
// var sCustomerNumber = scScreenFunctions.GetContextParameter(window, "idCustomerNumber" + iLoop);
// 		if (sCustomerNumber != "")
// 		{	
// 			XML.ActiveTag = tagLIST;
// 			XML.CreateActiveTag("CUSTOMER");			
// 			XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);		
// 			XML.CreateTag("CUSTOMERVERSIONNUMBER", scScreenFunctions.GetContextParameter(window, "idCustomerVersionNumber" + iLoop));		
// 		}
// }
if(m_sApplicationMode == "Quick Quote")
	// 	XML.RunASP(document,"GetQQValidatedQuotation.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"GetQQValidatedQuotation.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

else
	// 	XML.RunASP(document,"GetAQValidatedQuotation.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"GetAQValidatedQuotation.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

// DLM 10/07/00 SYS0949 Changed check not to be specific for Loan components
if (XML.IsResponseOK())
{
	
	scScreenFunctions.SetContextParameter(window,"idQuotationNumber",XML.GetTagText("ACTIVEQUOTENUMBER"));
	<%//scScreenFunctions.SetContextParameter(window,"idLifeSubQuoteNumber",XML.GetTagText("LIFESUBQUOTENUMBER"));%>
	<%//GD BMIDS00341%>
	scScreenFunctions.SetContextParameter(window,"idLifeSubQuoteNumber","");
	scScreenFunctions.SetContextParameter(window,"idBCSubQuoteNumber",XML.GetTagText("BCSUBQUOTENUMBER"));
	scScreenFunctions.SetContextParameter(window,"idPPSubQuoteNumber",XML.GetTagText("PPSUBQUOTENUMBER"));
	scScreenFunctions.SetContextParameter(window,"idMortgageSubQuoteNumber",XML.GetTagText("MORTGAGESUBQUOTENUMBER"));
	
	if(XML.GetTagText("ACTIVEQUOTENUMBER") == "")
	{
		DisableSubQuoteButton(frmScreen.btnMortgage, "msgMortgageDisabled");
		<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
		DisableSubQuoteButton(frmScreen.btnLife, "msgLifeDisabled"); */ %>
		DisableSubQuoteButton(frmScreen.btnBuildingsAndContents, "msgBuildingsContentsDisabled");		
		DisableSubQuoteButton(frmScreen.btnPaymentProtection, "msgPaymentProtectionDisabled");
		DisableSubQuoteButton(frmScreen.btnAffordability, "msgAffordabilityDisabled");
		// JD MAR1100 DisableSubQuoteButton(frmScreen.btnMaximumBorrowing, "msgMaxBorrowingDisabled");
		//DisableCheckBox(frmScreen.btnLifeRequired);	- BMIDS00084 - DPF 04/07/02 - Removal of life cover
		DisableCheckBox(frmScreen.btnBuildingsAndContentsRequired);		
		DisableCheckBox(frmScreen.btnPaymentProtectionRequired);		
		frmScreen.btnRecommend.disabled = true;
		frmScreen.btnAccept.disabled = true;
		frmScreen.btnSummary.disabled = true;
		frmScreen.btnStoredQuotes.disabled = true;
		frmScreen.btnCalculate.disabled = true;		
	}
	else
	{
		if(XML.GetTagText("QUOTATIONCOMPLETE") == "1")
		{
			if(XML.GetTagText("VALIDMORTGAGESUBQUOTE") == "0" || XML.GetTagText("VALIDBCSUBQUOTE") == "0"
			   || XML.GetTagText("VALIDPPSUBQUOTE") == "0")
			{
				XML.SetTagText("ACTIVEQUOTENUMBER","");
				UpdateActiveQuoteNumber(true, "");
				//BG 18/08/00 SYS0918 Disable the majority of the screen if there is an invalid quote.
				DisableSubQuoteButton(frmScreen.btnMortgage, "msgMortgageDisabled");
				<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
				DisableSubQuoteButton(frmScreen.btnLife, "msgLifeDisabled"); */ %>
				DisableSubQuoteButton(frmScreen.btnBuildingsAndContents, "msgBuildingsContentsDisabled");		
				DisableSubQuoteButton(frmScreen.btnPaymentProtection, "msgPaymentProtectionDisabled");
				DisableSubQuoteButton(frmScreen.btnAffordability, "msgAffordabilityDisabled");
				// JD MAR1100 DisableSubQuoteButton(frmScreen.btnMaximumBorrowing, "msgMaxBorrowingDisabled");
				//DisableCheckBox(frmScreen.btnLifeRequired);	- BMIDS00084 - DPF 04/07/02 - Removal of life cover
				DisableCheckBox(frmScreen.btnBuildingsAndContentsRequired);		
				DisableCheckBox(frmScreen.btnPaymentProtectionRequired);		
				frmScreen.btnRecommend.disabled = true;
				frmScreen.btnAccept.disabled = true;
				frmScreen.btnSummary.disabled = true;
				frmScreen.btnStoredQuotes.disabled = true;
				frmScreen.btnCalculate.disabled = true;		
				return;
			}
			else
			{				
				if(XML.GetTagText("ACTIVEQUOTENUMBER") == XML.GetTagText("ACCEPTEDQUOTENUMBER")) frmScreen.btnAccept.disabled = true;
				if(XML.GetTagText("ACTIVEQUOTENUMBER") == XML.GetTagText("RECOMMENDEDQUOTENUMBER")) frmScreen.btnRecommend.disabled = true;
				//MV - 29/01/2003 - BM0300 
				//frmScreen.txtMonthlyMortgageCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALNETMONTHLYCOST"),2);
				frmScreen.txtMonthlyMortgageCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALGROSSMONTHLYCOST"),2);
				<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
				frmScreen.txtMonthlyLifeCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALLIFEMONTHLYCOST"), 2); */ %>
				frmScreen.txtMonthlyBuildingsContentsCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALBCMONTHLYCOST"), 2);
				frmScreen.txtMonthlyPaymentProtectionCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALPPMONTHLYCOST"), 2);
				frmScreen.txtTotalMonthlyCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALQUOTATIONCOST"), 2);
				DisableSubQuoteButton(frmScreen.btnMortgage, "msgMortgageDisabled");
				<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
				DisableSubQuoteButton(frmScreen.btnLife, "msgLifeDisabled"); */%>
				DisableSubQuoteButton(frmScreen.btnBuildingsAndContents, "msgBuildingsContentsDisabled");
				DisableSubQuoteButton(frmScreen.btnPaymentProtection, "msgPaymentProtectionDisabled");
				// AS - Spec says to disable Affordability on complete quotation but it makes not sense ask BC or AP
				//DisableSubQuoteButton(frmScreen.btnAffordability, "msgAffordabilityDisabled");
				//DisableCheckBox(frmScreen.btnLifeRequired);			- BMIDS00084 - DPF 04/07/02 - Removal of life cover
				DisableCheckBox(frmScreen.btnBuildingsAndContentsRequired);		
				DisableCheckBox(frmScreen.btnPaymentProtectionRequired);				
				frmScreen.btnCalculate.disabled = true;
			}
		}
		else
		{
			if(XML.GetTagText("VALIDMORTGAGESUBQUOTE") == "0")
			{				
				XML.SetTagText("ACTIVEQUOTENUMBER","");
				UpdateActiveQuoteNumber(true, "");
				DisableSubQuoteButton(frmScreen.btnMortgage, "msgMortgageDisabled");
				<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
				DisableSubQuoteButton(frmScreen.btnLife, "msgLifeDisabled");*/ %>
				DisableSubQuoteButton(frmScreen.btnBuildingsAndContents, "msgBuildingsContentsDisabled");		
				DisableSubQuoteButton(frmScreen.btnPaymentProtection, "msgPaymentProtectionDisabled");
				DisableSubQuoteButton(frmScreen.btnAffordability, "msgAffordabilityDisabled");
				DisableCheckBox(frmScreen.btnBuildingsAndContentsRequired);		
				DisableCheckBox(frmScreen.btnPaymentProtectionRequired);
				//DisableCheckBox(frmScreen.btnLifeRequired);	- BMIDS00084 - DPF 04/07/02 - Removal of life cover
				frmScreen.btnCalculate.disabled = true;
				// DLM
				// return;
			}
			else
			{								
				EnableSubQuoteButton(frmScreen.btnMortgage, "msgMortgage");		
				<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 		
				EnableCheckBox(frmScreen.btnLifeRequired);		
				if (IsChecked(frmScreen.btnLifeRequired)) EnableSubQuoteButton(frmScreen.btnLife, "msgLife"); */ %>
				EnableCheckBox(frmScreen.btnBuildingsAndContentsRequired);		
				if (IsChecked(frmScreen.btnBuildingsAndContentsRequired)) EnableSubQuoteButton(frmScreen.btnBuildingsAndContents, "msgBuildingsContents");		
				EnableCheckBox(frmScreen.btnPaymentProtectionRequired);
				if (IsChecked(frmScreen.btnPaymentProtectionRequired)) EnableSubQuoteButton(frmScreen.btnPaymentProtection, "msgPaymentProtection");				
				if (m_sApplicationMode == "Cost Modelling") EnableSubQuoteButton(frmScreen.btnAffordability, "msgAffordability");
				// JD MAR1100 EnableSubQuoteButton(frmScreen.btnMaximumBorrowing, "msgMaxBorrowing");				
				frmScreen.btnSummary.disabled = false;
				frmScreen.btnStoredQuotes.disabled = false;	
				frmScreen.btnCalculate.disabled = false;
				frmScreen.btnCreateNewQuote.disabled = true;
				frmScreen.btnAccept.disabled = true;
				frmScreen.btnRecommend.disabled = true;
				var bClearBC = false;
				var bClearPP = false;
				if(XML.GetTagText("VALIDBCSUBQUOTE") == "0")
				{
					alert("The active Buildings & Contents sub-quote is no longer valid. You can either reinstate a previous sub-quote or create a new one");
					XML.SetTagText("BCSUBQUOTENUMBER","");
					scScreenFunctions.SetContextParameter(window,"idBCSubQuoteNumber","");
					bClearBC = true;
				}
				else frmScreen.txtMonthlyBuildingsContentsCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALBCMONTHLYCOST"), 2);

				if(XML.GetTagText("VALIDPPSUBQUOTE") == "0")
				{
					alert("The active Payment Protection sub-quote is no longer valid. You can either reinstate a previous sub-quote or create a new one");
					XML.SetTagText("PPSUBQUOTENUMBER","");					
					scScreenFunctions.SetContextParameter(window,"idPPSubQuoteNumber","");
					bClearPP = true;
				}
				else frmScreen.txtMonthlyPaymentProtectionCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALPPMONTHLYCOST"), 2);

				if(bClearBC || bClearPP)  // Update the Quotation record
				{
					var QuoteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					QuoteXML.CreateRequestTag(window,"UPDATE");
					QuoteXML.CreateActiveTag("QUOTATION");
					QuoteXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
					QuoteXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
					QuoteXML.CreateTag("QUOTATIONNUMBER",XML.GetTagText("QUOTATIONNUMBER"));
					if(bClearBC) QuoteXML.CreateTag("BCSUBQUOTENUMBER","");
					if(bClearPP) QuoteXML.CreateTag("PPSUBQUOTENUMBER","");
					// 					QuoteXML.RunASP(document,"UpdateQuotation.asp");
					// Added by automated update TW 09 Oct 2002 SYS5115
					switch (ScreenRules())
						{
						case 1: // Warning
						case 0: // OK
											QuoteXML.RunASP(document,"UpdateQuotation.asp");
							break;
						default: // Error
							QuoteXML.SetErrorResponse();
						}

					QuoteXML.IsResponseOK();
				}
				//MV - 29/01/2003 - BM0300 
				//var dblMortgageCost = XML.GetTagFloat("TOTALNETMONTHLYCOST");
				var dblMortgageCost = XML.GetTagFloat("TOTALGROSSMONTHLYCOST");   
				<%//var dblLifeCost = XML.GetTagFloat("TOTALLIFEMONTHLYCOST");%>
				<%//GD BMIDS00341%>
				var dblLifeCost = 0;
				if (dblMortgageCost > 0) frmScreen.txtMonthlyMortgageCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(dblMortgageCost, 2);
				else frmScreen.txtMonthlyMortgageCost.value = "";
				<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
				if (dblLifeCost > 0) frmScreen.txtMonthlyLifeCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(dblLifeCost, 2);
				else frmScreen.txtMonthlyLifeCost.value = ""; */ %>
			}			
		}
	}
	iBandCValue = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALBCMONTHLYCOST"), 2);
	<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
	iLifeValue = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALLIFEMONTHLYCOST"), 2); */ %>
	iPayProt = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALPPMONTHLYCOST"), 2);			
	//XML.WriteXMLToFile("c:\\ValidatedQuoation.xml");	
	
	<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
	if(XML.GetTagText("LIFESUBQUOTEREQUIRED") == "0") 
	{
		bLifeRequired = false;
		frmScreen.btnLifeRequired.className = "msgCross";
		frmScreen.txtMonthlyLifeCost.value = "";
		DisableSubQuoteButton(frmScreen.btnLife, "msgLifeDisabled");		
	} */ %>
	
	if((XML.GetTagText("BCSUBQUOTEREQUIRED") == "0") || (XML.GetTagText("BCSUBQUOTEREQUIRED") == "") )
	{
		frmScreen.btnBuildingsAndContentsRequired.className = "msgCross";
		frmScreen.txtMonthlyBuildingsContentsCost.value = "";
		DisableSubQuoteButton(frmScreen.btnBuildingsAndContents, "msgBuildingsContentsDisabled");				
	} else <% //GD BMIDS00320 Added Else %>
	{
		if(frmScreen.btnCalculate.disabled == false)
		{
			frmScreen.btnBuildingsAndContentsRequired.className = "msgTick";
			EnableSubQuoteButton(frmScreen.btnBuildingsAndContents, "msgBuildingsContents");	
		}	else frmScreen.btnBuildingsAndContentsRequired.className = "msgTickDisabled";		
	}
	
	if((XML.GetTagText("PPSUBQUOTEREQUIRED") == "0") || (XML.GetTagText("PPSUBQUOTEREQUIRED") == "") )	
	{
		frmScreen.btnPaymentProtectionRequired.className = "msgCross";
		frmScreen.txtMonthlyPaymentProtectionCost.value = "";
		DisableSubQuoteButton(frmScreen.btnPaymentProtection, "msgPaymentProtectionDisabled");
	} else <% //GD BMIDS00320 Added Else %>
	{
		if(frmScreen.btnCalculate.disabled == false)
		{
			frmScreen.btnPaymentProtectionRequired.className = "msgTick";
			EnableSubQuoteButton(frmScreen.btnPaymentProtection, "msgPaymentProtection");	
		} else frmScreen.btnPaymentProtectionRequired.className = "msgTickDisabled";
	}
	//BG SYS0964 21/08/00 If payment frequency in BC quote is annual, 
	//set variables for use in frmScreen.btnCalculate.onclick()
	var sBCPaymentFrequency = XML.GetTagText("FREQUENCY");
	if (sBCPaymentFrequency == "2")
	{	
		bAnnualPayment = true;
		sBCMonthlyCost = XML.GetTagText("TOTALBCMONTHLYCOST")		
		//ShowBCText(XML.GetTagText("TOTALBCMONTHLYCOST"));
	}
}

<% /*  MV  - 05/11/2002 - BMIDS00797 - Remove CalcAffordabilty 
	if (m_sApplicationMode == "Cost Modelling") 
	{
		CalculateAffordability(XML.GetTagText("QUOTATIONCOMPLETE") == 1);
	} */ %>
}
function UpdateActiveQuoteNumber(bDisplayMsg, sQuotationNumber)
{
var ApplicationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

if (bDisplayMsg == true)
	alert("The active quotation is no longer valid. You can either reinstate a previous quotation or create a new one");
	
ApplicationXML.CreateRequestTag(window,"UPDATE");
ApplicationXML.CreateActiveTag("APPLICATIONFACTFIND");
ApplicationXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
ApplicationXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
ApplicationXML.CreateTag("ACTIVEQUOTENUMBER",sQuotationNumber);
//	AW	18/11/2002	BMIDS00971
ApplicationXML.CreateTag("ACCEPTEDQUOTENUMBER","");
//	AW	18/11/2002	BMIDS00971 - End
ApplicationXML.RunASP(document,"UpdateApplicationFactFind.asp");
ApplicationXML.IsResponseOK();
scScreenFunctions.SetContextParameter(window,"idQuotationNumber",sQuotationNumber);
}
function DisplayAffordabilityGauge(bShow, iTotalCommitments, iTotalIncome)
{
<% /* BMIDS00268 MDC 31/10/2002 - Hide Affordability Gauge as requested */ %>
bShow = false;
<% /* BMIDS00268 MDC 31/10/2002 - End */ %>

if(!bShow)
{
	spnAffordabilityGaugeText.style.visibility = "hidden";
	spnAffordabilityIncomeMarker.style.visibility = "hidden";
	spnMaximumAffordability.style.visibility = "hidden";
	spnMininumAffordability.style.visibility = "hidden";
	divAffordabilityGauge.style.visibility	= "hidden";
	divAffordabilityIncomeMarker.style.visibility = "hidden";
	divInsideAffordability.style.visibility	= "hidden";
	divOutsideAffordability.style.visibility = "hidden";
	return;
}
var iGAUGEWIDTH	= 200;
var iIncomePosition = 0;
var iAffordabilityPosition = 0;
var iMaximumAvailable = parseInt(iTotalIncome * 1.5);
if (iMaximumAvailable != 0) iIncomePosition = iTotalIncome / iMaximumAvailable;
if (iMaximumAvailable != 0) iAffordabilityPosition = iTotalCommitments / iMaximumAvailable;

<% // we cannot exceed the dimension of the gauge %>
if (iAffordabilityPosition > 1) iAffordabilityPosition = 1;

<% // Income marker line %>
divAffordabilityIncomeMarker.style.width = (iIncomePosition * iGAUGEWIDTH);
<% // Income marker value %>
spnAffordabilityIncomeMarker.style.left	= ((iIncomePosition * iGAUGEWIDTH) - 10);
spnAffordabilityIncomeMarker.innerHTML	= "<strong>" + iTotalIncome + "</strong>";
spnMininumAffordability.innerHTML = "<strong>" + 0 + "</strong>";
<% // Set maximum affordability to 150% of the income %>
spnMaximumAffordability.innerHTML = "<strong>" + iMaximumAvailable + "</strong>";
if (iTotalCommitments > iTotalIncome)
{
	<% /*
		Due to the z order we require the inside div to be
		higher in the z order and therefore seen than the ouside div
		If we are here then we must be outside affordability; 
		we also need to show the inside affordability up to the income marker
	*/ %>
	divOutsideAffordability.style.width = ((iAffordabilityPosition * iGAUGEWIDTH) - 1);
	divInsideAffordability.style.width = (iIncomePosition * iGAUGEWIDTH);
}
else
{
	<% /*
		If we are here then we must be inside affordability and therefore
		we only need to display the Inside div
	*/ %>
	divInsideAffordability.style.width = (iAffordabilityPosition * iGAUGEWIDTH);
}
var sToolTip = "Total Commitments " + m_sCurrencyText + iTotalCommitments;
divAffordabilityIncomeMarker.title = sToolTip;
divOutsideAffordability.title = sToolTip;
divAffordabilityGauge.title	= sToolTip;
<% // Collection is first hidden then render then at this point the results are shown %>
spnAffordabilityGaugeText.style.visibility = "visible";
spnAffordabilityIncomeMarker.style.visibility = "visible";
spnMaximumAffordability.style.visibility = "visible";
spnMininumAffordability.style.visibility = "visible";
divAffordabilityGauge.style.visibility	= "visible";
divAffordabilityIncomeMarker.style.visibility = "visible";
divInsideAffordability.style.visibility	= "visible";
divOutsideAffordability.style.visibility = "visible";
}
<% /* JD MAR1100  remove max borrowing button
function frmScreen.btnMaximumBorrowing.onclick()
{
 // Route to the Maximum Borrowing screen 
	if (frmScreen.btnMaximumBorrowing.disabled != true)
	{
		 // BMIDS00654 MDC 04/11/2002  
		var sMSQNumber = scScreenFunctions.GetContextParameter(window,"idMortgageSubQuoteNumber", "");
		if (m_sApplicationMode != "Quick Quote" && sMSQNumber == "")
			alert("A mortgage sub quote must be created before calculating maximum borrowing.");
		else		
		{
			if(UpdateQuote())
				frmMaximumBorrowing.submit();
		}
		 // BMIDS00654 MDC 04/11/2002 - End  
	}
}
function frmScreen.btnMaximumBorrowing.onfocus()
{
if (frmScreen.btnMaximumBorrowing.disabled != true)
	frmScreen.btnMaximumBorrowing.className = "msgMaxBorrowingSelected";
}
function frmScreen.btnMaximumBorrowing.onblur()
{
frmScreen.btnMaximumBorrowing.className = "msgMaxBorrowing";
}
function frmScreen.btnMaximumBorrowing.onmouseover()
{
if (frmScreen.btnMaximumBorrowing.disabled != true)
	frmScreen.btnMaximumBorrowing.className = "msgMaxBorrowingSelected";
}
function frmScreen.btnMaximumBorrowing.onmouseout()
{
if (frmScreen.btnMaximumBorrowing.disabled != true)
	frmScreen.btnMaximumBorrowing.className = "msgMaxBorrowing";
}
*****************/ %>
function frmScreen.btnAffordability.onclick()
{
<% // Route to the Affordability screen 
%>

	if(UpdateQuote())
		frmAffordability.submit();
}
function frmScreen.btnAffordability.onmouseover()
{
if (frmScreen.btnAffordability.disabled != true)
	frmScreen.btnAffordability.className = "msgAffordabilitySelected";
}
function frmScreen.btnAffordability.onmouseout()
{
if (frmScreen.btnAffordability.disabled != true)
	frmScreen.btnAffordability.className = "msgAffordability";
}
function frmScreen.btnAffordability.onfocus()
{
frmScreen.btnAffordability.className = "msgAffordabilitySelected";
}
function frmScreen.btnAffordability.onblur()
{
frmScreen.btnAffordability.className = "msgAffordability";
}

function frmScreen.btnMortgage.onclick()
{

	//frmScreen.btnMortgage.disabled = true;
	
	if(m_bBusy==true)
		{
			<% /* EP529 / MAR1705 - Removed incorrect error message
			alert("BUSY!!!!! WHAT THE!");
			*/ %>
			alert("BUSY STATE DETECTED");
			return;
		}
	else
		m_bBusy=true;
		
	DisableSubQuoteButton(frmScreen.btnMortgage, "msgMortgageDisabled");	
	
	if(UpdateQuote())
		frmLoanComponents.submit();
	
	EnableSubQuoteButton(frmScreen.btnMortgage, "msgMortgageEnabled");
		
	m_bBusy=false;	
		
	//frmScreen.btnMortgage.disabled = false;	
}
<% //Mortgage Button %>
function frmScreen.btnMortgage.onfocus()
{
frmScreen.btnMortgage.className = "msgMortgageOver";
}
function frmScreen.btnMortgage.onblur()
{
frmScreen.btnMortgage.className = "msgMortgage";
}
function frmScreen.btnMortgage.onmouseover()
{
if (frmScreen.btnMortgage.disabled != true)
	frmScreen.btnMortgage.className = "msgMortgageOver";
}
function frmScreen.btnMortgage.onmouseout()
{
if (frmScreen.btnMortgage.disabled != true)
	frmScreen.btnMortgage.className = "msgMortgage";
}

<% //Life Button  %>
//function frmScreen.btnLife.onclick()	- BMIDS00084 - DPF 04/07/02 - Removal of life cover
{
//alert("Not Implemented Yet.");
}

<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null
<% /* function frmScreen.btnLife.onfocus()
{
frmScreen.btnLife.className = "msgLifeOver";
}
function frmScreen.btnLife.onblur()
{
frmScreen.btnLife.className = "msgLife";
}
function frmScreen.btnLife.onmouseover()
{
if (frmScreen.btnLife.disabled != true)
	frmScreen.btnLife.className = "msgLifeOver";
}
function frmScreen.btnLife.onmouseout()
{
if (frmScreen.btnLife.disabled != true)
	frmScreen.btnLife.className = "msgLife";
}
function frmScreen.btnLifeRequired.onclick()
{
if (frmScreen.btnLifeRequired.className == "msgTick")
{
	frmScreen.btnLifeRequired.className = "msgCross";
	frmScreen.txtMonthlyLifeCost.value = "";
	DisableSubQuoteButton(frmScreen.btnLife, "msgLifeDisabled");
}
else 
{	
	if(bLifeRequired == false)
	{
		alert("There is no life cover associated with this mortgage. Please remodel mortgage in order to reinstate life cover");
		frmScreen.btnMortgage.focus();
	}
	else
	{
	frmScreen.btnLifeRequired.className = "msgTick";
	frmScreen.txtMonthlyLifeCost.value = iLifeValue;
	EnableSubQuoteButton(frmScreen.btnLife, "msgLife");
	}
}
} */ %>
<% // BuildingsAndContents %>
function frmScreen.btnBuildingsAndContents.onclick()
{
	if (IsMortgageSubQuoteComplete()) 
	{
		if(UpdateQuote())
			frmBuildingAndContentsDetails.submit();
	}
<% // NOTE: On returning from cm020 a different sub-quote might now be active! %>
}
function frmScreen.btnBuildingsAndContents.onfocus()
{
frmScreen.btnBuildingsAndContents.className = "msgBuildingsContentsOver";
}
function frmScreen.btnBuildingsAndContents.onblur()
{
frmScreen.btnBuildingsAndContents.className = "msgBuildingsContents";
}
function frmScreen.btnBuildingsAndContents.onmouseover()
{
if (frmScreen.btnBuildingsAndContents.disabled != true)
	frmScreen.btnBuildingsAndContents.className = "msgBuildingsContentsOver";
}
function frmScreen.btnBuildingsAndContents.onmouseout()
{
if (frmScreen.btnBuildingsAndContents.disabled != true)
	frmScreen.btnBuildingsAndContents.className = "msgBuildingsContents";
}
function frmScreen.btnBuildingsAndContentsRequired.onclick()
{
if (frmScreen.btnBuildingsAndContentsRequired.className == "msgTick")
{
	frmScreen.btnBuildingsAndContentsRequired.className = "msgCross";
	frmScreen.txtMonthlyBuildingsContentsCost.value = "";
	DisableSubQuoteButton(frmScreen.btnBuildingsAndContents, "msgBuildingsContentsDisabled");
}
else 
{
	frmScreen.btnBuildingsAndContentsRequired.className = "msgTick";
	frmScreen.txtMonthlyBuildingsContentsCost.value = iBandCValue;
	EnableSubQuoteButton(frmScreen.btnBuildingsAndContents, "msgBuildingsContents");
}
}
<% //Payment Protection Button %>
function frmScreen.btnPaymentProtection.onclick()
{
if (IsMortgageSubQuoteComplete()) 
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window);
	if (m_sApplicationMode== "Quick Quote")
	{
		XML.CreateActiveTag("QUICKQUOTE");
		sASPFile = "CheckApplicantQQPPEligibility.asp";
	}
	else
	{
		XML.CreateActiveTag("APPLICATIONQUOTE");
		sASPFile = "CheckApplicantAQPPEligibility.asp";
	}		
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);	
	// 	XML.RunASP(document, sASPFile);
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, sASPFile);
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	//AQR SYS2203 DRC
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[0] == true)
    //AQR SYS2203 DRC		
		{
		var strApplicant1Eligible = XML.GetTagText("APPLICANT1ELIGIBLE");
		var strApplicant2Eligible = XML.GetTagText("APPLICANT2ELIGIBLE");

		scScreenFunctions.SetContextParameter(window,"idApplicant1PPEligible",strApplicant1Eligible);	
		scScreenFunctions.SetContextParameter(window,"idApplicant2PPEligible",strApplicant2Eligible);		

		// If either or the applicants are eligible then we routine to cm040
		if (strApplicant1Eligible=="1" || strApplicant2Eligible=="1")
			{
				if(UpdateQuote())
					frmPaymentProtectionDetails.submit();
			}
			else
			alert("No Applicants are eligible for Payment Protection");
		}
//AQR SYS2203 DRC		
		else
			alert("No Applicants are eligible for Payment Protection");
//AQR SYS2203 DRC			
   }			
	XML = null;
}
<% // NOTE: On returning from cm040 a different sub-quote might now be active! %>
}
function frmScreen.btnPaymentProtection.onfocus()
{
frmScreen.btnPaymentProtection.className = "msgPaymentProtectionOver";
}
function frmScreen.btnPaymentProtection.onblur()
{
frmScreen.btnPaymentProtection.className = "msgPaymentProtection";
}
function frmScreen.btnPaymentProtection.onmouseover()
{
if (frmScreen.btnPaymentProtection.disabled != true)
	frmScreen.btnPaymentProtection.className = "msgPaymentProtectionOver";
}
function frmScreen.btnPaymentProtection.onmouseout()
{
if (frmScreen.btnPaymentProtection.disabled != true)
	frmScreen.btnPaymentProtection.className = "msgPaymentProtection";
}
function frmScreen.btnPaymentProtectionRequired.onclick()
{
if (frmScreen.btnPaymentProtectionRequired.className == "msgTick") 
{
	frmScreen.btnPaymentProtectionRequired.className = "msgCross";
	frmScreen.txtMonthlyPaymentProtectionCost.value = "";
	DisableSubQuoteButton(frmScreen.btnPaymentProtection, "msgPaymentProtectionDisabled");
}
else
{
	frmScreen.btnPaymentProtectionRequired.className = "msgTick";
	frmScreen.txtMonthlyPaymentProtectionCost.value = iPayProt;
	EnableSubQuoteButton(frmScreen.btnPaymentProtection, "msgPaymentProtection");
}
}
function btnSubmit.onclick()
{
	if(UpdateQuote())
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
		{
			scScreenFunctions.SetContextParameter(window,"idReturnScreenId","");
			frmToGN300.submit();
		}
		else
		{
			scScreenFunctions.SetContextParameter(window,"idReturnScreenId","");
			frmApplicationMenu.submit();
		}
	}
}
function frmScreen.btnCreateNewQuote.onclick()
{
var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
XML.CreateRequestTag(window,"CREATE");
XML.CreateActiveTag("QUOTATION");
XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
XML.CreateTag("QUOTATIONNUMBER",scScreenFunctions.GetContextParameter(window,"idQuotationNumber",""));
// XML.RunASP(document,"CreateNewQuotation.asp");
// Added by automated update TW 09 Oct 2002 SYS5115
switch (ScreenRules())
	{
	case 1: // Warning
	case 0: // OK
	XML.RunASP(document,"CreateNewQuotation.asp");
		break;
	default: // Error
		XML.SetErrorResponse();
	}

if(XML.IsResponseOK())
{
	var sQuotationNumber = XML.GetTagText("QUOTATIONNUMBER");
	scScreenFunctions.SetContextParameter(window,"idQuotationNumber",sQuotationNumber);	
	UpdateActiveQuoteNumber(false, sQuotationNumber);	
	PopulateScreen();
	frmScreen.btnCreateNewQuote.disabled = true;
	frmScreen.txtTotalMonthlyCost.value = "";
}
}
function frmScreen.btnCalculate.onclick()
{
	var sMortgageSubQuoteNumber = scScreenFunctions.GetContextParameter(window,"idMortgageSubQuoteNumber");
	var sBCSubQuoteNumber = scScreenFunctions.GetContextParameter(window,"idBCSubQuoteNumber");
	var sPPSubQuoteNumber = scScreenFunctions.GetContextParameter(window,"idPPSubQuoteNumber");
	if (sMortgageSubQuoteNumber == "" || frmScreen.txtMonthlyMortgageCost.value == "")
	{
		alert("This quotation cannot be calculated without a completed mortgage sub-quote.");
		return
	}
	var dblTotalMonthlyCost = 0;	
	dblTotalMonthlyCost = parseFloat(frmScreen.txtMonthlyMortgageCost.value);
	
	<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
	if (IsChecked(frmScreen.btnLifeRequired) && frmScreen.txtMonthlyLifeCost.value != "")
	{
		dblTotalMonthlyCost += parseFloat(frmScreen.txtMonthlyLifeCost.value);		
	}
	else scScreenFunctions.SetContextParameter(window,"idLifeSubQuoteNumber",""); */ %>
		
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "SEARCH");
	XML.CreateActiveTag("MORTGAGESUBQUOTE");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("MORTGAGESUBQUOTENUMBER", sMortgageSubQuoteNumber);
	// 	XML.RunASP(document, "ValidateCompulsoryProducts.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "ValidateCompulsoryProducts.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if (XML.IsResponseOK() == true)
	{
		if (IsChecked(frmScreen.btnBuildingsAndContentsRequired) || 
			(XML.GetTagBoolean("COMPULSORYBCFLAG")==true))
		{
			if (sBCSubQuoteNumber = "" || frmScreen.txtMonthlyBuildingsContentsCost.value == "")
			{
				alert("A completed Buildings & Contents sub-quote is required before the Quotation can be calculated.");
				return;
			}
			else dblTotalMonthlyCost += parseFloat(frmScreen.txtMonthlyBuildingsContentsCost.value);
		}
		else scScreenFunctions.SetContextParameter(window,"idBCSubQuoteNumber","");

		if (IsChecked(frmScreen.btnPaymentProtectionRequired) ||
			(XML.GetTagBoolean("COMPULSORYPPFLAG")==true))
		{
			if (sPPSubQuoteNumber == "" || frmScreen.txtMonthlyPaymentProtectionCost.value == "")
			{
				alert("A completed Payment Protection sub-quote is required before the Quotation can be calculated.");
				return;
			}
			else dblTotalMonthlyCost += parseFloat(frmScreen.txtMonthlyPaymentProtectionCost.value);
		}
		else scScreenFunctions.SetContextParameter(window,"idPPSubQuoteNumber","");
		
		frmScreen.txtTotalMonthlyCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(dblTotalMonthlyCost.toString(), 2);
		frmScreen.btnCalculate.disabled = true;				
		frmScreen.btnSummary.disabled = false;
		frmScreen.btnCreateNewQuote.disabled = false;
		frmScreen.btnStoredQuotes.disabled = false;
		DisableSubQuoteButton(frmScreen.btnMortgage, "msgMortgageDisabled");		
		
		<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
		DisableSubQuoteButton(frmScreen.btnLife, "msgLifeDisabled"); */ %>
		
		DisableSubQuoteButton(frmScreen.btnPaymentProtection, "msgPaymentProtectionDisabled");
		DisableSubQuoteButton(frmScreen.btnBuildingsAndContents, "msgBuildingsContentsDisabled");
		
		<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
		DisableCheckBox(frmScreen.btnLifeRequired); */ %>
		
		DisableCheckBox(frmScreen.btnBuildingsAndContentsRequired);		
		DisableCheckBox(frmScreen.btnPaymentProtectionRequired);
		// JD MAR1100 EnableSubQuoteButton(frmScreen.btnMaximumBorrowing, "msgMaxBorrowing");
		if (m_sApplicationMode == "Cost Modelling") 
		{
			EnableSubQuoteButton(frmScreen.btnAffordability, "msgAffordability");
			frmScreen.btnAccept.disabled = false;
			frmScreen.btnRecommend.disabled = false;		
		}		

			
		XML.ResetXMLDocument();
		XML.CreateRequestTag(window, "UPDATE");
		XML.CreateActiveTag("QUOTATION");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idQuotationNumber"));
		XML.CreateTag("TOTALQUOTATIONCOST", dblTotalMonthlyCost.toString());
		<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
		XML.CreateTag("LIFESUBQUOTENUMBER", scScreenFunctions.GetContextParameter(window,"idLifeSubQuoteNumber",""));
		*/ %>
		<%//GD BMIDS00341%>
		XML.CreateTag("LIFESUBQUOTENUMBER", "");
		XML.CreateTag("BCSUBQUOTENUMBER", scScreenFunctions.GetContextParameter(window,"idBCSubQuoteNumber",""));
		XML.CreateTag("PPSUBQUOTENUMBER", scScreenFunctions.GetContextParameter(window,"idPPSubQuoteNumber",""));
		<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
		if (IsChecked(frmScreen.btnLifeRequired))XML.CreateTag("LIFESUBQUOTEREQUIRED", "1");
		else XML.CreateTag("LIFESUBQUOTEREQUIRED", "0"); */ %>
		<%//GD BMIDS00341%>
		XML.CreateTag("LIFESUBQUOTEREQUIRED", "0");
		if (IsChecked(frmScreen.btnBuildingsAndContentsRequired))XML.CreateTag("BCSUBQUOTEREQUIRED", "1");
		else XML.CreateTag("BCSUBQUOTEREQUIRED", "0");
		if (IsChecked(frmScreen.btnPaymentProtectionRequired))XML.CreateTag("PPSUBQUOTEREQUIRED", "1");
		else XML.CreateTag("PPSUBQUOTEREQUIRED", "0");
		
		// 		XML.RunASP(document, "StoreQuotation.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "StoreQuotation.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		XML.IsResponseOK();				
		<% /* MV - 05/11/2002 - BMIDs00797 - Remove CalcAffordability
		if (m_sApplicationMode == "Cost Modelling")
			CalculateAffordability(true);	// Calculate and update affordability gauge */ %>
	} 
}

function frmScreen.btnAccept.onclick()
{
	
	var bSuccess = GetApplicationData()
	if (bSuccess)
	{
		AppXML.SelectTag(null, "APPLICATIONFACTFIND");
		if (AppXML.GetTagText("KFIRECEIVEDBYALLAPPS")== "0")
		{
			alert("KFI has not yet been produced");
			frmScreen.btnStoredQuotes.onfocus();
			return false;
		}
	}
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window);
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "AcceptQuotation.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
	}

	if (XML.IsResponseOK())
	{
		frmScreen.btnAccept.disabled = true;
		UpdateStageSequence();
	}
	
	<% /* BMIDS903  Run Risk Assessment if the LTV has changed  */ %>
	if (m_sLTVChanged == '1')
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, "EXECUTE")
		XML.CreateActiveTag("RISKASSESSMENT");  //JLD SYS2982
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("STAGENUMBER", m_sActiveStageNumber);
			
		XML.RunASP(document, "RunRiskAssessment.asp");
		if(XML.IsResponseOK()) <% /* Call screen RA010 */ %>
		{
			var ArrayArguments = new Array();
			ArrayArguments[0] = scScreenFunctions.GetContextParameter(window,"idUserId", null);
			ArrayArguments[1] = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[2] = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
			ArrayArguments[3] = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
			ArrayArguments[4] = scScreenFunctions.GetContextParameter(window,"idStageId",null);
			ArrayArguments[5] = XML.CreateRequestAttributeArray(window);
			ArrayArguments[6] = scScreenFunctions.IsMainScreenReadOnly(window, "RA010");
			ArrayArguments[7] = scScreenFunctions.GetContextParameter(window,"idFreezeDataIndicator",null);
			ArrayArguments[8] = m_bCalledFromAcceptQuote;           // BMIDS977

			scScreenFunctions.DisplayPopup(window, document, "RA010.asp", ArrayArguments, 630, 380);
				
			<% /* BMIDS977  Now route to PP010 to review the fees */ %>
			scScreenFunctions.SetContextParameter(window,"idMetaAction","fromAcceptQuote");
			frmToPP010.submit();
				
			<% /* BMIDS903  Clear the context parameter */ %>
			scScreenFunctions.SetContextParameter(window,"idLTVChanged","");			
		} 
	}
	else
	{
		<% /*BMIDS977  Route to PP010 to review fees 
					   Set the context parameter to indicate that Accept Quote has been done */ %>
						   
		scScreenFunctions.SetContextParameter(window,"idMetaAction","fromAcceptQuote");
		frmToPP010.submit();
	}
}

function frmScreen.btnRecommend.onclick()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window);
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.RunASP(document, "RecommendQuotation.asp");
	if (XML.IsResponseOK()) frmScreen.btnRecommend.disabled = true;
}

function frmScreen.btnStoredQuotes.onclick()
{
	if(UpdateQuote())
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction",null);
		frmStoredQuotes.submit();
	}
}

function frmScreen.btnSummary.onclick()
{
	if(UpdateQuote())
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","fromCM010");
		scScreenFunctions.SetContextParameter(window,"idXML",scScreenFunctions.GetContextParameter(window,"idQuotationNumber",""));
		<% /* BM0176 MDC 20/12/2002 */ %>
		scScreenFunctions.SetContextParameter(window,"idXML2",scScreenFunctions.GetContextParameter(window,"idMortgageSubQuoteNumber",""));
		<% /* BM0176 MDC 20/12/2002 - End */ %>
		frmQuotationSummary.submit();
	}
}

function btnCancel.onclick()
{
	if(UpdateQuote())
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
		{
			scScreenFunctions.SetContextParameter(window,"idReturnScreenId","");
			frmToGN300.submit();
		}
		else
		{
			<% /* SR 21/05/2004 : BMIDS772 - route to MN060 instead of 'Attitude To Borrowing'. */ %>
			//frmMN060.submit();
			<%/*=[MC]-Redirect screen to MN060 in default circumstances, If it routed from DC200, cancel action redirects
				to DC200. */%>
			switch(scScreenFunctions.GetContextParameter(window,"idReturnScreenId"))
			{
				case 'DC200':
					scScreenFunctions.SetContextParameter(window,"idReturnScreenId","");
					frmToDC200.submit();
					break;
				default:
					scScreenFunctions.SetContextParameter(window,"idReturnScreenId","");
					frmMN060.submit();
					break;
			}
		}
				
	}
}

function EnableSubQuoteButton(refField, sEnabledClassName)
{
	if(refField.tagName == "INPUT")
		if(refField.type == "button")
		{
			refField.className = sEnabledClassName;
			refField.disabled = false;
			refField.tabIndex = 0;
		}
	<% /* MO	11/07/2002	BMDIS00199 - Start - Remove this code to reenable buttons */ %>
	//if ((refField.id == "btnBuildingsAndContents") || (refField.id == "btnPaymentProtection")) {
		//DisableSubQuoteButton(refField, sEnabledClassName + "Disabled")
	//}
	<% /* MO	11/07/2002	BMDIS00199 - End */ %>

	//GD Reapply above
	//BMIDS00312 19/08/2002
	//if ((refField.id == "btnBuildingsAndContents") || (refField.id == "btnPaymentProtection")) {
	//	DisableSubQuoteButton(refField, sEnabledClassName + "Disabled")
	//}
}

function DisableSubQuoteButton(refField, sDisabledClassName)
{
if(refField.tagName == "INPUT")
	if(refField.type == "button")
	{
		refField.className = sDisabledClassName;
		refField.disabled = true;
		refField.tabIndex = -1;
	}
}

function EnableCheckBox(refField)
{
	<% /* MO	11/07/2002	BMDIS00199 - Start - Remove this code to reenable buttons */ %>
	//refField.className = "msgCross";
	<% /* MO	11/07/2002	BMDIS00199 - End */ %>
	//GD Reapply above
	//BMIDS00312
	//refField.className = "msgCross";
	if (refField.type == "button")
	{
		refField.disabled = false;
		refField.tabIndex = 0;
		if (refField.className == "msgTickDisabled")
			refField.className = "msgTick";
		else if (refField.className == "msgCrossDisabled")
			refField.className = "msgCross";
	}
	<% /* MO	11/07/2002	BMDIS00199 - Start - Remove this code to reenable buttons */ %>
	//if ((refField.id == "btnBuildingsAndContentsRequired") || (refField.id == "btnPaymentProtectionRequired")) {
		//DisableCheckBox(refField)
	//}
	<% /* MO	11/07/2002	BMDIS00199 - End */ %>
	//GD reapply above
	//BMIDS00312
	//if ((refField.id == "btnBuildingsAndContentsRequired") || (refField.id == "btnPaymentProtectionRequired")) {
	//	DisableCheckBox(refField)
	//}
}

function DisableCheckBox(refField)
{
if (refField.type == "button")
{
	refField.disabled = true;
	refField.tabIndex = -1;
	if (refField.className == "msgTick")
		refField.className = "msgTickDisabled";
	else if (refField.className == "msgCross")
		refField.className = "msgCrossDisabled";
}
}

function IsChecked(refField)
{
	return ((refField.className == "msgTick") || (refField.className == "msgTickDisabled"));
}

//function CalculateAffordability(bQuotationComplete)
//{
//	var iTotalIncome = 0;
//	var iTotalCommitments = 0;	
//	
//	if (bQuotationComplete == true)
//	{
//		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
//		var sTagValue;
//		XML.CreateRequestTag(window, "UPDATE");
//		var tagActiveList = XML.CreateActiveTag("AFFORDABILITY");
//		var tagLIST = XML.CreateActiveTag("CUSTOMERLIST");
//		for (var iLoop=1; iLoop<6; iLoop++)
//		{
//			var sCustomerNumber = scScreenFunctions.GetContextParameter(window, "idCustomerNumber" + iLoop);
//			if (sCustomerNumber != "")
//			{	
//				XML.ActiveTag = tagLIST;
//				XML.CreateActiveTag("CUSTOMER");			
//				XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);		
//				XML.CreateTag("CUSTOMERVERSIONNUMBER", scScreenFunctions.GetContextParameter(window, "idCustomerVersionNumber" + iLoop));		
//				XML.CreateTag("CUSTOMERROLETYPE", scScreenFunctions.GetContextParameter(window, "idCustomerRoleType" + iLoop));
//			}
//		}
//		XML.ActiveTag = tagActiveList; 
//		XML.CreateTag("TYPEOFAPPLICATION", scScreenFunctions.GetContextParameter(window, "idMortgageApplicationValue"));
//		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
//		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
//		XML.CreateTag("MORTGAGESUBQUOTENUMBER", scScreenFunctions.GetContextParameter(window, "idMortgageSubQuoteNumber"));
		<% /* MV - 28/05/2002 - BMIDS00013 - BMIDS/BM045 - Default Life Cover Check Box to Null 
		XML.CreateTag("LIFESUBQUOTENUMBER", scScreenFunctions.GetContextParameter(window, "idLifeSubQuoteNumber")); */ %>
		<%//GD BMIDS00341%>
//		XML.CreateTag("LIFESUBQUOTENUMBER","");
//		XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window, "idQuotationNumber"));
		
		// SR 17/10/00 SYS0883: Add PPSubQuoteNUmber, BCSubQuoteNumber to the request for Affordability calculations
//		XML.CreateTag("PPSUBQUOTENUMBER", scScreenFunctions.GetContextParameter(window, "idPPSubQuoteNumber"));
//		XML.CreateTag("BCSUBQUOTENUMBER", scScreenFunctions.GetContextParameter(window, "idBCSubQuoteNumber"));
		// SR 17/10/00 End SYS0883
		
		// 		XML.RunASP(document, "CalculateAffordability.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
	//	switch (ScreenRules())
//			{
//			case 1: // Warning
//			case 0: // OK
//					XML.RunASP(document, "CalculateAffordability.asp");
//				break;
//			default: // Error
//				XML.SetErrorResponse();
//			}

//		if (XML.IsResponseOK()==true)
//		{		
			//XML.WriteXMLToFile("c:\\affordability.xml");
//			iTotalCommitments = XML.GetTagInt("LOANSANDLIABILITIES") + XML.GetTagInt("MONTHLYMORTGAGEPAYMENTS")
//								+ XML.GetTagInt("MORTGAGERELATEDINSURANCE") + XML.GetTagInt("OTHEROUTGOINGS");
//			iTotalIncome = XML.GetTagInt("TOTALMONTHLYINCOME");		
			//DisplayAffordabilityGauge(bQuotationComplete, iTotalCommitments, iTotalIncome);
//		}	
//	}
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2DisplayAffordabilityGauge(bQuotationComplete, iTotalCommitments, iTotalIncome);%>
//}

function IsMortgageSubQuoteComplete()
{
if (scScreenFunctions.GetContextParameter(window,"idMortgageSubQuoteNumber") == "" ||
	frmScreen.txtMonthlyMortgageCost.value == "")
{
	alert("A completed Mortgage Sub-Quote must exist for this Quotation");
	return false;
}
else return true;	
}

function CostModelCompleteCheck()
{
	thisXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	thisXML.CreateRequestTag(window,null);
	thisXML.CreateActiveTag("COMPLETENESSCHECK");
	thisXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	thisXML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	thisXML.RunASP(document,"RunCompletenessCheck.asp");
	if(thisXML.IsResponseOK())
	{
		colStatus.innerText = "Please wait ... Analysing Completeness Check response";
		var xmlList = thisXML.XMLDocument.selectNodes("RESPONSE/COMPLETIONRULE[@COSTMODELLINGIND='Y']");
		if(xmlList.length != 0)
		{
			colStatus.innerText = "Completeness Check failed, routing to Completeness Check screen ...";
			window.setTimeout("CompletenessCheckRouting()",1000);
//			alert("Completeness Check failed, routing to Completeness Check screen ...");
			return false;
		}
	}
	return true;
}


function UpdateStageSequence() 
{
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var strTaskID = XML.GetGlobalParameterString(document, "TMAcceptQuote");
	var strCMStageID = XML.GetGlobalParameterAmount(document, "TMCostModellingStage"); 
	var strStageID = scScreenFunctions.GetContextParameter(window,"StageID");
	var strCaseID = scScreenFunctions.GetContextParameter(window,"idApplicationNumber");
	var strActivityID = scScreenFunctions.GetContextParameter(window,"idActivityId");
	var strStageSequenceNumber = scScreenFunctions.GetContextParameter(window,"idCurrentStageOrigSeqNo");
	
	<% /* PSC 13/12/01 SYS3425 - Start */ %>
	var strRQTaskId = XML.GetGlobalParameterString(document, "TMRemodelMortgage");
	var strCurrentStage; 
	
	<% /* Get tasks for the current stage */ %>
	TasksXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	TasksXML.CreateRequestTag(window,"GetCurrentStage");
	TasksXML.CreateActiveTag("CASEACTIVITY");
	TasksXML.SetAttribute("SOURCEAPPLICATION","Omiga");
	TasksXML.SetAttribute("CASEID",strCaseID);
	TasksXML.SetAttribute("ACTIVITYID",strActivityID);
	// 	TasksXML.RunASP(document,"MsgTMBO.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			TasksXML.RunASP(document,"MsgTMBO.asp");
			break;
		default: // Error
			TasksXML.SetErrorResponse();
		}

	
	if(TasksXML.IsResponseOK())
	{
		var iIndex = 0;
		var strStatus;
		var bOutstandingTask = false;
		var strPattern = "RESPONSE/CASESTAGE/CASETASK[@TASKID='" + strRQTaskId + "']";
		var TaskList = TasksXML.XMLDocument.selectNodes(strPattern);
		var ValidationList = new Array(1);
		
		if (TaskList.length > 0)
		{
			<% /* Get Current Stage */ %>
			var CaseStage;
			CaseStage = TasksXML.XMLDocument.selectSingleNode("RESPONSE/CASESTAGE");
			
			if (CaseStage != null)
				strCurrentStage = CaseStage.getAttribute("STAGEID");	

			<% /* Check if any of the tasks are outstanding RemodelMortgage */ %>
			ValidationList[0] = "I";
			
			for (iIndex = 0;iIndex < TaskList.length && bOutstandingTask == false;iIndex++)
			{
				strStatus = TaskList.item(iIndex).getAttribute("TASKSTATUS");
				if (TasksXML.IsInComboValidationList(document,"TaskStatus", strStatus, ValidationList))
					bOutstandingTask = true;	
			}
		
			<% /* If outstanding remodel mortgage task update this */ %>
			if (bOutstandingTask == true)
			{
				strTaskID = strRQTaskId;
				strStageId = strCurrentStage;
			}
		}
	}
	<% /* PSC 13/12/01 SYS3425 - End */ %>
	
	if (strStageSequenceNumber == strCMStageID || bOutstandingTask == true) 
	{
		//Make the XML
		thisXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		thisXML.CreateRequestTag(window,"CompleteSimpleCaseTask");
		thisXML.CreateActiveTag("CASETASK");
		thisXML.SetAttribute("SOURCEAPPLICATION","Omiga");
		thisXML.SetAttribute("CASEID",strCaseID);
		thisXML.SetAttribute("ACTIVITYID",strActivityID);
		thisXML.SetAttribute("STAGEID",strStageID);
		thisXML.SetAttribute("TASKID",strTaskID);
		//Call asp page
		// 		thisXML.RunASP(document,"MsgTMBO.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					thisXML.RunASP(document,"MsgTMBO.asp");
				break;
			default: // Error
				thisXML.SetErrorResponse();
			}

		if(thisXML.IsResponseOK())
		{
			return;
		}
		else
		{
			alert("Error updating the quotation.");
			return;	
		}
	}	
}
//JD BMIDS749 			
function CheckForChangedRates(sApplicationNumber, sApplicationFactFindNumber)
{	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"");
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	XML.RunASP(document, "HaveRatesChanged.asp");
	if (XML.IsResponseOK())
	{
		if (XML.GetTagAttribute("QUOTATION", "RATESINCONSISTENT") == "1")
			alert("Rates on this product are not consistent, please remodel.");
	}
}

function CompletenessCheckRouting()
{
	scScreenFunctions.SetContextParameter(window,"idReturnScreenId","MN060");
	scScreenFunctions.SetContextParameter(window, "idProcessContext", "CompletenessCheck");
	scScreenFunctions.SetContextParameter(window,"idXML",thisXML.XMLDocument.xml);
	frmToGN300.submit();
}

function GetApplicationData()
{
	AppXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AppXML.CreateRequestTag(window, null);
	AppXML.CreateActiveTag("APPLICATION");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
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
		return true;
	
	return false;
}

-->
</script>
</body>
</html>
<% /* OMIGA BUILD VERSION 045.02.07.15.00 */ %>
<% /* OMIGA BUILD VERSION 046.02.08.08.00 */ %>




