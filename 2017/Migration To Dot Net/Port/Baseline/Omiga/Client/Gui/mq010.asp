<%@ LANGUAGE="JSCRIPT" codePage="28591"%>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<%/*
Workfile:      MQ010.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Loan Composition
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
HMA		20/07/05	Created (MAR18)
HMA     10/08/05    MAR18 Improvements from review.
HMA     07/09/05	Temporarily remove MSGAlert
  
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
MF		15/09/2005	MAR30		IncomeCalcs modified to send ActivityID into calculation
MV		22/10/2005	MAR109		Amended Initialise()
PCT		05/01/2006	MAR100		Creation of Variables required for incentives screen cm150
								Removed duplicate var declarations in "Initialise" fn
DRC     11/02/2006  MAR1257     Moved the saving of Level of advice from combo onchange event  to 	onsubmit
PJO     10/03/2006  MAR1378     Disable buttons in read only  
PE		13/03/2006	MAR1061		APP04a - On 'Quote Summary ' clicking Get New Quote removes the previous 'Property purchase pric/estimated valuation (£)'
HMA     22/03/2006  MAR1452     Ensure that PurchasePriceOrEstimatedValue gets saved.
HMA     23/03/2006  MAR1061     Send PurchasePriceOrEstimatedValue correctly when purchase price or amount requested changes.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
SAB		21/06/2006	EP837		Rollback MAR1707 - We want to format LTV to 2 decimal places
AShaw	08/11/2006	EP2_8		Add new button and related code - Application Source.
								Add extra params when loading MQ020 (Add/Edit).
								Added default value for new params.
								Passes params to and from MQ020.
GHun	20/11/2006	EP2_123		Changed Initialise to fix spelling error in m_sApplicationFFNumber
GHun	23/11/2006	EP2_171		Fixed saving of TypeOfValuation, PropertyLocation and PurchasePrice
AShaw	16/11/2006	EP2_55		New code for PSW / TOE apps.
GHun	29/11/2006	EP2_216		Fix purchase price being lost
MAH		04/11/2006	EP2_274		Removed EP2_55 changes
MAH		04/11/2006	EP2_274		Reestablish EP2_55 changes
AShaw	28/12/2006	EP2_56		New method for Adding components.
INR		12/01/2007	EP2_697		Increase size of MQ020 popup
INR		25/01/2007	EP2_780		Added Right To Buy Discount
******	NB. DRAWDOWN needs uncommenting when col in table. ******
AShaw	26/01/2007	EP2_771		DisableAmountRequested call moved.
INR		28/01/2007	EP2_1040	Update RTBDiscount for btnAdd and btnEdit
AShaw	31/01/2007	EP2_1023	The ReturnIndicator method MUST return a non empty 
								(nor a single space) string.
PSC		02/02/2007	EP2_1113	Correct screen functionality for further borrowing, switches etc	
PSC		07/02/2007	EP2_1271	Set up start date on loan component when generating from account															
PSC		12/02/2007	EP2_1314	Generate new quotation when GenerateQuote is pressed together with other corrections
INR		16/02/2007	EP2_1464	Should be using Combo "PurposeOfLoanTrfOfEquity" for Validation
MAH		07/03/2007	EP2_1614	Disabled automatic Calculation processing on flexible products	
MAH		08/03/2007	EP2_1730	Altered Right to buy discount validation												
PE		30/03/2007	EP2_1888	Ensure "Application Source" screen has been populated.
DS		30/03/2007	EP2_1943	Disabled 'Add' and 'New Quote' buttons for application types TOE,PSW or NP 
AShaw	02/04/2007	EP2_2168	Add new method to prevent Drawdown > Requested amount.
DS		10/04/2007	EP2_1943	On calculate button click disabled 'Add' and 'New Quote' buttons for application types TOE,PSW or NP 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Scriptlets - remove any which are not required */ %>
<span style="TOP: 271px; LEFT: 310px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="DISPLAY: none; HEIGHT: 24px; WIDTH: 304px; VISIBILITY: hidden" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToCM010" method="post" action="cm010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM065" method="post" action="cm065.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC010" method="post" action="DC010.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange">

<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 350px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Purchase Price
		<span style="TOP: 0px; LEFT: 110px; POSITION: ABSOLUTE">
			<label style="LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtPurchasePrice" maxlength="10" style="TOP: -3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt" type="text">
		</span>
	</span>
<% /* AShaw - EP2_8 - Reposition and add new button */ %>
	<span id="spnRTBDiscount" style="TOP: 10px; LEFT: 236px; POSITION: ABSOLUTE" class="msgLabel">
		Right To Buy Discount
		<span style="TOP: 0px; LEFT: 120px; POSITION: ABSOLUTE">
			<label style="LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtRightToBuyDisc" maxlength="10" style="TOP: -3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt" NAME="txtRightToBuyDisc">
		</span>
	</span>	
	<span style="TOP: 6px; LEFT: 476px; POSITION: ABSOLUTE">
		<input id="btnApplicationSrc" value="Application Source" type="button" style="WIDTH: 120px" class="msgButton" NAME="btnApplicationSrc">
	</span>

	<span style="TOP: 42px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Amount Requested
		<span style="TOP: 0px; LEFT: 110px; POSITION: ABSOLUTE">
			<label style="LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtAmtRequested" maxlength="10" style="TOP: -3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 42px; LEFT: 192px; POSITION: ABSOLUTE" class="msgLabel">
		LTV 
		<span style="TOP: -3px; LEFT: 25px; POSITION: ABSOLUTE">
			<input id="txtLTVPercent" maxlength="10" style="WIDTH: 38px; POSITION: ABSOLUTE" class="msgTxt">
			<span style="LEFT: 41px; POSITION: absolute; TOP: 3px" class="msgLabel">
				%
			</span>
		</span>
	</span>	
	<span style="TOP: 42px; LEFT: 283px; POSITION: ABSOLUTE" class="msgLabel">
		Draw Down
		<span style="TOP: 0px; LEFT: 74px; POSITION: ABSOLUTE">
			<label style="LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtDrawDown" maxlength="10" style="TOP: -3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>	
	<span style="TOP: 42px; LEFT: 430px; POSITION: ABSOLUTE" class="msgLabel">
		Total Loan Amount
		<span style="TOP: 0px; LEFT: 108px; POSITION: ABSOLUTE">
			<label style="LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtTotalLoanAmt" maxlength="10" style="TOP: -3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>	
	<div id="spnTable" style="TOP: 60px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblTable" width="596" border="0" cellspacing="0"cellpadding="0" class="msgTable">
			<tr id="rowTitles">
				<td width="5%" class="TableHead">Ind</td>	
				<td width="10%" class="TableHead">Product Number</td>	
				<td width="8%" class="TableHead">Product Name</td>	
				<td width="8%" class="TableHead">Initial Type</td>		
				<td width="10%" class="TableHead">Interest Rate&nbsp;Step</td>	
				<td width="10%" class="TableHead">Resolved Rate</td>
				<td width="10%" class="TableHead">Loan Amt</td>  
				<td width="8%" class="TableHead">Repay Type</td>  
				<td width="10%" class="TableHead">Monthly Cost</td>
				<!--PSC		10/07/02	BMIDS00062 - Start -->
				<td width="8%" class="TableHead">APR</td>
				<td width="10%" class="TableHead">Cost at Final Rate</td></tr>						
			<tr id="row01">		<td width="5%" class="TableTopLeft">&nbsp;</td>	<td width="10%" class="TableTopCenter">&nbsp;</td>	<td width="8%" class="TableTopCenter">&nbsp;</td>	<td width="8%" class="TableTopCenter">&nbsp;</td> <td width="10%" class="TableTopCenter">&nbsp;</td>		<td width="10%" class="TableTopCenter">&nbsp;</td>		<td width="10%" class="TableTopCenter">&nbsp;</td>  <td width="10%" class="TableTopCenter">&nbsp;</td>  <td width="10%" class="TableTopCenter">&nbsp;</td>		<td width="8%" class="TableTopCenter">&nbsp;</td>		<td width="10%" class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="5%" class="TableBottomLeft">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="8%" class="TableBottomCenter">&nbsp;</td>	<td width="8%" class="TableBottomCenter">&nbsp;</td> <td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>   <td width="10%" class="TableBottomCenter">&nbsp;</td>   <td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="8%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomRight">&nbsp;</td></tr>
			<!--PSC		10/07/02	BMIDS00062 - End-->
		</table>
	</div>
	<span style="TOP: 260px; LEFT: 4px; POSITION: ABSOLUTE">
		<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	<span style="TOP: 260px; LEFT: 70px; POSITION: ABSOLUTE">
		<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	<span style="TOP: 260px; LEFT: 140px; POSITION: ABSOLUTE">
		<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	<span style="TOP: 265px; LEFT: 332px; POSITION: ABSOLUTE" class="msgLabel">
		Total Monthly Cost
		<span style="TOP: 0px; LEFT: 106px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtTotalMnthCost" maxlength="10" style="TOP:-3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 290px; LEFT: 273px; POSITION: ABSOLUTE" class="msgLabel">
		Monthly Cost Less Drawdown
		<span style="TOP: 0px; LEFT: 165px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtMnthCostLessDD" maxlength="10" style="TOP:-3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 320px; LEFT: 130px; POSITION: ABSOLUTE">
		<input id="btnGenerateQuote" value="Generate Quote from Mtg Acc" type="button" style="WIDTH: 160px" class="msgButton" NAME="btnGenerateQuote">
	</span>
	<span style="TOP: 320px; LEFT: 294px; POSITION: ABSOLUTE">
		<input id="btnNewQuote" value="New Quote" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	<span style="TOP: 320px; LEFT: 397px; POSITION: ABSOLUTE">
		<input id="btnOneOffCosts" value="One Off Costs" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	<span style="TOP: 320px; LEFT: 500px; POSITION: ABSOLUTE">
		<input id="btnCalculate" value="Calculate" type="button" style="WIDTH: 100px" class="msgButton" NAME="btnCalculate">
	</span>	
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 412px; WIDTH: 612px">
<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!--  #include FILE="attribs/mq010attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sApplicationMode = "";
var m_sReadOnly = "";
var m_sMortgageSubquoteNumber = "";
var m_sLifeSubquoteNumber = "";
var m_sApplicationNumber = "";
var m_sApplicationFFNumber = "";
var m_sCurrency = "";
var subQuoteXML = null;
var m_sAmountRequested = "";
var m_sPurchasePrice = "";
var m_sMortgageType = "";
var m_sActiveQuoteNumber = "";

var m_iTableLength = 10;
var m_iNumOfLoanComponents = 0;
var m_sAmtRequested = "";
var m_sDrawDown = "";		// Drawdown amount
var m_iManualIncentive = 0;	// Manual incentive amount
var m_PopWinInd = "No";		// Marker for initialisation
var scScreenFunctions;
var m_blnReadOnly = false;
var m_iInterestOnlyAmount = 0;
var m_iCapitalInterestAmount = 0;
var m_XMLRepay = null;
var m_bPPPresent = false;
var m_bScreenEntry = false;
var m_sLTVChanged = "";

var m_sPropertyLocation;
var m_sValuationType;
var MQ010ParamsXML = null;   // EP2_8 - New Params for Product search on MQ020.
//EP2_780
var m_XMLTypeOfMortgage = null;
var m_XMLSpecialSchemes = null;
var mDiscountAmount = "";

<% /* PSC 12/02/2007 EP2_1314 - Start */ %>
var m_blnIsPort = false;
var m_blnIsSwitch = false;
var m_blnIsCLI = false;
<% /* PSC 12/02/2007 EP2_1314 - End */ %>
var m_FlexibleMortgageInd; <%/* EP2_1614 */%>

var scClientScreenFunctions;
function window.onload()
{
	
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();

	var sButtonList = new Array("Submit", "Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Loan Composition","MQ010",scScreenFunctions);
	m_sCurrency = scScreenFunctions.SetCurrency(window,frmScreen);
	RetrieveContextData();
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
	
	m_bScreenEntry = true;	

	PopulateCombos();
	
    <% /* PJO 10/03/2006 MAR1378 - Set read only flag before initialise */ %>
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "MQ010");

	Initialise();	

	m_bScreenEntry = false;	
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
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
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFFNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sMortgageSubquoteNumber = scScreenFunctions.GetContextParameter(window,"idMortgageSubQuoteNumber",null);
	m_sLifeSubquoteNumber = scScreenFunctions.GetContextParameter(window,"idLifeSubquoteNumber",null);	
	m_sAmountRequested = scScreenFunctions.GetContextParameter(window,"idAmountRequested","0");
	m_sMortgageType = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","");
	
}

function Initialise()
{
	var bSuccess = true;
	var sPropertyLocationValidation;
	var sValuationTypeValidation;
	<% /* Set the application mode to be Cost Modelling */ %>
	m_sApplicationMode = "Cost Modelling";
	
	<% /* PSC 12/02/2007 EP2_1314 */ %>
	frmScreen.btnAdd.disabled = false;

	<% /* Get the active quote */ %>

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("BASICQUOTATIONDETAILS");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);

	// EP2_55 (EP2_771 Moved to here) - Alter logic for Amount requested enabling.
	if (DisableAmountRequested() == true)
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAmtRequested");

	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"GetAQValidatedQuotation.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	bSuccess = XML.IsResponseOK();
	if (bSuccess == true)
	{
		scScreenFunctions.SetContextParameter(window,"idQuotationNumber",XML.GetTagText("ACTIVEQUOTENUMBER"));
		scScreenFunctions.SetContextParameter(window,"idLifeSubQuoteNumber","");
		scScreenFunctions.SetContextParameter(window,"idBCSubQuoteNumber",XML.GetTagText("BCSUBQUOTENUMBER"));
		scScreenFunctions.SetContextParameter(window,"idPPSubQuoteNumber",XML.GetTagText("PPSUBQUOTENUMBER"));
		scScreenFunctions.SetContextParameter(window,"idMortgageSubQuoteNumber",XML.GetTagText("MORTGAGESUBQUOTENUMBER"));
	
		m_sActiveQuoteNumber = XML.GetTagText("ACTIVEQUOTENUMBER");
		m_sMortgageSubquoteNumber = XML.GetTagText("MORTGAGESUBQUOTENUMBER");

		<%/* Need this for RTB Discount */%>
		var LoanPropertyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

		LoanPropertyXML.CreateRequestTag(window,null);
		LoanPropertyXML.CreateActiveTag("LOANPROPERTYDETAILS");
		LoanPropertyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		LoanPropertyXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		LoanPropertyXML.RunASP(document,"GetLoanPropertyDetails.asp");

		var ErrorTypes = new Array("RECORDNOTFOUND");
		var LoanPropertyErrorReturn = LoanPropertyXML.CheckResponse(ErrorTypes);

		if(m_sActiveQuoteNumber == "")
		{
			<% /* No active quote so create one */ %>
			bSuccess = CreateQuotation();
			
			if(bSuccess == true)
			{
				<% /* Set the property location and valuation type according to global parameters*/ %>
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
				sPropertyLocationValidation = XML.GetGlobalParameterString(document, "CMDefaultPropertyLocation");
				<% /* Note that the following line of code will bring back the FIRST value id for this validation type */ %>
				m_sPropertyLocation = XML.GetComboIdForValidation("PropertyLocation", sPropertyLocationValidation, null, document);

				var ValidationList = new Array(1);
				ValidationList[0] = "N";  // New Loan
				XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				if (XML.IsInComboValidationList(document,"TypeOfMortgage",m_sMortgageType, ValidationList))
				{
					XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					sValuationTypeValidation = XML.GetGlobalParameterString(document, "CMDefaultNewLoanValType");
				}
				else 
				{
					ValidationList[0] = "R";   // Remortgage
					XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					if (XML.IsInComboValidationList(document,"TypeOfMortgage",m_sMortgageType, ValidationList))
					{
						sValuationTypeValidation = XML.GetGlobalParameterString(document, "CMDefaultRemortgageValType");
					}
					else
					{
						sValuationTypeValidation = XML.GetGlobalParameterString(document, "CMDefaultAdditionalValType");
					}
				}
				m_sValuationType = XML.GetComboIdForValidation("ValuationType", sValuationTypeValidation, null, document);

				<% /* Create an entry in the NewProperty and SharedOwnersipDetails tables
					Default the Valuation Type and Property Location */ %>
			<%/* 	var LoanPropertyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

				LoanPropertyXML.CreateRequestTag(window,null);
				LoanPropertyXML.CreateActiveTag("LOANPROPERTYDETAILS");
				LoanPropertyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
				LoanPropertyXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
				LoanPropertyXML.RunASP(document,"GetLoanPropertyDetails.asp");

				var ErrorTypes = new Array("RECORDNOTFOUND");
				var LoanPropertyErrorReturn = LoanPropertyXML.CheckResponse(ErrorTypes);*/%>

				<%/* Record found */%>
				if (LoanPropertyErrorReturn[0] == true)
				{
					LoanPropertyXML.SelectTag(null, "LOANPROPERTYDETAILS")
					
					var PropertyLocation  = LoanPropertyXML.GetTagText("PROPERTYLOCATION");
					var ValuationType = LoanPropertyXML.GetTagText("VALUATIONTYPE");

					if (PropertyLocation == "" || ValuationType == "")
					{
						var UpdateLoanPropertyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
						
						UpdateLoanPropertyXML.CreateRequestTag(window,null);
						UpdateLoanPropertyXML.CreateActiveTag("NEWPROPERTY");
						UpdateLoanPropertyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
						UpdateLoanPropertyXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);	//EP2_123 GHun
						
						if (PropertyLocation == "" )
							UpdateLoanPropertyXML.CreateTag("PROPERTYLOCATION", m_sPropertyLocation);
							
						if (ValuationType == "")
							UpdateLoanPropertyXML.CreateTag("VALUATIONTYPE", m_sValuationType);
						
						<% /* EP2_171 GHun
						UpdateLoanPropertyXML.RunASP(document,"UpdateLoanProperty.asp");	 */ %>
						UpdateLoanPropertyXML.RunASP(document,"UpdateNewPropertyGeneral.asp");
						bSuccess = UpdateLoanPropertyXML.IsResponseOK();	
						<% /* EP2_171 End */ %>
					}
				}
				else if(LoanPropertyErrorReturn[1] == ErrorTypes[0]) <%/* Record not found */%>
				{
				
					XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					XML.CreateRequestTag(window,null)

					<%/* NEWPROPERTY */%>
					XML.CreateActiveTag("NEWPROPERTY");
					XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
					XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
					XML.CreateTag("VALUATIONTYPE", m_sValuationType);
					XML.CreateTag("PROPERTYLOCATION", m_sPropertyLocation);

					<%/* SHAREOWNERSHIPDETAILS 
					XML.CreateActiveTag("SHAREDOWNERSHIPDETAILS");
					XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
					XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
					XML.CreateTag("SHAREDAMOUNT", "");
					XML.CreateTag("SHAREDOWNERSHIPINDICATOR", "");
					XML.CreateTag("SHAREDPERCENTAGE","");
					XML.CreateTag("SHAREDOWNERSHIPTYPE", "");*/%>
		
					switch (ScreenRules())
					{
						case 1: // Warning
						case 0: // OK
						XML.RunASP(document,"CreateNewPropertyDetails.asp");
							break;
						default: // Error
							XML.SetErrorResponse();
					}
	
					bSuccess = XML.IsResponseOK();
				}
				
				<% /* Set up a default value of 0 for the Purchase Price in Application Fact Find */ %>
				UpdatePurchasePrice("0");
			}
		}
	}
	
	<% /* If there is no mortgage sub-quote, create one */ %>
	if (bSuccess == true)
	{
		if (m_sMortgageSubquoteNumber == "" || m_sMortgageSubquoteNumber == "0")
		{
			if (m_sReadOnly != "1") 
			{
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
				XML.CreateRequestTag(window,"CREATE");
				XML.CreateActiveTag("LIFESUBQUOTES");
				XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
				XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
				XML.CreateTag("QUOTATIONTYPE", "2");
				XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idQuotationNumber","1"));
		
				if(m_bScreenEntry)
					XML.RunASP(document, "AQCreateFirstMortgageLifeSubquote.asp");
				else
				{
					switch (ScreenRules())
					{
						case 1: // Warning
						case 0: // OK
							XML.RunASP(document, "AQCreateFirstMortgageLifeSubquote.asp");
							break;
						default: // Error
							XML.SetErrorResponse();
					}
				}

				bSuccess = XML.IsResponseOK();
				if(bSuccess == true)
				{
					XML.SelectTag(null, "RESPONSE");
					m_sMortgageSubquoteNumber = XML.GetTagText("MORTGAGESUBQUOTENUMBER");
					m_sLifeSubquoteNumber = XML.GetTagText("LIFESUBQUOTENUMBER");
					scScreenFunctions.SetContextParameter(window,"idMortgageSubQuoteNumber",m_sMortgageSubquoteNumber);
					scScreenFunctions.SetContextParameter(window,"idLifeSubquoteNumber",m_sLifeSubquoteNumber);
					
					<% /* Disable the appropriate fields and buttons */ %>
					scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTotalLoanAmt");
					scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtLTVPercent");
					scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTotalMnthCost");
					scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtMnthCostLessDD");
					scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDrawDown");

					frmScreen.btnEdit.disabled = true;
					frmScreen.btnDelete.disabled = true;
					frmScreen.btnOneOffCosts.disabled = true;
					frmScreen.btnNewQuote.disabled = true;
					frmScreen.btnCalculate.disabled = true;
	
					<% /* Initialise the subQuoteXML to be used in Add 
						  Cannot use GetLoanComposition as some data (eg Purchase Price) is not available yet */ %>
					subQuoteXML = null;
					subQuoteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					subQuoteXML.CreateRequestTag(window,null);
					subQuoteXML.CreateActiveTag("MORTGAGESUBQUOTE");
					subQuoteXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
					subQuoteXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
					subQuoteXML.CreateTag("APPLICATIONDATE", scScreenFunctions.GetContextParameter(window,"idApplicationDate","")); 
					subQuoteXML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
					subQuoteXML.CreateTag("AMOUNTREQUESTED", "");
					subQuoteXML.CreateTag("LTV", "");
					subQuoteXML.CreateTag("DRAWDOWN", "");
					
					
				}
				else
				{
					alert("Failed to create first subquote");
					//scScreenFunctions.MSGAlert("Failed to create first subquote");
				}
			}
		}
		else
		{
			<% /* Retrieve Loan Composition details */ %>
			subQuoteXML = null;
			subQuoteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			subQuoteXML.CreateRequestTag(window,null);
			subQuoteXML.CreateActiveTag("LOANCOMPOSITION");
			subQuoteXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			subQuoteXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
			subQuoteXML.CreateTag("APPLICATIONDATE", scScreenFunctions.GetContextParameter(window,"idApplicationDate","")); 
			subQuoteXML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
		
			subQuoteXML.CreateTag("QUOTATIONTYPE", "2");
			if(m_bScreenEntry)
				subQuoteXML.RunASP(document, "AQGetLoanCompositionDetails.asp");
			else
			{
				switch (ScreenRules())
				{
					case 1: // Warning
					case 0: // OK
						subQuoteXML.RunASP(document, "AQGetLoanCompositionDetails.asp");
						break;
					default: // Error
						subQuoteXML.SetErrorResponse();
				}
			}
			
			bSuccess = subQuoteXML.IsResponseOK();		
			if(bSuccess == true) 
			{
				<%/* Record found */%>
				if (LoanPropertyErrorReturn[0] == true)
				{
					LoanPropertyXML.SelectTag(null, "LOANPROPERTYDETAILS")
					mDiscountAmount  = LoanPropertyXML.GetTagText("DISCOUNTAMOUNT");
				}

				PopulateScreen();
			
				m_iManualIncentive = subQuoteXML.GetTagText("MANUALINCENTIVEAMOUNT");
			
				if (m_iManualIncentive != "" && m_iManualIncentive > 0 && m_PopWinInd == "No")
				{	
					alert('Please review the manual incentive amount in one-off costs');
					// scScreenFunctions.MSGAlert("Please review the manual incentive amount in one-off costs");
				}
			
				m_iInterestOnlyAmount = subQuoteXML.GetTagText("INTERESTONLYELEMENT");
				m_iCapitalInterestAmount = subQuoteXML.GetTagText("CAPITALANDINTERESTELEMENT");
			}
		}
	}
	
	<% /* PSC 02/02/2007 EP2_1113 */ %> 
	ResetGenerateQuote();
	
	if (bSuccess == false)
	{
		scScreenFunctions.SetCollectionToReadOnly(divBackground);
		DisableButtons();
	}
	<% /* PJO 10/03/2006 MAR1378 - disable buttons in read only */ %>
	if (m_blnReadOnly == true)
	{
		DisableButtons();
	}
	
	var specialScheme = "";
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	<%/* EP2_1614 Start*/%>
	m_FlexibleMortgageInd = "0";
	subQuoteXML.SelectTag(null,"MORTGAGEPRODUCT");
	if(subQuoteXML.SelectSingleNode("/RESPONSE/MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT/MORTGAGEPRODUCTDETAILS/MORTGAGEPRODUCT/FLEXIBLEMORTGAGEPRODUCT"))
	{
		m_FlexibleMortgageInd = subQuoteXML.SelectSingleNode("/RESPONSE/MORTGAGESUBQUOTE/LOANCOMPONENTLIST/LOANCOMPONENT/MORTGAGEPRODUCTDETAILS/MORTGAGEPRODUCT/FLEXIBLEMORTGAGEPRODUCT").text;
	}
	<%/* EP2_1614 End*/%>
	subQuoteXML.SelectTag(null,"APPLICATIONFACTFIND");
	if(subQuoteXML.SelectSingleNode("/RESPONSE/APPLICATIONFACTFIND/SPECIALSCHEME"))
	{
		specialScheme = subQuoteXML.SelectSingleNode("/RESPONSE/APPLICATIONFACTFIND/SPECIALSCHEME").text;
	}

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	<% /* PSC 12/02/2007 EP2_1314 - Start */ %>
	m_blnIsPort = XML.IsInComboValidationList(document,"TypeOfMortgage", sTypeOfMortgage, Array("NP"));
	m_blnIsSwitch = XML.IsInComboValidationList(document,"TypeOfMortgage", sTypeOfMortgage, Array("PSW"));
	m_blnIsCLI = XML.IsInComboValidationList(document,"TypeOfMortgage", sTypeOfMortgage, Array("CLI")); 
	<% /* PSC 12/02/2007 EP2_1314 - End */ %>
	
	if( (sTypeOfMortgage != "" && XML.IsInComboValidationList(document,"TypeOfMortgage", sTypeOfMortgage, Array("RTB"))) 
	|| (specialScheme != "" && XML.IsInComboValidationList(document,"SpecialSchemes", specialScheme, Array("RTB"))))
	{
		<% // populate Right To Buy Discount %>
		frmScreen.txtRightToBuyDisc.setAttribute("required", "true");
		frmScreen.txtRightToBuyDisc.value = mDiscountAmount;
	}
	else
	{
		<% // Hide Right To Buy Discount %>
		scScreenFunctions.HideCollection(spnRTBDiscount);
		frmScreen.txtRightToBuyDisc.setAttribute("required", "false");
	}

}
function PopulateCombos()
{
	<% /* Remove Level Of Advice Add TypeOfMortgage & SpecialSchemes */ %>
	
	var sGroupList = new Array("RepaymentType","TypeOfMortgage","SpecialSchemes");
	//var XMLLevelOfAdvice = null;

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	if(XML.GetComboLists(document,sGroupList) == true)
	{
		m_XMLRepay       = XML.GetComboListXML("RepaymentType");
		<% /*EP2_780 */ %>
		m_XMLTypeOfMortgage = XML.GetComboListXML("TypeOfMortgage");
		m_XMLSpecialSchemes = XML.GetComboListXML("SpecialSchemes");
	}
}

function DisableButtons()
{
<%	/* When something has gone wrong, disable all buttons */%>  
	frmScreen.btnAdd.disabled = true;
	frmScreen.btnCalculate.disabled = true;
	frmScreen.btnDelete.disabled = true;
	frmScreen.btnEdit.disabled = true;
	frmScreen.btnNewQuote.disabled = true;
	frmScreen.btnOneOffCosts.disabled = true;
	frmScreen.btnGenerateQuote.disabled = true;  // EP2_55

}
function PopulateScreen()
{
	<% /* PSC 12/02/2007 EP2_1314 - Start */ %>
	subQuoteXML.SelectTag(null,"MORTGAGESUBQUOTE");
	
	<% // MAR1061 - Peter Edney - 13/03/2006 %>
	var m_sPurchasePrice = subQuoteXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE");
	
	if(m_sPurchasePrice == "" || m_sPurchasePrice == "0")
	{
		<% //m_sPurchasePrice = subQuoteXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE"); %>
		subQuoteXML.SelectTag(null,"APPLICATIONFACTFIND");
		m_sPurchasePrice = subQuoteXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE");	
	}
	<% /* PSC 12/02/2007 EP2_1314 - End */ %>
	
	<% /* If the default value of 0 is held in the Purchase Price, display nothing */ %>
	if (m_sPurchasePrice == "0")
		frmScreen.txtPurchasePrice.value = "";
	else
		frmScreen.txtPurchasePrice.value = m_sPurchasePrice;

	subQuoteXML.SelectTag(null,"MORTGAGESUBQUOTE");
	m_sAmtRequested = subQuoteXML.GetTagText("AMOUNTREQUESTED");
	<% /* Assign Drawdown to variable */ %>
	m_sDrawDown = subQuoteXML.GetTagText("DRAWDOWN");
	
	if(m_sAmtRequested != "0")
		frmScreen.txtAmtRequested.value = m_sAmtRequested;
	else if(m_sAmountRequested != "0")
	{
		frmScreen.txtAmtRequested.value = m_sAmountRequested;
		m_sAmtRequested = m_sAmountRequested;
	}
	
	frmScreen.txtTotalLoanAmt.value = subQuoteXML.GetTagText("TOTALLOANAMOUNT");
	
	if(subQuoteXML.GetTagText("LTV") != "0")
		frmScreen.txtLTVPercent.value = top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("LTV"),2);
	
	frmScreen.txtTotalMnthCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("TOTALGROSSMONTHLYCOST"), 2);
	frmScreen.txtMnthCostLessDD.value = top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("MONTHLYCOSTLESSDRAWDOWN"), 2);
	
	m_bPPPresent = false;
	PopulateListBox(0);
	
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTotalLoanAmt");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtLTVPercent");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTotalMnthCost");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtMnthCostLessDD");
	if(scScrollTable.getRowSelected() == -1)
	{
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
	}
	else
	{
		frmScreen.btnEdit.disabled = false;
		frmScreen.btnDelete.disabled = false;
		
	}
	subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE");
	
	<% /* Set DrawDown Field read/write property
		  Disable if Part and part present and default to zero */ %>
	if ((subQuoteXML.GetTagText("FLEXIBLEMORTGAGEPRODUCT").length == 0) || 
		(subQuoteXML.GetTagText("FLEXIBLEMORTGAGEPRODUCT") == "0") ||
		m_bPPPresent)
	{
		frmScreen.txtDrawDown.value = "0";
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDrawDown");
	}
	else
	{
		frmScreen.txtDrawDown.value = m_sDrawDown;
		scScreenFunctions.SetFieldState(frmScreen, "txtDrawDown", "W");
	}
	
	if(subQuoteXML.GetTagText("TOTALNETMONTHLYCOST") == "" || subQuoteXML.GetTagText("TOTALNETMONTHLYCOST") == "0")
	{
		frmScreen.btnOneOffCosts.disabled = true;
		frmScreen.btnNewQuote.disabled = true;
		if(subQuoteXML.GetTagText("TOTALLOANAMOUNT") != subQuoteXML.GetTagText("AMOUNTREQUESTED"))
		{
			frmScreen.btnCalculate.disabled = true;
		}	
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtPurchasePrice");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAmtRequested");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtRightToBuyDisc");<%/*EP2_1730*/%>
		<% /* Set drawdown field to read only */ %>
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDrawDown");
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnCalculate.disabled = true;
	}	
		
	if(m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnCalculate.disabled = true;
		frmScreen.btnNewQuote.disabled = true;
		frmScreen.btnGenerateQuote.disabled = true;  // EP2_55
	}
	
	<%/*Begin: DS - EP2_1943*/%>
	var sMortgType ;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	
	if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW')))
		sMortgType = "PSW";
	else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE')))
		sMortgType = "TOE";
	else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP')))
		sMortgType = "NP";
	if ((sMortgType != null)&&(sMortgType == 'TOE')||(sMortgType == 'PSW')||(sMortgType == 'NP'))
	{
		frmScreen.btnNewQuote.disabled = true;
		if(m_iNumOfLoanComponents == 0)
		{
			frmScreen.btnAdd.disabled = true;
		}
	}
	<%/* End: DS - EP2_1943*/%>

}


function PopulateListBox(nStart)
{
<%  /* Populate the listbox with values from subQuoteXML */
%>	subQuoteXML.ActiveTag = null;
	subQuoteXML.CreateTagList("LOANCOMPONENT");
	m_iNumOfLoanComponents = subQuoteXML.ActiveTagList.length;
	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, subQuoteXML.ActiveTagList.length);
	ShowList(nStart);
	if(m_iNumOfLoanComponents > 0)
	{ scScrollTable.setRowSelected(1);	frmScreen.btnCalculate.disabled = false;}
	else
	{frmScreen.btnCalculate.disabled = true;}
}

function ShowList(nStart)
{
	scScrollTable.clear();
	for (var iCount = 0; iCount < subQuoteXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		subQuoteXML.SelectTagListItem(iCount + nStart);
		
		<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),subQuoteXML.GetTagText("PRODUCTNAME"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("RESOLVEDRATE"), 2));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(6),subQuoteXML.GetTagText("TOTALLOANCOMPONENTAMOUNT"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(7),GetAbbrevRepayMethod(subQuoteXML.GetTagText("REPAYMENTMETHOD")));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(8),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("NETMONTHLYCOST"), 2));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(9),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("APR"), 2));
				
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(10),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("FINALRATEMONTHLYCOST"), 2));
		<% /* PSC 02/02/2007 EP2_1113 - End */ %> 

		// EP2_55 - Add Indicator
		var sIndicator = ReturnIndicator(subQuoteXML.GetTagBoolean("PRODUCTSWITCHRETAINPRODUCTIND"),subQuoteXML.GetTagBoolean("MANUALPORTEDLOANIND"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sIndicator);

		if (subQuoteXML.GetTagText("PORTEDLOAN") == "0" || subQuoteXML.GetTagText("PORTEDLOAN") == "")
		{
			var arrayReturn = GetProductAndRate()
			if( arrayReturn != null)
			{
				<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
				scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),arrayReturn[0]);
				scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),arrayReturn[1]);
				<% /* PSC 02/02/2007 EP2_1113 - End */ %> 
			}
		}
		
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}	
}
function GetAbbrevRepayMethod(sRepayMethod)
{
<%	/* to condense the information in the listbox, output an abbreviated string */
%>	var sAbbrev;
	
	var xmlNode = null;
	var sSearch = "LISTNAME/LISTENTRY[VALUEID='" + sRepayMethod + "' and VALIDATIONTYPELIST/VALIDATIONTYPE='I']";	
	xmlNode = m_XMLRepay.selectSingleNode(sSearch)
	
	<% /* Interest Only */ %>
	if (xmlNode != null)
		sAbbrev = "I/O";	
	else
	{
		sSearch = "LISTNAME/LISTENTRY[VALUEID='" + sRepayMethod + "' and VALIDATIONTYPELIST/VALIDATIONTYPE='C']";	
		xmlNode = m_XMLRepay.selectSingleNode(sSearch)
		<% /* Capital And Interest */ %>
		if (xmlNode != null)
			sAbbrev = "C&I";
		else
		{
			sSearch = "LISTNAME/LISTENTRY[VALUEID='" + sRepayMethod + "' and VALIDATIONTYPELIST/VALIDATIONTYPE='P']";	
			xmlNode = m_XMLRepay.selectSingleNode(sSearch)
			<% /* Part and Part */ %>
			if (xmlNode != null)
			{
				sAbbrev = "P&P";
				m_bPPPresent = true;
			}
			else
				sAbbrev = "??";
		}
	}
	
	return sAbbrev;
}

function GetProductAndRate()
{
<%  /* All the information for this loancomponent is listed under the currently selected 
       loancomponent in subQuoteXML so there is no need to select a tag. We remember which
       taglist was active when we came in so we can reset it on the way out. */
%>  var arrayReturn = new Array(3);           
	var bSuccess = false;
	var currentActiveTagList = subQuoteXML.ActiveTagList;
	var currentActiveTag = subQuoteXML.ActiveTag;
	var strProductType = "";
	 
	<% /* Retrieve the Rate as well as the Step so that it can be used in CM101 */ %> 
	var strRate = "";                        
	var strStep = "";
	
	if(subQuoteXML.GetTagText("NONPANELLENDEROPTION") == "0")
	{
<%		/* get the baserate rate for future reference */
%>		var strBaseRate = "0";
		if(subQuoteXML.SelectTag(currentActiveTag, "BASERATEBAND") != null)
			strBaseRate = subQuoteXML.GetTagText("RATE");
		
		subQuoteXML.ActiveTag = currentActiveTag;
		subQuoteXML.CreateTagList("INTERESTRATETYPE");
		
			subQuoteXML.SelectTagListItem(0);
			strStep = subQuoteXML.GetTagText("INTERESTRATETYPESEQUENCENUMBER");
			
				var strRateType = subQuoteXML.GetTagText("RATETYPE");
				if(strRateType == "F"){
					strProductType = "Fixed";
					strRate = subQuoteXML.GetTagText("RATE");   
				}
				else if(strRateType == "B"){
					strProductType = "Base";
					strRate = strBaseRate;                      
				}
				else if(strRateType == "D"){
					strProductType = "Discounted";
					strRate = parseFloat(strBaseRate) - parseFloat(subQuoteXML.GetTagText("RATE"));  
				}
				else if(strRateType == "C"){
					strProductType = "Capped/Floored";
					strRate = parseFloat(strBaseRate) - parseFloat(subQuoteXML.GetTagText("RATE"));  
					if(parseFloat(strRate) < parseFloat(subQuoteXML.GetTagText("FLOOREDRATE")))
						strRate = subQuoteXML.GetTagText("FLOOREDRATE");
					if(parseFloat(strRate) > parseFloat(subQuoteXML.GetTagText("CEILINGRATE")))
						strRate = subQuoteXML.GetTagText("CEILINGRATE");					
				}
				bSuccess = true;
	}
	else
	{
		strProductType = subQuoteXML.GetTagText("MORTGAGEPRODUCTTYPE");
		strRate = subQuoteXML.GetTagText("INTERESTRATE1");
		bSuccess = true;
	}
	subQuoteXML.ActiveTagList = currentActiveTagList;
	if(bSuccess)
	{
		arrayReturn[0] = strProductType;
		arrayReturn[1] = strStep;
		arrayReturn[2] = top.frames[1].document.all.scMathFunctions.RoundValue(strRate, 2);
		
		return(arrayReturn);
	}
	else return(null);
}

<% /*EP2_780 Update the RTBDiscount on the NewProperty table */ %>
function UpdateRTBDiscount()
{
	var bContinue = true;
	if(frmScreen.txtRightToBuyDisc.value.length > 0)
	{	
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window);
		var xn = XML.XMLDocument.documentElement;
		xn.setAttribute("CRUD_OP","UPDATE");
		xn.setAttribute("SCHEMA_NAME","omCRUD");
		xn.setAttribute("ENTITY_REF","NEWPROPERTY");
		var xe = XML.XMLDocument.createElement("NEWPROPERTY");
		xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		xe.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		xe.setAttribute("DISCOUNTAMOUNT", frmScreen.txtRightToBuyDisc.value);

		xn.appendChild(xe);
		XML.RunASP(document, "omCRUDIf.asp");
		if (!XML.IsResponseOK())
		{	
			alert ("Error creating LoanComponent entries in table");
			bContinue = false;
		}
	}
	return bContinue;
}

function frmScreen.txtAmtRequested.onblur()
{
	<% /* If the amount requested changes (and a purchase price is present) then reset */ %>
	if(frmScreen.txtAmtRequested.value != m_sAmtRequested && 
	   frmScreen.txtAmtRequested.value != "")
	{
		if(frmScreen.txtPurchasePrice.value != "" && parseInt(frmScreen.txtPurchasePrice.value) != 0)
		{
			<% /* PSC 12/02/2007 EP2_1314 - Start */ %>
			if(m_sAmtRequested != "" && m_sAmtRequested != "0" && !m_blnIsPort)
				ResetAmtRequested();
			else
			{
				if (frmScreen.txtAmtRequested.value != "" && frmScreen.txtTotalLoanAmt.value != "")
				{
					if (parseInt(frmScreen.txtAmtRequested.value) < parseInt(frmScreen.txtTotalLoanAmt.value))
					{
						alert("The amount requested cannot be less than the existing mortgage balance"); 
						frmScreen.txtAmtRequested.value = m_sAmtRequested;

						return;
					}
				}
				CalcAndPopulateLTV();
				
				if (m_blnIsPort)
				{
					SaveMortgageSubQuoteDetails();
					Initialise();
				}
			}
			<% /* PSC 12/02/2007 EP2_1314 - End */ %>

			m_sAmtRequested = frmScreen.txtAmtRequested.value;
			<%/* EP2_1730  No point in resetting if other fields are empty*/%>
			if(!m_blnIsPort && (frmScreen.txtRightToBuyDisc.value != "") && (frmScreen.txtPurchasePrice.value != ""))
			{
				Reset();
			}
		}
	}
	
	// EP2_2168 - Check Drawdown < AmtReq
	CheckDrawDownValid();
}
		
function frmScreen.txtPurchasePrice.onblur()
{
	<% /* If the purchase price changes (and an amount requested is present) then reset */ %>
	if(frmScreen.txtPurchasePrice.value != m_sPurchasePrice && 
	   frmScreen.txtPurchasePrice.value != "")
	{
		<% /* PSC 12/02/2007 EP2_1314 - Start */ %>
		var sPurchasePrice = frmScreen.txtPurchasePrice.value;
		var sDiscountAmount = frmScreen.txtRightToBuyDisc.value;
		
		if (sPurchasePrice == "")
			sPurchasePrice = "0";
			
		if (sDiscountAmount == "")
			sDiscountAmount = "0";
			
		if (parseInt(sDiscountAmount) > parseInt(sPurchasePrice))
		{
			alert("The discount amount cannot be greater than the purchase price");
			frmScreen.txtPurchasePrice.value = m_sPurchasePrice;
			frmScreen.txtPurchasePrice.focus();
			return;
		}
		<% /* PSC 12/02/2007 EP2_1314 - End */ %>
	
		if(m_sPurchasePrice != "" && m_sPurchasePrice != "0")ResetPurchasePrice();
		else
		{
			<% /* PSC 12/02/2007 EP2_1314 */ %>
			UpdatePurchasePrice(frmScreen.txtPurchasePrice.value);
			
			if(frmScreen.txtAmtRequested.value != "" && parseInt(frmScreen.txtAmtRequested.value) != 0)
			{
				CalcAndPopulateLTV();
			}
		}
		m_sPurchasePrice = frmScreen.txtPurchasePrice.value;
	}
	
	<% /* PSC 02/02/2007 EP2_1113 */ %> 
	ResetGenerateQuote();
}

function ResetAmtRequested()
{
	<%/* Resets the screen after a change to AmountRequested so that the loan composition
         for the mortgage subquote may be re-established. */ %>
      
	if (confirm("Changing Amount Requested will reset the screen and will clear any loan components already established. Do you wish to continue?") == true)
	{
		<% /* PSC 12/02/2007 EP2_1314 */ %>
		Reset();	
	}
	else  
	{
		frmScreen.txtAmtRequested.value = m_sAmtRequested;
	}
}

function ResetPurchasePrice()
{
	<%/* Resets the screen after a change to Purchase Price so that the loan composition
         for the mortgage subquote may be re-established. */ %>
      
	if (confirm("Changing Purchase Price will reset the screen and will clear any loan components already established. Do you wish to continue?") == true)
	{
		<% /* Save the property value into ApplicationFactFind before resetting the subquote */ %>
		UpdatePurchasePrice(frmScreen.txtPurchasePrice.value);
		<% /* PSC 12/02/2007 EP2_1314 */ %>
		Reset();	
	}
	else  
	{
		<% /* PSC 12/02/2007 EP2_1314 */ %>
		frmScreen.txtPurchasePrice.value = m_sPurchasePrice;
	}
}

function AddCustomerList(XML)
{
	var tagCustomerList = XML.CreateActiveTag("CUSTOMERLIST");
	for(var nCount = 1; nCount < 6; nCount++)
	{
		var sCustomerNumber			= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nCount, "rf1111");
		var sCustomerVersionNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nCount, "1");
		if(sCustomerNumber != "" && sCustomerVersionNumber != "")
		{
			XML.ActiveTag = tagCustomerList;
			XML.CreateActiveTag("CUSTOMER");
			XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
		}
	}
}
function spnTable.onclick()
{
	if (scScrollTable.getRowSelected() != -1)
	{
		if(frmScreen.txtAmtRequested.readOnly == false)
		{
			<%/* Only enable editing if subquote has been calculated */%>
			frmScreen.btnEdit.disabled = false;
			frmScreen.btnDelete.disabled = false;
		}
	}
	else
	{
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
	}
}

function CalcAndPopulateLTV()
{
	var bAmountOK = ((frmScreen.txtAmtRequested.value != "") && (frmScreen.txtAmtRequested.value != 0));
	var bPurchasePriceOK = ((frmScreen.txtPurchasePrice.value != "") && (frmScreen.txtPurchasePrice.value != 0));
	
	if(bAmountOK && bPurchasePriceOK)
	{
		<% /* Save the property value into ApplicationFactFind before calculating LTV */ %>
		UpdatePurchasePrice(frmScreen.txtPurchasePrice.value);

		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("LTV");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		XML.CreateTag("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
		
		AddCustomerList(XML);

		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				XML.RunASP(document, "AQCalcCostModelLTV.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
		}

		if(XML.IsResponseOK())
		{
			XML.SelectTag(null, "RESPONSE"); 
			frmScreen.txtLTVPercent.value = XML.GetTagText("LTV");
		}
	}
}

function GetFloat(sString)
{
<% /* This function returns a float representation of the sString passed in.
      It does full checks for NaN and 0 */
%>
	if( sString.length == 0 ) return 0.0;
	if( isNaN(parseFloat(sString))) return 0.0;
	else
	{	
		return( top.frames[1].document.all.scMathFunctions.RoundValue(sString, 2));
	}
}

function frmScreen.btnAdd.onclick()
{
	var bContinue = true;

	if(subQuoteXML != null)
	{
		if( subQuoteXML.SelectTag(null, "MORTGAGELENDER") != null)
		{
			if(m_iNumOfLoanComponents == parseInt(subQuoteXML.GetTagText("MAXNOLOANS")))
			{
				alert("Maximum number of loan components allowed already created on this mortgage sub-quote");
				//scScreenFunctions.MSGAlert("Maximum number of loan components allowed already created on this mortgage sub-quote");
				bContinue = false;			
			}
		}
		if( bContinue && (parseInt(frmScreen.txtTotalLoanAmt.value) == parseInt(frmScreen.txtAmtRequested.value)) )
		{
			alert("Total of loans already equals amount requested; it is not possible to add another loan component");
			//scScreenFunctions.MSGAlert("Total of loans already equals amount requested; it is not possible to add another loan component");
			bContinue = false;
		}
	}
	if(bContinue == true)
	{
		if(frmScreen.txtPurchasePrice.value == "" || frmScreen.txtPurchasePrice.value == "0")
		{
			alert("Please enter Purchase Price before adding loan components");
			//scScreenFunctions.MSGAlert("Please enter Purchase Price before adding loan components");
			frmScreen.txtPurchasePrice.focus();
		}
		else if(frmScreen.txtAmtRequested.value == "" || frmScreen.txtAmtRequested.value == "0")
		{
			alert("Please enter Amount Requested before adding loan components");
			//scScreenFunctions.MSGAlert("Please enter Amount Requested before adding loan components");
			frmScreen.txtAmtRequested.focus();
		}
		else
		{
<%			/* First make sure there is an LTV present */
%>			if(frmScreen.txtLTVPercent.value == "") CalcAndPopulateLTV();

			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

			if(subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE") != null)
			{
				subQuoteXML.SetTagText("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
				subQuoteXML.SetTagText("LTV", frmScreen.txtLTVPercent.value);
				subQuoteXML.SetTagText("DRAWDOWN", frmScreen.txtDrawDown.value);

				<% /* MAR1452 */ %>
				<% /* EP2_171 GHun
				subQuoteXML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value);	*/ %>
				subQuoteXML.SetTagText("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value);
				<% /* EP2_171 End */ %>
				
				<%/* Update the mortgagesubquote on the database */%>				

				XML.CreateRequestTag(window,null);
				XML.CreateActiveTag("MORTGAGESUBQUOTE");
				XML.AddXMLBlock(subQuoteXML.CreateFragment());
								
				switch (ScreenRules())
				{
					case 1: // Warning
					case 0: // OK
						XML.RunASP(document, "AQUpdateMortgageSubQuote.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
				}

				if(XML.IsResponseOK() != true)
				{
					bContinue = false;
					btnSubmit.focus();
				}
			}
			<% /* EP2_1040 */ %>
			if(!UpdateRTBDiscount())
			{
				bContinue = false;
			}
			if(bContinue)
			{
				CreateMQ020ParamsXML(); // EP2_8 - Create XML with new params
				var sReturn = null;
				var ArrayArguments = new Array(14); //EP2_8 - Redefine. // AS EP2_8 Extra Params.
				
				<% // pct - MAR100 - Creation of Variables required for incentives screen %>
				m_sPurchasePrice = frmScreen.txtPurchasePrice.value;
					
				subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE");
				subQuoteXML.CreateActiveTag("APPLICATIONFACTFIND");					
				subQuoteXML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE", m_sPurchasePrice);
				subQuoteXML.CreateTag("VALUATIONTYPE", m_sValuationType);
				subQuoteXML.CreateTag("PROPERTYLOCATION", m_sPropertyLocation);
				subQuoteXML.CreateTag("TYPEOFAPPLICATION", m_sMortgageType);
				<% // end pct - MAR100 %>
				
				ArrayArguments[0] = m_sApplicationMode;
				ArrayArguments[1] = scScreenFunctions.GetContextParameter(window,"idUserId",null);
				ArrayArguments[2] = scScreenFunctions.GetContextParameter(window,"idUserType",null);
				ArrayArguments[3] = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
				ArrayArguments[4] = m_sReadOnly;
				if(subQuoteXML != null)	ArrayArguments[5] = subQuoteXML.XMLDocument.xml;
				else ArrayArguments[5] = null;
				ArrayArguments[6] = null; //Loan Component sequence number
				ArrayArguments[7] = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","30");
				ArrayArguments[8] = m_iNumOfLoanComponents;
				ArrayArguments[9] = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
				ArrayArguments[10] = XML.CreateRequestAttributeArray(window);
				ArrayArguments[11] = m_sCurrency;
				ArrayArguments[12] = scScreenFunctions.GetContextParameter(window,"idApplicationDate","");  
				ArrayArguments[13] = "Add"  // EP2_8 - Add or Edit - Need mode on MQ020.
				if(MQ010ParamsXML != null)	ArrayArguments[14] = MQ010ParamsXML.XMLDocument.xml;
				else ArrayArguments[14] = null; // EP2_8 - Need on MQ020 
				<% // EP2_697 %>
				
				sReturn = scScreenFunctions.DisplayPopup(window, document, "MQ020.asp", ArrayArguments, 630, 560);
				if(sReturn != null && m_sReadOnly != "1")
				{
					m_PopWinInd = "Yes";
					FlagChange(sReturn[0]);
					Initialise();
					
					<% /* If the total loan amount is now equal to the amount requested, automatically perform Calculation */ %>
					<%/* if(parseFloat(frmScreen.txtTotalLoanAmt.value) == parseFloat(frmScreen.txtAmtRequested.value)) EP2_1614 */%> 
					if((m_FlexibleMortgageInd != "1") && (parseFloat(frmScreen.txtTotalLoanAmt.value) == parseFloat(frmScreen.txtAmtRequested.value))) <%/* EP2_1614 */%> 
					{
						frmScreen.style.cursor = "wait";
						
						<% /* Remove focus from Calculate so that Enter cannot be pressed multiple times */ %>
						frmScreen.btnCalculate.blur();

						<% /* Call Calculate Processing after a timeout to allow the cursor time to change. */ %>
						window.setTimeout(CalculateProcessing, 0)
					}
					
				}
			}
		}		
	}
}

function frmScreen.btnEdit.onclick()
{
<%  /* locate the LOANCOMPONENT XML section for the currently selected loancomponent */
%>  var bContinue = true;
	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	subQuoteXML.ActiveTag = null;			
	subQuoteXML.CreateTagList("LOANCOMPONENT");
	subQuoteXML.SelectTagListItem(nRowSelected -1);
	if(subQuoteXML.GetTagText("PORTEDLOAN") == "1")
				alert("Ported loan components cannot be edited or deleted.");
				//scScreenFunctions.MSGAlert("Ported loan components cannot be edited or deleted.");
	else
	{
		<% /* Update the mortgagesubquote on the database */ %>
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if(subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE") != null)
		{
			subQuoteXML.SetTagText("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
			subQuoteXML.SetTagText("LTV", frmScreen.txtLTVPercent.value);
			<% /* save drawdown amount */ %>
			subQuoteXML.SetTagText("DRAWDOWN", frmScreen.txtDrawDown.value);

			<% /* MAR1452 */ %>
			<% /* EP2_216 GHun
			subQuoteXML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value); */ %>
			subQuoteXML.SetTagText("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value);
			<% /* EP2_216 End */ %>
			
			XML.CreateRequestTag(window,null);
			XML.CreateActiveTag("MORTGAGESUBQUOTE");
			XML.AddXMLBlock(subQuoteXML.CreateFragment());
			
			switch (ScreenRules())
			{
				case 1: // Warning
				case 0: // OK
					XML.RunASP(document, "AQUpdateMortgageSubquote.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
			}

			if(XML.IsResponseOK() != true)
			{
				bContinue = false;
				btnSubmit.focus();
			}
		}
		<% /* EP2_1040 */ %>
		if(!UpdateRTBDiscount())
		{
			bContinue = false;
		}
		if(bContinue)
		{
			CreateMQ020ParamsXML(); // EP2_8 - Create XML with new params
			subQuoteXML.ActiveTag = null;
			subQuoteXML.CreateTagList("LOANCOMPONENT");
			subQuoteXML.SelectTagListItem(nRowSelected -1);
			var sReturn = null;
			var ArrayArguments = new Array(14); //EP2_8 - Redefine.
			ArrayArguments[0] = m_sApplicationMode;
			ArrayArguments[1] = scScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[2] = scScreenFunctions.GetContextParameter(window,"idUserType",null);
			ArrayArguments[3] = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[4] = m_sReadOnly;
			ArrayArguments[5] = subQuoteXML.XMLDocument.xml;
			ArrayArguments[6] = subQuoteXML.GetTagText("LOANCOMPONENTSEQUENCENUMBER");
			ArrayArguments[7] = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","30");
			ArrayArguments[8] = m_iNumOfLoanComponents;
			ArrayArguments[9] = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
			ArrayArguments[10] = XML.CreateRequestAttributeArray(window);
			ArrayArguments[11] = m_sCurrency;
			ArrayArguments[12] = scScreenFunctions.GetContextParameter(window,"idApplicationDate","");  
			ArrayArguments[13] = "Edit"  // EP2_8 - Add or Edit - Need mode on MQ020.
			if(MQ010ParamsXML != null)	ArrayArguments[14] = MQ010ParamsXML.XMLDocument.xml;
			else ArrayArguments[14] = null; // EP2_8 - Need on MQ020 
			sReturn = scScreenFunctions.DisplayPopup(window, document, "MQ020.asp", ArrayArguments, 630, 480);
			
			if(sReturn != null && m_sReadOnly != "1")
			{
				m_PopWinInd = "Yes";
				FlagChange(sReturn[0]);
				Initialise();
				
				<% /* If the total loan amount is now equal to the amount requested, automatically perform Calculation */ %>
				<% /* PSC 12/02/2007 EP2_1314*/ %>
				<% /*if(!m_blnIsSwitch && !m_blnIsCLI && parseFloat(frmScreen.txtTotalLoanAmt.value) == parseFloat(frmScreen.txtAmtRequested.value)) */%> 
				if(!m_blnIsSwitch && !m_blnIsCLI && (m_FlexibleMortgageInd != "1") && parseFloat(frmScreen.txtTotalLoanAmt.value) == parseFloat(frmScreen.txtAmtRequested.value)) <%/* EP2_1614 */%> 
				{
					frmScreen.style.cursor = "wait";

					<% /* Remove focus from Calculate so that Enter cannot be pressed multiple times */ %>
					frmScreen.btnCalculate.blur();

					<% /* Call Calculate Processing after a timeout to allow the cursor time to change. */ %>
					window.setTimeout(CalculateProcessing, 0)						
				}
			}
		}	
	}
}

function frmScreen.btnDelete.onclick()
{
<%  /* locate the LOANCOMPONENT XML section for the currently selected loancomponent */
%>  var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	subQuoteXML.ActiveTag = null;			
	subQuoteXML.CreateTagList("LOANCOMPONENT");
	subQuoteXML.SelectTagListItem(nRowSelected -1);
	if(subQuoteXML.GetTagText("PORTEDLOAN") == "1")
				alert("Ported loan components cannot be edited or deleted.");
				//scScreenFunctions.MSGAlert("Ported loan components cannot be edited or deleted.");
	else
	{
		if(confirm("This will delete the selected loan component. Are you sure?") == true)
		{
			var sLoanAmount = subQuoteXML.GetTagText("LOANAMOUNT");
			var sMnthlyCost = subQuoteXML.GetTagText("NETMONTHLYCOST");
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTag(window,"DELETE");
			XML.CreateActiveTag("MORTGAGESUBQUOTE");
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
			XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
			XML.CreateTag("LOANCOMPONENTSEQUENCENUMBER", subQuoteXML.GetTagText("LOANCOMPONENTSEQUENCENUMBER"));
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document,"DeleteLoanComponent.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

			if(XML.IsResponseOK())
			{
				<%/* Delete this loancomponent from the subQuoteXML and repopulate the listbox.
				  	Also, need to correct the Total Loan Amount field as we are not re-getting
					the mortgage subquote. The DeleteLoanComponent() function should update the
					database with the new total loan amount */ %>
				
				m_PopWinInd = "Yes";
				
				<% /* Reinitialise subQuoteXML */ %>
				Initialise();
				
				frmScreen.btnCalculate.disabled = true;
			}
		}
	}
}

function frmScreen.btnCalculate.onclick()
{
	frmScreen.btnCalculate.style.cursor = "wait";

	<% /* Remove focus from Calculate so that Enter cannot be pressed multiple times */ %>
	frmScreen.btnCalculate.blur();

	<% /* Call Calculate Processing after a timeout to allow the cursor time to change. */ %>
	window.setTimeout(CalculateProcessing, 0)
	
	
	
}

function CalculateProcessing()
{	

	// EP2_1888 - 30/03/2007
	subQuoteXML.SelectTag(null, "APPLICATIONFACTFIND")
	var sLevelOfAdvice = subQuoteXML.GetTagText("LEVELOFADVICE")
	if(sLevelOfAdvice == "")
	{
		alert("Please select 'Application Source' to enter application data before calculating the quotation");
		frmScreen.btnApplicationSrc.focus();
	}
	else
	{	
		var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
		var XMLTemp = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject(); 
		var xmlMainTag
		var blnFurtherAdvance = XMLTemp.IsInComboValidationList(document, "TypeOfMortgage", sTypeOfMortgage, Array('GA'));
		
		XML.CreateRequestTag(window, null);
		if(blnFurtherAdvance) xmlMainTag = XML.CreateActiveTag("REFRESHMORTGAGEACCOUNTDATA");
		else xmlMainTag = XML.CreateActiveTag("GETANDSAVEPORTEDSTEPANDPERIODFROMMORTGAGEACCOUNT"); 
			
		if(blnFurtherAdvance) XML.CreateTag("PORTINGINDICATOR", "0");
		XML.CreateActiveTag("APPLICATION");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		XML.CreateTag("TYPEOFAPPLICATION", scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue",null));
		XML.CreateTag("BMACCOUNTNUMBER", scScreenFunctions.GetContextParameter(window,"idOtherSystemAccountNumber",null));
		
		XML.ActiveTag = xmlMainTag;	
		AddOtherSystemCustomerNumbers(XML)
		
		XML.ActiveTag = xmlMainTag ;
		XML.CreateActiveTag("MORTGAGESUBQUOTE");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
		XML.CreateTag("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
		XML.CreateTag("LTV", frmScreen.txtLTVPercent.value);
			
		if(blnFurtherAdvance)
		{ <% /* Refresh MortgageAccount data from Admin system */ %>
			<% /* Set window status as account refresh can take a while. */ %>
			window.status = "Refreshing account data ...";
			XML.RunASP(document,"RefreshMortgageAccountData.asp");
			window.status = "";
			if(!XML.IsResponseOK())
			{
				frmScreen.btnCalculate.style.cursor = "default";
				frmScreen.style.cursor = "default";

				return;
			}
			else
			{
				if(XML.SelectSingleNode("//LTVCHANGED"))
				{
					m_sLTVChanged = XML.ActiveTag.text ;
					scScreenFunctions.SetContextParameter(window,"idLTVChanged", m_sLTVChanged);
					
					XML.ActiveTag = null ;
					if(XML.SelectSingleNode("//LTV"))
					{
						var sLTV = XML.ActiveTag.text ;
						frmScreen.txtLTVPercent.value = sLTV;
					}				
				}			
			}
		}
		else
		{
			window.status = "Refreshing account data ..."; 
			XML.RunASP(document,"GetAndSavePortedStepAndPeriodFromMortAcc.asp");
			window.status = "";
			if(!XML.IsResponseOK())
			{
				frmScreen.btnCalculate.style.cursor = "default";
				frmScreen.style.cursor = "default";
				
				return;
			}
		}
		
		SaveMortgageSubQuoteDetails();
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("MORTGAGECOSTS");
		XML.CreateTag("CONTEXT", m_sApplicationMode);
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		XML.CreateTag("APPLICATIONDATE", scScreenFunctions.GetContextParameter(window,"idApplicationDate","")); 
		XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idQuotationNumber","1"));
		XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
		XML.CreateTag("LIFESUBQUOTENUMBER", m_sLifeSubquoteNumber);
		CustomerList(XML);

		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				<% /* Set window status as calculation can take a while. */ %>
				window.status = "Calculating costs ...";
				XML.RunASP(document, "AQCalculateMortgageCosts.asp");
				window.status = "";
				break;
			default: // Error
				XML.SetErrorResponse();
		}

		if(XML.IsResponseOK())
		{
			m_PopWinInd = "Yes";
			Initialise();
						
			scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAmtRequested");
			scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtPurchasePrice");
			scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtRightToBuyDisc"); <%/* EP2_1730*/%>

			frmScreen.btnOneOffCosts.disabled = false;
			frmScreen.btnNewQuote.disabled = false;
			frmScreen.btnAdd.disabled = true;
			frmScreen.btnEdit.disabled = true;
			frmScreen.btnDelete.disabled = true;
			frmScreen.btnCalculate.disabled = true;
			RunIncomeCalculations();
		}
		
		<%/*Begin: DS - EP2_1943*/%>
		var sMortgType ;
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
		
		if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW')))
			sMortgType = "PSW";
		else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE')))
			sMortgType = "TOE";
		else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP')))
			sMortgType = "NP";
		if ((sMortgType != null)&&(sMortgType == 'TOE')||(sMortgType == 'PSW')||(sMortgType == 'NP'))
		{
			frmScreen.btnNewQuote.disabled = true;
			if(m_iNumOfLoanComponents == 0)
			{
				frmScreen.btnAdd.disabled = true;
			}
		}
		<%/* End: DS - EP2_1943*/%>
		XML = null;
	}

	frmScreen.btnCalculate.style.cursor = "default";
	frmScreen.style.cursor = "default";	
}

function AddOtherSystemCustomerNumbers(XML)
{
	var sCustomerNumber, sCustomerVersionNumber, sCustomerRoleType, sOtherSystemCustomerNumber ;
	
	var ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(window,"SEARCH");
	ListXML.CreateActiveTag("APPLICATION");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	ListXML.RunASP(document,"FindCustomersForApplication.asp");

	if (ListXML.IsResponseOK())
	{
		var tagCustomerList = XML.CreateActiveTag("CUSTOMERLIST");
		ListXML.CreateTagList("CUSTOMERROLE");
		var iNoOfCustomers = ListXML.ActiveTagList.length;
		for(var nCount = 0; nCount < iNoOfCustomers; nCount++)
		{	
			ListXML.SelectTagListItem(nCount);
			sCustomerNumber = ListXML.GetTagText("CUSTOMERNUMBER");
			sCustomerVersionNumber = ListXML.GetTagText("CUSTOMERVERSIONNUMBER");
			sCustomerRoleType = ListXML.GetTagText("CUSTOMERROLETYPE");
			sOtherSystemCustomerNumber = ListXML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER");
			if(sCustomerNumber != "" && sCustomerVersionNumber != "" && sCustomerRoleType == "1")
			{
				XML.ActiveTag = tagCustomerList;
				XML.CreateActiveTag("CUSTOMER");
				XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
				XML.CreateTag("OTHERSYSTEMCUSTOMERNUMBER", sOtherSystemCustomerNumber);
			}
		}
	}
}

function frmScreen.btnNewQuote.onclick()
{
	<% /* Calculate and save the quotation record */ %>
	if (UpdateQuotation())
	{
		<% /* Create a new Quotation */ %>
		if(CreateQuotation())
		{
			<% /* Create a new Mortgage Sub Quote */ %>
			CreateNewMortgageSubQuote();
		}
	} 	
}

function CreateNewMortgageSubQuote()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idQuotationNumber","1"));
	XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
	XML.CreateTag("LIFESUBQUOTENUMBER", m_sLifeSubquoteNumber);
	XML.CreateTag("QUOTATIONTYPE", "2");
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "CreateNewMortgageLifeSubquote.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if(XML.IsResponseOK())
	{
		XML.SelectTag(null, "MORTGAGESUBQUOTE");
		m_sMortgageSubquoteNumber = XML.GetTagText("MORTGAGESUBQUOTENUMBER");
		m_sLifeSubquoteNumber = XML.GetTagText("LIFESUBQUOTENUMBER");
		scScreenFunctions.SetContextParameter(window,"idMortgageSubQuoteNumber",m_sMortgageSubquoteNumber);
		scScreenFunctions.SetContextParameter(window,"idLifeSubquoteNumber",m_sLifeSubquoteNumber);
		
		m_PopWinInd = "Yes";
		Initialise();
		
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtAmtRequested");
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtPurchasePrice");
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtRightToBuyDisc"); <%/*EP2_1730*/%>
		<% /* new quote so empty monthly cost less drawdown */ %>
		frmScreen.txtMnthCostLessDD.value = "";
		frmScreen.btnAdd.disabled = false;
		frmScreen.btnCalculate.disabled = false;
		if(scScrollTable.getRowSelected() == -1)
		{
			frmScreen.btnEdit.disabled = true;
			frmScreen.btnDelete.disabled = true;
		}
		frmScreen.btnOneOffCosts.disabled = true;
		frmScreen.btnNewQuote.disabled = true;
	}
	XML = null; 
}

function frmScreen.btnOneOffCosts.onclick()
{
	var sReturn = null;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sApplicationType;

	var ArrayArguments = new Array(12);
	ArrayArguments[0] = m_sApplicationMode;
	ArrayArguments[1] = m_sApplicationNumber;
	ArrayArguments[2] = m_sApplicationFFNumber;
	ArrayArguments[3] = m_sMortgageSubquoteNumber;
	ArrayArguments[4] = m_sLifeSubquoteNumber;
	ArrayArguments[5] = m_sReadOnly;
	ArrayArguments[6] = XML.CreateRequestAttributeArray(window);
	ArrayArguments[7] = m_sCurrency;
	XML.ResetXMLDocument();
	AddCustomerList(XML);
	ArrayArguments[8] = XML.XMLDocument.xml;
	ArrayArguments[9] = m_iInterestOnlyAmount;
	ArrayArguments[10] = m_iCapitalInterestAmount;
	ArrayArguments[11] = frmScreen.txtDrawDown.value;
	
	<% /* Get Application type from context */ %>
	sApplicationType = scScreenFunctions.GetContextParameter(window, "idMortgageApplicationValue", null);
	
	ArrayArguments[12] = sApplicationType;
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "CM130.asp", ArrayArguments, 425, 550);

	if(sReturn != null)
	{
		m_PopWinInd = "Yes";
		FlagChange(sReturn[0]);
		Initialise();
	}
	XML = null; 
	
}

function SaveMortgageSubQuoteDetails()
{
	<% /* First make sure there is an LTV present */ %>
	if(frmScreen.txtLTVPercent.value == "") CalcAndPopulateLTV();

	<% /* Update the mortgagesubquote on the database */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if(subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE") != null)
	{
		subQuoteXML.SetTagText("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
		subQuoteXML.SetTagText("LTV", frmScreen.txtLTVPercent.value);
		<% /* Save drawdown amount */ %>
		subQuoteXML.SetTagText("DRAWDOWN", frmScreen.txtDrawDown.value);
		subQuoteXML.SetTagText("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value);	<% /* EP2_216 GHun */ %>
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("MORTGAGESUBQUOTE");
		XML.AddXMLBlock(subQuoteXML.CreateFragment());
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
				XML.RunASP(document, "AQUpdateMortgageSubQuote.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}
		
		if(XML.IsResponseOK() != true)
		{
			bContinue = false;
			btnSubmit.focus();
			return false;
		}
	}
	return true;

}
function btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
	{
		<% /* EP2_780 */ %>
		if(UpdateRTBDiscount())
	    {
			if (parseInt(frmScreen.txtTotalMnthCost.value) > 0)
			{
				<% /* Calculation has been performed - update quotation */ %>
				UpdateQuotation();
			}
			
			scScreenFunctions.SetContextParameter(window,"idNoCompleteCheck","1");
			scScreenFunctions.SetContextParameter(window, "idCallingScreenID", "MQ010");

			if(scScreenFunctions.CompletenessCheckRouting(window))
				frmToGN300.submit();				
			else 
				frmToCM065.submit();   // Stored Quotes
		}
	}
}


function CustomerList(XML)
{
	var tagCustomerList = XML.CreateActiveTag("CUSTOMERLIST");
	var nCustCount = 0;
	for(var nCount = 1; nCount < 6; nCount++)
	{
		if(nCustCount < "2")
		{
			var sCustomerNumber			= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nCount, "rf1111");
			var sCustomerVersionNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nCount, "1");
			var sCustomerRoleType		= scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nCount, "1");
			if(sCustomerNumber != "" && sCustomerVersionNumber != "" && sCustomerRoleType == "1")
			{
				XML.ActiveTag = tagCustomerList;
				XML.CreateActiveTag("CUSTOMER");
				XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
				nCustCount = nCustCount + 1;
			}
		}
	}
}

function btnCancel.onclick()
{
	<% /* If cancel clicked ensure that completeness check is not re-run in cm010. */ %>
	scScreenFunctions.SetContextParameter(window,"idNoCompleteCheck","1");
	scScreenFunctions.SetContextParameter(window,"idMetaAction","MQ010");
	
	scScreenFunctions.SetContextParameter(window, "idCallingScreenID", "MQ010");
	
	frmToMN060.submit();
}

// EP2_2168 - New Method - Check Drawdown < AmtReq
function frmScreen.txtDrawDown.onblur()
{
	CheckDrawDownValid();
}

// EP2_2168 - New Method - Check Drawdown < AmtReq
function CheckDrawDownValid()
{
	// Drawdown is present.
	if(frmScreen.txtDrawDown.value != "" && frmScreen.txtDrawDown.value != "0")
	{
		// Amount Requested is present.
		if(frmScreen.txtAmtRequested.value != "" && frmScreen.txtAmtRequested.value != "0")
		{
			// Is Drawdown > Amount Requested - This is a NoNo!
			if (parseInt(frmScreen.txtDrawDown.value) > parseInt(frmScreen.txtAmtRequested.value))
			{
				alert("The Drawdown amount cannot be greater than the amount requested");
				frmScreen.txtDrawDown.focus();
			}
		}
	}
}

function RunIncomeCalculations()
{
	var AllowableIncXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AllowableIncXML.CreateRequestTag(window,null);
	<% /* MAR30 */ %> 
	AllowableIncXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));
		
	<% /* Set up request for Income Calculation */ %>
	var TagRequest = AllowableIncXML.CreateActiveTag("INCOMECALCULATION");
	AllowableIncXML.SetAttribute("CALCULATEMAXBORROWING","1");
	AllowableIncXML.SetAttribute("QUICKQUOTEMAXBORROWING","0");
	AllowableIncXML.CreateActiveTag("APPLICATION");
	AllowableIncXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AllowableIncXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	AllowableIncXML.ActiveTag = TagRequest;
	
	var TagCUSTOMERLIST = AllowableIncXML.CreateActiveTag("CUSTOMERLIST");

	for(var nLoop = 1; nLoop <= 2; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		<% /* The customer must exist */ %>
		if (sCustomerNumber.trim().length > 0)
		{
			var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
			var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);
			var sCustomerOrder = scScreenFunctions.GetContextParameter(window,"idCustomerOrder" + nLoop);

			AllowableIncXML.CreateActiveTag("CUSTOMER");
			AllowableIncXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			AllowableIncXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
			AllowableIncXML.CreateTag("CUSTOMERROLETYPE", sCustomerRoleType);
			AllowableIncXML.CreateTag("CUSTOMERORDER", sCustomerOrder);
			AllowableIncXML.ActiveTag = TagCUSTOMERLIST;
		}
	}
	
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			AllowableIncXML.RunASP(document,"RunIncomeCalculations.asp");
			break;
		default: // Error
			AllowableIncXML.SetErrorResponse();
	}
	
	if(AllowableIncXML.IsResponseOK())
	{
		return(true);
	}

	return(false);
}

function UpdateActiveQuoteNumber(sQuotationNumber)
{
	var ApplicationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	ApplicationXML.CreateRequestTag(window,"UPDATE");
	ApplicationXML.CreateActiveTag("APPLICATIONFACTFIND");
	ApplicationXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ApplicationXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFFNumber);
	ApplicationXML.CreateTag("ACTIVEQUOTENUMBER",sQuotationNumber);
	ApplicationXML.CreateTag("ACCEPTEDQUOTENUMBER","");
	
	ApplicationXML.RunASP(document,"UpdateApplicationFactFind.asp");
	ApplicationXML.IsResponseOK();
	scScreenFunctions.SetContextParameter(window,"idQuotationNumber",sQuotationNumber);
	m_sActiveQuoteNumber = sQuotationNumber;
}

function UpdatePurchasePrice(sPurchasePrice)
{
	var ApplicationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	ApplicationXML.CreateRequestTag(window,"UPDATE");
	ApplicationXML.CreateActiveTag("APPLICATIONFACTFIND");
	ApplicationXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ApplicationXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFFNumber);
	ApplicationXML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE",sPurchasePrice);
	
	ApplicationXML.RunASP(document,"UpdateApplicationFactFind.asp");
	ApplicationXML.IsResponseOK();
}

function UpdateQuotation()
{
	var dblTotalQuotationCost = 0;
	var sActiveBCSubQuoteNumber = scScreenFunctions.GetContextParameter(window,"idBCSubQuoteNumber",null);
	var sActivePPSubQuoteNumber = scScreenFunctions.GetContextParameter(window,"idPPSubQuoteNumber",null);
	var sActiveQuoteNumber = scScreenFunctions.GetContextParameter(window,"idQuotationNumber",null);
		
	if (m_sMortgageSubquoteNumber == "" || frmScreen.txtTotalMnthCost.value == "")
	{
		alert("This quotation cannot be calculated without a completed mortgage sub-quote.");
		//scScreenFunctions.MSGAlert("This quotation cannot be calculated without a completed mortgage sub-quote.");
		return
	}
	
	dblTotalQuotationCost = parseFloat(frmScreen.txtTotalMnthCost.value);
				
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "UPDATE");
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	XML.CreateTag("QUOTATIONNUMBER", sActiveQuoteNumber);
	XML.CreateTag("TOTALQUOTATIONCOST", dblTotalQuotationCost.toString());
	XML.CreateTag("LIFESUBQUOTENUMBER", "");
	XML.CreateTag("BCSUBQUOTENUMBER", sActiveBCSubQuoteNumber);
	XML.CreateTag("PPSUBQUOTENUMBER", sActivePPSubQuoteNumber);
	XML.CreateTag("LIFESUBQUOTEREQUIRED", "0");
		
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
				XML.RunASP(document, "StoreQuotation.asp");
		break;
		default: // Error
			XML.SetErrorResponse();
		}

	if(XML.IsResponseOK())
	{
		return(true);
	}

	return(false);	
}

function CreateQuotation()
{
		<% /* Create a new Quotation */ %>

		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window,"CREATE");
		XML.CreateActiveTag("QUOTATION");
		XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFFNumber);
		XML.CreateTag("QUOTATIONNUMBER",m_sActiveQuoteNumber);

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
		UpdateActiveQuoteNumber(sQuotationNumber);	
	
		return(true);
	}

	return(false);	
}
<% /* EP2_8 - New method */ %>
function frmScreen.btnApplicationSrc.onclick()
{
	if(frmScreen.onsubmit())
	{
		<% /* EP2_780 */ %>
		if(UpdateRTBDiscount())
	    {
			if (parseInt(frmScreen.txtTotalMnthCost.value) > 0)
			{
				<% /* Calculation has been performed - update quotation */ %>
				UpdateQuotation();
			}

			//Set Calling screen params
			scScreenFunctions.SetContextParameter(window,"idNoCompleteCheck","1");
			scScreenFunctions.SetContextParameter(window, "idCallingScreenID", "MQ010");
			// Goto DC010 form.
			frmToDC010.submit();				
		}
	}
}

<%/* EP2_8 - New function.*/%>
function CreateMQ020ParamsXML()
{
	// Create doc basics
	MQ010ParamsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	MQ010ParamsXML.CreateActiveTag("PARAMS");
	
	// NATUREOFLOAN, CREDITSCHEME and APPLICATIONINCOMESTATUS flags.
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window);
	var xn = XML.XMLDocument.documentElement;
	xn.setAttribute("CRUD_OP","READ");
	xn.setAttribute("SCHEMA_NAME","omCRUD");
	xn.setAttribute("ENTITY_REF","APPLICATIONFACTFIND");
	var xe = XML.XMLDocument.createElement("APPLICATIONFACTFIND");
	xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	xe.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	xn.appendChild(xe);
	XML.RunASP(document, "omCRUDIf.asp");
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null,"APPLICATIONFACTFIND");
		MQ010ParamsXML.CreateTag("NATUREOFLOAN", XML.GetAttribute("NATUREOFLOAN"));
		MQ010ParamsXML.CreateTag("APPLICATIONINCOMESTATUS", XML.GetAttribute("APPLICATIONINCOMESTATUS"));
		MQ010ParamsXML.CreateTag("CREDITSCHEME", XML.GetAttribute("PRODUCTSCHEME"));
	}
	
	// CREDITSCHEME flag.
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.RunASP(document,"GetApplicationData.asp");
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null,"APPLICATION");
		var AccountGUID = XML.GetTagText("ACCOUNTGUID");
	}
	
	// GUARANTORIND flag.
	var IsThereAGuarantor = "0"
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"CustomerNumber" + nLoop);
		if (sCustomerNumber.trim().length > 0)
		{
			var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"CustomerRoleType" + nLoop);
			if(sCustomerRoleType == "2") // Guarantor
			{
				IsThereAGuarantor = "1";
				break;
			}
		}
		else
		break;   // No value => no more customers.
	}
	MQ010ParamsXML.CreateTag("GUARANTORIND", IsThereAGuarantor);

	// FLEXIBLEPRODUCTS / NONFLEXIBLEPRODUCTS
	
	// Set defaults.
	MQ010ParamsXML.CreateTag("FLEXIBLEPRODUCTS", "-1");
	MQ010ParamsXML.CreateTag("NONFLEXIBLEPRODUCTS", "-1");

	// If we have existing components.
	if (m_iNumOfLoanComponents > 0) 
	{
		// Get first row and find ProductCode
		subQuoteXML.SelectTagListItem(0);
		var ProdCode = subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE");
		// Now get the FLEXIBLEMORTGAGEPRODUCT flag from MortgageProduct table.
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window);
		var xn = XML.XMLDocument.documentElement;
		xn.setAttribute("CRUD_OP","READ");
		xn.setAttribute("SCHEMA_NAME","epsomCRUD");
		xn.setAttribute("ENTITY_REF","MORTGAGEPRODUCT");
		var xe = XML.XMLDocument.createElement("MORTGAGEPRODUCT");
		xe.setAttribute("MORTGAGEPRODUCTCODE", ProdCode);
		xn.appendChild(xe);

		XML.RunASP(document, "omCRUDIf.asp");
		if (XML.IsResponseOK())
		{
			XML.SelectTag(null,"MORTGAGEPRODUCT");
			var FlexibleMortFlag = XML.GetAttribute("FLEXIBLEMORTGAGEPRODUCT")
			// Flexible Products flag.
			MQ010ParamsXML.SetTagText("FLEXIBLEPRODUCTS", FlexibleMortFlag);
			// NONFLEXIBLEPRODUCTS Products flag.
			// Checked with Robin 25Oct06 - If Null then treat as non-flexible. 
			if (FlexibleMortFlag == "1")
			MQ010ParamsXML.SetTagText("NONFLEXIBLEPRODUCTS", "0");
			else
			MQ010ParamsXML.SetTagText("NONFLEXIBLEPRODUCTS", "1");
		}
	}	
	else // NO existing Loan Components
	{
		// If: Is additional lending (use Validation types)
		//	Then: Get Product code from MortgageLoan via application.AccountGUID)
		// Gather round data folks
		var sTypeOfMortgage = scScreenFunctions.GetContextParameter(window,"MORTGAGEAPPLICATIONVALUE");
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject(); 
		var blnValidationF = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('F'));
		var blnValidationM = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('M'));
		var blnValidationTOE = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE'));
		var blnValidationPSW = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW'));
	
		// Now check our conditions F&M or TOE or PSW
		if (((blnValidationF && blnValidationM) || blnValidationTOE || blnValidationPSW) && AccountGUID != "")
		{
			// Use AccountGUID from Application (retrieved above with CREDITSCHEME)
			// to retrieve the MORTGAGEPRODUCTCODE from MortgageLoan table.
			XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTag(window);
			var xn = XML.XMLDocument.documentElement;
			xn.setAttribute("CRUD_OP","READ");
			xn.setAttribute("SCHEMA_NAME","epsomCRUD");
			xn.setAttribute("ENTITY_REF","MORTGAGELOAN");
			var xe = XML.XMLDocument.createElement("MORTGAGELOAN");
			xe.setAttribute("ACCOUNTGUID", AccountGUID);
			xn.appendChild(xe);

			XML.RunASP(document, "omCRUDIf.asp");
			if (XML.IsResponseOK())
			{
				XML.SelectTag(null,"MORTGAGELOAN");
				var ProdCode = XML.GetAttribute("MORTGAGEPRODUCTCODE");
			}
		
			// Now get the FLEXIBLEMORTGAGEPRODUCT flag from MortgageProduct table.
			XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTag(window);
			var xn = XML.XMLDocument.documentElement;
			xn.setAttribute("CRUD_OP","READ");
			xn.setAttribute("SCHEMA_NAME","epsomCRUD");
			xn.setAttribute("ENTITY_REF","MORTGAGEPRODUCT");
			var xe = XML.XMLDocument.createElement("MORTGAGEPRODUCT");
			xe.setAttribute("MORTGAGEPRODUCTCODE", ProdCode);
			xn.appendChild(xe);

			XML.RunASP(document, "omCRUDIf.asp");
			if (XML.IsResponseOK())
			{
				XML.SelectTag(null,"MORTGAGEPRODUCT");
				var FlexibleMortFlag = XML.GetAttribute("FLEXIBLEMORTGAGEPRODUCT")
				// FLEXIBLEPRODUCTS / NONFLEXIBLEPRODUCTS flags.
				// Checked with RobinH 25Oct06 - If Null then treat as non-flexible. 
				if (FlexibleMortFlag == "1")
				{
					MQ010ParamsXML.SetTagText("FLEXIBLEPRODUCTS", "1");
					MQ010ParamsXML.SetTagText("NONFLEXIBLEPRODUCTS", "0");
				}
				else
				{
					MQ010ParamsXML.SetTagText("FLEXIBLEPRODUCTS", "0");
					MQ010ParamsXML.SetTagText("NONFLEXIBLEPRODUCTS", "1");
				}
			}	
		}
	}

}

// EP2_55 - New Method.
function DisableAmountRequested()
{
	// Method to decide whether we are disabling AmountRequested field.
	var lDisable = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	
	// Check PSW
	var bDisablePSW = XML.GetGlobalParameterBoolean(document,"PSWDisableAmtRequested");
	var bIsPSWType = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW'));
	if((bDisablePSW == true) && (bIsPSWType == true))
		lDisable = true;

	// Check TOE
	var bDisableTOE = XML.GetGlobalParameterBoolean(document,"TOEDisableAmtRequested");
	var bIsTOEType = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE'));
	if((bDisableTOE == true) && (bIsTOEType == true))
		lDisable = true;
	
	return lDisable;	
	
}

// EP2_55 - New Method.
<% /* PSC 02/02/2007 EP2_1113 - Start */ %>
function ReturnIndicator(bIsPSWRetain, bIsManualPort)
{
	// Method to create indicator for each row.
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	
	// Check PSW
	var bIsPSWType = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW'));
	// Check TOE
	var bIsTOEType = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE'));
	// Check NP
	var bIsNPType = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP'));
	
	if (bIsPSWType == true)
		if (bIsPSWRetain)
			return "R";
		else
			return "S";
	else if (bIsTOEType == true)
		return "T";
	else if (bIsNPType == true)
		if (bIsManualPort)
			return "P";
		else
			return "N";
	else
		return "N"		
}
<% /* PSC 02/02/2007 EP2_1113 - End */ %> 

// EP2_55 - New Method.
function frmScreen.btnGenerateQuote.onclick()
{
	<% /* PSC 02/02/2007 EP2_1113 - Start */ %>
	var TotalNetMonthlycost = frmScreen.txtTotalMnthCost.value
	
	if (m_iNumOfLoanComponents > 0 && (TotalNetMonthlycost == "" || TotalNetMonthlycost == 0))
	{
		alert("Delete existing loancomponents or calculate");
		return;
	}
	
	if (m_iNumOfLoanComponents > 0 && TotalNetMonthlycost != "" && TotalNetMonthlycost > 0)
	{
		<% /* PSC 12/02/2007 EP2_1314 - Start */ %>
		if (UpdateQuotation())
		{
			if(CreateQuotation())
			{
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XML.CreateRequestTag(window,null);
				XML.CreateActiveTag("QUOTATION");
				XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
				XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
				XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idQuotationNumber","1"));
				XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
				XML.CreateTag("LIFESUBQUOTENUMBER", m_sLifeSubquoteNumber);
				XML.CreateTag("QUOTATIONTYPE", "2");
				XML.CreateTag("CREATESUBQUOTES", "0");

				switch (ScreenRules())
				{
					case 1: // Warning
					case 0: // OK
						XML.RunASP(document, "CreateNewMortgageLifeSubquote.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
				}

				if(XML.IsResponseOK()) 
				{	// Reset m_sMortgageSubquoteNumber before call.
						m_sMortgageSubquoteNumber = XML.GetTagText("MORTGAGESUBQUOTENUMBER")
						scScreenFunctions.SetContextParameter(window,"idMortgageSubQuoteNumber",m_sMortgageSubquoteNumber);
				}
			}
		}
		<% /* PSC 12/02/2007 EP2_1314 - End */ %>
	}
		
	CreateComponentsFromExistingAcc();
	
	<% /* PSC 12/02/2007 EP2_1314 - Start */ %>
	scScreenFunctions.SetFieldToWritable(frmScreen, "txtPurchasePrice");
	scScreenFunctions.SetFieldToWritable(frmScreen, "txtAmtRequested");
	<% /* PSC 12/02/2007 EP2_1314 - End */ %>
	scScreenFunctions.SetFieldToWritable(frmScreen, "txtRightToBuyDisc"); <%/*EP2_1730*/%>
	// Now reset the form
	m_PopWinInd = 'Yes';
	Initialise();
	
	XML = null;	
	<% /* PSC 02/02/2007 EP2_1113 - End */ %> 
}

// EP2_55 - New Method.  
// Specced in EP2_56 Middle tier, but implemented using CRUD as per 'new code' dictat.
function CreateComponentsFromExistingAcc()
{
	//Set up variables
	var AccountGUID= "";		// Account GUID from Application table. 
	var MortList;               // List fo Mortgageloan details
	var bContinue = true;
	
	// Is our Application PSW, TOE or NP? Return if not.
	var sMortType;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW')))
		sMortType = "PSW";
	else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE')))
		sMortType = "TOE";
	else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP')))
		sMortType = "NP";
	else
	{
		alert("Mortgage Type is not valid for this operation. Operation cancelled");
		return ;
	}	
	// Get AccountGUID flag.
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.RunASP(document,"GetApplicationData.asp");
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null,"APPLICATION");
		AccountGUID = XML.GetTagText("ACCOUNTGUID");
	}

	// Now retrieve the Mortgageloan details
	// Use AccountGUID from Application to retrieve the MortgageLoan data.
	if (AccountGUID != "")
	{
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window);
		var xn = XML.XMLDocument.documentElement;
		xn.setAttribute("CRUD_OP","READ");
		xn.setAttribute("SCHEMA_NAME","epsomCRUD");
		<% /* PSC 02/02/2007 EP2_1271 */ %> 
		xn.setAttribute("ENTITY_REF","MORTGAGELOANANDPRODUCT");
		var xe = XML.XMLDocument.createElement("MORTGAGELOAN");
		xe.setAttribute("ACCOUNTGUID", AccountGUID);
		xn.appendChild(xe);

		XML.RunASP(document, "omCRUDIf.asp");
		if (XML.IsResponseOK())
		{
				<% /* PSC 07/02/2007 EP2_1271 - Start */ %> 
				var invalidLoan = XML.XMLDocument.selectSingleNode("RESPONSE/MORTGAGELOAN[not(MORTGAGEPRODUCT)]");
	
				if (invalidLoan != null)
				{
					var productCode = invalidLoan.getAttribute("MORTGAGEPRODUCTCODE");
					
					alert("Unable to create loan components. Mortgage loan has an invalid product code " + productCode);
					return;
				}
				<% /* PSC 07/02/2007 EP2_1271 - End */ %> 
		
				if (sMortType == 'PSW')
				{
					var sPSRValue = XML.GetComboIdForValidation("RedemptionStatus", "PSR", null, document);
					var sPSValue = XML.GetComboIdForValidation("RedemptionStatus", "PS", null, document);
					var strPattern = "RESPONSE/MORTGAGELOAN[@REDEMPTIONSTATUS='" + sPSRValue + "'or @REDEMPTIONSTATUS='" + sPSValue + "']";
					MortList = XML.XMLDocument.selectNodes(strPattern);
				}
				else if (sMortType == 'NP')
				{
					var sTBPValue = XML.GetComboIdForValidation("RedemptionStatus", "TBP", null, document);
					var strPattern = "RESPONSE/MORTGAGELOAN[@REDEMPTIONSTATUS='" + sTBPValue + "']";
					MortList = XML.XMLDocument.selectNodes(strPattern);
				}
				else if (sMortType == 'TOE')
				{
					var sTOEValue = XML.GetComboIdForValidation("RedemptionStatus", "TOE", null, document);
					var strPattern = "RESPONSE/MORTGAGELOAN[@REDEMPTIONSTATUS='" + sTOEValue + "']";
					MortList = XML.XMLDocument.selectNodes(strPattern);
				}
				
		}

		// Only continue if we've found Mortgage loan details
		if (MortList != null)
		{
			// Set variables
			var iMorts = MortList.length
			var TotalAmountRequested = 0;
			
			// Create new Loancomponents for each Mortgageloan.
			for(var nLoop = 0; nLoop < iMorts; nLoop++)
			{
				if(bContinue == true)
				{
					var CurrMortgageLoan = MortList.item(nLoop);
					var sRedemptionStatus = (CurrMortgageLoan.getAttribute("REDEMPTIONSTATUS")? CurrMortgageLoan.getAttribute("REDEMPTIONSTATUS"): 0) ;
					var gMortgageLoanGUID = (CurrMortgageLoan.getAttribute("MORTGAGELOANGUID")? CurrMortgageLoan.getAttribute("MORTGAGELOANGUID"): "") ;
					var OSBalance = parseFloat(CurrMortgageLoan.getAttribute("OUTSTANDINGBALANCE")? CurrMortgageLoan.getAttribute("OUTSTANDINGBALANCE"): 0) ;
					var AvailDisburse = parseFloat(CurrMortgageLoan.getAttribute("AVAILABLEFORDISBURSEMENT")? CurrMortgageLoan.getAttribute("AVAILABLEFORDISBURSEMENT"): 0) ;
					var OriginalLTV = (CurrMortgageLoan.getAttribute("ORIGINALLTV")? CurrMortgageLoan.getAttribute("ORIGINALLTV"): 0) ;
					var RepayType = (CurrMortgageLoan.getAttribute("REPAYMENTTYPE")? CurrMortgageLoan.getAttribute("REPAYMENTTYPE"): 0) ;
					var nOrigPandPIntOnly = (CurrMortgageLoan.getAttribute("ORIGINALPARTANDPARTINTONLYAMT")? CurrMortgageLoan.getAttribute("ORIGINALPARTANDPARTINTONLYAMT"): 0) ;
					var sMortgageProdCode = CurrMortgageLoan.getAttribute("MORTGAGEPRODUCTCODE") ;
					var dStartDate = CurrMortgageLoan.getAttribute("STARTDATE");
					
					<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
					var sLoanStartDate = CurrMortgageLoan.getAttribute("STARTDATE");
					var lYears = CurrMortgageLoan.getAttribute("ORIGINALTERMYEARS");
					if (lYears == "")
						lYears = "0";
					var lMonths = CurrMortgageLoan.getAttribute("ORIGINALTERMMONTHS");
					if (lMonths == "") 
						lMonths = "0";
					
					var l_sOutStandingTerm = CalculateOutstandingTerm(sLoanStartDate, lYears, lMonths);
					
					lYears = Math.floor(l_sOutStandingTerm /12);
					lMonths = l_sOutStandingTerm % 12;
					
					<% /* PSC 07/02/2007 EP2_1271 - Start */ %> 
					var mortgageProductNode = CurrMortgageLoan.selectSingleNode("MORTGAGEPRODUCT");
					var dProductStartDate = mortgageProductNode.getAttribute("STARTDATE");
					<% /* PSC 02/02/2007 EP2_1113 - End */ %>
					<% /* PSC 07/02/2007 EP2_1271 - End */ %> 
					
					var lXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					lXML.CreateRequestTag(window);
					var xn = lXML.XMLDocument.documentElement;
					xn.setAttribute("CRUD_OP","CREATE");
					xn.setAttribute("SCHEMA_NAME","omCRUD");
					xn.setAttribute("ENTITY_REF","LOANCOMPONENT");
					var xe = lXML.XMLDocument.createElement("LOANCOMPONENT");
					xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
					xe.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
					xe.setAttribute("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
					xe.setAttribute("MORTGAGELOANGUID", gMortgageLoanGUID);
	//				"LOANCOMPONENTSEQUENCENUMBER" is a Generated field. 
					xe.setAttribute("LOANAMOUNT",OSBalance );
					var TotalBalance = OSBalance + AvailDisburse;
					xe.setAttribute("TOTALLOANCOMPONENTAMOUNT", TotalBalance );
					// Add to Total amount
					TotalAmountRequested = TotalAmountRequested + TotalBalance;

	//				THIS NEEDS TO BE ADDED AFTER Column added to table.
	//				xe.setAttribute("DRAWDOWN",CurrMortgageLoan.getAttribute(""));
					xe.setAttribute("ORIGINALLTV", OriginalLTV);
					xe.setAttribute("REPAYMENTMETHOD", RepayType );
					
					<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
					xe.setAttribute("TERMINYEARS", lYears);
					xe.setAttribute("TERMINMONTHS", lMonths);
					<% /* PSC 02/02/2007 EP2_1113 - End */ %> 

					xe.setAttribute("PRODUCTCODESEARCHIND", 1 );

					// If Part and Part payment
					if (XML.IsInComboValidationList(document, 'RepaymentType', RepayType, Array('P')))
					{
						xe.setAttribute("INTERESTONLYELEMENT", nOrigPandPIntOnly );
						var CapLessInt = OSBalance - nOrigPandPIntOnly;
						xe.setAttribute("CAPITALANDINTERESTELEMENT", CapLessInt);
						xe.setAttribute("NETINTONLYELEMENT", nOrigPandPIntOnly );
						xe.setAttribute("NETCAPANDINTELEMENT", CapLessInt );
					}
					// Product Switch Retained
					if (XML.IsInComboValidationList(document, 'RedemptionStatus', sRedemptionStatus, Array('PSR')))
						xe.setAttribute("PRODUCTSWITCHRETAINPRODUCTIND", 1 );
					else 
						xe.setAttribute("PRODUCTSWITCHRETAINPRODUCTIND", 0 );

					// To be ported
					if (XML.IsInComboValidationList(document, 'RedemptionStatus', sRedemptionStatus, Array('TBP')))
						xe.setAttribute("MANUALPORTEDLOANIND", 1 );
					else 
						xe.setAttribute("MANUALPORTEDLOANIND", 0 );

					// Application Type
					if (sMortType == 'PSW')
					{
						var sPurposeOfLoan = XML.GetComboIdForValidation("PurposeOfLoanNew", "PSWD", null, document);
						xe.setAttribute("PURPOSEOFLOAN", sPurposeOfLoan);
					}
					else if (sMortType == 'NP')
					{
						var sPurposeOfLoan = XML.GetComboIdForValidation("PurposeOfLoanNew", "NPD", null, document);
						xe.setAttribute("PURPOSEOFLOAN", sPurposeOfLoan);
					}
					else if (sMortType == 'TOE')
					{
						<% /* EP2_1464 */ %> 
						var sPurposeOfLoan = XML.GetComboIdForValidation("PurposeOfLoanTrfOfEquity", "TOED", null, document);
						xe.setAttribute("PURPOSEOFLOAN", sPurposeOfLoan);
					}
					if (sMortgageProdCode != null)
						xe.setAttribute("MORTGAGEPRODUCTCODE", sMortgageProdCode );
					if (dStartDate != null)
						xe.setAttribute("STARTDATE", dStartDate );
						
					<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
					if (dProductStartDate != null)
						xe.setAttribute("STARTDATE", dProductStartDate);
					<% /* PSC 02/02/2007 EP2_1113 - End */ %> 
						
					// Append to Request
					xn.appendChild(xe);

					lXML.RunASP(document, "omCRUDIf.asp");
					if (!lXML.IsResponseOK())
					{	
						alert ("Error creating LoanComponent entries in table");
						bContinue = false;
					}

				} // End If bContinue == true

			}  // End nLoop < iMorts
			
			// Lastly Update the mortgage subquote Total
			if (bContinue == true)
			{
				<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
				frmScreen.txtAmtRequested.value = TotalAmountRequested;
				CalcAndPopulateLTV();
				<% /* PSC 02/02/2007 EP2_1113 - End */ %> 

				XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XML.CreateRequestTag(window);
				var xn = XML.XMLDocument.documentElement;
				xn.setAttribute("CRUD_OP","UPDATE");
				xn.setAttribute("SCHEMA_NAME","omCRUD");
				xn.setAttribute("ENTITY_REF","MORTGAGESUBQUOTE");
				var xe = XML.XMLDocument.createElement("MORTGAGESUBQUOTE");
				xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
				xe.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
				xe.setAttribute("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
				xe.setAttribute("AMOUNTREQUESTED", TotalAmountRequested);
				
				<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
				xe.setAttribute("TOTALLOANAMOUNT", TotalAmountRequested);
				xe.setAttribute("LTV", frmScreen.txtLTVPercent.value);
				<% /* PSC 02/02/2007 EP2_1113 - End */ %> 

				xn.appendChild(xe);

				XML.RunASP(document, "omCRUDIf.asp");
				if (!XML.IsResponseOK())
					Alert ("Failed to update AmountRequested in MortgageSubquote table");
			} // End update subquote
		  
		} // End (MortList != null)
	} // End (AccountGUID != "")
	else
		alert("No existing Account details found. Operation cancelled");
	
	// Clear up the XML object.		
	XML = null;
}

<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
function ResetGenerateQuote()
{
	var tempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
		
	if (tempXML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW'))
			|| tempXML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE'))
			|| tempXML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP')))
			
	{
			var purchasePrice = frmScreen.txtPurchasePrice.value;
			var totalMonthlyCost = frmScreen.txtTotalMnthCost.value;
			
			if (m_iNumOfLoanComponents == 0 && purchasePrice != "" && purchasePrice > 0 &&
			    (totalMonthlyCost == "" || totalMonthlyCost == 0))
				frmScreen.btnGenerateQuote.disabled = false;
			else if (m_iNumOfLoanComponents > 0 && purchasePrice != "" && purchasePrice > 0 &&
			         (totalMonthlyCost == "" || totalMonthlyCost > 0))
				frmScreen.btnGenerateQuote.disabled = false;
			else
				frmScreen.btnGenerateQuote.disabled = true;
	}
	else
		frmScreen.btnGenerateQuote.disabled = true;
}

function CalculateOutstandingTerm(strStartDate, strOriginalTermYears, strOriginalTermMonths)
{
	var lOutstandingMonths = 0;
		
	var dtStartDate = scScreenFunctions.GetDateObjectFromString(strStartDate);
	var sOriginalTermYears = strOriginalTermYears;
	var sOriginalTermMonths = strOriginalTermMonths;
		
	if(dtStartDate != null)
	{
		var nStartYear = dtStartDate.getFullYear();
		var nStartMonth = dtStartDate.getMonth();
			
		<% /* Get the current date (but strip out the time info) */ %>
		<% /* MO - BMIDS00807 */ %>
		var dtTempCurrentDate = scScreenFunctions.GetAppServerDate(true);
		<% /* var dtTempCurrentDate	= new Date(); */ %>
		var dtCurrentDate		= new Date(dtTempCurrentDate.getFullYear(),
								           dtTempCurrentDate.getMonth(),
								           dtTempCurrentDate.getDate());

		var nCurrentYear = dtCurrentDate.getFullYear();
		var nCurrentMonth = dtCurrentDate.getMonth();

		if(dtCurrentDate > dtStartDate)
		{
			var nYearsPassed = nCurrentYear - nStartYear;
			var nMonthsPassed = 0;
				
			if(nStartMonth > nCurrentMonth)
			{
				nMonthsPassed = 12 - (nStartMonth - nCurrentMonth);
				nYearsPassed--;
			}
			else
				nMonthsPassed = nCurrentMonth - nStartMonth;

			if(sOriginalTermYears != "")
			{
				var nYearsToGo = sOriginalTermYears - nYearsPassed;
				var nMonthsToGo = 0;
					
				if(nMonthsPassed > sOriginalTermMonths)
				{
					nMonthsToGo = 12 - (nMonthsPassed - sOriginalTermMonths);
					nYearsToGo--;
				}
				else
					nMonthsToGo = sOriginalTermMonths - nMonthsPassed;
					
				if((nYearsToGo == 0 && nMonthsToGo > 0) || nYearsToGo > 0)
				{
					lOutstandingMonths = nYearsToGo * 12 + nMonthsToGo;
				}
			}
		}
	}
	
	return lOutstandingMonths;
}
<% /* PSC 02/02/2007 EP2_1113 - End */ %>

<% /* PSC 12/02/2007 EP2_1314 - Start */ %>
function frmScreen.txtRightToBuyDisc.onblur()
{
	if (frmScreen.txtRightToBuyDisc.value != mDiscountAmount)
	{
		var sPurchasePrice = frmScreen.txtPurchasePrice.value;
		var sDiscountAmount = frmScreen.txtRightToBuyDisc.value;
		
		if (sPurchasePrice == "")
			sPurchasePrice = "0";
			
		if (sDiscountAmount == "")
			sDiscountAmount = "0";
			
		if (parseInt(sDiscountAmount) > parseInt(sPurchasePrice))
		{
			alert("The discount amount cannot be greater than the purchase price");
			frmScreen.txtRightToBuyDisc.value = mDiscountAmount;
			frmScreen.txtRightToBuyDisc.focus();
			return;
		}
		
		if (UpdateRTBDiscount())
		{
			mDiscountAmount = frmScreen.txtRightToBuyDisc.value;
			<%/* EP2_1730  No point in updating the MortgageSubQuote if Amount Requested is empty*/%>
			if(frmScreen.txtAmtRequested.value != ""){Reset();}
		}	
	} 
}

function Reset()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	XML.CreateRequestTag(window,null);
			
	XML.CreateActiveTag("APPLICATIONQUOTE");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
	XML.CreateTag("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
	// MAR1061 - Peter Edney - 13/03/2006
	XML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value);
	XML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value);		
	XML.CreateTag("DRAWDOWN", frmScreen.txtDrawDown.value);
			
	AddCustomerList(XML);
			
	XML.RunASP(document, "AQResetMortgageSubQuote.asp");

	if(XML.IsResponseOK()) 
		Initialise();
			
	XML = null;
}
<% /* PSC 12/02/2007 EP2_1314 - End */ %>
 


-->
</script>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/Jscript"></script>
</body>
</html>


