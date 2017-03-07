<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP030.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Application Summary - Loan & Risks
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		27/02/01	Created	
CL		05/03/01	SYS1920 Read only functionality added
JLD		07/03/01	SYS1879 Plug methods in.
SA		23/07/01	SYS2204 Altered captions Calculator and L.T.V
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog    Date        AQR		Description
TW      09/10/2002  SYS5115 Modified to incorporate client validation  
MV		04/02/2003	BM0288	Amended PopulateScreen() and HTML 
GD		15/07/2003	BM0591  Indemnity Amount to be displayed as zero if negative
SR		27/07/2004	BMIDS822 Change lables 'Indemnity Amount' To 'Higher Lending Charge Mortgage Amount'
							 'Indemnity Premium' To 'Higher Lending Charge Premium'
SR		27/07/2004	BMIDS822 Re-aligned lables							 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS 

Prog    Date        AQR		Description
JD		01/12/2005	MAR722	Add Hometrack Valuation field
JD		15/02/2006	MAR1040	Change to screen fields and display values.
JD		05/05/2006	MAR1700 alter where ltv property price is populated from
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
<form id="frmToAP020" method="post" action="AP020.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP010" method="post" action="AP010.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div id="divLoanDetails" style="HEIGHT: 140px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabelHead">
	Loan Details:
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Purpose of Loan
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtPurposeOfLoan" maxlength="50" style="WIDTH: 150px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 65px" class="msgLabel">
	Term
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtTermYears" maxlength="3" style="POSITION: absolute; WIDTH: 40px" class="msgTxt">
		<span style="LEFT: 50px; POSITION: absolute; TOP: 0px" class="msgLabel">Yrs</span>
	</span>
	<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
		<input id="txtTermMonths" maxlength="3" style="POSITION: absolute; WIDTH: 40px" class="msgTxt">
		<span style="LEFT: 50px; POSITION: absolute; TOP: 0px" class="msgLabel">Mths</span>
	</span>		
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 90px" class="msgLabel">
	Life Policy Type
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtLifePolicy" maxlength="50" style="WIDTH: 150px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 339px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Total Loan Amount
	<span style="LEFT: 130px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
		<input id="txtTotalLoanAmount" maxlength="10" style="LEFT: 4px; POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 339px; POSITION: absolute; TOP: 65px" class="msgLabel">
	Outstanding Loans
	<span style="LEFT: 130px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
		<input id="txtOutstandingLoans" maxlength="10" style="LEFT: 4px; POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 339px; POSITION: absolute; TOP: 90px" class="msgLabel">
	Total
	<span style="LEFT: 130px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
		<input id="txtTotal" maxlength="10" style="LEFT: 4px; POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 120px" class="msgLabel">
	Sub Purpose Of Loan
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtSubPurposeOfLoan" maxlength="5" style="WIDTH: 150px" class="msgTxt">
	</span>
</span>
</div>
<div id="divRisks" style="HEIGHT: 195px; LEFT: 10px; POSITION: absolute; TOP: 210px; WIDTH: 604px" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabelHead">
	Risks:
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Direct/Indirect Application
	<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
		<input id="txtDirectIndirectApp" maxlength="50" style="LEFT: 20px; POSITION: absolute; TOP: 1px; WIDTH: 100px" class="msgTxt">
	</span>
</span>

<span style="LEFT: 4px; POSITION: absolute; TOP: 65px" class="msgLabel">
	Purchase Price
	<span style="LEFT: 140px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: 11px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
		<input id="txtPurchasePrice" maxlength="10" style="LEFT: 20px; POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 90px" class="msgLabel">
	Present Valuation
	<span style="LEFT: 140px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: 11px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
		<input id="txtPresentValue" maxlength="10" style="LEFT: 20px; POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 245px; POSITION: absolute; TOP: -6px">
		<input id="btnDetails" value="Details" type="button" style="LEFT: 19px; POSITION: absolute; TOP: 1px; VISIBILITY: hidden; WIDTH: 55px" class="msgButton">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 115px" class="msgLabel">
	Post Works Valuation
	<span style="LEFT: 140px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: 11px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
		<input id="txtPostValue" maxlength="10" style="LEFT: 20px; POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 140px" class="msgLabel">
	Property val to be used in LTV
	<span style="LEFT: 140px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: 10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
		<input id="txtCalculator" maxlength="10" style="LEFT: 20px; POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 165px" class="msgLabel">
	LTV
	<span style="LEFT: 140px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: 83px; POSITION: absolute; TOP: 0px" class="msgLabel">%</label>
		<input id="txtLTV" maxlength="3" style="LEFT: 20px; POSITION: absolute; TOP: -3px; WIDTH: 60px" class="msgTxt">
	</span>
</span>
<span id="spnHomeTrack" style="LEFT: 4px; POSITION: absolute; TOP: 190px" class="msgLabel">
	HomeTrack Valuation
	<span style="LEFT: 140px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: 11px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
		<input id="txtHomeTrackVal" maxlength="10" style="LEFT: 20px; POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT:280px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Risk Assessment Decision
	<span style="LEFT: 215px; POSITION: absolute; TOP: -3px">
		<input id="txtRiskScore" maxlength="10" style="LEFT: 4px; TOP: 1px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 280px; POSITION: absolute; TOP: 65px" class="msgLabel">
	Declared Income Multiple
	<span style="LEFT: 215px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" ></label>
		<input id="txtDeclIncMult" maxlength="10" style="LEFT: 0px; POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 280px; POSITION: absolute; TOP: 90px" class="msgLabel">
	Confirmed Income Multiple
	<span style="LEFT: 215px; POSITION: absolute; TOP: -3px">
		<input id="txtIncomeMultiple" maxlength="10" style="LEFT: 4px; TOP: 1px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<% /* MAR1040 <span style="LEFT: 280px; POSITION: absolute; TOP: 115px" class="msgLabel">
	% of Disp. Net Income
	<span style="LEFT: 215px; POSITION: absolute; TOP: -3px">
		<input id="txtDispNetIncome" maxlength="5" style="LEFT: 4px; TOP: 1px; WIDTH: 100px" class="msgTxt">
	</span> 
</span>  */ %>
<span style="LEFT: 280px; POSITION: absolute; TOP: 140px" class="msgLabel">
	Higher Lending Charge Mortgage Amount
	<span style="LEFT: 215px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
		<input id="txtIndemnityAmount" maxlength="10" style="LEFT: 0px; POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 280px; POSITION: absolute; TOP: 165px" class="msgLabel">
	Higher Lending Charge Premium
	<span style="LEFT: 215px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
		<input id="txtIndemnityPremium" maxlength="10" style="LEFT: 0px; POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 435px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP030attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sMortgageAppValue = "";
var m_sReadOnly = "";
var scScreenFunctions;
var m_AppSummaryXML = null;
var m_sInstructionSeqNo = "";
var m_blnReadOnly = false;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Back");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Application Summary - Loan & Risks","AP030",scScreenFunctions);

	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
	SetMasks();
	Validation_Init();
	PopulateScreen();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP030");
	
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
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sMortgageAppValue = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
}
function PopulateScreen()
{
	scScreenFunctions.SetCollectionState(divLoanDetails,"R");
	scScreenFunctions.SetCollectionState(divRisks,"R");
	frmScreen.btnDetails.disabled = true;
	
	//MAR1700
	var sTypeOfApp = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","");
	var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	
	m_AppSummaryXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_AppSummaryXML.CreateRequestTag(window, "GetLoanAndRisksData");
	m_AppSummaryXML.CreateActiveTag("APPLICATION");
	m_AppSummaryXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	m_AppSummaryXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	m_AppSummaryXML.RunASP(document, "omAppProc.asp");
	//m_AppSummaryXML.LoadXML(m_AppSummaryXML.ReadXMLFromFile("d:\\projects\\XMLTEST\\GetLoanAndRiskData.xml"));
	if(m_AppSummaryXML.IsResponseOK())
	{
		m_AppSummaryXML.SelectTag(null, "LOANDATA");
		frmScreen.txtPurposeOfLoan.value = m_AppSummaryXML.GetAttribute("PURPOSEOFLOAN_TEXT");
		frmScreen.txtTotalLoanAmount.value = m_AppSummaryXML.GetAttribute("TOTALLOANAMOUNT");
		frmScreen.txtTermYears.value = m_AppSummaryXML.GetAttribute("TERMINYEARS");
		frmScreen.txtTermMonths.value = m_AppSummaryXML.GetAttribute("TERMINMONTHS");
		frmScreen.txtOutstandingLoans.value = m_AppSummaryXML.GetAttribute("LOANSANDLIABILITIES");
		frmScreen.txtSubPurposeOfLoan.value =  m_AppSummaryXML.GetAttribute("SUBPURPOSEOFLOAN_TEXT");
		var nTotal = 0;
		nTotal = GetStringValue(frmScreen.txtTotalLoanAmount.value) + GetStringValue(frmScreen.txtOutstandingLoans.value);
		frmScreen.txtTotal.value = nTotal.toString();
		frmScreen.txtLifePolicy.value = m_AppSummaryXML.GetAttribute("BENEFITTYPE_TEXT");
		frmScreen.txtDirectIndirectApp.value = m_AppSummaryXML.GetAttribute("DIRECTINDIRECTBUSINESS_TEXT");
		frmScreen.txtPurchasePrice.value = m_AppSummaryXML.GetAttribute("PURCHASEPRICEORESTIMATEDVALUE");
		frmScreen.txtPresentValue.value = m_AppSummaryXML.GetAttribute("PRESENTVALUATION");
		frmScreen.txtPostValue.value = m_AppSummaryXML.GetAttribute("POSTWORKSVALUATION");
		frmScreen.txtLTV.value = m_AppSummaryXML.GetAttribute("LTV");
		frmScreen.txtRiskScore.value = m_AppSummaryXML.GetAttribute("RISKASSESSMENTDECISION"); //JD MAR1040
		<% //BM0591 START %>
		<%//frmScreen.txtIndemnityAmount.value = m_AppSummaryXML.GetAttribute("INDEMNITYAMOUNT");%>
		var tempIndemnityAmount = m_AppSummaryXML.GetAttribute("INDEMNITYAMOUNT")
		if (isNaN(tempIndemnityAmount))
		{
			tempIndemnityAmount = "0";
		} else
		{
			if (parseFloat(tempIndemnityAmount) < 0)
			{
				tempIndemnityAmount = "0";
			}
		}
		frmScreen.txtIndemnityAmount.value = tempIndemnityAmount;
		<% //BM0591 END %>
		<% /* JD MAR1040 set as 0 if -ve */ %>
		if(parseIntSafe(m_AppSummaryXML.GetAttribute("INDEMNITYPREMIUM")) >= 0)
			frmScreen.txtIndemnityPremium.value = m_AppSummaryXML.GetAttribute("INDEMNITYPREMIUM");
		
		var nPresentValuation = GetStringValue(m_AppSummaryXML.GetAttribute("PRESENTVALUATION"));
		var nPurchasePrice = GetStringValue(m_AppSummaryXML.GetAttribute("PURCHASEPRICEORESTIMATEDVALUE"));
		<% /* MAR1700 check change of property and change logic for setting txtCalculator */ %>
		if(m_AppSummaryXML.GetAttribute("CHANGEOFPROPERTY") == "1")
		{
			frmScreen.txtCalculator.value = nPurchasePrice.toString();
		}
		else if(parseIntSafe(m_AppSummaryXML.GetAttribute("POSTWORKSVALUATION")) > 0)
		{
			frmScreen.txtCalculator.value = m_AppSummaryXML.GetAttribute("POSTWORKSVALUATION");
		}
		else if (TempXML.IsInComboValidationList(document,"TypeOfMortgage",sTypeOfApp,["N"])) //new loan
		{
			if(nPresentValuation > 0.0 && nPresentValuation < nPurchasePrice)frmScreen.txtCalculator.value = nPresentValuation.toString();
			else frmScreen.txtCalculator.value = nPurchasePrice.toString();
		}
		else //remortgage 
		{
			if(nPresentValuation > 0.0) frmScreen.txtCalculator.value = nPresentValuation.toString();
			else frmScreen.txtCalculator.value = nPurchasePrice.toString();
		}
		
		frmScreen.txtDeclIncMult.value = m_AppSummaryXML.GetAttribute("INCOMEMULTIPLE"); //JD MAR1040
		frmScreen.txtIncomeMultiple.value = m_AppSummaryXML.GetAttribute("CONFIRMEDCALCULATEDINCMULTIPLE"); //JD MAR1040
		<% /*  MAR1040 remove % of disp income and gross income
		var nDispIncome;
		nDispIncome = GetStringValue(m_AppSummaryXML.GetAttribute("LOANSANDLIABILITIES")) + 
					  GetStringValue(m_AppSummaryXML.GetAttribute("MONTHLYMORTGAGEPAYMENTS")) + 
					  GetStringValue(m_AppSummaryXML.GetAttribute("MORTGAGERELATEDINSURANCE")) + 
					  GetStringValue(m_AppSummaryXML.GetAttribute("OTHEROUTGOINGS"));
		nDispIncome = (nDispIncome / GetStringValue(m_AppSummaryXML.GetAttribute("TOTALMONTHLYINCOME"))) * 100;
		frmScreen.txtDispNetIncome.value = top.frames[1].document.all.scMathFunctions.RoundValue(nDispIncome.toString(), 2);
		
		GetHighestEarners();
		*/ %>
		//GetValuerInstructions();  Middle tier method not available and route to AP205 is disabled so we don't need to do this at the moment.
	}
	//Get the hometrack valuation amount if app is a remortgage   MAR722
	if (TempXML.IsInComboValidationList(document,"TypeOfMortgage",sTypeOfApp,["R"]))
	{
		var ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var reqTag = ValuationXML.CreateRequestTag(window , "GetPresentValuation");
	
		ValuationXML.CreateActiveTag("HOMETRACKVALUATION");
		ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		ValuationXML.RunASP(document, "GetPresentValuation.aspx");  //MAR722
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = ValuationXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{
			//record not found - valuation not carried out
		}
		else if(ErrorReturn[0] == true)
		{
			ValuationXML.SelectTag(null, "HOMETRACKVALUATION");
			frmScreen.txtHomeTrackVal.value = ValuationXML.GetAttribute("VALUATIONAMOUNT");
			
			<% /* MAR1700 if valuation type is AU then set the txtCalculator field too */ %>
			if (TempXML.IsInComboValidationList(document,"ValuationType",m_AppSummaryXML.GetAttribute("VALUATIONTYPE"),["AU"]))
			{
				var nHometrackValuation = parseIntSafe(ValuationXML.GetAttribute("VALUATIONAMOUNT"));
				if(nHometrackValuation > 0.0 && nHometrackValuation < nPurchasePrice)frmScreen.txtCalculator.value = nHometrackValuation.toString();
				else frmScreen.txtCalculator.value = nPurchasePrice.toString();

			}
		}
	}
	else
	{
		//hide the field
		scScreenFunctions.HideCollection(spnHomeTrack);
	}
	
}
function GetHighestEarners()
{
<% /* JD MAR1040 method not called */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "");
	var tagLIST = XML.CreateActiveTag("CUSTOMERLIST");
	for(nCust = 1; nCust <= 5; nCust++)
	{
		var sCustNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nCust,null);
		var sCustVersNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nCust,null);
		var sCustRole = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nCust,null);
		if(sCustNumber != "")
		{
			XML.ActiveTag = tagLIST;
			XML.CreateActiveTag("CUSTOMER");
			XML.CreateTag("CUSTOMERNUMBER", sCustNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustVersNumber);
			XML.CreateTag("CUSTOMERROLETYPE", sCustRole);
		}
	}
	XML.RunASP(document, "GetHighestEarners.asp");
	if(XML.IsResponseOK())
	{
		var nTotal = 0;
		XML.SelectTag(null, "EMPLOYMENTANDINCOME");
		XML.CreateTagList("INCOMESUMMARY");
		for(var nIncome = 0; nIncome < XML.ActiveTagList.length; nIncome++)
		{
			XML.SelectTagListItem(nIncome);
			nTotal += GetStringValue(XML.GetTagText("TOTALGROSSEARNEDINCOME"));
		}
		frmScreen.txtDeclIncMult.value = nTotal.toString();
	}
	XML = null;
}
function GetValuerInstructions()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "GetValuerInstructions");
	XML.CreateActiveTag("VALUERINSTRUCTIONS");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	//XML.RunASP(document, "omAppProc.asp");
	XML.LoadXML(XML.ReadXMLFromFile("d:\\projects\\XMLTEST\\GetValuerInstr.xml"));
	if(XML.IsResponseOK())
	{
		frmScreen.btnDetails.disabled = false;
		XML.SelectTag(null, "VALUERINSTRUCTIONS");
		m_sInstructionSeqNo = XML.GetAttribute("INSTRUCTIONSEQUENCENO");
	}
}
function frmScreen.btnDetails.onclick()
{
// THIS BUTTON HAS BEEN MADE INVISIBLE FOR THE TIME BEING TO AVOID ANY ROUTING CONFLICTS
	alert("go to valuation report screen when available");
}
function GetStringValue(sString)
{
<%	/* returns the float value of the string checking for null and empty string in the process*/
%>	if(sString == "") return(0.0);
	if(isNaN(parseFloat(sString))) return(0.0);
	return(parseFloat(sString));
}
function btnSubmit.onclick()
{
	frmToAP010.submit();
}
function btnBack.onclick()
{
	frmToAP020.submit();
}
function parseIntSafe(sText)
{
	if (sText=="")
	   return(0);
	else
	   return(parseInt(sText, 10));
}
-->
</script>
</body>
</html>


