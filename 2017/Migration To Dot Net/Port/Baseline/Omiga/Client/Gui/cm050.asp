<%@ LANGUAGE="JSCRIPT" %>

<% /*
Workfile:		cm050.asp
Copyright:		Copyright © 1999 Marlborough Stirling

Description:	Affordability Screen.
				The purpose of this screen is to provide the Advisor and the 
				Applicants with a clear and concise view about the issue of
				Affordability, on a monthly basis.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DPol	17/11/1999	Created.
DPol	17/11/1999	Still in the process of creating it, but as the
					Business Object isn't ready, this is as far as
					I can go at the moment, so checked in for safety.
MCS		24/11/1999	Modified and finished 					
RF		28/01/00	Reworked for IE5 and performance.
AY		14/02/00	Change to msgButtons button types
AY		29/03/00	New top menu/scScreenFunctions change
AY		12/04/00	SYS0328 - Dynamic currency display
					GetAIPAffordabilityData.asp wasn't in ""
CL		05/03/01	SYS1920 Read only functionality added
IK		22/03/01	SYS2145 Add idNoCompleteCheck
JLD		04/12/01	SYS2806 added completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<html>

<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<script src="validation.js" language="JScript"></script>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<object data="scTableListScroll.asp" id="scScrollTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px; height:24; width:304" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>

<% // Specify Forms Here %>
<form id="frmToCM010" method="post" action="cm010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% // Specify Screen Layout Here %>
<form id="frmScreen" mark validate="onchange" year4>
	<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 160px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
		
		<% // Monthly mortgage repayments %>
		<span style="TOP: 10px; LEFT: 120px; POSITION: ABSOLUTE" class="msgLabel">
			Monthly Mortgage Repayments
			<span style="TOP: 0px; LEFT: 200px; POSITION: ABSOLUTE">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtMonthMortRepayments" name="MonthMortRepayments"  
					maxlength="10" style="TOP: -3px; WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<% // Monthly mortgage related insurance %>
		<span style="TOP: 35px; LEFT: 120px; POSITION: ABSOLUTE" class="msgLabel">
			Monthly Mortgage Related Insurance
			<span style="TOP: 0px; LEFT: 200px; POSITION: ABSOLUTE">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtMonthMortRelatedInsurance" name="MonthMortRelatedInsurance" 
					maxlength="10" style="TOP: -3px; WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<% // Monthly Loans and Liabilities %>
		<span style="TOP: 60px; LEFT: 120px; POSITION: ABSOLUTE" class="msgLabel">
			Monthly Loans and Liabilities
			<span style="TOP: 0px; LEFT: 200px; POSITION: ABSOLUTE">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtMonthLoansAndLiabilities" name="MonthLoansAndLiabilities" 
					maxlength="10" style="TOP: -3px; WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<% // Monthly Other Outgoings %>
		<span style="TOP: 85px; LEFT: 120px; POSITION: ABSOLUTE" class="msgLabel">
			Monthly Other Outgoings
			<span style="TOP: 0px; LEFT: 200px; POSITION: ABSOLUTE">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtMonthOtherOutgoings" name="MonthOtherOutgoings" 
					maxlength="10" style="TOP: -3px; WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<% // Total Monthly Outgoings %>
		<span style="TOP: 110px; LEFT: 120px; POSITION: ABSOLUTE" class="msgLabel">
			Total Monthly Outgoings
			<span style="TOP: 0px; LEFT: 200px; POSITION: ABSOLUTE">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtTotalMonthOutgoings" name="TotalMonthOutgoings" 
					maxlength="10" style="TOP: -3px; WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<% // Total Monthly Income %>
		<span style="TOP: 135px; LEFT: 120px; POSITION: ABSOLUTE" class="msgLabel">
			Total Monthly Income
			<span style="TOP: 0px; LEFT: 200px; POSITION: ABSOLUTE">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtTotalMonthIncome" name="TotalMonthIncome" 
					maxlength="10" style="TOP: -3px; WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
	</div>
</form>

<% // Main Buttons %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 12px; POSITION: absolute; TOP: 230px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% // File containing field attributes (remove if not required) %>
<!--  #include FILE="attribs/cm050attribs.asp" -->

<% // Specify Code Here %>
<script language="JScript">
<!-- // JScript Code
var m_sMetaAction = null;

var m_XMLAffordability = null;
var m_sUserId	= null;
var m_sUnitId	= null;
var m_sUserType	= null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sQuotationNumber = null;
var scScreenFunctions;
var m_blnReadOnly = false;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Affordability","CM050",scScreenFunctions);

	scScreenFunctions.SetCurrency(window,frmScreen);
	RetrieveContextData();
	SetMasks();
	PopulateScreen();
	scScreenFunctions.SetScreenToReadOnly(frmScreen);
	btnSubmit.focus();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CM050");
	
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
	m_sUserId = scScreenFunctions.GetContextParameter(window,"idUserId","USER0001");
	m_sUnitId = scScreenFunctions.GetContextParameter(window,"idUnitId","UNIT1");
	m_sUserType	= scScreenFunctions.GetContextParameter(window,"idUserType","BRANCH");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B3069");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	m_sQuotationNumber = scScreenFunctions.GetContextParameter(window,"idQuotationNumber","1");
}

function PopulateScreen()
{
	m_XMLAffordability = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	m_XMLAffordability.CreateRequestTag(window, "SEARCH")
							
	m_XMLAffordability.CreateActiveTag("QUOTATION");
	m_XMLAffordability.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	m_XMLAffordability.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	m_XMLAffordability.CreateTag("QUOTATIONNUMBER",m_sQuotationNumber);

	<% /* 
		// populate with the affordability data from the Quotation table.
		//
		// using Application Number, Application FactFindNumber and Quotation Number
		// <in context>
	*/ %>

	 m_XMLAffordability.RunASP(document, "GetAIPAffordabilityData.asp");

	if (m_XMLAffordability.IsResponseOK())
		CalculateAndPopulateValues()
}

					
function btnSubmit.onclick()
{
	<% /* Routes back to the calling screen, CM010 */%>
	scScreenFunctions.SetContextParameter(window,"idNoCompleteCheck","1");
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmToCM010.submit();
}

function CalculateAndPopulateValues()
{
	var iMonthlyMortgageRepayments			= 0; 
	var iMonthlyMortgageRelatedInsurance	= 0;
	var iMonthlyLoansAndLiabilities			= 0;
	var iMonthlyOtherOutgoings				= 0;
	var iTotalMonthlyOutgoings				= 0;
	var iTotalMonthlyIncome					= 0;
							
	m_XMLAffordability.ActiveTag = null;
							
	var RootTag = m_XMLAffordability.CreateTagList("QUOTATION");
							
	var bReturn = m_XMLAffordability.SelectTagListItem(0);
			
	<% // check we have a Quotation %>
	if(m_XMLAffordability.ActiveTagList.length  > 0)
	{
		iMonthlyMortgageRepayments	= parseFloat(m_XMLAffordability.GetTagText("MONTHLYMORTGAGEPAYMENTS"));
		iMonthlyMortgageRelatedInsurance = parseFloat(m_XMLAffordability.GetTagText("MORTGAGERELATEDINSURANCE"));
		iMonthlyLoansAndLiabilities	= parseFloat(m_XMLAffordability.GetTagText("LOANSANDLIABILITIES"));
		iMonthlyOtherOutgoings = parseFloat(m_XMLAffordability.GetTagText("OTHEROUTGOINGS"));
		iTotalMonthlyIncome	= parseFloat(m_XMLAffordability.GetTagText("TOTALMONTHLYINCOME"));
						
		<% //If there is no values set to 0 %>
		if (isNaN(iMonthlyMortgageRepayments))
			iMonthlyMortgageRepayments = 0;
		if (isNaN(iMonthlyMortgageRelatedInsurance))
			iMonthlyMortgageRelatedInsurance = 0;
		if (isNaN(iMonthlyLoansAndLiabilities))
			iMonthlyLoansAndLiabilities = 0;
		if (isNaN(iMonthlyOtherOutgoings))
			iMonthlyOtherOutgoings = 0;
		if (isNaN(iTotalMonthlyIncome))
			iTotalMonthlyIncome = 0;

		<% //add the values up %>
		iTotalMonthlyOutgoings = iMonthlyMortgageRepayments + iMonthlyMortgageRelatedInsurance 
									+ iMonthlyLoansAndLiabilities + iMonthlyOtherOutgoings;
	}

	frmScreen.txtMonthMortRepayments.value	= iMonthlyMortgageRepayments;
	frmScreen.txtMonthMortRelatedInsurance.value = iMonthlyMortgageRelatedInsurance;			
	frmScreen.txtMonthLoansAndLiabilities.value	= iMonthlyLoansAndLiabilities;	
	frmScreen.txtMonthOtherOutgoings.value	= iMonthlyOtherOutgoings;
	frmScreen.txtTotalMonthIncome.value	= iTotalMonthlyIncome;
				
	frmScreen.txtTotalMonthOutgoings.value	= iTotalMonthlyOutgoings;
}

-->
</script>
</body>
</html>


