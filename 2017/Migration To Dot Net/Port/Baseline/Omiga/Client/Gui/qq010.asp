<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      qq010.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: Quick Quote
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		03/04/00	New top menu/scScreenFunctions change
AY		13/04/00	SYS0328 - Dynamic currency display
BG		17/05/00	SYS0752 Removed Tooltips
APS		23/05/00	SYS0780 Changed process to Show 2nd App Controls if Joint Application
DJP		13/06/00	SYS0536 Changed TotalMonthlyRepayments length to include floating point
					numbers
CL		05/03/01	SYS1920 Read only functionality added
IK		15/03/01	SYS1924 Completeness Check routing
JLD		4/12/01		SYS2806 use screenFunctions for completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
DB		29/05/02	SYS4767 MSMS to Core Integration
SG		05/06/02	SYS4818 Error on SYS4767
SG		06/06/02	SYS4828 MSMS to Core integration error
SG		10/06/02	SYS4848 Error on SYS4818
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MO		04/10/2002	BMIDS00578	INWP1, BM061, To remove the intermediary details from the screen
SA		24/10/2002	BMIDS00670  ReadApplicationSourceDetails altered to cater for no records returned from GetQuickQuoteDetails
MDC		01/11/2002	BMIDS00654  ICWP BM088 Income Calculations
MC		20/04/2004	BMIDS517	QQ020,QQ030 Popup dialogs height incr. by 10px
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
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
<link href="stylesheet.css" rel="STYLESHEET" type="text/css"><title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<script src="validation.js" language="JScript"></script>
<% //DB SYS4767 - MSMS Integration %>
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<% //DB End %>
<!-- Specify Forms Here -->
<form id="frmApplicationMenu" method="post" action="mn060.asp" STYLE="DISPLAY: none"></form>
<form id="frmAttitudeToBorrowing" method="post" action="dc330.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange" year4>
<div style="TOP: 60px; LEFT: 10px; HEIGHT: px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<% /* DB SYS4767 - MSMS Integration */ %>
<% /*<div id="divApplicantDetails" style="HEIGHT: 340px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 604px" class="msgGroup"> */ %>
<% /* STB - MSMS0011: Application Source Details */ %>
<div id="divApplicationDetails" style="HEIGHT: 65px; LEFT: 0px; TOP: 0px; WIDTH: 604px; POSITION: absolute;" class="msgGroup">
	<span style="LEFT: 4px; TOP: 10px; POSITION: absolute;" class="msgLabel">
		<strong>Application Source Details...</strong>
	</span>
	<span id="spnApplicationSource" style="LEFT: 4px; TOP: 20px; WIDTH: 200px; POSITION: absolute;" class="msgLabel">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 15px" class="msgLabel">
			Direct/Indirect
			<span style="LEFT: 100px; TOP: -3px; POSITION: absolute;">
				<select id="cboDirectIndirect" name="DirectIndirect" style="WIDTH: 100px" class="msgCombo"></select>
			</span>
		</span>		
	</span>
	<% /* MO 4/10/2002 BMIDS00578 <span id="spnIndividualName" style="LEFT: 220px; TOP: 20px; WIDTH: 300px; POSITION: absolute; VISIBILITY: hidden;" class="msgLabel">
		<span style="LEFT: 0px; TOP: 15px; POSITION: absolute;" class="msgLabel">
			Name of Individual
			<span style="LEFT: 100px; TOP: 0px; POSITION: absolute;">
				<input id="txtIndividualName" name="IndividualName" msg="Please search for an Intermediary" maxlength="82"style="POSITION: absolute; WIDTH: 200px; TOP: -3px" class="msgTxt">
			</span>
		</span>		
	</span> 
	<span id="spnIntermediarySearch" style="LEFT: 530px; TOP: 30px; POSITION: ABSOLUTE; VISIBILITY: hidden;">
		<input id="btnIntermedSearch" value="Search" type="button" style="WIDTH: 60px" class="msgButton">
	</span> */ %>
</div>
<div id="divApplicantDetails" style="HEIGHT: 300px; LEFT: 0px; POSITION: absolute; TOP: 70px; WIDTH: 604px" class="msgGroup">
<% /* DB End */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Applicant Details...</strong>
	</span>
	<span id="spnApplicant1" style="LEFT: 4px; POSITION: absolute; TOP: 20px; width: 300px" class="msgLabel">
		<span id="txtApplicant1Name" name="Applicant1Name" 
			style="LEFT: 150px; OVERFLOW: hidden; WIDTH: 250px; POSITION: absolute; TOP: -3px" tabindex = "-1" class="msglabel">
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 15px" class="msgLabel">
			Employment Status
			<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
				<select id="cboApplicant1EmploymentStatus" name="Applicant1EmploymentStatus" style="WIDTH: 150px" class="msgCombo"></select>
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 40px" class="msgLabel">
			Annual Gross <br>Income/Net Profit
			<span style="LEFT: 150px; POSITION: absolute; TOP: 5px">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtApplicant1GrossIncOrNetProfit" maxlength="10"	name="Applicant1GrossIncOrNetProfit" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 75px" class="msgLabel">
			Total Monthly Outgoings
			<span style="LEFT: 150px; POSITION: absolute; TOP: 0px">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtApplicant1TotalMonthlyOutgoings" maxlength="5" name="Applicant1TotalMonthlyOutgoings" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt">
			</span>
			<span style="TOP: -10px; LEFT: 240px; POSITION: ABSOLUTE">
				<input id="btnApplicant1Outgoings" value="" type="button" 
					style="WIDTH: 32px; HEIGHT: 32px" class="msgDDButton">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 100px" class="msgLabel">
			Total Monthly Personal <br>Debt Repayments
			<span style="LEFT: 150px; POSITION: absolute; TOP: 5px">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtApplicant1TotalMonthlyDebtRepayments" maxlength="5" name="Applicant1TotalMonthlyDebtRepayments" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt">
			</span>
			<span style="TOP: -5px; LEFT: 240px; POSITION: ABSOLUTE">
				<input id="btnApplicant1PersonalDebts" value="" type="button" style="WIDTH: 32px; HEIGHT: 32px" class="msgDDButton">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 135px" class="msgLabel">
			Total O/S Personal Debts
			<span style="LEFT: 150px; POSITION: absolute; TOP: 0px">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtApplicant1TotalOSBalance" maxlength="10" name="Applicant1TotalOSBalance" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt">
			</span>
		</span>
		<span id="spnTaxPayer1" style="LEFT: 0px; POSITION: absolute; TOP: 165px" class="msgLabel">
			UK Tax Payer?
			<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant1UKTaxPayerYes" name="Applicant1UKTaxPayerGroup" type="radio" value="1" checked>
				<label for="optApplicant1UKTaxPayerYes" class="msgLabel">Yes</label>
			</span> 
			<span style="LEFT: 210px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant1UKTaxPayerNo" name="Applicant1UKTaxPayerGroup" type="radio" value="0">
				<label for="optApplicant1UKTaxPayerNo" class="msgLabel">No</label> 
			</span> 
		</span>
		<span id="spnSmoker1" style="LEFT: 0px; POSITION: absolute; TOP: 195px" class="msgLabel">
			Smoker?
			<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant1SmokerYes" name="Applicant1SmokerGroup" type="radio" value="1">
				<label for="optApplicant1SmokerYes" class="msgLabel">Yes</label> 
			</span> 
			<span style="LEFT: 210px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant1SmokerNo" name="Applicant1SmokerGroup" type="radio" value="0" checked>
				<label for="optApplicant1SmokerNo" class="msgLabel">No</label> 
			</span>
		</span>
		<span id="spnGoodHealth1" style="LEFT: 0px; POSITION: absolute; TOP: 225px" class="msgLabel">
			Good Health?
			<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant1GoodHealth" name="Applicant1GoodHealthGroup" type="radio" value="1" checked>
				<label for="optApplicant1GoodHealth" class="msgLabel">Yes</label> 
			</span>
			<span style="LEFT: 210px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant1GoodHealthNo" name="Applicant1GoodHealthGroup" type="radio" value="0">
				<label for="optApplicant1GoodHealthNo" class="msgLabel">No</label> 
			</span>
		</span>
		<span id="spnCCJHistory1" style="LEFT: 0px; POSITION: absolute; TOP: 255px" class="msgLabel">
			CCJ History
			<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
				<select id="cboApplicant1CCJHistory" name="Applicant1CCJHistory" 
					style="WIDTH: 75px" class="msgCombo">
				</select>
			</span> 
		</span> 
	</span>
	<span id="spnApplicant2" style="LEFT: 350px; POSITION: absolute; TOP: 20px; visibility:hidden" class="msgLabel">
		<span id="txtApplicant2Name" name="Applicant2Name" style="LEFT: 0px; WIDTH: 250px; POSITION: absolute; TOP: -3px" tabindex = "-1" class="msgLabel"></span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 15px">
			<span style="LEFT: 0px; POSITION: absolute; TOP: -3px">
				<select id="cboApplicant2EmploymentStatus" name="Applicant2EmploymentStatus" style="WIDTH: 150px" class="msgCombo">
				</select>
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 45px">
			<span style="LEFT: 0px; POSITION: absolute">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtApplicant2GrossIncOrNetProfit" maxlength="10"	name="Applicant2GrossIncOrNetProfit" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt"> 
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 75px">
			<span style="LEFT: 0px; POSITION: absolute">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtApplicant2TotalMonthlyOutgoings" maxlength="5" name="Applicant2TotalMonthlyOutgoings" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt">
			</span>
			<span style="TOP: -10px; LEFT: 90px; POSITION: ABSOLUTE">
				<input id="btnApplicant2Outgoings" value="" type="button" style="WIDTH: 32px; HEIGHT: 32px" class="msgDDButton">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 105px">
			<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtApplicant2TotalMonthlyDebtRepayments" maxlength="5" name="Applicant2TotalMonthlyDebtRepayments" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt"> 
			</span>
			<span style="TOP: -10px; LEFT: 90px; POSITION: ABSOLUTE">
				<input id="btnApplicant2PersonalDebts" value="" type="button" style="WIDTH: 32px; HEIGHT: 32px" class="msgDDButton">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 135px">
			<span style="LEFT: 0px; POSITION: absolute">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtApplicant2TotalOSBalance" maxlength="10" name="Applicant2TotalOSBalance" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt">
			</span> 
		</span>
		<span id="spnTaxPayer2" style="LEFT: 0px; POSITION: absolute; TOP: 165px">
			<span style="LEFT: 0px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant2UKTaxPayerYes" name="Applicant2UKTaxPayerGroup" type="radio" value="1" checked>
				<label for="optApplicant2UKTaxPayerYes" class="msgLabel">Yes</label> 
			</span> 
			<span style="LEFT: 60px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant2UKTaxPayerNo" name="Applicant2UKTaxPayerGroup" type="radio" value="0">
				<label for="optApplicant2UKTaxPayerNo" class="msgLabel">No</label> 
			</span> 
		</span>
		<span id="spnSmoker2" style="LEFT: 0px; POSITION: absolute; TOP: 195px">
			<span style="LEFT: 0px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant2SmokerYes" name="Applicant2SmokerGroup" type="radio" value="1"> 
				<label for="optApplicant2SmokerYes" class="msgLabel">Yes</label>
			</span> 
			<span style="LEFT: 60px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant2SmokerNo" name="Applicant2SmokerGroup" type="radio" value="0" checked>
				<label for="optApplicant2SmokerNo" class="msgLabel">No</label> 
			</span> 
		</span>
		<span id="spnGoodHealth2" style="LEFT: 0px; POSITION: absolute; TOP: 225px">
			<span style="LEFT: 0px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant2GoodHealthYes" name="Applicant2GoodHealthGroup" type="radio" value="1" checked>
				<label for="optApplicant2GoodHealthYes" class="msgLabel">Yes</label> 
			</span> 
			<span style="LEFT: 60px; POSITION: absolute; TOP: -3px">
				<input id="optApplicant2GoodHealthNo" name="Applicant2GoodHealthGroup" type="radio" value="0">
				<label for="optApplicant2GoodHealthNo" class="msgLabel">No</label> 
			</span> 
		</span>
		<span id="spnCCJHistory2" style="LEFT: 0px; POSITION: absolute; TOP: 255px">
			<span style="LEFT: 0px; POSITION: absolute; TOP: -3px">
				<select id="cboApplicant2CCJHistory" name="Applicant2CCJHistory" title  ="CCJ History Help" style="WIDTH: 75px" class="msgCombo">
				</select>
			</span> 
		</span> 
	</span> 
</div>
<div id="divLoanDetails" style="HEIGHT: 80px; LEFT: 0px; POSITION: absolute; TOP: 375px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Loan Details...</strong>
		<% /* BMIDS00654 MDC 04/11/2002 */ %>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 20px" class="msgLabel">
			Amount Requested
			<span style="LEFT: 120px; POSITION: absolute">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtAmountRequested" name="AmountRequested" maxlength="10"  
							 style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt"> 
			</span>
		</span>
		<% /* BMIDS00654 MDC 04/11/2002 */ %>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 55px" class="msgLabel">
		Purchase Price
		<span style="LEFT: 120px; POSITION: absolute">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtPurchasePrice" name="PurchasePrice" maxlength="10"  
						 style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt"> 
		</span>
	</span>
	<span id="spnOSLoanAndMonthlyPayments" style="LEFT: 4px; POSITION: absolute; TOP: 50px; visibility:hidden" class="msgLabel">
		<span style="LEFT: 210px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Outstanding <br>Loan Amount
			<span style="LEFT: 100px; POSITION: absolute; TOP: 5px">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtOSLoanAmount"  maxlength="10" name="OSLoanAmount" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 410px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Total Monthly <br>Repayments
			<span style="LEFT: 100px; POSITION: absolute; TOP: 5px">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtTotalMonthlyRepayments" maxlength="10"	name="TotalMonthlyRepayments" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt">
			</span>
		</span> 
	</span> 
</div> 
</div>
</form>
<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 520px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<!-- File containing field attributes (remove if not required) -->
<!--  #include FILE="attribs/qq010attribs.asp" -->
<!--  #include FILE="Customise/QQ010Customise.asp" -->

<!-- Specify Code Here -->
<script language="JScript">
<!--
var m_sMetaAction					= null;
var m_sUserId						= null;
var m_sUserType						= null;
var m_sUnitId						= null;
var m_sMortgageApplicationValue		= null;
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
var m_sCustomer1Name				= null;
var m_sCustomer1Number				= null;
var m_sCustomer1VersionNumber		= null;
var m_sCustomer1RoleType			= null;
var m_sCustomer1Order				= null;
var m_sCustomer2Name				= null;
var m_sCustomer2Number				= null;
var m_sCustomer2VersionNumber		= null;
var m_sCustomer2RoleType			= null;
var m_sCustomer2Order				= null;
var m_sReadOnly						= null;
var m_sCurrency						= null;
var m_XMLBeforeImage				= null;
var m_sXMLApplicant1PersonalDebts	= null;
var m_sXMLApplicant2PersonalDebts	= null;
var m_sXMLApplicant1Outgoings		= null;
var m_sXMLApplicant2Outgoings		= null;
var scScreenFunctions;
var m_blnReadOnly = false;

<% /* MO 4/10/2002 BMIDS00578
//DB SYS4767 - MSMS Integration
// STB: MSMS0011 - Required for Intermediary search popup DC015Popup.asp
var m_sDistributionChannelId		= null;
var m_sDepartmentId					= null;

// STB: MSMS0011 - Store the IntermediaryGUID and name.
var m_sIntermediaryGUID				= null;
var m_sIntermediaryName				= null;
//DB End
*/ %>

var m_sIntermediaryGUID				= null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	<%/* scScreenFunctions = new scScrFunctions.ScreenFunctionsObject(); */%>
	<%/* DB End */%>	
 	<%/* SG 05/06/02 SYS4818 */%>	

	FW030SetTitles("Quick Quote","QQ010",scScreenFunctions);
	scScreenFunctions.SetCurrency(window,frmScreen);
	RetrieveContextData();
	SetMasks();
	Validation_Init();
	PopulateScreen();

	// MC SYS2564/SYS2757 for client customisation
	Customise();

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "QQ010");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	SetScreenOnReadOnly();
	
	<% /* MO 4/10/2002 BMIDS00578
	//DB SYS4767 - MSMS Integration
	// STB - MSMS0011: Intermediaries name is populated by searching from DC015Popup.
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtIndividualName");	
	//DB End
	*/ %>
	
	
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
	m_sUserId = scScreenFunctions.GetContextParameter(window,"idUserId",null);
	m_sUserType = scScreenFunctions.GetContextParameter(window,"idUserType",null);
	m_sUnitId = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	m_sMortgageApplicationValue = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","30");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B3069");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	m_sCustomer1Name = scScreenFunctions.GetContextParameter(window,"idCustomerName1","Applicant One");
	m_sCustomer1Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber1","70270");
	m_sCustomer1VersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1","1");
	m_sCustomer1RoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType1","1");
	m_sCustomer1Order = scScreenFunctions.GetContextParameter(window,"idCustomerOrder1","1");
	m_sCustomer2Name = scScreenFunctions.GetContextParameter(window,"idCustomerName2","Applicant Two");
	m_sCustomer2Number = scScreenFunctions.GetContextParameter(window,"idCustomerNumber2","70289");
	m_sCustomer2VersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber2","1");
	m_sCustomer2RoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType2","1");
	m_sCustomer2Order = scScreenFunctions.GetContextParameter(window,"idCustomerOrder2","2");
	m_sReadOnly	= scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sCurrency	= scScreenFunctions.GetContextParameter(window,"idCurrency","");
	<% /* MO 4/10/2002 BMIDS00578
	//DB SYS4767 - MSMS Integration
	// STB: MSMS0011 - Required for Intermediary search popup DC015Popup.asp
	m_sDistributionChannelId = scScreenFunctions.GetContextParameter(window, "idDistributionChannelId", null);
	m_sDepartmentId = scScreenFunctions.GetContextParameter(window, "idDepartmentId", null);
	//DB End
	*/ %>
}
function btnSubmit.onclick()
{
	var blnSaveQuickQuoteOk = true;
	
	if (m_sReadOnly != "1") {
	
		if (frmScreen.onsubmit() == true){
			<%
			// APS 01/03/00 - Make sure that we only route to DC330 if we have no errors from
			// WriteQuickQuoteDetails		
			%>
			<% /* MO 4/10/2002 BMIDS00578
			//DB SYS4767 - MSMS Integration
			//if (IsChanged() == true)
			//	blnSaveQuickQuoteOk = WriteQuickQuoteDetails();
			// STB: MSMS0011 - Ensure the intermediary is active.
			if (ValidateIntermediary(m_sIntermediaryGUID) == true)
			{
				if (IsChanged() == true)
				{				
					blnSaveQuickQuoteOk = WriteQuickQuoteDetails();		
				}
			}
			else
			{
				// If the intermediary is not active stay on page.
				blnSaveQuickQuoteOk = false;
			}				
			//DB End
			*/ %>
			<% /* BMIDS00654 MDC 04/11/2002	 */ %>
			scScreenFunctions.SetContextParameter(window, "idAmountRequested", frmScreen.txtAmountRequested.value);				
			<% /* BMIDS00654 MDC 04/11/2002 - End */ %>

			if (IsChanged() == true)
			{				
				blnSaveQuickQuoteOk = WriteQuickQuoteDetails();	
			}
		}
		else
			blnSaveQuickQuoteOk = false;
	}	
	
	if (blnSaveQuickQuoteOk == true) 
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
		{
			frmToGN300.submit();
			return;
		}
		frmAttitudeToBorrowing.submit();
	}
}
function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmApplicationMenu.submit();
}
function PopulateScreen()
{		
	<% /* BMIDS00654 MDC 04/11/2002	 */ %>
	scScreenFunctions.SetContextParameter(window, "idAmountRequested", "");				
	<% /* BMIDS00654 MDC 04/11/2002	- End */ %>

	SetApplicantNames();
	GetComboLists();
	GetQuickQuoteDetails();
	ShowSecondApplicantDetails();
	ShowLoanDetails();
	//DB SYS4767 - MSMS Integration
	// STB - MSMS0011: If source is indirect, show the intermediary.
	<% /* MO 4/10/2002 BMIDS00578 ShowOrHideiaryDetails(); */ %>
	//DB End
}
function GetQuickQuoteDetails()
{
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var XML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>
	XML.CreateRequestTag(window, null);
	var tagCUSTOMERROLELIST = XML.CreateActiveTag("CUSTOMERROLELIST");
	AddApplicantToRequest(XML, tagCUSTOMERROLELIST, "CUSTOMERROLE", 1);
	if (IsJointApplication() == true)AddApplicantToRequest(XML, tagCUSTOMERROLELIST, "CUSTOMERROLE", 2);
	XML.RunASP(document,"GetQuickQuoteDetails.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		ReadQuickQuoteApplicantDetails(XML);
		ReadLoanDetails(XML);
		//DB SYS4767 - MSMS Integration
		// STB - MSMS0011: Populate the application source controls.
		ReadApplicationSourceDetails(XML);
		//DB End
	}
	else alert(ErrorReturn[2]);
	ErrorTypes = null;
	ErrorReturn = null;
	XML = null;
}
function AddApplicantToRequest(XML, tagActiveTag, sTagName, iApplicantNumber)
{
	var sCustomerNumber, sCustomerVersionNumber;
	var sCustomerRoleType, sCustomerOrder;
	switch (iApplicantNumber)
	{
		case 1:	sCustomerNumber	= m_sCustomer1Number;
				sCustomerVersionNumber = m_sCustomer1VersionNumber;
				sCustomerRoleType = m_sCustomer1RoleType;
				sCustomerOrder = m_sCustomer1Order;
				break;
		case 2:	sCustomerNumber = m_sCustomer2Number;
				sCustomerVersionNumber = m_sCustomer2VersionNumber;
				sCustomerRoleType = m_sCustomer2RoleType;
				sCustomerOrder = m_sCustomer2Order;
				break;
		default: break;
	}
	XML.ActiveTag = tagActiveTag;
	XML.CreateActiveTag(sTagName);
	XML.CreateTag("APPLICATIONNUMBER",			m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",	m_sApplicationFactFindNumber);
	XML.CreateTag("CUSTOMERNUMBER",				sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER",		sCustomerVersionNumber);
	XML.CreateTag("CUSTOMERROLETYPE",			sCustomerRoleType);
	XML.CreateTag("CUSTOMERORDER",				sCustomerOrder);
}
function ReadQuickQuoteApplicantDetails(XML)
{
	var tagQUICKQUOTEFIRSTAPPLICANT = null;
	var	tagQUICKQUOTESECONDAPPLICANT = null;
	var tagQUICKQUOTEAPPLICANTDETAILS = XML.CreateTagList("QUICKQUOTEAPPLICANTDETAILS");
	if (XML.SelectTagListItem(0) == true)
	{
		tagQUICKQUOTEFIRSTAPPLICANT = XML.ActiveTag;
		ReadApplicantDetails(XML, 1);
		XML.CreateTagList("QUICKQUOTEPERSONALDEBTSLIST");
		if (XML.SelectTagListItem(0) == true)m_sXMLApplicant1PersonalDebts = XML.ActiveTag.xml;
		XML.ActiveTag = tagQUICKQUOTEFIRSTAPPLICANT;
		XML.CreateTagList("QUICKQUOTEOUTGOINGSLIST");
		if (XML.SelectTagListItem(0) == true)m_sXMLApplicant1Outgoings = XML.ActiveTag.xml;
	}
	if (IsJointApplication() == true)
	{
		XML.ActiveTagList = tagQUICKQUOTEAPPLICANTDETAILS;
		if (XML.SelectTagListItem(1) == true)
		{
			tagQUICKQUOTESECONDAPPLICANT = XML.ActiveTag;
			ReadApplicantDetails(XML, "2");
			XML.CreateTagList("QUICKQUOTEPERSONALDEBTSLIST");
			if (XML.SelectTagListItem(0) == true)m_sXMLApplicant2PersonalDebts = XML.ActiveTag.xml;
			XML.ActiveTag = tagQUICKQUOTESECONDAPPLICANT;
			XML.CreateTagList("QUICKQUOTEOUTGOINGSLIST");
			if (XML.SelectTagListItem(0) == true)m_sXMLApplicant2Outgoings = XML.ActiveTag.xml;
		}
	}
}
function ReadApplicantDetails(XML, sApplicantNumber)
{
	var sFieldName;
	sFieldName = "Applicant" + sApplicantNumber + "GrossIncOrNetProfit";
	frmScreen(sFieldName).value = XML.GetTagText("ANNUALINCOMEORNETPROFIT");
	sFieldName = "Applicant" + sApplicantNumber + "TotalMonthlyOutgoings";
	frmScreen(sFieldName).value = XML.GetTagText("TOTALMONTHLYOUTGOINGS");
	sFieldName = "Applicant" + sApplicantNumber + "TotalMonthlyDebtRepayments";
	frmScreen(sFieldName).value = XML.GetTagText("TOTALDEBTSMONTHLYPAYMENTS");
	sFieldName = "Applicant" + sApplicantNumber + "TotalOSBalance";
	frmScreen(sFieldName).value = XML.GetTagText("TOTALDEBTSOUTSTANDINGBALANCE");
	sFieldName = "Applicant" + sApplicantNumber + "CCJHistory";
	frmScreen(sFieldName).value = XML.GetTagText("CCJDETAILSSEQUENCENUMBER");
	sFieldName = "Applicant" + sApplicantNumber + "EmploymentStatus";
	frmScreen(sFieldName).value = XML.GetTagText("EMPLOYMENTSTATUS");
	sFieldName = "Applicant" + sApplicantNumber + "GoodHealthGroup";
	scScreenFunctions.SetRadioGroupValue(frmScreen, sFieldName, XML.GetTagText("GOODHEALTH"));
	sFieldName = "Applicant" + sApplicantNumber + "SmokerGroup";
	scScreenFunctions.SetRadioGroupValue(frmScreen, sFieldName, XML.GetTagText("SMOKER"));
	sFieldName = "Applicant" + sApplicantNumber + "UKTaxPayerGroup";
	scScreenFunctions.SetRadioGroupValue(frmScreen, sFieldName, XML.GetTagText("UKTAXPAYER"));
	<% /* BMIDS00654 MDC 01/11/2002 */ %>
	if (sApplicantNumber == "1")
		frmScreen.txtAmountRequested.value = XML.GetTagText("AMOUNTREQUESTED");
	<% /* BMIDS00654 MDC 01/11/2002 - End */ %>
}

function ReadLoanDetails(XML)
{
	XML.ActiveTag = null;
	XML.CreateTagList("APPLICATIONFACTFIND");
	if (XML.SelectTagListItem(0) == true)
	{
		m_XMLBeforeImage = XML.CreateFragment();
		frmScreen.txtPurchasePrice.value			= XML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE");
		frmScreen.txtOSLoanAmount.value				= XML.GetTagText("OUTSTANDINGLOANAMOUNT");
		frmScreen.txtTotalMonthlyRepayments.value	= XML.GetTagText("TOTALMONTHLYREPAYMENT");
	}
}
function GetComboLists()
{
	<%/* SG 10/06/02 SYS4848 Start */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("EmploymentStatus", "CCJHistory", "Direct/Indirect"); 
	<%/* SG 10/06/02 SYS4848 End */%>
	
	var bSuccess = false;
	if (XML.GetComboLists(document, sGroups) == true)
	{
		bSuccess = true;
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboApplicant1EmploymentStatus,"EmploymentStatus",true);
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboApplicant2EmploymentStatus,"EmploymentStatus",true);
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboApplicant1CCJHistory,"CCJHistory",true);
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboApplicant2CCJHistory,"CCJHistory",true);
		//DB SYS4767 - MSMS Integration
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboDirectIndirect,"Direct/Indirect",true);
		//DB End
	}
	if(bSuccess == false)
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		DisableMainButton("Submit");
	}
}
function ShowSecondApplicantDetails()
{
	if (IsJointApplication() == true)scScreenFunctions.ShowCollection(spnApplicant2);
	else scScreenFunctions.HideCollection(spnApplicant2);
}
function frmScreen.btnApplicant1Outgoings.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sXMLApplicant1Outgoings;
	ArrayArguments[1] = m_sCustomer1Name;
	ArrayArguments[2] = m_sCustomer1Number;
	ArrayArguments[3] = m_sCustomer1VersionNumber;
	ArrayArguments[4] = m_sReadOnly;
	ArrayArguments[5] = m_sApplicationNumber;
	ArrayArguments[6] = m_sApplicationFactFindNumber;
	ArrayArguments[7] = m_sCurrency;
	sReturn = scScreenFunctions.DisplayPopup(window, document, "qq020.asp", ArrayArguments, 470, 420);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sXMLApplicant1Outgoings = sReturn[1];
		frmScreen.txtApplicant1TotalMonthlyOutgoings.value = sReturn[2];
	}
}
function frmScreen.btnApplicant2Outgoings.onclick() 
{
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sXMLApplicant2Outgoings;
	ArrayArguments[1] = m_sCustomer2Name;
	ArrayArguments[2] = m_sCustomer2Number;
	ArrayArguments[3] = m_sCustomer2VersionNumber;
	ArrayArguments[4] = m_sReadOnly;
	ArrayArguments[5] = m_sApplicationNumber;
	ArrayArguments[6] = m_sApplicationFactFindNumber;
	ArrayArguments[7] = m_sCurrency;
	sReturn = scScreenFunctions.DisplayPopup(window, document, "qq020.asp", ArrayArguments, 470, 420);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sXMLApplicant2Outgoings = sReturn[1];
		frmScreen.txtApplicant2TotalMonthlyOutgoings.value = sReturn[2];
	}
}
function frmScreen.btnApplicant1PersonalDebts.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sXMLApplicant1PersonalDebts;
	ArrayArguments[1] = m_sCustomer1Name;
	ArrayArguments[2] = m_sCustomer1Number;
	ArrayArguments[3] = m_sCustomer1VersionNumber;
	ArrayArguments[4] = m_sReadOnly;
	ArrayArguments[5] = m_sApplicationNumber;
	ArrayArguments[6] = m_sApplicationFactFindNumber;
	ArrayArguments[7] = m_sCurrency;
	sReturn = scScreenFunctions.DisplayPopup(window, document, "qq030.asp", ArrayArguments, 470, 440);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sXMLApplicant1PersonalDebts = sReturn[1];
		frmScreen.txtApplicant1TotalMonthlyDebtRepayments.value = sReturn[2];
		frmScreen.txtApplicant1TotalOSBalance.value = sReturn[3];
	}
}
function frmScreen.btnApplicant2PersonalDebts.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sXMLApplicant2PersonalDebts;
	ArrayArguments[1] = m_sCustomer2Name;
	ArrayArguments[2] = m_sCustomer2Number;
	ArrayArguments[3] = m_sCustomer2VersionNumber;
	ArrayArguments[4] = m_sReadOnly;
	ArrayArguments[5] = m_sApplicationNumber;
	ArrayArguments[6] = m_sApplicationFactFindNumber;
	ArrayArguments[7] = m_sCurrency;
	sReturn = scScreenFunctions.DisplayPopup(window, document, "qq030.asp", ArrayArguments, 470, 440);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sXMLApplicant2PersonalDebts = sReturn[1];
		frmScreen.txtApplicant2TotalMonthlyDebtRepayments.value = sReturn[2];
		frmScreen.txtApplicant2TotalOSBalance.value = sReturn[3];
	}
}
function ShowLoanDetails()
{
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var XML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>
	var ValidationList = new Array("F", "T");
	var blnValidationList = false;
	blnValidationList = XML.IsInComboValidationList(document,"TypeOfMortgage",m_sMortgageApplicationValue, ValidationList);
	if (blnValidationList == true)scScreenFunctions.ShowCollection(spnOSLoanAndMonthlyPayments);
	else scScreenFunctions.HideCollection(spnOSLoanAndMonthlyPayments);
	XML = null;
	ValidationList = null;
}
function SetScreenOnReadOnly()
{
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetCollectionToReadOnly(spnApplicant1);
		scScreenFunctions.SetCollectionToReadOnly(divLoanDetails);
		if (IsJointApplication() == true)
			scScreenFunctions.SetCollectionToReadOnly(spnApplicant2);
		btnSubmit.focus();	
	}
	else
		<%/* SG 06/06/02 SYS4828 - added If statement */%>
		if(frmScreen.cboDirectIndirect.disabled != true)
		{	
			//DB SYS4767 - MSMS Integration
			//frmScreen.cboApplicant1EmploymentStatus.focus();
			frmScreen.cboDirectIndirect.focus();
			//DB End
		}
}
function SetApplicantNames()
{
	txtApplicant1Name.innerHTML = "<strong>" + m_sCustomer1Name + "</strong>";
	if (IsJointApplication() == true)txtApplicant2Name.innerHTML = "<strong>" + m_sCustomer2Name + "</strong>";
}
function WriteQuickQuoteDetails()
{
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var XML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>
	var tagREQUEST = XML.CreateRequestTag(window,"SAVE");
	var tagQUICKQUOTEAPPLICANTLIST = XML.CreateActiveTag("QUICKQUOTEAPPLICANTDETAILSLIST");
	AddApplicantToRequest(XML, tagQUICKQUOTEAPPLICANTLIST, "QUICKQUOTEAPPLICANTDETAILS", 1);
	WriteApplicantDetails(XML, 1);
	WritePersonalDebts(XML, 1);
	WriteOutgoings(XML, 1);
	if (IsJointApplication() == true)
	{
		AddApplicantToRequest(XML, tagQUICKQUOTEAPPLICANTLIST, "QUICKQUOTEAPPLICANTDETAILS", 2);
		WriteApplicantDetails(XML, 2);
		WritePersonalDebts(XML, 2);
		WriteOutgoings(XML, 2);
	}
	WriteLoanDetails(XML, tagREQUEST);
	//DB SYS4767 - MSMS Integration
	// STB: MSMS0011 - Append the Application Source information to the save request.
	WriteApplicationSourceDetails(XML, tagREQUEST);
	//DB End
	// 	XML.RunASP(document, "WriteQuickQuoteDetails.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "WriteQuickQuoteDetails.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	var blnSaveQuickQuoteDetails = XML.IsResponseOK();
	XML = null;
	return blnSaveQuickQuoteDetails;
}
function WriteApplicantDetails(XML, iApplicantNumber)
{
	var sFieldName;
	sFieldName = "Applicant" + iApplicantNumber + "GrossIncOrNetProfit";
	XML.CreateTag("ANNUALINCOMEORNETPROFIT", frmScreen(sFieldName).value);
	sFieldName = "Applicant" + iApplicantNumber + "TotalMonthlyOutgoings";
	XML.CreateTag("TOTALMONTHLYOUTGOINGS", frmScreen(sFieldName).value);
	sFieldName = "Applicant" + iApplicantNumber + "TotalMonthlyDebtRepayments";
	XML.CreateTag("TOTALDEBTSMONTHLYPAYMENTS", frmScreen(sFieldName).value);
	sFieldName = "Applicant" + iApplicantNumber + "TotalOSBalance";
	XML.CreateTag("TOTALDEBTSOUTSTANDINGBALANCE", frmScreen(sFieldName).value);
	sFieldName = "Applicant" + iApplicantNumber + "CCJHistory";
	XML.CreateTag("CCJDETAILSSEQUENCENUMBER", frmScreen(sFieldName).value);
	sFieldName = "Applicant" + iApplicantNumber + "EmploymentStatus";
	XML.CreateTag("EMPLOYMENTSTATUS", frmScreen(sFieldName).value);
	sFieldName = "Applicant" + iApplicantNumber + "GoodHealthGroup";
	XML.CreateTag("GOODHEALTH", scScreenFunctions.GetRadioGroupValue(frmScreen, sFieldName));
	sFieldName = "Applicant" + iApplicantNumber + "SmokerGroup";
	XML.CreateTag("SMOKER", scScreenFunctions.GetRadioGroupValue(frmScreen, sFieldName));
	sFieldName = "Applicant" + iApplicantNumber + "UKTaxPayerGroup";
	XML.CreateTag("UKTAXPAYER", scScreenFunctions.GetRadioGroupValue(frmScreen, sFieldName));
	<% /* BMIDS00654 MDC 01/11/2002 */ %>
	if (iApplicantNumber == 1)
		XML.CreateTag("AMOUNTREQUESTED", frmScreen.txtAmountRequested.value)
	<% /* BMIDS00654 MDC 01/11/2002 - End */ %>
}
function WritePersonalDebts(XML, iApplicantNumber)
{
	var sXML;
	switch (iApplicantNumber)
	{
		case 1:	sXML = m_sXMLApplicant1PersonalDebts;
				break;
		case 2:	sXML = m_sXMLApplicant2PersonalDebts;
				break;
		default: break;
	}
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XMLPersonalDebts = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var XMLPersonalDebts = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>
	XMLPersonalDebts.LoadXML(sXML);
	XML.AddXMLBlock(XMLPersonalDebts.XMLDocument);
	XMLPersonalDebts = null;
}
function WriteOutgoings(XML, iApplicantNumber)
{
	var sXML;
	switch (iApplicantNumber)
	{
		case 1:	sXML = m_sXMLApplicant1Outgoings;
				break;
		case 2:	sXML = m_sXMLApplicant2Outgoings;
				break;
		default: break;
	}
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XMLOutgoings = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var XMLOutgoings = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>
	XMLOutgoings.LoadXML(sXML);
	XML.AddXMLBlock(XMLOutgoings.XMLDocument);
	XMLOutgoings = null;
}
function WriteLoanDetails(XML, tagREQUEST)
{
	XML.ActiveTag = tagREQUEST;
	XML.CreateActiveTag("UPDATE");
	XML.SetAttribute("TYPE","BEFORE");
	XML.AddXMLBlock(m_XMLBeforeImage);
	XML.ActiveTag = tagREQUEST;
	XML.CreateActiveTag("UPDATE");
	XML.SetAttribute("TYPE","AFTER");
	XML.CreateActiveTag("APPLICATIONFACTFIND");
	XML.CreateTag("APPLICATIONNUMBER",				m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",		m_sApplicationFactFindNumber);
	XML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE",	frmScreen.txtPurchasePrice.value);
	XML.CreateTag("OUTSTANDINGLOANAMOUNT",			frmScreen.txtOSLoanAmount.value);
	XML.CreateTag("TOTALMONTHLYREPAYMENT",			frmScreen.txtTotalMonthlyRepayments.value);
	//DB SYS4767 - MSMS Integration
	// STB: MSMS0011 - Append the Application source.
	XML.CreateTag("DIRECTINDIRECTBUSINESS", frmScreen.cboDirectIndirect.value);	
	//DB End
}
function IsJointApplication()
{
	var blnJointApplication = false;
	blnJointApplication = ((m_sCustomer2Number != "") && (m_sCustomer2RoleType == "1"))
	return blnJointApplication;
	}

<% /* MO 4/10/2002 BMIDS00578
//DB SYS4767 - MSMS Integration
// STB - MSMS0011: Hide or show the intermediary based upon whether the       
// application source is direct or indirect.                                 
function ShowOrHideIntermediaryDetails()
{
	var iSelectedIdx = frmScreen.cboDirectIndirect.selectedIndex;
				
	if((iSelectedIdx != -1) && (scScreenFunctions.IsOptionValidationType(frmScreen.cboDirectIndirect, iSelectedIdx, "I")))
	{
		// Indirect application, show intermediary details.
		scScreenFunctions.ShowCollection(spnIndividualName);
		scScreenFunctions.ShowCollection(spnIntermediarySearch);				
	}
	else
	{
		// Direct application, hide the intermediary details.
		scScreenFunctions.HideCollection(spnIndividualName);
		scScreenFunctions.HideCollection(spnIntermediarySearch);				
	}
} */
%>


/* STB - MSMS0011: Store XML pertaining to the Application Source and         */
/* populate the screen elements.                                             */
function ReadApplicationSourceDetails(XML)
{
	// Attempt to find the application element beneath the latest details element.
	var LatestDetailsXML = XML.SelectTag(null, "APPLICATIONLATESTDETAILS");
	var ApplicationXML = XML.SelectTag(LatestDetailsXML, "APPLICATION");
	
	<% /* MO 4/10/2002 BMIDS00578  if (ApplicationXML == null)
	{		
		// MSMS0036 - If no application XML found then default app source.
		ReadIntermediaryDetails(null);
	}
	else
	{ 
	// Obtain an Intermediary XML block from it's GUID.
	var IntermediaryXML = GetIntermediaryXML(XML.GetTagText("INTERMEDIARYGUID"));												
	*/ %>	
	<%/*BMIDS00670 Check to see if we hace records. */%>
	if (ApplicationXML != null)
	{
		// Move the active tag down to the APPLICATIONFACTFIND.
		XML.SelectTag(ApplicationXML, "APPLICATIONFACTFIND");
	}	
	// Populate the screen elements with the data.		
	<%/*BMIDS00670 Check to see if we hace records. */%>
	if (XML.SelectTagListItem(0) == true) 
	{
		frmScreen.cboDirectIndirect.value = XML.GetTagText("DIRECTINDIRECTBUSINESS");
	}
	<% /* MO 4/10/2002 BMIDS00578  
	// Populate the intermediary screen elements from this cloned XML block.				
	ReadIntermediaryDetails(IntermediaryXML);		
	*/ %>
}


<% /* MO 4/10/2002 BMIDS00578  
// STB - MSMS0011: Store XML pertaining to the Intermediary and populate the  
// screen elements.                                                         
function ReadIntermediaryDetails(IntermediaryXML)
{
	if (IntermediaryXML != null)
	{		
		// Store the intermediary details for saving later.		
		IntermediaryXML.SelectTag(null, "INTERMEDIARYINDIVIDUAL");
		m_sIntermediaryGUID = IntermediaryXML.GetTagText("INTERMEDIARYGUID");
		m_sIntermediaryName = IntermediaryXML.GetTagText("FORENAME") + ' ' + IntermediaryXML.GetTagText("SURNAME");			
	}
	// MSMS0036 - Default values if no XML found.
	else
	{
		m_sIntermediaryGUID = "";
		m_sIntermediaryName = "";	
	}
	
	// Populate the screen elements with the data.
	frmScreen.txtIndividualName.value = m_sIntermediaryName;
}
*/ %>

<% /*
// STB - MSMS0011: Build an Intermediary XML object from the given GUID 
function GetIntermediaryXML(sIntermediaryGUID)
{
	if (sIntermediaryGUID == "")
	{
		return null;
	}
	else
	{
		IntermedXML = new scXMLFunctions.XMLObject();
		IntermedXML.CreateRequestTag(window, null)
		IntermedXML.CreateActiveTag("INTERMEDIARY");
		IntermedXML.CreateTag("INTERMEDIARYGUID", sIntermediaryGUID);
		IntermedXML.RunASP(document, "GetIndividualIntermediary.asp");

		return IntermedXML;
	}	
}
*/
%>

<% /* MO 4/10/2002 BMIDS00578
// STB - MSMS0011: Open DC015Popup to select an Intermediary and then call  
// ReadIntermediaryDetails with the returned XML.                            
function frmScreen.btnIntermedSearch.onclick()
{	
	var ArrayArguments = new Array();
	
	// All of these are obtained from context parameters.
	ArrayArguments[0] = m_sDistributionChannelId;
	ArrayArguments[1] = m_sDepartmentId;
	ArrayArguments[2] = m_sUnitId;
	ArrayArguments[3] = m_sUserId;
	
	// Display the Intermediary Search Popup window.
	var sReturn = scScreenFunctions.DisplayPopup(window, document, "DC015Popup.asp", ArrayArguments, 650, 420);

	if (sReturn != null)
	{								
		// Set the changed flag to true.
		FlagChange(true);
		
		// If an Intermediary was selected then populate the controls with it.
		var IntermediaryXML = new scXMLFunctions.XMLObject();
		
		// Build an Intermediary XML element from the returned values.
		IntermediaryXML.CreateActiveTag("INTERMEDIARYINDIVIDUAL");
		IntermediaryXML.CreateTag("INTERMEDIARYGUID", sReturn[2]);
		
		// sReturn[1] contains the full, concatonated name.
		IntermediaryXML.CreateTag("FORENAME", sReturn[1]);
		IntermediaryXML.CreateTag("SURNAME", " ");
		
		// Populate the intermediary controls.
		ReadIntermediaryDetails(IntermediaryXML);			
	}		
}
*/ %>

<% /* MO 4/10/2002 BMIDS00578  STB - MSMS0011: Delegate the call to ShowOrHideIntermediaryDetails().     */
/* function frmScreen.cboDirectIndirect.onclick()
{
	ShowOrHideIntermediaryDetails();
} */ %>



// STB - MSMS0011: Append the Application Source - specific XML data to the  
// save request.                                                             
function WriteApplicationSourceDetails(RequestXML, tagREQUEST)
{			
	// Append an UPDATE tag to the Application table onto the request.	
	RequestXML.ActiveTag = tagREQUEST;
	RequestXML.CreateActiveTag("UPDATE");
	RequestXML.SetAttribute("TYPE", "AFTER");
	
	// Append the primary key and the field(s) we want to update.
	RequestXML.CreateActiveTag("APPLICATION");
	RequestXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);	
	RequestXML.CreateTag("INTERMEDIARYGUID", m_sIntermediaryGUID);
}

<% /* MO 4/10/2002 BMIDS00578  
// STB - MSMS0011: The Intermediary must have a status of Active and an	    
// DRC - MSMS0046: ActiveToDate must be now or future
// ActiveTo date greater than or equal to today's date.                         
function ValidateIntermediary()
{		
	var bValid = true;
	
	// If the source is direct then Intermediaries are irrelevant (return true).
	if (scScreenFunctions.GetComboValidationType(frmScreen, "cboDirectIndirect") == "I")
	{					
		// Load the required intermediary record.
		var IntermedXML = GetIntermediaryXML(m_sIntermediaryGUID);

		// Get the ActiveTo date.
		IntermedXML.SelectTag(null, "INTERMEDIARY");
		var sActiveToDate = IntermedXML.GetTagText("INTERMEDIARYACTIVETO");
		
		// Get the IntermediaryStatus Value ID.
		IntermedXML.SelectTag(null, "INTERMEDIARYINDIVIDUAL");
		var lIntermediaryStatus = IntermedXML.GetTagText("INTERMEDIARYSTATUS");
		
		// Attempt to get the validation type if we have a ValueID.
		if (lIntermediaryStatus != "")
		{
			// Get the Validation for the Intermediary status (use private routine rather
			// than the common routine as we have no <SELECT> control to use.
			var sIntermediaryStatus = GetValidationFromValue("IntermediaryStatus", lIntermediaryStatus);
		}
					
		// If the ActiveTo date is not null and LESS than today or the 
		// status is not (A)ctive then it is INvalid.		
		if ((sActiveToDate != "" && scScreenFunctions.CompareDateStringToSystemDate(sActiveToDate, "<")) | 
			(sIntermediaryStatus != "A"))
		{
		    // DRC MS0048 - allow continuation even if invalid
			bValid = confirm("Warning - The Intermediary is not active. Continue anyway?");
		}		
	}
	
	return bValid;
}
*/ %>

/* STB - MSMS0011: Return the first Validation type for specified group and  */
/* value ID - used when we have no combo control to use the generic XML      */
/* functions.                                                                */
function GetValidationFromValue(sGroupName, lValueID)
{
	var sReturn;
	var XML = new scXMLFunctions.XMLObject();			
	var sGroups = new Array(sGroupName);
		
	// Get the combo data for the named group.
	if (XML.GetComboLists(document, sGroups) == true)
	{		
		// Find the first validation for the specified valueID.
		var sPattern = "//LISTENTRY[VALUEID=" + lValueID + "]" + "/VALIDATIONTYPELIST"		
		XML.SelectSingleNode(sPattern);		
		sReturn = XML.GetTagText("VALIDATIONTYPE");
	}
	else
	{
		alert("Unable to locate information for combogroup " + sGroupName);
	}

	XML = null;	
	return sReturn;
//DB End
}
-->
</script>
</body>
</html>

