<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      dc074.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Add/Edit Mortgage Loan
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		02/02/2000	ie5 changes 
AY		14/02/00	Change to msgButtons button types
APS		15/02/2000	Change to only allow redemption dates in the past
AD		02/03/2000	Fixed SYS0172
AY		30/03/00	New top menu/scScreenFunctions change
MH      04/05/00    SYS0661 Redemption date validation and tabbing.
MC		12/05/00	Fixed SYS0688 
BG		17/05/00	SYS0752 Removed Tooltips
MC		08/06/00	SYS0866 Standardise behaviour of Balance & Payment fields
MH      22/06/00    SYS0933 Readonly processing
CL		05/03/01	SYS1920 Read only functionality added
AT		10/04/02	SYS4359 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
PSC		30/07/2002	BMIDS00006  Amend to use new Mortgage Account structure
MDC		30/08/2002	BMIDS00336	CCWP1 Credit Check and Bureau Download
GHun	05/09/2002	BMIDS00406	Remove the word "mortgage" from the screen
TW      09/10/2002  Modified to incorporate client validation - SYS5115
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
AW		04/12/2002	BM0128		Changed orig loan amount, amount outstanding field lengths to 9
MV		19/02/2003	BM0356		Amended PopulateScreen(); RetrieveContext()
GD		10/06/2003	BM0356		Use validation types, not value ids
GD		18/06/2003	BM0356		Use validation types, not value ids 
GD		23/06/2003	BM0356		Stop XML being overwritten
MC		17/05/2004	bmids756	New Fields (ProductStartDate,IndexCode, Variance) added - Regulation change
KW      26/05/2004  BMIDS773    Regulation Amendments to text fields
MC		09/07/2004	BMIDS756	Part and Part lable to hide if right option not selected.
GHun	16/07/2004	BMIDS756	Don't clear Part&Part amount when hiding it
SR		10/08/2004	BMIDS815	Add two new fields 'CurrentInterestRateStep','RemainingStepDuration'
SR		01/09/2004  BMIDS815	Add two new fields 'Remaining Int Only Amount','Remaining Cap & Int Amount'
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
GHun	26/07/2005	MAR10		Made admin system specific label more generic
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
MF		15/09/2005	MAR30		IncomeCalcs modified to send ActivityID into calculation
HMA     26/09/2005  MAR46       WP15. Add Rate expiry Date
GHun	12/10/2005	MAR46		Reordered fields, tidied layout and fixed ProductCode maxlength
PSC		16/05/2006	MAR1798		Run Critical Data 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

EPSOM Specific History:

Prog	Date		AQR			Description
DRC     08/06/2006  EP669       Remove the 'Another' Button altogether
AShaw	14/11/2006	EP2_55		Add code for Original LTV and Avail for Disbursement.
PSC		15/01/2007	EP2_741		Add Flexible Product Indicator and Original Income Status
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex="-1" 
type="text/x-scriptlet" width="1" viewastext></object>
<script src="Validation.js" language="JScript" type="text/javascript"></script>	

<form id="frmToDC073" method="post" action="dc073.asp" style="DISPLAY: none"></form>

<% /*span to keep tabbing within this screen */%>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" year4 validate="onchange" mark>
	<div id="divBackground" style="HEIGHT: 620px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
		
		<table style="MARGIN-LEFT: 10px; WIDTH: 594px" class="msgLabel" border="0">
			<tr>
				<td width="30%">Loan Account Number</td>
				<td width="70%">
					<input id="txtLoanAccountNumber" type="text" maxlength="20" style="WIDTH: 150px" class="msgTxt">
				</td>
			</tr>
			<tr>
				<td>Loan Start Date</td>
				<td>
					<input id="txtStartDate" type="text" maxlength="10" style="WIDTH: 75px" class="msgTxt">
				</td>
			</tr>
			<tr>
				<td>Product Start Date</td>
				<td>
					<input id="txtProductStartDate" type="text" maxlength="10" style="WIDTH: 75px" class="msgTxt">
				</td>
			</tr>
			<tr>
				<td>Product Code</td>
				<td>
					<input id="txtMortgageProductCode" type="text" maxlength="6" style="WIDTH: 75px" class="msgTxt">
				</td>
			</tr>
			<tr>
				<td>Product Description</td>
				<td>
					<input id="txtMortgageProductDescription" type="text" maxlength="50" style="WIDTH: 220px" class="msgTxt" name="txtMortgageProductDescription">
				</td>
			</tr>
			<tr>
				<td>Interest Rate</td>
				<td>
					<input id="txtInterestRate" type="text" maxlength="7" style="WIDTH: 75px" class="msgTxt" name="txtInterestRate">
				</td>
			</tr>
			<% /* MAR46  Remove Remaining Duration and Current Step. Add Current Rate Expiry Date */ %>
			<tr>
				<td>Current Rate Expiry Date</td>
				<td>
					<input id="txtCurrentRateExpiryDate" type="text" maxlength="10" style="WIDTH: 75px" class="msgTxt" name="txtCurrentRateExpiryDate">
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<table class="msgLabel" width="100%" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td width="30%">Outstanding Balance</td>
							<td width="28%" style="padding-left: 2px">
								<input id="txtOutstandingBalance" type="text" maxlength="9" style="WIDTH: 75px" class="msgTxt" NAME="txtOutstandingBalance">
							</td>
							<td>
							<td width="26%">Available for Disbursement</td>
							<td width="16%">
								<input id="txtAvailableForDisbursement" type="text" maxlength="9" style="WIDTH: 75px" class="msgTxt" NAME="txtAvailableForDisbursement">
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<table class="msgLabel" width="100%" border="0" cellpadding="0" cellspacing="0" ID="Table2">
						<tr>
							<td width="30%">Redemption Status</td>
							<td width="28%" style="padding-left: 2px">
								<select id="cboRedemptionStatus" style="WIDTH: 155px" class="msgCombo" menusafe="true" NAME="cboRedemptionStatus"></select>
							</td>
							<td>
								<div id="spnRedemptionDate" class="msgLabel">
									<table class="msgLabel" width="100%" cellpadding="0" cellspacing="0" ID="Table3">
										<tr>
											<td width="62%">Redemption Date</td>
											<td>
												<input id="txtRedemptionDate" type="text" maxlength="10" style="WIDTH: 75px" class="msgTxt" NAME="txtRedemptionDate">
											</td>
										</tr>
									</table>
								</div>
							</td>
						</tr>
					</table>
				</td>
			</tr>

			<tr>
				<td colspan="2">
					<table class="msgLabel" width="100%" cellpadding="0" cellspacing="0" ID="Table1">
						<tr>
							<td width="30%">Original Loan Amount</td>
							<td width="28%" style="padding-left: 2px">
								<input id="txtOriginalLoanAmount" type="text" maxlength="9" style="WIDTH: 75px" class="msgTxt" name="txtOriginalLoanAmount">
							</td>
							<td width="26%">Original LTV</td>
							<td width="16%">
								<input id="txtOriginalLTV" type="text" maxlength="10" style="WIDTH: 75px" class="msgTxt" NAME="txtOriginalLTV">
							</td>
						</tr>
					</table>
				</td>
			</tr>

			<%/* SR 01/09/2004 : BMIDS815 */%>
			<tr>
				<td colspan="2">
					<table id="tblRemainingBalance" class="msgLabel" width="100%" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td width="30%">Remaining Int Only Balance</td>
							<td width="28%" style="padding-left: 2px">
								<input id="txtRemainingIntOnlyBalance" type="text" maxlength="6" style="WIDTH: 75px" class="msgTxt" name="txtRemainingIntOnlyBalance">
							</td>
							<td>
							<td width="26%">Remaining Cap &amp; Int Amount</td>
							<td width="16%">
								<input id="txtRemainingCapAndIntAmount" type="text" maxlength="6" style="WIDTH: 75px" class="msgTxt">
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<%/*  SR 01/09/2004 : BMIDS815 - End  */%>
			<tr>
				<td>Purpose of Loan</td>
				<td>
					<select id="cboPurposeOfLoan" style="WIDTH: 220px" class="msgCombo" name="cboPurposeOfLoan" menusafe="true">
					</select>
				</td>
			</tr>
			<tr>
				<td><label id="idOriginalTerm">Original Term</label></td>
				<td>
					<input id="txtOriginalTermYears" type="text" maxlength="3" style="WIDTH: 35px" class ="msgTxt">
					Years &nbsp;
					<input id="txtOriginalTermMonths" type="text" maxlength="2" style="WIDTH: 35px" class="msgTxt">
					Months
				</td>
			</tr>
			<tr>
				<td><label id="idOutstandingTerm">Outstanding Term</label></td>
				<td>
					<input id="txtOutstandingTermYears" type="text" maxlength="3" style="WIDTH: 35px" class="msgTxt">
					Years &nbsp;
					<input id="txtOutstandingTermMonths" type="text" maxlength="2" style="WIDTH: 35px" class="msgTxt">
					Months	
				</td>
			</tr>
			<tr>
				<td>Admin Calculated Outstanding Term</td>
				<td>
					<input id="txtAdminOutstandingTerm" type="text" maxlength="5" style="WIDTH: 75px" class="msgTxt">
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<table id="tlbRepaymentType" class="msgLabel" width="100%" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td width="30%">Repayment Type</td>
							<td width="28%" style="padding-left: 2px">
								<select id="cboRepaymentType" style="WIDTH: 155px" class="msgCombo" menusafe="true"></select>
							</td>
							<td valign="top">
								<div id="spnOriginalPartAndPartIntOnlyAmount" class="msgLabel">
									<table id="tlbOriginalPartAndPartIntOnlyAmount" cellspacing="0" cellpadding="0" border="0" class="msgLabel">
										<tr>
											<td width="62%">Original Part and Part Int Only Amount</td>
											<td>
												<input id="txtOriginalPartAndPartIntOnlyAmount" type="text" maxlength="6" style="WIDTH: 75px" class="msgTxt">
											</td>
										</tr>
									</table>
								</div>
							</td>
						</tr>
					</table>
				</td>
				
			</tr>
			<tr>
				<td>Monthly Repayment</td>
				<td>
					<input id="txtMonthlyRepayment" type="text" maxlength="9" style="WIDTH: 75px" class="msgTxt">
				</td>
			</tr>
			<tr>
				<td>Flexible Product</td>
				<td>
					<input id="optFlexibleProductYes" name="FlexibleProductIndicator" type="radio" value="1"><label for="optFlexibleProductYes" class="msgLabel">Yes</label>
					<input id="optFlexibleProductNo" name="FlexibleProductIndicator" type="radio" value="0" checked ><label for="optFlexibleProductNo" class="msgLabel">No</label>
				</td>
			</tr>			

			<tr>
				<td>Original Income Status</td>
				<td>
					<select id="cboOriginalIncomeStatus" style="WIDTH: 220px" class="msgCombo" name="cboOriginalIncomeStatus" menusafe="true">
					</select>
				</td>
			</tr>
			
			<tr>
				<td>Index Code</td>
				<td>
					<input id="txtIndexCode" type="text" maxlength="2" style="WIDTH: 75px" class="msgTxt">
				</td>
			</tr>
			<tr>
				<td>Variance</td>
				<td>
					<input id="txtVariance" type="text" maxlength="8" style="WIDTH: 75px" class="msgTxt">
				</td>
			</tr>
			<tr>
				<td>Funds Disbursed</td>
				<td>
					<input id="txtDisbursedAmount" type="text" maxlength="9" style="WIDTH: 75px" class="msgTxt">
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<table class="msgLabel" width="100%" cellpadding="0" cellspacing="0">
						<tr>
							<td width="30%">Collections CCI Indicator</td>
							<td width="28%" style="padding-left: 2px">
								<input id="txtCCIIndicator" type="text" maxlength="1" style="WIDTH: 75px" class="msgTxt">
							</td>
							<td width="26%">
								CCA Indicator
							</td>
							<td width="16%">
								<input id="txtCCAIndicator" type="text" maxlength="1" style="WIDTH: 75px" class="msgTxt">
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<div id=lblPenaltyPlanDesc class="msgLabel">
						<table class="msgLabel" width="100%" cellpadding="0" cellspacing="0">
							<tr>
								<td width="30%">Early Repayment Charge<br>Description</td>
								<td style="padding-left: 2px">
									<textarea class="msgTxt" id="txtPenaltyPlanDesc" rows="3" style="WIDTH: 391px"></textarea>
								</td>
							</tr>
						</table>
					</div>
				</td>
				
			</tr>
			<tr>
				<td colspan="2">
					<table class="msgLabel" width="100%" cellpadding="0" cellspacing="0">
						<tr>
							<td width="30%">Loan Class Type</td>
							<td width="28%" style="padding-left: 2px">
								<input id="txtLoanClass" type="text" maxlength="5" style="WIDTH: 75px" class="msgTxt">
							</td>
							<td width="26%">
								Early Repayment Charge Code
							</td>
							<td width="16%">	
								<input id="txtPenaltyPlanCode" type="text" maxlength="3" style="WIDTH: 75px" class="msgTxt">
							</td>
						</tr>
					</table>
				
				</td>
			</tr>
			
			<tr>
				<td colspan="2">
					<table class="msgLabel" width="100%" cellpadding="0" cellspacing="0">
						<tr>
							<td width="30%">Early Repayment Charge End Date</td>
							<td width="28%" style="padding-left: 2px">
								<input id="txtPenaltyPlanEndDate" type="text" maxlength="10" style="WIDTH: 75px" class="msgTxt">
							</td>
							<td width="26%">Overpayments
							</td>
							<td width="16%">	
								<input id="txtOverpayments" type="text" maxlength="9" style="WIDTH: 75px" class="msgTxt">
							</td>
						</tr>
					</table>
				
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<table class="msgLabel" width="100%" cellpadding="0" cellspacing="0">
						<tr>
							<td width="30%">Loan Type</td>
							<td width="28%" style="padding-left: 2px">
								<input id="txtLoanType" type="text" maxlength="5" style="WIDTH: 75px" class ="msgTxt">
							</td>
							<td width="26%">
								Outstanding Retentions
							</td>
							<td width="16%">	
								<input id="txtExistingRetentions" type="text" maxlength="9" style="WIDTH: 75px" class="msgTxt">
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<table class="msgLabel" width="100%" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td width="30%">Early Repayment Charge Amount</td>
							<td width="28%" style="padding-left: 2px">
								<input id="txtRedemptionPenalty" type="text" maxlength="9" style="WIDTH: 75px" class="msgTxt">
							</td>
							<td width="26%">
								Loan End Date
							</td>
							<td width="16%">
								<input id="txtLoanEndDate" type="text" maxlength="10" style="WIDTH: 75px" class="msgTxt">
							</td>
						</tr>
					</table>
				
				</td>
			</tr>
		</table>
	</div>
</form><!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 12px; POSITION: absolute; TOP: 825px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!--  #include FILE="attribs/dc074attribs.asp" --><!--  #include FILE="Customise/DC074Customise.asp" -->

<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = null;
var m_sAccountGUID = null;
var m_sMortgageLoanGUID	= null;
var m_sInitialRedemptionStatus = null;
var m_sCustomerNumber = null;
var m_sCustomerVersionNumber = null;
var scScreenFunctions;
var m_sAppType = null;
var m_sDirectoryGUID = null;
var m_bIsSubmit = false;
var m_sReadOnly = null;
var m_blnReadOnly = false;
var m_sOtherSystemAccountNumber = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	RetrieveContextData();
	// AT SYS4359 for client customisation
	Customise();

	<% /* Make the required buttons available on the bottom of the screen
	(see msgButtons.asp for details) */ %>
	<% /*EP669 Remove the Another button"  
	var sButtonList = new Array("Submit","Cancel","Another");
     
	
	if(m_sMetaAction != "Add")*/ %>
	
	sButtonList = new Array("Submit","Cancel");
	

	ShowMainButtons(sButtonList);
	<% /* BMIDS00406 remove mortgage
	//FW030SetTitles("Add/Edit Mortgage Loan","DC074",scScreenFunctions);
	*/ %>
	FW030SetTitles("Add/Edit Loan","DC074",scScreenFunctions);
	<% /*BMIDS00406 End */ %>

	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtOutstandingTermYears");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtOutstandingTermMonths");

	GetComboLists();
	PopulateScreen();

	frmScreen.cboRepaymentType.onchange();
	frmScreen.cboRedemptionStatus.onchange();
	CalculateOutstandingTerm();
			
	//This function is contained in the field attributes file (remove if not required)
	SetMasks();

	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");

	if (m_sReadOnly=="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		<% /*EP669 Remove the Another button"  
		DisableMainButton("Another");*/ %>
	}
	
	Validation_Init();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC074");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

<% /* keep the focus within this screen when using the tab key */%>

function spnToFirstField.onfocus ()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<% /* keep the focus within this screen when using the tab key */%>
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

<% /* Retrieves all context data required for use within this screen. */%>
function RetrieveContextData()
{
	m_sMetaAction	= scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sAppType		= scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue",null);
	m_sOtherSystemAccountNumber =  scScreenFunctions.GetContextParameter(window,"idOtherSystemAccountNumber",null);
}

<% /* Populates all combos with their options */%>
function GetComboLists()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	<% /* PSC 15/01/2007 EP2_741 */ %>
	var sGroups = new Array("PurposeOfLoan","RepaymentType","RedemptionStatus", "ApplicationIncomeStatus");
			
	if(XML.GetComboLists(document, sGroups) == true)
	{
		XML.PopulateCombo(document, frmScreen.cboPurposeOfLoan, "PurposeOfLoan"   ,true);
		XML.PopulateCombo(document, frmScreen.cboRepaymentType, "RepaymentType"   ,true);
		XML.PopulateCombo(document, frmScreen.cboRedemptionStatus, "RedemptionStatus",true);
		<% /* PSC 15/01/2007 EP2_741 */ %>
		XML.PopulateCombo(document, frmScreen.cboOriginalIncomeStatus, "ApplicationIncomeStatus",true);
		<% /* If type of application is not New Mortgage, remove 'To be ported' from the list */ %>
		<% // GD BM0356 %>
		<%//if(m_sAppType != "10") %><% /* New Loan */ %>
		if (!(XML.IsInComboValidationList(document,"TypeOfMortgage",m_sAppType,["N"])))
		
		{
			for(var nIndex=0; nIndex < frmScreen.cboRedemptionStatus.length; nIndex++)
			{
				<%//GD BM0356 %>
				if(scScreenFunctions.IsOptionValidationType(frmScreen.cboRedemptionStatus,nIndex,"P"))
					frmScreen.cboRedemptionStatus.remove(nIndex);	
					
					
			}
		}
	}
}

<% /* Checks that the indemnity date isn't in the future */ %>
function CheckRedemptionDate()
{
	var bOK = true;

	if(scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtRedemptionDate,">") == true)
	{
		alert("Redemption Date cannot be in the future");
		frmScreen.txtRedemptionDate.focus();
		bOK = false;
	}
	
	if (frmScreen.txtStartDate.value!="" && frmScreen.txtRedemptionDate.value!="")
		if (scScreenFunctions.GetDateObject(frmScreen.txtRedemptionDate)<
						scScreenFunctions.GetDateObject(frmScreen.txtStartDate))
		{
			alert("Redemption Date cannot be earlier than Loan Start Date");
			frmScreen.txtRedemptionDate.focus();
			return false;
		}	
	return bOK;
}


<% /* Retrieves the data and sets the screen accordingly */ %>
function PopulateScreen()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	XML.LoadXML(sXML);
			
	if(XML.SelectTag(null,"MORTGAGEACCOUNT") != null)
	{
		m_sAccountGUID = XML.GetTagText("ACCOUNTGUID");
		m_sCustomerNumber = XML.GetTagText("CUSTOMERNUMBER");
		m_sCustomerVersionNumber = XML.GetTagText("CUSTOMERVERSIONNUMBER");
		m_sDirectoryGUID = XML.GetTagText("DIRECTORYGUID");
	}

	if(m_sMetaAction == "Edit")
	{
		sXML = scScreenFunctions.GetContextParameter(window,"idXML2",null);
				
		XML.LoadXML(sXML);
		if(XML.SelectTag(null,"MORTGAGELOAN") != null)
		{
			frmScreen.txtLoanAccountNumber.value = XML.GetTagText("LOANACCOUNTNUMBER");
			frmScreen.txtMortgageProductDescription.value = XML.GetTagText("MORTGAGEPRODUCTDESCRIPTION");
			<% /* MAR46 Add Product Code and Expiry Date */ %>
			frmScreen.txtMortgageProductCode.value = XML.GetTagText("MORTGAGEPRODUCTCODE");
			frmScreen.txtCurrentRateExpiryDate.value = XML.GetTagText("CURRENTRATEEXPIRYDATE");
			
			frmScreen.cboPurposeOfLoan.value = XML.GetTagText("PURPOSEOFLOAN");
			frmScreen.cboRepaymentType.value = XML.GetTagText("REPAYMENTTYPE");
			frmScreen.txtOriginalPartAndPartIntOnlyAmount.value	= XML.GetTagText("ORIGINALPARTANDPARTINTONLYAMT");
			frmScreen.txtOriginalLoanAmount.value = XML.GetTagText("ORIGINALLOANAMOUNT");
			frmScreen.txtInterestRate.value = XML.GetTagText("INTERESTRATE");
			frmScreen.txtOutstandingBalance.value = XML.GetTagText("OUTSTANDINGBALANCE");
			frmScreen.txtMonthlyRepayment.value = XML.GetTagText("MONTHLYREPAYMENT");
			frmScreen.txtStartDate.value = XML.GetTagText("STARTDATE");
			frmScreen.txtOriginalTermYears.value = XML.GetTagText("ORIGINALTERMYEARS");
			frmScreen.txtOriginalTermMonths.value = XML.GetTagText("ORIGINALTERMMONTHS");
			frmScreen.cboRedemptionStatus.value = XML.GetTagText("REDEMPTIONSTATUS");
			// Assign value to initial redemption status
			m_sInitialRedemptionStatus = XML.GetTagText("REDEMPTIONSTATUS");		
			
			// EP2_55
			frmScreen.txtOriginalLTV.value = XML.GetTagText("ORIGINALLTV");
			
			<% /* PSC 15/01/2007 EP2_741 - Start */ %>
			scScreenFunctions.SetRadioGroupValue(frmScreen, "FlexibleProductIndicator", XML.GetTagText("FLEXIBLEPRODUCTINDICATOR"));
			frmScreen.cboOriginalIncomeStatus.value = XML.GetTagText("ORIGINALINCOMESTATUS");
			<% /* PSC 15/01/2007 EP2_741 - End */ %>

			var strAccountNumber = XML.GetTagText("LOANACCOUNTNUMBER");
			var nCharIndex = strAccountNumber.indexOf('/');
			strAccountNumber  = parseFloat(strAccountNumber.substring(0,nCharIndex));
			if ( m_sOtherSystemAccountNumber ==  strAccountNumber)
			{
				<%//GD BM0356 %>
				<%//if ( m_sAppType ==  '30' ||  m_sAppType ==  '50')%>
				var combXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				if ((combXML.IsInComboValidationList(document,"TypeOfMortgage",m_sAppType,["F"])) && (combXML.IsInComboValidationList(document,"TypeOfMortgage",m_sAppType,["M"])) )
				{
					scScreenFunctions.SetFieldState(frmScreen, "cboRedemptionStatus", "R");
				}
			}

			frmScreen.txtRedemptionDate.value = XML.GetTagText("REDEMPTIONDATE");
			
			<% /* SR 10/08/2004 : BMIDS815 */ %>
			<% /* MAR46 Remove Remaining Duration and Current Step 
			frmScreen.txtCurrentInterestRateStep.value = XML.GetTagText("PRODUCTSTEP"); 
			frmScreen.txtRemainingStepDuration.value = XML.GetTagText("REMAININGSTEPDURATION");  */ %>
			
			<% /* SR 10/08/2004 : BMIDS815 - End */ %>
			<% /* SR 01/09/2004 : BMIDS815 */ %>
			frmScreen.txtRemainingIntOnlyBalance.value = XML.GetTagText("REMAININGINTERESTONLYAMOUNT");
			frmScreen.txtRemainingCapAndIntAmount.value = XML.GetTagText("REMAININGCAPITALINTERESTAMOUNT");
			<% /* SR 01/09/2004 : BMIDS815 - End */ %>
		
			<% /* PSC 09/08/2002 BMIDS00006 - Start */ %>
			frmScreen.txtAdminOutstandingTerm.value = XML.GetTagText("ICBSCALCULATEDOUTSTANDINGTERM");
			frmScreen.txtDisbursedAmount.value = XML.GetTagText("DISBURSEDAMOUNT");
			frmScreen.txtCCIIndicator.value = XML.GetTagText("CCIINDICATOR");
			frmScreen.txtCCAIndicator.value = XML.GetTagText("CCAINDICATOR");
			frmScreen.txtPenaltyPlanDesc.value = XML.GetTagText("PENALTYPLANDESCRIPTION");
			frmScreen.txtLoanClass.value = XML.GetTagText("LOANCLASSTYPE");
			frmScreen.txtPenaltyPlanCode.value = XML.GetTagText("PENALTYPLANCODE");
			frmScreen.txtPenaltyPlanEndDate.value = XML.GetTagText("PENALTYPLANENDDATE");
			frmScreen.txtOverpayments.value = XML.GetTagText("OVERPAYMENTS");
			frmScreen.txtLoanType.value = XML.GetTagText("LOANTYPE");
			frmScreen.txtExistingRetentions.value = XML.GetTagText("EXISTINGRETENTIONS");
			frmScreen.txtAvailableForDisbursement.value = XML.GetTagText("AVAILABLEFORDISBURSEMENT");
			frmScreen.txtRedemptionPenalty.value = XML.GetTagText("REDEMPTIONPENALTY");
			<% /* PSC 09/08/2002 BMIDS00006 - End */ %>
			<% /* BMIDS00336 MDC 30/08/2002 */ %>
			frmScreen.txtLoanEndDate.value = XML.GetTagText("LOANENDDATE");
			<% /* BMIDS00336 MDC 30/08/2002 */ %>
			
			<%/*START:- AQR: BMIDS756*/%>
			//use try catch block
			try
			{
			frmScreen.txtProductStartDate.value = XML.GetTagText("PRODUCTSTARTDATE");
			frmScreen.txtIndexCode.value = XML.GetTagText("INDEXCODE");
			frmScreen.txtVariance.value = XML.GetTagText("VARIANCE");
			}
			catch(e)
			{
			}
			<%/*END:- AQR: BMIDS756*/%>
				
					
//			scScreenFunctions.SetRadioGroupValue(frmScreen, "ToBeRepaidIndicator", XML.GetTagText("TOBEREPAIDINDICATOR"));

			m_sMortgageLoanGUID = XML.GetTagText("MORTGAGELOANGUID");
		}
	}
}

function frmScreen.cboRepaymentType.onchange()
{
	<%/*[MC]*=Part and Part label visible/invisible code*/%>
	if(scScreenFunctions.IsValidationType(frmScreen.cboRepaymentType, "P"))
	{
		<% /* BMIDS756 GHun
		scScreenFunctions.ShowCollection(spnOriginalPartAndPartIntOnlyAmount);
		*/ %>
		spnOriginalPartAndPartIntOnlyAmount.style.visibility="";
		tlbOriginalPartAndPartIntOnlyAmount.style.visibility="";
	}
	else
	{
		<% /* BMIDS756 GHun Don't call HideCollection as it clears the fields
		scScreenFunctions.HideCollection(spnOriginalPartAndPartIntOnlyAmount);
		*/ %>
		spnOriginalPartAndPartIntOnlyAmount.style.visibility="hidden";
		tlbOriginalPartAndPartIntOnlyAmount.style.visibility="hidden";
	}
}

function frmScreen.cboRedemptionStatus.onchange()
{
	if(scScreenFunctions.IsValidationType(frmScreen.cboRedemptionStatus, "A"))
		scScreenFunctions.ShowCollection(spnRedemptionDate);
	else
		scScreenFunctions.HideCollection(spnRedemptionDate);
}
	
<% /* MAR30 Process income calculations */ %>
function RunIncomeCalculations()
{
	var AllowableIncXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AllowableIncXML.CreateRequestTag(window,null);
	<% /* MAR30 */ %> 
	AllowableIncXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));
		
	var sAppFactFindNo = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	var sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
		
	// Set up request for Income Calculation
	var TagRequest = AllowableIncXML.CreateActiveTag("INCOMECALCULATION");
	AllowableIncXML.SetAttribute("CALCULATEMAXBORROWING","1");
	AllowableIncXML.SetAttribute("QUICKQUOTEMAXBORROWING","0");
	AllowableIncXML.CreateActiveTag("APPLICATION");
	AllowableIncXML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	AllowableIncXML.CreateTag("APPLICATIONFACTFINDNUMBER", sAppFactFindNo);
	AllowableIncXML.ActiveTag = TagRequest;
	
	var TagCUSTOMERLIST = AllowableIncXML.CreateActiveTag("CUSTOMERLIST");

	for(var nLoop = 1; nLoop <= 2; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		<% /* BM0457 The customer must exist */ %>
		if (sCustomerNumber.trim().length > 0)
		{
		<% /* BM0457 End */ %>
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
	
	<% /* PSC 16/05/2006 MAR1798 - Start */ %>
	AllowableIncXML.SelectTag(null,"REQUEST");
	AllowableIncXML.SetAttribute("OPERATION","CriticalDataCheck");	
	AllowableIncXML.CreateActiveTag("CRITICALDATACONTEXT");
	AllowableIncXML.SetAttribute("APPLICATIONNUMBER",sApplicationNumber);
	AllowableIncXML.SetAttribute("APPLICATIONFACTFINDNUMBER",sAppFactFindNo);
	AllowableIncXML.SetAttribute("SOURCEAPPLICATION","Omiga");
	AllowableIncXML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
	AllowableIncXML.SetAttribute("ACTIVITYINSTANCE","1");
	AllowableIncXML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
	AllowableIncXML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
	AllowableIncXML.SetAttribute("COMPONENT","omIC.IncomeCalcsBO");
	AllowableIncXML.SetAttribute("METHOD","RunIncomeCalculation");	

	window.status = "Critical Data Check - please wait";

	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
					AllowableIncXML.RunASP(document,"OmigaTMBO.asp");
			break;
		default: // Error
			AllowableIncXML.SetErrorResponse();
	}

	window.status = "";	
	<% /* PSC 16/05/2006 MAR1798 - End */ %>
	
	AllowableIncXML.IsResponseOK();	
	return(true);	
}	
		
function btnSubmit.onclick()
{
	if(m_bIsSubmit)
		return;
		
	m_bIsSubmit = true;
	if (frmScreen.onsubmit() == true)
		if (CheckRedemptionDate() == true){
			<% /* MAR30 If changed then call income calcs */ %>			
			if(CommitScreen() == true){
				var bContinue = true;
				if(IsChanged())
				{
					bContinue = RunIncomeCalculations();							
				}
				if(bContinue)
					frmToDC073.submit();
			}
			else
				m_bIsSubmit = false;
		}			
		else
			m_bIsSubmit = false;
	else
		m_bIsSubmit = false;
}
		
function btnCancel.onclick()
{
	frmToDC073.submit();
}

<% /* Saves the record and clears all fields for new input */%>
<% /* BMIDS756 GHun btnAnother does not exist 
function btnAnother.onclick()
{
	if(frmScreen.onsubmit() == true)
	{
		if (CheckRedemptionDate() == true)
		{
			if(CommitScreen() == true)
			{
				scScreenFunctions.ClearCollection(frmScreen);
				scScreenFunctions.HideCollection(spnOriginalPartAndPartIntOnlyAmount);
				scScreenFunctions.HideCollection(spnRedemptionDate);
//				frmScreen.optToBeRepaidNo.checked = true;
				scScreenFunctions.SetFocusToFirstField(frmScreen);
			}
		}
	}
}
*/ %>

<% /* Commits the screen data to the database, either by a create or update*/%>
function CommitScreen()
{
	var bOK = false;
	var TagRequestType = null;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	if(m_sMetaAction == "Add")
		TagRequestType = XML.CreateRequestTag(window, "CREATE");
	if(m_sMetaAction == "Edit")
		TagRequestType = XML.CreateRequestTag(window, "UPDATE");

	if(TagRequestType != null)
	{				
		XML.CreateActiveTag("MORTGAGELOAN");				
				
		<% /* For an update we need to specify the Mortgage Loan GUID */ %>
		if(m_sMetaAction == "Edit")
			XML.CreateTag("MORTGAGELOANGUID",m_sMortgageLoanGUID);
		
		XML.CreateTag("ACCOUNTGUID",m_sAccountGUID);
		XML.CreateTag("LOANACCOUNTNUMBER",frmScreen.txtLoanAccountNumber.value);
		XML.CreateTag("MORTGAGEPRODUCTDESCRIPTION",frmScreen.txtMortgageProductDescription.value);
		<% /* MAR46 Add Product Code and Expiry Date */ %>
		XML.CreateTag("MORTGAGEPRODUCTCODE",frmScreen.txtMortgageProductCode.value);
		XML.CreateTag("CURRENTRATEEXPIRYDATE",frmScreen.txtCurrentRateExpiryDate.value);
		
		XML.CreateTag("PURPOSEOFLOAN",frmScreen.cboPurposeOfLoan.value);
		XML.CreateTag("REPAYMENTTYPE",frmScreen.cboRepaymentType.value);
		XML.CreateTag("ORIGINALPARTANDPARTINTONLYAMT",frmScreen.txtOriginalPartAndPartIntOnlyAmount.value);
		XML.CreateTag("ORIGINALLOANAMOUNT",frmScreen.txtOriginalLoanAmount.value);
		XML.CreateTag("INTERESTRATE",frmScreen.txtInterestRate.value);
		XML.CreateTag("OUTSTANDINGBALANCE",frmScreen.txtOutstandingBalance.value);
		XML.CreateTag("MONTHLYREPAYMENT",frmScreen.txtMonthlyRepayment.value);
		XML.CreateTag("STARTDATE",frmScreen.txtStartDate.value);
		XML.CreateTag("ORIGINALTERMYEARS",frmScreen.txtOriginalTermYears.value);
		XML.CreateTag("ORIGINALTERMMONTHS",frmScreen.txtOriginalTermMonths.value);
		XML.CreateTag("REDEMPTIONSTATUS",frmScreen.cboRedemptionStatus.value);
		XML.CreateTag("INITIALREDEMPTIONSTATUS", m_sInitialRedemptionStatus) ;
		XML.CreateTag("REDEMPTIONDATE",frmScreen.txtRedemptionDate.value);
		XML.CreateTag("CUSTOMERNUMBER",m_sCustomerNumber);
		XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber) ;
//		XML.CreateTag("TOBEREPAIDINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"ToBeRepaidIndicator"));
		XML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID);
		
		<% /* PSC 09/08/2002 BMIDS00006 - Start */ %>
		XML.CreateTag("ICBSCALCULATEDOUTSTANDINGTERM", frmScreen.txtAdminOutstandingTerm.value);
		XML.CreateTag("DISBURSEDAMOUNT", frmScreen.txtDisbursedAmount.value);
		XML.CreateTag("CCIINDICATOR", frmScreen.txtCCIIndicator.value);
		XML.CreateTag("CCAINDICATOR", frmScreen.txtCCAIndicator.value);
		XML.CreateTag("PENALTYPLANDESCRIPTION", frmScreen.txtPenaltyPlanDesc.value);
		XML.CreateTag("LOANCLASSTYPE", frmScreen.txtLoanClass.value);
		XML.CreateTag("PENALTYPLANCODE", frmScreen.txtPenaltyPlanCode.value);
		XML.CreateTag("PENALTYPLANENDDATE", frmScreen.txtPenaltyPlanEndDate.value);
		XML.CreateTag("OVERPAYMENTS", frmScreen.txtOverpayments.value);
		XML.CreateTag("LOANTYPE", frmScreen.txtLoanType.value);
		XML.CreateTag("EXISTINGRETENTIONS", frmScreen.txtExistingRetentions.value);
		XML.CreateTag("AVAILABLEFORDISBURSEMENT", frmScreen.txtAvailableForDisbursement.value);
		XML.CreateTag("REDEMPTIONPENALTY", frmScreen.txtRedemptionPenalty.value);
		
		<%/*START:- AQR: BMIDS756*/%>
		XML.CreateTag("PRODUCTSTARTDATE", frmScreen.txtProductStartDate.value);
		XML.CreateTag("INDEXCODE", frmScreen.txtIndexCode.value);
		XML.CreateTag("VARIANCE", frmScreen.txtVariance.value);
		<%/*END:- AQR: BMIDS756*/%>
		<% /* PSC 09/08/2002 BMIDS00006 - End */ %>

		<% /* SR 10/08/2004 : BMIDS815 */ %>
		
		<% /* MAR46 Remove Remaining Duration and Current Step 
		XML.CreateTag("PRODUCTSTEP", frmScreen.txtCurrentInterestRateStep.value);
		XML.CreateTag("REMAININGSTEPDURATION", frmScreen.txtRemainingStepDuration.value);  */ %>
		XML.CreateTag("PRODUCTSTEP", "");
		XML.CreateTag("REMAININGSTEPDURATION", "");  
		
		<% /* SR 10/08/2004 : BMIDS815 - End */ %>
		<% /* SR 01/09/2004 : BMIDS815 */ %>
		XML.CreateTag("REMAININGINTERESTONLYAMOUNT", frmScreen.txtRemainingIntOnlyBalance.value);
		XML.CreateTag("REMAININGCAPITALINTERESTAMOUNT", frmScreen.txtRemainingCapAndIntAmount.value);
		<% /* SR 01/09/2004 : BMIDS815 - End */ %>
		
		<% /* BMIDS00336 MDC 30/08/2002 */ %>
		XML.CreateTag("LOANENDDATE", frmScreen.txtLoanEndDate.value);
		<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
		
		// EP2_55
		XML.CreateTag("ORIGINALLTV", frmScreen.txtOriginalLTV.value);
		
		<% /* PSC 15/01/2007 EP2_741 - Start */ %>
		XML.CreateTag("ORIGINALINCOMESTATUS", frmScreen.cboOriginalIncomeStatus.value);
		XML.CreateTag("FLEXIBLEPRODUCTINDICATOR", scScreenFunctions.GetRadioGroupValue(frmScreen,"FlexibleProductIndicator"));
		<% /* PSC 15/01/2007 EP2_741 - End */ %>

		if(m_sMetaAction == "Add")
			// 			XML.RunASP(document, "CreateMortgageLoan.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "CreateMortgageLoan.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

		else
			// 			XML.RunASP(document, "UpdateMortgageLoan.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "UpdateMortgageLoan.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}


		bOK = XML.IsResponseOK();
	}
	else
		alert("CommitScreen - Invalid MetaAction");
	return bOK;
}

function CalculateOutstandingTerm()
{
	var sOutstandingYears = "";
	var sOutstandingMonths = "";
		
	var dtStartDate = scScreenFunctions.GetDateObject(frmScreen.txtStartDate);
	var sOriginalTermYears = frmScreen.txtOriginalTermYears.value;
	var sOriginalTermMonths = frmScreen.txtOriginalTermMonths.value;
		
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
					sOutstandingYears = nYearsToGo;
					sOutstandingMonths = nMonthsToGo;
				}
			}
		}
	}
	frmScreen.txtOutstandingTermYears.value = sOutstandingYears;
	frmScreen.txtOutstandingTermMonths.value = sOutstandingMonths;
}
	
function frmScreen.txtStartDate.onchange()
{
	CalculateOutstandingTerm();
}
	
function frmScreen.txtOriginalTermYears.onchange()
{
	CalculateOutstandingTerm();
}

function frmScreen.txtOriginalTermMonths.onchange()
{
	CalculateOutstandingTerm();
}

function frmScreen.txtInterestRate.onchange()
{
	<%/* Pad interest rate to two decimal places */%>
	var sNumber = frmScreen.txtInterestRate.value;

	if (sNumber != "")
	{
		sNumber = top.frames[1].document.all.scMathFunctions.RoundValue(sNumber,2);
		frmScreen.txtInterestRate.value = top.frames[1].document.all.scMathFunctions.PadToPrecision(sNumber,2);
	}
}

function frmScreen.txtPenaltyPlanDesc.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtPenaltyPlanDesc", 255, true);
}

-->
</script>
</body>
</html>
