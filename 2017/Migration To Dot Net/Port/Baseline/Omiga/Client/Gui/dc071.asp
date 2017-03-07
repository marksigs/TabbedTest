<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      dc071.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Add/Edit Mortgage Account
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		02/02/2000	ie5 changes 
AY		14/02/00	Change to msgButtons button types
AD		17/03/2000	Incorporated third party include files.
IW		22/03/2000  SYS0314 - Address Combo 
AY		06/04/00	New top menu/scScreenFunctions change
IVW		12/04/2000	Fixed Combo to include guarantors
IVW		03/05/2000  ------- - Added all addresses into combo/saved/returned that info.
IVW		03/05/2000  ------- - Added 'Not Applicable into address combo'.
IVW		04/05/2000	SYS0486 - Set applicant combo toggle to be correct value
MH      08/05/2000  SYS0672 Routing of OK button
MC		12/05/2000	SYS0688 Check Lender in Directory for Ported Loans
MC		15/05/2000	SYS0694 Remove tooltips and amend maxlength of Additional Details
							and Indemnity Mortgage Amount
MH		15/05/2000  SYS0738 Saving with defaults
BG		17/05/00	SYS0752 Removed Tooltips
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
MC		31/05/00	Added check on length of text in Additional Details textarea
MC		05/06/00	SYS0694 Remove <select> from Country combo and default to UK
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
MC		08/06/00	SYS0866 Standardise behaviour of Balance & Payment fields
MH      22/06/00    SYS0933 Readonly processing
BG		15/08/00	SYS1428 Added code to ensure Lender name is completed if any lender details
							are completed.
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
BG		04/11/00	SYS1037 Remove validation for "0" in ValidateThirdPartyDetails.
CL		05/03/01	SYS1920 Read only functionality added
JR		14/09/01	Omiplus24 - Removed FaxNumber, TelephoneNumber and EmailAddress. Now included
					in ThirdpartyDetails include.
MDC		01/10/01	SYS2785 Enable client versions to override labels
JLD		22/01/02	SYS3851 Get the contactdetails based on the thirdparty or directory guid.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
PSC		30/07/2002	BMIDS00006  Amend to use new Mortgage Account structure
MV 		08/08/2002 	BMIDS00302 	Core Ref : SYS4728 remove non-style sheet styles
PSC		13/08/2002  BMIDS00006	Set address combo from MAADDRESSGUID rather than ADDRESSGUID
PSC		21/08/2002	BMIDS00350	Cater for Property Details not being present if user has not
                                gone into DC072
MDC		29/08/2002	BMIDS00336	Credit Check & Full Bureau Download
GHun	05/09/2002	BMIDS00406	Remove the word "mortgage" from the screen
MDC		30/09/2002	BMIDS00510	Security Address for downloaded accounts
TW      09/10/2002  SYS5115     Modified to incorporate client validation - 
CL		14/10/2002  BMIDS00439	Change to function btnLoans.onclick
MDC		24/10/2002	BMIDS00682	Enable Property Details drilldown button even when Property 
								Address combo is read only.
MV		31/10/2002	BMIDS00355	Modified HTML
MO		06/11/2002	BMIDS00847  Modified after errors were received at BM, when using the method SetProperty of the DOM
GHun	11/11/2002	BMIDS00444	Added Remortgage account indicator
GD		17/11/2002	BMIDS00376 Disable screen if readonly
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
MV		25/03/2003	BM00063		AMended HTML Text for Option buttons 
HMA		16/09/2003	BM00063		Amended HTML Text for Option buttons 
KRW     25/09/03    BM00063     Correction to screen alignment
DRC     23/12/2003  BMIDS667    Added Repossession Radio buttons
KRW     29/03/2004  BMIDS733    Set unassigned to Yes on screen entry
MC		19/04/04	CC057	Popup Dialogs height increased by 10px
MC		17/05/2004	BMIDS756	New Fields (collateralID,Collateral Balance, etc.,) Added
KW      26/05/2004  BMIDS773    Regulation Amendments to text fields
SR		14/06/2004	BMIDS772	Update FinancialSummary record on Submit (only for create)	
MC		09/07/2004	BMIDS772	Srini's Financial Summary code block moved.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
GHun	22/07/2005	MAR10		Made client specific text more generic (Own Account)
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
MF		02/08/2005	MAR20		WP06 Add new fields and modify field names
MF		15/09/2005	MAR30		IncomeCalcs modified to send ActivityID into calculation
HM		03/11/2005	MAR343		Payment due date max length cnanged from 2 to 10
MV		04/10/2005  MAR252		Amended PaymentDueDate
PE		25/01/2006	MAR1109		DC070 - Changed to display the outstanding balance (Collateral Balance) from Mortgage Account
PE		07/02/2006	MAR1189		Make various fields mandatory if mortgage type is remortgage
JD		08/02/2006	MAR1040		If there are mortgage loans for this account, disable the monthly cost and redemption status fields
								as they will be taken from the loan for income calcs.
PE		08/02/2006	MAR1189		Make various fields mandatory if mortgage type is remortgage (bugfix)	
MH&GH	21/04/2006	MAR1448		If indemnity insurance not present create a working record else update what's there
PSC		16/05/2006	MAR1798		Run Critical Data 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
PB		21/07/2006	EP543		Added 'TITLEOTHER' field for third parties
AW		24/07/2006	EP1023		Amended xml doc reference for ContactOtherTitle
MAH		21/11/2006	E2Cr35		Added OTHERMORTGAGEPROPERTY nb Pre-emption EndDate added to get spacing, but set to readonly
AShaw	24/11/2006	E2CR18		Add Pre-emption end date for Right to Buy.
PSC		15/01/2007	EP2_741		Add Bank Name, Nature Of Loan, and Credit Scheme
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp 
viewastext></OBJECT>
<script src="validation.js" type="text/javascript"></script>

<form id="frmToDC070" method="post" action="dc070.asp" style="DISPLAY: none"></form>
<form id="frmToDC073" method="post" action="dc073.asp" style="DISPLAY: none"></form>
<form id="frmToDC130" method="post" action="dc130.asp" style="DISPLAY: none"></form>

<% /* PSC 05/08/2002 BMIDS00006 */ %>
<form id="frmToDC075" method="post" action="dc075.asp" style="DISPLAY: none"></form>

<span id="spnListScroll">
	<span style="LEFT: 301px; POSITION: absolute; TOP: 128px">
<OBJECT id=scScrollTable style="WIDTH: 304px; HEIGHT: 24px" tabIndex=-1 
type=text/x-scriptlet data=scTableListScroll.asp 
viewastext></OBJECT>
	</span> 
</span>

<% /* Span field to keep tabbing within this screen */%>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange" year4>
    <div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 533px" class="msgGroup">
		<div id="lblOwners" style="LEFT: 10px; POSITION: absolute; TOP: 8px" class="msgLabel">
			<% /* BMIDS00406 remove mortgage
			// Owner(s) of<br>Mortgage<br>Account
			*/ %>
			Owner(s) of<br>Account
			<% /* BMIDS406 End */ %>
			<div id="spnOwners" style="LEFT: 85px; POSITION: absolute; TOP: -3px">
				<table id="tblOwners" width="500" border="0" cellspacing="0" cellpadding="0" class="msgTable">
				<tr id="rowTitles">
					<td width="70%" class="TableHead">Name</td>	
					<td width="30%" class="TableHead">Role</td>
				</tr>
				<tr id="row01">		
					<td class="TableTopLeft">&nbsp;</td>		
					<td class="TableTopRight">&nbsp;</td>
				</tr>
				<tr id="row02">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row03">		
					<td class="TableBottomLeft">&nbsp;</td>		
					<td class="TableBottomRight">&nbsp;</td>
				</tr>
				</table>
			</div>
		</div>
		<span style="LEFT: 95px; POSITION: absolute; TOP: 68px">
			<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="LEFT: 160px; POSITION: absolute; TOP: 68px">
			<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="LEFT: 225px; POSITION: absolute; TOP: 68px">
			<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
		</span><!--MC Change start-->
		<table id="tlbExistingAccount" border="0" cellpadding="1" cellspacing="1" class="msgLabel" style="MARGIN-LEFT: 0px; WIDTH: 593px; POSITION: absolute; TOP: 105px; HEIGHT: 318px">
			<tr>
				<td width="27%"> Collateral ID</td>
				<td width="27%"><input id="txtCallateralID" maxlength="12" style="WIDTH: 130px" class ="msgTxt"></td>
				<td width="27%">Own Account?</td>
				<td width="18%">
					<input id="optOwnAccountYes" name="OwnAccount" type="radio" value="1"><label for="optOwnAccountYes" class="msgLabel">Yes</label>
					<input id="optOwnAccountNo" name="OwnAccount" type="radio" value="0" checked><label for="optOwnAccountNo" class="msgLabel">No</label>
				</td>
			</tr>
			<tr>
				<td>Balance</td>
				<td> 
					<input id="txtCollateralBalance" maxlength="14" class="msgTxt"  >
				</td>
				
				<td colspan="2">
						<table id="spnRMC" width="100%" height="20" class="msgLabel" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td>RMC flag</td>
								<td width="105">
									<input id="optRMCFlagYes" name="RMCFlag" type="radio" value="1" ><label for="optRMCFlagYes" class="msgLabel" >Yes</label>
						            <input id="optRMCFlagNo" name="RMCFlag" type="radio" value="0" checked><label for="optRMCFlagNo" class="msgLabel">No</label>
								</td>
							</tr>
						</table>
				</td>
				
				
			</tr>
			<tr>
				<td>Account Number</td>
				<td> 
					<input id="txtAccountNumber" name="AccountNumber" maxlength="20" class="msgTxt" >
				</td>
				
				<td colspan="2">
					<table id="spnRemortgageIndicator" width="100%" height="20" class="msgLabel" border="0" cellspacing="0" cellpadding="0" color="red">
						<tr>
							<td> Is this account to be <br>remortgaged?</td>
							<td width="105">
								<input id="optRemortgageIndicatorYes" name="RemortgageIndicator" style="MARGIN-LEFT: 0px" type="radio" value="1"><label for="optRemortgageIndicatorYes" class="msgLabel">Yes</label>
								<input id="optRemortgageIndicatorNo" name="RemortgageIndicator" type="radio" value="0"><label for="optRemortgageIndicatorNo" class="msgLabel">No</label>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<% /* MF MARS20 New fields: Redemption status & Total monthly cost */ %>
			<tr>
				<td>Nature Of Loan</td>
				<td  colspan=3>  
					<select id="cboNatureOfLoan" style="WIDTH: 200px; HEIGHT: 20px" class="msgCombo" NAME="cboNatureOfLoan"></select>
				</td>		
			</tr>
			<tr>
				<td>Credit Scheme</td>
				<td  colspan=3>  
					<select id="cboCreditScheme" style="WIDTH: 200px; HEIGHT: 20px" class="msgCombo" NAME="cboCreditScheme"></select>
				</td>		
			</tr>
			<tr>
				<td>Account Redemption Status</td>
				<td  colspan=3>  
					<select id="cboRedemptionStatus" style="WIDTH: 200px; HEIGHT: 20px" class="msgCombo" NAME="cboRedemptionStatus"></select>
				</td>		
			</tr>
			<tr>
				<td>Total Monthly Cost</td>
				<td>  
					<input id="txtTotalMonthly" name="txtTotalMonthly" maxlength="9" class="msgTxt" >
				</td>
				<td colspan="2">
					<table id="spnSecondCharge" width="100%" height="20" class="msgLabel" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td>Is this a Second Charge?</td>
							<td width="105">
								<input id="optSecondChargeYes" name="SecondChargeIndicator" type="radio" value="1" ><label for="optSecondChargeYes" class="msgLabel">Yes</label>
								<input id="optSecondChargeNo" name="SecondChargeIndicator" type="radio" value="0" checked><label for="optSecondChargeNo" class="msgLabel">No</label>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td nowrap>Daily/Monthly/Annual Interest</td>
				<td> 
					<input id="txtInterestRate" maxlength="1" class="msgTxt" >
				</td>
				<td>DSS flag</td>
				<td>
					<input id="optDSSFlagYes" name="DSSFlag" type="radio" value="1" ><label for="optDSSFlagYes" class="msgLabel">Yes</label>
					<input id="optDSSFlagNo" name="DSSFlag" type="radio" value="0" checked><label for="optDSSFlagNo" class="msgLabel">No</label>
				</td>
			</tr>
			<tr>
				<td>Payment Due Day</td>
				<td> 
					<input id="txtPaymentDueDate" class="msgTxt" maxLength=10>
				</td>
				<td>Credit Search</td>
				<td>
					<input id="optCreditCheckYes" name="CreditCheckIndicator" type="radio" value="1"><label for="optCreditCheckYes" class="msgLabel">Yes</label>
					<input id="optCreditCheckNo" name="CreditCheckIndicator" type="radio" value="0" checked><label for="optCreditCheckNo" class="msgLabel">No</label>
				</td>
			</tr>
			<tr>
				<td>Property Address</td>
				<td colspan=2>
					<select id="cboPropertyAddress" style="WIDTH: 315px; HEIGHT: 20px" class="msgCombo" NAME="cboPropertyAddress"></select>					
				</td>
				<td>
					<input id="btnPropertyDetails" type="button" style="WIDTH: 63px; HEIGHT: 26px" class ="msgDDButton" onclick="btnPropertyDetails.onClick()" NAME="btnPropertyDetails">				
				</td>
			</tr>
			<tr>
				<td>Pre-emption End Date</td>
				<td> 
					<input id="txtPreEmptionEndDate" class="msgTxt" maxLength=10 NAME="txtPreEmptionEndDate">
				</td>
				<td>Monthly Rental</td>
				<td>
					<input id="txtRentalIncome" maxlength="20" style="WIDTH: 75px" class ="msgTxt">
				</td>
			</tr>		
			<tr>
				<td>Bank Name</td>
				<td colspan=3> 
					<input id="txtBankName" class="msgTxt" style="WIDTH: 200px; HEIGHT: 18px">
				</td>
			</tr>
			<tr>
				<td>Bank Sort Code</td>
				<td> 
					<input id="txtSortCode" maxlength="8" class="msgTxt" >
				</td>
				<td>Repossession?</td>
				<td>
					<input id="optRepossessionIndicatorYes" name="RepossessionIndicator" type="radio" value="1"><label for="optRepossessionIndicatorYes" class="msgLabel">Yes</label>
					<input id="optRepossessionIndicatorNo" name="RepossessionIndicator" type="radio" value="0" checked><label for="optRepossessionIndicatorNo" class="msgLabel">No</label>
				</td>
			</tr>
			<tr>
				<td>Bank Account Number</td>
				<td> 
					<input id="txtBankAccountNumber" name="BankAccountNumber" maxlength="18"  class="msgTxt">
				</td>				
				<td>Unassigned</td>
				<td>
					<input id="optUnassignedYes" name="UnassignedIndicator" type="radio" value="1" CHECKED size=22><label for="optUnassignedYes" class="msgLabel">Yes</label>
					<input id="optUnassignedNo" name="UnassignedIndicator" type="radio" value="0" ><label for="optUnassignedNo" class="msgLabel">No</label>
				</td>
			</tr>			
			<tr>
				<td>Bank Account Name</td>
				<td colspan=3> 
					<input id="txtBankAccountName" class="msgTxt" style="WIDTH: 398px; HEIGHT: 18px">
				</td>
			</tr>
			<tr>
				<td>Business Channel</td>
				<td colspan=3> 
					<input id="txtBusinessChannel" class="msgTxt" style="WIDTH: 398px; HEIGHT: 40px">
				</td>
			</tr>			
			<tr>
				<td>
					<span id=lblAdditionalDetails class="msgLabel">	
						Additional Details
					</span>
				</td>
				<td colspan="3">
					<TEXTAREA class=msgTxt id=txtAdditionalDetails style="WIDTH: 398px"></TEXTAREA>
				</td>
			</tr>
			<tr>
				<td colspan="3">Do you own/partly own any other property or are you party to any other mortgage?</td>
				<td>
					<input id="optOtherPropertyMortgageYes" name="OtherPropMortInd" type="radio" value="1"><label for="optOtherPropertyMortgageYes" class="msgLabel">Yes</label>
					<input id="optOtherPropertyMortgageNo" name="OtherPropMortInd" type="radio" value="0"><label for="optOtherPropertyMortgageNo" class="msgLabel">No</label>
				</td>
			</tr>
		</table>
	</div><!--Old data replacement end-->
	<div id="Lender" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 637px; HEIGHT: 274px" class="msgGroup">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
			<strong>Lender Details</strong> 
		</span>
		
		<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
			Lender
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<input id="txtCompanyName" maxlength="45" style="WIDTH: 200px; POSITION: absolute" class="msgTxt" onchange="ContactDetailsChanged()">
			</span> 
		</span>

		<span style="LEFT: 330px; POSITION: absolute; TOP: 31px" class="msgLabel">
			<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton" onclick="btnDirectorySearch.onClick()">
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
			Is in the directory?
			<span style="LEFT: 120px; WIDTH: 60px; POSITION: absolute; TOP: -3px">
				<input id="optInDirectoryYes" name="InDirectoryGroup" type="radio" value="1" onclick="InDirectoryChanged()"><label for="optInDirectoryYes" class="msgLabel">Yes</label> 
			</span> 

			<span style="LEFT: 174px; WIDTH: 60px; POSITION: absolute; TOP: -3px">
				<input id="optInDirectoryNo" name="InDirectoryGroup" type="radio" value="0" onclick="InDirectoryChanged()"><label for="optInDirectoryNo" class="msgLabel">No</label> 
			</span> 
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 82px" class="msgLabel">
			Organisation Type
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<select id="cboOrganisationType" style="WIDTH: 200px" class="msgCombo">
				</select>
			</span>
		</span>
	</div> 

	<div style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 730px; HEIGHT: 232px" class="msgGroup"><!-- #include FILE="includes/thirdpartydetails.htm" -->
	</div> 

	<div style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 960px; HEIGHT: 120px" class="msgGroup">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
			<strong>Higher Lending Charge Details</strong>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel">
			Original Higher Lending Charge Company
			<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
				<input id="txtIndemnityCompanyName" maxlength="50" style="LEFT: 93px; WIDTH: 200px; POSITION: absolute; TOP: 0px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 48px" class="msgLabel">
			Original Higher Lending Charge Premium
			<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
				<input id="txtIndemnityAmount" maxlength="9" style="LEFT: 94px; WIDTH: 70px; POSITION: absolute; TOP: 0px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 72px" class="msgLabel">
			Original Higher Lending Charge Mortgage Amount
			<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
				<input id="txtIndemnityMortgageAmount" maxlength="6" style="LEFT: 94px; WIDTH: 70px; POSITION: absolute; TOP: 0px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 96px" class="msgLabel">
			Higher Lending Charge Date Paid
			<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
				<input id="txtIndemnityDatePaid" maxlength="10" style="LEFT: 94px; WIDTH: 70px; POSITION: absolute; TOP: 0px" class="msgTxt">
			</span>
		</span>
	</div>
</form><!-- Main Buttons -->
<div id="msgButtons" style="LEFT: 8px; WIDTH: 604px; POSITION: absolute; TOP: 1087px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% //Span to keep tabbing within this screen %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- #include FILE="attribs/dc071attribs.asp" --><!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = null;
var XMLOnEntry = null;
var m_sAccountGUID = null;
var scScreenFunctions;
var m_sDirectoryGUID = null;
var m_sReadOnly = null;
var XMLSeqNo = null;
var m_blnReadOnly = false;
var m_iTableLength = 3;
var m_XMLRelationships = null;
var m_XMLApplicants = null;
var m_nNoOfCustomers = null;
var m_XMLAddressList = null;
var m_sSelPostCode = "";
var m_sSelHouseNumber = "";
var m_sSelHouseName = "";
var m_XMLPropertyDetails = null;
<% /* BMIDS00510 MDC 30/09/2002 */ %>
var m_bDownloadedOwnAccount = false;
var m_xmlSecurityAddress = null;
<% /* BMIDS00510 MDC 30/09/2002 - End */ %>

<% /* SR 14/06/2004 : BMIDS772 */ %>
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
<% /* SR 14/06/2004 : BMIDS772 - End */ %>

var m_blnAccountNumberMandatory = false;
var m_blnCreateIndeminityInsuranceOnUpdate = false; <% //MAR1706 MH 04/05/2006 %>

function frmScreen.txtAdditionalDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true);
}

function btnCancel.onclick()
{
	frmToDC070.submit();
}

function btnLoans.onclick()
{
	if(CommitChanges())
	{
		<% /* PSC 12/08/2002 BMIDS00006 - Start */%>
		var sName = "";
		m_XMLRelationships.ActiveTag = null;
		m_XMLRelationships.CreateTagList("ACCOUNTRELATIONSHIP");

<%		//BMIDS00439 CL 10/10/02
		//for (var nIndex=1; nIndex <= m_XMLRelationships.ActiveTagList.length; nIndex++)
		//{
		//	if (nIndex > 1)
		//		if (nIndex == m_XMLRelationships.ActiveTagList.length)
		//			sName = sName + " & ";
		//		else
		//			sName = sName + ", ";
		//
		//	sName = sName + tblOwners.rows(nIndex).cells(0).innerHTML		
		//}
%>		
		
		for (var nIndex=0; nIndex < m_XMLRelationships.ActiveTagList.length; nIndex++)
		{
			if (nIndex > 0)
				if (nIndex == m_XMLRelationships.ActiveTagList.length)
					sName = sName + " & ";
				else
					sName = sName + ", ";
					m_XMLRelationships.SelectTagListItem(nIndex)
		    		sName =  sName  +  m_XMLRelationships.ActiveTag.selectSingleNode("FIRSTFORENAME").text + ' ' +
					m_XMLRelationships.ActiveTag.selectSingleNode("SURNAME").text;
		}
		//BMIDS00439 CL 10/10/02 END
		
		
		<% /* PSC 12/08/2002 BMIDS00006 - End */%>
		
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateActiveTag("MORTGAGEACCOUNT");
		
		<% /* PSC 12/08/2002 BMIDS00006 */%>
		XML.CreateTag("CUSTOMERNAMES" ,sName);	
		XML.CreateTag("ORGANISATIONNAME",frmScreen.txtCompanyName.value);
		XML.CreateTag("ACCOUNTNUMBER" ,frmScreen.txtAccountNumber.value);
		XML.CreateTag("ACCOUNTGUID" ,m_sAccountGUID);
		XML.CreateTag("DIRECTORYGUID" ,m_sDirectoryGUID);
						
		scScreenFunctions.SetContextParameter(window,"idXML",XML.XMLDocument.xml);
						
		frmToDC073.submit();
	}
}
function btnSubmit.onclick()
{
//BG SYS1428 15/08/00 ensure lender name is completed if lender details have been included
	if(frmScreen.txtCompanyName.value == "" && IsScreenClear()== false)	
	{
		alert("Please enter a value");
		frmScreen.txtCompanyName.focus();
	}
	else
	{
		if (m_sMetaAction == "Add")
			btnLoans.onclick();
		else if (CommitChanges())
			frmToDC070.submit();
	}	
//BG SYS1428
}

function frmScreen.cboPropertyAddress.onchange()
{

	<% /* PSC 07/08/2002 BMIDS00006 - Start */ %>
	var nSelectedindex = frmScreen.cboPropertyAddress.selectedIndex;
	m_sSelPostCode = frmScreen.cboPropertyAddress(nSelectedindex).getAttribute("PostCode");
	m_sSelHouseNumber = frmScreen.cboPropertyAddress(nSelectedindex).getAttribute("HouseNumber");
	m_sSelHouseName = frmScreen.cboPropertyAddress(nSelectedindex).getAttribute("HouseName");
	<% /* PSC 07/08/2002 BMIDS00006 - End */ %>
	
	if(frmScreen.cboPropertyAddress.value == "")
		scScreenFunctions.DisableDrillDown(frmScreen.btnPropertyDetails);
	else					
		scScreenFunctions.EnableDrillDown(frmScreen.btnPropertyDetails);		
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	RetrieveContextData();
	
	<% /* PSC 30/07/2002 BMIDS00006 - Start */ %>
	m_XMLRelationships = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_XMLRelationships.CreateActiveTag("ACCOUNTRELATIONSHIPLIST");
	<% /* PSC 30/07/2002 BMIDS00006 - End */ %>

	<% /* Make the required buttons available on the bottom of the screen
		(see msgButtons.asp for details) */ %>
	sButtonList = new Array("Submit","Cancel","Loans","Features","Arrears");
	ShowMainButtons(sButtonList);

	// MDC SYS2564 / SYS2785 Client Customisation
	var sGroups = new Array("Country","OrganisationType");
	objDerivedOperations = new DerivedScreen(sGroups);
	
	<%/* Parameters for thirdpartydetails.asp */%>
	m_sThirdPartyType = "3";
	m_ctrOrganisationType = frmScreen.cboOrganisationType;
	m_fValidateScreen = ValidateThirdPartyDetails;
	m_sUnitId = scScreenFunctions.GetContextParameter(window, "idUnitId",null);
	m_sUserId = scScreenFunctions.GetContextParameter(window, "idUserId",null);

	FW030SetTitles("Add/Edit Existing Account","DC071",scScreenFunctions);

	<% /* Default state on screen entry */ %>
	scScreenFunctions.SetFieldToDisabled(frmScreen,"cboPropertyAddress");
	// IVW Enable the drill button to go to DC072
	scScreenFunctions.DisableDrillDown(frmScreen.btnPropertyDetails);

	// MC SYS2564/SYS2785 for client customisation
	ThirdPartyCustomise();

	PopulateCombos();
	// BG 08/08/00 disable the drill button to go to DC072 if not applicable is default
	scScreenFunctions.DisableDrillDown(frmScreen.btnPropertyDetails);
	
	m_nNoOfCustomers = GetCustomerCount();
	
	if (m_sMetaAction == "Edit")
		PopulateScreen();
	else
	{
		<% /* BMIDS00336 MDC 29/08/2002 */ %>
		scScreenFunctions.SetRadioGroupToReadOnly(frmScreen, "UnassignedIndicator");
		scScreenFunctions.SetRadioGroupValue(frmScreen, "UnassignedIndicator", "0"); //BMIDS733 KRW
		<% /* BMIDS00336 MDC 29/08/2002 - End */ %>
		PopulateListBox(0);
	}
	
	SetRemortgageIndicatorVisibility(); <% /* BMIDS00444 */ %>
				
	<% /* PSC 08/08/2002 BMIDS00006 */ %>	
	SetupPropertyDetails();		
			
	<% /* Always mark the record as changed on an add */ %>	
	FlagChange(m_sMetaAction == "Add");
				
	//This function is contained in the field attributes file (remove if not required)
	SetThirdPartyDetailsMasks();
	SetMasks();

	Validation_Init();
	SetAvailableFunctionality();
	<% /* Reset label colour to 616161 as only mandatory if radio button set to yes */ %>
	lblAdditionalDetails.style.color = "#616161";
	//GD BMIDS00376
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC071");

	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		//GD BMIDS00376
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
		if (frmScreen.btnDirectorySearch.style.visibility != 'hidden')
		{
			frmScreen.btnDirectorySearch.disabled = true;
		}
		if (frmScreen.btnClear.style.visibility != 'hidden')
		{
			frmScreen.btnClear.disabled = true;
		}
		if (frmScreen.btnPAFSearch.style.visibility != 'hidden')
		{
			frmScreen.btnPAFSearch.disabled = true;
		}		
		if (frmScreen.btnAddToDirectory.style.visibility != 'hidden')
		{
			frmScreen.btnAddToDirectory.disabled = true;
		}		
	}

	<% /* BMIDS00336 MDC 29/08/2002 */ %>
	scScreenFunctions.SetRadioGroupToReadOnly(frmScreen, "CreditCheckIndicator");
	<% /* BMIDS00336 MDC 29/08/2002 - End */ %>
		
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	//m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC071");
		
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
//	scScreenFunctions.SetFieldToReadOnly(frmScreen,"Text1");  // Removed EP2_18
}

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
			{
				if (ValidateScreen())
				{
					bSuccess = SaveMortgageAccount();
					if(bSuccess)
					{				
						bSuccess = RunIncomeCalculations();
					}
				}
				else
					bSuccess = false
			}
		}
		else
			bSuccess = false;
	return(bSuccess);
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

function PopulateCombos()
// Populates all combos on the screen
{	
	PopulateTPTitleCombo();
	var sControlList = new Array(frmScreen.cboCountry, frmScreen.cboOrganisationType)
	objDerivedOperations.GetComboLists(sControlList);


	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<% /* PSC 15/01/2007 EP2_741 */ %>
	var sGroupList = new Array("Country","OrganisationType","RedemptionStatus", "NatureOfLoan", "SpecialGroup");

	if(XML.GetComboLists(document,sGroupList))
	{
		// Populate Country combo
		blnSuccess = XML.PopulateCombo(document,frmScreen.cboCountry,"Country",false);
		blnSuccess = blnSuccess && XML.PopulateCombo(document,frmScreen.cboOrganisationType,"OrganisationType",true);
		blnSuccess = blnSuccess && XML.PopulateCombo(document,frmScreen.cboRedemptionStatus,"RedemptionStatus",true);
		
		<% /* PSC 15/01/2007 EP2_741 - Start */ %>
		blnSuccess = blnSuccess && XML.PopulateCombo(document,frmScreen.cboNatureOfLoan,"NatureOfLoan",true);
		blnSuccess = blnSuccess && XML.PopulateCombo(document,frmScreen.cboCreditScheme,"SpecialGroup",true);
		<% /* PSC 15/01/2007 EP2_741 - End */ %>
		if(blnSuccess == false)
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}

	XML = null;
}
		
function PopulatePropertyAddressCombo()
{
		// Code not yet written for Address popup! hence:
		//if (frmScreen.cboPropertyAddress.disabled)
		//scScreenFunctions.SetFieldToWritable(frmScreen,"cboPropertyAddress");
						
		while(frmScreen.cboPropertyAddress.options.length > 0) frmScreen.cboPropertyAddress.remove(0);

		<% /* PSC 07/08/2002 BMIDS00006 - Start */ %>
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

		var sPostCode;
		var sHouseName;
		var sHouseNumber;

		m_XMLAddressList = new top.frames[1].document.all.scXMLFunctions.XMLObject();

		<% /* BMIDS00510 MDC 30/09/2002 */ %>
		if(m_bDownloadedOwnAccount && m_xmlSecurityAddress != null)
		{
			m_XMLAddressList.LoadXML(m_xmlSecurityAddress.xml);
			m_XMLAddressList.CreateTagList("SECURITYADDRESS");
			m_XMLAddressList.SelectTagListItem(0);
		
			var sItem = frmScreen.document.createElement("OPTION");
			sItem.text = (GetAddressString())
			sItem.value = m_XMLAddressList.GetTagText("ADDRESSGUID");
			frmScreen.cboPropertyAddress.add(sItem);
			frmScreen.cboPropertyAddress(0).setAttribute("PostCode", "");
			frmScreen.cboPropertyAddress(0).setAttribute("HouseNumber", "");
			frmScreen.cboPropertyAddress(0).setAttribute("HouseName", "");
		
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboPropertyAddress");
			m_XMLAddressList = null; 
		}
		else
		{
		<% /* BMIDS00510 MDC 30/09/2002  - End */ %>

			m_XMLRelationships.ActiveTag = null;
			m_XMLRelationships.CreateTagList("ACCOUNTRELATIONSHIP");
		
			<% /* Get the customer addresses if there are customers attached to the account */ %>
			if (m_XMLRelationships.ActiveTagList.length > 0)
			{
				<% /* Set namespace and language properties so that we can use the translate function */ %>
				<% /* MO - 06/11/2002 - BMIDS00847 - No longer needed */ %>
				<% /* m_XMLAddressList.XMLDocument.setProperty("SelectionNamespaces", "xmlns:xsl='http://www.w3.org/1999/XSL/Transform'");
				m_XMLAddressList.XMLDocument.setProperty("SelectionLanguage","XPath");*/ %>
				m_XMLAddressList.CreateActiveTag("CUSTOMERADDRESSLIST")
	
				<% /* Get addresses for all attached customers */ %>
				XML.CreateRequestTag(window, null);
				XML.CreateActiveTag("CUSTOMERADDRESSLIST");

				for (var nIndex = 0; nIndex < m_XMLRelationships.ActiveTagList.length; nIndex++)
				{
					m_XMLRelationships.SelectTagListItem(nIndex);
					
					XML.CreateActiveTag("CUSTOMERADDRESS");
					XML.CreateTag("CUSTOMERNUMBER"  ,m_XMLRelationships.GetTagText("CUSTOMERNUMBER"));
					XML.CreateTag("CUSTOMERVERSIONNUMBER" ,m_XMLRelationships.GetTagText("CUSTOMERVERSIONNUMBER"));
					XML.SelectTag(null,"CUSTOMERADDRESSLIST");
				}
		
				// 				XML.RunASP(document, "FindCustomerAddressList.asp");				
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									XML.RunASP(document, "FindCustomerAddressList.asp");				
						break;
					default: // Error
						XML.SetErrorResponse();
					}

									
				var ErrorTypes = new Array("RECORDNOTFOUND");
				var ErrorReturn = XML.CheckResponse(ErrorTypes);
				
				<% /* Sort into CUSTOMERNUMBER, CUSTOMERADDRESSSEQUENCENUMBER order */ %>
				var xslPattern = "<?xml version=\"1.0\"?>" +
								 "<xsl:stylesheet version=\"1.0\" " +
								 "xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\">" +
									"<xsl:template match=\"RESPONSE/CUSTOMERADDRESSLIST\">" +
										"<CUSTOMERADDRESSLIST>" +
											"<xsl:for-each select=\"CUSTOMERADDRESS\"> " +
												"<xsl:sort select=\"CUSTOMERNUMBER\"/>" +
												"<xsl:sort select=\"CUSTOMERADDRESSSEQUENCENUMBER\"/>" +
												"<xsl:copy-of select=\".\"/>" +
											"</xsl:for-each>" +
										"</CUSTOMERADDRESSLIST>" +
									"</xsl:template>" +
								"</xsl:stylesheet>";
								
				var xslDoc = new ActiveXObject("Microsoft.XMLDOM");
				xslDoc.loadXML(xslPattern);

				var sSorted = XML.XMLDocument.transformNode(xslDoc);
				
				XML.LoadXML(sSorted);
				
				XML.CreateTagList("CUSTOMERADDRESS");

				var sSearch;
				         
				<% /* De Duplicate addresses */ %>
				for (var nIndex1 = 0; nIndex1 < XML.ActiveTagList.length; nIndex1++)
				{
					XML.SelectTagListItem(nIndex1);
					XML.SelectTag(XML.ActiveTag, "ADDRESS");
						
					sPostCode = XML.GetTagText("POSTCODE");
					sHouseNumber = XML.GetTagText("BUILDINGORHOUSENUMBER");
					sHouseName = XML.GetTagText("BUILDINGORHOUSENAME");
					
					<% /* MO - 06/11/2002 - BMIDS00847 - Modified the search so that the translate method isnt needed */ %>
					<% /* sSearch = "CUSTOMERADDRESS[ADDRESS[POSTCODE[translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '" + sPostCode.toUpperCase() + 
							  "'] and BUILDINGORHOUSENAME[translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')= '" + sHouseName.toUpperCase() + 
							  "'] and BUILDINGORHOUSENUMBER[translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='" + sHouseNumber.toUpperCase() + "']]]";  */ %>
					
					sSearch = "CUSTOMERADDRESS[ADDRESS[POSTCODE $ieq$ '" + sPostCode.toUpperCase() + 
							  "' and BUILDINGORHOUSENAME $ieq$ '" + sHouseName.toUpperCase() + 
							  "' and BUILDINGORHOUSENUMBER = '" + sHouseNumber.toUpperCase() + "']]"; 
					
					
					<% /* See if we already have the address by trying to match on house name
					      house number and postcode" */ %>
					var XMLNode = m_XMLAddressList.ActiveTag.selectSingleNode(sSearch);
						
					<% /* If not in the list already add it in */ %>
					if (!XMLNode)
					{
						XML.SelectTagListItem(nIndex1);
						m_XMLAddressList.ActiveTag.appendChild (XML.ActiveTag.cloneNode(true));
					}
				}
			}		

			// Populate the combo with the XML:
			m_XMLAddressList.CreateTagList("CUSTOMERADDRESS"); 
			<% /* PSC 07/08/2002 BMIDS00006 - End */ %>
		
			var sDefault = ""
			var nAddressCount = 0;
			var sSelectedGuid = "";
		
			for (var nLoop=0; m_XMLAddressList.SelectTagListItem(nLoop) != false && nLoop < 10; nLoop++)
			{	
				var sItem = frmScreen.document.createElement("OPTION");
				//sItem.value = XML.GetTagText("ADDRESSGUID") //nLoop
				// IVW
				<% /* PSC 07/08/2002 BMIDS00006 */ %>
				sItem.value = m_XMLAddressList.GetTagText("ADDRESSGUID")
				sItem.text = (GetAddressString())
				<% /* PSC 07/08/2002 BMIDS00006*/ %>
				sItem.tag = m_XMLAddressList.GetTagText("ADDRESSTYPE")
						
				if (sItem.tag == "1") 
					sDefault = sItem.value;
					
				frmScreen.cboPropertyAddress.add(sItem);
				
				sPostCode = m_XMLAddressList.GetTagText("POSTCODE");
				sHouseName = m_XMLAddressList.GetTagText("BUILDINGORHOUSENAME");
				sHouseNumber = m_XMLAddressList.GetTagText("BUILDINGORHOUSENUMBER");

				<% /* PSC 07/08/2002 BMIDS00006 - Start*/ %>
				<% /* Set tha attributes for each row so that we can use them to set 
				      the combo */ %>
				frmScreen.cboPropertyAddress(nAddressCount).setAttribute("PostCode", sPostCode);
				frmScreen.cboPropertyAddress(nAddressCount).setAttribute("HouseNumber", sHouseNumber);
				frmScreen.cboPropertyAddress(nAddressCount).setAttribute("HouseName", sHouseName);
				
				<% /* Check to see if this address is the one that was selected before
				      the combo was re-populated. The guid of the address may be 
				      different as the original may have been removed during de-duping.
				      Therefore the only way of seeing if it is the same address is to match
				      on postcode, house number and house name. If it matches store the Guid */ %>
				if (sPostCode.toUpperCase() == m_sSelPostCode.toUpperCase() && 
				    sHouseName.toUpperCase() == m_sSelHouseName.toUpperCase() &&
				    sHouseNumber.toUpperCase() == m_sSelHouseNumber.toUpperCase())
					sSelectedGuid = m_XMLAddressList.GetTagText("ADDRESSGUID")
				<% /* PSC 07/08/2002 BMIDS00006 - End*/ %>

				nAddressCount++;
			}
		
			// IVW add Not Applicable into the address combo
		
			var sItemNotApplicable = frmScreen.document.createElement("OPTION");
			<% /* MF MARS20 Change text */ %>
			sItemNotApplicable.text = "Other Property";
			sItemNotApplicable.value = "";
			frmScreen.cboPropertyAddress.add(sItemNotApplicable);
			frmScreen.cboPropertyAddress(nAddressCount).setAttribute("PostCode", "");
			frmScreen.cboPropertyAddress(nAddressCount).setAttribute("HouseNumber", "");
			frmScreen.cboPropertyAddress(nAddressCount).setAttribute("HouseName", "");
		
			if (nAddressCount != 0)
				scScreenFunctions.SetFieldToWritable(frmScreen,"cboPropertyAddress");
			else			
				scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboPropertyAddress");	
					
				
			<% /* PSC 07/08/2002 BMIDS00006 - Start 
			    If a match of address was found then use the guid to set the combo */ %>
			if (sSelectedGuid != "")
					frmScreen.cboPropertyAddress.value = sSelectedGuid;
			else if (sDefault != "")	
				frmScreen.cboPropertyAddress.value = sDefault;
			else
				frmScreen.cboPropertyAddress.value = ""; // Not Applicable	
			<% /* PSC 07/08/2002 BMIDS00006 - End */ %>				

			<% /* BMIDS00682 MDC 25/10/2002 				
			frmScreen.cboPropertyAddress.onchange()	*/ %>
		
			m_XMLAddressList = null; 
		}

		<% /* BMIDS00682 MDC 25/10/2002 */ %>				
		frmScreen.cboPropertyAddress.onchange()
		<% /* BMIDS00682 MDC 25/10/2002 - End */ %>				

}

function GetAddressString()
{		
	function AddComma(sAddress, sTagValue, bAddComma)
	{																						
		if ((bAddComma==true) && (sTagValue != "") && (sAddress != "")) sTagValue = ", " + sTagValue;
		sAddress += sTagValue;
		return sAddress;
	}	

	var sAddress = "";
	
	<% /* PSC 07/08/2002 BMIDS00006 - Start */ %>					
	sAddress = AddComma(sAddress, m_XMLAddressList.GetTagText("POSTCODE"), false);
	sAddress = AddComma(sAddress, m_XMLAddressList.GetTagText("FLATNUMBER"), true);
	sAddress = AddComma(sAddress, m_XMLAddressList.GetTagText("BUILDINGORHOUSENAME"), true);
	sAddress = AddComma(sAddress, m_XMLAddressList.GetTagText("BUILDINGORHOUSENUMBER"), true); <% /* APS UNIT TEST REF 21 */ %>
	sAddress = AddComma(sAddress, m_XMLAddressList.GetTagText("STREET"), true);
	sAddress = AddComma(sAddress, m_XMLAddressList.GetTagText("DISTRICT"), true);
	sAddress = AddComma(sAddress, m_XMLAddressList.GetTagText("TOWN"), true);
	sAddress = AddComma(sAddress, m_XMLAddressList.GetTagText("COUNTY"), true);
	sAddress = AddComma(sAddress, m_XMLAddressList.GetTagAttribute("COUNTRY","TEXT"), true);
	<% /* PSC 07/08/2002 BMIDS00006 - End */ %>
						
	return sAddress;
}


<% /* Retrieves the data and sets the screen accordingly */%>
function PopulateScreen()
{
	XMLOnEntry	= new top.frames[1].document.all.scXMLFunctions.XMLObject();

	<% /* PSC 07/08/2002 BMIDS00006 - Start*/ %>
	XMLOnEntry.CreateRequestTag(window, null);
	XMLOnEntry.CreateActiveTag("MORTGAGEACCOUNT");
	XMLOnEntry.CreateTag("ACCOUNTGUID", m_sAccountGUID);
	XMLOnEntry.RunASP(document, "GetMortgageAccountDetails.asp");
	
	if(XMLOnEntry.IsResponseOK())
	{
		<% /* Take a copy of the account relationships */ %>
		XMLOnEntry.SelectTag(null, "ACCOUNTRELATIONSHIPLIST");
		m_XMLRelationships.AddXMLBlock(XMLOnEntry.CreateFragment());  
	}			
	<% /* PSC 30/07/2002 BMIDS00006 - End */ %>
		
	if(XMLOnEntry.SelectTag(null,"MORTGAGEACCOUNT") != null)
	{	
		<% /* BMIDS00510 MDC 30/09/2002 */ %>
		if(m_sMetaAction == "Edit" && XMLOnEntry.GetTagText("BMIDSACCOUNT") == "1" &&
					 	XMLOnEntry.GetTagText("IMPORTEDINDICATOR") == "1")
		{
			m_bDownloadedOwnAccount = true;
			var xmlSelectedTag = XMLOnEntry.ActiveTag;
			m_xmlSecurityAddress = XMLOnEntry.SelectTag(null,"SECURITYADDRESS");
			XMLOnEntry.ActiveTag = xmlSelectedTag;
		}
		<% /* BMIDS00510 MDC 30/09/2002 - End */ %>
	
		PopulateListBox(0);
		PopulatePropertyAddressCombo();

		frmScreen.txtAccountNumber.value = XMLOnEntry.GetTagText("ACCOUNTNUMBER");
		frmScreen.txtAdditionalDetails.value = XMLOnEntry.GetTagText("ADDITIONALDETAILS");

		<% /* PSC 06/08/2002 BMIDS00006 */ %>
		frmScreen.txtRentalIncome.value = XMLOnEntry.GetTagText("MONTHLYRENTALINCOME");
		
		<% /* AShaw 24/11/2006 Ep2_18 Add Pre-emptive end date. */ %>
		frmScreen.txtPreEmptionEndDate.value = XMLOnEntry.GetTagText("PREEMPTIONENDDATE");

		<% /*MAH 21/11/2006 E2CR35*/ %>
		scScreenFunctions.SetRadioGroupValue(frmScreen, "OtherPropMortInd", XMLOnEntry.GetTagText("OTHERPROPERTYMORTGAGE"));
	
		<% /* MF 10/08/2005 MAR20 */ %>		
		frmScreen.cboRedemptionStatus.value = XMLOnEntry.GetTagText("REDEMPTIONSTATUS");
		frmScreen.txtTotalMonthly.value = XMLOnEntry.GetTagText("TOTALMONTHLYCOST");

		<% /* PSC 15/01/2007 EP2_741 - Start */ %>
		frmScreen.cboNatureOfLoan.value = XMLOnEntry.GetTagText("ORIGINALNATUREOFLOAN");
		frmScreen.cboCreditScheme.value = XMLOnEntry.GetTagText("ORIGINALCREDITSCHEME");
		frmScreen.txtBankName.value = XMLOnEntry.GetTagText("BANKNAME");
		<% /* PSC 15/01/2007 EP2_741 - End */ %>

		// IVW set combo to original value
		frmScreen.cboPropertyAddress.value = XMLOnEntry.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER");
		 
		<% /* PSC 06/08/2002 BMIDS00006 */ %>
		scScreenFunctions.SetRadioGroupValue(frmScreen, "OwnAccount", XMLOnEntry.GetTagText("BMIDSACCOUNT"));
		scScreenFunctions.SetRadioGroupValue(frmScreen, "SecondChargeIndicator", XMLOnEntry.GetTagText("SECONDCHARGEINDICATOR"));

		<% /* BMIDS00336 MDC 29/08/2002 */ %>

		if(XMLOnEntry.GetTagText("CREDITSEARCH") == "1")
		{
			scScreenFunctions.SetRadioGroupValue(frmScreen, "CreditCheckIndicator", "1");
			if(XMLOnEntry.GetTagText("UNASSIGNED") == "1")
				scScreenFunctions.SetRadioGroupValue(frmScreen, "UnassignedIndicator", "1");
			else
			{
				scScreenFunctions.SetRadioGroupValue(frmScreen, "UnassignedIndicator", "0");
				scScreenFunctions.SetRadioGroupToReadOnly(frmScreen, "UnassignedIndicator");
			}
		}
		else
		{
			scScreenFunctions.SetRadioGroupValue(frmScreen, "CreditCheckIndicator", "0");
			scScreenFunctions.SetRadioGroupValue(frmScreen, "UnassignedIndicator", "0");
			scScreenFunctions.SetRadioGroupToReadOnly(frmScreen, "UnassignedIndicator");
		}
		<% /* BMIDS00336 MDC 29/08/2002 - End */ %>
		
		<% /* BMIDS00444 */ %>
		if(XMLOnEntry.GetTagText("REMORTGAGEINDICATOR") == "1")
			scScreenFunctions.SetRadioGroupValue(frmScreen, "RemortgageIndicator", "1");
		else
		{
			scScreenFunctions.SetRadioGroupValue(frmScreen, "RemortgageIndicator", "0");
			SetRemortgageIndicatorStatus();
		}
		<% /* BMIDS00444 End */ %>
		
		<%/*START:- AQR:BMIDS756 */%>
			try
			{
				frmScreen.txtCallateralID.value = XMLOnEntry.GetTagText("COLLATERALID");
				frmScreen.txtCollateralBalance.value = XMLOnEntry.GetTagText("TOTALCOLLATERALBALANCE");
				frmScreen.txtPaymentDueDate.value = XMLOnEntry.GetTagText("PAYMENTDUEDATE");
				
				if(XMLOnEntry.GetTagText("RMCFLAG") == "1")
				{
					scScreenFunctions.SetRadioGroupValue(frmScreen,"RMCFlag","1");
				}
				else
				{
					scScreenFunctions.SetRadioGroupValue(frmScreen,"RMCFlag","0");
				}
		
				frmScreen.txtInterestRate.value = XMLOnEntry.GetTagText("DAILYMONTHLYANNUALINTEREST");
				if(XMLOnEntry.GetTagText("DSSFLAG") == "1")
				{
					scScreenFunctions.SetRadioGroupValue(frmScreen,"DSSFlag","1");
				}
				else
				{
					scScreenFunctions.SetRadioGroupValue(frmScreen,"DSSFlag","0");
				}		
				frmScreen.txtSortCode.value = XMLOnEntry.GetTagText("BANKSORTCODE");
				frmScreen.txtBankAccountNumber.value = XMLOnEntry.GetTagText("BANKACCOUNTNUMBER");
				frmScreen.txtBankAccountName.value = XMLOnEntry.GetTagText("BANKACCOUNTNAME");
				frmScreen.txtBusinessChannel.value = XMLOnEntry.GetTagText("BUSINESSCHANNEL");
			}
			catch(e)
			{
			}
		<%/*END:- AQR:BMIDS756 */%>

		
		<% /* BMIDS667 */ %>
		scScreenFunctions.SetRadioGroupValue(frmScreen, "RepossessionIndicator", XMLOnEntry.GetTagText("REPOSSESSIONFLAG"));
		<% /* BMIDS667 End */ %>
		
		<% /* Indemnity details */ %>
		if(XMLOnEntry.SelectTag(null,"INDEMNITYINSURANCE") != null)
		{
			frmScreen.txtIndemnityCompanyName.value = XMLOnEntry.GetTagText("INDEMNITYCOMPANYNAME");
			frmScreen.txtIndemnityAmount.value = XMLOnEntry.GetTagText("INDEMNITYAMOUNT");
			frmScreen.txtIndemnityMortgageAmount.value = XMLOnEntry.GetTagText("INDEMNITYMORTGAGEAMOUNT");
			frmScreen.txtIndemnityDatePaid.value = XMLOnEntry.GetTagText("INDEMNITYDATEPAID");
		}
		<% /* MAR1706 M Heys 04/05/2006 */ %>
		else
		{
			m_blnCreateIndeminityInsuranceOnUpdate = true;
		}

		<%/* Lender details */%>
		XMLOnEntry.SelectTag(null,"MORTGAGEACCOUNT");
		m_sDirectoryGUID = XMLOnEntry.GetTagText("DIRECTORYGUID");
		with (frmScreen)
		{
			txtCompanyName.value = XMLOnEntry.GetTagText("COMPANYNAME");
			txtContactForename.value = XMLOnEntry.GetTagText("CONTACTFORENAME");
			txtContactSurname.value = XMLOnEntry.GetTagText("CONTACTSURNAME");
			cboTitle.value = XMLOnEntry.GetTagText("CONTACTTITLE");
			<% /* PB 07/07/2006 EP543 Begin */ %>
			checkOtherTitleField();
			<% /* AW 24/07/2006 EP1023 */ %>
			txtTitleOther.value = XMLOnEntry.GetTagText("CONTACTTITLEOTHER"); 
			<% /* EP543 End */ %>

			// MDC SYS2564/SYS2785 
			objDerivedOperations.LoadCounty(XMLOnEntry);
			//txtCounty.value = XMLOnEntry.GetTagText("COUNTY");
			
			txtDistrict.value = XMLOnEntry.GetTagText("DISTRICT");
			//txtEMailAddress.value = XMLOnEntry.GetTagText("EMAILADDRESS");
			//txtFaxNo.value = XMLOnEntry.GetTagText("FAXNUMBER");
			txtFlatNumber.value = XMLOnEntry.GetTagText("FLATNUMBER");
			txtHouseName.value = XMLOnEntry.GetTagText("BUILDINGORHOUSENAME");
			txtHouseNumber.value = XMLOnEntry.GetTagText("BUILDINGORHOUSENUMBER");
			txtPostcode.value = XMLOnEntry.GetTagText("POSTCODE");
			txtStreet.value = XMLOnEntry.GetTagText("STREET");
			//txtTelephoneNo.value = XMLOnEntry.GetTagText("TELEPHONENUMBER");
			txtTown.value = XMLOnEntry.GetTagText("TOWN");
			cboCountry.value = XMLOnEntry.GetTagText("COUNTRY");
			cboOrganisationType.value = XMLOnEntry.GetTagText("ORGANISATIONTYPE");
			<% /* PSC 07/08/2002 BMIDS00006 */ %>
			cboPropertyAddress.value = XMLOnEntry.GetTagText("MAADDRESSGUID");
		}

		m_sDirectoryGUID = XMLOnEntry.GetTagText("DIRECTORYGUID");
		m_bDirectoryAddress = (m_sDirectoryGUID != "");
		m_sThirdPartyGUID = XMLOnEntry.GetTagText("THIRDPARTYGUID");
		
<%/*	JLD SYS3851
		// 14/09/2001 JR OmiPlus 24 
		var TempXML = XMLOnEntry.ActiveTag;
		var ContactXML = XMLOnEntry.SelectTag(null, "CONTACTDETAILS");
		if(ContactXML != null)
			m_sXMLContact = ContactXML.xml;
		XMLOnEntry.ActiveTag = TempXML;  */
%>		//JLD Get the thirdparty/directory details using the relevant guid
		var xmlContactDetails = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		xmlContactDetails.CreateRequestTag(window, "");
		if(m_sDirectoryGUID != "")
		{
			xmlContactDetails.CreateActiveTag("NAMEANDADDRESSDIRECTORY");
			xmlContactDetails.CreateTag("DIRECTORYGUID", m_sDirectoryGUID);
		}
		else
		{
			xmlContactDetails.CreateActiveTag("THIRDPARTY");
			xmlContactDetails.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID);
		}
		
		xmlContactDetails.RunASP(document, "GetThirdParty.asp");
		if(xmlContactDetails.IsResponseOK())
		{
			xmlContactDetails.SelectTag(null, "CONTACTDETAILS");
			m_sXMLContact = xmlContactDetails.ActiveTag.xml;
		}
		
	}
	<% /* JD MAR1040 Check if we have any mortgageloans. If we have, disable the monthly cost and redemption status */ %>
	if(HaveMortgageLoans())
	{
		frmScreen.cboRedemptionStatus.disabled = true;
		frmScreen.txtTotalMonthly.disabled = true;
		frmScreen.optSecondChargeNo.disabled = true;
		frmScreen.optSecondChargeYes.disabled = true;
	}
}
<% /* JD MAR1040 */ %>
function HaveMortgageLoans()
{
	var bLoansFound = false;
	
	var ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(window, "SEARCH");
	ListXML.CreateActiveTag("MORTGAGELOANLIST");
	ListXML.CreateActiveTag("MORTGAGELOAN");
	ListXML.CreateTag("ACCOUNTGUID",m_sAccountGUID);			
	ListXML.RunASP(document, "FindMortgageLoanList.asp");
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);
	
	if(	sResponseArray[0] == true )
	{
		//check we have the mandatory mortgageloan info
		ListXML.ActiveTag = null;
		ListXML.CreateTagList("MORTGAGELOAN");
			
		for (var nLoop=0; ListXML.SelectTagListItem(nLoop) != false ; nLoop++)
		{				
			if(ListXML.GetTagText("MONTHLYREPAYMENT") != "")
				bLoansFound = true;
		}
	}

	return bLoansFound;
}
function RetrieveContextData()
{
<%	/*JR - Test
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber","02013118");
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber","1");
	End */
%>	
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber",null);
	m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber",null);
	
	<% /* PSC 07/08/2002 BMIDS00006 */ %>	
	m_sAccountGUID = scScreenFunctions.GetContextParameter(window,"idAccountGuid",null);
	
	var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bCanAddToDirectory = (sUserRole >= XML.GetGlobalParameterAmount(document,"AddToDirectoryRole"));

	<% /* SR 14/06/2004 : BMIDS772 */ %>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber");
	<% /* SR 14/06/2004 : BMIDS772 - End */ %>

	XML = null;
}

function SaveMortgageAccount()
{
	
	var bOK = false;			
	var TagRequestType = null; 
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	if(m_sMetaAction == "Add")
		TagRequestType = XML.CreateRequestTag(window, "CREATE");
	else if(m_sMetaAction == "Edit")
		TagRequestType = XML.CreateRequestTag(window, "UPDATE");
	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");

	<%/* MORTGAGEACCOUNT */%>
	var TagMORTGAGEACCOUNT = XML.CreateActiveTag("MORTGAGEACCOUNT");
	if(m_sMetaAction == "Edit")
		XML.CreateTag("ACCOUNTGUID",m_sAccountGUID);
	XML.CreateTag("SECONDCHARGEINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"SecondChargeIndicator"));
	
	<% /* BMIDS00444 */ %>
	XML.CreateTag("REMORTGAGEINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"RemortgageIndicator"));
	<% /* BMIDS00444 End */ %>
	<% /* BMIDS667 */ %>
	XML.CreateTag("REPOSSESSIONFLAG",scScreenFunctions.GetRadioGroupValue(frmScreen,"RepossessionIndicator"));
	<% /* BMIDS00444 End */ %>
	
	<% /* PSC 06/08/2002 BMIDS00006 - Start */ %>
	XML.CreateTag("BMIDSACCOUNT",scScreenFunctions.GetRadioGroupValue(frmScreen,"OwnAccount"));
	XML.CreateTag("ADDITIONALDETAILS",frmScreen.txtAdditionalDetails.value);
	if (frmScreen.txtAdditionalDetails.value == "")
		XML.CreateTag("ADDITIONALINDICATOR",0);
	else
		XML.CreateTag("ADDITIONALINDICATOR",1);
		
	XML.CreateTag("MONTHLYRENTALINCOME",frmScreen.txtRentalIncome.value);
	<% /* MF 10/08/2005 MAR20 */ %>
	XML.CreateTag("REDEMPTIONSTATUS",frmScreen.cboRedemptionStatus.value);	
	XML.CreateTag("TOTALMONTHLYCOST",frmScreen.txtTotalMonthly.value);	
	XML.CreateTag("ADDRESSGUID",frmScreen.cboPropertyAddress.value);
	
	<% /* AShaw 24/11/2006 Ep2_18 Add Pre-emptive end date. */ %>
	XML.CreateTag("PREEMPTIONENDDATE",frmScreen.txtPreEmptionEndDate.value);	
	
	<% /* PSC 15/01/2007 EP2_741 - Start */ %>
	XML.CreateTag("ORIGINALNATUREOFLOAN",frmScreen.cboNatureOfLoan.value);	
	XML.CreateTag("ORIGINALCREDITSCHEME",frmScreen.cboCreditScheme.value);	
	XML.CreateTag("BANKNAME",frmScreen.txtBankName.value);
	<% /* PSC 15/01/2007 EP2_741 - End */ %>

	<% /*MAH 21/11/2006 E2CR35*/ %>
	XML.CreateTag("OTHERPROPERTYMORTGAGE",scScreenFunctions.GetRadioGroupValue(frmScreen,"OtherPropMortInd"));
	
	<%/*START:- AQR:BMIDS756 */%>
	XML.CreateTag("COLLATERALID",frmScreen.txtCallateralID.value);
	XML.CreateTag("TOTALCOLLATERALBALANCE",frmScreen.txtCollateralBalance.value);
	XML.CreateTag("PAYMENTDUEDATE",frmScreen.txtPaymentDueDate.value);
	XML.CreateTag("RMCFLAG",scScreenFunctions.GetRadioGroupValue(frmScreen,"RMCFlag"));
	XML.CreateTag("DAILYMONTHLYANNUALINTEREST",frmScreen.txtInterestRate.value);
	XML.CreateTag("DSSFLAG",scScreenFunctions.GetRadioGroupValue(frmScreen,"DSSFlag"));
	XML.CreateTag("BANKSORTCODE",frmScreen.txtSortCode.value);
	XML.CreateTag("BANKACCOUNTNUMBER",frmScreen.txtBankAccountNumber.value);
	XML.CreateTag("BANKACCOUNTNAME",frmScreen.txtBankAccountName.value);
	XML.CreateTag("BUSINESSCHANNEL",frmScreen.txtBusinessChannel.value);
	<%/*END:- AQR:BMIDS756 */%>

	m_XMLPropertyDetails.SelectTag(null, "PROPERTYDETAILS");
	
	<% /* PSC 21/08/2002 BMIDS00350 */ %>
	if (m_XMLPropertyDetails.ActiveTag != null)
	{
		XML.CreateTag("LASTVALUERNAME",m_XMLPropertyDetails.GetTagText("LASTVALUERNAME"));
		XML.CreateTag("LASTVALUATIONDATE",m_XMLPropertyDetails.GetTagText("LASTVALUATIONDATE"));
		XML.CreateTag("LASTVALUATIONAMOUNT",m_XMLPropertyDetails.GetTagText("LASTVALUATIONAMOUNT"));
		XML.CreateTag("REINSTATEMENTAMOUNT",m_XMLPropertyDetails.GetTagText("REINSTATEMENTAMOUNT"));
		XML.CreateTag("ESTIMATEDCURRENTVALUE",m_XMLPropertyDetails.GetTagText("ESTIMATEDCURRENTVALUE"));
		XML.CreateTag("BUILDINGSSUMINSURED",m_XMLPropertyDetails.GetTagText("BUILDINGSSUMINSURED"));
		XML.CreateTag("HOMEINSURANCETYPE",m_XMLPropertyDetails.GetTagText("HOMEINSURANCETYPE"));
	}
	<% /* PSC 06/08/2002 BMIDS00006 - End */ %>
		
	<%/* ACCOUNT */%>
	XML.ActiveTag = TagMORTGAGEACCOUNT;
	XML.CreateActiveTag("ACCOUNT");
	if(m_sMetaAction == "Edit")
		XML.CreateTag("ACCOUNTGUID",m_sAccountGUID);
	XML.CreateTag("ACCOUNTNUMBER",frmScreen.txtAccountNumber.value);
	<% /* BMIDS00336 MDC 30/08/2002 */ %>
	XML.CreateTag("UNASSIGNED",scScreenFunctions.GetRadioGroupValue(frmScreen,"UnassignedIndicator"));
	<% /* BMIDS00336 MDC 30/08/2002 */ %>
	
	
	
	if (m_sMetaAction == "Edit")
	{
		// Retrieve the original third party/directory GUIDs
		var sOriginalThirdPartyGUID = XMLOnEntry.GetTagText("THIRDPARTYGUID");
		var sOriginalDirectoryGUID = XMLOnEntry.GetTagText("DIRECTORYGUID");
		// Only retrieve the address/contact details GUID if we are updating an existing third party record
		var sAddressGUID = (sOriginalThirdPartyGUID != "") ? XMLOnEntry.GetTagText("ADDRESSGUID") : "";
		var sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? XMLOnEntry.GetTagText("CONTACTDETAILSGUID") : "";
	}

	// If changing from a directory to a third party, or vice-versa, the old directory/thirdparty GUID
	// should still be specified to alert the middler tier to the fact that the old link needs deleting
	XML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID != "" ? m_sDirectoryGUID : sOriginalDirectoryGUID);
	XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);

	if ((m_bDirectoryAddress==false) && (AllFieldsEmpty()==false)) // Note that AllFieldsEmpty is in the ThirdPartyDetails.asp
	{
		// Store the third party details
		XML.CreateActiveTag("THIRDPARTY");
		XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
		XML.CreateTag("THIRDPARTYTYPE", 3); // Bank/Building society
		XML.CreateTag("COMPANYNAME", frmScreen.txtCompanyName.value);
		XML.CreateTag("ORGANISATIONTYPE", frmScreen.cboOrganisationType.value);

		// Address
		XML.CreateTag("ADDRESSGUID", sAddressGUID);
		SaveAddress(XML, sAddressGUID);

		// Contact Details
		XML.SelectTag(null, "THIRDPARTY");
		XML.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
		SaveContactDetails(XML, sContactDetailsGUID);
	}
	
	<% /* ACCOUNTRELATIONSHIP */ %>
	XML.ActiveTag = TagMORTGAGEACCOUNT;
	XML.CreateActiveTag("ACCOUNTRELATIONSHIPLIST");
	m_XMLRelationships.SelectTag(null, "ACCOUNTRELATIONSHIPLIST");
	XML.AddXMLBlock(m_XMLRelationships.CreateFragment());
	
	<%/* INDEMNITYINSURANCE */%>
	XML.ActiveTag = TagMORTGAGEACCOUNT;
	XML.CreateActiveTag("INDEMNITYINSURANCE");
	<% //MAR1448 M Heys & G Hun 21/04/2006 start %>
	<% //if(XMLOnEntry.SelectTag(null,"INDEMNITYINSURANCE") == null) -- MAR1706 MH 04/05/2006 %>
	if (m_blnCreateIndeminityInsuranceOnUpdate) <% // ++ MAR1706 %>
		XML.SetAttribute("CREATE", "1");
	<% //MAR1448 M Heys & G Hun 21/04/2006 end %>
	if(m_sMetaAction == "Edit")
		XML.CreateTag("ACCOUNTGUID",m_sAccountGUID);

	XML.CreateTag("INDEMNITYCOMPANYNAME",frmScreen.txtIndemnityCompanyName.value);
	XML.CreateTag("INDEMNITYAMOUNT",frmScreen.txtIndemnityAmount.value);
	XML.CreateTag("INDEMNITYMORTGAGEAMOUNT",frmScreen.txtIndemnityMortgageAmount.value);
	XML.CreateTag("INDEMNITYDATEPAID",frmScreen.txtIndemnityDatePaid.value);

	
	<% /* SR 14/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
						  node to CustomerFinancialBO. 
		NOTE: [MC]*= SRINI'S code block moved here. ThirdParty Details wasn't stored for new financial records..*/ %>
	if(m_sMetaAction == "Add")
	{
		XML.ActiveTag = TagRequestType
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("EXISTINGMORTGAGEINDICATOR", 1);
	}
	<% /* SR 14/06/2004 : BMIDS772 - End */ %>
	

	if(m_sMetaAction == "Add")
		// 		XML.RunASP(document, "CreateMortgageAccount.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "CreateMortgageAccount.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	else
		// 		XML.RunASP(document, "UpdateMortgageAccount.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "UpdateMortgageAccount.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}


	if(XML.IsResponseOK() == true)
	{
		bOK = true;					
		if(m_sMetaAction == "Add")
		{
			m_sAccountGUID = XML.GetTagText("ACCOUNTGUID");		
			<% /* PSC 07/08/2002 BMIDS00006 */ %>
			scScreenFunctions.SetContextParameter(window,"idAccountGuid",m_sAccountGUID);
		}
	}

	return bOK;
}

function ValidateScreen()
{

	<% /* PSC 07/08/2002 BMIDS00006 - Start */ %>
	m_XMLRelationships.ActiveTag = null;
	m_XMLRelationships.CreateTagList("ACCOUNTRELATIONSHIP");
	
	<% /* Check there are applicants assigned to the account */ %>
	if (m_XMLRelationships.ActiveTagList.length == 0)
	{
		alert("No applicants have been linked to this account");
		return(false);
	}
	
	<% /* Warn if applicants are present but no address selected */ %>
	if (m_XMLRelationships.ActiveTagList.length > 0 &&
	    frmScreen.cboPropertyAddress.value == "")
	{
		var bOk = false;
		<% /* BMIDS00406 remove mortgage
		// bOk = confirm("No address details exist for this mortgage account");
		*/ %>
		bOk = confirm("No address details exist for this account");
		<% /* BMIDS00406 End */ %>
		return (bOk);
	}	  
	<% /* PSC 07/08/2002 BMIDS00006 - End */ %>
	

	<% /* Check that the indemnity date isn't in the future */ %>
	if(scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtIndemnityDatePaid,">") == true)
	{
		alert("Indemnity Date cannot be in the future");
		frmScreen.txtIndemnityDatePaid.focus();
		return(false);
	}
	
	if(scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true))
		return false;

	<% //MAR1189/2161 %>
	var sMortgageType = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","");
	var ValidationList = new Array(1);
	ValidationList[0] = "R";
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var blnIsRemortgage = XML.IsInComboValidationList(document,"TypeOfMortgage",sMortgageType, ValidationList);
	
    if(scScreenFunctions.GetRadioGroupValue(frmScreen,"RemortgageIndicator") == null && blnIsRemortgage)
    {
		alert("Please state whether the account is to be remortgaged");
		frmScreen.optRemortgageIndicatorYes.focus();
		return(false);
    }

	if(scScreenFunctions.ValidatePostcode(frmScreen.txtPostcode))
		return(ValidateThirdPartyDetails);
	else
		return(false);
		
	<% /*MAH 21/11/2006 E2CR35*/ %> 
	if(scScreenFunctions.GetRadioGroupValue(frmScreen,"OtherPropMortInd") == null)
	{
		alert("Please indicate if there is another mortgaged property");
		frmScreen.optOtherPropertyMortgageYes.focus();
		return(false);
	}
	
}

function frmScreen.btnPropertyDetails.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array(4);
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	<% /* PSC 07/08/2002 BMIDS00006 - Start */ %>
	ArrayArguments[0] = m_XMLPropertyDetails.XMLDocument.xml;
	//GD BMIDS00376
	ArrayArguments[1] = m_sReadOnly;	
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "dc072.asp", ArrayArguments, 491, 335);
	
	if (sReturn != null)
	{
		m_XMLPropertyDetails.LoadXML(sReturn[1]);
		FlagChange(sReturn[0]);
	}
	<% /* PSC 07/08/2002 BMIDS00006 - End */ %>

	return(true);
}

function ValidateThirdPartyDetails()
{
	<%/* If third party details have been specified then so must a lender name (note that the
	     call to AllFieldsEmpty only checks the THIRDPARTY related fields) */%>
	if((frmScreen.txtCompanyName.value == "") && !AllFieldsEmpty())
	{
		alert("Lender details have been specified, therefore the Lender Name cannot be blank.");
		scScreenFunctions.DoFocusProcessing(frmScreen,"txtCompanyName");
		return(false);
	}

	//if((frmScreen.txtCompanyName.value != "") && (frmScreen.cboCountry.selectedIndex == 0))
	//{
	//	alert("The Country must be specified for the lender.");
	//	scScreenFunctions.DoFocusProcessing(frmScreen,"txtCompanyName");
	//	return(false);
	//}

	return(true);
}

function PopulateListBox(nStart)
{
<%  /* Populate the listbox with values from m_XMLRelationships */ 
%>	m_XMLRelationships.ActiveTag = null;
	m_XMLRelationships.CreateTagList("ACCOUNTRELATIONSHIP");

	var nNumberOfOwners = m_XMLRelationships.ActiveTagList.length;
	scScrollTable.initialiseTable(tblOwners, 0, "", ShowList, m_iTableLength, nNumberOfOwners);
	ShowList(nStart);
	

	if (m_sReadOnly == "1")
	{
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnAdd.disabled = true;
	}
	else
	{
		if(nNumberOfOwners == 0)
		{
			frmScreen.btnDelete.disabled = true;
			frmScreen.btnEdit.disabled = true;
			frmScreen.btnAdd.disabled = false;
		}
		else
		{
			if (nNumberOfOwners < m_nNoOfCustomers)
				frmScreen.btnAdd.disabled = false;
			else
				frmScreen.btnAdd.disabled = true;
			
			frmScreen.btnDelete.disabled = false;
			frmScreen.btnEdit.disabled = false;
			scScrollTable.setRowSelected(1);		
		}	
	}
}

function ShowList(nStart)
{
<%  /* Populate the listbox with values from m_XMLRelationships Also used by the list scroll object */
%> 
	scScrollTable.clear();
	
	for (var iCount = 0; iCount < m_XMLRelationships.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		m_XMLRelationships.SelectTagListItem(iCount + nStart);
		
		var Role = m_XMLRelationships.ActiveTag.selectSingleNode("CUSTOMERROLETYPE");
		var sRole = m_XMLRelationships.GetNodeAttribute(Role, "TEXT");
		var sName = m_XMLRelationships.ActiveTag.selectSingleNode("FIRSTFORENAME").text + ' ' +
					m_XMLRelationships.ActiveTag.selectSingleNode("SURNAME").text;
		var sCustomerNo = m_XMLRelationships.ActiveTag.selectSingleNode("CUSTOMERNUMBER").text;
		
		scScreenFunctions.SizeTextToField(tblOwners.rows(iCount+1).cells(0),sName);
		scScreenFunctions.SizeTextToField(tblOwners.rows(iCount+1).cells(1),sRole);
		tblOwners.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}
}

function frmScreen.btnAdd.onclick()
{
	var sReturn = "";
	var ArrayArguments = new Array(4);

	ArrayArguments[0] = "Add";
	ArrayArguments[1] = m_XMLRelationships.XMLDocument.xml;
	ArrayArguments[2] = GetApplicants("");
	ArrayArguments[3] = m_sAccountGUID;
	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	ArrayArguments[4] = m_sReadOnly;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "dc077.asp", ArrayArguments, 377, 245);
	
	if (sReturn != null)
	{
		m_XMLRelationships.LoadXML(sReturn[1]);
		PopulateListBox(0);
		FlagChange(sReturn[0]);
		PopulatePropertyAddressCombo();
	}
}

function frmScreen.btnDelete.onclick()
{
	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	m_XMLRelationships.ActiveTag = null;
	m_XMLRelationships.CreateTagList("ACCOUNTRELATIONSHIP");
	m_XMLRelationships.SelectTagListItem(nRowSelected -1);
	
	m_XMLRelationships.RemoveActiveTag();
	PopulateListBox(0);	
	FlagChange(true);
	PopulatePropertyAddressCombo();	
}


function frmScreen.btnEdit.onclick()
{
	var sReturn = "";
	var ArrayArguments = new Array(4);
	
	ArrayArguments[0] = "Edit";
	ArrayArguments[1] = m_XMLRelationships.XMLDocument.xml;

	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	m_XMLRelationships.ActiveTag = null;
	m_XMLRelationships.CreateTagList("ACCOUNTRELATIONSHIP");
	m_XMLRelationships.SelectTagListItem(nRowSelected -1);

	var sCustomerNo = m_XMLRelationships.GetTagText("CUSTOMERNUMBER");

	ArrayArguments[2] = GetApplicants(sCustomerNo);
	ArrayArguments[3] = nRowSelected -1;
	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	ArrayArguments[4] = m_sReadOnly;
	sReturn = scScreenFunctions.DisplayPopup(window, document, "dc077.asp", ArrayArguments, 377, 245);

	if (sReturn != null)
	{
		m_XMLRelationships.LoadXML(sReturn[1]);
		PopulateListBox(0);
		FlagChange(sReturn[0]);
	}
}

function GetApplicants(sInCustomerNo)
{

	m_XMLApplicants = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	m_XMLApplicants.CreateActiveTag("CUSTOMERLIST");
	
	m_XMLRelationships.SelectTag(null,"ACCOUNTRELATIONSHIPLIST");
	
	var sSearch = "";

	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		m_XMLApplicants.SelectTag(null, "CUSTOMERLIST");
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
	
		if (sCustomerNumber != "")
		{
			sSearch = "ACCOUNTRELATIONSHIP[CUSTOMERNUMBER = '" + sCustomerNumber + "']";
		
			var Relationship = m_XMLRelationships.ActiveTag.selectSingleNode(sSearch);
		
			if (Relationship == null || sCustomerNumber == sInCustomerNo)
			{
				var sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + nLoop);
				var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
				var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
				m_XMLApplicants.CreateActiveTag("CUSTOMER");
				m_XMLApplicants.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
				m_XMLApplicants.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
				m_XMLApplicants.CreateTag("CUSTOMERNAME", sCustomerName);
			}
		}
	}
	
	return(m_XMLApplicants.XMLDocument.xml);

}

function GetCustomerCount()
{
	var nCustomerCount = 0;
	
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
	
		if (sCustomerNumber != "")
			nCustomerCount++;		
	}

	return (nCustomerCount);

}

<% /* PSC 08/08/2002 BMIDS00006 - Start */ %>
function btnFeatures.onclick()
{
	if(CommitChanges())
	{
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateActiveTag("MORTGAGEACCOUNT");		
		XML.CreateTag("ORGANISATIONNAME",frmScreen.txtCompanyName.value);
		XML.CreateTag("ACCOUNTNUMBER" ,frmScreen.txtAccountNumber.value);
		XML.CreateTag("ACCOUNTGUID" ,m_sAccountGUID);
		XML.CreateTag("DIRECTORYGUID" ,m_sDirectoryGUID);
						
		scScreenFunctions.SetContextParameter(window,"idXML",XML.XMLDocument.xml);
		frmToDC075.submit();
	}
}
function btnArrears.onclick()
{
	if(CommitChanges())
	{
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateActiveTag("MORTGAGEACCOUNT");		
		XML.CreateTag("ORGANISATIONNAME",frmScreen.txtCompanyName.value);
		XML.CreateTag("ACCOUNTNUMBER" ,frmScreen.txtAccountNumber.value);
		XML.CreateTag("ACCOUNTGUID" ,m_sAccountGUID);
		XML.CreateTag("DIRECTORYGUID" ,m_sDirectoryGUID);
						
		scScreenFunctions.SetContextParameter(window,"idXML",XML.XMLDocument.xml);
		frmToDC130.submit();
	}
}

function SetupPropertyDetails()
{
	m_XMLPropertyDetails = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	if (XMLOnEntry != null)
	{
		XMLOnEntry.SelectTag(null, "MORTGAGEACCOUNT");	
		m_XMLPropertyDetails.CreateActiveTag("PROPERTYDETAILS");
		m_XMLPropertyDetails.CreateTag("LASTVALUERNAME",XMLOnEntry.GetTagText("LASTVALUERNAME"));
		m_XMLPropertyDetails.CreateTag("LASTVALUATIONDATE",XMLOnEntry.GetTagText("LASTVALUATIONDATE"));
		m_XMLPropertyDetails.CreateTag("LASTVALUATIONAMOUNT",XMLOnEntry.GetTagText("LASTVALUATIONAMOUNT"));
		m_XMLPropertyDetails.CreateTag("REINSTATEMENTAMOUNT",XMLOnEntry.GetTagText("REINSTATEMENTAMOUNT"));
		m_XMLPropertyDetails.CreateTag("ESTIMATEDCURRENTVALUE",XMLOnEntry.GetTagText("ESTIMATEDCURRENTVALUE"));
		m_XMLPropertyDetails.CreateTag("BUILDINGSSUMINSURED",XMLOnEntry.GetTagText("BUILDINGSSUMINSURED"));
		m_XMLPropertyDetails.CreateTag("HOMEINSURANCETYPE",XMLOnEntry.GetTagText("HOMEINSURANCETYPE"));
	}
}
<% /* PSC 08/08/2002 BMIDS00006 - End */ %>

<% /* BMIDS00444 Set whether the remortgage account indicator should be visible or not */ %>
function SetRemortgageIndicatorVisibility()
{
	<% /* Indicator should only be visible if TypeOfMortgage is "Remortgage" (validation type "R") */ %>
	var sMortgageType = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","");
	var ValidationList = new Array(1);
	ValidationList[0] = "R";
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var blnIsRemortgage = XML.IsInComboValidationList(document,"TypeOfMortgage",sMortgageType, ValidationList);
	
	if (blnIsRemortgage)
	{
		spnRemortgageIndicator.style.visibility = "visible";
		m_blnAccountNumberMandatory = true;		  

		<% //MAR1189/2161 %>
		frmScreen.optRemortgageIndicatorYes.setAttribute("required","true");
		frmScreen.optRemortgageIndicatorNo.setAttribute("required","true");		

		if(scScreenFunctions.GetRadioGroupValue(frmScreen,"RemortgageIndicator") == 1)
		{
			frmScreen.optRemortgageIndicatorYes.onclick()
		}
		else
		{
			frmScreen.optRemortgageIndicatorNo.onclick()
		}				
	} 
	else
	{
		spnRemortgageIndicator.style.visibility = "hidden";
		m_blnAccountNumberMandatory = false;

		<% //MAR1189/2161 %>
		frmScreen.optRemortgageIndicatorYes.removeAttribute("required");
		frmScreen.optRemortgageIndicatorNo.removeAttribute("required");				
	}
}
<% /* BMIDS00444 End */ %>

<% /* BMIDS00444 Set whether the remortgage account indicator should be enabled or not */ %>
function SetRemortgageIndicatorStatus()
{
	<% /* The indicator should only be enabled if it is NOT already true on any other
		  MortgageAccount(s) held by the customer(s) on the application */ %>		
	
	var sCustomerNumber = "";
	var sCustomerVersionNumber = "";
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "FindRemortgageAccountAddress");
	XML.CreateActiveTag("MORTGAGEACCOUNT");
	
	<% /* Add Customers to the query */ %>
	for(var nLoop = 1; nLoop <= m_nNoOfCustomers; nLoop++)
	{
		XML.CreateActiveTag("ACCOUNTRELATIONSHIP");
		sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
		XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
		XML.ActiveTag = XML.ActiveTag.parentNode;
	}
	
	XML.RunASP(document, "FindRemortgageAccountAddress.asp");				
	
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = XML.CheckResponse(sErrorArray);
			
	if ((sResponseArray[0] == false) && (sResponseArray[1] == sErrorArray[0]))
	  {
		scScreenFunctions.SetRadioGroupState(frmScreen, "RemortgageIndicator", "W");
		m_blnAccountNumberMandatory = true;
	  }	
	else
	  {
		scScreenFunctions.SetRadioGroupState(frmScreen, "RemortgageIndicator", "R");
		m_blnAccountNumberMandatory = false;
      }		
}

function frmScreen.txtPaymentDueDate.onblur()
{
	if ((parseInt(frmScreen.txtPaymentDueDate.value) < 0 ) || (parseInt(frmScreen.txtPaymentDueDate.value) > 31))
	{
		frmScreen.txtPaymentDueDate.focus();
		alert('PaymentDueDay should be between 1 to 31');
	}
}
<% //MAR1189/2161 %>
function frmScreen.optRemortgageIndicatorYes.onclick()
{    
    m_blnAccountNumberMandatory = true;    
    frmScreen.txtAccountNumber.setAttribute("required","true");		
    frmScreen.txtAccountNumber.parentElement.parentElement.style.color = "red";	
} 
<% //MAR1189/2161 %>
function frmScreen.optRemortgageIndicatorNo.onclick()
{
	m_blnAccountNumberMandatory = false;    
	frmScreen.txtAccountNumber.removeAttribute("required");
	frmScreen.txtAccountNumber.parentElement.parentElement.style.color = "";	
}
-->
</script>
<!-- #include FILE="includes/thirdpartydetails.asp" -->
</body>
</html>
