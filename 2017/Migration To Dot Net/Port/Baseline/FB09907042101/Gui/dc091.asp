<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      dc091.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Add/Edit Loans/Liabilities
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		03/02/2000	ie5 changes 
AY		14/02/00	Change to msgButtons button types
AD		20/03/2000	Incorporated third party include files.
AY		30/03/00	New top menu/scScreenFunctions change
MH      02/05/00    SYS0618 Postcode validation
BG		08/05/2000	SYS0673	Wording changed to "Is the Balance to be repaid.."
MH      02/05/00    SYS0721 Tooltips and stuff
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
MC		08/06/00	SYS0865 Add Agreement Type combo
MC		09/06/00	SYS0866 Standardise format of balance and payment fields
MH		12/06/00	SYS0721 Default country to UK
MH      20/06/00    SYS0933 ReadOnly
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
BG		04/11/00	SYS1037 When populating country combo, set final parameter to false so there is
							no "SELECT" option and remove validation for "0" in ValidateThirdPartyDetails.
CL		05/03/01	SYS1920 Read only functionality added
SA		06/06/01	SYS1672 Guarantors added to Applicants combo.
JR		14/09/01	Omiplus24 - Removed FaxNumber, TelephoneNumber & EmailAddress. Now included in TPD.asp
MDC		01/10/01	SYS2785 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
GHun	10/07/2002	BMIDS00190	DCWP3 BM076 Support multiple customers per loan/liability
GHun	20/09/2002	BMIDS00392	Grey out monthly repayment for credit cards
TW      09/10/2002    Modified to incorporate client validation - SYS5115
GHun	30/10/2002	BMIDS00731	Customers with alphas in the customer number are not displayed
SA		07/11/2002	BMIDS00832	SelectSingleNode pattern needs to deal with alpha customer numbers
MV		25/03/2003	BM00063		AMended HTML Text for Option buttons 
HMA  	16/09/2003	BM00063		Amended HTML Text for Option buttons 
HMA     26/09/2003  BM0063      Disable Directory search button when data is frozen
HMA     01/10/2003  BM0063      Disable Directory search button when user in a non-processing unit.
SR		01/06/2004	BMIDS772	Update FinancialSummary record on Submit (only for create)	
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
HM		26/07/2005	MAR250		txtLoanAccountNumber max length set to 20 chars using attr file
SR		09/11/2005	MAR496		Added combo 'Credit Card Type'
SD		14/11/2005	MAR258		Critical data check changes
HMA     14/11/2005  MAR102      Retrieve and Save OriginalLoanAmount in LoansLiabilities table.
SD		17/11/2005	MAR610		Fixing Script Error
SD		17/11/2005	MAR613		MAR610 had a duplicate
DRC     03/02/2006  MAR1189     Balance is mandatory for ceratin aggreement types
PE		07/02/2006	MAR1189		Slight reworking of the above
PE		20/03/2006	MAR1437		Associated CardType with LoansLiabilities.CardType
JD		21/04/2006	MAR1371		save bankcreditcard to the correct customers
INR		24/04/2006	MAR1306		Disable the button and display an hourglass until completenesscheck has run
GHun	17/05/2006  MAR1781		Changed ConfirmProcessing to stop if validation fails
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
SW		29/06/2006	EP898		Added CARDPROVIDER tag to SaveLiability()
LDM		19/07/2006	EP1003		If the agreement type has a validation code of "M" dont make lender name mandetory
PB		21/07/2006	EP543		Added 'TITLEOTHER' field for third parties
AW		31/07/2006	EP1059		Amended xml doc reference for ContactOtherTitle
MAH		21/11/2006	E2CR35		Added LoansLiabilities.PaidForByBusiness
SR		15/02/2007  EP2_946		modified ConfirmProcessing(). Call return after the frmToDC90.submit
LDM		02/04/2007	EP2_2164	Disable the additional details if none there and additionalind radio btns not set
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
VIEWASTEXT></OBJECT>
<script src="Validation.js" type="text/javascript"></script>
<%/* BMIDS00190 */%>
<OBJECT id=scTable style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 data=scTable.htm VIEWASTEXT></OBJECT>

<form id="frmToDC090" method="post" action="dc090.asp" STYLE="DISPLAY: none"></form>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" year4 validate="onchange" mark>
<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 504px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Applicant
	</span>
	<% /* BMIDS00190 DCWP3 BM076 replace combo with table
		<span style="TOP: -3px; LEFT: 300px; POSITION: ABSOLUTE">
			<select id="cboApplicant" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	*/ %>
	
	<div id="spnApplicant" style="LEFT: 304px; POSITION: absolute; TOP: 10px">
		<table id="tblApplicant" width="280" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">
				<td width="80%" class="TableHead">Name&nbsp;</td>
				<td width="20%" class="TableHead">Selected&nbsp;</td>
			</tr>
			<tr id="row01">
				<td width="80%" class="TableTopLeft">&nbsp;</td>
				<td width="20%" class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">
				<td width="80%" class="TableLeft">&nbsp;</td>
				<td width="20%" class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">
				<td width="80%" class="TableLeft">&nbsp;</td>
				<td width="20%" class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">
				<td width="80%" class="TableLeft">&nbsp;</td>
				<td width="20%" class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row05">
				<td width="80%" class="TableBottomLeft">&nbsp;</td>
				<td width="20%" class="TableBottomRight">&nbsp;</td>
			</tr>
		</table>
	</div>

	<span id="spnButtons" style="LEFT: 310px; POSITION: absolute; TOP: 103px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnSelectDeselect" value="Select/De-select" type="button" style="WIDTH: 100px" class="msgButton">
		</span>
	</span>
	<% /* BMIDS00190 End */ %>
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 136px" class="msgLabel">
		Agreement Type
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<select id="cboAgreementType" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 160px" class="msgLabel">
		Credit Card Type
		<span style="TOP: -3px; LEFT: 300px; POSITION: ABSOLUTE">
			<select id="cboCreditCardType" style="WIDTH: 200px" class="msgCombo" NAME="cboCreditCardType">
			</select>
		</span>
	</span>
	<span style="TOP: 184px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Loan Account Number
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="txtLoanAccountNumber" style="WIDTH: 200px; POSITION: absolute" class ="msgTxt">
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 208px" class="msgLabel">
		Total Outstanding Balance
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="txtTotalOutstandingBalance" maxlength="6" style="WIDTH: 70px; POSITION: absolute" class ="msgTxt">
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 232px" class="msgLabel">
		Monthly Repayment
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="txtMonthlyRepayment" maxlength="9" style="WIDTH: 70px; POSITION: absolute" class ="msgTxt">
		</span>
	</span>
	<% /* BMIDS00336 MDC 30/08/2002 */ %>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 256px" class="msgLabel">
		Original Loan Amount
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="txtOriginalLoanAmount" maxlength="7" style="WIDTH: 70px; POSITION: absolute" class ="msgTxt">
		</span>
	</span>
	<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 280px" class="msgLabel">
		End Date
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="txtEndDate" maxlength="10" style="WIDTH: 70px; POSITION: absolute" class ="msgTxt">
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 304px" class="msgLabel">
		Is the Balance to be repaid?
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="optLoanRepaymentYes" name="LoanRepaymentInd" type="radio" value="1"><label for="optLoanRepaymentYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 355px; POSITION: absolute; TOP: -3px">
			<input id="optLoanRepaymentNo" name="LoanRepaymentInd" type="radio" value="0" checked><label for="optLoanRepaymentNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span style="LEFT: 10px; POSITION: absolute; TOP: 328px" class="msgLabel">
			Is it paid for by your business?
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="optPaidForByBusinessYes" name="PaidForByBusinessInd" type="radio" value="1"><label for="optPaidForByBusinessYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 355px; POSITION: absolute; TOP: -3px">
			<input id="optPaidForByBusinessNo" name="PaidForByBusinessInd" type="radio" value="0" checked><label for="optPaidForByBusinessNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span style="LEFT: 10px; POSITION: absolute; TOP: 352px" class="msgLabel">
		Are there any additional loan/liability holders?
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="optAdditionalIndYes" name="AdditionalInd" type="radio" value="1"><label for="optAdditionalIndYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 355px; POSITION: absolute; TOP: -3px">
			<input id="optAdditionalIndNo" name="AdditionalInd" type="radio" value="0" checked><label for="optAdditionalIndNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span id=lblAdditionalDetails style="LEFT: 10px; POSITION: absolute; TOP: 376px" class="msgLabel">
		Additional Loan/Liability Holder(s) Details
		<span style="LEFT: 300px; POSITION: absolute; TOP: 0px"><TEXTAREA class=msgTxt id=txtAdditionalDetails style="WIDTH: 280px; POSITION: absolute" name=AdditionalDetails rows=5 maxlength="255"></TEXTAREA>
		</span>
	</span>
	<% /* BMIDS00336 MDC 30/08/2002 */ %>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 458px" class="msgLabel">
		Credit Search
		<span style="LEFT: 230px; POSITION: relative; TOP: -3px">
			<input id="optCreditCheckYes" name="CreditCheckIndicator" type="radio" value="1"><label for="optCreditCheckYes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 355px; POSITION: absolute; TOP: -3px">
			<input id="optCreditCheckNo" name="CreditCheckIndicator" type="radio" value="0" checked><label for="optCreditCheckNo" class="msgLabel">No</label>
		</span>
	</span>			
	<span style="LEFT: 10px; POSITION: absolute; TOP: 483px" class="msgLabel">
		Unassigned
		<span style="LEFT: 240px; POSITION: relative; TOP: -3px">
			<input id="optUnassignedYes" name="UnassignedIndicator" type="radio" value="1"><label for="optUnassignedYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 355px; POSITION: absolute; TOP: -3px">
			<input id="optUnassignedNo" name="UnassignedIndicator" type="radio" value="0" checked><label for="optUnassignedNo" class="msgLabel">No</label>
		</span> 
	</span>	
	<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
</div>

<div style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 568px; HEIGHT: 304px" class="msgGroup">
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

<div style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 644px; HEIGHT: 232px" class="msgGroup"><!-- #include FILE="includes/thirdpartydetails.htm" -->
</div> 

</form><!-- Main Buttons -->
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 884px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- #include FILE="attribs/dc091attribs.asp" --><!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = null;
var m_sReadOnly = "";

<% /* BM0063  Check Processing Unit and Data Freeze indicators */ %>
var m_sProcessInd = "";
var m_sDataFreezeInd = "";

var XMLOnEntry = null;
//var m_sSequenceNumber = null;		
var scScreenFunctions;
var m_blnReadOnly = false;
<% /* BMIDS00190 */ %>
var m_iNumCustomers = 0;
<% /* BMIDS00190 End */ %>

<% /* SR 11/06/2004 : BMIDS772 */ %>
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
<% /* SR 11/06/2004 : BMIDS772 - End */ %>

var blnBalanceMandatory = false;
<% /* MAR1306 */ %>
var m_isBtnSubmit = false;

<% /* Event handler for the Another frame button. Saves the record 
	and clears all fields for new input */ %>
function btnAnother.onclick()
{
	<% /* MAR1306 Disable the button and display an hourglass until completenesscheck finishes */ %>
	btnAnother.disabled = true;
	btnAnother.blur();
	btnAnother.style.cursor = "wait";
	btnSubmit.disabled = true;
	btnSubmit.blur();
	btnSubmit.style.cursor = "wait";

	m_isBtnSubmit = false;
	<% /* MAR1306 Moved processing from here to CommitChanges & ConfirmProcessing. */ %>
	CommitChanges();
}

<% /* MAR1306 */ %>
function finishProcessing(commitChanges)
{
	if (commitChanges)
	{
		PopulateTable();
		//PopulateApplicantCombo();
		scScreenFunctions.SetFocusToFirstField(frmScreen);
		DefaultFields();

		m_sDirectoryGUID = "";
		m_sThirdPartyGUID = "";
		m_bDirectoryAddress = false;
		m_bPAFIndicator = false;
		SetAvailableFunctionality();
	}

}

function frmScreen.txtAdditionalDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true);
}

function btnCancel.onclick()
{
	frmToDC090.submit();
}

function btnSubmit.onclick()
{
	<% /* if (CommitChanges())
		frmToDC090.submit();*/ %>
	<% /* MAR1306 frmToDC090.submit() moved to CommitChanges. */ %>
	<% /* MAR1306 Disable the button and display an hourglass until completenesscheck finishes */ %>
	btnSubmit.disabled = true;
	btnSubmit.blur();
	btnSubmit.style.cursor = "wait";
	btnAnother.disabled = true;
	btnAnother.blur();
	btnAnother.style.cursor = "wait";

	m_isBtnSubmit = true;
	CommitChanges();

}

<% /* Removes the <SELECT> option and runs the Mortgage Account combo functionality */ %>
<% /* BMIDS00190 No longer required 
function frmScreen.cboApplicant.onchange()
{
	if(frmScreen.cboApplicant.value != "")
		if(frmScreen.cboApplicant.item(0).text == "<SELECT>")
			frmScreen.cboApplicant.remove(0);
}
*/ %>

function frmScreen.optAdditionalIndYes.onclick()
{
	if(frmScreen.optAdditionalIndYes.checked)
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtAdditionalDetails");
}

<% /* Sets the Additional Card Holder field to read only */ %>
function frmScreen.optAdditionalIndNo.onclick()
{
	if(frmScreen.optAdditionalIndNo.checked)
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtAdditionalDetails");
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
	<% /* BM0063 */ %>
	var bDisableButtons = false;
		
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	RetrieveContextData();
	<% /* BMIDS00190 */ %>
	scTable.initialise(tblApplicant, 0, "");

	<% /* Make the required buttons available on the bottom of the screen
		(see msgButtons.asp for details) */ %>
	var sButtonList = new Array("Submit","Cancel","Another");
	if(m_sMetaAction != "Add")
		sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	FW030SetTitles("Add/Edit Loans/Liabilities","DC091",scScreenFunctions);

	// MDC SYS2564 / SYS2785 Client Customisation
	var sGroups = new Array("Country","OrganisationType","AgreementType");
	objDerivedOperations = new DerivedScreen(sGroups);

	<%/* Parameters for thirdpartydetails.asp */%>
	m_sThirdPartyType = "3";
	m_ctrOrganisationType = frmScreen.cboOrganisationType;
	m_fValidateScreen = ValidateThirdPartyDetails;

	// MC SYS2564/SYS2785 for client customisation
	ThirdPartyCustomise();

	PopulateCombos();
	PopulateTable();
	//PopulateApplicantCombo();			
	if (m_sMetaAction == "Edit")
		PopulateScreen();
	<% /* BMIDS00336 MDC 29/08/2002 */ %>
	else
		scScreenFunctions.SetRadioGroupToReadOnly(frmScreen, "UnassignedIndicator");

	scScreenFunctions.SetRadioGroupToReadOnly(frmScreen, "CreditCheckIndicator");
	<% /* BMIDS00336 MDC 29/08/2002 - End */ %>
		
	setCreditCardTypeCombo(); <% /* SR 09/11/2005 : MAR496 */%>
	//frmScreen.cboApplicant.onchange();
	frmScreen.optAdditionalIndNo.onclick();
					
	SetThirdPartyDetailsMasks();
	SetMasks();
	Validation_Init();
	SetAvailableFunctionality();
	
	UpdateLenderNameStatus();<% /* LDM 19/7/2006 EP1003 */ %>
						
	<% /* Reset label colour to 616161 as only mandatory if radio button set to yes */ %>
	lblAdditionalDetails.style.color = "#616161";

	<% /* BM0063  Context Data is retrieved in function RetrieveContextData() 
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0"); 
	
	BM0063 Need to check Processing Unit and datafreeze indicators too. */ %>
	bDisableButtons = ((m_sReadOnly == "1") || (m_sProcessInd != "1"));
		
	if (bDisableButtons == false)
	{
		bDisableButtons = (m_sDataFreezeInd == "1");
		m_sReadOnly = m_sDataFreezeInd;
	}	

	if (bDisableButtons == true) 
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		ThirdPartyDetailsDisableButtons();
		DisableMainButton("Another");
	}
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC091");
	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function CommitChanges()
{
	var bSuccess = true;
	<% /* BM0063 Add check of Processing Unit */ %>
	if ((m_sReadOnly != "1") && (m_sProcessInd == 1))
	{

		if(frmScreen.onsubmit())
		{

			<% /* Call ConfirmProcessing after a timeout to allow the cursor time to change. */ %>
			window.setTimeout(ConfirmProcessing, 0)
		}
		else
		{
			<% /* MAR1306 Restore the mouse cursor for the buttons and reenable */ %>
			btnAnother.style.cursor = "hand";
			btnAnother.disabled = false;
			btnSubmit.style.cursor = "hand";
			btnSubmit.disabled = false;
			m_isBtnSubmit = false;
		}
	}

}

<% /* MAR1306 */ %>
function ConfirmProcessing()
{
	var bSuccess = true;

	if (IsChanged())
		if (ValidateScreen())
		{	<% /* MAR1781 GHun */ %>
			bSuccess = SaveLiability();
			
			<% /* PSC 16/05/2006 MAR1798 - Start */ %>
			if (bSuccess)
				bSuccess = RunIncomeCalculations();
			<% /* PSC 16/05/2006 MAR1798 - End */ %>
		}	<% /* MAR1781 GHun */ %>
		else
			bSuccess = false;
	
	<% /* PSC 16/05/2006 MAR1798 - Start */ %>
	if (bSuccess)
	{
		if(m_isBtnSubmit)
		{
			frmToDC090.submit();
			return ;  //SR 15/02/2007 : EP2_946 
		}
		else
			finishProcessing(bSuccess);
	}
	<% /* PSC 16/05/2006 MAR1798 - End */ %>

	<% /* MAR1306 Restore the mouse cursor for the buttons and reenable */ %>
	btnAnother.style.cursor = "hand";
	btnAnother.disabled = false;
	btnSubmit.style.cursor = "hand";
	btnSubmit.disabled = false;
	m_isBtnSubmit = false;
			
}
function DefaultFields()
{
	scScreenFunctions.ClearCollection(frmScreen);
	frmScreen.optLoanRepaymentNo.checked = true;
	frmScreen.optPaidForByBusinessNo.checked = true; <%/*MAH E2CR35 21/11/2006*/%>
	frmScreen.optAdditionalIndNo.checked = true;
	frmScreen.optAdditionalIndNo.onclick();
	ClearFields(true,true);
}

<% /* BMIDS00190 DCWP3 BM076 */ %>
function ShowRow(nIndex,sCustomerName,sSelected,sCustomerNumber,sCustomerVersionNumber)
{
	<%	// Set the table fields %>	
	scScreenFunctions.SizeTextToField(tblApplicant.rows(nIndex).cells(0),sCustomerName);
	scScreenFunctions.SizeTextToField(tblApplicant.rows(nIndex).cells(1),sSelected);
	<%	// Set the invisible context for each row %>	
	tblApplicant.rows(nIndex).setAttribute("CustomerNumber", sCustomerNumber);
	tblApplicant.rows(nIndex).setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
	tblApplicant.rows(nIndex).setAttribute("Selected", sSelected);
}

function PopulateTable()
{
	scTable.clear();
	m_iNumCustomers = 0;
	
	// Loop through all customer context entries
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);

		<% /* BMIDS00731   if (sCustomerNumber > 0) */ %>
		if (sCustomerNumber.length > 0)
		{	
			var sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + nLoop);
			var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		
			ShowRow(nLoop,sCustomerName,"No",sCustomerNumber,sCustomerVersionNumber);
			m_iNumCustomers++;
		}
	}
	
	<% /* If there is only one customer then select it by default */ %>
	if (m_iNumCustomers == 1)
	{
		scScreenFunctions.SizeTextToField(tblApplicant.rows(1).cells(1),"Yes");
		tblApplicant.rows(1).setAttribute("Selected", "Yes");	
	}
	
	frmScreen.btnSelectDeselect.disabled = true;
}
<% /* BMIDS00190 End */ %>

<% /* BMIDS00190 No longer required
function PopulateApplicantCombo()
{	*/ %>
	<% /* Clear any <OPTION> elements from the combo */ %>
<% /*	while(frmScreen.cboApplicant.options.length > 0)
	{
		frmScreen.cboApplicant.options.remove(0);
	}
	*/ %>
	<% /* Add a <SELECT> option */ %>
<% /*	var TagOPTION = document.createElement("OPTION");
	TagOPTION.value	= "";
	TagOPTION.text = "<SELECT>";
	frmScreen.cboApplicant.add(TagOPTION);

	var nCustomerCount = 0;
*/%>
	<% /* Loop through all customer context entries */ %>
<%/*	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + nLoop);
		var sCustomerNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);
*/ %>
		<% /* If the customer is an applicant, add him/her as an option */ %>
		<% /* SYS1672 - or guarantor */ %>
<% /*		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= sCustomerNumber;
			TagOPTION.text = sCustomerName;
			TagOPTION.setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
			frmScreen.cboApplicant.add(TagOPTION);
			nCustomerCount++;
		}
	}
			
	if(nCustomerCount == 1)
	{
		frmScreen.cboApplicant.remove(0);
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboApplicant");
	}
*/ %>
	<% /* Default to the first option */ %>
<% /*	frmScreen.cboApplicant.selectedIndex = 0;
} */ %>

function PopulateCombos()
// Populates all combos on the screen
{	
	PopulateTPTitleCombo();

	// MDC SYS2564 / SYS2785 Client Customisation
	var sControlList = new Array(frmScreen.cboCountry, frmScreen.cboOrganisationType, frmScreen.cboAgreementType)
	objDerivedOperations.GetComboLists(sControlList);
	
	PopulateCreditCardType(); <% /* SR 09/11/2005 : MAR496 - Populate combo CreditCardType */ %>
	
	function PopulateCreditCardType()
	{
		var blnSuccess = true;
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var sGroupList = new Array("CreditCardType");
		if(XML.GetComboLists(document, sGroupList))
		{
			var CreditCardTypeXML = XML.GetComboListXML("CreditCardType");
			var xmlNodeList = CreditCardTypeXML.selectNodes("//LISTENTRY/VALIDATIONTYPELIST");
			var xmlNode = null ;
			var xmlValidationType , xmlValidationTypeListNode ;
			var blnCCFound = false ;
			for (var iCount=0 ; iCount < xmlNodeList.length; iCount++)
			{
				xmlNode = xmlNodeList.item(iCount);
				xmlValidationTypeListNode = xmlNode.selectNodes(".//VALIDATIONTYPE");
				for(var index=0; index < xmlValidationTypeListNode.length; index++)
				{
					xmlValidationType = xmlValidationTypeListNode.item(index);
					if(xmlValidationType.text == "CC")
					{
						blnCCFound = true ;
						break ;
					}
				}
				
				if(!blnCCFound) CreditCardTypeXML.firstChild.removeChild(xmlNode.parentNode) ;
				else blnCCFound = false;
			}
			
			blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,
							frmScreen.cboCreditCardType, CreditCardTypeXML,true);
		}	
		return blnSuccess;
	}
}

function setCreditCardTypeCombo()  <%/* SR 09/11/2005 : MAR496  */ %>
{
	if(scScreenFunctions.IsValidationType(frmScreen.cboAgreementType,"CC"))
	 {
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboCreditCardType");	
     }		
	else 
	 {
		scScreenFunctions.SetFieldToDisabled(frmScreen,"cboCreditCardType");
	 }
	
	 // MAR1189 DRC	
	 if(scScreenFunctions.IsValidationType(frmScreen.cboAgreementType,"MBM"))
	 {
	    blnBalanceMandatory = true;
	    frmScreen.txtTotalOutstandingBalance.setAttribute("required","true");	    
	    frmScreen.txtTotalOutstandingBalance.parentElement.parentElement.style.color = "red";	
	 }
	 else   
	 {
	    blnBalanceMandatory = false;
	    frmScreen.txtTotalOutstandingBalance.removeAttribute("required");	    	    
	    frmScreen.txtTotalOutstandingBalance.parentElement.parentElement.style.color = "";	
	 }	
	  	 	 	 
}
		
function PopulateScreen()
{
	if(m_sMetaAction == "Edit")
	{
		var sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
		XMLOnEntry = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				
		XMLOnEntry.LoadXML(sXML);
				
		if(XMLOnEntry.SelectTag(null,"LOANSLIABILITIES") != null)
		{
			<% /* BMIDS00190 */ %>
			if (m_iNumCustomers > 1)
			{
				var CustomersXML = XMLOnEntry.ActiveTag.cloneNode(true).selectNodes("ACCOUNTRELATIONSHIP");
				for(var nCust=0; nCust < CustomersXML.length; nCust++)
				{
					sCustomerNumber = CustomersXML.item(nCust).selectSingleNode("CUSTOMERNUMBER").text;
					for(var nLoop=1; nLoop <= m_iNumCustomers; nLoop++)
					{
						if (tblApplicant.rows(nLoop).getAttribute("CUSTOMERNUMBER") == sCustomerNumber)
						{
							tblApplicant.rows(nLoop).setAttribute("Selected", "Yes");
							scScreenFunctions.SizeTextToField(tblApplicant.rows(nLoop).cells(1),"Yes");	
						}
					}
				}
			}
			<%
			//var sCustomerNumber = XMLOnEntry.GetTagText("CUSTOMERNUMBER");
			//frmScreen.cboApplicant.value = sCustomerNumber;
			%>
			<% /* BMIDS00190 End */ %>	
			
			frmScreen.cboAgreementType.value = XMLOnEntry.GetTagText("AGREEMENTTYPE");
			frmScreen.txtLoanAccountNumber.value = XMLOnEntry.GetTagText("ACCOUNTNUMBER");
			frmScreen.txtTotalOutstandingBalance.value = XMLOnEntry.GetTagText("TOTALOUTSTANDINGBALANCE");
			frmScreen.txtMonthlyRepayment.value	= XMLOnEntry.GetTagText("MONTHLYREPAYMENT");
			
			<% /* MAR102 */ %>
			frmScreen.txtOriginalLoanAmount.value = XMLOnEntry.GetTagText("ORIGINALLOANAMOUNT");
			
			<% // MAR1437 - 20/03/2006 %>
			frmScreen.cboCreditCardType.value = XMLOnEntry.GetTagText("CARDTYPE");
			
			frmScreen.txtEndDate.value = XMLOnEntry.GetTagText("ENDDATE");
			frmScreen.txtAdditionalDetails.value = XMLOnEntry.GetTagText("ADDITIONALDETAILS");
					
			scScreenFunctions.SetRadioGroupValue(frmScreen, "LoanRepaymentInd", XMLOnEntry.GetTagText("LOANREPAYMENTINDICATOR"));
			scScreenFunctions.SetRadioGroupValue(frmScreen, "PaidForByBusinessInd", XMLOnEntry.GetTagText("PAIDFORBYBUSINESS"));<%/*MAH E2CR35 21/11/2006*/%>
			
			<% /* EP2_2164 additionalindicator is not entered */ %>
			if(XMLOnEntry.GetTagText("ADDITIONALINDICATOR") != "")
			{
				scScreenFunctions.SetRadioGroupValue(frmScreen, "AdditionalInd", XMLOnEntry.GetTagText("ADDITIONALINDICATOR"));
			}

			<% /* BMIDS00190
			//frmScreen.cboApplicant.onchange();
			//m_sSequenceNumber = XMLOnEntry.GetTagText("SEQUENCENUMBER");
			*/ %>

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

			<%/* Lender details */%>
			with (frmScreen)
			{
				txtCompanyName.value = XMLOnEntry.GetTagText("COMPANYNAME");
				txtContactForename.value = XMLOnEntry.GetTagText("CONTACTFORENAME");
				txtContactSurname.value = XMLOnEntry.GetTagText("CONTACTSURNAME");
				cboTitle.value = XMLOnEntry.GetTagText("CONTACTTITLE");
				<% /* PB 07/07/2006 EP543 Begin */ %>
				checkOtherTitleField();
				<% /* AW 31/07/2006 EP1059 */ %>
				txtTitleOther.value = XMLOnEntry.GetTagText("CONTACTTITLEOTHER");
				<% /* EP543 End */ %>

				// MDC SYS2564/SYS2785 
				objDerivedOperations.LoadCounty(XMLOnEntry);
				// txtCounty.value = XMLOnEntry.GetTagText("COUNTY");
				
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
			}

			CalculateMonthlyRepayment();	<% /* BMIDS00190 */ %>

			m_sDirectoryGUID = XMLOnEntry.GetTagText("DIRECTORYGUID");
			m_bDirectoryAddress = (m_sDirectoryGUID != "");
			m_sThirdPartyGUID = XMLOnEntry.GetTagText("THIRDPARTYGUID");
			
			// 14/09/2001 JR OmiPlus 24 
			var TempXML = XMLOnEntry.ActiveTag;
			var ContactXML = XMLOnEntry.SelectTag(null, "CONTACTDETAILS");
			if(ContactXML != null)
				m_sXMLContact = ContactXML.xml;
			XMLOnEntry.ActiveTag = TempXML;
		}
	}
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	<% /* BM0063 Get Processing Unit indicator and Data Freeze Indicator */ %>
	m_sProcessInd = scScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null); 
	m_sDataFreezeInd = scScreenFunctions.GetContextParameter(window,"idFreezeDataIndicator","0");

	var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bCanAddToDirectory = (sUserRole >= XML.GetGlobalParameterAmount(document,"AddToDirectoryRole"));
	XML = null;
	
	<% /* SR 11/06/2004 : BMIDS772 */ %>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber");
	<% /* SR 11/06/2004 : BMIDS772 - End */ %>
}

<% /* Commits the screen data to the database, either by a create or update */ %>
function SaveLiability()
{
	var bOK = false;
	var TagRequestType = null;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	<% /* BMIDS00190 */ %>
	var sCustomerNumber;
	var sCustomerVersionNumber;
	var blnCustomerPassedIn;
	var DeleteXML = null;
	<% /* BMIDS00190 End */ %>
			
	if(m_sMetaAction == "Add")
		TagRequestType = XML.CreateRequestTag(window, "CREATE");
	else if(m_sMetaAction == "Edit")
		TagRequestType = XML.CreateRequestTag(window, "UPDATE");
	TagRequestType = XML.CreateRequestTag(window, null);
	XML.SetAttribute("OPERATION","CriticalDataCheck");	
	//SD MAR258 End
		
	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");

	if(TagRequestType != null)
	{
		XML.CreateActiveTag("LOANSLIABILITIES");
		
		<% /* BMIDS00190 No longer used
		var strCustomerNumber = frmScreen.cboApplicant.value;
		XML.CreateTag("CUSTOMERNUMBER", strCustomerNumber);

		var nSelectedCustomer = frmScreen.cboApplicant.selectedIndex;
		var TagOption = frmScreen.cboApplicant.options.item(nSelectedCustomer);
								
		XML.CreateTag("CUSTOMERVERSIONNUMBER",TagOption.getAttribute("CustomerVersionNumber"));

		if(m_sMetaAction == "Edit")
			XML.CreateTag("SEQUENCENUMBER",m_sSequenceNumber);
		*/ %>
		
		<% /* BMIDS00190 */ %>
		if(m_sMetaAction == "Add")
		{	
			for(var nLoop=1;nLoop <= m_iNumCustomers; nLoop++)
			{
				if (tblApplicant.rows(nLoop).getAttribute("Selected") == "Yes")
				{
					sCustomerNumber = tblApplicant.rows(nLoop).getAttribute("CustomerNumber");
					sCustomerVersionNumber = tblApplicant.rows(nLoop).getAttribute("CustomerVersionNumber");
					XML.CreateActiveTag("ACCOUNTRELATIONSHIP");
					XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
					XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);										
					XML.ActiveTag = XML.ActiveTag.parentNode;
				}
			}
		}
		else
		{
			XML.CreateTag("ACCOUNTGUID", XMLOnEntry.GetTagText("ACCOUNTGUID"));
			
			for(var nLoop=1;nLoop <= m_iNumCustomers; nLoop++)
			{
				sCustomerNumber = tblApplicant.rows(nLoop).getAttribute("CustomerNumber");
				sCustomerVersionNumber = tblApplicant.rows(nLoop).getAttribute("CustomerVersionNumber");			
				//BMIDS00832 Add in single quotes to pattern below
				if (XMLOnEntry.ActiveTag.selectSingleNode(".//CUSTOMERNUMBER[.='" + sCustomerNumber + "']") == null)
					blnCustomerPassedIn = false;
				else
					blnCustomerPassedIn = true;
				
				if (tblApplicant.rows(nLoop).getAttribute("Selected") == "Yes")
				{
					if (!blnCustomerPassedIn) <% /* Add newly selected customers */ %>
					{
						XML.CreateActiveTag("ACCOUNTRELATIONSHIP");
						XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
						XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);						
						XML.ActiveTag = XML.ActiveTag.parentNode;
					}
				}
				else	<% /* Customer was selected, but is no longer, so delete the link */ %>
				{
					if (blnCustomerPassedIn)
					{
						if (DeleteXML == null)
							DeleteXML = XML.CreateActiveTag("DELETE");
						
						XML.ActiveTag = DeleteXML;
						XML.CreateActiveTag("ACCOUNTRELATIONSHIP");
						XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
						XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
						XML.ActiveTag = XML.ActiveTag.parentNode.parentNode;
					}
				}
			}
		}
		<% /* BMIDS00190 End */ %>
		
		XML.CreateTag("ADDITIONALDETAILS",frmScreen.txtAdditionalDetails.value);
		XML.CreateTag("ADDITIONALINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"AdditionalInd"));
		XML.CreateTag("ENDDATE",frmScreen.txtEndDate.value);
		XML.CreateTag("LOANREPAYMENTINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"LoanRepaymentInd"));
		XML.CreateTag("PAIDFORBYBUSINESS",scScreenFunctions.GetRadioGroupValue(frmScreen,"PaidForByBusinessInd"));<%/*MAH E2CR35 21/11/2006*/%>
		XML.CreateTag("MONTHLYREPAYMENT",frmScreen.txtMonthlyRepayment.value);
		XML.CreateTag("TOTALOUTSTANDINGBALANCE",frmScreen.txtTotalOutstandingBalance.value);
		XML.CreateTag("AGREEMENTTYPE",frmScreen.cboAgreementType.value);
				
		<% /* MAR102 Save Original Loan Amount */ %>
		XML.CreateTag("ORIGINALLOANAMOUNT",frmScreen.txtOriginalLoanAmount.value);

		<% // MAR1437 - 20/03/2006 %>
		XML.CreateTag("CARDTYPE", frmScreen.cboCreditCardType.value);	
	
		<%/* Lender details */%>
		if (m_sMetaAction == "Edit")
		{
			// Retrieve the original third party/directory GUIDs
			var sOriginalThirdPartyGUID = XMLOnEntry.GetTagText("THIRDPARTYGUID");
			var sOriginalDirectoryGUID = XMLOnEntry.GetTagText("DIRECTORYGUID");
			// Only retrieve the address/contact details GUID if we are updating an existing third party record
			var sAddressGUID = (sOriginalThirdPartyGUID != "") ? XMLOnEntry.GetTagText("ADDRESSGUID") : "";
			var sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? XMLOnEntry.GetTagText("CONTACTDETAILSGUID") : "";
		}

		<% /* BMIDS00190 Account details*/ %>
		XML.CreateActiveTag("ACCOUNT");
		XML.CreateTag("ACCOUNTNUMBER",frmScreen.txtLoanAccountNumber.value);
		<% /* BMIDS00336 MDC 30/08/2002 */ %>
		XML.CreateTag("UNASSIGNED",scScreenFunctions.GetRadioGroupValue(frmScreen,"UnassignedIndicator"));
		<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
		<% /* BMIDS00190 End */ %>

		<% /* BMIDS000190
		// If changing from a directory to a third party, or vice-versa, the old directory/thirdparty GUID
		// should still be specified to alert the middler tier to the fact that the old link needs deleting
		//XML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID != "" ? m_sDirectoryGUID : sOriginalDirectoryGUID);
		//XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
		*/ %>
		
		<% /* BMIDS00190 */ %>
		XML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID);
		XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID);
		
		XML.ActiveTag = XML.ActiveTag.parentNode;
		<% /* BMIDS00190 End */ %>

		if (!m_bDirectoryAddress && !AllFieldsEmpty()) // Note that AllFieldsEmpty is in the ThirdPartyDetails.asp
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
		
		<% /* SR 09/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
			 node to CustomerFinancialBO. 
		*/ %>
		if(m_sMetaAction == "Add")
		{
			XML.ActiveTag = TagRequestType
			XML.CreateActiveTag("FINANCIALSUMMARY");
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
			XML.CreateTag("LOANLIABILITYINDICATOR", 1);
		}
		
			<% /* A PREVIOUS KEY section may need to be created because we may have 
				changed the customer to who the loan or liability applies*/ %>					
			<% /* BMIDS00190 no longer required
			var strOrigCustomerNumber = XMLOnEntry.GetTagText("CUSTOMERNUMBER");					
			if (strOrigCustomerNumber != strCustomerNumber)
			{						
				XML.CreateActiveTag("PREVIOUSKEY");
				XML.CreateActiveTag("LOANSLIABILITIES");
				XML.CreateTag("CUSTOMERNUMBER", strOrigCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER", XMLOnEntry.GetTagText("CUSTOMERVERSIONNUMBER"));
				XML.CreateTag("SEQUENCENUMBER", XMLOnEntry.GetTagText("SEQUENCENUMBER"));
			}
			 */ %>
			// 			XML.RunASP(document, "UpdateLiability.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			
			//Commented to fix MAR610
			/*
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "UpdateLiability.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}
			*/
		
		
		<% /* SR 09/06/2004 : BMIDS772 - End */ %>
		<% /* SR 09/11/2005 : MAR486 */ %>
		//SD CardType is not being updated, so changed this line
		//if(m_sMetaAction == "Add" && scScreenFunctions.IsValidationType(frmScreen.cboAgreementType,"CC"))
		if(scScreenFunctions.IsValidationType(frmScreen.cboAgreementType,"CC"))
		{
			XML.ActiveTag = TagRequestType;
			<% /* JD MAR1371 save bankcreditcard to the correct customers */ %>
			for(var nLoop=1;nLoop <= m_iNumCustomers; nLoop++)
			{
				if (tblApplicant.rows(nLoop).getAttribute("Selected") == "Yes")
				{
					sCustomerNumber = tblApplicant.rows(nLoop).getAttribute("CustomerNumber");
					sCustomerVersionNumber = tblApplicant.rows(nLoop).getAttribute("CustomerVersionNumber");
					XML.CreateActiveTag("BANKCREDITCARD");
					XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
					XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
					XML.CreateTag("CARDTYPE", frmScreen.cboCreditCardType.value);
					<% /* SW 29/06/2006 EP898 */ %>			
					XML.CreateTag("CARDPROVIDER", frmScreen.txtCompanyName.value);
					XML.CreateTag("TOTALOUTSTANDINGBALANCE", frmScreen.txtTotalOutstandingBalance.value);
					XML.CreateTag("AVERAGEMONTHLYREPAYMENT", frmScreen.txtMonthlyRepayment.value);
					XML.CreateTag("ADDITIONALINDICATOR", scScreenFunctions.GetRadioGroupValue(frmScreen,"AdditionalInd"));
					XML.CreateTag("ADDITIONALDETAILS", frmScreen.txtAdditionalDetails.value);		
					XML.ActiveTag = TagRequestType;
				}
			}
		} <% /* SR 09/11/2005 : MAR486 - End */ %>
		
		var sMethod = "";
				
		if(m_sMetaAction == "Add")
			sMethod = "CreateLiability.asp";	
		else
			sMethod = "UpdateLiability.asp";
				 
		// 			XML.RunASP(document, "UpdateLiability.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
							
		switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
				XML.RunASP(document, sMethod);
			break;
		default: // Error
			XML.SetErrorResponse();
		}

		bOK = XML.IsResponseOK();
	}

	return bOK;
}

function ValidateScreen()
{
	<% /* BMIDS00190 At least one customer must be selected */ %>
	var nSelected = 0;
	for(var nLoop=1; nLoop <= m_iNumCustomers; nLoop++)
	{
		if (tblApplicant.rows(nLoop).getAttribute("Selected") == "Yes")
			nSelected++;
	}
	if (nSelected == 0)
	{
		alert("At least one customer must be selected");
		tblApplicant.focus();
		return false;
	}
	<% /* BMIDS00190 End */ %>
	
	<% /* Check that the loan date isn't in the past */ %>
	if(scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtEndDate,"<") == true)
	{
		alert("Loan End Date cannot be in the past.");
		frmScreen.txtEndDate.focus();
		return(false);
	}
	
	if (!scScreenFunctions.ValidatePostcode(frmScreen.txtPostcode)) return false;

	if(scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true))
		return false;
	<%/*MAH 22/11/2006 E2CR35 Start*/%>
	if(scScreenFunctions.GetRadioGroupValue(frmScreen,"PaidForByBusinessInd") == null)
	{
		alert("Please indicate if the loan is being paid for by your business");
		frmScreen.optPaidForByBusinessNo.focus();
		return(false);
	}
	<%/*MAH 22/11/2006 E2CR35 End*/%>
	return(ValidateThirdPartyDetails());
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

<% /* BMIDS00190 DCWP3 BM076 - Start */ %>

function ToggleSelection()
{
	var iRowSelected = scTable.getRowSelected();

	if ((iRowSelected > -1) && (m_iNumCustomers > 1))
	{
		if (tblApplicant.rows(iRowSelected).getAttribute("Selected") == "Yes")
		{
			scScreenFunctions.SizeTextToField(tblApplicant.rows(iRowSelected).cells(1),"No");
			tblApplicant.rows(iRowSelected).setAttribute("Selected", "No");
		}
		else
		{
			scScreenFunctions.SizeTextToField(tblApplicant.rows(iRowSelected).cells(1),"Yes");
			tblApplicant.rows(iRowSelected).setAttribute("Selected", "Yes");
		}
		FlagChange(true);
	}
}

function frmScreen.btnSelectDeselect.onclick()
{
	ToggleSelection();
}

function spnApplicant.onclick()
{
	if ((scTable.getRowSelectedId() != null) && (m_iNumCustomers > 1))
		frmScreen.btnSelectDeselect.disabled = false;
}

function spnApplicant.ondblclick()
{
	ToggleSelection();
}

function frmScreen.cboAgreementType.onchange()
{
	CalculateMonthlyRepayment();
	setCreditCardTypeCombo(); <% /* SR 09/11/2005 : MAR496 */ %>
	UpdateLenderNameStatus(); <% /* LDM 19/07/2006 : EP1003 */ %>
}

function frmScreen.txtTotalOutstandingBalance.onchange()
{
	CalculateMonthlyRepayment();
}

function CalculateMonthlyRepayment()
{
	if (scScreenFunctions.IsValidationType(frmScreen.cboAgreementType,"CC"))
	{
		<% /* BMIDS00392 */ %>
		scScreenFunctions.SetFieldState(frmScreen, "txtMonthlyRepayment", "R"); 
		frmScreen.txtMonthlyRepayment.value = 0;
		<% /* BMIDS00392 End */ %>
		
		if (frmScreen.txtTotalOutstandingBalance.value.length > 0)
		{
			var XMLRepayment = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XMLRepayment.CreateRequestTag(window, null);
			XMLRepayment.CreateTag("TOTALOUTSTANDINGBALANCE", frmScreen.txtTotalOutstandingBalance.value);
			XMLRepayment.RunASP(document, "CalculateCreditCardRepayment.asp");
				
			if (XMLRepayment.IsResponseOK())
			{
				XMLRepayment.SelectTag(null, "RESPONSE");
				frmScreen.txtMonthlyRepayment.value = XMLRepayment.GetTagText("MONTHLYREPAYMENT");
			}
		}
	}
	else
		scScreenFunctions.SetFieldState(frmScreen, "txtMonthlyRepayment", "W"); <% /* BMIDS00392 */ %>
}
<% /* BMIDS00190 - End */ %>
<% /* PSC 16/05/2006 MAR1798 - Start */ %>
function RunIncomeCalculations()
{
  	if (m_sReadOnly =="1")
  		return(true) ;

	var AllowableIncXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AllowableIncXML.CreateRequestTag(window,null);
	AllowableIncXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));	
		
	// Set up request for Income Calculation
	var TagRequest = AllowableIncXML.CreateActiveTag("INCOMECALCULATION");
	AllowableIncXML.SetAttribute("CALCULATEMAXBORROWING","1");
	AllowableIncXML.SetAttribute("QUICKQUOTEMAXBORROWING","0");
	AllowableIncXML.CreateActiveTag("APPLICATION");
	AllowableIncXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AllowableIncXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	AllowableIncXML.ActiveTag = TagRequest;
	
	var TagCUSTOMERLIST = AllowableIncXML.CreateActiveTag("CUSTOMERLIST");

	for(var nLoop = 1; nLoop <= 2; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
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
	
	AllowableIncXML.SelectTag(null,"REQUEST");
	AllowableIncXML.SetAttribute("OPERATION","CriticalDataCheck");	
	AllowableIncXML.CreateActiveTag("CRITICALDATACONTEXT");
	AllowableIncXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	AllowableIncXML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
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

	AllowableIncXML.IsResponseOK()
	return(true);
}
<% /* PSC 16/05/2006 MAR1798 - End */ %>

<% /* LDM 19/7/2006 EP1003  
remove the mandatory requirement on lender name if the validation type of the agrement type combo is "M" eg child maintenance */ %>
function UpdateLenderNameStatus() 
{
	if(scScreenFunctions.IsValidationType(frmScreen.cboAgreementType,"NL"))
	{
		m_ctrOrganisationType.value = ""; <% /* Organisation combo */ %>
		frmScreen.txtCompanyName.value = "";
		frmScreen.txtCompanyName.setAttribute("required", null); 
		frmScreen.txtCompanyName.parentElement.parentElement.style.color = "";
	}
	else
	{
		frmScreen.txtCompanyName.setAttribute("required", "true");
		frmScreen.txtCompanyName.parentElement.parentElement.style.color = "red";
	}
}
-->
</script>
<!--  #include FILE="includes/ThirdPartyDetails.asp" -->
</body>
</html>


