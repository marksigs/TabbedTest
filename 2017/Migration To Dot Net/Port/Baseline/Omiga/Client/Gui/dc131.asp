<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<% /*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DC131.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Arrears History Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		28/09/1999	Created
AD		02/02/2000	Rework
AY		14/02/00	Change to msgButtons button types
AY		30/03/00	New top menu/scScreenFunctions change
MC		08/05/2000	SYS0679 Populate Mortgage Account combo, correct
					error message and format of amounts
BG		17/05/00	SYS0752 Removed Tooltips
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
MC		09/06/00	SYS0866 Standardise format of balance and payment fields
MH      23/06/00    SYS0933 Readonly stuff
CL		05/03/01	SYS1920 Read only functionality added
SA		06/06/01	SYS1672 Guarantors added to applicants combo
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
GHun	15/07/2002	BMIDS00190	DCWP3 BM076 support a many to many relationship between customers and arrears
GHun	29/08/2002	BMIDS00384	Add validation for LoansLiabilities and Mortgage selection
								Support multiple descriptions of loans with validation type 'O'
								Fixed problem checking validation type before combo was populated
MDC		30/08/2002	BMIDS00336	CCWP1 BM062 Credit Check & Bureau Download
GHun	04/09/2002	BMIDS00404	Disable the applicant table on screen entry if a mortgage/loan is populated
TW      09/10/2002  SYS5115		Modified to incorporate client validation
GHun	30/10/2002	BMIDS00731	Customers with alphas in the customer number are not displayed
GHun	31/10/2002	BMIDS00447	Added Last 2 years in arrears
SA		07/11/2002	BMIDS00832	DEal with alphas in customer numbers
GHun	16/11/2002	BMIDS00950	Only imported BMids mortgage accounts should be readonly
MV		25/03/2003	BM00063		AMended HTML Text for Option buttons 
HMA		16/09/2003	BM00063		Amended HTML Text for Option buttons 
SR		16/06/2004	BMIDS772	Update FinancialSummary record on Submit (only for create)	
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
INR		24/04/2006	MAR1306		Disable the button and display an hourglass onAnother & Submit to stop multiple records.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<%/* BMIDS00190 */%>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>

<%/* FORMS */%>
<form id="frmToDC130" method="post" action="dc130.asp" style="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 626px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Applicant
	</span>
<% /* BMIDS00190 DCWP3 BM076 replace combo with table
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<select id="cboApplicant" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
*/ %>
	<div id="spnApplicant" style="LEFT: 220px; POSITION: absolute; TOP: 10px">
		<table id="tblApplicant" width="300" border="0" cellspacing="0" cellpadding="0" class="msgTable">
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

	<span id="spnButtons" style="LEFT: 220px; POSITION: absolute; TOP: 103px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnSelectDeselect" value="Select/De-select" type="button" style="WIDTH: 100px" class="msgButton">
		</span>
	</span>
	
<% /* BMIDS00190 End */ %>
	</span>

	<span style="TOP: 136px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Description of Loan
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<select id="cboDescriptionOfLoan" style="WIDTH: 120px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="TOP: 162px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Account In Arrears
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<select id="cboMortgageAccount" style="WIDTH: 300px" class="msgCombo">
			</select>
		</span>
	</span>

	<% /* BMIDS00190 */ %>
	<span style="TOP: 188px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Other Arrears Account
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<input id="txtOtherArrears" type="text" maxlength="50" style="WIDTH: 300px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<% /* BMIDS00190 End */ %>

	<span style="TOP: 214px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Maximum Arrears Balance
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<input id="txtMaximumBalance" type="text" maxlength="6" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 240px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Monthly Repayment
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<input id="txtMonthlyRepayment" type="text" maxlength="9" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 266px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Maximum Number of Months in Arrears
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<input id="txtMaximumNumberOfMonths" type="text" maxlength="3" msg="Maximum Number of Arears must be greater than zero" style="WIDTH: 35px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<% /* BMIDS00336 MDC 30/08/2002 */ %>
	<span style="TOP: 292px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Current Years In Arrears
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<input id="txtCurrYearsInArrears" type="text" maxlength="12" style="WIDTH: 80px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
	
	<% /* BMIDS00447 */ %>
	<span style="TOP: 318px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Last 2 Years In Arrears
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<input id="txtLast2YearsInArrears" type="text" maxlength="24" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<% /* BMIDS00447 - End */ %>
	
	<span style="TOP: 344px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Date Cleared
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<input id="txtDateCleared" type="text" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<% /* BMIDS00336 MDC 30/08/2002 */ %>
	<span style="TOP:370px; LEFT:10px; POSITION:ABSOLUTE" class="msgLabel">
		Credit Search
		<span style="TOP:-3px; LEFT:135px; POSITION:RELATIVE">
			<input id="optCreditCheckYes" name="CreditCheckIndicator" type="radio" value="1"><label for="optCreditCheckYes" class="msgLabel">Yes</label>
		</span>
		<span style="TOP:-3px; LEFT:260px; POSITION:ABSOLUTE">
			<input id="optCreditCheckNo" name="CreditCheckIndicator" type="radio" value="0" checked><label for="optCreditCheckNo" class="msgLabel">No</label>
		</span>
	</span>		
	<span style="LEFT: 10px; POSITION: absolute; TOP: 396px" class="msgLabel">
		Unassigned
		<span style="LEFT: 145px; POSITION: RELATIVE; TOP: -3px">
			<input id="optUnassignedYes" name="UnassignedIndicator" type="radio" value="1"><label for="optUnassignedYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 260px; POSITION: absolute; TOP: -3px">
			<input id="optUnassignedNo" name="UnassignedIndicator" type="radio" value="0" checked><label for="optUnassignedNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span style="LEFT: 10px; POSITION: absolute; TOP: 422px" class="msgLabel">
		Repossession
		<span style="LEFT: 206px; POSITION: absolute; TOP: -3px">
			<input id="optRepossessionYes" name="RepossessionIndicator" type="radio" value="1"><label for="optRepossessionYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 260px; POSITION: absolute; TOP: -3px">
			<input id="optRepossessionNo" name="RepossessionIndicator" type="radio" value="0" checked><label for="optRepossessionNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
	
	<span style="TOP: 448px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Other Details
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<textarea id="txtOtherDetails" rows="5" style="WIDTH: 300px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>

	<span style="TOP: 522px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Are there any additional Arrears Holders?
		<span style="TOP: -3px; LEFT: 206px; POSITION: ABSOLUTE">
			<input id="optAdditionalIndYes" name="AdditionalInd" type="radio" value="1"><label for="optAdditionalIndYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: -3px; LEFT: 260px; POSITION: ABSOLUTE">
			<input id="optAdditionalIndNo" name="AdditionalInd" type="radio" value="0" checked><label for="optAdditionalIndNo" class="msgLabel">No</label>
		</span> 
	</span>		

	<span style="TOP: 548px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Additional Arrears Holder(s) Details
		<span style="TOP: -3px; LEFT: 210px; POSITION: ABSOLUTE">
			<textarea id="txtAdditionalDetails" rows="5" style="WIDTH: 300px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 693px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/dc131attribs.asp" -->

<!-- Specify Code Here -->
<script language="JScript">
<!--
var m_sMetaAction = null;
var XMLOnEntry = null;
var MortgageComboXML = null;
var InXML = null;
var m_sReadOnly="";
var scScreenFunctions;
var m_blnReadOnly = false;
<% /* BMIDS00190 */ %>
var m_iNumCustomers = 0;
var m_sArrearsHistoryGUID = "";
var m_sOtherArrearsAccountGUID = "";
var m_blnOtherOnEntry = false;	<% /* BMIDS00384 */ %>
var LoansComboXML = null;
<% /* var m_sSequenceNumber = null; */ %>
<% /* BMIDS00190 End */ %>
var m_blnImported = false;	<% /* BMIDS00447 */ %>
var m_blnIsBMidsMortgage = false; <% /* BMIDS00950 */ %>
<% /* SR 16/06/2004 : BMIDS772 */ %>
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
<% /* SR 16/06/2004 : BMIDS772 - End */ %>
<% /* MAR1306 */ %>
var m_isBtnSubmit = false;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	RetrieveContextData();
	
	scTable.initialise(tblApplicant, 0, "");	<% /* BMIDS00190 */ %>

	var sButtonList = new Array("Submit","Cancel","Another");

	if(m_sMetaAction != "Add")
		sButtonList = new Array("Submit","Cancel");

	ShowMainButtons(sButtonList);
	FW030SetTitles("Add/Edit Arrears History","DC131",scScreenFunctions);

	// Initial state of the Mortgage Account combo is disabled
	scScreenFunctions.SetFieldToDisabled(frmScreen,"cboMortgageAccount");
	scScreenFunctions.SetFieldToDisabled(frmScreen,"txtOtherArrears");
	frmScreen.btnSelectDeselect.disabled = true;

	GetComboLists();
	PopulateScreen();

	frmScreen.optAdditionalIndNo.onclick();

	//This function is contained in the field attributes file (remove if not required)
	SetMasks();
	Validation_Init();

	<% /* Reset label colour to 616161 as only mandatory if radio button set to yes */ %>
	frmScreen.txtAdditionalDetails.parentElement.parentElement.style.color = "#616161";

	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	<% /* BMIDS00447 Imported BMIDS accounts should be treated as read-only */ %>
	<% /* BMIDS00950 This only applies to BMids mortgage accounts 
	if (m_blnImported) */ %>
	if (m_blnImported && m_blnIsBMidsMortgage)
		m_sReadOnly = "1";
	<% /* BMIDS00447 End */ %>

	if (m_sReadOnly=="1") 
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		DisableMainButton("Another");

	}
	
	<% /* BMIDS00336 MDC 30/08/2002 */ %>
	//scScreenFunctions.SetFieldToDisabled(frmScreen,"txtCurrYearsInArrears");
	<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC131");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function frmScreen.txtAdditionalDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true);
}

function frmScreen.txtOtherDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtOtherDetails", 255, true);
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
	m_sMetaAction	= scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	<% /* SR 16/06/2004 : BMIDS772 */ %>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber");
	<% /* SR 16/06/2004 : BMIDS772 - End */ %>
}

function GetComboLists()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("ArrearsLoanType");
	if(XML.GetComboLists(document,sGroups))
		XML.PopulateCombo(document,frmScreen.cboDescriptionOfLoan,"ArrearsLoanType",true);

	GetMortgageAccountDetails(); // Doesn't populate the combo but generates the list
	GetLoansLiabilitiesDetails(); <% /* BMIDS00190 */ %>
	<% /* BMIDS00190 No longer required
	SetMortgageAccountCombo();
	PopulateApplicantCombo();
	*/ %>
	PopulateTable();
}

function GetMortgageAccountDetails()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	var TagSEARCH = XML.CreateActiveTag("SEARCH");

	XML.ActiveTag = TagSEARCH;
	var TagMORTGAGEACCOUNTLIST = XML.CreateActiveTag("MORTGAGEACCOUNTLIST");

	// Loop through all customer context entries
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		// If the customer is an applicant, add him/her to the search
		<% /* BMIDS00190 or Guarantor */ %>
		if ((sCustomerRoleType == "1") || (sCustomerRoleType == "2"))
		{
			XML.CreateActiveTag("MORTGAGEACCOUNT");
			XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			XML.ActiveTag = TagMORTGAGEACCOUNTLIST;
		}
	}
	XML.RunASP(document,"FindMortgageListForArrears.asp");

	// A record not found error is valid
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = XML.CheckResponse(sErrorArray);

	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		var ComboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateTagList("MORTGAGEACCOUNTARREARS");

		var TagMORTGAGEACCOUNTLIST = ComboXML.CreateActiveTag("MORTGAGEACCOUNTLIST");

		// Loop through the list of Mortgage Account records returned
		for(var nLoop = 0;XML.SelectTagListItem(nLoop) == true;nLoop++)
		{
			// Get all the data required for the combo entry
			// var sOrganisationName = XML.GetTagText("ORGANISATIONNAME");
			var sOrganisationName = XML.GetTagText("COMPANYNAME");
			var sAccountNumber = XML.GetTagText("ACCOUNTNUMBER");
			var sAccountGUID = XML.GetTagText("ACCOUNTGUID");
			
			// Do not generate an entry if Organisation Name AND Account Number are missing
			if(sOrganisationName != "" || sAccountNumber != "")
			{
				ComboXML.CreateActiveTag("LISTENTRY");
				ComboXML.CreateTag("VALUEID",sAccountGUID);

				// If Organisation Name or Account Number is missing, replace with suitable text
				if(sOrganisationName == "")
					sOrganisationName = "<No Organisation Details>";
				if(sAccountNumber == "")
					sAccountNumber = "<No Account Number>";

				ComboXML.CreateTag("VALUENAME",sOrganisationName + " " + sAccountNumber);
				ComboXML.CreateActiveTag("VALIDATIONTYPELIST");
				
				<% /* BMIDS00190 Add customers to ValidationType list */ %>
				var CustomersXML = XML.ActiveTag.cloneNode(true).selectNodes("ACCOUNTRELATIONSHIP");
				for(var nCust=0; nCust < CustomersXML.length; nCust++)
				{		
					sCustomerNumber = CustomersXML.item(nCust).selectSingleNode("CUSTOMERNUMBER").text;
					ComboXML.CreateTag("VALIDATIONTYPE",sCustomerNumber);
				}
				<% /* BMIDS00190 End */ %>
			
				<% /* BMIDS00950 */ %>
				if (XML.GetTagText("BMIDSACCOUNT") == "1")
					ComboXML.CreateTag("VALIDATIONTYPE","ISBMIDSACCOUNT");
				<% /* BMIDS00950 End */ %>
				
				ComboXML.ActiveTag = TagMORTGAGEACCOUNTLIST;
			}
		}
		MortgageComboXML = ComboXML.XMLDocument;
	}
}

function SetMortgageAccountCombo()
{
	if(MortgageComboXML != null)
	{
		// Populate the combo with all the records
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.PopulateComboFromXML(document,frmScreen.cboMortgageAccount,MortgageComboXML,true);
	}

	<% /* BMIDS00190 DCWP3 BM076 */ %>
<% /*	var blnCustomerSelected = false;
	var nLoop = 1;
	
	scTable.EnableTable();
	
	while(!blnCustomerSelected && (nLoop <= 5))
	{
		if (tblApplicant.rows(nLoop).getAttribute("Selected") == "Yes")
			blnCustomerSelected = true;
		nLoop++;
	}

	// If the combo should be enabled
	if(blnCustomerSelected
	   && scScreenFunctions.IsValidationType(frmScreen.cboDescriptionOfLoan,"M") == true)
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboMortgageAccount");

		if(MortgageComboXML != null)
		{
			// Populate the combo with all the records
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.PopulateComboFromXML(document,frmScreen.cboMortgageAccount,MortgageComboXML,true);
			XML = null;

			// Now remove all options which don't belong to the selected customer
			var nOption = 0;
			var sCustomerNumber;
			var blnCustomerFound;
			
			while(frmScreen.cboMortgageAccount.options.length > 0
			      && nOption < frmScreen.cboMortgageAccount.options.length)
			{
				// Make sure the <SELECT> option isn't removed.  If this is <SELECT> move
				// to the next option
				if(frmScreen.cboMortgageAccount.options.item(nOption).text != "<SELECT>")
				{
					// Remove this entry if the associated customer number does not match
					// the selected customer, otherwise move on to the next option
					
					blnCustomerFound = false;		
					var nLoop = 1;
					while (!blnCustomerFound && (nLoop <= 5))
					{
						if (tblApplicant.rows(nLoop).getAttribute("Selected") == "Yes")
						{
							sCustomerNumber = tblApplicant.rows(nLoop).getAttribute("CustomerNumber");
							if (scScreenFunctions.IsOptionValidationType(frmScreen.cboMortgageAccount,nOption,sCustomerNumber))
								blnCustomerFound = true;
						}
						nLoop++;
					}
					
					if (blnCustomerFound)
						nOption++;
					else
						frmScreen.cboMortgageAccount.options.remove(nOption);
				}
				else
					nOption++;
			} */ %>
			<% /* Disabled the combo if <SELECT> is the only option */ %>
<%/*			if (frmScreen.cboMortgageAccount.options.length == 1)
				scScreenFunctions.SetFieldToDisabled(frmScreen,"cboMortgageAccount");
		}
	}
	else
		scScreenFunctions.SetFieldToDisabled(frmScreen,"cboMortgageAccount");
	*/ %>
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

<% /* BMIDS00190 No longer used
function PopulateApplicantCombo()
{
	while(frmScreen.cboApplicant.options.length > 0)
		frmScreen.cboApplicant.options.remove(0);

	// Add a <SELECT> option
	var TagOPTION = document.createElement("OPTION");
	TagOPTION.value	= "";
	TagOPTION.text = "<SELECT>";
	frmScreen.cboApplicant.add(TagOPTION);

	var nCustomerCount = 0;

	// Loop through all customer context entries
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + nLoop);
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		// If the customer is an applicant, add him/her as an option */ %>
		<% /* SYS1672 - or guarantor */ %>
<% /*		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= sCustomerNumber;
			TagOPTION.text	= sCustomerName;
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

	// Default to the first option
	frmScreen.cboApplicant.selectedIndex = 0;
}

*/
%>

function PopulateScreen()
{	
	if(m_sMetaAction == "Edit")
	{
		var sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
		InXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

		InXML.LoadXML(sXML);
		XMLOnEntry = InXML.XMLDocument;

		if(InXML.SelectTag(null,"ARREARSHISTORY") != null)
		{
			<% /* BMIDS00190 */ %>
			if (m_iNumCustomers > 1)
			{
				var CustomersXML = InXML.ActiveTag.cloneNode(true).selectNodes("ACCOUNTRELATIONSHIP");
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
			//var sCustomerNumber = InXML.GetTagText("CUSTOMERNUMBER");
			//frmScreen.cboApplicant.value = sCustomerNumber;
			%>

			<% /* BMIDS00190 End */ %>
			
			frmScreen.txtMaximumBalance.value = InXML.GetTagText("MAXIMUMBALANCE");
			frmScreen.txtMonthlyRepayment.value = InXML.GetTagText("MONTHLYREPAYMENT");
			frmScreen.cboDescriptionOfLoan.value = InXML.GetTagText("DESCRIPTIONOFLOAN");
			frmScreen.txtMaximumNumberOfMonths.value = InXML.GetTagText("MAXIMUMNUMBEROFMONTHS");
			frmScreen.txtDateCleared.value = InXML.GetTagText("DATECLEARED");
			frmScreen.txtOtherDetails.value = InXML.GetTagText("OTHERDETAILS");
			frmScreen.txtAdditionalDetails.value = InXML.GetTagText("ADDITIONALDETAILS");
			<% /* BMIDS00447 */ %>
			frmScreen.txtCurrYearsInArrears.value = InXML.GetTagText("CURRENTYEARSINARREARS");
			frmScreen.txtLast2YearsInArrears.value = InXML.GetTagText("LASTTWOYEARSINARREARS");

			if (InXML.GetTagText("IMPORTEDINDICATOR") == '1')
				m_blnImported = true;
			<% /* BMIDS00447 End */ %>
			
			if (InXML.GetTagText("ADDITIONALINDICATOR").length > 0) <% /* BMIDS00447 only populate if there is a value */ %>
				scScreenFunctions.SetRadioGroupValue(frmScreen, "AdditionalInd", InXML.GetTagText("ADDITIONALINDICATOR"));
			//frmScreen.cboApplicant.onchange();

			if (InXML.GetTagText("REPOSSESSIONIND").length > 0)	<% /* BMIDS00447 only populate if there is a value */ %>
			<% /* BMIDS00336 MDC 02/09/2002 */ %>			
				scScreenFunctions.SetRadioGroupValue(frmScreen, "RepossessionIndicator", InXML.GetTagText("REPOSSESSIONIND"));
			<% /* BMIDS00336 MDC 02/09/2002 - End */ %>

			<% /* BMIDS00190 */ %>

			frmScreen.cboDescriptionOfLoan.onchange();	<% /* BMIDS00384 This must be called before checking the validation type below */ %>
			
			if(scScreenFunctions.IsValidationType(frmScreen.cboDescriptionOfLoan,"O"))
			{
				m_blnOtherOnEntry = true;	<% /* BMIDS00384 */ %>
				frmScreen.txtOtherArrears.value = InXML.GetTagText("DESCRIPTION");
				m_sOtherArrearsAccountGUID = InXML.GetTagText("ACCOUNTGUID");
			}
			else
			{
				frmScreen.cboMortgageAccount.value = InXML.GetTagText("ACCOUNTGUID");
				<% /* BMIDS00950 */ %>
				frmScreen.cboMortgageAccount.onchange()
				<% /* BMIDS00950 End */ %>
				<% /* BMIDS00404 Disable the applicant table */ %>
				scTable.setRowSelected(-1);
				scTable.DisableTable();
				frmScreen.btnSelectDeselect.disabled = true;
				<% /* BMIDS00404 End */ %>
			}
			
			<% /* BMIDS00190 End*/ %>

			<% /* BMIDS00336 MDC 29/08/2002 */ %>
			if(InXML.GetTagText("CREDITSEARCH") == "1")
			{
				scScreenFunctions.SetRadioGroupValue(frmScreen, "CreditCheckIndicator", "1");
				if(InXML.GetTagText("UNASSIGNED") == "1")
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
	
			//m_sSequenceNumber = InXML.GetTagText("ARREARSSEQUENCENUMBER");
			m_sArrearsHistoryGUID = InXML.GetTagText("ARREARSHISTORYGUID");
		}
	}
	<% /* BMIDS00336 MDC 29/08/2002 */ %>
	else
		scScreenFunctions.SetRadioGroupToReadOnly(frmScreen, "UnassignedIndicator");

	scScreenFunctions.SetRadioGroupToReadOnly(frmScreen, "CreditCheckIndicator");
	<% /* BMIDS00336 MDC 29/08/2002 - End */ %>
	
}

function frmScreen.optAdditionalIndYes.onclick()
{
	if(frmScreen.optAdditionalIndYes.checked)
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtAdditionalDetails");
}

function frmScreen.optAdditionalIndNo.onclick()
{
	if(frmScreen.optAdditionalIndNo.checked)
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtAdditionalDetails");
}

function frmScreen.cboDescriptionOfLoan.onchange()
{
	scTable.EnableTable();
	<% /* BMIDS00190 */ %>
	scScreenFunctions.SetFieldToDisabled(frmScreen,"txtOtherArrears");
	scScreenFunctions.SetFieldToDisabled(frmScreen,"cboMortgageAccount");
	frmScreen.cboMortgageAccount.setAttribute("required","false"); <% /* BMIDS00384 */ %>
	
	if (scScreenFunctions.IsValidationType(frmScreen.cboDescriptionOfLoan,"M"))
	{
		SetMortgageAccountCombo();
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboMortgageAccount");
		frmScreen.cboMortgageAccount.setAttribute("required","true"); <% /* BMIDS00384 */ %>
	}
	else
	{
		if (scScreenFunctions.IsValidationType(frmScreen.cboDescriptionOfLoan,"L"))
		{
			SetLoansLiabilitiesCombo();
			scScreenFunctions.SetFieldToWritable(frmScreen,"cboMortgageAccount");
			frmScreen.cboMortgageAccount.setAttribute("required","true"); <% /* BMIDS00384 */ %>
		}
		else
			if (scScreenFunctions.IsValidationType(frmScreen.cboDescriptionOfLoan,"O"))
				scScreenFunctions.SetFieldToWritable(frmScreen,"txtOtherArrears");

	}
	<% /* BMIDS00190 End */ %>
}

<% /*
function frmScreen.cboApplicant.onchange()
{
	// If the selection isn't <SELECT>, remove it
	if(frmScreen.cboApplicant.value != "")
		if(frmScreen.cboApplicant.item(0).text == "<SELECT>")
			frmScreen.cboApplicant.remove(0);
	SetMortgageAccountCombo();
}
*/ %>

function btnSubmit.onclick()
{
	<% /* MAR1306 Disable the button and display an hourglass until completenesscheck finishes */ %>
	btnSubmit.disabled = true;
	btnSubmit.blur();
	btnSubmit.style.cursor = "wait";
	btnAnother.disabled = true;
	btnAnother.blur();
	btnAnother.style.cursor = "wait";

	m_isBtnSubmit = true;

	if(frmScreen.onsubmit() && ValidateScreen())
	{
		<% /* MAR1306 Moved processing from here to CommitScreen & finishProcessing. */ %>
		<% /* Call CommitScreen after a timeout to allow the cursor time to change. */ %>
		window.setTimeout("CommitScreen()", 0)
	}
	else
	{
		btnAnother.style.cursor = "hand";
		btnAnother.disabled = false;
		btnSubmit.style.cursor = "hand";
		btnSubmit.disabled = false;

		m_isBtnSubmit = false;
	}

}

function btnCancel.onclick()
{
	frmToDC130.submit();
}

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

	if(frmScreen.onsubmit() && ValidateScreen())
	{
		<% /* MAR1306 Moved processing from here to CommitScreen & finishProcessing. */ %>
		<% /* Call CommitScreen after a timeout to allow the cursor time to change. */ %>
		window.setTimeout("CommitScreen()", 0)
	}
	else
	{
		btnAnother.style.cursor = "hand";
		btnAnother.disabled = false;
		btnSubmit.style.cursor = "hand";
		btnSubmit.disabled = false;

		m_isBtnSubmit = false;
	}

}

<% /* MAR1306 */ %>
function finishProcessing(changesOK)
{
	if (changesOK)
	{
		if(m_isBtnSubmit)
			frmToDC130.submit();
		else
		{
				PopulateTable();
				scScreenFunctions.ClearCollection(frmScreen);
				scScreenFunctions.SetFieldToDisabled(frmScreen,"cboMortgageAccount");
				frmScreen.optAdditionalIndNo.checked = true;
				frmScreen.optAdditionalIndNo.onclick();
				scScreenFunctions.SetFocusToFirstField(frmScreen);
		}
	}
	btnAnother.style.cursor = "hand";
	btnAnother.disabled = false;
	btnSubmit.style.cursor = "hand";
	btnSubmit.disabled = false;

	m_isBtnSubmit = false;

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

	if(scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true))
		return false;

	if(scScreenFunctions.RestrictLength(frmScreen, "txtOtherDetails", 255, true))
		return false;

	return true;
}

function CommitScreen()
{
	if (m_sReadOnly=="1") return true;
	
	var bOK = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var TagRequestType = null;
	<% /* BMIDS00190 */ %>
	var sCustomerNumber;
	var sCustomerVersionNumber;
	var blnCustomerPassedIn;
	var DeleteXML = null;
	var blnOtherArrearsAccounts = false;
	<% /* BMIDS00190 End */ %>

	var TagRequestType=XML.CreateRequestTag(window,null);

	<% /* BMIDS00190
	var nSelectedCustomer = frmScreen.cboApplicant.selectedIndex;
	var sCustomerNumber = frmScreen.cboApplicant.value;
	var sCustomerVersionNumber = frmScreen.cboApplicant.item(nSelectedCustomer).getAttribute("CustomerVersionNumber");
	*/ %>

	XML.CreateActiveTag("ARREARSHISTORY");
	//XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
	//XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
	XML.CreateTag("DATECLEARED", frmScreen.txtDateCleared.value);
	XML.CreateTag("DESCRIPTIONOFLOAN", frmScreen.cboDescriptionOfLoan.value);
	XML.CreateTag("MAXIMUMBALANCE", frmScreen.txtMaximumBalance.value);
	XML.CreateTag("MONTHLYREPAYMENT", frmScreen.txtMonthlyRepayment.value);
	XML.CreateTag("OTHERDETAILS", frmScreen.txtOtherDetails.value);
	XML.CreateTag("MAXIMUMNUMBEROFMONTHS", frmScreen.txtMaximumNumberOfMonths.value);
	XML.CreateTag("ADDITIONALDETAILS", frmScreen.txtAdditionalDetails.value);
	XML.CreateTag("ADDITIONALINDICATOR", scScreenFunctions.GetRadioGroupValue(frmScreen,"AdditionalInd"));
	<% /* BMIDS00447 */ %>
	XML.CreateTag("CURRENTYEARSINARREARS", frmScreen.txtCurrYearsInArrears.value);
	XML.CreateTag("LASTTWOYEARSINARREARS", frmScreen.txtLast2YearsInArrears.value);
	<% /* BMIDS00447 End */ %>

	<% /* BMIDS00190 */ %>
	if (scScreenFunctions.IsValidationType(frmScreen.cboDescriptionOfLoan,"O"))
	{
		blnOtherArrearsAccounts = true;
		XML.CreateActiveTag("OTHERARREARSACCOUNT");
		XML.CreateTag("DESCRIPTION", frmScreen.txtOtherArrears.value);
		XML.ActiveTag = XML.ActiveTag.parentNode;
	}
	else
	{
		XML.CreateTag("ACCOUNTGUID", frmScreen.cboMortgageAccount.value);
		
		if (m_sOtherArrearsAccountGUID.length > 0)
		{
			<% /* Account has changed and previous Other Arrears Account must be deleted */ %>	
			XML.CreateActiveTag("DELETE")
			XML.CreateActiveTag("OTHERARREARSACCOUNT")
			XML.CreateTag("ACCOUNTGUID", m_sOtherArrearsAccountGUID)
			XML.ActiveTag = XML.ActiveTag.parentNode.parentNode;
		}
	}
	<% /* BMIDS00190 End */ %>

	if(m_sMetaAction == "Add")
	{	
		<% /* BMIDS00190 Only save customers for Other Arrears Accounts */ %>
		if (blnOtherArrearsAccounts)
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
		
		<% /* SR 16/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
						  node to CustomerFinancialBO. */ %>
		XML.ActiveTag = TagRequestType ;
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("ARREARSHISTORYINDICATOR", 1);
		<% /* SR 16/06/2004 : BMIDS772 - End */ %>
		
		<% /* BMIDS00190 End */ %>
		// 		XML.RunASP(document,"CreateArrearsHistory.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateArrearsHistory.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	}
	else
	{	
		<% /* BMIDS00190 Only update customers for Other Arrears Accounts */ %>
		if (blnOtherArrearsAccounts)
		{
			XML.CreateTag("ACCOUNTGUID", m_sOtherArrearsAccountGUID);
			
			for(var nLoop=1;nLoop <= m_iNumCustomers; nLoop++)
			{
				sCustomerNumber = tblApplicant.rows(nLoop).getAttribute("CustomerNumber");
				sCustomerVersionNumber = tblApplicant.rows(nLoop).getAttribute("CustomerVersionNumber");			
				//BMIDS00832 add in single quotes to deal with alpha customer numbers
				if (InXML.ActiveTag.selectSingleNode(".//CUSTOMERNUMBER[.='" + sCustomerNumber + "']") == null)
					blnCustomerPassedIn = false;
				else
					blnCustomerPassedIn = true;
					
				if (tblApplicant.rows(nLoop).getAttribute("Selected") == "Yes")
				{
					<% /* Add newly selected customers, or customers that were already selected when the 
					description of loan was changed to Other Arrears Account */ %>
					<% /* BMIDS00384 */ %>
					var blnChangedToOther = ((!m_blnOtherOnEntry) && scScreenFunctions.IsValidationType(frmScreen.cboDescriptionOfLoan,"O"))
					if ((!blnCustomerPassedIn) || blnChangedToOther)
					<% /* BMIDS00384 End */ %>
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

		<% /* BMIDS00336 MDC 02/09/2002 */ %>
		XML.CreateTag("UNASSIGNED", scScreenFunctions.GetRadioGroupValue(frmScreen,"UnassignedIndicator"));
		XML.CreateTag("REPOSSESSIONIND", scScreenFunctions.GetRadioGroupValue(frmScreen,"RepossessionIndicator"));
		<% /* BMIDS00336 MDC 02/09/2002 - End */ %>
				
		XML.CreateTag("ARREARSHISTORYGUID", m_sArrearsHistoryGUID);
		<% /* BMIDS00190 End */ %>
	
		<% /*
		XML.CreateTag("ARREARSSEQUENCENUMBER",m_sSequenceNumber);
		var sOldCustomerNumber = InXML.GetTagText("CUSTOMERNUMBER");
		var sOldCustomerVersionNumber = InXML.GetTagText("CUSTOMERVERSIONNUMBER");
		
		// Compare the screen details with the original XML passed in to ascertain whether key details have been changed
		if ((sOldCustomerNumber != sCustomerNumber) | (sOldCustomerVersionNumber != sCustomerVersionNumber))
		{
			// Key fields have been changed - pass in the old key values as part of the update
			XML.CreateActiveTag("PREVIOUSKEY");
			XML.CreateActiveTag("ARREARSHISTORY");
			XML.CreateTag("CUSTOMERNUMBER", sOldCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", sOldCustomerVersionNumber);
			XML.CreateTag("ARREARSSEQUENCENUMBER", m_sSequenceNumber);
		}
		*/ %>
		
		// 		XML.RunASP(document,"UpdateArrearsHistory.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateArrearsHistory.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	}
	bOK = XML.IsResponseOK();

	<% /* MAR1306 return bOK; */ %>
	finishProcessing(bOK);
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
		//SetMortgageAccountCombo();
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

function frmScreen.cboMortgageAccount.onchange()
{
	var sCustomerNumer;
	var sSelected;
	
	if (frmScreen.cboMortgageAccount.value.length > 0)
	{	
		<% /* Select customers linked to the mortgage account and disable the table */ %>
		for(var nLoop = 1; nLoop <= m_iNumCustomers; nLoop++)
		{
			sCustomerNumber = tblApplicant.rows(nLoop).getAttribute("CustomerNumber");
			if (sCustomerNumber != null)
			{
				if (scScreenFunctions.IsOptionValidationType(frmScreen.cboMortgageAccount,frmScreen.cboMortgageAccount.selectedIndex,sCustomerNumber))
					sSelected = "Yes";
				else
					sSelected = "No";
			
				tblApplicant.rows(nLoop).setAttribute("Selected", sSelected);
				scScreenFunctions.SizeTextToField(tblApplicant.rows(nLoop).cells(1), sSelected);
			}
		}
		
		scTable.setRowSelected(-1);
		scTable.DisableTable();
		frmScreen.btnSelectDeselect.disabled = true;
		
		<% /* BMIDS00950 */ %>
		if (scScreenFunctions.IsValidationType(frmScreen.cboDescriptionOfLoan,"M"))
		{
			if (scScreenFunctions.IsOptionValidationType(frmScreen.cboMortgageAccount,frmScreen.cboMortgageAccount.selectedIndex,"ISBMIDSACCOUNT"))
				m_blnIsBMidsMortgage = true;
			else
				m_blnIsBMidsMortgage = false;
		}
		<% /* BMIDS00950 End */ %>
	}
	else
		scTable.EnableTable();
}

function GetLoansLiabilitiesDetails()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	var TagSEARCH = XML.CreateActiveTag("SEARCH");

	XML.ActiveTag = TagSEARCH;
	var TagMORTGAGEACCOUNTLIST = XML.CreateActiveTag("LOANSLIABILITIESLIST");

	// Loop through all customer context entries
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		// If the customer is an applicant or guarantor, add him/her to the search
		if ((sCustomerRoleType == "1") || (sCustomerRoleType == "2"))
		{
			XML.CreateActiveTag("LOANSLIABILITIES");
			XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			XML.ActiveTag = TagMORTGAGEACCOUNTLIST;
		}
	}
	XML.RunASP(document,"FindLoansListForArrears.asp");

	// A record not found error is valid
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = XML.CheckResponse(sErrorArray);

	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		var ComboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateTagList("LOANSLIABILITIESARREARS");

		var TagMORTGAGEACCOUNTLIST = ComboXML.CreateActiveTag("LOANSLIABILITIESLIST");

		// Loop through the list of Mortgage Account records returned
		for(var nLoop = 0;XML.SelectTagListItem(nLoop) == true;nLoop++)
		{
			// Get all the data required for the combo entry
			// var sOrganisationName = XML.GetTagText("ORGANISATIONNAME");
			var sOrganisationName = XML.GetTagText("COMPANYNAME");
			var sAccountNumber = XML.GetTagText("ACCOUNTNUMBER");
			var sAccountGUID = XML.GetTagText("ACCOUNTGUID");

			// Do not generate an entry if Organisation Name AND Account Number are missing
			if(sOrganisationName != "" || sAccountNumber != "")
			{
				ComboXML.CreateActiveTag("LISTENTRY");
				ComboXML.CreateTag("VALUEID",sAccountGUID);

				// If Organisation Name or Account Number is missing, replace with suitable text
				if(sOrganisationName == "")
					sOrganisationName = "<No Organisation Details>";
				if(sAccountNumber == "")
					sAccountNumber = "<No Account Number>";

				ComboXML.CreateTag("VALUENAME",sOrganisationName + " " + sAccountNumber);
				ComboXML.CreateActiveTag("VALIDATIONTYPELIST");
				
				<% /* Add customers to ValidationType list */ %>
				var CustomersXML = XML.ActiveTag.cloneNode(true).selectNodes("ACCOUNTRELATIONSHIP");
				for(var nCust=0; nCust < CustomersXML.length; nCust++)
				{		
					sCustomerNumber = CustomersXML.item(nCust).selectSingleNode("CUSTOMERNUMBER").text;
					ComboXML.CreateTag("VALIDATIONTYPE",sCustomerNumber);
				}
				
				ComboXML.ActiveTag = TagMORTGAGEACCOUNTLIST;
			}
		}
		LoansComboXML = ComboXML.XMLDocument;
	}
}

function SetLoansLiabilitiesCombo()
{
	if(LoansComboXML != null)
	{
		// Populate the combo with all the records
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.PopulateComboFromXML(document,frmScreen.cboMortgageAccount,LoansComboXML,true);
	}
}

<% /* BMIDS00190 - End */ %>
-->
</script>
</body>
</html>
