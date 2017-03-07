<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      DC141.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Bankruptcy History Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		28/09/1999	Created
AD		03/02/2000	Rework
AY		14/02/00	Change to msgButtons button types
AY		30/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
MC		09/06/00	SYS0866 Standardise format of balance and payment fields
MH      23/06/00    SYS0933 ReadOnly stuff
CL		05/03/01	SYS1920 Read only functionality added
SA		06/06/01	SYS1672 Guarantors added to applicants combo
LD		23/05/02	SYS4727 Use cached versions of frame functions


BMIDS Specific History:

Prog	Date		AQR			Description
GHun	15/07/2002	BMIDS00190	DCWP3 BM076 support a many to many relationship between customers and bankruptcy
MDC		02/09/2002	BMIDS00336	Credit Check & Bureau Download
TW      09/10/2002    Modified to incorporate client validation - SYS5115
GHun	30/10/2002	BMIDS00731	Customers with alphas in the customer number are not displayed
SA		07/11/2002	BMIDS00832	Deal with alpha customer numbers in pattern matching
MV		25/03/2003	BM0063		AMended HTML Text for Option buttons 
MV		08/04/2003	BM0063		AMended HTML Text for Option buttons 
HMA     16/09/2003  BM0063      Amended HTML Text for Option buttons
SR		16/06/2004	BMIDS772	Update FinancialSummary record on Submit (only for create)	
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
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

<form id="frmToDC140" method="post" action="dc140.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 397px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Applicant
	</span>
	<% /* BMIDS00190 DCWP3 BM076 Replace combo with table
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<select id="cboApplicant" name="Applicant" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	*/ %>
	<span id="spnApplicant" style="LEFT: 130px; POSITION: absolute; TOP: 10px">
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
	</span>
	
	<span id="spnButtons" style="LEFT: 130px; POSITION: absolute; TOP: 103px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnSelectDeselect" value="Select/De-select" type="button" style="WIDTH: 100px" class="msgButton">
		</span>
	</span>
	<% /* BMIDS00190 End */ %>
		
	<span style="TOP: 136px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Amount of Debt
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtAmountOfDebt" name="AmountOfDebt" maxlength="6" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 162px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Monthly Repayment
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtMonthlyRepayment" name="MonthlyRepayment" maxlength="9" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 188px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Date Declared
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtDateDeclared" name="DateDeclared" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 214px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Date of Discharge
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtDateOfDischarge" name="DateOfDischarge" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<% /* BMIDS00336 MDC 30/08/2002 */ %>
	<span style="TOP:240px; LEFT:10px; POSITION:ABSOLUTE" class="msgLabel">
		Credit Search
		<span style="TOP:-3px; LEFT:50px; POSITION:RELATIVE">
			<input id="optCreditCheckYes" name="CreditCheckIndicator" type="radio" value="1"><label for="optCreditCheckYes" class="msgLabel">Yes</label>
		</span>
		<span style="TOP:-3px; LEFT:180px; POSITION:ABSOLUTE">
			<input id="optCreditCheckNo" name="CreditCheckIndicator" type="radio" value="0" checked><label for="optCreditCheckNo" class="msgLabel">No</label>
		</span>
	</span>		
	<span style="LEFT: 10px; POSITION: absolute; TOP: 266px" class="msgLabel">
		Unassigned
		<span style="LEFT: 60px; POSITION: RELATIVE; TOP: -3px">
			<input id="optUnassignedYes" name="UnassignedIndicator" type="radio" value="1"><label for="optUnassignedYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="optUnassignedNo" name="UnassignedIndicator" type="radio" value="0" checked><label for="optUnassignedNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span style="LEFT: 10px; POSITION: absolute; TOP: 292px" class="msgLabel">
		IVA
		<span style="LEFT: 100px; POSITION: RELATIVE; TOP: -3px">
			<input id="optIVAYes" name="IVAIndicator" type="radio" value="1"><label for="optIVAYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 115px; POSITION: RELATIVE; TOP: -3px">
			<input id="optIVANo" name="IVAIndicator" type="radio" value="0" checked><label for="optIVANo" class="msgLabel">No</label>
		</span> 
	</span>		
	<% /* BMIDS00336 MDC 30/08/2002 - End */ %>

	<span style="TOP: 318px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Other Details
		<span style="TOP: 0px; LEFT: 120px; POSITION: ABSOLUTE">
			<textarea id="txtOtherDetails" name="OtherDetails" rows="5" style="WIDTH: 300px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 480px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<!-- File containing field attributes (remove if not required) -->
<!--  #include FILE="attribs/dc141attribs.asp" -->

<!-- Specify Code Here -->
<script language="JScript">
<!--
var m_sMetaAction = null;
var InXML = null;
var scScreenFunctions;
var m_sReadOnly ="";
var m_blnReadOnly = false;
var m_iNumCustomers = 0;

<% /* SR 16/06/2004 : BMIDS772 */ %>
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
<% /* SR 16/06/2004 : BMIDS772 - End */ %>

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

	// If not in add mode then the another button is not required
	if(m_sMetaAction != "Add")
		sButtonList = new Array("Submit","Cancel");

	ShowMainButtons(sButtonList);
	FW030SetTitles("Add/Edit Bankruptcy History","DC141",scScreenFunctions);

	GetComboLists();
	PopulateScreen();
	
	<% /* BMIDS00190
	frmScreen.cboApplicant.onchange();
	*/ %>

	//This function is contained in the field attributes file (remove if not required)
	SetMasks();

	Validation_Init();
	
	if (m_sReadOnly=="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		DisableMainButton("Another");
	}
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC141");
	
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

function frmScreen.txtOtherDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtOtherDetails", 255, true);
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	
	<% /* SR 16/06/2004 : BMIDS772 */ %>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber");
	<% /* SR 16/06/2004 : BMIDS772 - End */ %>
}

function GetComboLists()
{
	PopulateTable();
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
	// Clear any <OPTION> elements from the combo
	while(frmScreen.cboApplicant.options.length > 0)
		frmScreen.cboApplicant.options.remove(0);

	var TagOPTION	= document.createElement("OPTION");
	TagOPTION.value	= "";
	TagOPTION.text	= "<SELECT>";
	frmScreen.cboApplicant.add(TagOPTION);

	var nCustomerCount = 0;

	// Loop through all customer context entries
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + nLoop);
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		// If the customer is an applicant, add him/her as an option
		/* %>
		<% /* SYS1672 - or guarantor */ %>
		<% /*
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
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
*/ %>

<% /* BMIDS00190
function frmScreen.cboApplicant.onchange()
{
	// If the selection isn't <SELECT>, remove it
	if(frmScreen.cboApplicant.value != "")
		if(frmScreen.cboApplicant.item(0).text == "<SELECT>")
			frmScreen.cboApplicant.remove(0);
}
*/ %>

function PopulateScreen()
{
	var sCustomerNumber;
	
	if(m_sMetaAction == "Edit")
	{
		var sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
		InXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

		InXML.LoadXML(sXML);

		if(InXML.SelectTag(null,"BANKRUPTCYHISTORY") != null)
		{
			<% /* BMIDS00190 */ %>
			if (m_iNumCustomers > 1)
			{
				var CustomersXML = InXML.ActiveTag.cloneNode(true).selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY");
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
			<% /* BMIDS00190 End */ %>
			frmScreen.txtAmountOfDebt.value = InXML.GetTagText("AMOUNTOFDEBT");
			frmScreen.txtMonthlyRepayment.value = InXML.GetTagText("MONTHLYREPAYMENT");
			frmScreen.txtDateDeclared.value = InXML.GetTagText("DATEDECLARED");
			frmScreen.txtDateOfDischarge.value = InXML.GetTagText("DATEOFDISCHARGE");
			frmScreen.txtOtherDetails.value = InXML.GetTagText("OTHERDETAILS");

			<% /* BMIDS00336 MDC 02/09/2002 */ %>
			scScreenFunctions.SetRadioGroupValue(frmScreen, "IVAIndicator", InXML.GetTagText("IVA"));
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
			<% /* BMIDS00336 MDC 02/09/2002 - End */ %>
			
		}
	}
	<% /* BMIDS00336 MDC 29/08/2002 */ %>
	else
		scScreenFunctions.SetRadioGroupToReadOnly(frmScreen, "UnassignedIndicator");

	scScreenFunctions.SetRadioGroupToReadOnly(frmScreen, "CreditCheckIndicator");
	<% /* BMIDS00336 MDC 29/08/2002 - End */ %>
	
}

function btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
		if(ValidateScreen())
			if(CommitScreen())
				frmToDC140.submit();
}

function btnCancel.onclick()
{
	frmToDC140.submit();
}

function btnAnother.onclick()
{
	if(frmScreen.onsubmit())
		if(ValidateScreen())
			if(CommitScreen())
			{
				PopulateTable();
				scScreenFunctions.ClearCollection(frmScreen);
				scScreenFunctions.SetFocusToFirstField(frmScreen);
			}
}

function ValidateScreen()
{
	if (m_sReadOnly=="1") return true;
	
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
	
	if(scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtDateDeclared,">"))
	{
		alert("Date Declared cannot be in the future");
		frmScreen.txtDateDeclared.focus();
		return false;
	}
	if(scScreenFunctions.CompareDateFields(frmScreen.txtDateDeclared,">",frmScreen.txtDateOfDischarge))
	{
		alert("Date Of Discharge cannot be before Date Declared");
		frmScreen.txtDateOfDischarge.focus();
		return false;
	}
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
	var sCustomer;
	var sCustomerNumber;
	var blnCustomerPassedIn;
	var DeleteXML = null;
	<% /* BMIDS00190 End */ %>

	var TagRequestType = XML.CreateRequestTag(window,null)
	
	XML.CreateActiveTag("BANKRUPTCYHISTORY");
	XML.CreateTag("AMOUNTOFDEBT",frmScreen.txtAmountOfDebt.value);
	XML.CreateTag("DATEDECLARED",frmScreen.txtDateDeclared.value);
	XML.CreateTag("DATEOFDISCHARGE",frmScreen.txtDateOfDischarge.value);
	XML.CreateTag("MONTHLYREPAYMENT",frmScreen.txtMonthlyRepayment.value);
	XML.CreateTag("OTHERDETAILS",frmScreen.txtOtherDetails.value);

	<% /* BMIDS00336 MDC 02/09/2002 */ %>	
	XML.CreateTag("UNASSIGNED", scScreenFunctions.GetRadioGroupValue(frmScreen,"UnassignedIndicator"));
	XML.CreateTag("IVA", scScreenFunctions.GetRadioGroupValue(frmScreen,"IVAIndicator"));
	<% /* BMIDS00336 MDC 02/09/2002 - End */ %>	
	
	<% /* BMIDS00190 */ %>
	if(m_sMetaAction == "Add")
	{
		for(var nLoop=1;nLoop <= m_iNumCustomers; nLoop++)
		{
			if (tblApplicant.rows(nLoop).getAttribute("Selected") == "Yes")
			{
				sCustomerNumber = tblApplicant.rows(nLoop).getAttribute("CustomerNumber");
				sCustomerVersionNumber = tblApplicant.rows(nLoop).getAttribute("CustomerVersionNumber");
				XML.CreateActiveTag("CUSTOMERVERSIONBANKRUPTCYHISTORY");
				XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
				XML.ActiveTag = XML.ActiveTag.parentNode;
			}
		}	
		<% /* SR 15/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
						  node to CustomerFinancialBO. */ %>
		XML.ActiveTag = TagRequestType ;
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("BANKRUPTCYHISTORYINDICATOR", 1);
		<% /* SR 15/06/2004 : BMIDS772 - End */ %>
		
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateBankruptcyHistory.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	}
	else
	{
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
				if (!blnCustomerPassedIn) <% /* Add newly selected customers */ %>
				{
					XML.CreateActiveTag("CUSTOMERVERSIONBANKRUPTCYHISTORY");
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
					XML.CreateActiveTag("CUSTOMERVERSIONBANKRUPTCYHISTORY");
					XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
					XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
					XML.ActiveTag = XML.ActiveTag.parentNode.parentNode;
				}
			}
		}
		
		<% /*
		XML.CreateTag("SEQUENCENUMBER",m_sSequenceNumber);
		var sOldCustomerNumber = InXML.GetTagText("CUSTOMERNUMBER");
		var sOldCustomerVersionNumber = InXML.GetTagText("CUSTOMERVERSIONNUMBER");
		*/ %>
		// Compare the screen details with the original XML passed in to ascertain whether key details have been changed
		<% /*
		if ((sOldCustomerNumber != sCustomerNumber) | (sOldCustomerVersionNumber != sCustomerVersionNumber))
		{
			// Key fields have been changed - pass in the old key values as part of the update
			XML.CreateActiveTag("PREVIOUSKEY");
			XML.CreateActiveTag("BANKRUPTCYHISTORY");
			XML.CreateTag("CUSTOMERNUMBER", sOldCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", sOldCustomerVersionNumber);
			XML.CreateTag("SEQUENCENUMBER", m_sSequenceNumber);
		}
		*/ %>

		XML.CreateTag("BANKRUPTCYHISTORYGUID", InXML.GetTagText("BANKRUPTCYHISTORYGUID"));

		// 		XML.RunASP(document,"UpdateBankruptcyHistory.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateBankruptcyHistory.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	}
	<% /* BMIDS00190 End */ %>
	bOK = XML.IsResponseOK();

	return bOK;
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

<% /* BMIDS00190 - End */ %>
-->
</script>
</body>
</html>


