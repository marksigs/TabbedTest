<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<% /*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DC151.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   CCJ History Details
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
APS		11/07/00	Removed reference to scTable.asp
CL		05/03/01	SYS1920 Read only functionality added
SA		06/06/01	SYS1672 Guarantors added to applicants combo
LD		23/05/02	SYS4727 Use cached versions of frame functions


BMIDS Specific History:

Prog	Date		AQR			Description
GHun	15/07/2002	BMIDS00190	DCWP3 BM076 support a many to many relationship between customers and CCJ history
MDC		02/09/2002	BMIDS00336	Credit Check & Bureau Download
GHun	30/10/2002	BMIDS00731	Customers with alphas in the customer number are not displayed
TW      09/10/2002  SYS5115		Modified to incorporate client validation - 
MV		31/10/2002	BMIDS00355	Modified HTML
SA		07/11/2002	BMIDS00832	Deal with alpha customer numbers
HMA     17/09/2003  BM0063      Amend HTML text for radio buttons
SR		16/06/2004	BMIDS772	Update FinancialSummary record on Submit (only for create)	
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<% /* BMIDS00190 */ %>
<object data="scTable.htm" height="1" id="scTable" 
style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex="-1" type="text/x-scriptlet" 
viewastext></object>

<%/* FORMS */%>
<form id="frmToDC150" method="post" action="dc150.asp" style="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" year4 validate="onchange" mark>
<div id="divBackground" style="HEIGHT: 555px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Applicant
	</span>
	<% /* BMIDS00190 DCWP3 BM076 Replace combo with table
	<span style="TOP: -3px; LEFT: 220px; POSITION: ABSOLUTE">
		<select id="cboApplicant" name="Applicant" style="WIDTH: 200px" class="msgCombo">
		</select>
	</span>
	*/ %>

	<div id="spnApplicant" style="LEFT: 230px; POSITION: absolute; TOP: 10px">
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
	
	<span id="spnButtons" style="LEFT: 230px; POSITION: absolute; TOP: 103px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnSelectDeselect" value="Select/De-select" type="button" style="WIDTH: 100px" class="msgButton">
		</span>
	</span>

	<% /* BMIDS00190 End */ %>

	<% /* BMIDS00336 MDC 02/09/2002 */ %>
	<span style="TOP: 136px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Type
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px" class="msgLabel">
			<select id="cboCCJType" name="CCJType" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	</span>
	<% /* BMIDS00336 MDC 02/09/2002 - End */ %>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 162px" class="msgLabel">
		Value of Judgement
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="txtValueOfJudgement" name="ValueOfJudgement" maxlength="6" style="POSITION: absolute; WIDTH: 65px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 188px" class="msgLabel">
		Monthly Repayment
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="txtMonthlyRepayment" name="MonthlyRepayment" maxlength="9" style="POSITION: absolute; WIDTH: 65px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 214px" class="msgLabel">
		Date of Judgement
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="txtDateOfJudgement" name="DateOfJudgement" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 240px" class="msgLabel">
		Plaintiff
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="txtPlaintiff" name="Plaintiff" maxlength="50" style="POSITION: absolute; WIDTH: 200px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 266px" class="msgLabel">
		Date Cleared
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="txtDateCleared" name="DateCleared" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span>
	</span>

	<% /* BMIDS00336 MDC 30/08/2002 */ %>
	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span style="TOP:292px; LEFT:10px; POSITION:ABSOLUTE" class="msgLabel">
		Credit Search
		<span style="TOP:-3px; LEFT:150px; POSITION:RELATIVE">
			<input id="optCreditCheckYes" name="CreditCheckIndicator" type="radio" value="1"><label for="optCreditCheckYes" class="msgLabel">Yes</label>
		</span>
		<span style="TOP:-3px; LEFT:160px; POSITION:RELATIVE">
			<input id="optCreditCheckNo" name="CreditCheckIndicator" type="radio" value="0" checked><label for="optCreditCheckNo" class="msgLabel">No</label>
		</span>
	</span>		
	<span style="LEFT: 10px; POSITION: absolute; TOP: 318px" class="msgLabel">
		Unassigned
		<span style="LEFT: 158px; POSITION: RELATIVE; TOP: -3px">
			<input id="optUnassignedYes" name="UnassignedIndicator" type="radio" value="1"><label for="optUnassignedYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 168px; POSITION: RELATIVE; TOP: -3px">
			<input id="optUnassignedNo" name="UnassignedIndicator" type="radio" value="0" checked><label for="optUnassignedNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span style="LEFT: 10px; POSITION: absolute; TOP: 344px" class="msgLabel">
		Default Record
		<span style="LEFT: 218px; POSITION: absolute; TOP: -3px">
			<input id="optDefaultYes" name="DefaultIndicator" type="radio" value="1"><label for="optDefaultYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 272px; POSITION: absolute; TOP: -3px">
			<input id="optDefaultNo" name="DefaultIndicator" type="radio" value="0" checked><label for="optDefaultNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<% /* BMIDS00336 MDC 30/08/2002 - End */ %>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 370px" class="msgLabel">
		Other Details
		<span style="LEFT: 220px; POSITION: absolute; TOP: 0px"><TEXTAREA class=msgTxt id=txtOtherDetails name=OtherDetails rows=5 style="POSITION: absolute; WIDTH: 300px"></TEXTAREA>
		</span>
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 450px" class="msgLabel">
		Are there any additional CCJ Holders?
		<span style="LEFT: 216px; POSITION: absolute; TOP: -3px">
			<input id="optAdditionalIndYes" name="AdditionalInd" type="radio" value="1"><label for="optAdditionalIndYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 270px; POSITION: absolute; TOP: -3px">
			<input id="optAdditionalIndNo" name="AdditionalInd" type="radio" value="0" checked><label for="optAdditionalIndNo" class="msgLabel">No</label>
		</span> 
	</span>		

	<span style="LEFT: 10px; POSITION: absolute; TOP: 476px" class="msgLabel">
		Additional CCJ Holder(s) Details
		<span style="LEFT: 220px; POSITION: absolute; TOP: 0px"><textarea class=msgTxt id=txtAdditionalDetails name=AdditionalDetails rows=5 style="POSITION: absolute; WIDTH: 300px"></textarea>
		</span>
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 625px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>
<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/dc151attribs.asp" -->
<!-- Specify Code Here -->
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = null;
var XMLOnEntry = null;	
var InXML = null;
var scScreenFunctions;
var m_sReadOnly="";
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
	<% /* BMIDS00190 */ %>
	scTable.initialise(tblApplicant, 0, "");

	var sButtonList = new Array("Submit","Cancel","Another");

	// If not in add mode then the another button is not required
	if(m_sMetaAction != "Add")
		sButtonList = new Array("Submit","Cancel");

	ShowMainButtons(sButtonList);
	FW030SetTitles("Add/Edit CCJ History","DC151",scScreenFunctions);

	GetComboLists();
	PopulateScreen();

	<% /* BMIDS00190
	frmScreen.cboApplicant.onchange();
	*/ %>
	
	frmScreen.optAdditionalIndNo.onclick();

	//This function is contained in the field attributes file (remove if not required)
	SetMasks();

	Validation_Init();

	<% /* Reset label colour to 616161 as only mandatory if radio button set to yes */ %>
	frmScreen.txtAdditionalDetails.parentElement.parentElement.style.color = "616161";
	
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");

	if (m_sReadOnly=="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		DisableMainButton("Another");
	}
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC151");
	
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

function frmScreen.txtAdditionalDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true);
}
function frmScreen.txtOtherDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtOtherDetails", 255, true);
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	<% /* SR 16/06/2004 : BMIDS772 */ %>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber");
	<% /* SR 16/06/2004 : BMIDS772 - End */ %>
}

function GetComboLists()
{
	<% /* BMIDS00336 MDC 02/09/2002 */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("CCJType");
	if(XML.GetComboLists(document,sGroups))
		XML.PopulateCombo(document,frmScreen.cboCCJType,"CCJType",true);
	<% /* BMIDS00336 MDC 02/09/2002  - End */ %>

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

		// If the customer is an applicant, add him/her as an option
		*/ %>
		<% /* SYS1672 - or guarantor */ %>
		<% /* 
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
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

	frmScreen.cboApplicant.selectedIndex = 0;
}


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
	var sCustomernumber;
	
	if(m_sMetaAction == "Edit")
	{
		var sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
		InXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

		InXML.LoadXML(sXML);
		XMLOnEntry = InXML.XMLDocument;

		if(InXML.SelectTag(null,"CCJHISTORY") != null)
		{
			<% /* BMIDS00190 */ %>
			if (m_iNumCustomers > 1)
			{
				var CustomersXML = InXML.ActiveTag.cloneNode(true).selectNodes("CUSTOMERVERSIONCCJHISTORY");
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
			frmScreen.txtValueOfJudgement.value = InXML.GetTagText("VALUEOFJUDGEMENT");
			frmScreen.txtMonthlyRepayment.value = InXML.GetTagText("MONTHLYREPAYMENT");
			frmScreen.txtDateOfJudgement.value = InXML.GetTagText("DATEOFJUDGEMENT");
			frmScreen.txtPlaintiff.value = InXML.GetTagText("PLAINTIFF");
			frmScreen.txtDateCleared.value = InXML.GetTagText("DATECLEARED");
			frmScreen.txtOtherDetails.value = InXML.GetTagText("OTHERDETAILS");
			frmScreen.txtAdditionalDetails.value = InXML.GetTagText("ADDITIONALDETAILS");
			scScreenFunctions.SetRadioGroupValue(frmScreen, "AdditionalInd", InXML.GetTagText("ADDITIONALINDICATOR"));

			<% /* BMIDS00336 MDC 02/09/2002 */ %>
			frmScreen.cboCCJType.value = InXML.GetTagText("CCJTYPE");
			scScreenFunctions.SetRadioGroupValue(frmScreen, "DefaultIndicator", InXML.GetTagText("DEFAULTRECORD"));
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

function btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
		if(ValidateFields())
			if(CommitScreen())
				frmToDC150.submit();
}

function btnCancel.onclick()
{
	frmToDC150.submit();
}

function btnAnother.onclick()
{
	if(frmScreen.onsubmit())
		if(ValidateFields())
			if(CommitScreen())
			{
				PopulateTable();
				scScreenFunctions.ClearCollection(frmScreen);
				frmScreen.optAdditionalIndNo.checked = true;
				frmScreen.optAdditionalIndNo.onclick();
				scScreenFunctions.SetFocusToFirstField(frmScreen);
			}
}

function ValidateFields()
{
	var bOK = true;

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
		bOK = false;
	}
	<% /* BMIDS00190 End */ %>

	if(scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtDateOfJudgement,">"))
	{
		alert("Date Of Judgement cannot be in the future");
		frmScreen.txtDateOfJudgement.focus();
		bOK = false;
	}
	else if(scScreenFunctions.CompareDateFields(frmScreen.txtDateOfJudgement,">",frmScreen.txtDateCleared))
	{
		alert("Date Cleared cannot be before Date Of Judgement");
		frmScreen.txtDateCleared.focus();
		bOK = false;
	}
	else if((frmScreen.txtValueOfJudgement.value != "") & (parseFloat(frmScreen.txtValueOfJudgement.value) == 0))
	{
		alert("Value of Judgment must be greater than 0");
		frmScreen.txtValueOfJudgement.focus();
		bOK = false;
	}
	else if(scScreenFunctions.RestrictLength(frmScreen, "txtOtherDetails", 255, true))
		bOK = false;
	else if(scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true))
		bOK = false;
		
	return bOK;
}

function CommitScreen()
{
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
	
	XML.CreateActiveTag("CCJHISTORY");
	XML.CreateTag("ADDITIONALDETAILS",frmScreen.txtAdditionalDetails.value);
	XML.CreateTag("ADDITIONALINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"AdditionalInd"));
	XML.CreateTag("DATECLEARED",frmScreen.txtDateCleared.value);
	XML.CreateTag("DATEOFJUDGEMENT",frmScreen.txtDateOfJudgement.value);
	XML.CreateTag("MONTHLYREPAYMENT",frmScreen.txtMonthlyRepayment.value);
	XML.CreateTag("OTHERDETAILS",frmScreen.txtOtherDetails.value);
	XML.CreateTag("PLAINTIFF",frmScreen.txtPlaintiff.value);
	XML.CreateTag("VALUEOFJUDGEMENT",frmScreen.txtValueOfJudgement.value);

	<% /* BMIDS00336 MDC 02/09/2002 */ %>	
	XML.CreateTag("UNASSIGNED", scScreenFunctions.GetRadioGroupValue(frmScreen,"UnassignedIndicator"));
	XML.CreateTag("DEFAULTRECORD", scScreenFunctions.GetRadioGroupValue(frmScreen,"DefaultIndicator"));
	XML.CreateTag("CCJTYPE", frmScreen.cboCCJType.value);
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
				XML.CreateActiveTag("CUSTOMERVERSIONCCJHISTORY");
				XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
				XML.ActiveTag = XML.ActiveTag.parentNode;
			}
		}			
		<% /* SR 16/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
						  node to CustomerFinancialBO. */ %>
		XML.ActiveTag = TagRequestType ;
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("CCJHISTORYINDICATOR", 1);
		<% /* SR 16/06/2004 : BMIDS772 - End */ %>

		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateCCJHistory.asp");
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
					XML.CreateActiveTag("CUSTOMERVERSIONCCJHISTORY");
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
					{	DeleteXML = XML.CreateActiveTag("DELETE");
					}
					
					XML.ActiveTag = DeleteXML;
					XML.CreateActiveTag("CUSTOMERVERSIONCCJHISTORY");
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
			XML.CreateActiveTag("CCJHISTORY");
			XML.CreateTag("CUSTOMERNUMBER", sOldCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", sOldCustomerVersionNumber);
			XML.CreateTag("SEQUENCENUMBER", m_sSequenceNumber);
		}
		*/ %>
		XML.CreateTag("CCJHISTORYGUID", InXML.GetTagText("CCJHISTORYGUID"));
	
		// 		XML.RunASP(document,"UpdateCCJHistory.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateCCJHistory.asp");
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
