<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<% /*
Workfile:      DC121.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Declined Mortgage Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		28/09/1999	Created
AD		02/02/2000	Rework
AY		14/02/00	Change to msgButtons button types
IW		23/02/00	SYS0140 - additional applicants text changed.
IW		23/02/00	SYS0140 - Toggling of Mandatory text colours defered pending tech meet
AY		30/03/00	New top menu/scScreenFunctions change
IW		03/05/00	SYS0140 - Fixed [3]
MC		16/05/00	SYS0733 - Fixed problems with textareas > 255 chars.
BG		17/05/00	SYS0752 Removed Tooltips
MC		05/06/00	SYS0820 - Fixed data loss in Additional Details
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
CL		05/03/01	SYS1920 Read only functionality added
SA		06/06/01	SYS1672 Guarantors added to applicants combo					
LD		23/05/02	SYS4727 Use cached versions of frame functions
TW      09/10/2002  SYS5115 Modified to incorporate client validation
HMA     17/09/2003  BM0063  Amend HTML text for radio buttons
SR		16/06/2004	BMIDS772 Update FinancialSummary record on Submit (only for create)	
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>

<%/* FORMS */%>
<form id="frmToDC120" method="post" action="dc120.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 260px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Declined Mortgage Owner
		<span style="TOP: -3px; LEFT: 190px; POSITION: ABSOLUTE">
			<select id="cboApplicant" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="TOP: 36px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Date Declined
		<span style="TOP: -3px; LEFT: 190px; POSITION: ABSOLUTE">
			<input id="txtDateDeclined" type="text" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 62px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Details
		<span style="TOP: 0px; LEFT: 190px; POSITION: ABSOLUTE">
			<textarea id="txtMortgageDeclinedDetails" rows="5" style="WIDTH: 300px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>
	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span style="TOP: 150px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel" >
		Are there any Additional Applicants?
		<span style="TOP: -3px; LEFT: 186px; POSITION: ABSOLUTE">
			<input id="optAdditionalIndYes" name="AdditionalInd" type="radio" value="1"><label for="optAdditionalIndYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: -3px; LEFT: 240px; POSITION: ABSOLUTE">
			<input id="optAdditionalIndNo" name="AdditionalInd" type="radio" value="0" checked><label for="optAdditionalIndNo" class="msgLabel">No</label>
		</span> 
	</span>		

	<span style="TOP: 176px; LEFT: 10px; POSITION: ABSOLUTE"  class="msgLabel">
		Additional Applicant(s) Details
		<span style="TOP: 0px; LEFT: 190px; POSITION: ABSOLUTE">
			<textarea id="txtAdditionalDetails" rows="5" style="WIDTH: 300px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/dc121attribs.asp" -->

<!-- Specify Code Here -->
<script language="JScript">
<!--
var m_sMetaAction = null;
var XMLOnEntry = null;
var m_sSequenceNumber = null;
var InXML = null;
var scScreenFunctions;
var m_blnReadOnly = false;
<% /* SR 14/06/2004 : BMIDS772 */ %>
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
<% /* SR 14/06/2004 : BMIDS772 - End */ %>

function frmScreen.txtMortgageDeclinedDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtMortgageDeclinedDetails", 255, true);
}

function frmScreen.txtAdditionalDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true);
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

	var sButtonList = new Array("Submit","Cancel","Another");

	// If not in add mode then the another button is not required
	if(m_sMetaAction != "Add")
		sButtonList = new Array("Submit","Cancel");

	ShowMainButtons(sButtonList);
	FW030SetTitles("Add/Edit Declined Mortgage Details","DC121",scScreenFunctions);

	GetComboLists();
	PopulateScreen();

	frmScreen.cboApplicant.onchange();

	//This function is contained in the field attributes file (remove if not required)
	SetMasks();

	Validation_Init();
	
	<% /* Reset label colour to 616161 as only mandatory if radio button set to yes */ %>
	frmScreen.txtAdditionalDetails.parentElement.parentElement.style.color = "#616161";
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC121");
	
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
	<% /* SR 14/06/2004 : BMIDS772 */ %>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber");
	<% /* SR 14/06/2004 : BMIDS772 - End */ %>
}

function GetComboLists()
{
	PopulateApplicantCombo();
}

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
		<% /* SYS1672 - or guarantor */ %>
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
	// Default to the first option
	frmScreen.cboApplicant.selectedIndex = 0;
}

function frmScreen.cboApplicant.onchange()
{
	// If the selection isn't <SELECT>, remove it
	if(frmScreen.cboApplicant.value != "")
		if(frmScreen.cboApplicant.item(0).text == "<SELECT>")
			frmScreen.cboApplicant.remove(0);
}

function PopulateScreen()
{
	if(m_sMetaAction == "Edit")
	{
		var sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
		InXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

		InXML.LoadXML(sXML);
		XMLOnEntry = InXML.XMLDocument;

		if(InXML.SelectTag(null,"DECLINEDMORTGAGE") != null)
		{
			var sCustomerNumber = InXML.GetTagText("CUSTOMERNUMBER");
			frmScreen.cboApplicant.value = sCustomerNumber;
			frmScreen.txtDateDeclined.value = InXML.GetTagText("DATEDECLINED");
			frmScreen.txtMortgageDeclinedDetails.value = InXML.GetTagText("DECLINEDDETAILS");
			frmScreen.txtAdditionalDetails.value = InXML.GetTagText("ADDITIONALDETAILS");

			scScreenFunctions.SetRadioGroupValue(frmScreen, "AdditionalInd", InXML.GetTagText("ADDITIONALINDICATOR"));
			if(scScreenFunctions.GetRadioGroupValue(frmScreen, "AdditionalInd") == "0")
				frmScreen.optAdditionalIndNo.onclick();
		
			m_sSequenceNumber = InXML.GetTagText("SEQUENCENUMBER");
		}
	}
	else if(m_sMetaAction == "Add")
		frmScreen.optAdditionalIndNo.onclick();
	
}

function frmScreen.optAdditionalIndYes.onclick()
{
	if(frmScreen.optAdditionalIndYes.checked)
		frmScreen.txtAdditionalDetails.setAttribute("required", "true");
		<% /*SYS0140 lblAdditionalDetails.runtimeStyle.color = 'red'*/ %>
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtAdditionalDetails");

}

function frmScreen.optAdditionalIndNo.onclick()
{
	if(frmScreen.optAdditionalIndNo.checked)
		frmScreen.txtAdditionalDetails.setAttribute("required", "false");
		<% /*SYS0140 lblAdditionalDetails.runtimeStyle.color = 'darkblue'*/ %>
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtAdditionalDetails");
}

function btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
		if(ValidateScreen())
			if(CommitScreen())
				frmToDC120.submit();
}

function btnCancel.onclick()
{
	frmToDC120.submit();
}

function btnAnother.onclick()
{
	if(frmScreen.onsubmit())
		if(ValidateScreen())
			if(CommitScreen())
			{
				PopulateApplicantCombo();
				scScreenFunctions.ClearCollection(frmScreen);
				frmScreen.optAdditionalIndNo.checked = true;
				frmScreen.optAdditionalIndNo.onclick();
				scScreenFunctions.SetFocusToFirstField(frmScreen);
			}
}

function ValidateScreen()
{
	if(scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtDateDeclined,">"))
	{
		alert("Date Declined cannot be in the future");
		frmScreen.txtDateDeclined.focus();
		return false;
	}

	if(scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true))
		return false;

	if(scScreenFunctions.RestrictLength(frmScreen, "txtMortgageDeclinedDetails", 255, true))
		return false;

	return true;
}

function CommitScreen()
{
	var bOK = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var TagRequestType = null;

	var nSelectedCustomer = frmScreen.cboApplicant.selectedIndex;
	var sCustomerNumber = frmScreen.cboApplicant.value;
	var sCustomerVersionNumber = frmScreen.cboApplicant.item(nSelectedCustomer).getAttribute("CustomerVersionNumber");
	
	var TagRequestType = XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("DECLINEDMORTGAGE");
	XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
	XML.CreateTag("ADDITIONALDETAILS",frmScreen.txtAdditionalDetails.value);
	XML.CreateTag("ADDITIONALINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"AdditionalInd"));
	XML.CreateTag("DATEDECLINED",frmScreen.txtDateDeclined.value);
	XML.CreateTag("DECLINEDDETAILS",frmScreen.txtMortgageDeclinedDetails.value);
	
	<% /* SR 14/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
						  node to CustomerFinancialBO. */ %>
	if(m_sMetaAction == "Add")
	{
		XML.ActiveTag = TagRequestType ;
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("DECLINEDMORTGAGEINDICATOR", 1);
	}
	<% /* SR 14/06/2004 : BMIDS772 - End */ %>

	if(m_sMetaAction == "Add")
		// 		XML.RunASP(document,"CreateDeclinedMortgage.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateDeclinedMortgage.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	else
	{
		XML.CreateTag("SEQUENCENUMBER",m_sSequenceNumber);
		var sOldCustomerNumber = InXML.GetTagText("CUSTOMERNUMBER");
		var sOldCustomerVersionNumber = InXML.GetTagText("CUSTOMERVERSIONNUMBER");

		// Compare the screen details with the original XML passed in to ascertain whether key details have been changed
		if ((sOldCustomerNumber != sCustomerNumber) | (sOldCustomerVersionNumber != sCustomerVersionNumber))
		{
			// Key fields have been changed - pass in the old key values as part of the update
			XML.CreateActiveTag("PREVIOUSKEY");
			XML.CreateActiveTag("DECLINEDMORTGAGE");
			XML.CreateTag("CUSTOMERNUMBER", sOldCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", sOldCustomerVersionNumber);
			XML.CreateTag("SEQUENCENUMBER", m_sSequenceNumber);
		}

		// 		XML.RunASP(document,"UpdateDeclinedMortgage.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateDeclinedMortgage.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	}
	bOK = XML.IsResponseOK();

	return bOK;
}
-->
</script>
</body>
</html>


