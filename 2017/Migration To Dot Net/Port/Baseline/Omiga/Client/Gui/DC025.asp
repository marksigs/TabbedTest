<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/* 
Workfile:      DC025.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Add/Edit Group connection details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		17/11/99	Created
JLD		17/12/1999	SYS0073 - error message given on text in additional details > 255
					SYS0071 - Default radio if only one applicant
JLD		01/02/2000	rework
AY		14/02/00	Change to msgButtons button types
IW		08/03/00	SYS0095 - Readonly operation corected
AD		21/03/00	SYS0411 - Focus set to account type when only one applicant
AY		30/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
CL		05/03/01	SYS1920 Read only functionality added
SA		23/05/01	SYS2187 Commitchanges() was not being called when adding new group connection
SR		02/08/01	SYS2531 - handle read-only processing in function CommitChanges; call this 
					function in all situations. 					
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MV		22/08/2002	BMIDS00355  IE 5.5 Upgrade Errors , amended the positions 
HMA     17/09/2003  BM0063      Amend HTML text for radio buttons
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
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
<!-- Specify Forms Here -->
<form id="frmToDC020" method="post" action="DC020.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 250px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Group Connection Details</strong>
	</span>
	<span style="TOP: 36px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Connection Owner
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<select id="cboApplicantName" name="ApplicantName" style="WIDTH: 200px" class="msgCombo"></select>
		</span>
	</span>
	<span style="TOP: 62px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Account Type
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<select id="cboAccountType" name="AccountType" style="WIDTH: 100px" class="msgCombo"></select>
		</span>
	</span>
	<span style="TOP: 98px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Account Number
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtAccountNumber" name="AccountNumber" maxlength="20" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span style="TOP: 124px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Are there any Additional Account Holders on this account?
		<span style="TOP: -3px; LEFT: 300px; POSITION: ABSOLUTE">
			<input id="optOthersResponsibleYes" name="OthersResponsibleGroup" type="radio" value="1"><label for="optOthersResponsibleYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: -3px; LEFT: 360px; POSITION: ABSOLUTE">
			<input id="optOthersResponsibleNo" name="OthersResponsibleGroup" type="radio" value="0"><label for="optOthersResponsibleNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span style="TOP: 150px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Additional Account Holder Details
		<span style="TOP: 16px; LEFT: 0px; POSITION: ABSOLUTE">
			<textarea id="txtAdditionalDetails" name="AdditionalDetails" rows="5" style="WIDTH: 300px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>
</div>
</form>
<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 320px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<!-- File containing field attributes (remove if not required) -->
<!--  #include FILE="attribs/DC025attribs.asp" -->
<!-- Specify Code Here -->
<script language="JScript">
<!--
var m_sMetaAction		= ""; 
var m_sReadOnly			= "";
var m_sXML				= "";
var editXML						= null;
var m_nApplicantNameIndex		= -1;
var	m_nNumOfApplicants			= 0;
var scScreenFunctions;
var m_blnReadOnly = false;



<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel","Another");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Add/Edit Group Connection Details","DC025",scScreenFunctions);

	RetrieveContextData();
	SetMasks();
	Validation_Init();
	Initialise(true);
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC025")
	
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
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	if(m_sMetaAction == "Edit")	m_sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
}		
function DefaultFields()
{
<%  /* reset field values */
%>	if(m_nNumOfApplicants == 1)
	{
		frmScreen.cboApplicantName.selectedIndex = 1;
		frmScreen.optOthersResponsibleNo.checked = true;
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAdditionalDetails");
	}
	else
	{
		frmScreen.cboApplicantName.selectedIndex = 0;
		frmScreen.optOthersResponsibleYes.checked = false;
		frmScreen.optOthersResponsibleNo.checked = false;
	}
	frmScreen.cboAccountType.selectedIndex = 0;
	frmScreen.txtAccountNumber.value = "";
	frmScreen.txtAdditionalDetails.value = "";
}
function Initialise(bOnLoad)
{
	if(bOnLoad == true)PopulateCombos();
	if(m_sMetaAction == "Edit")
	{
		PopulateScreen();
		DisableMainButton("Another");
	}
	else DefaultFields();
	if(frmScreen.optOthersResponsibleYes.checked == false)
	{
		frmScreen.txtAdditionalDetails.disabled = true;
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAdditionalDetails");
	}

	if(m_nNumOfApplicants == 1)
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboApplicantName");

	if(m_sReadOnly == true)
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
}
function PopulateCombos()
{
	var XMLAccountType	= null;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
<%	/* populate Applicant name from the customers in context. Add a <SELECT> option */
%>	var TagOPTION	= document.createElement("OPTION");
	TagOPTION.value	= "";
	TagOPTION.text	= "<SELECT>";
	frmScreen.cboApplicantName.add(TagOPTION);
	m_nNumOfApplicants = 0;
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName			= scScreenFunctions.GetContextParameter(window,"idCustomerName" + nLoop);
		var sCustomerNumber			= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		if(sCustomerName != "" && sCustomerNumber != "")
		{
			m_nNumOfApplicants++;
			TagOPTION		= document.createElement("OPTION");
			TagOPTION.value	= sCustomerNumber;
			TagOPTION.text	= sCustomerName;
			TagOPTION.setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
			frmScreen.cboApplicantName.add(TagOPTION);
		}
	}
<%	// Default to SELECT or the only option if there is only one
%>	if(m_nNumOfApplicants == 1)frmScreen.cboApplicantName.selectedIndex = 1;
	else frmScreen.cboApplicantName.selectedIndex = 0;
<%	// populate AccountType from the database
%>	var sGroupList = new Array("GroupConnectionAccountType");
	if(XML.GetComboLists(document, sGroupList) == true)
	{
		XMLAccountType = XML.GetComboListXML("GroupConnectionAccountType");
		var blnSuccess = XML.PopulateComboFromXML(document, frmScreen.cboAccountType,XMLAccountType,true);
		if(blnSuccess == false)	scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
	frmScreen.cboAccountType.selectedIndex = 0;
	XML = null;
}
function PopulateScreen()
{
<%	//we are in edit mode. Load the groupconnection XML from context
%>	editXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	editXML.LoadXML(m_sXML);
	editXML.SelectTag(null,"GROUPCONNECTION");
	frmScreen.cboApplicantName.value = editXML.GetTagText("CUSTOMERNUMBER");
	frmScreen.cboAccountType.value = editXML.GetTagText("ACCOUNTTYPE");
	frmScreen.txtAccountNumber.value = editXML.GetTagText("ACCOUNTNUMBER");
	var sOthersResponsibleYesNo = editXML.GetTagText("ADDITIONALINDICATOR");
	if(sOthersResponsibleYesNo == "0")frmScreen.optOthersResponsibleNo.checked = true;
	if(sOthersResponsibleYesNo == "1")
	{
		frmScreen.optOthersResponsibleYes.checked = true;
		frmScreen.txtAdditionalDetails.value = editXML.GetTagText("ADDITIONALDETAILS");
	}
<%	//save the chosen index of applicant name
%>	m_nApplicantNameIndex = frmScreen.cboApplicantName.selectedIndex;
}
function WriteGroupConnection(bInsertMode)
{
<% /* Creates new group connection record.
		Args Passed:	bInsertMode - true if we are inserting a new record
						false if we are updating an existing record
		Returns:		false if an error occurred*/
%>	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var tagREQUEST = XML.CreateRequestTag(window, "SAVE");
	var tagGROUPCONNECTION = XML.CreateActiveTag("GROUPCONNECTION");
	XML.CreateTag("CUSTOMERNUMBER", frmScreen.cboApplicantName.value);
	var TagOption = frmScreen.cboApplicantName.options.item(frmScreen.cboApplicantName.selectedIndex);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", TagOption.getAttribute("CustomerVersionNumber"));
<%	// in update mode we need the GroupConnectionSequenceNumber which we can get
	// from the idXML xml
%>	if(bInsertMode == false) XML.CreateTag("GROUPCONNECTIONSEQUENCENUMBER", editXML.GetTagText("GROUPCONNECTIONSEQUENCENUMBER"));
	XML.CreateTag("ACCOUNTTYPE", frmScreen.cboAccountType.value);
	XML.CreateTag("ACCOUNTNUMBER", frmScreen.txtAccountNumber.value);
	var sAdditionalDetails = "";
	if(frmScreen.optOthersResponsibleYes.checked == true)
	{
		XML.CreateTag("ADDITIONALINDICATOR", "1");
		sAdditionalDetails = frmScreen.txtAdditionalDetails.value;
	}
	else if(frmScreen.optOthersResponsibleNo.checked == true)XML.CreateTag("ADDITIONALINDICATOR", "0");
	else XML.CreateTag("ADDITIONALINDICATOR", "");
	XML.CreateTag("ADDITIONALDETAILS", sAdditionalDetails);
	// 	XML.RunASP(document, "SaveGroupConnection.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "SaveGroupConnection.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if (XML.IsResponseOK() == false)bSuccess = false;
	XML = null;
	return(bSuccess);
}
function DeleteGroupConnection()
{
<%  /* Only called in edit mode. Generate XML to send to delete procedure  */
%>	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var tagREQUEST = XML.CreateRequestTag(window, "SAVE");
	var tagGROUPCONNECTION = XML.CreateActiveTag("GROUPCONNECTION");
	XML.CreateTag("CUSTOMERNUMBER", editXML.GetTagText("CUSTOMERNUMBER"));
	XML.CreateTag("CUSTOMERVERSIONNUMBER", editXML.GetTagText("CUSTOMERVERSIONNUMBER"));
	XML.CreateTag("GROUPCONNECTIONSEQUENCENUMBER", editXML.GetTagText("GROUPCONNECTIONSEQUENCENUMBER"));
	// 	XML.RunASP(document, "SaveGroupConnection.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "SaveGroupConnection.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if (XML.IsResponseOK() == false)bSuccess = false;
	XML = null;
	return(bSuccess);
}
function CommitChanges()
{
	var bSuccess = false;
	if(m_sReadOnly == "1") bSuccess = true ;
	else
	{
		if ((frmScreen.onsubmit() == true))
		{
			bSuccess = true;
			if(scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true))
				bSuccess = false;
			if (IsChanged() == true && bSuccess == true)
			{
				var bInsertMode = false;
				if( m_sMetaAction == "Edit" && m_nApplicantNameIndex != frmScreen.cboApplicantName.selectedIndex )
				{
					<% /*delete this groupconnection record*/%>		
					bSuccess = DeleteGroupConnection();
					bInsertMode = true;
				}
				if(m_sMetaAction == "Add")bInsertMode = true;
				if(bSuccess == true)bSuccess = WriteGroupConnection(bInsertMode);
			}
		}
	}
	return(bSuccess);
}
function frmScreen.optOthersResponsibleYes.onclick()
{
	frmScreen.txtAdditionalDetails.disabled = false;
	scScreenFunctions.SetFieldToWritable(frmScreen, "txtAdditionalDetails");
}
function frmScreen.optOthersResponsibleNo.onclick()
{
	frmScreen.txtAdditionalDetails.value = "";
	frmScreen.txtAdditionalDetails.disabled = true;
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAdditionalDetails");
}
function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
	scScreenFunctions.SetContextParameter(window,"idXML", "");
	frmToDC020.submit();
}
function btnSubmit.onclick()
{
	<%
	// SArn 23/5/01 SYS2187. Even though it wasn't read only, the commitchanges
	// function was not being called. Added some more brackets!
	//if( m_sReadOnly ||CommitChanges() == true )
	%>
	<% /* SR 02/08/01 : SYS2531 - handle read-only processing in Commit Changes function  */ %>
	if( (CommitChanges() == true))
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		scScreenFunctions.SetContextParameter(window,"idXML", "");
		frmToDC020.submit();
	}
}
function btnAnother.onclick()
{
	if( CommitChanges() == true)Initialise(false);
}

function frmScreen.txtAdditionalDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true);
}

-->
</script>
</body>
</html>


