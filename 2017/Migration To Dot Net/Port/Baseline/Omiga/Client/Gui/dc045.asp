<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DC045.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Add/Edit Group connection details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		01/12/99	Created
JLD		10/12/1999	DC/023 - on OK/Another, don't route if fails.
					DC/029 and 032 - attribs change
JLD		16/12/1999	SYS0078 - removed title and other fields
					SYS0074 - get error if additional details text exceed 255 chars
					SYS0072 - default additional applicants.
JLD		17/12/1999	SYS0079 - check validity of date of birth field
JLD		01/02/2000	rework fr performance
AY		14/02/00	Change to msgButtons button types
APS		23/02/00	Change to disabled processing on Additional Applicant Relatives which
					now works correctly on editing existing dependent
IW		15/03/00	SYS0499 [1] & [2]					
AY		30/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
MH      22/06/00    SYS0933 ReadOnluy
MC		07/07/00	SYS1168 Allow Dependant Relative Of field to be updated
CL		05/03/01	SYS1920 Read only functionality added
TJ		30/03/01	SYS2050 Critical Data functionality added
GD		11/05/01	SYS2050 Critical Data functionality ROLLED BACK
SA		24/05/01	SYS0499 DOB not mandatory - only call CalculateCustomerAge if dob is input
SA		24/05/01	SYS0994 Increase length of surname to 40 and forenames to 30
							Disable cboApplicantsName if only one.
							Added CheckName function - checking for first char of names being numeric.
DC      20/07/01    SYS2038 Critical Data functionality ROLLED FORWARD AGAIN							
LD		23/05/02	SYS4727 Use cached versions of frame functions
TW      09/10/02    SYS5115 Modified to incorporate client validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :

Prog	Date		Description
GHun	22/07/2005	MAR14 Apply ING Style Sheet and GUI Images
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title></title>
</head>
<body><!-- Form Manager Object - Controls Soft Coded Field Attributes -->

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<!-- Specify Forms Here -->
<form id="frmToDC040" method="post" action="DC040.asp" style="DISPLAY: none"></form><!-- File containing field attributes (remove if not required) --><!-- #include FILE="attribs/dc045attribs.asp" -->
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="HEIGHT: 370px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Dependant Details</strong>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Dependant relative of:
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboApplicantName" name="ApplicantName" style="WIDTH: 210px" class="msgCombo"></select>
		</span> 
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Forenames
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<input id="txtFirstForename" name="FirstForename" maxlength="30" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
		</span> 
		<span style="LEFT: 230px; POSITION: absolute; TOP: -3px">
			<input id="txtSecondForename" name="SecondForename" maxlength="30" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
		</span> 
		<span style="LEFT: 340px; POSITION: absolute; TOP: -3px">
			<input id="txtOtherForenames" name="OtherForenames" maxlength="30" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 94px" class="msgLabel">
		Surname
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtSurname" name="Surname" maxlength="40" style="POSITION: absolute; WIDTH: 100px" class="msgTxt"> 
		</span> 
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 118px" class="msgLabel">
		Date of Birth
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtDateOfBirth" name="DateOfBirth" maxlength="20" style="POSITION: absolute; WIDTH: 100px" class="msgTxt"> 
		</span> 
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 142px" class="msgLabel">
		Sex
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboGender" name="Gender" style="WIDTH: 100px" class="msgCombo"></select>
		</span> 
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 166px" class="msgLabel">
		Dependant is your:
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboRelationship" name="Relationship" style="WIDTH: 100px" class="msgCombo"></select>
		</span> 
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 190px" class="msgLabel">
		Resident at new Address?
		<span style="LEFT: 310px; POSITION: absolute; TOP: -3px">
			<input id="optResidentAtNewAddressYes" name="ResidentAtNewAddressGroup" type="radio" value="1">
			<label for="optResidentAtNewAddressYes" class="msgLabel">Yes</label> 
		</span> 
		<span style="LEFT: 360px; POSITION: absolute; TOP: -3px">
			<input id="optResidentAtNewAddressNo" name="ResidentAtNewAddressGroup" type="radio" value="0">
			<label for="optResidentAtNewAddressNo" class="msgLabel">No</label> 
		</span> 
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 214px" class="msgLabel">
		Are there any Additional Applicants the Dependant is related to?
		<span style="LEFT: 310px; POSITION: absolute; TOP: -3px">
			<input id="optAdditionalApplicantsYes" name="AdditionalApplicantsGroup" type="radio" value="1">
			<label for="optAdditionalApplicantsYes" class="msgLabel">Yes</label> 
		</span> 
		<span style="LEFT: 360px; POSITION: absolute; TOP: -3px">
			<input id="optAdditionalApplicantsNo" name="AdditionalApplicantsGroup" type="radio" value="0">
			<label for="optAdditionalApplicantsNo" class="msgLabel">No</label> 
		</span> 
	</span>		
	<span id=lblAdditionalRelatives style="LEFT: 4px; POSITION: absolute; TOP: 238px" class="msgLabel">
		Additional Applicant Relatives
		<span style="LEFT: 0px; POSITION: absolute; TOP: 16px"><TEXTAREA class=msgTxt id=txtAdditionalRelatives name=AdditionalRelatives rows=5 style="POSITION: absolute; WIDTH: 300px"></TEXTAREA> 
		</span> 
	</span> 
</div> 
</form><!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 440px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div> 
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- Specify Code Here -->
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction			= "";
var m_sReadOnly				= "";
var m_sXML					= "";
var editXML					= null;
var m_nApplicantNameIndex	= -1;
var	m_nNumOfApplicants		= 0;
var m_sApplicationNumber		   = "";
var m_sApplicationFactFindNumber   = "";
var m_sPersonGUID				   = "";
var m_sOtherResidentSequenceNumber = "";
var scScreenFunctions;
var m_blnReadOnly = false;


function btnAnother.onclick()
{
	if (CommitChanges())Initialise(false);		
}				
function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
	scScreenFunctions.SetContextParameter(window,"idXML", "");
	frmToDC040.submit();
}
function btnSubmit.onclick()
{
	
	if (CommitChanges())
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		scScreenFunctions.SetContextParameter(window,"idXML", "");
		frmToDC040.submit();			
	}
}
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}
function frmScreen.optAdditionalApplicantsYes.onclick()
{
	frmScreen.txtAdditionalRelatives.disabled = false;
	scScreenFunctions.SetFieldToWritable(frmScreen, "txtAdditionalRelatives");
}
function frmScreen.optAdditionalApplicantsNo.onclick()
{
<%	// No additional relatives should be specified if the 'No' option button is selected.
%>	scScreenFunctions.SetFieldToDisabled(frmScreen, "txtAdditionalRelatives");
}

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
	FW030SetTitles("Add/Edit Dependant Details","DC045",scScreenFunctions);
	
	RetrieveContextData();
	SetMasks();
	Validation_Init();
	Initialise(true);

	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");

	if (m_sReadOnly=="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		DisableMainButton("Another");
	}
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC045");
	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}
function CalculateCustomerAge()
{
<%  /* Description:	sets the Age field based on the date of birth field
       Returns:		age in years, or -1 if field is invalid */
%>	var nAge = -1;
	var dteBirthdate = scScreenFunctions.GetDateObject(frmScreen.txtDateOfBirth);
	if(dteBirthdate != null)
	{
		<% /* MO - BMIDS00807 */ %>
		var dteToday = scScreenFunctions.GetAppServerDate(true);
		<% /* var dteToday = new Date(); */ %>
		nAge = top.frames[1].document.all.scMathFunctions.GetYearsBetweenDates(dteBirthdate, dteToday);
	}
	return(nAge);
}
function CommitChanges()
{
	var bSuccess = true;
	if (m_sReadOnly != "1")
	{ 
		if(frmScreen.onsubmit() == true)			
		{
			if (IsChanged())
			{
				<% /* SYS0994 First name + surname should not contain numerics in first letter */ %>
				if (CheckNames(frmScreen.txtSurname.value) == false)
				{
					alert("First character of surname must be alpha");
					frmScreen.txtSurname.focus();
					return false;
				}
				if (CheckNames(frmScreen.txtFirstForename.value) == false)
				{		
					alert("First character of first forename must be alpha");
					frmScreen.txtFirstForename.focus(); 
					return false;
				}
				if (CheckNames(frmScreen.txtSecondForename.value) == false)
				{
					alert("First character of second forename must be alpha");
					frmScreen.txtSecondForename.focus(); 
					return false;
				}
				if (CheckNames(frmScreen.txtOtherForenames.value) == false)
				{
					alert("First character of other forenames must be alpha");
					frmScreen.txtOtherForenames.focus(); 
					return false;
				}
		
				<% /* SYS0499 Only call calculate customer age if date of birth input (not a mandatory field) */ %>
				if (frmScreen.txtDateOfBirth.value != "")
				{
					var nAge = CalculateCustomerAge();
					if(nAge < 0 || nAge > 99)
					{
						bSuccess = false;
						alert("Date of birth is invalid");
						frmScreen.txtDateOfBirth.focus();
					}
				}
				if( bSuccess == true && scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalRelatives", 255, true))
					bSuccess = false;

				if(bSuccess == true)bSuccess = SaveDependant();
			}	
		}
		else bSuccess = false;
	}
	return(bSuccess);					
}
function DefaultFields()
{
	if(m_nNumOfApplicants == 1)
	{
		frmScreen.cboApplicantName.selectedIndex = 1;
		<% /* SYS0994 Disable field also */ %>
		frmScreen.cboApplicantName.disabled = true;
	}
	else frmScreen.cboApplicantName.selectedIndex = 0;
	frmScreen.txtAdditionalRelatives.value = "";
	frmScreen.optAdditionalApplicantsYes.checked = false;
	frmScreen.optAdditionalApplicantsNo.checked = true;
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAdditionalRelatives");			
	frmScreen.optResidentAtNewAddressYes.checked = false;
	frmScreen.optResidentAtNewAddressNo.checked = true;			
	frmScreen.txtFirstForename.value	= "";
	frmScreen.txtSecondForename.value	= "";
	frmScreen.txtOtherForenames.value	= "";
	frmScreen.txtSurname.value			= "";
	frmScreen.txtDateOfBirth.value		= "";
	frmScreen.cboGender.selectedIndex	= 0;
	frmScreen.cboRelationship.selectedIndex	= 0;
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
	if(frmScreen.optAdditionalApplicantsYes.checked == false)
	{
		scScreenFunctions.SetFieldToDisabled(frmScreen, "txtAdditionalRelatives");
	}
	<% /* Reset label colour to 616161 as only mandatory if radio button set to yes */ %>
	lblAdditionalRelatives.style.color = "616161";
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}
function PopulateCombos()
{
	var blnSuccess = true;
	var XMLAccountType	= null;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
<%	// populate Applicant name from the customers in context
%>	// Add a <SELECT id=select1 name=select1> option
	var TagOPTION	= document.createElement("OPTION");
	TagOPTION.value	= "";
	TagOPTION.text	= "<SELECT>";
	frmScreen.cboApplicantName.add(TagOPTION);
	m_nNumOfApplicants = 0;
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName			= scScreenFunctions.GetContextParameter(window,"idCustomerName"	+ nLoop);
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
%>	if(m_nNumOfApplicants == 1)	
	{	
		<% /* SYS0994 Disable if only one applicant */ %>
		frmScreen.cboApplicantName.selectedIndex = 1;
		frmScreen.cboApplicantName.disabled = true;
	}
	else frmScreen.cboApplicantName.selectedIndex = 0;
<%	// Populate the rest of the combos
%>	var sGroupList = new Array("Sex","DependantRelationship");
	if(XML.GetComboLists(document, sGroupList) == true)
	{
		blnSuccess = blnSuccess & XML.PopulateCombo(document, frmScreen.cboGender,"Sex",true);
		blnSuccess = blnSuccess & XML.PopulateCombo(document, frmScreen.cboRelationship,"DependantRelationship",true);
		if(blnSuccess == false)	scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
	XML = null;
}
function PopulateScreen()
{
<%	//we are in edit mode. Load the dependant XML from context
%>	editXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	editXML.LoadXML(m_sXML);
	editXML.SelectTag(null,"DEPENDANT");
	frmScreen.cboApplicantName.value	= editXML.GetTagText("CUSTOMERNUMBER");
	frmScreen.txtFirstForename.value	= editXML.GetTagText("FIRSTFORENAME");
	frmScreen.txtSecondForename.value	= editXML.GetTagText("SECONDFORENAME");
	frmScreen.txtOtherForenames.value	= editXML.GetTagText("OTHERFORENAMES");
	frmScreen.txtSurname.value			= editXML.GetTagText("SURNAME");												
	frmScreen.txtDateOfBirth.value		= editXML.GetTagText("DATEOFBIRTH");
	frmScreen.cboGender.value			= editXML.GetTagText("GENDER");
	frmScreen.cboRelationship.value		= editXML.GetTagText("DEPENDANTRELATIONSHIP");
	m_sPersonGUID					= editXML.GetTagText("PERSONGUID");
	m_sOtherResidentSequenceNumber	= editXML.GetTagText("OTHERRESIDENTSEQUENCENUMBER");
	if(m_sOtherResidentSequenceNumber == "")frmScreen.optResidentAtNewAddressNo.checked = true;
	else frmScreen.optResidentAtNewAddressYes.checked = true;
	var sOthersResponsibleYesNo = editXML.GetTagText("ADDITIONALINDICATOR");
	if(sOthersResponsibleYesNo == "1")
	{
		frmScreen.optAdditionalApplicantsYes.checked = true;
		frmScreen.txtAdditionalRelatives.value = editXML.GetTagText("ADDITIONALDETAILS");
	}
	if(sOthersResponsibleYesNo == "0")frmScreen.optAdditionalApplicantsNo.checked = true;
<%	//save the chosen index of applicant name
%>	m_nApplicantNameIndex = frmScreen.cboApplicantName.selectedIndex;
}
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Add");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	if(m_sMetaAction == "Edit")	m_sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B00003743");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");  
}
function SaveDependant()
{
<%  /*Description:	If in "Edit" mode then an existing dependant record
					is updated. If in "Add" mode then a new dependant record
					is created.
					An OTHERRESIDENT child node is created within the DEPENDANT XML
					if the dependant is 'resident at the new address'.
					Note that the details pertaining to a Person record
					(i.e. Title, Surname etc.) are given their own child node.
					If an OTHERRESIDENT child node exists then the Person sub-node
					becomes a child of that XML; otherwise it becomes a child
					of the DEPENDANT XML.*/
%>	var bSuccess = true;
	var bChangeKey = false;
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var tagREQUEST = XML.CreateRequestTag(window, "SAVE");
	//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD ROLLBACK 11/05/01 SYS2050
	// Add OPERATION TJ_30/03/2001 SYS2050
	XML.SetAttribute("OPERATION","CriticalDataCheck");
	XML.CreateActiveTag("CUSTOMER");
	// End OPERATION
	// END ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD ROLLBACK 11/05/01 SYS2050 
	var tagDEPENDANT = XML.CreateActiveTag("DEPENDANT");
	XML.CreateTag("CUSTOMERNUMBER", frmScreen.cboApplicantName.value);
	var TagOption = frmScreen.cboApplicantName.options.item(frmScreen.cboApplicantName.selectedIndex);
	var sAttribute = TagOption.getAttribute("CustomerVersionNumber");
	XML.CreateTag("CUSTOMERVERSIONNUMBER", sAttribute);
	
	<% /* SYS1168 */ %>
	if(m_sMetaAction == "Edit")
	{	
		var sOldCustomerNumber = editXML.GetTagText("CUSTOMERNUMBER");
		var sOldCustomerVersionNumber = editXML.GetTagText("CUSTOMERVERSIONNUMBER");

		// Compare the screen details with the original XML passed in to ascertain whether key details have been changed
		if ((sOldCustomerNumber != frmScreen.cboApplicantName.value) | (sOldCustomerVersionNumber != sAttribute))
			bChangeKey = true;
	}

	
	if(m_sMetaAction == "Edit" && !bChangeKey)XML.CreateTag("DEPENDANTSEQUENCENUMBER", editXML.GetTagText("DEPENDANTSEQUENCENUMBER"));
	
	XML.CreateTag("DEPENDANTRELATIONSHIP", frmScreen.cboRelationship.value);
	var sAdditionalDetails = "";
	if(frmScreen.optAdditionalApplicantsYes.checked == true)
	{
		XML.CreateTag("ADDITIONALINDICATOR", "1");
		sAdditionalDetails = frmScreen.txtAdditionalRelatives.value;
	}
	else if(frmScreen.optAdditionalApplicantsNo.checked == true)XML.CreateTag("ADDITIONALINDICATOR", "0");
	else XML.CreateTag("ADDITIONALINDICATOR", "");
	XML.CreateTag("ADDITIONALDETAILS", sAdditionalDetails);

	<% /* SYS1168 */ %>
	if(bChangeKey)
	{	
		// Key fields have been changed - pass in the old key values as part of the update
		XML.CreateActiveTag("PREVIOUSKEY");
		XML.CreateActiveTag("DEPENDANT");
		XML.CreateTag("CUSTOMERNUMBER", sOldCustomerNumber);
		XML.CreateTag("CUSTOMERVERSIONNUMBER", sOldCustomerVersionNumber);
		XML.CreateTag("DEPENDANTSEQUENCENUMBER", editXML.GetTagText("DEPENDANTSEQUENCENUMBER"));
		XML.SelectTag(null, "DEPENDANT");
	}
	
	if(frmScreen.optResidentAtNewAddressYes.checked | (m_sOtherResidentSequenceNumber != ""))
	{
<%		// Either the PERSON details are to relate to an OTHERRESIDENT or they already do
%>		var tagOTHERRESIDENT = XML.CreateActiveTag("OTHERRESIDENT");
		if (m_sOtherResidentSequenceNumber != "" && !bChangeKey)XML.CreateTag("OTHERRESIDENTSEQUENCENUMBER", m_sOtherResidentSequenceNumber);
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		if((m_sPersonGUID != "") && (frmScreen.optResidentAtNewAddressYes.checked) && !bChangeKey)
			XML.CreateTag("PERSONGUID", m_sPersonGUID);
	}
	if (frmScreen.optResidentAtNewAddressYes.checked == false)XML.ActiveTag = tagDEPENDANT
	var tagPERSON = XML.CreateActiveTag("PERSON");
	if(m_sPersonGUID != "" && !bChangeKey)	XML.CreateTag("PERSONGUID", m_sPersonGUID);
	XML.CreateTag("DATEOFBIRTH",	frmScreen.txtDateOfBirth.value);
	XML.CreateTag("FIRSTFORENAME",	frmScreen.txtFirstForename.value);
	XML.CreateTag("SECONDFORENAME", frmScreen.txtSecondForename.value);
	XML.CreateTag("OTHERFORENAMES", frmScreen.txtOtherForenames.value);
	XML.CreateTag("SURNAME",		frmScreen.txtSurname.value);
	XML.CreateTag("GENDER",			frmScreen.cboGender.value);
	//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD ROLLBACK 11/05/01 SYS2050

	// Add CRITICALDATACONTEXT TJ_30/03/2001 AQR SYS2050
	XML.SelectTag(null,"REQUEST");
	XML.CreateActiveTag("CRITICALDATACONTEXT");
	XML.SetAttribute("APPLICATIONNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null));
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null));
	XML.SetAttribute("SOURCEAPPLICATION","Omiga");
	XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
	XML.SetAttribute("ACTIVITYINSTANCE","1");
	XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
	XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
	XML.SetAttribute("COMPONENT","omApp.ApplicationBO");
	XML.SetAttribute("METHOD","SaveDependantForCustomer");
		
	window.status = "Critical Data Check - please wait";
		
	// 	XML.RunASP(document,"OmigaTMBO.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"OmigaTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	window.status = "";
	// 	//* XML.RunASP(document, "SaveDependant.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			//* XML.RunASP(document, "SaveDependant.asp");
			break;
		default: // Error
			//* XML.SetErrorResponse();
		}

	//end CRITICALDATACONTEXT  TJ_29/03/2001 
	// END ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD ROLLBACK 11/05/01 SYS2050 */
	//XML.RunASP(document, "SaveDependant.asp");
	bSuccess = XML.IsResponseOK();
	XML = null;
	return(bSuccess);
}

<% /* SYS0994 SA 24/5/01 New function to check 1st character of names */ %>
function CheckNames(sNameToCheck)
{
	var CheckStr = sNameToCheck.substr(0,1)
	if (_validation.isNum(CheckStr))
	{
		return false;
	}
	
	return true;	
	
}
function frmScreen.txtAdditionalRelatives.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalRelatives", 255, true);
}

-->
</script>
</body>
</html>
