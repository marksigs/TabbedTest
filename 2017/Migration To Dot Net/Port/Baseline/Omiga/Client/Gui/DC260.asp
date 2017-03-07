<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/*
Workfile:      DC260.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Builder Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		06/03/2000	Created
AD		15/03/2000	Incorporated third party include files.
AY		03/04/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
BG		22/05/00	SYS0686	Set focus to postcode after OK on error, page to readonly when entering in readonly mode
					validate expected completion date, limit field to 255.
MC		22/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
CL		05/03/01	SYS1920 Read only functionality added
JR		17/09/01	Omiplus24 - Removed FaxNumber, TelephoneNumber & EmailAddress. Now included in TPD.asp
MDC		01/10/01	SYS2785 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

EPSOM Specific History

Prog	Date		AQR			Description
PB		07/07/2006	EP543		Populate 'Title-Other' field from XML
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<% /* FORMS */ %>
<form id="frmToDC240" method="post" action="DC240.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
<div style="HEIGHT: 168px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Builder Details</strong> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Company Name
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyName" maxlength="45" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" onchange="ContactDetailsChanged()">
		</span> 
	</span>

	<span style="LEFT: 370px; POSITION: absolute; TOP: 31px" class="msgLabel">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton" onclick="btnDirectorySearch.onclick()">
	</span>

	<span style="TOP: 60px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Is in the directory?
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryYes" name="InDirectoryGroup" type="radio" value="1" onclick="InDirectoryChanged()">
			<label for="optInDirectoryYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryNo" name="InDirectoryGroup" type="radio" value="0" onclick="InDirectoryChanged()">
			<label for="optInDirectoryNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 84px" class="msgLabel">
		Expected Completion Date
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtExpectedCompletionDate" maxlength="10" style="POSITION: absolute; WIDTH: 80px" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 290px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Building work required prior to completion
		<span style="LEFT: 0px; POSITION: absolute; TOP: 21px">
			<TEXTAREA class=msgTxt id="txtBuilderPropertyDescription" rows=5 style="POSITION: absolute; WIDTH: 300px"></TEXTAREA> 
		</span> 
	</span> 
</div>

<div style="HEIGHT: 232px; LEFT: 10px; POSITION: absolute; TOP: 239px; WIDTH: 604px" class="msgGroup">
<!-- #include FILE="includes/thirdpartydetails.htm" -->
</div> 
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 500px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/DC260attribs.asp" -->
<!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sXML = "";
var scScreenFunctions;

var BuilderXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber  = "";
var m_blnReadOnly = false;


/* EVENTS */

function frmScreen.txtBuilderPropertyDescription.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtBuilderPropertyDescription", 255, true);
}

function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
	scScreenFunctions.SetContextParameter(window,"idXML", "");
	frmToDC240.submit();
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		//clear the contexts
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		scScreenFunctions.SetContextParameter(window,"idXML", "");
		frmToDC240.submit();
	}
}

function frmScreen.btnClear.onclick()
{
	ClearFields(true, false, false);
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
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	<%/* Parameters for thirdpartydetails.asp */%>
	m_sThirdPartyType = "4";

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Add/Edit Builder Details","DC260",scScreenFunctions);

	// MDC SYS2564 / SYS2785 Client Customisation
	var sGroups = new Array("Country");
	objDerivedOperations = new DerivedScreen(sGroups);

	RetrieveContextData();
	SetThirdPartyDetailsMasks();
	SetMasks();

	// MC SYS2564/SYS2785 for client customisation
	ThirdPartyCustomise();

	Validation_Init();	
	Initialise(true);

	//BG 22/05/2000 SYS0686 If Edit in Readonly mode, make screen readonly
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC260");
	if (m_blnReadOnly == true) 
	{
		m_sReadOnly = "1";
		<% /* MO - 18/11/2002 - BMIDS00376 */ %>
		if (frmScreen.btnDirectorySearch.style.visibility != "hidden") frmScreen.btnDirectorySearch.disabled = true;
	}
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				if(scScreenFunctions.RestrictLength(frmScreen, "txtBuilderPropertyDescription", 255, true))
					return false;
			
				if (scScreenFunctions.ValidatePostcode(frmScreen.txtPostcode))
					bSuccess = SaveBuilder();
				else
					bSuccess=false;
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function DefaultFields()
// Inserts default values into all fields
{
	ClearFields(true, true);

	with (frmScreen)
	{
		txtBuilderPropertyDescription.value = "";
		txtExpectedCompletionDate.value = "";
	}

	m_sDirectoryGUID = "";
	m_bDirectoryAddress = false;
}

function Initialise()
// Initialises the screen
{
	PopulateCombos();

	if(m_sMetaAction == "Edit")
		PopulateScreen();
	else
		DefaultFields();

	SetAvailableFunctionality();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateCombos()
// Populates all combos on the screen
{
	PopulateTPTitleCombo();

	// MDC SYS2564 / SYS2785 Client Customisation
	var sControlList = new Array(frmScreen.cboCountry);
	objDerivedOperations.GetComboLists(sControlList);
/*
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("Country");

	if(XML.GetComboLists(document,sGroupList))
	{
		blnSuccess = XML.PopulateCombo(document,frmScreen.cboCountry,"Country",false);

		if(blnSuccess == false)
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}

	XML = null;
*/
}

function PopulateScreen()
// Populates the screen with details of the item selected in 
{
	//we are in edit mode. Load the dependant XML from context
	BuilderXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	BuilderXML.LoadXML(m_sXML);
	BuilderXML.SelectTag(null,"APPLICATIONTHIRDPARTY");

	with (frmScreen)
	{
		txtCompanyName.value = BuilderXML.GetTagText("COMPANYNAME");
		txtBuilderPropertyDescription.value = BuilderXML.GetTagText("BUILDERPROPERTYDESCRIPTION");
		txtExpectedCompletionDate.value = BuilderXML.GetTagText("EXPECTEDCOMPLETIONDATE");
		txtContactForename.value = BuilderXML.GetTagText("CONTACTFORENAME");
		txtContactSurname.value = BuilderXML.GetTagText("CONTACTSURNAME");
		cboTitle.value = BuilderXML.GetTagText("CONTACTTITLE");
		<% /* PB 07/07/2006 EP543 Begin */ %>
		checkOtherTitleField();
		txtTitleOther.value = BuilderXML.GetTagText("CONTACTTITLEOTHER");
		<% /* EP543 End */ %>

		// MDC SYS2564/SYS2785 
		objDerivedOperations.LoadCounty(BuilderXML);
		// txtCounty.value = BuilderXML.GetTagText("COUNTY");

		txtDistrict.value = BuilderXML.GetTagText("DISTRICT");
		//txtEMailAddress.value = BuilderXML.GetTagText("EMAILADDRESS");
		//txtFaxNo.value = BuilderXML.GetTagText("FAXNUMBER");
		txtFlatNumber.value = BuilderXML.GetTagText("FLATNUMBER");
		txtHouseName.value = BuilderXML.GetTagText("BUILDINGORHOUSENAME");
		txtHouseNumber.value = BuilderXML.GetTagText("BUILDINGORHOUSENUMBER");
		txtPostcode.value = BuilderXML.GetTagText("POSTCODE");
		txtStreet.value = BuilderXML.GetTagText("STREET");
		//txtTelephoneNo.value = BuilderXML.GetTagText("TELEPHONENUMBER");
		txtTown.value = BuilderXML.GetTagText("TOWN");
		cboCountry.value = BuilderXML.GetTagText("COUNTRY");
	}

	m_sDirectoryGUID = BuilderXML.GetTagText("DIRECTORYGUID");
	m_sThirdPartyGUID = BuilderXML.GetTagText("THIRDPARTYGUID");
	m_bDirectoryAddress = (m_sDirectoryGUID != "");
	
	// 17/09/2001 JR OmiPlus24 
	var TempXML = BuilderXML.ActiveTag;
	var ContactXML = BuilderXML.SelectTag(null, "CONTACTDETAILS");
	if(ContactXML != null)
		m_sXMLContact = ContactXML.xml;
	BuilderXML.ActiveTag = TempXML;
	//End
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");

	if(m_sMetaAction == "Edit")
		m_sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);

	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B00003743");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");  

	var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bCanAddToDirectory = (sUserRole >= XML.GetGlobalParameterAmount(document,"AddToDirectoryRole"));

	XML = null;
}

function SaveBuilder()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,null)
	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");

	XML.CreateActiveTag("APPLICATIONBUILDER");
	// Create tags for the key
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("BUILDERPROPERTYDESCRIPTION",frmScreen.txtBuilderPropertyDescription.value);
	XML.CreateTag("EXPECTEDCOMPLETIONDATE",frmScreen.txtExpectedCompletionDate.value);

	if (m_sMetaAction == "Edit")
	{
		// Retrieve the original third party/directory GUIDs
		var sOriginalThirdPartyGUID = BuilderXML.GetTagText("THIRDPARTYGUID");
		var sOriginalDirectoryGUID = BuilderXML.GetTagText("DIRECTORYGUID");
		// Only retrieve the address/contact details GUID if we are updating an existing third party record
		var sAddressGUID = (sOriginalThirdPartyGUID != "") ? BuilderXML.GetTagText("ADDRESSGUID") : "";
		var sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? BuilderXML.GetTagText("CONTACTDETAILSGUID") : "";
	}

	// If changing from a directory to a third party, or vice-versa, the old directory/thirdparty GUID
	// should still be specified to alert the middler tier to the fact that the old link needs deleting
	XML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID != "" ? m_sDirectoryGUID : sOriginalDirectoryGUID);
	XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);

	if (!m_bDirectoryAddress)
	{
		// Store the third party details
		XML.CreateActiveTag("THIRDPARTY");
		XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
		XML.CreateTag("THIRDPARTYTYPE", 4); // Builder
		XML.CreateTag("COMPANYNAME", frmScreen.txtCompanyName.value);

		// Address
		XML.CreateTag("ADDRESSGUID", sAddressGUID);
		SaveAddress(XML, sAddressGUID);

		// Contact Details
		XML.SelectTag(null, "THIRDPARTY");
		XML.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
		SaveContactDetails(XML, sContactDetailsGUID);
	}

	// Save the details
	if (m_sMetaAction == "Edit")
		// 		XML.RunASP(document,"UpdateBuilder.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateBuilder.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	else
		// 		XML.RunASP(document,"CreateBuilder.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateBuilder.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}


	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}
function SetScreenToReadOnly()
{
<%	// Set all controls on the screen to read only
	// The buttons are treated separately
%>	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnDirectorySearch.disabled = true;
		frmScreen.btnClear.disabled = true;
		frmScreen.btnPAFSearch.disabled = true;
		DisableMainButton("Cancel");
	}
}
function frmScreen.txtExpectedCompletionDate.onblur()
{
	<% /* BG Check that the expected completion date isn't in the past */ %>
	if(scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtExpectedCompletionDate,"<") == true)
	{
		alert("Expected completion date should be in the future");
		frmScreen.txtExpectedCompletionDate.focus();
	}
}
-->
</script>
<!-- #include FILE="includes/thirdpartydetails.asp" -->
</body>
</html>


