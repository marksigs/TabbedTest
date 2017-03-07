<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/*
Workfile:      DC250.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Estate Agent Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		20/01/2000	Created
AD		02/02/2000	Rework
AY		14/02/00	Change to msgButtons button types
AD		15/03/2000	Incorporated third party include files.
AY		03/04/00	New top menu/scScreenFunctions change
MH		03/05/00	SYS0618 Validate postcode
MH      17/05/00    SYS0683 Various bits of validation 'n' stuff
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
CL		05/03/01	SYS1920 Read only functionality added
JR		17/09/01	Omiplus24 - Removed FaxNumber, TelephoneNumber & EmailAddress. Now included in TPD.asp
MDC		01/10/01	SYS2785 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

EPSOM Specific History

Prog	Date		AQR			Description
PB		07/07/2006	EP543		Populate 'Title-Other' field from XML
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
<div style="HEIGHT: 88px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Estate Agent Details</strong> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Company Name
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyName" name="CompanyName" maxlength="45" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" onchange="ContactDetailsChanged()">
		</span> 
	</span>

	<span style="LEFT: 330px; POSITION: absolute; TOP: 31px" class="msgLabel">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton" onclick="btnDirectorySearch.onClick()">
	</span>

	<span style="TOP: 60px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Is in the directory?
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryYes" name="InDirectoryGroup" type="radio" value="1" onclick="InDirectoryChanged()">
			<label for="optInDirectoryYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 174px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryNo" name="InDirectoryGroup" type="radio" value="0" onclick="InDirectoryChanged()">
			<label for="optInDirectoryNo" class="msgLabel">No</label> 
		</span> 
	</span>
</div>

<div style="HEIGHT: 232px; LEFT: 10px; POSITION: absolute; TOP: 154px; WIDTH: 604px" class="msgGroup">
<!-- #include FILE="includes/thirdpartydetails.htm" -->
</div> 
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 440px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/DC250Attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sXML = "";
var scScreenFunctions;

var EstateAgentXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber  = "";
var m_blnReadOnly = false;

/* EVENTS */

function btnCancel.onclick()
{
	//clear the contexts
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
	m_sThirdPartyType = "6";

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Add/Edit Estate Agent Details","DC250",scScreenFunctions);

	// MDC SYS2564 / SYS2785 Client Customisation
	var sGroups = new Array("Country");
	objDerivedOperations = new DerivedScreen(sGroups);

	RetrieveContextData();
	SetThirdPartyDetailsMasks();

	// MC SYS2564/SYS2785 for client customisation
	ThirdPartyCustomise();

	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();	
	Initialise(true);

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC250");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	SetScreenOnReadOnly();
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */
function SetScreenOnReadOnly()
{
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnDirectorySearch.disabled=true;
		frmScreen.btnClear.disabled=true;
		frmScreen.btnPAFSearch.disabled=true;
		frmScreen.btnAddToDirectory.disabled=true;
	}
}

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				if (scScreenFunctions.ValidatePostcode(frmScreen.txtPostcode))
					bSuccess = SaveEstateAgent();
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
}

function Initialise(bOnLoad)
// Initialises the screen
{
	if(bOnLoad == true)
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
		// Populate Country combo
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
	EstateAgentXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	EstateAgentXML.LoadXML(m_sXML);
	EstateAgentXML.SelectTag(null,"APPLICATIONTHIRDPARTY");

	with (frmScreen)
	{
		txtCompanyName.value = EstateAgentXML.GetTagText("COMPANYNAME");
		txtContactForename.value = EstateAgentXML.GetTagText("CONTACTFORENAME");
		txtContactSurname.value = EstateAgentXML.GetTagText("CONTACTSURNAME");
		cboTitle.value = EstateAgentXML.GetTagText("CONTACTTITLE");
		<% /* PB 07/07/2006 EP543 Begin */ %>
		checkOtherTitleField();
		txtTitleOther.value = EstateAgentXML.GetTagText("CONTACTTITLEOTHER");
		<% /* EP543 End */ %>

		// MDC SYS2564/SYS2785 
		objDerivedOperations.LoadCounty(EstateAgentXML);
		// txtCounty.value = EstateAgentXML.GetTagText("COUNTY");

		txtDistrict.value = EstateAgentXML.GetTagText("DISTRICT");
		//txtEMailAddress.value = EstateAgentXML.GetTagText("EMAILADDRESS");
		//txtFaxNo.value = EstateAgentXML.GetTagText("FAXNUMBER");
		txtFlatNumber.value = EstateAgentXML.GetTagText("FLATNUMBER");
		txtHouseName.value = EstateAgentXML.GetTagText("BUILDINGORHOUSENAME");
		txtHouseNumber.value = EstateAgentXML.GetTagText("BUILDINGORHOUSENUMBER");
		txtPostcode.value = EstateAgentXML.GetTagText("POSTCODE");
		txtStreet.value = EstateAgentXML.GetTagText("STREET");
		//txtTelephoneNo.value = EstateAgentXML.GetTagText("TELEPHONENUMBER");
		txtTown.value = EstateAgentXML.GetTagText("TOWN");
		cboCountry.value = EstateAgentXML.GetTagText("COUNTRY");
	}

	m_sDirectoryGUID = EstateAgentXML.GetTagText("DIRECTORYGUID");
	m_bDirectoryAddress = (m_sDirectoryGUID != "");
	m_sThirdPartyGUID = EstateAgentXML.GetTagText("THIRDPARTYGUID");
	
	// 17/09/2001 JR OmiPlus24 
	var TempXML = EstateAgentXML.ActiveTag;
	var ContactXML = EstateAgentXML.SelectTag(null, "CONTACTDETAILS");
	if(ContactXML != null)
		m_sXMLContact = ContactXML.xml;
	EstateAgentXML.ActiveTag = TempXML;
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

function SaveEstateAgent()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,null)
	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");

	var tagEstateAgent = XML.CreateActiveTag("APPLICATIONESTATEAGENT");
	// Create tags for the key
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);

	if (m_sMetaAction == "Edit")
	{
		// Retrieve the original third party/directory GUIDs
		var sOriginalThirdPartyGUID = EstateAgentXML.GetTagText("THIRDPARTYGUID");
		var sOriginalDirectoryGUID = EstateAgentXML.GetTagText("DIRECTORYGUID");
		// Only retrieve the address/contact details GUID if we are updating an existing third party record
		var sAddressGUID = (sOriginalThirdPartyGUID != "") ? EstateAgentXML.GetTagText("ADDRESSGUID") : "";
		var sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? EstateAgentXML.GetTagText("CONTACTDETAILSGUID") : "";
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
		XML.CreateTag("THIRDPARTYTYPE", 6); // Estate Agent
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
		// 		XML.RunASP(document,"UpdateEstateAgent.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateEstateAgent.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	else
		// 		XML.RunASP(document,"CreateEstateAgent.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateEstateAgent.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}


	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}
-->
</script>

<!-- #include FILE="includes/thirdpartydetails.asp" -->
</body>
</html>




