<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/*
Workfile:      DC260.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Architect Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		06/03/2000	Created
AD		15/03/2000	Incorporated third party include files.
AD		21/03/2000	SYS0479 - OK/Cancel buttons moved up
AY		03/04/00	New top menu/scScreenFunctions change
MH		03/05/00	SYS0618 Postcode validation
MH      16/05/00    SYS0700 Behaviour
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
CL		05/03/01	SYS1920 Read only functionality added
JR		17/09/01	Omiplus24 - Removed FaxNumber, TelephoneNumber & EmailAddress. Now included in TPD.asp
MDC		01/10/01	SYS2785 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
HMA     17/09/2003  BM0063      Amend HTML text for radio buttons
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<div style="HEIGHT: 160px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Architect Details</strong> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Company Name
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyName" maxlength="45" style="POSITION: absolute; WIDTH: 200px" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 370px; POSITION: absolute; TOP: 31px" class="msgLabel">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton">
	</span>

	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span style="TOP: 60px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Is in the directory?
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryYes" name="InDirectoryGroup" type="radio" value="1" onclick="InDirectoryChanged()"><label for="optInDirectoryYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryNo" name="InDirectoryGroup" type="radio" value="0" onclick="InDirectoryChanged()"><label for="optInDirectoryNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 84px" class="msgLabel">
		Architect's Qualications
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtArchitectQualifications" maxlength="50" style="POSITION: absolute; WIDTH: 200px" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 108px" class="msgLabel">
		RIBA qualified?
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="optRIBAQualifiedYes" name="RIBAQualifiedGroup" type="radio" value="1"><label for="optRIBAQualifiedYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px">
			<input id="optRIBAQualifiedNo" name="RIBAQualifiedGroup" type="radio" value="0"><label for="optRIBAQualifiedNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="TOP: 132px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Level of Indemnity Cover
		<span style="TOP: -3px; LEFT: 160px; POSITION: ABSOLUTE">
			<select id="cboLevelOfIndemnityCover" style="WIDTH: 200px" class="msgCombo"></select>
		</span>
	</span>
</div>

<div style="HEIGHT: 232px; LEFT: 10px; POSITION: absolute; TOP: 226px; WIDTH: 604px" class="msgGroup">
<!-- #include FILE="includes/thirdpartydetails.htm" -->
</div> 
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 500px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->
<!-- #include FILE="includes/thirdpartydetails.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/DC290Attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sXML = "";
var scScreenFunctions;

var ArchitectXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber  = "";
var m_blnReadOnly = false;

/* EVENTS */

<% /* This event is also in Thirdpartydetails.asp. In Ie5/iis4 search precedence 
  finds this one first. If this changes it wil need to be moved */ %>
function frmScreen.btnDirectorySearch.onclick()
{
	<% /* Call the routine in TPD and then update the screen when it returns */ %>
	ThirdPartyDetailsDirectorySearch();

	if (m_sMetaAction!="Add") 
	{
		with (frmScreen)
		{
			txtArchitectQualifications.value = "";
			scScreenFunctions.SetRadioGroupValue(frmScreen,"RIBAQualifiedGroup","0");
			cboLevelOfIndemnityCover.selectedIndex = 0;
		}
	}
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
	m_sThirdPartyType = "2";

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Add/Edit Architect Details","DC290",scScreenFunctions);

	// MDC SYS2564 / SYS2785 Client Customisation
	var sGroups = new Array("Country","IndemnityCoverLevel");
	objDerivedOperations = new DerivedScreen(sGroups);

	RetrieveContextData();
	SetThirdPartyDetailsMasks();

	// MC SYS2564/SYS2785 for client customisation
	ThirdPartyCustomise();

	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();	
	Initialise(true);
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC290");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	SetFieldsToReadOnly();
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
				if (scScreenFunctions.ValidatePostcode(frmScreen.txtPostcode))
					bSuccess = SaveArchitect();
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
		txtArchitectQualifications.value = "";
		scScreenFunctions.SetRadioGroupValue(frmScreen,"RIBAQualifiedGroup","0");
		cboLevelOfIndemnityCover.selectedIndex = 0;
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

}

function PopulateCombos()
// Populates all combos on the screen
{
	PopulateTPTitleCombo();

	// MDC SYS2564 / SYS2785 Client Customisation
	var sControlList = new Array(frmScreen.cboCountry, frmScreen.cboLevelOfIndemnityCover);
	objDerivedOperations.GetComboLists(sControlList);
/*
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("Country","IndemnityCoverLevel");

	if(XML.GetComboLists(document,sGroupList))
	{
		blnSuccess = XML.PopulateCombo(document,frmScreen.cboCountry,"Country",false);
		blnSuccess = blnSuccess & XML.PopulateCombo(document,frmScreen.cboLevelOfIndemnityCover,"IndemnityCoverLevel",true);

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
	ArchitectXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ArchitectXML.LoadXML(m_sXML);
	ArchitectXML.SelectTag(null,"APPLICATIONTHIRDPARTY");

	with (frmScreen)
	{
		txtCompanyName.value = ArchitectXML.GetTagText("COMPANYNAME");
		txtArchitectQualifications.value = ArchitectXML.GetTagText("QUALIFICATIONS");
		cboLevelOfIndemnityCover.value = ArchitectXML.GetTagText("LEVELOFINDEMNITYCOVER");
		scScreenFunctions.SetRadioGroupValue(frmScreen,"RIBAQualifiedGroup",ArchitectXML.GetTagText("RIBAQUALIFIEDINDICATOR"));
		txtContactForename.value = ArchitectXML.GetTagText("CONTACTFORENAME");
		txtContactSurname.value = ArchitectXML.GetTagText("CONTACTSURNAME");
		cboTitle.value = ArchitectXML.GetTagText("CONTACTTITLE");
		<% /* PB 07/07/2006 EP543 Begin */ %>
		checkOtherTitleField();
		txtTitleOther.value = ArchitectXML.GetTagText("CONTACTTITLEOTHER");
		<% /* EP543 End */ %>

		// MDC SYS2564/SYS2785 
		objDerivedOperations.LoadCounty(ArchitectXML);
		// txtCounty.value = ArchitectXML.GetTagText("COUNTY");

		txtDistrict.value = ArchitectXML.GetTagText("DISTRICT");
		//txtEMailAddress.value = ArchitectXML.GetTagText("EMAILADDRESS");
		//txtFaxNo.value = ArchitectXML.GetTagText("FAXNUMBER");
		txtFlatNumber.value = ArchitectXML.GetTagText("FLATNUMBER");
		txtHouseName.value = ArchitectXML.GetTagText("BUILDINGORHOUSENAME");
		txtHouseNumber.value = ArchitectXML.GetTagText("BUILDINGORHOUSENUMBER");
		txtPostcode.value = ArchitectXML.GetTagText("POSTCODE");
		txtStreet.value = ArchitectXML.GetTagText("STREET");
		//txtTelephoneNo.value = ArchitectXML.GetTagText("TELEPHONENUMBER");
		txtTown.value = ArchitectXML.GetTagText("TOWN");
		cboCountry.value = ArchitectXML.GetTagText("COUNTRY");
	}
	
	m_sDirectoryGUID = ArchitectXML.GetTagText("DIRECTORYGUID");
	m_sThirdPartyGUID = ArchitectXML.GetTagText("THIRDPARTYGUID");
	m_bDirectoryAddress = (m_sDirectoryGUID != "");
	
	// 17/09/2001 JR OmiPlus24 
	var TempXML = ArchitectXML.ActiveTag;
	var ContactXML = ArchitectXML.SelectTag(null, "CONTACTDETAILS");
	if(ContactXML != null)
		m_sXMLContact = ContactXML.xml;
	ArchitectXML.ActiveTag = TempXML;
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

function SaveArchitect()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,null)
	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");

	XML.CreateActiveTag("APPLICATIONARCHITECT");
	// Create tags for the key
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("QUALIFICATIONS",frmScreen.txtArchitectQualifications.value);
	XML.CreateTag("LEVELOFINDEMNITYCOVER",frmScreen.cboLevelOfIndemnityCover.value);
	XML.CreateTag("RIBAQUALIFIEDINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"RIBAQualifiedGroup"));

	if (m_sMetaAction == "Edit")
	{
		// Retrieve the original third party/directory GUIDs
		var sOriginalThirdPartyGUID = ArchitectXML.GetTagText("THIRDPARTYGUID");
		var sOriginalDirectoryGUID = ArchitectXML.GetTagText("DIRECTORYGUID");
		// Only retrieve the address/contact details GUID if we are updating an existing third party record
		var sAddressGUID = (sOriginalThirdPartyGUID != "") ? ArchitectXML.GetTagText("ADDRESSGUID") : "";
		var sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? ArchitectXML.GetTagText("CONTACTDETAILSGUID") : "";
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
		XML.CreateTag("THIRDPARTYTYPE", 2); // Architect
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
		// 		XML.RunASP(document,"UpdateArchitect.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateArchitect.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	else
		// 		XML.RunASP(document,"CreateArchitect.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateArchitect.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}


	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}

function SetFieldsToReadOnly()
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
-->
</script>
</body>
</html>




