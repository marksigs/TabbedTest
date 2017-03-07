<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<% /*
Workfile:      DC215.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Vendor Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MC		08/05/2000	Created
BG		17/05/00	SYS0752 Removed Tooltips
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
MC		08/06/00	SYS0886 Default Vendor Address to New Property Address
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
CL		05/03/01	SYS1920 Read only functionality added
JR		01/10/01	Omiplus24, Telephone Number changes
MDC		01/10/01	SYS2785 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
TW      09/10/02    SYS5115 Modified to incorporate client validation

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
MO		20/11/2002	BMIDS00376	Fixed bug noted after the first release of this AQR
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Mars History:

Prog	Date		Description
GHun	16/05/2006	MAR1795 Changed DefaultFields to skip over a blank address
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
PB		21/07/2006	EP543		Added 'TITLEOTHER' field for third parties
IK		28/07/2006	EP543		fix
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* FORMS */ %>
<form id="frmToDC210" method="post" action="DC210.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" validate ="onchange" mark>

<div style="HEIGHT: 25px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<font face="MS Sans Serif" size="1">
			<strong>Vendor Details</strong> 
		</font>
	</span>
</div> 

<div style="HEIGHT: 232px; LEFT: 10px; POSITION: absolute; TOP: 85px; WIDTH: 608px" class="msgGroup">
	<!-- #include FILE="includes/thirdpartydetails.htm" -->
</div> 

<!-- Company Name etc required by Third Party includes. Therefore added to form but hidden -->
<div style="HEIGHT: 88px; LEFT: 10px; POSITION: absolute; TOP: 305px; VISIBILITY: hidden; WIDTH: 604px" class="msgGroup">

	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Company Name
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyName" name="CompanyName" maxlength="45" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" onchange="ContactDetailsChanged()">
		</span> 
	</span>

	<span style="LEFT: 330px; POSITION: absolute; TOP: 31px" class="msgLabel">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton" onclick="btnDirectorySearch.onClick()">
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Is in the directory?
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px; WIDTH: 60px">
			<input id="optInDirectoryYes" name="InDirectoryGroup" type="radio" value="1" onclick="InDirectoryChanged()">
			<label for="optInDirectoryYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 174px; POSITION: absolute; TOP: -3px; WIDTH: 60px">
			<input id="optInDirectoryNo" name="InDirectoryGroup" type="radio" value="0" onclick="InDirectoryChanged()">
			<label for="optInDirectoryNo" class="msgLabel">No</label> 
		</span> 
	</span>
</div>

</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 440px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/DC215Attribs.asp" -->

<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sXML = "";
var scScreenFunctions;

var VendorXML = null;
var m_bVendorExists = false;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber  = "";
var m_blnReadOnly = false;


<%/* EVENTS */%>

function btnCancel.onclick()
{
	frmToDC210.submit();
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		frmToDC210.submit();
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
	
	<%// Added by automated update TW 09 Oct 2002 SYS5115 %>
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	<% /* Hide the Add to Directory button as Vendor cannot be added to directory */ %>
	<% // frmScreen.btnAddToDirectory.style.visibility = "hidden"; %>

	<%/* Parameters for thirdpartydetails.asp */%>
	m_sThirdPartyType = "12"; <% /* Vendor */ %>

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Add/Edit Vendor Details","DC215",scScreenFunctions);

	<%// MDC SYS2564 / SYS2785 Client Customisation %>
	var sGroups = new Array("Country");
	objDerivedOperations = new DerivedScreen(sGroups);

	RetrieveContextData();
	SetThirdPartyDetailsMasks();

	<%// MC SYS2564/SYS2785 for client customisation %>
	ThirdPartyCustomise();

	SetMasks();
	Validation_Init();	
	Initialise(true);

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC215");
	if (m_blnReadOnly == true) 
	{
		m_sReadOnly = "1";
		<% /* MO - 18/11/2002 - BMIDS00376 */ %>
		if (frmScreen.btnDirectorySearch.style.visibility != "hidden") frmScreen.btnDirectorySearch.disabled = true;
		if (frmScreen.btnPAFSearch.style.visibility != "hidden") frmScreen.btnPAFSearch.disabled = true;
		if (frmScreen.btnClear.style.visibility != "hidden") frmScreen.btnClear.disabled = true;
	}
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
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
					bSuccess = SaveVendor();
				else
					bSuccess=false;
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function DefaultFields()
{
	ClearFields(true, true);

	<% /* Default to New Property Address */ %>
	sXML = scScreenFunctions.GetContextParameter(window, "idXML", "");
	if(sXML != "")
	{
		var AddressXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		AddressXML.LoadXML(sXML);
		AddressXML.SelectTag(null, "ADDRESS");
		
		if (AddressXML.ActiveTag) <% /* MAR1795 GHun Only continue if an address exists */ %>
		{
			with (frmScreen)
			{
				txtFlatNumber.value = AddressXML.GetTagText("FLATNUMBER");
				txtHouseName.value = AddressXML.GetTagText("BUILDINGORHOUSENAME");
				txtHouseNumber.value = AddressXML.GetTagText("BUILDINGORHOUSENUMBER");
				txtPostcode.value = AddressXML.GetTagText("POSTCODE");
				txtStreet.value = AddressXML.GetTagText("STREET");
				txtTown.value = AddressXML.GetTagText("TOWN");
				txtDistrict.value = AddressXML.GetTagText("DISTRICT");
				
				<%// MDC SYS2564/SYS2785 %>
				objDerivedOperations.ClearCounty();
				<%//txtCounty.value = AddressXML.GetTagText("COUNTY"); %>

				cboCountry.value = AddressXML.GetTagText("COUNTRY");
			}
		}
		FlagChange(true);
	}
}

function Initialise(bOnLoad)
{
	if(bOnLoad == true)
		PopulateCombos();
	
	if(GetVendorData() && m_bVendorExists)
	{
		m_sMetaAction = "Edit";
		PopulateScreen();
	}
	else
	{
		m_sMetaAction = "Add";
		DefaultFields();
	}
	
	SetAvailableFunctionality();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateCombos()
{
	PopulateTPTitleCombo();

	<%// MDC SYS2564 / SYS2785 Client Customisation %>
	var sControlList = new Array(frmScreen.cboCountry);
	objDerivedOperations.GetComboLists(sControlList);
}

function GetVendorData()
{
var bSuccess = false;

	<% /* Retrieve Vendor data */ %>
	VendorXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	with (VendorXML)
	{
		<% /* Create request block */ %>
		CreateRequestTag(window);
		CreateActiveTag("NEWPROPERTYVENDOR");
		CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		RunASP(document, "GetVendorDetails.asp");

		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = CheckResponse(ErrorTypes);
				
		if(ErrorReturn[1] == ErrorTypes[0])
		{
			<%/* Record not found */%>
			m_bVendorExists = false;
			bSuccess = true;
		}
		else if (ErrorReturn[0] == true)
		{
			<%/* Record found */%>
			m_bVendorExists = true;
			bSuccess = true;
		}
		
		return bSuccess;
	}
}

function PopulateScreen()
{
	VendorXML.SelectTag(null,"THIRDPARTY");
	with (frmScreen)
	{
		txtCompanyName.value = VendorXML.GetTagText("COMPANYNAME");
		txtContactForename.value = VendorXML.GetTagText("CONTACTFORENAME");
		txtContactSurname.value = VendorXML.GetTagText("CONTACTSURNAME");
		cboTitle.value = VendorXML.GetTagText("CONTACTTITLE");
		<% /* PB 07/07/2006 EP543 Begin */ %>
		checkOtherTitleField();
		txtTitleOther.value = VendorXML.GetTagText("CONTACTTITLEOTHER");
		<% /* EP543 End */ %>

		<%// MDC SYS2564/SYS2785 %>
		objDerivedOperations.LoadCounty(VendorXML);
		<%// txtCounty.value = VendorXML.GetTagText("COUNTY");%>

		txtDistrict.value = VendorXML.GetTagText("DISTRICT");
		<%//txtEMailAddress.value = VendorXML.GetTagText("EMAILADDRESS"); %>
		<%//txtFaxNo.value = VendorXML.GetTagText("FAXNUMBER"); %>
		txtFlatNumber.value = VendorXML.GetTagText("FLATNUMBER");
		txtHouseName.value = VendorXML.GetTagText("BUILDINGORHOUSENAME");
		txtHouseNumber.value = VendorXML.GetTagText("BUILDINGORHOUSENUMBER");
		txtPostcode.value = VendorXML.GetTagText("POSTCODE");
		txtStreet.value = VendorXML.GetTagText("STREET");
		<%//txtTelephoneNo.value = VendorXML.GetTagText("TELEPHONENUMBER"); %>
		txtTown.value = VendorXML.GetTagText("TOWN");
		cboCountry.value = VendorXML.GetTagText("COUNTRY");
		
		<%// 01/10/2001 JR OmiPlus 24 %>
		var TempXML = VendorXML.ActiveTag;
		var ContactXML = VendorXML.SelectTag(null, "CONTACTDETAILS");
		if(ContactXML != null)
			m_sXMLContact = ContactXML.xml;
		VendorXML.ActiveTag = TempXML;
	}

	m_sDirectoryGUID = VendorXML.GetTagText("DIRECTORYGUID");
	m_bDirectoryAddress = (m_sDirectoryGUID != "");
	m_sThirdPartyGUID = VendorXML.GetTagText("THIRDPARTYGUID");
}

function RetrieveContextData()
{
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","C00027871");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");  

	<% /* This will hide the Add To Directory button */ %>
	m_bCanAddToDirectory = false;

}

function SaveVendor()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null)
	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");

	var tagVendor = XML.CreateActiveTag("NEWPROPERTYVENDOR");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);

	if (m_sMetaAction == "Edit")
	{
		<% // Retrieve the original third party/directory GUIDs %>
		var sOriginalThirdPartyGUID = VendorXML.GetTagText("THIRDPARTYGUID");
		
		<%// Only retrieve the address/contact details GUID if we are updating an existing third party record %>
		var sAddressGUID = (sOriginalThirdPartyGUID != "") ? VendorXML.GetTagText("ADDRESSGUID") : "";
		var sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? VendorXML.GetTagText("CONTACTDETAILSGUID") : "";
	}

	<%// If changing from a directory to a third party, or vice-versa, the old directory/thirdparty GUID
	// should still be specified to alert the middler tier to the fact that the old link needs deleting %>
	XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);

	if (!m_bDirectoryAddress)
	{
		<% /* Third Party Details */ %>
		XML.CreateActiveTag("THIRDPARTY");
		XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
		XML.CreateTag("THIRDPARTYTYPE", 12); // Vendor
		<% /* Use dummy value for company as not required for Vendors */ %>
		XML.CreateTag("COMPANYNAME", "Vendor");

		<% /* Address */ %>
		XML.CreateTag("ADDRESSGUID", sAddressGUID);
		SaveAddress(XML, sAddressGUID);

		<% /* Contact Details */ %>
		XML.SelectTag(null, "THIRDPARTY");
		XML.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
		SaveContactDetails(XML, sContactDetailsGUID);
	}

	<% /* Save the details */ %>
	if (m_sMetaAction == "Edit")
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateVendor.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	else
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateVendor.asp");
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
<script src="validation.js" language="JScript" type="text/javascript"></script>
</body>
</html>
