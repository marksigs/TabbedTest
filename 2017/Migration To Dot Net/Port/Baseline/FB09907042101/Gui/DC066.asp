<%@ LANGUAGE="JSCRIPT" %>
<html>
<comment>
Workfile:      DC066.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Tenancy Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		01/12/99	Created
JLD		13/12/1999	Corrected routing if errors found
JLD		20/12/1999	SYS0091 - Changed look of screen (to match address on CR020)
					SYS0089 - If no records found on directory search, still go to ZA020
AD		02/02/2000	Rework
AY		30/03/00	New top menu/scScreenFunctions change
MH      02/05/00    SYS0618 Postcode validation
BG		17/05/00	SYS0752 Removed Tooltips
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
CL		05/03/01	SYS1920 Read only functionality added
SA		23/04/01	SYS1262 Field label changed from "Rent" to "Monthly Rent"
SA		31/05/01	SYS
MDC		15/06/01	OmiPlus24 Telephone Numbers
MDC		01/10/01	SYS2785 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS HISTORY
GD		17/11/2002	BMIDS00376 Disable screen if readonly
BS		10/06/2003	BM0521 Disable screen if screen is read only

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS HISTORY
PJO		19/12/2005 	MAR67 Third Party type should be '8' - Landlord

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
PB		12/07/2006	EP543		Populate 'Title-Other' field from XML
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

</comment>
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
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>

<%/* FORMS */%>
<form id="frmToGN030" method="post" action="GN030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC065" method="post" action="DC065.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
<div id="divLandlord" style="TOP: 60px; LEFT: 10px; HEIGHT: 132px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Landlord Type
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<select id="cboLandlordType" name="LandlordType" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="TOP: 36px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Landlord Name
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtCompanyName" maxlength="45" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt" onchange="AddressChanged()">
		</span>
	</span>			

	<span style="TOP: 31px; LEFT: 330px; POSITION: ABSOLUTE">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton">
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

	<span style="TOP: 84px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Monthly Rent
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtRent" name="Rent" maxlength="10" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 108px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Account Number
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtAccountNo" name="AccountNo" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
</div>

<div style="HEIGHT: 232px; LEFT: 10px; POSITION: absolute; TOP: 198px; WIDTH: 604px" class="msgGroup">
<!-- #include FILE="includes/thirdpartydetails.htm" -->
</div> 
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 460px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/DC066attribs.asp" -->
<!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->

<!-- Specify Code Here -->
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sXML = "";
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sCustomerAddressSeqNum = "";
var m_bPreviousAddress = false;
var TenancyXML = null;
var scScreenFunctions;
var m_blnReadOnly = false;

function btnCancel.onclick()
{
	//clear the contexts
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit")
	frmToDC065.submit();
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		//clear the contexts
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit")
		frmToDC065.submit();
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

	<% // MDC SYS2564/SYS2785 %>
	var sGroup = new Array("Country","LandlordType");
	objDerivedOperations = new DerivedScreen(sGroup);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Tenancy Details","DC066",scScreenFunctions);

	<%/* Parameters for thirdpartydetails.asp */%>
	m_fSetAvailableFunctionalityOverride = SetAvailableFunctionalityOverride;
	m_sThirdPartyType = "8";

	RetrieveContextData();
	SetThirdPartyDetailsMasks();
	SetMasks();

	// MC SYS2564/SYS2785 for client customisation
	ThirdPartyCustomise();

	Validation_Init();	
	Initialise(true);

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	//GD BMIDS00376 m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC066");
	//BS BM0521 10/06/03 - reinstate call here as the Initialise function re-enables the screen
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC066");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
<% // GD BMIDS00376 START %>
	if (m_sReadOnly == true)
	{
		frmScreen.btnDirectorySearch.disabled = true;
		if (frmScreen.btnClear.style.visibility != 'hidden')
		{
			frmScreen.btnClear.disabled = true;
		}
		if (frmScreen.btnPAFSearch.style.visibility != 'hidden')
		{
			frmScreen.btnPAFSearch.disabled = true;
		}		
		if (frmScreen.btnAddToDirectory.style.visibility != 'hidden')
		{
			frmScreen.btnAddToDirectory.disabled = true;
		}		
	}
	
<% // GD BMIDS00376 END %>	
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
					bSuccess = SaveTenancy();
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

	if (!PopulateScreen())
		DefaultFields();

	SetAvailableFunctionality();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateCombos()
// Populates all combos on the screen
{	
	PopulateTPTitleCombo();
	
	var sControlList = new Array(frmScreen.cboCountry, frmScreen.cboLandlordType)
	objDerivedOperations.GetComboLists(sControlList);
}

function PopulateScreen()
// Populates the screen with details of the item selected in 
{	
	//we are in edit mode. Load the dependant XML from context
	TenancyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	TenancyXML.CreateRequestTag(window,null);
	TenancyXML.CreateActiveTag("TENANCY");
	TenancyXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	TenancyXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	TenancyXML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER", m_sCustomerAddressSeqNum);
	TenancyXML.RunASP(document,"GetTenancyDetails.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = TenancyXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[1] == ErrorTypes[0]))
	{
		//Error: record not found
		m_sMetaAction = "Add";
	}
	else if (ErrorReturn[0] == true)
	{
		// No error
		m_sMetaAction = "Edit";

		if(TenancyXML.SelectTag(null, "TENANCY") != null)
		{
			with (frmScreen)
			{
				txtCompanyName.value = TenancyXML.GetTagText("COMPANYNAME");
				txtContactForename.value = TenancyXML.GetTagText("CONTACTFORENAME");
				txtContactSurname.value = TenancyXML.GetTagText("CONTACTSURNAME");
				cboTitle.value = TenancyXML.GetTagText("CONTACTTITLE");
				<% /* PB 07/06/2007 EP543 Begin */ %>
				checkOtherTitleField();
				txtTitleOther.value = TenancyXML.GetTagText("CONTACTTITLEOTHER");
				<% /* EP543 End */ %>
		
				// MDC SYS2564/SYS2785 
				objDerivedOperations.LoadCounty(TenancyXML);
				
				txtDistrict.value = TenancyXML.GetTagText("DISTRICT");
				//txtEMailAddress.value = TenancyXML.GetTagText("EMAILADDRESS");
				//txtFaxNo.value = TenancyXML.GetTagText("FAXNUMBER");
				txtFlatNumber.value = TenancyXML.GetTagText("FLATNUMBER");
				txtHouseName.value = TenancyXML.GetTagText("BUILDINGORHOUSENAME");
				txtHouseNumber.value = TenancyXML.GetTagText("BUILDINGORHOUSENUMBER");
				txtPostcode.value = TenancyXML.GetTagText("POSTCODE");
				txtStreet.value = TenancyXML.GetTagText("STREET");
				//txtTelephoneNo.value = TenancyXML.GetTagText("TELEPHONENUMBER");
				txtTown.value = TenancyXML.GetTagText("TOWN");
				cboCountry.value = TenancyXML.GetTagText("COUNTRY");
				txtRent.value = TenancyXML.GetTagText("MONTHLYRENTAMOUNT");
				txtAccountNo.value = TenancyXML.GetTagText("ACCOUNTNUMBER");
				cboLandlordType.value = TenancyXML.GetTagText("TENANCYTYPE");

				m_sDirectoryGUID = TenancyXML.GetTagText("DIRECTORYGUID");
				m_bDirectoryAddress = (m_sDirectoryGUID != "");
				m_sThirdPartyGUID = TenancyXML.GetTagText("THIRDPARTYGUID");
			}
			
			// 15/06/2001 MDC OmiPlus 24 
			var TempXML = TenancyXML.ActiveTag;
			var ContactXML = TenancyXML.SelectTag(null, "CONTACTDETAILS");
			if(ContactXML != null)
				m_sXMLContact = ContactXML.xml;
			TenancyXML.ActiveTag = TempXML;

		}
	}

	ErrorTypes = null;
	ErrorReturn = null;

	return(m_sMetaAction == "Edit");
}

function RetrieveContextData()
{
	//GD BMIDS00376 m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	<% /* BS BM0521 10/06/03
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC066");
	m_blnReadOnly = m_sReadOnly; */ %>
	var sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	if(sXML != "")
	{
		//Should always have XML passed from DC065
		XML.LoadXML(sXML);
		XML.SelectTag(null, "CUSTOMERADDRESS");
		m_sCustomerNumber = XML.GetTagText("CUSTOMERNUMBER");
		m_sCustomerVersionNumber = XML.GetTagText("CUSTOMERVERSIONNUMBER");
		m_sCustomerAddressSeqNum = XML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER");
	}

	var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	var sCustomerAddressType = scScreenFunctions.GetContextParameter(window,"idCustomerAddressType","0");
	m_bCanAddToDirectory = (sUserRole >= XML.GetGlobalParameterAmount(document,"AddToDirectoryRole"));
	m_bPreviousAddress = XML.IsInComboValidationList(document,"CustomerAddressType",sCustomerAddressType,["P"])

	XML = null;
}

function SaveTenancy()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,null)
	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");

	XML.CreateActiveTag("TENANCY");
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	XML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER", m_sCustomerAddressSeqNum);
	XML.CreateTag("TENANCYTYPE", frmScreen.cboLandlordType.value);
	XML.CreateTag("MONTHLYRENTAMOUNT", frmScreen.txtRent.value);
	XML.CreateTag("ACCOUNTNUMBER", frmScreen.txtAccountNo.value);

	if (m_sMetaAction == "Edit")
	{
		// Retrieve the original third party/directory GUIDs
		var sOriginalThirdPartyGUID = TenancyXML.GetTagText("THIRDPARTYGUID");
		var sOriginalDirectoryGUID = TenancyXML.GetTagText("DIRECTORYGUID");
		// Only retrieve the address/contact details GUID if we are updating an existing third party record
		var sAddressGUID = (sOriginalThirdPartyGUID != "") ? TenancyXML.GetTagText("ADDRESSGUID") : "";
		var sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? TenancyXML.GetTagText("CONTACTDETAILSGUID") : "";
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
		XML.CreateTag("THIRDPARTYTYPE", 8); // Landlord
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
		// 		XML.RunASP(document,"UpdateTenancy.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateTenancy.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	else
		// 		XML.RunASP(document,"CreateTenancy.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateTenancy.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}


	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}

function SetAvailableFunctionalityOverride()
{
	if (m_bPreviousAddress)
	{
		frmScreen.txtAccountNo.value = "";
		frmScreen.txtRent.value = "";
		<% /* SYS1034 These should not be read only
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtAccountNo");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtRent");*/ %>
	}
}
-->
</script>
<!-- #include FILE="includes/thirdpartydetails.asp" -->
</body>

</html>


