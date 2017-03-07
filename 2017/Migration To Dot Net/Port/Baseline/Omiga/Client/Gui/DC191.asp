<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /* 
Workfile:      DC191.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Accountant Details (Popup)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		18/02/2000	Created
AD		15/03/2000	Incorporated third party include files.
AY		31/03/00	scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
MH      14/06/00    SYS0791 Address not saving
MC		21/06/00	SYS0756 If Read-only mode, disabled fields
BG		30/08/00	SYS1453	Added Title
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
JR		14/09/01	Omiplus24 - Removed FaxNumber, TelephoneNumber & EmailAddress. Now included in TPD.asp
MDC		01/10/01	SYS2785 Enable client versions to override labels
DPF		20/06/02	BMIDS00077  Changes made to file to bring in line with Core V7.0.2, changes are...
					SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		17/05/02	BMIDS00008	Modified SaveAccountant()- Amended with new Tag EMPLOYMENT
								Amended m_sAccountantGUID in PopulateScreen()
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MC		19/04/2004	BMIDS517	White space padded to the title text (To Hide default IE text for dialog windows - 'Web Page Dialog..')
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %> 
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
PB		21/07/2006	EP543		Added 'TITLEOTHER' field for third parties
AW		01/08/2006	EP1060		Amended xml doc ref for 'CONTACTTITLEOTHER'
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
	<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Accountant Details  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>
<% /* removed as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
<OBJECT data="scScreenFunctions.asp" height=1 id=scScrFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data="scXMLFunctions.asp" height=1 id=scXMLFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
*/ %>

<form id="frmScreen" mark validate="onchange">
<div style="HEIGHT: 132px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Accountant Details</strong> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Company Name
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyName" name="CompanyName" maxlength="45" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" onchange="AddressChanged()">
		</span> 
	</span>

	<span style="LEFT: 390px; POSITION: absolute; TOP: 31px" class="msgLabel">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton" onclick="btnDirectorySearch.onClick()">
	</span>

	<span style="TOP: 60px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Is in the directory?
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryYes" name="InDirectoryGroup" type="radio" value="1" onclick="InDirectoryChanged()">
			<label for="optInDirectoryYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 234px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryNo" name="InDirectoryGroup" type="radio" value="0" onclick="InDirectoryChanged()">
			<label for="optInDirectoryNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="TOP: 84px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Qualifications
		<span style="TOP: -3px; LEFT: 180px; POSITION: ABSOLUTE">
			<select id="cboQualifications" style="WIDTH: 283px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 108px" class="msgLabel">
		Number of Years Acting for Client
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtYearsActingForCustomer" maxlength="3" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span> 
	</span>
</div>

<div style="HEIGHT: 232px; LEFT: 10px; POSITION: absolute; TOP: 148px; WIDTH: 604px" class="msgGroup">
<!-- #include FILE="includes/thirdpartydetails.htm" -->
</div> 

<span style="TOP: 386px; LEFT: 10px; POSITION: ABSOLUTE">
	<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
	<span style="TOP: 0px; LEFT: 70px; POSITION: ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
</span>
</form>

<!-- #include FILE="attribs/DC191attribs.asp" -->
<!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;
	
var AccountantXML = null;
//var m_sAccountantGUID = "";
var m_aArgArray = null;
var m_sCallingScreen = "";
var m_sCustomerNumber = ""
var m_sCustomerVersionNumber = ""
var m_sEmploymentSequenceNumber = ""
var m_sEmploymentRelationshipInd = ""

/* EVENTS */

function frmScreen.btnCancel.onclick()
{
	window.close();
}

function frmScreen.btnOK.onclick()
{
	if (CommitChanges())
		window.close();
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
		
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	//next 3 lines added as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	m_aArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	<%/* Parameters for thirdpartydetails.asp */%>
	m_bIsPopup = true;
	m_sThirdPartyType = "1";
	
	var sGroups = new Array("Country","AccountantQualifications");
	objDerivedOperations = new DerivedScreen(sGroups);

	m_aArgArray = sArguments[4];
	m_sAccountantGUID = m_aArgArray[4];
	m_bCanAddToDirectory = m_aArgArray[5];
	m_sReadOnly = m_aArgArray[6];
	m_sCallingScreen = 	m_aArgArray[7];
	m_sCustomerNumber = m_aArgArray[8];
	m_sCustomerVersionNumber = m_aArgArray[9] ;
	m_sEmploymentSequenceNumber= m_aArgArray[10];
	m_sEmploymentRelationshipInd = m_aArgArray[11] ;
	
	SetThirdPartyDetailsMasks();
	SetMasks();
	
	ThirdPartyCustomise();

	Validation_Init();	
	Initialise();

	if(m_sReadOnly == "1")
		SetScreenToReadOnly();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function SetScreenToReadOnly()
{
	scScreenFunctions.SetScreenToReadOnly(frmScreen);
	frmScreen.btnDirectorySearch.disabled = true;
	frmScreen.btnOK.disabled = true;
	frmScreen.btnClear.disabled = true;
	frmScreen.btnPAFSearch.disabled = true;
	frmScreen.btnAddToDirectory.disabled = true;
}

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				bSuccess = SaveAccountant();
		}
		else
			bSuccess = false;
	return(bSuccess);
}

<% /*  Inserts default values into all fields */ %>
function DefaultFields()
{
	ClearFields(true, true);
	with (frmScreen)
	{
		txtYearsActingForCustomer.value = "";
		cboQualifications.selectedIndex = 0;
	}
}

function Initialise()
<% /* Initialises the screen */ %>
{
	PopulateCombos();
	PopulateScreen();

	if(m_sMetaAction != "Edit")
		DefaultFields();

	SetAvailableFunctionality();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<%/* Populates all combos on the screen */ %>
function PopulateCombos()
{
	PopulateTPTitleCombo();

	var sControlList = new Array(frmScreen.cboCountry, frmScreen.cboQualifications);
	objDerivedOperations.GetComboLists(sControlList);
}

<% /* Populates the screen with details of the item selected in  */ %>
function PopulateScreen()
{
	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//AccountantXML = new scXMLFunctions.XMLObject();
	AccountantXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	AccountantXML.CreateRequestTagFromArray(m_aArgArray,null);
	AccountantXML.CreateActiveTag("ACCOUNTANT");
	AccountantXML.CreateTag("ACCOUNTANTGUID", m_sAccountantGUID);

	if (m_sAccountantGUID == "")
	{
		m_sMetaAction = "Add";
		return;
	}
	else
		AccountantXML.RunASP(document,"GetAccountantDetails.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = AccountantXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[1] == ErrorTypes[0]))
	{
		m_sMetaAction = "Add";
		m_sAccountantGUID = "";
	}
	else if (ErrorReturn[0] == true)
	{
		m_sMetaAction = "Edit";

		if(AccountantXML.SelectTag(null, "ACCOUNTANT") != null)
			with (frmScreen)
			{
				m_sAccountantGUID = AccountantXML.GetTagText("ACCOUNTANTGUID"); 
				txtYearsActingForCustomer.value = AccountantXML.GetTagText("YEARSACTINGFORCUSTOMER");
				cboQualifications.value = AccountantXML.GetTagText("QUALIFICATIONS");
				txtCompanyName.value = AccountantXML.GetTagText("COMPANYNAME");
				txtContactForename.value = AccountantXML.GetTagText("CONTACTFORENAME");
				txtContactSurname.value = AccountantXML.GetTagText("CONTACTSURNAME");
				cboTitle.value = AccountantXML.GetTagText("CONTACTTITLE");
				<% /* PB 07/07/2006 EP543 Begin */ %>
				checkOtherTitleField();
				<% /* AW 01/08/2006 EP1060  */ %>
				txtTitleOther.value = AccountantXML.GetTagText("CONTACTTITLEOTHER");
				<% /* EP543 End */ %>

				objDerivedOperations.LoadCounty(AccountantXML);

				txtDistrict.value = AccountantXML.GetTagText("DISTRICT");
				txtFlatNumber.value = AccountantXML.GetTagText("FLATNUMBER");
				txtHouseName.value = AccountantXML.GetTagText("BUILDINGORHOUSENAME");
				txtHouseNumber.value = AccountantXML.GetTagText("BUILDINGORHOUSENUMBER");
				txtPostcode.value = AccountantXML.GetTagText("POSTCODE");
				txtStreet.value = AccountantXML.GetTagText("STREET");
				txtTown.value = AccountantXML.GetTagText("TOWN");
				cboCountry.value = parseInt(AccountantXML.GetTagText("COUNTRY"));
			}
			
			var TempXML = AccountantXML.ActiveTag;
			var ContactXML = AccountantXML.SelectTag(null, "CONTACTDETAILS");
			if(ContactXML != null)
				m_sXMLContact = ContactXML.xml;
			AccountantXML.ActiveTag = TempXML;
	}

	m_sDirectoryGUID = AccountantXML.GetTagText("DIRECTORYGUID");
	m_sThirdPartyGUID = AccountantXML.GetTagText("THIRDPARTYGUID");
	m_bDirectoryAddress = (m_sDirectoryGUID != "");

	ErrorTypes = null;
	ErrorReturn = null;
}

function SaveAccountant()
{
	var bSuccess = true;
	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTagFromArray(m_aArgArray,null)
	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");
	XML.CreateActiveTag("ACCOUNTANT");
	XML.CreateTag("YEARSACTINGFORCUSTOMER", frmScreen.txtYearsActingForCustomer.value);
	XML.CreateTag("QUALIFICATIONS", frmScreen.cboQualifications.value);

	if (m_sMetaAction == "Edit")
	{
		XML.CreateTag("ACCOUNTANTGUID", m_sAccountantGUID);
		
		<% /* Retrieve the original third party/directory GUIDs */%>
		var sOriginalThirdPartyGUID = AccountantXML.GetTagText("THIRDPARTYGUID");
		var sOriginalDirectoryGUID = AccountantXML.GetTagText("DIRECTORYGUID");
		
		<% /*  Only retrieve the address/contact details GUID if we are updating an existing third party record */ %>
		var sAddressGUID = (sOriginalThirdPartyGUID != "") ? AccountantXML.GetTagText("ADDRESSGUID") : "";
		var sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? AccountantXML.GetTagText("CONTACTDETAILSGUID") : "";
	}

	<% /*  If changing from a directory to a third party, or vice-versa, the old directory/thirdparty GUID
	 should still be specified to alert the middler tier to the fact that the old link needs deleting */ %>
	XML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID != "" ? m_sDirectoryGUID : sOriginalDirectoryGUID);
	XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
	if (!m_bDirectoryAddress)
	{
		<% /*  Store the third party details */ %>
		XML.CreateActiveTag("THIRDPARTY");
		XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
		XML.CreateTag("THIRDPARTYTYPE", 1); // Accountant
		XML.CreateTag("COMPANYNAME", frmScreen.txtCompanyName.value);

		<% /*  Address */ %>
		XML.CreateTag("ADDRESSGUID", sAddressGUID);
		SaveAddress(XML, sAddressGUID);

		<% /*  Contact Details */ %>
		XML.SelectTag(null, "THIRDPARTY");
		XML.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
		SaveContactDetails(XML, sContactDetailsGUID);
	}
	<% /* Accountant Details - amended by DPF 26/06/02 - BMIDS00077 */ %>
	if (m_sCallingScreen == "DC181" )
	{
		//XML.CreateActiveTag("EMPLOYMENT"); - needs to be EMPLOYEDDETAILS
		XML.CreateActiveTag("EMPLOYEDDETAILS");
		XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber );
		XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber );
		XML.CreateTag("EMPLOYMENTSEQUENCENUMBER", m_sEmploymentSequenceNumber );
		XML.CreateTag("EMPLOYMENTRELATIONSHIPIND", m_sEmploymentRelationshipInd);
		//XML.CreateTag("SHARESOWNEDIND", '1');
	
		if (m_sMetaAction == "Edit")
			XML.CreateTag("ACCOUNTANTGUID", m_sAccountantGUID);
	} 
		
	<% /*  Save the details */ %>
	// 	XML.RunASP(document,"SaveAccountantDetails.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"SaveAccountantDetails.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	bSuccess = XML.IsResponseOK();

	if (bSuccess)
	{
		<% /*  Return the new DirectoryGUID and OrganisationID to the calling screen */ %>
		var sReturn = new Array();
		sReturn[0] = IsChanged();
		if (m_sMetaAction == "Edit")
			sReturn[1] = m_sAccountantGUID;
		else
			sReturn[1] = XML.GetTagText("ACCOUNTANTGUID");
		sReturn[2] = frmScreen.txtCompanyName.value;

		window.returnValue = sReturn;
	}

	XML = null;
	return(bSuccess);
}
-->
</script>
<!-- #include FILE="includes/thirdpartydetails.asp" -->
</body>
</html>


