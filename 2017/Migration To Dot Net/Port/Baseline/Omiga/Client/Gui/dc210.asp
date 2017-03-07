<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:		DC210.asp
Copyright:		Copyright © 1999 Marlborough Stirling

Description:	Property Address Screen.
				This frame captures the new property address details 
				required to complete the mortgage application.
				For a further advance, remortgage or transfer of equity, 
				the customer registration process will have set up the 
				property address, but not the telephone number and 
				arrangements for access.
LD		23/05/02	SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		18/01/00	Created, based on DC065.asp.
RF		01/02/00	Reworked for IE5 and performance.
AY		14/02/00	Change to msgButtons button types
AD		28/02/00	Completed work.
AD		21/03/00	SYS0448 - Tabbing order changed from postcode-PAF search
AY		31/03/00	New top menu/scScreenFunctions change
MH		03/05/00	SYS0618 Postcode validation
IVW		04/05/00	SYS0654 - text change and focus field default
IVW		09/05/00	SYS0617 - remorting of screen, field resizea etc . see aqr for full list.
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
MC		08/06/00	SYS0886 Default Vendor Address to New Property Address
BG		25/07/00	SYS1199 - if mortgage type is further advance, remortgage or transfer of equity
					address details are populated with current address and screen set to readonly.
BG		11/09/00	SYS1199 fixed error message appearing on clicking Vendor Details Button.
CL		05/03/01	SYS1920 Read only functionality added
CP		28/04/01	SYS2050 Critical Data functionality added
GD		10/05/01	SYS2050	Roll back Critical data code as above
SA		30/05/01	SYS0932 Clear & PAF buttons should be disabled + focus on OK.
DC      20/07/01    SYS2038 Critical Data functionality ROLLED FORWARD AGAIN
JR		10/10/01	SYSOmiplus24 Include fields for COUNTRYCODE, AREACODE and EXTENSIONNUMBER
JLD		10/12/01	SYS2806 completeness check routing
JLD		17/12/01	SYS2620 cancel routing error

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
DRC     23/09/02    BMIDS00481 - needed a declaration for m_bIsPopup used in the PAFSearch include - see core AQR SYS4930 
GHun	11/11/02	BMIDS00444 - changed the way Remortgages and Further Advances set the address record
GHun	19/11/02	BMIDS01016 - MetaAction set incorrectly on GetRemortgageAddress and GetFurtherAdvanceAddress
SA		20/11/02	BMIDS01021 - Screen Rules Processing added
GHun	20/11/02	BMIDS01009 - Address does not always save for Further Advances and Remortgage
MV		06/12/2002	BMIDS0020 - Modified PAF Search tab order
GHun	06/01/2003	BM0062 - Addresses with NULL country codes use default country to avoid PAF errors
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
MF		04/08/2005	MAR020		Changed routing 
MF		08/08/2005	MAR20		Parameterised routing for Global Parameters
								NewPropertySummary & ThirdPartySummary
PJO     10/11/2005  MAR400      Add <SELECT> to country combo		
SD		29/11/2005	   Check NewProperty PostCode to see if it's Mainland UK
DRC     02/12/2005  MAR794      New Property Mainland check
PE		14/12/2005	MAR868		When global parameter "NewPropertySummary" is 0, the screen flow is changed.
PE		20/02/2006	MAR1224		Paf search results not being saved
PE		27/02/2006	MAR1328		Country details not as per test script
PE		15/03/2006	MAR1387		Enhance DC210 to default the arrangement for access details
PE		29/03/2006	MAR1481		Default access details to the details of the main applicant of the Remortgaged Property 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
TLiu	02/09/2005	   Changed layout for Flat No., House Name & No. - MAR38

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom Specific History:

Prog	Date		AQR			Description
AShaw	01/02/2007	EP2_1002	Move Business connection question from DC200.
								Rename form.
AShaw	02/02/2007	EP2_1184	Business connection question null value issue.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/

%>
<HEAD>
<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>

<% // Specify Forms Here %>
<form id="frmToDC220" method="post" action="dc220.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC215" method="post" action="dc215.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC201" method="post" action="DC201.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC200" method="post" action="DC200.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% // Specify Screen Layout Here %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground1" style="HEIGHT: 210px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<% // PostCode %>
	<span style="LEFT: 420px; POSITION: absolute; TOP: 5px">
		<input id="btnClear" value="Clear" type="button" style="WIDTH: 100px" class="msgButton"> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Postcode
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtPostcode" name="Postcode" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgTxtUpper" onchange="ClearAddressPAFIndicator()"> 
		</span> 
	</span>

	<% // Flat Number %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Flat No./ Name
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtFlatNo" name="FlatNumber" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgTxt" onchange="ClearAddressPAFIndicator()"> 
		</span> 
	</span>

	<% // House Name %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
		House No. &amp; Name
		<% // House Number %>
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtHouseNumber" name="HouseNumber" maxlength="10" style="POSITION: absolute; WIDTH: 45px" class="msgTxt"
				 onchange="ClearAddressPAFIndicator()"> 
		</span> 
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtHouseName" name="HouseName" maxlength="40" style="POSITION: absolute; WIDTH: 230px" class="msgTxt"
				 onchange="ClearAddressPAFIndicator()"> 
		</span>
	</span>
	
	
	<span style="LEFT: 420px; POSITION: absolute; TOP: 29px">
		<input id="btnPAFSearch" value="P.A.F.Search" type="button" style="WIDTH: 100px" class="msgButton"> 
	</span>
	<% // Street %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 85px" class="msgLabel">
		Street
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtStreet" name="Street" maxlength="40" style="POSITION: absolute; WIDTH: 280px" class="msgTxt"
				 onchange="ClearAddressPAFIndicator()"> 
		</span> 
	</span>
	<% // District %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 110px" class="msgLabel">
		District
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtDistrict" name="District" maxlength="40" style="POSITION: absolute; WIDTH: 280px" class="msgTxt"
				 onchange="ClearAddressPAFIndicator()"> 
		</span> 
	</span>
	<% // Town %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 135px" class="msgLabel">
		Town
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtTown" name="Town" maxlength="40" style="POSITION: absolute; WIDTH: 280px" class="msgTxt"
				 onchange="ClearAddressPAFIndicator()"> 
		</span> 
	</span>
	<% // County %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 160px" class="msgLabel">
		County
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtCounty" name="County" maxlength="40" 
				style="POSITION: absolute; WIDTH: 280px" class="msgTxt"> 
		</span> 
	</span>
	<% // Country %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 185px" class="msgLabel">
		Country
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<select id="cboCountry" name="Country" style="WIDTH: 280px" class="msgCombo"
				 onchange="ClearAddressPAFIndicator()">
			</select>
 
		</span> 
	</span> 
</div>
		
<div id="divBackground2" style="HEIGHT: 240px; LEFT: 10px; POSITION: absolute; TOP: 275px; WIDTH: 604px" class="msgGroup">
	<% // Access Arrangements %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 7px" class="msgLabel">
		Arrangements for access
		<span style="LEFT: 130px; POSITION: absolute; TOP: 0px">
			<select id="cboAccessArrangements" name="AccessArrangements" 
				style="WIDTH: 150px" class="msgCombo">
			</select>
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Name of Contact
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtAccessContactName" name="AccessContactName" maxlength="40" 
				style="POSITION: absolute; WIDTH: 280px" class="msgTxt"> 
		</span> 
	</span>
	
	<!-- JR - SYSOmiplus24 -->
	<span style="LEFT: 135px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Country
	</span>
	
	<span style="LEFT: 135px; POSITION: absolute; TOP: 73px" class="msgLabel">
		Code
	</span>
	
	<span style="LEFT: 180px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Area Code
	</span>
	
	<span style="LEFT: 235px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Telephone Number
	</span>
	
	<span style="LEFT: 343px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Extension
	</span>
	
	<span style="LEFT: 343px; POSITION: absolute; TOP: 73px" class="msgLabel">
		Number
	</span>	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 90px" class="msgLabel">
		Contact's Tel.
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtCountryCode" name="CountryCode" maxlength="3" 
				style="POSITION: absolute; WIDTH: 40px" class="msgTxt"> 
		</span>		
		<span style="LEFT: 172px; POSITION: absolute; TOP: -3px">
			<input id="txtAreaCode" name="AreaCode" maxlength="6" 
				style="POSITION: absolute; WIDTH: 55px" class="msgTxt"> 
		</span>
		<span style="LEFT: 229px; POSITION: absolute; TOP: -3px">
			<input id="txtAccessTelephoneNumber" name="AccessTelephoneNumber" maxlength="30" 
				style="POSITION: absolute; WIDTH: 105px" class="msgTxt"> 
		</span> 
		<span style="LEFT: 336px; POSITION: absolute; TOP: -3px">
			<input id="txtTelephoneExtensionNum" name="ExtensionNumber" maxlength="10" 
				style="POSITION: absolute; WIDTH: 75px" class="msgTxt"> 
		</span>
	</span>
		
	
	<% // Other Arrangements %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 120px" class="msgLabel">
		Details
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px"><TEXTAREA class=msgTxt id=txtOtherArrangements rows=5 style="POSITION: absolute; WIDTH: 280px"></TEXTAREA>
		</span>
	</span>

	<% // Vendor Details button %>
	<span style="LEFT: 420px; POSITION: absolute; TOP: 4px">
		<input id="btnVendorDetails" value="Vendor Details" type="button" 
			style="WIDTH: 100px" class="msgButton"> 
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 200px" class="msgLabel">
		Do you have any business connection<br>with or are you related to the vendor?
		<span style="LEFT: 310px; POSITION: absolute; TOP: 3px">
			<input id="optBusinessConnectionYes" name="BusinessConnectionGroup" type="radio" value="1"><label for="optBusinessConnectionYes" class="msgLabel">Yes</label> 
		</span> 
		<span style="LEFT: 375px; POSITION: absolute; TOP: 3px">
			<input id="optBusinessConnectionNo" name="BusinessConnectionGroup" type="radio" value="0"><label for="optBusinessConnectionNo" class="msgLabel">No</label> 
		</span>		 
	</span>	

</div> 
</form>

<% // Main Buttons %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 520px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% // File containing field attributes (remove if not required) %>
<!--  #include FILE="attribs/dc210attribs.asp" -->
<!--  #include FILE="includes/pafsearch.asp" -->

<% // Specify Code Here %>
<script language="JScript">
<!--	// JScript Code
var m_bReadOnlyIndicator  = null;
var m_bPAFIndicator = false;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sAddressGUID = "";
var m_sMetaAction = "";
var scScreenFunctions;
var m_sXMLInAddress = null;
var m_sMortgageApplicationValue = "";
var m_sCustomerNumber = null;
var m_sCustomerVersionNumber = null;
var m_blnReadOnly = false;
// AQR BMIDS00481 DRC - decalration needed in PAFSearch
var m_bIsPopup = false;
var m_XMLTypeOfMortgage = null;	<% /* BMIDS00444 */ %>
var m_XMLINGNonUKPostCodes = null;


function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		// MAR868
		// Peter Edney - 14/12/2005
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var bNewPropertySummary = XML.GetGlobalParameterBoolean(document, "NewPropertySummary");					
		if (bNewPropertySummary)
		{
			frmToMN060.submit();
		}
		else
		{
			frmToDC200.submit();
		}

}

function btnSubmit.onclick()
{	
	if (DoOKProcessing())
	{

		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else{
			<% /* MF Read Parameter to decide route */ %>
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var bNewPropSummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");	
			if(bNewPropSummary){
				frmToDC201.submit();
			}else{
				frmToDC220.submit();
			}
		}				
	}
}

function frmScreen.btnClear.onclick()
{
	ClearFields(true,false);
	FlagChange(true);
	frmScreen.txtPostcode.focus();
}

function frmScreen.btnPAFSearch.onclick()
{
	with (frmScreen)
		m_bPAFIndicator = PAFSearch(txtPostcode,txtHouseName,txtHouseNumber,txtFlatNo,txtStreet,txtDistrict,
			txtTown,txtCounty,cboCountry);
			
		<% // MAR1224 - Paf search results not being saved %>
		<% // Peter Edney - 20/02/2006 %>
		if(m_bPAFIndicator) FlagChange(m_bPAFIndicator);
}

function frmScreen.btnVendorDetails.onclick()
{
	if (DoOKProcessing())
		StoreNewPropertyAddressData();
		frmToDC215.submit();
}

function frmScreen.txtOtherArrangements.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtOtherArrangements", 255, true);
}

function frmScreen.cboAccessArrangements.onchange()
{
	UnhideOtherArrangements();			
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

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Property Address","DC210",scScreenFunctions);

	RetrieveContextData();

	SetMasks();
	Validation_Init();
	Initialise(true);
	//scScreenFunctions.SetFocusToFirstField(frmScreen);
	frmScreen.txtPostcode.focus();

	<% /* frmScreen.btnVendorDetails.style.visibility = "hidden"; */%>
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC210");
	if (m_blnReadOnly == true) 
	{
		m_bReadOnlyIndicator = "1";
		<%/*SYS0932 Clear & PAF buttons should be disabled + focus on OK.*/%>
		frmScreen.btnClear.disabled =true;
		frmScreen.btnPAFSearch.disabled =true;
		btnSubmit.focus();	
	}
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function ClearAddressPAFIndicator()
{
	m_bPAFIndicator	= false;
}

function ClearFields(bAddress, bOther)
// Clears all fields of data
{
	with (frmScreen)
	{
		if (bAddress)
		{
			txtFlatNo.value = "";
			txtHouseName.value = "";
			txtHouseNumber.value = "";
			txtPostcode.value = "";
			txtStreet.value = "";
			txtDistrict.value = "";
			txtTown.value = "";
			txtCounty.value = "";

			<% // MAR1328 - Peter Edney - 27/02/2006 %>
			<% // Get the global to confirm if the Country should be disabled. %>
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var bNewPropertyAddDisableCountry = XML.GetGlobalParameterBoolean(document, "NewPropertyAddDisableCountry");						
			if(bNewPropertyAddDisableCountry == false) 
			{		
				cboCountry.selectedIndex = 0;
			}
			
		}

		if (bOther)
		{
			cboAccessArrangements.selectedIndex = 0;
			txtAccessContactName.value  = "";
			txtAccessTelephoneNumber.value  = "";
			txtOtherArrangements.value = "";
			//JR - Omiplus24
			txtCountryCode.value = "";
			txtAreaCode.value = "";
			txtTelephoneExtensionNum.value = "";
			
			
		}
	}
}

function CommitChanges(sTransactionType)
{
	function SaveAddress(XML, sAddressGUID)
	{
		XML.CreateActiveTag("ADDRESS");
		XML.CreateTag("ADDRESSGUID", sAddressGUID);
		XML.CreateTag("BUILDINGORHOUSENAME", frmScreen.txtHouseName.value);
		XML.CreateTag("BUILDINGORHOUSENUMBER", frmScreen.txtHouseNumber.value);
		XML.CreateTag("FLATNUMBER", frmScreen.txtFlatNo.value);
		XML.CreateTag("STREET", frmScreen.txtStreet.value);
		XML.CreateTag("DISTRICT", frmScreen.txtDistrict.value);
		XML.CreateTag("TOWN", frmScreen.txtTown.value);
		XML.CreateTag("COUNTY", frmScreen.txtCounty.value);
		XML.CreateTag("COUNTRY", frmScreen.cboCountry.value);
		XML.CreateTag("POSTCODE", frmScreen.txtPostcode.value);
		XML.CreateTag("PAFINDICATOR", m_bPAFIndicator ? "1" : "0");
	}

	var bSuccess = true;

	// EP2_1002 - Save BUSINESSCONNECTION field back to NEWPROPERTY table.
	// EP2_1184 - Only if not null.
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if (scScreenFunctions.GetRadioGroupValue(frmScreen,"BusinessConnectionGroup") != null)
	{
		XML.CreateRequestTag(window);
		var xn = XML.XMLDocument.documentElement;
		xn.setAttribute("CRUD_OP","UPDATE");
		xn.setAttribute("SCHEMA_NAME","omCRUD");
		xn.setAttribute("ENTITY_REF","NEWPROPERTY");
		var xe = XML.XMLDocument.createElement("NEWPROPERTY");
		xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		xe.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		xe.setAttribute("BUSINESSCONNECTION", scScreenFunctions.GetRadioGroupValue(frmScreen,"BusinessConnectionGroup"));
		xn.appendChild(xe);
		XML.RunASP(document, "omCRUDIf.asp");
	}


	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	if (sTransactionType == "Delete")
	{
		<%/* All address fields are empty, so try to delete this address */%>
		XML.CreateRequestTag(window);
		XML.CreateActiveTag("NEWPROPERTYADDRESS");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("ADDRESSGUID", m_sAddressGUID);
		//BMIDS01021 Screen Rules Processing
		//XML.RunASP(document, "DeleteNewPropertyAddress.asp");
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
				XML.RunASP(document, "DeleteNewPropertyAddress.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		bSuccess = XML.IsResponseOK();
	}
	else
	{
		XML.CreateRequestTag(window);
		//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308 ROLL BACK GD - 10/05/2001 SYS2050
		//start OPERATION  CP_28/04/2001 AQR SYS2050
				
		if (sTransactionType != "Add")
		{
			XML.SetAttribute("OPERATION","CriticalDataCheck"); 
		}
		//end OPERATION  CP_28/04/2001 AQR SYS2050
	
		//END ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  ROLLBACK GD - 10/05/2001
		XML.CreateActiveTag("NEWPROPERTYADDRESS");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("ARRANGEMENTSFORACCESS", frmScreen.cboAccessArrangements.value);
		XML.CreateTag("ACCESSCONTACTNAME", frmScreen.txtAccessContactName.value);
		XML.CreateTag("ACCESSTELEPHONENUMBER", frmScreen.txtAccessTelephoneNumber.value);
		XML.CreateTag("OTHERARRANGEMENTSFORACCESS", frmScreen.txtOtherArrangements.value);

		//JR - SYSOmiplus24, include CountryCode, AreaCode and ExtensionNumber
		XML.CreateTag("COUNTRYCODE", frmScreen.txtCountryCode.value);
		XML.CreateTag("AREACODE", frmScreen.txtAreaCode.value);
		XML.CreateTag("EXTENSIONNUMBER", frmScreen.txtTelephoneExtensionNum.value);
		
		<%/* Address */%>
		XML.CreateTag("ADDRESSGUID", m_sAddressGUID);
		SaveAddress(XML, m_sAddressGUID);

		<%/* Call the business object */%>			
		if (sTransactionType == "Add")
		{
			//BMIDS01021 Add Screen Rules Processing
			//XML.RunASP(document, "CreateNewPropertyAddress.asp");
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
					XML.RunASP(document, "CreateNewPropertyAddress.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

		}
		else
		{	
			//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308 ROLL BACK GD - 10/05/2001 SYS2050

			//start OPERATION  CP_28/04/2001 AQR SYS2050			
			XML.SelectTag(null,"REQUEST");
			XML.CreateActiveTag("CRITICALDATACONTEXT");
			XML.SetAttribute("APPLICATIONNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null));
			XML.SetAttribute("APPLICATIONFACTFINDNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null));
			XML.SetAttribute("SOURCEAPPLICATION","Omiga");
			XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
			XML.SetAttribute("ACTIVITYINSTANCE","1");
			XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
			XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
			XML.SetAttribute("COMPONENT","omApp.NewPropertyBO");
			XML.SetAttribute("METHOD","UpdateNewPropertyAddress");
			window.status = "Critical Data Check - please wait";
			
			//BMIDS01021 Add SCreen Rules Processing
			//XML.RunASP(document,"OmigaTMBO.asp");
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
			bSuccess = XML.IsResponseOK();		
			
			//XML.RunASP(document, "UpdateNewPropertyAddress.asp");
			//end OPERATION  CP_28/04/2001 AQR SYS2050
			//END ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  ROLLBACK GD - 10/05/2001

				
		}	
		bSuccess = XML.IsResponseOK();			
	}
				
	XML = null;
	return(bSuccess);		
}

function DefaultFields()
{
	ClearFields(true,true);
}

function DoOKProcessing()
{
	var bSuccess = false;
	var bDoSave = false;
	var bAllAddressFieldsEmpty = false;
	var sTransactionType = "Edit";

	if(scScreenFunctions.RestrictLength(frmScreen, "txtOtherArrangements", 255, true))
		return false;

	with (frmScreen)
		bAllAddressFieldsEmpty = (txtHouseNumber.value + txtHouseName.value + txtFlatNo.value + txtStreet.value +
			txtDistrict.value + txtTown.value + txtCounty.value + txtPostcode.value == "");
			
	if(frmScreen.onsubmit())
	{
		bSuccess = true;
		bDoSave=false;
		
		if ((m_bReadOnlyIndicator != "1") & IsChanged())
		{
			
			bSuccess = scScreenFunctions.ValidatePostcode(frmScreen.txtPostcode);				
			
			if (bSuccess) 
				bSuccess = IsMainlandPostCode(frmScreen.txtPostcode);				
						  
			
			if (bSuccess) 
			{
				if ((bAllAddressFieldsEmpty) & (m_sMetaAction == "Edit"))
				{
					bDoSave = confirm("Because all address fields are blank, continuing will delete this property " +
								"address. Do you wish to continue?");
					sTransactionType = "Delete";
				}
				else if ((bAllAddressFieldsEmpty) & (m_sMetaAction == "Add"))
				{
					bDoSave = false;
					if (!confirm("Because all address fields are blank, continuing will not actually save any data " +
								"to the database. Do you wish to continue?"))
						bSuccess = false;
				}
				else
				{
					bDoSave = true;
					sTransactionType = m_sMetaAction;
				}
			}
		}
	}

	if (bDoSave)
		bSuccess = CommitChanges(sTransactionType);
	return(bSuccess);
}

function GetComboLists()
{
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("Country","AccessArrangements","TypeOfMortgage","INGNonUKPostCodes");

	if(XML.GetComboLists(document,sGroupList))
	{
		// Populate Country combo - PJO 10/11/2005 MAR400 Add <SELECT>
		blnSuccess = XML.PopulateCombo(document, frmScreen.cboCountry, "Country", true);
		blnSuccess = blnSuccess & XML.PopulateCombo(document, frmScreen.cboAccessArrangements, "AccessArrangements", true);
		<% /* BMIDS00444 */ %>
		m_XMLTypeOfMortgage = XML.GetComboListXML("TypeOfMortgage");
		m_XMLINGNonUKPostCodes = XML.GetComboListXML("INGNonUKPostCodes");
						
		if(!blnSuccess) scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}

	XML = null;
}

function Initialise()
{
	GetComboLists();
	DefaultFields();
	PopulateScreen();
	UnhideOtherArrangements();

	if (m_bReadOnlyIndicator == "1")
	{
		scScreenFunctions.SetCollectionToReadOnly(divBackground1);
		scScreenFunctions.SetCollectionToReadOnly(divBackground2);
	}

	<% // MAR1328 - Peter Edney - 02/03/06 %>
	<% // Get the global to confirm if the Country should be disabled. %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bNewPropertyAddDisableCountry = XML.GetGlobalParameterBoolean(document, "NewPropertyAddDisableCountry");						
	if((bNewPropertyAddDisableCountry == true) && (frmScreen.cboCountry.value == "")) 
	{		
		scScreenFunctions.SetComboOnValidationType(frmScreen,"cboCountry","UK");
		frmScreen.cboCountry.disabled = true;
	}
	
}

<% /* BMIDS00444 no longer used - updated version below.
function PopulateScreen()
{	
	var m_sMortgageApplicationValue = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue");
	
	if (m_sMortgageApplicationValue == "10")
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window);
		XML.CreateActiveTag("NEWPROPERTYADDRESS");
		XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.RunASP(document, "GetNewPropertyAddress.asp");
	
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = XML.CheckResponse(ErrorTypes);
				
		if(ErrorReturn[1] == ErrorTypes[0])
			// Record not found
			m_sMetaAction = "Add";
			
		else if (ErrorReturn[0] == true)
			// Record found
		{
			m_sMetaAction = "Edit";
			XML.SelectTag(null, "NEWPROPERTYADDRESS");
			frmScreen.txtPostcode.value	= XML.GetTagText("POSTCODE");
			frmScreen.txtFlatNo.value = XML.GetTagText("FLATNUMBER");
			frmScreen.txtHouseName.value = XML.GetTagText("BUILDINGORHOUSENAME");
			frmScreen.txtHouseNumber.value = XML.GetTagText("BUILDINGORHOUSENUMBER");
			frmScreen.txtStreet.value = XML.GetTagText("STREET");
			frmScreen.txtDistrict.value = XML.GetTagText("DISTRICT");
			frmScreen.txtTown.value	= XML.GetTagText("TOWN");
			frmScreen.txtCounty.value = XML.GetTagText("COUNTY");
			frmScreen.cboCountry.value = XML.GetTagText("COUNTRY");
			frmScreen.cboAccessArrangements.value = XML.GetTagText("ARRANGEMENTSFORACCESS");
			frmScreen.txtAccessContactName.value = XML.GetTagText("ACCESSCONTACTNAME");
			frmScreen.txtAccessTelephoneNumber.value = XML.GetTagText("ACCESSTELEPHONENUMBER");
			//JR - Omiplus24
			frmScreen.txtTelephoneExtensionNum.value = XML.GetTagText("EXTENSIONNUMBER");
			frmScreen.txtCountryCode.value = XML.GetTagText("COUNTRYCODE");
			frmScreen.txtAreaCode.value = XML.GetTagText("AREACODE");
			//End
			frmScreen.txtOtherArrangements.value = XML.GetTagText("OTHERARRANGEMENTSFORACCESS");
			m_bPAFIndicator = (XML.GetTagText("PAFINDICATOR") == "1");
			m_sAddressGUID = XML.GetTagText("ADDRESSGUID");
		}
	}
	// BG 25/07/00 SYS1199 if loan type is remortgage, further advance or Transfer of Equity, 
	// populate screen with current address and disable screen and buttons.
	else
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window);
		XML.CreateActiveTag("NEWPROPERTYADDRESS");
		XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.RunASP(document, "GetNewPropertyAddress.asp");
	
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = XML.CheckResponse(ErrorTypes);
				
		if(ErrorReturn[1] == ErrorTypes[0])
			// Record not found
		{	
			m_sMetaAction = "Add";
						
			XML = null;			
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber");
			var m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber");
			XML.CreateRequestTag(window,"SEARCH");
			XML.CreateActiveTag("SEARCH");
			XML.CreateActiveTag("CUSTOMER");
			XML.CreateTag("CUSTOMERNUMBER",m_sCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER",m_sCustomerVersionNumber);

			XML.RunASP(document,"GetCustomerDetails.asp");

			if(XML.IsResponseOK())
			XML.CreateTagList("CUSTOMERADDRESS");
			
			if(XML.GetTagText("ADDRESSTYPE") == "1")
			for(var nLoop = 0;nLoop < XML.ActiveTagList.length;nLoop++)
			{
				XML.SelectTagListItem(nLoop);
				if(XML.GetTagText("ADDRESSTYPE") == "1")
				{
					frmScreen.txtPostcode.value = XML.GetTagText("POSTCODE");
					frmScreen.txtFlatNo.value = XML.GetTagText("FLATNUMBER");
					frmScreen.txtHouseName.value = XML.GetTagText("BUILDINGORHOUSENAME");
					frmScreen.txtHouseNumber.value = XML.GetTagText("BUILDINGORHOUSENUMBER");
					frmScreen.txtStreet.value = XML.GetTagText("STREET");
					frmScreen.txtDistrict.value = XML.GetTagText("DISTRICT");
					frmScreen.txtTown.value = XML.GetTagText("TOWN");
					frmScreen.txtCounty.value = XML.GetTagText("COUNTY");
					frmScreen.cboCountry.value = XML.GetTagText("COUNTRY");
					m_bPAFIndicator = (XML.GetTagText("PAFINDICATOR") == "1");
					m_sAddressGUID = XML.GetTagText("ADDRESSGUID");
					scScreenFunctions.SetCollectionToReadOnly(divBackground1);
					frmScreen.btnPAFSearch.disabled =true;
					frmScreen.btnClear.disabled =true;
					frmScreen.btnVendorDetails.disabled =false;
			
				}

			}
		}	
		else if (ErrorReturn[0] == true)
			// Record found
		{
			m_sMetaAction = "Edit";
			XML.SelectTag(null, "NEWPROPERTYADDRESS");
			frmScreen.cboAccessArrangements.value = XML.GetTagText("ARRANGEMENTSFORACCESS");
			frmScreen.txtAccessContactName.value = XML.GetTagText("ACCESSCONTACTNAME");
			frmScreen.txtAccessTelephoneNumber.value = XML.GetTagText("ACCESSTELEPHONENUMBER");
			frmScreen.txtOtherArrangements.value = XML.GetTagText("OTHERARRANGEMENTSFORACCESS");
			//JR - Omiplus24
			frmScreen.txtCountryCode.value = XML.GetTagText("COUNTRYCODE");
			frmScreen.txtAreaCode.value = XML.GetTagText("AREACODE");
			frmScreen.txtTelephoneExtensionNum.value = XML.GetTagText("EXTENSIONNUMBER");
			//End
			
			
			XML = null;			
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber");
			var m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber");
			XML.CreateRequestTag(window,"SEARCH");
			XML.CreateActiveTag("SEARCH");
			XML.CreateActiveTag("CUSTOMER");
			XML.CreateTag("CUSTOMERNUMBER",m_sCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER",m_sCustomerVersionNumber);

			XML.RunASP(document,"GetCustomerDetails.asp");

			if(XML.IsResponseOK())
			XML.CreateTagList("CUSTOMERADDRESS");
			
			if(XML.GetTagText("ADDRESSTYPE") == "1")
			for(var nLoop = 0;nLoop < XML.ActiveTagList.length;nLoop++)
			{
				XML.SelectTagListItem(nLoop);
				if(XML.GetTagText("ADDRESSTYPE") == "1")
				{
					frmScreen.txtPostcode.value = XML.GetTagText("POSTCODE");
					frmScreen.txtFlatNo.value = XML.GetTagText("FLATNUMBER");
					frmScreen.txtHouseName.value = XML.GetTagText("BUILDINGORHOUSENAME");
					frmScreen.txtHouseNumber.value = XML.GetTagText("BUILDINGORHOUSENUMBER");
					frmScreen.txtStreet.value = XML.GetTagText("STREET");
					frmScreen.txtDistrict.value = XML.GetTagText("DISTRICT");
					frmScreen.txtTown.value = XML.GetTagText("TOWN");
					frmScreen.txtCounty.value = XML.GetTagText("COUNTY");
					frmScreen.cboCountry.value = XML.GetTagText("COUNTRY");
					m_bPAFIndicator = (XML.GetTagText("PAFINDICATOR") == "1");
					m_sAddressGUID = XML.GetTagText("ADDRESSGUID");
					scScreenFunctions.SetCollectionToReadOnly(divBackground1);
					frmScreen.btnPAFSearch.disabled =true;
					frmScreen.btnClear.disabled =true;
					frmScreen.btnVendorDetails.disabled =false;
			
				}

			}
		}
	}	

	XML = null;
	ErrorTypes = null;
	ErrorReturn = null;
}
*/ %>

<% /* BMIDS00444 */ %>
function PopulateScreen()
{	
	<% /* Look for a NewProperty address */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window);
	XML.CreateActiveTag("NEWPROPERTYADDRESS");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	XML.RunASP(document, "GetNewPropertyAddress.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
				
	if(ErrorReturn[1] == ErrorTypes[0])
	{
		<%/* NewPropertyAddress not found */%>
		m_sMetaAction = "Add";
		
		var m_sMortgageApplicationValue = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue");
		
		<% /* Determine Type Of Mortgage */ %>
		var XMLType = m_XMLTypeOfMortgage.selectSingleNode("LISTNAME/LISTENTRY[VALUEID[.='" + m_sMortgageApplicationValue + "']]");
	
		<% /* For new loans the address will be blank so there is nothing to do */ %>
		
		if (XMLType.selectSingleNode("VALIDATIONTYPELIST/VALIDATIONTYPE[.='R']") != null)
		{
			<% /* Remortgage */ %>
			GetRemortgageAddress();
			
			<% // MAR1387 %>
			frmScreen.cboAccessArrangements.value = 1;
						
		}	
		else
		{
			XMLType = XMLType.selectSingleNode("VALIDATIONTYPELIST[VALIDATIONTYPE[.='M']]");
			if (XMLType != null)
			{
				if (XMLType.selectSingleNode("VALIDATIONTYPE[.='F' or .='T']") != null)
					<% /* Further Advance, SPL, EP, or ToE */ %>
					GetFurtherAdvanceAddress();
			}	
		}
	}
	else if (ErrorReturn[0] == true)
	{
		<%/* NewPropertyAddress found */%>
		m_sMetaAction = "Edit";
		XML.SelectTag(null, "NEWPROPERTYADDRESS");
		frmScreen.txtPostcode.value	= XML.GetTagText("POSTCODE");
		frmScreen.txtFlatNo.value = XML.GetTagText("FLATNUMBER");
		frmScreen.txtHouseName.value = XML.GetTagText("BUILDINGORHOUSENAME");
		frmScreen.txtHouseNumber.value = XML.GetTagText("BUILDINGORHOUSENUMBER");
		frmScreen.txtStreet.value = XML.GetTagText("STREET");
		frmScreen.txtDistrict.value = XML.GetTagText("DISTRICT");
		frmScreen.txtTown.value	= XML.GetTagText("TOWN");
		frmScreen.txtCounty.value = XML.GetTagText("COUNTY");
		<% /* BM0062 Only assign country combo value if it is not blank */ %>
		if (XML.GetTagText("COUNTRY").length > 0)
		<% /* BM0062 End */ %>
			frmScreen.cboCountry.value = XML.GetTagText("COUNTRY");
		frmScreen.cboAccessArrangements.value = XML.GetTagText("ARRANGEMENTSFORACCESS");
		frmScreen.txtAccessContactName.value = XML.GetTagText("ACCESSCONTACTNAME");
		frmScreen.txtAccessTelephoneNumber.value = XML.GetTagText("ACCESSTELEPHONENUMBER");
		//JR - Omiplus24
		frmScreen.txtTelephoneExtensionNum.value = XML.GetTagText("EXTENSIONNUMBER");
		frmScreen.txtCountryCode.value = XML.GetTagText("COUNTRYCODE");
		frmScreen.txtAreaCode.value = XML.GetTagText("AREACODE");
		//End
		frmScreen.txtOtherArrangements.value = XML.GetTagText("OTHERARRANGEMENTSFORACCESS");
		m_bPAFIndicator = (XML.GetTagText("PAFINDICATOR") == "1");
		m_sAddressGUID = XML.GetTagText("ADDRESSGUID");
		// EP2_1002 BusinessConnectionGroup fields moved from DC200
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window);
		var xn = XML.XMLDocument.documentElement;
		xn.setAttribute("CRUD_OP","READ");
		xn.setAttribute("SCHEMA_NAME","omCRUD");
		xn.setAttribute("ENTITY_REF","NEWPROPERTY");
		var xe = XML.XMLDocument.createElement("NEWPROPERTY");
		xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		xe.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		xn.appendChild(xe);
		XML.RunASP(document, "omCRUDIf.asp");
		if(XML.IsResponseOK())
		{
		XML.SelectTag(null,"NEWPROPERTY");
		scScreenFunctions.SetRadioGroupValue(frmScreen,"BusinessConnectionGroup",XML.GetAttribute("BUSINESSCONNECTION"));
		}


	}
}
<% /* BMIDS00444 End */ %>

function RetrieveContextData()
{
	m_bReadOnlyIndicator = scScreenFunctions.GetContextParameter(window, "idReadOnly", "0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","E2003");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
}
		
function UnhideOtherArrangements()
{
	<%/* unhide the Access Arrangements fields if ArrangementsForAccess is of type "Other". */%>
	if(scScreenFunctions.IsValidationType(frmScreen.cboAccessArrangements,"O"))
	{
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtOtherArrangements");
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtAccessContactName");
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtAccessTelephoneNumber");
		//JR - Omiplus24
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtCountryCode");
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtAreaCode");
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtTelephoneExtensionNum")
		//End
		
		//MAR1387
		if(scScreenFunctions.IsValidationType(frmScreen.cboAccessArrangements,"V"))
			PopulateAccessArrangements();			
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtOtherArrangements");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAccessContactName");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAccessTelephoneNumber");
		//JR - Omiplus24
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtCountryCode");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAreaCode");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTelephoneExtensionNum");
		//End
		frmScreen.txtOtherArrangements.value = "";
		frmScreen.txtAccessContactName.value = "";
		frmScreen.txtAccessTelephoneNumber.value = "";
		frmScreen.txtCountryCode.value = "";
		frmScreen.txtAreaCode.value = "";
		frmScreen.txtTelephoneExtensionNum.value = "";
	}
}

function StoreNewPropertyAddressData()
{
var AddressXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	<% /* Stores New Property Address in context for use in DC215 */ %>
	AddressXML.CreateActiveTag("ADDRESS");
	AddressXML.CreateTag("ADDRESSGUID", m_sAddressGUID);
	AddressXML.CreateTag("BUILDINGORHOUSENAME", frmScreen.txtHouseName.value);
	AddressXML.CreateTag("BUILDINGORHOUSENUMBER", frmScreen.txtHouseNumber.value);
	AddressXML.CreateTag("FLATNUMBER", frmScreen.txtFlatNo.value);
	AddressXML.CreateTag("STREET", frmScreen.txtStreet.value);
	AddressXML.CreateTag("DISTRICT", frmScreen.txtDistrict.value);
	AddressXML.CreateTag("TOWN", frmScreen.txtTown.value);
	AddressXML.CreateTag("COUNTY", frmScreen.txtCounty.value);
	AddressXML.CreateTag("COUNTRY", frmScreen.cboCountry.value);
	AddressXML.CreateTag("POSTCODE", frmScreen.txtPostcode.value);
	AddressXML.CreateTag("PAFINDICATOR", m_bPAFIndicator ? "1" : "0");

	scScreenFunctions.SetContextParameter(window,"idXML", AddressXML.XMLDocument.xml);
	
}

<% /* BMIDS00444 Display the security address of the mortgage account on which the further advance is based */ %>
function GetFurtherAdvanceAddress()
{
	<% /* First get the AccountGuid for account on which further advance is based. */ %>
	var AppXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AppXML.CreateRequestTag(window, null);
	AppXML.CreateActiveTag("APPLICATION");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AppXML.RunASP(document,"GetApplicationData.asp");
	if(AppXML.IsResponseOK())
	{
		var sAccountGuid = AppXML.GetTagText("ACCOUNTGUID");
		<% /* Now get the original Mortgage Account record. */ %>
		var XMLMortgageAccount = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XMLMortgageAccount.CreateRequestTag(window, null);
		XMLMortgageAccount.CreateActiveTag("MORTGAGEACCOUNT");
		XMLMortgageAccount.CreateTag("ACCOUNTGUID", sAccountGuid);
		XMLMortgageAccount.RunASP(document,"GetMortgageAccountDetails.asp");
		if(XMLMortgageAccount.IsResponseOK())			
		{
			if(XMLMortgageAccount.SelectTag(null, "SECURITYADDRESS") != null)
			{
				frmScreen.txtPostcode.value = XMLMortgageAccount.GetTagText("POSTCODE");
				frmScreen.txtFlatNo.value = XMLMortgageAccount.GetTagText("FLATNUMBER");
				frmScreen.txtHouseName.value = XMLMortgageAccount.GetTagText("BUILDINGORHOUSENAME");
				frmScreen.txtHouseNumber.value = XMLMortgageAccount.GetTagText("BUILDINGORHOUSENUMBER");
				frmScreen.txtStreet.value = XMLMortgageAccount.GetTagText("STREET");
				frmScreen.txtDistrict.value = XMLMortgageAccount.GetTagText("DISTRICT");
				frmScreen.txtTown.value = XMLMortgageAccount.GetTagText("TOWN");
				frmScreen.txtCounty.value = XMLMortgageAccount.GetTagText("COUNTY");
				frmScreen.cboCountry.value = XMLMortgageAccount.GetTagText("COUNTRY");
				m_bPAFIndicator = (XMLMortgageAccount.GetTagText("PAFINDICATOR") == "1");
				m_sAddressGUID = XMLMortgageAccount.GetTagText("ADDRESSGUID");
				<% /* BMIDS01016
				m_sMetaAction = "Edit"; */ %>
				<% /* BMIDS01009 Must save even if there are no changes, as this creates the NewPropertyAddress record */ %>
				FlagChange(true);
			}
		}
	}
}
<% /* BMIDS00444 End */ %>

<% /* BMIDS00444 Display the address of the mortgage account which has its remortgage indicator set to true */ %>
function GetRemortgageAddress()
{
	var sCustomerNumber = "";
	var sCustomerVersionNumber = "";
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "FindRemortgageAccountAddress");
	XML.CreateActiveTag("MORTGAGEACCOUNT");
	
	<% /* Add Customers to the query */ %>
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		XML.CreateActiveTag("ACCOUNTRELATIONSHIP");
		sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
		XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
		XML.ActiveTag = XML.ActiveTag.parentNode;
	}
	
	XML.RunASP(document, "FindRemortgageAccountAddress.asp");				
	
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = XML.CheckResponse(sErrorArray);
			
	if (sResponseArray[0] == true)
	{
		if(XML.SelectTag(null, "ADDRESS") != null)
		{
			frmScreen.txtPostcode.value = XML.GetTagText("POSTCODE");
			frmScreen.txtFlatNo.value = XML.GetTagText("FLATNUMBER");
			frmScreen.txtHouseName.value = XML.GetTagText("BUILDINGORHOUSENAME");
			frmScreen.txtHouseNumber.value = XML.GetTagText("BUILDINGORHOUSENUMBER");
			frmScreen.txtStreet.value = XML.GetTagText("STREET");
			frmScreen.txtDistrict.value = XML.GetTagText("DISTRICT");
			frmScreen.txtTown.value = XML.GetTagText("TOWN");
			frmScreen.txtCounty.value = XML.GetTagText("COUNTY");
			frmScreen.cboCountry.value = XML.GetTagText("COUNTRY");
			m_bPAFIndicator = (XML.GetTagText("PAFINDICATOR") == "1");
			m_sAddressGUID = XML.GetTagText("ADDRESSGUID");
			<% /* BMIDS01016
			m_sMetaAction = "Edit"; */ %>
			<% /* BMIDS01009 Must save even if there are no changes, as this creates the NewPropertyAddress record */ %>
			FlagChange(true);
		}
	}
}

function IsMainlandPostCode(vPostCode)
{
	var bSuccess;
	var XMLPostCodeList = null;
	var iCount = 0;
	var strPostCode = frmScreen.txtPostcode.value;
	var strPostCodePart = strPostCode.substring(0,4);
	var bValid = true;
	var XMLPostCode = null;
		
	//If PostCode is not found anywhere, it is valid
		
		//Check for INGNonUKPostCodes - with 4 characters of postcode
		XMLPostCodeList = m_XMLINGNonUKPostCodes.selectNodes("LISTNAME/LISTENTRY");
		for (iCount = XMLPostCodeList.length - 1; iCount >= 0; iCount--)
		{
			XMLPostCode = XMLPostCodeList[iCount];
			if (XMLPostCode.selectSingleNode("VALIDATIONTYPELIST/VALIDATIONTYPE[.='" +strPostCodePart  +"']") != null)
			{
				bValid = false;
				
			}
		}	
		
		XMLPostCodeList = null;
		if (bValid)	
		{
			strPostCodePart = strPostCode.substring(0,2);
		
		//Check for INGNonUKPostCodes - with 2 characters of postcode
			XMLPostCodeList = m_XMLINGNonUKPostCodes.selectNodes("LISTNAME/LISTENTRY");
			for (iCount = XMLPostCodeList.length - 1; iCount >= 0; iCount--)
			{
				XMLPostCode = XMLPostCodeList[iCount];
				if (XMLPostCode.selectSingleNode("VALIDATIONTYPELIST/VALIDATIONTYPE[.='" +strPostCodePart  +"']") != null)
				{
					bValid = false;
				
				}
			}	
		
		XMLPostCodeList = null;
		}
	
        if (!bValid)
		{
		    alert("Invalid Property Location - Must be Mainland Britain");
			frmScreen.txtPostcode.focus();
		}
		return bValid;
	
}

<% // MAR1387 - 15/03/2006 - Peter Edney %>
<% // MAR1481 - 29/03/2006 - Peter Edney %>
function PopulateAccessArrangements()
{

	<% /* Determine Type Of Mortgage */ %>
	var sMortgageApplicationValue = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue");
	var XMLType = m_XMLTypeOfMortgage.selectSingleNode("LISTNAME/LISTENTRY[VALUEID[.='" + sMortgageApplicationValue + "']]");
	var bReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC210");	
	
	if ((XMLType.selectSingleNode("VALIDATIONTYPELIST/VALIDATIONTYPE[.='R']") != null) && (!bReadOnly))
	{
	
		<% /* Get the mortage accounts */ %>
		var CustomerXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		with (CustomerXML)
		{
			
			var TagSEARCH = CreateRequestTag(window, "SEARCH")			
			CreateActiveTag("FINANCIALSUMMARY");
			CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
			CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
			ActiveTag = TagSEARCH;
			var TagMORTGAGEACCOUNTLIST = CreateActiveTag("MORTGAGEACCOUNTLIST");

			
			<% /* Loop through all customer context entries */ %>
			for(var nLoop = 1; nLoop <= 5; nLoop++)
			{
				var sCustomerNumber			= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
				var sCustomerVersionNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
				var sCustomerRoleType		= scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

				<% /* If the customer is an applicant, add him/her to the search */ %>
				if(sCustomerRoleType != "")
				{
					CreateActiveTag("MORTGAGEACCOUNT");
					CreateTag("CUSTOMERNUMBER",sCustomerNumber);
					CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
					ActiveTag = TagMORTGAGEACCOUNTLIST;
				}
			}
			
			ActiveTag = TagSEARCH;
			CreateTag("SHOWNOTREDEEMED","TRUE");
			
			RunASP(document, "FindMortgageAccountSummary.asp");
			bSuccess = IsResponseOK();
			
			if(bSuccess)
			{					
				<% /* Get the contacts for the remortgaged accounts */ %>
				var RelationList = XMLDocument.selectNodes("/RESPONSE/MORTGAGEACCOUNTLIST/MORTGAGEACCOUNT[REMORTGAGEINDICATOR=1]/ACCOUNTRELATIONSHIPLIST/ACCOUNTRELATIONSHIP");
				if(RelationList.length > 0)	
				{
					sCustomerNumber = RelationList.item(0).selectSingleNode("CUSTOMERNUMBER").text;
					sCustomerVersionNumber = RelationList.item(0).selectSingleNode("CUSTOMERVERSIONNUMBER").text;
				}
				else
				{
					<% /* If there are no remortgaged accounts then use the first contact */ %>
					sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber1");
					sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1");					
				}

				CustomerXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				CreateRequestTag(window, null)
				CreateActiveTag("CUSTOMERVERSION");
				CreateTag("CUSTOMERNUMBER", sCustomerNumber);
				CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
				RunASP(document, "GetPersonalDetails.asp");
				bSuccess = IsResponseOK();

				if(bSuccess)
				{						
					var telNode;
					var TelList;
					
					<% /* Get the preffered telephone number */ %>
					TelList = XMLDocument.selectNodes("/RESPONSE/CUSTOMER/CUSTOMERTELEPHONENUMBERLIST/CUSTOMERTELEPHONENUMBER[PREFERREDMETHODOFCONTACT=1]");
					if(TelList.length > 0)	
					{
						telNode = TelList.item(0);
					}
					else
					{
						<% /* If there is no preffered telephone number, just use the first one */ %>
						TelList = XMLDocument.selectNodes("/RESPONSE/CUSTOMER/CUSTOMERTELEPHONENUMBERLIST/CUSTOMERTELEPHONENUMBER");
						if(TelList.length > 0)	
						{
							telNode = TelList.item(0);
						}
					}
					
					
					var sEmail = ""
					if(SelectTag(null, "CUSTOMERVERSION") != null)
					{
						if(frmScreen.txtAccessContactName.value == "")
							frmScreen.txtAccessContactName.value = GetTagText("CORRESPONDENCESALUTATION");				
							
						sEmail = GetTagText("CONTACTEMAILADDRESS");																											
					}
					

					<% /* Only set the contact details if they are empty */ %>
					if(frmScreen.txtCountryCode.value == "" && frmScreen.txtAreaCode.value == ""
						&& frmScreen.txtAccessTelephoneNumber.value == "" && frmScreen.txtTelephoneExtensionNum.value == ""
						&& frmScreen.txtOtherArrangements.value == "")
					{

						if(telNode != null)
						{
								frmScreen.txtCountryCode.value = telNode.selectSingleNode("COUNTRYCODE").text;	
								frmScreen.txtAreaCode.value = telNode.selectSingleNode("AREACODE").text;	
								frmScreen.txtAccessTelephoneNumber.value = telNode.selectSingleNode("TELEPHONENUMBER").text;
								frmScreen.txtTelephoneExtensionNum.value = telNode.selectSingleNode("EXTENSIONNUMBER").text;

								var sContactTime = telNode.selectSingleNode("CONTACTTIME").text;
								if(sContactTime != "")			
									frmScreen.txtOtherArrangements.value = "Contact time: " + sContactTime;								
						}
						
						if(sEmail != "")
						{
							if(frmScreen.txtOtherArrangements.value != "")	
								frmScreen.txtOtherArrangements.value = frmScreen.txtOtherArrangements.value + "\n";
							frmScreen.txtOtherArrangements.value = frmScreen.txtOtherArrangements.value + "Email: " + sEmail;
						}
						
					}
				}
			}
		}
	}
}
-->
</script>
</body>
</html>


