<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      dc220.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:	Additional New Property Details Screen.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		19/01/00	Created.
RF		01/02/00	Reworked for IE5 and performance.
AD		08/03/00	Interface added to DC222 and DC223.
RF		08/03/00	Added DoButtonEnabling() for leasehold details button.
AD		22/03/00	Completed most of SYS0449
IW		27/03/00	GUI Homezone Rework - Added EntirelyOwnUseIndicator + ResidentialUseIndicator
AY		31/03/00	New top menu/scScreenFunctions change
AY		12/04/00	SYS0328 - Currency passed into popups
IW		26/04/00	OK Navigates to DC225
IW		28/04/00	SYS0621 (see Deuce for multiple fixes)
IW		03/05/00	SYS0655 Validation changes
IW		03/05/00	SYS0623 Popup (DC221) Discrepancies Fixed.
IW		10/05/00	SYS0722 Resolved ReadOnly Screen issues
SR		02/06/00	SYS0802 Screen title - changed to 'New Property Description '
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
BG		16/08/00	SYS0621 Fixed routing to next screen on clicking OK on validation error message.
BG		07/09/00	SYS0959 No of storeys disabled if property type is NOT Flat or Maisonette.
CL		05/03/01	SYS1920 Read only functionality added
CP		28/04/01	SYS2050 Critical Data functionality added
GD		11/05/01	SYS2050 Critical Data functionality ROLLED BACK
SA		30/05/01	SYS1279 Set focus to txtResidentialPctgeDetails if enabled.
SA		05/06/01	SYS1058 If residential use only, set percentage to 100.
SA		05/06/01	SYS0925 Deal with a non percentage value.
DC      20/07/01    SYS2038 Critical Data functionality ROLLED FORWARD AGAIN
JLD		10/12/01	SYS2806 completeness check routing
JLD		15/01/02	SYS3536 read only routing correction
DRC     27/02/02    SYS4049 Changed logic to do Critical Data check only when the Year Built box has been populated
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		29/05/02	SYS4767 MSMS to Core integration
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
DPF		28/06/2002	BMIDS00140	Have fixed bug relating to Lease start date when submitting a freehold application
MV		22/08/2002	BMIDS00355	IE 5.5 upgrade Errors - Modified the HTML layout , Modified the Popwindows Height and width 
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
MDC		17/12/2002	BM0146		Make 'Flat Serviced By Lift' non mandatory
HMA     17/09/2003  BM0063      Amend HTML text for radion buttons
KRW     25/09/03    BM00063     Corrections to screen alignment
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
PB		23/05/2006	EP570		Should allow future year built (not built yet)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom Changes
Prog    Date           Description
MAH		23/11/2006		E2CR35 Added new Property fields
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

%>
<HEAD>
<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
data=scClientFunctions.asp width=1 height=1 type=text/x-scriptlet 
VIEWASTEXT></OBJECT>
<script src="validation.js" language="JScript"></script>

<% // Specify Forms Here %>
<form id="frmCancel" method="post" action="DC210.asp" STYLE="DISPLAY: none"></form>
<form id="frmSubmit" method="post" action="DC225.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% // Specify Screen Layout Here %>
<form id="frmScreen" mark validate="onchange" year4>
	<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 460px" class="msgGroup">
	<% // left hand column %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
		Tenure of Property
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<select id="cboTenureOfProperty" style="WIDTH: 144px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Property Type
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<select id="cboPropertyType" style="WIDTH: 144px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 100px" class="msgLabel">
		Floor No.
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="txtFloorNo" maxlength="5" style="WIDTH: 40px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 130px" class="msgLabel">
		Flat serviced by Lift?
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="optFlatServicedByLiftYes" name="FlatServicedByLiftRadioGroup" type="radio" value="1"><label for="optFlatServicedByLiftYes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="optFlatServicedByLiftNo" name="FlatServicedByLiftRadioGroup" type="radio" value="0"><label for="optFlatServicedByLiftNo" class="msgLabel">No</label>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 160px" class="msgLabel">
		Number of flats in the block?
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="txtNoOfFlats" maxlength="5" style="WIDTH: 40px; POSITION: absolute" class="msgTxt" NAME="txtNoOfFlats">
		</span>
	</span>	
	<span style="LEFT: 4px; WIDTH: 200px; POSITION: absolute; TOP: 185px" class="msgLabel">
		Is the flat above<br> commercial premises?
		<span style="LEFT: 140px; WIDTH: 50px; POSITION: absolute; TOP: 0px">
			<input id="optAboveCommercialYes" name="FlatAboveCommercialGroup" type="radio" value="1"><label for="optAboveCommercialYes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 180px; WIDTH: 50px; POSITION: absolute; TOP: 0px">
			<input id="optAboveCommercialNo" name="FlatAboveCommercialGroup" type="radio" value="0"><label for="optAboveCommercialNo" class="msgLabel">No</label>
		</span>
	</span>	<span style="LEFT: 4px; POSITION: absolute; TOP: 220px" class="msgLabel">
		No. of Storeys
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="txtNumStoreys" maxlength="5" style="WIDTH: 40px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; WIDTH: 200px; POSITION: absolute; TOP: 245px" class="msgLabel">
		Property Entirely For<br>Own Use?
		<span style="LEFT: 140px; WIDTH: 50px; POSITION: absolute; TOP: 0px">
			<input id="optEntirelyOwnUseIndicatorYes" name="EntirelyOwnUseIndicatorGroup" type="radio" value="1"><label for="optEntirelyOwnUseIndicatorYes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 180px; WIDTH: 50px; POSITION: absolute; TOP: 0px">
			<input id="optEntirelyOwnUseIndicatorNo" name="EntirelyOwnUseIndicatorGroup" type="radio" value="0"><label for="optEntirelyOwnUseIndicatorNo" class="msgLabel">No</label>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 280px" class="msgLabel">
		Residential Use Only? 
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="optResidentialUseOnlyIndicatorYes" name="ResidentialUseOnlyIndicatorGroup" type="radio" value="1"><label for="optResidentialUseOnlyIndicatorYes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="optResidentialUseOnlyIndicatorNo" name="ResidentialUseOnlyIndicatorGroup" type="radio" value="0"><label for="opResidentialUseOnlyIndicatorNo" class="msgLabel">No</label>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 310px" class="msgLabel">
		New Property?
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="optNewPropertyYes" name="NewPropertyRadioGroup" type="radio" value="1"><label for="optNewPropertyYes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="optNewPropertyNo" name="NewPropertyRadioGroup" type="radio" value="0"><label for="optNewPropertyNo" class="msgLabel">No</label>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 340px" class="msgLabel">
		Year Built
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="txtYearBuilt" maxlength="4" style="WIDTH: 40px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 370px" class="msgLabel">
		House Builder's Guarantee
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<select id="cboHouseBuildersGuarantee" style="WIDTH: 146px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 400px" class="msgLabel">
		Guarantee Expiry Date
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="txtGuaranteeExpiryDate" maxlength="10" style="WIDTH: 70px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
	<% // right hand column %>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 40px" class="msgLabel">
		Building Construction
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboBuildingConstruction" style="WIDTH: 180px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Roof Construction
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboRoofConstruction" style="WIDTH: 180px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 100px" class="msgLabel">
		Total Land
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboTotalLand" style="WIDTH: 180px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 130px" class="msgLabel">
		Parking Arrangements
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboParkingArrangements" style="WIDTH: 180px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 160px" class="msgLabel">
		Location of Main Garage
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboLocationofMainGarage" style="WIDTH: 180px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 190px" class="msgLabel">
		No. of Garages
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtNumGarages" maxlength="5" style="WIDTH: 50px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 220px" class="msgLabel">
		No. of Outbuildings
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtNumOutbuildings" maxlength="5" style="WIDTH: 50px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 250px" class="msgLabel">
		Residential Percentage
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtResidentialPctge" maxlength="3" style="WIDTH: 50px; POSITION: absolute" class="msgTxt">
			<span style="LEFT: 50px; POSITION: absolute; TOP: 3px" class="msgLabel">%</span>
		</span>
	</span>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 280px" class="msgLabel">
		Residential Percentage Details
		<span style="LEFT: 0px; POSITION: absolute; TOP: 16px"><TEXTAREA class=msgTxt id=txtResidentialPctgeDetails style="WIDTH: 260px; POSITION: absolute"></TEXTAREA>
		</span>
	</span>
	<% // SG 29/05/02 SYS4767 START %>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 340px" class="msgLabel">
		Date Of Entry
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtDateOfEntry" maxlength="10" style="WIDTH: 70px; POSITION: absolute" class="msgTxt">
		</span>
	</span>	
	<% // SG 29/05/02 SYS4767 END %>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 364px" class="msgLabel">
		Is the property affected by an<br>agricultural restriction/covenant?
		<span style="LEFT: 170px; POSITION: absolute; TOP: 0px">
			<input id="optAgricRestrictYes" name="AgriculturalRestrictionsGroup" type="radio" value="1"><label for="optAgricRestrictYes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 210px; POSITION: absolute; TOP: 0px">
			<input id="optAgricRestrictNo" name="AgriculturalRestrictionsGroup" type="radio" value="0"><label for="optAgricRestrictNo" class="msgLabel">No</label>
		</span>
	</span>
	<% // buttons for popups %>
	<span style="LEFT: 19px; POSITION: absolute; TOP: 430px">
		<input id="btnLeaseholdDetails" value="Leasehold Details" type="button" style="WIDTH: 105px" class="msgButton">
	</span>
	<span style="LEFT: 250px; POSITION: absolute; TOP: 430px">
		<input id="btnRooms" value="Rooms" type="button" style="WIDTH: 105px" class="msgButton">
	</span>
	<span style="LEFT: 478px; POSITION: absolute; TOP: 430px">
		<input id="btnDepositDetails" value="Deposit Details" type="button" style="WIDTH: 105px" class="msgButton">
	</span>
	</div>
</form>
	
<% // Main Buttons %>
<div id="msgButtons" style="LEFT: 8px; WIDTH: 604px; POSITION: absolute; TOP: 524px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% // File containing field attributes %>
<!--  #include FILE="attribs/dc220attribs.asp" -->

<% // Specify Code Here %>
<script language="JScript">
<!--	// JScript Code
var m_sMetaAction = null;
var m_sDebug = null;
var	m_sApplicationNumber = null;
var	m_sApplicationFactFindNumber = null;
var m_sCurrency = null;
var m_sRoomTypesXML = null;
var m_sLeaseholdXML = null;
// BMIDS00140 - DPF 28/06/02 - 2nd variable to hold leasehold information taken when first entering into screen
var m_sLeaseholdHoldingXML = null;
var m_sDepositXML = null;
// SYS4049 changed flag  m_bNewPropertyRecordExistsOnScreenEntry
// to m_bYearBuiltExists  
var m_bYearBuiltExists = false;
var scScreenFunctions;
var m_blnReadOnly = false;
var m_sReadOnly = "";


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
	FW030SetTitles("New Property Description","DC220",scScreenFunctions);

	RetrieveContextData();
	SetMasks();
	Validation_Init();
	InitialiseScreen();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC220");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	scScreenFunctions.SetFocusToFirstField(frmScreen);	
	
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

function frmScreen.txtResidentialPctgeDetails.onkeyup()
{
	scScreenFunctions.RestrictLength(frmScreen, "txtResidentialPctgeDetails", 255, true);
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sDebug = scScreenFunctions.GetContextParameter(window,"idDebug","True");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","C50000002283");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	m_sCurrency = scScreenFunctions.GetContextParameter(window,"idCurrency","");
}

function frmScreen.txtNumStoreys.onblur()
{
	if (frmScreen.txtNumStoreys.value == "0")
	{
		alert("Invalid Number");
		frmScreen.txtNumStoreys.focus();
	}
}
function frmScreen.btnDepositDetails.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sDepositXML;
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = m_sApplicationNumber;
	ArrayArguments[3] = m_sApplicationFactFindNumber;
	ArrayArguments[4] = m_sCurrency;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "DC223.asp", ArrayArguments, 575, 330);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sDepositXML = sReturn[1];
	}
}

function frmScreen.btnLeaseholdDetails.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sLeaseholdXML;
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = m_sApplicationNumber;
	ArrayArguments[3] = m_sApplicationFactFindNumber;
	ArrayArguments[4] = m_sCurrency;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "DC222.asp", ArrayArguments, 340, 369);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sLeaseholdXML = sReturn[1];
	}
}

function frmScreen.btnRooms.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sRoomTypesXML;
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = m_sApplicationNumber;
	ArrayArguments[3] = m_sApplicationFactFindNumber;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "dc221.asp", ArrayArguments, 255, 300);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sRoomTypesXML = sReturn[1];
	}
}

function InitialiseScreen()
{
	PopulateCombos();
	GetNewPropertyDescription();
	DoFieldEnablingForFlat();
	DoFieldEnablingForResPctge();
	DoFieldEnablingForGuaranteeExpiryDate();
	DoButtonEnabling();
}

function DoButtonEnabling()
{
	if (scScreenFunctions.IsValidationType(frmScreen.cboTenureOfProperty, "L"))
		frmScreen.btnLeaseholdDetails.disabled = false;
	else
		frmScreen.btnLeaseholdDetails.disabled = true;	
}

function frmScreen.cboTenureOfProperty.onchange()
{
	DoButtonEnabling();
}

function DoFieldEnablingForFlat()
{
	if (m_sReadOnly != "1")
	{
		if (scScreenFunctions.IsValidationType(frmScreen.cboPropertyType, "F"))
		{
			scScreenFunctions.SetFieldState(frmScreen, "txtFloorNo", "W");
			scScreenFunctions.SetRadioGroupState(frmScreen, "FlatServicedByLiftRadioGroup", "W");
			scScreenFunctions.SetFieldState(frmScreen, "txtNoOfFlats", "W");<%/*MAH 23/11/2006 E2CR35*/%>
			scScreenFunctions.SetRadioGroupState(frmScreen, "FlatAboveCommercialGroup", "W");<%/*MAH 23/11/2006 E2CR35*/%>
		}
		else if(scScreenFunctions.IsValidationType(frmScreen.cboPropertyType, "M"))
		{
			scScreenFunctions.SetFieldState(frmScreen, "txtNoOfFlats", "W");<%/*MAH 23/11/2006 E2CR35*/%>
			scScreenFunctions.SetRadioGroupState(frmScreen, "FlatAboveCommercialGroup", "W");<%/*MAH 23/11/2006 E2CR35*/%>
		}
		else
		{
			scScreenFunctions.SetFieldState(frmScreen, "txtFloorNo", "D");
			scScreenFunctions.SetRadioGroupState(frmScreen, "FlatServicedByLiftRadioGroup", "D");
			frmScreen.txtNoOfFlats.value = "";<%/*MAH 23/11/2006 E2CR35*/%>
			scScreenFunctions.SetFieldState(frmScreen, "txtNoOfFlats", "D");<%/*MAH 23/11/2006 E2CR35*/%>
			scScreenFunctions.SetRadioGroupValue(frmScreen, "FlatAboveCommercialGroup","");<%/*MAH 23/11/2006 E2CR35*/%>
			scScreenFunctions.SetRadioGroupState(frmScreen, "FlatAboveCommercialGroup", "D");<%/*MAH 23/11/2006 E2CR35*/%>
		}
		//BG SYS0959 07/09/00
		if ((scScreenFunctions.IsValidationType(frmScreen.cboPropertyType, "H")) ||
			(scScreenFunctions.IsValidationType(frmScreen.cboPropertyType, "B")))	
		{
			frmScreen.txtNumStoreys.value = "";
			scScreenFunctions.SetFieldState(frmScreen, "txtNumStoreys", "D");
			frmScreen.txtNumStoreys.setAttribute("required", "false");
		}
		else
		{
			scScreenFunctions.SetFieldState(frmScreen, "txtNumStoreys", "W");
			frmScreen.txtNumStoreys.setAttribute("required", "true");
		}
		//BG SYS0959 END
	}
}

function DoFieldEnablingForResPctge()
{
	var sFieldState = null;
	
	if (frmScreen.txtResidentialPctge.value < 100)
	{
		if (m_sReadOnly == "1")
			sFieldState = "R";
		else
			sFieldState = "W";
	}
	else
		sFieldState = "D";
		
	scScreenFunctions.SetFieldState(frmScreen, "txtResidentialPctgeDetails", sFieldState);	<%/* SYS1279 If its now writable - set the focus */ %>
	if (sFieldState == "W")
	{
		frmScreen.txtResidentialPctgeDetails.focus();
	}
}

function frmScreen.cboPropertyType.onchange()
{
	DoFieldEnablingForFlat();
}

function frmScreen.txtResidentialPctge.onblur()
{
	DoFieldEnablingForResPctge();
}

//SYS1058 
function frmScreen.optResidentialUseOnlyIndicatorYes.onclick() 
{
	frmScreen.txtResidentialPctge.value = "100";
	DoFieldEnablingForResPctge();
}	
function frmScreen.optResidentialUseOnlyIndicatorNo.onclick()
{
	frmScreen.txtResidentialPctge.value = "";
	DoFieldEnablingForResPctge();
}
//SYS1058 end

function frmScreen.cboHouseBuildersGuarantee.onchange()
{
	DoFieldEnablingForGuaranteeExpiryDate();
}

function DoFieldEnablingForGuaranteeExpiryDate()
{
	if (m_sReadOnly != "1")
	{
		if (scScreenFunctions.IsValidationType(frmScreen.cboHouseBuildersGuarantee, "G"))
		{
			scScreenFunctions.SetFieldState(frmScreen, "txtGuaranteeExpiryDate", "W");
		}
		else
		{
			scScreenFunctions.SetFieldState(frmScreen, "txtGuaranteeExpiryDate", "D");
		}
	}	
}

function GetNewPropertyDescription()
{
	var XML = null;
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "SEARCH");

	XML.CreateActiveTag("NEWPROPERTY");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);

	XML.RunASP(document,"GetNewPropertyDescription.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);

	if (ErrorReturn[0] == true)
	{
		
		PopulateScreenFromXML(XML);
		// AQR SYS4049 - check for critical data 
		if (frmScreen.txtYearBuilt.value.length > 0 )
			m_bYearBuiltExists = true;
		if (XML.SelectTag(null,"NEWPROPERTYROOMTYPELIST") != null)
			m_sRoomTypesXML = XML.ActiveTag.xml;

		if (XML.SelectTag(null,"NEWPROPERTYLEASEHOLD") != null)
		{	
			m_sLeaseholdHoldingXML = XML.ActiveTag.xml; // BMIDS00140 - assign to 2nd variable to fix bug
			m_sLeaseholdXML = XML.ActiveTag.xml;
		}
		if (XML.SelectTag(null,"NEWPROPERTYDEPOSITLIST") != null)
			m_sDepositXML = XML.ActiveTag.xml;
	}

	ErrorTypes = null;
	ErrorReturn = null;
}

function PopulateScreenFromXML(XML)
{
	XML.ActiveTag = null;
	XML.CreateTagList("NEWPROPERTY");
	if (XML.SelectTagListItem(0))
	{
		frmScreen.txtNumStoreys.value = XML.GetTagText("NUMBEROFSTOREYS");
		scScreenFunctions.SetRadioGroupValue(frmScreen, "NewPropertyRadioGroup", XML.GetTagText("NEWPROPERTYINDICATOR"));
		frmScreen.txtYearBuilt.value = XML.GetTagText("YEARBUILT");
		frmScreen.cboHouseBuildersGuarantee.value = XML.GetTagText("HOUSEBUILDERSGUARANTEE");
		frmScreen.txtGuaranteeExpiryDate.value = XML.GetTagText("GUARANTEEEXPIRYDATE");
		frmScreen.cboTenureOfProperty.value = XML.GetTagText("TENURETYPE");
		frmScreen.cboPropertyType.value = XML.GetTagText("TYPEOFPROPERTY");
		scScreenFunctions.SetRadioGroupValue(frmScreen, "EntirelyOwnUseIndicatorGroup", XML.GetTagText("ENTIRELYOWNUSEINDICATOR"));
		scScreenFunctions.SetRadioGroupValue(frmScreen, "ResidentialUseOnlyIndicatorGroup", XML.GetTagText("RESIDENTIALUSEONLYINDICATOR"));
		frmScreen.txtFloorNo.value = XML.GetTagText("FLOORNUMBER");
		scScreenFunctions.SetRadioGroupValue(frmScreen, "FlatServicedByLiftRadioGroup", XML.GetTagText("FLATSERVICEDBYLIFTINDICATOR"));
		frmScreen.cboBuildingConstruction.value = XML.GetTagText("BUILDINGCONSTRUCTIONTYPE");
		frmScreen.cboRoofConstruction.value = XML.GetTagText("ROOFCONSTRUCTIONTYPE");
		frmScreen.cboTotalLand.value = XML.GetTagText("TOTALLANDTYPE");
		frmScreen.cboParkingArrangements.value = XML.GetTagText("PARKINGARRANGEMENTSTYPE");
		frmScreen.cboLocationofMainGarage.value = XML.GetTagText("LOCATIONOFGARAGETYPE");
		frmScreen.txtNumGarages.value = XML.GetTagText("NUMBEROFGARAGES");
		frmScreen.txtNumOutbuildings.value = XML.GetTagText("NUMBEROFOUTBUILDINGS");
		frmScreen.txtResidentialPctge.value = XML.GetTagText("RESIDENTIALPERCENTAGE");
		frmScreen.txtResidentialPctgeDetails.value = XML.GetTagText("RESIDENTIALPERCENTAGEDETAILS");
		
		<%/*MAH 23/11/2006 E2CR35*/%>
		frmScreen.txtNoOfFlats.value = XML.GetTagText("NOFLATSINBLOCK");
		scScreenFunctions.SetRadioGroupValue(frmScreen, "FlatAboveCommercialGroup", XML.GetTagText("FLATABOVECOMMERCIAL"));
		scScreenFunctions.SetRadioGroupValue(frmScreen, "AgriculturalRestrictionsGroup", XML.GetTagText("AGRICULTURALRESTRICTIONS"));

		//SG 29/05/02 SYS4767
		frmScreen.txtDateOfEntry.value = XML.GetTagText("DATEOFENTRY");
	}
}

function PopulateCombos()
{
	var XMLHouseBuildersGuarantee = null;
	var XMLTenureOfProperty = null;
	var XMLPropertyType = null;
	var XMLBuildingConstruction = null;
	var XMLRoofConstruction = null;
	var XMLTotalLand = null;
	var XMLParkingArrangements = null;
	var XMLLocationofMainGarage = null;

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	var sGroupList = new Array(
			"HouseBuildersGuarantee", "PropertyTenure",
			"PropertyType",	"BuildingConstruction", 
			"RoofConstruction",	"TotalLand",
			"ParkingArrangements", "GarageLocation");

	if(XML.GetComboLists(document,sGroupList))
	{
		if (m_sDebug == "True")
			XML.WriteXMLToFile("C:\\temp\\dc220GetComboLists.xml");
		XMLHouseBuildersGuarantee = XML.GetComboListXML("HouseBuildersGuarantee");
		XMLTenureOfProperty = XML.GetComboListXML("PropertyTenure");
		XMLPropertyType = XML.GetComboListXML("PropertyType");
		XMLBuildingConstruction = XML.GetComboListXML("BuildingConstruction");
		XMLRoofConstruction = XML.GetComboListXML("RoofConstruction");
		XMLTotalLand = XML.GetComboListXML("TotalLand");
		XMLParkingArrangements = XML.GetComboListXML("ParkingArrangements");
		XMLLocationofMainGarage = XML.GetComboListXML("GarageLocation");

		var blnSuccess = true;
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,
			frmScreen.cboHouseBuildersGuarantee,XMLHouseBuildersGuarantee,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,
			frmScreen.cboTenureOfProperty,XMLTenureOfProperty,true);			
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,
			frmScreen.cboPropertyType,XMLPropertyType,true);			
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,
			frmScreen.cboBuildingConstruction,XMLBuildingConstruction,true);			
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,
			frmScreen.cboRoofConstruction,XMLRoofConstruction,true);		
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,
			frmScreen.cboTotalLand,XMLTotalLand,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,
			frmScreen.cboParkingArrangements,XMLParkingArrangements,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,
			frmScreen.cboLocationofMainGarage,XMLLocationofMainGarage,true);

		if(!blnSuccess)
		{
			alert('Failed to populate combos');
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			DisableMainButton("Submit");
		}
	}

	XML = null;
}

function btnCancel.onclick()
{
	<% //clear down any Context params %>
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","");
		frmCancel.submit();
	}
}

function CalcPropertyAge()
{
	<% /* MO - BMIDS00807 */ %>
	var dtNow = scScreenFunctions.GetAppServerDate(true);
	<% /* var dtNow = new Date(); */ %>
	return (dtNow.getFullYear() - frmScreen.txtYearBuilt.value);
}

function btnSubmit.onclick()
{
	var bContinue = true;
	
	if (m_sReadOnly != "1" )
	{
		if(frmScreen.onsubmit())
		{
			if (frmScreen.cboPropertyType.value == "30")
			{
				if (frmScreen.txtFloorNo.value == "")
				{
					alert('Please enter a value');
					frmScreen.txtFloorNo.focus();
					bContinue = false;
				}
			}
			//SYS0925 Check valid percentage.
			if (frmScreen.txtResidentialPctge.value < 0 || frmScreen.txtResidentialPctge.value > 100)
			{	
				alert('Please enter a percentage in the range 0-100');
				frmScreen.txtResidentialPctge.focus();
				bContinue=false;
			}
			if (bContinue)
			{
				if (frmScreen.txtNumStoreys.value == "0")
				{
					alert('Invalid Number');
					frmScreen.txtNumStoreys.focus();
					bContinue = false;
				}	
			}	
			if (bContinue)
			{
				if (frmScreen.cboHouseBuildersGuarantee.value == "")
				{
					if (CalcPropertyAge() < 10)
					{
						alert('House Builder\'s Guarantee type is required where property is less than 10 years old');
						frmScreen.cboHouseBuildersGuarantee.focus();
						bContinue = false;
					}
				}
			}
			<% /* BM0146 MDC 17/12/2002
			if (bContinue)
			{
				if (!frmScreen.optFlatServicedByLiftYes.disabled)
				{
					if (scScreenFunctions.GetRadioGroupValue(frmScreen,"FlatServicedByLiftRadioGroup") == null)
					{
						alert('Please enter a value');
						frmScreen.optFlatServicedByLiftYes.focus();				
						bContinue = false;
					}
				}
			} */ %>
			if (bContinue)
			{
				if (scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtGuaranteeExpiryDate,"<"))
				{
					alert("Guarantee expiry date cannot be in the past.");
					frmScreen.txtGuaranteeExpiryDate.focus();
					bContinue = false;
				}
			}
			if (bContinue)
				if ((frmScreen.txtYearBuilt.value != "") && (frmScreen.txtYearBuilt.value < 1000))
					{
						alert("Build Year must be after 999AD.");
						frmScreen.txtYearBuilt.focus();
						bContinue = false;
					}
			if (bContinue)
			{
				<% /* MO - BMIDS00807 */ %>
				var dateNow = scScreenFunctions.GetAppServerDate(true);
				<% /* var dateNow = new Date(); */ %>
				var yearNow = dateNow.getYear();
			}
			<% /* PB 23/05/2006 EP570 - should allow future year built (not built yet)
			if ((frmScreen.txtYearBuilt.value != "") && (frmScreen.txtYearBuilt.value > yearNow))
			{
				alert("Year Built must not be later than current year");
				frmScreen.txtYearBuilt.focus();
				bContinue = false;
			} */ %>
			if (bContinue)
			{
								
				var XMLRoomTypes = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XMLRoomTypes.LoadXML(m_sRoomTypesXML);
				var XMLLeasehold = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				
				//BMIDS00140 check if Leasehold to ensure we don't pass it data that is not required
				if ((frmScreen.cboTenureOfProperty.value != 2) && (frmScreen.cboTenureOfProperty.value != 3))
				{
					XMLLeasehold.LoadXML(m_sLeaseholdHoldingXML);
				}
				else
				{
					XMLLeasehold.LoadXML(m_sLeaseholdXML);
				}
				var XMLDeposit = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XMLDeposit.LoadXML(m_sDepositXML);

				XMLLeasehold.SelectTag(null,"NEWPROPERTYLEASEHOLD");
			}
			if (bContinue)
			{
				if(scScreenFunctions.RestrictLength(frmScreen, "txtResidentialPctgeDetails", 255, true))
					bContinue = false;
			}
			if (bContinue && IsChanged())
			{			
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XML.CreateRequestTag(window, "UPDATE");		    
				XML.CreateActiveTag("NEWPROPERTY");
				XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
				XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
				XML.CreateTag("NUMBEROFSTOREYS",frmScreen.txtNumStoreys.value);
				XML.CreateTag("ENTIRELYOWNUSEINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"EntirelyOwnUseIndicatorGroup"));
				XML.CreateTag("RESIDENTIALUSEONLYINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"ResidentialUseOnlyIndicatorGroup"));
				XML.CreateTag("NEWPROPERTYINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"NewPropertyRadioGroup"));
				XML.CreateTag("YEARBUILT",frmScreen.txtYearBuilt.value);
				XML.CreateTag("HOUSEBUILDERSGUARANTEE",frmScreen.cboHouseBuildersGuarantee.value);
				XML.CreateTag("GUARANTEEEXPIRYDATE",frmScreen.txtGuaranteeExpiryDate.value);
				XML.CreateTag("TENURETYPE",frmScreen.cboTenureOfProperty.value);
				XML.CreateTag("TYPEOFPROPERTY",frmScreen.cboPropertyType.value);
				XML.CreateTag("FLOORNUMBER",frmScreen.txtFloorNo.value);
				XML.CreateTag("FLATSERVICEDBYLIFTINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"FlatServicedByLiftRadioGroup"));
				XML.CreateTag("BUILDINGCONSTRUCTIONTYPE",frmScreen.cboBuildingConstruction.value);
				XML.CreateTag("ROOFCONSTRUCTIONTYPE",frmScreen.cboRoofConstruction.value);
				XML.CreateTag("TOTALLANDTYPE",frmScreen.cboTotalLand.value);
				XML.CreateTag("PARKINGARRANGEMENTSTYPE",frmScreen.cboParkingArrangements.value);
				XML.CreateTag("LOCATIONOFGARAGETYPE",frmScreen.cboLocationofMainGarage.value);
				XML.CreateTag("NUMBEROFGARAGES",frmScreen.txtNumGarages.value);
				XML.CreateTag("NUMBEROFOUTBUILDINGS",frmScreen.txtNumOutbuildings.value);
				XML.CreateTag("RESIDENTIALPERCENTAGE",frmScreen.txtResidentialPctge.value);	
				XML.CreateTag("RESIDENTIALPERCENTAGEDETAILS",frmScreen.txtResidentialPctgeDetails.value);
				
				//SG 29/05/02 SYS4767
				XML.CreateTag("DATEOFENTRY",frmScreen.txtDateOfEntry.value);
				<%/*MAH 23/11/2006 E2CR35*/%>
				XML.CreateTag("NOFLATSINBLOCK",frmScreen.txtNoOfFlats.value);
				XML.CreateTag("FLATABOVECOMMERCIAL",scScreenFunctions.GetRadioGroupValue(frmScreen,"FlatAboveCommercialGroup"));
				XML.CreateTag("AGRICULTURALRESTRICTIONS",scScreenFunctions.GetRadioGroupValue(frmScreen,"AgriculturalRestrictionsGroup"));
												
				XML.AddXMLBlock(XMLRoomTypes.XMLDocument);
				XML.AddXMLBlock(XMLLeasehold.XMLDocument);					
				XML.AddXMLBlock(XMLDeposit.XMLDocument);
				
		
				if (m_bYearBuiltExists)
				{
					if (m_sDebug == "True")
						XML.WriteXMLToFile("C:\\temp\\dc220UpdateNewPropertyDetails.xml");
					//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD ROLLBACK 11/05/01 SYS2050
					//start CRITICAL DATA CHECK CP_28/04/2001 AQR SYS2050
						// MOVED DRC SYS4049 - So that Crit Data Check stuff only done for new record		
					XML.SelectTag(null,"REQUEST");
					XML.SetAttribute("OPERATION","CriticalDataCheck"); 
					XML.CreateActiveTag("CRITICALDATACONTEXT");
					XML.SetAttribute("APPLICATIONNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null));
					XML.SetAttribute("APPLICATIONFACTFINDNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null));
					XML.SetAttribute("SOURCEAPPLICATION","Omiga");
					XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
					XML.SetAttribute("ACTIVITYINSTANCE","1");
					XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
					XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
					XML.SetAttribute("COMPONENT","omApp.NewPropertyBO");
					XML.SetAttribute("METHOD","UpdateNewPropertyDetails");
						//END  MOVED DRC SYS4049 
					
					window.status = "Critical Data Check - please wait";
					// 					XML.RunASP(document,"OmigaTMBO.asp");
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
					
					//end CRITICAL DATA CHECK CP_28/04/2001 AQR SYS2050
					//XML.RunASP(document,"UpdateNewPropertyDetails.asp");
					
					// END ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD ROLLBACK 11/05/01 SYS2050 
				}
				else
				{
					if (m_sDebug == "True")
						XML.WriteXMLToFile("C:\\temp\\dc220CreateNewPropertyDetails.xml");
                    // DRC SYS4049 Create changed to Update	since the record must already exist when in this screen					
					// 					XML.RunASP(document,"UpdateNewPropertyDetails.asp");
					// Added by automated update TW 09 Oct 2002 SYS5115
					switch (ScreenRules())
						{
						case 1: // Warning
						case 0: // OK
											XML.RunASP(document,"UpdateNewPropertyDetails.asp");
							break;
						default: // Error
							XML.SetErrorResponse();
						}

					// make sure this is only done once without crit data check DRC SYS4049
					m_bYearBuiltExists = true;
				}
				bContinue = XML.IsResponseOK();
			}
			XML = null;		
		}
		else bContinue = false;
	}
	if (bContinue)
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else
			frmSubmit.submit();
	}
}


-->
</script>
</body>
</html>


