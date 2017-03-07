<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
 /*
Workfile:      DC065.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   This frame is used to add details of a new customer
				address or to edit details of an existing one.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DPol	16/11/1999	Created.
DPol	17/11/1999	Still in the process of creating it, but as the
					Business Object isn't ready, this is as far as
					I can go at the moment, so checked in for safety.
JLD		09/12/99	Finished off
JLD		13/12/1999	DC/028 - implemented change
					Also, corrected routing when errors occurred.
JLD		17/12/1999	SYS0080 - populate the customer combo with all customers - don't need to check for the role type.										
JLD		17/12/1999	SYS0082 - Flatnumber field addressed incorrectly.
JLD		20/12/1999	SYS0090 - Fields to be made readonly rather than hidden.
AD		02/02/2000	Rework
AY		14/02/00	Change to msgButtons button types
AD		21/03/2000	SYS0111 - Done
AD		21/03/2000	SYS0484 - Default country combo to 'UK'
MC		30/03/2000	SYS0568 - Enable 'Address Of' combo when adding and more than 1 applicant
AY		30/03/00	New top menu/scScreenFunctions change
MH      02/05/00    SYS0618 Postcode validation
BG		17/05/00	SYS0752 Removed Tooltips
BG		18/05/00	SYS0671	When Home\Current is selected from combo, txtDateMovedOut is cleared then set to read only.
MH      22/06/00    SYS0933 Readonly processing
BG		25/07 00	SYS0958 if Additional Current is selected, clear date moved out and make it read only.
CL		05/03/01	SYS1920 Read only functionality added
TJ		30/03/01	SYS2050 Added Critical data functionality
GD		11/05/01	SYS2050 ROLL BACK - removed Critical data functionality as above
PF		11/05/01	SYS2300	Changed idXML to idXML2
SA		31/05/01	SYS1034 Tenancy details button should be enabled in ready only
							Rent and account number fields enabled even when "previous" selected
SA		06/06/01	SYS1007 Various minor fixes.
DJP		27/09/01	SYS2564/SYS2752 (child) Change DC065 to enable client versions to override 
					the county textbox, and various labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		29/05/02	SYS4767 MSMS to Core integration
GD		26/06/02	BMIDS0077 applied SG 25/06/02 SYS4930
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMids History:

Prog	Date		Description
MDC		27/08/2002	BMIDS00336 - Credit Score and Bureau Download
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
MDC		15/11/2002	BMIDS00965 - Remove Screen Rules from PopulateScreen
GD		17/11/2002	BMIDS00376 Disable screen if readonly
MDC		18/11/2002	BMIDS00965 - Added Screen Rules processing to Save
MV		06/12/2002	BMIDS0020	Modified the PAF Search tab Order
KRW     14/04/2004  BMIDS710 -  Added new flag bUpdatePAFIndicator to indicate when PAF search has been completed
							-   Use this value to indicate wheter we need to update fields esp PAFINDICATOR
KRW     13/05/2004  BMIDS745    cboCountry no longer defaults to England 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History

Prog    Date			AQR		Description
MF		22/07/2005		MAR19	IA_WP01 disabling fields.
MF		27/07/2005		MAR19	change control (code mainly inside attribs\dc065attribs.asp)
TLiu	02/09/2005		MAR38	Changed layout for Flat No., House Name & No.
MF		26/09/2005		MAR19	Modifications for MAR19. Displaying of various controls
								now dependant on global paramater RestrictUseOfCurrentOrCorAddress
GHun	01/12/2005		MAR777	Temporarily enable country combo
PEdney	10/03/2006		NA		Change requested and supplied by ING (Joyce Joseph)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


EPSOM Specific History

Prog    Date			AQR		Description
pct		07/03/2006		EP211	Postcode Field - set to Required/Not Required depending on BFPO Flag 
pct		13/03/2006		EP8		Various behaviour changes for BFPO Addresses
PB		13/04/2006		EP211	Enabled/Disabled BFPO checkbox depending on address type
PB		20/04/2006		EP211	Recoded to better integrate with existing code
PB		20/04/2006		EP389	(1) 'Date moved out' made mandatory if 'Previous address' selected
								(2) 'Date moved in' must now be before 'Date moved out'
								(3) 'Date moved in' no longer mandatory for correspondence addresses
PB		21/04/2006		EP422	Commented out unneccesary code which caused unexpected results
PB		26/04/2006		EP389	Date Moved In should not be mandatory for a correspondence address
PB		28/04/2006		EP467	Postcode field is not mandatory when page loads - should be based on BFPO flag.
PB		05/05/2006		EP467	(...cont) should also take account of country - no postcode req if overseas.
PB		10/05/2006		EP524	Removed "Either postcode or street..." validation as redundant now
PB		11/05/2006		EP538	PostCode field not always re-enabling when 'Clear' clicked.
PB		23/05/2006		EP586	Don't do EnableDisablePostcode if readonly, otherwise will enable fields which should be disabled
LH		25/05/2006		EP594	Current address must be a UK address
SAB		13/07/2006		EP991	Allow a user to click OK if no change made or the screen
								is in read-only mode.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ */

%>

<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>

<%/* FORMS */%>
<form id="frmToDC060" method="post" action="dc060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC066" method="post" action="dc066.asp" STYLE="DISPLAY: none"></form>
<form id="frmToZA010" method="post" action="za010.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 50px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 5px" class="msgLabel">
		Address of:
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<select id="cboApplicantName" name="ApplicantName" style="WIDTH: 280px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
		Address Type
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<select id="cboCustomerAddressType" name="CustomerAddressType" style="WIDTH: 280px" class="msgCombo">
			</select>
		</span>
	</span>
</div>

<div style="HEIGHT: 200px; LEFT: 10px; POSITION: absolute; TOP: 115px; WIDTH: 604px" class="msgGroup">
	<%/* Clear button */%>
	<span style="LEFT: 475px; POSITION: absolute; TOP: 5px">
		<input id="btnClearButton" value="Clear" type="button" style="WIDTH: 75px" class="msgButton">
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 5px" class="msgLabel">
		<LABEL id="idPostcode"></LABEL>
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtPostcode" name="Postcode" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgTxtUpper"
			 onchange="ClearAddressPAFIndicator()">
		</span>
	</span>

	<% /* BMIDS00336 MDC 27/08/2002 */ %>
	<% /* EP8 pct 13/03/2006 */ %>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 5px" class="msgLabel">
	<% /* EP8 pct 13/03/2006 - END */ %>
		BFPO
		<span style="LEFT: 10px; POSITION: absolute; TOP: -3px">
			<input id="chkBFPO" type="checkbox" name="BFPO" tabIndex="-1" style="POSITION: absolute; WIDTH: 70px" onclick="EnableDisablePostcode()"  disabled>
		</span>
	</span>
	<% /* BMIDS00336 MDC 27/08/2002 - End */ %>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
		<LABEL id="idFlatNo"></LABEL>
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtFlatNo" name="FlatNumber" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgTxt"
			 onchange="ClearAddressPAFIndicator()">
		</span>
	</span>

	<span id="spnHouseNameOuter" style="LEFT: 4px; POSITION: absolute; TOP: 50px" class="msgLabel">
		<LABEL id="idHouseName"></LABEL>
		<span id="spnHouseNumberInner" style="LEFT: 170px; POSITION: absolute; TOP: 3px">
			<input id="txtHouseNumber" name="HouseNumber" maxlength="10" style="POSITION: absolute; WIDTH: 45px" class="msgTxt"
			 onchange="ClearAddressPAFIndicator()">
		</span>
		<span id="spnHouseNameInner" style="LEFT: 220px; POSITION: absolute; TOP: 3px">
			<input id="txtHouseName" name="HouseName" maxlength="40" style="POSITION: absolute; WIDTH: 230px" class="msgTxt"
			 onchange="ClearAddressPAFIndicator()">
		</span>
	</span>
	
	<%/* PAF Search Button */%>
	<span style="LEFT: 475px; POSITION: absolute; TOP: 30px">
		<input id="btnPAFSearch" value="P.A.F.Search" type="button" style="WIDTH: 100px" class="msgButton">
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 80px" class="msgLabel">
		Street
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtStreet" name="Street" maxlength="40" style="POSITION: absolute; WIDTH: 280px" class="msgTxt"
			 onchange="ClearAddressPAFIndicator()">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 105px" class="msgLabel">
		<LABEL id="idDistrict"></LABEL>
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtDistrict" name="District" maxlength="40" style="POSITION: absolute; WIDTH: 280px" class="msgTxt"
			 onchange="ClearAddressPAFIndicator()">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 130px" class="msgLabel">
		<LABEL id="idTown"></LABEL>
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtTown" name="Town" maxlength="40" style="POSITION: absolute; WIDTH: 280px" class="msgTxt"
			 onchange="ClearAddressPAFIndicator()">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 155px" class="msgLabel">
		<LABEL id="idCounty"></LABEL>
		<span id="spnCounty" style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtCounty" name="County" maxlength="40" style="POSITION: absolute; WIDTH: 280px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 180px" class="msgLabel" >
		Country
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<select id="cboCountry" name="Country" style="WIDTH: 280px" class="msgCombo">
			</select>
		</span>
	</span>
</div>

<div style="HEIGHT: 125px; LEFT: 10px; POSITION: absolute; TOP: 320px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 5px" class="msgLabel">
		Nature of Occupancy
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<select id="cboOccupancy" name="Occupancy" style="WIDTH: 280px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
		Current/Estimated Property Value
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtPropertyValue" name="PropertyValue" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 50px" class="msgLabel">
		What are you going to do<br>with this property?
		<span style="LEFT: 170px; POSITION: absolute; TOP: 3px">
			<select id="cboIntendedAction" name="IntendedAction" style="WIDTH: 280px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 80px" class="msgLabel">
		Date Moved In
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtDateMovedIn" name="DateMovedIn" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 105px" class="msgLabel">
		Date Moved Out
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtDateMovedOut" name="DateMovedOut" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 475px; POSITION: absolute; TOP: 98px">
		<input id="btnTenancyButton" name="btnTenancyButton" value="Tenancy" type="button" style="WIDTH: 70px" class="msgButton">
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 460px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- #include FILE="attribs/dc065attribs.asp" --><!-- #include FILE="includes/pafsearch.asp" --><!-- #include FILE="customise/dc065customise.asp" --><!-- Specify Code Here -->
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_sXMLInAddress = null;
var m_bAddressPAFIndicator = false;
var bUpdatePAFIndicator = false;
var InAddressXML = null;
var SaveXML = null;
var scScreenFunctions;
var m_sReadOnly = null;
var m_blnReadOnly = false;

<% /* LH - EP495 */ %>
var m_isBtnSubmit = false;

<%/* SG 25/06/02 SYS4930 */%>
var m_bIsPopup = false;

<% 
/* SYS2564/SYS2752 Allow client variations. This class contains core versions of methods that 
	may need to be overriden by a client specific version.
*/ 
%>

<% /* PB EP467 - Update fields according to country */ %>
function frmScreen.cboCountry.onchange()
{
	ClearAddressPAFIndicator();
	EnableDisablePostcode();
}
<% /* EP467 End */ %>

<% /* PB EP211 13/04/2006 - Enable/Disable BFPO checkbox according to whether this is a previous address */ %>
function enableBFPOIfPreviousAddress()
{
	if(frmScreen.cboCustomerAddressType.value==3)
	{
		frmScreen.chkBFPO.disabled=false;
	}
	else
	{
		frmScreen.chkBFPO.disabled=true;
		<% /* PB EP211 13/04/2006 - Tidy-up in case box was checked when address-type was changed */ %>
		if(frmScreen.chkBFPO.checked==true)
		{
			frmScreen.chkBFPO.checked=false;
			EnableDisablePostcode();
		}
	}
}

function BaseScreen()
{
	<% // Base methods %>
	this.saveCounty = SaveCounty;
	this.loadCounty = LoadCounty;
	this.clearCounty = ClearCounty;
	this.getComboLists = GetComboLists;
	this.populateDerivedCombos = PopulateDerivedCombos;
	<% // Allow clients to override the number and type of combos to be obtained from the middle tier %>
	this.sGroupList = new Array("CustomerAddressType", "Country", "NatureOfOccupancy", "CurrentPropertyIntendedAction");
	<% // Allow clients to populate their own combos over and above the ones used already in this screen %>
	function PopulateDerivedCombos(XML){ return(true);}
	function SaveCounty(XML)
	{
		XML.CreateTag("COUNTY", frmScreen.txtCounty.value);	
	}
	function LoadCounty(XML)
	{
		frmScreen.txtCounty.value = XML.GetTagText("COUNTY");
	}
	function ClearCounty()
	{
		frmScreen.txtCounty.value = "";	
	}
	function GetComboLists()
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

		if(XML.GetComboLists(document,this.sGroupList))
		{
			XML.PopulateCombo(document,frmScreen.cboCustomerAddressType, "CustomerAddressType", true);
			XML.PopulateCombo(document,frmScreen.cboCountry, "Country", false);  //KRW BMIDS745
			//XML.PopulateCombo(document,frmScreen.cboCountry, "Country", true); 
			XML.PopulateCombo(document,frmScreen.cboOccupancy, "NatureOfOccupancy", true);
			XML.PopulateCombo(document,frmScreen.cboIntendedAction, "CurrentPropertyIntendedAction", true);
			this.populateDerivedCombos(XML);
		}
	}
}

var objDerivedOperations;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;

function window.onload()
{	
	<% /* MF This variable is to be set through work on Change control MARS004 */ %>
	var bPreviousAddress=true;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel","Another");
	<% // DJP SYS2564/SYS2752 %>
	objDerivedOperations = new DerivedScreen();
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Address Details","DC065",scScreenFunctions);

	RetrieveContextData();
	SetMasks(bPreviousAddress);

	<% // DJP SYS2564/SYS2752 %>
	Customise();
	Validation_Init();
	Initialise(true);

	//GD BMIDS00376 m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC065");
	if (m_sReadOnly=="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnClearButton.disabled =true;
		frmScreen.btnPAFSearch.disabled =true;
		<%/*frmScreen.btnTenancyButton.disabled=true;*/%>
		
		<% /* LH - EP595  */ %>
		btnAnother.disabled=false; 
		btnSubmit.disabled=false;
		
	} else
	{
		scScreenFunctions.SetFocusToFirstField(frmScreen);	
	}

	//GD BMIDS00376 m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC065");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
	if(!bPreviousAddress) DisableAddressFields();
	
	<% /* PB 23/05/2006 EP586
	// PB EP467 28/04/2006 Postcode field is not mandatory when page loads.
	EnableDisablePostcode();
	// EP467 End */ %>
	if(m_sReadOnly!="1") EnableDisablePostcode();
	<% /* EP586 End */ %>
}

<% /* MF MAR19 WP01 Check state of Global parameter RestrictUseOfCurrentOrCorAddress */ %>
function IsAddressRestriction()
{
	<% /* MF Read Global Parameter */ %>
	var ParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bRestrict = ParamXML.GetGlobalParameterBoolean(document,"RestrictUseOfCurrOrCorrAddress");	
	ParamXML = null;
	return bRestrict;
}

<% /* MF MAR19 WP01 Disable most of the address fields */ %>
function DisableFieldsForAddress()
{	
	frmScreen.btnClearButton.disabled=true;
	frmScreen.btnPAFSearch.disabled=true;
	//frmScreen.chkBFPO.disabled=true; // pct EP8 13/06/2006
	var aFields = new Array ("cboApplicantName","cboCustomerAddressType",
		"txtPostcode", "txtFlatNo", 
		"txtHouseName", "txtHouseNumber",
		"txtStreet", "txtDistrict", "txtTown", 
		"txtCounty" <% /* MAR777 GHun ,"cboCountry" */ %>);
	for(var i=0;i<aFields.length;i++){						
		document.all(aFields[i]).disabled=true;				
		switch (document.all(aFields[i]).tagName){
			case "select":
				document.all(aFields[i]).className="msgReadOnlyCombo";
				document.all(aFields[i]).disabled=true;
				break;				
			case "input":
				document.all(aFields[i]).disabled=true;					
				document.all(aFields[i]).className="msgReadOnly";					
				break;
			default:
				document.all(aFields[i]).className="msgReadOnly";
				break;
		}						
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

<% /* BMIDS00336 MDC 27/08/2002 */ %>
<% /* EP8 pct 13/03/2006 */ %>
function EnableDisablePostcode() <% /* PB 05/05/2006 EP467 - Renamed function to clarify purpose */ %>
{
	// If BFPO then postcode is not required, otherwise it is mandatory.
	if(frmScreen.chkBFPO.checked)
	{
		frmScreen.txtPostcode.value = "";
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtPostcode");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboCountry");
		frmScreen.btnPAFSearch.disabled=true;
		frmScreen.txtPostcode.removeAttribute("required");
		frmScreen.txtPostcode.parentElement.parentElement.style.color = "";
		frmScreen.txtFlatNo.focus();
	}
	else
	{
		// If overseas address then postcode is optional
		if(frmScreen.cboCountry.value==20)
		{	scScreenFunctions.SetFieldToWritable(frmScreen, "txtPostcode");
			scScreenFunctions.SetFieldToWritable(frmScreen, "cboCountry");
			frmScreen.btnPAFSearch.disabled=true;
			frmScreen.txtPostcode.removeAttribute("required");
			frmScreen.txtPostcode.parentElement.parentElement.style.color = "";
		}
		else
		{
			scScreenFunctions.SetFieldToWritable(frmScreen, "txtPostcode");
			//EP594 current address must be a UK address
			if(frmScreen.cboCustomerAddressType.value > 1)
			{
				scScreenFunctions.SetFieldToWritable(frmScreen, "cboCountry");
			}
			frmScreen.btnPAFSearch.disabled=false;
			frmScreen.txtPostcode.setAttribute("required", "true");
			frmScreen.txtPostcode.parentElement.parentElement.style.color = "Red";
			frmScreen.txtPostcode.focus();
		}
	}
}
<% /* EP8 pct 13/03/2006 END*/ %>
<% /* BMIDS00336 MDC 27/08/2002 - End */ %>

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	
	//SG 29/05/02 SYS4767 START
	//SG 27/02/02 SYS4186
	//if (m_sMetaAction == "Edit")
	if (m_sMetaAction == "Edit" | m_sMetaAction == "Copy")
	//SG 29/05/02 SYS4767 END
		//SYS2300
		//SYS1034 To make it consistent with other screens - changed back to idXML from idXML2
		m_sXMLInAddress	= scScreenFunctions.GetContextParameter(window, "idXML", null);
}

function Initialise(bOnLoad)
{
	if(bOnLoad)
	{
		<% // DJP SYS2564/SYS2752 %>
		objDerivedOperations.getComboLists();
		PopulateApplicantNameCombo();
	}		

	if(m_sMetaAction == "Edit")
	{
		PopulateScreen();
		DisableMainButton("Another");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboApplicantName");		
	}
	//SG 29/05/02 SYS4767 START
	else if (m_sMetaAction == "Copy")	//SG 27/02/02 SYS4186
	{									
		DefaultFields();				//Set the defaults
		PopulateScreen();				//In Copy mode we only populate some of them		
		DisableMainButton("Another");		
	}
	//SG 29/05/02 SYS4767 END
	else
		DefaultFields();

	// SYS0568 MC 30/03/2000. Amend = to == in the comparison
	if (frmScreen.cboApplicantName.options.length == 2)
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboApplicantName");

	UnhideOwnerOccupierDetails();
	DisableOrEnableTenancyButton();
	setupDateFieldsForAddressType();

	if (m_sReadOnly=="1") {
		frmScreen.btnClearButton.disabled = true;
		DisableMainButton("ANOTHER");
		}
	
	scScreenFunctions.SetFocusToFirstField(frmScreen); 
	
}

function DefaultFields()
{
	with (frmScreen)
	{	
		if (m_sMetaAction != "Edit")
			// SYS0568 MC 30/03/2000. Amend comparison to check no of items in list
			// if (cboApplicantName.length == 2)
			if (frmScreen.cboApplicantName.options.length == 2)
				<%/* Only one applicant in the dropdown */%>
				cboApplicantName.selectedIndex = 1;
			else
				cboApplicantName.selectedIndex = 0;
		cboCustomerAddressType.selectedIndex = 0;
		cboCustomerAddressType.value = "";
		txtPostcode.value = "";
		txtFlatNo.value = "";
		txtHouseName.value = "";
		txtHouseNumber.value = "";
		txtStreet.value = "";
		txtDistrict.value = "";
		txtTown.value = "";
		<% // DJP SYS2564/SYS2752 %>
		objDerivedOperations.clearCounty();
		cboCountry.selectedIndex = 0;
		cboOccupancy.value = "";
		txtPropertyValue.value = "";
		cboIntendedAction.value = "";
		txtDateMovedIn.value = "";
		txtDateMovedOut.value = "";		
				
		if (IsAddressRestriction()){
			RemoveHomeAddressFromCombo();		
		}
	}
}

function PopulateScreen()
{
	//we are in edit mode. Load the XML from context
	//OR we are in copy mode.
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(m_sXMLInAddress);
	XML.SelectTag(null, "CUSTOMERADDRESS");

	InAddressXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	InAddressXML.CreateRequestTag(window,null)
	InAddressXML.CreateActiveTag("CUSTOMERADDRESS");
	InAddressXML.CreateTag("CUSTOMERNUMBER", XML.GetTagText("CUSTOMERNUMBER"));
	InAddressXML.CreateTag("CUSTOMERVERSIONNUMBER", XML.GetTagText("CUSTOMERVERSIONNUMBER"));
	InAddressXML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER", XML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER"));							
	<% /* BMIDS00965 MDC 15/11/2002 - Remove screen rule processing from PopulateScreen */ %>
	InAddressXML.RunASP(document,"GetCustomerAddress.asp");
	<% /*
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			InAddressXML.RunASP(document,"GetCustomerAddress.asp");
			break;
		default: // Error
			InAddressXML.SetErrorResponse();
		}
	*/ %>
	<% /* BMIDS00965 MDC 15/11/2002 - End */ %>

	if(InAddressXML.SelectTag(null, "CUSTOMERADDRESS") != null)
	{
		
		
		//SG 29/05/02 SYS4767 START
		//SG 27/02/02 SYS4186
		//Placed IF statement around existing code - don't populate these if in Copy mode. 		
		if (m_sMetaAction != "Copy")
		{			
			// Populate the screen fields
			frmScreen.cboApplicantName.value = InAddressXML.GetTagText("CUSTOMERNUMBER");
			frmScreen.cboCustomerAddressType.value = InAddressXML.GetTagText("ADDRESSTYPE");
			
			<% /* PB EP211 20/04/2006 */ %>
			enableBFPOIfPreviousAddress();
			<% /* PB End of EP211 */ %>
			
			frmScreen.txtDateMovedIn.value = InAddressXML.GetTagText("DATEMOVEDIN");
			frmScreen.txtDateMovedOut.value = InAddressXML.GetTagText("DATEMOVEDOUT");
			frmScreen.cboOccupancy.value = InAddressXML.GetTagText("NATUREOFOCCUPANCY");

			frmScreen.txtPropertyValue.value = InAddressXML.GetTagText("ESTIMATEDCURRENTVALUE");
			frmScreen.cboIntendedAction.value = InAddressXML.GetTagText("INTENDEDACTION");
			m_bAddressPAFIndicator = (InAddressXML.GetTagText("PAFINDICATOR") == "1");
		}
		//SG 29/05/02 SYS4767 END

		frmScreen.txtPostcode.value = InAddressXML.GetTagText("POSTCODE");
		frmScreen.txtFlatNo.value = InAddressXML.GetTagText("FLATNUMBER");
		frmScreen.txtHouseName.value = InAddressXML.GetTagText("BUILDINGORHOUSENAME");
		frmScreen.txtHouseNumber.value = InAddressXML.GetTagText("BUILDINGORHOUSENUMBER");
		frmScreen.txtStreet.value = InAddressXML.GetTagText("STREET");
		frmScreen.txtDistrict.value = InAddressXML.GetTagText("DISTRICT");
		frmScreen.txtTown.value = InAddressXML.GetTagText("TOWN");
		<% // DJP SYS2564/SYS2752 %>
		objDerivedOperations.loadCounty(InAddressXML);
		frmScreen.cboCountry.value = InAddressXML.GetTagText("COUNTRY");

		<% /* BMIDS00336 MDC 27/08/2002 */ %>
		if(InAddressXML.GetTagText("BFPO") == "1") {
			frmScreen.chkBFPO.checked = true;
			}
		else {
			frmScreen.chkBFPO.checked = false;
			}
		EnableDisablePostcode();
		<% /* BMIDS00336 MDC 27/08/2002 - End */ %>
		
		<% /* MF 26/09/2005 WP01 MAR19 modifications */ %>
		if (m_sMetaAction == "Edit"){
	
			if (IsAddressRestriction()){				
				if ((scScreenFunctions.IsOptionValidationType(frmScreen.cboCustomerAddressType, 
						frmScreen.cboCustomerAddressType.selectedIndex, 'H')) || 
					scScreenFunctions.IsOptionValidationType(frmScreen.cboCustomerAddressType, 
						frmScreen.cboCustomerAddressType.selectedIndex, 'C'))
						
				{
					<% /* MF Disable address fields */ %>						
					MakeFieldsReadOnly(true);
			
				} else {
					RemoveHomeAddressFromCombo();
				}
			}
		}
				
		//SG 29/05/02 SYS54767
		//frmScreen.txtPropertyValue.value = InAddressXML.GetTagText("ESTIMATEDCURRENTVALUE");
		//frmScreen.cboIntendedAction.value = InAddressXML.GetTagText("INTENDEDACTION");
		//m_bAddressPAFIndicator = (InAddressXML.GetTagText("PAFINDICATOR") == "1");
	}		
}

<% /* MF 09/09/2005 Created rather than use scSCreenFunctions.html method as this
		did not give the desired appearance. This call gives a different look 
		to the background:
			scScreenFunctions.SetFieldToReadOnly(frmScreen,sField);  
*/ %>

function MakeFieldReadOnly(sField, bReadOnly){		
	document.all(sField).disabled=bReadOnly;				
	if(bReadOnly){
		document.all(sField).tabIndex=-1;
	}
	switch (document.all(sField).tagName){
		case "select":
			if(bReadOnly){
				document.all(sField).className="msgReadOnlyCombo";
			}else{
				document.all(sField).className="msgCombo";
			}
			break;				
		default:			
			if(bReadOnly){
				document.all(sField).className="msgReadOnly";
			}else{
				document.all(sField).className="msgTxt";
			}
			break;
	}						
}

function MakeFieldsReadOnly(bReadOnly){
	var aFields = new Array ("cboCustomerAddressType","txtPostcode",
							"txtFlatNo","txtHouseNumber","txtHouseName",
							"txtStreet","txtDistrict","txtTown",
							"txtCounty","cboCountry","txtTown");
							//EP594 - add cboCountry to array - current address must be a UK address
			
	for(var i=0;i<aFields.length;i++){						
		MakeFieldReadOnly(aFields[i],bReadOnly);						
	}
	<% /* MF Set the buttons readonly */ %>
	frmScreen.btnClearButton.disabled = bReadOnly;
	frmScreen.btnPAFSearch.disabled = bReadOnly;	
	
}

<% /* MF 26/09/2005 WP01 MAR19 modifications */ %>
function RemoveHomeAddressFromCombo()
{
	var iCount=0;
	for (iCount = frmScreen.cboCustomerAddressType.length - 1; iCount > 0; iCount--){
		<% /* MF - if this validation type is 'H' or 'C' remove it */ %>
		if (scScreenFunctions.IsOptionValidationType(frmScreen.cboCustomerAddressType, iCount, 'H')) {
			frmScreen.cboCustomerAddressType.remove(iCount);
		} else {
			if (scScreenFunctions.IsOptionValidationType(frmScreen.cboCustomerAddressType, iCount, 'C')) {
				frmScreen.cboCustomerAddressType.remove(iCount);
			}		
		}
	}
}

function PopulateApplicantNameCombo()
{
	// Add a <SELECT> option
	var TagOPTION	= document.createElement("OPTION");
	TagOPTION.value	= "";
	TagOPTION.text	= "<SELECT>";
	frmScreen.cboApplicantName.add(TagOPTION);

	// Loop through all customer context entries
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + nLoop, "One");
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop, "rf1111");
		var sCustomerVersionNumber= scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop, "2");
		if (sCustomerNumber != "")
		{
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= sCustomerNumber;
			TagOPTION.text = sCustomerName;
			TagOPTION.setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
			frmScreen.cboApplicantName.add(TagOPTION);
		}
	}
}

function ClearAddressPAFIndicator()
{
	m_bAddressPAFIndicator = false;
}

function frmScreen.cboOccupancy.onchange()
{
	UnhideOwnerOccupierDetails();
	DisableOrEnableTenancyButton();
}

function frmScreen.cboCustomerAddressType.onchange()
{
	UnhideOwnerOccupierDetails();
	
	// PB EP211 20/04/2006
	enableBFPOIfPreviousAddress();
	// PB End of EP211
	
	// PB EP389 20/04/2006
	setupDateFieldsForAddressType();
	// PB End of EP389
}

function setupDateFieldsForAddressType()
{
	// PB EP389 20/04/2006 26/04/2006 - rewrote to cover all options
	if(scScreenFunctions.IsValidationType(frmScreen.cboCustomerAddressType,"P")|scScreenFunctions.IsValidationType(frmScreen.cboCustomerAddressType,"TP"))
	{
		// Address type is 'previous address', so 'moved out' and 'moved in' dates should be mandatory
		frmScreen.txtDateMovedIn.setAttribute("required", "true");
		frmScreen.txtDateMovedIn.parentElement.parentElement.style.color = "Red";
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtDateMovedIn")
		frmScreen.txtDateMovedOut.setAttribute("required", "true");
		frmScreen.txtDateMovedOut.parentElement.parentElement.style.color = "Red";
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtDateMovedOut")
	}
	else
	{
		if(scScreenFunctions.IsValidationType(frmScreen.cboCustomerAddressType,"C"))
		{
			// Address type is 'Correspondence', so moved-in and moved-out dates should be optional
			frmScreen.txtDateMovedIn.removeAttribute("required");
			frmScreen.txtDateMovedIn.parentElement.parentElement.style.color = "";
			scScreenFunctions.SetFieldToWritable(frmScreen, "txtDateMovedIn")
			frmScreen.txtDateMovedOut.removeAttribute("required");
			frmScreen.txtDateMovedOut.parentElement.parentElement.style.color = "";
			scScreenFunctions.SetFieldToWritable(frmScreen, "txtDateMovedOut")
		}
		else
		{
			if(scScreenFunctions.IsValidationType(frmScreen.cboCustomerAddressType,"H")|scScreenFunctions.IsValidationType(frmScreen.cboCustomerAddressType,"TH"))
			{
				// Address type is current home or targeted home - mandatory moved-in but no moved-out
				frmScreen.txtDateMovedIn.setAttribute("required", "true");
				frmScreen.txtDateMovedIn.parentElement.parentElement.style.color = "Red";
				scScreenFunctions.SetFieldToWritable(frmScreen, "txtDateMovedIn")
				frmScreen.txtDateMovedOut.removeAttribute("required");
				frmScreen.txtDateMovedOut.value='';
				frmScreen.txtDateMovedOut.parentElement.parentElement.style.color = "";
				scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDateMovedOut");
			}
			else
			{
				if(scScreenFunctions.IsValidationType(frmScreen.cboCustomerAddressType,"S"))
				{
					// Address type is security address - both dates are optional
					frmScreen.txtDateMovedIn.removeAttribute("required");
					frmScreen.txtDateMovedIn.parentElement.parentElement.style.color = "";
					scScreenFunctions.SetFieldToWritable(frmScreen, "txtDateMovedIn")
					frmScreen.txtDateMovedOut.removeAttribute("required");
					frmScreen.txtDateMovedOut.parentElement.parentElement.style.color = "";
					scScreenFunctions.SetFieldToWritable(frmScreen, "txtDateMovedOut")
				}
				else
				{
					// Can only be 'Targeted previous but 1'
					frmScreen.txtDateMovedIn.setAttribute("required", "true");
					frmScreen.txtDateMovedIn.parentElement.parentElement.style.color = "Red";
					scScreenFunctions.SetFieldToWritable(frmScreen, "txtDateMovedIn")
					frmScreen.txtDateMovedOut.setAttribute("required", "true");
					frmScreen.txtDateMovedOut.parentElement.parentElement.style.color = "Red";
					scScreenFunctions.SetFieldToWritable(frmScreen, "txtDateMovedOut")
				}	
			}
		}
	}
	// PB End of EP389
}

function DisableOrEnableTenancyButton()
{
	if(scScreenFunctions.IsValidationType(frmScreen.cboOccupancy,"TU") || 
	   scScreenFunctions.IsValidationType(frmScreen.cboOccupancy,"TF"))
		//Tenancy type, so enable the next button
		frmScreen.btnTenancyButton.disabled = false;
	else
		frmScreen.btnTenancyButton.disabled = true;
}

function UnhideOwnerOccupierDetails()
//BG 18/05/00 if Home/Current is selected, clear date moved out and make it read only.
//BG 25/07 00 SYS0958 if Additional Current is selected, clear date moved out and make it read only. 
//SA 07/06/01 SYS1007 Changed If statement below as more than 1 condition could be true and requires action
{
	if(scScreenFunctions.IsValidationType(frmScreen.cboOccupancy,"O") & 
		!scScreenFunctions.IsValidationType(frmScreen.cboCustomerAddressType,"C") &
		!scScreenFunctions.IsValidationType(frmScreen.cboCustomerAddressType,"P"))
	{
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtPropertyValue");
		scScreenFunctions.SetFieldToWritable(frmScreen, "cboIntendedAction");
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtPropertyValue");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboIntendedAction");
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtDateMovedOut");
		frmScreen.txtPropertyValue.value = "";
		frmScreen.cboIntendedAction.selectedIndex = 0;
	}
	
	// EP499 PB 04/05/2006 - Merged in changes from MAR1598
	// MAR1598 / UAT 2423
	if (scScreenFunctions.IsValidationType(frmScreen.cboCustomerAddressType,"H") ||
	    scScreenFunctions.IsValidationType(frmScreen.cboCustomerAddressType,"A"))
	{
		frmScreen.txtDateMovedOut.value = "";
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDateMovedOut");
	}
}

function frmScreen.btnPAFSearch.onclick()
{

	bUpdatePAFIndicator == false;
	
	with (frmScreen)
		m_bAddressPAFIndicator = PAFSearch(txtPostcode,txtHouseName,txtHouseNumber,txtFlatNo,txtStreet,txtDistrict,
			txtTown,txtCounty,cboCountry);
			
	if(m_bAddressPAFIndicator == true)		// BMIDS710 KRW 14/04/2004
		bUpdatePAFIndicator = true;   
}

function ValidateScreen()
{
	function ValidateAddress()
	{
		with (frmScreen) {
			if (chkBFPO.checked == true) {
				if ((txtDistrict.value == "") & (txtTown.value == "") & (txtCounty.value == "")) {
					alert("A BFPO Address must have at least one of: District, Town or County populated with BFPO Address details.");
					return(false);
				}
			}
			else {
				<%/* Flat No., House Name, or House Number must be specified AND Postcode or Street+Town */%>
				if ((txtFlatNo.value == "") & (txtHouseName.value == "") & (HouseNumber.value == "")) {
					alert("At least one of Flat No., House Name or House Number must be specified.");
					frmScreen.txtFlatNo.focus();
					return(false);
				}	
				/* PB EP524 10/05/2006
				else if ((txtPostcode.value == "") & ((txtStreet.value == "") | (txtTown.value == ""))) {
					alert("Either Postcode or Street + Town must be specified.");
					frmScreen.txtPostcode.focus();
					return(false);
				}
				*/
				else if (!scScreenFunctions.ValidatePostcode(txtPostcode)) {
					return false;
				}
			}	
			return(true);
		}
	}

	function ValidateDates()
	{
		<% /* MO - BMIDS00807 */ %>
		var dteToday = scScreenFunctions.GetAppServerDate(true);
		<% /* var dteToday = new Date(); */ %>
		var dtActiveFrom = scScreenFunctions.GetDateObject(frmScreen.txtDateMovedIn);
		var dtActiveTo = scScreenFunctions.GetDateObject(frmScreen.txtDateMovedOut);

		//SYS1007 If Date moved in is Null - its an invalid format.
		// EP389 PB 26/04/2006 - Date Moved In is not mandatory for a correspondence address
		//if (dtActiveFrom == null)
		//	return(false);
		// EP389		
		
		<%/* Date moved in must not be in the future or more than a 100 years in the past */%>
		if (dtActiveFrom != null)
			if (scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtDateMovedIn,">"))
			{
				alert("Date moved in cannot be in the future.");
				frmScreen.txtDateMovedIn.focus();
				return(false);
			}
			else if (top.frames[1].document.all.scMathFunctions.GetYearsBetweenDates(dtActiveFrom,dteToday) >= 100)
			{
				alert("Date moved in cannot more than 100 years in the past.");
				frmScreen.txtDateMovedIn.focus();
				return(false);
			}

		<%/* Date moved out must not be in the future or more than a 100 years in the past */%>
		if (dtActiveTo != null)
			if (scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtDateMovedOut,">"))
			{
				alert("Date moved out cannot be in the future.");
				frmScreen.txtDateMovedOut.focus();
				return(false);
			}
			else if (top.frames[1].document.all.scMathFunctions.GetYearsBetweenDates(dtActiveTo,dteToday) >= 100)
			{
				alert("Date moved in cannot more than 100 years in the past.");
				frmScreen.txtDateMovedOut.focus();
				return(false);
			}

		<%/* The moved in date cannot be after the moved out date */%>
		if ((dtActiveFrom != null) & (dtActiveTo != null))
			<% /* PB EP389(2) 20/04/2006 */ %>
			if (scScreenFunctions.CompareDateFields(frmScreen.txtDateMovedIn,">=",frmScreen.txtDateMovedOut))
			{
				alert("'Date moved in' must be before 'date moved out'.");
			<% /* PB End of EP389(2) */ %>
				frmScreen.txtDateMovedIn.focus();
				return(false);
			}
		return(true);
	}

	var bReturn = false;

	if (ValidateAddress())
		bReturn = ValidateDates();
	else
		bReturn = false

	return(bReturn);
}

function frmScreen.btnTenancyButton.onclick()
{
	var bSuccess = false;
	if(frmScreen.onsubmit())
	{
		bSuccess = true;
		// if any changes have been made
		if (m_sReadOnly != "1" && IsChanged())
			if (ValidateScreen())
				bSuccess = CommitChanges();
			else
				bSuccess = false;
	}

	if(bSuccess)
	{
		scScreenFunctions.SetContextParameter(window, "idCustomerAddressType", frmScreen.cboCustomerAddressType.value);
		if (SaveXML != null)
			scScreenFunctions.SetContextParameter(window, "idXML", SaveXML.XMLDocument.xml);
		else
			scScreenFunctions.SetContextParameter(window, "idXML", InAddressXML.XMLDocument.xml);
		frmToDC066.submit();
	}
}

function btnSubmit.onclick()
{
	
	<% /* LH - EP595 Disable the button and display an hourglass until critical data check finishes */ %>
	btnSubmit.disabled = true;
	btnSubmit.blur();
	btnSubmit.style.cursor = "wait";
	btnAnother.disabled = true;
	btnAnother.blur();
	btnAnother.style.cursor = "wait";
	m_isBtnSubmit = true;
	
	var bSuccess = false;
	
	if (frmScreen.onsubmit() == true)
	{
		
		if (m_sReadOnly != "1" && IsChanged() || m_sReadOnly != "1" && bUpdatePAFIndicator == true) // BMIDS710 KRW 14/04/2004
		{
			if (ValidateScreen())
			{
				bSuccess = CommitChanges();
			}
			else
			{
				bSuccess = false;
			}
		}
		else
		{
			<%/* EP991 - SAB - Route to DC060 if data is unchanged or screen is read-only */%>
			bSuccess = true;
		}
	}
	else
	{

		/* EP595 LH */
		btnAnother.style.cursor = "hand";
		btnAnother.disabled = false;
		btnSubmit.style.cursor = "hand";
		btnSubmit.disabled = false;

		m_isBtnSubmit = false;
	}

	<% /* EP595 LH - Return to DC060 else stay on DC065 and change cursors back as before*/%>
	if (m_isBtnSubmit == true)
	{
		if(bSuccess)
		{
			frmToDC060.submit();
		}
		else
		{
			btnAnother.style.cursor = "hand";
			btnAnother.disabled = false;
			btnSubmit.style.cursor = "hand";
			btnSubmit.disabled = false;
		}
	}
}

function frmScreen.btnClearButton.onclick()
{
	with (frmScreen)
	{
		txtPostcode.value = "";
		txtFlatNo.value = "";
		txtHouseName.value = "";
		txtHouseNumber.value = "";
		txtStreet.value = "";
		txtDistrict.value = "";
		txtTown.value = "";
		<% // DJP SYS2564/SYS2752 %>
		objDerivedOperations.clearCounty();
		cboCountry.value = "";
		//SYS1007 Country should default to UK + set focus to postcode
		cboCountry.selectedIndex = 0;
		<% /* BMIDS00336 MDC 28/08/2002 */ %>
		chkBFPO.checked = false;
		<% /* EP538 PB 11/05/2006 */ %>
		EnableDisablePostcode();
		<% /* BMIDS00336 MDC 28/08/2002 */ %>
		txtPostcode.focus()
		m_bAddressPAFIndicator = false;
	}
}

function btnCancel.onclick()
{
	frmToDC060.submit();
}

function btnAnother.onclick()
{
	// disabled if meta action in context <> add.
	// calls save and then calls initialise for a new record...
	//SYS1007 First call validation methods
	var bSuccess = true;
	// if any changes have been made
	if (m_sReadOnly != "1" && IsChanged())
		if (ValidateScreen())
			if(CommitChanges())
				Initialise(false);
		
	
}

function CommitChanges()
{
	var bSuccess = true;
	
	//SYS1007 SA Before we save - need to check if a current address already exists
	
	if ( (scScreenFunctions.IsValidationType(frmScreen.cboCustomerAddressType,"H"))
	     && (m_sMetaAction == "Add"))
	{
		xmlAddress = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var xmlRequest, xmlCustomerAddressList ;
		// Get current address data by default; 
		xmlCustomerAddressList = xmlAddress.CreateActiveTag("CUSTOMERADDRESSLIST");
		xmlAddress.CreateActiveTag("CUSTOMERADDRESS");
		if(InAddressXML != null)
		{
			// edit mode and we haven't changed the applicant, so we have details we can use
			InAddressXML.SelectTag(null, "CUSTOMERADDRESS");
			xmlAddress.CreateTag("CUSTOMERNUMBER", InAddressXML.GetTagText("CUSTOMERNUMBER"));
			xmlAddress.CreateTag("CUSTOMERVERSIONNUMBER", InAddressXML.GetTagText("CUSTOMERVERSIONNUMBER"));
			xmlAddress.CreateTag("ADDRESSTYPE", 1);
		}
		else
		{
			xmlAddress.CreateTag("CUSTOMERNUMBER", frmScreen.cboApplicantName.value);
			var TagOption = frmScreen.cboApplicantName.options.item(frmScreen.cboApplicantName.selectedIndex);
			var sAttribute = TagOption.getAttribute("CustomerVersionNumber");
			xmlAddress.CreateTag("CUSTOMERVERSIONNUMBER", sAttribute);
			xmlAddress.CreateTag("ADDRESSTYPE", 1);
		}

		xmlAddress.RunASP(document,"FindCustomerAddressList.asp");
		
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = xmlAddress.CheckResponse(ErrorTypes);
		if ((ErrorReturn[1] == ErrorTypes[0]) | (xmlAddress.XMLDocument.text == ""))
		{
			//Error: record not found
			bSuccess =  true
		}
		else if (ErrorReturn[0] == true)
		{
			// No error
			xmlAddress.ActiveTag = null;
			xmlAddress.CreateTagList("CUSTOMERADDRESS");
			if (xmlAddress.ActiveTagList.length > 0)
			{
				alert("You have already entered a Home/Current address. Subsequent current addresses should be entered using address type Additional Current.");				
				frmScreen.cboCustomerAddressType.focus();				
				return(false);
			}
		}
	}
	//SYS1007 end
	SaveXML = null;
	SaveXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	SaveXML.CreateActiveTag("CUSTOMERADDRESS");

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null)

	//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD ROLL BACK 11/05/01 SYS2050
	// Add OPERATION TJ_30/03/2001 SYS2050
	XML.SetAttribute("OPERATION","CriticalDataCheck");
	XML.CreateActiveTag("CUSTOMER");
	// End OPERATION
	//end ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD END ROLL BACK SYS2050

	XML.CreateActiveTag("CUSTOMERADDRESSADDRESS");
	var tagCUSTOMERADDRESS = XML.CreateActiveTag("CUSTOMERADDRESS");
	
	//SG 29/05/02 SYS4767 START
	//SG 28/02/02 SYS4186
	//if(InAddressXML != null)
	if(InAddressXML != null && m_sMetaAction != "Copy")
	//SG 29/05/02 SYS4767 END
	{
		// edit mode and we haven't changed the applicant, so we have details we can use
		InAddressXML.SelectTag(null, "CUSTOMERADDRESS");
		XML.CreateTag("CUSTOMERNUMBER", InAddressXML.GetTagText("CUSTOMERNUMBER"));
		XML.CreateTag("CUSTOMERVERSIONNUMBER", InAddressXML.GetTagText("CUSTOMERVERSIONNUMBER"));
		XML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER", InAddressXML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER"));
	}
	else
	{
		XML.CreateTag("CUSTOMERNUMBER", frmScreen.cboApplicantName.value);
		var TagOption = frmScreen.cboApplicantName.options.item(frmScreen.cboApplicantName.selectedIndex);
		var sAttribute = TagOption.getAttribute("CustomerVersionNumber");
		XML.CreateTag("CUSTOMERVERSIONNUMBER", sAttribute);
		XML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER", "");
	}
	XML.CreateTag("ADDRESSTYPE", frmScreen.cboCustomerAddressType.value);
	XML.CreateTag("DATEMOVEDIN", frmScreen.txtDateMovedIn.value);
	XML.CreateTag("DATEMOVEDOUT", frmScreen.txtDateMovedOut.value);
	XML.CreateTag("NATUREOFOCCUPANCY", frmScreen.cboOccupancy.value);

	//SG 29/05/02 SYS4767 START
	//SG 28/02/02 SYS4186
	//if(InAddressXML != null)
	if(InAddressXML != null && m_sMetaAction != "Copy")
	//SG 29/05/02 SYS4767 END
		XML.CreateTag("ADDRESSGUID", InAddressXML.GetTagText("ADDRESSGUID"));
	else
		XML.CreateTag("ADDRESSGUID", "");
	//Set up SaveXML for use in Next processing
	SaveXML.CreateTag("CUSTOMERNUMBER", XML.GetTagText("CUSTOMERNUMBER"));
	SaveXML.CreateTag("CUSTOMERVERSIONNUMBER", XML.GetTagText("CUSTOMERVERSIONNUMBER"));
	if(m_sMetaAction == "Edit")
		SaveXML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER", XML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER"));

	XML.CreateActiveTag("ADDRESS");

	//SG 29/05/02 SYS4767 START
	//SG 28/02/02 SYS4186
	//if(InAddressXML != null)
	if(InAddressXML != null && m_sMetaAction != "Copy")
	//SG 29/05/02 SYS4767 END
		XML.CreateTag("ADDRESSGUID", InAddressXML.GetTagText("ADDRESSGUID"));
	else
		XML.CreateTag("ADDRESSGUID", "");
	XML.CreateTag("BUILDINGORHOUSENAME", frmScreen.txtHouseName.value);
	XML.CreateTag("BUILDINGORHOUSENUMBER", frmScreen.txtHouseNumber.value);
	XML.CreateTag("FLATNUMBER", frmScreen.txtFlatNo.value);
	XML.CreateTag("STREET", frmScreen.txtStreet.value);
	XML.CreateTag("DISTRICT", frmScreen.txtDistrict.value);
	XML.CreateTag("TOWN", frmScreen.txtTown.value);
	<% // DJP SYS2564/SYS2752 %>
	objDerivedOperations.saveCounty(XML);
	XML.CreateTag("COUNTRY", frmScreen.cboCountry.value);
	XML.CreateTag("POSTCODE", frmScreen.txtPostcode.value);
	XML.CreateTag("PAFINDICATOR", m_bAddressPAFIndicator ? "1" : "0");
	<% /* BMIDS00336 MDC 28/08/2002 */ %>
	XML.CreateTag("BFPO", frmScreen.chkBFPO.checked ? "1" : "0");
	<% /* BMIDS00336 MDC 28/08/2002 - End */ %>
			
	XML.ActiveTag = tagCUSTOMERADDRESS;
	XML.CreateActiveTag("CURRENTPROPERTY");
	
	//SG 29/05/02 SYS4767 START
	//SG 28/02/02 SYS4186
	//if(InAddressXML != null)
	if(InAddressXML != null && m_sMetaAction != "Copy")
	//SG 29/05/02 SYS4767 END
	{
		// edit mode and we haven't changed the applicant, so we have details we can use
		XML.CreateTag("CUSTOMERNUMBER", InAddressXML.GetTagText("CUSTOMERNUMBER"));
		XML.CreateTag("CUSTOMERVERSIONNUMBER", InAddressXML.GetTagText("CUSTOMERVERSIONNUMBER"));
		XML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER", InAddressXML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER"));
	}
	else
	{
		XML.CreateTag("CUSTOMERNUMBER", frmScreen.cboApplicantName.value);
		var TagOption = frmScreen.cboApplicantName.options.item(frmScreen.cboApplicantName.selectedIndex);
		var sAttribute = TagOption.getAttribute("CustomerVersionNumber");
		XML.CreateTag("CUSTOMERVERSIONNUMBER",sAttribute);
		XML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER", "");
	}
	XML.CreateTag("ESTIMATEDCURRENTVALUE", frmScreen.txtPropertyValue.value);
	XML.CreateTag("INTENDEDACTION", frmScreen.cboIntendedAction.value);

	if(m_sReadOnly != "1" && IsChanged() || m_sReadOnly != "1" && bUpdatePAFIndicator == true )// BMIDS710 KRW 14/04/2004
	{
		
			//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD ROLL BACK 11/05/01 SYS2050

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
		XML.SetAttribute("COMPONENT","omCust.CustomerBO");
		XML.SetAttribute("METHOD","SaveCustomerAddress");
		
		window.status = "Critical Data Check - please wait";
		
		// Added by automated update TW 09 Oct 2002 SYS5115
		// XML.RunASP(document,"OmigaTMBO.asp");
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
		//* XML.RunASP(document,"SaveCustomerAddress.asp");
		//end CRITICALDATACONTEXT  TJ_29/03/2001 
		//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD END ROLL BACK SYS2050 
		
		bSuccess = XML.IsResponseOK();
	}
	//SG 29/05/02 SYS4767 START
	// MSMS0042 Pass address sequence number to tenancy screen if copying.
	// if((m_sMetaAction == "Add") && bSuccess)
	if(((m_sMetaAction == "Add") || (m_sMetaAction == "Copy")) && bSuccess)
	// MSMS0042 - End.
	//SG 29/05/02 SYS4767 END
	{
		// CUSTOMERADDRESSSEQUENCENUMBER in add mode is returned here
		XML.SelectTag(null, "CUSTOMERADDRESS");
		SaveXML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER", XML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER"));
	}

	XML = null;
	return(bSuccess);
}
-->
</script>
</body>

</html>


<% /* OMIGA BUILD VERSION 045.04.05.10.01 */ %>
