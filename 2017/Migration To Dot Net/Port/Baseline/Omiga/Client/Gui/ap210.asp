<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP210.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		01/02/01	SYS1839 Created
DJP		19/03/01	SYS2114 Address buttons always enabled
ADP		22/03/01	SYS2151 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
DRC     21/11/01    SYS3048 Get Instruction Sequence No if not present in the Task Context from
                    the omiga menu context					
BG		03/11/01	SYS3048	moved "var m_ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();"
					from declarations at top of JScript to GetValuationReport.
AT		30/04/02	SYS4359 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
BG		28/06/02	SYS4767 MSMS/Core integration
SG		01/03/02	SYS4187 Added "Number of Occupants" and "Number of Kitchens"
SA		17/05/02	MSMS0071 Only allow 5 digits to be input in new fields added for SYS4187
BG		28/06/02	SYS4767 MSMS/Core integration - END

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

GD		26/06/2002	BMIDS00077 - applied SG 25/06/02 SYS4930
MV		05/07/2002	BMIDS00162 - Modified the Position of MsgButtons
DPF		29/08/2002	BMIDS00344 - Added / removed fields for APWP3
MO		13/11/2002	BMIDS00723	Made changes after the font sizes were changed
MO		19/11/2002	BMIDS01006  Made change for the flat validation type
GHun	16/01/2003	BM0243 No data displayed when InstructionSequenceNo is blank
DB		27/02/2003	BM0090 - Changed maxlength property of year built = 4
HMA     16/09/2003  BM0063      Amended HTML text for radio buttons
KRW     25/09/03    BMO063   Now makes screen read only when application locked out by other user
MC		20/04/2004	BMIDS517	AP225 popup dialog size changed.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
INGDUK Specific History:
JD		22/09/2005	MAR40 - ESurv valuation report changes
JD		23/11/2005	MAR571 - add New Property Indicator radio group and assoc. code.
PE		06/02/2006	MAR1352 - Need to make Property description mandatory
DRC     27/03/2006  MAR1402 - ING have made fields TypeofProperty mandatory - so must check if its empty on loading
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
TLiu	02/09/2005	   Changed layout for Flat No., House Name & No. - MAR38
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

EPSOM Specific History:
AW		03/05/2006	EP436 - Population of description combo related to property type 
PB		05/05/2006	EP429 - Added 'next' button (saves & jumps to ap220.asp)
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css"><title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT><!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<script src="validation.js" language="JScript"></script>
<!-- Screen Functions Object - see comments within scScreenFunctions.htm for details -->

<% /* Specify Forms Here */ %>
<form id="frmToAP205" method="post" action="AP205.asp" STYLE="DISPLAY: none"></form>
<% /* PB 05/05/2006 EP429 */ %>
<form id="frmToAP220" method="post" action="AP220.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate ="onchange"><!-- <div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 350px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
-->
	<div id="divPropertyDetails" style ="HEIGHT: 380px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel" >	
			<strong>Property Details
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 25px" class="msgLabel">
		Type of Property
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<select id="cboTypeOfProperty" style="WIDTH: 120px" class="msgCombo"></select>
			</span>
			<span style="LEFT: 240px; POSITION: absolute; TOP: -8px">
			<input id="btnFlatDetails" type ="button" style="HEIGHT: 26px; WIDTH: 26px" class="msgDDButton">
		</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 50px" class="msgLabel">
		Description
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<select id="cboDescription" style="WIDTH: 120px" class="msgCombo"></select>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 75px" class="msgLabel">	
			Year Built
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<input id="txtYearBuilt"  style="WIDTH: 40px"  class="msgTxt" maxLength=4>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 100px" class="msgLabel">
			Is this a New Property?
			<span style="LEFT: 116px; POSITION: absolute; TOP: -3px">
				<input id="NewProperty_Yes" name="NewPropertyInd" type="radio" value="1"><label for="NewProperty_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 184px; POSITION: absolute; TOP: -3px">
				<input id="NewProperty_No" name="NewPropertyInd" type="radio" value="0"><label for="NewProperty_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 125px" class="msgLabel">
		Tenure
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<select id="cboTenure" style="WIDTH: 120px" class="msgCombo"></select>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 150px" class="msgLabel" >	
			Unexpired Lease
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<input id="txtUnexpiredLease"  style="WIDTH: 100px"  class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 175px" class="msgLabel">
			Ground Rent
			<span style="LEFT: 115px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtGroundRent" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 200px" class="msgLabel" >	
			Service Charge
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">			
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtServiceCharge" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 225px" class="msgLabel" >	
			Chief Rent/Feu Duty
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtChiefRent" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px"  class="msgTxt">
			</span>
		</span>
<!--		<span style="LEFT: 4px; POSITION: absolute; TOP: 250px" class="msgLabel" >	
			Rental Income
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtRentalIncome" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px"  class="msgTxt">
			</span>
		</span>
-->
		<span style="LEFT: 4px; POSITION: absolute; TOP: 250px" class="msgLabel" >	
			Est Road Charge
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtEstRoadCharge" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px"  class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 275px" class="msgLabel">
		Central Heating
			<span style="LEFT: 116px; POSITION: absolute; TOP: -3px">
				<input id="CentralHeating_Yes" name="CentralHeating" type="radio" value="1"><label for="CentralHeating_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 184px; POSITION: absolute; TOP: -3px">
				<input id="CentralHeating_No" name="CentralHeating" type="radio" value="0"><label for="CentralHeating_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 295px" class="msgLabel">
			Ex Local Authority
			<span style="LEFT: 116px; POSITION: absolute; TOP: -3px">
				<input id="ExLocalAuthority_Yes" name="ExLocalAuthority" type="radio" value="1"><label for="ExLocalAuthority_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 184px; POSITION: absolute; TOP: -3px">
				<input id="ExLocalAuthority_No" name="ExLocalAuthority" type="radio" value="0"><label for="ExLocalAuthority_No" class="msgLabel">No</label>
			</span>
			<!-- EP2_2 New field -->
			<span style="LEFT: 300px;  POSITION: absolute; TOP: -3px" class="msgLabel">	
				% in Private Ownership
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtPrivateOwnership"  style="WIDTH: 155px"  class="msgTxt" NAME="txtPrivateOwnership">
				</span>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 318px" class="msgLabel">
			Instruction<BR>Address Correct?
			<span style="LEFT: 116px; POSITION: absolute; TOP: 3px">
				<input id="InstructionAddressCorrect_Yes" name="InstructionAddressCorrect" type="radio" value="1"><label for="InstructionAddressCorrect_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 184px; POSITION: absolute; TOP: 3px">
				<input id="InstructionAddressCorrect_No" name="InstructionAddressCorrect" type="radio" value="0"><label for="InstructionAddressCorrect_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 350px" class="msgLabel">
			Building Standard<BR>Indemnity Scheme
			<span style="LEFT: 115px; POSITION: absolute; TOP: 2px">
				<select id="cboBuildingScheme" style="WIDTH: 120px" class="msgCombo"></select>
			</span>
		</span>
		<!-- EP2_2 New options -->
		<span style="LEFT: 300px; POSITION: absolute; TOP: 350px" class="msgLabel">
		Sales incentive included<br>in purchase price?
			<span style="LEFT: 135px; POSITION: absolute; TOP: 3px">
				<input id="SalesIncentive_Yes" name="SalesIncentive" type="radio" value="1"><label for="SalesIncentive_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 189px; POSITION: absolute; TOP: 3px">
				<input id="SalesIncentive_No" name="SalesIncentive" type="radio" value="0"><label for="SalesIncentive_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 300px; POSITION: absolute; TOP: 385px" class="msgLabel">
		Commercial Usage?
			<span style="LEFT: 135px; POSITION: absolute; TOP: -3px">
				<input id="CommercialUse_Yes" name="CommercialUse" type="radio" value="1"><label for="CommercialUse_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 189px; POSITION: absolute; TOP: -3px">
				<input id="CommercialUse_No" name="CommercialUse" type="radio" value="0"><label for="CommercialUse_No" class="msgLabel">No</label>
			</span>
		</span>
		<!-- address details -->
		<span id="spnAddress">
			<span style="LEFT: 300px;  POSITION: absolute; TOP: 4px" class="msgLabel">	
				PostCode
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtPostcode"  style="HEIGHT: 18px;  WIDTH: 60px" class="msgTxt">
				</span>
			</span>         
			<span style="LEFT: 300px;  POSITION: absolute; TOP: 25px" class="msgLabel">	
				Flat No./ Name
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtFlatNo"  style="WIDTH: 60px"  class="msgTxt">
				</span>
			</span>
	
			<span style="LEFT: 300px;  POSITION: absolute; TOP: 50px" class="msgLabel">	
				House No. and Name
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtHouseNo"  style="WIDTH: 35px"  class="msgTxt">
				</span>
				<span style="LEFT: 158px; POSITION: absolute; TOP: -3px">
					<input id="txtHouseName"  style="WIDTH: 117px"  class="msgTxt">
				</span>
			</span>
			<span style="LEFT: 300px;  POSITION: absolute; TOP: 75px" class="msgLabel">	
				Street
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtStreet"  style="WIDTH: 155px"  class="msgTxt">
				</span>
			</span>
		
			<span style="LEFT: 300px;  POSITION: absolute; TOP: 100px" class="msgLabel">	
				District
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtDistrict"  style="WIDTH: 155px"  class="msgTxt">
				</span>
			</span>
			<span style="LEFT: 300px;  POSITION: absolute; TOP: 125px" class="msgLabel">	
				Town
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtTown"  style="WIDTH: 155px"  class="msgTxt">
				</span>
			</span>
	
			<span style="LEFT: 300px;  POSITION: absolute; TOP: 150px" class="msgLabel">	
				County
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtCounty"  style="WIDTH: 155px"  class="msgTxt">
				</span>
			</span>
			<span style="LEFT: 300px;  POSITION: absolute; TOP: 175px" class="msgLabel">	
				Country
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<select id="cboCountry" style="WIDTH: 120px" class="msgCombo"></select>
				</span>
			</span>
		</span>
		
		<span style="LEFT: 490px;  POSITION: absolute; TOP: 26px">
			<input id="btnPAFSearch" value="P.A.F. Search" type="button" style="HEIGHT: 20px; WIDTH: 83px" class="msgButton">
		</span>
		<span style="LEFT: 490px; POSITION: absolute; TOP: 2px">
			<input id="btnClearAddress" value="Clear" type="button" style="HEIGHT: 20px; WIDTH: 83px" class="msgButton">
		</span>
		<span style="LEFT: 300px; POSITION: absolute; TOP: 220px">
			<input id="btnRooms" value="Rooms" type="button" style="HEIGHT: 30px; WIDTH: 103px" class="msgButton" NAME="btnRooms">
		</span><!-- </div> --><!-- Floor Areas Part -->
		<!-- EP2_2 New button -->
		<span style="LEFT: 420px; POSITION: absolute; TOP: 220px">
			<input id="btnBuyToLet" value="Buy To Let" type="button" style="HEIGHT: 30px; WIDTH: 103px" class="msgButton" NAME="btnBuyToLet">
		</span><!-- </div> --><!-- Floor Areas Part -->
    <div id="divFloorAreas" style="HEIGHT: 200px; LEFT: 0px; POSITION: absolute; TOP: 405px; WIDTH: 604px" class="msgGroup">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">	
			<strong>Floor Area in Sq. Meters
		</span>	    		
		<!-- EP2_2 New option -->
		<span style="LEFT: 300px; POSITION: absolute; TOP: 3px" class="msgLabel">
		Live work unit with a<br>residential element > 70%?
			<span style="LEFT: 135px; POSITION: absolute; TOP: 3px">
				<input id="HighResUse_Yes" name="HighResUse" type="radio" value="1"><label for="HighResUse_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 189px; POSITION: absolute; TOP: 3px">
				<input id="HighResUse_No" name="HighResUse" type="radio" value="0"><label for="HighResUse_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 25px" class="msgLabel">	
			Gross External Floor Area
			<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
				<input id="txtExtFloorArea"  style="WIDTH: 100px"  class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 50px" class="msgLabel">	
		    Land Area if &gt;1 acre
			<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
				<input id="txtLandArea"  style="WIDTH: 120px"  class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 75px" class="msgLabel">	
			<strong>Construction
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 90px" class="msgLabel">
		Structure
			<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
				<select id="cboStructure" style="WIDTH: 120px" class="msgCombo"></select>
			</span>
		</span>
		<span style="LEFT: 300px; POSITION: absolute; TOP: 90px" class="msgLabel">	
		    Non Standard
			<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
				<input id="txtNonStandard"  style="WIDTH: 155px"  class="msgTxt" maxLength=60>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 115px" class="msgLabel">
		Main Roof
			<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
				<select id="cboMainRoof" style="WIDTH: 120px" class="msgCombo"></select>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 140px" class="msgLabel">
			Single Skin Construction
			<span style="LEFT: 126px; POSITION: absolute; TOP: -3px">
				<input id="SingleSkinConstruction_Yes" name="SingleSkinConstruction" type="radio" value="1"><label for="SingleSkinConstruction_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 194px; POSITION: absolute; TOP: -3px">
				<input id="SingleSkinConstruction_No" name="SingleSkinConstruction" type="radio" value="0"><label for="SingleSkinConstruction_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 165px" class="msgLabel">
			Single Skin 2 Storey
			<span style="LEFT: 126px; POSITION: absolute; TOP: -3px">
				<input id="SingleSkin2Storey_Yes" name="SingleSkin2Storey" type="radio" value="1"><label for="SingleSkin2Storey_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 194px; POSITION: absolute; TOP: -3px">
				<input id="SingleSkin2Storey_No" name="SingleSkin2Storey" type="radio" value="0"><label for="SingleSkin2Storey_No" class="msgLabel">No</label>
			</span>
		</span>
	</div>
</div>				
</form>
<%/* Main Buttons */ %>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 650px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- #include FILE="includes/pafsearch.asp" -->
<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP210attribs.asp" --><!-- #include FILE="Customise/AP210Customise.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
//BG 03/11/2001 SYS3048 can't var as new here
var m_ValuationXML = null;
var m_TaskXML = null;
var m_sRoomsXML = null; //DPF APWP3 29/9/02
var m_sFlatXML = null; //DPF APWP3 29/9/02
var m_sBTLXML = "";  //EP2_2 Buy to Let XML
var m_sTaskXML = "";
var m_sAppNo = "";
var m_sAppFactFindNo = "";
var m_sTaskID = "";
var m_sInsSeqNo = "";
var csValuerTypeStaff = "10";
var csTaskComplete = "40";
var csTaskPending = "20";
var csTaskNotStarted = "10";
var csTenureLeashold = "2";
var m_sCallingScreen = "";
var m_bPAFIndicator = "0";
var m_sReadOnly = null;
var m_sReportNotFound = "0";
var XMLRooms = new top.frames[1].document.all.scXMLFunctions.XMLObject(); // DPF APWP3 2/9/02
var XMLFlat = new top.frames[1].document.all.scXMLFunctions.XMLObject(); // DPF APWP3 2/9/02
var XMLBTL = new top.frames[1].document.all.scXMLFunctions.XMLObject(); // EP2_2 AShaw 23/11/2006
//GD BMIDS0077
<%/* SG 25/06/02 SYS4930 */%>
var m_bIsPopup = false;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Cancel","Next");

	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Valuation Report - Property Details","AP210",scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
	Customise();
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();

	PopulateScreen();
 	scScreenFunctions.SetFocusToFirstField(frmScreen);

	if(m_sReadOnly != "1")  // This if statement was rem'd out ? , replaced to reflect read only status KRW 25/09/03
	{
		var bReadOnly = false;
		bReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP210");	
		if(bReadOnly)
		{
			m_sReadOnly = "1";		
		}
	}
	if(	m_sReadOnly == "1")
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			frmScreen.btnPAFSearch.disabled = true;
			frmScreen.btnClearAddress.disabled = true;    
		}	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function PopulateScreen()
{
	var bSuccess;

	PopulateCombos();
	bSuccess = GetValuationReport();

	if(bSuccess)
	{
		// Populate the screen fields (DPF 29/08/02 - changed order to be same as screen)
		frmScreen.cboTypeOfProperty.value = m_ValuationXML.GetAttribute("TYPEOFPROPERTY");
		frmScreen.txtYearBuilt.value = m_ValuationXML.GetAttribute("YEARBUILT");
		<% /* JD MAR571 */ %>
		SetOptionValue(frmScreen.NewProperty_Yes, frmScreen.NewProperty_No, m_ValuationXML.GetAttribute("NEWPROPERTYINDICATOR"));
		frmScreen.cboTenure.value = m_ValuationXML.GetAttribute("TENURE");
		frmScreen.txtUnexpiredLease.value = m_ValuationXML.GetAttribute("UNEXPIREDLEASE");	
		frmScreen.txtGroundRent.value = m_ValuationXML.GetAttribute("GROUNDRENT");
		frmScreen.txtServiceCharge.value  = m_ValuationXML.GetAttribute("SERVICECHARGE");
		//APWP3 - DPF 29/08/02 - 3 new text fields & 1 radio option added
		frmScreen.txtChiefRent.value  = m_ValuationXML.GetAttribute("CHIEFRENT");
//EP2_2	frmScreen.txtRentalIncome.value  = m_ValuationXML.GetAttribute("RENTALINCOME");
		frmScreen.txtEstRoadCharge.value  = m_ValuationXML.GetAttribute("ESTROADCHARGE");
		
		SetOptionValue(frmScreen.CentralHeating_Yes, frmScreen.CentralHeating_No, m_ValuationXML.GetAttribute("HOTWATERCENTRALHEATING"));
		SetOptionValue(frmScreen.ExLocalAuthority_Yes, frmScreen.ExLocalAuthority_No, m_ValuationXML.GetAttribute("EXLOCALAUTHORITY")); //JD MAR40
		SetOptionValue(frmScreen.InstructionAddressCorrect_Yes, frmScreen.InstructionAddressCorrect_No, m_ValuationXML.GetAttribute("INSTRUCTIONADDRESSCORRECT")); //JD MAR40
		frmScreen.cboBuildingScheme.value = m_ValuationXML.GetAttribute("CERTIFICATIONTYPE"); //JD MAR40
				
		frmScreen.txtPostcode.value = m_ValuationXML.GetAttribute("POSTCODE");
		frmScreen.txtFlatNo.value = m_ValuationXML.GetAttribute("FLATNUMBER");
		frmScreen.txtHouseName.value = m_ValuationXML.GetAttribute("BUILDINGORHOUSENAME");		
		frmScreen.txtHouseNo.value = m_ValuationXML.GetAttribute("BUILDINGORHOUSENUMBER");	
		frmScreen.txtStreet.value = m_ValuationXML.GetAttribute("STREET");						
		frmScreen.txtDistrict.value = m_ValuationXML.GetAttribute("DISTRICT");		
		frmScreen.txtTown.value = m_ValuationXML.GetAttribute("TOWN");
		frmScreen.txtCounty.value = m_ValuationXML.GetAttribute("COUNTY");
		frmScreen.cboCountry.value = m_ValuationXML.GetAttribute("COUNTRY");
		
		frmScreen.txtLandArea.value = m_ValuationXML.GetAttribute("RESIDENTIALOUTBUILDINGSAREA");
		frmScreen.txtExtFloorArea.value = m_ValuationXML.GetAttribute("RESIDENCEAREA");
				
		frmScreen.cboStructure.value = m_ValuationXML.GetAttribute("STRUCTURE");
		frmScreen.txtNonStandard.value = m_ValuationXML.GetAttribute("NONSTANDARDCONSTRUCTIONTYPE"); //JD MAR40
		frmScreen.cboMainRoof.value = m_ValuationXML.GetAttribute("MAINROOF");
		SetOptionValue(frmScreen.SingleSkinConstruction_Yes, frmScreen.SingleSkinConstruction_No, m_ValuationXML.GetAttribute("ISSINGLESKINCONSTRUCTION")); //JD MAR40
		SetOptionValue(frmScreen.SingleSkin2Storey_Yes, frmScreen.SingleSkin2Storey_No, m_ValuationXML.GetAttribute("ISSINGLESKINTWOSTOREY")); //JD MAR40
		
		//APWP3 - DPF 02/09/02 - Check what type of property we're dealing with
		<% /* MAR1402 - Fields are mandatory in MARS */ %>
		if (m_ValuationXML.GetAttribute("TYPEOFPROPERTY").length > 0)
			frmScreen.cboTypeOfProperty.onchange();
		if (m_ValuationXML.GetAttribute("STRUCTURE").length > 0)	
			frmScreen.cboStructure.onchange();//JD MAR40
		<% /* MAR1402 - End */ %>
		
		<%/*AW	EP436 03/05/2006 */%>
		frmScreen.cboDescription.value = m_ValuationXML.GetAttribute("PROPERTYDESCRIPTION");
		
		//EP2_2 New fields
		frmScreen.txtPrivateOwnership.value = m_ValuationXML.GetAttribute("EXLAPERCENTAGEPRIVATEOWNER"); 
		SetOptionValue(frmScreen.SalesIncentive_Yes, frmScreen.SalesIncentive_No, m_ValuationXML.GetAttribute("PPINCLSALESINCENTIVES")); 
		SetOptionValue(frmScreen.CommercialUse_Yes, frmScreen.CommercialUse_No, m_ValuationXML.GetAttribute("COMMERCIALUSAGE")); 
		SetOptionValue(frmScreen.HighResUse_Yes, frmScreen.HighResUse_No, m_ValuationXML.GetAttribute("LIVEWORKUNITRESELEMENT")); 
		// Enable Private Ownership % if ExLocalAuthority_Yes
		var bEnableELA = frmScreen.ExLocalAuthority_Yes.checked;
		ExLAPercent(bEnableELA);

	}	
	else
		m_sReportNotFound = "1";
	HandleTenure();
	return(bSuccess);
}

function GetValuationReport()
{
	//BG 03/11/2001 SYS3048 
	m_ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = m_ValuationXML.CreateRequestTag(window , "GetValuationReport");
	var bSuccess = false;
	
	m_ValuationXML.CreateActiveTag("VALUATION");
	m_ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	m_ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
	m_ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
	// AQR SYS3048 - if Instruction sequence number is not in CASETASK context then
	//               get it from the main context (set in AP205)
	if (m_sInsSeqNo == "" )
		m_sInsSeqNo = scScreenFunctions.GetContextParameter(window,"idInstructionSequenceNo",null)
	
	<% /* BM0243 Only set InstructionSequenceNo if it is not blank */ %>
	if (m_sInsSeqNo.length > 0)
		m_ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO",m_sInsSeqNo);
	<% /* BM0243 End */ %>

	m_ValuationXML.RunASP(document, "omAppProc.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = m_ValuationXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)	
	{
		if(m_ValuationXML.SelectTag(null, "GETVALUATIONREPORT") != null)
		{
			bSuccess = true;
		}
	}
	return(bSuccess);
}

function PopulateCombos()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("PropertyType","PropertyDescription","PropertyTenure","BuildingConstruction","RoofConstruction","Country", "BuildingsCertificationType");

	if(XML.GetComboLists(document,sGroupList))
	{
		XML.PopulateCombo(document,frmScreen.cboTypeOfProperty,"PropertyType",true);
		XML.PopulateCombo(document,frmScreen.cboDescription ,"PropertyDescription",true);
		XML.PopulateCombo(document,frmScreen.cboTenure,"PropertyTenure",true);
		XML.PopulateCombo(document,frmScreen.cboStructure,"BuildingConstruction",true);
		XML.PopulateCombo(document,frmScreen.cboMainRoof ,"RoofConstruction",true);
		XML.PopulateCombo(document,frmScreen.cboCountry ,"Country",true);
		XML.PopulateCombo(document,frmScreen.cboBuildingScheme ,"BuildingsCertificationType",true);//JD MAR40
	}

	XML = null;
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function RetrieveContextData()
{
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	if(m_sTaskXML != "")
	{
		m_TaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_TaskXML.LoadXML(m_sTaskXML);
		m_TaskXML.SelectTag(null, "CASETASK");
		m_sInsSeqNo = m_TaskXML.GetAttribute("CONTEXT");
		m_sTaskID = m_TaskXML.GetAttribute("TASKID")	
	}

	m_sAppNo = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sAppFactFindNo = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sCallingScreen = scScreenFunctions.GetContextParameter(window,"idREturnScreenId",null);

//	if(m_sReadOnly != "1")
//	{
//		var sRet;
		sRet = scScreenFunctions.GetContextParameter(window,"idMetaAction", null);	
		if(sRet == "1")
		{
			m_sReadOnly = "1";
//			scScreenFunctions.SetScreenToReadOnly(frmScreen);
//			frmScreen.btnPAFSearch.disabled = true;
//			frmScreen.btnClearAddress.disabled = true;
		}
//	}
}

function CommitChanges()
{
	var bSuccess = true;
	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				bSuccess = SaveValuation();
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function SaveScreenValues( XML )
{
	var bSuccess;
	var sVal;
	
	bSuccess = false
	if(XML != null)
	{
		XML.SetAttribute("TYPEOFPROPERTY", frmScreen.cboTypeOfProperty.value);
		XML.SetAttribute("PROPERTYDESCRIPTION", frmScreen.cboDescription.value);
		XML.SetAttribute("YEARBUILT",frmScreen.txtYearBuilt.value);
		<% /* JD MAR571 */ %>
		XML.SetAttribute("NEWPROPERTYINDICATOR", GetOptionValue(frmScreen.NewProperty_Yes));
		XML.SetAttribute("TENURE", frmScreen.cboTenure.value);
		XML.SetAttribute("UNEXPIREDLEASE",frmScreen.txtUnexpiredLease.value);
		XML.SetAttribute("GROUNDRENT",frmScreen.txtGroundRent.value);
		XML.SetAttribute("SERVICECHARGE",frmScreen.txtServiceCharge.value);
		
		//APWP3 - DPF 30/08/2002 - 3 new text fields & 1 new radio option added
		XML.SetAttribute("CHIEFRENT",frmScreen.txtChiefRent.value);
//EP2_2	XML.SetAttribute("RENTALINCOME",frmScreen.txtRentalIncome.value);
		XML.SetAttribute("ESTROADCHARGE",frmScreen.txtEstRoadCharge.value);
		
		XML.SetAttribute("HOTWATERCENTRALHEATING", GetOptionValue(frmScreen.CentralHeating_Yes));
		XML.SetAttribute("EXLOCALAUTHORITY", GetOptionValue(frmScreen.ExLocalAuthority_Yes)); //JD MAR40
		XML.SetAttribute("INSTRUCTIONADDRESSCORRECT", GetOptionValue(frmScreen.InstructionAddressCorrect_Yes)); //JD MAR40
		XML.SetAttribute("CERTIFICATIONTYPE", frmScreen.cboBuildingScheme.value); //JD MAR40

								
		XML.SetAttribute("POSTCODE",frmScreen.txtPostcode.value);
		XML.SetAttribute("FLATNUMBER", frmScreen.txtFlatNo.value);
		XML.SetAttribute("BUILDINGORHOUSENAME",frmScreen.txtHouseName.value);
		XML.SetAttribute("BUILDINGORHOUSENUMBER",frmScreen.txtHouseNo.value);
		XML.SetAttribute("STREET",frmScreen.txtStreet.value);
		XML.SetAttribute("DISTRICT", frmScreen.txtDistrict.value);
		XML.SetAttribute("TOWN",frmScreen.txtTown.value);
		XML.SetAttribute("COUNTY", frmScreen.txtCounty.value);
		XML.SetAttribute("COUNTRY", frmScreen.cboCountry.value);
		
		XML.SetAttribute("RESIDENCEAREA",frmScreen.txtExtFloorArea.value);
		XML.SetAttribute("RESIDENTIALOUTBUILDINGSAREA",frmScreen.txtLandArea.value);
				
		XML.SetAttribute("STRUCTURE", frmScreen.cboStructure.value);
		XML.SetAttribute("MAINROOF", frmScreen.cboMainRoof.value);
		XML.SetAttribute("NONSTANDARDCONSTRUCTIONTYPE", frmScreen.txtNonStandard.value); //JD MAR40
		XML.SetAttribute("ISSINGLESKINCONSTRUCTION", GetOptionValue(frmScreen.SingleSkinConstruction_Yes)); //JD MAR40
		XML.SetAttribute("ISSINGLESKINTWOSTOREY", GetOptionValue(frmScreen.SingleSkin2Storey_Yes)); //JD MAR40

		//EP2_2 New fields
		XML.SetAttribute("EXLAPERCENTAGEPRIVATEOWNER", frmScreen.txtPrivateOwnership.value); 
		XML.SetAttribute("PPINCLSALESINCENTIVES", GetOptionValue(frmScreen.SalesIncentive_Yes));
		XML.SetAttribute("COMMERCIALUSAGE", GetOptionValue(frmScreen.CommercialUse_Yes));
		XML.SetAttribute("LIVEWORKUNITRESELEMENT", GetOptionValue(frmScreen.HighResUse_Yes)); 

		//DPF 02/09/2002 - save Room Details
		SaveRoomDetails(XML);
		
		//DPF 02/09/2002 - if property type is a flat then save flat details
		if (scScreenFunctions.GetComboValidationType(frmScreen, "cboTypeOfProperty") == "F")
		{	
			SaveFlatDetails(XML);
		}
		
		//EP2_2 If type is a Buy To Let then save BTL details
		if (m_sBTLXML != "")
			SaveBTLDetails(XML);
		
		bSuccess = true;
	}
	
	return(bSuccess);
}

function SaveRoomDetails(XML)
{
//DPF 02/09/02 - save rooms details record on AP215.asp
	XMLRooms.LoadXML(m_sRoomsXML);

	if(XMLRooms.SelectTag(null, "ROOMS") != null)
	{
		XML.SetAttribute("NUMBEROFBEDROOMS", XMLRooms.GetAttribute("BEDROOMS"));
		XML.SetAttribute("LIVINGROOMS", XMLRooms.GetAttribute("LIVINGROOMS"));
		XML.SetAttribute("BATHROOMS", XMLRooms.GetAttribute("BATHROOMS"));
		XML.SetAttribute("SEPERATEWCS", XMLRooms.GetAttribute("SEPERATEWCS"));
		XML.SetAttribute("NUMBEROFKITCHENS", XMLRooms.GetAttribute("NUMBEROFKITCHENS")); //JD MAR40
		XML.SetAttribute("GARAGES", XMLRooms.GetAttribute("GARAGES"));
		XML.SetAttribute("PARKINGSPACES", XMLRooms.GetAttribute("PARKINGSPACES"));
		XML.SetAttribute("HABITABLEROOMS", XMLRooms.GetAttribute("HABITABLEROOMS"));
		XML.SetAttribute("CONSERVATORY", XMLRooms.GetAttribute("CONSERVATORY"));
		XML.SetAttribute("GARDENINCLUDED", XMLRooms.GetAttribute("GARDENINCLUDED"));
		//EP2_ Save new values.
		XML.SetAttribute("FLOORS", XMLRooms.GetAttribute("FLOORS"));
		XML.SetAttribute("BALCONYACCESS", XMLRooms.GetAttribute("BALCONYACCESS"));
	}
}

function SaveFlatDetails(XML)
{
//DPF 02/09/02 - save flat details record on AP225.asp.  If it's not a flat ensure blank details are passed thru
	XMLFlat.LoadXML(m_sFlatXML);
	
	if (scScreenFunctions.GetComboValidationType(frmScreen, "cboTypeOfProperty") == "F")
	{	
		if(XMLFlat.SelectTag(null, "FLAT") != null)
		{
			XML.SetAttribute("STOREYSINBLOCK", XMLFlat.GetAttribute("STOREYSINBLOCK"));
			XML.SetAttribute("FLOORINBLOCK", XMLFlat.GetAttribute("FLOORINBLOCK"));
			XML.SetAttribute("BUSINESSINBLOCK", XMLFlat.GetAttribute("BUSINESSINBLOCK"));
			//EP2_2
			XML.SetAttribute("NUMBEROFUNITS", XMLFlat.GetAttribute("NUMBEROFUNITS"));
			XML.SetAttribute("BUSINESSINBLOCKTYPEOFBUSINESS", XMLFlat.GetAttribute("BUSINESSINBLOCKTYPEOFBUSINESS"));
		}
	}
	else
	{
		XML.SetAttribute("STOREYSINBLOCK", "");
		XML.SetAttribute("FLOORINBLOCK", "");
		XML.SetAttribute("BUSINESSINBLOCK", "");
		//EP2_2
		XML.SetAttribute("NUMBEROFUNITS", "");
		XML.SetAttribute("BUSINESSINBLOCKTYPEOFBUSINESS", "");
		
	}	
}

function SaveValuation()
{
	var bSuccess = false;
	var ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTaskStatus;
	var sASPFile;
		
	sTaskStatus = m_TaskXML.GetAttribute("TASKSTATUS");
	
	if(m_sReportNotFound == "0") 
	{
		var reqTag = ValuationXML.CreateRequestTag(window , "UpdateValuationReport");
		ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");

		sASPFile = "omAppProc.asp"
		ValuationXML.CreateActiveTag("VALUATION");
		ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
		ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
		ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInsSeqNo);

		if(SaveScreenValues(ValuationXML))
		{
			bSuccess = true;
		}			
	}
	else if(m_sReportNotFound == "1")
	{
		var reqTag = ValuationXML.CreateRequestTag(window , "CreateValuationReport");
		ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");

		ValuationXML.CreateActiveTag("VALUATION");
		if(SaveScreenValues(ValuationXML))
		{
			sASPFile = "OmigaTmBO.asp"
			ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
			ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
			ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInsSeqNo);

			ValuationXML.ActiveTag = reqTag;
			ValuationXML.CreateActiveTag("CASETASK");
			ValuationXML.SetAttribute("ACTIVITYID", m_TaskXML.GetAttribute("ACTIVITYID"));
			ValuationXML.SetAttribute("STAGEID", m_TaskXML.GetAttribute("STAGEID"));
			ValuationXML.SetAttribute("CASESTAGESEQUENCENO", m_TaskXML.GetAttribute("CASESTAGESEQUENCENO"));
			ValuationXML.SetAttribute("CASEID", m_TaskXML.GetAttribute("CASEID"));
			ValuationXML.SetAttribute("TASKID", m_TaskXML.GetAttribute("TASKID"));
			ValuationXML.SetAttribute("TASKINSTANCE", m_TaskXML.GetAttribute("TASKINSTANCE"));
			
			bSuccess = true;
		}	
	}

	if(bSuccess)
	{
		// 		ValuationXML.RunASP(document, sASPFile);
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					ValuationXML.RunASP(document, sASPFile);
				break;
			default: // Error
				ValuationXML.SetErrorResponse();
			}

		if(ValuationXML.IsResponseOK())
		{
			// Update the context
			m_TaskXML.SetAttribute("TASKSTATUS",csTaskPending);
			scScreenFunctions.SetContextParameter(window,"idTaskXML", m_TaskXML.XMLDocument.xml);
			bSuccess = true;
		}
		else
		{
			bSuccess = false
		}
	}	

	return(bSuccess);
}

function HandleTenure()
{
	var sTenure;
	var bEnable = false;
	sTenure = frmScreen.cboTenure.value;
	
	if(sTenure == csTenureLeashold)
	{
		bEnable = true;
	}
	if(bEnable && m_sReadOnly != "1" )
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtUnexpiredLease")			
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtGroundRent")			
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtServiceCharge")			
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtUnexpiredLease")			
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtGroundRent")			
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtServiceCharge")			
	}
}

function frmScreen.cboTenure.onclick()
{
	HandleTenure();
}

function frmScreen.cboTenure.onchange()
{
	HandleTenure();
}

function GetOptionValue( objOption )
{
	var sVal = "0";
	
	if(objOption.checked == true)
	{
		sVal = "1";
	}
	return(sVal)
}

function SetOptionValue( objOptionYes, objOptionNo , sVal )
{
	if( sVal == "1" )
	{
		objOptionYes.checked = true;
	}
	else
	{
		objOptionNo.checked = true;
	}
}

function btnCancel.onclick()
{
	frmToAP205.submit();
}

function btnSubmit.onclick()
{
	if(CommitChanges())
	{
		frmToAP205.submit();
	}
}

/* PB 05/05/2006 EP429 */
function btnNext.onclick()
{
	if(CommitChanges())
	{
		frmToAP220.submit();
	}
}
/* EP429 End */

function frmScreen.btnPAFSearch.onclick()
{
	with (frmScreen)
		m_bPAFIndicator = PAFSearch(txtPostcode,txtHouseName,
			txtHouseNo,txtFlatNo,txtStreet,txtDistrict,
			txtTown,txtCounty,cboCountry);
		
}

function frmScreen.btnClearAddress.onclick()
{
<%	// APS 09/09/1999 - UNIT TEST REF 73
	// Add clear button to clear out possible wrong PAF Search
	// AY 15/09/99 - clearing of correspondence address fields now handled by a common function
	// Make sure that any changes of field content are flagged
%>	FlagChange(scScreenFunctions.ClearCollection(spnAddress));
	frmScreen.txtPostcode.focus();
}

function frmScreen.btnRooms.onclick()
//DPF 02/09/02 - this button calls pop window for rooms.
{
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_ValuationXML.XMLDocument.xml;
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = m_sAppNo;
	ArrayArguments[3] = m_sAppFactFindNo;
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "ap215.asp", ArrayArguments, 250, 420);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sRoomsXML = sReturn[1];
		SaveRoomDetails(m_ValuationXML);
	}
}

function frmScreen.cboTypeOfProperty.onchange()
{	

	<% /* MO - 19/11/2002 - BMIDS01006 */ %>
	<% /* if (scScreenFunctions.GetComboValidationType(frmScreen, "cboTypeOfProperty") == "F") */ %>
	if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfProperty, frmScreen.cboTypeOfProperty.selectedIndex, "F"))
	{
		frmScreen.btnFlatDetails.disabled = false;
	}
	else
	{
		frmScreen.btnFlatDetails.disabled = true;
		m_ValuationXML.SetAttribute("STOREYSINBLOCK", "");
		m_ValuationXML.SetAttribute("FLOORINBLOCK", "");
		m_ValuationXML.SetAttribute("BUSINESSINBLOCK", "");
	}
	
	<% //MAR1352 %>
	<% //Peter Edney - 06/03/2006 %>
	var sPropertyTypeIndex = frmScreen.cboTypeOfProperty.selectedIndex;	
	if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfProperty, sPropertyTypeIndex, "MPD"))
	{
		frmScreen.cboDescription.setAttribute("required","true");	    
		frmScreen.cboDescription.parentElement.parentElement.style.color = "red";	
	}
	else   
	{
		frmScreen.cboDescription.removeAttribute("required");	    	    
		frmScreen.cboDescription.parentElement.parentElement.style.color = "";	
	}	
	
	<%/*AW	EP436 03/05/2006 */%>
	populateDescriptionCombo();
}

function frmScreen.btnFlatDetails.onclick()
{
	ShowFlatDetailsPopUp();
}
function frmScreen.cboStructure.onchange()
{
	//JD MAR40 if the structure is non traditional enable the NonStandard text field
	//BC MAR1620 Include additional validation type
	if ((scScreenFunctions.GetComboValidationType(frmScreen, "cboStructure") == "NS") ||
	   (scScreenFunctions.GetComboValidationType(frmScreen, "cboStructure") == "NTC"))
		frmScreen.txtNonStandard.disabled = false;
	else
	{
		frmScreen.txtNonStandard.value = "";
		frmScreen.txtNonStandard.disabled = true;
	}
}
function ShowFlatDetailsPopUp()
//DPF 02/09/02 - this function calls pop window for flat details.
{ 
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_ValuationXML.XMLDocument.xml;
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = m_sAppNo;
	ArrayArguments[3] = m_sAppFactFindNo;
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "ap225.asp", ArrayArguments, 260, 230);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sFlatXML = sReturn[1];
		SaveFlatDetails(m_ValuationXML);
	}
}


function populateDescriptionCombo()
<%/*AW	EP436 03/05/2006 */%>
{
	//get current selected property validation type
	var strValidationType = frmScreen.cboTypeOfProperty.options[frmScreen.cboTypeOfProperty.selectedIndex].ValidationType0;
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("PropertyDescription");

	if(XML.GetComboLists(document,sGroupList))
	{
		XML.PopulateCombo(document,frmScreen.cboDescription,"PropertyDescription",true);
		 
		// remove non-matching entries
		var iCount = 0;
		for (iCount = frmScreen.cboDescription.length - 1; iCount > 0; iCount--)
		{
			if (scScreenFunctions.IsOptionValidationType(frmScreen.cboDescription, iCount, strValidationType) == false)
				frmScreen.cboDescription.remove(iCount);
		}
	}
}

//EP2_2 New methods
function frmScreen.ExLocalAuthority_Yes.onclick()
{
	ExLAPercent(true);
}	

function frmScreen.ExLocalAuthority_No.onclick()
{
	ExLAPercent(false);
}
	
function ExLAPercent(bFlag)
{
	// Enable if true.
	if (bFlag == true)
		frmScreen.txtPrivateOwnership.disabled = false;	
	else
	{
		frmScreen.txtPrivateOwnership.value = "0";	
		frmScreen.txtPrivateOwnership.disabled = true;	
	}
}	

function frmScreen.btnBuyToLet.onclick()
{
	ShowBTLDetailsPopUp();
}

function ShowBTLDetailsPopUp()
// This function calls pop window for Buy to Let details.
{ 
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_ValuationXML.XMLDocument.xml;
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = m_sAppNo;
	ArrayArguments[3] = m_sAppFactFindNo;
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "ap211.asp", ArrayArguments, 400, 400);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sBTLXML = sReturn[1];
		SaveBTLDetails(m_ValuationXML);
	}
}
function SaveBTLDetails(XML)
{
// Save Buy to Let details record on AP211.asp
	XMLBTL.LoadXML(m_sBTLXML);

	if(XMLBTL.SelectTag(null, "BTL") != null)
	{
		XML.SetAttribute("RENTALDEMAND", XMLBTL.GetAttribute("RENTALDEMAND"));
		XML.SetAttribute("RENTALINCOME", XMLBTL.GetAttribute("RENTALINCOME"));
		XML.SetAttribute("RENTABLECONDITION", XMLBTL.GetAttribute("RENTABLECONDITION"));
	}
}


//EP2_2 - END New methods


-->
</script>
</STRONG></STRONG></STRONG>
</body>
</html>


