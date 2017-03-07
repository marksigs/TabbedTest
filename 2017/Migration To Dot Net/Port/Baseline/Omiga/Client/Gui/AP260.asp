<%@ LANGUAGE="JSCRIPT" %>
<%
/*
Workfile:      AP260.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Valuation Instructions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
		16/01/2001	CREATED By  GD  SYS2039
JR		18/09/01	Omiplus24 - included Country/Area Code and renamed
					function PopulateDirectoryAddress to PopulateDirectoryAddressDetails and 
					to btnDirectorySearch to btnValuerDirectorySearch toavoid conflict with 
					ThirdPartyDetails function.
MDC		01/10/01	SYS2785 Enable client versions to override labels
PSC		11/12/01	SYS3367 Add button to clear search criteria
sys2100	19/03/2001	V2 and V3 changed valuation type from 10 to 20
DRC     30/01/02    SYS3203 & SYS2092 Ensure that Valuation status does not regress
DRC     30/01/02    SYS3203 & SYS2662 Apointment Date can't be earlier than Date of Instruction
DRC     04/02/02    SYS3203 & SYS3070 Pass Current StageID from context to ap265.asp popup
DRC     13/02/02    SYS4058 change valuationstatus to be passed as a string
SA		22/04/02	SYS4439 County field set to read only.
AT		30/04/02	SYS4359 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
GD		16/01/2001	BMIDS00154	Remove PAF,Clear and Add To Directory
MV		05/07/2002	BMIDS00103  Amended window.onload()and GetComboLists
MV		07/08/2002	BMIDS00302	Remove non-style sheet styles
DPF		13/09/2002	APWP3/BM007 Added new field - 'Town' into search criteria
SA		30/10/2002	BMIDS00657	Clear down all address fields too when clear button is pressed - plus telephone no's
SA		31/10/2002	BMIDS00658  Pass in ValuerType as 10th param to ZA020
SA		08/11/2002	BMIDS00657  Disable Clear Search button as well if readonly or assigned
PSC		03/12/2002	BM0105      Amend to always go to ZA020 on search and rewrite PopulateDirectoryAddressDetails
                                to access new structure
BS		17/02/2003	BM0187		Remove valuation types with validation type 'N' from combo
BS		03/03/2003	BM0187		Forgot to declare variable
KRW     25/09/2003  BM0063      Made change to make screen read only when locked by other user
HMA     03/10/2003  BMIDS640    Do not allow multiple presses of OK.
mc		19/04/2004	BMIDS517	PropertyValuation dialog(AP265) height incr. by 22px
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
PB		21/07/2006	EP543		Added 'TITLEOTHER' field for third parties
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/	
%>
<html>
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
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<OBJECT data=scTable.htm height=1 id=scTable style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet VIEWASTEXT></OBJECT>

<% /* Specify Forms Here */ %>
<form id="frmGoToPreviousScreen" method="post" action="AP250.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" year4 validate="onchange" mark>

<div id="divBackground" style="HEIGHT: 105px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">

	<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel">
		<LABEL id="idValuationType"></LABEL>
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
			<select id="cboValuationType" style="POSITION: absolute; WIDTH: 150px" class="msgCombo">

			</select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 48px" class="msgLabel">
		Company
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyName" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 320px; POSITION: absolute; TOP: 24px" class="msgLabel">
		<LABEL id="idValuerType"></LABEL>
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
			<select id="cboValuerType" style="POSITION: absolute; WIDTH: 120px" class="msgCombo">

			</select>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 72px" class="msgLabel">
		Town
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
			<input id="txtSearchTown" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 320px; POSITION: absolute; TOP: 48px" class="msgLabel">
		Panel No.
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
			<input id="txtPanelNo"  maxlength="12" style="POSITION: absolute; WIDTH: 80px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 390px; POSITION: absolute; TOP: 72px">
		<input id="btnClearSearch" value="Clear" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	
	<span style="LEFT: 495px; POSITION: absolute; TOP: 72px">
		<input id="btnValuerDirectorySearch" value="Directory Search" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	
	<!-- This button is specifically used for ThirdPartyDetails, not currently active in this form-->
	<span style="TOP: 92px; LEFT: 495px; POSITION: ABSOLUTE">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px;  visibility:hidden" class="msgButton">
	</span>
	

</div>

<div style="HEIGHT: 130px; LEFT: 10px; POSITION: absolute; TOP: 170px; WIDTH: 604px" class="msgGroup">


	<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel">
		Appointment Date
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
			<input id="txtAppointmentDate"  maxlength="10" style="POSITION: absolute; WIDTH: 120px" class="msgTxt">
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 48px" class="msgLabel">
		Key Taken?
		<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
			<input id="optKeyTaken_Yes" name="optKeyTaken" type="radio"><label for="optRadio1" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="optKeyTaken_No" name="optKeyTaken" type="radio" checked><label for="optRadio2" class="msgLabel">No</label>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 72px" class="msgLabel">
		<LABEL id="idValuationStatus"></LABEL>
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
			<input id="txtValuationStatus"  maxlength="40" style="POSITION: absolute; WIDTH: 120px" class="msgTxt">
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 96px" class="msgLabel">
	<!--	hiddenDirectoryGUID-->
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
			<input id="txtDirectoryGUID"  type = hidden style="POSITION: absolute; WIDTH: 120px" class="msgTxt" tabindex = -1>
		</span>
	</span>
	<!--
	<span style="LEFT: 4px; POSITION: absolute; TOP: 120px" class="msgLabel">
		txtChanged
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
			<input id="txtChanged" style="POSITION: absolute; WIDTH: 120px" class="msgTxt" tabindex = -1>
		</span>
	</span>
	-->
	

	<span style="LEFT: 320px; POSITION: absolute; TOP: 24px" class="msgLabel">
		Date Of Instruction
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
			<input id="txtDateOfInstruction"  maxlength="10" style="POSITION: absolute; WIDTH: 120px" class="msgTxt">
		</span>
	</span>
<!--GD Removed until decision made about this field-->
<!--
	<span style="LEFT: 320px; POSITION: absolute; TOP: 48px" class="msgLabel">
		Valuation Fee
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
			<LABEL style="POSITION: absolute; LEFT: -6px; TOP: 2px " class="msgCurrency"></LABEL>-->
			<input id="txtValuationFee"  type=hidden maxlength="8" style="POSITION: absolute; WIDTH: 120px" class="msgTxt">
<!--
		</span>
	</span>
-->	
<!--Use these positionings if above text box is replaced
	<span style="LEFT: 320px; POSITION: absolute; TOP: 72px" class="msgLabel">
		Additional Notes
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px"><TEXTAREA class=msgTxt id=txtAdditionalNotes rows=3 style="POSITION: absolute; WIDTH: 150px"></TEXTAREA>
		</span>
	</span>
-->
<!--replace the span below with the span above if valuation fee put back in-->
	<span style="LEFT: 320px; POSITION: absolute; TOP: 48px" class="msgLabel">
		Additional Notes
		<span style="LEFT: 95px; POSITION: absolute; TOP: -3px"><TEXTAREA class=msgTxt id=txtAdditionalNotes rows=3 style="POSITION: absolute; WIDTH: 150px"></TEXTAREA>
		</span>
	</span>
	<!--GD SYS2047 - enable Property Details Button-->
	<!--JR Omiplus24 - re-positioned it here-->
	<span style="LEFT: 495px; POSITION: absolute; TOP: 102px">
		<input id="btnPropertyDetails" value="Property Details"  type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	
</div>

<% /* Address Details and Contact Details*/ %>
<div style="HEIGHT: 230px; LEFT: 10px; POSITION: absolute; TOP: 305px; WIDTH: 604px" class="msgGroup">
	<!-- #include FILE="includes/thirdpartydetails.htm" -->
	
	
	<!--<% /* Address */ %>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 0px" class="msgLabel">
			<strong>Address</strong>
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
			Postcode
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtPostcode" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
			</span>
		</span>


		<span style="LEFT: 4px; POSITION: absolute; TOP: 58px" class="msgLabel">
			Building<br>Name &amp; No.
			<span style="LEFT: 95px; POSITION: absolute; TOP: 3px">
				<input id="txtBuildingName" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
			</span>

			<span style="LEFT: 249px; POSITION: absolute; TOP: 3px">
				<input id="txtBuildingNumber" maxlength="10" style="POSITION: absolute; WIDTH: 45px" class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
			</span>
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 88px" class="msgLabel">
			Street
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtStreet" maxlength="40" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
			</span>
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 112px" class="msgLabel">
			District
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtDistrict" maxlength="40" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
			</span>
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 136px" class="msgLabel">
			Town
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtTown" maxlength="40" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
			</span>
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 160px" class="msgLabel">
			County
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtCounty" maxlength="40" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" onchange="ClearCurrentAddressPAFIndicator()">
			</span>
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 184px" class="msgLabel">
			Country
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtCountry"  maxlength="40" style="WIDTH: 200px" class="msgTxt" >
			</span>
		</span>
<!-- +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

	<!-- <% /* Contact Details */ %>
		<span style="LEFT: 320px; POSITION: absolute; TOP: 0px" class="msgLabel">
			<strong>Contact Details</strong>
		</span>

		<span style="LEFT: 320px; POSITION: absolute; TOP: 40px" class="msgLabel">
			Title
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtTitle" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgTxt" >
			</span>
		</span>

		<span style="LEFT: 320px; POSITION: absolute; TOP: 64px" class="msgLabel">
			Forename
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtForename" maxlength="40" style="POSITION: absolute; WIDTH: 70px" class="msgTxt" >
			</span>
		</span>

		<span style="LEFT: 320px; POSITION: absolute; TOP: 88px" class="msgLabel">
			Surname
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtSurname" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgTxt" onchange="ClearCorrespondenceAddressPAFIndicator()">
			</span>
		</span>
		
		<span style="LEFT: 320px; POSITION: absolute; TOP: 112px" class="msgLabel">
			Telephone No.
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtTelephoneNumber" maxlength="30" style="POSITION: absolute; WIDTH: 150px" class="msgTxt" onchange="ClearCorrespondenceAddressPAFIndicator()">
			</span>
		</span>

		<span style="LEFT: 320px; POSITION: absolute; TOP: 136px" class="msgLabel">
			Fax No.
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtFaxNo" maxlength="30" style="POSITION: absolute; WIDTH: 150px" class="msgTxt" onchange="ClearCorrespondenceAddressPAFIndicator()">
			</span>
		</span>

		<span style="LEFT: 320px; POSITION: absolute; TOP: 160px" class="msgLabel">
			E-mail Address:
		</span>	
		
		<span style="LEFT: 320px; POSITION: absolute; TOP: 182px">
			<input id="txtEmail" maxlength="40" style="POSITION: absolute; WIDTH: 248px" class="msgTxt" onchange="ClearCorrespondenceAddressPAFIndicator()">
		</span>	
	<!--</span>-->
	<!--GD SYS2047 - enable Property Details Button
	<span style="LEFT: 495px; POSITION: absolute; TOP: 220px">
		<input id="btnPropertyDetails" value="Property Details"  type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	-->

</div>

</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 560px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>


<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP260attribs.asp" -->
<!-- #include FILE="Customise/AP260Customise.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction				= null;
var m_sInstructionSequenceNo	= null;
var m_sApplicationNumber		= null;
var	m_sApplicationFactFindNumber= null;
var m_sValuationStatus			= "";
var m_sValuerType				= null;
var m_sValuationType			= null;
//JR - Omiplus24
var m_sReadOnly					= null;
var m_ctrSortCode				= null;
var m_ctrOrganisationType		= null;
//End

var XML							= null;

var bValuationTypeChanged;
var scScreenFunctions;

//JR - Omiplus24
var m_sComboWorkId = "";
var m_sComboFaxId = "";


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>

	var sButtonList = new Array("Submit", "Cancel");

	ShowMainButtons(sButtonList);
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();	
	FW030SetTitles("Valuation Instruction","AP260",scScreenFunctions);
	scScreenFunctions.SetCurrency(window,frmScreen);	

	// MDC SYS2564 / SYS2785 Client Customisation
	var sGroups = new Array("ValuationType", "ValuerType", "Country","Title");
	objDerivedOperations = new DerivedScreen(sGroups);

	RetrieveContextData();
	Customise();
	// MC SYS2564/SYS2785 for client customisation
	ThirdPartyCustomise();
	var sControlList = new Array(frmScreen.cboValuationType, frmScreen.cboValuerType, frmScreen.cboCountry,frmScreen.cboTitle);
	objDerivedOperations.GetComboLists(sControlList);
			
	<% /* BS BM0187 17/02/03 
	Remove ValuationTypes with a ValidationType = 'N' from the combo */ %>
	var iCount = 0;
	for (iCount = frmScreen.cboValuationType.length - 1; iCount > 0; iCount--)
	{
		if (scScreenFunctions.IsOptionValidationType(frmScreen.cboValuationType, iCount, "N") )
			frmScreen.cboValuationType.remove(iCount);
	}		
	<% /* BS BM0187 End 17/02/03 */ %>
	// GetComboLists(); 	
	
	SetThirdPartyDetailsMasks(); //JR - Omiplus24
	Initialise();
	SetReadOnlyFields();
	
	if(m_sReadOnly = "1") // KRW 25/09/03 Added to make screen read only when locked by other user
	{
		var bReadOnly = false;
		bReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP205");	
		if(bReadOnly)
		{
			m_sReadOnly = "1";		
		}
	}
		
	SetMasks();	
	Validation_Init(); 
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	<%// GD BMIDS00154 START %>
	frmScreen.btnClear.disabled = true;
	frmScreen.btnPAFSearch.disabled = true;
	frmScreen.btnAddToDirectory.disabled = true;
	<%// GD BMIDS00154 END %>
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function SetReadOnlyFields()
{
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtValuationStatus");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtValuationFee");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtPostcode");
	//scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtBuildingName");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtHouseName");
	//scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtBuildingNumber");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtHouseNumber");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtFlatNumber");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtStreet");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDistrict");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTown");
	//scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtCounty");
	//SYS4439 County is a client configurable field
	objDerivedOperations.SetCountyToReadOnly();
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboCountry");
	//scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtCountry");
	//scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTitle");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboTitle");
	//scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtForename");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtContactForename");
	//scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtSurname");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtContactSurname");
	/*scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTelephoneNumber");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtFaxNo");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtEmail");	*/
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function Initialise()
{
	if (m_sInstructionSequenceNo.length > 0)
	{//sInstructionSequenceNo IS in context
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, "GetValuerInstructions");
		XML.CreateActiveTag("VALUERINSTRUCTIONS");
		XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInstructionSequenceNo);
		XML.SetAttribute("_COMBOLOOKUP_", "1");		
		//GetValuerInstructions
		
		XML.RunASP(document,"omAppProc.asp");		
		//CHECK RESPONSE before...
		
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = XML.CheckResponse(ErrorTypes);
		
		if(ErrorReturn[1] == ErrorTypes[0])
		{
			//alert("No Valuer instructions found");
			return;
		}
		else
		
		{
			PopulateScreen();
			if (m_sMetaAction =="READONLY")
			{
				scScreenFunctions.SetScreenToReadOnly(frmScreen);
				DisableMainButton("Submit");
				frmScreen.btnValuerDirectorySearch.disabled = true;
				//JR - Disable selective ThirdPartyDetails buttons
				frmScreen.btnClear.disabled=true;
				frmScreen.btnPAFSearch.disabled=true;
				frmScreen.btnAddToDirectory.disabled = true;
				//End
				//BMIDS00657 Disable Clearsearch button too
				frmScreen.btnClearSearch.disabled = true;
			} else
			{
		//SYS2078 19/03/2001
				if (m_sValuationStatus=="20") //Assigned
				{
					scScreenFunctions.SetScreenToReadOnly(frmScreen);
					scScreenFunctions.SetFieldToWritable(frmScreen,"txtAppointmentDate");
					scScreenFunctions.SetFieldToWritable(frmScreen,"optKeyTaken_No");
					scScreenFunctions.SetFieldToWritable(frmScreen,"optKeyTaken_Yes");
					scScreenFunctions.SetFieldToWritable(frmScreen,"optKeyTaken_Yes");
					scScreenFunctions.SetFieldToWritable(frmScreen,"txtAdditionalNotes");
					frmScreen.btnValuerDirectorySearch.disabled = true;	
					//JR - Disable selective ThirdPartyDetails buttons
					frmScreen.btnClear.disabled=true;
					frmScreen.btnPAFSearch.disabled=true;
					frmScreen.btnAddToDirectory.disabled = true;
					//End
					//BMIDS00657 Disable Clearsearch button too
					frmScreen.btnClearSearch.disabled = true;

				}
			}
		}
	}
	else
	{
		//SEE ADAM TO CLARIFY
		//display valuation type = RESPONSE.valuerInstructions.NewPropertyValuationTypeText
		
		var NewPropertyGeneralXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		NewPropertyGeneralXML.CreateRequestTag(window,null);
		NewPropertyGeneralXML.CreateActiveTag("NEWPROPERTY");
		NewPropertyGeneralXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		NewPropertyGeneralXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		NewPropertyGeneralXML.RunASP(document,"GetValuationTypeAndLocation.asp");
	
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = NewPropertyGeneralXML.CheckResponse(ErrorTypes);
		
		if(ErrorReturn[1] == ErrorTypes[0])
		{
			<%/* Record not found */%>
			//alert("No New Property details found");
			return;
		} else
		{
			var sValuationType
			NewPropertyGeneralXML.SelectTag(null,"NEWPROPERTY");
			sValuationType = NewPropertyGeneralXML.GetTagText("VALUATIONTYPE");
			if (sValuationType.length != 0)
			{
				<% /* BS BM0187 17/02/03 
				Valuation Types with ValidationType = 'N' have been removed from the combo
				so only initialise combo if sValuationType matches a value, else it will
				default to the first in the list */ %>
				var iCount=0;
				for (iCount=0; iCount < frmScreen.cboValuationType.length; iCount++) 
				{ 
					if (frmScreen.cboValuationType.options(iCount).value == sValuationType)
					{
						frmScreen.cboValuationType.value = sValuationType;
						break;
					}
				}
				<% /* BS BM0187 End 17/02/03 */ %>
			} else 
			{
				//alert("No valuation type found on the NewProperty table");
			}
		}
	}
}

function PopulateScreen()
{
	m_sComboWorkId = XML.GetComboIdForValidation("ContactTelephoneUsage", "W", null, document);
	m_sComboFaxId = XML.GetComboIdForValidation("ContactTelephoneUsage", "F", null, document);
	
	/* 02/10/2001 JR OmiPlus 24,
	Ensure the correct xml is passed to TP in the form: 
	<CONTACTDETAILS><CONTACTTELEPHONEDETAILS/></CONTACTDETAILS>*/ 
	var ThirdPartyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
	//Setup required xml to be passed to ThirdPartyDetails to display ContactDetails(DC241 - popup)
	ThirdPartyXML.CreateActiveTag("CONTACTDETAILS");
	var RootXML = ThirdPartyXML.ActiveTag;
	
	//Create EmailAddress on ContactDetails tag
	XML.SelectTag(null, "VALUERINSTRUCTIONS");
	ThirdPartyXML.CreateTag("EMAILADDRESS", XML.GetAttribute("EMAILADDRESS"));
	
	//Create CONTACTTELEPHONEDETAILS tag for work(W)
	XML.ActiveTag = null;
	XML.SelectTag(null,"VALUERINSTRUCTIONS[@USAGE='" + m_sComboWorkId + "']");
	if (XML.ActiveTag != null)
	{
		ThirdPartyXML.CreateActiveTag("CONTACTTELEPHONEDETAILS");		
		ThirdPartyXML.CreateTag("COUNTRYCODE", XML.GetAttribute("COUNTRYCODE"));			
		ThirdPartyXML.CreateTag("AREACODE", XML.GetAttribute("AREACODE"));		
		ThirdPartyXML.CreateTag("TELENUMBER", XML.GetAttribute("TELENUMBER"));
		ThirdPartyXML.CreateTag("EXTENSIONNUMBER", XML.GetAttribute("EXTENSIONNUMBER"));
		ThirdPartyXML.CreateTag("USAGE", XML.GetAttribute("USAGE"));				
		XML.ActiveTag = null;
	}
	//Create CONTACTTELEPHONEDETAILS tag for fax(F)
	XML.SelectTag(null,"VALUERINSTRUCTIONS[@USAGE='" + m_sComboFaxId + "']");
	if (XML.ActiveTag != null)
	{
		ThirdPartyXML.ActiveTag = RootXML;
		ThirdPartyXML.CreateActiveTag("CONTACTTELEPHONEDETAILS");		
		ThirdPartyXML.CreateTag("COUNTRYCODE", XML.GetAttribute("COUNTRYCODE"));			
		ThirdPartyXML.CreateTag("AREACODE", XML.GetAttribute("AREACODE"));		
		ThirdPartyXML.CreateTag("TELENUMBER", XML.GetAttribute("TELENUMBER"));
		ThirdPartyXML.CreateTag("EXTENSIONNUMBER", XML.GetAttribute("EXTENSIONNUMBER"));
		ThirdPartyXML.CreateTag("USAGE", XML.GetAttribute("USAGE"));
		XML.ActiveTag = null;
	}
	//Use the details for ThirdPartyDetails screen
	var ContactXML = ThirdPartyXML.SelectTag(null, "CONTACTDETAILS");	
	if(ContactXML != null)
		m_sXMLContact = ContactXML.xml;
	//JR - End
	
	XML.ActiveTag = null;
	XML.SelectTag(null,"VALUERINSTRUCTIONS");
	//Text Boxes
	frmScreen.txtAdditionalNotes.value	= XML.GetAttribute("VALUERNOTES"); 
	frmScreen.txtAppointmentDate.value	= XML.GetAttribute("APPOINTMENTDATE");
	frmScreen.txtHouseName.value		= XML.GetAttribute("BUILDINGORHOUSENAME");
	frmScreen.txtHouseNumber.value		= XML.GetAttribute("BUILDINGORHOUSENUMBER");
	frmScreen.txtCompanyName.value			= XML.GetAttribute("COMPANYNAME");
	frmScreen.cboCountry.value			= XML.GetAttribute("COUNTRY");	

	// MDC SYS2564/SYS2785 
	objDerivedOperations.LoadCounty(XML);
	//frmScreen.txtCounty.value			= XML.GetAttribute("COUNTY");

	frmScreen.txtDateOfInstruction.value= XML.GetAttribute("DATEOFINSTRUCTION");
	frmScreen.txtDirectoryGUID.value	= XML.GetAttribute("DIRECTORYGUID");
	frmScreen.txtDistrict.value			= XML.GetAttribute("DISTRICT");
	frmScreen.txtContactForename.value	= XML.GetAttribute("CONTACTFORENAME");
	frmScreen.txtPanelNo.value			= XML.GetAttribute("VALUERPANELNO");
	frmScreen.txtPostcode.value			= XML.GetAttribute("POSTCODE");
	frmScreen.txtStreet.value			= XML.GetAttribute("STREET");
	frmScreen.txtContactSurname.value	= XML.GetAttribute("CONTACTSURNAME");
	frmScreen.cboTitle.value			= XML.GetAttribute("CONTACTTITLE");
	<% /* PB 07/07/2006 EP543 Begin */ %>
	checkOtherTitleField();
	txtTitleOther.value = ArchitectXML.GetTagText("CONTACTTITLEOTHER");
	<% /* EP543 End */ %>
	frmScreen.txtTown.value				= XML.GetAttribute("TOWN");
	frmScreen.txtValuationFee.value		= XML.GetAttribute("VALUATIONFEE");
	frmScreen.txtValuationStatus.value  = XML.GetAttribute("VALUATIONSTATUS_TEXT");
	//frmScreen.txtCountry.value		= XML.GetAttribute("COUNTRY_TEXT");		
	//frmScreen.txtTelephoneNumber.value= sTelephone;
	//frmScreen.txtTitle.value			= XML.GetAttribute("CONTACTTITLE_TEXT");	
	//frmScreen.txtEmail.value			= XML.GetAttribute("EMAILADDRESS");
	//frmScreen.txtFaxNo.value			= sFaxNo;  
	//BMIDS00658/660 populate search town
	frmScreen.txtSearchTown.value				= XML.GetAttribute("TOWN");
	
	m_sValuationStatus = XML.GetAttribute("VALUATIONSTATUS");

	//Set Radios
	if (XML.GetAttribute("KEYTAKEN") == "1")
	{
		frmScreen.optKeyTaken_Yes.checked=true;
	} else
	{
		frmScreen.optKeyTaken_No.checked=true;
	}
	//Set Combos
	frmScreen.cboValuationType.value	= XML.GetAttribute("VALUATIONTYPE");	
	frmScreen.cboValuerType.value		= XML.GetAttribute("VALUERTYPE");	
}



function RetrieveContextData()
{
	/*Debug
	scScreenFunctions.SetContextParameter(window, "idApplicationNumber", "01160028");
	scScreenFunctions.SetContextParameter(window, "idApplicationFactFindNumber", "1");
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
	//End*/	
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sInstructionSequenceNo = scScreenFunctions.GetContextParameter(window, "idXML2", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
	//JR - Omiplus24, set to ReadOnly for use in ThirdPartyDetails.asp include file
	m_sReadOnly = "1";
}



function btnCancel.onclick()
{
	
	frmGoToPreviousScreen.submit();
}

<% /* BMIDS640 Disable the OK button and display the 'wait' cursor while processing is taking place */ %>
               
function btnSubmit.onclick()
{
	frmScreen.style.cursor = "wait";
	
	var bKeyTaken = false;
	if (frmScreen.onsubmit())
	{
		DisableMainButton("Submit");

		window.setTimeout("SubmitProcessing()", 0);
	}	
}

function SubmitProcessing()
{
		if (frmScreen.txtDateOfInstruction.value == "")
		{
			frmScreen.txtDateOfInstruction.focus;
			alert("Please enter the Instruction Date");
			return false;
		}
		if (frmScreen.optKeyTaken_Yes.checked==true)
		{
			bKeyTaken=1;
		}
		else
		{
			bKeyTaken=0;
		}

		if (frmScreen.txtDirectoryGUID.value == "")
		{
			alert("Please select a valuer");
			frmScreen.cboValuerType.focus;
			return false;
		}
	
		//Check that appointment date is valid
		// AQR SYS2662 & SYS3203 DRC 30/01/02 make sure that both fields are not blank
		if (!(frmScreen.txtAppointmentDate.value == "") && !(frmScreen.txtDateOfInstruction.value == ""))
		{
			if (scScreenFunctions.CompareDateFields(frmScreen.txtAppointmentDate,"<",frmScreen.txtDateOfInstruction))
			{
				alert("Appointment Date cannot be less than Date of Instruction");
				return;
			} 
		}
		//If we reach here, then all essential items have been entered
	
		if (m_sInstructionSequenceNo !="")
		{
			//Edit Mode
			
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
			XML.CreateRequestTag(window, "UpdateValuerInstructions");
			XML.CreateActiveTag("VALUERINSTRUCTION");
			XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
			XML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInstructionSequenceNo);
			XML.SetAttribute("DIRECTORYGUID", frmScreen.txtDirectoryGUID.value);
			XML.SetAttribute("VALUERTYPE", frmScreen.cboValuerType.value);
			XML.SetAttribute("VALUATIONTYPE", frmScreen.cboValuationType.value);
			XML.SetAttribute("APPOINTMENTDATE", frmScreen.txtAppointmentDate.value);
			XML.SetAttribute("KEYTAKEN", bKeyTaken);
			
			//GD &  
			//XML.SetAttribute("VALUATIONSTATUS", 20);
			// DC SYS2092 &  SYS3203 
		    // DC SYS4058 - changed integers to strings
			if ((m_sValuationStatus == "")||(parseInt(m_sValuationStatus) < 20))
			  XML.SetAttribute("VALUATIONSTATUS", "10");//New
			else
			  XML.SetAttribute("VALUATIONSTATUS", m_sValuationStatus);//Assigned
			  
			XML.SetAttribute("DATEOFINSTRUCTION", frmScreen.txtDateOfInstruction.value);
			XML.SetAttribute("VALUATIONFEE", frmScreen.txtValuationFee.value);
			XML.SetAttribute("VALUERNOTES", frmScreen.txtAdditionalNotes.value);
			XML.SetAttribute("VALUERPANELNO", frmScreen.txtPanelNo.value);
			
			// 			XML.RunASP(document,"omAppProc.asp");		
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document,"omAppProc.asp");		
					break;
				default: // Error
					XML.SetErrorResponse();
				}

			if(XML.IsResponseOK() != true)
			{
				alert("Error updating Valuation Instructions record.");
				return;
			}

		} else
		{
			
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			//CreateValuerInstructions
			XML.CreateRequestTag(window, "CreateValuerInstructions");
			XML.CreateActiveTag("VALUERINSTRUCTION");
			XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
			XML.SetAttribute("DIRECTORYGUID", frmScreen.txtDirectoryGUID.value);
			XML.SetAttribute("VALUERTYPE", frmScreen.cboValuerType.value );
			XML.SetAttribute("VALUATIONTYPE",frmScreen.cboValuationType.value );
			XML.SetAttribute("APPOINTMENTDATE", frmScreen.txtAppointmentDate.value);
			var bKeyTaken = frmScreen.optKeyTaken_Yes.checked;
			if (bKeyTaken == true)
			{
				XML.SetAttribute("KEYTAKEN", "1");
			} else
			{
				XML.SetAttribute("KEYTAKEN", "0");
			}
			//SYS2100 GD 19/03/2001
			//GD SYS2092
			//XML.SetAttribute("VALUATIONSTATUS", 20);
			// DC SYS2092 & SYS3203
			// DC SYS4058 - changed integers to strings
			if ((m_sValuationStatus == "")||(parseInt(m_sValuationStatus) < 20))
			  XML.SetAttribute("VALUATIONSTATUS", "10");//New
			else
			  XML.SetAttribute("VALUATIONSTATUS", m_sValuationStatus);//Assigned
			  
			XML.SetAttribute("DATEOFINSTRUCTION",frmScreen.txtDateOfInstruction.value );
			XML.SetAttribute("VALUATIONFEE", frmScreen.txtValuationFee.value);
			XML.SetAttribute("VALUERNOTES",frmScreen.txtAdditionalNotes.value);
			XML.SetAttribute("VALUERPANELNO", frmScreen.txtPanelNo.value);
			
			// 			XML.RunASP(document,"OmAppProc.asp");	
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document,"OmAppProc.asp");	
					break;
				default: // Error
					XML.SetErrorResponse();
				}

			if(XML.IsResponseOK() != true)
			{
				alert("Error creating Valuer Instructions record.");
				return;
			}

		} //End Create
		
		/*
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		//OmigaTMBO.CompleteValuerInstructions
		var ReqTag = XML.CreateRequestTag(window, "CompleteValuerInstructions");
		XML.CreateActiveTag("CASETASK");
		XML.SetAttribute("SOURCEAPPLICATION","Omiga" );
		XML.SetAttribute("CASEID", m_sApplicationNumber);		
		XML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",null));	
		XML.SetAttribute("STAGEID", scScreenFunctions.GetContextParameter(window,"idStageId",null));

		//Get idTaskXML from context, if there
		var sXMLCaseInfo = scScreenFunctions.GetContextParameter(window,"idTaskXML",null)
		if (sXMLCaseInfo !="")
		{
			var XMLcontext = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XMLcontext.LoadXML(sXMLCaseInfo);
			XMLcontext.SelectTag(null,"CASETASK");
			var sTaskInstance = XMLcontext.GetAttribute("TASKINSTANCE");
			var sActivityInstance = XMLcontext.GetAttribute("ACTIVITYINSTANCE");
			var sCaseStageSequenceNo = XMLcontext.GetAttribute("CASESTAGESEQUENCENO");	
			XML.SetAttribute("TASKINSTANCE", sTaskInstance );
			XML.SetAttribute("ACTIVITYINSTANCE", sActivityInstance);
			XML.SetAttribute("CASESTAGESEQUENCENO", sCaseStageSequenceNo);			
		}
		//End of Get idTaskXML from context	

		XML.ActiveTag=ReqTag;
		XML.CreateActiveTag("APPLICATION");
		XML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null) );
		XML.SetAttribute("APPLICATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null) );
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null) );
		XML.RunASP(document,"OmigaTmBO.asp");
		
		if (XML.IsResponseOK() == false)
		{
			alert("Error Completing Valuer Instructions");
			return;
		}
		*/
		frmGoToPreviousScreen.submit();	

}

function frmScreen.cboValuationType.onblur()
{
	//NOT IMNPLEMENTED
	//combo has been left (blurred) - now check to see if valuation fee needs to be populated
	
	if		(((frmScreen.cboValuationType.value !="") && (frmScreen.txtValuationFee.value =="") )
	||
			((bValuationTypeChanged==true) && (frmScreen.cboValuationType.value !="")) )
	{
		//GO GET VALUATION FEE
		//NOT IMPLEMENTED
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_sValuationType =	frmScreen.cboValuationType.value
		XML.CreateRequestTag(window, "GetValuationFee");
		
		XML.CreateActiveTag("VALUATION");
		XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.SetAttribute("VALUATIONTYPE", m_sValuationType);
		
		//XML.RunASP(document,"Valuation.asp");	
		//if XML.IsResponseOK()
		//{
		//	frmScreen.txtValuationFee.value=XML.GetAttribute("VALUATIONFEE");
		//}
		
	}

	//Reset changed flag...
	
	bValuationTypeChanged=false;
	//frmScreen.txtChanged.value=bValuationTypeChanged; GD DEBUG
}
function frmScreen.cboValuationType.onchange()
{
	//NOT IMPLEMENTED
	
	bValuationTypeChanged=true
	//frmScreen.txtChanged.value=bValuationTypeChanged; //GD DEBUG
	//if (frmScreen.cboValuationType.value !="")
	//{
	//	alert("GO GET VALUATION FEE - on change");
	//}	
}

function frmScreen.btnPropertyDetails.onclick()
{

	var sReturn = null;
	var ArrayArguments = new Array();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	
	ArrayArguments[0] = XML.CreateRequestAttributeArray(window);	
	ArrayArguments[1] = m_sApplicationNumber;
	ArrayArguments[2] = m_sApplicationFactFindNumber;
	ArrayArguments[3] = "1"//the read only flag - always pass this...

	
	//GS SYS2047
	ArrayArguments[4] = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue",null);	
	ArrayArguments[5] = scScreenFunctions.GetContextParameter(window,"idCurrency","");//the current currency
	//DRC SYS3070
    ArrayArguments[6] = scScreenFunctions.GetContextParameter(window,"idStageId","");//Current Stage
	sReturn = scScreenFunctions.DisplayPopup(window, document, "AP265.asp", ArrayArguments, 630, 550);
	
	
}

function frmScreen.btnValuerDirectorySearch.onclick()
{
	<% /* PSC 03/12/2002 BM0105 - Start */ %>
	if (frmScreen.cboValuationType.value=="")
	{
		alert("Valuation Type is required for a directory search.");
		return;
	}
	var XMLPanelValuerList 	= new top.frames[1].document.all.scXMLFunctions.XMLObject();

	if (frmScreen.txtCompanyName.value != "")
		if (!CheckField(frmScreen.txtCompanyName, "Company Name")) return;
		
	if (frmScreen.txtPanelNo.value != "")
		if (!CheckField(frmScreen.txtPanelNo, "Panel Number")) return;
		
	if (frmScreen.txtSearchTown.value != "")
		if (!CheckField(frmScreen.txtSearchTown, "Search Town")) return;
				
	var XMLreturned = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
	var ArrayArguments = XMLPanelValuerList.CreateRequestAttributeArray(window)

	/* Route to ZA020 */
	ArrayArguments[4] = 11;//a valuer
	ArrayArguments[5] = frmScreen.txtCompanyName.value;
	ArrayArguments[6] = true;
	ArrayArguments[7] = frmScreen.cboValuationType.value;
	ArrayArguments[8] = frmScreen.txtSearchTown.value;
	ArrayArguments[9] = frmScreen.txtPanelNo.value;
	//BMIDS00658 Pass in Valuer Type
	ArrayArguments[10] = frmScreen.cboValuerType.value;
	 							
	var sReturn = scScreenFunctions.DisplayPopup(window, document, "ZA020.asp", ArrayArguments, 630, 405);
	if (sReturn != null)
	{
					
		XMLreturned.LoadXML(sReturn[1]);
		PopulateDirectoryAddressDetails(XMLreturned);
		XMLreturned=null;
	}
	<% /* PSC 03/12/2002 BM0105 - End */ %>			
}

function CheckField(refField,sName)
{
	var sField=refField.value ;
	if(sField.indexOf("*") == 0)
	{
		alert("You must enter at least 1 preceding character to wildcard the " + sName);
		refField.focus();
		return false;
	}
	else 
		return true;

}

function PopulateDirectoryAddressDetails(XML)
{
	XML.ActiveTag = null;
	XML.SelectTag(null,"NAMEANDADDRESSDIRECTORY");
	
	<% /* PSC 03/12/2002 BM0105 - Start */ %>
	frmScreen.txtDirectoryGUID.value = XML.GetTagText("DIRECTORYGUID");
	frmScreen.txtCompanyName.value	 = XML.GetTagText("COMPANYNAME");
	
	XML.SelectTag(null,"ADDRESS");
	
	if (XML.ActiveTag != null)
	{
		frmScreen.txtPostcode.value	   = XML.GetTagText("POSTCODE");
		frmScreen.txtHouseName.value   = XML.GetTagText("BUILDINGORHOUSENAME");	
		frmScreen.txtHouseNumber.value = XML.GetTagText("BUILDINGORHOUSENUMBER");
		frmScreen.txtStreet.value	   = XML.GetTagText("STREET");
		frmScreen.txtDistrict.value	   = XML.GetTagText("DISTRICT");
		frmScreen.txtTown.value		   = XML.GetTagText("TOWN");	
		frmScreen.txtSearchTown.value  = XML.GetTagText("TOWN");
		frmScreen.cboCountry.value	   = XML.GetTagText("COUNTRY");
		objDerivedOperations.LoadCounty(XML);
	}

	XML.SelectTag(null,"CONTACTDETAILS");
	
	if (XML.ActiveTag != null)
	{
		frmScreen.cboTitle.value		   = XML.GetTagText("CONTACTTITLE");	
		frmScreen.txtContactForename.value = XML.GetTagText("CONTACTFORENAME");	
		frmScreen.txtContactSurname.value  = XML.GetTagText("CONTACTSURNAME");
		m_sXMLContact = XML.ActiveTag.xml
	}
	
	XML.SelectTag(null,"PANELVALUER");
	
	if (XML.ActiveTag != null)
	{
		frmScreen.cboValuerType.value = XML.GetTagText("VALUERTYPE");
	}
	<% /* PSC 03/12/2002 BM0105 - End */ %>
}


function GetComboLists()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("ValuationType", "ValuerType", "Country","Title");
	var bSuccess = false;

	if(XML.GetComboLists(document,sGroups))
	{
		bSuccess = true;
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboValuationType,"ValuationType",false);
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboValuerType,"ValuerType",true);
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboCountry,"Country",true);
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboTitle,"Title",true);
	}

	if(!bSuccess)
	{
	//	scScreenFunctions.SetScreenToReadOnly(frmScreen);
		DisableMainButton("Submit");
	}
	bValuationTypeChanged = false; //Initialise changed flag.
	//frmScreen.txtChanged.value=bValuationTypeChanged;//GD Debug
	
	//JR - Omiplus24, Call ThirdPartyDetails to populate contact title combo
	//PopulateTPTitleCombo();	
	
}

function frmScreen.btnClearSearch.onclick()
{
	frmScreen.txtCompanyName.value="";
	frmScreen.txtPanelNo.value="";
	//BMIDS00657 SA Clear all address fields too.
	frmScreen.txtSearchTown.value		= "";
	frmScreen.txtDirectoryGUID.value	= "";
	frmScreen.txtPostcode.value			= "";
	frmScreen.txtHouseName.value		= "";	
	frmScreen.txtHouseNumber.value		= "";
	frmScreen.txtStreet.value			= "";
	frmScreen.txtDistrict.value			= "";
	frmScreen.txtTown.value				= "";	
	objDerivedOperations.ClearCounty();
	frmScreen.cboCountry.value			= "";
	frmScreen.cboTitle.value			= "";	
	frmScreen.txtContactForename.value	= "";	
	frmScreen.txtContactSurname.value	= "";
	//BMIDS00657 - Don't forget the contact details!!
	m_sXMLContact = "";	
}		
-->
</script>
<!-- #include FILE="includes/thirdpartydetails.asp" -->
</body>
</html>





