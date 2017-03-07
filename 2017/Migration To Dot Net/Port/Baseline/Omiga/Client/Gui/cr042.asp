<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      ???.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   ???
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
BG		27/12/00	Created Page - SYS1860
CL		05/03/01	SYS1920 Read only functionality added
JR		12/09/01	Omiplus24 include Country/Area Code for telephone number change
DRC     11/04/02    AQR 4369 Stop trying to unlock application when navigating away
JLD		21/05/02	SYS4664 select the correct component when populating the list.
LD		23/05/02	SYS4727 Use cached versions of frame functions


BMIDS Specific History:

Prog	Date		AQR			Description
MO		30/10/2002				Made change to fix Validation_Init call error, by adding validation.js to ASP
GHun	10/11/2002	BMIDS00304	Functions in the top menu were enabled when they should not be
DB		25/02/2003	BM0059		Screen loading incorrectly, scrolls to the bottom of the page
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
PSC		18/10/2005	MAR57		Change Routing
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM 2 Specific History:

Prog	Date		AQR			Description
PE		03/04/2007	EP2_2095	Contact details - Only display extension number if it is present.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
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
<% /* Scriptlets */ %>
<span style="TOP: 680px; LEFT: 310px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>

<% /* Specify Forms Here */ %>
<form id="frmToCR040" method="post" action="cr040.asp" STYLE="DISPLAY: none"></form>
<% /* PSC 18/10/2005 MAR57 */ %>
<form id="frmToCR041" method="post" action="cr041.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen">

<% /* Applicant Details */ %>
<div style="TOP: 60px; LEFT: 10px; HEIGHT: 300px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<% /* Current Address */ %>
	<span id="spnCorrespondenceAddress" style="TOP: 0px; LEFT: 10px; POSITION: ABSOLUTE">
		<span style="TOP: 15px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Applicant Name
				<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
					<input id="txtApplicantName" style="WIDTH: 200px; POSITION: ABSOLUTE">
				</span>	
		</span>

		<span style="TOP: 40px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Postcode
			<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
				<input id="txtPostcode" maxlength="8" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxtUpper">
			</span>
		</span>

		<span style="TOP: 64px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Flat No.
			<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
				<input id="txtFlatNumber" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 82px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			House<br>Name &amp; No.
			<span style="TOP: 3px; LEFT: 80px; POSITION: ABSOLUTE">
				<input id="txtHouseName" maxlength="40" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
			</span>

			<span style="TOP: 3px; LEFT: 234px; POSITION: ABSOLUTE">
				<input id="txtHouseNumber" maxlength="10" style="WIDTH: 45px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 112px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Street
			<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
				<input id="txtStreet" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 136px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			District
			<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
				<input id="txtDistrict" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 160px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Town
			<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
				<input id="txtTown" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 184px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			County
			<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
				<input id="txtCounty" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 208px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Country
			<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
				<input id="txtCountry" style="WIDTH: 200px" class="msgTxt">
			</span>
		</span>
		<span style="TOP: 244px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Customer Contact Details
			<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
				<input id="txtCustomerContactDetails" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		
		<span style="TOP: 244px; LEFT: 360px; POSITION: ABSOLUTE" class="msgLabel">
			Contact Type
			<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
				<input id="txtContactType" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		
		<span style="TOP: 270px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Applicants
			<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
				<select id="cboApplicants" style="WIDTH: 200px" class="msgCombo"></select>
			</span>
		</span>
	</span>
	
	<% /* Other info */ %>
	<span id="spnAdditionalInfo" style="TOP: 0px; LEFT: 320px; POSITION: ABSOLUTE">
		<span style="TOP: 15px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Legal Rep(s)
				<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
					<input id="txtLegalRep" style="WIDTH: 200px; POSITION: ABSOLUTE">
				</span>
		</span>
		
		<span style="TOP: 40px; LEFT: 0px; POSITION: ABSOLUTE">
				<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
					<input id="txtLegalRepCompany" style="WIDTH: 200px; POSITION: ABSOLUTE">
				</span>
		</span>
		
		<span style="TOP: 74px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			DX Number
			<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
				<input id="txtDXNumber" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 97px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Town
			<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
				<input id="txtLegalRepTown" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 120px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			Contact Number
			<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
				<input id="txtLegalRepContactNumber" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		
	</span>
	
</div>

<% /* Application Details */ %>
<div style="TOP: 365px; LEFT: 10px; HEIGHT: 65px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
		<span id="spnApplicationDetails" style="TOP: 0px; LEFT: 10px; POSITION: ABSOLUTE">
			<span style="TOP: 15px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
				Type of Application
					<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
						<input id="txtTypeofApplication" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
					</span>	
			</span>

			<span style="TOP: 40px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
				Business Type
				<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
					<input id="txtBusinessType" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
				</span>
			</span>

			<span style="TOP: 15px; LEFT: 310px; POSITION: ABSOLUTE" class="msgLabel">
				Business Source
				<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
					<input id="txtBusinessSource" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
				</span>
			</span>
		
			<span style="TOP: 40px; LEFT: 310px; POSITION: ABSOLUTE" class="msgLabel">
				Source Contact Details
				<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
					<input id="txtSourceContactDetails" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
				</span>
			</span>
		</span>
</div>

<% /* Business Details */ %>
<div style="TOP: 435px; LEFT: 10px; HEIGHT: 277px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
		
		<span id="spnQuoteDetails" style="TOP: -3px; LEFT: 10px; POSITION: ABSOLUTE">
			<span style="TOP: 15px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
				Home Insurance Quotation
					<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
						<input id="txtHomeInsuranceQuote" style="WIDTH: 150px; POSITION: ABSOLUTE">
					</span>	
			</span>

			<span style="TOP: 40px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
				Payment Protection Quotation
				<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
					<input id="txtPaymentProtectionQuote" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
				</span>
			</span>

			<span style="TOP: 15px; LEFT: 310px; POSITION: ABSOLUTE" class="msgLabel">
				Application Stage
				<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
					<input id="txtApplicationStage" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
				</span>
			</span>		
		</span>
		
		<span id="spnExistingBusiness" style="TOP: 70px; LEFT: 4px; POSITION: ABSOLUTE">
			<table id="tblExistingBusiness" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
				<tr id="rowTitles">	<td width="30%" class="TableHead" >Product Description</td>	<td width="15%" class="TableHead">Deadline</td>		<td width="10%" class="TableHead">Interest Rate</td>	<td width="20%" class="TableHead">Loan Value</td>			<td width="20%" class="TableHead">Term</td>					<td width="10%" class="TableHead">Repayment Type</td>			<td class="TableHead">Loan Instalment</td></tr>
				<tr id="row01">		<td width="30%" class="TableTopLeft">&nbsp</td>			<td width="15%" class="TableTopCenter">&nbsp</td>		<td width="10%" class="TableTopCenter">&nbsp</td>		<td width="20%" class="TableTopCenter">&nbsp</td>		<td width="20%" class="TableTopCenter">&nbsp</td>	<td width="10%" class="TableTopCenter">&nbsp</td>	<td class="TableTopRight">&nbsp</td></tr>
				<tr id="row02">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>				
				<tr id="row03">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row04">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row05">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row06">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row07">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row08">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row09">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>			<td width="20%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row10">		<td width="30%" class="TableBottomLeft">&nbsp</td>		<td width="15%" class="TableBottomCenter">&nbsp</td>	<td width="10%" class="TableBottomCenter">&nbsp</td>	<td width="20%" class="TableBottomCenter">&nbsp</td>	<td width="20%" class="TableBottomCenter">&nbsp</td>	<td width="10%" class="TableBottomCenter">&nbsp</td>	<td class="TableBottomRight">&nbsp</td></tr>
			</table>
		</span>
		
</div>

</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 720px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cr042Attribs.asp" -->


<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var scScreenFunctions;
var m_iTableLength = 10;
var m_blnReadOnly = false;
<% /* PSC 18/10/2005 MAR57 */%>
var	m_sCallingScreen = null;

	

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

<% /*Set context for testing purposes.
	
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
	scScreenFunctions.SetContextParameter(window,"idUserId", "");
	scScreenFunctions.SetContextParameter(window,"idRole", "");
	scScreenFunctions.SetContextParameter(window,"idUnitId", "");
	//scScreenFunctions.SetContextParameter(window,"idApplicationNumber", "C00041890");
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber", "C00066036");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber", "1");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber", "rf1116");
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber", "1");
*/%>

<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
	// AQR 4369 no lock set on application, so this prevents an unlock when you navigate away
	ClearApplicationContext();
	<% /* BMIDS00304 FW030SetTitles enables menu functions that are not appropriate for this
	screen if the application number is in context, so it must only be called after the 
	application context has been cleared */ %>
	FW030SetTitles("Application Summary","CR042",scScreenFunctions);
	
	PopulateScreen();
	PopulateApplicantsCombo();
	
			
	scScreenFunctions.SetCollectionState(spnCorrespondenceAddress, "R");
	frmScreen.cboApplicants.disabled = false;	
	scScreenFunctions.SetCollectionState(spnAdditionalInfo, "R");
	//DB BM0059 25/02/03 - This is causing focus to be set to the 'OK' button and therefore
	//the gui loads showing the bottom of the page, you have to scroll up to the top !		
	//scScreenFunctions.SetCollectionState(spnApplicationDetails, "R");	
	scScreenFunctions.SetCollectionState(spnQuoteDetails, "R");
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	scScreenFunctions.SetCollectionState(spnApplicationDetails, "R");
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CR042");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}
function ClearApplicationContext()
{
	scScreenFunctions.SetContextParameter(window, "idApplicationNumber", "");
	scScreenFunctions.SetContextParameter(window, "idApplicationFactFindNumber", "");
	scScreenFunctions.SetContextParameter(window, "idApplicationNumber", "");
}

function PopulateScreen()
{
	<%/* Go get the data*/%>
	AppSummaryXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AppSummaryXML.CreateRequestTag(window,null);
	AppSummaryXML.CreateActiveTag("APPLICATIONDETAILS");
	AppSummaryXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AppSummaryXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	AppSummaryXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	AppSummaryXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	
	AppSummaryXML.RunASP(document,"GetApplicationSummary.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = AppSummaryXML.CheckResponse(ErrorTypes);

	if (ErrorReturn[0] == true)PopulateListBox(0);
	else if(ErrorReturn[1] == ErrorTypes[0]) alert("No entries found for this application.");
	ErrorTypes = null;
	
	<%/* Populate screen fields*/%>
	AppSummaryXML.SelectTag(null, "CUSTOMERNAME")
	if (AppSummaryXML.ActiveTag != null)
	{ 
		frmScreen.txtApplicantName.value = AppSummaryXML.GetTagAttribute("TITLE", "TEXT") + " " + AppSummaryXML.GetTagText("FIRSTFORENAME") + " " + AppSummaryXML.GetTagText("SURNAME");
	}
	AppSummaryXML.SelectTag(null, "PROPERTYADDRESS");
	if (AppSummaryXML.ActiveTag != null)
	{
		frmScreen.txtPostcode.value = AppSummaryXML.GetTagText("POSTCODE");
		frmScreen.txtFlatNumber.value = AppSummaryXML.GetTagText("FLATNUMBER");
		frmScreen.txtHouseName.value = AppSummaryXML.GetTagText("BUILDINGORHOUSENAME");
		frmScreen.txtHouseNumber.value = AppSummaryXML.GetTagText("BUILDINGORHOUSENUMBER");
		frmScreen.txtStreet.value = AppSummaryXML.GetTagText("STREET");
		frmScreen.txtDistrict.value = AppSummaryXML.GetTagText("DISTRICT");
		frmScreen.txtTown.value = AppSummaryXML.GetTagText("TOWN");
		frmScreen.txtCounty.value = AppSummaryXML.GetTagText("COUNTY");
		frmScreen.txtCountry.value = AppSummaryXML.GetTagAttribute("COUNTRY", "TEXT");
	}
	AppSummaryXML.SelectTag(null, "CUSTOMERCONTACTDETAILS");
	if (AppSummaryXML.ActiveTag != null)
	{
		var sExt = "";
		
		if(AppSummaryXML.GetTagText("TELEPHONEEXTENSIONNUMBER") != "")
		{
			sExt = " Ext " + AppSummaryXML.GetTagText("TELEPHONEEXTENSIONNUMBER");
		}
		frmScreen.txtCustomerContactDetails.value = AppSummaryXML.GetTagText("COUNTRYCODE") + " " + AppSummaryXML.GetTagText("AREACODE") + " " + AppSummaryXML.GetTagText("TELEPHONENUMBER") + sExt;
		frmScreen.txtContactType.value = AppSummaryXML.GetTagAttribute("USAGE", "TEXT");
		
	}
	AppSummaryXML.SelectTag(null, "TYPEOFAPPLICATIONDETAILS");
	if (AppSummaryXML.ActiveTag != null)
	{
		frmScreen.txtTypeofApplication.value = AppSummaryXML.GetTagAttribute("TYPEOFAPPLICATION", "TEXT");
	}	
	AppSummaryXML.SelectTag(null, "BUSINESSTYPE");
	if (AppSummaryXML.ActiveTag != null)
	{
		frmScreen.txtBusinessType.value = AppSummaryXML.GetTagAttribute("DIRECTINDIRECTBUSINESS", "TEXT");
	}	
	AppSummaryXML.SelectTag(null, "BUSINESSSOURCE");
	if (AppSummaryXML.ActiveTag != null)
	{
		var sSource = AppSummaryXML.GetTagText("NAME");
		
		if(sSource != "")
		{
			frmScreen.txtBusinessSource.value = sSource;
		}
		else
		{
			frmScreen.txtBusinessSource.value = AppSummaryXML.GetTagText("UNITNAME");
		}
	}
	//JR - Omiplus24, get work(W) contact telephone details 
	var sComboId = AppSummaryXML.GetComboIdForValidation("ContactTelephoneUsage", "W", null, document);
	AppSummaryXML.SelectTag(null, "BUSINESSSOURCECONTACTDETAILS/CONTACTTELEPHONEDETAILS[USAGE='" + sComboId + "']");
	//JR End
	if (AppSummaryXML.ActiveTag != null)
	{
		var sPhoneNo = AppSummaryXML.GetTagText("COUNTRYCODE") + " " + AppSummaryXML.GetTagText("AREACODE") + " " + AppSummaryXML.GetTagText("TELENUMBER");
		if (AppSummaryXML.GetTagText("EXTENSIONNUMBER") != "")
			sPhoneNo = sPhoneNo + " Ext " + AppSummaryXML.GetTagText("EXTENSIONNUMBER");
		frmScreen.txtSourceContactDetails.value = sPhoneNo;
	}	
	AppSummaryXML.SelectTag(null, "LEGALREPS");
	if (AppSummaryXML.ActiveTag != null)
	{
		frmScreen.txtLegalRep.value = AppSummaryXML.GetTagText("CONTACTFORENAME") + " " + AppSummaryXML.GetTagText("CONTACTSURNAME");
		frmScreen.txtLegalRepCompany.value = AppSummaryXML.GetTagText("COMPANYNAME");
		frmScreen.txtDXNumber.value = AppSummaryXML.GetTagText("DXID");
		frmScreen.txtLegalRepTown.value = AppSummaryXML.GetTagText("DXLOCATION");
		frmScreen.txtLegalRepCompany.value = AppSummaryXML.GetTagText("COMPANYNAME");
		//JR - Omiplus, Get Work(W) telephonedetails from within LegalReps node
		AppSummaryXML.SelectTag(null, "LEGALREPS/CONTACTTELEPHONEDETAILS[USAGE='" + sComboId + "']");		
		if (AppSummaryXML.ActiveTag != null)
		{
			var sPhoneNo = AppSummaryXML.GetTagText("COUNTRYCODE") + " " + AppSummaryXML.GetTagText("AREACODE") + " " + AppSummaryXML.GetTagText("TELENUMBER");
			if (AppSummaryXML.GetTagText("EXTENSIONNUMBER") != "")
				sPhoneNo = sPhoneNo + " Ext " + AppSummaryXML.GetTagText("EXTENSIONNUMBER");
			frmScreen.txtLegalRepContactNumber.value = sPhoneNo;
			
		}
	}	
	AppSummaryXML.SelectTag(null, "HOMEINSQUOTE");
	if (AppSummaryXML.ActiveTag != null)
	{
		frmScreen.txtHomeInsuranceQuote.value = AppSummaryXML.GetTagText("TOTALBCMONTHLYCOST");
	}	
	AppSummaryXML.SelectTag(null, "PAYMENTPROTECTIONQUOTE");
	if (AppSummaryXML.ActiveTag != null)
	{
		frmScreen.txtPaymentProtectionQuote.value = AppSummaryXML.GetTagText("TOTALPPMONTHLYCOST");
	}	
	
	frmScreen.txtApplicationStage.value = m_sStageName 
}


function PopulateApplicantsCombo()
{
	AppSummaryXML.CreateTagList("APPLICANTDETAILS");
	for(var iCount = 0; iCount < AppSummaryXML.ActiveTagList.length; iCount++)
	{
		AppSummaryXML.SelectTagListItem(iCount);
		TagOPTION = document.createElement("OPTION");
		TagOPTION.text = AppSummaryXML.GetTagAttribute("TITLE", "TEXT") + " " + AppSummaryXML.GetTagText("FIRSTFORENAME") + " " +  AppSummaryXML.GetTagText("SECONDFORENAME") +  " " + AppSummaryXML.GetTagText("SURNAME");
		frmScreen.cboApplicants.add(TagOPTION);
	}
}

function PopulateListBox()
{
	<%  /* Populate the listbox with values from AppSummaryXML */%>	
	AppSummaryXML.SelectTag(null,"PRODUCTANDLOANDETAILSLIST");
	
	AppSummaryXML.CreateTagList("PRODUCTANDLOANDETAILS");
	var iNumberOfUsers = AppSummaryXML.ActiveTagList.length;
	
	scScrollTable.initialiseTable(tblExistingBusiness, 0, "", ShowList, m_iTableLength, iNumberOfUsers);
	ShowList(0);
}

function ShowList(nStart)
{
	var iCount;
	scScrollTable.clear();
	for (iCount = 0; iCount < AppSummaryXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		AppSummaryXML.SelectTagListItem(iCount + nStart);  //JLD SYS4664
		var sTerm = AppSummaryXML.GetTagText("TERMINYEARS") + " yrs " + AppSummaryXML.GetTagText("TERMINMONTHS") + " mths";
		
		var sDescription = ""
		var sTag = AppSummaryXML.GetTagText("PORTED");
		if (sTag == "1")
			{
			sDescription = AppSummaryXML.GetTagText("MORTGAGEPRODUCTDESCRIPTION") + "(Ported)";
				if(sDescription == "(Ported)")
					 sDescription = "No Description - " + sDescription
			}
		else
			{	
			sDescription = AppSummaryXML.GetTagText("PRODUCTNAME");
			if(sDescription == "")
			  	sDescription = "No Description available"
			}
		scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(iCount+1).cells(0),sDescription);
		scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(iCount+1).cells(1),AppSummaryXML.GetTagText("WITHDRAWNDATE"));
		scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(iCount+1).cells(2),AppSummaryXML.GetTagText("INTERESTRATE"));
		scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(iCount+1).cells(3),AppSummaryXML.GetTagText("LOANAMOUNT"));
		scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(iCount+1).cells(4),sTerm);
		scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(iCount+1).cells(5),AppSummaryXML.GetTagAttribute("REPAYMENTMETHOD", "TEXT"));
		scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(iCount+1).cells(6),AppSummaryXML.GetTagText("GROSSMONTHLYCOST"));
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

function RetrieveContextData()
{	
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFFNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber",null);
	m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber",null);
	m_sStageName = scScreenFunctions.GetContextParameter(window,"idStageName",null);
	<% /* PSC 18/10/2005 MAR57 */%>
	m_sCallingScreen = scScreenFunctions.GetContextParameter(window,"idCallingScreenID",null);
}


function btnSubmit.onclick()
{	
    ClearApplicationContext();
    
	<% /* PSC 18/10/2005 MAR57 - Start */ %>
    if (m_sCallingScreen == "CR041")
		frmToCR041.submit();
	else
		frmToCR040.submit();
	<% /* PSC 18/10/2005 MAR57 - End */ %>

}

-->
</script>
</body>
</html>




