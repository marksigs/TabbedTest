<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      Ap020.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Application Summary - Application
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		26/02/01	Created
CL		05/03/01	SYS1920 Read only functionality added
JLD		07/03/01	SYS1879 add routing to AP011
PSC		16/03/01	SYS2090 Populate EARNEDINCOMEAMOUNT based on
                            Applicant selected 
JR		02/07/01	Omiplus24 Include CountryCode and AreaCode fields
PSC		12/12/01	SYS3318 Amend to return Net Annual Income
LD		23/05/02	SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
ASu		04/09/02	BMIDS00394 AnnualIncome not populated correctly where 
					income is zero & change monthly income to Annual.
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
MO		26/11/2002	BMIDS01089 - Change hard code 910, 920 for cancel and decline stages to 
								  global parameters.
INR		01/06/2004	BMIDS744	ThirdPartyData, Show OptOutIndicator
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:

PJO     15/11/2005  MAR555 Avoid setting age to -1 if DOB is null
HMA     20/02/2006  MAR1040 Add NetAllowableAnnualIncome and NetConfirmedAllowableIncome
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
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 175px; LEFT: 300px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToAP010" method="post" action="AP010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP030" method="post" action="AP030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP011" method="post" action="AP011.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divApplicants" style="TOP: 60px; LEFT: 10px; HEIGHT: 235px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
	Applicants (select to view details):
</span>
<span id="spnTable" style="TOP: 25px; LEFT: 4px; POSITION: ABSOLUTE">
	<table id="tblTable" width="596px" border="0" cellspacing="0" cellpadding="0" class="msgTable">
	<tr id="rowTitles">
		<td width="30%" class="TableHead">Surname</td>
		<td width="30%" class="TableHead">Forenames</td>
		<td width="15%" class="TableHead">Title</td>
		<td width="13%" class="TableHead">Age</td>
		<td width="12%" class="TableHead">Role</td>
	</tr>
	<tr id="row01">
		<td class="TableTopLeft">&nbsp;</td>
		<td class="TableTopCenter">&nbsp;</td>
		<td class="TableTopCenter">&nbsp;</td>
		<td class="TableTopCenter">&nbsp;</td>
		<td class="TableTopRight">&nbsp;</td>
	</tr>
	<tr id="row02">
		<td class="TableLeft">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row03">
		<td class="TableLeft">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row04">
		<td class="TableLeft">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row05">
		<td class="TableBottomLeft">&nbsp;</td>
		<td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomRight">&nbsp;</td>
	</tr></table>
</span>
<span style="TOP:155px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Occupation
	<span style="TOP:-3px; LEFT:150px; POSITION:ABSOLUTE">
		<input id="txtOccupation" maxlength="50" style="WIDTH:140px" class="msgTxt">
	</span>
</span>
<span style="TOP:180px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Employment Status
	<span style="TOP:-3px; LEFT:150px; POSITION:ABSOLUTE">
		<input id="txtEmployStatus" maxlength="50" style="WIDTH:140px" class="msgTxt">
	</span>
</span>
<span style="TOP:205px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Date Started or Established
	<span style="TOP:-3px; LEFT:150px; POSITION:ABSOLUTE">
		<input id="txtDateStarted" maxlength="10" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:155px; LEFT:315px; POSITION:ABSOLUTE" class="msgLabel">
	Net Allowable Annual income
	<span style="TOP:0px; LEFT:155px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtAnnualIncome" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<span style="TOP:180px; LEFT:315px; POSITION:ABSOLUTE" class="msgLabel">
	Net Confirmed Annual income
	<span style="TOP:0px; LEFT:155px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtConfirmedAnnualIncome" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<span style="TOP:205px; LEFT:315px; POSITION:ABSOLUTE" class="msgLabel">
	Marital Status
	<span style="TOP:-3px; LEFT:155px; POSITION:ABSOLUTE">
		<input id="txtMaritalStatus" maxlength="50" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:230px; LEFT:415px; POSITION:ABSOLUTE">
	<input style="TOP:-3px; LEFT:50px; POSITION:absolute" id="chkMemberOfStaff" type="checkbox" value="1">
	<label style="TOP:0px; LEFT:-100px; POSITION:absolute" for="chkMemberOfStaff" class="msgLabel">Member of Staff</label>
</span>
</div>
<div id="divProperty" style="TOP: 305px; LEFT: 10px; HEIGHT: 310px; WIDTH: 297px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
	Mortgage Property Address
</span>
<span style="TOP:35px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Postcode
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtPostcode" maxlength="8" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:60px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Flat No.
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtFlatNo" maxlength="10" style="WIDTH:60px" class="msgTxt">
	</span>
</span>
<span style="TOP:85px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Building Name & No.
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtBuildingName" maxlength="40" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:82px; LEFT:230px; POSITION:ABSOLUTE">
	<input id="txtBuildingNo" maxlength="10" style="WIDTH:50px" class="msgTxt">
</span>
<span style="TOP:110px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Street
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtStreet" maxlength="40" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:135px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	District
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtDistrict" maxlength="40" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:160px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Town
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtTown" maxlength="40" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:185px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	County
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtCounty" maxlength="40" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:210px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Country
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtCountry" maxlength="50" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:235px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Country Code
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtCountryCode" maxlength="40" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:260px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Area Code
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtAreaCode" maxlength="40" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:285px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Telephone No.
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtTelephoneNo" maxlength="40" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
</div>
<div id="divAppDetails" style="TOP: 305px; LEFT: 317px; HEIGHT: 310px; WIDTH: 297px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
	Application Details
</span>
<span style="TOP:35px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Type of Application
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtTypeOfApp" maxlength="50" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:60px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Type of Buyer
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtTypeOfBuyer" maxlength="50" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:85px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Intermediary
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtIntermediary" maxlength="10" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:110px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Application No.
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtApplicationNo" maxlength="12" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:135px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Application Date
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtApplicationDate" maxlength="10" style="WIDTH:60px" class="msgTxt">
	</span>
</span>
<span style="TOP:160px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Application Status
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtApplicationStatus" maxlength="20" style="WIDTH:80px" class="msgTxt">
	</span>
</span>
<span style="TOP:156px; LEFT:210px; POSITION:ABSOLUTE">
	<input id="btnReview" value="Review" type="button" style="WIDTH:60px" class="msgButton">
</span>
<span style="TOP:185px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Application Priority
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<select id="cboApplicationPriority" style="WIDTH:100px" class="msgCombo"></select>
	</span>
</span>
<span style="TOP:210px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Application Stage
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtApplicationStage" maxlength="50" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
<span id="lblCreditCheckOptOutIndicator" name="lblCreditCheckOptOutIndicator" style="LEFT: 5px; POSITION: absolute; TOP: 235px" class="msgLabel">
	Opted out?
</span>
<span style="LEFT: 120px; POSITION: absolute; TOP: 235px" class="msgLabel">
	<input type="radio" value="1" id="idCreditCheckOptOutYes" name="CreditCheckOptOutIndicator" > Yes <input type="radio" value="0" id="idCreditCheckOptOutNo" name="CreditCheckOptOutIndicator"> No
</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 625px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP020attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_sReadOnly = "";
var m_sUnderReview = "";
var m_sStageName = "";
var m_sStageId = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_AppSummaryXML = null;
var m_sOrigAppPriority = "";
var m_iTableLength = 5;
var scScreenFunctions;
var m_blnReadOnly = false;
var m_sCancelStageId = "";
var m_sDeclineStageId = "";


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Next");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Application Summary - Application","AP020",scScreenFunctions);

	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
	SetMasks();
	Validation_Init();
	PopulateScreen();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP020");
	
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

function RetrieveContextData()
{
	/*JR TEST
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","00005010");	
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
	END*/

	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sUnderReview = scScreenFunctions.GetContextParameter(window,"idAppUnderReview",null);
	m_sStageName = scScreenFunctions.GetContextParameter(window,"idStageName",null);
	m_sStageId = scScreenFunctions.GetContextParameter(window,"idStageId",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	<% /* MO - 26/11/2002 - BMIDS01089 */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_sDeclineStageId = XML.GetGlobalParameterAmount(document, "DeclinedStageValue");
	m_sCancelStageId = XML.GetGlobalParameterAmount(document, "CancelledStageValue");
	XML = null;
}
function PopulateScreen()
{
	// Set fields readonly
	scScreenFunctions.SetCollectionState(divApplicants,"R");
	scScreenFunctions.SetCollectionState(divProperty,"R");
	scScreenFunctions.SetCollectionState(divAppDetails,"R");
	if(m_sReadOnly != "1" && m_sUnderReview != "1")
		scScreenFunctions.SetFieldState(frmScreen, "cboApplicationPriority", "W");
		
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();			
	var sGroups = new Array("ApplicationPriority");
	if (XML.GetComboLists(document, sGroups) == true)
		XML.PopulateCombo(document, frmScreen.cboApplicationPriority, "ApplicationPriority", false);
	
	if(GetApplicationSummary())
	{
		PopulateListBox();
		PopulateDetails();
		m_sOrigAppPriority = frmScreen.cboApplicationPriority.value;
		frmScreen.txtApplicationStage.value = m_sStageName;
		frmScreen.txtApplicationNo.value = m_sApplicationNumber;
		if(m_sUnderReview == "1")
			frmScreen.txtApplicationStatus.value = "Under Review";
		else
		{
			if(m_sStageId == m_sCancelStageId) frmScreen.txtApplicationStatus.value = "Cancelled";
			else if(m_sStageId == m_sDeclineStageId) frmScreen.txtApplicationStatus.value = "Declined";
			else
			{
				m_AppSummaryXML.SelectTag(null, "APPLICATION");
				if(m_AppSummaryXML.GetTagText("APPLICATIONRECOMMENDEDDATE") != "" &&
				   m_AppSummaryXML.GetTagText("APPLICATIONAPPROVALDATE") == ""       )
					frmScreen.txtApplicationStatus.value = "Recommended";
				else if(m_AppSummaryXML.GetTagText("APPLICATIONAPPROVALDATE") != "")
					frmScreen.txtApplicationStatus.value = "Approved";
				else
					frmScreen.txtApplicationStatus.value = "Active";
			}
		}
	}
}
function GetApplicationSummary()
{
	m_AppSummaryXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_AppSummaryXML.CreateRequestTag(window, "")
	m_AppSummaryXML.CreateActiveTag("APPLICATION");
	m_AppSummaryXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	m_AppSummaryXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	m_AppSummaryXML.RunASP(document,"GetApplicationSummaryData.asp");
	if(m_AppSummaryXML.IsResponseOK())
		return true;
	return false;
}
function PopulateListBox()
{
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("APPLICANT");
	var iNumberOfApplicants = m_AppSummaryXML.ActiveTagList.length;

	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfApplicants);
	ShowList(0);
	if(iNumberOfApplicants > 0) scScrollTable.setRowSelected(1);
}
function ShowList(nStart)
{
	var iCount;
	scScrollTable.clear();
	for (iCount = 0; iCount < m_AppSummaryXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		m_AppSummaryXML.SelectTagListItem(iCount + nStart);
		var sName = m_AppSummaryXML.GetTagText("FIRSTFORENAME");
		if(m_AppSummaryXML.GetTagText("SECONDFORENAME") != "")
			sName += " " + m_AppSummaryXML.GetTagText("SECONDFORENAME");
		if(m_AppSummaryXML.GetTagText("OTHERFORENAMES") != "")
			sName += " " + m_AppSummaryXML.GetTagText("OTHERFORENAMES");
		var sTitle = m_AppSummaryXML.GetTagText("TITLEOTHER");
		if(sTitle == "")
			sTitle = m_AppSummaryXML.GetTagAttribute("TITLE","TEXT");
		var sAge = CalculateCustomerAge(m_AppSummaryXML.GetTagText("DATEOFBIRTH"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),m_AppSummaryXML.GetTagText("SURNAME"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),sName );
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),sTitle );
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),sAge);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),m_AppSummaryXML.GetTagAttribute("CUSTOMERROLETYPE", "TEXT"));
		tblTable.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));
	}
}
function PopulateDetails()
{
	var dblAnnualIncome
	var dblConfirmedAnnualIncome
	
	m_AppSummaryXML.ActiveTag = null;
	m_AppSummaryXML.CreateTagList("APPLICANT");
	if(m_AppSummaryXML.ActiveTagList.length > 0)
	{
		var iRowSelected = scScrollTable.getRowSelected();
		var nApplicant = parseInt(tblTable.rows(iRowSelected).getAttribute("TagListItem"));
		m_AppSummaryXML.SelectTagListItem(nApplicant);
		frmScreen.txtOccupation.value = m_AppSummaryXML.GetTagAttribute("OCCUPATIONTYPE", "TEXT");
		frmScreen.txtEmployStatus.value = m_AppSummaryXML.GetTagAttribute("EMPLOYMENTSTATUS","TEXT");
		frmScreen.txtDateStarted.value = m_AppSummaryXML.GetTagText("DATESTARTEDORESTABLISHED");
		
		<% /* MAR1040 Populate fields from IncomeSummary table */ %>
		dblAnnualIncome = parseFloat(m_AppSummaryXML.GetTagText("NETALLOWABLEANNUALINCOME"))
		dblConfirmedAnnualIncome = parseFloat(m_AppSummaryXML.GetTagText("NETCONFIRMEDALLOWABLEINCOME"))
		
		if (! isNaN(dblAnnualIncome))
			frmScreen.txtAnnualIncome.value = dblAnnualIncome; 	
		else
			frmScreen.txtAnnualIncome.value = 0;

		if (! isNaN(dblConfirmedAnnualIncome))
			frmScreen.txtConfirmedAnnualIncome.value = dblConfirmedAnnualIncome; 	
		else
			frmScreen.txtConfirmedAnnualIncome.value = 0;
			
		frmScreen.txtMaritalStatus.value = m_AppSummaryXML.GetTagAttribute("MARITALSTATUS","TEXT");
		if(m_AppSummaryXML.GetTagText("MEMBEROFSTAFF") == "1")
			frmScreen.chkMemberOfStaff.checked = true;
		else frmScreen.chkMemberOfStaff.checked = false;
	}
	if(m_AppSummaryXML.SelectTag(null, "APPLICATION") != null)
	{
		frmScreen.txtPostcode.value = m_AppSummaryXML.GetTagText("POSTCODE");
		frmScreen.txtFlatNo.value = m_AppSummaryXML.GetTagText("FLATNUMBER");
		frmScreen.txtBuildingName.value = m_AppSummaryXML.GetTagText("BUILDINGORHOUSENAME");
		frmScreen.txtBuildingNo.value = m_AppSummaryXML.GetTagText("BUILDINGORHOUSENUMBER");
		frmScreen.txtStreet.value = m_AppSummaryXML.GetTagText("STREET");
		frmScreen.txtDistrict.value = m_AppSummaryXML.GetTagText("DISTRICT");
		frmScreen.txtTown.value = m_AppSummaryXML.GetTagText("TOWN");
		frmScreen.txtCounty.value = m_AppSummaryXML.GetTagText("COUNTY");
		frmScreen.txtCountry.value = m_AppSummaryXML.GetTagAttribute("COUNTRY","TEXT");
		frmScreen.txtCountryCode.value = m_AppSummaryXML.GetTagText("COUNTRYCODE");
		frmScreen.txtAreaCode.value = m_AppSummaryXML.GetTagText("AREACODE");
		frmScreen.txtTelephoneNo.value = m_AppSummaryXML.GetTagText("TELEPHONENUMBER");
		frmScreen.txtTypeOfApp.value = m_AppSummaryXML.GetTagAttribute("TYPEOFAPPLICATION","TEXT");
		frmScreen.txtTypeOfBuyer.value = m_AppSummaryXML.GetTagAttribute("TYPEOFBUYER","TEXT");
		frmScreen.txtIntermediary.value = m_AppSummaryXML.GetTagText("INTERMEDIARYPANELID");
		frmScreen.txtApplicationDate.value = m_AppSummaryXML.GetTagText("APPLICATIONDATE");
		frmScreen.cboApplicationPriority.value = m_AppSummaryXML.GetTagText("APPLICATIONPRIORITYVALUE");
		<% /* INR BMIDS744 */ %>	
		//scScreenFunctions.SetRadioGroupValue(frmScreen, "CreditCheckOptOutIndicator", m_AppSummaryXML.GetTagText("OPTOUTINDICATOR")) ;
		if(m_AppSummaryXML.GetTagText("OPTOUTINDICATOR") != "")
		{
			if(m_AppSummaryXML.GetTagText("OPTOUTINDICATOR") == "1")
				frmScreen.idCreditCheckOptOutYes.checked = true;
			else
				frmScreen.idCreditCheckOptOutNo.checked = true;
		}

	}	
	<% /*  BMIDS744 And don't let anyone change it */%>
	scScreenFunctions.SetRadioGroupState(frmScreen, "CreditCheckOptOutIndicator", "R");

}
function CalculateCustomerAge(sDateOfBirth)
{
	var nAge = -1;
	var dteBirthdate = scScreenFunctions.GetDateObjectFromString(sDateOfBirth);
	if(dteBirthdate != null)
	{
		<% /* MO - BMIDS00807 */ %>
		var dteToday = scScreenFunctions.GetAppServerDate(true);
		<% /* var dteToday = new Date(); */ %>
		nAge = top.frames[1].document.all.scMathFunctions.GetYearsBetweenDates(dteBirthdate, dteToday);
		<% /* PJO MAR555 15/11/2005 - Avoid setting age to -1 */ %>
		return(nAge.toString());		
	}
	else
	{
    	return("");
    }
}

function spnTable.onclick()
{
	PopulateDetails();
}
function OKProcessing()
{
	var bSuccess = true;
	if(frmScreen.onsubmit())
	{
		if(IsChanged())
		{
			if(frmScreen.cboApplicationPriority.value != m_sOrigAppPriority)
			{
				if(confirm("The application priority has been changed. Do you wish to continue?"))
				{
					var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					XML.CreateRequestTag(window, "")
					XML.CreateActiveTag("APPLICATIONPRIORITY");
					XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
					XML.CreateTag("APPLICATIONPRIORITYVALUE", frmScreen.cboApplicationPriority.value);
					// 					XML.RunASP(document,"CreateApplicationPriority.asp");
					// Added by automated update TW 09 Oct 2002 SYS5115
					switch (ScreenRules())
						{
						case 1: // Warning
						case 0: // OK
											XML.RunASP(document,"CreateApplicationPriority.asp");
							break;
						default: // Error
							XML.SetErrorResponse();
						}

					if(XML.IsResponseOK())
					{
						// change the idApplicationPriority context param
						scScreenFunctions.SetContextParameter(window,"idApplicationPriority",frmScreen.cboApplicationPriority.value);
						XML.ResetXMLDocument();
						var tagReq = XML.CreateRequestTag(window, "UpdateCaseTaskPriority")
						XML.CreateActiveTag("CASESTAGE");
						XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
						XML.SetAttribute("CASEID", m_sApplicationNumber);
						XML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",null));
						XML.ActiveTag = tagReq;
						XML.CreateActiveTag("PRIORITY");
						XML.SetAttribute("APPLICATIONPRIORITY", frmScreen.cboApplicationPriority.value);
						// 						XML.RunASP(document,"OmigaTMBO.asp");
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

						if(XML.IsResponseOK() == false)
							bSuccess = false;
					}
					else bSuccess = false;
					XML = null;
				}
				else bSuccess = false;
			}
		}
	}
	else bSuccess = false;
	
	return bSuccess;
}
function frmScreen.btnReview.onclick()
{
	frmToAP011.submit();
}
function btnSubmit.onclick()
{
	if(OKProcessing()) frmToAP010.submit();
}
function btnNext.onclick()
{
	if(OKProcessing()) frmToAP030.submit();
}
-->
</script>
</body>
</html>


