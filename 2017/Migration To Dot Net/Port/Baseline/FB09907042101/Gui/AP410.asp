<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP410.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Applicants and Property Address
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Descrip
JR		31/1/01		Screen Design
JR		06/1/01		Added functionality
CL		12/03/01	SYS2034 Read only functionality added
JR		15/03/01	SYS1878 Change create report on title to call omigatmbo.asp
JR		19/03/01	SYS2089 Only retrieve ROT data if Task status not 10 (task already started)
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
JR		22/03/01	SYS2048 Allow middle tier to display error msg & not GUI
JR		10/04/01	SYS2251 Translate middle tier error msg using XML functions and implemented
JR		20/04/01	SYS2048 Added AddressType attribute when setting address attributes
JR		11/05/01	SYS2048	Added STAGEID and CASESTAGESEQUENCENO attributes to CreateReportonTitle request and
					set the ROTGUID attribute only on a UpdateReportonTitle. 
BG		26/11/01	SYS3107 need to change the TaskXML held in context to reflect the task 
					has been moved to pending from not started in ReportOnTitleAction function.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History : 

Prog	Date		AQR			Description
MV		07/08/2002	BMIDS0302	Core Ref : SYS4728 remove non-style sheet styles
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History : 

Prog	Date		AQR			Description
LDM		27/04/2006	MAR1624		Only create 1 report on title record
JD		18/05/2006	MAR1698		Check for create or update of address and ROTaddress.
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
<script src="validation.js" language="JScript"></script>

<% /* Specify Forms Here */ %>
<form id="frmToAP400" method="post" action="AP400.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP420" method="post" action="AP420.asp" STYLE="DISPLAY: none"></form>


<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<form id="frmScreen"  mark validate="onchange">

<div id="divBackground" style="HEIGHT: 365px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Full Names of Applicant(s):</strong>
	</span>
	<span id="spnApplicantList" style="LEFT: 10px; POSITION: absolute; TOP: 30px">
		<Table id="tblApplicantList" width="450" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="10%" class="TableHead">No.</td>
				<td width="40%" class="TableHead">Surname</td>	
				<td width="40%" class="TableHead">Forenames</td>	
				<td class="TableHead">Title</td>
			</tr>
			<tr id="row01">		
				<td width="10%" class="TableTopLeft">&nbsp</td>		
				<td width="40%" class="TableTopCenter">&nbsp</td>		
				<td width="40%" class="TableTopCenter">&nbsp</td>		
				<td class="TableTopRight">&nbsp</td>
			</tr>
			<tr id="row02">		
				<td width="10%" class="TableLeft">&nbsp</td>		
				<td width="40%" class="TableCenter">&nbsp</td>		
				<td width="40%" class="TableCenter">&nbsp</td>		
				<td class="TableRight">&nbsp</td>
			</tr>
			<tr id="row03">		
				<td width="10%" class="TableLeft">&nbsp</td>		
				<td width="40%" class="TableCenter">&nbsp</td>
				<td width="40%" class="TableCenter">&nbsp</td>
				<td class="TableRight">&nbsp</td>
			</tr>
			<tr id="row04">		
				<td width="10%" class="TableLeft">&nbsp</td>		
				<td width="40%" class="TableCenter">&nbsp</td>		
				<td width="40%" class="TableCenter">&nbsp</td>
				<td class="TableRight">&nbsp</td>
			</tr>			
			<tr id="row5">		
				<td width="10%" class="TableBottomLeft">&nbsp</td>		
				<td width="40%" class="TableBottomCenter">&nbsp</td>		
				<td width="40%" class="TableBottomCenter">&nbsp</td>
				<td class="TableBottomRight">&nbsp</td>
			</tr>
		</table>
	</span>
	<span style="TOP: 140px; LEFT:10px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Security Address</strong>
	</span>
	<span style="TOP: 140px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Completion Correspondence Address (if different)</strong>
	</span>
	<span style="TOP: 165px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Postcode
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtPostcode" maxlength="8" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt" readonly>
		</span>
	</span>
	<span style="TOP: 190px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Flat No.
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtFlatNumber" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt" readonly>
		</span>
	</span>
	<span style="TOP: 215px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		House Name
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtHouseName" maxlength="26" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt" readonly>
		</span>
		<span style="TOP: -3px; LEFT: 220px; POSITION: ABSOLUTE">
			<input id="txtHouseNumber" maxlength="10" style="WIDTH: 50px; POSITION: ABSOLUTE" class="msgTxt" readonly>
		</span>
	</span>
	<span style="TOP: 225px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		& No.
	</span>
	<span style="TOP: 240px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Street
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtStreet" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt" readonly>
		</span>
	</span>
	<span style="TOP: 265px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		District
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtDistrict" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt" readonly>
		</span>
	</span>
	<span style="TOP: 290px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Town
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtTown" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt" readonly>
		</span>
	</span>
	<span style="TOP: 315px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		County
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtCounty" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt" readonly>
		</span>
	</span>
	<span style="TOP: 340px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Country
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtCountry" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt" readonly>
		</span>
	</span>
	<span style="TOP: 165px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Postcode
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtPostcodeCorrespondence" maxlength="8" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 190px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Flat No.
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtFlatNumberCorrespondence" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 215px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		House Name
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtHouseNameCorrespondence" maxlength="26" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
		<span style="TOP: -3px; LEFT: 220px; POSITION: ABSOLUTE">
			<input id="txtHouseNumberCorrespondence" maxlength="10" style="WIDTH: 50px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 225px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		& No.
	</span>
	<span style="TOP: 240px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Street
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtStreetCorrespondence" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 265px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		District
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtDistrictCorrespondence" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 290px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Town
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtTownCorrespondence" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 315px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		County
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtCountyCorrespondence" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 340px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Country
		<span style="LEFT: 70px; POSITION: absolute; TOP: -3px">
			<select id="cboCountryCorrespondence" style="WIDTH: 200px" class="msgCombo"></select>
		</span>
	</span>
</div>	
</form>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 450px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/AP410Attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sUserRole = "" ;
var m_sUserId = "" ;
var m_sUnitId = "" ;
var m_sTaskXML = "";
var taskXML = null;
var XML = null;
var scScreenFunctions;

var strAddressGuid ;
var strRotGuid ;
var scScreenFunctions;
var m_blnReadOnly = false;

var m_sApplicationNumber, m_sApplicationFactFindNumber ;

var iCurrentRow ;
var m_iTableLength = 10;
var m_EntryStatus = null;

/** EVENTS **/


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel","Next");
	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Applicants & Property Address","AP410",scScreenFunctions);

	PopulateCountryCombo() ;
	RetrieveContextData() ;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();	
	Initialise() ;

	scScreenFunctions.SetFocusToFirstField(frmScreen);
		
	// Fetch the combo values and the respective validations for the follwing combos ;
	// later used to populate value name in the form
	xmlCombos = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("SpecInst", 'CMLAddrFlag', 'StandardStatus');
	xmlCombos.GetComboLists(document,sGroupList)
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP410");
	
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
	/* TEST
	
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","C00059072");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
	//scScreenFunctions.SetContextParameter(window,"idCustomerRoleType","1");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber","00000094");
	scScreenFunctions.SetContextParameter(window,"idMortgageApplicationValue","10");
	
	//END TEST */
		
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	
	m_sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	m_sUserId = scScreenFunctions.GetContextParameter(window,"idUserId","");
	m_sUnitId =	scScreenFunctions.GetContextParameter(window,"idUnitId","");
	
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = 	scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","");
		
//DEBUG
//m_sReadOnly = "0" ;
//m_sTaskXML = "<CASETASK TASKID=\"22\" CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"2\" TASKSTATUS=\"20\"/>";	
}

function Initialise()
{
	if(m_sTaskXML.length > 0)
	{
		taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject() ;
		taskXML.LoadXML(m_sTaskXML) ;
		taskXML.ActiveTag = null;
		taskXML.SelectTag(null, "CASETASK");
		m_EntryStatus = taskXML.GetAttribute ("TASKSTATUS");<%/* LDM 27/04/2006 MAR1624 need to know what the task status on entry to the screen */%> 
		PopulateScreen() ;	
	}
}

function PopulateScreen()
{
	XML = null ;
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window , null);
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);

	//Pass XML to ApplicationManager BO
	XML.RunASP(document,"FindCustomersForApplication.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{		
		PopulateApplicantTable() ;
		PopulateSecurityAddress() ;			
		PopulateCorrespondenceAddress() ;<%/* LDM 27/04/2006 MAR1624 */%> 	
	}
	
	//Set Read only for Security Address fields
	scScreenFunctions.SetFieldState(frmScreen, "txtFlatNumber", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtHouseName", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtHouseNumber", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtStreet", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtDistrict", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtTown", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtCounty", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtCountry", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtPostcode", "R");	
	
	if (m_sReadOnly == "1") scScreenFunctions.SetScreenToReadOnly(frmScreen);
}

function PopulateApplicantTable()
{
	var iRowIndex = 0;
	
	XML.ActiveTag = null ;
	XML.SelectTag(null, "CUSTOMERROLELIST") ;
	XML.CreateTagList("CUSTOMERROLE") ;
	
	for (var iCount=0; iCount < XML.ActiveTagList.length ; iCount++)
	{
		XML.SelectTagListItem(iCount) ;
		if (XML.GetTagText("CUSTOMERROLETYPE") == "1")
		{
			sApplicantNumber	= XML.GetTagText (".//CUSTOMERNUMBER") ;
			sApplicantSurname	= XML.GetTagText (".//SURNAME") ;
			sApplicantForename	= XML.GetTagText (".//FIRSTFORENAME") ;
			sApplicantForename += " " + XML.GetTagText (".//SECONDFORENAME") ;
			sApplicantTitle     = XML.GetTagAttribute(".//TITLE","TEXT")
						
			scScreenFunctions.SizeTextToField(tblApplicantList.rows(iRowIndex+1).cells(0),sApplicantNumber);
			scScreenFunctions.SizeTextToField(tblApplicantList.rows(iRowIndex+1).cells(1),sApplicantSurname);
			scScreenFunctions.SizeTextToField(tblApplicantList.rows(iRowIndex+1).cells(2),sApplicantForename);
			scScreenFunctions.SizeTextToField(tblApplicantList.rows(iRowIndex+1).cells(3),sApplicantTitle);
			
			++iRowIndex ;
		}
	}
}

function PopulateSecurityAddress ()
{
	XML = null ; 
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window , "GetMortgagePropertyAddress");
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("TYPEOFAPPLICATION", taskXML.GetAttribute ("TASKSTATUS"));

	//Pass XML to GetNewPropertyAddress BO
	XML.RunASP(document,"GetMortgagePropertyAddress.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)	
	{
		if(XML.SelectTag(null, "ADDRESS") != null)
		{
			frmScreen.txtFlatNumber.value	= XML.GetTagText("FLATNUMBER") ;
			frmScreen.txtHouseName.value	= XML.GetTagText("BUILDINGORHOUSENAME") ;
			frmScreen.txtHouseNumber.value	= XML.GetTagText("BUILDINGORHOUSENUMBER") ;
			frmScreen.txtStreet.value		= XML.GetTagText("STREET") ;
			frmScreen.txtDistrict.value		= XML.GetTagText("DISTRICT") ;
			frmScreen.txtTown.value			= XML.GetTagText("TOWN") ;
			frmScreen.txtCounty.value		= XML.GetTagText("COUNTY") ;
			frmScreen.txtCountry.value		= XML.GetTagAttribute("COUNTRY","TEXT") ;
			frmScreen.txtPostcode.value		= XML.GetTagText("POSTCODE") ;
		}
	}
	else
	{
		if (ErrorReturn[1] == "RECORDNOTFOUND")
			alert("No Security Address exists.") ;
	}	
}

function PopulateCorrespondenceAddress()
{
	XML = null ;
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject() ; 
	var reqTag = XML.CreateRequestTag(window, "GetReportOnTitleData");
	XML.CreateActiveTag("REPORTONTITLE");
	XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	
	//Pass XML to GetCorrespondenceAddress BO
	XML.RunASP(document,"ReportOnTitle.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)	
	{
		taskXML.SetAttribute("TASKSTATUS",20); <%/* LDM 27/04/2006 MAR1624 force to be a status of update if ROT rec is there already */%> 
		scScreenFunctions.SetContextParameter(window,"idTaskXML",taskXML.XMLDocument.xml); <%/* LDM 27/04/2006 MAR1624 */%>
		XML.SelectTag(null, "REPORTONTITLE");
		strRotGuid = XML.GetAttribute("ROTGUID") ;
		strAddressGuid = "";
		<% /* JD MAR1698 get the correspondence address from the list*/ %>
		var xmlAddrList = XML.CreateTagList("ADDRESS")
		for (var iCount=0; iCount < XML.ActiveTagList.length ; iCount++)
		{
			XML.SelectTagListItem(iCount) ;
			//Populate if address type is (C)ompletion
			if (XML.GetAttribute("ADDRESSTYPE") == "10")
			{
				strAddressGuid = XML.GetAttribute("ADDRESSGUID") ;
				if (XML.GetAttribute("FLATNUMBER") != null) frmScreen.txtFlatNumberCorrespondence.value = XML.GetAttribute("FLATNUMBER");
				if (XML.GetAttribute("BUILDINGORHOUSENAME")!= null) frmScreen.txtHouseNameCorrespondence.value = XML.GetAttribute("BUILDINGORHOUSENAME");
				if (XML.GetAttribute("BUILDINGORHOUSENUMBER") != null) frmScreen.txtHouseNumberCorrespondence.value = XML.GetAttribute("BUILDINGORHOUSENUMBER");
				if (XML.GetAttribute("STREET") != null) frmScreen.txtStreetCorrespondence.value = XML.GetAttribute("STREET");
				if (XML.GetAttribute("DISTRICT") != null) frmScreen.txtDistrictCorrespondence.value = XML.GetAttribute("DISTRICT");
				if (XML.GetAttribute("TOWN") != null) frmScreen.txtTownCorrespondence.value = XML.GetAttribute("TOWN") ;
				if (XML.GetAttribute("COUNTY") != null) frmScreen.txtCountyCorrespondence.value = XML.GetAttribute("COUNTY");
				if (XML.GetAttribute("COUNTRY") !=null) frmScreen.cboCountryCorrespondence.value = XML.GetAttribute("COUNTRY");
				if (XML.GetAttribute("POSTCODE") != null) frmScreen.txtPostcodeCorrespondence.value = XML.GetAttribute("POSTCODE");
			}
		}
	}
}

function ReportOnTitleAction ()
{
	XML = null ;
	var reqTag ;

	if ( (taskXML.GetAttribute("TASKID") == null) || (m_sReadOnly == "1") ) return true ;
		
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject() ;
	
	//Update report on title record
	if (taskXML.GetAttribute ("TASKSTATUS") == "20")
	{
		reqTag = XML.CreateRequestTag(window , "UpdateReportOnTitle");
	
		//Complete the XML request 
		SetAddressAttributes(reqTag, "Update") ; 
			
		//Pass XML to ReportOnTitleBO
		XML.RunASP(document,"ReportOnTitle.asp");
			
		//Verify XML response
		<% /* LDM 27/04/2006 MAR1624 */ %>
		if(XML.IsResponseOK())
		{
			if (m_EntryStatus == "20")
			{
				return true;
			}
			else (m_EntryStatus == "10")
			{
				<% /* LDM 27/04/2006 MAR1624 - If everything worked ok, need to set the task to pending.
					If this is the first time into the cot screens by "any" action; then the status will be automatically moved
					to pending, when the createreportontitle is run.
					If we got here by the details btn in the application stage. Then the status is already pending, so do not reset to pending
					If we got here by creating another cot action and this is the first time in need to reset the status to pending  */ %>
				SetToPendingXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				SetToPendingXML.CreateRequestTag(window , "UpdateCaseTask");
				SetToPendingXML.CreateActiveTag("CASETASK");
				SetToPendingXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
				SetToPendingXML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
				SetToPendingXML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
				SetToPendingXML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
				SetToPendingXML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
				SetToPendingXML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
				SetToPendingXML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
				SetToPendingXML.SetAttribute("TASKSTATUS", "20"); <% /* Pending */ %>
				
				SetToPendingXML.RunASP(document, "msgTMBO.asp") ;
				
				if (SetToPendingXML.IsResponseOK())
				{
					return true;
				}
				else
				{
					return false;
				}
			}
		}
		else
		{
			return false;
		}
		<% /* LDM 27/04/2006 MAR1624 */ %>
	}
	
	//Create a Report on Title Record	
	if (taskXML.GetAttribute ("TASKSTATUS") == "10")
	{					
		reqTag = XML.CreateRequestTag(window , "CreateReportOnTitle");
		XML.SetAttribute("USERID", m_sUserId) ;
		XML.SetAttribute("UNITID", m_sUnitId) ;
		XML.SetAttribute("USERAUTHORITYLEVEL", m_sUserRole) ;
						
		XML.ActiveTag = reqTag ;
		XML.CreateActiveTag("CASETASK");
		XML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
		XML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
		XML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
		XML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
		XML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
		// JR - SYS2048, add the following attributes to Request
		XML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
		XML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
							
		//Complete the XML request
		SetAddressAttributes(reqTag, "Create") ;
		
		//Pass XML to OmTmBO.CreateReportonTitleData
		XML.RunASP(document,"OmigaTMBO.asp");
			
		//Verify XML response
		
		//BG 26/11/01 SYS3107 need to change the TaskXML held in context to reflect
		//the task has been moved to pending from not started.
		//if(!XML.IsResponseOK())	return false ;
		
		if(!XML.IsResponseOK())	
		{
			return false ;
		}
		else
		{
			taskXML.SetAttribute("TASKSTATUS",20);
			scScreenFunctions.SetContextParameter(window,"idTaskXML",taskXML.XMLDocument.xml);
		}				
	}	
	return true ;
}

function SetAddressAttributes(reqTag, sOperation) 
{
	XML.ActiveTag = reqTag ;
	XML.CreateActiveTag("REPORTONTITLE") ;
	XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber) ;
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber) ;
	// JR SYS2048 - only set ROTGUID attribute on Updates
	if (sOperation == "Update") XML.SetAttribute("ROTGUID", strRotGuid) ;
		
	XML.CreateActiveTag("ADDRESS");
	if (sOperation == "Update" && strAddressGuid != "") XML.SetAttribute("ADDRESSGUID", strAddressGuid);
	XML.SetAttribute("BUILDINGORHOUSENAME", frmScreen.txtHouseNameCorrespondence.value);
	XML.SetAttribute("BUILDINGORHOUSENUMBER", frmScreen.txtHouseNumberCorrespondence.value);
	XML.SetAttribute("FLATNUMBER", frmScreen.txtFlatNumberCorrespondence.value);
	XML.SetAttribute("STREET", frmScreen.txtStreetCorrespondence.value);
	XML.SetAttribute("DISTRICT", frmScreen.txtDistrictCorrespondence.value);
	XML.SetAttribute("TOWN", frmScreen.txtTownCorrespondence.value);
	XML.SetAttribute("COUNTY", frmScreen.txtCountyCorrespondence.value);
	XML.SetAttribute("COUNTRY", frmScreen.cboCountryCorrespondence.value);
	XML.SetAttribute("POSTCODE", frmScreen.txtPostcodeCorrespondence.value);
	
	var sROTAddressTypeValue = GetRotAddressTypeValue() ;
	XML.SetAttribute("ADDRESSTYPE", sROTAddressTypeValue) ;	
}

function GetRotAddressTypeValue()
{
	var addressXML = null ;
	addressXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("ROTAddressType");
	var sRecipientList = addressXML.GetComboLists(document, sGroups);
	var sValidationType = "C"; //this is the validation type for Completion Correspondence Address 
	var sROTAddressType = addressXML.GetTagText("LIST/LISTNAME/LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='" + sValidationType + "']/VALUEID");		

	return sROTAddressType ;	
}

function PopulateCountryCombo()
{
	XML = null ;
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("Country");
	var bSuccess = false;
	if (XML.GetComboLists(document, sGroups) == true)
	{
		bSuccess = true;
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboCountryCorrespondence,"Country",true);
	}
}

function CommitChanges()
{
	var bSuccess = true;
	//if (m_sReadOnly != "1")
	if(frmScreen.onsubmit())
	{
		if (IsChanged()) bSuccess = ReportOnTitleAction();
	}
	else
		bSuccess = false;
	return(bSuccess);
}

function btnSubmit.onclick() 
{
	if (CommitChanges()) frmToAP400.submit() ;
}

function btnNext.onclick()
{
	if (CommitChanges()) frmToAP420.submit() ;
}

function btnCancel.onclick()
{
	taskXML.SetAttribute("TASKSTATUS",m_EntryStatus); <%/* LDM 02/04/2006 MAR1624 reset to initial status */%> 
	scScreenFunctions.SetContextParameter(window,"idTaskXML",taskXML.XMLDocument.xml); <%/* LDM 02/05/2006 MAR1624 */%>
	frmToAP400.submit() ;
}


-->
</script>
</BODY>
</HTML>



