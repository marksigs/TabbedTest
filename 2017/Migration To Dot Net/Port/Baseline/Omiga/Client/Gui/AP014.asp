<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP014.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Approve/Recommend reason POP-UP
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		23/02/01	Created
JLD		09/03/01	SYS1879 added method call
STB		05/04/02	SYS3395 Allow filtering on other reason field.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:
Prog	Date		Description
MO		14/11/2002	BMIDS00807  - Made change to take the date from the app server not the client
MC		05/05/2004	BMIDS751	- White space added to the title
KRW     20/10/2004  BM0519      - Enable Other Reason Text box on entry if Validation Type = "O" thats letter O

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
<title>Approve/Recommend Application <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate ="onchange">
<div id="divBackground" style="HEIGHT: 160px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 510px" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Reason
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<select id="cboReason" style="WIDTH: 300px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Other Reason Text
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px"><TEXTAREA class=msgTxt id=txtOtherReasonText rows=5 style="WIDTH: 300px"></TEXTAREA>
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 130px">
	<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
</span>
<span style="LEFT: 80px; POSITION: absolute; TOP: 130px">
	<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
</span>
</div>
</form>

<%/* SYS3395: File containing field attributes (remove if not required) */%>
<!--  #include FILE="attribs/AP014attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sRequestArray = "";
var m_sContext = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sUserId = "";
var m_sAuthorisedUserId = "";
var m_sUnitId = "";
var m_sTaskXML = "";
var m_sAuthorityLevel = "";
var m_sAppPriority = "";
var m_sComboGroup = "";
var m_BaseNonPopupWindow = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	// SYS3395: This function is contained in the field attributes file (remove if not required)
	SetMasks();
	
	RetrieveData();
	Validation_Init();
	PopulateScreen();
	SetScreenOnReadOnly();
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}
function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
}
function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters		= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sRequestArray		= sParameters[0];
	m_sContext			= sParameters[1];
	m_sApplicationNumber = sParameters[2];
	m_sApplicationFactFindNumber = sParameters[3];
	m_sUserId			= sParameters[4];
	m_sAuthorisedUserId = sParameters[5];
	m_sUnitId			= sParameters[6];
	m_sTaskXML			= sParameters[7];
	m_sAuthorityLevel   = sParameters[8];
	m_sAppPriority      = sParameters[9];
}
function PopulateScreen()
{
	m_sComboGroup = "";
	if(m_sContext == "Approve") m_sComboGroup = "ApprovalReason";
	else m_sComboGroup = "RecommendReason";
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();			
	var sGroups = new Array(m_sComboGroup);
	if (XML.GetComboLists(document, sGroups) == true)
	{
		XML.PopulateCombo(document, frmScreen.cboReason, m_sComboGroup, false);
	}
	// KRW BM0519 20/10/2004
	if(XML.IsInComboValidationList(document,m_sComboGroup, frmScreen.cboReason.value, Array("O")))
	{
		scScreenFunctions.SetFieldState(frmScreen, "txtOtherReasonText", "W");
	}
	else
	{
		scScreenFunctions.SetFieldState(frmScreen, "txtOtherReasonText", "R");
	}
	XML = null;
}
function frmScreen.cboReason.onchange()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	if(XML.IsInComboValidationList(document,m_sComboGroup, frmScreen.cboReason.value, Array("O")))
		scScreenFunctions.SetFieldState(frmScreen, "txtOtherReasonText", "W");
	else
	{
		scScreenFunctions.SetFieldState(frmScreen, "txtOtherReasonText", "R");
		frmScreen.txtOtherReasonText.value = "";
	}
	XML = null;
}
function frmScreen.btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}
function frmScreen.btnOK.onclick()
{
	var sReturn = new Array();
	
	<% /* MO - BMIDS00807 */ %>
	<% /* var sDateTime = scScreenFunctions.DateTimeToString(scScreenFunctions.GetSystemDateTime()); */ %>
	var sDateTime = scScreenFunctions.DateTimeToString(scScreenFunctions.GetAppServerDate(true));
	<% /* var sDate = scScreenFunctions.DateToString(scScreenFunctions.GetSystemDate()); */ %>
	var sDate = scScreenFunctions.DateToString(scScreenFunctions.GetAppServerDate());
	var taskXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	taskXML.LoadXML(m_sTaskXML);
	taskXML.SelectTag(null, "CASETASK");
	var updateXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = updateXML.CreateRequestTagFromArray(m_sRequestArray,"ApproveRecommendApplication");
	updateXML.CreateActiveTag("APPROVREC");
	updateXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	updateXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	updateXML.SetAttribute("ACTIONNAME", m_sContext);
	updateXML.SetAttribute("ENTRYDATETIME", sDateTime);
	if(frmScreen.txtOtherReasonText.value.length == 0)
		updateXML.SetAttribute("MEMOENTRY", frmScreen.cboReason.options(frmScreen.cboReason.selectedIndex).text);
	else
		updateXML.SetAttribute("MEMOENTRY", frmScreen.txtOtherReasonText.value);
	updateXML.SetAttribute("SCREENDESCRIPTION", "Approval Summary");
	updateXML.SetAttribute("SCREENREFERENCE", "AP010");
	updateXML.SetAttribute("ENTRYTYPE", "8");
	if(m_sContext == "Approve")
	{
		updateXML.SetAttribute("APPLICATIONAPPROVALDATE", sDate);
		updateXML.SetAttribute("APPLICATIONAPPROVALUSERID", m_sAuthorisedUserId);
		updateXML.SetAttribute("APPLICATIONAPPROVALUNITID", m_sUnitId);
	}
	else
	{
		updateXML.SetAttribute("APPLICATIONRECOMMENDEDDATE", sDate);
		updateXML.SetAttribute("APPLICATIONRECOMMENDEDUSERID", m_sAuthorisedUserId);
		updateXML.SetAttribute("APPLICATIONRECOMMENDEDUNITID", m_sUnitId);
	}
	updateXML.SetAttribute("UNITID", m_sUnitId);
	updateXML.SetAttribute("USERID", m_sUserId);
	updateXML.SetAttribute("USERAUTHORITYLEVEL", m_sAuthorityLevel);
	updateXML.SetAttribute("APPLICATIONPRIORITY", m_sAppPriority);
	updateXML.ActiveTag = reqTag;
	updateXML.CreateActiveTag("CASETASK");
	updateXML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
	updateXML.SetAttribute("CASEID", m_sApplicationNumber);
	updateXML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
	updateXML.SetAttribute("ACTIVITYINSTANCE", taskXML.GetAttribute("ACTIVITYINSTANCE"));
	updateXML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
	updateXML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
	updateXML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
	updateXML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
	if(m_sContext == "Approve")
		updateXML.SetAttribute("TASKSTATUS", "40"); // Complete
	else updateXML.SetAttribute("TASKSTATUS", "20"); // Pending
	// 	updateXML.RunASP(document, "OmigaTMBO.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			updateXML.RunASP(document, "OmigaTMBO.asp");
			break;
		default: // Error
			updateXML.SetErrorResponse();
		}

	if(updateXML.IsResponseOK())
	{
		sReturn[0] = IsChanged();
		window.returnValue	= sReturn;
		window.close();
	}
}
-->
</script>
</body>
</html>


