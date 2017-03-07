<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      SP093.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Sanction/Unsanction Authorisation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
CL		14/02/2001	SYS1982 Initial creation 
CL		05/03/01	SYS1920 Read only functionality added
CL		07/03/01	SYS2002 Removed Context detail
CL		07/03/01	SYS2002 Removed FW030 detail
CL		07/03/01	SYS2002 Adjustment for Popup functionality 
APS		14/03/01	SYS2061 Make the password control mask out the chars as *
CL		15/03/01	SYS2069 Remove acceptance message on password entry
CL		20/03/01	SYS2069 Added screen title
APS		22/03/01	SYS2146
SR		06/09/01	SYS2412
LD		23/05/02	SYS4727 Use cached versions of frame functions
PJO     13/10/2003  BMIDS643	- Password length increased to 15
MC		21/04/2004	BMIDS517	- Background DIV size changed to fix title text (IE 6.0 Display bigfont)
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
<title>Sanction/Unsanction Authorisation  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange" year4>
	<div id="divSanctionDetails" style="TOP: 10px; LEFT: 10px; HEIGHT: 140px; WIDTH: 365px; POSITION: ABSOLUTE" class="msgGroup">
		<span style="TOP:15px; LEFT:30px; POSITION:ABSOLUTE" class="msgLabel">
			Total number of payments
			<span style="TOP:-3px; LEFT:150px; POSITION:ABSOLUTE">
				<input id="txtNumberOfPayments" maxlength="10" style="WIDTH:60px" class="msgReadOnly" readonly tabindex="-1">
			</span>
		</span>
		<span style="TOP:40px; LEFT:30px; POSITION:ABSOLUTE" class="msgLabel">
			Total value of payments
			<span style="TOP:-3px; LEFT:150px; POSITION:ABSOLUTE">
				<input id="txtValueOfPayments" maxlength="10" style="WIDTH:60px" class="msgReadOnly" readonly tabindex="-1">
			</span>
		</span>
		<span style="TOP:75px; LEFT:30px; POSITION:ABSOLUTE" class="msgLabel">
			User Password
	<% /* PJO 13/10/2003 BMIDS643 - Password length increased to 15 */ %>
			<span style="TOP:-3px; LEFT:150px; POSITION:ABSOLUTE">
				<input id="txtPassword" type="PASSWORD" maxlength="15" style="WIDTH:100px" class="msgTxt">
			</span>
		</span>
		<!-- Buttons -->
		<span id="spnButtons" style="LEFT: 60px; POSITION: absolute; TOP: 110px">
			<span style="LEFT: 1px; POSITION: absolute; TOP: 0px">
				<input id="btnOK" value="OK" type="button" 
					style="WIDTH: 80px" class="msgButton">
				</span>
			<span style="LEFT: 104px; POSITION: absolute; TOP: 0px">
				<input id="btnCancel" value="Cancel" type="button" 
					style="WIDTH: 80px" class="msgButton">
			</span>
		</span>
	</div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/SP093Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_ArraySelectedRows
var m_blnReadOnly = false;
var m_sUserID = "0";
var m_UnitID = "0";
var m_BaseNonPopupWindow = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;

function window.onunload()
{
	//invoke onclose event
	frmScreen.btnCancel.onclick();
}

function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();

	//frmScreen.txtValueOfPayments = null;
	//frmScreen.txtNumberOfPayments = null; 
	RetrieveData();
	scScreenFunctions.ShowCollection(frmScreen);
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveData()
{
	var ArraySelectedRows = new Array();
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters	= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	frmScreen.txtValueOfPayments.value = sParameters[0];
	frmScreen.txtNumberOfPayments.value = sParameters[1];
	m_ArraySelectedRows = sParameters[2];
	m_sUserID = sParameters[3];
	m_UnitID = sParameters[4];
}

function frmScreen.btnOK.onclick()
{
	ValidatePassword();	
}

function frmScreen.btnCancel.onclick()
{
	var sReturn = new Array();
	sReturn[0] = false;
	window.returnValue = sReturn;
	window.close();
}

function ValidatePassword()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateActiveTag("REQUEST");
	XML.SetAttribute("USERID", m_sUserID);
	XML.CreateActiveTag("OMIGAUSER");
	XML.CreateTag("USERID", m_sUserID);
	XML.CreateTag("PASSWORDVALUE", frmScreen.txtPassword.value);
	XML.CreateTag("UNITID", m_UnitID);
	XML.CreateTag("AUDITRECORDTYPE", "1");
	// 	XML.RunASP(document, "ValidateCurrentPassword.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "ValidateCurrentPassword.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}


	if (XML.IsResponseOK() == true)
	{
		var bAccessAuditAttempt = XML.GetTagBoolean("ACCESSAUDITATTEMPT");
		//Now go and update the rows
		UpdateRows();
	}
}

function UpdateRows()
{
	XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(m_ArraySelectedRows);
		
	//Run the update method
	// 	XML.RunASP(document,"UpdateDisbursement.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"UpdateDisbursement.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	if (XML.IsResponseOK() == true) {	
		var sReturn = new Array();

		sReturn[0] = true;

		window.returnValue = sReturn;
		window.close();		
	}
}		

-->
</script>
</body>
</html>





