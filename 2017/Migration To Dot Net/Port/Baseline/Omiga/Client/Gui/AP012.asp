<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP012.asp
Copyright:     Copyright © 2001 Marlborough Stirling

Description:   Password
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
PSC		23/02/01	Created
JLD		06/03/01	Added method call.
JLD		09/03/01	check for empty fields on OK
LD		23/05/02	SYS4727 Use cached versions of frame functions
PJO     13/10/2003  BMIDS643	- Password length increased to 15

MARS Specific changes
----------------------

PJO     21/11/2005  MAR417 - Increase User ID length to 30 chars
PJO     30/11/2005  MAR417 - Increase User ID length to 64 chars
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MC		19/04/2004		White space padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Password <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 200px; WIDTH: 400px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP:40px; LEFT:40px; POSITION:ABSOLUTE" class="msgLabel">
		User Id
	<% /* PJO 30/10/2003 MAR417 - User ID length increased to 64 */ %>
		<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
			<input id="txtUserId" maxlength="64" style="WIDTH:240px" class="msgTxt">
		</span>
	</span>
	<span style="TOP:70px; LEFT:40px; POSITION:ABSOLUTE" class="msgLabel">
		Unit Name
		<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
			<input id="txtUnitName" maxlength="10" style="WIDTH:100px" class="msgTxt">
		</span>
	</span>
	<span style="TOP:100px; LEFT:40px; POSITION:ABSOLUTE" class="msgLabel">
		User Password
	<% /* PJO 13/10/2003 BMIDS643 - Password length increased to 15 */ %>
		<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
			<input id="txtUserPassword" type=PASSWORD maxlength="15" style="WIDTH:100px" class="msgTxt">
		</span>
	</span>
	<span id="spnButtons" style="TOP: 150px; LEFT: 40px; POSITION: ABSOLUTE">
		<span style="TOP:0px; LEFT:0px; POSITION:ABSOLUTE">
			<input id="btnOK" value="OK" type="button" style="WIDTH:60px" class="msgButton">
		</span>
		<span style="TOP:0px; LEFT:64px; POSITION:ABSOLUTE">
			<input id="btnCancel" value="Cancel" type="button" style="WIDTH:60px" class="msgButton">
		</span>
	</span>
</div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/AP012Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_scScreenFunctions;
var m_sReadOnly = "";
var m_sUnitId = "";
var m_sUnitName = "";
var m_sRequestArray = "";
var m_sApplicationNumber = "";
var m_sApplicationFFNumber = "";
var m_BaseNonPopupWindow = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	RetrieveData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	//AQR SYS3134 DRC - SetScreenOnReadOnly Removed
	scScreenFunctions.ShowCollection(frmScreen);
	scScreenFunctions.SetFieldState(frmScreen, "txtUnitName", "R");
	frmScreen.txtUnitName.value = m_sUnitName;
	frmScreen.txtUserId.focus();

	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
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
	m_sRequestArray = sParameters[0];
	m_sReadOnly	= sParameters[1];
	m_sUnitId =  sParameters[2];
	m_sUnitName = sParameters[3];
	m_sApplicationNumber = sParameters[4];
	m_sApplicationFFNumber = sParameters[5];
}
function frmScreen.btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}
function frmScreen.btnOK.onclick()
{
	if(frmScreen.txtUserId.value.length == 0 || frmScreen.txtUserPassword.value.length == 0)
		alert("UserId and Password must be entered");
	else
	{
		var sReturn = new Array();
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		var tagReq = XML.CreateRequestTagFromArray(m_sRequestArray,"ValidateUserLogon");
		XML.CreateActiveTag("USER");
		XML.SetAttribute("USERID", frmScreen.txtUserId.value);
		XML.SetAttribute("UNITID", m_sUnitId);
		XML.SetAttribute("PASSWORDVALUE", frmScreen.txtUserPassword.value);
		XML.ActiveTag = tagReq;
		XML.CreateActiveTag("APPLICATION");
		XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		// 		XML.RunASP(document, "omAppProc.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "omAppProc.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		if(XML.IsResponseOK())
		{
			XML.SelectTag(null, "USER");
			sReturn[0] = IsChanged();
			if(XML.GetAttribute("VALIDUSER") == "1")
				sReturn[1] = frmScreen.txtUserId.value;
			else
			{
				sReturn[1] = "";
				alert("User " + frmScreen.txtUserId.value + " is not authorised.");
			}
			window.returnValue	= sReturn;
		}
		else window.returnValue	= null;
		window.close();
	}
}
-->
</script>
</body>
</html>




