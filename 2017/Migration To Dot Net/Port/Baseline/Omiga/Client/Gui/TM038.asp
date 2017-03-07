<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      TM038.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Reason for Decline/Cancel POP-UP
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		01/03/01	Created
JLD		02/03/01	Return the reason combo value not the text
LD		23/05/02	SYS4727 Use cached versions of frame functions
MC		06/05/2004	white space added to page title
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Exception Reason <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange" >
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 125px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:30px; LEFT:50px; POSITION:ABSOLUTE" class="msgLabel">
	Reason
	<span style="TOP:-3px; LEFT:60px; POSITION:ABSOLUTE">
		<select id="cboReason" style="WIDTH:300px" class="msgCombo"></select>
	</span>
</span>
<span style="TOP:90px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnOK" value="OK" type="button" style="WIDTH:60px" class="msgButton">
</span>
<span style="TOP:90px; LEFT:70px; POSITION:ABSOLUTE">
	<input id="btnCancel" value="Cancel" type="button" style="WIDTH:60px" class="msgButton">
</span>
</div>
</form>

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReason = "";
var m_BaseNonPopupWindow = null;

function window.onload()
{
	RetrieveData();
	Validation_Init();
	PopulateScreen();
	SetScreenOnReadOnly();
	window.returnValue = null;
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
	m_BaseNonPopupWindow = sArguments[5];
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
}
function PopulateScreen()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();			
	var sGroups = new Array("ExceptionReason");
	if (XML.GetComboLists(document, sGroups) == true)
		XML.PopulateCombo(document, frmScreen.cboReason, "ExceptionReason", false);
}
function frmScreen.btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}
function frmScreen.btnOK.onclick()
{
	var sReturn = new Array();
	sReturn[0] = IsChanged();
	//sReturn[1] = frmScreen.cboReason.options(frmScreen.cboReason.selectedIndex).text;
	sReturn[1] = frmScreen.cboReason.value;
	window.returnValue	= sReturn;
	window.close();
}
-->
</script>
</body>
</html>
