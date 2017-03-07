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

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 370px; LEFT: 310px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Forms Here. GN300 is needed for MAIN screens to route back to completeness check screen.
      A MAIN screen is not a pop-up or a screen routed to via any button other than the main submit/cancel buttons */ %>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToXX999" method="post" action="XX999.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 340px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/xx999attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_blnReadOnly = false;
var m_sReadOnly = "";
var scScreenFunctions;

function window.onload()
{
	var sScreenId = "XX000";
	
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("A Screen Title",sScreenId,scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
		
	m_bReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, sScreenId);
	if (m_blnReadOnly == true) m_sReadOnly = "1";

	scScreenFunctions.SetFocusToFirstField(frmScreen);
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
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
}

function btnSubmit.onclick()
{
<% // An example submit function showing the use of the validation functions 
%>	if(frmScreen.onsubmit())
	{
		if(IsChanged())
<%			// Do some processing
%>
<%		//for a MAIN screen, always check if we've been called from completeness check.
		//A Main screen is not a pop-up or a screen routed to from a button or drill down other than the main buttons.
%>		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else
			frmToXX999.submit();
	}
}
function btnCancel.onclick()
{
<% //if a cancel button is required, always check if we've been called from completeness check and route accordingly
%>
<%	//for a MAIN screen, always check if we've been called from completeness check.
	//A Main screen is not a pop-up or a screen routed to from a button or drill down other than the main buttons.
%>	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else		
		frmToXX999.submit();
}
-->
</script>
</body>
</html>
