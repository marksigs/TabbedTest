<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      MsgBox.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   MsgBox Functionality
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog		Date		Description
TW		09 Oct 2002	SYS5115 Initial creation 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		Description
SD		25 Oct 2005	MAR227 Warning, If Ok button is clicked before Send Details. 
					Changed the height of the message box and the position of Yes, No buttons
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Omiga 4</title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; 

VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen">
	<div id="divMsgBoxDetails" style="TOP: 10px; LEFT: 10px; HEIGHT: 140px; WIDTH: 200px; POSITION: ABSOLUTE" 

class="msgPageInfo">

		<img alt border="0" id="imgIcon" src="" style="LEFT: 1px; POSITION: absolute; TOP: 1px; Z-INDEX: 2" 

WIDTH="30" HEIGHT="30">				

		<label id="lblMessage" style="LEFT: 2px; POSITION: absolute; TOP: 40px; WIDTH: 194px; HEIGHT: 100px" 

class="msgLabelHead">
		</label>
		<input id="btnOK" value="OK" type="button" style="LEFT: 15px; POSITION: absolute; TOP: 145px; WIDTH: 

80px" class="msgButton">
		<input id="btnCancel" value="Cancel" type="button" style="LEFT: 105px; POSITION: absolute; TOP: 

145px; WIDTH: 80px" class="msgButton">
	</div>
</form>

<% /* Specify Code Here */ %>
<script language="JScript">
<!--

var sType = null; 

function window.onload()
{
	var sArguments = window.dialogArguments;
	var m_BaseNonPopupWindow = sArguments[0];
	var sMessage = sArguments[1];
	sType = sArguments[2];
	frmScreen.imgIcon.src = sArguments[3];
	if (sType == 1)
	{
		frmScreen.btnOK.value = sArguments[4];
		frmScreen.btnCancel.value = sArguments[5];
	}
	else
	{
		frmScreen.btnOK.style.left = "60px";
		frmScreen.btnCancel.style.visibility="hidden";
	}

	window.returnValue = new Array(2);

	window.returnValue[0] = 2;		// Default to 'Error'
	window.returnValue[1] = sMessage;	// User defined message

	dialogTop = (window.screen.height - 200) / 2 + "px" 

	window.dialogTop = ((window.screen.height - 200) / 2) + "px";
	window.dialogLeft = ((window.screen.width - 230)) / 2 + "px";
	window.dialogWidth = "230px";
	window.dialogHeight = "265px";
	  
	var scScreenFunctions = new 		

m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	scScreenFunctions.SizeTextToField(lblMessage,sMessage);

}

function frmScreen.btnOK.onclick()
{
	window.returnValue[0] = sType;
	window.close();
}

function frmScreen.btnCancel.onclick()
{
	window.close();
}

-->
</script>
</body>
</html>
