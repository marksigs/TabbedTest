<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      ?????.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   ?????
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

/*	
	IMPORTANT INFORMATION

	Pop-up's have no access to the context parameters, so data must be passed from the calling screen.
	
	When calling a Middle Tier method from a main GUI screen we make a call to :-
	
	XML.CreateRequestTag(window, sOperation);
	
	This method will set-up the <REQUEST> tag and also assign a standard set of attributes to the request, at the time
	of writing, the tag generated would look like this:-
	
	<REQUEST USERID="" UNITID="" MACHINEID="" CHANNELID="" USERAUTHORITYLEVEL="" OPERATION="">	
	
	When calling a Middle Tier method from a popup GUI screen we do not have access the context parameters, therefore
	the call to XML.CreateRequestTag will NOT work. To get around this we need to pass into the popup all the
	information the popup requires to run a middle tier method. 
	
	The XML.CreateRequestTag method is broken down into two methods:-
	
		1. XML.CreateRequestAttributeArray();
		2. XML.CreateRequestTagFromArray(AttributeArray,sOperation);
	
	We use these two methods to call a Middle Tier method from a popup screen, as described below:-
	
	FROM A MAIN SCREEN...
	
	From the main GUI screen that is calling the popup we create the request attribute array and pass this
	array through to the popup as another element in the array:-
	
	var sReturn = null;
	var ArrayArguments = new Array();
	
	ArrayArguments[0] = XML.CreateRequestAttributeArray(window);	// returns an attribute array
	ArrayArguments[1] = m_sCustomer1Name;							// other parameters required for screen...
	ArrayArguments[2] = m_sCustomer1Number;
	ArrayArguments[3] = m_sCustomer1VersionNumber;
	ArrayArguments[4] = m_sReadOnly;								// always pass through read only flag

	sReturn = scScreenFunctions.DisplayPopup(window, document, "nameOfPopup.asp", ArrayArguments, <width>470, <height>400);
	
	if (sReturn != null)
	{
		// set dirty flag as we may not want to re-display screen and/or save data if the user has changed/added data
		FlagChange(sReturn[0]);											
	}
	
	FROM A POPUP...
	
	Then to create a <REQUEST> tag from within the Popup screen, we need to call the second method to use the attribute 
	array:-
	
	XML.CreateRequestTagFromArray(ArrayArgument[0], sOperation);
	XML.CreateActiveTag(sTagName);
	XML.CreateTag(sTagName);
	XML.RunASP(sASPFile);
	...

	This approach allows us to be more flexible on the attributes sent to a Middle Tier call, if we require to extend
	the set of attributes then we only need to change the CreateRequestAttributeArray method.
	
	See scXMLFunctions.asp for more details on the XML scriplet methods...
*/
%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<%/* FILL IN THE TITLE OF THE SCREEN HERE */%>
<title>TITLE</title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 250px; LEFT: 210px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 280px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">

</div>
</form>

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_BaseNonPopupWindow = null;

function window.onload()
{
	RetrieveData();
	Validation_Init();
	SetScreenOnReadOnly();
	window.returnValue = null;
}
function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
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
	m_sReadOnly			= sParameters[0];
<% /* fetch any other parameters passed by the calling screen here */ %>
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
	window.returnValue	= sReturn;
	window.close();
}
-->
</script>
</body>
</html>
