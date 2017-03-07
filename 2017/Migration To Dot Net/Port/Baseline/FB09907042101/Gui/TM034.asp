<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      TM034.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Chase up contact details POP-UP
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		30/10/00	Created (screen paint)
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MC		05/05/2004	BMIDS751	- White space added to the title
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Chase Up Contact Details<!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 170px; WIDTH: 410px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Name
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtName" maxlength="10" style="WIDTH:250px" class="msgTxt">
	</span>
</span>
<span style="TOP:35px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Company
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtCompany" maxlength="10" style="WIDTH:250px" class="msgTxt">
	</span>
</span>
<span style="TOP:60px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Telephone
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtTelephone" maxlength="10" style="WIDTH:250px" class="msgTxt">
	</span>
</span>
<span style="TOP:85px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Fax
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtFax" maxlength="10" style="WIDTH:250px" class="msgTxt">
	</span>
</span>
<span style="TOP:110px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	E-mail
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtEmail" maxlength="10" style="WIDTH:250px" class="msgTxt">
	</span>
</span>
<span style="TOP:140px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnOK" value="OK" type="button" style="WIDTH:60px" class="msgButton">
</span>
<span style="TOP:140px; LEFT:70px; POSITION:ABSOLUTE">
	<input id="btnCancel" value="Cancel" type="button" style="WIDTH:60px" class="msgButton">
</span>
</div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/TM034Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_sTaskXML = "";
var m_bDisplay = false;
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
	SetScreenOnReadOnly();
	PopulateScreen();
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
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
	m_sTaskXML			= sParameters[1];
	m_sRequestAttributes = sParameters[2];
}
function PopulateScreen()
{
	scScreenFunctions.SetFieldState(frmScreen, "txtName", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtCompany", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtTelephone", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtFax", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtEmail", "R");
	
	taskXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	taskXML.LoadXML(m_sTaskXML);
	taskXML.SelectTag(null, "CASETASK");
	
	ContactDetailsXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	ContactDetailsXML.CreateRequestTagFromArray(m_sRequestAttributes, "GetTaskContactDetails");
	ContactDetailsXML.CreateActiveTag("CASETASK");
	
	ContactDetailsXML.SetAttribute("CONTACTTYPE", taskXML.GetAttribute("CONTACTTYPE"));
	ContactDetailsXML.SetAttribute("CUSTOMERIDENTIFIER", taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
	ContactDetailsXML.SetAttribute("CONTEXT", taskXML.GetAttribute("CONTEXT"));
	ContactDetailsXML.RunASP(document,"OmigaTMBO.asp");

	var ErrorOwnerTypes = new Array("RECORDNOTFOUND");
	var ErrorOwnerReturn = ContactDetailsXML.CheckResponse(ErrorOwnerTypes);
	if (ErrorOwnerReturn[0] == true) 
	{
		ContactDetailsXML.SelectTag(null,"CONTACTDETAILS");
		frmScreen.txtName.value = ContactDetailsXML.GetAttribute("CONTACTNAME");
		frmScreen.txtCompany.value = ContactDetailsXML.GetAttribute("CONTACTCOMPANY");
		frmScreen.txtTelephone.value = ContactDetailsXML.GetAttribute("TELEPHONENUMBER");
		frmScreen.txtFax.value = ContactDetailsXML.GetAttribute("FAXNUMBER");
		frmScreen.txtEmail.value = ContactDetailsXML.GetAttribute("EMAILADDRESS");
		m_bDisplay = true;
	}

	
	 
}
function frmScreen.btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}
function frmScreen.btnOK.onclick()
{
  if (m_bDisplay)	 
	if (window.confirm("Update task as complete?"))
	 {
	  	taskXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		taskXML.LoadXML(m_sTaskXML);
		taskXML.SelectTag(null, "CASETASK");
	
		UpDateTaskXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		UpDateTaskXML.CreateRequestTagFromArray(m_sRequestAttributes, "UpdateCaseTask");
		UpDateTaskXML.CreateActiveTag("CASETASK");
	    UpDateTaskXML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
		UpDateTaskXML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
		UpDateTaskXML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
		UpDateTaskXML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
		UpDateTaskXML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
		UpDateTaskXML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
		UpDateTaskXML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
		UpDateTaskXML.SetAttribute("TASKSTATUS", "40");

		// 		UpDateTaskXML.RunASP(document,"MsgTMBO.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					UpDateTaskXML.RunASP(document,"MsgTMBO.asp");
				break;
			default: // Error
				UpDateTaskXML.SetErrorResponse();
			}


		var ErrorOwnerTypes = new Array("RECORDNOTFOUND");
		var ErrorOwnerReturn = UpDateTaskXML.CheckResponse(ErrorOwnerTypes);
		var sReturn = new Array();
		sReturn[0] = IsChanged();
		window.returnValue	= sReturn;
	 }
    
	window.close();
}
-->
</script>
</body>
</html>




