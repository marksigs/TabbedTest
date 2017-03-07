<%@ LANGUAGE="JSCRIPT" %>
<html>
<%

/*
Workfile:      pp200.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Application Enquiry Quick Search screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MV		30/10/00	Created
MV		14/11/00	Added Cancel Button 
					Added Code to Build XML string on btnSearch_Click event
ASm		22/11/00	CORE000012: Added Message box to inform user that 
					Application Number takes precedence
CL		05/03/01	SYS1920 Read only functionality added
				
LD		23/05/02	SYS4727 Use cached versions of frame functions
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
<form id="frmTogn215" method="post" action="pp215.asp" STYLE="DISPLAY: none"></form>
<form id="frmTogn210" method="post" action="pp210.asp" STYLE="DISPLAY: none"></form>
<form id="frmWelcomeMenu" method="post" action="mn010.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span> 
<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate ="onchange">

<% /* Application Details */ %>
<div style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 143px" class="msgGroup" id="DIV1">
<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel">
	Application Number
	<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
		<input id="txtApplicationNumber" maxlength="12" style="LEFT: 30px; WIDTH: 141px; POSITION: absolute; TOP: 0px; HEIGHT: 20px" class="msgTxt" size="14">
	</span>
</span>
<span style="LEFT: 287px; POSITION: absolute; TOP: 26px" class="msgLabel">
	Account Number
	<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
		<input id="txtAccountNumber" maxlength="20" style="LEFT: 30px; WIDTH: 150px; POSITION: absolute; TOP: -2px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 72px" class="msgLabel">
	Unit Name
	<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
		<input id="txtUnitName" maxlength="10" style="LEFT: 30px; WIDTH: 141px; POSITION: absolute; TOP: -2px; HEIGHT: 20px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 288px; POSITION: absolute; TOP: 70px" class="msgLabel">
	User Name
	<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">&nbsp;
		<input id="txtUserName" maxlength="10" style="LEFT: 30px; WIDTH: 150px; POSITION: absolute; TOP: -2px" class="msgTxt">
	</span>
</span>
<span style="LEFT: -4px; WIDTH: 260px; POSITION: absolute; TOP:  107px; HEIGHT: 34px" class="msgLabelHead" id=SPAN1>&nbsp; 
	<span>
		<input id="ChkCancelledorDeclinedApplications" style="WIDTH: 27px; HEIGHT: 20px" type="checkbox" value="on" For="ChkCancelledorDeclinedApplications"
		 <LABEL>Include Cancelled/Declined 
Applications </LABEL>
	</span>
</span>
<span style="LEFT: 350px; POSITION: absolute; TOP: 107px">
	<input id="btnSearch" value="Search" type="button" style="WIDTH: 75px" class="msgButton">
</span>
<span style="LEFT: 450px; POSITION: absolute; TOP: 107px">
	<input id="btnExtendedSearch" value="ExtendedSearch" type="button" style="WIDTH: 100px" class="msgButton">
</span>
</div> 
</form>
<%/* Main Buttons */ %>
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 210px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen  */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- File containing field attributes (remove if not required) --><!-- #include FILE="attribs/pp200attribs.asp"  --><!-- Specify Code here -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var scScreenFunctions;
var m_blnReadOnly = false;


function window.onload()
{
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Payment Processing Enquiry - Quick Search ","PP200",scScreenFunctions); 

	RetrieveContextData();
	//This function is contained in the field attributes file (remove if not required)
	SetMasks();
	Validation_Init();
	//scScreenFunctions.SetFieldState(frmScreen, "txtUnitName", ((frmScreen.txtApplicationNumber.value) == "") ? "R":"D");
	//scScreenFunctions.SetFieldState(frmScreen, "txtUserName", ((frmScreen.txtApplicationNumber.value) == "") ? "R":"D");
	scScreenFunctions.SetFieldState(frmScreen, "txtUnitName",  "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtUserName", "R");
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP200");
	
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
	//scScreenFunctions.SetContextParameter(window,"idXML","");
	var sXML  = scScreenFunctions.GetContextParameter(window,"idXML","");
	if (sXML != "" )
	{
		PopulateScreen(sXML);
	}
	//initialising  UnitName and UserName with Context Parameters
	frmScreen.txtUnitName.value =  scScreenFunctions.GetContextParameter(window,"idUnitName",null);
	frmScreen.txtUserName.value =  scScreenFunctions.GetContextParameter(window,"idUserName",null);
}

function frmScreen.btnSearch.onclick()
{	
	//save the combo settings in idXML context to use when we return
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	if ((frmScreen.txtApplicationNumber.value != "") &&
		(frmScreen.txtAccountNumber.value != "")) {
		alert ("Application Number and Account number have both been specified. Application Number will take precedence.");	
		frmScreen.txtAccountNumber.value = "";
	}
	
	//Preparing XML Request string 
	XML.CreateActiveTag("REQUEST");
	XML.SetAttribute("OPERATION", "QuickSearch");
	XML.SetAttribute("UnitId",scScreenFunctions.GetContextParameter(window,"idUnitId",null));
	XML.SetAttribute("UserId",scScreenFunctions.GetContextParameter(window,"idUserId",null));
	XML.SetAttribute("ChannelId",scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null));
	XML.SetAttribute("MachineId",scScreenFunctions.GetContextParameter(window,"idMachineId",null));
	XML.CreateTag("APPLICATIONNUMBER",frmScreen.txtApplicationNumber.value);
	XML.CreateTag("ACCOUNTNUMBER",frmScreen.txtAccountNumber.value);
	if (frmScreen.ChkCancelledorDeclinedApplications.checked == true) 
	{
		XML.CreateTag("INCLUDECANCELLEDAPPS","1");
		XML.CreateTag("INCLUDEDECLINEDAPPS","1");
	}
	else
	{
		XML.CreateTag("INCLUDECANCELLEDAPPS","0");
		XML.CreateTag("INCLUDEDECLINEDAPPS","0");
	}
	XML.CreateTag("CHANNELID",null);
	XML.CreateTag("DEPARTMENTID",null);
	XML.CreateTag("UNITNAME",frmScreen.txtUnitName.value);
	XML.CreateTag("USERNAME",frmScreen.txtUserName.value);
	XML.CreateTag("LENDERNAME",null);
	XML.CreateTag("PRODUCTNAME",null);
	XML.CreateTag("TYPEOFAPPLICATION",null);
	XML.CreateTag("FROMSTAGE",null);
	XML.CreateTag("TOSTAGE",null);
	XML.CreateTag("DATETO",null);
	XML.CreateTag("APPAPPROVEDMONTH",null);
	XML.CreateTag("APPAPPROVEDYEAR",null);
	XML.CreateTag("SOBDIRECTORINDIRECTt",null);
	XML.CreateTag("SOBCHANNELID",null);
	XML.CreateTag("SOBNAME",null);
	//alert(XML.ActiveTag.xml);
	scScreenFunctions.SetContextParameter(window,"idXML",XML.ActiveTag.xml);
	scScreenFunctions.SetContextParameter(window,"idMetaAction","QuickSearch");
	frmTogn215.submit();
}

function frmScreen.btnExtendedSearch.onclick()
{	
	// On Click Extended Search Button Navigate to GN210 Screen	
	alert("Under construction.");
	frmTogn210.submit();
}

function btnCancel.onclick()
{
	// On Click cancel button Navigate to Main Menu
	frmWelcomeMenu.submit();				
}

function PopulateScreen(sXML) 
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	XML.LoadXML(sXML);
	if (XML.SelectTag(null,"REQUEST") != null){
		frmScreen.txtApplicationNumber.value  = XML.GetTagText("APPLICATIONNUMBER");
		frmScreen.txtAccountNumber.value = XML.GetTagText("ACCOUNTNUMBER");
		frmScreen.txtUnitName.value = XML.GetTagText("UNITNAME");
		frmScreen.txtUserName.value = XML.GetTagText("USERNAME");
		if ( XML.GetTagText("INCLUDEDECLINEDAPPS")== "1" ) 	frmScreen.ChkCancelledorDeclinedApplications.checked = true; 
	}
	XML = null;	
}

-->
</script>
</body>
</html>
