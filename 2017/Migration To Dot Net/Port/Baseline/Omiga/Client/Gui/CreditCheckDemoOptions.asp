<%@ Language=JScript %>
<HTML>
<%
/*
Workfile:      CreditCheckDemoOptions.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Options screen to allow the Credit Check Demo options
			   to be set/amended
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MDC		26/04/2000	Original
LD		23/05/02	SYS4727 Use cached versions of frame functions
JLD		12/06/02	sys4728 use stylesheet at all times
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
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Credit Check Demo Options</title>
</HEAD>

<BODY>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<form id="frmScreen">
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 170px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Credit Check Demo Options</strong>
	</span>

	<span style="TOP: 40px; LEFT: 50px; POSITION: ABSOLUTE" class="msgLabel">
		Demo Mode
		<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
			<input id="optDemoModeYes" name="DemoModeGroup" type="radio" value="1">
			<label for="optDemoModeYes" class="msgLabel">Yes</label>
		</span>
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="optDemoModeNo" name="DemoModeGroup" type="radio" value="0">
			<label for="optDemoModeNo" class="msgLabel">No</label>
		</span>
	</span>

	<span style="TOP: 70px; LEFT: 50px; POSITION: ABSOLUTE" class="msgLabel">
		Application Number
		<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
			<input id="txtAppNum" maxlength="12" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 100px; LEFT: 50px; POSITION: ABSOLUTE" class="msgLabel">
		Application Fact Find Number
		<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
			<input id="txtAppFFNum" maxlength="2" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 130px; LEFT: 50px; POSITION: ABSOLUTE" class="msgLabel">
		Sequence Number
		<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
			<input id="txtSeqNum" maxlength="2" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/CreditCheckDemoOptionsAttribs.asp" -->

<script language="JScript">
<!--
var DemoOptionsXML = "";
var scScreenFunctions;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	Initialise();

	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function PopulateScreen()
{
	<% /* populate fields in form */ %>
	scScreenFunctions.SetRadioGroupValue(frmScreen, "DemoModeGroup", DemoOptionsXML.GetTagText("DEMOMODEIND"));
	frmScreen.txtAppNum.value = DemoOptionsXML.GetTagText("APPLICATIONNUMBER");
	frmScreen.txtAppFFNum.value = DemoOptionsXML.GetTagText("APPLICATIONFACTFINDNUMBER");
	frmScreen.txtSeqNum.value = DemoOptionsXML.GetTagText("SEQUENCENUMBER");
}

function GetDemoOptions()
{
	<% /* Retrieve Credit Check data */ %>
	DemoOptionsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	with (DemoOptionsXML)
	{
		<% /* Create request block */ %>
		// CreateRequestTag(window, "SEARCH");
		CreateActiveTag("REQUEST");
		CreateActiveTag("SEARCH");

		<% /* Run the corresponding ASP page script to retrieve the data */ %>
		RunASP(document, "GetDemoOptions.asp");
		<% /* Check response */ %>
		if (IsResponseOK())
		{
			PopulateScreen();
		}
		else
			DisableMainButton("Submit");
	}
}
function Initialise()
{
	GetDemoOptions();
}

function btnSubmit.onclick()
{
	if(CommitScreen())			
		alert("Credit Check Demo Options have been updated successfully.");
			
}

function CommitScreen()
{
<%	/*	XML structure:
		<REQUEST USERID=?,USERTYPE=?,UNIT=?,ACTION="UPDATE">
			Demo Options
		<REQUEST>
	*/
%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	// var TagREQUEST = XML.CreateRequestTag(window,null);
	// var TagRequestType = null;
	// var TagREQUEST = null;

	// TagRequestType = XML.CreateActiveTag("UPDATE");
	<%	/*	XML structure:
			<DEMOOPTIONS>
				fields
			</DEMOOPTIONS>
		*/
	%>	
		XML.CreateActiveTag("REQUEST")
		XML.CreateActiveTag("UPDATE")
		// XML.ActiveTag = TagRequestType;
		XML.CreateActiveTag("DEMOOPTIONS");
		XML.CreateTag("APPLICATIONNUMBER", frmScreen.txtAppNum.value);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", frmScreen.txtAppFFNum.value);
		XML.CreateTag("SEQUENCENUMBER", frmScreen.txtSeqNum.value);
		
		if(frmScreen.optDemoModeYes.checked)
			XML.CreateTag("DEMOMODEIND", "1");
		else if(frmScreen.optDemoModeNo.checked)
			XML.CreateTag("DEMOMODEIND", "0");
		else
			XML.CreateTag("DEMOMODEIND", "");

		// 		XML.RunASP(document,"UpdateDemoOptions.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateDemoOptions.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		if(XML.IsResponseOK())
			return true;
		else
			return false;

}
-->
</script>

</BODY>
</HTML>




