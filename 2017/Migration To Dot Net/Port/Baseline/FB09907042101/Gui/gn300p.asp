<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      gn300.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Completeness Check screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		17/05/00	Initial revision
MH      31/05/00    Added Memopad screen and finished off.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MC		19/04/2004	BMIDS517	MemoPad dialog height incr. by 10px
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<title>Completeness Check <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<span id="spnRulesScroll" style="TOP: 388px; LEFT: 310px; POSITION: absolute; VISIBILITY: hidden">
	<object data="scTableListScroll.asp" id="scRulesScroll" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Screen Layout Here */ %>
<div id="divStatus" style="TOP: 10px; LEFT: 10px; WIDTH: 604px; HEIGHT: 100px; POSITION: absolute" class="msgGroup">
	<table id="tblStatus" width="604px" height="100px" border="0" cellspacing="0" cellpadding="0">
		<tr align="center"><td id="colStatus" align="center" class="msgLabelWait" >Please Wait...</td></tr>
	</table>
</div>

<div id="divBackground" style="TOP: 0px; LEFT: 10px; HEIGHT: 450px; WIDTH: 604px; POSITION: ABSOLUTE; VISIBILITY: HIDDEN" class="msgGroup">
	<form id="frmScreen" style="VISIBILITY: hidden">
		<div id="divRules" style="TOP: 4px; LEFT: 4px; WIDTH: 604px; POSITION: ABSOLUTE">
			<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
				<strong>Completion Rules</strong>
			</span>

			<span id="spnRulesTable" style="TOP: 16px; LEFT: 0px; POSITION: ABSOLUTE">
				<table id="tblRules" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
					<tr id="rowTitles">	<td width="12%" class="TableHead">ID</td>			<td class="TableHead">Description</td> <td width="12%" class="TableHead">Cost Modelling?</td></tr>
					<tr id="row01">		<td width="12%" class="TableTopLeft">&nbsp</td>		<td class="TableTopCenter">&nbsp</td> <td width="12%" class="TableTopRight">&nbsp</td></tr>
					<tr id="row02">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row03">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row04">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row05">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row06">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row07">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row08">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row09">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row10">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row11">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row12">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row13">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row14">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row15">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row16">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row17">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row18">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row19">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row20">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row21">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row22">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row23">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
					<tr id="row24">		<td width="12%" class="TableBottomLeft">&nbsp</td>	<td class="TableBottomCenter">&nbsp</td> <td width="12%" class="TableBottomRight">&nbsp</td></tr>
				</table>
			</span>
		</div>

		<div id="divButtons" style="TOP: 455px; LEFT: 4px; WIDTH: 604px; POSITION: ABSOLUTE; VISIBILITY: hidden">
			<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
				<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
			</span>

			<span style="TOP: 0px; LEFT: 66px; POSITION: ABSOLUTE">
				<input id="btnMemoPad" value="Memo Pad" type="button" style="WIDTH: 60px" class="msgButton">
			</span>

	<% /* Not used in cost modelling. Not yet supported. Maybe one day
			<span style="TOP: 0px; LEFT: 132px; POSITION: ABSOLUTE">
				<input id="btnPrint" value="Print" type="button" style="WIDTH: 60px" class="msgButton">
			</span>
			<span style="TOP: 0px; LEFT: 198px; POSITION: ABSOLUTE">
				<input id="btnDetails" value="Details" type="button" style="WIDTH: 60px" class="msgButton">
			</span>
	 */ %>
		</div>
	</form>
</div>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/gn300pAttribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_bIsCostModelling = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sApplicationStageArray = null;
var m_sUserId = null;
var m_sActivityId = null;
var m_bReadOnly = null;
var m_bApplicationReadOnly = null;
var m_sCompletenessTaskId = null;
var m_sStageId = null;
var m_sRequestArray = null;
var scScreenFunctions;
var ListXML = null;
var m_BaseNonPopupWindow = null;


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
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_bIsCostModelling = sArgArray[0];
	m_sApplicationNumber = sArgArray[1];
	m_sApplicationFactFindNumber = sArgArray[2];
	m_sApplicationStageArray = sArgArray[3];
	m_sUserId = sArgArray[4];
	m_sRequestArray = sArgArray[5];
	m_sActivityId = "1"; // temporary hard coded value. This should be passed in from calling screen
	m_bReadOnly = false;	// temporary hard coded value. This should be passed in from calling screen
	m_bApplicationReadOnly = false;	// temporary hard coded value. This should be passed in from calling screen
	m_sCompletenessTaskId = "1";	// temporary hard coded value. This should be passed in from calling screen
	m_sStageId = "1";	// temporary hard coded value. This should be passed in from calling screen
	
	var sText = "Please wait ... Running Completeness Check for";
	if(m_bIsCostModelling) sText += "Cost Modelling";
	else sText += "Application";
	colStatus.innerText = sText;
	
	window.setTimeout("CompleteInitialisation()",100);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}	
	
function CompleteInitialisation()	
{
	// Get data and populate the tables
	RunCompletenessCheck();
	// GetStageAndTasks();
	// GetConditions();

	scScreenFunctions.ShowCollection(frmScreen);
	
	divStatus.style.visibility = "hidden";
	divBackground.style.visibility = "visible";
	spnRulesScroll.style.visibility = "visible";

	divButtons.style.visibility = "visible";

	window.returnValue = null;
}

function RunCompletenessCheck()
{
	ListXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTagFromArray(m_sRequestArray,"SEARCH");
	ListXML.CreateActiveTag("COMPLETENESSCHECK");
	ListXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	ListXML.RunASP(document,"RunCompletenessCheck.asp");
	if(ListXML.IsResponseOK())
	{
		scRulesScroll.initialiseTable(tblRules,0,"",PopulateRulesTable,24,PopulateRulesTable(0));
	}
}

function PopulateRulesTable(nStart)
{
	var strText = "";
	
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("COMPLETIONRULE");
			
	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) && nLoop < 24; nLoop++)
	{								
		scScreenFunctions.SizeTextToField(tblRules.rows(nLoop + 1).cells(0), ListXML.GetAttribute("RULEID"));
		scScreenFunctions.SizeTextToField(tblRules.rows(nLoop + 1).cells(1), ListXML.GetAttribute("DESCRIPTION"));
		if(ListXML.GetAttribute("COSTMODELLINGIND") == "Y")
			strText = "Yes";
		else
			strText = "";
		scScreenFunctions.SizeTextToField(tblRules.rows(nLoop + 1).cells(2), strText);
	}

	return ListXML.ActiveTagList.length;
}

function frmScreen.btnOK.onclick()
{
	window.close();
}

function frmScreen.btnMemoPad.onclick()
{
	var ArrayArguments=new Array(5);
	
	ArrayArguments[0] = m_sApplicationNumber;
	ArrayArguments[1] = m_sApplicationFactFindNumber;
	ArrayArguments[2] = m_sUserId;
	ArrayArguments[3] = "GN300";
	ArrayArguments[4] = frmScreen.title;
	ArrayArguments[5] = m_sRequestArray;

	scScreenFunctions.DisplayPopup(window, document,"GN100.asp", ArrayArguments, 630, 485);	

}

-->
</script>
</body>
</html>




