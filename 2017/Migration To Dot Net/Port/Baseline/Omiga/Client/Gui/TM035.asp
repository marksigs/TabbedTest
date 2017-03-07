<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      TM035.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Task History  POP-UP
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		30/10/00	Created (screen paint)
JLD		16/11/00	implementation added (not complete)
DRC     07/03/01    SYS1787 Ownership history added and completed
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
AW		27/09/06	   EP1160	Use case task name for editable tasks.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Task History  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 160px; LEFT: 210px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTableStatus" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>	
<span style="TOP: 295px; LEFT: 210px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTableOwner" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 310px; WIDTH: 520px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Task
	<span style="TOP:-3px; LEFT:50px; POSITION:ABSOLUTE">
		<input id="txtTask" maxlength="100" style="WIDTH:400px" class="msgTxt">
	</span>
</span> 
<span style="TOP:45px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
    Status History:
</span>   
<span id="spnStatusTable" style="TOP: 60px; LEFT: 4px; POSITION: ABSOLUTE">

	<table id="tblStatusTable" width="496px" border="0" cellspacing="0" cellpadding="0" class="msgTable">
	<tr id="rowTitles">
		<td width="20%" class="TableHead">Status</td>	
		<td width="27%" class="TableHead">Date Set</td>	
		<td width="33%" class="TableHead">Actioning User</td>		
		<td width="20%" class="TableHead">Unit</td>
	</tr>
	<tr id="row01">		
		<td class="TableTopLeft">&nbsp;</td>		
		<td class="TableTopCenter">&nbsp;</td>		
		<td class="TableTopCenter">&nbsp;</td>		
		<td class="TableTopRight">&nbsp;</td>
	</tr>
	<tr id="row02">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row03">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row04">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row05">		
		<td class="TableBottomLeft">&nbsp;</td>	
		<td class="TableBottomCenter">&nbsp;</td>	
		<td class="TableBottomCenter">&nbsp;</td>	
		<td class="TableBottomRight">&nbsp;</td>
	</tr>

	</table>
</span>
<span style="TOP:180px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
    Ownership History:
</span>
<span id="spnOwnerTable" style="TOP: 195px; LEFT: 4px; POSITION:RELATIVE">
	<table id="tblOwnerTable" width="496px" border="0" cellspacing="0" cellpadding="0" class="msgTable">
	<tr id="rowTitles">
		<td width="10%" class="TableHead">No.</td>	
		<td width="20%" class="TableHead">Date</td>	
		<td width="25%" class="TableHead">Owning User</td>		
		<td width="20%" class="TableHead">Unit</td>
		<td width="25%" class="TableHead">Allocated by</td>
	</tr>
	<tr id="row01">		
		<td class="TableTopLeft">&nbsp;</td>		
		<td class="TableTopCenter">&nbsp;</td>
		<td class="TableTopCenter">&nbsp;</td>		
		<td class="TableTopCenter">&nbsp;</td>		
		<td class="TableTopRight">&nbsp;</td>
	</tr>
	<tr id="row02">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>	
	    <td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row03">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>	
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row04">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>	
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row05">		
		<td class="TableBottomLeft">&nbsp;</td>	
		<td class="TableBottomCenter">&nbsp;</td>	
		<td class="TableBottomCenter">&nbsp;</td>	
		<td class="TableBottomCenter">&nbsp;</td>	
		<td class="TableBottomRight">&nbsp;</td>
	</tr>
	</table>
</span>
<span style="TOP:320px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnOK" value="OK" type="button" style="WIDTH:60px" class="msgButton">
</span>
</div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/TM035Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_sTaskXML = "";
var m_sRequestAttributes = "";
var taskXML = null;
var historyXML = null;
var m_iTableLength = 5; // Both tables to have 5 rows
var TaskStatusXML = null;
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
	else scScreenFunctions.SetFieldState(frmScreen, "txtTask", "R");
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
	// Get the TaskStatus combo information
	TaskStatusXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("TaskStatus");
	TaskStatusXML.GetComboLists(document, sGroups);
	
	taskXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	taskXML.LoadXML(m_sTaskXML);
	taskXML.SelectTag(null, "CASETASK");
	
	<% /* AW   27/09/06 EP1160 */ %>
	if( taskXML.GetAttribute("EDITABLETASKIND") == "1" )
	{
		frmScreen.txtTask.value = taskXML.GetAttribute("CASETASKNAME");
	}
	else
	{
		frmScreen.txtTask.value = taskXML.GetAttribute("TASKNAME");
	}
	
    StatushistoryXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	StatushistoryXML.CreateRequestTagFromArray(m_sRequestAttributes, "FindTaskStatusList");
	StatushistoryXML.CreateActiveTag("TASKSTATUS");
	StatushistoryXML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
	StatushistoryXML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
	StatushistoryXML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
	StatushistoryXML.SetAttribute("ACTIVITYINSTANCE", taskXML.GetAttribute("ACTIVITYINSTANCE"));
	StatushistoryXML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
	StatushistoryXML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
	StatushistoryXML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
	StatushistoryXML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
	StatushistoryXML.RunASP(document,"MsgTMBO.asp");

	var ErrorStatusTypes = new Array("RECORDNOTFOUND");
	var ErrorStatusReturn = StatushistoryXML.CheckResponse(ErrorStatusTypes);
	if ((ErrorStatusReturn[0] == true) || (ErrorStatusReturn[1] == ErrorStatusTypes[0]))
	{
		if(ErrorStatusReturn[1] == ErrorStatusTypes[0])
			alert("No Status history found for this task.");
		else PopulateStatusListBox();
	}
	
	OwnerhistoryXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	OwnerhistoryXML.CreateRequestTagFromArray(m_sRequestAttributes, "FindTaskOwnershipList");
	OwnerhistoryXML.CreateActiveTag("TASKOWNERSHIPHISTORY");
	OwnerhistoryXML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
	OwnerhistoryXML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
	OwnerhistoryXML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
	OwnerhistoryXML.SetAttribute("ACTIVITYINSTANCE", taskXML.GetAttribute("ACTIVITYINSTANCE"));
	OwnerhistoryXML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
	OwnerhistoryXML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
	OwnerhistoryXML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
	OwnerhistoryXML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
    OwnerhistoryXML.RunASP(document,"MsgTMBO.asp");

	var ErrorOwnerTypes = new Array("RECORDNOTFOUND");
	var ErrorOwnerReturn = OwnerhistoryXML.CheckResponse(ErrorOwnerTypes);
	if ((ErrorOwnerReturn[0] == true) || (ErrorOwnerReturn[1] == ErrorOwnerTypes[0]))
	{
		if(ErrorOwnerReturn[1] == ErrorOwnerTypes[0])
			alert("No Owner History found for this task.");
		else PopulateOwnerListBox();
	}		
}
function PopulateStatusListBox()
{
	StatushistoryXML.ActiveTag = null;
	StatushistoryXML.CreateTagList("TASKSTATUSHISTORY");
	var iNumberOfTasks = StatushistoryXML.ActiveTagList.length;

	scScrollTableStatus.initialiseTable(tblStatusTable, 0, "", ShowStatusList, m_iTableLength, iNumberOfTasks);
	ShowStatusList(0);	
}
function ShowStatusList(nStart)
{
	var iCount;
	scScrollTableStatus.clear();
	for (iCount = 0; iCount < StatushistoryXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		StatushistoryXML.SelectTagListItem(iCount + nStart);
		scScreenFunctions.SizeTextToField(tblStatusTable.rows(iCount+1).cells(0),GetComboText(StatushistoryXML.GetAttribute("TASKSTATUS")));
		scScreenFunctions.SizeTextToField(tblStatusTable.rows(iCount+1).cells(1),StatushistoryXML.GetAttribute("TASKSTATUSSETDATETIME") );
		scScreenFunctions.SizeTextToField(tblStatusTable.rows(iCount+1).cells(2),StatushistoryXML.GetAttribute("TASKSTATUSSETBYUSERID") );
		scScreenFunctions.SizeTextToField(tblStatusTable.rows(iCount+1).cells(3),StatushistoryXML.GetAttribute("TASKSTATUSSETBYUNITID") );
	}
}
function PopulateOwnerListBox()
{
	OwnerhistoryXML.ActiveTag = null;
	OwnerhistoryXML.CreateTagList("TASKOWNERSHIPHISTORY");
	var iNumberOfTasks = OwnerhistoryXML.ActiveTagList.length;
 
	scScrollTableOwner.initialiseTable(tblOwnerTable, 0, "", ShowOwnerList, m_iTableLength, iNumberOfTasks);
	ShowOwnerList(0);	
}
function ShowOwnerList(nStart)
{
	var iCount;
	scScrollTableOwner.clear();
	for (iCount = 0; iCount < OwnerhistoryXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		OwnerhistoryXML.SelectTagListItem(iCount + nStart);
		scScreenFunctions.SizeTextToField(tblOwnerTable.rows(iCount+1).cells(0),OwnerhistoryXML.GetAttribute("TASKOWNERSEQUENCENO"));
		scScreenFunctions.SizeTextToField(tblOwnerTable.rows(iCount+1).cells(1),OwnerhistoryXML.GetAttribute("DATEOFOWNERSHIP") );
		scScreenFunctions.SizeTextToField(tblOwnerTable.rows(iCount+1).cells(2),OwnerhistoryXML.GetAttribute("OWNINGUSERID") );
		scScreenFunctions.SizeTextToField(tblOwnerTable.rows(iCount+1).cells(3),OwnerhistoryXML.GetAttribute("OWNINGUNITID") );
		scScreenFunctions.SizeTextToField(tblOwnerTable.rows(iCount+1).cells(4),OwnerhistoryXML.GetAttribute("ALLOCATEDBYUSERID") );
	}
}
function GetComboText(sTaskStatusValue)
{
<%	// return the valuename from the TaskStatus combo for the valueid sTaskStatusValue
%>	TaskStatusXML.SelectTag(null, "LISTNAME");
	TaskStatusXML.CreateTagList("LISTENTRY");
	var sValueName = "";
	for(var iCount = 0; iCount < TaskStatusXML.ActiveTagList.length && sValueName == ""; iCount++)
	{
		TaskStatusXML.SelectTagListItem(iCount);
		if(TaskStatusXML.GetTagText("VALUEID") == sTaskStatusValue)
			sValueName = TaskStatusXML.GetTagText("VALUENAME");
	}
	return sValueName;
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




