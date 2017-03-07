<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      TM037.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Stage History 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		17/11/00	Created
JLD		20/11/00	implementation added (not complete)
JLD		24/11/00	Added BO calls.
JLD		09/01/01	SYS1799 Add a friendly message if there are no previous stages to show.
CL		13/03/01	SYS2037 Screen converted to NON popup + Table length change
CL		19/03/01	SYS2037	Change to form calling
ADP		23/03/01	SYS2037	Change to screen layout
SG		11/12/01	SYS3346/3347 Use CaseTaskName instead of TaskName in Stage History listbox.
SG		12/12/01	SYS3346/3105 Ensure correct task is selected when clickign DETAILS
SG		12/12/01	SYS3346/3349 Correct typo in RouteToInputProcess
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
ASu		12/09/2002	BMIDS00424 - Disable <Details> & <Memo> buttons where task status is not set to complete.
ASu		02/10/2002	BMIDS00424 - Enable <Memo> button for all instances.
ASu		03/10/2002	BMIDS00426 - Disable <Details> button where task is 'Completeness Check'(It is causing update errors because it runs as new task) 
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
BS		11/06/2003	BM0521	Enable Stage combo when case locked
HMA     27/09/2004  BMIDS879   - Correct error when no row selected.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:
Prog	Date		Description
HMA     07/07/2005  MAR11      - Correct scrolling/sorting when large number of tasks present (BMIDS1017)
PJO     21/12/2005  MAR882     To force code change promotion for MAR11 - no ther changes made
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom History:

Prog	Date		Description
PE		13/07/2006	EP990 - Change "Due/complete Date" from TASKDUEDATEANDTIME to TASKSTATUSSETDATETIME
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<title></title>
</head>
<STRONG></STRONG>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 400px; LEFT: 310px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>


<% /* Specify Forms Here */ %>

<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM036" method="post" action="TM036.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP010" method="post" action="AP010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP100" method="post" action="AP100.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP110" method="post" action="AP110.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP120" method="post" action="AP120.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP130" method="post" action="AP130.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP140" method="post" action="AP140.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP150" method="post" action="AP150.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP160" method="post" action="AP160.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP205" method="post" action="AP205.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP250" method="post" action="AP250.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP300" method="post" action="AP300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP400" method="post" action="AP400.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>


<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: Hidden" mark validate="onchange" year4>
	<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 380px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	
		<span style="TOP:10px; LEFT:40px; POSITION:ABSOLUTE" class="msgLabel">
			Stage
			<span style="TOP:-3px; LEFT:50px; POSITION:ABSOLUTE">
				<select id="cboStage" style="WIDTH:200px" class="msgCombo"></select>
			</span>
		</span>
		<span style="TOP:10px; LEFT:350px; POSITION:ABSOLUTE" class="msgLabel">
			Sort Tasks by
			<span style="TOP:-3px; LEFT:90px; POSITION:ABSOLUTE">
				<select id="cboOrderBy" style="WIDTH:100px" class="msgCombo"></select>
			</span>
		</span>
		<span id="spnTable" style="TOP: 35px; LEFT: 15px; POSITION: ABSOLUTE">
			<table id="tblTable" width="575px" border="0" cellspacing="0" cellpadding="0" class="msgTable">
				<tr id="rowTitles">
					<td width="40%" class="TableHead">Task</td>	
					<td width="20%" class="TableHead">Status</td>	
					<td width="25%" class="TableHead">Due/Completed Date</td>		
					<td width="15%" class="TableHead">Owner</td>
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
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row06">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row07">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row08">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row09">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row10">		
					<td class="TableLeft">&nbsp;</td>	
					<td class="TableCenter">&nbsp;</td>	
					<td class="TableCenter">&nbsp;</td>	
					<td class="TableRight">&nbsp;</td>
				</tr>
					<tr id="row11">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>	
					<td class="TableCenter">&nbsp;</td>	
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row12">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row13">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row14">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row15">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row16">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row17">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row18">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row19">		
					<td class="TableLeft">&nbsp;</td>		
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableCenter">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row20">		
					<td class="TableBottomLeft">&nbsp;</td>	
					<td class="TableBottomCenter">&nbsp;</td>	
					<td class="TableBottomCenter">&nbsp;</td>	
					<td class="TableBottomRight">&nbsp;</td>
				</tr>
			</table>
		</span>
	</div>
	<span style="TOP:400px; LEFT:25px; POSITION:ABSOLUTE">
		<input id="btnDetails" value="Details" type="button" style="WIDTH:80px" class="msgButton">
		<input id="btnMemo" value="Memo" type="button" style="WIDTH:80px" class="msgButton">
	</span>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 10px; POSITION: absolute; TOP: 450px; WIDTH: 612px">
<!-- #include FILE="msgButtons.asp" -->
</div>


<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/TM037Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--

var scScreenFunctions;
var m_sReadOnly = "";
var m_sRequestAttributes = "";
var m_sApplicationNumber = "";
var m_sActivityId = "";
var XML = null;
var taskStatusXML = null;
var m_iTableLength = 20;
var m_taskXML = null;
var XML = null;
var stageXML = null;



<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	frmScreen.btnDetails.disabled = true;
	frmScreen.btnMemo.disabled = true;
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	var sButtonList = new Array("Submit");
	var sScreenId = "TM037";
	FW030SetTitles("Stage History",sScreenId,scScreenFunctions);
	RetrieveContextData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	//scTable.initialise(tblTable,0,"");
	PopulateScreen();
	SetScreenOnReadOnly();
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function spnTable.onclick()
{
	<% /* BMIDS879 Get table index and check that a row is selected before continuing */ %>

	var iIndex = scScrollTable.getRowSelected();  // returns the table index 
	
	if (iIndex > 0)
	{
		var aSelectedRows = tblTable.rows(iIndex).getAttribute("InputProcess");
		var aTaskStatus = tblTable.rows(iIndex).getAttribute("TaskStatus");
		<% /* ASu - BMIDS00424 - Get the task id from the returned xml - Start */ %>	
		var aTaskId = tblTable.rows(iIndex).getAttribute("TaskId");
		<% /* ASu - End */ %>	
		<% /* ASu - BMIDS00424 - Add additional check for completed tasks & BMIDS00426 - Disable Details if Complete_Check */ %>
		if ((aSelectedRows != "") && (aTaskStatus == 40) && (aTaskId != "Complete_Check")) <%/* 40 = Task Complete */ %>
		{ 
			frmScreen.btnDetails.disabled = false;
			frmScreen.btnMemo.disabled = false;
		}	
		else
		{
			frmScreen.btnDetails.disabled = true;
			frmScreen.btnMemo.disabled = false;
		}	
	}
}


function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	<% /* BS BM0521 11/06/03
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	} */ %>
}


function RetrieveContextData()
{
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sActivityId = scScreenFunctions.GetContextParameter(window,"idActivityId","1");
}


function PopulateScreen()
{
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var XMLRequestTag = XML.CreateRequestTag(window,"FindArchiveStageList"); 
	XML.CreateActiveTag("CASEACTIVITY");
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("CASEID", m_sApplicationNumber);
	XML.SetAttribute("ACTIVITYID", m_sActivityId);
	XML.SetAttribute("ACTIVITYINSTANCE", "1");
	XML.RunASP(document, "MsgTMBO.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{
		SetScreenOnReadOnly();
		alert("There are no previous stages so no stage history can be shown.");
	}
	else if(ErrorReturn[0] == true)
	{
		//Add an ORDER tag to each CASETASK of each CASESTAGE for populating the table
		XML.CreateTagList("CASESTAGE");
		for(var iCount = 0 ; iCount < XML.ActiveTagList.length; iCount++)
		{
			XML.SelectTagListItem(iCount);
			XML.CreateTagList("CASETASK");		
		
			var iNoOfTasks = XML.ActiveTagList.length;
			if (iNoOfTasks > 0)
			{
				for(var iLoop = 0; iLoop < iNoOfTasks; iLoop++)
				{
					XML.SelectTagListItem(iLoop);
					XML.CreateTag("ORDER", iLoop+1);
				}
			}
			XML.ActiveTag = null;
			XML.CreateTagList("CASESTAGE");
		}
		PopulateCombos();
		PopulateListBox();
		//scTable.setRowSelected(1);
		//spnTable.onclick();
		
	}
}


function OrderByAddOption(sName)
{
	TagOPTION = document.createElement("OPTION");
	TagOPTION.value	= sName;
	TagOPTION.text = sName;
	frmScreen.cboOrderBy.add(TagOPTION);
}


function PopulateCombos()
{
	// Get the TaskStatus combo information
	TaskStatusXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("TaskStatus");
	TaskStatusXML.GetComboLists(document, sGroups);
	
	// order by
	OrderByAddOption("Task");
	OrderByAddOption("Status");
	OrderByAddOption("Date");
	OrderByAddOption("Owner");

	//Stage
	XML.CreateTagList("CASESTAGE");
	var iNoOfStages = XML.ActiveTagList.length;

	if (iNoOfStages > 0)
	{
		var dateDiff = -1;
		var SelectedIndex = -1;
		<% /* MO - BMIDS00807 */ %>
		<% /* var dtToday = scScreenFunctions.GetSystemDate(); */ %>
		var dtToday = scScreenFunctions.GetAppServerDate();
		
		for(var iLoop = 0; iLoop < iNoOfStages; iLoop++)
		{
			XML.SelectTagListItem(iLoop);
			var sStageName = XML.GetAttribute("STAGENAME");
			var dtStageCompletion = scScreenFunctions.GetDateObjectFromString(XML.GetAttribute("STAGECOMPLETIONDATETIME"));
			var diff = dtToday.valueOf() - dtStageCompletion.valueOf();
			if(dateDiff == -1 || diff < dateDiff)
			{
				dateDiff = diff;
				SelectedIndex = iLoop;
			}
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= sStageName;
			TagOPTION.text = sStageName;
			TagOPTION.setAttribute("StageId", XML.GetAttribute("STAGEID"));
			TagOPTION.setAttribute("TagListItem", iLoop);
			frmScreen.cboStage.add(TagOPTION);
		}
		frmScreen.cboStage.selectedIndex = SelectedIndex;
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


function PopulateListBox()
{
	XML.ActiveTag = null;
	XML.CreateTagList("CASESTAGE");
	XML.SelectTagListItem(frmScreen.cboStage.selectedIndex);
	XML.CreateTagList("CASETASK");
	var iNumberOfTasks = XML.ActiveTagList.length;

	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfTasks);
	ShowList(0);	
}


function ShowList(nStart)
{
<%  //This method populates the listbox based on the ORDER tag on each CASETASK. We must cater
	//for scrolling of the listbox too.
%>	var iCount;
	var iDisplayed = 0;

	scScrollTable.clear();		
	for (iCount = 0; iCount < XML.ActiveTagList.length && iDisplayed < m_iTableLength; iCount++)
	{
		XML.SelectTagListItem(iCount);
		var nOrder = parseInt(XML.GetTagText("ORDER"));
		if((nOrder - nStart) <= 0 | (nOrder - nStart) > m_iTableLength)
		{
			//ignore it
		}
		else
		{
			<% /* BMIDS1017  Increment the count of displayed rows. */ %>
			iDisplayed = iDisplayed + 1;

		<%	//SG 11/12/01 SYS3346/3347
			//scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(0),XML.GetAttribute("TASKNAME"));
		%>	scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(0),XML.GetAttribute("CASETASKNAME"));
			scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(1),GetComboText(XML.GetAttribute("TASKSTATUS")));
			
			<%
			// Peter Edney - 13/07/2006 - EP990
			// scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(2),XML.GetAttribute("TASKDUEDATEANDTIME"));
			%>
			scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(2),XML.GetAttribute("TASKSTATUSSETDATETIME"));						
			
			scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(3),XML.GetAttribute("OWNINGUSERID"));
			//Assign the INPUTPROCESS as an atribute of the table
			var varInputProcess = XML.GetAttribute("INPUTPROCESS");
			
			tblTable.rows(nOrder - nStart).setAttribute("InputProcess", varInputProcess);
			<% /* ASu - BMIDS00426 - Start */ %>
			tblTable.rows(nOrder - nStart).setAttribute("TaskId", XML.GetAttribute("TASKID"));
			<% /* ASu - BMIDS00426 - End */ %>
			<% /* ASu - BMIDS00424 - Start */ %>
			tblTable.rows(nOrder - nStart).setAttribute("TaskStatus", XML.GetAttribute("TASKSTATUS"));
			<% /* ASu - BMIDS00424 - End */ %>
		}
	}
}

function frmScreen.cboOrderBy.onchange()
{
	OrderTaskList();
}

function OrderTaskList()
{
<%	//re-order the listbox entries by re-setting the ORDER tag text on each of the casetask entries.
	//To do this, use the Array object Sort method. Fill the array with the attribute chosen in the
	//combo. Sort it alphabetically (or by ascending date order). Loop the XML again and reset
	//the order of each casetask based on it's value and position in the Array. Call PopulateTable.
%>	var aSortArray = new Array();
	var sSortBy = "";
	XML.ActiveTag = null;
	XML.CreateTagList("CASESTAGE");
	XML.SelectTagListItem(frmScreen.cboStage.selectedIndex);
	XML.CreateTagList("CASETASK");
	switch(frmScreen.cboOrderBy.value)
	{
		case "Task" : sSortBy = "TASKNAME"; break;
		case "Status" : sSortBy = "TASKSTATUS"; break;
		case "Date" : sSortBy = "TASKDUEDATEANDTIME"; break;
		case "Owner": sSortBy = "OWNINGUSERID"; break;
		default: sSortBy = "TASKNAME";
	}
	
	for (var iCount = 0; iCount < XML.ActiveTagList.length; iCount++)
	{
		XML.SelectTagListItem(iCount);
		aSortArray[iCount] = XML.GetAttribute(sSortBy);
	}
	
	if(sSortBy == "TASKDUEDATEANDTIME")
		aSortArray.sort(sortByDateAsc);
	else
		aSortArray.sort();
	
	for (var iCount = 0; iCount < XML.ActiveTagList.length; iCount++)
	{
		XML.SelectTagListItem(iCount);
		var sAttrValue = XML.GetAttribute(sSortBy);
		XML.SetTagText("ORDER", GetOrderFromArray(sAttrValue, aSortArray));
	}
	
	PopulateListBox();
}


function GetOrderFromArray(sValue, sArray)
{
<%	//loops the array to match the value passed in. Return the array position.
	//If there are identical values in the array we want to get each array position and not just
	//the first one we hit, so when we've found a value, alter the array value for that position
	//so that the sValue will not match that position again and will match the next one.
%>	var sOrder = "";
	for(var iCount = 0; iCount < sArray.length && sOrder == ""; iCount++)
	{
		if(sValue == sArray[iCount]) 
		{
			sOrder = (iCount+1);
			sArray[iCount] = sValue + "  ZZZ";
		}
	}
	return sOrder;
}


function sortByDateAsc(sArg1, sArg2)
{
	var sDate1 = new Date(sArg1.substr(6,4), (sArg1.substr(3,2) -1), sArg1.substr(0,2), sArg1.substr(11,2), sArg1.substr(14,2), sArg1.substr(17,2));
	var sDate2 = new Date(sArg2.substr(6,4), (sArg2.substr(3,2) -1), sArg2.substr(0,2), sArg2.substr(11,2), sArg2.substr(14,2), sArg2.substr(17,2));
	if(sDate1 < sDate2) return -1;
	else if(sDate1 > sDate2) return 1;
	else return 0;
}


function frmScreen.cboStage.onchange()
{
	<% /* BMIDS1017 Order task list and Populate List Box using value in OrderBy combo */ %>
	OrderTaskList();
}

function frmScreen.btnMemo.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idTaskXML",GetTasksXML());
	scScreenFunctions.SetContextParameter(window,"idReturnScreenId","TM037");
	frmToTM036.submit();
}


function RouteToInputProcess(sInputProcess)
{
	scScreenFunctions.SetContextParameter(window,"idReturnScreenId","TM030");
	switch(sInputProcess)
	{
		case "TM030": frmToTM030.submit(); break;
		case "TM036": frmToTM036.submit(); break;
		case "AP010": frmToAP010.submit(); break;
		case "AP100": frmToAP100.submit(); break;
		case "AP110": frmToAP110.submit(); break;
		case "AP120": frmToAP120.submit(); break;
		case "AP130": frmToAP130.submit(); break;
		case "AP140": frmToAP140.submit(); break;
		case "AP150": frmToAP150.submit(); break;
		case "AP160": frmToAP160.submit(); break;
		case "AP205": frmToAP205.submit(); break;
		case "AP250": frmToAP250.submit(); break;
		case "AP300": frmToAP300.submit(); break;
	<%	// SG 12/12/01 SYS3346/3349
		// AP400 and GN300 used to be routed to AP300 - typo?
	%>	case "AP400": frmToAP400.submit(); break;
		case "GN300": frmToGN300.submit(); break;
		default : alert("screen " + sInputProcess + " not found");
	}
}


function frmScreen.btnDetails.onclick()
{
	//Check that a row has been selected
	if (scScrollTable.getRowSelected() != null) ;	
	{
		frmScreen.btnDetails.disabled = false;
	<%	//SG 12/12/01 SYS3346/3105
		//GetTasksXML();
	%>	
		//get the CASETASK xml element for the selected rows
		scScreenFunctions.SetContextParameter(window,"idTaskXML",GetTasksXML());
		XML.SelectTag(null, "RESPONSE");
		XML.CreateTagList("CASESTAGE/CASETASK[@INPUTPROCESS]");
		var aFormToCall = tblTable.rows(scScrollTable.getRowSelected()).getAttribute("InputProcess");
		if(aFormToCall != "")
		{
			RouteToInputProcess(aFormToCall);
		}	
	else
		{
			frmScreen.btnDetails.disabled = true;
		}
	}
}

function GetTasksXML()
{
<%	//For each row of the table selected get the associated CASETASK xml and return them all
%>	var saRowArray = scScrollTable.getArrayofRowsSelected(); //THIS MIGHT ONLY BE THE VISIBLE ONES
	var SelectedXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var tagRequest = SelectedXML.CreateRequestTag(window , "");
	for(var iCount = 0; iCount < saRowArray.length; iCount++)
	{
		var nOrderNum = saRowArray[iCount];
		var newNode = SetActiveTask(nOrderNum);
		var node = newNode.cloneNode(true);
		SelectedXML.ActiveTag.appendChild(node);
	}
	return(SelectedXML.XMLDocument.xml);
}


function SetActiveTask(nOrderNum)
{
<%	//SG 12/12/01 SYS3346/3105 - The ActiveTag and TagList are already correctly set
	//XML.ActiveTag = null;
	//XML.CreateTagList("CASETASK");
%>	var node = null;
	for (iCount = 0; iCount < XML.ActiveTagList.length && node == null; iCount++)
	{
		XML.SelectTagListItem(iCount); //sets this tag as the active tag
		var nThisOrder = parseInt(XML.GetTagText("ORDER"));
		if(nThisOrder == nOrderNum)
			node = XML.GetTagListItem(iCount);
	}
	return node;
}

function btnSubmit.onclick()
{
	var sReturn = new Array();
	sReturn[0] = IsChanged();
	window.returnValue	= sReturn;
	frmToTM030.submit();
}


-->
</script>
</body>
</html>




