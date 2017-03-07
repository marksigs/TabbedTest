<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      TM040.asp
Copyright:     Copyright © 2006 Marlborough Stirling

Description:   Application Progress Chart
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		AQR			Description
AW		14/09/2006	EP1103      CC78 New screen
AW		14/09/2006	EP1150      CC78 Default orderby to Date
									 New arguments passed to TM036P
AW		20/09/2006	EP1159      CC78 Increased height co-ordinate of TM036P
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title>Progress Chart  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<OBJECT id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp 
VIEWASTEXT></OBJECT>

<% /* Scriptlets - remove any which are not required */ %>
<OBJECT id=scTable style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 data=scTable.htm VIEWASTEXT></OBJECT>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="LEFT: 314px; POSITION: absolute; TOP: 455px">
<OBJECT id=scScrollTable style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" 
tabIndex=-1 type=text/x-scriptlet data=scTableListScroll.asp 
VIEWASTEXT></OBJECT>
</span>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>

<div id="divReorder" style="LEFT: 10px; WIDTH: 608px; POSITION: absolute; TOP: 5px; HEIGHT: 38px" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 11px" class="msgLabel">
	Current Stage
	<span style="LEFT: 75px; POSITION: absolute; TOP: -3px">
		<input id="txtCurrentStage" maxlength="50" style="WIDTH: 200px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 381px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Sort Tasks by
	<span style="LEFT: 74px; POSITION: absolute; TOP: -3px">
		<select id="cboOrderBy" style="WIDTH: 150px" class="msgCombo"></select>
	</span>
</span>
</div>

<div id="divTaskList" style="LEFT: 10px; WIDTH: 608px; POSITION: absolute; TOP: 50px; HEIGHT: 550px" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 0px" class="msgLabelHead">
	Task List
</span>
<div id="spnTable" style="LEFT: 4px; WIDTH: 601px; POSITION: absolute; TOP: 17px; HEIGHT: 400px">
	<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable" style="WIDTH: 596px; HEIGHT: 322px">
			<tr id="rowTitles">	
				<td width="80%" class="TableHead">Task</td>	
				<td width="10%" class="TableHead">Status</td>	
				<td width="10%" class="TableHead">Task Note?</td>
			</tr>
			<tr id="row01">		
				<td width="15%" class="TableTopLeft">&nbsp;</td>		
				<td width="70%" class="TableTopCenter">&nbsp;</td>		
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">		
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">		
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row05">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row06">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row07">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row08">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row09">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row10">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row11">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row12">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row13">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row14">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row15">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row16">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row17">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row18">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row19">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row20">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row21">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row22">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row23">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row124">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row25">
				<td width="15%" class="TableBottomLeft">&nbsp;</td>	
				<td width="70%" class="TableBottomCenter">&nbsp;</td>		
				<td class="TableBottomRight">&nbsp;</td>
			</tr> 
		</table>
</div>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 435px">
		<input id="btnAdd" value="Add Task" type="button" style="WIDTH: 80px" class="msgButton">
	</span>

	<% /* BM0493 MDC 03/04/2003 */ %>
	<span style="LEFT: 88px; POSITION: absolute; TOP: 435px">
		<input id="btnMemo" value="Memo" type="button" style="WIDTH: 80px" class="msgButton">
	</span>

	<% /* BM0493 MDC 03/04/2003 - End */ %>

</div>
<div id="divProcessTask" style="LEFT: 10px; WIDTH: 608px; POSITION: absolute; TOP: 530px; HEIGHT: 90px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 0px" class="msgLabelHead">
		Process Task
	</span>
	<div class="msgInput" style="LEFT: 4px; POSITION: absolute; TOP: 15px">
		<table id="tblTaskDetail" width="594" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr>
				<td width="80"><b>Task:</b></td><td width="168" id="tblTDName"></td>
				<td width="120"><b>Status:</b></td><td id="tblTDStatus"></td>
			</tr>
			<tr>
				<td><b>Due Date:</b></td><td id="tblTDDueDate"></td>
				<td><b>Last Actioned Date:</b></td><td id="tblTDActionedDate"></td>
			</tr>
			<tr>
				<td><b>Owner:</b></td><td id="tblTDOwner"></td>
				<td><b>Last Updated By:</b></td><td id="tblTDUpdatedBy"></td>
			</tr>
		</table>
	</div>
	<SPAN style="LEFT: 0px; POSITION: absolute; TOP: 88px">
			<INPUT class=msgButton id="btnOK" style="WIDTH: 60px" type=button value="OK" name="btnOK"> 
	</SPAN>
</div>
</form>

<script src="includes/Documents.js" language="javascript" type="text/javascript"></script> <% /* MAR7 GHun */ %>

<% /* Specify Code Here */ %>
<script language="JScript" type="text/javascript">
<!--

var m_sApplicationNumber = "";
var m_sApplicationPriority = "" ;
var m_sActivityId = "";
var m_sReadOnly = "";
var m_iUserRole = 0;
var m_sCurrentStageNum = "0";
var m_iTableLength = 25;
var m_sCurrStageId = "";
var scScreenFunctions;
var m_stageXML = null;
var TaskStatusXML = null;

var m_blnReadOnly = false;
var m_ParamXML = null;

var m_sReadOnlyAsCaseLocked = "";
var m_sProcessInd = "";

var scClientScreenFunctions;

var m_taskXML = null;

var AttribsXML = null;
var node = null;
var m_UserId = "";
var m_UnitId = ""; 
var m_MachineId  = "";
var m_DistributionChannelId = "";
var sTaskName = "";

var m_sRequestArray = null;
var m_BaseNonPopupWindow = null;
var m_sCustomerNameArray = null;
var m_sCustomerNumberArray = null;
var m_sCustomerVersionNumberArray = null;

function window.onload()
{	
	
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	m_sRequestArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	m_sApplicationNumber = m_sRequestArray[0];
	m_iUserRole = m_sRequestArray[8];
	m_sCustomerNameArray	= m_sRequestArray[13]; // Customer Name Array from context customers no. 1-5
	m_sCustomerNumberArray	= m_sRequestArray[14]; // Customer Number Array from context customers no. 1-5
	m_sCustomerVersionNumberArray = m_sRequestArray[15];// Customer Version Number Array from context customers no. 1-5
	m_sApplicationPriority = m_sRequestArray[16];
	
	PopulateScreen();
	
}


function PopulateScreen()
{
	scScreenFunctions.SetFieldState(frmScreen, "txtCurrentStage", "R");

	GetAllGlobalParameters();
	GetStageInfo();
	PopulateCombos();

	frmScreen.cboOrderBy.onchange()

}

function GetStageInfo()
{
	m_stageXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_stageXML.CreateRequestTagFromArray(m_sRequestArray[11], "GetProgressTasks");
	m_stageXML.CreateActiveTag("CASEACTIVITY");
	m_stageXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	m_stageXML.SetAttribute("CASEID", m_sApplicationNumber);
	m_stageXML.SetAttribute("ACTIVITYID", m_sActivityId);
	m_stageXML.SetAttribute("ACTIVITYINSTANCE", "1");

	m_stageXML.RunASP(document, "MsgTMBO.asp");

	var sActivitySequenceNo = "";
	if(m_stageXML.IsResponseOK())
	{
		m_stageXML.SelectTag(null, "CASESTAGE");
		m_sCurrStageId = m_stageXML.GetAttribute("STAGEID");

		scScreenFunctions.SetContextParameter(window,"idStageId", m_sCurrStageId);

		frmScreen.txtCurrentStage.value = m_stageXML.GetAttribute("STAGENAME");
		sActivitySequenceNo = m_stageXML.GetAttribute("ACTIVITYSEQUENCENO");
	
		//Add an ORDER tag to each CASETASK for populating the table
		//Add a SYMBOL tag also to store a symbol representing the tasks lateness.
		m_stageXML.CreateTagList("CASETASK");
		var iNoOfTasks = m_stageXML.ActiveTagList.length;

		if (iNoOfTasks > 0)
		{
			var iLoop;

			for(iLoop = 0; iLoop < iNoOfTasks; iLoop++)
			{
				m_stageXML.SelectTagListItem(iLoop);
				m_stageXML.CreateTag("ORDER", iLoop+1);
				
			}			
		}
	}
	return(sActivitySequenceNo);
}

function ResetCombos()
{
	var nLength = frmScreen.cboOrderBy.options.length;
	for(var iCount = nLength -1; iCount >= 0 ; iCount--)
	{
		frmScreen.cboOrderBy.options.remove(iCount);
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
	ResetCombos();
	// Get the TaskStatus combo information
	
	TaskStatusXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("TaskStatus");
	TaskStatusXML.GetComboLists(document, sGroups);

    // OrderBy
    <% /* AW 19/09/06	EP1150 - Default order by date */ %>
    OrderByAddOption("Date");
	OrderByAddOption("Status");
	OrderByAddOption("Task");
	OrderByAddOption("Owner");
}


function GetComboText(sTaskStatusValue)
{
<%	/* return the valuename from the TaskStatus combo for the valueid sTaskStatusValue */ %>	
	TaskStatusXML.SelectTag(null, "LISTNAME");
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

function PopulateTable()
{
	m_stageXML.ActiveTag = null;
	m_stageXML.CreateTagList("CASETASK");
	var iNumberOfTasks = m_stageXML.ActiveTagList.length;

	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfTasks);
	ShowList(0);
}

function GetComboValidationType(sTaskStatusValue)
{
<% /* Returns the ComboValidation code for a given Task Status */ %>	
	TaskStatusXML.SelectTag(null, "LISTNAME");
	TaskStatusXML.CreateTagList("LISTENTRY");
	var sValueName = "";
	for(var iCount = 0; iCount < TaskStatusXML.ActiveTagList.length && sValueName == ""; iCount++)
	{
		TaskStatusXML.SelectTagListItem(iCount);
		if(TaskStatusXML.GetTagText("VALUEID") == sTaskStatusValue)
			sValueName = TaskStatusXML.GetTagText("VALIDATIONTYPE");
	}
	return sValueName;
}

function ShowList(nStart)
{
<%  //This method populates the listbox based on the ORDER tag on each CASETASK. We must cater
	//for scrolling of the listbox too.
%>	var iCount;
	scScrollTable.clear();	
	
	for (iCount = 0; iCount < m_stageXML.ActiveTagList.length; iCount++)
	{
		m_stageXML.SelectTagListItem(iCount);
		var nOrder = parseInt(m_stageXML.GetTagText("ORDER"));
		if((nOrder - nStart) <= 0 | (nOrder - nStart) > m_iTableLength)
		{
			//ignore it
		}
		else
		{

			sTaskName = m_stageXML.GetAttribute("CASETASKNAME");
			if(sTaskName.length == 0) sTaskName = m_stageXML.GetAttribute("TASKNAME");
			
			var oRow = document.getElementById('tblTable').rows(nOrder - nStart).cells(0);
 			
			scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(0),sTaskName);
			
			scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(1),GetComboText(m_stageXML.GetAttribute("TASKSTATUS")) );
						
			if(m_stageXML.GetAttribute("TASKNOTESIND") != null && m_stageXML.GetAttribute("TASKNOTESIND") == "1")
				tblTable.rows(nOrder - nStart).cells(2).innerHTML="<img src=\"images/arrow_up.gif\" height=\"11px\" width=\"11px\">";
			
			tblTable.rows(nOrder - nStart).setAttribute("TaskId",m_stageXML.GetAttribute("TASKID"));

		}
	}
}

function GetMandatory(sIndicator)
{
	var sMandatory = "";
	if(sIndicator == "1") sMandatory = "Yes";
	return sMandatory;
}

function frmScreen.cboOrderBy.onchange()
{
<%	//re-order the listbox entries by re-setting the ORDER tag text on each of the casetask entries.
	//To do this, use the Array object Sort method. Fill the array with the attribute chosen in the
	//combo. Sort it alphabetically (or by ascending date order). Loop the m_stageXML again and reset
	//the order of each casetask based on it's value and position in the Array. Call PopulateTable.
%>	var aSortArray = new Array();
	var sSortBy = "";
	m_stageXML.ActiveTag = null;
	m_stageXML.CreateTagList("CASETASK");
	switch(frmScreen.cboOrderBy.value)
	{
		case "Task" : sSortBy = "TASKNAME"; break;
		case "Status" : sSortBy = "TASKSTATUS"; break;
		case "Mandatory" : sSortBy = "MANDATORYINDICATOR"; break;
		case "Date" : sSortBy = "TASKDUEDATEANDTIME"; break;
		case "Owner": sSortBy = "OWNINGUSERID"; break;
		default: sSortBy = "TASKSTATUS";
	}
	
	for (var iCount = 0; iCount < m_stageXML.ActiveTagList.length; iCount++)
	{
		m_stageXML.SelectTagListItem(iCount);
		if(sSortBy == "SYMBOL")
			aSortArray[iCount] = m_stageXML.GetTagText("SYMBOL");
		else
		{
			if(m_stageXML.GetAttribute(sSortBy) == null)
				aSortArray[iCount] = " ";
			else
				aSortArray[iCount] = m_stageXML.GetAttribute(sSortBy);
		}
	}
	
	if(sSortBy == "TASKDUEDATEANDTIME")
		aSortArray.sort(sortByDateAsc);
	else if(sSortBy == "MANDATORYINDICATOR") 
		aSortArray.sort(sortByAlphaDesc);
	else if(sSortBy == "SYMBOL")
		aSortArray.sort(sortBySymbol);
	else
		aSortArray.sort();
	
	for (var iCount = 0; iCount < m_stageXML.ActiveTagList.length; iCount++)
	{
		m_stageXML.SelectTagListItem(iCount);
		var sAttrValue;
		if(sSortBy == "SYMBOL")
			sAttrValue = m_stageXML.GetTagText("SYMBOL");
		else
		{
			sAttrValue = m_stageXML.GetAttribute(sSortBy);
			if( sAttrValue == null) sAttrValue = " ";
		}
		m_stageXML.SetTagText("ORDER", GetOrderFromArray(sAttrValue, aSortArray));
	}
	
	PopulateTable();

	disableProcessButtons();

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
	if(sArg1 == "" || sArg2 == "")
	{
		return -1;
	}
	var sDate1 = new Date(sArg1.substr(6,4), (sArg1.substr(3,2) -1), sArg1.substr(0,2), sArg1.substr(11,2), sArg1.substr(14,2), sArg1.substr(17,2));
	var sDate2 = new Date(sArg2.substr(6,4), (sArg2.substr(3,2) -1), sArg2.substr(0,2), sArg2.substr(11,2), sArg2.substr(14,2), sArg2.substr(17,2));
	if(sDate1 < sDate2)
	{
		return -1;
	}
	else if(sDate1 > sDate2) 
	{
		return 1;
	}
	return 0;
}

function sortByAlphaDesc(sArg1, sArg2)
{
	if(sArg2 < sArg1) 
	{
		return -1;
	}
	else if(sArg2 > sArg1) 
	{
		return 1;
	}
	return 0;
}

function sortBySymbol(sArg1, sArg2)
{
<%	// order as "!!", "*", ""
%>	var nNumArg1;
	var nNumArg2;
	switch(sArg1)
	{
		case "!!": nNumArg1 = 1; break;
		case "*": nNumArg1 = 2; break;
		default: nNumArg1 = 3; break;
	}
	switch(sArg2)
	{
		case "!!": nNumArg2 = 1; break;
		case "*": nNumArg2 = 2; break;
		default: nNumArg2 = 3; break;
	}
	
	if(nNumArg1 < nNumArg2) 
	{
		return -1;
	}
	else if(nNumArg1 > nNumArg2) 
	{
		return 1;
	}
	return 0;
}

function spnTable.onclick()
{
	var saRowArray = scScrollTable.getArrayofRowsSelected();
	if(saRowArray.length > 0)
	{

		var sXML = GetTasksXML();
		m_taskXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_taskXML.LoadXML(sXML);
		scScreenFunctions.SetContextParameter(window,"idTaskXML",sXML);
		initTaskDetail();
		<% /*  If ReadOnly because it's a non-processing unit or the case is locked then enable 
		the Details and Memo buttons, otherwise call enableProcessButtons */ %>
		if((m_sReadOnlyAsCaseLocked == "1") || (m_sProcessInd == "0"))
		{
			frmScreen.btnMemo.disabled = false;
		}
		else enableProcessButtons();

	}
}

function frmScreen.btnAdd.onclick()
{

	var sReturn = null;
	m_stageXML.SelectTag(null, "CASESTAGE");
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var bUnderwritingTasksOnly = false;
	var bProgressTaskMode = true;
	
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sReadOnly;
	ArrayArguments[1] = m_sCurrStageId;
	ArrayArguments[2] = m_sApplicationNumber;
	ArrayArguments[3] = m_sApplicationPriority ;
	ArrayArguments[4] = m_iUserRole ;
	ArrayArguments[5] = m_sActivityId;
	ArrayArguments[6] = GetCustomerXML();
	ArrayArguments[7] = m_sRequestArray[11];
	ArrayArguments[8] = m_stageXML.GetAttribute("CASESTAGESEQUENCENO");
	ArrayArguments[9] = bUnderwritingTasksOnly;
	ArrayArguments[10] = bProgressTaskMode;
	sReturn = scScreenFunctions.DisplayPopup(window, document, "TM031.asp", ArrayArguments, 538, 418);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		var sTaskXML = sReturn[1];
		if(sTaskXML != "")
		{
			// we have some tasks to add to the current stage.
			AddTasksToStage(sTaskXML);

		}
	}
}

function GetCustomerXML()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_sRequestArray[11]);
	var tagCustomers = XML.CreateActiveTag("CUSTOMERS");
	for(var nLoop = 0; nLoop < 5; nLoop++)
	{
	
		var sCustomerName			= m_sCustomerNameArray[nLoop];
		var sCustomerNumber			= m_sCustomerNumberArray[nLoop];
		var sCustomerVersionNumber	= m_sCustomerVersionNumberArray[nLoop];
		if(sCustomerName != null && sCustomerName != "" && sCustomerNumber != "")
		{
			XML.ActiveTag = tagCustomers;
			XML.CreateActiveTag("CUSTOMER");
			XML.SetAttribute("NUMBER", sCustomerNumber);
			XML.SetAttribute("NAME", sCustomerName);	
			XML.SetAttribute("VERSION", sCustomerVersionNumber);
		}
	}
	return(XML.XMLDocument.xml);
}

function AddTasksToStage(sTaskXML)
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(sTaskXML);
	
	XML.RunASP(document, "OmigaTMBO.asp");

	var ErrorTypes = new Array("NOTASKAUTHORITY");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{
		alert("User does not have the authority to create this task");
	}
	else if(ErrorReturn[0] == true)
	{		
		updateList();

	}
}

function GetTasksXML()
{
<%	//For each row of the table selected get the associated CASETASK xml and return them all
%>	var saRowArray = scScrollTable.getArrayofRowsSelected(); //THIS MIGHT ONLY BE THE VISIBLE ONES
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var tagRequest = XML.CreateRequestTagFromArray(m_sRequestArray[11]);
	for(var iCount = 0; iCount < saRowArray.length; iCount++)
	{
		var nOrderNum = saRowArray[iCount];
		var newNode = SetActiveTask(nOrderNum);
		var node = newNode.cloneNode(true);
		XML.ActiveTag.appendChild(node);
	}
	return(XML.XMLDocument.xml);
}

function SetActiveTask(nOrderNum)
{
	m_stageXML.ActiveTag = null;
	m_stageXML.CreateTagList("CASETASK");
	var node = null;
	for (iCount = 0; iCount < m_stageXML.ActiveTagList.length && node == null; iCount++)
	{
		m_stageXML.SelectTagListItem(iCount); //sets this tag as the active tag
		var nThisOrder = parseInt(m_stageXML.GetTagText("ORDER"));
		if(nThisOrder == nOrderNum)
			node = m_stageXML.GetTagListItem(iCount);
	}
	return node;
}



function frmScreen.btnOK.onclick()
{
	window.close();
}


function enableProcessButtons()
{
	frmScreen.btnMemo.disabled = false;
	//frmScreen.btnAdd.disabled = false;
}

function disableProcessButtons()
{
	frmScreen.btnMemo.disabled = true;
	//frmScreen.btnAdd.disabled = true;
}

function updateList()
{
	m_stageXML = null;
	GetStageInfo();

	<% /* Sort by Status */ %>
	frmScreen.cboOrderBy.selectedIndex = 0;
	frmScreen.cboOrderBy.onchange();
	
}

function initTaskDetail()
{
	if (scScrollTable.getRowSelectedIndex() == null) 
	{
		tblTDName.innerText="";
		tblTDActionedDate.innerText="";
		tblTDDueDate.innerText="";
		tblTDOwner.innerText="";
		tblTDStatus.innerText="";
		tblTDUpdatedBy.innerText="";
	}
	else if(scScrollTable.getArrayofRowsSelected().length > 1)
	{
		// Multiple tasks selected
		tblTDName.innerText="Multiple Selection";
		tblTDActionedDate.innerText="";
		tblTDDueDate.innerText="";
		tblTDOwner.innerText="";
		tblTDStatus.innerText="";
		tblTDUpdatedBy.innerText="";
	}
	else
	{
		SetActiveTask(scScrollTable.getRowSelectedIndex());
		
		<%  /*Allow for a longer task name (up to 100 characters) */ %>
		scScreenFunctions.SizeTextToField(tblTDName,m_stageXML.GetAttribute("TASKNAME"));
		
		tblTDStatus.innerText = GetComboText(m_stageXML.GetAttribute("TASKSTATUS"));
		tblTDDueDate.innerText = m_stageXML.GetAttribute("TASKDUEDATEANDTIME");
		tblTDActionedDate.innerText = m_stageXML.GetAttribute("TASKSTATUSSETDATETIME");
		tblTDOwner.innerText = m_stageXML.GetAttribute("OWNINGUSERID");
		tblTDUpdatedBy.innerText = m_stageXML.GetAttribute("TASKSTATUSSETBYUSERID");
	}
}


function frmScreen.btnMemo.onclick()
{	
	var sReturn = null;
	var ArrayArguments = new Array();
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	ArrayArguments[0] = m_sReadOnly;
	ArrayArguments[1] = m_taskXML.XMLDocument.xml;
	ArrayArguments[2] = m_sRequestArray[11];

	ArrayArguments[3] = m_sReadOnlyAsCaseLocked;
	ArrayArguments[4] = m_sProcessInd;
	<% /*AW 19/09/06	EP1150 */ %>
	var bProgressTaskMode = true;
	ArrayArguments[5] = m_sCurrStageId;
	ArrayArguments[6] = bProgressTaskMode;
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "TM036P.asp", ArrayArguments, 540, 500);
	updateList();

}

function GetAllGlobalParameters()
{
	m_ParamXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_ParamXML.CreateRequestTagFromArray(m_sRequestArray[11]);		
	var ListTag = m_ParamXML.CreateActiveTag("GLOBALPARAMETERLIST");

	m_ParamXML.ActiveTag = ListTag;
	m_ParamXML.CreateActiveTag("GLOBALPARAMETER");
	m_ParamXML.CreateTag("NAME", "TMOmigaActivity");
	
	m_ParamXML.RunASP(document, "GetCurrentParameterListEx.asp");

	if (m_ParamXML.IsResponseOK())
	{
		// Get param values	
		m_sActivityId = TM040GetGlobalParameterValue("TMOmigaActivity", "STRING");
	}	
}

function TM040GetGlobalParameterValue(sParameterName, sParameterType)
{
	var sRet = "";
	
	<% /* Reset the ActiveTag so that the whole document is searched */ %>
	m_ParamXML.ActiveTag = null;
	
	<% /* Find the parameter matching the name and type provided */ %>
	m_ParamXML.CreateTagList("GLOBALPARAMETER[NAME='" + sParameterName + "']");

	if(m_ParamXML.SelectTagListItem(0)) 
		sRet = m_ParamXML.GetTagText(sParameterType);

	return sRet
}


-->
</script>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
</body>
</html>
