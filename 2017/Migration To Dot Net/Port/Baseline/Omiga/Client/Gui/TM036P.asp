<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      TM036P.asp
Copyright:     Copyright © 2003 Marlborough Stirling

Description:   Task Notes Popup
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MDC		03/04/2003	Created 
BS		10/06/2003	BM0521 Disable Add when screen is in read only mode
MC		05/05/2004	BMIDS751	- White space added to the title
AW		14/09/2006	EP1150      CC78 New arguments passed from TM040/TM030
									 For previous stages retrieve/save to TASKNOTARCHIVE table
AW		20/09/2006	EP1159		Limit the length of task note to database column size.
AW		27/09/06	EP1160		Use case task name for editable tasks.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Task Notes<!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 370px; LEFT: 220px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<% /* <form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange" year4> */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 335px; WIDTH: 518px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
	Task Notes
</span>
<span style="TOP:35px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Task
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtTask" maxlength="10" style="WIDTH:300px" class="msgTxt">
	</span>
</span>
<span style="TOP:60px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Note Type
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<select id="cboNoteType" style="WIDTH:200px" class="msgCombo"></select>
	</span>
</span>
<span style="TOP:85px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Reason
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<select id="cboReason" style="WIDTH:200px" class="msgCombo"></select>
	</span>
</span>
<span style="TOP:110px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Note Entry
	<span style="TOP:0px; LEFT:100px; POSITION:ABSOLUTE">
		<textarea id="txtNote" rows="10" style="WIDTH:300px" class="msgTxt"></textarea>
	</span>
</span>
<span style="TOP:110px; LEFT:410px; POSITION:ABSOLUTE">
	<input id="btnAdd" value="Add" type="button" style="WIDTH:60px" class="msgButton">
</span>
<span style="TOP:135px; LEFT:410px; POSITION:ABSOLUTE">
	<input id="btnClear" value="Clear" type="button" style="WIDTH:60px" class="msgButton">
</span>
<span id="spnTable" style="TOP: 270px; LEFT: 4px; POSITION: ABSOLUTE">
	<table id="tblTable" width="506px" border="0" cellspacing="0" cellpadding="0" class="msgTable">
	<tr id="rowTitles">
		<td width="15%" class="TableHead">Date</td>
		<td width="15%" class="TableHead">Author</td>
		<td width="15%" class="TableHead">Note Type</td>
		<td width="55%" class="TableHead">Details</td>
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
<span style="TOP:360px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnViewDetails" value="View Details" type="button" style="WIDTH:80px" class="msgButton">
</span>

<span style="TOP:400px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnOK" value="OK" type="button" style="WIDTH:60px" class="msgButton">
</span>

</div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/TM036Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_sTaskXML = "";
var taskXML = null;
var m_iTableLength = 5; // Both tables to have 5 rows
var m_BaseNonPopupWindow = null;
var scClientScreenFunctions;
var m_sRequestAttributes = "";
var taskNoteXML = null;
var m_comboXML = null;
<% /* BS BM0521 10/06/03 */ %>
var	m_sReadOnlyAsCaseLocked = "";
var	m_sProcessInd = "";
<% /* BS BM0521 End 10/06/03 */ %>
<% /*AW 19/09/06	EP1150 */ %>
var	m_sCurrStageId = null;
var	m_ProgressTaskMode = null;
	
function window.onload()
{
	
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	RetrieveData();
	
	SetMasks();
	Validation_Init();
	<% /* BS BM0521 10/06/03 Move SetScreenOnReadOnly to after PopulateScreen */ %>
	//SetScreenOnReadOnly();
	PopulateScreen();
	SetScreenOnReadOnly();
	
	window.returnValue = null;
	
	ClientPopulateScreen();
}

function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	<% /* BS BM0521 10/06/03 */ %>
	//if (m_sReadOnly == "1")
	if (m_sReadOnlyAsCaseLocked == "1" || m_sProcessInd == "0")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		<% /* BS BM0521 10/06/03 */ %>
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnClear.disabled = true;
		<% /* BS BM0521 End 10/06/03 */ %>
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
	<% /* BS BM0521 10/06/03 */ %>
	m_sReadOnlyAsCaseLocked = sParameters[3];
	m_sProcessInd = sParameters[4];
	<% /* BS BM0521 End 10/06/03 */ %>
	<% /*AW 19/09/06	EP1150 */ %>
	m_sCurrStageId = sParameters[5];
	m_ProgressTaskMode = sParameters[6];
}

function PopulateScreen()
{
	PopulateCombos();
		
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

	RefreshData(false);
}

function RefreshData(blnRunRules)
{
	scScreenFunctions.SetFieldState(frmScreen, "txtTask", "R");
	frmScreen.btnViewDetails.disabled = true;
	ClearData();
	
	taskNoteXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	<% /* AW 19/09/06	EP1150 - Start */ %>
	<% /* If the stage is a previous one, retrieve task note details from the archive */ %>
	var sStageID = taskXML.GetAttribute("STAGEID");
	
	if(m_ProgressTaskMode)
	{
		if(parseInt(sStageID) < parseInt(m_sCurrStageId))
		{
			taskNoteXML.CreateRequestTagFromArray(m_sRequestAttributes, "FindTaskNoteArchiveList");
		}
		else
		{
			taskNoteXML.CreateRequestTagFromArray(m_sRequestAttributes, "FindTaskNoteList");
		}
	}
	else
	{
		taskNoteXML.CreateRequestTagFromArray(m_sRequestAttributes, "FindTaskNoteList");
	}
	<% /*AW 19/09/06	EP1150  - End */ %>
	
	taskNoteXML.CreateActiveTag("TASKNOTE");
	taskNoteXML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
	taskNoteXML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
	taskNoteXML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
	taskNoteXML.SetAttribute("ACTIVITYINSTANCE", taskXML.GetAttribute("ACTIVITYINSTANCE"));
	taskNoteXML.SetAttribute("STAGEID", sStageID);	<% /*AW 19/09/06	EP1150 */ %>
	taskNoteXML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
	taskNoteXML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
	taskNoteXML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));	
	
	if(blnRunRules == false)
		taskNoteXML.RunASP(document, "MsgTMBO.asp");
	else
	{	
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
				taskNoteXML.RunASP(document, "MsgTMBO.asp");
				break;
			default: // Error
				taskNoteXML.SetErrorResponse();
			}
	}
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = taskNoteXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
		PopulateListBox();
}

function PopulateCombos()
{
	m_comboXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("TaskNoteType", "TaskContactReason");

	if(m_comboXML.GetComboLists(document,sGroups))
	{
		m_comboXML.PopulateCombo(document,frmScreen.cboNoteType, "TaskNoteType", true);
		m_comboXML.PopulateCombo(document,frmScreen.cboReason, "TaskContactReason", true);
	}
}

function PopulateListBox()
{
	taskNoteXML.ActiveTag = null;
	taskNoteXML.CreateTagList("TASKNOTE");
	var iNoOfNotes = taskNoteXML.ActiveTagList.length;

	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNoOfNotes);
	ShowList(0);
}

function ShowList(iStart)
{			
	taskNoteXML.ActiveTag = null;
	taskNoteXML.CreateTagList("TASKNOTE");
	for (var iLoop = 0; iLoop < taskNoteXML.ActiveTagList.length && iLoop < m_iTableLength; iLoop++)
	{				
		taskNoteXML.SelectTagListItem(iLoop + iStart);

		scScreenFunctions.SizeTextToField(tblTable.rows(iLoop+1).cells(0),taskNoteXML.GetAttribute("NOTEDATEANDTIME"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iLoop+1).cells(1),taskNoteXML.GetAttribute("NOTEORIGINATINGUSERID"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iLoop+1).cells(2),GetNoteTypeText(taskNoteXML.GetAttribute("NOTETYPE")));
		scScreenFunctions.SizeTextToField(tblTable.rows(iLoop+1).cells(3),taskNoteXML.GetAttribute("NOTEENTRY"));
		tblTable.rows(iLoop+1).setAttribute("TagListItem", (iLoop + iStart));
	}											
}

function GetNoteTypeText(sNoteTypeValue)
{
<%	//Return the value name for the given valueid from the TaskNoteType combo
%>	var tagLIST = m_comboXML.SelectTag(null, "LIST");
	m_comboXML.CreateTagList("LISTNAME");
	var sValueName = "";
	for(var iList = 0; iList < m_comboXML.ActiveTagList.length && sValueName == ""; iList++)
	{
		m_comboXML.SelectTagListItem(iList);
		if(m_comboXML.GetAttribute("NAME") == "TaskNoteType")
		{
			m_comboXML.CreateTagList("LISTENTRY");
			for(var iCount = 0; iCount < m_comboXML.ActiveTagList.length && sValueName == ""; iCount++)
			{
				m_comboXML.SelectTagListItem(iCount);
				if(m_comboXML.GetTagText("VALUEID") == sNoteTypeValue)
					sValueName = m_comboXML.GetTagText("VALUENAME");
			}
			m_comboXML.ActiveTag = tagLIST;
			m_comboXML.CreateTagList("LISTNAME");
		}
	}
	return sValueName;
}

function ClearData()
{
	frmScreen.txtNote.value = "";
	frmScreen.cboNoteType.value = "";
	frmScreen.cboReason.value = "";
	frmScreen.btnAdd.disabled = false;
	scScreenFunctions.SetFieldState(frmScreen, "cboNoteType", "W");
	scScreenFunctions.SetFieldState(frmScreen, "cboReason", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtNote", "W");
}

function frmScreen.btnClear.onclick()
{
	ClearData();
}

function spnTable.onclick()
{
	frmScreen.btnViewDetails.disabled = false;
}

function frmScreen.btnAdd.onclick()
{
	if(frmScreen.cboNoteType.value == "")
	{
		alert("Please enter a note type.");
		frmScreen.cboNoteType.focus();
	}
	else if(frmScreen.cboNoteType.value == "20" &&
	   frmScreen.cboReason.value == ""        )
	{
		alert("Please enter a reason for this contact.");
		frmScreen.cboReason.focus();
	}
	else if(frmScreen.txtNote.value.length == 0)
	{
		alert("Please enter the details of the Note");
		frmScreen.txtNote.focus();
	}
	else
	{
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		
		<% /* AW 19/09/06	EP1150 - Start */ %>
		<% /* If the stage is a previous one, save task note details to the archive */ %>
		var sStageID = taskXML.GetAttribute("STAGEID");
	
		if(m_ProgressTaskMode)
		{
			if(parseInt(sStageID) < parseInt(m_sCurrStageId))
			{
				XML.CreateRequestTagFromArray(m_sRequestAttributes, "CreateTaskNoteArchive");
			}
			else
			{
				XML.CreateRequestTagFromArray(m_sRequestAttributes, "CreateTaskNote");
			}
		}
		else
		{
			XML.CreateRequestTagFromArray(m_sRequestAttributes, "CreateTaskNote");
		}
		<% /* AW 19/09/06	EP1150 - End */ %>
		
		XML.CreateActiveTag("TASKNOTE");
		XML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
		XML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
		XML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
		XML.SetAttribute("ACTIVITYINSTANCE", taskXML.GetAttribute("ACTIVITYINSTANCE"));
		XML.SetAttribute("STAGEID", sStageID);	<% /* AW 19/09/06	EP1150 */ %>
		XML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
		XML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
		XML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
		XML.SetAttribute("NOTEENTRY", frmScreen.txtNote.value);
		XML.SetAttribute("NOTEORIGINATINGUSERID", m_sRequestAttributes[0]);
		XML.SetAttribute("NOTETYPE", frmScreen.cboNoteType.value);
		if(frmScreen.cboNoteType.value == "20") //Contact
			XML.SetAttribute("CONTACTNOTEREASON", frmScreen.cboReason.value);
		// 		XML.RunASP(document, "MsgTMBO.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "MsgTMBO.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		XML.IsResponseOK();
		RefreshData(true);
	}
}

function frmScreen.btnViewDetails.onclick()
{
	var iRowSelected = scScrollTable.getRowSelected();
	if (iRowSelected > -1)
	{
		var nNoteItem = parseInt(tblTable.rows(iRowSelected).getAttribute("TagListItem"));
		taskNoteXML.ActiveTag = null;
		taskNoteXML.CreateTagList("TASKNOTE");
		taskNoteXML.SelectTagListItem(nNoteItem);
		frmScreen.cboNoteType.value = taskNoteXML.GetAttribute("NOTETYPE");
		frmScreen.cboReason.value = taskNoteXML.GetAttribute("CONTACTNOTEREASON");
		frmScreen.txtNote.value = taskNoteXML.GetAttribute("NOTEENTRY");
		scScreenFunctions.SetFieldState(frmScreen, "cboNoteType", "R");
		scScreenFunctions.SetFieldState(frmScreen, "cboReason", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtNote", "R");
		frmScreen.btnAdd.disabled = true;
	}
}

function frmScreen.cboNoteType.onchange()
{
	if(frmScreen.cboNoteType.value == "20")  // "Contact"
		scScreenFunctions.SetFieldState(frmScreen, "cboReason", "W");
	else
	{
		scScreenFunctions.SetFieldState(frmScreen, "cboReason", "R");
		frmScreen.cboReason.value = "";
	}
}

function frmScreen.btnOK.onclick()
{
	var sReturn = new Array();
	sReturn[0] = IsChanged();
	window.returnValue	= sReturn;
	window.close();
}

<% /* AW 20/09/06 EP1159 */ %>
function frmScreen.txtNote.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtNote", 2000, true);
}	

-->
</script>
</body>
</html>




