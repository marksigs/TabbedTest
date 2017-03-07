<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      TM036.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Task Notes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		30/10/00	Created (screen paint)
JLD		16/11/00	Implementation added (not complete)
JLD		22/11/00	Added BO calls.
JLD		24/11/00	fix to refreshing screen after add
JLD		05/01/01	SYS1792 Don't show Record Not Found error if there are no current notes for a task
CL		05/03/01	SYS1920 Read only functionality added
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
CL		21/03/01	SYS2135 New calling from facility					
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
BS		10/06/2003		BM0521 Disable Add when screen is in read only mode
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 450px; LEFT: 310px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToTM032" method="post" action="TM032.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM037" method="post" action="TM037.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 420px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
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
		<textarea id="txtNote" rows="5" style="WIDTH:300px" class="msgTxt"></textarea>
	</span>
</span>
<span style="TOP:110px; LEFT:410px; POSITION:ABSOLUTE">
	<input id="btnAdd" value="Add" type="button" style="WIDTH:60px" class="msgButton">
</span>
<span style="TOP:135px; LEFT:410px; POSITION:ABSOLUTE">
	<input id="btnClear" value="Clear" type="button" style="WIDTH:60px" class="msgButton">
</span>
<span id="spnTable" style="TOP: 210px; LEFT: 4px; POSITION: ABSOLUTE">
	<table id="tblTable" width="596px" border="0" cellspacing="0" cellpadding="0" class="msgTable">
	<tr id="rowTitles">
		<td width="20%" class="TableHead">Date</td>
		<td width="15%" class="TableHead">Author</td>
		<td width="15%" class="TableHead">Note Type</td>
		<td width="50%" class="TableHead">Details</td>
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
		<td class="TableBottomLeft">&nbsp;</td>
		<td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomRight">&nbsp;</td>
	</tr>
	</table>
</span>
<span style="TOP:390px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnViewDetails" value="View Details" type="button" style="WIDTH:80px" class="msgButton">
</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 480px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/TM036Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_sTasksXML = "";
var m_iTableLength = 10;
var m_comboXML = null;
var taskXML = null;
var taskNoteXML = null;
var scScreenFunctions;
var m_blnReadOnly = false;
<% /* BS BM0521 10/06/03 */ %>
var	m_sReadOnly = "";
var	m_sProcessInd = "";
<% /* BS BM0521 End 10/06/03 */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Task Notes","TM036",scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);
	
	RetrieveContextData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	PopulateScreen();
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "TM036");
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	<% /* BS BM0521 10/06/03 */ %>
	if(m_sReadOnly == "1" || m_sProcessInd != "1")
	{
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnClear.disabled = true;
	}
	<% /* BS BM0521 End 10/06/03 */ %>
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
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
	m_sTasksXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	<% /* BS BM0521 10/06/03 */ %>
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sProcessInd = scScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null); 
	<% /* BS BM0521 End 10/06/03 */ %>
}
function PopulateScreen()
{
	PopulateCombos();
		
	taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	taskXML.LoadXML(m_sTasksXML);
	taskXML.SelectTag(null, "CASETASK");
	frmScreen.txtTask.value = taskXML.GetAttribute("TASKNAME");

	RefreshData();
}
function RefreshData()
{
	scScreenFunctions.SetFieldState(frmScreen, "txtTask", "R");
	frmScreen.btnViewDetails.disabled = true;
	ClearData();
	
	taskNoteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	taskNoteXML.CreateRequestTag(window, "FindTaskNoteList");
	taskNoteXML.CreateActiveTag("TASKNOTE");
	taskNoteXML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
	taskNoteXML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
	taskNoteXML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
	taskNoteXML.SetAttribute("ACTIVITYINSTANCE", taskXML.GetAttribute("ACTIVITYINSTANCE"));
	taskNoteXML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
	taskNoteXML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
	taskNoteXML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
	taskNoteXML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));	
	// 	taskNoteXML.RunASP(document, "MsgTMBO.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			taskNoteXML.RunASP(document, "MsgTMBO.asp");
			break;
		default: // Error
			taskNoteXML.SetErrorResponse();
		}

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = taskNoteXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
		PopulateListBox();
}
function PopulateCombos()
{
	m_comboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("TaskNoteType", "TaskContactReason");

	if(m_comboXML.GetComboLists(document,sGroups))
	{
		m_comboXML.PopulateCombo(document,frmScreen.cboNoteType, "TaskNoteType", true);
		m_comboXML.PopulateCombo(document,frmScreen.cboReason, "TaskContactReason", true);
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
function frmScreen.btnClear.onclick()
{
	ClearData();
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
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, "CreateTaskNote");
		XML.CreateActiveTag("TASKNOTE");
		XML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
		XML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
		XML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
		XML.SetAttribute("ACTIVITYINSTANCE", taskXML.GetAttribute("ACTIVITYINSTANCE"));
		XML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
		XML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
		XML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
		XML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
		XML.SetAttribute("NOTEENTRY", frmScreen.txtNote.value);
		XML.SetAttribute("NOTEORIGINATINGUSERID", scScreenFunctions.GetContextParameter(window, "idUserId", ""));
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
		RefreshData();
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
function btnSubmit.onclick()
{
<% // An example submit function showing the use of the validation functions 
%>	if(frmScreen.onsubmit())
	{
		var strCalledFrom = scScreenFunctions.GetContextParameter(window,"idReturnScreenId");
		switch(strCalledFrom)
		{
			case "TM030": frmToTM030.submit(); break;
			case "TM037": frmToTM037.submit(); break;
			case "TM032": frmToTM032.submit(); break;
			default : alert("screen " + sInputProcess + " not found");
		}
	}
}
-->
</script>
</body>
</html>






