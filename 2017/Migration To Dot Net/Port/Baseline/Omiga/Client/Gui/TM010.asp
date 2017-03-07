<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      TM010.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Task Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		23/10/00	Created (screen paint)
JLD		22/11/00	Added implementation
JLD		30/11/00	right justify text in totals textboxes.
CL		05/03/01	SYS1920 Read only functionality added
MEVA	19/04/02	SYS2233 Widen Unit combo
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
GHun	16/06/2003	BMIDS585	Performance improvements based on MCAP00433 & MCAP00436
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
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<OBJECT data=scTable.htm height=1 id=scTable 
style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet 
VIEWASTEXT></OBJECT>
<span style="LEFT: 300px; POSITION: absolute; TOP: 330px">
<OBJECT data=scTableListScroll.asp id=scScrollTable 
style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 
type=text/x-scriptlet VIEWASTEXT></OBJECT>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToMN015" method="post" action="MN015.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" validate ="onchange" mark>
<div id="divUnit" style="HEIGHT: 70px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
	<strong>Select Unit...</strong>
</span>
<span style="LEFT: 24px; POSITION: absolute; TOP: 35px" class="msgLabel">
	Unit
	<span style="LEFT: 40px; POSITION: absolute; TOP: -3px">
		<select id="cboUnitCombo" style="WIDTH: 180px" class="msgCombo"></select>
	</span>
</span>
</div>

<div id="divUsers" style="HEIGHT: 220px; LEFT: 10px; POSITION: absolute; TOP: 140px; WIDTH: 604px" class="msgGroup">
<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 10px">
	<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
	<tr id="rowTitles">
		<td width="34%" class="TableHead">User</td>	
		<td width="22%" class="TableHead">No. Applications</td>	
		<td width="22%" class="TableHead">No. Tasks</td>		
		<td width="22%" class="TableHead">No. Outstanding Tasks</td>
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
</div>

<div id="divTotals" style="HEIGHT: 70px; LEFT: 10px; POSITION: absolute; TOP: 370px; WIDTH: 604px" class="msgGroup">
<span style="LEFT: 24px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Totals
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtTotalApps" maxlength="10" style="WIDTH: 100px; text-align: right" class="msgTxt">
	</span>
	<span style="LEFT: 230px; POSITION: absolute; TOP: -3px">
		<input id="txtTotalTasks" maxlength="10" style="WIDTH: 100px; text-align: right" class="msgTxt">
	</span>
	<span style="LEFT: 360px; POSITION: absolute; TOP: -3px">
		<input id="txtTotalOutstanding" maxlength="10" style="WIDTH: 100px; text-align: right" class="msgTxt">
	</span>
</span>
<span style="LEFT: 24px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Average per user
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtAverApps" maxlength="10" style="WIDTH: 100px; text-align: right" class="msgTxt">
	</span>
	<span style="LEFT: 230px; POSITION: absolute; TOP: -3px">
		<input id="txtAverTasks" maxlength="10" style="WIDTH: 100px; text-align: right" class="msgTxt">
	</span>
	<span style="LEFT: 360px; POSITION: absolute; TOP: -3px">
		<input id="txtAverOutstanding" maxlength="10" style="WIDTH: 100px; text-align: right" class="msgTxt">
	</span>
</span>
</div>

</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 450px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/TM010attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sActivityId = "";
var m_comboXML = null;
var m_usersXML = null;
var m_sActivityId = "";
var m_sDistributionChannel = "";
var m_iTableLength = 10;
var scScreenFunctions;


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
	FW030SetTitles("Task Summary","TM010",scScreenFunctions);

	RetrieveContextData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	PopulateCombo();
	SetFieldsToReadOnly();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
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
	m_sActivityId = scScreenFunctions.GetContextParameter(window,"idActivityId","10");
	m_sDistributionChannel = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId","1");
}

function SetFieldsToReadOnly()
{
	scScreenFunctions.SetFieldState(frmScreen, "txtTotalApps", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtTotalTasks", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtTotalOutstanding", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtAverApps", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtAverTasks", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtAverOutstanding", "R");
}
function PopulateCombo()
{
	m_comboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_comboXML.CreateRequestTag(window , "");
	m_comboXML.CreateActiveTag("UNIT");
	m_comboXML.CreateTag("CHANNELID", m_sDistributionChannel);
	m_comboXML.RunASP(document, "FindUnitList.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = m_comboXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[1] == ErrorTypes[0])
			alert("No Units could be found for the logged on Distribution Channel.");
		else
		{
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= "";
			TagOPTION.text = "<SELECT>";
			frmScreen.cboUnitCombo.add(TagOPTION);
			
			m_comboXML.CreateTagList("UNIT");
			for(var iCount = 0; iCount < m_comboXML.ActiveTagList.length; iCount++)
			{
				m_comboXML.SelectTagListItem(iCount);
				TagOPTION = document.createElement("OPTION");
				TagOPTION.value	= m_comboXML.GetTagText("UNITID");
				TagOPTION.text =  m_comboXML.GetTagText("UNITNAME");
				frmScreen.cboUnitCombo.add(TagOPTION);
			}
		}
	}
}
function frmScreen.cboUnitCombo.onchange()
{
	m_usersXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = m_usersXML.CreateRequestTag(window , "FindUnitTaskSummary");
	m_usersXML.CreateActiveTag("UNIT");
	m_usersXML.SetAttribute("UNITID", frmScreen.cboUnitCombo.value);
	m_usersXML.ActiveTag = reqTag;

	<% /* BMIDS585 Activity is never used
	m_usersXML.CreateActiveTag("ACTIVITY");
	m_usersXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	m_usersXML.SetAttribute("ACTIVITYID", m_sActivityId);
	*/ %>

	<% /* BMIDS585 Don't search if no unit is selected */ %>
	if (frmScreen.cboUnitCombo.value == "")
	{
		PopulateListBox();
		ClearTotals();
		return;
	}
	<% /* BMIDS585 End */ %>
	
	m_usersXML.RunASP(document, "OmigaTMBO.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = m_usersXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		PopulateListBox();
		if(ErrorReturn[0] == true)
		{
			m_usersXML.SelectTag(null, "UNITTOTALS");
			frmScreen.txtTotalApps.value = m_usersXML.GetAttribute("APPLICATIONS");
			frmScreen.txtTotalTasks.value = m_usersXML.GetAttribute("TASKS");
			frmScreen.txtTotalOutstanding.value = m_usersXML.GetAttribute("OUTSTANDINGTASKS");
			m_usersXML.SelectTag(null, "USERAVERAGES");
			frmScreen.txtAverApps.value = m_usersXML.GetAttribute("APPLICATIONS");
			frmScreen.txtAverTasks.value = m_usersXML.GetAttribute("TASKS");
			frmScreen.txtAverOutstanding.value = m_usersXML.GetAttribute("OUTSTANDINGTASKS");
		}
		else
		{
			ClearTotals();
			alert("No tasks found on this unit");
		}
	}
}
function PopulateListBox()
{
	m_usersXML.ActiveTag = null;
	m_usersXML.CreateTagList("OMIGAUSER");
	var iNumberOfUsers = m_usersXML.ActiveTagList.length;

	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfUsers);
	ShowList(0);
}
function ShowList(nStart)
{
	var iCount;
	scScrollTable.clear();
	for (iCount = 0; iCount < m_usersXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		m_usersXML.SelectTagListItem(iCount + nStart);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),m_usersXML.GetAttribute("USERID"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),m_usersXML.GetAttribute("APPLICATIONS"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),m_usersXML.GetAttribute("TASKS"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),m_usersXML.GetAttribute("OUTSTANDINGTASKS"));
	}
}
function btnSubmit.onclick()
{
	frmToMN015.submit();
}

<% /* BMIDS585 */ %>
function ClearTotals()
{
	frmScreen.txtTotalApps.value = "";
	frmScreen.txtTotalTasks.value = "";
	frmScreen.txtTotalOutstanding.value = "";
	frmScreen.txtAverApps.value = "";
	frmScreen.txtAverTasks.value = "";
	frmScreen.txtAverOutstanding.value = "";
}
<% /* BMIDS585 End */ %>
-->
</script>
</body>
</html>



