<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP011P.asp
Copyright:     Copyright © 2001 Marlborough Stirling

Description:   Application Review
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
CL		23/03/01	Created as popup
AT		09/05/02	SYS4548 - Remove reference to btnView which doesn't exist here
LD		23/05/02	SYS4727 Use cached versions of frame functions
MC		05/05/2004	BMIDS751	- White space added to the title
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
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>AP011P Application Review <!-- #include file="includes/TitleWhiteSpace.asp" -->       </title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 210px; LEFT: 210px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: visible" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 230px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span id="spnReasons" style="TOP: 20px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblTable" width="596px" border="0" cellspacing="0" cellpadding="0" class="msgTable">
		<tr id="rowTitles">
			<td width="30%" class="TableHead">Reason</td>	
			<td width="30%" class="TableHead">Comment</td>	
			<td width="10%" class="TableHead">Date</td>		
			<td width="10%" class="TableHead">UserId</td>
			<td width="10%" class="TableHead">Override Date</td>
			<td width="10%" class="TableHead">Override UserId</td>
		</tr>
		<tr id="row01">		
			<td class="TableTopLeft">&nbsp;</td>		
			<td class="TableTopCenter">&nbsp;</td>		
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
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>
		</tr>
		<tr id="row03">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
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
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>
		</tr>
		<tr id="row05">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>
		</tr>
		<tr id="row06">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>
		</tr>
		<tr id="row07">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>
		</tr>
		<tr id="row08">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>
		</tr>
		<tr id="row09">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableRight">&nbsp;</td>
		</tr>
		<tr id="row10">		
			<td class="TableBottomLeft">&nbsp;</td>	
			<td class="TableBottomCenter">&nbsp;</td>	
			<td class="TableBottomCenter">&nbsp;</td>	
			<td class="TableBottomCenter">&nbsp;</td>	
			<td class="TableBottomCenter">&nbsp;</td>	
			<td class="TableBottomRight">&nbsp;</td>
		</tr>
		</table>
	</span>
</div>
</form>


<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 270px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/AP011PAttribs.asp" -->



<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sApplicationNumber = null;
var	m_sApplicationFactFindNumber = null;
var scScreenFunctions;
var m_sUnitName = "";
var m_sUnitId = "";
var m_sReadOnly = "";
var m_appReviewXML = null;
var m_iTableLength = 10;
var m_RequestArray = null;
var m_BaseNonPopupWindow = null;


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

	RetrieveData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}


function RetrieveData()
{
	//var ArraySelectedRows = new Array();
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters	= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();	
	m_RequestArray = sParameters[0] // array
	m_sApplicationNumber = sParameters[1] //idApplicationNumber
	m_sApplicationFactFindNumber = sParameters[2] //idApplicationFactFindNumber
	m_sUnitId = sParameters[3] //idUnitId
	m_sUnitName = sParameters[4] //idUnitName
	m_sReadOnly = sParameters[5] //idReadOnly
}

function PopulateScreen()
{
	m_appReviewXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_appReviewXML.CreateRequestTagFromArray(m_RequestArray, "FindApplicationReviewHistoryList");
	m_appReviewXML.CreateActiveTag("APPLICATIONREVIEWHISTORY");
	m_appReviewXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	m_appReviewXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	m_appReviewXML.SetAttribute("_COMBOLOOKUP_","y");
	m_appReviewXML.RunASP(document, "omAppProc.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = m_appReviewXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
		alert("No review reasons found for this application.");
	else if(ErrorReturn[0] == true)
	{
<%		//Find out how many review reasons are outstanding (haven't been overridden). This number
		//is passed to AP013 and used to decide if the application is still under review.
%>		SetOutstandingReasons();
		PopulateListBox();
	}
}
function SetOutstandingReasons()
{
	m_appReviewXML.ActiveTag = null;
	m_appReviewXML.CreateTagList("APPLICATIONREVIEWHISTORY");
	m_nOutstandingReasons = 0;
	for (iCount = 0; iCount < m_appReviewXML.ActiveTagList.length; iCount++)
	{
		m_appReviewXML.SelectTagListItem(iCount);
		if(m_appReviewXML.GetAttribute("OVERRIDEREASON") == "")
			m_nOutstandingReasons++;
	}

}
function PopulateListBox()
{
	m_appReviewXML.ActiveTag = null;
	m_appReviewXML.CreateTagList("APPLICATIONREVIEWHISTORY");
	var iNumberOfReasons = m_appReviewXML.ActiveTagList.length;

	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfReasons);
	ShowList(0);
	//if(iNumberOfReasons > 0) scScrollTable.setRowSelected(1);
}
function ShowList(nStart)
{
	var iCount;
	scScrollTable.clear();
	for (iCount = 0; iCount < m_appReviewXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		m_appReviewXML.SelectTagListItem(iCount + nStart);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),m_appReviewXML.GetAttribute("REVIEWREASON_TEXT"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),m_appReviewXML.GetAttribute("REVIEWCOMMENTS"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),m_appReviewXML.GetAttribute("REVIEWDATETIME"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),m_appReviewXML.GetAttribute("REVIEWUSERID"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),m_appReviewXML.GetAttribute("OVERRIDEDATETIME"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5),m_appReviewXML.GetAttribute("OVERRIDENBYUSERID"));
		tblTable.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));
	}
}


function btnSubmit.onclick()
{
	window.close();	
}


function CheckPassword()
{
	var sAuthorisedUserId = "";
	var sReturn = null;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ArrayArguments = new Array();
	ArrayArguments[0] = XML.CreateRequestAttributeArray(window);
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = m_sUnitId;
	ArrayArguments[3] = m_sUnitName;
	ArrayArguments[4] = m_sApplicationNumber;
	ArrayArguments[5] = m_sApplicationFactFindNumber;
	sReturn = scScreenFunctions.DisplayPopup(window, document, "AP012.asp", ArrayArguments, 325, 265);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		sAuthorisedUserId = sReturn[1];
	}
	XML = null;
	
	return(sAuthorisedUserId);
}


function getReasonXML()
{
	//Add m_nOutstandingReasons as an attribute
	m_appReviewXML.SetAttribute("OUTSTANDINGREASONS", m_nOutstandingReasons.toString());
	var sXMLString = m_appReviewXML.ActiveTag.xml;
	return sXMLString;
}
function spnReasons.onclick()
{
	var iRowSelected = scScrollTable.getRowSelected();
	if(iRowSelected != -1)
	{
		var nReason = parseInt(tblTable.rows(iRowSelected).getAttribute("TagListItem"));
		m_appReviewXML.ActiveTag = null;
		m_appReviewXML.CreateTagList("APPLICATIONREVIEWHISTORY");
		m_appReviewXML.SelectTagListItem(nReason);
	}
}
-->
</script>
</body>
</html>




