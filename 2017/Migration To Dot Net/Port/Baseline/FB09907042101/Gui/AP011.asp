<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP011.asp
Copyright:     Copyright © 2006 Marlborough Stirling

Description:   Application Review
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
LH		14/09/06	Created
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
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 260px; LEFT: 310px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToAP013" method="post" action="AP013.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP040" method="post" action="AP040.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 240px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
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
	<span id="spnButtons" style="TOP: 200px; LEFT: 4px; POSITION: ABSOLUTE">
		<span style="TOP:0px; LEFT:0px; POSITION:ABSOLUTE">
			<input id="btnAdd" value="Add" type="button" style="WIDTH:60px" class="msgButton">
		</span>
		<span style="TOP:0px; LEFT:64px; POSITION:ABSOLUTE">
			<input id="btnView" value="View" type="button" style="WIDTH:60px" class="msgButton">
		</span>
		<span style="TOP:0px; LEFT:128px; POSITION:ABSOLUTE">
			<input id="btnOverride" value="Override" type="button" style="WIDTH:60px" class="msgButton">
		</span>
	</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 310px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/AP011Attribs.asp" -->

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
var m_blnReadOnly = false; //JR BM0271


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
	FW030SetTitles("Application Review","AP011",scScreenFunctions);


	RetrieveContextData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	PopulateScreen();

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP011"); //JR BM0271 Prefixed with m_sblnReadOnly
	
	//BG 11/11/01 SYS3455 Disable add button if readonly
	//if (m_sReadOnly == "1")
	if(m_blnReadOnly) //JR BM0271
	{
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnOverride.disabled = true; // KRW Override Disabled while case Frozen\Cancelled
		m_sReadOnly = "1";
		
	}
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
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sUnitId = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	m_sUnitName = scScreenFunctions.GetContextParameter(window,"idUnitName",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
}
function PopulateScreen()
{
	m_appReviewXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_appReviewXML.CreateRequestTag(window, "FindApplicationReviewHistoryList");
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
	frmScreen.btnView.disabled = true;
	frmScreen.btnOverride.disabled = true;
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
	//GD SYS2174
	var sOverRideDateTime;
	scScrollTable.clear();
	for (iCount = 0; iCount < m_appReviewXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		//GD SYS2174
		m_appReviewXML.SelectTagListItem(iCount + nStart);
		sOverRideDateTime = m_appReviewXML.GetAttribute("OVERRIDEDATETIME");
		<%//GD BMIDS01057 START
		/*
		if (sOverRideDateTime != "")
		{
			scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),m_appReviewXML.GetAttribute("OVERRIDEREASON_TEXT"));
			scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),m_appReviewXML.GetAttribute("OVERRIDECOMMENTS"));
		} else
		{			
			scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),m_appReviewXML.GetAttribute("REVIEWREASON_TEXT"));
			scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),m_appReviewXML.GetAttribute("REVIEWCOMMENTS"));
		}
		*/
		//GD BMIDS01057 END %>
		<% // GD BMIDS01057 START %>
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),m_appReviewXML.GetAttribute("REVIEWREASON_TEXT"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),m_appReviewXML.GetAttribute("REVIEWCOMMENTS"));
		<% //GD BMIDS01057 END %>
		
		
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),m_appReviewXML.GetAttribute("REVIEWDATETIME"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),m_appReviewXML.GetAttribute("REVIEWUSERID"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),sOverRideDateTime);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5),m_appReviewXML.GetAttribute("OVERRIDENBYUSERID"));
		tblTable.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));
	}
}
function btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
		frmToAP040.submit();
}
function CheckPassword()
{
	var sAuthorisedUserId = "";
	var sReturn = null;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ArrayArguments = new Array();
	ArrayArguments[0] = XML.CreateRequestAttributeArray(window);
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = m_sUnitId;
	ArrayArguments[3] = m_sUnitName;
	ArrayArguments[4] = m_sApplicationNumber;
	ArrayArguments[5] = m_sApplicationFactFindNumber;
	sReturn = scScreenFunctions.DisplayPopup(window, document, "AP012.asp", ArrayArguments, 328, 275);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		sAuthorisedUserId = sReturn[1];
	}
	XML = null;
	
	return(sAuthorisedUserId);
}
function frmScreen.btnAdd.onclick()
{
	var sAuthorisedUserId = CheckPassword();
	if(sAuthorisedUserId != "")
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Add_" + sAuthorisedUserId);
		frmToAP013.submit();
	}
}

function frmScreen.btnView.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","View_");
	scScreenFunctions.SetContextParameter(window,"idXML",getReasonXML());
	frmToAP013.submit();
}

function frmScreen.btnOverride.onclick()
{
	var sAuthorisedUserId = CheckPassword();
	if(sAuthorisedUserId != "")
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Override_" + sAuthorisedUserId);
		scScreenFunctions.SetContextParameter(window,"idXML",getReasonXML());
		frmToAP013.submit();
	}
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
		frmScreen.btnView.disabled = false;
		var nReason = parseInt(tblTable.rows(iRowSelected).getAttribute("TagListItem"));
		m_appReviewXML.ActiveTag = null;
		m_appReviewXML.CreateTagList("APPLICATIONREVIEWHISTORY");
		m_appReviewXML.SelectTagListItem(nReason);
		<% /* BS BM0271 15/04/03
		if(m_appReviewXML.GetAttribute("OVERRIDEREASON") == "") */ %>
		if((m_appReviewXML.GetAttribute("OVERRIDEREASON") == "") && (m_sReadOnly != "1"))
			frmScreen.btnOverride.disabled = false;
		else frmScreen.btnOverride.disabled = true;
	}
}
-->
</script>
</body>
</html>




