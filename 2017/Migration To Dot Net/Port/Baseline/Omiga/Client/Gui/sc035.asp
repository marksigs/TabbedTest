<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<% /*
Workfile:      sc035.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   About Omiga 4 popup
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		06/12/99	Created.
RF		28/01/00	Reworked for IE5 and performance.
AY		03/04/00	scScreenFunctions change
BG		04/10/00	SYS1603 Removed spnToLastField.onfocus()
CL		26/03/01	SYS2042	New table layout
CL		30/03/01	SYS2042 Copyright info added
CL		30/03/01	SYS2042	Resize screen
CL		30/03/01	SYS2042 Add hidden attribute
CL		03/04/01	SYS2042	Change to table column widths and work on scroll bar.
MDC		17/04/02	SYS4396 Enable client customisation of system name
LD		23/05/02	SYS4727 Use cached versions of frame functions
TW      09/10/02	SYS5115 Modified to incorporate client validation
MC		19/04/04	CC057	white spaces padded to the title text to hide the standard IE text "Web Dialog...."
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS history:

GHun	01/08/2005	MAR14	Apply ING Style Sheet and GUI Images
PJO     15/12/2005  MAR896  Copyright notice rebrand
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<html>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<link href="stylesheet.css" rel="stylesheet" type="text/css">
	<title>About <!--#include file="customise/sc035customise.asp"--> <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" viewastext tabIndex="-1"></object>
<% /* Scriptlets - remove any which are not required */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<object data="scXMLFunctions.asp" height="1" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" viewastext tabindex="-1"></object>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" viewastext tabindex="-1"></object>

<% // List Box object %>
<span id="spnListScroll">
	<span style="LEFT: 145px; POSITION: absolute; TOP: 182px">
		<object data="scTableListScroll.asp" id="scScrollPlus" style="LEFT: 0px; TOP: 0px; height:24; width:304" type="text/x-scriptlet" viewastext tabindex="-1"></object>
	</span>
</span>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Forms Here */ %>
<form id="frmScreen" mark validate="onchange" year4>
	<% // Versions %>
	<div style="HEIGHT: 205px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 443px" class="msgGroup">
		<div id="spnVersionsTable" style="LEFT: 4px; POSITION: absolute; TOP: 10px">
			<table id="m_tblVersions" width="435" border="0" cellspacing="0" cellpadding="0" class="msgTable">
				<tr>
					<td width="30%" class="TableHead">Component</td>
					<td width="20%" class="TableHead">ID</td>
					<td width="30%" class="TableHead">Version</td>
					<td width="20%" class="TableHead">Date</td>
				</tr>
				<tr>
					<td class="TableLeft">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr>
					<td class="TableLeft">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr>
					<td class="TableLeft">&nbsp;</td>	
					<td class="TableCenter">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr>
					<td class="TableLeft">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr>
					<td class="TableLeft">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr>
					<td class="TableLeft">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr>
					<td class="TableLeft">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr>
					<td class="TableLeft">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr>
					<td class="TableLeft">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableCenter">&nbsp;</td>
					<td class="TableRight">&nbsp;</td>	
				</tr>
				<tr>
					<td class="TableBottomLeft">&nbsp;</td>		
					<td class="TableBottomCenter">&nbsp;</td>
					<td class="TableBottomCenter">&nbsp;</td>	
					<td class="TableBottomRight">&nbsp;</td>
				</tr>
			</table>
		</div>
	</div>
<% // Copyright statement %>
	<div style="HEIGHT: 80px; LEFT: 10px; POSITION: absolute; TOP: 219px; WIDTH: 443px" class="msgGroup">
		<div id="spnStatementTable" style="LEFT: 4px; POSITION: absolute; TOP: 10px">
			<table id="m_tblVersions1" width="435" border="0" cellspacing="0" cellpadding="0" class="msgTable">
				<tr>
					<td class="TableTopLeft">Copyright</td>	
					<td class="TableTopCenter">&nbsp;</td>
					<td class="TableTopRight">&nbsp;</td>
				</tr>
				<tr>
					<td class="TableLeft">&nbsp;</td>
					<td class="TableCenter">Vertex Financial Services Ltd. 2005. All rights reserved.</td>	
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr>
					<td class="TableLeft">&nbsp;</td>	
					<td class="TableCenter">&nbsp;</td>
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr>
					<td class="TableBottomLeft">&nbsp;</td>
					<td class="TableBottomCenter">&nbsp;</td>
					<td class="TableBottomRight">&nbsp;</td>
				</tr>
			</table>
		</div>
	</div>
	<% // OK button %>
	<div style="LEFT: 10px; POSITION: absolute; TOP: 303px; WIDTH: 443px; HEIGHT: 50px" class="msgGroup">	
		<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE">
			<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</div>
</form>
	
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/sc035Attribs.asp" -->

<% // Specify Code Here %>
<script language="JScript" type="text/javascript">
<!--	// JScript Code
var m_VersionXML = null;
var m_iTableLength = 10;
var scScreenFunctions;
var m_iNumberOfComponents = null;

function frmScreen.btnOK.onclick()
{
	window.close();
}
		
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	scScreenFunctions.SetCollectionToReadOnly(spnVersionsTable);
	InitialiseTable();
	window.returnValue = null;
	scScreenFunctions.ShowCollection(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<%/* BG 04/10/00 SYS1603 Commented out function
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}*/%>

function InitialiseTable()
{
	if (GetComponentVersionInfo() == true)
	{
		m_VersionXML.ActiveTag = null;
		var tagCOMPONENTLIST = m_VersionXML.CreateTagList("COMPONENT");
		m_iNumberOfComponents = tagCOMPONENTLIST.length;
		scScrollPlus.initialiseTable(m_tblVersions, 0, "", DisplayVersions, m_iTableLength, m_iNumberOfComponents);
		DisplayVersions(0);
	}
}
		
function DisplayVersions(nStart)
{			
	<% // n.b. nStart is zero based %>

	var sDISPLAYNAME			= null;
	var sCOMPONENTNAME			= null;
	var sVERSION				= null;
	var sBUILD					= null;
				
	m_VersionXML.ActiveTag = null;
	var tagCOMPONENTLIST = m_VersionXML.CreateTagList("COMPONENT");
	m_iNumberOfComponents = tagCOMPONENTLIST.length;
				
	for (var nComponent = 0;
			nComponent < m_iNumberOfComponents && nComponent < m_iTableLength;
			nComponent++)
	{
		m_VersionXML.ActiveTagList = tagCOMPONENTLIST;
		if (m_VersionXML.SelectTagListItem(nComponent + nStart) == true)
		{
			sDISPLAYNAME = m_VersionXML.GetTagText("DISPLAYNAME");
			sCOMPONENTNAME = m_VersionXML.GetTagText("COMPONENTNAME");
			sVERSION = m_VersionXML.GetTagText("VERSION");
			sBUILD = m_VersionXML.GetTagText("BUILDDATE");		
			
			ShowRow(nComponent+1, sDISPLAYNAME, sCOMPONENTNAME, sVERSION, sBUILD);
			}
	}
	return m_iNumberOfComponents;
}
		
function ShowRow(iRow, sDISPLAYNAME,sCOMPONENTNAME, sVERSION, sBUILD)
{
	var iCellIndex = 0;
	scScreenFunctions.SizeTextToField(
		m_tblVersions.rows(iRow).cells(iCellIndex++),sDISPLAYNAME);
	scScreenFunctions.SizeTextToField(
		m_tblVersions.rows(iRow).cells(iCellIndex++),sCOMPONENTNAME);
	scScreenFunctions.SizeTextToField(
		m_tblVersions.rows(iRow).cells(iCellIndex++),sVERSION);
	scScreenFunctions.SizeTextToField(
		m_tblVersions.rows(iRow).cells(iCellIndex++),sBUILD);									
}

function GetComponentVersionInfo()
{
	m_VersionXML = new scXMLFunctions.XMLObject();
	
	<% // create request tag; no child nodes required %>
	m_VersionXML.CreateActiveTag("REQUEST");
	
	m_VersionXML.RunASP(document, "GetVersionList.asp");
	
	return m_VersionXML.IsResponseOK();
}

-->
</script>
</body>
</html>
