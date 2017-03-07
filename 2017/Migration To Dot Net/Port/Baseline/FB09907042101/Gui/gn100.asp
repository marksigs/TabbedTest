<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
Workfile:      gn100.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Memo Pad 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
BG		17/05/00	SYS0752 Removed Tooltips
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
DPF		16/09/2002	APWP1/BM067	Add Sort By combo list to screen to sort entries in 
								list box
TW      09/10/2002  SYS5115		Modified to incorporate client validation
MO		13/11/2002	BMIDS00723	Made changes after the font sizes were changed
GD		20/11/2002	BMIDS00961	Default SortBy to Date, on validation type
GD		24/01/2003	BM0265		Ensure memopad details don't get bigger than the database definition of them
GD		27/01/2003	BM0275		Remove Sort By button (does the same as Search)
BS		20/02/2003	BM0271		Disable buttons in ReadOnly mode
BS		09/04/2003	BM0271		Remove erroneous "then"
BS		24/04/2003	BM0271		Enable screen if data frozen (case at offer stage)
MC		19/04/2004	CC057		white spaces are padded to the title text
HMA     05/08/2004  BMIDS748    Display Action Name in list.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom History:

Prog	Date		AQR			Description
GHun	22/03/2007	EP2_2067	Made entry type combos wider
CC		26/03/2007	EP2_1544	Correct -ve index use in list box double click
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Memo Pad  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Scriptlets */ %>

<span id="spnEntrySummaryListScroll">
	<span style="LEFT: 250px; POSITION: absolute; TOP: 398px">
		<object data="scTableListScroll.asp" id="scEntrySummaryTable" style="LEFT: 0px; TOP: 0px; height:24; width:304" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
	</span> 
</span>	
	
<% /* Specify Forms Here */ %>


<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" mark validate="onchange">
<% /* Search Criteria */ %>
<div style="TOP: 10px; LEFT: 10px; HEIGHT: 135px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Current Memo Pad Entry </strong>
	</span>

	<div style="TOP: 24px; LEFT: 10px; HEIGHT: 58px; WIDTH: 500px; POSITION: ABSOLUTE" class="msgGroup">	
		<span style="TOP: 10px; LEFT: 5px; POSITION: ABSOLUTE" class="msgLabel">
			Type
			<span style="TOP: -3px; LEFT: 55px; POSITION: ABSOLUTE">
				<select id="cboEntryType" style="WIDTH: 280px" class="msgCombo" disabled></select>
			</span>
		</span>
		
		<span style="TOP: 35px; LEFT: 5px; POSITION: ABSOLUTE" class="msgLabel">
			Entry
			<span style="TOP: -3px; LEFT: 55px; POSITION: ABSOLUTE">
				<TEXTAREA class=msgTxt id="txtMemoEntry" name=MemoEntry rows=5 style="POSITION: absolute; WIDTH: 450px"></TEXTAREA>  
			</span>
		</span>
		
		<span style="TOP: 35px; LEFT: 521px; POSITION: ABSOLUTE">
			<input id="btnNew" value="New" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		
		<span style="TOP: 60px; LEFT: 521px; POSITION: ABSOLUTE">
			<input id="btnSave" value="Save" type="button" style="WIDTH: 60px" class="msgButton" disabled onclick="return btnSave_onclick()">
		</span>
	</div>
</div>

<% /* Search Results */ %>
<div style="TOP: 155px; LEFT: 10px; HEIGHT: 266px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Search for existing entries</strong>
	</span>

	<span style="TOP: 4px; LEFT: 200px; POSITION: ABSOLUTE">
		<input id="btnSearch" value="Search" type="button" style="LEFT:10px; WIDTH:60px" class="msgButton">
	</span>
	
	<span style="TOP: 30px; LEFT: 15px; POSITION: ABSOLUTE" class="msgLabel">
		Search by
		<span style="TOP: -3px; LEFT: 55px; POSITION: ABSOLUTE">
			<select id="cboSearchType" style="WIDTH: 120px" class="msgCombo">
				<option value=" "> &lt;SELECT&gt; 
				<option value="1">Show All
				<option value="2">By Entry Type
				<option value="3">By Current Screen
				<option value="4">By Current Action
			</select>
		</span>
	</span>
	
	<span style="TOP: 55px; LEFT: 15px; POSITION: ABSOLUTE" class="msgLabel">
		Type
		<span style="TOP: -3px; LEFT: 55px; POSITION: ABSOLUTE">
			<select id="cboEntryTypeSearch" style="WIDTH: 280px" class="msgCombo" disabled></select>
		</span>
	</span>
<% /*GD BM0275 START
	<span style="TOP: 4px; LEFT: 300px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Sort existing entries</strong>
	</span>

	<span style="TOP: 4px; LEFT: 484px; POSITION: ABSOLUTE">
		<input id="btnSort" value="Sort" type="button" style="LEFT:10px; WIDTH:60px" class="msgButton">
	</span>
	GD BM0275 END */%>	
	<span style="TOP: 30px; LEFT: 300px; POSITION: ABSOLUTE" class="msgLabel">
		Order By 
		<span style="TOP: -3px; LEFT: 55px; POSITION: ABSOLUTE">
			<select id="cboSortBy" style="WIDTH: 120px" class="msgCombo"></select>
		</span>
	</span>

	<div id="spnTable" style="TOP: 84px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblEntrySummary" width="596" border="0" cellspacing="0"cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="17%" class="TableHead">Type</td> <td width="15%" class="TableHead">Screen/Action</td>	<td width="10%" class="TableHead">User Id</td><td width="20%" class="TableHead">Entry Date</td><td class="TableHead">Entry</td></tr>
			<tr id="row01">		<td width="17%" class="TableTopLeft">&nbsp;</td>		<td width="15%" class="TableTopCenter">&nbsp;</td>		<td width="10%" class="TableTopCenter">&nbsp;</td>  <td width="20%" class="TableTopCenter">&nbsp;</td> <td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="17%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="17%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="17%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="17%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="17%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="17%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>  <td width="20%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="17%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="17%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="17%" class="TableBottomLeft">&nbsp;</td>	<td width="15%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>  <td width="20%" class="TableBottomCenter">&nbsp;</td>		<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</div>

	<span id="spnButtons" style="TOP: 244px; LEFT: 4px; POSITION: ABSOLUTE">
		<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
			<input id="btnClose" value="Close" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>

</div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/gn100Attribs.asp" -->

<script language="JScript" type="text/javascript">
<!--
var XML;		
var scScreenFunctions ;
		
<%	/* form frmContext information */ %>
var m_sApplicationNumber = null;
var m_sAFFNumber = null;
var m_sUserId = null;
var m_sScreenId = null;
var m_sScreenDesc = null;
var m_sActionRef = null;

var m_aArgArray = null;
var m_iTableLength = 10;
var DOWNLOADXTABLE = 10;
var m_BaseNonPopupWindow = null;
<% /* BS BM0271 24/04/03 Remove m_blnReadOnly, add m_sProcessInd and m_sReadOnly*/ %>
<% /* BS BM0271 20/02/03 */ %>
//var m_blnReadOnly = false;
var m_sProcessInd = ""; 
var m_sReadOnly = "";
		

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
	var sButtonList = new Array("Cancel");

	m_aArgArray = sArguments[4];
	m_sApplicationNumber = m_aArgArray[0];
	m_sAFFNumber = m_aArgArray[1];
	m_sUserId = m_aArgArray[2];
	m_sScreenId = m_aArgArray[3];
	m_sScreenDesc = m_aArgArray[4];
<% /*	m_sActionRef = m_aArgArray[5];	*/ %>
	<% /* BS BM0271 24/04/2003 Remove ReadOnly indicator, add ProcessingInd and DataFreezeInd */ %>
	<% /* BS BM0271 20/02/03 */ %>
	//m_blnReadOnly = m_aArgArray[6];
	m_sReadOnly = m_aArgArray[6];
	m_sProcessInd = m_aArgArray[7];
	<% /* BS BM0271 End 24/04/2003 */ %>
	
	XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	GetComboLists();
	
	<% /* BS BM0271 24/04/2003 */ %>
	<% /* BS BM0271 20/02/03 */ %>
	<% /* BS BM0271 09/04/03 */ %>
	//if (m_blnReadOnly)
	if ((m_sProcessInd != "1") || (m_sReadOnly == "1")) 
	{
		setMemoPadReadOnly();
		frmScreen.btnNew.disabled = true;
	}
	<% /* BS BM0271 End 20/02/03 */ %>
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function GetComboLists()
{
	<% //GD BMIDS00961 %>
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("MemoPadEntryType","MemoPadSortBy");
	<% //var sSortBy = new Array("MemoPadSortBy"); %>
	var bSuccess = false;

	if(XML.GetComboLists(document,sGroups))
	{
		bSuccess = true;
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboEntryType,"MemoPadEntryType",true);
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboEntryTypeSearch, "MemoPadEntryType",true);
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboSortBy, "MemoPadSortBy",false);
	}

<% //APWP1/BM067 - Populate 'Sort By' combo list & then preset to sort by date %>

	if (!(bSuccess))
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	} else
	{
		<% //GD BMIDS00961 Set on validation type, not value id. %>
		scScreenFunctions.SetComboOnValidationType(frmScreen, "cboSortBy", "D");
	}
}

function frmScreen.btnNew.onclick()
{
	frmScreen.cboEntryType.disabled = false;
	frmScreen.txtMemoEntry.value = "";
	frmScreen.cboEntryType.focus();
	
	setMemoPadWritable();
}

function frmScreen.btnSearch.onclick()
{
	var nMaxRows = null;
	XML.ResetXMLDocument();
	
	nMaxRows = FindMemoPadList();
	scEntrySummaryTable.initialiseTable(tblEntrySummary, 0, " ", ShowList,10, nMaxRows);
	scEntrySummaryTable.clear();
	ShowList(0);
	clearMemoPad();
	setMemoPadReadOnly();
}

<% //APWP1/BM067 - DPF 18/09/2002 - New button, on click this will re-sort entries in list %>
<%/*GD BM0275 START 
function frmScreen.btnSort.onclick()
{
	var nMaxRows = null;
	XML.ResetXMLDocument();
	
	nMaxRows = FindMemoPadList();
	scEntrySummaryTable.initialiseTable(tblEntrySummary, 0, " ", ShowList,10, nMaxRows);
	scEntrySummaryTable.clear();
	ShowList(0);
	clearMemoPad();
	setMemoPadReadOnly();
}
GD BM0275 END */%>
function FindMemoPadList()
{
	var	sASPFile;
	var memoPadTagList;
			
	
<%	/* Validate the entries before searching. */	  
%>		
	if(frmScreen.cboSearchType.value == 2 && frmScreen.cboEntryTypeSearch.value == '')
	{
		alert ('Please select a memo pad type on which to search');
		frmScreen.cboEntryTypeSearch.focus();
		return ;
	}

<%  /* create the 'Request' Tag*/
%>
	XML.ResetXMLDocument();
	XML.CreateRequestTagFromArray(m_aArgArray,"SEARCH");
	XML.CreateActiveTag("MEMOPAD"); 
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAFFNumber );
	
<% //APWP1/BM067 - DPF 18/09/02 - If we have a sort by value selected, sort by that, else by date %>
	<%//GD BMIDS00961 There will always be something selected %>

	XML.CreateTag("SORTORDER", frmScreen.cboSortBy.value);
	
	switch(frmScreen.cboSearchType.value)
	{
		case '1':
			break;
		case '2':
			XML.CreateTag("ENTRYTYPE", frmScreen.cboEntryTypeSearch.value);
			break;
		case '3':
			XML.CreateTag("SCREENREFERENCE", m_sScreenId);
			break;
		case '4' :
<%  /* Action Ref comes only in Second phase, so is the tag to be added here  */
%>
			break;
	}

	sASPFile = "FindMemoPadList.asp";
	
	// 	XML.RunASP(document,sASPFile);
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,sASPFile);
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	
	if ( ErrorReturn[0] == true )
	{
		XML.ActiveTag = null;
		memoPadTagList = XML.CreateTagList("MEMOPAD");
		return memoPadTagList.length;
	}

	ErrorTypes = null;
	ErrorReturn = null;
}

function ShowList(iStart)
{			
	var sEntryType, sScreenId, sUserId, sEntryDateTime, sEntry, sAction;   // BMIDS748

	XML.ActiveTag = null;
	XML.CreateTagList("MEMOPAD");
	for (var iLoop = 0; iLoop < XML.ActiveTagList.length && iLoop < m_iTableLength; iLoop++)
	{				
		XML.SelectTagListItem(iLoop + iStart);

		sEntryType = XML.GetTagAttribute("ENTRYTYPE", "TEXT");
		sScreenId = XML.GetTagText("SCREENREFERENCE");
		sUserId = XML.GetTagText("USERID");
		sEntryDateTime = XML.GetTagText("ENTRYDATETIME");
		sEntry = XML.GetTagText("MEMOENTRY");
		sAction = XML.GetTagText("ACTIONNAME");                           // BMIDS748
		ShowRow(iLoop+1, sEntryType, sScreenId, sUserId, sEntryDateTime, sEntry, sAction);   // BMIDS748
	}											
}

function ShowRow(nIndex, sEntryType, sScreenId, sUserId, sEntryDateTime, sEntry, sAction)
{			 					
	scScreenFunctions.SizeTextToField(tblEntrySummary.rows(nIndex).cells(0),sEntryType);
	
	<% /* BMIDS748 If the ScreenReference is empty, display the ActionName */ %>
	if (sScreenId != "")
	{
		scScreenFunctions.SizeTextToField(tblEntrySummary.rows(nIndex).cells(1),sScreenId);
	}
	else
	{
		scScreenFunctions.SizeTextToField(tblEntrySummary.rows(nIndex).cells(1),sAction);
	}	
	scScreenFunctions.SizeTextToField(tblEntrySummary.rows(nIndex).cells(2),sUserId);
	scScreenFunctions.SizeTextToField(tblEntrySummary.rows(nIndex).cells(3),sEntryDateTime);
	scScreenFunctions.SizeTextToField(tblEntrySummary.rows(nIndex).cells(4),sEntry);
}

function frmScreen.cboSearchType.onchange()
{
	if(frmScreen.cboSearchType.value == '2')
	{
		frmScreen.cboEntryTypeSearch.value = '';
		frmScreen.cboEntryTypeSearch.disabled = false;
	}
	else
	{	
		frmScreen.cboEntryTypeSearch.value = '';
		frmScreen.cboEntryTypeSearch.disabled = true;
	}
}

function frmScreen.btnSave.onclick()
{
	if(validteCurrentEntry())
		 if (CreateMemoPadEntry()) 
			setMemoPadReadOnly() ;
}

function validteCurrentEntry()
{
	if(frmScreen.cboEntryType.value == '')
	{
		alert ('Plese Select a Type for this entry');
		frmScreen.cboEntryType.focus();
		return false;
	}
	if(frmScreen.txtMemoEntry.value == '' )
	{
		alert ('There is no memo entry to save');
		frmScreen.txtMemoEntry.focus();
		return false;
	}
	return true;
}

function CreateMemoPadEntry()
{
	var bSuccess = true;
	var memoPadXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	memoPadXML.CreateRequestTagFromArray(m_aArgArray[5],"CREATE");
	memoPadXML.CreateActiveTag("MEMOPAD");
	memoPadXML.CreateTag("MEMOPADID", null);
	memoPadXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	memoPadXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAFFNumber);
	memoPadXML.CreateTag("ENTRYTYPE", frmScreen.cboEntryType.value);
	memoPadXML.CreateTag("MEMOENTRY", frmScreen.txtMemoEntry.value);
	<% // GD 24/01/03 BM0265 START %>
	<%/*  Ensure SCREENREFERENCE,SCREENDESCRIPTION and USERID aren't bigger than the database fields
	memoPadXML.CreateTag("SCREENREFERENCE", m_sScreenId);
	memoPadXML.CreateTag("SCREENDESCRIPTION", m_sScreenDesc);
	memoPadXML.CreateTag("USERID", m_sUserId);
	*/%>
	memoPadXML.CreateTag("SCREENREFERENCE", m_sScreenId.substr(0, 5));
	memoPadXML.CreateTag("SCREENDESCRIPTION", m_sScreenDesc.substr(0, 30));
	memoPadXML.CreateTag("USERID", m_sUserId.substr(0, 10));	
	<% // GD 24/01/03 BM0265 END %>	
	// 	memoPadXML.RunASP(document,"CreateMemoPad.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			memoPadXML.RunASP(document,"CreateMemoPad.asp");
			break;
		default: // Error
			memoPadXML.SetErrorResponse();
		}

	
	bSuccess = memoPadXML.IsResponseOK();
	memoPadXML = null;
	return(bSuccess);
}
		
function setMemoPadReadOnly()
{
	frmScreen.cboEntryType.disabled = true;
<% /* Memo text may need scrolling while viewing the old entries. Do not disable it.
	frmScreen.txtMemoEntry.disabled = true;
	*/
%>
	frmScreen.btnSave.disabled = true;
}	

function setMemoPadWritable()
{
	frmScreen.cboEntryType.disabled = false;
	frmScreen.txtMemoEntry.disabled = false;
	frmScreen.btnSave.disabled = false;
}	

function clearMemoPad()
{
	frmScreen.cboEntryType.value = '';
	frmScreen.txtMemoEntry.value = '';
}

function frmScreen.btnClose.onclick()
{
	window.close();
}

function spnTable.ondblclick()
{
	var nRowSelected = scEntrySummaryTable.getOffset() + scEntrySummaryTable.getRowSelected();

	if(nRowSelected > 0) // CC 26/03/2007 EP2_1544
	{
		XML.ActiveTag = null;
		XML.CreateTagList("MEMOPAD");
	
		XML.SelectTagListItem(nRowSelected - 1);

		frmScreen.cboEntryType.value = XML.GetTagText("ENTRYTYPE");
		frmScreen.txtMemoEntry.value = XML.GetTagText("MEMOENTRY");	
		setMemoPadReadOnly();
	}	
}
<% // GD 24/01/03 BM0265 START %>
function frmScreen.txtMemoEntry.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtMemoEntry", 2000, true);
}	
<% // GD 24/01/03 BM0265 END %>
-->		
</script>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
</body>
</html>
