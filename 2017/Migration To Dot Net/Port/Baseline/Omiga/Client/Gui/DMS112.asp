<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DMS112.asp
Copyright:     Copyright © 2006 Vertex

Description:   GUI Interface into the DMS Admin subsystem
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AW		27/10/2006	EP1240 CC056 Created Page - Initial layout
AW		07/12/2006	EP1240 CC056 Further development
AS		05/01/2007	EP1273 Gemini printing WIP
AS		15/01/2007	EP1286 DMS112 always displays Archive Location for latest document version.
AS		16/01/2007	EP1288 Removed redundant alerts - now in getArchivedDocument.
AS		17/01/2007	EP1300 DMS110/DMS112 list box navigation details not always displayed.
AS		17/01/2007	EP1299 DMS112 - missing xml element when printing a document
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<head>
	<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title>DMS Event Screen <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<object id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp 
viewastext></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<span style="LEFT: 308px; POSITION: absolute; TOP: 330px">
	<object id=scScrollTable style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" 
		tabIndex=-1 type=text/x-scriptlet data=scTableListScroll.asp 
		viewastext>
	</object>
</span>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* AM		MARS 13/905 */ %>
<form id="frmToDMS109" method="post" action="DMS109.asp" STYLE="DISPLAY: none"></form>
<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>

<form id="frmScreen" year4 validate="onchange" mark>
<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 10px; HEIGHT: 345px" class="msgGroup">

	<div id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 20px">
		<table id="tblTable" width="596"  border="0" cellspacing="0" cellpadding="0" class="msgTable">
			
			<tr id="rowTitles">
				<td class="TableHead" width="25%">Date/Time</td>
				<td class="TableHead" width="25%">Event</td>
				<td class="TableHead" width="6%">Version</td>
				<td class="TableHead" width="18%">User Name</td>
				<td class="TableHead" width="18%">Unit Name</td>
				<td class="TableHead" width="8%">Copies</td>
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
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row11">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row12">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row13">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row14">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row15">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row16">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row17">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row18">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row19">
				<td class="TableLeft">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row20">
				<td class="TableBottomLeft">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomRight">&nbsp;</td>
			</tr>
		</table>
		
		<span style="TOP: 340px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Document Detail</strong>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 360px" class="msgLabel">
			Application Number
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtApplicationNumber" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 315px; POSITION: absolute; TOP: 360px" class="msgLabel">
			Account Name
			<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
				<input id="txtAccountName" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 380px" class="msgLabel">
			Document Group
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtDocumentGroup" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 315px; POSITION: absolute; TOP: 380px" class="msgLabel">
			Document Name
			<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
				<input id="txtDocumentName" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 400px" class="msgLabel">
			Output Location
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtOutputLocation" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 315px; POSITION: absolute; TOP: 400px" class="msgLabel">
			Archive Location
			<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
				<input id="txtArchiveLocation" maxlength="10" style="WIDTH: 190px" class="msgTxt" >
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 420px" class="msgLabel">
			Customer Name
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtCustomerName" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 315px; POSITION: absolute; TOP: 420px" class="msgLabel">
			Recipient Name
			<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
				<input id="txtRecipientName" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 515px; POSITION: absolute; TOP: 450px">
			<input id="btnView" value="View" type="button" style="WIDTH: 75px" class="msgButton">
		</span>
	</div>	
</div>

</form>

<% /* File containing field attributes (remove if not required) */ %>
<!--  #include FILE="attribs/DMS112attribs.asp" -->

<%/* Main Buttons */ %>
<div id="divPrint" style="LEFT: 10px; WIDTH: 1px; POSITION: absolute; TOP: 520px; HEIGHT: 1px">
	<input id="btnExit" value="Exit" type="button" style="LEFT: 0px; WIDTH: 71px; POSITION: absolute; TOP: 0px; HEIGHT: 25px" class="msgButton">
</div>
<% /* Specify Code Here */ %>

<script src="includes/Documents.js" language="jscript" type="text/javascript"></script>

<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = null;
var m_blnReadOnly = false;
var m_sReadOnly = "";
var scScreenFunctions = null;
var m_XML = null;
var m_iTableLength = 20;
var m_aArgArray = null;
var m_sApplicationNumber = null;
var m_DocumentGuid = null;
var m_sDocumentName = null;
var m_sDocumentDescription = null;
var m_sAccountName = null;
var m_sDocumentGroup = null;
var m_sOutputLocation = null;
var m_sCustomerName = null;
var m_sRecipientName = null;
var m_BaseNonPopupWindow = null;
var m_HostTemplateID = null;
var m_sPrinterType = null;
var m_sMachineId = null;
var m_sDistributionChannelId = null;
var m_archiveFileExtension = null;

var scClientScreenFunctions;

function window.onload()
{
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	
	GetRulesData();
	RetrieveContextData();
	
	m_XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
					
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	refreshScreen();
		
	ClientPopulateScreen();
}

function RetrieveContextData()
{	
	window.dialogTop	= window.dialogArguments[0] + "px";
	window.dialogLeft	= window.dialogArguments[1] + "px";
	window.dialogWidth	= window.dialogArguments[2] + "px";
	window.dialogHeight = window.dialogArguments[3] + "px";
	m_aArgArray = window.dialogArguments[4];
	m_BaseNonPopupWindow = window.dialogArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();	
 
	m_sApplicationNumber = m_aArgArray[0];
	m_sUserId = m_aArgArray[2];
	m_sUnitId = m_aArgArray[3];
	m_sMachineId = m_aArgArray[9];
	m_sDistributionChannelId = m_aArgArray[10];
	
	m_DocumentGuid = m_aArgArray[11];
	m_sDocumentGroup = m_aArgArray[12];
	m_sDocumentName = m_aArgArray[13];
	m_sDocumentDescription = m_aArgArray[14];
	m_HostTemplateID = m_aArgArray[15];
	m_sPrinterType = m_aArgArray[16];
	m_sAccountName = m_aArgArray[17];
	m_sOutputLocation = m_aArgArray[18];
	m_sCustomerName = m_aArgArray[19];
	m_sRecipientName = m_aArgArray[20];
}

function refreshScreen()
{
	frmScreen.txtApplicationNumber.value = m_sApplicationNumber;
	frmScreen.txtAccountName.value = m_sAccountName;
	frmScreen.txtDocumentGroup.value = m_sDocumentGroup;
	frmScreen.txtDocumentName.value = m_sDocumentName;		
	frmScreen.txtOutputLocation.value = m_sOutputLocation;
	frmScreen.txtCustomerName.value = m_sCustomerName;
	frmScreen.txtRecipientName.value = m_sRecipientName;

	if (frmScreen.txtRecipientName.value == "")
	{
		frmScreen.txtRecipientName.value = "None";
	}
	
	m_XML.ResetXMLDocument();

	var tagREQUEST = m_XML.CreateActiveTag("REQUEST");
	m_XML.SetAttribute("CRUD_OP", "READ");
	m_XML.SetAttribute("SCHEMA_NAME", "DmsSchema");
	m_XML.SetAttribute("ENTITY_REF", "DOCUMENTEVENTLIST");
	var documentEventListNode = m_XML.CreateActiveTag("DOCUMENTEVENTLIST");
	documentEventListNode.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	documentEventListNode.setAttribute("DOCUMENTGUID", m_DocumentGuid);
	
	m_XML.RunASP(document, "omCRUDIf.asp");
	
	if (m_XML.IsResponseOK())
	{
		PopulateScreen();
	}
}

function PopulateScreen()
{
	m_XML.CreateTagList("EVENTDETAILS") ;
	var iNumberOfRows = m_XML.ActiveTagList.length;
	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows, true);
	PopulateListBox(0);
}

function PopulateListBox(nStart)
{
	ShowList(nStart);
	
	<% /* Only auto-select the first row if it contains data */ %>
	if (tblTable.rows(1).getAttribute("Version"))
	{
		scScrollTable.setRowSelected(1);		
	}
	
	spnTable.onclick();
}

function ShowList(nStart)
{
	var iCount, xmlNodeList, xmlNode ;
	var sEventTimeStamp,sEvent,sVersion,sUserName,sUnitName,sNumCopies,sOutputLocation,sFileGuid,sDocGuid;
	
	scScrollTable.clear();
	
	for (iCount = 0; iCount < m_XML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		m_XML.SelectTagListItem(nStart+iCount);	
		
		sEventTimeStamp = m_XML.GetAttribute("EVENTDATE");
		sEvent = m_XML.GetAttribute("EVENTKEYTEXT");
		sVersion = m_XML.GetAttribute("FILEVERSION");
		sUserName = m_XML.GetAttribute("USERID");
		sUnitName = m_XML.GetAttribute("UNITID");
		sOutputLocation = m_XML.GetAttribute("PRINTLOCATION");
		sNumCopies = m_XML.GetAttribute("NUMBEROFCOPIES");
		sFileGuid = m_XML.GetAttribute("FILEGUID");
		sDocGuid = m_XML.GetAttribute("DOCUMENTGUID");
		sArchiveLocation = m_XML.GetAttribute("ARCHIVELOCATION");
				
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sEventTimeStamp);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sEvent);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sVersion);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), sUserName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4), sUnitName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5), sNumCopies);
		
		tblTable.rows(iCount+1).setAttribute("Version", sVersion);
		tblTable.rows(iCount+1).setAttribute("FileGuid", sFileGuid);
		tblTable.rows(iCount+1).setAttribute("DocGuid", sDocGuid);
		tblTable.rows(iCount+1).setAttribute("Event", sEvent);
		tblTable.rows(iCount+1).setAttribute("OutputLocation", sOutputLocation);
		tblTable.rows(iCount+1).setAttribute("ArchiveLocation", sArchiveLocation);		
	}
}

function spnTable.onclick()
{
	var iCurrentRow = scScrollTable.getRowSelected();	
	var sEvent = "";

	if (iCurrentRow != -1)
	{
		frmScreen.txtArchiveLocation.value = tblTable.rows(iCurrentRow).getAttribute("ArchiveLocation");
		sEvent = tblTable.rows(iCurrentRow).getAttribute("Event");
		
		if (sEvent == "SMS" || sEvent == "EMAIL")
		{
			frmScreen.btnView.disabled = true;
		}
		else
		{
			frmScreen.btnView.disabled = false;
		}
	}
	else
	{
		frmScreen.btnView.disabled = true;
	}	
}

function frmScreen.btnView.onclick()
{

	var iCurrentRow = scScrollTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		alert("Select a document to view.");
		return;
	}

	var readOnly = true;
	var printOnly = false;

	var prXML = prXML = getArchivedDocumentForRow(tblTable, iCurrentRow, readOnly, printOnly);

	if (prXML != null && m_fileContents != null && m_fileContents != "")
	{
		var printDocumentData = displayArchivedDocument(m_fileContents, m_HostTemplateID, m_sDocumentName, m_archiveDeliveryType, m_sPrinterType, readOnly, printOnly, m_archiveFileExtension);
		if (printDocumentData != null && printDocumentData.get_success())
		{
		}

		refreshScreen();
	}
	else if (prXML != null && m_fileContentsUrl != null && m_fileContentsUrl != "")
	{
		window.open(m_fileContentsUrl, "", "");
	}
}

function btnExit.onclick()
{	
	window.close();
}

-->
</script>
</body>
</html>
