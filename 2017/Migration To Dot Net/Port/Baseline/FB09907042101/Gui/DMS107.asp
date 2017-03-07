<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DMS107.asp
Copyright:     Copyright © 2001 Marlborough Stirling

Description:   ???
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AT		10 July 01	Created
DR		25/01/2002	DMSSYS0003 Amended to work as pop-up within Omiga4.
DR		21/02/2002	DMSSYS0003 Fixed typo and added global declarations.
DR		15/03/02	SYS4278 removed axword cab/xml changes
DR      20/03/02    Fixing print to file
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog    Date        AQR			Description
TW      09/10/2002	SYS1115		Modified to incorporate client validation - SYS5115
MV		17/10/2002	BMIDS00664	Removed the comment tags around the Attribs include File Declaration
DPF		19/11/2002	BMIDS00887	Removed the RePrint button
BS		13/05/2003	BM0548		Use axWord2 to avoid previous versions of axWord
MC		19/04/2004	CC057		white spaces padded to the title to hide std IE "webpage dialog" message.
JD		07/07/04	BMIDS787	New version of axword 5,0,0,0
HMA 	29/09/04	BMIDS850	New version of axword 5,0,0,1
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog    Date        AQR			Description
GHun	15/07/2005	MAR7		Integrate local printing
JD		09/12/2005	MAR838		Add FileNet info from DMS105 to view properly.
PSC		16/01/2006	MAR1054		Amend to parameterise filenet user
PSC		06/02/2006	MAR1197		Amend to use file extension for filenet documents
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:
IK		25/04/2006	EP462		amend for omGemini interface
AS		03/05/2006	EP462		Amend for omGemini interface.
PB		09/05/2006	EP507		Merge in changes from MAR1296
AS		04/08/2006	EP1068		omGemini: Add support for scanned documents
AS		16/01/2007	EP1288		Removed redundant alerts - now in getArchivedDocument.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>DMS Event Screen <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex="-1" data="scTable.htm" height="1" type="text/x-scriptlet" viewastext></object>

<span style="TOP: 285px; LEFT: 310px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" viewastext tabindex="-1"></object>
</span>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" validate="onchange" mark>

<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 340px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Application Number
	<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
		<input id="txtApplicationNumber" maxlength="10" style="WIDTH: 150px" class="msgTxt" READONLY>
	</span>
</span>
<span style="LEFT: 315px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Recipient Name
	<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
		<input id="txtRecipientName" maxlength="10" style="WIDTH: 150px" class="msgTxt" READONLY>
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Document Name
	<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
		<input id="txtDocumentName" maxlength="10" style="WIDTH: 150px" class="msgTxt" READONLY>
	</span>
</span>

<span style="LEFT: 315px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Customer Name
	<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
		<input id="txtCustomerName" maxlength="10" style="WIDTH: 150px" class="msgTxt" READONLY>
	</span>
</span>

<!-- DR Doesnt work
<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
	Document Type
	<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
		<input id="txtDocumentType" maxlength="10" style="WIDTH: 100px" class="msgTxt" READONLY>
	</span>
</span>
-->

<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
	Stage Name
	<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
		<input id="txtStageName" maxlength="10" style="WIDTH: 150px" class="msgTxt" READONLY>
	</span>
</span>
<div id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 100px">
	<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
	<tr id="rowTitles">
		<td width="20%" class="TableHead">Date/Time</td>
		<td width="12%" class="TableHead">Event</td>
		<td width="7%" class="TableHead">Version</td>
		<td width="16%" class="TableHead">User Name</td>
		<td width="12%" class="TableHead">Unit Name</td>
		<td width="8%" class="TableHead">Copies</td>
		<td width="25%" class="TableHead">Output Location</td>
	</tr>
	<tr id="row01">
		<td class="TableTopLeft">&nbsp;</td><td class="TableTopCenter">&nbsp;</td>
		<td class="TableTopCenter">&nbsp;</td><td class="TableTopCenter">&nbsp;</td>
		<td class="TableTopCenter">&nbsp;</td><td class="TableTopCenter">&nbsp;</td>
		<td class="TableTopRight">&nbsp;</td>
	</tr>
	<tr id="row02">
		<td class="TableLeft">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row03">
		<td class="TableLeft">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row04">
		<td class="TableLeft">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row05">
		<td class="TableLeft">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row06">
		<td class="TableLeft">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row07">
		<td class="TableLeft">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row08">
		<td class="TableLeft">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row09">
		<td class="TableLeft">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td><td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row10">
		<td class="TableBottomLeft">&nbsp;</td><td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomCenter">&nbsp;</td><td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomCenter">&nbsp;</td><td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomRight">&nbsp;</td>
	</tr>
	</table>
<!-- 
BMIDS00887  -  DPF 19/11/2002 - removed reprint button and moved View button across
	<span style="LEFT: 470px; POSITION: absolute; TOP: 210px">
		<input id="btnRePrint" value="Reprint" type="button" style="WIDTH: 60px" class="msgButton">		
	</span>
-->
	<span style="LEFT: 530px; POSITION: absolute; TOP: 210px">
		<input id="btnView" value="View" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
</div>
</div>
</form>

<%/* Main Buttons */ %>
<div id="divPrint" style="TOP: 370px; LEFT: 10px;HEIGHT:1px; WIDTH: 1px; POSITION: ABSOLUTE">
	<input id="btnExit" value="Exit" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 71px; HEIGHT: 25px; POSITION: ABSOLUTE" class="msgButton">
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>

<!-- #include FILE="attribs/DMS107Attribs.asp" -->

<% /* Specify Code Here */ %>
<% /* MAR7 GHun */ %>
<script src="includes/Documents.js" language="jscript" type="text/javascript"></script>
<% /* MAR7 End */ %>
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = null;
var m_blnReadOnly = false;
var m_sReadOnly = "";
var scScreenFunctions = null;
var m_XML = null;
var m_iTableLength = 10;
var m_DocumentGUID = null;


var m_aArgArray = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sUserId = null;
var m_sUnitId = null;
var m_sStageId = null;
var m_sStageName = null;
var m_sCustomerNumber = null;
var m_sCustomerVersionNumber = null;
var m_sRole = null;
var m_sMachineId = null;
var m_sDistributionChannelId = null;
var m_sCustomerNumber = null;
var m_sXML = null;
var m_sMachineName = null;
var m_sChannelId = null;
var g_sDocumentName = null;
var g_sDocumentDescription = null;
var m_BaseNonPopupWindow = null;
<% /* MAR7 GHun */ %>
var m_HostTemplateID = null;
var m_sPrinterType = null;
<% /* MAR7 End */ %>
var m_sFileNetGuid = ""; // JD MAR838

<% /* AS 03/05/2006 EP462 */ %>
var m_archiveFileExtension = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	//var sScreenId = "DMS107";
	
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	//var sButtonList = new Array("Submit");
	//ShowMainButtons(sButtonList);

	RetrieveContextData();
	m_XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
					
	//FW030SetTitles("Document Events",sScreenId,scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	//scScreenFunctions.SetCurrency(window,frmScreen);
	
<%	//This function is contained in the field attributes file (remove if not required)
%>	//SetMasks();
	//Validation_Init();
		
//	m_bReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, sScreenId);
//	if (m_blnReadOnly == true) m_sReadOnly = "1";
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	refreshScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function refreshScreen()
{
	m_XML.LoadXML(m_sXML);
	m_XML.CreateTagList("REQUEST");

	if(m_XML.SelectTagListItem(0))
	{

		m_DocumentGUID = m_XML.GetAttribute("DOCUMENTGUID");
		frmScreen.txtApplicationNumber.value = m_sApplicationNumber;
		g_sDocumentName = m_XML.GetAttribute("DOCUMENTNAME");
		g_sDocumentDescription = m_XML.GetAttribute("DOCUMENTDESCRIPTION");
		
		frmScreen.txtDocumentName.value = g_sDocumentName;		
		frmScreen.txtRecipientName.value = m_XML.GetAttribute("RECIPIENTNAME");
		
		if (frmScreen.txtRecipientName.value == "")
		{
			frmScreen.txtRecipientName.value = "None";
		}
			
		frmScreen.txtCustomerName.value = m_XML.GetAttribute("CUSTOMERNAME");
		
		frmScreen.txtStageName.value = m_sStageName;
		
		<% /* MAR7 GHun */ %>
		m_HostTemplateID = m_XML.GetAttribute("HOSTTEMPLATEID");
		m_sPrinterType = m_XML.GetAttribute("PRINTERTYPE");
		<% /* MAR7 End */ %>
		//JD MAR838
		m_sFileNetGuid = m_XML.GetAttribute("FILENETGUID");
	}
	
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			m_XML.RunASP(document, "omPMRequest.asp");
			break;
		default: // Error
			m_XML.SetErrorResponse();
		}

	PopulateScreen();	
	
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
	m_sApplicationFactFindNumber = m_aArgArray[1];	
	m_sUserId = m_aArgArray[2];
	m_sUnitId = m_aArgArray[3];
	m_sStageId = m_aArgArray[4];
	m_sStageName = m_aArgArray[5]; 
	m_sCustomerNumber = m_aArgArray[6];//Customer Number from context
	m_sCustomerVersionNumber = m_aArgArray[7];//Customer Version Number from context
	m_sRole = m_aArgArray[8];//User Role from context
	m_sMachineId = m_aArgArray[9];
	m_sDistributionChannelId = m_aArgArray[10];	
	m_sXML = m_aArgArray[11];//Any XML passed through
}

function PopulateScreen()
{
	m_XML.CreateTagList("DOCUMENTDETAILS") ;
	var iNumberOfRows = m_XML.ActiveTagList.length;
	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	PopulateListBox(0);
}

function PopulateListBox(nStart)
{
	ShowList(nStart);
}

function ShowList(nStart)
{
	var iCount, xmlNodeList, xmlNode ;
	var sEventTimeStamp,sEvent,sVersion,sUserName,sUnitName,sNumCopies,sOutputLocation,sFileGuid;
	
	scScrollTable.clear();
	
	for(iCount = 0; iCount < m_XML.ActiveTagList.length  && iCount < m_iTableLength; iCount++)
	{
		m_XML.SelectTagListItem(nStart+iCount);	
		
		sEventTimeStamp = m_XML.GetAttribute("EVENTTIMESTAMP");
		sEvent = m_XML.GetAttribute("EVENT");
		sVersion = m_XML.GetAttribute("DOCUMENTVERSION");
		sUserName = m_XML.GetAttribute("USERNAME");
		sUnitName = m_XML.GetAttribute("UNITNAME");
		sOutputLocation = m_XML.GetAttribute("PRINTLOCATION");
		sNumCopies = m_XML.GetAttribute("NUMCOPIES");
		sFileGuid = m_XML.GetAttribute("FILEGUID");
				
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sEventTimeStamp);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sEvent);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sVersion);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), sUserName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4), sUnitName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5), sNumCopies);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(6), sOutputLocation);
		
		tblTable.rows(iCount+1).setAttribute("Version", sVersion);			
		tblTable.rows(iCount+1).setAttribute("FileGuid", sFileGuid);
		tblTable.rows(iCount+1).setAttribute("DocGuid", m_DocumentGUID);
		tblTable.rows(iCount+1).setAttribute("DocumentName", g_sDocumentName);
		tblTable.rows(iCount+1).setAttribute("HostTemplateId", m_HostTemplateID);
		tblTable.rows(iCount+1).setAttribute("FileNetGUID",m_sFileNetGuid);
		<% /* EP507 PB */ %>
		tblTable.rows(iCount+1).setAttribute("Event", sEvent);
		<% /* EP507 PB End */ %>
		
	}
}

<% /* EP507 PB */ %>
function spnTable.onclick()
{
	var iCurrentRow = scScrollTable.getRowSelected();	
	var sEvent = "";

	if (iCurrentRow != -1)
	{
		sEvent = tblTable.rows(iCurrentRow).getAttribute("Event");
	}

	if (sEvent == "SMS" || sEvent == "EMAIL")
	{
		frmScreen.btnView.disabled = true;
	}
	else
	{
		frmScreen.btnView.disabled = false;
	}
}
<% /* EP507 PB End */ %>

function frmScreen.btnView.onclick()
{

	var iCurrentRow = scScrollTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		alert("Select a document to view.");
		return;
	<% /* MAR7 GHun */ %>
	}

	var readOnly = true;
	var printOnly = false;
	<% /* JD MAR838 09/12/2005 */ %>
	var prXML = null
	var documentHostTemplateID = tblTable.rows(iCurrentRow).getAttribute("HostTemplateId");
	var documentName = tblTable.rows(iCurrentRow).getAttribute("DocumentName");
	var fileNetGuid = tblTable.rows(iCurrentRow).getAttribute("FileNetGUID");
	if (fileNetGuid != "")
		{
			<% /* PSC 16/01/2006 MAR1054 - Start */ %>
			GlobalXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			sFileNetUserId = GlobalXML.GetGlobalParameterString(document,'FileNetUserId');		
			prXML = getDocumentFromFileNet(tblTable, iCurrentRow, sFileNetUserId);
		}	
	else
		{
		prXML = getArchivedDocumentForRow(tblTable, iCurrentRow, readOnly, printOnly);
		}
	<% /* JD MAR838 09/12/2005 End */ %>
	//var prXML = getArchivedDocumentForRow(tblTable, iCurrentRow, readOnly, printOnly);

	<% /* EP462 AS */ %>
	if (prXML != null && m_fileContents != null && m_fileContents != "")
	{
		<% /* PSC MAR1197 06/02/2006 */ %>
		var printDocumentData = displayArchivedDocument(m_fileContents, m_HostTemplateID, g_sDocumentName, m_archiveDeliveryType, m_sPrinterType, readOnly, printOnly, m_archiveFileExtension);
		if (printDocumentData != null && printDocumentData.get_success())
		{
		}
<% /* MAR7 End */ %>
		refreshScreen();
	}
	else if (prXML != null && m_fileContentsUrl != null && m_fileContentsUrl != "")
	{
		<% /* AS 04/08/2006 EP1068 omGemini: Add support for scanned documents */ %>
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
