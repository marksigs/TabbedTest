<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DMS110.asp
Copyright:     Copyright © 2006 Vertex

Description:   GUI Interface into the DMS Admin subsystem
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AW		27/10/06	EP1240 CC056 Created Page - Initial layout
AW		07/12/06	EP1240 CC056 Further development
AS		28/12/2006	EP1270 Gemini printing
AS		05/01/2007	EP1273 Gemini printing WIP
AS		10/01/2007	EP1284 Prevent editing a document that has been Gemini printed.
AS		15/01/2007	EP1286 Use omCRUD for getting document list.
AS		16/01/2007	EP1288 Removed redundant alerts - now in getArchivedDocument.
AS		16/01/2007	EP1296 Fixed disabling of Edit button.
AS		17/01/2007	EP1300 DMS110/DMS112 list box navigation details not always displayed.
AS		18/01/2007	EP1301 Use correspondence salutation in context.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<head>
	<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title>DMS History Screen <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<object id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp 
viewastext></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<span style="LEFT: 300px; POSITION: absolute; TOP: 304px">
	<object id=scTable style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" 
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
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		User Name
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtUserName" maxlength="20" style="WIDTH: 150px" class="msgTxt" >
		</span>
	</span>
	<span style="LEFT: 350px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Status
		<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
			<select id="cboStatus" style="WIDTH: 150px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
		Application Number
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtAppNumberSearch" maxlength="12" style="WIDTH: 150px" class="msgTxt" >
		</span>
	</span>
	<span style="LEFT: 350px; POSITION: absolute; TOP: 40px" class="msgLabel">
		Account Name
		<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountNameSearch" maxlength="20" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt" >
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Document Group
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboDocumentGroup" style="WIDTH: 150px" class="msgCombo"></select>
		</span>
	</span>
		
	<span style="LEFT: 4px; POSITION: absolute; TOP: 100px" class="msgLabel">
		Document Name
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboDocumentName" style="WIDTH: 150px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 510px; POSITION: absolute; TOP: 100px">
			<input id="btnSearch" value="Search" type="button" style="WIDTH: 75px" class="msgButton">
		</span>
	<div id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 135px">
		<table id="tblTable" width="590"  border="0" cellspacing="0" cellpadding="0" class="msgTable" >
			
			<tr id="rowTitles">
				<td class="TableHead" width="6%">Version</td>
				<td class="TableHead" width="38%">Document Name</td>
				<td class="TableHead" width="25%">Date/Time</td>
				<td class="TableHead" width="15%">User Name</td>
				<td class="TableHead" width="16%">Status</td>
			</tr>	
						
			<tr id="row01">
				<td class="TableTopLeft">&nbsp;</td>
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
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">
				<td class="TableLeft">&nbsp;</td>
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
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row05">
				<td class="TableLeft">&nbsp;</td>
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
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row07">
				<td class="TableLeft">&nbsp;</td>
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
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row09">
				<td class="TableLeft">&nbsp;</td>
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
				<td class="TableBottomRight">&nbsp;</td>
			</tr>
		</table>
		
		<span style="TOP: 200px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Document Detail</strong>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 220px" class="msgLabel">
			Application Number
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtApplicationNumber" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 315px; POSITION: absolute; TOP: 220px" class="msgLabel">
			Account Name
			<span style="LEFT: 85px; POSITION: absolute; TOP: -3px">
				<input id="txtAccountName" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 240px" class="msgLabel">
			Document Group
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtDocumentGroup" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 315px; POSITION: absolute; TOP: 240px" class="msgLabel">
			Document Name
			<span style="LEFT: 85px; POSITION: absolute; TOP: -3px">
				<input id="txtDocumentName" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 260px" class="msgLabel">
			Output Location
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtOutputLocation" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 315px; POSITION: absolute; TOP: 260px" class="msgLabel">
			Archive Location
			<span style="LEFT: 85px; POSITION: absolute; TOP: -3px">
				<input id="txtArchiveLocation" maxlength="10" style="WIDTH: 190px" class="msgTxt" >
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 280px" class="msgLabel">
			Customer Name
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtCustomerName" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 315px; POSITION: absolute; TOP: 280px" class="msgLabel">
			Recipient Name
			<span style="LEFT: 85px; POSITION: absolute; TOP: -3px">
				<input id="txtRecipientName" maxlength="10" style="WIDTH: 190px" class="msgReadOnly" readonly>
			</span>
		</span>
		<span style="LEFT: 245px; POSITION: absolute; TOP: 310px">
			<input id="btnEvents" value="Events" type="button" style="WIDTH: 75px" class="msgButton">
		</span>
		<span style="LEFT: 335px; POSITION: absolute; TOP: 310px">
			<input id="btnView" value="View" type="button" style="WIDTH: 75px" class="msgButton">
		</span>
		<span style="LEFT: 425px; POSITION: absolute; TOP: 310px">
			<input id="btnEdit" value="Edit" type="button" style="WIDTH: 75px" class="msgButton">
		</span>
		<span style="LEFT: 515px; POSITION: absolute; TOP: 310px">
			<input id="btnRePrint" value="Reprint" type="button" style="WIDTH: 75px" class="msgButton">
		</span>
		<span style="LEFT: 515px; POSITION: absolute; TOP: 340px">
			<input id="btnGeminiPrint" value="Gemini Print" type="button" style="WIDTH: 75px" class="msgButton">
		</span>
	</div>	
</div>

</form>

<% /* File containing field attributes (remove if not required) */ %>
<!--  #include FILE="attribs/DMS110attribs.asp" -->

<%/* Main Buttons */ %>
<div id="divPrint" style="LEFT: 10px; WIDTH: 1px; POSITION: absolute; TOP: 520px; HEIGHT: 1px">
	<input id="btnExit" value="Exit" type="button" style="LEFT: 0px; WIDTH: 71px; POSITION: absolute; TOP: 0px; HEIGHT: 25px" class="msgButton">
</div>
<% /* Specify Code Here */ %>

<script src="includes/Documents.js" language="jscript" type="text/javascript"></script>

<script language="JScript" type="text/javascript">
<!--
var scScreenFunctions;
var m_sApplicationNumber = null;
var m_sStageName = null;

var m_sReadOnly = "";

var m_sAFFNumber = null;
var m_sUserId = null;
var m_sUnitId = null;
var m_sStageId = null;
var m_sCustomerNumber = null;
var m_sCustomerVersionNumber = null;
var m_sRole = null;
var m_sMachineId = null;
var m_sDistributionChannelId = null;
var m_XML = null;

var m_aArgArray = null;	
var m_sRequestAttribs = null;
var m_iTableLength = 10;
var m_sCustomerNameArray = null;
var m_sCustomerNumberArray = null;
var m_sCustomerVersionNumberArray = null;
var m_sLastRowSelected = 0;
var m_sPrinterXML = null;
var m_sAttributesXML = null;
var scScreenFunctions;
var m_BaseNonPopupWindow = null;
var m_sProcessingIndicator = null;
var m_sEmailAdministrator = "";
var m_sDeliveryType = "";
var m_sTemplateID = "";
var m_sPrinterDestinationType = ""; 
var m_sRemotePrinterLocation = "";
var m_sPrinterType = "";
var m_sFirstPageTray = "";
var m_sOtherPagesTray = "";
var m_xmlLocalPrinters = null;
var m_aCustomerNameArray = null;
var m_aCustomerNumberArray = null;
var m_aCustomerVersionNumberArray = null;
var m_bIsUserQualityChecker  = null;
var m_bIsUserFulfillApproved  = null;
var m_bInApplication = false;
var m_sSalutation = null;

var m_XMLDocuments = null;
var XMLStatus = null;
	
var EVENTKEY_CREATED = "0";
var EVENTKEY_EDITED = "1";
var EVENTKEY_VIEWED = "2";
var EVENTKEY_REPRINTED = "3";
var EVENTKEY_FULFILMENTSEND = "4";
var EVENTKEY_SMS = "7";
var EVENTKEY_EMAIL = "8";

var iDMSReprintButtonAuthorityLevel = 0;
var iDMSCancelButtonAuthorityLevel = 0;

var m_archiveFileExtension = null;

var m_GlobalXML = null;
var m_scannedDocumentGroupValueId = null;

var scClientScreenFunctions;

function window.onload()
{

	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();

	RetrieveContextData();
	
	PopulateDocumentGroupCombo();
	PopulateStatusCombo();
	
	//Add a single <ALL> entry in document name combo
	InitialiseDocNameCombo();

	displayContextInfo();
	SetMasks();
	SetButtonState();	
	Validation_Init();
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);

	//ClientPopulateScreen();
}

function RetrieveContextData()
{	
	window.dialogTop	= window.dialogArguments[0] + "px";
	window.dialogLeft	= window.dialogArguments[1] + "px";
	window.dialogWidth	= window.dialogArguments[2] + "px";
	window.dialogHeight = window.dialogArguments[3] + "px";
	m_aArgArray = window.dialogArguments[4];
	m_BaseNonPopupWindow = window.dialogArguments[5];

	m_GlobalXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	iDMSReprintButtonAuthorityLevel = m_GlobalXML.GetGlobalParameterAmount(document,'DMSRePrintButtonAuthLevel');
	iDMSCancelButtonAuthorityLevel = m_GlobalXML.GetGlobalParameterAmount(document,'DMSCancelFulfilmentAuthLevel');
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sApplicationNumber = m_aArgArray[0];
	m_bInApplication = m_sApplicationNumber != null &&  m_sApplicationNumber.length > 0;
	
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
	m_sRequestAttribs = m_aArgArray[11];
	m_aCustomerNameArray = m_aArgArray[12];
	m_aCustomerNumberArray = m_aArgArray[13];
	m_aCustomerVersionNumberArray = m_aArgArray[14];
	m_sProcessingIndicator = m_aArgArray[15];
	m_sReadOnly = m_aArgArray[16];

	m_sEmailAdministrator = m_aArgArray[17];
	m_bIsUserQualityChecker = m_aArgArray[18];
	m_bIsUserFulfillApproved = m_aArgArray[19];
	m_sSalutation = m_aArgArray[20];		
}

function displayContextInfo()
{
	//If the user has selected an application, disable this search criteria
	frmScreen.txtAppNumberSearch.value = m_sApplicationNumber;
	frmScreen.txtAccountNameSearch.value = "";
	
	if (m_bInApplication) 
	{
		frmScreen.txtAppNumberSearch.disabled = true;
		frmScreen.txtAccountNameSearch.value = m_sSalutation;
		frmScreen.txtAccountNameSearch.disabled = true;
	}
	
	frmScreen.txtUserName.value = m_sUserId;
	
	if (!m_bInApplication && m_bIsUserQualityChecker == "1" ) 
	{
		frmScreen.txtUserName.disabled = false;
	}
	else
	{
		frmScreen.txtUserName.disabled = true;
	}
}

function frmScreen.btnEvents.onclick()
{	
	var iCurrentRow = scTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		alert("Select a document for events.");
		return ;
	}

	if (!m_bInApplication)
	{
		m_aArgArray[0] = m_sApplicationNumber;
	}

	m_aArgArray[11] = tblTable.rows(iCurrentRow).getAttribute("DocGuid");
	m_aArgArray[12] = frmScreen.txtDocumentGroup.value;
	m_aArgArray[13] = tblTable.rows(iCurrentRow).getAttribute("DocumentName");
	m_aArgArray[14] = tblTable.rows(iCurrentRow).getAttribute("DocumentDescription");
	m_aArgArray[15] = tblTable.rows(iCurrentRow).getAttribute("HostTemplateId");
	m_aArgArray[16] = m_sPrinterType;
	m_aArgArray[17] = tblTable.rows(iCurrentRow).getAttribute("AccountName");
	m_aArgArray[18] = tblTable.rows(iCurrentRow).getAttribute("OutputLocation");
	m_aArgArray[19] = tblTable.rows(iCurrentRow).getAttribute("CustomerName");
	m_aArgArray[20] = tblTable.rows(iCurrentRow).getAttribute("RecipientName");
	
	var nWidth = window.dialogArguments[2].replace("px","");
	var nHeight = window.dialogArguments[3].replace("px","");
	var ArrayArguments = m_aArgArray;

	scScreenFunctions.DisplayPopup(window, document, "DMS112.asp", ArrayArguments, nWidth, nHeight);
}

function frmScreen.cboDocumentGroup.onchange()
{	
	if(frmScreen.cboDocumentGroup.value != "999")
	{ 
		FilterDocuments();
	}
	else
	{
		frmScreen.cboDocumentName.length = 0;
		InitialiseDocNameCombo();
	}	
}

function FilterDocuments()
{
	//Dependent upon the option selected in the group combo, gets the list of documents
	//from omPrintBO.FindDocumentNameList
	DocumentXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	DocumentXML.CreateRequestTagFromArray(m_sRequestAttribs, "FindDocumentNameList");
	DocumentXML.CreateActiveTag("FINDDOCUMENTS");
	DocumentXML.SetAttribute("DOCUMENTGROUP", frmScreen.cboDocumentGroup.value);
	DocumentXML.SetAttribute("STAGEID", "");
	DocumentXML.SetAttribute("USERROLE", "");
	DocumentXML.SetAttribute("PRINTMENUACCESS", "");
	DocumentXML.SetAttribute("INACTIVEINDICATOR", "");
	

	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			DocumentXML.RunASP(document, "PrintManager.asp");
			break;
		default: // Error
			DocumentXML.SetErrorResponse();
	}

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = DocumentXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{		
		alert("There are currently no documents available for the selected document group.");
	}
	else if(ErrorReturn[0] == true)
	{
		PopulateDocumentNames(DocumentXML);
	}	
}

function PopulateDocumentNames(XML)
{	
	frmScreen.cboDocumentName.length = 0;
		
	XML.CreateTagList("DOCUMENTS");
	
	// Loop through all entries and only add relevant entries to combo
	for(var iLoop = 0; iLoop < XML.ActiveTagList.length; iLoop++)
	{
		XML.SelectTagListItem(iLoop);
		var sTemplateName = XML.GetAttribute("HOSTTEMPLATENAME");
		var sTemplateID = XML.GetAttribute("HOSTTEMPLATEID");
		TagOPTION = document.createElement("OPTION");
		TagOPTION.value	= sTemplateID;
		TagOPTION.text = sTemplateName;
					
		frmScreen.cboDocumentName.add(TagOPTION);
	}
							
	TagOPTION	= document.createElement("OPTION");
	TagOPTION.value	= "999";
	TagOPTION.text	= "<ALL>";
	frmScreen.cboDocumentName.add(TagOPTION);
			
	frmScreen.cboDocumentName.selectedIndex = iLoop;
}

function PopulateDocumentGroupCombo()
{
	<% //Populate document group combo %>
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("PrintDocumentType");

	if(XML.GetComboLists(document,sGroupList))
	{
		XML.CreateTagList("LISTENTRY");
		<% // Loop through all entries and only add relevant entry to combo %>
		for(var iLoop = 0; iLoop < XML.ActiveTagList.length; iLoop++)
		{	
			XML.SelectTagListItem(iLoop);

			XML.SelectTagListItem(iLoop);
			var sGroupName	= XML.GetTagText("VALUENAME");
			var sValueID = XML.GetTagText("VALUEID");
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= sValueID;
			TagOPTION.text = sGroupName;					
			frmScreen.cboDocumentGroup.add(TagOPTION);
		}
		
		TagOPTION	= document.createElement("OPTION");
		TagOPTION.value	= "999";
		TagOPTION.text	= "<ALL>";
		frmScreen.cboDocumentGroup.add(TagOPTION);
			
		frmScreen.cboDocumentGroup.selectedIndex = iLoop;
	}
	
	XML = null;
}

function PopulateStatusCombo()
{
	<% //Populate document group combo %>
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("DmsDocumentStatus");

	if(XML.GetComboLists(document,sGroupList))
	{
		XML.CreateTagList("LISTENTRY");
		<% // Loop through all entries and only add relevant entry to combo %>
		for(var iLoop = 0; iLoop < XML.ActiveTagList.length; iLoop++)
		{	
			XML.SelectTagListItem(iLoop);

			XML.SelectTagListItem(iLoop);
			var sGroupName	= XML.GetTagText("VALUENAME");
			var sValueID = XML.GetTagText("VALUEID");
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= sValueID;
			TagOPTION.text = sGroupName;					
			frmScreen.cboStatus.add(TagOPTION);
		}
		
		TagOPTION	= document.createElement("OPTION");
		TagOPTION.value	= "999";
		TagOPTION.text	= "<ALL>";
		frmScreen.cboStatus.add(TagOPTION);
			
		frmScreen.cboStatus.selectedIndex = iLoop;
	}
	
	XML = null;
}

function InitialiseDocNameCombo()
{	
	TagOPTION	= document.createElement("OPTION");
	TagOPTION.value	= "999";
	TagOPTION.text	= "<ALL>";
	frmScreen.cboDocumentName.add(TagOPTION);
}

function frmScreen.btnSearch.onclick()
{
	if(CriteriaValid())
	{
		if(IsCorrectWildcard())
		{
			//Preparing XML Request string
			m_XMLDocuments = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

			var tagREQUEST = m_XMLDocuments.CreateActiveTag("REQUEST");
			m_XMLDocuments.SetAttribute("CRUD_OP", "READ");
			m_XMLDocuments.SetAttribute("SCHEMA_NAME", "DmsSchema");
			m_XMLDocuments.SetAttribute("ENTITY_REF", "DOCUMENTLIST");
			var documentListNode = m_XMLDocuments.CreateActiveTag("DOCUMENTLIST");
			documentListNode.setAttribute("USERID", frmScreen.txtUserName.value);
			documentListNode.setAttribute("GEMINIPRINTSTATUS", frmScreen.cboStatus.value != "999" ? frmScreen.cboStatus.value : "");
			documentListNode.setAttribute("APPLICATIONNUMBER", frmScreen.txtAppNumberSearch.value);
			documentListNode.setAttribute("CORRESPONDENCESALUTATION", !m_bInApplication ? frmScreen.txtAccountNameSearch.value : "");
			documentListNode.setAttribute("DOCUMENTGROUP", frmScreen.cboDocumentGroup.value != "999" ? frmScreen.cboDocumentGroup.value : "");
			documentListNode.setAttribute("HOSTTEMPLATEID", frmScreen.cboDocumentName.value != "999" ? frmScreen.cboDocumentName.value : "");
			
			m_XMLDocuments.RunASP(document, "omCRUDIf.asp");
			
			if (m_XMLDocuments.IsResponseOK())
			{
				PopulateListBox(0);
			}
		}
	}
}

function PopulateListBox(nStart)
{
	m_XMLDocuments.CreateTagList("DOCUMENTDETAILS");
	var iNumberOfRows = m_XMLDocuments.ActiveTagList.length;
	scTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows, true);
	
	ShowList(nStart);
	
	if (iNumberOfRows == 0)
	{
		alert("Unable to find any documents for your search criteria.");
	}
	<% /* Only auto-select the first row if it contains data */ %>
	else if (tblTable.rows(1).getAttribute("DocGuid"))
	{
		scTable.setRowSelected(1);		
	}
	
	spnTable.onclick();
}

function ShowList(nStart)
{
	var iCount, xmlNode;
	var sVersion,sDocumentName,sDateTime,sUserName,sStatus;
	var sEventDate = "";

	scTable.clear();
	
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("DmsDocumentStatus");
		
	if(XML.GetComboLists(document,sGroupList))
	{
		for(iCount = 0; iCount < m_XMLDocuments.ActiveTagList.length  && iCount < m_iTableLength; iCount++)
		{
			m_XMLDocuments.SelectTagListItem(nStart+iCount);

			sVersion = m_XMLDocuments.GetAttribute("DOCUMENTVERSION");
			sDocumentName = m_XMLDocuments.GetAttribute("DOCUMENTDESCRIPTION");			
			sDateTime = m_XMLDocuments.GetAttribute("EVENTDATE");
			sUserName = m_XMLDocuments.GetAttribute("USERNAME");
			sStatus = m_XMLDocuments.GetAttribute("STATUS");		

			scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sVersion);
			scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sDocumentName);
			scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sDateTime);
			scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), sUserName);
			
			xmlNode = XML.SelectSingleNode("//LISTENTRY[VALUEID='" +  sStatus + "']/VALUENAME");		
			if(xmlNode != null)
				scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4), xmlNode.text);
			
			xmlNode = XML.SelectSingleNode("//LISTENTRY[VALUEID='" +  sStatus + "']/VALIDATIONTYPE");
			if(xmlNode != null)
				tblTable.rows(iCount+1).setAttribute("StatusValidationType", xmlNode.text);
				
			tblTable.rows(iCount+1).setAttribute("UserName", sUserName);
			tblTable.rows(iCount+1).setAttribute("ApplicationNumber", m_XMLDocuments.GetAttribute("APPLICATIONNUMBER"));
			tblTable.rows(iCount+1).setAttribute("AccountName", m_XMLDocuments.GetAttribute("ACCOUNTNAME"));
			tblTable.rows(iCount+1).setAttribute("DocumentGroup", m_XMLDocuments.GetAttribute("DOCUMENTGROUP"));				
			tblTable.rows(iCount+1).setAttribute("DocumentName", m_XMLDocuments.GetAttribute("DOCUMENTNAME"));								
			tblTable.rows(iCount+1).setAttribute("OutputLocation", m_XMLDocuments.GetAttribute("PRINTLOCATION"));
			tblTable.rows(iCount+1).setAttribute("ArchiveLocation", m_XMLDocuments.GetAttribute("ARCHIVELOCATION"));
			tblTable.rows(iCount+1).setAttribute("CustomerName", m_XMLDocuments.GetAttribute("CUSTOMERNAME"));
			tblTable.rows(iCount+1).setAttribute("RecipientName", m_XMLDocuments.GetAttribute("RECIPIENTNAME"));
			
			tblTable.rows(iCount+1).setAttribute("Version", sVersion);	
			tblTable.rows(iCount+1).setAttribute("FileGuid", m_XMLDocuments.GetAttribute("FILEGUID"));
			tblTable.rows(iCount+1).setAttribute("DocGuid", m_XMLDocuments.GetAttribute("DOCUMENTGUID"));
	
			tblTable.rows(iCount+1).setAttribute("DocType", m_XMLDocuments.GetAttribute("DOCUMENTDELIVERYTYPE"));
			tblTable.rows(iCount+1).setAttribute("DocumentDescription", m_XMLDocuments.GetAttribute("DOCUMENTDESCRIPTION"));				
			tblTable.rows(iCount+1).setAttribute("HostTemplateId", m_XMLDocuments.GetAttribute("HOSTTEMPLATEID"));
			tblTable.rows(iCount+1).setAttribute("ArchiveDate", m_XMLDocuments.GetAttribute("ARCHIVEDATE"));
			tblTable.rows(iCount+1).setAttribute("DocumentPurpose", m_XMLDocuments.GetAttribute("DOCUMENTPURPOSE"));
			tblTable.rows(iCount+1).setAttribute("DocumentLanguage", m_XMLDocuments.GetAttribute("LANGUAGE"));
			tblTable.rows(iCount+1).setAttribute("PrintDate", m_XMLDocuments.GetAttribute("PRINTDATE"));
			tblTable.rows(iCount+1).setAttribute("SourceSystem", m_XMLDocuments.GetAttribute("SOURCESYSTEM"));
			tblTable.rows(iCount+1).setAttribute("SearchKey1", m_XMLDocuments.GetAttribute("SEARCHKEY1"));
			tblTable.rows(iCount+1).setAttribute("SearchKey2", m_XMLDocuments.GetAttribute("SEARCHKEY2"));
			tblTable.rows(iCount+1).setAttribute("SearchKey3", m_XMLDocuments.GetAttribute("SEARCHKEY3"));
			tblTable.rows(iCount+1).setAttribute("TemplateId", m_XMLDocuments.GetAttribute("TEMPLATEID"));
			tblTable.rows(iCount+1).setAttribute("EventKey", m_XMLDocuments.GetAttribute("EVENTKEY"));
			tblTable.rows(iCount+1).setAttribute("EventDate", m_XMLDocuments.GetAttribute("EVENTDATE"));
			tblTable.rows(iCount+1).setAttribute("InboundDocument", m_XMLDocuments.GetAttribute("INBOUNDDOCUMENT"));
			// AS 10/01/2007 EP1284
			tblTable.rows(iCount+1).setAttribute("GeminiPrintStatus", sStatus);
		}
	}
}

function spnTable.onclick()
{
	// Moved so it does not override CheckAttributes.
	SetButtonState();

	var iRowSelected = scTable.getRowSelected();
	
	if(iRowSelected != -1)
	{
		var hostTemplateId = tblTable.rows(iRowSelected).getAttribute("HostTemplateId");
		//Need to update global app number variable, it is used in document.js
		m_sApplicationNumber = tblTable.rows(iRowSelected).getAttribute("ApplicationNumber");
		
		if (hostTemplateId && (hostTemplateId != ""))
		{	
			CheckAttributes(tblTable.rows(iRowSelected).getAttribute("HostTemplateId"), 
							tblTable.rows(iRowSelected).getAttribute("OutputLocation"));
		}
	}
						
	InitTaskDetail();
}

function InitTaskDetail()
{

	if (scTable.getRowSelected() == null || scTable.getRowSelected() == -1) 
	{
		frmScreen.txtApplicationNumber.value="";
		frmScreen.txtAccountName.value="";
		frmScreen.txtDocumentGroup.value="";
		frmScreen.txtDocumentName.value="";
		frmScreen.txtOutputLocation.value="";
		frmScreen.txtArchiveLocation.value="";
		frmScreen.txtCustomerName.value="";
		frmScreen.txtRecipientName.value="";
	}
	else if (scTable.getRowSelected() > 0) 
	{
		m_XMLDocuments.SelectTag(null, "DOCUMENTS");
		m_XMLDocuments.CreateTagList("DOCUMENTDETAILS");
		node = m_XMLDocuments.GetTagListItem(scTable.getRowSelected()); 
		
		frmScreen.txtApplicationNumber.value = tblTable.rows(scTable.getRowSelected()).getAttribute("ApplicationNumber");
		frmScreen.txtAccountName.value = tblTable.rows(scTable.getRowSelected()).getAttribute("AccountName");
		
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		var sGroupList = new Array("PrintDocumentType");

		if(XML.GetComboLists(document,sGroupList))
		{
			var sDocGroupID = tblTable.rows(scTable.getRowSelected()).getAttribute("DocumentGroup");
			xmlNode = XML.SelectSingleNode("//LISTENTRY[VALUEID='" +  sDocGroupID + "']/VALUENAME");		
			if(xmlNode != null)
				frmScreen.txtDocumentGroup.value = xmlNode.text
		}
		
		frmScreen.txtDocumentName.value = tblTable.rows(scTable.getRowSelected()).getAttribute("DocumentName");
		frmScreen.txtOutputLocation.value = tblTable.rows(scTable.getRowSelected()).getAttribute("OutputLocation");
		frmScreen.txtArchiveLocation.value = tblTable.rows(scTable.getRowSelected()).getAttribute("ArchiveLocation");
		frmScreen.txtCustomerName.value = tblTable.rows(scTable.getRowSelected()).getAttribute("CustomerName");
		frmScreen.txtRecipientName.value = tblTable.rows(scTable.getRowSelected()).getAttribute("RecipientName");
		
	}
}

function CriteriaValid()
{
	if (!m_bInApplication) 
	{
		if( (frmScreen.txtUserName.value == "") &&
			(frmScreen.txtAppNumberSearch.value == "") &&
			(frmScreen.txtAccountNameSearch.value == "") &&
			(frmScreen.cboDocumentGroup.value == "999") && 
			(frmScreen.cboDocumentName.value == "999") &&
			(frmScreen.cboStatus.value == "999") )
		{
			alert("Insufficient criteria for Search.");
			return false;
		}
	}
	
	return true;
}

function IsCorrectWildcard()
{
	<% /* Account Name checks */ %>
	var sField = frmScreen.txtAccountNameSearch.value;
	if (sField.length > 0)
	{
		var iMinChars = parseInt(m_GlobalXML.GetGlobalParameterAmount(document,"MinCharsBeforeSalutationWcard"));
	
		var iWildcardPos = sField.indexOf("*");
		if (iWildcardPos != -1) {
			if(iWildcardPos < iMinChars)
			{
				alert("You must enter at least " + iMinChars + " preceding character(s) to wildcard the Surname.");
				frmScreen.txtAccountNameSearch.focus();
				return false;
			}

			if(iWildcardPos < sField.length - 1)
			{
				alert("There must be no characters after the wildcard for the Surname.");
				frmScreen.txtAccountNameSearch.focus();
				return false;
			}
		}
	}

	return true;
}

function SetButtonState()
{
	var nSelectedRow = scTable.getRowSelected();
	if (nSelectedRow == null || nSelectedRow == -1) 
	{
		frmScreen.btnEvents.disabled = true;
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnRePrint.disabled = true;
		frmScreen.btnView.disabled = true;
	}
	else  
	{
		frmScreen.btnEvents.disabled = false;
		frmScreen.btnEdit.disabled = false;
		frmScreen.btnRePrint.disabled = false;
		frmScreen.btnView.disabled = false;
	}
	
	//Gemini Printing
	frmScreen.btnGeminiPrint.disabled = false;	
}

function frmScreen.btnView.onclick()
{
	var iCurrentRow = scTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		alert("Select a document for viewing.");
		return;
	}

	var readOnly = true;
	var printOnly = false;

	var prXML = null

	prXML = getArchivedDocumentForRow(tblTable, iCurrentRow, readOnly, printOnly);

	if (prXML != null && m_fileContents != null && m_fileContents != "")
	{
		var documentHostTemplateID = tblTable.rows(iCurrentRow).getAttribute("HostTemplateId");
		var documentName = tblTable.rows(iCurrentRow).getAttribute("DocumentName");

		var printDocumentData = displayArchivedDocument(m_fileContents, documentHostTemplateID, documentName, m_archiveDeliveryType, m_sPrinterType, readOnly, printOnly, m_archiveFileExtension);
		if (printDocumentData != null && printDocumentData.get_success())
		{
			frmScreen.btnSearch.onclick();	
		}
	}
	else if (prXML != null && m_fileContentsUrl != null && m_fileContentsUrl != "")
	{
		window.open(m_fileContentsUrl, "", "");
	}
}

function frmScreen.btnEdit.onclick()
{
	var iCurrentRow = scTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		alert("Select a document for editing.");
		return;
	}
	else if (tblTable.rows(iCurrentRow).getAttribute("GeminiPrintStatus") == "40")
	{
		// AS 10/01/2007 EP1284
		alert("Can not edit a document that has been Gemini Printed.");
		return;
	}

	var readOnly = false;
	var printOnly = false;
	
	var prXML = getArchivedDocumentForRow(tblTable, iCurrentRow, readOnly, printOnly);

	if (prXML != null && m_fileContents != null && m_fileContents != "")
	{
		var documentHostTemplateID = tblTable.rows(iCurrentRow).getAttribute("HostTemplateId");
		var documentName = tblTable.rows(iCurrentRow).getAttribute("DocumentName");
		var printDocumentData = displayArchivedDocument(m_fileContents, documentHostTemplateID, documentName, m_archiveDeliveryType, m_sPrinterType, readOnly, printOnly);

		if (printDocumentData != null && printDocumentData.get_success())
		{
			var sVersion = tblTable.rows(iCurrentRow).cells(0).innerText;
			var sFileGUID = tblTable.rows(iCurrentRow).getAttribute("FileGuid");
			var sDocGUID = tblTable.rows(iCurrentRow).getAttribute("DocGuid");


			prXML.RemoveActiveTag();
			var tagREQUEST = prXML.CreateActiveTag("REQUEST");

			if(printDocumentData.get_fileEdited())
				prXML.SetAttribute("OPERATION", "SAVEDOCUMENT");
			else
				prXML.SetAttribute("OPERATION", "UNLOCKDOCUMENT");

			prXML.SetAttribute("USERID", m_sUserId);
			prXML.SetAttribute("UNITID", m_sUnitId);
			prXML.SetAttribute("MACHINENAME", m_sMachineId);
			prXML.SetAttribute("CHANNELID", m_sDistributionChannelId);
			prXML.SetAttribute("USERAUTHORITYLEVEL", "10");


			if(printDocumentData.get_fileEdited())	
			{
				prXML.SetAttribute("EDITDOCUMENT", "TRUE");
				prXML.SetAttribute("DOCUMENTGUID", sDocGUID);
				prXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
			}

			var tagPRINTDATA = prXML.CreateTag("PRINTDOCUMENTDATA");

			prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "DELIVERYTYPE", m_archiveDeliveryType);
			prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "COMPRESSIONMETHOD", "");
			prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "COMPRESSED", "0");
			prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "FILEGUID", sFileGUID);
			prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "FILEVERSION", sVersion);

			if(printDocumentData.get_fileEdited())
			{
				prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "FILESIZE", printDocumentData.get_fileSize());	<% /* MAR7 GHun */ %>
				prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "FILECONTENTS_TYPE", "BIN.BASE64");
				prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "FILECONTENTS", printDocumentData.get_fileContents());	<% /* MAR7 GHun */ %>
			}
		
			switch (ScreenRules())
			{
				case 1: // Warning
				case 0: // OK
							prXML.RunASP(document, "omPMRequest.asp");
					break;
				default: // Error
					prXML.SetErrorResponse();
			}
			
			var errorTypes = new Array("TSQLSEVERITY16");
			errorReturn = prXML.CheckResponse(errorTypes);
			if (errorReturn[0])
			{
				frmScreen.btnSearch.onclick();
			}
			else if (errorReturn[1] == "TSQLSEVERITY16")
			{	
				alert("Cannot save your changes as the document is not locked out to you.");
			}
		}
	}
}

function frmScreen.btnRePrint.onclick()
{
	var iCurrentRow = scTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		alert("Select a document for printing.");
		return;
	}

	if (m_sPrinterType == "W" )
	{
		var readOnly = true;
		var printOnly = true;

		var prXML  = getArchivedDocumentForRow(tblTable, iCurrentRow, readOnly, printOnly);

		if (prXML != null && m_fileContents != null && m_fileContents != "")
		{
			var documentHostTemplateID = tblTable.rows(iCurrentRow).getAttribute("HostTemplateId");
			var documentName = tblTable.rows(iCurrentRow).getAttribute("DocumentName");
			
			var printDocumentData = displayArchivedDocument(m_fileContents, documentHostTemplateID, documentName, m_archiveDeliveryType, m_sPrinterType, readOnly, printOnly);
			if (printDocumentData != null && printDocumentData.get_success())
			{
				// Create reprint event in the audit trail.
				createAuditTrail(EVENTKEY_REPRINTED, iCurrentRow);
				alert("Your document has been sent to the printer.");
				frmScreen.btnSearch.onclick();
			}
			else
			{
				alert("Error sending document to the printer.");
			}
		}
		else if (prXML != null && m_fileContentsUrl != null && m_fileContentsUrl != "")
		{
			window.open(m_fileContentsUrl, "", "");
		}
	}
	else
	{

		var prXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		
		var sVersion = tblTable.rows(iCurrentRow).cells(0).innerText;
		var sFileGUID = tblTable.rows(iCurrentRow).getAttribute("FileGuid");
		var sDocGUID = tblTable.rows(iCurrentRow).getAttribute("DocGuid");
		var sDocumentName = tblTable.rows(iCurrentRow).getAttribute("DocumentName");
		var sDocumentDescription = tblTable.rows(iCurrentRow).getAttribute("DocumentDescription");

		var tagREQUEST = prXML.CreateActiveTag("REQUEST");
		prXML.SetAttribute("OPERATION", "REPRINTDOCUMENT");
		prXML.SetAttribute("USERID", m_sUserId);
		prXML.SetAttribute("UNITID", m_sUnitId);
		prXML.SetAttribute("MACHINENAME", m_sMachineId);
		prXML.SetAttribute("CHANNELID", m_sDistributionChannelId);
		prXML.SetAttribute("USERAUTHORITYLEVEL", "10");
		prXML.SetAttribute("EDITDOCUMENT", "FALSE");
		prXML.SetAttribute("FREEFORMAT", "0");
		prXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		prXML.SetAttribute("DOCUMENTGUID",sDocGUID);

		var tagGETCRITERIA = prXML.CreateTag("DOCUMENTDETAILS");
		prXML.SetAttributeOnTag("DOCUMENTDETAILS", "FILEGUID", sFileGUID);
		prXML.SetAttributeOnTag("DOCUMENTDETAILS", "FILEVERSION", sVersion);
		prXML.SetAttributeOnTag("DOCUMENTDETAILS", "DOCUMENTNAME", sDocumentName);	
		prXML.SetAttributeOnTag("DOCUMENTDETAILS", "DOCUMENTDESCRIPTION", sDocumentDescription);
		prXML.SetAttributeOnTag("DOCUMENTDETAILS", "EMAILADMINISTRATOR", m_sEmailAdministrator); <% /* MAR7 GHun */ %>

		// Re-print to default (local) printer
		m_xmlLocalPrinters = getLocalPrinters(m_xmlLocalPrinters);
		if(m_xmlLocalPrinters != null)
		{			
			var LocalPrintersXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			LocalPrintersXML.LoadXML(m_xmlLocalPrinters.xml);	
						
			LocalPrintersXML.CreateTagList("PRINTER[DEFAULTINDICATOR='1']");
								
			if(LocalPrintersXML.ActiveTagList.length == 0)
			{					
				alert("You do not have a default printer set on your PC.");	
			}
			else	
			{
				var sDefaultPrinter = LocalPrintersXML.ActiveTagList.item(0).selectSingleNode("PRINTERNAME").text;
				
				prXML.ActiveTag = tagREQUEST;
				prXML.CreateActiveTag("CONTROLDATA");
				prXML.SetAttribute("DESTINATIONTYPE", "L");
				prXML.SetAttribute("DELIVERYTYPE", m_sDeliveryType);
				prXML.SetAttribute("TEMPLATEID", m_sTemplateID);
				prXML.SetAttribute("FIRSTPAGEPRINTERTRAY", m_sFirstPageTray);
				prXML.SetAttribute("OTHERPAGESPRINTERTRAY", m_sOtherPagesTray);
				prXML.CreateActiveTag("OUTPUTTYPE");
				prXML.CreateActiveTag("PRINTER");
				prXML.CreateTag("PRINTERNAME",sDefaultPrinter);
				prXML.CreateTag("NUMBEROFCOPIES","1");

				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
						prXML.RunASP(document, "PrintManager.asp");
						break;
					default: // Error
						prXML.SetErrorResponse();
					}

				prXML.SelectTag (null,"RESPONSE");
				if (!prXML.IsResponseOK()) alert(prXML.XMLDocument.xml);	// debugging

				alert("Your document is queued on the printer");
				
				frmScreen.btnSearch.onclick();	// Refresh the screen
			}
		}
	}
}

function CheckAttributes(sDocID, outputLocation)	
{	
	var m_sEditBeforePrinting = 0;
	AttribsXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var tagREQUEST = AttribsXML.CreateActiveTag("REQUEST");
	AttribsXML.SetAttribute("OPERATION", "GetPrintAttributes");
	AttribsXML.CreateActiveTag("FINDATTRIBUTES");
	AttribsXML.SetAttribute("HOSTTEMPLATEID", sDocID);
				
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			AttribsXML.RunASP(document, "PrintManager.asp");
			break;
		default: // Error
			AttribsXML.SetErrorResponse();
		}

	if(AttribsXML.IsResponseOK())
	{	
		m_sEditBeforePrinting = AttribsXML.GetTagAttribute("ATTRIBUTES", "EDITBEFOREPRINT");
		m_sDeliveryType = AttribsXML.GetTagAttribute("ATTRIBUTES", "DELIVERYTYPE");
		m_sTemplateID = AttribsXML.GetTagAttribute("ATTRIBUTES", "DPSTEMPLATEID");
		m_sFirstPageTray = AttribsXML.GetTagAttribute("ATTRIBUTES", "FIRSTPAGEPRINTERTRAY");
		m_sOtherPagesTray = AttribsXML.GetTagAttribute("ATTRIBUTES", "OTHERPAGESPRINTERTRAY");
		m_sPrinterDestinationType = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
		m_sRemotePrinterLocation = AttribsXML.GetTagAttribute("ATTRIBUTES", "REMOTEPRINTERLOCATION");
		m_sPrinterType = getPrinterType(m_BaseNonPopupWindow, m_sPrinterDestinationType, outputLocation, m_sRemotePrinterLocation);
	}

	if (m_sEditBeforePrinting == "1" && m_sReadOnly != "1" && m_sProcessingIndicator == "1")
	{
		var preventEditInDMS = AttribsXML.GetTagAttribute("ATTRIBUTES", "PREVENTEDITINDMS");

		if(preventEditInDMS == "1")
			frmScreen.btnEdit.disabled = true;
		else
			frmScreen.btnEdit.disabled = false;
	}
	else
	{
		frmScreen.btnEdit.disabled = true;
	}
}

function createAuditTrail(eventKey, iCurrentRow) 
{
	var success = false;
		
	var documentVersion = tblTable.rows(iCurrentRow).cells(0).innerText;
	var fileGuid = tblTable.rows(iCurrentRow).getAttribute("FileGuid");
	var documentGuid = tblTable.rows(iCurrentRow).getAttribute("DocGuid");
	var documentName = tblTable.rows(iCurrentRow).getAttribute("DocumentName");
	var documentDescription = tblTable.rows(iCurrentRow).getAttribute("DocumentDescription");
	var documentHostTemplateID = tblTable.rows(iCurrentRow).getAttribute("HostTemplateId");
	
	var xmlRequestObj = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequestDocument = xmlRequestObj.XMLDocument;
	
	var xmlRequest = xmlRequestDocument.createElement("REQUEST");		
	xmlRequestDocument.appendChild(xmlRequest);
	xmlRequest.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	xmlRequest.setAttribute("DOCUMENTGUID", documentGuid);
	xmlRequest.setAttribute("USERID", m_sUserId);
	xmlRequest.setAttribute("UNITID", m_sUnitId);
	xmlRequest.setAttribute("MACHINENAME", m_sMachineId);
	xmlRequest.setAttribute("CHANNELID", m_sDistributionChannelId);
	xmlRequest.setAttribute("USERAUTHORITYLEVEL", "10");
	xmlRequest.setAttribute("OPERATION", "CREATEAUDITTRAIL");	
		
	var xmlElement = xmlRequestDocument.createElement("EVENTDETAIL");
	xmlRequest.appendChild(xmlElement);	
	xmlElement.setAttribute("DOCUMENTVERSION", documentVersion);
	xmlElement.setAttribute("EVENTKEY", eventKey);
	xmlElement.setAttribute("FILEGUID", fileGuid);

	xmlElement = xmlRequestDocument.createElement("DOCUMENTDETAILS");
	xmlRequest.appendChild(xmlElement);	
	xmlElement.setAttribute("ARCHIVETIMESTAMP", tblTable.rows(iCurrentRow).getAttribute("ArchiveDate"));
	xmlElement.setAttribute("ARCHIVEDATE", tblTable.rows(iCurrentRow).getAttribute("ArchiveDate"));
	xmlElement.setAttribute("CUSTOMERNAME", tblTable.rows(iCurrentRow).getAttribute("CustomerName"));
	xmlElement.setAttribute("CUSTOMERNO", tblTable.rows(iCurrentRow).getAttribute("CustomerName"));
	xmlElement.setAttribute("DOCUMENTDESCRIPTION", documentDescription);
	xmlElement.setAttribute("DOCUMENTGROUP", tblTable.rows(iCurrentRow).getAttribute("DocumentGroup"));
	xmlElement.setAttribute("DOCUMENTGUID", documentGuid);
	xmlElement.setAttribute("DOCUMENTNAME", documentName);
	xmlElement.setAttribute("HOSTTEMPLATEID", documentHostTemplateID);
	xmlElement.setAttribute("LANGUAGE", tblTable.rows(iCurrentRow).getAttribute("DocumentLanguage"));
	xmlElement.setAttribute("PRINTTIMESTAMP", tblTable.rows(iCurrentRow).getAttribute("PrintDate"));
	xmlElement.setAttribute("PRINTDATE", tblTable.rows(iCurrentRow).getAttribute("PrintDate"));
	xmlElement.setAttribute("RECIPIENTNAME", tblTable.rows(iCurrentRow).getAttribute("RecipientName"));
	xmlElement.setAttribute("SOURCESYSTEM", tblTable.rows(iCurrentRow).getAttribute("SourceSystem"));

	xmlElement = xmlRequestDocument.createElement("APPLICATIONDETAIL");
	xmlRequest.appendChild(xmlElement);	
	xmlElement.setAttribute("HOSTTEMPLATEID", documentHostTemplateID);
	xmlElement.setAttribute("SEARCHKEY1", tblTable.rows(iCurrentRow).getAttribute("SearchKey1"));
	xmlElement.setAttribute("SEARCHKEY2", tblTable.rows(iCurrentRow).getAttribute("SearchKey2"));
	xmlElement.setAttribute("SEARCHKEY3", tblTable.rows(iCurrentRow).getAttribute("SearchKey3"));
	xmlElement.setAttribute("TEMPLATEID", tblTable.rows(iCurrentRow).getAttribute("TemplateId"));

	switch (ScreenRules())
	{
	case 1: <% // Warning %>
	case 0: <% // OK %>
		xmlRequestObj.RunASP(document, "omPMRequest.asp");
		break;
	default: <% // Error %>
		xmlRequestObj.SetErrorResponse();
	}

	xmlRequestObj.SelectTag(null, "RESPONSE");
	if (!xmlRequestObj.IsResponseOK()) 
	{
		alert("Error auditing document.");
	}
	else
	{
		success = true;
	}
	
	return success;
}


function frmScreen.btnGeminiPrint.onclick()
{
	var row = scTable.getRowSelected();	

	var errorIfGeminiPrinted = m_GlobalXML.GetGlobalParameterBoolean(document, "GeminiErrorIfGeminiPrinted");
	
	if (row < 0)
	{
		alert("Select a document to send to Gemini printing.");
	}
	else if (tblTable.rows(row).getAttribute("StatusValidationType") == "GP" && errorIfGeminiPrinted)
	{
		alert("This document has already been sent to Gemini printing.");
	}
	else if (confirm("You are about to send this document (and any associated documents) to Gemini printing and dispatch.\r\nAre you sure you wish to continue?"))
	{
		var xmlRequestObj = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		var xmlRequestDocument = xmlRequestObj.XMLDocument;
		
		var xmlRequest = xmlRequestDocument.createElement("REQUEST");		
		xmlRequestDocument.appendChild(xmlRequest);
		xmlRequest.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		xmlRequest.setAttribute("USERID", m_sUserId);
		xmlRequest.setAttribute("UNITID", m_sUnitId);
		xmlRequest.setAttribute("MUSTEXIST", "1");
		xmlRequest.setAttribute("PRINTONHOLD", "1");
		xmlRequest.setAttribute("ERRORIFGEMINIPRINTNEVER", "1");
		xmlRequest.setAttribute("ERRORIFGEMINIPRINTED", errorIfGeminiPrinted ? "1" : "0");
		xmlRequest.setAttribute("ERRORIFDOCUMENTNOTINPACK", "1");
		xmlRequest.setAttribute("OPERATION", "GEMINISENDTOFULFILLMENT");

		var xmlPack = xmlRequestDocument.createElement("PACK");
		xmlRequest.appendChild(xmlPack);	
		xmlPack.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);

		var xmlDocumentList = xmlRequestDocument.createElement("DOCUMENTLIST");
		xmlPack.appendChild(xmlDocumentList);	

		var xmlDocument = xmlRequestDocument.createElement("DOCUMENT");
		xmlDocumentList.appendChild(xmlDocument);
		xmlDocument.setAttribute("DOCUMENTGUID", tblTable.rows(row).getAttribute("DocGuid"));
		
		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				xmlRequestObj.RunASP(document, "omPMRequest.asp");
				break;
			default: // Error
				xmlRequestObj.SetErrorResponse();
		}
		
		var errorTypes = new Array("RECORDNOTFOUND");
		var errorReturn = xmlRequestObj.CheckResponse(errorTypes);
		if (errorReturn[0])
		{
			alert("Your document(s) have been sent to Gemini Printing.");
			frmScreen.btnSearch.onclick();
		}	
		else
		{
			alert("Unable to send your document(s) to Gemini Printing.\r\nError: " + errorReturn[2]);
		}
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
