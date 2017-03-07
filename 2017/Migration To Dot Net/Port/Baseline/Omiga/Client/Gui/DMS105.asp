<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DMS105.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   GUI Interface into the DMS Admin subsystem
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DR		24/01/02	DMSSYS0003 Created Page from PM010.asp and AT's test system
DR      15/02/02	DMSSYS0003 You cannot use the innerText values - they are sometimes 
					shortened with ...
DR		18/02/02	DMSSYS0003 Made Application Number input field read only.
DR		15/03/02	SYS4278 axword cab/xml changes
DR      20/03/02    Fixing print to file
SG		23/04/02	SYS4407 Display EventDate rather than CreatedDate in table
LD		23/05/02	SYS4727 Use cached versions of frame functions
LD		17/10/02	BMIDS00573 Modify version of axword.cab
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog    Date        AQR			Description
TW      09/10/2002	SYS1115		Modified to incorporate client validation - SYS5115
MV		17/10/2002	BMIDS00664	Removed the comment tags around the Attribs include File Declaration
IK		15/11/2002	BMIDS00885	re-printer to default (local) printer
DPF		18/11/2002	BMIDS00887	Have added check for EditBeforePrinting attribute to enable/disable Edit button
IK		12/12/2002	BM0037		just unlock document if not saved (i.e. no version update)
BS		20/02/2003	BM0271		Disable Edit and Reprint if user not in a processing unit
IK		21/02/2003	BM0363		reference axWord 4.1.0.3
IK		21/02/2003	BM0364		remove reference to default printer in re-print alert
BS		26/02/2003	BM0271		Comment out previous changes as not going live yet
BS		10/03/2003	BM0271		Re-instate changes for BM0271
IK		17/03/2002	BM0455		use axWord2 to avoid previous versions of axWord
BS		16/04/2003	BM0271		Disable Reprint if ReadOnly because case locked
AD		05/09/2003	BMIDS634	Amend the version number for re-issue of recertified axword2
MC		19/04/2004	CC057		white spaces padded to the title to hide std IE "webpage dialog" message.
								'Reset DMS107 dialog to required. see comment at the code.
MC		24/06/2004	BMIDS775	DMS Preventing Edit, After Review stage bug fix.	
JD		07/07/04	BMIDS787	New version of axword 5,0,0,0
HMA  	29/09/04	BMIDS850 	New version of axword 5,0,0,1
HMA     08/12/04    BMIDS957    Remove VBScript code.
AS		14/01/05	BMIDS968:	New version of axword 5,0,0,2
								Fixes Axword preventing ALT+TAB

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:
GHun	15/07/2005	MAR7		Integrate local printing
AM		13/09/2005  MAR7		Adding two buttons and Calling the Re-Categorization
								Screen.
GHun	15/10/2005	MAR141		Fixed reprint error message when no document selected
TW		31/10/2005	MAR211		Disable Print button if user's access level < global parameter DMSReprintAuthorityLevel
								Add in the Cancel Fulfilment and the ReSend Fulfilment interfaces
TW 		19/11/2005	MAR581		Display FileNet records
GHun	05/12/2005	MAR796		Turn off document compression
JD		08/12/2005	MAR837		Don't enable the edit of filenet records.
JD		09/12/2005	MAR838		Send FileNet info to DMS107. Moved function GetDocumentFromFileNet to document.js so DMS107 can use it too.
HMA     05/01/2006  MAR883      Correct enabling of 'Resend To Fulfillment' button.
PSC		16/01/2006	MAR1054		Amend to parameterise filenet user
PSC		06/02/2006	MAR1197		Amend to use file extension for filenet documents
PSC		06/02/2006	MAR1233		Amend to not get print attributes if templateid is empty
GHun	24/02/2006	MAR1309		Amend spnTable.onclick to fix an error introduced by MAR883
GHun	27/02/2006	MAR1332		Fix problems related to recategorisation
PE		25/03/2006	MAR1378		Not all screens are frozen when the case is in the decline stage
PE		05/04/2006	MAR1578		Increase size of DMS109 popup.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:
IK		11/04/2006	EP378		fix call to displayArchivedDocument
IK		25/04/2006	EP462		amend for omGemini interface
AS		03/05/2006	EP462		Amend for omGemini interface.
PB		09/05/2006	EP507		Merged-in MARS changes MAR1296
PB		10/06/2006	EP529		Merged in MAR1701 - Need to add security to the button that suppresses fulfillment
PB		11/05/2006	EP529		Merged in MAR1638 - Fix enabling of CancelFulfilment button
PB		25/05/2006	EP613		FIxed bug caused by previous change
AS		04/08/2006	EP1068		omGemini: Add support for scanned documents
AS		10/08/2006	EP1080		omGemini: DMS105 disable reprint button for scanned documents
AS		16/01/2007	EP1288		Removed redundant alerts - now in getArchivedDocument.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<head>
	<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title>DMS History Screen <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp 
viewastext></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<span style="LEFT: 310px; POSITION: absolute; TOP: 280px">
<object id=scTable style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" 
tabIndex=-1 type=text/x-scriptlet data=scTableListScroll.asp 
viewastext></object>
</span>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* AM		MARS 13/905 */ %>
<form id="frmToDMS109" method="post" action="DMS109.asp" STYLE="DISPLAY: none"></form>
<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>

<form id="frmScreen" year4 validate="onchange" mark>
<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 10px; HEIGHT: 345px" class="msgGroup">
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Application Number
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtApplicationNumber" maxlength="10" style="WIDTH: 150px" class="msgReadOnly" readonly>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
		Document Group
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboDocumentGroup" style="WIDTH: 150px" class="msgCombo"></select>
		</span>
	</span>
		
	<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Stage Name
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtStageName" maxlength="10" style="WIDTH: 150px" class="msgReadOnly" readonly>
		</span>
	</span>

	<div id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 95px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			
			<tr id="rowTitles">
				<td class="TableHead" width="6%">Version</td>
				<td class="TableHead" width="20%">Document Name</td>
				<td class="TableHead" width="20%">Date</td>
				<td class="TableHead" width="15%">User Name</td>
				<td class="TableHead" width="10%">Output Location</td>
				<td class="TableHead" width="15%">Customer Name</td>
				<td class="TableHead" width="14%">Recipient</td>
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
		
		<%/*AM	 Adding two buttons Re-Categorise and Re-sendtoFulfilment and setting it to non-modal	*/ %>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 210px">
			<input id="btnRecategorise" value="Recategorise" type="button" style="LEFT: 0px; WIDTH: 120px; TOP: -1px" class="msgButton" disabled>
		 </span>
		<span style="LEFT: 130px; POSITION: absolute; TOP: 210px">
			<input id="btnEvents" value="Events" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="LEFT: 190px; POSITION: absolute; TOP: 210px">
			<input id="btnView" value="View" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="LEFT: 250px; POSITION: absolute; TOP: 210px">
			<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="LEFT: 310px; POSITION: absolute; TOP: 210px">
			<input id="btnPrint" value="Reprint" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="LEFT: 370px; POSITION: absolute; TOP: 210px">
			<input id="btnResendtoFulfilment" value="Resend to Fulfilment" type="button" style="WIDTH: 120px" class="msgButton" >
		</span>	
		<span style="LEFT: 490px; POSITION: absolute; TOP: 210px">
			<input id="btnCancelFulfilment" value="Cancel Fulfilment" type="button" style="WIDTH: 90px" class="msgButton">
		</span>	
	</div>	
</div>

</form>

<% /* File containing field attributes (remove if not required) */ %>
<!--  #include FILE="attribs/DMS105attribs.asp" -->

<%/* Main Buttons */ %>
<div id="divPrint" style="LEFT: 10px; WIDTH: 1px; POSITION: absolute; TOP: 370px; HEIGHT: 1px">
	<input id="btnExit" value="Exit" type="button" style="LEFT: 0px; WIDTH: 71px; POSITION: absolute; TOP: 0px; HEIGHT: 25px" class="msgButton">
</div>
<% /* Specify Code Here */ %>
<% /* MAR7 GHun */ %>
<script src="includes/Documents.js" language="jscript" type="text/javascript"></script>
<% /* MAR7 End */ %>
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
var m_sCaseTaskXML = null;
var m_sPrinterXML = null;
var m_sAttributesXML = null;
var scScreenFunctions;
var m_BaseNonPopupWindow = null;
<% /* BS BM0271 20/02/03 */ %>
var m_sProcessingIndicator = null;
<% /* MAR7 GHun */ %>
var m_sEmailAdministrator = "";
var m_sDeliveryType = "";
var m_sTemplateID = "";
var m_sPrinterDestinationType = ""; 
var m_sRemotePrinterLocation = "";
var m_sPrinterType = "";
var m_sFirstPageTray = "";
var m_sOtherPagesTray = "";
var m_xmlLocalPrinters = null;
<%/* AM Added */%>
var m_aCustomerNameArray = null;
var m_aCustomerNumberArray = null;
var m_aCustomerVersionNumberArray = null;

<% /* MAR1332 GHun Make eventkeys consistent with other components */ %>
var EVENTKEY_CREATED = "0";
var EVENTKEY_EDITED = "1";
var EVENTKEY_VIEWED = "2";
var EVENTKEY_REPRINTED = "3";
var EVENTKEY_FULFILMENTSEND = "4";
var EVENTKEY_FULFILMENTRESEND = "5";
var EVENTKEY_FULFILMENTCANCEL = "6";
var EVENTKEY_SMS = "7";
var EVENTKEY_EMAIL = "8";
var EVENTKEY_RECATEGORISATION = "9";
<% /* MAR1332 GHun */ %>

var iDMSReprintButtonAuthorityLevel = 0;
var iDMSCancelButtonAuthorityLevel = 0;
<% /* EP529 / MAR1638 
var m_AccessType; */ %>
<% /* MAR211End  */%>

<% /* AS 03/05/2006 EP462 */ %>
var m_archiveFileExtension = null;

<% /* AS 10/08/2006 EP1080 */ %>
var m_GlobalXML = null;
var m_scannedDocumentGroupValueId = null;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();

	RetrieveContextData();
		
	displayContextInfo();
	
	PopulateCombos();
		
<%	//This function is contained in the field attributes file (remove if not required)
%>	//SetMasks();	
	
	Validation_Init();	
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);

<% /* MAR211 */%>
	// Disable Resend and Cancel fulfilment buttons
	frmScreen.btnResendtoFulfilment.disabled = true;
	frmScreen.btnCancelFulfilment.disabled = true;
<% /* MAR211 End */%>

	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function PopulateCombos()
{	
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("PrintDocumentType");

	if(XML.GetComboLists(document,sGroupList))
	{		
		XML.CreateTagList("LISTENTRY");
			TagOPTION	= document.createElement("OPTION");
			TagOPTION.value	= "";
			TagOPTION.text	= "<SELECT>";
			frmScreen.cboDocumentGroup.add(TagOPTION);
					
			<% /* AS 10/08/2006 EP1080 */ %>
			var geminiScannedDocumentGroup = m_GlobalXML.GetGlobalParameterString(document, "GeminiScannedDocumentGroup").toUpperCase();

			// Loop through all entries and only add relevant entries to combo
			for(var iLoop = 0; iLoop < XML.ActiveTagList.length; iLoop++)
			{
				XML.SelectTagListItem(iLoop);
				var sGroupName	= XML.GetTagText("VALUENAME");
				var sValueID = XML.GetTagText("VALUEID");
				TagOPTION = document.createElement("OPTION");
				TagOPTION.value	= sValueID;
				TagOPTION.text = sGroupName;
				
				frmScreen.cboDocumentGroup.add(TagOPTION);
				
				<% /* AS 10/08/2006 EP1080 */ %>
				if (sGroupName.toUpperCase() == geminiScannedDocumentGroup)
				{
					m_scannedDocumentGroupValueId = sValueID;
				}
			}
						
			TagOPTION	= document.createElement("OPTION");
			TagOPTION.value	= "999";
			TagOPTION.text	= "<ALL>";
			frmScreen.cboDocumentGroup.add(TagOPTION);
		
			frmScreen.cboDocumentGroup.selectedIndex = 0;
	}
	else
	{
		alert("There are currently no documents available for printing");
	}
	XML = null;
}

function displayContextInfo()
{
	//Set up the application number on screen
	frmScreen.txtApplicationNumber.value = m_sApplicationNumber;
	frmScreen.txtStageName.value = m_sStageName;	
}

function RetrieveContextData()
{	
	
	//m_aArgArray = [window.dialogArguments[0] + "px", window.dialogArguments[1] + "px", window.dialogArguments[2] + "px", window.dialogArguments[3] + "px", "1001", 1, "RobinC", "Unit1", 40, "100001645", 1, 1, "Cost Modelling", 1, "CH00006429", null];
	window.dialogTop	= window.dialogArguments[0] + "px";
	window.dialogLeft	= window.dialogArguments[1] + "px";
	window.dialogWidth	= window.dialogArguments[2] + "px";
	window.dialogHeight = window.dialogArguments[3] + "px";
	m_aArgArray = window.dialogArguments[4];
	m_BaseNonPopupWindow = window.dialogArguments[5];

	<% /* AS 10/08/2006 EP1080 */ %>
	m_GlobalXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	<% /* MAR211 TW */ %>
	iDMSReprintButtonAuthorityLevel = m_GlobalXML.GetGlobalParameterAmount(document,'DMSRePrintButtonAuthLevel');
	iDMSCancelButtonAuthorityLevel = m_GlobalXML.GetGlobalParameterAmount(document,'DMSCancelFulfilmentAuthLevel');
	<% /* MAR211End  */ %>
	
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
	
		<%/* AM  16/09/2005 */%>
	m_aCustomerNameArray = m_aArgArray[12];
	m_aCustomerNumberArray = m_aArgArray[13];
	m_aCustomerVersionNumberArray = m_aArgArray[14];
	
	
	<% /* BS BM0271 20/02/03 */ %>
	m_sProcessingIndicator = m_aArgArray[15];
	<% /* BS BM0271 16/04/03 
	if (m_sProcessingIndicator != "1")*/ %>
	m_sReadOnly = m_aArgArray[16];
	if ((m_sProcessingIndicator != "1") || (m_sReadOnly == "1"))	
	{
		<% /* MAR7 GHun */ %>
		if (m_sProcessingIndicator != "1")
			frmScreen.btnPrint.disabled = true;
		<% /* MAR7 End */ %>
		frmScreen.btnEdit.disabled = true;
	}
	<% /* MAR211 TW*/ %>
	<% /* EP529 / MAR1638 PB Use userRole in check and make sure the comparison is numeric
	m_AccessType = parseInt(scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idAccessType", "0")); */ %>
	<% /* PB 25/05/2006 EP613 */ %>
	if (parseInt(m_sRole, 10) < parseInt(iDMSReprintButtonAuthorityLevel, 10))
	<% /*if (m_AccessType < iDMSReprintButtonAuthorityLevel)
	EP613 End */ %>
	<% /* EP529 / MAR1638 End */ %>
		{
		frmScreen.btnPrint.disabled = true;
		}
	<% /* MAR211 End */ %>
	<% /* BS BM0271 End 20/02/03 */ %>
	
	m_sEmailAdministrator = m_aArgArray[17];	<% /* MAR7 GHun */ %>
	
	frmScreen.txtApplicationNumber.value = m_sApplicationNumber;
		
}

function btnExit.onclick()
{	
	window.close();
}


function frmScreen.cboDocumentGroup.onchange()
{
	<% /* MAR211 TW */ %>
	frmScreen.btnResendtoFulfilment.disabled = true;
	frmScreen.btnCancelFulfilment.disabled = true;
	<% /* MAR211 End */ %>
	
	frmScreen.btnRecategorise.disabled = true;	<% /* MAR1332 GHun */ %>

	// m_XML is a global variable and it's set in this function!
	m_XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	if (frmScreen.txtApplicationNumber.value == "")
	{
		alert ("Please enter an Application Number.");
		frmScreen.txtApplicationNumber.focus();
		return;
	} 
	else
	{
		m_sApplicationNumber = frmScreen.txtApplicationNumber.value;
	}
	
	//Preparing XML Request string
	var tagREQUEST = m_XML.CreateActiveTag("REQUEST");
	m_XML.SetAttribute("OPERATION", "GETDOCUMENTHISTORYLIST");
	
	var sSearchKey1 = frmScreen.cboDocumentGroup.value;
	
	if (sSearchKey1 == "")
	{
		alert("Please select a document group.");
		frmScreen.cboDocumentGroup.focus();
		return false;
	}
	
	if (sSearchKey1 != "999")
		m_XML.SetAttribute("SEARCHKEY1", sSearchKey1);
		
	m_XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			m_XML.RunASP(document, "omPMRequest.asp");
			break;
		default: // Error
			m_XML.SetErrorResponse();
		}
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = m_XML.CheckResponse(ErrorTypes);

	if(ErrorReturn[0])
	{
		PopulateScreen();
		<% /* PSC 16/01/2006 MAR1054 */ %>

		<% /* MAR1332 GHun Only auto-select the first row if it contains data */ %>
		if (tblTable.rows(1).getAttribute("DocGuid"))
		{
			scTable.setRowSelected(1);
			spnTable.onclick();		<% /* MAR1332 GHun */ %>
		}
	}
	else
	{
		if(ErrorReturn[1] == ErrorTypes[0])
		{
			alert("Unable to find any documents for your search criteria.");
			PopulateScreen();			
		}
	}
}

function PopulateScreen()
{
	<% //DPF 18/11/2002 - BMIDS00887 - check EditBeforePrinting attribute %>
	frmScreen.btnEdit.disabled = true;
	PopulateListBox(0);
	
}

function PopulateListBox(nStart)
{
	m_XML.CreateTagList("DOCUMENTDETAILS");
	var iNumberOfRows = m_XML.ActiveTagList.length;
	scTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}

function TaskPopulateTable()
{
	scTable.initialiseTable(tblTable, 0, "", TaskShowList, m_iTableLength, "1");
	TaskShowList(0);	
}

function TaskShowList(nStart)
{	
	scScreenFunctions.SizeTextToField(tblTable.rows(1).cells(0),m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME"));
}

function PopulateTable()
{
	FilterXML.ActiveTag = null;
	FilterXML.CreateTagList("DOCUMENTS");
	var iNumberOfTasks = FilterXML.ActiveTagList.length;

	scTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfTasks);
	ShowList(0);	
}

function ShowList(nStart)
{
	var iCount, xmlNodeList, xmlNode;
	<% //DPF 18/11/2002 - BMIDS00887 - added sTemplateId %>
	var sVersion,sDocumentName,sDocumentDescription,sDate,sUserName,sOutputLocation,sCustomerName,sRecipient,sFileGuid,sDocGuid,sTemplateId;

	<% //SG 23/04/02 SYS4407 %>
	var sEventDate = "";

	scTable.clear();
	tblTable.rows(1).removeAttribute("DocGuid");	<% /* MAR1332 GHun */ %>	
	
	for(iCount = 0; iCount < m_XML.ActiveTagList.length  && iCount < m_iTableLength; iCount++)
	{
		m_XML.SelectTagListItem(nStart+iCount);

		sVersion = m_XML.GetAttribute("DOCUMENTVERSION");
		sDocumentName = m_XML.GetAttribute("DOCUMENTNAME");
		sDocumentDescription = m_XML.GetAttribute("DOCUMENTDESCRIPTION");
				
		//SG 23/04/02 SYS4407	
		//sDate = m_XML.GetAttribute("CREATETIMESTAMP");
		sDate = m_XML.GetAttribute("EVENTDATE");
		
		sUserName = m_XML.GetAttribute("USERNAME");		
		sOutputLocation = m_XML.GetAttribute("PRINTLOCATION");
		<% /* EP507 PB */ %>
		var iEventKey = "0";
		iEventKey = m_XML.GetAttribute("EVENTKEY")
		if (iEventKey == EVENTKEY_EMAIL)
		{
			sOutputLocation = "EMAIL";
		}

		if (iEventKey == EVENTKEY_SMS)
		{
			sOutputLocation = "SMS";
		}
		<% /* EP507 End */ %>
		sCustomerName = m_XML.GetAttribute("CUSTOMERNAME");		
		sRecipient = m_XML.GetAttribute("RECIPIENTNAME");

		sFileGuid = m_XML.GetAttribute("FILEGUID");
		sDocGuid = m_XML.GetAttribute("DOCUMENTGUID");
		//DPF 18/11/2002 - BMIDS00887 - added sTemplateId
		sTemplateId = m_XML.GetAttribute("HOSTTEMPLATEID");

		var preventEdit = m_XML.GetAttribute("PREVENTEDITINDMS");
		
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sVersion);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sDocumentName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sDate);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), sUserName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4), sOutputLocation);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5), sCustomerName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(6), sRecipient);
		
		tblTable.rows(iCount+1).setAttribute("Version", sVersion);	<% /* MAR7 GHun */ %>
		tblTable.rows(iCount+1).setAttribute("FileGuid", sFileGuid);
		tblTable.rows(iCount+1).setAttribute("DocGuid", sDocGuid);
		tblTable.rows(iCount+1).setAttribute("CustomerName", sCustomerName);
		tblTable.rows(iCount+1).setAttribute("RecipientName", sRecipient);		
		tblTable.rows(iCount+1).setAttribute("StageName", m_XML.GetAttribute("STAGEID"));
		tblTable.rows(iCount+1).setAttribute("DocType", m_XML.GetAttribute("DOCUMENTTYPE"));
		tblTable.rows(iCount+1).setAttribute("DocumentName", sDocumentName);
		tblTable.rows(iCount+1).setAttribute("DocumentDescription", sDocumentDescription);				
		//DPF 18/11/2002 - BMIDS00887 - added sTemplateId
		tblTable.rows(iCount+1).setAttribute("HostTemplateId", sTemplateId);
		<% /* MAR7 GHun */ %>
		tblTable.rows(iCount+1).setAttribute("OutputLocation", sOutputLocation);
		tblTable.rows(iCount+1).setAttribute("ArchiveDate", m_XML.GetAttribute("ARCHIVEDATE"));
		tblTable.rows(iCount+1).setAttribute("DocumentGroup", m_XML.GetAttribute("DOCUMENTGROUP"));
		<%/*AM 21st Sept 2005 MARS - */%>
		tblTable.rows(iCount+1).setAttribute("DocumentPurpose", m_XML.GetAttribute("DOCUMENTPURPOSE"));
		tblTable.rows(iCount+1).setAttribute("DocumentLanguage", m_XML.GetAttribute("LANGUAGE"));
		tblTable.rows(iCount+1).setAttribute("PrintDate", m_XML.GetAttribute("PRINTDATE"));
		tblTable.rows(iCount+1).setAttribute("SourceSystem", m_XML.GetAttribute("SOURCESYSTEM"));
		tblTable.rows(iCount+1).setAttribute("SearchKey1", m_XML.GetAttribute("SEARCHKEY1"));
		tblTable.rows(iCount+1).setAttribute("SearchKey2", m_XML.GetAttribute("SEARCHKEY2"));
		tblTable.rows(iCount+1).setAttribute("SearchKey3", m_XML.GetAttribute("SEARCHKEY3"));
		tblTable.rows(iCount+1).setAttribute("TemplateId", m_XML.GetAttribute("TEMPLATEID"));

		<% /* MAR211 TW */ %>
		tblTable.rows(iCount+1).setAttribute("EventKey", m_XML.GetAttribute("EVENTKEY"));
		tblTable.rows(iCount+1).setAttribute("EventDate", m_XML.GetAttribute("EVENTTIMESTAMP"));
		tblTable.rows(iCount+1).setAttribute("EventPackFulfilmentGUID", m_XML.GetAttribute("PACKFULFILLMENTGUID"));
		
		<% /* TW MAR581 19/11/2005 */ %>

		tblTable.rows(iCount+1).setAttribute("FileNetGUID",m_XML.GetAttribute("FILENETIMAGEREF"));
		<% /* TW MAR581 19/11/2005 End */ %>

		<% /* MAR211 End */ %>

		<% /* MAR7 End */ %>
		//MC PreventEditInDMS flag
		tblTable.rows(iCount+1).setAttribute("PREVENTEDITINDMS", preventEdit);
		<% /* MAR1332 GHun */ %>
		tblTable.rows(iCount+1).setAttribute("INBOUNDDOCUMENT", m_XML.GetAttribute("INBOUNDDOCUMENT"));
	}
}

function frmScreen.btnView.onclick()
{
	var iCurrentRow = scTable.getRowSelected();	
	if(iCurrentRow == -1)
	<% /* MAR7 GHun */ %>
	{
		alert("Select a document for viewing.");
		<% /* EP462 IK */ %>
		return;
	}

	var readOnly = true;
	var printOnly = false;
	<% /* TW MAR581 19/11/2005 */ %>
	var prXML = null
	var documentHostTemplateID = tblTable.rows(iCurrentRow).getAttribute("HostTemplateId");
	var documentName = tblTable.rows(iCurrentRow).getAttribute("DocumentName");
	var fileNetGuid = tblTable.rows(iCurrentRow).getAttribute("FileNetGuid");
	if (fileNetGuid != "")
	{
		<% /* PSC 16/01/2006 MAR1054 - Start */ %>
		sFileNetUserId = m_GlobalXML.GetGlobalParameterString(document,'FileNetUserId');
		prXML = getDocumentFromFileNet(tblTable, iCurrentRow, sFileNetUserId);
		<% /* PSC 16/01/2006 MAR1054 - End */ %>
	}	
	else
	{
		prXML = getArchivedDocumentForRow(tblTable, iCurrentRow, readOnly, printOnly);
	}
	<% /* TW MAR581 19/11/2005 End */ %>

	<% /* EP462 AS */ %>
	if (prXML != null && m_fileContents != null && m_fileContents != "")
	{
		<% /* TW MAR581 19/11/2005 */ %>
		var documentHostTemplateID = tblTable.rows(iCurrentRow).getAttribute("HostTemplateId");
		var documentName = tblTable.rows(iCurrentRow).getAttribute("DocumentName");
		<% /* TW MAR581 19/11/2005 End */ %>

		<% /* PSC MAR1197 06/02/2006 */ %>
		var printDocumentData = displayArchivedDocument(m_fileContents, documentHostTemplateID, documentName, m_archiveDeliveryType, m_sPrinterType, readOnly, printOnly, m_archiveFileExtension);
		if (printDocumentData != null && printDocumentData.get_success())
		{
			frmScreen.cboDocumentGroup.onchange();	
		}
	}
	else if (prXML != null && m_fileContentsUrl != null && m_fileContentsUrl != "")
	{
		<% /* AS 04/08/2006 EP1068 omGemini: Add support for scanned documents */ %>
		window.open(m_fileContentsUrl, "", "");
	}
	<% /* MAR7 End */ %>
}

function frmScreen.btnEdit.onclick()
{
	var iCurrentRow = scTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		alert("Select a document for editing.");
		return;
	<% /* MAR7 GHun */ %>
	}

	var readOnly = false;
	var printOnly = false;
	var prXML = getArchivedDocumentForRow(tblTable, iCurrentRow, readOnly, printOnly);

	<% /* EP462 AS */ %>
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
<% /* MAR7 End */ %>

			prXML.RemoveActiveTag();
			var tagREQUEST = prXML.CreateActiveTag("REQUEST");
			// ik_bm0037
			if(printDocumentData.get_fileEdited())	<% /* MAR7 GHun */ %>
				prXML.SetAttribute("OPERATION", "SAVEDOCUMENT");
			else
				prXML.SetAttribute("OPERATION", "UNLOCKDOCUMENT");
			// ik_bm0037_ends
			prXML.SetAttribute("USERID", m_sUserId);
			prXML.SetAttribute("UNITID", m_sUnitId);
			prXML.SetAttribute("MACHINENAME", m_sMachineId);
			prXML.SetAttribute("CHANNELID", m_sDistributionChannelId);
			prXML.SetAttribute("USERAUTHORITYLEVEL", "10");

			// ik_bm0037
			if(printDocumentData.get_fileEdited())	<% /* MAR7 GHun */ %>
			{
				prXML.SetAttribute("EDITDOCUMENT", "TRUE");
				prXML.SetAttribute("DOCUMENTGUID", sDocGUID);
				prXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
			}
			// ik_bm0037_ends

			var tagPRINTDATA = prXML.CreateTag("PRINTDOCUMENTDATA");
			<% /* MAR7 GHun */ %>
			prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "DELIVERYTYPE", m_archiveDeliveryType);
			prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "COMPRESSIONMETHOD", "");	<% /* MAR796 GHun turn off compression */ %>
			prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "COMPRESSED", "0");
			<% /* MAR7 End */ %>
			prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "FILEGUID", sFileGUID);
			prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "FILEVERSION", sVersion);

			// ik_bm0037
			if(printDocumentData.get_fileEdited())
			{
				prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "FILESIZE", printDocumentData.get_fileSize());	<% /* MAR7 GHun */ %>
				prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "FILECONTENTS_TYPE", "BIN.BASE64");
				prXML.SetAttributeOnTag("PRINTDOCUMENTDATA", "FILECONTENTS", printDocumentData.get_fileContents());	<% /* MAR7 GHun */ %>
			}
			// ik_bm0037_ends
		
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							prXML.RunASP(document, "omPMRequest.asp");
					break;
				default: // Error
					prXML.SetErrorResponse();
				}
			prXML.IsResponseOK();	<% /* MAR7 GHun */ %>
			
			frmScreen.cboDocumentGroup.onchange();
		}
	}
}

function frmScreen.btnPrint.onclick()
{
	var iCurrentRow = scTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
	<% /* MAR7 GHun */ %>
		alert("Select a document for printing.");
		return;
	}

	if (m_sPrinterType == "W" )
	{
		var readOnly = true;
		var printOnly = true;

		<% /* TW MAR581 19/11/2005 */ %>
		// var prXML = getArchivedDocumentForRow(tblTable, iCurrentRow, readOnly, printOnly);
		var prXML = null
		var fileNetGuid = tblTable.rows(iCurrentRow).getAttribute("FileNetGuid");
		if (fileNetGuid != "")
			{
				<% /* PSC 16/01/2006 MAR1054 - Start */ %>
				sFileNetUserId = m_GlobalXML.GetGlobalParameterString(document,'FileNetUserId');
				prXML = getDocumentFromFileNet(tblTable, iCurrentRow, sFileNetUserId);
				<% /* PSC 16/01/2006 MAR1054 - End */ %>
			}	
		else
			{
			prXML = getArchivedDocumentForRow(tblTable, iCurrentRow, readOnly, printOnly);
			}
		<% /* TW MAR581 19/11/2005 End */ %>
	
		<% /* EP462 AS */ %>
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
				frmScreen.cboDocumentGroup.onchange();
			}
			else
			{
				alert("Error sending document to the printer.");
			}
		}
		else if (prXML != null && m_fileContentsUrl != null && m_fileContentsUrl != "")
		{
			<% /* AS 04/08/2006 EP1068 omGemini: Add support for scanned documents */ %>
			window.open(m_fileContentsUrl, "", "");
		}
	}
	else
	{
		<% /* TW MAR581 19/11/2005 */
//		var fileNetGuid = tblTable.rows(iCurrentRow).getAttribute("FileNetGuid");
//		if (fileNetGuid != "")
//		{
//			printFileNetDocument(tblTable, iCurrentRow);
//			return;
//		}
		/* TW MAR581 19/11/2005 End */ %>
		<% /* MAR7 End */ %>
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

		//	ik_BMIDS00885, re-print to default (local) printer
		<% /* MAR7 GHun */ %>
		m_xmlLocalPrinters = getLocalPrinters(m_xmlLocalPrinters);
		if(m_xmlLocalPrinters != null)
		{		
			<% /* MAR7 End */ %>	
			var LocalPrintersXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			LocalPrintersXML.LoadXML(m_xmlLocalPrinters.xml);	<% /* MAR7 GHun */ %>
						
			LocalPrintersXML.CreateTagList("PRINTER[DEFAULTINDICATOR='1']");
								
			if(LocalPrintersXML.ActiveTagList.length == 0)
			{					
				alert("You do not have a default printer set on your PC.");	
			}
			else	<% /* MAR7 GHun */ %>
			{
				var sDefaultPrinter = LocalPrintersXML.ActiveTagList.item(0).selectSingleNode("PRINTERNAME").text;
				
				prXML.ActiveTag = tagREQUEST;
				prXML.CreateActiveTag("CONTROLDATA");
				prXML.SetAttribute("DESTINATIONTYPE", "L");
				<% /* MAR7 GHun */ %>
				prXML.SetAttribute("DELIVERYTYPE", m_sDeliveryType);
				prXML.SetAttribute("TEMPLATEID", m_sTemplateID);
				prXML.SetAttribute("FIRSTPAGEPRINTERTRAY", m_sFirstPageTray);
				prXML.SetAttribute("OTHERPAGESPRINTERTRAY", m_sOtherPagesTray);
				<% /* MAR7 End */ %>
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
				
				frmScreen.cboDocumentGroup.onchange();	// Refresh the screen
			}
		}
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
	// Set required values in context and call DMS107
	var XML107 = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();	
	var tagREQUEST = XML107.CreateActiveTag("REQUEST");
	
	XML107.SetAttribute("OPERATION","GETDOCUMENTEVENTLIST");
	XML107.SetAttribute("DOCUMENTGUID",tblTable.rows(iCurrentRow).getAttribute("DocGuid"));
	XML107.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML107.SetAttribute("DOCUMENTNAME",tblTable.rows(iCurrentRow).getAttribute("DocumentName"));
	XML107.SetAttribute("DOCUMENTDESCRIPTION",tblTable.rows(iCurrentRow).getAttribute("DocumentDescription"));	
	XML107.SetAttribute("DOCUMENTTYPE",tblTable.rows(iCurrentRow).getAttribute("DocType"));
	
	//DR DMSSYS0003 You cannot use the innerText values - they have are sometimes shortened with ...
	XML107.SetAttribute("RECIPIENTNAME",tblTable.rows(iCurrentRow).getAttribute("RecipientName"));
	XML107.SetAttribute("CUSTOMERNAME",tblTable.rows(iCurrentRow).getAttribute("CustomerName"));
	XML107.SetAttribute("STAGEID",tblTable.rows(iCurrentRow).getAttribute("StageName"));
	<% /* MAR7 GHun */ %>
	XML107.SetAttribute("HOSTTEMPLATEID",tblTable.rows(iCurrentRow).getAttribute("HostTemplateId"));
	XML107.SetAttribute("PRINTERTYPE",m_sPrinterType);
	<% /* MAR7 End */ %>
	//JD MAR838
	XML107.SetAttribute("FILENETGUID",tblTable.rows(iCurrentRow).getAttribute("FileNetGuid"));

	//DR Save the XML in the dialog argument array cos the om context parameters are unavailable.
	m_aArgArray[11] = XML107.XMLDocument.xml;
		
	//Create a Events popup
	var nWidth = window.dialogArguments[2].replace("px","");
	var nHeight = window.dialogArguments[3].replace("px","");
	var ArrayArguments = m_aArgArray;

	scScreenFunctions.DisplayPopup(window, document, "DMS107.asp", ArrayArguments, nWidth, nHeight);

}

<% /*AM  13/09/05 - MARS -Routing to Reconfigurable Screen */> %>
function frmScreen.btnRecategorise.onclick()
{	
	var iCurrentRow = scTable.getRowSelected();
	if(iCurrentRow == -1)
	{
		alert("Select a document to Re-categorise.");
		return;
	}
		
	// Set required values in context and call DMS109
	var XML109 = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();	
	var tagREQUEST = XML109.CreateActiveTag("REQUEST");
	var ArrayArguments = m_aArgArray;

	XML109.SetAttribute("DOCUMENTGUID", tblTable.rows(iCurrentRow).getAttribute("DocGuid"));
	XML109.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML109.SetAttribute("DOCUMENTNAME", tblTable.rows(iCurrentRow).getAttribute("DocumentName"));
	XML109.SetAttribute("DOCUMENTGROUP", tblTable.rows(iCurrentRow).getAttribute("DocumentGroup"));
	XML109.SetAttribute("DOCUMENTTYPE", tblTable.rows(iCurrentRow).getAttribute("DocType"));
	
	<% /* MAR1332 GHun */ %>
	XML109.CreateActiveTag("EVENTDETAIL");
	XML109.SetAttribute("DOCUMENTVERSION", tblTable.rows(iCurrentRow).cells(0).innerText);
	XML109.SetAttribute("EVENTKEY", EVENTKEY_RECATEGORISATION);
	XML109.SetAttribute("FILEGUID", tblTable.rows(iCurrentRow).getAttribute("FileGuid"));
	var packGuid = tblTable.rows(iCurrentRow).getAttribute("EventPackFulfilmentGUID");
	if (packGuid.length == 0)
		packGuid = "00000000000000000000000000000000";
	XML109.SetAttribute("PACKFULFILLMENTGUID", packGuid);
	XML109.SetAttribute("FILENETIMAGEREF", tblTable.rows(iCurrentRow).getAttribute("FileNetGUID"));
	<% /* MAR1332 End */ %>
	
	// Save the XML in the dialog argument array cos the om context parameters are unavailable.		
	
	var m_aCustomerNameArray = new Array();
	var m_aCustomerNumberArray = new Array();
	var m_aCustomerVersionNumberArray = new Array();			
	var iCount = 0;
	var sCustomerName = "";
	var sCustomerNumber = "";
	var sCustomerVersionNumber = "";
	for (iCount = 1; iCount <= 5; iCount++)
	{								
		sCustomerNumber = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerNumber" + iCount,null);
		if (sCustomerNumber != "") 
		{				
			sCustomerName = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerName" + iCount,null);
			sCustomerVersionNumber = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerVersionNumber" + iCount,null);
			m_aCustomerNameArray[iCount-1] = sCustomerName;
			m_aCustomerNumberArray[iCount-1] = sCustomerNumber;
			m_aCustomerVersionNumberArray[iCount-1] = sCustomerVersionNumber;			
		}
	}	
	ArrayArguments[12] = m_aCustomerNameArray;
	ArrayArguments[13] = m_aCustomerNumberArray;
	ArrayArguments[14] = m_aCustomerVersionNumberArray;
	m_aArgArray[11] = XML109.XMLDocument.xml;
	ArrayArguments[15] = tblTable.rows(iCurrentRow).getAttribute("DocumentName");
	ArrayArguments[16] = tblTable.rows(iCurrentRow).getAttribute("DocumentGroup");
	ArrayArguments[17] = tblTable.rows(iCurrentRow).getAttribute("DocumentPurpose");
	//Create a Events popup
	var ArrayArguments = m_aArgArray;
	
	<% /* MAR1332 GHun 
	var sReturn = scScreenFunctions.DisplayPopup(window, document, "DMS109.asp", m_aArgArray, 360, 300); */ %>
	<% /* MAR1578 PE */ %>
	var sReturn = scScreenFunctions.DisplayPopup(window, document, "DMS109.asp", m_aArgArray, 470, 300);
	
	if (sReturn == 1)
		frmScreen.cboDocumentGroup.onchange();
	<% /* MAR1332 End */ %>
}

<% //DPF 18/11/2002 - BMIDS00887 - check EditBeforePrinting attribute %>
function CheckAttributes(sDocID, outputLocation)	<% /* MAR7 GHun */ %>
{	
	var m_sEditBeforePrinting = 0;
	AttribsXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var tagREQUEST = AttribsXML.CreateActiveTag("REQUEST");
	AttribsXML.SetAttribute("OPERATION", "GetPrintAttributes");
	//AttribsXML.CreateRequestTagFromArray(m_sRequestAttribs, "GetPrintAttributes");
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
		<% /* MAR7 GHun */ %>
		m_sDeliveryType = AttribsXML.GetTagAttribute("ATTRIBUTES", "DELIVERYTYPE");
		m_sTemplateID = AttribsXML.GetTagAttribute("ATTRIBUTES", "DPSTEMPLATEID");
		m_sFirstPageTray = AttribsXML.GetTagAttribute("ATTRIBUTES", "FIRSTPAGEPRINTERTRAY");
		m_sOtherPagesTray = AttribsXML.GetTagAttribute("ATTRIBUTES", "OTHERPAGESPRINTERTRAY");
		m_sPrinterDestinationType = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
		m_sRemotePrinterLocation = AttribsXML.GetTagAttribute("ATTRIBUTES", "REMOTEPRINTERLOCATION");
		m_sPrinterType = getPrinterType(m_BaseNonPopupWindow, m_sPrinterDestinationType, outputLocation, m_sRemotePrinterLocation);
		<% /* MAR7 End */ %>
	}

	<% /* MAR7 GHun
	if (m_sEditBeforePrinting == "1")
	*/ %>
	if (m_sEditBeforePrinting == "1" && m_sReadOnly != "1" && m_sProcessingIndicator == "1")
	<% /* MAR7 End */ %>
	{
		<% /* MAR7 GHun */ %>
		var preventEditInDMS = AttribsXML.GetTagAttribute("ATTRIBUTES", "PREVENTEDITINDMS");

		if(preventEditInDMS == "1")
			frmScreen.btnEdit.disabled = true;
		else
			frmScreen.btnEdit.disabled = false;
		<% /* MAR7 End */ %>
	}
	else
	{
		frmScreen.btnEdit.disabled = true;
	}
}

function spnTable.onclick()
{

	<% /* MAR7 GHun */ %>
	var iRowSelected = scTable.getRowSelected();

<% /* MAR211 TW */ %>
	var iEventKey = "0";
<% /* MAR211 End */ %>

	var bCancelledExists = false;
	var bInboundDocument = false;	<% /* MAR1332 GHun */ %>

	if (iRowSelected > 0 && iRowSelected != null)
	{
		<% /* PSC MAR1233 07/02/2006 */ %>
		var hostTemplateId = tblTable.rows(iRowSelected).getAttribute("HostTemplateId");	<% /* MAR1332 GHun */ %>
		if (hostTemplateId && (hostTemplateId != ""))	
			CheckAttributes(tblTable.rows(iRowSelected).getAttribute("HostTemplateId"), 
					tblTable.rows(iRowSelected).getAttribute("OutputLocation"));
<% /* MAR211 TW */ %>
		iEventKey = tblTable.rows(iRowSelected).getAttribute("EventKey");
<% /* MAR211 End */ %>

		<% /* MAR1309 GHun Only call GetDocumentEventList if a row is selected */ %>
		<% /* MAR883  Enable Resend button depending on events.
	      If the latest event is SendFulfilment or ResendFulfilment and 
	      no CancelFulfilment exists in the list of events for this document then enable the Resend button */ %>
	
		var EventXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

		// Prepare XML Request string
		var tagREQUEST = EventXML.CreateActiveTag("REQUEST");
		
		EventXML.SetAttribute("OPERATION","GETDOCUMENTEVENTLIST");
		EventXML.SetAttribute("DOCUMENTGUID",tblTable.rows(iRowSelected).getAttribute("DocGuid"));
		EventXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
			
		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				EventXML.RunASP(document, "omPMRequest.asp");
				break;
			default: // Error
				EventXML.SetErrorResponse();
		}

		EventXML.CreateTagList("DOCUMENTDETAILS");
		
		for(iCount = 0; iCount < EventXML.ActiveTagList.length; iCount++)
		{
			EventXML.SelectTagListItem(iCount);	
			
			if (EventXML.GetAttribute("EVENT") == "FulfilmentCancel")
				bCancelledExists = true;
		}
		
		<% /* MAR1332 GHun */ %>
		if (tblTable.rows(iRowSelected).getAttribute("INBOUNDDOCUMENT") == "1")
			bInboundDocument = true;
		<% /* MAR1332 End */ %>
	}
	
	<%/* MAR1332 GHun */%>
	if (bInboundDocument)
	{
		frmScreen.btnRecategorise.disabled = false;
	}
	else 
	{
		frmScreen.btnRecategorise.disabled = true;		
	}
	<% /* MAR1332 End */ %>

<% /* MAR211 TW */ %>
	if ((iEventKey == EVENTKEY_FULFILMENTSEND || iEventKey == EVENTKEY_FULFILMENTRESEND) && bCancelledExists == false)
	{
		<%/* EP529 PB */%>
		<% /* frmScreen.btnCancelFulfilment.disabled = (m_AccessType < iDMSCancelButtonAuthorityLevel); */ %>
		frmScreen.btnCancelFulfilment.disabled = (m_sRole != iDMSCancelButtonAuthorityLevel);
		<%/* EP529 End */%>
		frmScreen.btnResendtoFulfilment.disabled = false;
	}
	else
	{
		frmScreen.btnResendtoFulfilment.disabled = true;
	}

<% /* MAR211 End */ %>

<% /* TW MAR581 21/11/2005 */ %>
	if (iRowSelected > 0 && iRowSelected != null)
	{
		var fileGuid = tblTable.rows(iRowSelected).getAttribute("FileGuid");
		if(fileGuid == "" || fileGuid == "00000000000000000000000000000000")
		{
			frmScreen.btnPrint.disabled = true;
			frmScreen.btnEdit.disabled = true;  //MAR837
		}
		
		<% /* AS 10/08/2006 EP1080 omGemini: DMS105 disable reprint button for scanned documents */ %>
		var documentGroup = tblTable.rows(iRowSelected).getAttribute("DocumentGroup");
		if (documentGroup == m_scannedDocumentGroupValueId)
		{
			frmScreen.btnPrint.disabled = true;
			frmScreen.btnEdit.disabled = true;
		}	
		
		<% /* EP507 PB */ %>
		if (iEventKey == EVENTKEY_EMAIL || iEventKey == EVENTKEY_SMS)
		{
			frmScreen.btnView.disabled = true;
		}
		else
		{
			frmScreen.btnView.disabled = false;
		}
		<% /* EP507 End */ %>
	}
<% /* TW MAR581 21/11/2005 End */ %>

	// MAR1378 - Not all screens are frozen when the case is in the decline stage
	// Peter Edney - 25/03/2006	
	if(scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idCancelDeclineFreezeDataIndicator", "0"))
	{
		frmScreen.btnResendtoFulfilment.disabled = true;
	}

}


<% /* MAR7 GHun */ %>
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
	// Both ARCHIVETIMESTAMP & ARCHIVEDATE are used in omPM.
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
	// Both PRINTTIMESTAMP & PRINTDATE are used in omPM.
	xmlElement.setAttribute("PRINTTIMESTAMP", tblTable.rows(iCurrentRow).getAttribute("PrintDate"));
	xmlElement.setAttribute("PRINTDATE", tblTable.rows(iCurrentRow).getAttribute("PrintDate"));
	xmlElement.setAttribute("RECIPIENTNAME", tblTable.rows(iCurrentRow).getAttribute("RecipientName"));
	xmlElement.setAttribute("STAGEID", tblTable.rows(iCurrentRow).getAttribute("StageName"));
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

<% /* MAR7 End */ %>

<% /* MAR211 TW */ %>

function frmScreen.btnResendtoFulfilment.onclick()
{
	var iCurrentRow = scTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		alert("Select a document for Re-sending to Fulfilment");
	}
	else
	{
		SetUpPackRequest(iCurrentRow, EVENTKEY_FULFILMENTRESEND);
	}
	return;
}

function frmScreen.btnCancelFulfilment.onclick()
{
	var iCurrentRow = scTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		alert("Select a document to Cancel Fulfilment");
	}
	else
	{
		SetUpPackRequest(iCurrentRow, EVENTKEY_FULFILMENTCANCEL);
	}
	return;
}

function SetUpPackRequest(iCurrentRow, iEventKey)
{
	var success = false;
		
	var documentVersion = tblTable.rows(iCurrentRow).cells(0).innerText;
	var fileGuid = tblTable.rows(iCurrentRow).getAttribute("FileGuid");
	var documentGuid = tblTable.rows(iCurrentRow).getAttribute("DocGuid");
	var documentHostTemplateID = tblTable.rows(iCurrentRow).getAttribute("HostTemplateId");
	var documentEventDate = tblTable.rows(iCurrentRow).getAttribute("EventDate");
	var documentPackFulfilmentGUID = tblTable.rows(iCurrentRow).getAttribute("EventPackFulfilmentGUID");
	var fileNetGuid = tblTable.rows(iCurrentRow).getAttribute("FileNetGuid");     // MAR883
	
	var xmlRequestObj = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequestDocument = xmlRequestObj.XMLDocument;
	
	var xmlRequest = xmlRequestDocument.createElement("REQUEST");
	var xmlElement = xmlRequestDocument.createElement("EVENTDETAIL");
	
	xmlRequestDocument.appendChild(xmlRequest);
	xmlRequest.appendChild(xmlElement);

	switch (iEventKey)
	{
		case EVENTKEY_FULFILMENTRESEND: // Resend
			xmlRequest.setAttribute("OPERATION", "RESENDPACK");	
			break;
		case EVENTKEY_FULFILMENTCANCEL: // Cancel
			xmlRequest.setAttribute("OPERATION", "CANCELPACK");	
			break;
		default: // Error
			xmlRequestObj.SetErrorResponse();
	}
	xmlRequest.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	xmlRequest.setAttribute("DOCUMENTGUID", documentGuid);
	xmlRequest.setAttribute("UNITID", m_sUnitId);
	xmlRequest.setAttribute("USERID", m_sUserId);

	xmlElement.setAttribute("DOCUMENTVERSION", documentVersion);
	xmlElement.setAttribute("EVENTTIMESTAMP", documentEventDate);
	xmlElement.setAttribute("FILEGUID", fileGuid);
	xmlElement.setAttribute("EVENTKEY", iEventKey);	
	xmlElement.setAttribute("HOSTTEMPLATEID", documentHostTemplateID);
	xmlElement.setAttribute("PACKFULFILLMENTGUID", documentPackFulfilmentGUID);
	xmlElement.setAttribute("FILENETIMAGEREF", fileNetGuid);    // MAR883
	
	switch (ScreenRules())
	{
	case 1: <% // Warning %>
	case 0: <% // OK %>
		xmlRequestObj.RunASP(document, "omPackRequest.asp");
		break;
	default: <% // Error %>
		xmlRequestObj.SetErrorResponse();
	}
	xmlRequestObj.SelectTag(null, "RESPONSE");
	if (xmlRequestObj.IsResponseOK()) 
	{
		success = true;
	}
	
	return success;
}

<% /* MAR211 TW End */ %>

<% /* TW MAR581 19/11/2005 */ %>
<% /* JD MAR838 FUNCTION MOVED TO DOCUMENT.JS : function getDocumentFromFileNet(tblTable, iCurrentRow) */ %>


function printFileNetDocument(tblTable, iCurrentRow)
{

<% /* TW This code is incomplet. It will need more work to make it possible to print data returned from 
	 FileNet. */ %>
	var readOnly = false;
	var printOnly = false;

	<% /* PSC 16/01/2006 MAR1054 - Start */ %>
	sFileNetUserId = m_GlobalXML.GetGlobalParameterString(document,'FileNetUserId');

	var prXML = getDocumentFromFileNet(tblTable, iCurrentRow, sFileNetUserId);
	<% /* PSC 16/01/2006 MAR1054 - End */ %>
	
	if (prXML == null)
	{
		alert("No document returned from FileNet");
	}


	var documentHostTemplateID = tblTable.rows(iCurrentRow).getAttribute("HostTemplateId");
	var documentName = tblTable.rows(iCurrentRow).getAttribute("DocumentName");
	var sVersion = tblTable.rows(iCurrentRow).cells(0).innerText;
	var sFileGUID = tblTable.rows(iCurrentRow).getAttribute("FileGuid");
	var sDocGUID = tblTable.rows(iCurrentRow).getAttribute("DocGuid");

	var xmlRequestObj = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequestDocument = xmlRequestObj.XMLDocument;

	var xmlRequest = xmlRequestDocument.createElement("REQUEST");
	xmlRequestDocument.appendChild(xmlRequest);

	xmlRequest.setAttribute("OPERATION", "PRINTDOCUMENT");
	xmlRequest.setAttribute("USERID", m_sUserId);
	xmlRequest.setAttribute("UNITID", m_sUnitId);
	xmlRequest.setAttribute("MACHINENAME", m_sMachineId);
	xmlRequest.setAttribute("CHANNELID", m_sDistributionChannelId);
	xmlRequest.setAttribute("USERAUTHORITYLEVEL", "10");
	xmlRequest.setAttribute("FREEFORMAT", "1");

	var xmlDocumentDetails = xmlRequestDocument.createElement("DOCUMENTDETAILS");
	xmlRequest.appendChild(xmlDocumentDetails);
	xmlDocumentDetails.setAttribute("EMAILADMINISTRATOR", ""); 

	var xmlDocumentContents = xmlRequestDocument.createElement("DOCUMENTCONTENTS");
	xmlRequest.appendChild(xmlDocumentContents);
	xmlDocumentContents.setAttribute("DELIVERYTYPE", m_archiveDeliveryType);
	xmlDocumentContents.setAttribute("COMPRESSIONMETHOD", "");
	xmlDocumentContents.setAttribute("COMPRESSED", "0");
	xmlDocumentContents.setAttribute("FILEGUID", sFileGUID);
	xmlDocumentContents.setAttribute("FILEVERSION", sVersion);
	xmlDocumentContents.setAttribute("FILESIZE", "");
	xmlDocumentContents.setAttribute("FILECONTENTS_TYPE", "BIN.BASE64");
	xmlDocumentContents.setAttribute("FILECONTENTS", m_fileContents);

	var xmlControlData = xmlRequestDocument.createElement("CONTROLDATA");
	xmlRequest.appendChild(xmlControlData);

	var xmlOutputType = xmlRequestDocument.createElement("OUTPUTTYPE");
	xmlControlData.appendChild(xmlOutputType);

	var xmlPrinter = xmlRequestDocument.createElement("PRINTER");
	xmlOutputType.appendChild(xmlPrinter);

	var xmlPrinterName = xmlRequestDocument.createElement("PRINTERNAME");
	xmlPrinter.appendChild(xmlPrinterName);

	var xmlCopies = xmlRequestDocument.createElement("COPIES");
	xmlPrinter.appendChild(xmlCopies);
	xmlCopies.Text = "1";
		
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			xmlRequestObj.RunASP(document, "omPMRequest.asp");
			break;
		default: // Error
			xmlRequestObj.SetErrorResponse();
		}
	xmlRequestObj.IsResponseOK();
	frmScreen.cboDocumentGroup.onchange();
}
<% /* TW MAR581 19/11/2005 End */ %>
-->
</script>
</body>
</html>
