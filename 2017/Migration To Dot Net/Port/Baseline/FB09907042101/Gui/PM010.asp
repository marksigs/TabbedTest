<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      PM010.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Prints a variety of documents
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :

Prog	Date		AQR			Description
GHun	15/07/2005	MAR7		Integrate local printing
TW		21/11/2005	MAR579		Put pdf documents in FileNet
TW		29/11/2005	MAR716		MAR579 now done in omPM
PE		05/01/2006	MAR906		Set CustomerNumber and CustomerVersionNumber to thir correct values.
PSC		27/01/2006	MAR1116		Amend btnPrint.onclick to print editted document if it has changed
PSC		30/01/2006	MAR1147		Amend btnPrint.onclick to set return to a string rather than an array
RF      16/02/2006  MAR1251     Allow printer destination of "Document Store Only"
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History :

Prog	Date		AQR		Description
PB		13/07/2006	EP986	Make sure m_sCustomerNumber and m_sCustomerVersionNumber are not null before assigning to XML
PB		17/07/2006	EP986	Added additional checking to above.
AS		12/12/2006	EP1262	Gemini Printing.
AS		15/12/2006	EP1269	Gemini Printing: second release to system test.
GHun	19/02/2007	EP2_1903	Add support for Email Fulfilment
GHun	19/03/2007	EP2_2026	Only create DOCUMENTLOCATION if there is a value for it
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title>Print Menu <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<span style="TOP: 220px; LEFT: 210px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 245px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">
	<span id="spnDocGroup" style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Document Group
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<select id="cboDocumentGroup" name="DocumentGroup" title="" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	</span>
	<div id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 45px">
		<table id="tblTable" width="502" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			
			<tr id="rowTitles">	<td width="99%" class="TableHead">Document Name</td><td width="1%" class="TableHead"></td></tr>
			<tr id="row01">	<td class="TableTopLeft">&nbsp;</td><td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">	<td class="TableLeft">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">	<td class="TableBottomLeft">&nbsp;</td><td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</div>
</div>

<div id="divAttributesBackground" style="TOP: 260px; LEFT: 10px; HEIGHT: 140px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">

	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Print Attributes</strong>
	</span>
	<span style="TOP: 30px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<span id="spnOutputLocation" class="msgLabel">
		Output Location
		</span>
		<span style="TOP: -3px; LEFT: 155px; POSITION: ABSOLUTE">
			<select id="cboOutputLocation" name="OutputLocation" title="" style="WIDTH: 340px" class="msgCombo"></select>
			<input id="txtOutputLocation" name="txtOutputLocation" title="" style="WIDTH: 0px; LEFT: -0px; POSITION: ABSOLUTE" class="msgTxt">
		</span>		
	</span>

	<span style="TOP: 56px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Number of Copies
		<span style="TOP: -3px; LEFT: 155px; POSITION: ABSOLUTE">
			<input id="txtNumberofCopies" name="NumberofCopies" title="" maxlength="5" style="WIDTH: 40px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 84px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Customer Name
		<span style="TOP: -3px; LEFT: 155px; POSITION: ABSOLUTE">
			<select id="cboCustomerName" name="CustomerName" title="" style="WIDTH: 320px" class="msgCombo">
			</select>
		</span>
	</span>
	<span style="TOP: 110px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Recipient
		<span style="TOP: -3px; LEFT: 155px; POSITION: ABSOLUTE">
			<select id="cboRecipient" name="Recipient" title="" style="WIDTH: 320px" class="msgCombo">
			</select>
		</span>
	</span>
</div>	
	<%/*BG 14/10/02 BMIDS00592*/%>
<div id="divViewEditBackground" style="TOP: 405px; LEFT: 10px; HEIGHT: 75px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">
	
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Override</strong>
	</span>
	
	<span style="TOP: 30px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<span style="TOP: 0px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Edit Before Printing
		</span>			
		
		<span style="TOP: -3px; LEFT: 140px; POSITION: ABSOLUTE">
			<input id="optEditBeforePrinting_Yes" name="grpEdit" type="radio" value="0">
		</span>
		<span style="TOP: 0px; LEFT: 165px; POSITION: ABSOLUTE">
			<label for="optEditBeforePrinting_Yes" class="msgLabel">Yes</label>
		</span>
		
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="optEditBeforePrinting_No" name="grpEdit" type="radio" value="1" checked>
		</span>
		<span style="TOP: 0px; LEFT: 225px; POSITION: ABSOLUTE">
			<label for="optEditBeforePrinting_No" class="msgLabel">No</label>
		</span>
	</span>
	
	<span style="TOP: 50px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<span style="TOP: 0px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			View Before Printing
		</span>			
		
		<span style="TOP: -3px; LEFT: 140px; POSITION: ABSOLUTE">
			<input id="optViewBeforePrinting_Yes" name="grpView" type="radio" value="0" >
		</span>
		<span style="TOP: 0px; LEFT: 165px; POSITION: ABSOLUTE">
			<label for="optViewBeforePrinting_Yes" class="msgLabel">Yes</label>
		</span>
		
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="optViewBeforePrinting_No" name="grpView" type="radio" value="1" checked>
		</span>
		<span style="TOP: 0px; LEFT: 225px; POSITION: ABSOLUTE">
			<label for="optViewBeforePrinting_No" class="msgLabel">No</label>
		</span>
	</span>
</div>
	<%/*BG 14/10/02 BMIDS00592 END*/%>

</form>
<!--  #include FILE="attribs/PM010attribs.asp" -->
<%/* Main Buttons */ %>
<div id="divPrint" style="TOP: 485px; LEFT: 10px;HEIGHT:1px; WIDTH: 1px; POSITION: ABSOLUTE">
	<input id="btnPrint" value="Print" type="button" style="TOP: 0px; LEFT: 0px; WIDTH: 71px; HEIGHT: 25px; POSITION: ABSOLUTE" class="msgButton">
	<% /* MAR7 GHun changed button from cancel to close */ %>
	<input id="btnClose" value="Close" type="button" style="TOP: 0px; LEFT: 80px; WIDTH: 71px; HEIGHT: 25px; POSITION: ABSOLUTE" class="msgButton">
</div>
<% /* Specify Code Here */ %>
<% /* MAR7 GHun */ %>
<script src="includes/Documents.js" language="jscript" type="text/javascript"></script>
<% /* MAR7 End */ %>
<script language="JScript" type="text/javascript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_aArgArray = null;
var m_sApplicationNumber = null;
var m_sAFFNumber = null;
var m_sUserId = null;
var m_sUnitId = null;
var m_sStageId = null;
var m_sCustomerNumber = null;
var m_sCustomerVersionNumber = null;
var m_sRole = null;
var m_sRequestAttribs = null;
var m_iTableLength = 10;
var m_sCustomerNameArray = null;
var m_sCustomerNumberArray = null;
var m_sCustomerVersionNumberArray = null;
var m_sLastRowSelected = 0;
var m_sCaseTaskXML = null;
var m_sPrinterXML = null;
var m_sAttributesXML = null;
var m_sPrinterType = "";
var m_BaseNonPopupWindow = null;
var m_sEditBeforePrinting = "";
var m_sViewBeforePrinting = "";

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;

<%/*AS 12/12/2006 EP1262 Gemini Printing */%>
var m_geminiPrintMode = null;

<% /* MAR7 GHun */ %>
var m_sEmailAdministrator = "";
var sDeliveryType = "";
var sDocumentID = "";
var sHostTemplateName = null;
var sCompressionMethod = "";
var sQuotationNumber = "";
var sMortgageSubQuoteNo = "";
var sPrintType = "";
var xmlControlDataNode = null;
var xmlPrintDataNode = null;
var xmlTemplateDataNode = null;
var sPrinterDestType = "";
var sPrinterValidationType= "";
var m_sMachineId = "";
var	m_sDistributionChannelId = "";	
var	m_sProcessingIndicator = "";
var PrintXML = null;
var m_xmlLocalPrinters = null;
var bRoutingFromTaskManagment  = false;
<% /* MAR7 End */ %>

function window.onload()
{	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	RetrieveData();
	Initialise();
	scScreenFunctions.SetCollectionState(spnDocGroup, "D");
	if(m_sAttributesXML != null)
	{
		if(m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATEID") != "")
		{		
			bRoutingFromTaskManagment = true;	<% /* MAR7 GHun */ %>
			TaskManagementProcessing();	
			
			scTable.DisableTable();
		}
	}
	else
	{
		PopulateCombos();
		SetMasks();	
		scScreenFunctions.SetCollectionState(spnDocGroup, "W");	
	}
	window.returnValue = null;
	Validation_Init();	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function Initialise()
{
	scScreenFunctions.SetCollectionState(divAttributesBackground, "D");
	btnPrint.disabled = true;
	
	<%/*BG 14/10/02 BMIDS00592*/%>
	scScreenFunctions.SetCollectionState(divViewEditBackground, "R");
	<%/*BG 14/10/02 BMIDS00592 - END*/%>
}

function PopulateCombos()
{	
	var XMLDocumentType = null;
	var List;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("PrintDocumentType");

	if(XML.GetComboLists(document,sGroupList))
	{		
		XML.CreateTagList("LISTENTRY");
			TagOPTION	= document.createElement("OPTION");
			TagOPTION.value	= "";
			TagOPTION.text	= "<SELECT>";
			frmScreen.cboDocumentGroup.add(TagOPTION);
					
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

function frmScreen.cboDocumentGroup.onchange()
{	
	frmScreen.cboOutputLocation.length = 0;
	frmScreen.txtNumberofCopies.value = "";	
	frmScreen.txtOutputLocation.value = "";	
	frmScreen.cboRecipient.length = 0;
	frmScreen.cboCustomerName.length = 0;
	scScreenFunctions.SetCollectionState(divAttributesBackground, "R");
	scScreenFunctions.SetCollectionState(divAttributesBackground, "D");
	m_sLastRowSelected = 0;
	spnOutputLocation.innerText = "Output Location"
	
	if(frmScreen.cboDocumentGroup.value != "") 
		FilterDocuments();	
}

function FilterDocuments()
{
	//Dependent upon the option selected in the group combo, gets the list of documents
	//from omPrintBO.FindDocumentNameList
	FilterXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	FilterXML.CreateRequestTagFromArray(m_sRequestAttribs, "FindDocumentNameList");
	FilterXML.CreateActiveTag("FINDDOCUMENTS");
	FilterXML.SetAttribute("DOCUMENTGROUP", frmScreen.cboDocumentGroup.value);
	FilterXML.SetAttribute("STAGEID", m_sStageId);
	FilterXML.SetAttribute("USERROLE", m_sRole);
	FilterXML.SetAttribute("PRINTMENUACCESS", "1");
	FilterXML.SetAttribute("INACTIVEINDICATOR", "0");
	
	// Added by automated update TW 01 Oct 2002 SYS5115
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			FilterXML.RunASP(document, "PrintManager.asp");
			break;
		default: // Error
			FilterXML.SetErrorResponse();
	}

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = FilterXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{	
		frmScreen.cboOutputLocation.selectedIndex = 0;
		frmScreen.txtNumberofCopies.value = "";
		frmScreen.txtOutputLocation.value = "";
		frmScreen.cboCustomerName.selectedIndex = 0;
		frmScreen.cboRecipient.selectedIndex = 0;
	
		scScreenFunctions.SetCollectionState(divAttributesBackground, "D");
		scTable.clear();
		alert("There are currently no documents available for the selected document group.");
	}
	else if(ErrorReturn[0] == true)
	{
		PopulateTable();	
	}	
}

function frmScreen.txtNumberofCopies.onblur()
{
	MaxCopies();
}

function RetrieveData()
{	
	var sArguments = window.dialogArguments;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	m_aArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
	
	<%/* SG 06/06/02 SYS4807 */%>
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject(); 
	
	m_sApplicationNumber = m_aArgArray[0];
	m_sAFFNumber = m_aArgArray[1];//ApplicationFactFindNumber
	m_sUserId = m_aArgArray[2];
	m_sUnitId = m_aArgArray[3];
	m_sStageId = m_aArgArray[4];
	m_sCustomerNumber = m_aArgArray[5];//Customer Number from context
	m_sCustomerVersionNumber = m_aArgArray[6];//Customer Version Number from context
	m_sRole = m_aArgArray[7];//User Role from context
	m_sRequestAttribs = m_aArgArray[8];//Request Attributes XML required for "CreateRequestTagFromArray"
	m_sCustomerNameArray	= m_aArgArray[9]; // Customer Name Array from context customers no. 1-5
	m_sCustomerNumberArray	= m_aArgArray[10]; // Customer Number Array from context customers no. 1-5
	m_sCustomerVersionNumberArray = m_aArgArray[11];// Customer Version Number Array from context customers no. 1-5
	
	var m_AttXML = m_aArgArray[12];//Attributes XML passed through from a task
	if(m_AttXML != null)
	{
		m_sAttributesXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_sAttributesXML.LoadXML(m_AttXML);
		m_sAttributesXML.SelectTag(null, "RESPONSE");
	}
	
	m_CaseXML = m_aArgArray[13];//Case Task XML passed through from a task
	if(m_CaseXML != null)
	{
		m_sCaseTaskXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_sCaseTaskXML.LoadXML(m_CaseXML);
		m_sCaseTaskXML.SelectTag(null, "CASETASK");
	}
	
	m_PrintXML = m_aArgArray[14];//Printer XML passed through from a task
	if(m_PrintXML != null)
	{
		m_sPrinterXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_sPrinterXML.LoadXML(m_PrintXML);
		m_sPrinterXML.SelectTag(null, "RESPONSE");
	}
	
	m_sRecipientKey = m_aArgArray[15];//Recipient Key only passed through from a task

	<% /* MAR7 GHun */ %>
	m_sEmailAdministrator = m_aArgArray[16]; // EmailAdministrator
	m_sMachineId = m_aArgArray[17];
	m_sDistributionChannelId = m_aArgArray[18];	
	m_sProcessingIndicator = m_aArgArray[19];
	<% /* MAR7 End */ %>
}

function TaskManagementProcessing()
{	
	//Populate document table		
	TaskPopulateTable();
		
	<% //Populate Document Group Combo %>	
	PopulateDocumentGroupCombo();
	
	<% //Populate Output Location %>	
	PopulateOutputLocation();
		
	<% //Populate NumberOfCopies %>
	PopulateNoOfCopies();
		
	//Populate Customer name
	var sCustomer = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "CUSTOMERSPECIFICIND");
	frmScreen.cboCustomerName.length = 0;		
	
	if(sCustomer == "1")
	{		
		var sCustIdentifier = m_sCaseTaskXML.GetAttribute("CUSTOMERIDENTIFIER");
		if(sCustIdentifier != "")
		{	
			for(var nLoop = 0;nLoop < m_sCustomerNameArray.length;nLoop++)
			{											
				if(m_sCustomerNumberArray[nLoop] == sCustIdentifier)
				{
					var objOPTION = document.createElement("OPTION");
					objOPTION.text = m_sCustomerNameArray[nLoop];
					objOPTION.value = m_sCustomerNumberArray[nLoop];
					//GD BMIDS01041 
					objOPTION.setAttribute("attCustVersionNumber", m_sCustomerVersionNumberArray[nLoop]);
					frmScreen.cboCustomerName.add(objOPTION);	
					//GD BMIDS01041 frmScreen.cboCustomerName.setAttribute("attCustVersionNumber", m_sCustomerVersionNumberArray[nLoop]);

				}
			}
		}
	}	
					
	//Find Recipients if required
	var sRecipient = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "RECIPIENTTYPE");
	if(sRecipient != "")
	{	
		FindRecipients(sRecipient);						
	}	
	
	//enable print button
	btnPrint.disabled = false;
	
	<%/*BG 14/10/02 BMIDS00592 - set radio buttons*/%>
	Override(m_sAttributesXML);		
	<%/*BG 14/10/02 BMIDS00592 - END*/%>	
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
function btnClose.onclick()
{	
	window.close();
}
function btnPrint.onclick()
{
	var success = "";	<% /* MAR7 GHun */ %>
	
	if(MaxCopies())
	{
		if(frmScreen.cboOutputLocation.value == "" && frmScreen.txtOutputLocation.value == "" )
		{
			alert("You have not selected a printer");
			return;
		}
		if(frmScreen.txtNumberofCopies.value == "" || frmScreen.txtNumberofCopies.value =="0")
		{
			alert("You have not entered number of copies required");
			btnClose.focus();	<% /* MAR7 GHun */ %>
			//BG 29/03/01 specifically not set focus to recipient combo as the recipients
			//may not be set up for the selected customer, in which case the combo will be disabled - 
			//setting focus to it would then cause the screen to fall over.
			return;
		}
		if(m_sAttributesXML == null)
		{
			var bCustomer = AttribsXML.GetTagAttribute("ATTRIBUTES", "CUSTOMERSPECIFICIND");
			if(bCustomer == 1 && frmScreen.cboCustomerName.value == "")
			{
				alert("You have not selected a customer name");
				frmScreen.cboCustomerName.focus();
				return;
			}
				
			var sRecipientType = AttribsXML.GetTagAttribute("ATTRIBUTES", "RECIPIENTTYPE");
			if(sRecipientType != "" && frmScreen.cboRecipient.value == "")
			{
				alert("A recipient is required to print this document");
				btnClose.focus();	<% /* MAR7 GHun */ %>
				//BG 29/03/01 specifically not set focus to recipient combo as the recipients
				//may not be set up for the selected customer, in which case the combo will be disabled - 
				//setting focus to it would then cause the screen to fall over.
				return;
			}		
		}
		
		sPrinterValidationType = getPrinterTypeForAttributes(sPrinterDestType);	<% /* MAR7 GHun */ %>

		PrintXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		PrintXML.CreateRequestTagFromArray(m_sRequestAttribs, "PrintDocument");
		//BG 10/10/02 BMIDS00612 
		PrintXML.SetAttribute("PRINTINDICATOR", "1");
		//BG 10/10/02 BMIDS00612 - END

		<% /* MAR7 GHun */ %>
		BuildPrintDataBlock(PrintXML);
		BuildTemplateDataBlock(PrintXML);
		BuildControlDataBlock(PrintXML);
		<% /* MAR7 End */ %>
		
		PrintXML.RunASP(document, "PrintManager.asp");
		
		var ErrorTypes = new Array("NOTIMPLEMENTED");
		var ErrorReturn = PrintXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{				
			alert("Your print template has been set with an incorrect Method name. Please contact your System Administrator");
		}
		else if(ErrorReturn[0] == true)
		<% /* MAR7 GHun */ %>
		{			
			if (frmScreen.optEditBeforePrinting_Yes.checked == true || frmScreen.optViewBeforePrinting_Yes.checked == true || sPrinterValidationType == "W")
			{	
				var fileContents = PrintXML.GetTagAttribute("DOCUMENTCONTENTS", "FILECONTENTS");
				if (fileContents == "")
				{
					// omDPS returns a DOC in a PRINTDOCUMENT attribute of a PRINTDOCUMENTDETAILS node.
					fileContents = PrintXML.GetTagAttribute("PRINTDOCUMENTDETAILS", "PRINTDOCUMENT");
				}
				var printDocumentData = 
					printDocument(
						fileContents, 
						sDocumentID, 
						sHostTemplateName, 
						sDeliveryType, 
						sCompressionMethod, 
						sPrinterValidationType, 
						frmScreen.txtNumberofCopies.value,
						true, true, 
						!frmScreen.optEditBeforePrinting_Yes.checked, 
						!(frmScreen.optEditBeforePrinting_Yes.checked || frmScreen.optViewBeforePrinting_Yes.checked),
						null,
						m_geminiPrintMode)

				if (printDocumentData != null && printDocumentData.get_success())
				{
					if (printDocumentData.get_fileContents() != null)
					{	
						if ((frmScreen.optEditBeforePrinting_Yes.checked == true || frmScreen.optViewBeforePrinting_Yes.checked == true) && sPrinterValidationType == "W")
						{
							var printDocumentData = 
										<% /* PSC 27/01/2006 MAR1116 */ %>
										printDocument(
											printDocumentData.get_fileContents(),  
											sDocumentID, 
											sHostTemplateName, 
											sDeliveryType, 
											sCompressionMethod, 
											sPrinterValidationType, 
											frmScreen.txtNumberofCopies.value,
											true, 
											true, 
											true, 
											true,
											null,
											m_geminiPrintMode);
						}

						success = 
							savePrintedDocument(
								m_BaseNonPopupWindow, m_sUserId, m_sUnitId, m_sMachineId, m_sDistributionChannelId, "10", 
								xmlControlDataNode, xmlPrintDataNode, xmlTemplateDataNode, 
								printDocumentData.get_fileSize(), 
								printDocumentData.get_fileContents() != null && frmScreen.optEditBeforePrinting_Yes.checked ? printDocumentData.get_fileContents() : fileContents, 
								sCompressionMethod);

// TW 29/11/2005 MAR716
//<% /* TW MAR579 21/11/2005 */ %>
//						if (success && sDeliveryType == "2")
//						{
//							// Send pdf's to FileNet
//							success = sendDocumentToFileNet(fileContents);
//						}
//<% /* TW MAR579 21/11/2005 End */ %>
// TW 29/11/2005 MAR716 End				
					
					}
					else
					{
						success = true;
					}
				}
				else if (printDocumentData.get_success() == false)
				{
					alert("Printing was Cancelled by the User.");
					return;
				}
			}
			else
			{
				success = true;
			}
		}		
		
		if (success)
		{
			alert("Action completed successfully.");
		<% /* MAR7 End */ %>	
			
			<% /* PSC 30/01/2006 MAR1147 */ %>
			<% /*
			var sReturn = new Array();
			sReturn[0] = "Success";	*/ %>
			window.returnValue = "Success";
			
			if (bRoutingFromTaskManagment == true) <% /* MAR7 GHun */ %>
				window.close();
		}
		else
		{
			alert("The action failed to complete.");	<% /* MAR7 GHun */ %>
		}
	}
}

function MaxCopies()
{		

	var bMax = true;
	if(m_sAttributesXML == null)
	{
		var iMaxCopies = AttribsXML.GetTagAttribute("ATTRIBUTES", "MAXCOPIES")
	}
	else
	{
		var iMaxCopies = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "MAXCOPIES")
	}
		
	if (iMaxCopies != "")
	{
		//DR DMSSYS0003 pasrseInt() now used.
		if(parseInt(frmScreen.txtNumberofCopies.value) > parseInt(iMaxCopies))
		{
			alert("Your request has exceeded the permitted number of copies: " + iMaxCopies + ", please re-enter.");	
			var bMax = false;
			return(bMax);
		}
	}
		
	return(bMax);	
}
function frmScreen.cboRecipient.onchange()
{
	if(frmScreen.cboRecipient.value != "")
	{
		btnPrint.disabled = false;
	}
	else
	{
		btnPrint.disabled = true;
	}
}

function frmScreen.cboCustomerName.onchange()
{		
	frmScreen.cboRecipient.length = 0;
	//frmScreen.cboRecipient.disabled = true;
	scScreenFunctions.SetFieldState(frmScreen, "cboRecipient", "D");
	scScreenFunctions.SetFieldState(frmScreen, "cboRecipient", "R");	
	if(frmScreen.cboCustomerName.value != "")
	{
		var sRecipient = AttribsXML.GetTagAttribute("ATTRIBUTES", "RECIPIENTTYPE");
		if(sRecipient == "")
		{
			btnPrint.disabled = false;
		}
		else
		{
			if(sRecipient != "")
			{	
				FindRecipients(sRecipient);						
			}
		}
	}
}

function FindRecipients(sRecipientValue)
{
	
	TempXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("PrintRecipientType");
	var sRecipientList = TempXML.GetComboLists(document,sGroups);
	var sValueID = sRecipientValue;
	var sRecipientType = TempXML.GetTagText("LIST/LISTNAME/LISTENTRY[VALUEID='" + sValueID + "']/VALIDATIONTYPELIST/VALIDATIONTYPE");
	
	RecipientXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();	
	RecipientXML.CreateRequestTagFromArray(m_sRequestAttribs, "FindThirdPartyForCustomer");
	RecipientXML.CreateActiveTag("THIRDPARTY");	
	RecipientXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	RecipientXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAFFNumber);
	RecipientXML.CreateTag("CUSTOMERNUMBER", frmScreen.cboCustomerName.value);
	// GD BMIDS01041 RecipientXML.CreateTag("CUSTOMERVERSIONNUMBER", frmScreen.cboCustomerName.getAttribute("attCustVersionNumber"));
	RecipientXML.CreateTag("CUSTOMERVERSIONNUMBER",frmScreen.cboCustomerName.options.item(frmScreen.cboCustomerName.selectedIndex).getAttribute("attCustVersionNumber"));
	

	RecipientXML.CreateTag("THIRDPARTYTYPE", sRecipientType);
	
	// 	RecipientXML.RunASP(document, "FindThirdPartyForCustomer.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			RecipientXML.RunASP(document, "FindThirdPartyForCustomer.asp");
			break;
		default: // Error
			RecipientXML.SetErrorResponse();
		}

	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = RecipientXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{
		alert("Recipient details do not exist for this customer.");
	}
	else if(ErrorReturn[0] == true)
	{	
		// Loop through all entries and only add relevant entries to combo
		btnPrint.disabled = false;
		RecipientXML.CreateTagList("THIRDPARTY");
		if(RecipientXML.ActiveTagList.length > 1)
		{
			if(m_sAttributesXML == null)
			{
				TagOPTION	= document.createElement("OPTION");
				TagOPTION.value	= "";
				TagOPTION.text	= "<SELECT>";
				frmScreen.cboRecipient.add(TagOPTION);
				scScreenFunctions.SetFieldState(frmScreen, "cboRecipient", "W");
				frmScreen.cboRecipient.disabled = false;
				btnPrint.disabled = true;
			}			
		}
						
		for(var iLoop = 0; iLoop < RecipientXML.ActiveTagList.length; iLoop++)
		{	
			if(m_sAttributesXML == null)
			{
				RecipientXML.SelectTagListItem(iLoop);
				var sRecipientName	= RecipientXML.GetTagText("NAME");
				var sRecipientContext = RecipientXML.GetTagText("CONTEXT");
				TagOPTION = document.createElement("OPTION");
				TagOPTION.value	= sRecipientContext;
				TagOPTION.text = sRecipientName;
					
				frmScreen.cboRecipient.add(TagOPTION);	
			}
			else
			{					
				RecipientXML.SelectTagListItem(iLoop);
				var sRecipientName	= RecipientXML.GetTagText("NAME");
				var sRecipientContext = RecipientXML.GetTagText("CONTEXT");
				var sTaskContext = m_sCaseTaskXML.GetAttribute("CONTEXT");
				if(sRecipientContext == sTaskContext)
				{
					TagOPTION = document.createElement("OPTION");
					TagOPTION.value	= sRecipientContext;
					TagOPTION.text = sRecipientName;
						
					frmScreen.cboRecipient.add(TagOPTION);	
				}
			}								
		}						
	}	
}

function switchOffOutputLocationCombo(strMessage, strReadOnlyValue)
{	
	//Switch off the combo, switch on the text field	
	frmScreen.cboOutputLocation.style.width = 0;						
	frmScreen.txtOutputLocation.style.width = 340;							
	frmScreen.txtOutputLocation.disabled = false;
	if (strReadOnlyValue == "")
	{
		//its writeable
		scScreenFunctions.SetFieldState(frmScreen, "txtOutputLocation", "W");	
	}
	else
	{
		//its read only - just bung in the default value supplied
		frmScreen.txtOutputLocation.value = strReadOnlyValue;
	}
					
	spnOutputLocation.innerText = strMessage;
}

function EnableAttributes(sDocID)
{	
	frmScreen.cboOutputLocation.length = 0;
	frmScreen.txtNumberofCopies.value = "";
	frmScreen.cboRecipient.length = 0;
	frmScreen.cboCustomerName.length = 0;
	scScreenFunctions.SetCollectionState(divAttributesBackground, "R");
	scScreenFunctions.SetCollectionState(divAttributesBackground, "D");
			
	//Switch on the combo, switch off the text field
	frmScreen.txtOutputLocation.value = "";
	frmScreen.txtOutputLocation.style.width = 0;
	frmScreen.cboOutputLocation.style.width = 340;							
	frmScreen.cboOutputLocation.disabled = false;
	scScreenFunctions.SetFieldState(frmScreen, "cboOutputLocation", "W");	
		
	btnPrint.disabled = true;
	XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	AttribsXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	AttribsXML.CreateRequestTagFromArray(m_sRequestAttribs, "GetPrintAttributes");
	AttribsXML.CreateActiveTag("FINDATTRIBUTES");
	AttribsXML.SetAttribute("HOSTTEMPLATEID", sDocID);
				
	// 	AttribsXML.RunASP(document, "PrintManager.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
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
		sPrinterDestType = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE"); <% /* MAR7 GHun */ %>
		
		var TempXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		
		var sDestination = AttribsXML.GetTagAttribute("ATTRIBUTES", "REMOTEPRINTERLOCATION");
		//var sDestType = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
		
		var bDisplayMaxCopies = false;

		if (TempXML.IsInComboValidationList(document,"PrinterDestination",sPrinterDestType,["R"]))	<% /* MAR7 GHun */ %>
		{	
			bDisplayMaxCopies = true;
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= sPrinterDestType;	<% /* MAR7 GHun */ %>
			TagOPTION.text = sDestination;				
			frmScreen.cboOutputLocation.add(TagOPTION);			
			spnOutputLocation.innerText = "Output Location: Remote Printer";			
		}
		<% /* MAR7 GHun */ %>
		else if (TempXML.IsInComboValidationList(document,"PrinterDestination",sPrinterDestType,["W"]))
		{
			switchOffOutputLocationCombo("Output Location: ", "Workstation Printer")				
		}
		<% /* MAR7 End */ %>
		else if (TempXML.IsInComboValidationXML(["L"]))
		{
			bDisplayMaxCopies = true;			
		}
		else if (TempXML.IsInComboValidationXML(["F"]))
		{			
			switchOffOutputLocationCombo("Output Location: to FILE", sDestination + "\\" + m_sUserId);
		}
		else if (TempXML.IsInComboValidationXML(["E"]))
		{
			switchOffOutputLocationCombo("Output Location: to EMAIL", "")
		}
		else if (TempXML.IsInComboValidationXML(["X"]))
		{
			// ik_bmids00730
			// switchOffOutputLocationCombo("Output Location: to FAX", "");
			switchOffOutputLocationCombo("Output Location: to FAX", sDestination);
			// ik_bmids00730_ends
		}
		else if (TempXML.IsInComboValidationXML(["S"]))
		{
			switchOffOutputLocationCombo("Output Location: to SMS", "")
		}				
		<% // RF 16/02/2006 MAR1251 Allow printer destination of "Document Store Only" %>
		else if (TempXML.IsInComboValidationXML(["DMS"]))
		{
			switchOffOutputLocationCombo("Output Location: " , "Document Store Only");
		}
		<% /* EP2_1903 GHun */ %>
		else if (TempXML.IsInComboValidationXML(["EF"]))
		{
			switchOffOutputLocationCombo("Output Location: ", "Email Fulfilment");
		}
		<% /* EP2_1903 End */ %>

		if (TempXML.IsInComboValidationXML(["L"]))
		{
			<% /* MAR7 GHun */ %>
			m_xmlLocalPrinters = getLocalPrinters(m_xmlLocalPrinters);
			
			if(m_xmlLocalPrinters != null)
			{	
				if (m_xmlLocalPrinters.selectSingleNode("RESPONSE/PRINTER[DEFAULTINDICATOR='1']") == null)
			<% /*
			var sLocalPrinters = GetLocalPrinters();
			if(sLocalPrinters != "")
			{	
				var LocalPrintersXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
				LocalPrintersXML.LoadXML(sLocalPrinters);
				
				LocalPrintersXML.CreateTagList("PRINTER[DEFAULTINDICATOR='1']");
						
				if(LocalPrintersXML.ActiveTagList.length == 0)
				*/ %>
			<% /* MAR7 End */ %>
				{					
					alert("You do not have a default printer set on your PC.");				
				}		
				else
				{	
					<% /* MAR7 GHun
					LocalPrintersXML.CreateTagList("PRINTER");

					for(var iLoop = 0; iLoop < LocalPrintersXML.ActiveTagList.length; iLoop++)
					{
						LocalPrintersXML.SelectTagListItem(iLoop);
						var sDestination	= LocalPrintersXML.GetTagText("PRINTERNAME");
						var sDestType = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
						TagOPTION = document.createElement("OPTION");
						TagOPTION.value	= sDestType;
					*/ %>
					var printers = m_xmlLocalPrinters.getElementsByTagName("PRINTER");
					for (var printerIndex = 0; printerIndex < printers.length; printerIndex++)
					{
						var printer = printers.item(printerIndex);
						
						if (printer != null)
						{
							var element = null;
							var destination = (element = printer.selectSingleNode("PRINTERNAME")) ? element : printer.attributes.getNamedItem("DEVICENAME");
							var defaultIndicator = (element = printer.selectSingleNode("DEFAULTINDICATOR")) ? element  : printer.attributes.getNamedItem("DEFAULT");
							var sDestination = destination ? destination.text : "";
							TagOPTION = document.createElement("OPTION");
							TagOPTION.value	= sPrinterDestType;
					<% /* MAR7 End */ %>
							TagOPTION.text = sDestination;			

							// ik_bmids00730 (Local and Remote)
							if (TempXML.IsInComboValidationXML(["R"]))
							{
								var bHit=false;
								for(i=0;!bHit && i < frmScreen.cboOutputLocation.length;i++)
								{
									if(frmScreen.cboOutputLocation.item(i).text == sDestination)
									{
										bHit=true;
									}
								}
								if(!bHit)
								{
									frmScreen.cboOutputLocation.add(TagOPTION);
								}
							}
							else
							{					
								frmScreen.cboOutputLocation.add(TagOPTION);
								
								<% /* MAR7 GHun	
								if(LocalPrintersXML.GetTagText("DEFAULTINDICATOR") == 1) */ %>
								if (defaultIndicator && (defaultIndicator.text == "1"))
								<% /* MAR7 End */ %>
								{	
									if(frmScreen.cboOutputLocation.length > 1)
									{
										frmScreen.cboOutputLocation.selectedIndex = iLoop; <% /* MAR7 GHun changed from true to iLoop */ %>
									}
									spnOutputLocation.innerText = "Output Location: Local Printer";
									frmScreen.cboOutputLocation.disabled = false;
								}
							}
						}
					}
				
					if(frmScreen.cboOutputLocation.length > 1)
					{	
						frmScreen.cboOutputLocation.disabled = false;
						scScreenFunctions.SetFieldState(frmScreen, "cboOutputLocation", "W");
						spnOutputLocation.innerText = "Output Location";
					}
				}				
			}
		}		
			
		TempXML = null;
		frmScreen.txtNumberofCopies.value = 1;
		
		if (bDisplayMaxCopies)
		{
			//Populate Numberofcopies field
			var sMaxCopies = AttribsXML.GetTagAttribute("ATTRIBUTES", "MAXCOPIES");
			if(sMaxCopies == "")
			{
				frmScreen.txtNumberofCopies.value = AttribsXML.GetTagAttribute("ATTRIBUTES", "DEFAULTCOPIES")
			}
			else
			{
				frmScreen.txtNumberofCopies.value = AttribsXML.GetTagAttribute("ATTRIBUTES", "DEFAULTCOPIES")
				frmScreen.txtNumberofCopies.disabled = false;
				scScreenFunctions.SetFieldState(frmScreen, "txtNumberofCopies", "W");
			}
		}
		
		//Populate customer combos
		var sCustomer = AttribsXML.GetTagAttribute("ATTRIBUTES", "CUSTOMERSPECIFICIND");
		frmScreen.cboCustomerName.length = 0;		
		if(sCustomer == "0")
		{
			btnPrint.disabled = false;
		}
		else
		{
			if(m_sCustomerNumberArray.length > "1")
			{							
				TagOPTION = document.createElement("OPTION");
				TagOPTION.value	= "";
				TagOPTION.text = "<SELECT>";				
				frmScreen.cboCustomerName.add(TagOPTION);
				PopulateCustomerCombo();				
			}
			else if(m_sCustomerNumberArray.length == "1")
			{
				PopulateCustomerCombo();
										
				var sRecipient = AttribsXML.GetTagAttribute("ATTRIBUTES", "RECIPIENTTYPE");
				if(sRecipient == "")
				{
					btnPrint.disabled = false;
				}
				else
				{
					if(sRecipient != "")
					{	
						FindRecipients(sRecipient);						
					}	
				}
			}			
		}
	
		<%/*AS 12/12/2006 EP1262 Gemini Printing */%>
		m_geminiPrintMode = AttribsXML.GetTagAttribute("ATTRIBUTES", "GEMINIPRINTMODE");
			
		<%/*BG 14/10/02 BMIDS00592 - set radio buttons*/%>
		Override(AttribsXML);		
		<%/*BG 14/10/02 BMIDS00592 - END*/%>	
	}
}

<%/*BG 14/10/02 BMIDS00592 Added function*/%>
function Override(AttribsXML)
{
	m_sEditBeforePrinting = AttribsXML.GetTagAttribute("ATTRIBUTES", "EDITBEFOREPRINT");	
	m_sViewBeforePrinting = AttribsXML.GetTagAttribute("ATTRIBUTES", "VIEWBEFOREPRINT");

	frmScreen.optEditBeforePrinting_No.checked = true;
	frmScreen.optViewBeforePrinting_No.checked = true;
	frmScreen.optEditBeforePrinting_No.disabled = true;
	frmScreen.optEditBeforePrinting_Yes.disabled = true;
	frmScreen.optViewBeforePrinting_No.disabled = true;
	frmScreen.optViewBeforePrinting_Yes.disabled = true;
		
	if (m_sEditBeforePrinting == "1")
	{
		frmScreen.optEditBeforePrinting_No.disabled = false;
		frmScreen.optEditBeforePrinting_Yes.disabled = false;
		// ik_bm0416
		// frmScreen.optEditBeforePrinting_Yes.checked = true;
		//frmScreen.optEditBeforePrinting_No.checked = true;
		// ik_bm0416_ends
		//BMIDS719 Edit Before Printing = Yes
		frmScreen.optEditBeforePrinting_Yes.checked = true;
	}
	
	if (m_sViewBeforePrinting == "1")
	{
		frmScreen.optViewBeforePrinting_No.disabled = false;
		frmScreen.optViewBeforePrinting_Yes.disabled = false;
		// ik_bm0416
		// frmScreen.optViewBeforePrinting_Yes.checked = true;
		//frmScreen.optViewBeforePrinting_No.checked = true;		
		// ik_bm0416_ends
		// BMIDS719 View Before Printing = Yes
		frmScreen.optViewBeforePrinting_Yes.checked = true;
		
	}
}
<%/*BG 14/10/02 BMIDS00592 - END*/%>

function spnTable.onclick()
{
	var iRowSelected = scTable.getRowSelected();
	if(iRowSelected !=-1)
	{
		if(m_sLastRowSelected != iRowSelected)
		{
			EnableAttributes(tblTable.rows(iRowSelected).getAttribute("DocumentID"));
			
		}
		m_sLastRowSelected = iRowSelected;
	}
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
	var iCount;
	scTable.clear();
	for (iCount = 0; iCount < FilterXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		FilterXML.SelectTagListItem(iCount + nStart);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),FilterXML.GetAttribute("HOSTTEMPLATENAME"));
		tblTable.rows(iCount+1).setAttribute("DocumentID",FilterXML.GetAttribute("HOSTTEMPLATEID"));
		tblTable.rows(iCount+1).cells(0).title = FilterXML.GetAttribute("HOSTTEMPLATEDESCRIPTION");
	}
}

function PopulateCustomerCombo()
{
	//for length of array assign the values to the combo box	
	for(var nLoop = 0;nLoop < m_sCustomerNameArray.length;nLoop++)
	{
		var objOPTION = document.createElement("OPTION");
		objOPTION.text = m_sCustomerNameArray[nLoop];
		objOPTION.value = m_sCustomerNumberArray[nLoop];
		//GD BMIDS01041 Added :
		objOPTION.setAttribute("attCustVersionNumber", m_sCustomerVersionNumberArray[nLoop]);
		frmScreen.cboCustomerName.add(objOPTION);		
		//GD BMIDS01041
		//frmScreen.cboCustomerName.setAttribute("attCustVersionNumber", m_sCustomerVersionNumberArray[nLoop]);
	}
	if(frmScreen.cboCustomerName.length > 1)
	{
	scScreenFunctions.SetFieldState(frmScreen, "cboCustomerName", "W");	
	}
}

<% /* MAR7 GHun */ %>
function PopulateOutputLocation()
{
	sPrinterDestType = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
	var TempXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sDestination = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "REMOTEPRINTERLOCATION");
	
	if (TempXML.IsInComboValidationList(document,"PrinterDestination",sPrinterDestType,["R"]))
	{
		TagOPTION = document.createElement("OPTION");
		TagOPTION.value	= sPrinterDestType;
		TagOPTION.text = sDestination;
		frmScreen.cboOutputLocation.add(TagOPTION);
		spnOutputLocation.innerText = "Output Location: Remote Printer";
	}	
	else if (TempXML.IsInComboValidationXML(["F"]))
	{
		TagOPTION = document.createElement("OPTION");
		TagOPTION.value	= sPrinterDestType;
		TagOPTION.text = sDestination;
		frmScreen.cboOutputLocation.add(TagOPTION);
		spnOutputLocation.innerText = "Output Location: to FILE";
	}
	else if (TempXML.IsInComboValidationXML(["E"]))
	{
		EnableAttributes(m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATEID"));
		spnOutputLocation.innerText = "Output Location: to EMAIL";
	}
	else if (TempXML.IsInComboValidationXML(["W"]))
	{
		switchOffOutputLocationCombo("Output Location: " , "Workstation Printer");
	}
	<% // RF 16/02/2006 MAR1251 Allow printer destination of "Document Store Only" %>
	else if (TempXML.IsInComboValidationXML(["DMS"]))
	{
		switchOffOutputLocationCombo("Output Location: " , "Document Store Only");
	}
	<% /* EP2_1903 GHun */ %>
	else if (TempXML.IsInComboValidationXML(["EF"]))
	{
		switchOffOutputLocationCombo("Output Location: " , "Email Fulfilment");
	}
	<% /* EP2_1903 GHun */ %>
	
	if (TempXML.IsInComboValidationXML(["L"]))
	{
		if(m_sPrinterXML != "")
		{
			m_sPrinterXML.CreateTagList("PRINTER");

			for(var iLoop = 0; iLoop < m_sPrinterXML.ActiveTagList.length; iLoop++)
			{
				m_sPrinterXML.SelectTagListItem(iLoop);
				sDestination	= m_sPrinterXML.GetTagText("PRINTERNAME");
				
				TagOPTION = document.createElement("OPTION");
				TagOPTION.value	= sPrinterDestType;
				TagOPTION.text = sDestination;
				
				if (TempXML.IsInComboValidationXML(["R"]))
				{
					var bHit=false;
					for(i=0;!bHit && i < frmScreen.cboOutputLocation.length;i++)
					{
						if(frmScreen.cboOutputLocation.item(i).text == sDestination)
						{
							bHit=true;
						}
					}
					if(!bHit)
					{
						frmScreen.cboOutputLocation.add(TagOPTION);
					}
				}
				else
				{
					frmScreen.cboOutputLocation.add(TagOPTION);
					
					if(m_sPrinterXML.GetTagText("DEFAULTINDICATOR") == 1)
					{
						frmScreen.cboOutputLocation.selectedIndex = iLoop;
						spnOutputLocation.innerText = "Output Location: Local Printer";
					}
				}
			}
			
			if(frmScreen.cboOutputLocation.length > 1)
			{
				frmScreen.cboOutputLocation.disabled = false;
				scScreenFunctions.SetFieldState(frmScreen, "cboOutputLocation", "W");
				spnOutputLocation.innerText = "Output Location";
			}
		}
	}
	
	TempXML = null;
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
			if(XML.GetTagText("VALUEID") == m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "TEMPLATEGROUPID"))
			{
				XML.SelectTagListItem(iLoop);
				var sGroupName	= XML.GetTagText("VALUENAME");
				var sValueID = XML.GetTagText("VALUEID");
				TagOPTION = document.createElement("OPTION");
				TagOPTION.value	= sValueID;
				TagOPTION.text = sGroupName;					
				frmScreen.cboDocumentGroup.add(TagOPTION);
			}
		}
	}
	XML = null;
}

function PopulateNoOfCopies()
{
	<% //Populate Numberofcopies field %>
	var sMaxCopies = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "MAXCOPIES");
	if(sMaxCopies == "")
	{
		frmScreen.txtNumberofCopies.value = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "DEFAULTCOPIES");
	}
	else
	{
		frmScreen.txtNumberofCopies.value = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "DEFAULTCOPIES");
		frmScreen.txtNumberofCopies.disabled = false;
		scScreenFunctions.SetFieldState(frmScreen, "txtNumberofCopies", "W");
	}
}

function BuildTemplateDataBlock(PrintXML)
{
	//GetApplicationData
	var AppXML =  new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	AppXML.CreateRequestTagFromArray(m_sRequestAttribs, "GetAcceptedQuoteData");
	AppXML.CreateActiveTag("APPLICATION");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AppXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAFFNumber);
	
	AppXML.RunASP(document,"GetAcceptedOrActiveQuoteData.asp");
	
	if(AppXML.IsResponseOK())
	{
		if ( AppXML.SelectTag(null, "QUOTATION") != null )
		{
			sQuotationNumber = AppXML.GetTagText("QUOTATIONNUMBER");
			AppXML.SelectTag(null, "MORTGAGESUBQUOTE");
			sMortgageSubQuoteNo = AppXML.GetTagText("MORTGAGESUBQUOTENUMBER");
		
			if (sQuotationNumber != "" && sQuotationNumber != null)
			{
				PrintXML.SelectTag(null, "REQUEST");
				xmlTemplateDataNode = PrintXML.CreateActiveTag("TEMPLATEDATA");
				PrintXML.SetAttribute("QUOTATIONNUMBER", sQuotationNumber);
				PrintXML.SetAttribute("MORTGAGESUBQUOTENUMBER", sMortgageSubQuoteNo);
			}
		}
	}
	
	AppXML = null;
}

function BuildControlDataBlock(PrintXML)
{
	PrintXML.SelectTag(null, "REQUEST");
	xmlControlDataNode = PrintXML.CreateActiveTag("CONTROLDATA");
	PrintXML.SetAttribute("COPIES", frmScreen.txtNumberofCopies.value);

	if ( frmScreen.txtOutputLocation.value != "")
	{
		PrintXML.SetAttribute("PRINTER", frmScreen.txtOutputLocation.value);
	}
	else
	{
		if (frmScreen.cboOutputLocation.selectedIndex != -1)
			PrintXML.SetAttribute("PRINTER", frmScreen.cboOutputLocation.options(frmScreen.cboOutputLocation.selectedIndex).text);
		//else
			//PrintXML.SetAttribute("PRINTER", frmScreen.txtOutputLocation.value);
	}
	
	if(frmScreen.optEditBeforePrinting_Yes.checked == true)
		PrintXML.SetAttribute("EDITBEFOREPRINT", "1");
	else
		PrintXML.SetAttribute("EDITBEFOREPRINT", "0");
	
	if(frmScreen.optViewBeforePrinting_Yes.checked == true)
		PrintXML.SetAttribute("VIEWBEFOREPRINT", "1");
	else
		PrintXML.SetAttribute("VIEWBEFOREPRINT", "0");
	
	if (m_sAttributesXML == "" || m_sAttributesXML == null)
	{
		sDocumentID = AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATEID");
		PrintXML.SetAttribute("DOCUMENTID", sDocumentID);
		PrintXML.SetAttribute("DPSDOCUMENTID", AttribsXML.GetTagAttribute("ATTRIBUTES", "DPSTEMPLATEID"));
		PrintXML.SetAttribute("DOCUMENTNAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "DOCUMENTNAME"));
		
		sDeliveryType = AttribsXML.GetTagAttribute("ATTRIBUTES", "DELIVERYTYPE");
		PrintXML.SetAttribute("DELIVERYTYPE", sDeliveryType);

// TW 21/11/2005 MAR579 
// Can't compress when data will be going to fileNet
//		sCompressionMethod = "ZLIB";
		sCompressionMethod = "";
// TW 29/11/2005 MAR716
//		if (sDeliveryType == "2")
//		{
//			PrintXML.SetAttribute("VIEWBEFOREPRINT", "1");
//		}
// TW 29/11/2005 MAR716 End
// TW 21/11/2005 MAR579 End
		
		PrintXML.SetAttribute("COMPRESSIONMETHOD",sCompressionMethod);
		PrintXML.SetAttribute("TEMPLATEGROUPID", AttribsXML.GetTagAttribute("ATTRIBUTES", "TEMPLATEGROUPID"));
		sHostTemplateName = AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME");
		PrintXML.SetAttribute("HOSTTEMPLATENAME", sHostTemplateName);
		PrintXML.SetAttribute("HOSTTEMPLATEDESCRIPTION", AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATEDESCRIPTION"));
		<% /* EP2_1903 GHun */ %>
		PrintXML.SetAttribute("EMAILRECIPIENTTYPE", AttribsXML.GetTagAttribute("ATTRIBUTES", "EMAILRECIPIENTTYPE"));
		if (sPrinterValidationType == "EF" && m_sRecipientKey != null)	<% /* EP2_2026 GHun */ %>
			PrintXML.SetAttribute("DOCUMENTLOCATION", m_sRecipientKey);
		<% /* EP2_1903 End */ %>
	}
	else
	{
		sDocumentID = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATEID");
		PrintXML.SetAttribute("DOCUMENTID",sDocumentID );
		PrintXML.SetAttribute("DPSDOCUMENTID", m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "DPSTEMPLATEID"));
		PrintXML.SetAttribute("DOCUMENTNAME", m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "DOCUMENTNAME"));

		sDeliveryType = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "DELIVERYTYPE");
		PrintXML.SetAttribute("DELIVERYTYPE", sDeliveryType);
		
// TW 21/11/2005 MAR579 
// Can't compress when data will be going to fileNet
//		sCompressionMethod = "ZLIB";
		sCompressionMethod = "";
// TW 29/11/2005 MAR716
//		if (sDeliveryType == "2")
//		{
//			PrintXML.SetAttribute("VIEWBEFOREPRINT", "1");
//		}
// TW 29/11/2005 MAR716 End
// TW 21/11/2005 MAR579 End
		
		PrintXML.SetAttribute("COMPRESSIONMETHOD",sCompressionMethod);
		PrintXML.SetAttribute("TEMPLATEGROUPID", m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "TEMPLATEGROUPID"));
		sHostTemplateName = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME");
		PrintXML.SetAttribute("HOSTTEMPLATENAME", sHostTemplateName);
		PrintXML.SetAttribute("HOSTTEMPLATEDESCRIPTION", m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATEDESCRIPTION"));
		<% /* EP2_1903 GHun */ %>
		PrintXML.SetAttribute("EMAILRECIPIENTTYPE", m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "EMAILRECIPIENTTYPE"));
		if (sPrinterValidationType == "EF" && m_sRecipientKey != null)	<% /* EP2_2026 GHun */ %>
			PrintXML.SetAttribute("DOCUMENTLOCATION", m_sRecipientKey);
		<% /* EP2_1903 End */ %>
	}
	
	<% /* EP2_1903 GHun */ %>
	PrintXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	<% /* EP2_1903 End */ %>
	
	PrintXML.SetAttribute("DESTINATIONTYPE", sPrinterValidationType);
	PrintXML.SetAttribute("STAGEID", m_sStageId);
	
	PrintXML.SetAttribute("CUSTOMERNO", frmScreen.cboCustomerName.value);
	
	if (frmScreen.cboCustomerName.selectedIndex != -1)
	{
		PrintXML.SetAttribute("CUSTOMERNAME", frmScreen.cboCustomerName.options(frmScreen.cboCustomerName.selectedIndex).text);
	}

	if (frmScreen.cboRecipient.selectedIndex != -1)
	{
		PrintXML.SetAttribute("RECIPIENTNAME", frmScreen.cboRecipient.options(frmScreen.cboRecipient.selectedIndex).text);
	}

	if(m_sEmailAdministrator != null)
	{
		PrintXML.SetAttribute("EMAILADMINISTRATOR", m_sEmailAdministrator);
	}
	
	<%/* AS 15/12/2006 EP1269 Gemini Printing: second release to system test. */%>
	PrintXML.SetAttribute("GEMINIPRINTSTATUS", GEMINIPRINTSTATUS_APPROVED);	
}

function getPrinterTypeForAttributes(printerDestination)
{
	var remotePrinterLocation = null;
	var attribute = null;
	
	if (m_sAttributesXML != null && m_sAttributesXML != "")
	{
		printerDestination = (attribute = m_sAttributesXML.XMLDocument.selectSingleNode("RESPONSE/ATTRIBUTES/@PRINTERDESTINATIONTYPE")) != null ? attribute.text : null;
		remotePrinterLocation = (attribute = m_sAttributesXML.XMLDocument.selectSingleNode("RESPONSE/ATTRIBUTES/@REMOTEPRINTERLOCATION")) != null ? attribute.text : null;
	}
	else
	{
		remotePrinterLocation = (attribute = AttribsXML.XMLDocument.selectSingleNode("RESPONSE/ATTRIBUTES/@REMOTEPRINTERLOCATION")) != null ? attribute.text : null;
	}
	
	var userPrinterLocation = null;
	var outputLocationSelectedIndex = frmScreen.cboOutputLocation.selectedIndex;
	if (outputLocationSelectedIndex != -1)
	{
		userPrinterLocation = frmScreen.cboOutputLocation.options(outputLocationSelectedIndex).text;
	}
	
	return getPrinterType(m_BaseNonPopupWindow, printerDestination, userPrinterLocation, remotePrinterLocation);
}

function BuildPrintDataBlock(PrintXML)
{
	xmlPrintDataNode = PrintXML.CreateActiveTag("PRINTDATA");
	PrintXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	PrintXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAFFNumber);
	
	if (m_sAttributesXML == "" || m_sAttributesXML == null)
	{
		var strCustReq = AttribsXML.GetTagAttribute("ATTRIBUTES", "CUSTOMERSPECIFICIND");
		if(strCustReq == "1")
		{
			PrintXML.SetAttribute("CUSTOMERNUMBER", frmScreen.cboCustomerName.value);
			PrintXML.SetAttribute("CUSTOMERVERSIONNUMBER",frmScreen.cboCustomerName.options.item(frmScreen.cboCustomerName.selectedIndex).getAttribute("attCustVersionNumber"));
		}
		else
		{
			<% // MAR906 - Peter Edney - 05/01/2006
			// PrintXML.SetAttribute("CUSTOMERNUMBER", "1");
			// PrintXML.SetAttribute("CUSTOMERVERSIONNUMBER", "1"); %>
			<% /* PB 13/07/2006 EP986 Begin */ %>
			if( m_sCustomerNumber != null && m_sCustomerNumber != "" )
			{
				PrintXML.SetAttribute("CUSTOMERNUMBER", m_sCustomerNumber);
			}
			if( m_sCustomerVersionNumber != null && m_sCustomerVersionNumber != "" )
			{
				PrintXML.SetAttribute("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
			}
			<% /* PB EP986 End */ %>
		}
	}
	else
	{
		if (frmScreen.cboCustomerName.selectedIndex != -1)
		{
			PrintXML.SetAttribute("CUSTOMERNAME", frmScreen.cboCustomerName.options(frmScreen.cboCustomerName.selectedIndex).text);
		}
		
		PrintXML.SetAttribute("CUSTOMERNO", frmScreen.cboCustomerName.value);
		
		var strCustReq = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "CUSTOMERSPECIFICIND");
		if( (strCustReq == "1") && (frmScreen.cboCustomerName.selectedIndex != -1) )
		{
			PrintXML.SetAttribute("CUSTOMERNUMBER", frmScreen.cboCustomerName.value);
			PrintXML.SetAttribute("CUSTOMERVERSIONNUMBER",frmScreen.cboCustomerName.options.item(frmScreen.cboCustomerName.selectedIndex).getAttribute("attCustVersionNumber"));
		}
		else
		{
			<% // MAR906 - Peter Edney - 05/01/2006
			// PrintXML.SetAttribute("CUSTOMERNUMBER", "1");
			// PrintXML.SetAttribute("CUSTOMERVERSIONNUMBER", "1"); %>
			<% /* PB 17/07/2006 EP986 Begin */ %>
			if( m_sCustomerNumber != null && m_sCustomerNumber != "" )
			{
				PrintXML.SetAttribute("CUSTOMERNUMBER", m_sCustomerNumber);
			}
			if( m_sCustomerVersionNumber != null && m_sCustomerVersionNumber != "" )
			{
				PrintXML.SetAttribute("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
			}
			<% /* PB 17/07/2006 EP986 End */ %>
		}
	}
	
	if(!(frmScreen.cboCustomerName.selectedIndex < 0))
	{
		if 	(frmScreen.cboCustomerName.options.item(frmScreen.cboCustomerName.selectedIndex).getAttribute("attCustVersionNumber") != null)
		{
			PrintXML.SetAttribute("CUSTOMERVERSIONNUMBER", frmScreen.cboCustomerName.options.item(frmScreen.cboCustomerName.selectedIndex).getAttribute("attCustVersionNumber"));
		}
	}
	
	PrintXML.SetAttribute("RECIPIENTKEY",frmScreen.cboRecipient.value);
	
	if (m_sAttributesXML == "" || m_sAttributesXML == null)
	{
		PrintXML.SetAttribute("METHODNAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "PDMMETHOD"));
		PrintXML.SetAttribute("RECIPIENTTYPE", AttribsXML.GetTagAttribute("ATTRIBUTES", "RECIPIENTTYPE"));
	}
	else
	{
		PrintXML.SetAttribute("METHODNAME", m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "PDMMETHOD"));
		PrintXML.SetAttribute("RECIPIENTTYPE", m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "RECIPIENTTYPE"));
	}
	
	var TempRecTypeXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("PrintRecipientType");
	var sRecipientList = TempRecTypeXML.GetComboLists(document,sGroups);
	
	if (m_sAttributesXML == "" || m_sAttributesXML == null)
		var sValueID = AttribsXML.GetTagAttribute("ATTRIBUTES", "RECIPIENTTYPE");
	else
		var sValueID = m_sAttributesXML.GetTagAttribute("ATTRIBUTES", "RECIPIENTTYPE");
	
	var sRecipientPrintType = TempRecTypeXML.GetTagText("LIST/LISTNAME/LISTENTRY[VALUEID='" + sValueID + "']/VALUENAME");
	
	PrintXML.SetAttribute("REFTYPE", sRecipientPrintType);
}
<% /* MAR7 End */ %>

-->
</script>
</body>
</html>
