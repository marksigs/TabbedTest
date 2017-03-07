<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      cm065.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Stored Quotes Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		17/03/00	Initial Revision
AY		29/03/00	New top menu/scScreenFunctions change
MCS		26/05/99	SYS0724,SYS0759
CL		05/03/01	SYS1920 Read only functionality added
IK		22/03/01	SYS2145 Add idNoCompleteCheck
LD		23/05/02	SYS4727 Use cached versions of frame functions

BG		28/06/02	SYS4767 MSMS/Core integration
STB		28/02/02	SYS4144 Added APR column.
STB		25/04/02	MSMS0019 Reinstate button only enabled if selected quote not active.
BG		28/06/02	SYS4767 MSMS/Core integration
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

PSC		29/06/2005	MAR5 - KFI Printing
GHun	15/07/2005	MAR7 Integrate local printing
HMA     01/08/2005  MAR18  Change routing for KFI Quote
MV		19/09/2005	MAR35  Amended PopulateScreen to populate LastOffered Column 
MV		21/09/2005	MAR44  Amended PopulateTable(),ResetBlankRows(),BuildControlDataBlock(),
						   frmScreen.btnViewKFI.onclick(),GetApplicationRegulationIndicator(),
						   SetCheckedLastOffered(),SetUnCheckedLastOffered()
MV		30/09/2005	MAR44  Code Review fixes
HMA     10/10/2005  MAR136 Use AQFindStoredQuoteDetails.
                           Changed selection of Print Quote check box.
TW		28/10/2005	MAR211 Add KFI Required checkbox,
					Disable Print button as per MARS007 Change control
					Update quote information for selected quotes
MV		03/11/2005	MAR167	 Amended to enable viewKFI Button on seleting a row in the table
Maha T	18/11/2005	MAR273	Include Validation_Init() to handle back space.
GHun	05/12/2005	MAR796 GHun	turn off document compression
HMA     07/03/2006  MAR1103  Correct authority level checks on Print button and check box.
PJO     08/03/2006  MAR1359 Add progress message while generating KFI
PJO     09/03/2006  MAR1359 Parameterise Progress Message width
JD		10/03/2006	MAR1061 Added Amt Req, Prop Price and LTV to listbox. Added critical datacheck to reinstatequotation
GHun	20/03/2006	MAR1453 Close progress window before displaying error messages
JD		23/03/2006	MAR1518 set quotationNumber and mortgageSubQuoteNumber for viewKFI
HMA     28/03/2006  MAR1500 Check for changed rates on entry and on ViewKFI
IK		31/05/2006	EP500	(temporarily) hide KFI buttons
PSC		28/11/2006	EP2_218 Back out EP500
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex="-1" type="text/x-scriptlet" width="1" viewastext></object>
<% /* Scriptlets - remove any which are not required */ %>
<span style="LEFT: 310px; POSITION: absolute; TOP: 266px">
<object data="scTableListScroll.asp" id="scTable" style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex="-1" type="text/x-scriptlet" viewastext></object>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToCM010" method="post" action="CM010.asp" style="DISPLAY: none"></form>
<form id="frmToCM060" method="post" action="CM060.asp" style="DISPLAY: none"></form>
<form id="frmToMN060" method="post" action="MN060.asp" style="DISPLAY: none"></form>


<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen">
<div id="divBackground" style="HEIGHT: 340px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 670px" class="msgGroup">
	<div id="spnStoredQuotes" style="LEFT: 4px; POSITION: absolute; TOP: 4px; WIDTH: 650px">
		<table id="tblStoredQuotes" width="645" border="0" cellspacing="0" cellpadding="0" class="msgTable" style ="PADDING-LEFT: 2px; PADDING-RIGHT: 2px">
			<col width="2%"/>
			<col width="10%"/>
			<col width="5%"/>
			<col width="8%"/>
			<col width="8%"/>
			<col width="3%"/>
			<col width="9%"/>
			<col width="5%"/>
			<col width="5%"/>
			<col width="5%"/>
			<col width="4%"/>
			<col width="5%"/>
			<col width="5%"/>
			<col width="5%"/>
			<col width="9%"/>
			<col width="4%"/>
			<col width="4%"/>
			<col width="4%"/>
			<tr id="rowTitles">	
				<td class="TableHead">No.</td>
				<td class="TableHead">Product Details</td>
				<td class="TableHead">Repay Type</td>
				<td class="TableHead">Amt Req</td>	
				<td class="TableHead">Prop price</td>
				<td class="TableHead">LTV</td>
				<td class="TableHead">Mtg Cost</td>		
				<td class="TableHead">B&amp;C Cost</td>			
				<td class="TableHead">PP Cost</td>		
				<td class="TableHead">Total Cost</td>		
				<td class="TableHead">APR</td>
				<td class="TableHead">Act?</td>
				<td class="TableHead">Rec?</td>
				<td class="TableHead">Acc?</td>
				<td class="TableHead">Cost at Final Rate</td>
				<td class="TableHead">Last Offered</td>
				<td class="TableHead">Print Quote</td>
				<td class="TableHead">KFI Reqd</td>
			</tr>
			<!-- row data -->
			<tr id="row01">
				<td class="TableTopLeft">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td ondblclick="SelectPrintQuote()" class="TableTopCenter">&nbsp;</td>
				<td class="TableTopRight" ondblclick="SelectKFIRequired()">&nbsp;</td>
			</tr>
			<tr id="row02">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td ondblclick="SelectPrintQuote()">&nbsp;</td>
				<td class="TableRight" ondblclick="SelectKFIRequired()">&nbsp;</td>
			</tr>
			<tr id="row03">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td ondblclick="SelectPrintQuote()">&nbsp;</td>
				<td class="TableRight" ondblclick="SelectKFIRequired()">&nbsp;</td>
			</tr>
			<tr id="row04">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td ondblclick="SelectPrintQuote()">&nbsp;</td>
				<td class="TableRight" ondblclick="SelectKFIRequired()">&nbsp;</td>
			</tr>
			<tr id="row05">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td ondblclick="SelectPrintQuote()">&nbsp;</td>
				<td class="TableRight" ondblclick="SelectKFIRequired()">&nbsp;</td>
			</tr>
			<tr id="row06">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td ondblclick="SelectPrintQuote()">&nbsp;</td>
				<td class="TableRight" ondblclick="SelectKFIRequired()">&nbsp;</td>
			</tr>
			<tr id="row07">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td ondblclick="SelectPrintQuote()">&nbsp;</td>
				<td class="TableRight" ondblclick="SelectKFIRequired()">&nbsp;</td>
			</tr>
			<tr id="row08">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td ondblclick="SelectPrintQuote()">&nbsp;</td>
				<td class="TableRight" ondblclick="SelectKFIRequired()">&nbsp;</td>
			</tr>
			<tr id="row09">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td ondblclick="SelectPrintQuote()">&nbsp;</td>
				<td class="TableRight" ondblclick="SelectKFIRequired()">&nbsp;</td>
			</tr>
			<tr id="row10">
				<td class="TableBottomLeft">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td ondblclick="SelectPrintQuote()" class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomRight" ondblclick="SelectKFIRequired()">&nbsp;</td>
			</tr>
		</table>
	</div>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 236px">
		<input id="btnDetails" value="Details" type="button" style="WIDTH: 60px" class="msgButton" disabled>
		<span style="LEFT: 70px; POSITION: absolute; TOP: 0px">
			<input id="btnReinstate" value="Reinstate" type="button" style="WIDTH: 60px" class="msgButton" disabled>
		</span>
		<% /* EP500 IK */ %>
		<% /* PSC 28/11/2006 EP2_218 */ %>
		<span style="LEFT: 140px; POSITION: absolute; TOP: 0px;">
			<input id="btnPrint" value="Print" type="button" style="WIDTH: 60px" class="msgButton" disabled>
		</span>
		<% /* EP500 IK */ %>
		<% /* PSC 28/11/2006 EP2_218 */ %>
		<span style="LEFT: 210px; POSITION: absolute; TOP: 0px;">
			<input id="btnViewKFI" value="View KFI" type="button" style="WIDTH: 80px" class="msgButton" disabled>
		</span>		
	</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px">
<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cm065Attribs.asp" -->
<% /* MAR7 GHun */ %>
<script src="includes/Documents.js" language="javascript" type="text/javascript"></script>
<% /* MAR7 End */ %>

<% /* Specify Code Here */ %>
<script language="JScript" type="text/javascript">
<!--
var ListXML = null;
var m_sApplicationMode = null;
var m_sReadOnly = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sActiveQuotationNumber = null;
<%/* PSC 29/06/2005 MAR5 - Start */%>
var m_iActiveLocation = 0;
var m_iAcceptedLocation = 0;
<%/* PSC 29/06/2005 MAR5 - End */%>
var scScreenFunctions;
var m_blnReadOnly = false;
<%/* PSC 29/06/2005 MAR5 - Start */%>
var m_iTableLength = 10;    
var m_aCheckBoxArray;
var m_iPrintsRequired = 0;
var m_iCurrentMaximumNoOfCustomers = 5;
var m_sRegulationIndicator = "";
<%/* PSC 29/06/2005 MAR5 - End */%>

<% /* BG		28/06/02	SYS4767 MSMS/Core integration*/%>
<% /* MSMS0019 - Use a constant for the column as its referred to multiple times. */ %>
//BMIDS00209 - altered from 8 -7 by DPF 29/07/2002
<%/* PSC 29/06/2005 MAR5 - Start */%>
var COLUMN_ACTIVE_QUOTE = 11;
var COLUMN_REC_QUOTE = 12;
var COLUMN_ACC_QUOTE = 13;
<%/* PSC 29/06/2005 MAR5 - End */%>
<% /* BG		28/06/02	SYS4767 MSMS/Core integration - END*/%>
	

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;

<% /* MAR7 GHun */ %>
var m_aQuotationNumberArray;
var m_aSubQuoteArray;
var m_aLastOfferedQuoteArray;

<% /* PSC 28/11/2006 EP2_218 */ %>
var sDocumentID;
var m_deliveryType = "";
var m_compressionMethod  = "";
var m_xmlLocalPrinters = null;
var iNumberOfCopies = 1;
var m_printerType = "";
var sQuotationNumber = "";
var sMortgageSubQuoteNumber = ""; 
var m_defaultPrinter = "";
var bViewBeforePrint;
var xmlControlDataNode = null;
var xmlPrintDataNode = null;
var xmlTemplateDataNode = null;
var m_UserId = null;
var m_UnitId = null;
var m_DistributionChannelId = null;
var m_MachineId = null;
<% /* MAR7 End */ %>

<% /* MAR211 */%>
var m_aKFIReqdArray;
var iSQPrintButtonAuthorityLevel = 0;
var m_AccessType;
<% /* MAR211End  */%>

var bViewKFIClick = false;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	<%/* PSC 29/06/2005 MAR5 - Initialise CheckBox array */%>
	m_aCheckBoxArray = new Array();
	<% /* MAR7 GHun Initialise arrays */ %>
	m_aQuotationNumberArray = new Array();
	m_aSubQuoteArray = new Array();
	<% /* MAR7 End */ %>
	m_aLastOfferedQuoteArray = new Array();

	<% /* MAR211 */ %>
	m_aKFIReqdArray = new Array();
	var GlobalXML = null;
	GlobalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	iSQPrintButtonAuthorityLevel = GlobalXML.GetGlobalParameterAmount(document,'SQPrintButtonAuthorityLevel');
	<% /* MAR211End  */ %>
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	<% /* MAR273 - Maha T */ %>
	Validation_Init();
	
	<%/* PSC 29/06/2005 MAR5 - Start */%>
	var sGroups = new Array("RepaymentType");
	var XMLTemp = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	XMLTemp.GetComboLists(document,sGroups);
	m_XMLRepay = XMLTemp.GetComboListXML("RepaymentType")
	<%/* PSC 29/06/2005 MAR5 - End */%>
		
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Stored Quotes","CM065",scScreenFunctions);
	
	RetrieveContextData();
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CM065");
	
	PopulateScreen();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	//alert(m_blnReadOnly);
	<% // GD BMIDS000376 START %>
	if (m_blnReadOnly == true || m_AccessType < iSQPrintButtonAuthorityLevel)
	{
		frmScreen.btnReinstate.disabled = true;
		frmScreen.btnPrint.disabled = true; <%/* PSC 29/06/2005 MAR5 */%>
		frmScreen.btnViewKFI.disabled = true;
	}
	
	 <% // GD BMIDS000376 END %>	
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
	m_sApplicationMode = scScreenFunctions.GetContextParameter(window, "idApplicationMode", "Cost Modelling");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window, "idReadOnly", 0);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", "");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", "");
	m_sActiveQuotationNumber  = scScreenFunctions.GetContextParameter(window, "idQuotationNumber", "");
	
	<% /* MAR7 GHun */ %>
	m_UserId = scScreenFunctions.GetContextParameter(window,"idUserId",null);
	m_UnitId = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	m_DistributionChannelId = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
	m_MachineId = scScreenFunctions.GetContextParameter(window,"idMachineId",null);
	<% /* MAR7 End */ %>

	<% /* MAR1103 Use UserRole.Role to check authority levels */ %>
	m_AccessType = parseInt(scScreenFunctions.GetContextParameter(window, "idRole", "0"));
}

function spnStoredQuotes.onclick()
{
	if (scTable.getRowSelected() != -1) 
	{
		frmScreen.btnDetails.disabled = false;
		
		<% /* BG		28/06/02	SYS4767 MSMS/Core integration*/%>
		<% /* MSMS0011 - Only enable the reinstate button if the selected quote isn't active. */ %>
		if ((m_blnReadOnly != true) && (tblStoredQuotes.rows(scTable.getRowSelected()).cells(COLUMN_ACTIVE_QUOTE).innerText != "Yes"))
		{
			frmScreen.btnReinstate.disabled = false;
		}
		else
		{
			frmScreen.btnReinstate.disabled = true;
		}
		
		//disable reinstate when readonly
		if (m_blnReadOnly == true)
		{
			frmScreen.btnReinstate.disabled = true;
		}
		
		//disable reinstate button if it is cancel decline stage.
		//bmids 468
		if(scScreenFunctions.IsCancelDeclineStage(window))
		{
			frmScreen.btnReinstate.disabled=true;
		}
		
		<%/*if(m_sReadOnly != "1") frmScreen.btnReinstate.disabled = false;*/%>
		<% /* BG		28/06/02	SYS4767 MSMS/Core integration - END*/%>
		<% /* JD MAR1518 set quote and subquote number for viewKFI */ %>
		sMortgageSubQuoteNumber = tblStoredQuotes.rows(scTable.getRowSelected()).getAttribute("MortgageSubQuoteNumber");
		sQuotationNumber = tblStoredQuotes.rows(scTable.getRowSelected()).cells(0).innerText;
		frmScreen.btnViewKFI.disabled = false;
		
	}
}

function PopulateScreen()
{
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(window,"SEARCH");
	ListXML.CreateActiveTag("QUOTATION");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	
	<% /* MAR18 Remove Quick Quote */ %>
	<% /* MAR44 Remove AQ 
	<% /* MAR136 Do need to use AQFindStoredQuoteDetails because it brings back Application Data as well
	             as Stored Quote data */ %>
	ListXML.RunASP(document,"AQFindStoredQuoteDetails.asp");
	
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);
	if(sResponseArray[0])
	{
		PopulateTable(0);
		ListXML.ActiveTag = null;
		ListXML.CreateTagList("STOREDQUOTATION");
		scTable.initialiseTable(tblStoredQuotes,0,"",PopulateTable,10,ListXML.ActiveTagList.length);
		if(scScreenFunctions.GetContextParameter(window,"idMetaAction","") == "fromCM060")
		{
			sQuotationNumber = scScreenFunctions.GetContextParameter(window,"idXML","");
			var nTotal = 0;
			var bFound = false;
			while(nTotal <= ListXML.ActiveTagList.length && !bFound)
			{
				for(var nLoop = 1;nLoop <= 10;nLoop++)
				{
					if(tblStoredQuotes.rows(nLoop).cells(0).innerText == sQuotationNumber)
					{
						bFound = true;
						scTable.setRowSelected(nLoop);
					}
					nTotal++;
				}
				
				if(!bFound && nTotal < ListXML.ActiveTagList.length)
					scTable.pageDown();
			}
			
			spnStoredQuotes.onclick();
		}
		else
		{
			<% /* MAR1500 Check if the Active Quote has consistent rates */ %>
			sQuotationNumber = ""
			if (CheckForChangedRates(m_sApplicationNumber, m_sApplicationFactFindNumber, sQuotationNumber) == true)
			{
				alert("The Rates have changed since the accepted/active quote was created and the details held may be inconsisent. Please create a quote ");
			}
		}
	}
}

<%/* PSC 29/06/2005 MAR5 - Start */%>
function PopulateTable(nStart)
{
	ListXML.ActiveTag = null;
	var xmlAPPLICATIONFACTFIND = ListXML.SelectTag(null,"APPLICATIONFACTFIND");

	if(xmlAPPLICATIONFACTFIND != null)
	{	
		sActiveQuoteNumber =  ListXML.GetTagText("ACTIVEQUOTENUMBER");
		
		if(sActiveQuoteNumber != m_sActiveQuotationNumber)
		{
			sActiveQuoteNumber = m_sActiveQuotationNumber;
		}
		
		sRecommendedQuoteNumber =  ListXML.GetTagText("RECOMMENDEDQUOTENUMBER");
		sAcceptedQuoteNumber =  ListXML.GetTagText("ACCEPTEDQUOTENUMBER");
		m_sRegulationIndicator = ListXML.GetTagText("REGULATIONINDICATOR");
	}
	ListXML.ActiveTag = null;
	var TagListSTOREDQUOTATION = ListXML.CreateTagList("STOREDQUOTATION");
	//var sQuotationNumber;		MAR7
	var sActiveQuoteNumber;
	var sRecommendedQuoteNumber;
	var sAcceptedQuoteNumber;
	var sMortgageSubQuoteNumber;	<% /* MAR7 GHun */ %>
		
	<% /* Added for testing only !!
	if (m_sActiveQuotationNumber != sActiveQuoteNumber)
		alert ("The Active Quotation Number in Context is not the same as the Active Quotation Number returned from the search");
	*/ %>
	
	<% /* Populate the table with a set of records, starting with record nStart */ %>
	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) && nLoop < 10; nLoop++)
	{			
		<% /* MAR7 GHun */ %>
		sMortgageSubQuoteNumber = ListXML.GetTagText("MORTGAGESUBQUOTENUMBER");
		tblStoredQuotes.rows(nLoop+1).setAttribute("MortgageSubQuoteNumber", sMortgageSubQuoteNumber);
		m_aSubQuoteArray[nLoop + nStart] = sMortgageSubQuoteNumber;
		<% /* MAR7 End */ %>
		
		sQuotationNumber = ListXML.GetTagText("QUOTATIONNUMBER");					
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(0),sQuotationNumber );
		//BMIDS00209 - removed column (moved cells to the right across one).  DPF 29/07/2002 
		//scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(3), ListXML.GetTagText("LIFECOVER"));
		m_aQuotationNumberArray[nLoop + nStart] = sQuotationNumber; <% /* MAR7 GHun */ %>
		
		m_aLastOfferedQuoteArray[nLoop + nStart] = ListXML.GetTagText("LASTOFFERED");
		
		<% /* MAR211 TW */ %>
		m_aKFIReqdArray[nLoop + nStart] = ListXML.GetTagText("KFIREQUIREDINDICATOR");
		<% /* MAR211 TW End*/ %>
		
		<% /* MAR1061 amt req, Prop Price, LTV*/ %>
		//Amt Req
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(3),ListXML.GetTagText("AMOUNTREQUESTED") );
		
		//Prop Price
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(4),ListXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE") );
		
		//ltv
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(5),top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("LTV"), 2) );
		
		//Mtg Cost
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(6),top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("TOTALNETMONTHLYCOST"), 2) );
		
		//B&C Cost
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(7),top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("TOTALBCMONTHLYCOST"), 2) );

		//PP Cost
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(8),top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("TOTALPPMONTHLYCOST"), 2) );
				
		//TOTAL COST				
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(9),top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("TOTALQUOTATIONCOST"), 2) );

		//APR
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(10),top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("APR"), 2) );

		var sAQN = "";		
		if(sQuotationNumber == sActiveQuoteNumber) 
		{
			sAQN = "Yes";
			m_iActiveLocation = nLoop+1;
		}
		
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(COLUMN_ACTIVE_QUOTE),sAQN);
			
		var sRQN = "";
		if(sQuotationNumber == sRecommendedQuoteNumber) sRQN = "Yes";

		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(COLUMN_REC_QUOTE),sRQN);
		
		var sAccQN = "";
		if(sQuotationNumber == sAcceptedQuoteNumber)
		{
			sAccQN = "Yes";
			m_iAcceptedLocation = nLoop + 1; 
		}	
		
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(COLUMN_ACC_QUOTE),sAccQN);

		var sRepayTypeText = "Not applicable";
		var sRepayTypeValidation = "";
		
		var sProductDetail = "";
		var nComponents = "0";
		var dblFinalRMC;
		var nFinalRateCost = 0;  <% /* BMIDS00224 - DPF 30/07/02 - added FINALRATEMONTHLYCOST */ %>
		var sPPBreakdown = "";
		var dblInterestOnly;
		var nTotalInterestOnly = 0;
		var dblCapitalInterest;
		var nTotalCapitalInterest = 0;
		
		ListXML.CreateTagList("LOANCOMPONENT");
		for(var nComponentLoop = 0;ListXML.SelectTagListItem(nComponentLoop);nComponentLoop++)
		{
			nComponents++;
			var sThisRepayTypeText = ListXML.GetTagAttribute("REPAYMENTMETHOD","TEXT");
			var sThisRepayTypeValidation = ListXML.GetTagText("REPAYMENTMETHOD");
			
			var sSearch = "LISTNAME/LISTENTRY[VALUEID='" + sThisRepayTypeValidation + "']/VALIDATIONTYPELIST/VALIDATIONTYPE";
			sThisRepayTypeValidation = m_XMLRepay.selectSingleNode(sSearch).text;
		
			var sThisProductDetail = ListXML.GetTagText("PRODUCTNAME");

			dblFinalRMC = parseFloat(ListXML.GetTagText("FINALRATEMONTHLYCOST"));
			if ( ! isNaN(dblFinalRMC) )
			{
				nFinalRateCost = nFinalRateCost + dblFinalRMC;
			}
			
			// Repayment Type can be Interest Only(I), Capital & Interest (C) or Part and Part(P)
			// If Repayment Type Validation is different for different Loan Components, set it to "Multiple" overall (MARS1061 - was set to P&P previously)
			
			// If the Validation Type is the same for different Loan Components but the text is different,
			// set the text to be a general value.
			
			if(sThisRepayTypeValidation != "" && sRepayTypeText != "Part and Part") // Already decided this is P&P
			{
				if(sRepayTypeValidation != "")
				{
					if (sThisRepayTypeValidation != sRepayTypeValidation) // Validation Types differ - set to P&P
					{
						sRepayTypeText = "Multiple"; //MAR1061
					}
					else if (sThisRepayTypeValidation == sRepayTypeValidation) // Validation Types the same - check text
					{
						if (sThisRepayTypeText != sRepayTypeText)
						{	
							if (sThisRepayTypeValidation == "I") sRepayTypeText = "I/O"; //MAR1061
							else if (sThisRepayTypeValidation == "C") sRepayTypeText = "C&I"; //MAR1061
							else sRepayTypeText = "P&P"; //MAR1061
						}
					}	
				}
				else
				{
					sRepayTypeValidation = sThisRepayTypeValidation;
					sRepayTypeText = sThisRepayTypeText;
				}									
			}	
			
			//Keep a running total of Interest Only and Capital & Interest amounts.
			dblLoanAmount = parseFloat(ListXML.GetTagText("LOANAMOUNT"));
			dblInterestOnly = parseFloat(ListXML.GetTagText("INTERESTONLYELEMENT"));
			dblCapitalInterest = parseFloat(ListXML.GetTagText("CAPITALANDINTERESTELEMENT"));
			
			if (sThisRepayTypeValidation == "I")
			{	
				// Loan Amount holds the Interest Only element
				if (!isNaN(dblLoanAmount)) nTotalInterestOnly = nTotalInterestOnly + dblLoanAmount;
			}
			else if (sThisRepayTypeValidation == "C")
			{
				// Loan Amount holds the Capital & Interest element
				if (!isNaN(dblLoanAmount)) nTotalCapitalInterest = nTotalCapitalInterest + dblLoanAmount; 	
			}
			else
			{
				// Part & Part : Interest Only and Capital&Interest amounts are held separately
				if (!isNaN(dblInterestOnly))
					nTotalInterestOnly = nTotalInterestOnly + dblInterestOnly;
				if (!isNaN(dblCapitalInterest))
					nTotalCapitalInterest = nTotalCapitalInterest + dblCapitalInterest;
			}
						
			// If Product Name is different for different Loan Components, set it to "Not Applicable"
			if (sProductDetail == "" && sThisProductDetail != "")
			{
				sProductDetail = sThisProductDetail;
			}
			else
			{
				if (sProductDetail != "Not Applicable" && sThisProductDetail != sProductDetail) 
				{
					sProductDetail = "Multiple Component";
				}
			}		
		}

		ListXML.ActiveTagList = TagListSTOREDQUOTATION;		
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(1),sProductDetail);
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(2),GetShortVersion(sRepayTypeText));
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(14),top.frames[1].document.all.scMathFunctions.RoundValue(nFinalRateCost, 2));
		
		//Tool Tip 
		if (sRepayTypeText == "P&P") //MAR1061
		{
			sPPBreakdown = "I/O £" + nTotalInterestOnly + " - C/I £" + nTotalCapitalInterest;
			tblStoredQuotes.rows(nLoop+1).cells(2).title = sPPBreakdown;
		}	
		
		<% /* MAR44 MV   */ %>
		
		<% /* MAR35 MV */%>
		if (m_aLastOfferedQuoteArray[nLoop + nStart] == "1")
			SetChecked(tblStoredQuotes.rows(nLoop+1).cells(15), nLoop+1);
		else
			SetUnChecked(tblStoredQuotes.rows(nLoop+1).cells(15), nLoop+1);
			
		
		//Initialise Print Quote check box
		<% /* MAR7 GHun  Allow for scrolling */ %>
		if (m_aCheckBoxArray[nLoop + nStart] == "1")
		{
			SetChecked(tblStoredQuotes.rows(nLoop+1).cells(16), nLoop+1);
		}
		else
		{ 
		<% /* MAR7 End */ %>
			SetUnChecked(tblStoredQuotes.rows(nLoop+1).cells(16), nLoop+1);
		}
		
		<% /* MAR211 TW */%>
		if (m_aKFIReqdArray[nLoop + nStart] == "1")
			SetCheckedEndColumn(tblStoredQuotes.rows(nLoop+1).cells(17), nLoop+1);
		else
			SetUnCheckedEndColumn(tblStoredQuotes.rows(nLoop+1).cells(17), nLoop+1);
		<% /* MAR211 TW End */%>
	}
	if(nLoop < (m_iTableLength))
	{
		ResetBlankRows(nLoop);
	}	
}
<%/* PSC 29/06/2005 MAR5 - End */%>

function GetShortVersion(sRepayTypeText)
{
	var sShortVersion = "";
	switch (sRepayTypeText)
	{
		case "Interest Only": sShortVersion = "I/O"; break;
		case "Capital and Interest": sShortVersion = "C&I"; break;
		case "Part and Part": sShortVersion = "P&P"; break;
		default : sShortVersion = sRepayTypeText;
	}
	return sShortVersion;
}
function frmScreen.btnReinstate.onclick()
{

	<% /* BG		28/06/02	SYS4767 MSMS/Core integration*/%>
	if(tblStoredQuotes.rows(scTable.getRowSelected()).cells(COLUMN_ACTIVE_QUOTE).innerText == "Yes")
	<%/*if(tblStoredQuotes.rows(scTable.getRowSelected()).cells(7).innerText == "Yes")*/%>
	<% /* BG		28/06/02	SYS4767 MSMS/Core integration - END*/%>
	
	{
		alert("The selected quotation is already active");//error 215
		return;
	}
		
	//	AW	18/11/2002	BMIDS00971
	//	Clear down any accepted quote if the one being re-instated is not an accepted quote
	if(tblStoredQuotes.rows(scTable.getRowSelected()).cells(COLUMN_ACC_QUOTE).innerText != "Yes")
	{
		var ApplicationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
		ApplicationXML.CreateRequestTag(window,"UPDATE");
		ApplicationXML.CreateActiveTag("APPLICATIONFACTFIND");
		ApplicationXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		ApplicationXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);		
		ApplicationXML.CreateTag("ACCEPTEDQUOTENUMBER", "");		
		ApplicationXML.RunASP(document,"UpdateApplicationFactFind.asp");
		ApplicationXML.IsResponseOK();
	
		<%/* PSC 29/06/2005 MAR5 - Start */%>
		if (m_iAcceptedLocation != 0)
		{
			scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(m_iAcceptedLocation).cells(COLUMN_ACC_QUOTE),"");
		}	
		<%/* PSC 29/06/2005 MAR5 - End */%>
	
	}
	//	AW	18/11/2002	BMIDS00971 - End
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"UPDATE");
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	XML.CreateTag("QUOTATIONNUMBER",tblStoredQuotes.rows(scTable.getRowSelected()).cells(0).innerText);
	
	//MAR1061 add critical data check
	XML.SelectTag(null,"REQUEST");
	XML.SetAttribute("OPERATION","CriticalDataCheck"); 
	XML.CreateActiveTag("CRITICALDATACONTEXT");
	XML.SetAttribute("APPLICATIONNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null));
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null));
	XML.SetAttribute("SOURCEAPPLICATION","Omiga");
	XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
	XML.SetAttribute("ACTIVITYINSTANCE","1");
	XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
	XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
	XML.SetAttribute("COMPONENT","omAQ.ApplicationQuoteBO");
	XML.SetAttribute("METHOD","ReinstateQuotation");
			
	window.status = "Critical Data Check - please wait";
					
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"OmigaTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
	}

	window.status = "";

	<% /* MAR18 Remove Quick Quote */ %>

	if( XML.IsResponseOK() == true)
	{
		//Set ActiveQuoteNumber in context = Quote No. from quotation selected in <list> StoredQuotations.
				
		m_sActiveQuotationNumber =	tblStoredQuotes.rows(scTable.getRowSelected()).cells(0).innerText;
		scScreenFunctions.SetContextParameter(window,"idQuotationNumber",tblStoredQuotes.rows(m_sActiveQuotationNumber));
		//Locate quotation with Active? = "Yes" in <list> StoredQuotations and set Active? = blank.

		
		<% /* BG		28/06/02	SYS4767 MSMS/Core integration*/%>
		<%/* PSC 29/06/2005 MAR5 - Start */%>
		if (m_iActiveLocation != 0)
		{
			scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(m_iActiveLocation).cells(COLUMN_ACTIVE_QUOTE),"");
		}	
		<%/* PSC 29/06/2005 MAR5 - End */%>
		<%/*scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(m_iActiveLocation).cells(7),"");*/%>
		<% /* BG		28/06/02	SYS4767 MSMS/Core integration - END*/%>
			
		m_iActiveLocation = scTable.getRowSelected();
		
		<% /* BG		28/06/02	SYS4767 MSMS/Core integration*/%>
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(m_iActiveLocation).cells(COLUMN_ACTIVE_QUOTE),"Yes");	
		<%/*scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(m_iActiveLocation).cells(7),"Yes");	*/%>
		<% /* BG		28/06/02	SYS4767 MSMS/Core integration - END*/%>
		
		<% /* BMIDS624 */ %>
		// JD BMIDS749 removed   CheckForChangedRates(m_sApplicationNumber, m_sApplicationFactFindNumber);
		
		//Set Active? = "Yes" for quotation selected in <list> StoredQuotations.

	}
}	

function frmScreen.btnDetails.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","fromCM065");
<%	// Not strictly speaking XML, but it will be used for XML purposes!
%>	scScreenFunctions.SetContextParameter(window,"idXML",tblStoredQuotes.rows(scTable.getRowSelected()).cells(0).innerText);
	<% /* BM0176 MDC 19/12/2002 */ %>
	scScreenFunctions.SetContextParameter(window,"idXML2",tblStoredQuotes.rows(scTable.getRowSelected()).getAttribute("MortgageSubQuoteNumber"));
	<% /* BM0176 MDC 19/12/2002 */ %>
	
	frmToCM060.submit();
}

function btnSubmit.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","");
	scScreenFunctions.SetContextParameter(window,"idNoCompleteCheck","1");
	<% /* MAR18 If we have come here from MQ010, route to MN060 */ %>
	if (scScreenFunctions.GetContextParameter(window,"idCallingScreenID") == 'MQ010')
	{
		frmToMN060.submit();
	}
	else
	{
		frmToCM010.submit();
	
		<% /* BMIDS903 Set the CallingScreenID context parameter for use by CM100 */ %>
		if (scScreenFunctions.GetContextParameter(window,"idLTVChanged") == '1')
		{
			scScreenFunctions.SetContextParameter(window, "idCallingScreenID", "CM065");
		}		
	}
}

<% /* BMIDS624 End */ %>
<%/* PSC 29/06/2005 MAR5 - Start */%>
<%/*
Add functions to process check boxes in the last column.
These 2 functions alter the style on the cell (oCell)
We need to know which cell is having its style altered (iWhichCell) because
the very top and very bottom cell have slightly different styles which include more borders
*/%>

function SetCheckedEndColumn(oCell,iWhichCell)
{
	if (iWhichCell==1) //First Row
	{
		oCell.className = "msgRightYesTop";
	} else
	{
		if (iWhichCell == m_iTableLength) //Last Row
		{
			oCell.className = "msgRightYesBottom";
		
		} else
		{
			oCell.className = "msgRightYes";		
		}
	}
}
function SetUnCheckedEndColumn(oCell,iWhichCell)
{
	if (iWhichCell==1) //First Row
	{
		oCell.className = "msgRightNoTop";
	} else
	{
		if (iWhichCell == m_iTableLength) //Last Row
		{
			oCell.className = "msgRightNoBottom";
		
		} else
		{
			oCell.className = "msgRightNo";		
		}
	}
}

function GetArrayIndex(iRowSelected,iFirstVisible)
{
	return(iRowSelected + iFirstVisible - 2)
}

<% /* MAR136  Function used when Print Quote check box is double clicked */ %>
function SelectPrintQuote()
{
	// Do not do any processing if the screen is Read Only
	if (m_blnReadOnly == false && iSQPrintButtonAuthorityLevel <= m_AccessType)
	{
		var iIndex = scTable.getRowSelected();  // returns the table index 
		var iFirstVisible = scTable.getFirstVisibleRecord();	

		if (iIndex != -1)
		{
				var iArrayIndex = GetArrayIndex(iIndex,iFirstVisible);
				
				if(m_aCheckBoxArray[iArrayIndex] == "1")
				{
					m_aCheckBoxArray[iArrayIndex] = "";
					SetUnChecked(tblStoredQuotes.rows(iIndex).cells(16), iIndex);
					m_iPrintsRequired = m_iPrintsRequired - 1;
				
					//Disable the Print button if no prints are now selected
					if (m_iPrintsRequired == 0) frmScreen.btnPrint.disabled = true;       

				} 
				else
				{
					m_aCheckBoxArray[iArrayIndex] = "1";
					SetChecked(tblStoredQuotes.rows(iIndex).cells(16), iIndex);
					m_iPrintsRequired = m_iPrintsRequired + 1;
				
					//Enable the Print button
					frmScreen.btnPrint.disabled = false;
				}
				
				
				
		}
		document.selection.empty();
	}	
}

<% /* TW MAR211  Function used when KFI Required check box is double clicked */ %>
function SelectKFIRequired()
{
	// Do not do any processing if the screen is Read Only
	if (m_blnReadOnly == false)
	{
		var iIndex = scTable.getRowSelected();  // returns the table index 
		var iFirstVisible = scTable.getFirstVisibleRecord();	

		if (iIndex != -1)
		{
				var iArrayIndex = GetArrayIndex(iIndex,iFirstVisible);
				
				if(m_aKFIReqdArray[iArrayIndex] == "1")
				{
					m_aKFIReqdArray[iArrayIndex] = "";
					SetUnCheckedEndColumn(tblStoredQuotes.rows(iIndex).cells(17), iIndex);
					updateKFIRequired("0");
				} 
				else
				{
					m_aKFIReqdArray[iArrayIndex] = "1";
					SetCheckedEndColumn(tblStoredQuotes.rows(iIndex).cells(17), iIndex);
					updateKFIRequired("1");
				}
		}
		document.selection.empty();
	}	
}

function GetArrayOfRowsSelected()
{
	var aRetVal = new Array();
	var iRetValIndex = 0;
	var iIndex;

	for(iIndex = 0;iIndex < m_aCheckBoxArray.length;iIndex++)
	{
		if(m_aCheckBoxArray[iIndex] =="1")
		{
			aRetVal[iRetValIndex] = iIndex;
			iRetValIndex++;
		}
	}
	return(aRetVal);
}
<%
//This resets the style on the rows that aren't populated, to their original style, as per HTML definition
%>
function ResetBlankRows(iCount)
{
	for(iIndex = (iCount + 1); iIndex <= m_iTableLength; iIndex++)
	{
		if (iIndex == 1)
		{
			tblStoredQuotes.rows(iIndex).cells(15).className = "TableTopCenter";
			tblStoredQuotes.rows(iIndex).cells(16).className = "TableTopCenter";
			tblStoredQuotes.rows(iIndex).cells(17).className = "TableTopRight";
		} else
		{
			if (iIndex == m_iTableLength)
			{
				tblStoredQuotes.rows(iIndex).cells(15).className = "TableBottomCenter";
				tblStoredQuotes.rows(iIndex).cells(16).className = "TableBottomCenter";
				tblStoredQuotes.rows(iIndex).cells(17).className = "TableBottomRight";
			} else
			{
				tblStoredQuotes.rows(iIndex).cells(15).className = "TableCenter";
				tblStoredQuotes.rows(iIndex).cells(16).className = "TableCenter";
				tblStoredQuotes.rows(iIndex).cells(17).className = "TableRight";
			}
		}
	}
}

function frmScreen.btnPrint.onclick()
{
	var m_XMLAddressSummary;
	<% /* MAR7 GHun */ %>
	//var iNumberOfCopies = 1;
	
	frmScreen.style.cursor = "wait";
	<% /* MAR7 End */ %>
	
	m_XMLAddressSummary = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_XMLAddressSummary.CreateRequestTag(window,null)

	var tagLIST = m_XMLAddressSummary.CreateActiveTag("CUSTOMERADDRESSLIST");

	// Pass over the customers that have a value
	for (var iLoop=1; iLoop<=m_iCurrentMaximumNoOfCustomers; iLoop++)
	{
		var sLoop = iLoop; 
		// get data from context,
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + sLoop,null);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + sLoop,null);

		// now check that there is some data to be sent in the request.....
		if ((sCustomerNumber != "") && (sCustomerVersionNumber != ""))
		{
			m_XMLAddressSummary.ActiveTag = tagLIST;
			m_XMLAddressSummary.CreateActiveTag("CUSTOMERADDRESS");
			m_XMLAddressSummary.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			m_XMLAddressSummary.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
		}
	}

	// Run server-side asp script
	m_XMLAddressSummary.RunASP(document,"GetNumberOfCopiesForKFI.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = m_XMLAddressSummary.CheckResponse(ErrorTypes);

	if ( ErrorReturn[0] == true )
	{
		//Record was found
		m_XMLAddressSummary.SelectTag(null,"RESPONSE");
		iNumberOfCopies = m_XMLAddressSummary.GetTagText("NUMBEROFCOPIES");
	}	

	ListXML.ActiveTag = null;
	var xmlAPPLICATIONFACTFIND = ListXML.SelectTag(null,"APPLICATIONFACTFIND");

	if(xmlAPPLICATIONFACTFIND != null)
	{
		var sTypeOfApplication = ListXML.GetTagText("TYPEOFAPPLICATION").toUpperCase();
	}
		
	var ValidationList = new Array(1);
	ValidationList[0] = "EQ";   // Lifetime
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var m_bLifetime = false;
	<% /*
	if ( m_sSpecialGroup != "" )
	{
		if(XML.IsInComboValidationList(document,"SpecialGroup", m_sSpecialGroup , ValidationList)) 
			m_bLifetime = true;
	}	
	*/ %>
	//BBG573  Check Regulation Indicator
	ValidationList[0] = "R";
	var bRegulated = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if ( m_sRegulationIndicator != "" )
	{
		if(XML.IsInComboValidationList(document,"RegulationIndicator", m_sRegulationIndicator , ValidationList))
			bRegulated = true;
	}

	ListXML.ActiveTag = null;
	var TagListSTOREDQUOTATION = ListXML.CreateTagList("STOREDQUOTATION");
	var sTemplateID = "";
		
	// Loop over all stored quotes, printing those selected.		
	for (var nLoop=0; ListXML.SelectTagListItem(nLoop); nLoop++)
	{		
		//Has this quote been selected for print?
		if(m_aCheckBoxArray[nLoop] == "1")		
		{
			<% /* MAR7 GHun */ %>
			sQuotationNumber = m_aQuotationNumberArray[nLoop];
			sMortgageSubQuoteNumber = m_aSubQuoteArray[nLoop];
			<% /* MAR7 End */ %>
		
			if (bRegulated == true)
			{	
				if (m_bLifetime == false)
				{
					sTemplateID = "STKFIQuotationTemplateID";
				}
				else
				{
					sTemplateID = "LTKFIQuotationTemplateID";
				}
			}
			else
			{
				sTemplateID = "NRKFIQuotationTemplateID";
			}
			// Get the Document ID for the template.
			var GlobalXML = null;
			GlobalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			<% /* MAR7 GHun removed var */ %>
			<% /* PSC 28/11/2006 EP2_218 */ %>
			sDocumentID = GlobalXML.GetGlobalParameterString(document,sTemplateID);
			<% /* MAR7 End */ %>
			
			<% /* PSC 28/11/2006 EP2_218 */ %>
			if (sDocumentID != null)
			{
				//Get Print Attributes for this template
				<% /* MAR7 GHun */ %>
				var bSuccess = GetPrintAttributes();
			
				if (bSuccess)
				<% /* MAR7 End */ %>
				{
					if(AttribsXML.GetTagAttribute("ATTRIBUTES", "INACTIVEINDICATOR") == 1)
					{
						alert("This document template " + AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME") + " is inactive.");
						return false;	<% /* MAR7 GHun */ %>
					}
					else
					{	
						<% /* MAR7 GHun */ %>
						var printerTypeId = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
						m_printerType = getPrinterType(window, printerTypeId);
						m_defaultPrinter = getDefaultPrinterFromTypeId(window, printerTypeId);
			
						if (m_defaultPrinter != null)
						{
							var bSuccess = CallPrintManager();
							if (bSuccess) 
							{
								if (m_printerType == "W")
								{
									var fileContents = PrintXML.GetTagAttribute("DOCUMENTCONTENTS", "FILECONTENTS");
									if (fileContents == "")
									{
										fileContents = PrintXML.GetTagAttribute("PRINTDOCUMENTDETAILS", "PRINTDOCUMENT");
									}

									var documentName = AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME");
									
									<% /* PSC 28/11/2006 EP2_218 */ %>
									var printDocumentData = 
										printDocument(
											fileContents, 
											sDocumentID, 
											documentName,
											m_deliveryType, 
											m_compressionMethod, 
											m_printerType, 
											iNumberOfCopies, 
											true,
											true,
											true, 
											!bViewBeforePrint)																						

									if (printDocumentData != null && printDocumentData.get_success())
									{
										if (printDocumentData.get_fileContents() != null)
										{			
											bSuccess = 
												savePrintedDocument(
													window, m_UserId, m_UnitId, m_MachineId, m_DistributionChannelId, "10", 
													xmlControlDataNode, xmlPrintDataNode, xmlTemplateDataNode, 
													printDocumentData.get_fileSize(), 
													printDocumentData.get_fileContents(), 
													m_compressionMethod);
										}
									}	
									else
									{
										bSuccess = false;
									}
								}
										
								//Update the KFIPrintedIndicator in the Quotation table
								if (bSuccess)
								{
									bSuccess = updateQuotation();
								}
							}																	
							else

							{
								alert("Print action failed.");
							}
						} //EndIf default Printer set up
					}
				}
			}
			else
			{
				alert("The global parameter for " + sTemplateID + " is not defined. See your system administrator.");
			}					
		} // End if Print selected	
	} // End loop over stored quotes
	
	frmScreen.style.cursor = "default"
}
<% /* MAR7 End */ %>


<% /* MAR7 GHun */ %>
function BuildControlDataBlock(PrintXML)
{
	PrintXML.SelectTag(null, "REQUEST");			
	
	xmlControlDataNode = PrintXML.CreateActiveTag("CONTROLDATA");
							
	//Number of KFI copies is calculated above.
	PrintXML.SetAttribute("COPIES", iNumberOfCopies);
	PrintXML.SetAttribute("PRINTER", m_defaultPrinter);
	<% /* PSC 28/11/2006 EP2_218 - Start */ %>
	sDocumentID =  AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATEID");
	PrintXML.SetAttribute("DOCUMENTID", sDocumentID );
	<% /* PSC 28/11/2006 EP2_218 - End */ %>
	PrintXML.SetAttribute("DPSDOCUMENTID", AttribsXML.GetTagAttribute("ATTRIBUTES", "DPSTEMPLATEID"));
								
	m_deliveryType = AttribsXML.GetTagAttribute("ATTRIBUTES", "DELIVERYTYPE");
	m_compressionMethod = ""; //MAR796 GHun FileNet cannot support compression "ZLIB";
								
	PrintXML.SetAttribute("DELIVERYTYPE", m_deliveryType);
	PrintXML.SetAttribute("COMPRESSIONMETHOD", m_compressionMethod);
	PrintXML.SetAttribute("TEMPLATEGROUPID", AttribsXML.GetTagAttribute("ATTRIBUTES", "TEMPLATEGROUPID"));
	PrintXML.SetAttribute("HOSTTEMPLATENAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME"));
	PrintXML.SetAttribute("HOSTTEMPLATEDESCRIPTION", AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATEDESCRIPTION"));
	PrintXML.SetAttribute("DOCUMENTNAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME"));
	
	if (m_printerType == "W")
		PrintXML.SetAttribute("VIEWBEFOREPRINT", "1");
	else
		PrintXML.SetAttribute("VIEWBEFOREPRINT", AttribsXML.GetTagAttribute("ATTRIBUTES", "VIEWBEFOREPRINT"));

	if (AttribsXML.GetTagAttribute("ATTRIBUTES", "VIEWBEFOREPRINT") == "0")
		bViewBeforePrint = false;
	else
		bViewBeforePrint = true;

	if (bViewKFIClick == true)
	{
		bViewBeforePrint = true;
		PrintXML.SetAttribute("VIEWBEFOREPRINT", "1");
	}
	
	PrintXML.SetAttribute("URLPOSTIND", AttribsXML.GetTagAttribute("ATTRIBUTES", "URLPOSTIND"));
	PrintXML.SetAttribute("DESTINATIONTYPE",m_printerType);
}
<% /* MAR7 End */ %>

<% /* MAR7 GHun */ %>
function updateQuotation()
{
	var QuoteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	QuoteXML.CreateRequestTag(window, "UPDATE");
	QuoteXML.CreateActiveTag("QUOTATION");
	QuoteXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	QuoteXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	QuoteXML.CreateTag("QUOTATIONNUMBER", sQuotationNumber);
	QuoteXML.CreateTag("KFIPRINTEDINDICATOR", "1");
<% /* MAR211 TW */ %>
	QuoteXML.CreateTag("KFIREQUIREDINDICATOR", "0");
<% /* MAR211 TW End */ %>

	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			QuoteXML.RunASP(document, "UpdateQuotation.asp");
			break;
		default: // Error
			QuoteXML.SetErrorResponse();
	}																		
													
	if (QuoteXML.IsResponseOK())
	{
		alert("The document has been sent to the printer.");
		return true;
	}
	else
		return false;
}

<% /* MAR211 TW */ %>
function updateKFIRequired(v)
{

	var iIndex = scTable.getRowSelected();  // returns the table index 
	if (iIndex == -1)
	{
		return false;
	}
	var sQuotationNumber = tblStoredQuotes.rows(iIndex).cells(0).innerText;
	var QuoteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	QuoteXML.CreateRequestTag(window, "UPDATE");
	QuoteXML.CreateActiveTag("QUOTATION");
	QuoteXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	QuoteXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	QuoteXML.CreateTag("QUOTATIONNUMBER", sQuotationNumber);
	QuoteXML.CreateTag("KFIREQUIREDINDICATOR", v);


	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			QuoteXML.RunASP(document, "UpdateQuotation.asp");
			break;
		default: // Error
			QuoteXML.SetErrorResponse();
	}																		
													
	if (QuoteXML.IsResponseOK())
	{
		return true;
	}
	else
	{
		return false;
	}
}
<% /* MAR211 TW End */ %>

function CallPrintManager()
{
	PrintXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	PrintXML.CreateRequestTag(window, "PrintDocument");
													
	PrintXML.SelectTag(null, "REQUEST");
						
	PrintXML.SetAttribute("PRINTINDICATOR", "1");

	// CONTROLDATA element								
	BuildControlDataBlock(PrintXML)
							
	PrintXML.SelectTag(null, "REQUEST");	
									
	// PRINTDATA element					
	xmlPrintDataNode = PrintXML.CreateActiveTag("PRINTDATA");
	PrintXML.SetAttribute("METHODNAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "PDMMETHOD"));	
														
	PrintXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	PrintXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
											
	PrintXML.SelectTag(null, "REQUEST");						
									
	//TEMPLATEDATA element
	xmlTemplateDataNode = PrintXML.CreateActiveTag("TEMPLATEDATA");
	PrintXML.SetAttribute("MORTGAGESUBQUOTENUMBER", sMortgageSubQuoteNumber);	
	PrintXML.SetAttribute("QUOTATIONNUMBER", sQuotationNumber);	
									
	PrintXML.SelectTag(null, "REQUEST");
							
	//DOCUMENTDETAILS element
	PrintXML.CreateActiveTag("DOCUMENTDETAILS");
	PrintXML.SetAttribute("DOCUMENTNAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME"));	
	PrintXML.SetAttribute("DOCUMENTGROUP", AttribsXML.GetTagAttribute("ATTRIBUTES", "TEMPLATEGROUPID"));	
									
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			PrintXML.RunASP(document, "PrintManager.asp");
			break;
		default: // Error
			PrintXML.SetErrorResponse();
	}

	 return PrintXML.IsResponseOK();
		
}
function GetPrintAttributes()
{
	AttribsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AttribsXML.CreateRequestTag(window, "GetPrintAttributes");
	AttribsXML.CreateActiveTag("FINDATTRIBUTES");
	<% /* PSC 28/11/2006 EP2_218 */ %>
	AttribsXML.SetAttribute("HOSTTEMPLATEID", sDocumentID);
					
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			AttribsXML.RunASP(document, "PrintManager.asp");
			break;
		default: // Error
			AttribsXML.SetErrorResponse();
	}

	return AttribsXML.IsResponseOK();
	
}

function frmScreen.btnViewKFI.onclick()
{

	<% /* MAR1500 Check if the selected quote has consistent rates */ %>
	
	var sQuotationNumber;
	sQuotationNumber = tblStoredQuotes.rows(scTable.getRowSelected()).cells(0).innerText;
	
	if (CheckForChangedRates(m_sApplicationNumber, m_sApplicationFactFindNumber, sQuotationNumber) == true)
	{
		alert("The Rates have changed since the selected quote was created and the details held may be inconsisent.");
	}
	
	bViewKFIClick = true;
	frmScreen.style.cursor = "wait";
	
	<% /* PJO 08/03/2006 MAR1359 Add progress Message */ %>
	scScreenFunctions.progressOn ("Please wait ... Retrieving KFI", 400);

	var ValidationList = new Array(1);
	
	<% /* Get the Regualtion Indicator */ %>
	ValidationList[0] = "R";
	var bRegulated = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	m_sRegulationIndicator = GetApplicationRegulationIndicator();
	
	if ( m_sRegulationIndicator != "" )
	{
		if(XML.IsInComboValidationList(document,"RegulationIndicator", m_sRegulationIndicator , ValidationList))
			bRegulated = true;
	}

	var sTemplateID = "";
		
	<% /* Pick the TemplateId  */ %>
	if (bRegulated == true)
		sTemplateID = "STKFIQuotationTemplateID";
	else
		sTemplateID = "NRKFIQuotationTemplateID";
	
	<% /* Get the Document ID for the template. */ %>
	var GlobalXML = null;
	GlobalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<% /* PSC 28/11/2006 EP2_218 */ %>
	sDocumentID = GlobalXML.GetGlobalParameterString(document,sTemplateID);

	<% /* PSC 28/11/2006 EP2_218 */ %>
	if (sDocumentID != null)
	{
		<% /* Get Print Attributes for this template*/ %>
		var bSuccess = GetPrintAttributes();
		if (bSuccess)
		{
			if(AttribsXML.GetTagAttribute("ATTRIBUTES", "INACTIVEINDICATOR") == 1)
			{
				scScreenFunctions.progressOff();
				frmScreen.style.cursor = "default";
				alert("This document template " + AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME") + " is inactive.");
				return false;
			}
			else
			{
				var printerTypeId = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
				m_printerType = getPrinterType(window, printerTypeId);
				var bSuccess = CallPrintManager();
				if (bSuccess) 
				{
					var fileContents = PrintXML.GetTagAttribute("DOCUMENTCONTENTS", "FILECONTENTS");
					
					if (fileContents == "")
						fileContents = PrintXML.GetTagAttribute("PRINTDOCUMENTDETAILS", "PRINTDOCUMENT");
					
					var documentName = AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME");
					
					scScreenFunctions.progressOff();									
					<% /* PSC 28/11/2006 EP2_218 */ %>
					var printDocumentData = 
										printDocument(
												fileContents, 
												sDocumentID, 
												documentName,
												m_deliveryType, 
												m_compressionMethod, 
												m_printerType, 
												iNumberOfCopies, 
												true,
												true,
												true, 
												false);
				}
				else
				{
					scScreenFunctions.progressOff();
					frmScreen.style.cursor = "default";
					alert("View action failed.");
				}
			
			}
		}
	}
	else
	{
		scScreenFunctions.progressOff();
		frmScreen.style.cursor = "default";
		alert("The global parameter for " + sTemplateID + " is not defined. See your system administrator.");
	}					
	
	bViewKFIClick = false;
	<% /* PJO 08/03/2006 MAR1359 Add progress Message */ %>
	scScreenFunctions.progressOff();
	frmScreen.style.cursor = "default";
	
}

function GetApplicationRegulationIndicator()
{
	var AppXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AppXML.CreateRequestTag(window, null);
	AppXML.CreateActiveTag("APPLICATION");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			AppXML.RunASP(document,"GetApplicationData.asp");
			break;
		default: // Error
			AppXML.SetErrorResponse();
	}

	if(AppXML.IsResponseOK())
	{
		AppXML.SelectTag(null, "APPLICATIONFACTFIND");
		var sRegulationIndicator = AppXML.GetTagText("REGULATIONINDICATOR");
	}
	
	return sRegulationIndicator;
}

function SetChecked(oCell,iWhichCell)
{
	if (iWhichCell==1) //First Row
	{
		oCell.className = "msgCenterYesTop";
	} else
	{
		if (iWhichCell == m_iTableLength) //Last Row
		{
			oCell.className = "msgCenterYesBottom";
		
		} else
		{
			oCell.className = "msgCenterYes";		
		}
	}
}
function SetUnChecked(oCell,iWhichCell)
{
	if (iWhichCell==1) //First Row
	{
		oCell.className = "msgCenterNoTop";
	} else
	{
		if (iWhichCell == m_iTableLength) //Last Row
		{
			oCell.className = "msgCenterNoBottom";
		
		} else
		{
			oCell.className = "msgCenterNo";		
		}
	}
}

<% /* MAR1500 Check if rates have changed for the given quotation */ %>
function CheckForChangedRates(sApplicationNumber, sApplicationFactFindNumber, sQuotationNumber)
{	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"");
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	
	if (sQuotationNumber != "")
	{
		XML.CreateTag("QUOTATIONNUMBER", sQuotationNumber);
	}

	XML.RunASP(document, "HaveRatesChanged.asp");
	XML.IsResponseOK()
	
	if (XML.IsResponseOK())
	{
		if (XML.GetTagAttribute("QUOTATION", "RATESINCONSISTENT") == "1")
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
}
-->
</script>
<script src="validation.js" language="JScript" type="text/javascript"></script>
</body>
</html>
