<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      DC140.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Bankruptcy History Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		28/09/1999	Created
AD		03/02/2000	Rework
AY		14/02/00	Change to msgButtons button types
IVW		24/03/2000	SYS0438 - Default the selection to the first one in the list
AY		30/03/00	New top menu/scScreenFunctions change
BG		08/05/2000	SYS0680 - Text changed "...history of Bankruptcy?"
BG		17/05/00	SYS0752 Removed Tooltips 
MC		06/06/00	SYS0709 Fixed 'Yes' Radio button label moving to next line
MH      20/06/00    SYS0933 Readonly processing
CL		05/03/01	SYS1920 Read only functionality added
SA		06/06/01	SYS1672 Collect details for guarantors as well
DJP		06/08/01	SYS2564/SYS2787 (child) Client cosmetic customisation which in this case
					is routing.
JLD		4/12/01		SYS2806 use screenfunctions to check completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History

Prog	Date		AQR			Description
GHun	31/07/2002	BMIDS00190	Supoptr multiple customers per bankruptcy history record
GHun	29/08/2002	BMIDS00386  Removed alert message with debugging info
MDC		02/09/2002	BMIDS00336	Credit Check & Bureau Download
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
MDC		03/12/2002	BMIDS01124	Show 'Unassigned' as customer name if unassigned.
JR		23/01/2003	BM0271		Check for ReadOnly on clicking table.
SR		15/06/2004	BMIDS772	Remove the question and cancel button. Change processing accordingly
								Using asp files instead of html for ScrollList functionality
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog	Date		AQR			Description
MF		26/07/2005	MAR019		Changed screen routing to financial lists DC085

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
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>
<span id="spnListScroll">
	<span style="LEFT: 310px; POSITION: absolute; TOP: 245px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<%/* FORMS */%>
<form id="frmToDC141" method="post" action="dc141.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC130" method="post" action="dc130.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC085" method="post" action="DC085.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 220px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span id="spnBankruptcyHistory" style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblBankruptcyHistory" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
			<% /* BMIDS00336 MDC 02/09/2002 - Add Unassigned column */ %>
			<tr id="rowTitles">	<td width="30%" class="TableHead">Name&nbsp</td>		<td width="15%" class="TableHead">Amount<br>Of Debt&nbsp</td>	<td width="15%" class="TableHead">Monthly<br>Repayment&nbsp</td><td width="15%" class="TableHead">Date<br>Declared&nbsp</td><td width="15%" class="TableHead">Date Of<br>Discharge&nbsp</td><td class="TableHead">Unassigned&nbsp</td></tr>
			<tr id="row01">		<td width="30%" class="TableTopLeft">&nbsp</td>			<td width="15%" class="TableTopCenter">&nbsp</td>				<td width="15%" class="TableTopCenter">&nbsp</td>				<td width="15%" class="TableTopCenter">&nbsp</td>			<td width="15%" class="TableTopCenter">&nbsp</td>	<td class="TableTopRight">&nbsp</td></tr>
			<tr id="row02">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
			<tr id="row03">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
			<tr id="row04">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
			<tr id="row05">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
			<tr id="row06">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
			<tr id="row07">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
			<tr id="row08">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
			<tr id="row09">		<td width="30%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>					<td width="15%" class="TableCenter">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
			<tr id="row10">		<td width="30%" class="TableBottomLeft">&nbsp</td>		<td width="15%" class="TableBottomCenter">&nbsp</td>			<td width="15%" class="TableBottomCenter">&nbsp</td>			<td width="15%" class="TableBottomCenter">&nbsp</td>		<td width="15%" class="TableBottomCenter">&nbsp</td><td class="TableBottomRight">&nbsp</td></tr>
		</table>
	</span>

	<span id="spnButtons" style="TOP: 192px; LEFT: 4px; POSITION: ABSOLUTE">
		<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
			<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="TOP: 0px; LEFT: 64px; POSITION: ABSOLUTE">
			<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="TOP: 0px; LEFT: 128px; POSITION: ABSOLUTE">
			<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 320px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="routing/dc140routing.asp" -->	

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc140Attribs.asp" -->
<%/* CODE */%>
<script language="JScript">
<!--
var ListXML;
var m_sMetaAction = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var scScreenFunctions;
var m_sReadOnly="";
var m_blnReadOnly = false;
var m_iTableLength = 10; <% /* SR 14/06/2004 : BMIDS772 */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<% /* SR 15/06/2004 : BMIDS772 - Remove Cancel button */ %>
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Bankruptcy History","DC140",scScreenFunctions);
	
	// IVW - 24/04/2000 - SYS0438 - Default the first item in the list box.
	// IVW - Do before the populate functionality otherwise whatever defaults are set
	// are lost when we return here.

	RetrieveContextData();
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();

	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC140");
	if (m_sReadOnly=="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnAdd.disabled =true;
		frmScreen.btnEdit.disabled=true;
		frmScreen.btnDelete.disabled=true;
	}
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	<% /* MO - 18/11/2002 - BMIDS00376 m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC140"); */ %>
	
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
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
}

function spnBankruptcyHistory.onclick()
{
	if (scTable.getRowSelected() != -1) 
	{
		frmScreen.btnEdit.disabled = false;
		if(m_sReadOnly != "1") //JR BM0271
			frmScreen.btnDelete.disabled = false;
	}
}

function PopulateScreen()
{
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(window,null);

	var TagSEARCH = ListXML.CreateActiveTag("SEARCH");
	ListXML.CreateActiveTag("FINANCIALSUMMARY");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);

	ListXML.ActiveTag = TagSEARCH;
	var TagBANKRUPTCYHISTORYLIST = ListXML.CreateActiveTag("BANKRUPTCYHISTORYLIST");

	// Loop through all customer context entries
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		// If the customer is an applicant, add him/her to the search
		<% /* SYS1672 Collect for Guarantors as well */ %>
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			ListXML.CreateActiveTag("BANKRUPTCYHISTORY");
			ListXML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			ListXML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			ListXML.ActiveTag = TagBANKRUPTCYHISTORYLIST;
		}
	}

	ListXML.RunASP(document,"FindBankruptcyHistorySummary.asp");

	// A record not found error is valid
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);

	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		PopulateTable(0);
	}
}

function PopulateTable(nStart)
{
	var iNumberOfRows ;
	scTable.clear();

	ListXML.ActiveTag = null;
	ListXML.CreateTagList("BANKRUPTCYHISTORY");
	iNumberOfRows = ListXML.ActiveTagList.length
	
	scTable.initialiseTable(tblBankruptcyHistory, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}
function ShowList(nStart)
{	
	var nRow = 0;
	
	<% /* BMIDS00190 */ %>
	var sCustomers;
	var sName;
	var sAmountOfDebt, sMonthlyRepayment, sDateDeclared, sDateOfDischarge, sUnassigned ;
	var CustomersXML;
	<% /* BMIDS00190 End */ %>

	// Populate the table with a set of records, starting with record nStart
	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) != false && nLoop < 10; nLoop++)
	{	
		<% /* BMIDS00190 support display of multiple customers */ %>
		sCustomers = "";
		CustomersXML = ListXML.ActiveTag.cloneNode(true).selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY");
		
		for(var nCust=0; nCust < CustomersXML.length; nCust++)
		{
			sCustomerNumber = CustomersXML.item(nCust).selectSingleNode("CUSTOMERNUMBER").text;
			sName = scScreenFunctions.GetContextCustomerName(window,sCustomerNumber);
			if (sCustomers.length > 0)
				sCustomers = sCustomers + " & " + sName;
			else
				sCustomers = sName;
		}
		<% /* BMIDS00190 End */ %>
		
		sAmountOfDebt = ListXML.GetTagText("AMOUNTOFDEBT");
		sMonthlyRepayment = ListXML.GetTagText("MONTHLYREPAYMENT");
		sDateDeclared = ListXML.GetTagText("DATEDECLARED");
		sDateOfDischarge = ListXML.GetTagText("DATEOFDISCHARGE");

		<% /* BMIDS00336 MDC 30/08/2002 */ %>
		sUnassigned	= ListXML.GetTagText("UNASSIGNED");
		if(sUnassigned == "1")
		{
			sUnassigned = "Yes";
			sCustomers = "Unassigned";	<% /* BMIDS01124 MDC 03/12/2002 */ %>
		}
		else
			sUnassigned = "No";
		<% /* BMIDS00336 MDC 30/08/2002 - End */ %>

		// Display the details in the appropriate table row
		nRow = nLoop + 1;
		scScreenFunctions.SizeTextToField(tblBankruptcyHistory.rows(nRow).cells(0), sCustomers);
		scScreenFunctions.SizeTextToField(tblBankruptcyHistory.rows(nRow).cells(1), sAmountOfDebt);
		scScreenFunctions.SizeTextToField(tblBankruptcyHistory.rows(nRow).cells(2), sMonthlyRepayment);
		scScreenFunctions.SizeTextToField(tblBankruptcyHistory.rows(nRow).cells(3), sDateDeclared);
		scScreenFunctions.SizeTextToField(tblBankruptcyHistory.rows(nRow).cells(4), sDateOfDischarge);
		<% /* BMIDS00336 MDC 30/08/2002 */ %>
		scScreenFunctions.SizeTextToField(tblBankruptcyHistory.rows(nRow).cells(5), sUnassigned);
		<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
		
	}
	
	// IVW - 24/03/2000 - SYS0438 - Default the first item in the list box.
	
	if(nRow > 0)
	{
		scTable.setRowSelected(1);
		frmScreen.btnEdit.disabled = false;
		frmScreen.btnDelete.disabled = false;
	}
	else
	{
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;	
	}		
}

function frmScreen.btnEdit.onclick()
{
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("BANKRUPTCYHISTORY");

	// Get the index of the selected row
	var nRowSelected = scTable.getRowSelectedIndex();

	if(ListXML.SelectTagListItem(nRowSelected-1))
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		scScreenFunctions.SetContextParameter(window,"idXML",ListXML.ActiveTag.xml);
		frmToDC141.submit();
	}
}

function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToDC141.submit();
}

function frmScreen.btnDelete.onclick()
{	
	if (!confirm("Are you sure?")) return;

	ListXML.ActiveTag = null;
	ListXML.CreateTagList("BANKRUPTCYHISTORY");

	// Get the index of the selected row
	var nRowSelected =  scTable.getRowSelectedIndex();
	ListXML.SelectTagListItem(nRowSelected-1);

	// Set up the deletion XML
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequest = XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("DELETE");
	XML.CreateActiveTag("BANKRUPTCYHISTORY");
	
	<% /* BMIDS00190 */ %>
	XML.CreateTag("BANKRUPTCYHISTORYGUID", ListXML.GetTagText("BANKRUPTCYHISTORYGUID"));
	
	var CustomersXML = ListXML.ActiveTag.selectNodes("CUSTOMERVERSIONBANKRUPTCYHISTORY");
	for(var nCust=0; nCust < CustomersXML.length; nCust++)
		XML.ActiveTag.appendChild(CustomersXML.item(nCust).cloneNode(true));
	<%
	//XML.CreateTag("CUSTOMERNUMBER", ListXML.GetTagText("CUSTOMERNUMBER"));
	//XML.CreateTag("CUSTOMERVERSIONNUMBER", ListXML.GetTagText("CUSTOMERVERSIONNUMBER"));
	/* BMIDS00190 End */ %>

	<% /* SR 14/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
			node to CustomerFinancialBO. If the form accessed through completeness check, update FS always.
			Else, do it only when all the records are deleted from the list box
	*/ %>
	if( scScreenFunctions.CompletenessCheckRouting(window) || scTable.getTotalRecords() == 1)
	{
		XML.ActiveTag = xmlRequest ;
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		if(scTable.getTotalRecords() == 1) XML.CreateTag("BANKRUPTCYHISTORYINDICATOR", 0)
		else XML.CreateTag("BANKRUPTCYHISTORYINDICATOR", 1) 
	} 
	<% /* SR 14/06/2004 : BMIDS772 - End */ %>
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"DeleteBankruptcyHistory.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}


	// If the deletion is successful remove the entry from the list xml and the screen
	if(XML.IsResponseOK())
	{
		PopulateScreen();
	}
	XML = null;
}

function btnSubmit.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))	frmToGN300.submit();
	else RouteNext(); <% /*  SYS2564/SYS2787 Soft routing */ %>
}

-->
</script>
</body>
</html>




