<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /* 
Workfile:      DC130.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Arrears History List
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		28/09/1999	Created
AD		02/02/2000	Rework
AY		14/02/00	Change to msgButtons button types
IVW		24/03/2000	SYS0436 - Default the selection to the first one in the list
AY		30/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
MC		06/06/00	SYS0709 Fixed 'Yes' Radio button label moving to next line
MH      12/06/00    SYS0678 Capitals
MH      20/06/00    SYS0933 Readonly stuff
CL		05/03/01	SYS1920 Read only functionality added
IK		12/04/01	SYS1924 completeness check routing
SA		06/06/01	SYS1672 Collect details for guarantors as well
JLD		4/12/01		SYS2806 completeness check routing
DPF		20/06/02	BMIDS00077 Changes made to file to bring in line with Core V7.0.2, changes made...
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		17/05/02	BMIDS00008 - Modified Routing Screen on btnCancel.OnClick()
PSC		09/08/2002	BMIDS00006 - Amend to route back to DC071 if from there and to 
                                 get arrears list based on account guid passed in
GHun	20/08/2002	BMIDS00190 - DCWP3 BM076 support multiple customers per arrears history
PSC		21/08/2002	BMIDS00349 - Set meta action on submit and cancel
MDC		30/08/2002	BMIDS00336 - Credit Check & Full Bureau Download
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
GHun	22/11/2002	BMIDS01046 - Fix text alignment in the last row of the table
MDC		03/12/2002	BMIDS01124	Show 'Unassigned' as customer name if unassigned.
BS		10/06/2003	BM0521	Don't enable Delete when record selected if screen is in read only
SR		15/06/2004	BMIDS772	Remove the question and cancel button. Change processing accordingly
								Using asp files instead of html for ScrollList functionality
SR		19/06/2004  BMDIS772 - route to DC110 instead of DC140 (on OK) 								
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog	Date		AQR			Description
MF		26/07/2005	MAR019		Changed screen routing to financial lists DC085

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
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
<% /* removed as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
*/ %>
<script src="validation.js" language="JScript"></script>
<span id="spnListScroll">
	<span style="LEFT: 310px; POSITION: absolute; TOP: 245px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<%/* FORMS */%>
<form id="frmToDC131" method="post" action="dc131.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC090" method="post" action="dc090.asp" STYLE="DISPLAY: none"></form>
<% /* SR 19/06/2004 : BMDIS772 - route to DC110 instead of DC140 (on OK) */ %>
<form id="frmToDC110" method="post" action="dc110.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC071" method="post" action="dc071.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC085" method="post" action="DC085.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" validate ="onchange" mark>
<div id="divBackground" style="HEIGHT: 215px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span id="spnArrearsHistory" style="LEFT: 4px; POSITION: absolute; TOP: 10px">
		<table id="tblArrearsHistory" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<% /* BMIDS00336 MDC 30/08/2002 - Add Arrears Type & Unassigned columns, remove Description */ %>
			<tr id="rowTitles">	<td width="25%" class="TableHead">Name&nbsp;</td>		<td width="15%" class="TableHead">Arrears Type&nbsp;</td>		<td width="15%" class="TableHead">Arrears<br>Balance&nbsp;</td>		<td width="15%" class="TableHead">Monthly<br>Repayment&nbsp;</td>		<td width="15%" class="TableHead">Date<br>Cleared&nbsp;</td>	<td width="10%" class="TableHead">Others<br>Resp.&nbsp;</td>	<td class="TableHead">Unassigned</td></tr>
			<tr id="row01">		<td width="25%" class="TableTopLeft">	&nbsp;</td>		<td width="15%" class="TableTopCenter">&nbsp;</td>				<td width="15%" class="TableTopCenter">&nbsp;</td>					<td width="15%" class="TableTopCenter">&nbsp;</td>						<td width="15%" class="TableTopCenter">&nbsp;</td>			<td width="10%" class="TableTopCenter">&nbsp;</td>			<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="25%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>					<td width="15%" class="TableCenter">&nbsp;</td>						<td width="15%" class="TableCenter">&nbsp;</td>							<td width="15%" class="TableCenter">&nbsp;</td>				<td width="10%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="25%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>					<td width="15%" class="TableCenter">&nbsp;</td>						<td width="15%" class="TableCenter">&nbsp;</td>							<td width="15%" class="TableCenter">&nbsp;</td>				<td width="10%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="25%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>					<td width="15%" class="TableCenter">&nbsp;</td>						<td width="15%" class="TableCenter">&nbsp;</td>							<td width="15%" class="TableCenter">&nbsp;</td>				<td width="10%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="25%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>					<td width="15%" class="TableCenter">&nbsp;</td>						<td width="15%" class="TableCenter">&nbsp;</td>							<td width="15%" class="TableCenter">&nbsp;</td>				<td width="10%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="25%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>					<td width="15%" class="TableCenter">&nbsp;</td>						<td width="15%" class="TableCenter">&nbsp;</td>							<td width="15%" class="TableCenter">&nbsp;</td>				<td width="10%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="25%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>					<td width="15%" class="TableCenter">&nbsp;</td>						<td width="15%" class="TableCenter">&nbsp;</td>							<td width="15%" class="TableCenter">&nbsp;</td>				<td width="10%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="25%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>					<td width="15%" class="TableCenter">&nbsp;</td>						<td width="15%" class="TableCenter">&nbsp;</td>							<td width="15%" class="TableCenter">&nbsp;</td>				<td width="10%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="25%" class="TableLeft">&nbsp;</td>			<td width="15%" class="TableCenter">&nbsp;</td>					<td width="15%" class="TableCenter">&nbsp;</td>						<td width="15%" class="TableCenter">&nbsp;</td>							<td width="15%" class="TableCenter">&nbsp;</td>				<td width="10%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<% /* BMIDS01046 Removed align=right */ %>
			<tr id="row10">		<td width="25%" class="TableBottomLeft">&nbsp;</td>		<td width="15%" class="TableBottomCenter">&nbsp;</td>			<td width="15%" class="TableBottomCenter">&nbsp;</td>				<td width="15%" class="TableBottomCenter">&nbsp;</td>					<td width="15%" class="TableBottomCenter">&nbsp;</td>		<td width="10%" class="TableBottomCenter">&nbsp;</td>	<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>

	<span id="spnButtons" style="LEFT: 4px; POSITION: absolute; TOP:185px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">
			<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="LEFT: 128px; POSITION: absolute; TOP: 0px">
			<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 380px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc130Attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--
var ListXML;
var m_sMetaAction = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var scScreenFunctions;
var m_sReadOnly = "";
var m_blnReadOnly = false;
var m_sAccountGuid = "";
var m_iTableLength = 10; <% /* SR 15/06/2004 : BMIDS772 */ %>


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<% /* SR 16/06/2004 : BMIDS772 - Remove Cancel button */ %>
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Arrears History","DC130",scScreenFunctions);

<%
	// IVW - 24/03/2000 - SYS0436 - Default the first item in the list box.
	// IVW - Do before the populate functionality otherwise whatever defaults are set
	// are lost when we return here.
%>	
	RetrieveContextData();
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	
	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC130");

	if (m_sReadOnly=="1")
	{
		frmScreen.btnAdd.disabled =true;
		frmScreen.btnDelete.disabled=true;
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	<% /* MO - 18/11/2002 - BMIDS00376 m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC130"); */ %>
	
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
	m_sAccountGuid = scScreenFunctions.GetContextParameter(window, "idAccountGuid", null);
}

function spnArrearsHistory.onclick()
{
	if (scTable.getRowSelected() != -1) 
	{
		frmScreen.btnEdit.disabled = false;
		<% /* BS BM0521 Only enable delete if screen is in edit mode */%>
		if (m_sReadOnly != "1")
		frmScreen.btnDelete.disabled = false;
	}
}

function PopulateScreen()
{
<%	//replaced next line with line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//ListXML = new scXMLFunctions.XMLObject();
%>	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(window,null)

	var TagSEARCH = ListXML.CreateActiveTag("SEARCH");
	ListXML.CreateActiveTag("FINANCIALSUMMARY");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);

	ListXML.ActiveTag = TagSEARCH;
	var TagARREARSHISTORYLIST = ListXML.CreateActiveTag("ARREARSHISTORYLIST");

	<% /* PSC 09/09/2002 BMIDS00006 - Start*/ %>
	if (m_sAccountGuid != "")
	{
		ListXML.CreateActiveTag("ARREARSHISTORY");
		ListXML.CreateTag("ACCOUNTGUID",m_sAccountGuid);
	
	}
	else
	{
		// Loop through all customer context entries
		for(var nLoop = 1; nLoop <= 5; nLoop++)
		{
			var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
			var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
			var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

			// If the customer is an applicant, add him/her to the search
			//SYS1672 Guarantor details too..
			if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
			{
				ListXML.CreateActiveTag("ARREARSHISTORY");
				ListXML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
				ListXML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
				ListXML.ActiveTag = TagARREARSHISTORYLIST;
			}
		}
	}
	<% /* PSC 09/09/2002 BMIDS00006 - Start*/ %>	
	
	ListXML.RunASP(document,"FindArrearsHistorySummary.asp");

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
	ListXML.CreateTagList("ARREARSHISTORY");
	iNumberOfRows = ListXML.ActiveTagList.length
	
	scTable.initialiseTable(tblArrearsHistory, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}

function ShowList(nStart)
{	
	var nRow = 0;

	<% /* BMIDS00190 */ %>
	var sCustomers;
	var sName;
	var CustomersXML;
	<% /* BMIDS00190 End */ %>
	var sMaximumArrearsBalance, sMonthlyRepayment, sDescriptionOfLoan, sDateCleared
	var sOthersResponsible, sOthersResponsibleField, sUnassigned

	// Populate the table with a set of records, starting with record nStart
	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) != false && nLoop < 10; nLoop++)
	{	
		<% /* BMIDS00190 support display of multiple customers */ %>
		sCustomers = "";
		CustomersXML = ListXML.ActiveTag.cloneNode(true).selectNodes("ACCOUNTRELATIONSHIP");
		
		for(var nCust=0; nCust < CustomersXML.length; nCust++)
		{
			sCustomerNumber = CustomersXML.item(nCust).selectSingleNode("CUSTOMERNUMBER").text;
			sName = scScreenFunctions.GetContextCustomerName(window,sCustomerNumber);
			if (sCustomers.length > 0)
				sCustomers = sCustomers + " & " + sName;
			else
				sCustomers = sName;
		}
		
		//var sCustomerNumber = ListXML.GetTagText("CUSTOMERNUMBER");
		//var sName = scScreenFunctions.GetContextCustomerName(window,sCustomerNumber);
		<% /* BMIDS00190 End */ %>		
		
		sMaximumArrearsBalance = ListXML.GetTagText("MAXIMUMBALANCE");
		sMonthlyRepayment = ListXML.GetTagText("MONTHLYREPAYMENT");
		sDescriptionOfLoan = ListXML.GetTagAttribute("DESCRIPTIONOFLOAN","TEXT");
		sDateCleared = ListXML.GetTagText("DATECLEARED");
		sOthersResponsible = ListXML.GetTagText("ADDITIONALINDICATOR");
		sOthersResponsibleField = (sOthersResponsible == "1") ? "Yes" : "No";
		
		<% /* BMIDS00336 MDC 30/08/2002 */ %>
		sUnassigned	= ListXML.GetTagText("UNASSIGNED");
		if(sUnassigned == "1")
		{
			sUnassigned = "Yes";
			sCustomers = "Unassigned";	<% /* BMIDS01124 MDC 03/12/2002 */ %>
		}
		else
			sUnassigned = "No";
			
		// Display the details in the appropriate table row
		nRow = nLoop + 1;
		scScreenFunctions.SizeTextToField(tblArrearsHistory.rows(nRow).cells(0), sCustomers);
		scScreenFunctions.SizeTextToField(tblArrearsHistory.rows(nRow).cells(1), sDescriptionOfLoan);
		scScreenFunctions.SizeTextToField(tblArrearsHistory.rows(nRow).cells(2), sMaximumArrearsBalance);
		scScreenFunctions.SizeTextToField(tblArrearsHistory.rows(nRow).cells(3), sMonthlyRepayment);
		scScreenFunctions.SizeTextToField(tblArrearsHistory.rows(nRow).cells(4), sDateCleared);
		scScreenFunctions.SizeTextToField(tblArrearsHistory.rows(nRow).cells(5), sOthersResponsibleField);
		scScreenFunctions.SizeTextToField(tblArrearsHistory.rows(nRow).cells(6), sUnassigned);
		<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
	}
	
	// IVW - 24/03/2000 - SYS0436 - Default the first item in the list box.
	
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
	ListXML.CreateTagList("ARREARSHISTORY");

	// Get the index of the selected row
	var nRowSelected = scTable.getRowSelectedIndex();

	if(ListXML.SelectTagListItem(nRowSelected-1))
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		scScreenFunctions.SetContextParameter(window,"idXML",ListXML.ActiveTag.xml);
		frmToDC131.submit();
	}
}

function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToDC131.submit();
}

function frmScreen.btnDelete.onclick()
{	
	if (!confirm("Are you sure?")) return;

	ListXML.ActiveTag = null;
	ListXML.CreateTagList("ARREARSHISTORY");

	// Get the index of the selected row
	var nRowSelected = scTable.getRowSelectedIndex();
	ListXML.SelectTagListItem(nRowSelected-1);

	// Set up the deletion XML
	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequest = XML.CreateRequestTag(window,null)
	XML.CreateActiveTag("DELETE");
	XML.CreateActiveTag("ARREARSHISTORY");
	
	<% /* BMIDS00190 */ %>
	XML.CreateTag("ARREARSHISTORYGUID", ListXML.GetTagText("ARREARSHISTORYGUID"));
	
	var OtherArrearsXML = ListXML.ActiveTag.selectSingleNode("OTHERARREARSACCOUNT")
	if (OtherArrearsXML != null)
	{
		XML.CreateActiveTag("OTHERARREARSACCOUNT")
		XML.CreateTag("ACCOUNTGUID", ListXML.GetTagText("ACCOUNTGUID"));
		
		var CustomersXML = ListXML.ActiveTag.selectNodes("ACCOUNTRELATIONSHIP");
		for(var nCust=0; nCust < CustomersXML.length; nCust++)
			XML.ActiveTag.appendChild(CustomersXML.item(nCust).cloneNode(true));
		<%
		//XML.CreateTag("CUSTOMERNUMBER", ListXML.GetTagText("CUSTOMERNUMBER"));
		//XML.CreateTag("CUSTOMERVERSIONNUMBER", ListXML.GetTagText("CUSTOMERVERSIONNUMBER"));
		//XML.CreateTag("ARREARSSEQUENCENUMBER", ListXML.GetTagText("ARREARSSEQUENCENUMBER")); 
		%>
	}
	<% /* BMIDS00190 End */ %>
	
	<% /* SR 15/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
			node to CustomerFinancialBO. If the form accessed through completeness check, update FS always.
			Else, do it only when all the records are deleted from the list box
	*/ %>
	if( scScreenFunctions.CompletenessCheckRouting(window) || scTable.getTotalRecords() == 1)
	{
		XML.ActiveTag = xmlRequest ;
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		<% /* If routed from DC071, do not pass the new value of the Indicator. The list box shows arrears
			  corresponding to the account in context. Let the component find the value of the indicator.
			  This is done in stored procedure 'usp_SaveFinancialSummary'		
		*/ %>
		if (m_sAccountGuid != "") XML.CreateTag("ARREARSHISTORYINDICATOR", null)
		else <% /* All the arrears related to this case are displayed here  */ %>
		{
			if(scTable.getTotalRecords() == 1) XML.CreateTag("ARREARSHISTORYINDICATOR", 0)
			else XML.CreateTag("ARREARSHISTORYINDICATOR", 1) 
		}
	} 
	<% /* SR 15/06/2004 : BMIDS772 - End */ %>
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"DeleteArrearsHistory.asp");
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
	<% /* PSC 09/08/2002 BMIDS00006 Amend to route back to DC071 if from there */ %> 
	if(scScreenFunctions.CompletenessCheckRouting(window)) frmToGN300.submit();
	else if (m_sAccountGuid != "")
	{
		<% /* PSC 21/08/2002 BMIDS00349 Set meta action */ %>
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		frmToDC071.submit();
	}
	else frmToDC085.submit();<% /* MF 26/07/2005 changed routing frmToDC110.submit();  SR 19/06/2004 : BMIDs772 - route to DC110  */ %>
}

function btnCancel.onclick()
{
	<% /* PSC 09/08/2002 BMIDS00006 Amend to route back to DC071 if from there */ %>
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else if (m_sAccountGuid != "")
	{
		<% /* PSC 21/08/2002 BMIDS00349 Set meta action */ %>
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		frmToDC071.submit();
	}
	else		
		frmToDC090.submit();
}
-->
</script>
</body>
</html>




