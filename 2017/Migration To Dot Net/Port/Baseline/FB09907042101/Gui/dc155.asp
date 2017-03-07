<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      DC155.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Regular Outgoings
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
IW		??			Created
APS		12/07/2000	Removed FormManager asp reference
DJP		05/09/00	SYS1483 - Clear the listbox before populating, and
					set selection after editing.
CL		05/03/01	SYS1920 Read only functionality added
DJP		06/08/01	SYS2564/SYS2789 (child) Client cosmetic customisation which in this case
					is routing.
JLD		4/12/01		SYS2806 completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
GHun	17/07/2002	BMIDS00190	DCWP3 BM076 Support linking multiple customers to outgoings
AW		24/10/2002	BMIDS00653	BM029 Call to Allowable Income
MDC		01/11/2002	BMIDS00654  ICWP BM088 Income Calculations
GD		18/11/2002	BMIDS00376	Ensure ReadOnly set.
MDC		21/11/2002	BMIDS01034  Return true from RunIncomeCalculation allowing user to progress
GHun	13/03/2003	BM0457		Include guarantors in allowable income calculation
PJO     24/09/2003  BMIDS621    If screen is read only don't run Income Calcs on OK
SR		01/06/2004	BMIDS772	Update FinancialSummary on delete 
								Using asp files instead of html for ScrollList functionality
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog    Date			AQR		Description
MF		15/09/2005		MAR30	IncomeCalcs only called where data has changed
MF		15/09/2005		MAR30	IncomeCalcs modified to send ActivityID into calculation
Maha T	17/11/2005		MAR174	Reomve isChanged() for calling IncomeCalcs.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
	<script src="validation.js" language="JScript"></script>
	
<span id="spnListScroll">
	<span style="LEFT: 310px; POSITION: absolute; TOP: 335px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

	<!-- Specify Forms Here -->
	<form id="frmToDC156" method="post" action="dc156.asp" STYLE="DISPLAY: none"></form>
	<% /* SR 12/06/2004 : Route to DC085 instead of DC160 */ %>
	<form id="frmToOnSubmit" method="post" action="dc085.asp" STYLE="DISPLAY: none"></form>
	<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>

	<% /* Span to keep tabbing within this screen */ %>
	<span id="spnToLastField" tabindex="0"></span>

	<!-- Specify Screen Layout Here -->
	<form id="frmScreen" mark validate="onchange">
		<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 340px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
			<span id="spnRegularOutgoings" style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE">
				<table id="tblRegularOutgoings" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable"  LANGUAGE=javascript>
					<tr id="rowTitles">	<td width="40%" class="TableHead">Type</td>										<td width="40%" class="TableHead">Name</td>				<td class="TableHead">Monthly Amount</td></tr>
					<tr id="row01">		<td width="40%" class="TableTopLeftBold">&nbsp</td>		<td width="40%" class="TableTopCenter">&nbsp</td>		<td class="TableTopRight">&nbsp</td></tr>
					<tr id="row02">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row03">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row04">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row05">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row06">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row07">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row08">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row09">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row10">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row11">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row12">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row13">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row14">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row15">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row16">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row17">		<td width="40%" class="TableLeftBold">&nbsp</td>			<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="row18">		<td width="40%" class="TableBottomLeftBold">&nbsp</td>	<td width="40%" class="TableBottomCenter">&nbsp</td>	<td class="TableBottomRight">&nbsp</td></tr>
				</table>
			</span>

			<% /* BMIDS00190 DCWP3 BM076 */ %>
			<span id="spnButtons" style="TOP: 275px; LEFT: 4px; POSITION: ABSOLUTE">
				<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
					<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
			</span>
			<% /* BMIDS00190 End */ %>

			<span id="spnButtons" style="TOP: 275px; LEFT: 68px; POSITION: ABSOLUTE">
				<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
					<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
			</span>
			
			<% /* BMIDS00190 DCWP3 BM076 */ %>
			<span id="spnButtons" style="TOP: 275px; LEFT: 132px; POSITION: ABSOLUTE">
				<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
					<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
			</span>
			<% /* BMIDS00190 End */ %>

			<div style="TOP: 308px; LEFT: 350px; HEIGHT: 26px; WIDTH: 248px; POSITION: ABSOLUTE" class="msgGroupLight">
				<span style="TOP: 7px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
					Total Regular Outgoings (monthly)
					<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
						<input id="txtTotalRegularOutgoings" name="TotalRegularOutgoings" maxlength="7" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
					</span>
				</span>
			</div>
		</div>
	</form>
<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/DC155attribs.asp" -->
<!-- #include FILE="routing/DC155routing.asp" -->

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>


<!-- Specify Code Here -->
<script language="JScript">
<!--
// JScript Code
var ListXML;
var RecordXML;
		
// Context parameters
var m_sUserType						= null;
var m_sUnitId						= null;
var m_sUserId						= null;
var scScrnFunctions;
var m_sReadOnly						= "";
var m_bDirty						= false;
//var m_bIsSubmit					= false;
var m_blnReadOnly					= false;

<% /* BMIDS00190 */ %>
var ComboXML;
var m_iTableLength = 18; 


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	<% /* SR 12/06/2004 : BMIDS772 - Remove Cancal button */ %>
	var sButtonList = new Array("Submit");
	frmScreen.btnEdit.disabled = true;
	frmScreen.btnDelete.disabled = true;
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Regular Outgoings","DC155",scScreenFunctions);
	
	<% /* BMIDS00190 */ %>	
	scScreenFunctions.SetContextParameter(window, "idMetaAction", "AgreementInPrinciple");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTotalRegularOutgoings");
	RetrieveContextData();
	
	ComboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("RegularOutgoingsType");
	if(!ComboXML.GetComboLists(document,sGroups)) return
	
	PopulateScreen();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC155");
	//GD BMIDS00376
	if (m_blnReadOnly == true)
	{
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnDelete.disabled = true;
		m_sReadOnly = "1" ;    // PJO BMIDS621
	}
	
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
	m_sMetaAction	= scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	m_sReadOnly		= scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sUserId		= scScreenFunctions.GetContextParameter(window, "idUserId"    ,"USER0001");
	m_sUnitId		= scScreenFunctions.GetContextParameter(window, "idUnitId"    ,"UNIT1");
	m_sUserType		= scScreenFunctions.GetContextParameter(window, "idUserType"  ,"BRANCH");
	m_sApplicationNumber			= scScreenFunctions.GetContextParameter(window, "idApplicationNumber"        , null);
	m_sApplicationFactFindNumber	= scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
	m_sApplicationMode = scScreenFunctions.GetContextParameter(window, "idApplicationMode", "Quick Quote");
}

function spnRegularOutgoings.onclick()
{
	if (scTable.getRowSelected() != -1) 
	{
		frmScreen.btnEdit.disabled = false;
		//GD BMIDS00376
		if(m_blnReadOnly == false)
		{
			frmScreen.btnDelete.disabled = false;
		}
	}
}
function GetQuickQuoteOutgoings()
{
	RecordXML = null;
	RecordXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	RecordXML.CreateRequestTag(window,null);

	var TagREGULAROUTGOINGSLIST = RecordXML.CreateActiveTag("REGULAROUTGOINGSLIST");

	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		<%/* If the customer is an applicant, we want them  */%>
		<%/* AQR SYS3312 Or a Guarantor */%>
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			RecordXML.CreateActiveTag("REGULAROUTGOINGS");
			RecordXML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			RecordXML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			RecordXML.ActiveTag = TagREGULAROUTGOINGSLIST;
		}
	}
	RecordXML.RunASP(document,"FindRegularOutgoingsList.asp");

	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = RecordXML.CheckResponse(sErrorArray);
}

function PopulateScreen()
{
	RecordXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	RecordXML.CreateRequestTag(window,null);

	var TagREGULAROUTGOINGSLIST = RecordXML.CreateActiveTag("REGULAROUTGOINGSLIST");

	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		<%/* If the customer is an applicant, we want them  */%>
		<%/* AQR SYS3312 Or a Guarantor */%>
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			RecordXML.CreateActiveTag("REGULAROUTGOINGS");
			RecordXML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			RecordXML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			RecordXML.ActiveTag = TagREGULAROUTGOINGSLIST;
		}
	}
	
	RecordXML.RunASP(document,"FindRegularOutgoingsList.asp");

	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = RecordXML.CheckResponse(sErrorArray);

	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
		if(sResponseArray[1] == sErrorArray[0])
			GetQuickQuoteOutgoings();
	
	ListXML = RecordXML;
	ListXML.CreateTagList("REGULAROUTGOINGS");

	CalculateTotalMonthlyOutgoings();
	PopulateTable(0);
}

function CalculateTotalMonthlyOutgoings()
{
	var lTotalOutgoings = 0;
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("REGULAROUTGOINGS");
	for(var nLoop = 0;ListXML.SelectTagListItem(nLoop);nLoop++)
	{
		var sAmount = ListXML.GetTagText("AMOUNT");
		var sFrequency = ListXML.GetTagText("PAYMENTFREQUENCY");
		lTotalOutgoings += CalculateMonthlyOutgoing(sAmount,sFrequency);
	}
	frmScreen.txtTotalRegularOutgoings.value = lTotalOutgoings;
}
		
function CalculateMonthlyOutgoing(sAmount,sFrequency)
{
	if(sAmount == "") sAmount = "0";
	if(sFrequency == "") sFrequency = "0";
	return Math.ceil((sAmount * sFrequency) / 12);
}


function PopulateTable(nStart)
{
	var iNumberOfRows ;
	iNumberOfRows = ListXML.ActiveTagList.length
	
	scTable.initialiseTable(tblRegularOutgoings, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}
function ShowList(nStart)
{		
	var sValue;
	var sFrequency;
	var sCustomers;
	var sName;
	var sOutgoingsType;
	var sPreviousType = "";
	var sOutgoingsTypeDesc;
	var nRow;
	var CustomersXML;
	var sCustomerNumber;
	
	scTable.clear();

	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) != false && nLoop < m_iTableLength; nLoop++)
	{
		sOutgoingsType = ListXML.GetTagText("REGULAROUTGOINGSTYPE");
		if (sPreviousType != sOutgoingsType)
		{
			ComboXML.ActiveTag = null;
			sOutgoingsTypeDesc = ComboXML.SelectSingleNode("RESPONSE/LIST/LISTNAME/LISTENTRY[VALUEID[.='" + sOutgoingsType + "']]/VALUENAME").text;
			sPreviousType = sOutgoingsType;			
		}
		else
			sOutgoingsTypeDesc = "   ";
		
		sValue = ListXML.GetTagText("AMOUNT");
		sFrequency = ListXML.GetTagText("PAYMENTFREQUENCY");
		
		sCustomers = "";
		CustomersXML = ListXML.ActiveTag.cloneNode(true).selectNodes("CUSTOMERVERSIONREGULAROUTGOINGS");
				
		for(var nCust=0; nCust < CustomersXML.length; nCust++)
		{
			sCustomerNumber = CustomersXML.item(nCust).selectSingleNode("CUSTOMERNUMBER").text;
			sName = scScreenFunctions.GetContextCustomerName(window,sCustomerNumber);
			if (sCustomers.length > 0)
				sCustomers = sCustomers + " & " + sName;
			else
				sCustomers = sName;
		}
	
		nRow = nLoop + 1;	
		scScreenFunctions.SizeTextToField(tblRegularOutgoings.rows(nRow).cells(0), sOutgoingsTypeDesc);
		scScreenFunctions.SizeTextToField(tblRegularOutgoings.rows(nRow).cells(1), sCustomers);
		if(sValue != "")
		{
			sValue = CalculateMonthlyOutgoing(sValue,sFrequency);
			scScreenFunctions.SizeTextToField(tblRegularOutgoings.rows(nRow).cells(2), sValue);
		}
	}
	
	if (nLoop + nStart > 0)
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
	ListXML.CreateTagList("REGULAROUTGOINGS");

	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();
	
	if(ListXML.SelectTagListItem(nRowSelected-1))
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		scScreenFunctions.SetContextParameter(window,"idXML",ListXML.ActiveTag.xml);
		frmToDC156.submit();
	}
}

<% /* BMIDS00190 DCWP3 BM076 */ %>
function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToDC156.submit();
}

function frmScreen.btnDelete.onclick()
{	
	if (!confirm("Are you sure?")) return;

	ListXML.ActiveTag = null;
	ListXML.CreateTagList("REGULAROUTGOINGS");

	// Get the index of the selected row
	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();
	ListXML.SelectTagListItem(nRowSelected-1);

	// Set up the deletion XML
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var TagXmlRequest = XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("DELETE");
	XML.CreateActiveTag("REGULAROUTGOINGS");
	
	XML.CreateTag("REGULAROUTGOINGSGUID", ListXML.GetTagText("REGULAROUTGOINGSGUID"));
	
	var CustomersXML = ListXML.ActiveTag.selectNodes("CUSTOMERVERSIONREGULAROUTGOINGS");
	for(var nCust=0; nCust < CustomersXML.length; nCust++)
		XML.ActiveTag.appendChild(CustomersXML.item(nCust).cloneNode(true));
	
	<% /* SR 09/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
		  node to CustomerFinancialBO. If the form accessed through completeness check, update FS always.
		  Else, do it only when all the records are deleted from the list box
	*/ %>
		if( scScreenFunctions.CompletenessCheckRouting(window) || scTable.getTotalRecords() == 1)
		{
			XML.ActiveTag = TagXmlRequest ;
			XML.CreateActiveTag("FINANCIALSUMMARY");
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
			if(scTable.getTotalRecords() == 1) XML.CreateTag("REGULAROUTGOINGSINDICATOR", 0)
			else XML.CreateTag("REGULAROUTGOINGSINDICATOR", 1) 
		}
		<% /* SR 09/06/2004 : BMIDS772 - End */ %>
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"DeleteRegularOutgoings.asp");
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
<% /* BMIDS00190 End */ %>

function btnSubmit.onclick()
{
	<% /* BMIDS00654 MDC 01/11/2002 
	if (CalculateAllowableIncome())	//AW	24/10/2002	BMIDS00653  */ %>
	<% /* MF MAR30 Only call income calcs if data has changed */ %>
	var bContinue = true;
	<% /* START: MAR174 (Maha T)
				 As this screen is populated with computed values, IsChanged function will not work
	if(IsChanged())
	{
		bContinue = RunIncomeCalculations();
	}
		*/
	%>
	bContinue = RunIncomeCalculations();
	<% /* END: MAR72 */ %>
	
	if (bContinue)
	<% /* BMIDS00654 MDC 01/11/2002 - End */ %>
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else <% /* SR 12/06/2004 : Route to DC085 instead of DC160 */ %>		
			frmToOnSubmit.submit();
	}
}

<% /* BMIDS00190 no longer required
function CommitChanges()
{
	var bSuccess = true;
	if (m_sReadOnly != "1")	
		if (m_bDirty) 
			bSuccess = SaveRegularOutgoings();
	return(bSuccess);
}

function SaveRegularOutgoings()
{
	var bSuccess = true;
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("REGULAROUTGOINGS");
	for(var nLoop = 0;ListXML.SelectTagListItem(nLoop);nLoop++)
	{
		if (ListXML.ActiveTag.childNodes.length == 0)
		{
			ListXML.ActiveTag.parentNode.removeChild(ListXML.ActiveTagList(nLoop))
		}
	}
	DeleteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	DeleteXML.LoadXML(ListXML.XMLDocument.xml);
	DeleteXML.RunASP(document,"DeleteRegularOutgoings.asp");
	ListXML.RunASP(document,"CreateRegularOutgoings.asp");
	bSuccess = ListXML.IsResponseOK();
	
	DeleteXML = null;
	return(bSuccess);
}
*/ %>

<% /* BMIDS00654 MDC 01/11/2002 */ %>
<% /* AW	24/10/2002	BMIDS00653 */ %>
function RunIncomeCalculations()
{
    // PJO BMIDS621 Don't run income calcs in read only
  	if (m_sReadOnly =="1")
  		return(true) ;

	var AllowableIncXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AllowableIncXML.CreateRequestTag(window,null);
	<% /* MAR30 */ %> 
	AllowableIncXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));
		
	// Set up request for Income Calculation
	var TagRequest = AllowableIncXML.CreateActiveTag("INCOMECALCULATION");
	AllowableIncXML.SetAttribute("CALCULATEMAXBORROWING","1");
	AllowableIncXML.SetAttribute("QUICKQUOTEMAXBORROWING","0");
	AllowableIncXML.CreateActiveTag("APPLICATION");
	AllowableIncXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AllowableIncXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	AllowableIncXML.ActiveTag = TagRequest;
	
	var TagCUSTOMERLIST = AllowableIncXML.CreateActiveTag("CUSTOMERLIST");

	for(var nLoop = 1; nLoop <= 2; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		<% /* BM0457 The customer must exist */ %>
		if (sCustomerNumber.trim().length > 0)
		{
		<% /* BM0457 End */ %>
			var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
			var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);
			var sCustomerOrder = scScreenFunctions.GetContextParameter(window,"idCustomerOrder" + nLoop);

			<% /* BM0457
			if(sCustomerRoleType == "1")
			{ */ %>
			AllowableIncXML.CreateActiveTag("CUSTOMER");
			AllowableIncXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			AllowableIncXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
			AllowableIncXML.CreateTag("CUSTOMERROLETYPE", sCustomerRoleType);
			AllowableIncXML.CreateTag("CUSTOMERORDER", sCustomerOrder);
			AllowableIncXML.ActiveTag = TagCUSTOMERLIST;
		}
	}
	
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
					AllowableIncXML.RunASP(document,"RunIncomeCalculations.asp");
			break;
		default: // Error
			AllowableIncXML.SetErrorResponse();
	}

	<% /* BMIDS01034 MDC 21/11/2002 - Always return true to allow user to progress
	if(AllowableIncXML.IsResponseOK())
	{
		return(true);
	}

	return(false); */ %>
	AllowableIncXML.IsResponseOK()
	return(true);
	<% /* BMIDS01034 MDC 21/11/2002 - End */ %>
}
<%	/*BMIDS00653  - End */ %>
<% /* BMIDS00654 MDC 01/11/2002 - End */ %>

-->
</script>
</body>
</html>



<% /* OMIGA BUILD VERSION 045.03.09.26.01 */ %>
