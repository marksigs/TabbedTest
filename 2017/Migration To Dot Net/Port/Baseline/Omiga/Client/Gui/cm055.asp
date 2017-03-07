<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      cm055.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   A method of determining the maximum loan the Applicants
				can borrow, given the available information about the two
				highest earners.
				Maximum borrowing can be calculated in 2 different ways
				depending on whether in QuickQuote or Cost Modelling mode.
				Guarantors are currently not taken in preference to the applicants
				BUT are taken as an Applicant.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DPol	17/11/1999	Created.
DPol	17/11/1999	Still in the process of creating it, but as the
					Business Object isn't ready, this is as far as
					I can go at the moment, so checked in for safety.
DPol	23/11/1999	Business Object now created and put in request and response code.
MCS		24/11/1999	Modified and finished 
JLD		03/02/2000	Rework for performance.
APS		04/03/2000	Changed AiP meta-action to Cost Modelling meta-action
AY		17/03/2000	MetaAction changed to ApplicationMode
AY		29/03/2000	New top menu/scScreenFunctions change
AY		29/03/2000	Tabbing was not working correctly - msgButtons in the wrong place & no ID
JLD		26/05/00	SYS0501 - applicant2 was always missed from the maxborrowing calc because
					the customerlist XML was built incorrectly.
CL		05/03/01	SYS1920 Read only functionality added
IK		22/03/01	SYS2145 Add idNoCompleteCheck
CP		09/04/01	SYS1179 Increase Table Size
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMids Update History:

Prog    Date			Description
TW      09 Oct 2002		Modified to incorporate client validation - SYS5115
MDC     04/11/2002		BMIDS00654 BM088 Maximum Borrowing
MV		13/02/2003		BM0097	Amended HTML	
GHun	13/03/2003		BM0457	Include guarantors in allowable income calculation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
MF		15/09/2005	MAR30		IncomeCalcs modified to send ActivityID into calculation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css"><title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>

<% /* BMIDS00654 MDC 04/11/2002 

<!-- List Scroll object -->
<span id=spnMaxBorrowListScroll>
	<span style="LEFT: 302px; POSITION: absolute; TOP: 390px">
<OBJECT data=scTableListScroll.asp id=scMaxBorrowTable 
style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 
type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

*/ %>

<!-- Specify Forms Here -->
<form id="frmToCM010" method="post" action="cm010.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate ="onchange">
<div id="divBackground" style="HEIGHT: 170px; LEFT: 9px; POSITION: absolute; TOP: 60px; WIDTH: 605px" class="msgGroup"><!-- Table of lenders and their maximum borrowing amounts -->

<% /* BMIDS00654 MDC 04/11/2002 
	<span id="spnMaxBorrowTable" style="HEIGHT: 309px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 585px">
		<table id="tblMaxBorrowTable" width="584" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="70%" class="TableHead">Lenders Name</td>	<td width="30%" class="TableHead">Maxium Borrowing Amount</td></tr>
			<tr id="row01">		<td width="70%" class="TableTopLeft">&nbsp;</td>		<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<!-- 
				Start AQR SYS1179: 
				09/04/01 15:01:00 CP: Added additional rows to table
			-->
			<tr id="row11">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row12">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row13">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row14">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row15">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row16">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row17">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row18">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row19">		<td width="70%" class="TableLeft">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row20">		<td width="70%" class="TableBottomLeft">&nbsp;</td>	<td class="TableBottomRight">&nbsp;</td></tr>
			<!--
				End AQR SYS1179: 
			-->
		</table>
	</span>
*/ %>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Maximum Borrowing</strong> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 35px" class="msgLabel"> 
		Maximum Borrowing Amount
		<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
			<input id="txtMaxBorrowing" style="WIDTH: 100px" maxlength="10" class="msgcurrency" readonly tabindex=-1>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel"> 
		Income 1
		<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
			<input id="txtIncome1" style="WIDTH: 100px" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 85px" class="msgLabel"> 
		Income 2
		<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
			<input id="txtIncome2" style="WIDTH: 100px" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 110px" class="msgLabel"> 
		Income Multiplier Type
		<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
			<input id="txtIncMultType" style="WIDTH: 50px" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 135px" class="msgLabel"> 
		Income Multiple
		<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
			<input id="txtIncomeMultiple" style="WIDTH: 100px" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>
	
</div>
</form><!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 12px; POSITION: absolute; TOP: 240px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cm055Attribs.asp" -->

<script language="JScript">
<!--
var m_XMLMaxBorrowing;
var m_sApplicationMode = null;
var m_iTableLength = 20;
var m_sApplicationNumber = null
var m_sApplicationFactFindNumber = null
var m_iCurrentMaximumNoOfCustomers = 6;		<% //6 Current Maximum Number of Customers..... %>
var scScreenFunctions;
var m_blnReadOnly = false;
<% /* BMIDS00654 MDC 04/11/2002 */ %>
var IncomeCalcXML = null;
<% /* BMIDS00654 MDC 04/11/2002 */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Maximum Borrowing","CM055",scScreenFunctions);

	RetrieveContextData();
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	scScreenFunctions.SetScreenToReadOnly(frmScreen);
	btnSubmit.focus();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CM055");
	
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
	m_sApplicationMode = scScreenFunctions.GetContextParameter(window,"idApplicationMode",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B3069");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
}
function PopulateScreen()
{
	<% /* BMIDS00654 MDC 04/11/2002 */ %>
	IncomeCalcXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	if(RunIncomeCalculations())
	{
		frmScreen.txtMaxBorrowing.value = IncomeCalcXML.GetTagText("MAXIMUMBORROWINGAMOUNT");
		frmScreen.txtIncome1.value = IncomeCalcXML.GetTagText("INCOME1");
		frmScreen.txtIncome2.value = IncomeCalcXML.GetTagText("INCOME2");
		frmScreen.txtIncMultType.value = IncomeCalcXML.GetTagText("INCOMEMULTIPLIERTYPE");
		frmScreen.txtIncomeMultiple.value = IncomeCalcXML.GetTagText("INCOMEMULTIPLE");
	}
	<% /* BMIDS00654 MDC 04/11/2002 - End */ %>
}
function PopulateTable()
{
	var iNoOfLenderMaxBorrowAmounts = 0;
	m_XMLMaxBorrowing.ActiveTag = null;
	var TagListMAXBORROWING = m_XMLMaxBorrowing.CreateTagList("MAXBORROWING");
	iNoOfLenderMaxBorrowAmounts = m_XMLMaxBorrowing.ActiveTagList.length;
	if (iNoOfLenderMaxBorrowAmounts > 0)
	{
		scMaxBorrowTable.initialiseTable(tblMaxBorrowTable, 0, "",ShowMaxBorrowing,m_iTableLength,iNoOfLenderMaxBorrowAmounts);
		ShowMaxBorrowing(0);
		scMaxBorrowTable.DisableTable();
	}
}
function ShowRow(nIndex, sLenderName, sMaxBorrowingAmount)
{
	scScreenFunctions.SizeTextToField(tblMaxBorrowTable.rows(nIndex).cells(0),sLenderName);
	scScreenFunctions.SizeTextToField(tblMaxBorrowTable.rows(nIndex).cells(1),sMaxBorrowingAmount);
}
function ShowMaxBorrowing(iStart)
{
	var sLenderName, sMaxBorrowingAmount;
	for (var iLoop = 0; iLoop < m_XMLMaxBorrowing.ActiveTagList.length && iLoop < m_iTableLength;iLoop++)
	{
		m_XMLMaxBorrowing.SelectTagListItem(iLoop + iStart);
		sLenderName	= m_XMLMaxBorrowing.GetTagText("LENDERNAME");
		sMaxBorrowingAmount	= m_XMLMaxBorrowing.GetTagText("MAXBORROWINGAMOUNT");
		ShowRow(iLoop+1, sLenderName, sMaxBorrowingAmount);
	}
}
function btnSubmit.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idNoCompleteCheck","1");
	frmToCM010.submit();
}

<% /* BMIDS00654 MDC 04/11/2002 */ %>
function RunIncomeCalculations()
{
	IncomeCalcXML.CreateRequestTag(window,null);
	<% /* MAR30 */ %> 
	AllowableIncXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));
		
	// Set up request for Income Calculation
	var TagRequest = IncomeCalcXML.CreateActiveTag("INCOMECALCULATION");
	IncomeCalcXML.SetAttribute("CALCULATEMAXBORROWING","1");

	if (m_sApplicationMode == "Quick Quote")
		IncomeCalcXML.SetAttribute("QUICKQUOTEMAXBORROWING","1");
	else
		IncomeCalcXML.SetAttribute("QUICKQUOTEMAXBORROWING","0");

	IncomeCalcXML.CreateActiveTag("APPLICATION");
	IncomeCalcXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	IncomeCalcXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	IncomeCalcXML.ActiveTag = TagRequest;
	
	var TagCUSTOMERLIST = IncomeCalcXML.CreateActiveTag("CUSTOMERLIST");

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
			IncomeCalcXML.CreateActiveTag("CUSTOMER");
			IncomeCalcXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			IncomeCalcXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
			IncomeCalcXML.CreateTag("CUSTOMERROLETYPE", sCustomerRoleType);
			IncomeCalcXML.CreateTag("CUSTOMERORDER", sCustomerOrder);
			IncomeCalcXML.ActiveTag = TagCUSTOMERLIST;
		}
	}
	
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			IncomeCalcXML.RunASP(document,"RunIncomeCalculations.asp");
			break;
		default: // Error
			IncomeCalcXML.SetErrorResponse();
	}
	
	if(IncomeCalcXML.IsResponseOK())
	{
		return(true);
	}

	return(false);
}
<% /* BMIDS00654 MDC 04/11/2002 - End */ %>

-->
</script>
</body>
</html>




