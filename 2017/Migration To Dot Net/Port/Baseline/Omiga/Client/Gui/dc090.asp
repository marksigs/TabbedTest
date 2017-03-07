<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc090.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Loans & Liabilities
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		03/02/2000	ie5 changes 
AY		14/02/00	Change to msgButtons button types
MC		21/03/2000	Fixed SYS0432
AY		30/03/00	New top menu/scScreenFunctions change
BG		12/04/00	SYS0364 Changed capitalisation on Loans and Liabilities
BG		17/05/00	SYS0752 Removed Tooltips
MH      20/06/00    SYS0933 ReadOnly
CL		05/03/01	SYS1920 Read only functionality added
SA		31/05/01	SYS1865 On Submit - don't carry out check if read only
SA		06/06/01	SYS1672 Collect details for guarantors as well
DJP		06/08/01	SYS2564/SYS2784 (child) Client cosmetic customisation which in this case
					is routing.
JLD		4/12/01		SYS2806 Completeness check routing
JLD		21/01/02	SYS3844 output unique rows
DPF		20/06/02	BMIDS00077 Changes made to file to bring in line with V7.0.2 of Core, changes are...
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		17/05/2002	BMIDS00008	Modified btnCancel.Onclcik() Routing to MN060.asp
GHun	21/08/2002	BMIDS00190	DCWP3 BM076 support multiple customers per loans liabilities
MDC		30/08/2002	BMIDS00336	CCWP1 BM062 Credit Check & Bureau Download
MDC		08/10/2002	BMIDS00561	Set Name to 'Unassigned' if account is unassigned
AW		24/10/2002	BMIDS00653	BM029 Call to Allowable Income
MDC		01/11/2002	BMIDS00654  ICWP BM088 Income Calculations
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
MDC		21/11/2002	BMIDS01034  Return true from RunIncomeCalculation allowing user to progress
GHun	13/03/2003	BM0457		Include guarantors in allowable income calculation
BS		10/06/2003	BM0521		DOn't enable Delete when record selected if screen is read only
PJO     24/09/2003  BMIDS621    If screen is read only don't run Income Calcs on OK
SR		01/06/2004	BMIDS772	Remove the question and cancel button. Change processing accordingly
								Using asp files instead of html for ScrollList functionality
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog    Date			AQR		Description
MF		15/09/2005		MAR30	IncomeCalcs only called where data has changed
MF		15/09/2005		MAR30	IncomeCalcs modified to send ActivityID into calculation
Maha T	15/11/2005		MAR272	Reomve isChanged() for calling IncomeCalcs.
PSC		16/05/2006		MAR1798	Run Critical Data 
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
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<script src="Validation.js" language=""></script>
<% /* removed as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scMathFunctions.asp" height="1px" id="scMathFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
*/ %>

<span id="spnListScroll">
	<span style="LEFT: 310px; POSITION: absolute; TOP: 245px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<form id="frmToDC091" method="post" action="dc091.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
	
<form id="frmScreen" validate ="onchange" mark>
<div id="divBackground" style="HEIGHT: 280px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span id="spnLoansLiabilities" style="LEFT: 4px; POSITION: absolute; TOP: 10px">
		<table id="tblLoansLiabilities" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles"><td width="20%" class="TableHead">Name</td><td width="15%" class="TableHead">Lender</td><td width="10%" class="TableHead">Account<br>Number</td><td width="10%" class="TableHead">Outstanding<br>Balance</td><td width="10%" class="TableHead">Monthly<br>Repayment</td><td width="14%" class="TableHead">End Date</td><td width="10%" class="TableHead">Others<br>Resp.</td><td width="5%" class="TableHead">To Be<br>Repaid</td><td class="TableHead">Unassigned</td></tr>
			<tr id="row01"><td width="20%" class="TableTopLeft">&nbsp;</td><td width="15%" class="TableTopCenter">&nbsp;</td>	<td width="10%" class="TableTopCenter">&nbsp;</td><td width="10%" class="TableTopCenter">&nbsp;</td><td width="10%" class="TableTopCenter">&nbsp;</td><td width="14%" class="TableTopCenter">&nbsp;</td><td width="10%" class="TableTopCenter">&nbsp;</td><td width="5%" class="TableTopCenter">&nbsp;</td><td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02"><td width="20%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="14%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row03"><td width="20%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="14%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row04"><td width="20%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="14%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row05"><td width="20%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="14%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row06"><td width="20%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="14%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row07"><td width="20%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="14%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row08"><td width="20%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="14%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row09"><td width="20%" class="TableLeft">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="14%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row10"><td width="20%" class="TableBottomLeft">&nbsp;</td><td width="15%" class="TableBottomCenter">&nbsp;</td><td width="10%" class="TableBottomCenter">&nbsp;</td><td width="10%" class="TableBottomCenter">&nbsp;</td><td width="10%" class="TableBottomCenter">&nbsp;</td><td width="14%" class="TableBottomCenter">&nbsp;</td><td width="10%" class="TableBottomCenter">&nbsp;</td><td width="5%" class="TableBottomCenter">&nbsp;</td><td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>

	<span id="spnButtons" style="LEFT: 4px; POSITION: absolute; TOP: 185px">
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
	<div style="HEIGHT: 52px; LEFT: 390px; POSITION: absolute; TOP: 222px; WIDTH: 208px" class="msgGroupLight">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 7px" class="msgLabel">
			Current Total
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtTotalCurrentLiabilities" name="TotalCurrentLiabilities" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 33px" class="msgLabel">
			Future Total
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtTotalFutureLiabilities" name="TotalFutureLiabilities" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
			</span>
		</span>
	</div>
</div>
</form><!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 380px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- #include FILE="routing/dc090routing.asp" -->	

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc090Attribs.asp" -->
<script language="JScript">
<!--
var ListXML;
var m_sMetaAction					= null;
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
var scScreenFunctions;
var m_sReadOnly = "";
var m_blnReadOnly = false;
var XMLArray = new Array();
var m_iTableLength = 10;  

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<% /* Make the required buttons available on the bottom of the screen
		(see msgButtons.asp for details) */ %>
	<% /*  SR 01/06/2004 : BMIDS772 - Remove cancel button */ %>
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	FW030SetTitles("Loans & Liabilities","DC090",scScreenFunctions);

	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTotalCurrentLiabilities");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTotalFutureLiabilities");
			
	RetrieveContextData();
	
	frmScreen.btnEdit.disabled = true;
	frmScreen.btnDelete.disabled = true;

	PopulateScreen();		
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	
	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC090")

	if (m_sReadOnly =="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnAdd.disabled =true;
		frmScreen.btnDelete.disabled =true;
	}
		
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	<% /* MO - 18/11/2002 - BMIDS00376 m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC090") */ %>
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

<% /* keep the focus within this screen when using the tab key */ %>
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<% /* keep the focus within this screen when using the tab key */ %>
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

<% /* Retrieves all context data required for use within this screen */ %>
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber"        , null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
}

function spnLoansLiabilities.onclick()
{
	if (scTable.getRowSelected() != -1) 
	{
		frmScreen.btnEdit.disabled = false;
				
			<% /* BS BM0521 Only enable delete is screen is in edit mode */%>
			if (m_sReadOnly != "1")
			frmScreen.btnDelete.disabled = false;					
	}
}

<% /* Retrieves the data and sets the screen accordingly */ %>
function PopulateScreen()
{
	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//ListXML = new scXMLFunctions.XMLObject();
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	<% /* <REQUEST USERID=? USERTYPE=? UNIT=?> */ %>
	var TagSEARCH = ListXML.CreateRequestTag(window, "SEARCH");
	ListXML.CreateActiveTag("FINANCIALSUMMARY");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
			
	ListXML.ActiveTag = TagSEARCH;
	var TagLOANSLIABILITIESLIST = ListXML.CreateActiveTag("LOANSLIABILITIESLIST");
			
	<% /* Loop through all customer context entries */ %>
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		<% /* If the customer is an applicant, add him/her to the search */ %>
		//SYS1672 Get guarantor details too.
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{					
			ListXML.CreateActiveTag("LOANSLIABILITIES");
			ListXML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			ListXML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			ListXML.ActiveTag = TagLOANSLIABILITIESLIST;
		}
	}
	ListXML.RunASP(document, "FindLiabilitySummary.asp");
			
	<% /* A record not found error is valid */ %>
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);
			
	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		SetUpXMLArray();
		PopulateTable(0);

		if (ListXML.SelectTag(null,"LOANSLIABILITIESLIST") != null)
		{
			frmScreen.txtTotalCurrentLiabilities.value	= ListXML.GetTagText("TOTALCURRENTLIABILITIES");
			frmScreen.txtTotalFutureLiabilities.value	= ListXML.GetTagText("TOTALFUTURELIABILITIES");
		}				
	}
}
function SetUpXMLArray()
{
<% /* Only unique lender information to be displayed - based on sequencenumber - so
	  create an array linking the table row and the taglist index */
%>	ListXML.ActiveTag = null;
	ListXML.CreateTagList("LOANSLIABILITIES");
	var nSeqNum = 999;
	for(var nIdx = 0; ListXML.SelectTagListItem(nIdx); nIdx++)
	{
		if(parseInt(ListXML.GetTagText("SEQUENCENUMBER")) != nSeqNum)
		{
			nSeqNum = parseInt(ListXML.GetTagText("SEQUENCENUMBER"));
			XMLArray[XMLArray.length] = nIdx;
		}
	}
}
<% /* Displays a set of records in the table. This function is also 
	used by the scListScroll object. */ %>
function PopulateTable(nStart)
{
	var iNumberOfRows ;		
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("LOANSLIABILITIES");
	iNumberOfRows = ListXML.ActiveTagList.length
	
	scTable.initialiseTable(tblLoansLiabilities, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}

function ShowList(nStart)
{	
	<% /* BMIDS00190 */ %>
	var sCustomers;
	var sName;
	var CustomersXML;
	<% /* BMIDS00190 End */ %>
	
	scTable.clear();	
	
	<% /* Populate the table with a set of records, starting with 
		record nStart only output unique loansliabilities sequencenumbers */ %>
	for (var nLoop=0; ListXML.SelectTagListItem(XMLArray[nLoop + nStart]) != false && nLoop < 10; nLoop++)
	{
		<% /* BMIDS00561 MDC 08/10/2002 */ %>
		if(ListXML.GetTagText("UNASSIGNED") == "1")
		{
			<% /* Do not display customer name for Unassigned loans */ %>
			sCustomers = "Unassigned";
		}	
		else
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
		}
		<% /* BMIDS00561 MDC 08/10/2002 - End */ %>
				
		<% /*
		var sCustomerNumber = ListXML.GetTagText("CUSTOMERNUMBER");
		var sName = scScreenFunctions.GetContextCustomerName(window,sCustomerNumber);
		*/ %>
		<% /* BMIDS00190 End */ %>

		var sLender	= ListXML.GetTagText("COMPANYNAME");
		var sLoanAccountNumber = ListXML.GetTagText("ACCOUNTNUMBER");
		var sBalance = ListXML.GetTagText("TOTALOUTSTANDINGBALANCE");
		var sRepayment = ListXML.GetTagText("MONTHLYREPAYMENT");
		var sEndDate = ListXML.GetTagText("ENDDATE");
		var sOthersResponsible = ListXML.GetTagText("ADDITIONALINDICATOR");
		var sToBeRepaid	= ListXML.GetTagText("LOANREPAYMENTINDICATOR");
		
		<% /* BMIDS00336 MDC 30/08/2002 */ %>
		var sUnassigned	= ListXML.GetTagText("UNASSIGNED");
		if(sUnassigned == "1")
			sUnassigned = "Yes";
		else
			sUnassigned = "No";
		<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
		

		<% /* Display Yes or No for if there are others responsible or not */ %>
		var sOthersResponsibleField = "";
		if(sOthersResponsible == "1")
		{
			sOthersResponsibleField = "Yes";
		}

		if(sOthersResponsible == "0")
		{
			sOthersResponsibleField = "No";
		}

		<% /* Display Yes or No for if loan is to be repaid or not */ %>
		var sToBeRepaidField = "";
		if(sToBeRepaid == "1")
		{
			sToBeRepaidField = "Yes";
		}

		if(sToBeRepaid == "0")
		{
			sToBeRepaidField = "No";
		}

		<% /* Display the details in the appropriate table row */ %>
		var nRow = nLoop + 1;
		scScreenFunctions.SizeTextToField(tblLoansLiabilities.rows(nRow).cells(0), sCustomers);
		scScreenFunctions.SizeTextToField(tblLoansLiabilities.rows(nRow).cells(1), sLender);
		scScreenFunctions.SizeTextToField(tblLoansLiabilities.rows(nRow).cells(2), sLoanAccountNumber);
		scScreenFunctions.SizeTextToField(tblLoansLiabilities.rows(nRow).cells(3), sBalance);
		scScreenFunctions.SizeTextToField(tblLoansLiabilities.rows(nRow).cells(4), sRepayment);
		scScreenFunctions.SizeTextToField(tblLoansLiabilities.rows(nRow).cells(5), sEndDate);
		scScreenFunctions.SizeTextToField(tblLoansLiabilities.rows(nRow).cells(6), sOthersResponsibleField);
		scScreenFunctions.SizeTextToField(tblLoansLiabilities.rows(nRow).cells(7), sToBeRepaidField);
		<% /* BMIDS00336 MDC 30/08/2002 */ %>
		scScreenFunctions.SizeTextToField(tblLoansLiabilities.rows(nRow).cells(8), sUnassigned);
		<% /* BMIDS00336 MDC 30/08/2002 - End */ %>

	}
	// SYS0432 MDC 21/03/2000. Set focus to first item in list if an item exists
	//						   and enable Edit and Delete buttons. 	
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
	ListXML.CreateTagList("LOANSLIABILITIES");

	<% /* Get the index of the selected row */ %>
	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();

	if(ListXML.SelectTagListItem(XMLArray[nRowSelected-1]) == true)
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		scScreenFunctions.SetContextParameter(window,"idXML",ListXML.ActiveTag.xml);
		frmToDC091.submit();
	}
}

function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToDC091.submit();
}
		
function frmScreen.btnDelete.onclick()
{	
	var bAllowDelete = confirm("Are you sure?");
				
	if (bAllowDelete == true)
	{
		ListXML.ActiveTag = null;
		ListXML.CreateTagList("LOANSLIABILITIES");

		<% /* Get the index of the selected row */ %>
		var nRowSelected = scTable.getRowSelectedIndex();

		ListXML.SelectTagListItem(XMLArray[nRowSelected-1]);
				
		<% /* Get the values needed to alter the totals on the screen */ %>
		var dMonthlyRepayment = ListXML.GetTagText("MONTHLYREPAYMENT");
		var sLoanRepaymentIndicator	= ListXML.GetTagText("LOANREPAYMENTINDICATOR");

		<% /* Set up the deletion XML 
		next line replaced by line below as per V7.0.2 - DPF 20/06/02 - BMIDS00077*/ %>
		//var XML = new scXMLFunctions.XMLObject();
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				
		var xmlRequest = XML.CreateRequestTag(window, "DELETE");
		XML.CreateActiveTag("LOANSLIABILITIES");
		
		<% /* BMIDS00190 */ %>
		XML.CreateTag("ACCOUNTGUID", ListXML.GetTagText("ACCOUNTGUID"));
	
		var CustomersXML = ListXML.ActiveTag.selectNodes("ACCOUNTRELATIONSHIP");
		for(var nCust=0; nCust < CustomersXML.length; nCust++)
			XML.ActiveTag.appendChild(CustomersXML.item(nCust).cloneNode(true));
		<%
		//XML.CreateTag("CUSTOMERNUMBER", ListXML.GetTagText("CUSTOMERNUMBER"));
		//XML.CreateTag("CUSTOMERVERSIONNUMBER", ListXML.GetTagText("CUSTOMERVERSIONNUMBER"));
		//XML.CreateTag("SEQUENCENUMBER", ListXML.GetTagText("SEQUENCENUMBER")); 
		%>
		
		<% /* BMIDS00190 End */ %>
		
		XML.CreateTag("THIRDPARTYGUID", ListXML.GetTagText("THIRDPARTYGUID"));
	
		<% /* SR 09/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
			node to CustomerFinancialBO. If the form accessed through completeness check, update FS always.
			Else, do it only when all the records are deleted from the list box
		*/ %>
		if( scScreenFunctions.CompletenessCheckRouting(window) || scTable.getTotalRecords() == 1)
		{
			XML.ActiveTag = xmlRequest ;
			XML.CreateActiveTag("FINANCIALSUMMARY");
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
			if(scTable.getTotalRecords() == 1) XML.CreateTag("LOANLIABILITYINDICATOR", 0)
			else XML.CreateTag("LOANLIABILITYINDICATOR", 1) 
		} 
		<% /* SR 09/06/2004 : BMIDS772 - End */ %>
				
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "DeleteLiability.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

				
		<% /* If the deletion is successful remove the entry 
			from the list xml and the screen */ %>
		if(XML.IsResponseOK() == true)
		{					
			PopulateScreen();
		}
		XML = null;					
	}				
}

function btnSubmit.onclick()
{
	<% /* BMIDS00654 MDC 01/11/2002 
	if (CalculateAllowableIncome())	//AW	24/10/2002	BMIDS00653  */ %>
	<% /* MF MAR30 Only call income calcs if data has changed */ %>
	var bContinue = true;
	<% /* START: MAR272 (Maha T)
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
		{
		/*	if(UpdateFinancialSummary()) */ frmToGN300.submit();
		}
		else
			<% /* DJP SYS2564/SYS2784 (child) Remove hard coded routing */%>
			RouteNext();
	}
}

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
	
	<% /* PSC 16/05/2006 MAR1798 - Start */ %>
	AllowableIncXML.SelectTag(null,"REQUEST");
	AllowableIncXML.SetAttribute("OPERATION","CriticalDataCheck");	
	AllowableIncXML.CreateActiveTag("CRITICALDATACONTEXT");
	AllowableIncXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	AllowableIncXML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	AllowableIncXML.SetAttribute("SOURCEAPPLICATION","Omiga");
	AllowableIncXML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
	AllowableIncXML.SetAttribute("ACTIVITYINSTANCE","1");
	AllowableIncXML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
	AllowableIncXML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
	AllowableIncXML.SetAttribute("COMPONENT","omIC.IncomeCalcsBO");
	AllowableIncXML.SetAttribute("METHOD","RunIncomeCalculation");	

	window.status = "Critical Data Check - please wait";
	
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
					AllowableIncXML.RunASP(document,"OmigaTMBO.asp");
			break;
		default: // Error
			AllowableIncXML.SetErrorResponse();
	}
	
	window.status = "";	
	<% /* PSC 16/05/2006 MAR1798 - End */ %>

	
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




