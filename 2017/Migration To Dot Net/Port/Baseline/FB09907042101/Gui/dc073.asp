<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc073.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Mortgage Loan Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		02/02/2000	ie5 changes 
AY		14/02/00	Change to msgButtons button types
AY		30/03/00	New top menu/scScreenFunctions change
BG		12/05/2000	SYS0691 Changed column name, Totals default to "0.00", first row in table
					highlighted if there is one, double-clicking on a row opens option in Edit.
BG		17/05/00	SYS0752 Removed Tooltips
MC		08/06/00	SYS0866 Standardise behaviour of Balance & Payment fields
MH      22/06/00    SYS0933 Readonly processing
CL		05/03/01	SYS1920 Read only functionality added
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
PSC		12/08/2002	BMIDS00006	Amend to get customer names from context rather than
								customer number
GHun	05/09/2002	BMIDS00406	Remove the word "Mortgage" from the screen
ASu		05/09/2002	BMIDS00377	Include 'Cancel' button functionality.
GD		18/11/2002	BMIDS00376 Make sure goes to readonly if dataFreezeInd set
MO		20/11/2002	BMIDS00376	Fixed bug noted after the first release of this AQR 
BS		10/06/2003	BM0521	Don't enable Delete button when record selected if screen is read only
HMA     14/10/2004  BMIDS615    Allow row to be selected when Loan Account Number is empty.
PSC		16/05/2006	MAR1798	Run Critical Data 
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
<script src="validation.js" language="Jscript"></script>
<span style="LEFT: 310px; POSITION: absolute; TOP: 310px">
<OBJECT data=scTableListScroll.asp id=scTable 
style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 
type=text/x-scriptlet VIEWASTEXT></OBJECT>
</span>

<% /*ASu 05/09/02 - BMIDS00377 - Cancel button Routing - Start */%>
<form id="frmToDC071" method="post" action="dc071.asp" STYLE="DISPLAY: none"></form>
<% /*ASu BMIDS00377 - End */%>
<form id="frmToDC070" method="post" action="dc070.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC074" method="post" action="dc074.asp" STYLE="DISPLAY: none"></form>

<% //span to keep tabbing within this screen %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" validate ="onchange" mark>
<div id="divBackground" style="HEIGHT: 342px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 9px" class="msgLabel">
		<% /* BMIDS00406 Remove mortgage
		// Mortgage Account<br>Owner
		*/ %>
		Account Owner
		<% /* BMIDS00406 End */ %>
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtApplicant" name="Applicant" maxlength="50" style="POSITION: absolute; WIDTH: 200px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Lender
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtOrganisationName" name="OrganisationName" maxlength="50" style="POSITION: absolute; WIDTH: 200px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 340px; POSITION: absolute; TOP: 36px" class="msgLabel">
		<% /* BMIDS00406 Remove 
		// Mortgage Account<br>Number
		*/ %>
		Account Number
		<% /* BMIDS00406 End */ %>
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountNumber" name="AccountNumber" maxlength="20" style="POSITION: absolute; WIDTH: 150px" class="msgTxt">
		</span>
	</span>
	<span id="spnMortgageLoan" style="LEFT: 4px; POSITION: absolute; TOP: 62px">
		<table id="tblMortgageLoan" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="16%" class="TableHead">Loan Account<br>Number</td><td width="16%" class="TableHead">Original Loan<br>Amount</td><td width="15%" class="TableHead">Outstanding<br>Balance</td><td width="15%" class="TableHead">Monthly<br>Repayment</td><td width="25%" class="TableHead">Purpose Of<br>Loan</td>	<td width="20%" class="TableHead">Redemption<br>Status</td></tr>
			<tr id="row01">		<td width="16%" class="TableTopLeft">&nbsp;</td>			<td width="16%" class="TableTopCenter">&nbsp;</td>			 <td width="15%" class="TableTopCenter">&nbsp;</td>			 <td width="15%" class="TableTopCenter">&nbsp;</td>			<td width="25%" class="TableTopCenter">&nbsp;</td>			<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="16%" class="TableLeft">&nbsp;</td>			<td width="16%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="16%" class="TableLeft">&nbsp;</td>			<td width="16%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="16%" class="TableLeft">&nbsp;</td>			<td width="16%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="16%" class="TableLeft">&nbsp;</td>			<td width="16%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="16%" class="TableLeft">&nbsp;</td>			<td width="16%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="16%" class="TableLeft">&nbsp;</td>			<td width="16%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="16%" class="TableLeft">&nbsp;</td>			<td width="16%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="16%" class="TableLeft">&nbsp;</td>			<td width="16%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				 <td width="15%" class="TableCenter">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="16%" class="TableBottomLeft">&nbsp;</td>		<td width="16%" class="TableBottomCenter">&nbsp;</td>		 <td width="15%" class="TableBottomCenter">&nbsp;</td>		 <td width="15%" class="TableBottomCenter">&nbsp;</td>		<td width="25%" class="TableBottomCenter">&nbsp;</td>		<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>
	<span id="spnButtons" style="LEFT: 4px; POSITION: absolute; TOP: 255px">
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
	<div style="HEIGHT: 52px; LEFT: 250px; POSITION: absolute; TOP: 282px; WIDTH: 348px" class="msgGroupLight">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 7px" class="msgLabel">
			Total of Loans Not Redeemed
			<span style="LEFT: 240px; POSITION: absolute; TOP: -3px">
				<input id="txtLoansNotRedeemed" name="LoansNotRedeemed" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 33px" class="msgLabel">
			Total Monthly Payments on Loans Not Redeemed
			<span style="LEFT: 240px; POSITION: absolute; TOP: -3px">
				<input id="txtPaymentsNotRedeemed" name="PaymentsNotRedeemed" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
			</span>
		</span>
	</div>
</div>
</form><!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% //span to keep tabbing within this screen %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc073Attribs.asp" -->

<script language="JScript">
<!--
var ListXML;
var m_sMetaAction			= null;
var scScreenFunctions;
var m_sReadOnly = null;
var m_blnReadOnly = false;

<% /* PSC 16/05/2006 MAR1798 - Start */ %>
var m_sApplicationNumber		 = null;
var m_sApplicationFactFindNumber = null;
<% /* PSC 16/05/2006 MAR1798 - End */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<% /* Make the required buttons available on the bottom of the screen
		(see msgButtons.asp for details) ASu 05/09/02 BMIDS00377 - Added Cancel Button */ %>
	var sButtonList = new Array("Submit", "Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	<% /* BMIDS00406 Remove Mortgage from the title
	FW030SetTitles("Mortgage Loan Summary","DC073",scScreenFunctions);
	*/ %>
	FW030SetTitles("Loan Summary","DC073",scScreenFunctions);
	<% /* BMIDS00406 End */ %>

	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtApplicant");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtOrganisationName");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtAccountNumber");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtLoansNotRedeemed");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPaymentsNotRedeemed");

	RetrieveContextData();
	PopulateScreen();
			
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();

	//GD BMIDS00376 m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC073");
	if (m_sReadOnly=="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnAdd.disabled=true;
		frmScreen.btnDelete.disabled =true;
		<% /*ASu 05/09/02 BMIDS00377 */ %>
		<% /*GD BMIDS00376 frmScreen.btnCancel.disabled = false; */ %>
		<% /* MO 20/11/2002 BMIDS00376 */ %>
		<% /* DisableMainButton("Cancel"); */ %>
	}
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	//GD BMIDS00376 m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC073");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

<%	/* keep the focus within this screen when using the tab key */%>
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<%	/* keep the focus within this screen when using the tab key */%>
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

<%	/* Retrieves all context data required for use within this screen */%>
function RetrieveContextData()
{
	m_sMetaAction	= scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	<% /* PSC 16/05/2006 MAR1798 - Start */ %>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
	<% /* PSC 16/05/2006 MAR1798 - End */ %>
}

<%	/* Handles the onclick event from the span surrounding 
	the table. This is done here to handle the enabling 
	of buttons when a row is selected. Using the principle 
	of event bubbling we pick up the onclick event after 
	the table_onclick event in the scTable.htm scriptlet */%>
function spnMortgageLoan.onclick()
{
	if (scTable.getRowSelected() != -1) 
	{
		frmScreen.btnEdit.disabled = false;
				
<%
//				if (m_sReadOnly != "1")
//				{
%>
			<% /* BS BM0521 Only enable Delete if screen is in Edit mode */ %>
			if (m_sReadOnly != "1")
			frmScreen.btnDelete.disabled = false;					
<%
//				}
%>
	}
}
function spnMortgageLoan.ondblclick()
{
	if (scTable.getRowSelected() != null) frmScreen.btnEdit.onclick();
}		
<%	/* Retrieves the data and sets the screen accordingly */%>
function PopulateScreen()
{
	var nListLength = 0;
			
	<% /* Get the AccountGUID from the MortgageAccount XML */ %>
	var AccountXML	= new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sXML		= scScreenFunctions.GetContextParameter(window,"idXML");

	AccountXML.LoadXML(sXML);
	AccountXML.SelectTag(null,"MORTGAGEACCOUNT");
			
	<% /* PSC 12/08/2002 BMIDS00006 */ %>
	var sCustomerNames	= AccountXML.GetTagText("CUSTOMERNAMES");
	var sOrganisation	= AccountXML.GetTagText("ORGANISATIONNAME");
	var sAccountNo		= AccountXML.GetTagText("ACCOUNTNUMBER");

	<% /* PSC 12/08/2002 BMIDS00006 */ %>
	frmScreen.txtApplicant.value		= sCustomerNames
	frmScreen.txtOrganisationName.value	= sOrganisation;
	frmScreen.txtAccountNumber.value	= sAccountNo;
			
	var sAccountGUID	= AccountXML.GetTagText("ACCOUNTGUID");
	AccountXML			= null;
			
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	<% /* <REQUEST USERID=? USERTYPE=? UNIT=?>*/ %>
	ListXML.CreateRequestTag(window, "SEARCH");
	ListXML.CreateActiveTag("MORTGAGELOANLIST");
	ListXML.CreateActiveTag("MORTGAGELOAN");
	ListXML.CreateTag("ACCOUNTGUID",sAccountGUID);			
	// 	ListXML.RunASP(document, "FindMortgageLoanList.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			ListXML.RunASP(document, "FindMortgageLoanList.asp");
			break;
		default: // Error
			ListXML.SetErrorResponse();
		}

			
	<%/* A record not found error is valid */ %>
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);
			
	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		PopulateTable(0);
		

		if (ListXML.SelectTag(null,"MORTGAGELOANLIST") != null)
		{
			frmScreen.txtLoansNotRedeemed.value		= ListXML.GetTagText("LOANSNOTREDEEMED");
			frmScreen.txtPaymentsNotRedeemed.value	= ListXML.GetTagText("PAYMENTSNOTREDEEMED");
		}
		else
		{
			frmScreen.txtLoansNotRedeemed.value = "0";
			frmScreen.txtPaymentsNotRedeemed.value	= "0.00";
		}
						
		ListXML.ActiveTag = null;
		ListXML.CreateTagList("MORTGAGELOAN");
		nListLength = ListXML.ActiveTagList.length;
	}
	scTable.initialiseTable(tblMortgageLoan,0,"",PopulateTable,10,nListLength);
	// SYS0691 BG 11/05/2000. Set focus to first item in list if an item exists
	//						   and enable Edit and Delete buttons. 	
	if (ListXML.SelectTag(null,"MORTGAGELOANLIST") != null)
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

<%	/* Displays a set of records in the table. This function is also 
	used by the scListScroll object */%>
function PopulateTable(nStart)
{
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("MORTGAGELOAN");
			
	<% /* Populate the table with a set of records, starting with 
		record nStart */ %>
	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) != false && nLoop < 10; nLoop++)
	{				
		var sLoanAccountNumber	= ListXML.GetTagText("LOANACCOUNTNUMBER");
		
		<% /* BMIDS615 If the LoanAccountNumber is empty, set it to spaces so that this erroneous row can be deleted */ %>
		if (sLoanAccountNumber == "")
		{
			sLoanAccountNumber = "  ";
		}
		
		var sOriginalLoanAmount = ListXML.GetTagText("ORIGINALLOANAMOUNT");
		var sOutstandingBalance	= ListXML.GetTagText("OUTSTANDINGBALANCE");
		var sMonthlyRepayment	= ListXML.GetTagText("MONTHLYREPAYMENT");
		var sPurposeOfLoan		= ListXML.GetTagAttribute("PURPOSEOFLOAN","TEXT");
		var sRedemptionStatus	= ListXML.GetTagAttribute("REDEMPTIONSTATUS","TEXT");

		<% /* Display the details in the appropriate table row */ %>
		var nRow = nLoop + 1;
		scScreenFunctions.SizeTextToField(tblMortgageLoan.rows(nRow).cells(0), sLoanAccountNumber);
		scScreenFunctions.SizeTextToField(tblMortgageLoan.rows(nRow).cells(1), sOriginalLoanAmount);
		scScreenFunctions.SizeTextToField(tblMortgageLoan.rows(nRow).cells(2), sOutstandingBalance);
		scScreenFunctions.SizeTextToField(tblMortgageLoan.rows(nRow).cells(3), sMonthlyRepayment);
		scScreenFunctions.SizeTextToField(tblMortgageLoan.rows(nRow).cells(4), sPurposeOfLoan);
		scScreenFunctions.SizeTextToField(tblMortgageLoan.rows(nRow).cells(5), sRedemptionStatus);
	}
	
	
}

function frmScreen.btnEdit.onclick()
{
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("MORTGAGELOAN");

	<% /* Get the index of the selected row */ %>
	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();

	if(ListXML.SelectTagListItem(nRowSelected-1) == true)
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		scScreenFunctions.SetContextParameter(window,"idXML2",ListXML.ActiveTag.xml);
		frmToDC074.submit();
	}
}

function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToDC074.submit();
}
		
function frmScreen.btnDelete.onclick()
{	
	var bAllowDelete = confirm("Are you sure?");
				
	if (bAllowDelete == true)
	{
		ListXML.ActiveTag = null;
		ListXML.CreateTagList("MORTGAGELOAN");

		<% /* Get the index of the selected row */ %>
		var nRowSelected = scTable.getOffset() + scTable.getRowSelected();

		ListXML.SelectTagListItem(nRowSelected-1);
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				
		XML.CreateRequestTag(window, "DELETE");
		XML.CreateActiveTag("MORTGAGELOAN");
		XML.CreateTag("MORTGAGELOANGUID", ListXML.GetTagText("MORTGAGELOANGUID"));
		XML.CreateTag("ACCOUNTGUID"     , ListXML.GetTagText("ACCOUNTGUID"));
				
		// 		XML.RunASP(document, "DeleteMortgageLoan.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "DeleteMortgageLoan.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

				
		<% /* If the deletion is successful remove the entry from the 
			list xml and the screen */ %>
		if(XML.IsResponseOK())
		{
			<%/* Need to re-hit the database to re-calculate calculated values */%>
			scTable.clear();
			PopulateScreen();
			scScreenFunctions.SetFocusToFirstField(frmScreen);
		}
				
		XML = null;					
	}				
}

function btnSubmit.onclick()
{
	<% /* PSC 16/05/2006 MAR1798 - Start */ %>
	var bContinue = true;
	bContinue = RunIncomeCalculations();
	frmToDC070.submit();
	<% /* PSC 16/05/2006 MAR1798 - End */ %>
}

<%/* ASu 05/09/02 BMIDS00377 - On cancel ignore all changes and close the window. Start */%>
function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
	frmToDC071.submit();
}
<%/* ASu 05/09/02 BMIDS00377 - End */%>
<% /* PSC 16/05/2006 MAR1798 - Start */ %>
function RunIncomeCalculations()
{
  	if (m_sReadOnly =="1")
  		return(true) ;

	var AllowableIncXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AllowableIncXML.CreateRequestTag(window,null);
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
		if (sCustomerNumber.trim().length > 0)
		{
			var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
			var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);
			var sCustomerOrder = scScreenFunctions.GetContextParameter(window,"idCustomerOrder" + nLoop);

			AllowableIncXML.CreateActiveTag("CUSTOMER");
			AllowableIncXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			AllowableIncXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
			AllowableIncXML.CreateTag("CUSTOMERROLETYPE", sCustomerRoleType);
			AllowableIncXML.CreateTag("CUSTOMERORDER", sCustomerOrder);
			AllowableIncXML.ActiveTag = TagCUSTOMERLIST;
		}
	}
	
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

	AllowableIncXML.IsResponseOK()
	return(true);
}
<% /* PSC 16/05/2006 MAR1798 - End */ %>
-->
</script>
</body>
</html>




