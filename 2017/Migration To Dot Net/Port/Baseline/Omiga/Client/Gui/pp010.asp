<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
Workfile:      pp010.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Application Costs
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
IVW		01/01/2001	First version.	
SR		08/02/2001	Place OK & Cancel buttons after the main div 
SR		23/02/2001	SYS1951, SYS1867, SYS1968
CL		05/03/01	SYS1920 Read only functionality added
SR		15/03/01	SYS2063
SR		22/03/01	SYS2153 - increased the hieght of the calling window - Make Rebate/Addition
SR		25/05/01	SYS2298 - new column 'Refund Amount' in table 'FeePayment'
SR		08/06/01	SYS2298 - Layout of screen changed - two List boxes now.
SR		03/08/01	SYS2559	
SA		10/08/01	SYS2328 Default completion indicator to 1 - Not Interfaced
SR		06/09/01	SYS2412
PSC		23/09/01	SYS3164 Don't enable btnMakePayment if Outstanding Balance
							is zero
PSC		27/11/01	SYS3164 Don't enable edit or delete if payment is interfaced
DRC     16/01/02    SYS3528 Ensure that a Noname feetype can still be selected 
DRC     21/01/02    SYS3524 Allow for null payment method
LD		23/05/02	SYS4727 Use cached versions of frame functions
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		05/11/2002	BMIDS00741	- Renamed the Heading name "AMOUNT PAID" to "AMOUNT"
JR		23/01/2003	BM0271 Disable btnAddNewFeeType in ReadOnly mode.
MV		07/02/2003	BM0328  Amended frmSubmit.Onclick()and HTML Code 
BS		15/04/2003	BM0271	Disable Add/Edit/Delete buttons when readonly because case locked
GD		03/06/2003	BM0372	Handle PaymentEvents which have entries with more than 1 validation type the same(RA)
MC		19/04/2004	CC057	PP035 AND PP020 DialogS Height changed.
MC		30/04/2004	BMIDS0468	Cancel and Decline stage AQR
DC      14/05/2004  CC071BMIDS767 Call to PP035 - need environment varibles ( as per BBG234)
SR		06/08/2004	BMIDS813 - modified height of popup window when calling PP035
JD		15/11/2004	BMIDS945	Check GetAcceptedQuoteData for null activeTag before calling GetTagText
HMA     08/12/2004  BMIDS957    Remove VBScript function.
HMA     17/03/2005  BMIDS977    CC089 
HMA     18/04/2005  BMIDS977    CC089 Move "Re-Calculating Quotation" message to status bar.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history:

Prog	Date		Description
GHun	26/07/2005	MAR10	Remove scrollbars when popping up PP035
HMA     09/08/2005	MAR28	WP12. Add Pay Fees button.
HMA     03/10/2005  MAR28	Enable the Edit button for Valuation Refunds and Not To Be Refunded.
HMA     14/11/2005  MAR507  Modified popup window when calling PP020.
PE		13/02/2006  MAR1234	Payment issue on sumbit FMA case.
HMA     30/03/2006  MAR1517 Enable the Edit button for 'Already Refunded'
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific history:

Prog	Date		Description
SAB		05/04/2006	EP323	Enable/Disable the 'Pay Fees' button based on a global variable
AW		03/05/2006	EP323	Amended global parameter
PB		11/05/2006	EP529	MAR1652 Merged - Enable/disable the Pay Fees button based on user authority level.
PB      06/06/2006  EP696	MAR1761 Disable Pay Fees button if no fees are present.
IK		03/08/2006	EP955	MAR1929	Changed CreateListBoxXML to excluded deductions from credit card total
LH		10/08/2006	EP955	MAR1929	Changed CreateListBoxXML to exclude deductions, rebates and ROF from credit card total
PB		17/08/2006	EP1089	MAR1929	Changed CreateListBoxXML to exclude deductions, rebates and ROF from credit card total (again)
GHun	24/08/2006	EP1097	MAR1929 Change spnTable2.onclick to exclude deductions, rebates and ROF when checking if a GatewayPayment has been made
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>

	<% /* Validation script - Controls Soft Coded Field Attributes */ %>
	<script src="validation.js" language="JScript" type="text/javascript"></script>
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" viewastext tabIndex="-1"></object>
<% /* Scriptlets */ %>

<% /* Specify Forms Here 
 MV - 07/02/2003 - BM0328 */ %>
<form id="frmToMN070" method="post" action="mn070.asp" style="DISPLAY: none"></form>
<form id="frmClickDelete" method="post" action="PP010.asp" style="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" style="DISPLAY: none"></form>
<% /* DRC 14/05/2004 CR071 */ %>
<form id="frmToCM060" method="post" action="CM060.asp" style="DISPLAY: none"></form>
<span id="spnListScroll">
	<span style="LEFT: 306px; POSITION: absolute; TOP: 265px">
		<object data="scTableListScroll.asp" id="scTable" name="scScrollTable" style="HEIGHT: 24px; WIDTH: 304px" tabIndex="-1" type="text/x-scriptlet" viewastext></OBJECT>
	</span> 
</span>

<span id="spnListScroll2">
	<span style="LEFT: 306px; POSITION: absolute; TOP: 450px">
		<object data="scTableListScroll.asp" id="scTable2" name="scScrollTable2" style="HEIGHT: 24px; WIDTH: 304px" tabIndex="-1" type="text/x-scriptlet" viewastext></OBJECT>
	</span> 
</span>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 430px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<div id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 10px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="30%" class="TableHead">Fee Description</td>	
				<td width="20%" class="TableHead">Amount Due</td>	
				<td width="30%" class="TableHead">Amount Outstanding</td>
				<td width="25%" class="TableHead">Rebate/Addition</td></tr>
			<tr id="row01">		
				<td width="30%" class="TableTopLeft">&nbsp;</td>		
				<td width="20%" class="TableTopCenter">&nbsp;</td>
				<td width="30%" class="TableTopCenter">&nbsp;</td>		
				<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>	
				<td width="30%" class="TableCenter">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		
				<td width="30%" class="TableBottomLeft">&nbsp;</td>	
				<td width="20%" class="TableBottomCenter">&nbsp;</td>		
				<td width="30%" class="TableBottomCenter">&nbsp;</td>		
				<td class="TableBottomRight">&nbsp;</td></tr> 
		</table>
	</div>
	
	<span style="LEFT:20px; POSITION: absolute; TOP: 177px" class="msgLabel">
		Totals 
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<input id="txtTotalAmountDue" style="WIDTH: 80px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>	
		
		<span style="TOP: -3px; LEFT: 290px; POSITION: ABSOLUTE">
			<input id="txtTotalAmountOutstanding" style="WIDTH: 80px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
		
		<span style="TOP: -3px; LEFT: 470px; POSITION: ABSOLUTE">
			<input id="txtTotalRebateAddition" style="WIDTH: 80px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	
	<span style="LEFT:20px; POSITION: absolute; TOP: 205px">
 		<span style="LEFT:0px; POSITION: absolute; TOP: 0px">
			<input id="btnAddNewFeeType" value="Add New Fee Type" type="button" style="WIDTH: 120px" class="msgButton">
		</span>
	</span>
	
	<div id="spnTable2" style="LEFT: 4px; POSITION: absolute; TOP: 250px">
		<table id="tblTable2" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="15%" class="TableHead">Payment Type</td>	
				<td width="12%" class="TableHead">Date of Payment</td>	
				<td width="15%" class="TableHead">Amount</td>
				<td width="10%" class="TableHead">User Id</td>
				<td width="10%" class="TableHead">Payment Method</td>
				<td width="11%" class="TableHead">Cheque Number</td>
				<td width="15%" class="TableHead">Refund Amount</td>
				<td width="12%" class="TableHead">Date of Refund</td>
			</tr>
			<tr id="row11">		
				<td width="15%" class="TableTopLeft">&nbsp;</td>
				<td width="12%" class="TableTopCenter">&nbsp;</td>
				<td width="15%" class="TableTopCenter">&nbsp;</td>
				<td width="10%" class="TableTopCenter">&nbsp;</td>
				<td width="10%" class="TableTopCenter">&nbsp;</td>
				<td width="11%" class="TableTopCenter">&nbsp;</td>
				<td width="15%" class="TableTopCenter">&nbsp;</td>
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row12">
				<td width="15%" class="TableLeft">&nbsp;</td>
				<td width="12%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>
				<td width="11%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>		
			<tr id="row13">
				<td width="15%" class="TableLeft">&nbsp;</td>
				<td width="12%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>
				<td width="11%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row14">
				<td width="15%" class="TableLeft">&nbsp;</td>
				<td width="12%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>
				<td width="11%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row15">
				<td width="15%" class="TableBottomLeft">&nbsp;</td>
				<td width="12%" class="TableBottomCenter">&nbsp;</td>
				<td width="15%" class="TableBottomCenter">&nbsp;</td>
				<td width="10%" class="TableBottomCenter">&nbsp;</td>
				<td width="10%" class="TableBottomCenter">&nbsp;</td>
				<td width="11%" class="TableBottomCenter">&nbsp;</td>
				<td width="15%" class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomRight">&nbsp;</td>
			</tr>
		</table>
	</div>

	<span style="LEFT:20px; POSITION: absolute; TOP: 360px" class="msgLabel">
		Totals 
		<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
			<input id="txtTotalAmountPaid" style="WIDTH: 45px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>	
		
		<span style="TOP: -3px; LEFT: 425px; POSITION: ABSOLUTE">
			<input id="txtTotalRefundAmount" style="WIDTH: 45px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>	
	
	<span style="LEFT:20px; POSITION: absolute; TOP: 390px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnMakePayment" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		
		<span style="LEFT: 65px; POSITION: absolute; TOP: 0px">
			<input id="btnEditPayment" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="LEFT: 130px; POSITION: absolute; TOP: 0px">
			<input id="btnDeletePayment" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="LEFT: 195px; POSITION: absolute; TOP: 0px">
			<input id="btnPayFees" value="Pay Fees" type="button" style="WIDTH: 60px" class="msgButton">
		</span>		
	</span>		
</div>
</form>


<% /* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 500px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/pp010Attribs.asp" -->

<%/* CODE */ %>
<script language="JScript" type="text/javascript">
<!--
var scScreenFunctions ;
var m_iTableLength = 10 ;
var m_iPaymentsTableLength = 5 ;
var ApplFeeTypeXML = null;
var ApplFeeTypeXMLListBox = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = null;	
var m_ShowListRows = null;
var m_iShowListRowsLength = 0;                          // BMIDS977
var m_ShowListPaymentRows = null ;
var m_sPayMethod = null;
var m_iPaymentRole = 0;
var m_iFeeRole = 0;
var m_iUserRole=0;
var m_blnReadOnly = false;
var sDeductionId = "" ;
var sOtherFeeCode = "" ;
var sReturnOfFundsId = "";
var XMLFeePaymentMethod = null ;
var m_sNotInterfacedValue = ""
var m_sContext = null;
var m_sReadOnly = ""; //BS BM0271 15/04/03
var m_comboXML;
<% /* DRC 14/05/2004 CR071 Start */ %>
var m_sMortgageSubQuoteNumber = null;	
var m_sAcceptedQuoteNumber = null;	
<% /* DRC 14/05/2004 CR071 End */ %>
var m_sMetaAction = null;                          // BMIDS977
var m_dCurrentStateRebateAdditions = new Array();  // BMIDS977
var m_bFromOK = false;                             // BMIDS977
var m_bAwaitingPayment = false;                    // MAR28
var m_dCreditCardTotal;                            // MAR28
var m_sUserID;                                     // MAR28

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{	
	
	m_comboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Application Costs","PP010",scScreenFunctions)

	RetrieveContextData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	
	if(!AcceptedQuote())
	{
		alert("Unable to access Application Costs as there is no accepted quotation for this application.");
		frmToMN070.submit();
	}
	else
	{	
		<% /* Get combovalues for 'FeePaymentMethod' - used while populating screen */ %>
		XMLFeePaymentMethod = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var sGroups = new Array("FeePaymentMethod");
		XMLFeePaymentMethod.GetComboLists(document, sGroups);
		m_sNotInterfacedValue = XMLFeePaymentMethod.GetComboIdForValidation("CompletionIndicator", "N", null, document) ;

		<% /* BMIDS977  Get combovalues for OneOffCost  */ %>
		FeeApplicationValidationTypeXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		sGroups = new Array("OneOffCost");
		FeeApplicationValidationTypeXML.GetComboLists(document, sGroups);
	
		<% /* PB 06/06/06 EP696/MAR1761  Initialise the Pay Fees button to disabled */ %>
		frmScreen.btnPayFees.disabled = true;
		<% /* PB EP696/MAR1761 End */ %>
		
		PopulateScreen();
	
		<% /* SR 15/03/01 : SYS2063 Enable the button 'AddNewFeeType' only on initialisation  */ %>
		SetButtonStateOnInit();
		GetParameterValues();
		m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP010");	
		
		if(m_blnReadOnly)
		{
			frmScreen.btnAddNewFeeType.disabled = true; //JR BM0271
		}
		//BMIDS0468
		if(scScreenFunctions.IsCancelDeclineStage(window))
		{
			frmScreen.btnAddNewFeeType.disabled=true;
			frmScreen.btnDeletePayment.disabled=true;
			frmScreen.btnEditPayment.disabled=true;
			frmScreen.btnMakePayment.disabled=true;
		}
		
		scScreenFunctions.SetFocusToFirstField(frmScreen);	
		
		<% /* BMIDS977 */ %>
		<% /* Remodel when called from CM010 */ %>
		if (m_sMetaAction == "fromAcceptQuote")
		{
			RemodelQuotation();  
			scScreenFunctions.SetContextParameter(window,"idMetaAction",null);
		}	
				
		<% /* Save the rebate/addition values on screen entry */ %>
		SetCurrentStateRebateAddition();			
	} 
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}	

function AcceptedQuote()
{
	var AppXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AppXML.CreateRequestTag(window, null)
	AppXML.CreateActiveTag("APPLICATION");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AppXML.RunASP(document,"GetApplicationData.asp");
	if(AppXML.IsResponseOK())
	{
		AppXML.SelectTag(null, "APPLICATIONFACTFIND");
		if(AppXML.GetTagText("ACCEPTEDQUOTENUMBER")!= "")
		{
			m_sAcceptedQuoteNumber = AppXML.GetTagText("ACCEPTEDQUOTENUMBER");
			return true;
	    }		
	}

	return false;
}

function RetrieveContextData()
{
<%	
/*	// TEST
	//scScreenFunctions.SetContextParameter(window,"idApplicationNumber","C00078387");
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","ADP00108391");
	scScreenFunctions.SetContextParameter(window, "idApplicationFactFindNumber", 1);
	//scScreenFunctions.SetContextParameter(window,"idOtherSystemAccountNumber","ADP00108391");
	scScreenFunctions.SetContextParameter(window,"idCustomerName1","John Smith");	
	scScreenFunctions.SetContextParameter(window,"idRole","70");
	scScreenFunctions.SetContextParameter(window,"idUserId","SR");
	//END TEST
*/	
%>		
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);	
	m_iUserRole = scScreenFunctions.GetContextParameter(window,"idRole","");
	m_sContext = scScreenFunctions.GetContextParameter(window,"idProcessContext",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0"); //BS BM0271 15/04/03
	<% /* MAR28 Add context parameter */ %>
	m_sUserID = scScreenFunctions.GetContextParameter(window,"idUserId",null);
	
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);  // BMIDS977
	<% /* DRC 14/05/2004 CR071 Start */ %>
	m_sMortgageSubQuoteNumber = scScreenFunctions.GetContextParameter(window, "idMortgageSubQuoteNumber", null);	
	  // Is the mortgage SQ Number blank
	  if (m_sMortgageSubQuoteNumber.length == 0)
	  {
	     GetAcceptedQuoteData();
	  }
	<% /* DRC 14/05/2004 CR071 End */ %>
}

<%/* Events */ %>
function frmScreen.btnMakePayment.onclick()
{
	var iCurrentRow = scScrollTable.getRowSelected();
	
	if(iCurrentRow == -1)
	{
		window.alert("Select a fee type for making payment") ;
		return ;
	}
	
	if(m_iUserRole < m_iPaymentRole)
	{
		window.alert("You do not have the authority to add a payment.") ;
		return ;
	}
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sReturn = null;
	
	var ArrayArguments = new Array();
	
	ArrayArguments[0] = m_sApplicationNumber;
	ArrayArguments[1] = tblTable.rows(iCurrentRow).cells(0).innerText; // fee type
	ArrayArguments[2] = tblTable.rows(iCurrentRow).cells(1).innerText; // total amount due
	ArrayArguments[3] = tblTable.rows(iCurrentRow).cells(2).innerText; // total amount o/s
	ArrayArguments[4] = tblTable.rows(iCurrentRow).getAttribute("FEETYPE");
	ArrayArguments[5] = tblTable.rows(iCurrentRow).getAttribute("FEETYPESEQUENCENUMBER");
	ArrayArguments[6] = XML.CreateRequestAttributeArray(window);
	ArrayArguments[7] = "Add";
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "PP020.asp", ArrayArguments, 640, 375);
	
	if (sReturn != null)
	{
		PopulateScreen();
		PopulatePaymentsListBox();
	}
	
	XML = null;
}
<%/*
function spnTable.ondblclick()
{
	var iCurrentRow = scScrollTable.getRowSelected();
	
	if(iCurrentRow != -1)
	{
		if(tblTable.rows(iCurrentRow).getAttribute("LINETYPE")=="PAYHEADING")
			frmScreen.btnMakePayment.onclick();
		else 	
			frmScreen.btnEditPayment.onclick();
	}
}
*/ %>
function spnTable.onclick()
{
	var iCurrentRow = scScrollTable.getRowSelected();
	var sFeeType, sFeeTypeSeqNo ;	
	if(iCurrentRow != -1)
	{
		<% /* SR 22/02/01 - SYS1951 : Disable Make Payment button, if amount outstanding is zero */ %>
		var sAmountOutstanding = tblTable.rows(iCurrentRow).cells(2).innerText ;
		<% /*BS BM0271 15/04/03
		if(parseFloat(sAmountOutstanding) > 0) */ %>
		//BMIDS0468
		if((parseFloat(sAmountOutstanding) > 0) && (m_sReadOnly != "1"))
		{
			frmScreen.btnMakePayment.disabled = false;
		}
		else 
		{
			frmScreen.btnMakePayment.disabled = true;
		}
		
		if((parseFloat(sAmountOutstanding) > 0) && (m_blnReadOnly != true))
		{
			frmScreen.btnMakePayment.disabled = false;
		}
		else 
		{
			frmScreen.btnMakePayment.disabled = true;
		}
		
		if(m_blnReadOnly)
		{
			frmScreen.btnMakePayment.disabled=true;
			frmScreen.btnEditPayment.disabled = true ;
			frmScreen.btnDeletePayment.disabled = true ;
			
		}
		
		//BMIDS0468
		if(scScreenFunctions.IsCancelDeclineStage(window))
		{
			frmScreen.btnAddNewFeeType.disabled=true;
			frmScreen.btnDeletePayment.disabled=true;
			frmScreen.btnEditPayment.disabled=true;
			frmScreen.btnMakePayment.disabled=true;
		}

		
		<% /* Populate the Payments ListBox with related records   */ %>
		sFeeType = tblTable.rows(iCurrentRow).getAttribute("FEETYPE");;
		sFeeTypeSeqNo = tblTable.rows(iCurrentRow).getAttribute("FEETYPESEQUENCENUMBER");;
		PopulatePaymentsListBox(0, sFeeType, sFeeTypeSeqNo) ;
		
		frmScreen.btnEditPayment.disabled = true ;
		frmScreen.btnDeletePayment.disabled = true ;
	}
}

function spnTable2.onclick()
{
	var sPaymentEvent ;
	var sInterfaceStatus;
	var sTransactionReference;
	var sPaymentMethod;
	var bGatewayPaymentMade = false;
	var iCurrentRow = scScrollTable2.getRowSelected();
	
	if(iCurrentRow != -1)
	{
		sPaymentEvent = tblTable2.rows(iCurrentRow).getAttribute("PAYMENTEVENT");
		sInterfaceStatus = tblTable2.rows(iCurrentRow).getAttribute("COMPLETIONINDICATOR");
		
		<% /* MAR28 Check for successful payments already made */ %>
		sPaymentMethod = tblTable2.rows(iCurrentRow).getAttribute("PAYMENTMETHOD");
		sTransactionReference = tblTable2.rows(iCurrentRow).getAttribute("TRANSACTIONREFERENCE");  
		                        
		if (sPaymentMethod != " ")
		{
			<% /* EP1097 GHun Exclude deductions, rebate/additions and return of funds */ %>
			bGatewayPaymentMade = (sPaymentEvent != sDeductionId) && 
					(sPaymentEvent != sReturnOfFundsId) &&
					(!m_comboXML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["O", "RA"])) &&
					((m_comboXML.IsInComboValidationList(document,"FeePaymentMethod", sPaymentMethod, ["CD"])) &&
			                       (sTransactionReference != " "))
		}
		
		<% /*BS BM0271 15/04/03
		if(sInterfaceStatus != m_sNotInterfacedValue || sPaymentEvent == sDeductionId || sPaymentEvent == sReturnOfFundsId) */ %>
		<% //MAR1234 - Peter Edney - 13/02/2006 %>		
		<% //if((sInterfaceStatus != m_sNotInterfacedValue || sPaymentEvent == sDeductionId || sPaymentEvent == sReturnOfFundsId) %>
		if((sPaymentEvent == sDeductionId || sPaymentEvent == sReturnOfFundsId)
		|| (m_sReadOnly == "1"))
		{
			frmScreen.btnEditPayment.disabled = true ;
			frmScreen.btnDeletePayment.disabled = true ;
			
		}
		else if (bGatewayPaymentMade == true)
		{
			<% /* MAR28 Disable the Delete button if a credit card payment has been made 
						Allow the user to view details via the Edit button */ %>
			frmScreen.btnEditPayment.disabled = false ;
			frmScreen.btnDeletePayment.disabled = true ;		
		}
		else
		{
			frmScreen.btnEditPayment.disabled = false ;
			frmScreen.btnDeletePayment.disabled = false ;		
		}

		<% /* MAR75 MAR1517*/ %>
		if ((m_comboXML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["RFV"])) ||
			(m_comboXML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["NTR"])) ||
			(m_comboXML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["AR"])))
		{
			<% /* Enable the Edit button. Disable the Delete button */ %>
			frmScreen.btnEditPayment.disabled = false ;
			frmScreen.btnDeletePayment.disabled = true ;		
		}		


		
		//In all Readonly Entries, Make Edit and Delete Buttons ReadOnly
		if(m_blnReadOnly)
		{
			frmScreen.btnMakePayment.disabled=true;
			frmScreen.btnEditPayment.disabled = true ;
			frmScreen.btnDeletePayment.disabled = true ;
		}
		
		//BMIDS0468
		if(scScreenFunctions.IsCancelDeclineStage(window))
		{
			frmScreen.btnDeletePayment.disabled=true;
			frmScreen.btnEditPayment.disabled=true;
			frmScreen.btnMakePayment.disabled=true;
		}
		
	}
}

function frmScreen.btnEditPayment.onclick()
{
	var iCurrentRow = scScrollTable.getRowSelected();
	var iPaymentRow = scScrollTable2.getRowSelected();
	
	if(iCurrentRow == -1 || iPaymentRow == -1)
	{
		window.alert("Select a payment for editing payment") ;
		return ;
	}

	if(m_iUserRole < m_iPaymentRole)
	{
		window.alert("You do not have the authority to edit a payment.") ;
		return ;
	}
	
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sReturn = null;
	
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sApplicationNumber;
			
	ArrayArguments[1] = tblTable.rows(iCurrentRow).cells(0).innerText;
	ArrayArguments[2] = tblTable.rows(iCurrentRow).cells(1).innerText;
	ArrayArguments[3] = tblTable.rows(iCurrentRow).cells(2).innerText;
	ArrayArguments[4] = tblTable.rows(iCurrentRow).getAttribute("FEETYPE");
	ArrayArguments[5] = tblTable.rows(iCurrentRow).getAttribute("FEETYPESEQUENCENUMBER");
	ArrayArguments[6] = XML.CreateRequestAttributeArray(window);
	ArrayArguments[7] = "Edit";
	
	<% /* extra fields required for edit */ %>
	ArrayArguments[8]  = tblTable2.rows(iPaymentRow).getAttribute("FEEPAYMENTSEQNUMBER");	
	ArrayArguments[9]  = tblTable2.rows(iPaymentRow).cells(2).innerText;
	ArrayArguments[10] = tblTable2.rows(iPaymentRow).cells(1).innerText;
	ArrayArguments[11] = tblTable2.rows(iPaymentRow).getAttribute("CHEQUENUMBER");
	ArrayArguments[12] = tblTable2.rows(iPaymentRow).getAttribute("PAYMENTMETHOD");
	ArrayArguments[13] = tblTable2.rows(iPaymentRow).getAttribute("PAYMENTEVENT");
	ArrayArguments[14] = tblTable2.rows(iPaymentRow).cells(0).innerText;
	ArrayArguments[15] = tblTable2.rows(iPaymentRow).cells(7).innerText;
	ArrayArguments[16] = tblTable2.rows(iPaymentRow).getAttribute("NOTES");
	ArrayArguments[17] = tblTable2.rows(iPaymentRow).cells(6).innerText;
	ArrayArguments[18] = tblTable2.rows(iPaymentRow).getAttribute("SAVINGSACCOUNTNUMBER");
	ArrayArguments[19] = tblTable2.rows(iPaymentRow).getAttribute("TRANSACTIONREFERENCE");
	ArrayArguments[20] = tblTable2.rows(iPaymentRow).getAttribute("CARDFAILRESULTCODE");


	sReturn = scScreenFunctions.DisplayPopup(window, document, "PP020.asp", ArrayArguments, 640, 375);
	
	<% /* Refresh the screen */ %>
	if (sReturn != null)
	{
		PopulateScreen();
		PopulatePaymentsListBox();	
	}
	
	XML = null;
}

function frmScreen.btnDeletePayment.onclick()
{
	var sFile;
	var ErrorTypes = new Array("RECORDNOTFOUND");
	
	var iCurrentRow = scScrollTable.getRowSelected();
	var iPaymentRow = scScrollTable2.getRowSelected();
	
	if(iCurrentRow == -1 || iPaymentRow == -1)
	{
		window.alert("Select a payment for deleting payment") ;
		return ;
	}
	
	if(m_iUserRole < m_iPaymentRole)
	{
		window.alert("You do not have the authority to delete a payment.") ;
		return ;
	}

	ApplFeeTypeXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	
	ApplFeeTypeXML.ResetXMLDocument();
	var XMLActiveTag = ApplFeeTypeXML.CreateRequestTag(window , "DELETEFEETYPEPAYMENT");
	ApplFeeTypeXML.CreateActiveTag("FEEPAYMENT");
	ApplFeeTypeXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	ApplFeeTypeXML.SetAttribute("FEETYPE",tblTable.rows(iCurrentRow).getAttribute("FEETYPE"));
	ApplFeeTypeXML.SetAttribute("FEETYPESEQUENCENUMBER", tblTable.rows(iCurrentRow).getAttribute("FEETYPESEQUENCENUMBER"));
	ApplFeeTypeXML.SetAttribute("PAYMENTSEQUENCENUMBER",tblTable2.rows(iPaymentRow).getAttribute("FEEPAYMENTSEQNUMBER"));
	//SYS2328 default completion indicator
	ApplFeeTypeXML.SetAttribute("COMPLETIONINDICATOR", "1")

	ApplFeeTypeXML.ActiveTag = XMLActiveTag;
	ApplFeeTypeXML.CreateActiveTag("PAYMENTRECORD");
	ApplFeeTypeXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	ApplFeeTypeXML.SetAttribute("PAYMENTSEQUENCENUMBER",tblTable2.rows(iPaymentRow).getAttribute("FEEPAYMENTSEQNUMBER"));
	
	// 	ApplFeeTypeXML.RunASP(document,"PaymentProcessingRequest.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			ApplFeeTypeXML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			ApplFeeTypeXML.SetErrorResponse();
		}

	
	var ErrorReturn = ApplFeeTypeXML.CheckResponse(ErrorTypes);
		
	if(ErrorReturn[0] == false) return ;
	else
	{
		PopulateScreen();
		PopulatePaymentsListBox();
		<% /* frmClickDelete.submit();	*/ %>
	}
}

function frmScreen.btnAddNewFeeType.onclick()
{
	if(m_iUserRole < m_iFeeRole)
	{
		window.alert("You do not have the authority to add a new fee type.") ;
		return ;
	}

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sReturn = null;
	var ArrayArguments = new Array();
	
	ArrayArguments[0] = m_sApplicationNumber;
	ArrayArguments[1] = XML.CreateRequestAttributeArray(window);
	<% /* DRC  14/05/2004 CC71 Start */ %>
	ArrayArguments[2] = m_sApplicationFactFindNumber;
	ArrayArguments[3] = m_sMortgageSubQuoteNumber;
	ArrayArguments[4] = m_sAcceptedQuoteNumber;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "PP035.asp", ArrayArguments, 374, 260);
	
	// Refresh the screen
	if (sReturn != null) 
	{
		if (sReturn[0] == "CM060")
		{
		   scScreenFunctions.SetContextParameter(window,"idMetaAction","fromPP010");
		   scScreenFunctions.SetContextParameter(window,"idXML",m_sAcceptedQuoteNumber);
 		   scScreenFunctions.SetContextParameter(window,"idXML2",m_sMortgageSubQuoteNumber);
		  
			frmToCM060.submit();
		}
		else
			PopulateScreen();
	}
	<% /* DRC 14/05/2004 CC71 End */ %>
	XML = null;
}

function GetParameterValues()
{
	// This is a Function to retrieve the Global Parameter Values from database 
	//save the combo settings in idXML context to use when we return
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var blnReturn = false;	
	//Preparing XML Request string 

	var XMLActiveTag = XML.CreateRequestTag(window,"SEARCH");

	XML.CreateActiveTag("GLOBALPARAMETER");
	XML.CreateTag("NAME","PPROCMAINTAINFEEPAYMENTROLE");
	XML.ActiveTag = XMLActiveTag;
	
	XML.CreateActiveTag("GLOBALPARAMETER");
	XML.CreateTag("NAME","PPROCMAINTAINFEETYPEROLE");
	XML.ActiveTag = XMLActiveTag;
	
	XML.RunASP(document,"FindCurrentParameterList.asp");

	if (XML.IsResponseOK()== true)
	{	
		blnReturn = true;
		if(XML.SelectTag(null,"GLOBALPARAMETERLIST") != null)
			{
				XML.CreateTagList("GLOBALPARAMETER");
				iNumberOfParameters = XML.ActiveTagList.length;
				for (var i0 = 0;  i0 < iNumberOfParameters ; i0++)
				{
					if (XML.SelectTagListItem(i0) == true ) 
					{
						switch(XML.GetTagText("NAME").toUpperCase())
						{
							case "PPROCMAINTAINFEEPAYMENTROLE" :
								m_iPaymentRole = XML.GetTagText("AMOUNT");
								break;
							case "PPROCMAINTAINFEETYPEROLE":
								m_iFeeRole = XML.GetTagText("AMOUNT");
								break;
						}
					}
				}
			}		
	}
	else 
		blnReturn = false;

	XML = null; 
	return blnReturn;
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

<%/* Functions */ %>
function PopulatePaymentsListBox(nStart, sFeeType, sFeeTypeSeqNo)
{	
	if(nStart==null) nStart=0 ;
	if(sFeeType==null || sFeeTypeSeqNo==null)
	{
		var iFeeListRow	= scScrollTable.getRowSelected();;
		sFeeType		= tblTable.rows(iFeeListRow).getAttribute("FEETYPE");
		sFeeTypeSeqNo	= tblTable.rows(iFeeListRow).getAttribute("FEETYPESEQUENCENUMBER");
	}
	
	sPattern = "//LISTLINE[@LINETYPE='PAYMENT' $and$ @FEETYPE=" + sFeeType + " $and$ @FEETYPESEQUENCENUMBER=" + sFeeTypeSeqNo + "]";
	m_ShowListPaymentRows = ApplFeeTypeXMLListBox.selectNodes(sPattern);
	
	var iNumberOfRows = m_ShowListPaymentRows.length;
	scScrollTable2.clear(); 
	scScrollTable2.initialiseTable(tblTable2, 0, "", ShowPaymentsList, m_iPaymentsTableLength, iNumberOfRows);
	ShowPaymentsList(nStart);
	
	DisplayPaymentTotals() ;
}

function PopulateFeeTypesListBox(nStart)
{
	m_ShowListRows = ApplFeeTypeXMLListBox.selectNodes("//LISTLINE[@LINETYPE='PAYHEADING']");
	m_iShowListRowsLength = m_ShowListRows.length;    // BMIDS977

	scScrollTable.initialiseTable(tblTable, 0, "", ShowFeeTypesList, m_iTableLength, m_iShowListRowsLength);
	ShowFeeTypesList(nStart);
	
	DisplayFeeTotals();	
}

function DisplayFeeTotals()
{
	var iCount, ListRow ;
	var dTotalAmountDue, dTotalAmountOS, dTotalRebateOrAddition ;
	var sAmountDue, sAmountOS, sRebateOrAddition ;
	
	dTotalAmountDue = 0 ;
	dTotalAmountOS = 0 ;
	dTotalRebateOrAddition = 0 ;
	
	for (iCount = 0; iCount < m_iShowListRowsLength ; iCount++)
	{
		ListRow = m_ShowListRows.item(iCount);
		
		sAmountDue = ListRow.getAttribute("AMOUNT");
		sAmountOS = ListRow.getAttribute("AMOUNTOUTSTANDING") ;
		sRebateOrAddition = ListRow.getAttribute("REBATEORADDITION");
		
		if(sAmountDue != '') dTotalAmountDue = dTotalAmountDue + parseFloat(sAmountDue) ;
		if(sAmountOS != '') dTotalAmountOS = dTotalAmountOS + parseFloat(sAmountOS) ; 
		<% /* BMIDS957 Use Trim function from validation.js */ %>
		if(sRebateOrAddition.trim() != '')
			dTotalRebateOrAddition = dTotalRebateOrAddition + parseFloat(sRebateOrAddition);
	}
	
	frmScreen.txtTotalAmountDue.value			= dTotalAmountDue ;
	frmScreen.txtTotalAmountOutstanding.value	= dTotalAmountOS ;
	if(dTotalRebateOrAddition != 0) frmScreen.txtTotalRebateAddition.value = dTotalRebateOrAddition ;
}

function DisplayPaymentTotals()
{
	var iCount, ListRow;
	var dTotalAmountPaid, dTotalRefundAmount ;
	var sAmountPaid, sRefundAmount, sPaymentEvent ;
	var sPaymentMethod;

	dTotalAmountPaid = 0 ;
	dTotalRefundAmount = 0 ;

	for (iCount = 0; iCount < m_ShowListPaymentRows.length ; iCount++)
	{
		ListRow = m_ShowListPaymentRows.item(iCount);
		
		sAmountPaid		= ListRow.getAttribute("AMOUNTPAID");
		sRefundAmount	= ListRow.getAttribute("REFUNDAMOUNT") ;
		sPaymentEvent	= ListRow.getAttribute("PAYMENTEVENT") ;
		sPaymentMethod	= ListRow.getAttribute("PAYMENTMETHOD") ;

		<%//GD BM0372 START %>
		<%//if(sAmountPaid != '' && sPaymentEvent != sRebateAdditionId) dTotalAmountPaid = dTotalAmountPaid + parseFloat(sAmountPaid) ;%>
		if(sAmountPaid != '' && (!(IsRebateAddition(sPaymentEvent))) ) dTotalAmountPaid = dTotalAmountPaid + parseFloat(sAmountPaid) ;
		
		<%//GD BM0372 END %>
		if(sRefundAmount != '') dTotalRefundAmount = dTotalRefundAmount + parseFloat(sRefundAmount) ; 
	}
	
	<% /* MAR28 Change this column to hold the Total Amount Paid. Do not subtract refunds
	frmScreen.txtTotalAmountPaid.value	 = (dTotalAmountPaid-dTotalRefundAmount ==0) ? "" : (dTotalAmountPaid-dTotalRefundAmount);  */ %>
	frmScreen.txtTotalAmountPaid.value	 = dTotalAmountPaid;

	frmScreen.txtTotalRefundAmount.value = (dTotalRefundAmount == 0) ? "" : dTotalRefundAmount;
}

function PopulateScreen()
{
	// IVW  Set up XML for Phase 2 Payments
	
	var sFile;
	var ErrorTypes = new Array("RECORDNOTFOUND");
	
	ApplFeeTypeXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
		
	// Setup Application Fee Types
	ApplFeeTypeXML.CreateRequestTag(window , "FINDFEETYPELIST");
	ApplFeeTypeXML.CreateActiveTag("APPLICATIONFEETYPE");
	ApplFeeTypeXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	ApplFeeTypeXML.SetAttribute("_COMBOLOOKUP_","1");
	// 	ApplFeeTypeXML.RunASP(document,"PaymentProcessingRequest.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			ApplFeeTypeXML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			ApplFeeTypeXML.SetErrorResponse();
		}

	// IVW-REPLACE WITH

	var ErrorReturn = ApplFeeTypeXML.CheckResponse(ErrorTypes);
	
	<% /* Get the ComboValueIds for 'Rebate/Addition', 'Deduction' and 'OtherFeeType' */ %>
	sDeductionId	   = ApplFeeTypeXML.GetComboIdForValidation("PaymentEvent", "D", null, document) ;
	sReturnOfFundsId   = ApplFeeTypeXML.GetComboIdForValidation("PaymentEvent", "F", null, document) ;	
	sOtherFeeCode	   = ApplFeeTypeXML.GetComboIdForValidation("OneOffCost", "OTH", null, document) ;	
		
	<% /* SAB - 05/04/2006 - EP323 */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<% /* AW - 03/05/2006 - EP323 */ %>
	if (XML.GetGlobalParameterBoolean(document, "PP10EnablePayFee"))	
	{
		<% /* MAR28 Enable the Pay Fees button if any fees are Awaiting Payment by credit card */ %>		
		if (m_bAwaitingPayment == true)
			frmScreen.btnPayFees.disabled = false;
		else
			frmScreen.btnPayFees.disabled = true;
	}
	else
	{
		frmScreen.btnPayFees.disabled = true;
	}
	XML = null;
	
	if (ErrorReturn[0] == true) CreateListBoxXML();
	else return ;
	
	<% /* EP529 / MAR1652  Enable the Pay Fees button depending on User authority level */ %>
	
	if (m_comboXML.IsInComboValidationList(document,"UserRole", m_iUserRole, ["TF"]))
	{
		<% /* MAR28 Enable the Pay Fees button if any fees are Awaiting Payment by credit card */ %>		
		
		if (m_bAwaitingPayment == true)
			frmScreen.btnPayFees.disabled = false;
		else
			frmScreen.btnPayFees.disabled = true;
	}
	<% /* EP529 / MAR1652 End */ %>
				
	PopulateFeeTypesListBox(0);

<%	/* PSC 23/11/01 SYS3164 - Start */ %>
	if (m_iShowListRowsLength > 0)
	{
		spnTable.onclick()
	} 
<%	/* PSC 23/11/01 SYS3164 - End */ %>
}

function CreateListBoxXML()
{
<%	// IVW
	// The list returned from FindFeeTypeList contains FEE TYPES with sub Elements
	// containing payments for that FEE TYPE. this does not lend itself to list box 
	// displaying in ShowList(int) because it works on the basis of one active tag list.  
	// Therefore create a specific list for APPLICATIONFEETYPE using a true DOM document.
	
	// All payments are deducted for their respective FEE TYPE here also. Again this
	// would have been very messy with our XML functions as only one tag can be kept
	// active.
%>	
	var RootTag;
	var FeeLineTag;
	var PayLineTag;
	var iNumberOfRows;
	var iNumberOfPaymentsForFee;
	var sPaymentMethod, sPaymentMethodDesc;
	var sDateOfPayment, sUserId;
	var sRefundDate, sRefundAmount ;
	var sAmountOutstandingParse;
	var dPaymentTotal=0.000;
	var sFeePaySeqNum;
	var sAmountPaid;
	var sChequeNumber;
	var iLineCount = 0;
	var iCount3 = 0;
	var iPaymentHeader = 0;
	var sTotalAmountDue;
	var sFeeTypeComboSeq;
	var sRecordSeqNumber = 0;
	var xmlNode = null ;
	var sCondition ;
	var sPaymentEvent, sPaymentEventDesc ;
	var sRebateAdditionAmount, dRebateAdditionAmount ;
	var sInterfaceStatus;
	var sTransactionReference;   // MAR28
	var sSavingsAccountNumber;   // MAR28
	var sCardFailResultCode;     // MAR28
			
	ApplFeeTypeXMLListBox = new ActiveXObject("microsoft.XMLDOM");
	
	RootTag = ApplFeeTypeXMLListBox.createElement("LISTBOX");
	ApplFeeTypeXMLListBox.appendChild(RootTag);
	
	ApplFeeTypeXML.ActiveTag = null;	
	ApplFeeTypeXML.CreateTagList("APPLICATIONFEETYPE");
	
	iNumberOfRows = ApplFeeTypeXML.ActiveTagList.length;
	
	<% /* MAR28 Initialse Credit Card total awaiting payment */ %>
	m_dCreditCardTotal = 0;
	
	<% /* PB 06/06/06 EP696/MAR1761 Initialise Awaiting Payment flag */ %>
	m_bAwaitingPayment = false;	
	<% /* PB EP696/MAR1791 End */ %>
	
	for (iCount = 0; iCount < iNumberOfRows; iCount++)
	{
		sAmountOutstandingParse = "";
		
		ApplFeeTypeXML.SelectTagListItem(iCount);
		
		FeeLineTag = ApplFeeTypeXMLListBox.createElement("LISTLINE");
		RootTag.appendChild(FeeLineTag);
		
		FeeLineTag.setAttribute("LINETYPE","PAYHEADING");
		
		sFeeTypeDesc		= ApplFeeTypeXML.GetAttribute("FEETYPE_TEXT");
		sFeeTypeComboSeq	= ApplFeeTypeXML.GetAttribute("FEETYPE");	
		sRecordSeqNumber	= ApplFeeTypeXML.GetAttribute("FEETYPESEQUENCENUMBER");	
		
		if(sFeeTypeComboSeq == sOtherFeeCode) sFeeTypeDesc = ApplFeeTypeXML.GetAttribute("DESCRIPTION");
					
        <% /*  DRC     16/01/02    SYS3528 Ensure that a blank  feetype can still be selected */ %>
		if(sFeeTypeComboSeq=="11") // other, this hard check confirmed with KR/IK 15/12/2000 
		{
			sFeeTypeDesc	= ApplFeeTypeXML.GetAttribute("DESCRIPTION");
			// SR 23/02/01 SYS1968 : for otherfee type, show description. If description is empty show it as 'Other Fee'
			if(sFeeTypeDesc == null || sFeeTypeDesc == "" ) sFeeTypeDesc = "Other Fee";
		}
	
		if(sFeeTypeDesc == null || sFeeTypeDesc == "") sFeeTypeDesc = "   ";
	
		FeeLineTag.setAttribute("FEETYPE_TEXT",sFeeTypeDesc);	
		FeeLineTag.setAttribute("FEETYPE",sFeeTypeComboSeq);	
		FeeLineTag.setAttribute("FEETYPESEQUENCENUMBER",sRecordSeqNumber);
		
		sTotalAmountDue		= ApplFeeTypeXML.GetAttribute("AMOUNT");
		
			<% /* DRC 14/05/2004 CC71 */ %>
		if(sTotalAmountDue == null || sTotalAmountDue == "") sTotalAmountDue = "0";
		FeeLineTag.setAttribute("AMOUNT",sTotalAmountDue);
		
		<% /* extract Payments for the Fee type */ %>
		ApplFeeTypeXML.CreateTagList("FEEPAYMENT");	
		iNumberOfPaymentsForFee = ApplFeeTypeXML.ActiveTagList.length;
		
		<%/* total up amounts first! */ %>
		dPaymentTotal=0.000;		
		dRebateAdditionAmount = 0.000;
		
		for (iCount3 = 0; iCount3 < iNumberOfPaymentsForFee; iCount3++)
		{
			ApplFeeTypeXML.SelectTagListItem(iCount3);
			
			sPaymentEvent = ApplFeeTypeXML.GetAttribute("PAYMENTEVENT");
			sAmountPaid	= ApplFeeTypeXML.GetAttribute("AMOUNTPAID");
			
			<%//GD BM0372 START%>
			<%//if(sPaymentEvent==sRebateAdditionId)%>
			if ( IsRebateAddition(sPaymentEvent) )
			<%//GD BM0372 END%>
			{
				sRebateAdditionAmount = sAmountPaid ;
				dRebateAdditionAmount = dRebateAdditionAmount + parseFloat(sRebateAdditionAmount) ;	
			}
				
			<% /* MAR28 If FeePayment record has been created for Valuation Refund, Amount Paid will be 0 */ %>						
			if (sAmountPaid != "")						
				dPaymentTotal = dPaymentTotal + parseFloat(sAmountPaid);
		}
		
		for (iCount3 = 0; iCount3 < iNumberOfPaymentsForFee; iCount3++)
		{
			ApplFeeTypeXML.SelectTagListItem(iCount3);
			
			sPaymentEvent = ApplFeeTypeXML.GetAttribute("PAYMENTEVENT");
			
			if (sPaymentEvent != "")
			{
				if (m_comboXML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["POPA"]))
					sAmountOutstandingParse = "0";
			}
		}
		
		if (sAmountOutstandingParse == "")
			sAmountOutstandingParse = parseFloat(sTotalAmountDue) - dPaymentTotal;
		
		FeeLineTag.setAttribute("AMOUNTOUTSTANDING",sAmountOutstandingParse);
		FeeLineTag.setAttribute("AMOUNTPAID",dPaymentTotal);
		sRebateAdditionAmount = (dRebateAdditionAmount == 0) ? " " : (-1) * dRebateAdditionAmount ;
		FeeLineTag.setAttribute("REBATEORADDITION",sRebateAdditionAmount);
		
		for (iCount2 = 0; iCount2 < iNumberOfPaymentsForFee; iCount2++)
		{
			ApplFeeTypeXML.SelectTagListItem(iCount2);
			
			sInterfaceStatus = ApplFeeTypeXML.GetAttribute("COMPLETIONINDICATOR")

			PayLineTag = ApplFeeTypeXMLListBox.createElement("LISTLINE");
			RootTag.appendChild(PayLineTag);
			
			PayLineTag.setAttribute("LINETYPE","PAYMENT");
			
			PayLineTag.setAttribute("FEETYPE_TEXT",sFeeTypeDesc);	
			PayLineTag.setAttribute("FEETYPE",sFeeTypeComboSeq);	
			
			sFeePaySeqNum	= ApplFeeTypeXML.GetAttribute("PAYMENTSEQUENCENUMBER");
			PayLineTag.setAttribute("FEEPAYMENTSEQNUMBER",sFeePaySeqNum);	

			PayLineTag.setAttribute("FEETYPESEQUENCENUMBER",sRecordSeqNumber);
			
			sAmountPaid	= ApplFeeTypeXML.GetAttribute("AMOUNTPAID");
			PayLineTag.setAttribute("AMOUNTPAID",sAmountPaid);	
			
			sPaymentEvent		= ApplFeeTypeXML.GetAttribute("PAYMENTEVENT");
			sPaymentEventDesc	= ApplFeeTypeXML.GetAttribute("PAYMENTEVENT_TEXT");
			sRefundDate			= ApplFeeTypeXML.GetAttribute("REFUNDDATE");
			sRefundAmount		= ApplFeeTypeXML.GetAttribute("REFUNDAMOUNT");
			sCardFailResultCode = ApplFeeTypeXML.GetAttribute("CARDFAILRESULTCODE");   // MAR28

			PayLineTag.setAttribute("PAYMENTEVENT", sPaymentEvent);
			PayLineTag.setAttribute("PAYMENTEVENT_TEXT", sPaymentEventDesc);
			PayLineTag.setAttribute("REFUNDDATE", sRefundDate);
			PayLineTag.setAttribute("REFUNDAMOUNT", sRefundAmount);
			PayLineTag.setAttribute("CARDFAILRESULTCODE", sCardFailResultCode);       // MAR28
			
			sNotes				= ApplFeeTypeXML.GetAttribute("NOTES");	
			if(sNotes == null) 	sNotes = "  ";
			PayLineTag.setAttribute("NOTES",sNotes);

			sCondition = "//PAYMENTRECORD[@PAYMENTSEQUENCENUMBER= " + sFeePaySeqNum + "]"
			xmlNode = ApplFeeTypeXML.XMLDocument.selectSingleNode(sCondition);
			if (xmlNode == null)
			{
			 	alert('Error finding the payment details.')
				sDateOfPayment = " "	
			}
			else
			{
				sChequeNumber = xmlNode.getAttribute("CHEQUENUMBER");
				if(sChequeNumber == null) sChequeNumber = "  ";
				sPaymentMethod		= xmlNode.getAttribute("PAYMENTMETHOD");
<% /* AQR SYS3524 Allow for null payment method */ %>				
				if (sPaymentMethod == null) sPaymentMethod = " ";
				sPaymentMethodDesc	= xmlNode.getAttribute("PAYMENTMETHOD_TEXT");
				if (sPaymentMethodDesc == null) sPaymentMethodDesc = " ";
				sUserId				= xmlNode.getAttribute("USERID");
				if(sUserId == null) sUserId = " " ;

				sDateOfPayment = xmlNode.getAttribute('CREATIONDATETIME');
				if(sDateOfPayment == null) sDateOfPayment = " " ;
				else sDateOfPayment = sDateOfPayment.substr(0, 10);
				
				sTransactionReference = xmlNode.getAttribute("TRANSACTIONREFERENCE");  // MAR28
				if(sTransactionReference == null) sTransactionReference = " " ;        // MAR28
				sSavingsAccountNumber = xmlNode.getAttribute("SAVINGSACCOUNTNUMBER");  // MAR28
				if(sSavingsAccountNumber == null) sSavingsAccountNumber = " " ;        // MAR28
			}

			PayLineTag.setAttribute("CHEQUENUMBER",sChequeNumber);	
			PayLineTag.setAttribute("PAYMENTMETHOD",sPaymentMethod);
			PayLineTag.setAttribute("PAYMENTMETHOD_TEXT",sPaymentMethodDesc);
			PayLineTag.setAttribute("DATEOFPAYMENT",sDateOfPayment);
			PayLineTag.setAttribute("USERID",sUserId);
			
			PayLineTag.setAttribute("TOTALAMOUNTDUE",sTotalAmountDue);		
			PayLineTag.setAttribute("AMOUNTOUTSTANDING",sAmountOutstandingParse);
			PayLineTag.setAttribute("COMPLETIONINDICATOR",sInterfaceStatus);
			PayLineTag.setAttribute("TRANSACTIONREFERENCE",sTransactionReference);     // MAR28
			PayLineTag.setAttribute("SAVINGSACCOUNTNUMBER",sSavingsAccountNumber);     // MAR28
			PayLineTag.setAttribute("CARDFAILRESULTCODE",sCardFailResultCode);         // MAR28
			
			<% /* MAR28 Add up the feepayments of type CD (Credit or debit card) which have a payment
						event of type AW (Awaiting Payment) */ %>

			if (sPaymentMethod != " ")
			{
				<% /* EP1089/MAR1929 GHun Exclude deductions, rebate/additions and return of funds */ %>
				if ((sPaymentEvent != sDeductionId) && 
					(sPaymentEvent != sReturnOfFundsId) &&
					(!m_comboXML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["O", "RA"])) &&
					(m_comboXML.IsInComboValidationList(document,"FeePaymentMethod", sPaymentMethod, ["CD"])) &&
					(m_comboXML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["AW"])))
				{
					m_bAwaitingPayment = true;
					m_dCreditCardTotal = m_dCreditCardTotal + parseFloat(sAmountPaid);
				}			                     
			}	
		}
		
		// the original tag list of fees needs to be recreated...
		ApplFeeTypeXML.ActiveTag = null;	
		ApplFeeTypeXML.CreateTagList("APPLICATIONFEETYPE");
	}
}

<% /* Populate the ListBox displaying the FeeTypes */ %>
function ShowFeeTypesList(nStart)
{
	var sFeeTypeDesc, sFeeType, sAmountOwing, sFeeTypeSeqNum, sFeeTypeSeq, sRebateAdditionAmount ; 
	var iCount, ListRow;
	var sAmountOS,sTotalAmountDue;
<%		
	// IVW - this function is called upon entry to populate the list box. also gets called after add/edit etc
	
	// SR 23/02/01-SYS1867: Do not clear the table as it sets the selectedRowIndex to null. The null for this
	//						causes this problem 
	//scScrollTable.clear(); 
%>	
	for (iCount = 0; iCount < m_iShowListRowsLength && iCount < m_iTableLength; iCount++)
	{
		ListRow = m_ShowListRows.item(iCount + nStart);
		
		sLineType			= ListRow.getAttribute("LINETYPE");	

		sFeeTypeDesc		= "  " ;
		sAmountOwing		= "  " ;
		sAmountOS			= "  " ;
		sAmountPaid			= "  " ;
		
		sFeeTypeDesc		= ListRow.getAttribute("FEETYPE_TEXT");
		sFeeType			= ListRow.getAttribute("FEETYPE");
		sAmountOwing		= ListRow.getAttribute("AMOUNT");
		sAmountOS			= ListRow.getAttribute("AMOUNTOUTSTANDING");
		sAmountPaid			= ListRow.getAttribute("AMOUNTPAID");
		sRebateAdditionAmount = ListRow.getAttribute("REBATEORADDITION");
				
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sFeeTypeDesc);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sAmountOwing);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sAmountOS);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), sRebateAdditionAmount);

		sFeeTypeSeqNum = ListRow.getAttribute("FEETYPESEQUENCENUMBER");
		tblTable.rows(iCount+1).setAttribute("FEETYPESEQUENCENUMBER", sFeeTypeSeqNum);
		tblTable.rows(iCount+1).setAttribute("FEETYPE_TEXT", sFeeTypeDesc);
		tblTable.rows(iCount+1).setAttribute("FEETYPE", sFeeType);
		tblTable.rows(iCount+1).setAttribute("LINETYPE", sLineType);
	
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}
}

function ShowPaymentsList(nStart)
{
	var sPaymentEvent, sPaymentEventDesc, sAmountPaid, sRefundDate, sRefundAmount ; 
	var sUserId, sPaymentMethod, sPaymentMethodDesc, sChequeNumber, sDateOfPayment ;
	var sNotes, sFeePaymentSeqNumber, sInterfaceStatus;
	<% /* MAR28 */ %>
	var sTransactionReference;   
	var sSavingsAccountNumber;   
	var sCardFailResultCode;   
	
	var iCount, ListRow ;	
	for (iCount = 0; iCount < m_ShowListPaymentRows.length && iCount < m_iPaymentsTableLength; iCount++)
	{
		ListRow = m_ShowListPaymentRows.item(iCount + nStart);
		
		sPaymentEvent		  = ListRow.getAttribute("PAYMENTEVENT");
		sPaymentEventDesc	  = ListRow.getAttribute("PAYMENTEVENT_TEXT");
		sAmountPaid			  = ListRow.getAttribute("AMOUNTPAID");
		sDateOfPayment		  = ListRow.getAttribute("DATEOFPAYMENT");
		sUserId				  = ListRow.getAttribute("USERID");
		sPaymentMethod		  = ListRow.getAttribute("PAYMENTMETHOD");
		sPaymentMethodDesc	  = ListRow.getAttribute("PAYMENTMETHOD_TEXT");
		sChequeNumber		  = ListRow.getAttribute("CHEQUENUMBER");
		sRefundAmount		  = ListRow.getAttribute("REFUNDAMOUNT");
		sRefundDate			  = ListRow.getAttribute("REFUNDDATE");
		sNotes				  = ListRow.getAttribute("NOTES");
		sFeePaymentSeqNumber  = ListRow.getAttribute("FEEPAYMENTSEQNUMBER");
		sInterfaceStatus      = ListRow.getAttribute("COMPLETIONINDICATOR");
		sTransactionReference = ListRow.getAttribute("TRANSACTIONREFERENCE");  
		sSavingsAccountNumber = ListRow.getAttribute("SAVINGSACCOUNTNUMBER");
		sCardFailResultCode   = ListRow.getAttribute("CARDFAILRESULTCODE");
		
		<%//GD BM0372 START %>
		<%//if(sPaymentEvent == sRebateAdditionId) %>
		if ( IsRebateAddition(sPaymentEvent) )
		<%//GD BM0372 END %>
		{
			if(sAmountPaid >= 0) sPaymentEventDesc = 'Rebate' ;
			else
			{
				sPaymentEventDesc = 'Addition' ;
				sAmountPaid		  = sAmountPaid.substr(1) ; <%/* Show unsigned amount in the ListBox */%>
			}
		}
		else
		{
			<% /* if paymentEvent is not 'Deduction', get the description for PaymentMethod 
			      from XMLFeePaymentMethod  */ %>
			//If Paymentevent is anything other than a RA, then  get the payment method.
			
			//ORIGif(sPaymentEvent != sDeductionId) sPaymentMethodDesc = GetFeePaymentMethod(sPaymentMethod);
			if (sPaymentEvent != "")
			{
				if (!(m_comboXML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["RA"]))) 
					sPaymentMethodDesc = GetFeePaymentMethod(sPaymentMethod);
			}
		}
		
		scScreenFunctions.SizeTextToField(tblTable2.rows(iCount+1).cells(0), sPaymentEventDesc);
		scScreenFunctions.SizeTextToField(tblTable2.rows(iCount+1).cells(1), sDateOfPayment);
		scScreenFunctions.SizeTextToField(tblTable2.rows(iCount+1).cells(2), sAmountPaid);
		scScreenFunctions.SizeTextToField(tblTable2.rows(iCount+1).cells(3), sUserId);
		scScreenFunctions.SizeTextToField(tblTable2.rows(iCount+1).cells(4), sPaymentMethodDesc);
		scScreenFunctions.SizeTextToField(tblTable2.rows(iCount+1).cells(5), sChequeNumber);
		scScreenFunctions.SizeTextToField(tblTable2.rows(iCount+1).cells(6), sRefundAmount);
		scScreenFunctions.SizeTextToField(tblTable2.rows(iCount+1).cells(7), sRefundDate);
		
		tblTable2.rows(iCount+1).setAttribute("PAYMENTEVENT", sPaymentEvent);
		tblTable2.rows(iCount+1).setAttribute("PAYMENTMETHOD", sPaymentMethod);
		tblTable2.rows(iCount+1).setAttribute("CHEQUENUMBER", sChequeNumber);
		tblTable2.rows(iCount+1).setAttribute("USERID", sUserId);
		tblTable2.rows(iCount+1).setAttribute("NOTES", sNotes);
		tblTable2.rows(iCount+1).setAttribute("FEEPAYMENTSEQNUMBER", sFeePaymentSeqNumber);
		tblTable2.rows(iCount+1).setAttribute("COMPLETIONINDICATOR", sInterfaceStatus);
		tblTable2.rows(iCount+1).setAttribute("TRANSACTIONREFERENCE", sTransactionReference);
		tblTable2.rows(iCount+1).setAttribute("SAVINGSACCOUNTNUMBER", sSavingsAccountNumber);
		tblTable2.rows(iCount+1).setAttribute("CARDFAILRESULTCODE", sCardFailResultCode);
	}
}

function GetFeePaymentMethod(sFeePaymentMethod)
{
	var sPattern ;
	var xmlNode = null ;
	
	<% /* MAR28 Check that the Payment Method is not empty */ %>
	if (sFeePaymentMethod != " ")
	{
		sPattern = "//LISTENTRY[VALUEID=" + sFeePaymentMethod + "]" + "/VALUENAME"
		xmlNode = XMLFeePaymentMethod.XMLDocument.selectSingleNode(sPattern)
		if(xmlNode == null) return ""
		else return xmlNode.text ;
	}
	else
		return ""
}

function SetButtonStateOnInit()
{
	frmScreen.btnMakePayment.disabled = true ;
	frmScreen.btnEditPayment.disabled = true ;
	frmScreen.btnDeletePayment.disabled = true ;
}

function btnSubmit.onclick()
{
	// BMIDS977 set flag so that Remodel is not performed again on unload.
	m_bFromOK = true;

	if (m_sContext == "CompletenessCheck")
		frmToGN300.submit();
	else
	{
		//BMIDS977  Remodel
		RemodelOnLeavingPP010();
		
		frmToMN070.submit();
	}	
	
}
<%// GD BM0372 Added START %>
function IsRebateAddition(sPaymentEvent)
{
	if (sPaymentEvent != "")
	{
		if(m_comboXML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["RA"]) && m_comboXML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["O"])) 
			return(true);
		else
			return(false); 
	}
	else
	{
		return(false);
	}
}
<%// GD BM0372 Added END %>
function GetAcceptedQuoteData()
{
 var AppQuoteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
		
	// Setup Application data
	AppQuoteXML.CreateRequestTag(window , "GETACCEPTEDQUOTEDATA");
	AppQuoteXML.CreateActiveTag("APPLICATION");
	AppQuoteXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
	AppQuoteXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber );
	
	
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			AppQuoteXML.RunASP(document,"AQGetAcceptedQuoteData.asp");
			break;
		default: // Error
			AppQuoteXML.SetErrorResponse();
		}

   if(AppQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE") != null) //JD BMIDS945 
   {
		if(AppQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER")!= "")
		{
			m_sMortgageSubQuoteNumber  = AppQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER");
			return;
	    }		
   }
}
<% /* BMIDS957  Remove TrimIt function. Trim is available in validation.js */ %>

<% /* BMIDS977  Add new functions */ %>
function window.onunload()
{
	// If navigating away from the screen other than via the OK button, force a remodel.
	if (m_bFromOK == false)
	{
		RemodelOnLeavingPP010();
	}	
}

function RemodelOnLeavingPP010 ()
{
	<% /* Remodel if any of the rebate/addition values have changed. */ %>
	
	var iCount;
	var ListRow;
	var sRebateOrAddition;	
	var dRebateOrAddition;	
	var bRemodel = false;
	var sFeeType;
		
	for (iCount = 0; iCount < m_iShowListRowsLength ; iCount++)
	{
		ListRow = m_ShowListRows.item(iCount);
		
		sRebateOrAddition = ListRow.getAttribute("REBATEORADDITION");
		sFeeType = ListRow.getAttribute("FEETYPE");
		
		dRebateOrAddition = 0;
		if(sRebateOrAddition.trim() != '')
			dRebateOrAddition = parseFloat(sRebateOrAddition);
			
		if (dRebateOrAddition != m_dCurrentStateRebateAdditions[iCount]) 			
		{
			if (m_comboXML.IsInComboValidationList(document,"OneOffCost", sFeeType, ["APR"]))		
				bRemodel = true;
		}
	}			 

	if (bRemodel == true) 
	{
		// Alert the user that Remodelling is taking place
		window.status = "Re-calculating Quotation, please wait....";		
		RemodelQuotation();  
		window.status = "";
	}	
}
	
function RemodelQuotation()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window, null)
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("CALLEDFROM", "PP010");
	XML.RunASP(document,"RemodelQuotation.asp");
	if(XML.IsResponseOK())
	{
		return true;
	}
	else
	{
		return false;
	} 
}
function SetCurrentStateRebateAddition()
{
	var iCount;
	var ListRow;
	var sRebateOrAddition;
	
	for (iCount = 0; iCount < m_iShowListRowsLength ; iCount++)
	{
		ListRow = m_ShowListRows.item(iCount);
		
		sRebateOrAddition = ListRow.getAttribute("REBATEORADDITION");
		
		if(sRebateOrAddition.trim() != '')
		{
			m_dCurrentStateRebateAdditions[iCount] = parseFloat(sRebateOrAddition);   
		}
		else
		{	
			m_dCurrentStateRebateAdditions[iCount] = 0;
		}	
	}
}
function frmScreen.btnPayFees.onclick()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sReturn = null;
	
	var ArrayArguments = new Array();
	
	ArrayArguments[0]= XML.CreateRequestAttributeArray(window);	
	ArrayArguments[1]= m_sApplicationNumber;
	ArrayArguments[2]= m_dCreditCardTotal;	  // Total fee amount
	ArrayArguments[3]= m_sUserID;             

	sReturn = scScreenFunctions.DisplayPopup(window, document, "PP120.asp", ArrayArguments, 630, 365);
	
	<% /* If the card payment has been successful, refresh the screen */ %>

	if ((sReturn != null) && (sReturn[0] == true))
	{
		m_bAwaitingPayment = false;    // All payments have been made
		PopulateScreen();
		PopulatePaymentsListBox();
	}
	
	XML = null;
}
-->
</script>
</body>
</html>



