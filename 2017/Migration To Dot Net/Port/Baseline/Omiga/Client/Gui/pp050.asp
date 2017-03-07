<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<%/*
Workfile:      pp050.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Payment History
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog		Date			Description
IVW			31/01/2001		First version.		
CL			05/03/01		SYS1920 Read only functionality added
APS			15/03/01		SYS2083 Payee Record Not Found error handled
MC			19/03/01		SYS2073 Handle Payments with no Payee.
MV			19/03/01		SYS2049  Routing 
MV			21/03/01		SYS2072 Corrected the Return Of Funds 
MC			26/03/01		SYS2138 Amend label of 'Fees to be deducted...'
SR			06/09/01		SYS2412
SR			12/12/2001		SYS3391/ SYS2215
SR			13/12/01		SYS3363
JLD			06/03/02		SYS4177 Use NetPaymentAmount in listbox
LD			23/05/02		SYS4727 Use cached versions of frame functions
DB			29/05/02		SYS4767 - MSMS to Core Integration
SG			05/06/02		SYS4818 - Correct MSMS integration error
SG			10/06/02		SYS4845 - Correct MSMS integration error
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		20/07/2002	BMIDS0213	Modified - frmScreen.btnReturnFunds.onclick()
MV		06/08/2002	BMIDS0294	Core Upgrade inline with the functionality Ref AQRS SYS4728,SYS4559,SYS5129,SYS5110
								Modified - RetrieveContextData(); EnableDisableButtons();PopulateScreen()
MV		20/08/2002  BMIDS00294  Refresh the page 
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
PSC		25/10/2002	BMIDS00599  Amend enabling of Return Of funds button so that disbursements have to be
                                returned in reverse order
PSC		05/11/2002	BMIDS00599	Take into account incentive releases for disablement of Return of Funds button
PSC		11/11/2002	BMIDS00910	Allow return of funds in any order as long as Initial Release is last
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
PSC		28/11/2002	BMIDS01099  Amend to allow return of incentives
MDC		16/12/2002	BM0195		Replace Shortfall in table with Cheque Number
MDC		16/12/2002	BM0194		Add Screen Rules processing to all buttons
BS		25/02/2003	BM0271		Disable buttons in ReadOnly mode
MV		11/04/2003	BM0276		commented the Code for btnEditDisbursement.disabled = true 
BS		23/04/2003	BM0276		Corrected comment tags
PJO     04/07/2003  BMIDS576    btnEditDisbursement - one line left in
HMA     25/09/2003  BM0198      Do not include waived fees in total amounts.
MC		30/04/2004	BMIDS0468	Cancel & Decline stage freeze screens.
KRW     21/10/2004  BM0083      Enabled Return of Funds button for Interface Failed Cases
KRW     26/10/2004  BM0083      Return of Funds button now allows processing for PaymentStatus != m_sInterfaceFailed
HMA     08/12/2004  BMIDS957    Remove VBScript function.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :
HM		21/09/2005	MAR49		New controls ComplDate,Title Number103 were added. a new button Delay Completion was added
MV		19/10/2005	MAR178		Amended HTML Cosmetic Changes
MV		03/11/2005	MAR398		Amended btnSubmit.onclick();btnadddisbursement.onclick()
MV		04/11/2005	MAR407		Amended spnTable.onclick()
HM		09/11/2005	MAR378		Amended DisableButtons()
HMA     30/11/2005  MAR736      Use TMCompletionsStageID global parameter
PSC		01/12/2005	MAR728		Correct enabling and disabling ROF and Cancel Payment Buttons
HMA		01/03/2006  MAR1249     Change validation of Completion Date
HMA     15/03/2006  MAR1428     Correct initialisation after cancellation of payment.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History :
SAB		31/03/2006	EP8			Allow a shortfall payment
SAB		13/04/2006	EP385		MAR1608
SAB		13/04/2006	EP387		MAR1408
PB		20/04/2006	EP529		MAR1547 Changed LaunchAdHocTask to set the time on the TaskDueDateAndTime
PB		07/06/2006	EP696		MAR1803 Omiga - Delayed completion -date unknown - original date still displayed
PE		07/07/2006	EP953		Completion date in COT screen not being populated
PSC		26/02/2007	EP2_1347	Valuation Fee Refund processing
DS		23/03/2007	EP2_1920	Added condition in function ResetPostCompletionTasks(),performs certain tasks only
							     if there is a valid combo value for type "PC".
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<HEAD>
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>

<script src="validation.js" language="JScript"></script>
</HEAD>

<BODY>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets */ %>

<% //DB SYS4767 - MSMS Integration %>
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<% //DB End %>

<form id="frmClickOK" method="post" action="mn070.asp" STYLE="DISPLAY: none"></form>
<form id="frmToPP070" method="post" action="PP070.asp" STYLE="DISPLAY: none"></form>
<form id="frmToPP060" method="post" action="PP060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToPP080" method="post" action="PP080.asp" STYLE="DISPLAY: none"></form>
<form id="frmToPP090" method="post" action="PP090.asp" STYLE="DISPLAY: none"></form>

<span id="spnListScroll">
	<span style="LEFT: 305px; POSITION: absolute; TOP: 414px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 60px; LEFT: 10px; POSITION: absolute; TOP: 50px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Completion Date
		<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
			<input id="txtCompletionDate" type="text" class="msgTxt" style="HEIGHT:20px; WIDTH: 100px" maxlength="10">
		</span>
	</span>
	<span style="TOP: 10px; LEFT: 300px; POSITION: ABSOLUTE" class="msgLabel">
		Title Number 1
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtTitleNumber1" type="text" class="msgTxt" style="HEIGHT:20px; WIDTH: 200px">
		</span>
	</span>
	<span style="TOP: 35px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Title Number 2
		<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
			<input id="txtTitleNumber2" type="text" class="msgTxt" style="HEIGHT:20px; WIDTH: 200px">
		</span>
	</span>
	<span style="TOP: 35px; LEFT: 300px; POSITION: ABSOLUTE" class="msgLabel">
		Title Number 3
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtTitleNumber3" type="text" class="msgTxt" style="HEIGHT:20px; WIDTH: 200px" >
		</span>
	</span>
</div>
<div style="HEIGHT: 350px; LEFT: 10px; POSITION: absolute; TOP: 110px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 0px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Amount Requested
		<span style="TOP: -3px; LEFT: 180px; POSITION: ABSOLUTE">
			<input id="txtAmountRequested" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	<span style="TOP: 0px; LEFT: 300px; POSITION: ABSOLUTE" class="msgLabel">
		Total Fees added to Loan
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtFeesAddedToLoan" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>

	<span style="TOP: 20px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Drawdown
		<span style="TOP: -3px; LEFT: 180px; POSITION: ABSOLUTE">
			<input id="txtDrawDown" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	<span style="TOP: 20px; LEFT: 300px; POSITION: ABSOLUTE" class="msgLabel">
		Outstanding Fees to be Deducted
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtFeesDeductedFromAdvance" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>

	<span style="TOP: 40px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Retention Amount
		<span style="TOP: -3px; LEFT: 180px; POSITION: ABSOLUTE">
			<input id="txtRetentionAmount" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	<span style="TOP: 40px; LEFT: 300px; POSITION: ABSOLUTE" class="msgLabel">
		Total Completion Amount
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtTotalDisbursed" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>

	<span style="TOP: 60px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Balance
		<span style="TOP: -3px; LEFT: 180px; POSITION: ABSOLUTE">
			<input id="txtBalance" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	<span style="TOP: 60px; LEFT: 300px; POSITION: ABSOLUTE" class="msgLabel">
		Incentives to be Disbursed 
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtOutstandingIncentives" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<% //DB SYS4767 - MSMS Integration %>
	<% //<span style="LEFT: 4px; POSITION: absolute; TOP: 90px"> %>
	
	<!-- SG 26/03/02 MSMS009 -->
	<span style="TOP: 80px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Shortfall Payment
		<span style="TOP: -3px; LEFT: 180px; POSITION: ABSOLUTE">
			<input id="txtShortfallPayment" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt" maxlength="6">
		</span>
	</span>
	<!-- SG 26/03/02 MSMS009 -->
	<span style="TOP: 80px; LEFT: 300px; POSITION: ABSOLUTE" class="msgLabel">
		Total Fee Refunds
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtTotalRefund" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>

	<span style="TOP: 100px; LEFT: 300px; POSITION: ABSOLUTE" class="msgLabel">
		Valuation Fee Refund to be Disbursed
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtOutstandingValFeeRefund" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>

	<span style="LEFT: 104px; POSITION: absolute; TOP: 120px">
		<input id="btnAcceptShortfall" value="Add Shortfall to Inital Advance" type="button" style="WIDTH: 180px" class="msgButton">
	</span>
	
	<span style="LEFT: 500px; POSITION: absolute; TOP: 120px">
	<% //DB End %>
		<input id="btnCancelBalance" value="Cancel Balance" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	
	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 145px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="10%" class="TableHead">Issue  Date</td>	
				<td width="15%" class="TableHead">Payment Type</td>
				<% /* BM0195 MDC 16/12/2002 - End
				<td width="10%" class="TableHead">Shortfall</td>  */ %>
				<td width="10%" class="TableHead">Cheque Number</td>
				<% /* BM0195 MDC 16/12/2002 - End  */ %>
				<td width="10%" class="TableHead">Payment Amount</td>
				<td width="15%" class="TableHead">Payment Method</td>
				<td width="15%" class="TableHead">Payment Status</td>
				<td width="15%" class="TableHead">Payee Name</td>
				<td width="10%" class="TableHead">Advance Date</td>
				<!-- SG 10/06/02 SYS4845
				<td width="10%" class="TableHead">Cancel Date</td> -->
				</tr>
			<tr id="row01">		
				<td width="10%" class="TableTopLeft">&nbsp;</td>		
				<td width="15%" class="TableTopCenter">&nbsp;</td>
				<td width="10%" class="TableTopCenter">&nbsp;</td>		
				<td width="10%" class="TableTopCenter">&nbsp;</td>		
				<td width="15%" class="TableTopCenter">&nbsp;</td>	
				<td width="15%" class="TableTopCenter">&nbsp;</td>
				<td width="15%" class="TableTopCenter">&nbsp;</td>
		
				<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		
				<td width="10%" class="TableLeft">&nbsp;</td>			
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		
				<td width="10%" class="TableLeft">&nbsp;</td>			
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td></tr>	
			<tr id="row04">		
				<td width="10%" class="TableLeft">&nbsp;</td>
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		
				<td width="10%" class="TableLeft">&nbsp;</td>			
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>			
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		
				<td width="10%" class="TableLeft">&nbsp;</td>			
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		
				<td width="10%" class="TableLeft">&nbsp;</td>			
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		
				<td width="10%" class="TableLeft">&nbsp;</td>			
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="10%" class="TableMiddleCenter">&nbsp;</td>		
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>	
				<td width="15%" class="TableMiddleCenter">&nbsp;</td>
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		
				<td width="10%" class="TableBottomLeft">&nbsp;</td>	
				<td width="15%" class="TableBottomCenter">&nbsp;</td>
				<td width="10%" class="TableBottomCenter">&nbsp;</td>		
				<td width="10%" class="TableBottomCenter">&nbsp;</td>		
				<td width="15%" class="TableBottomCenter">&nbsp;</td>	
				<td width="15%" class="TableBottomCenter">&nbsp;</td>	
				<td width="15%" class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomRight">&nbsp;</td></tr> 
		</table>
	</span>
	
	<span id="spnButtons" style="LEFT:4px; POSITION: absolute; TOP: 330px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnAddDisbursement" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		
		<%/* <span style="LEFT: 70px; POSITION: absolute; TOP: 0px">
			<input id="btnEditDisbursement" value="Edit" type="button" style="WIDTH: 60px; Display:NONE" class="msgButton">
		</span> */ %>

		<span style="LEFT: 70px; POSITION: absolute; TOP: 0px">
			<input id="btnCopyDisbursement" value="Copy" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="LEFT: 150px; POSITION: absolute; TOP: 0px">
			<input id="btnDelayCompletion" value="Delay Completion" type="button" style="WIDTH: 100px" class="msgButton">
		</span> 
		<span style="LEFT: 260px; POSITION: absolute; TOP: 0px">
			<input id="btnPayeeDetails" value="Payee Details" type="button" style="WIDTH: 100px" class="msgButton">
		</span> 

 		<span style="LEFT: 370px; POSITION: absolute; TOP: 0px">
			<input id="btnReturnFunds" value="Return Funds" type="button" style="WIDTH: 100px" class="msgButton">
		</span>

 		<span style="LEFT: 480px; POSITION: absolute; TOP: 0px">
			<input id="btnCancelPayment" value="Cancel Payment" type="button" style="WIDTH: 100px" class="msgButton">
		</span>
	</span>
</div>

<div style="HEIGHT: 50px; LEFT: 8px; POSITION: absolute; TOP: 465px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Fee Payment Details</strong>
	</span>
	
	<span style="TOP: 24px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Total Payment Amount
		<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
			<input id="txtTotalFeePayment" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<span style="TOP: 24px; LEFT: 400px; POSITION: ABSOLUTE" class="msgLabel">
		Payment Type
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtFeePaymentType" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
</div>

</form>

<% /* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 520px; WIDTH: 600px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>
<!-- #include FILE="fw030.asp" -->

<% //DB SYS4767 - MSMS Integration %>
<% //SG 03/04/02 MSMS009 %>
<!-- #include FILE="attribs/PP050attribs.asp" -->
<% //DB End %>

<% /* BMIDS957  remove VBScript function */ %>

<%/* CODE */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var scClientScreenFunctions;

var m_iTableLength = 9;
var m_iTable2Length = 4;
var ApplFeeTypeXML = null;
var PayeeHistoryXML = null;
var XMLCombos = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = null;
var m_ShowListRows = null;
var m_sPayMethod = null;
var m_iDisbursementRole = 0;
var m_iCancelBalanceRole = 0;
var m_iReturnFundsRole = 0;
var m_iUserRole=0;
//WP13 MAR49
var m_bUseSanctionProcess = false;
var m_iUseSanctionProcess ;
var m_sCompletionDate = "";
var m_bCompletionDateValid = false;
var m_sRotGUID ="";

//MV - 05/08/2002
var m_sShortFallAmount = "0";
var m_TotalPaymentAmount = "0";

//DB SYS4767 - MSMS Integration
//SG 20/05/02 MSMS0097
var m_sOldBalance = 0;

//SG 27/03/02 MSMS009
var m_iShortfallMinimumAmount = 0;
var m_iShortfallMaximumAmount = 0;
//DB End

var m_blnReadOnly = false;
var m_sInitialAdvanceId , m_sReturnOfFundsId, m_sInstallmentId, m_sRetentionReleaseId, m_sBalanceCanellationId;
var m_sIncentiveReleaseId, m_sInterfaceFailed;
var m_sUnsanctionedId, m_sPaidId ;
var m_sSanctionedId, m_sAwaitingInterfaceResponseId, m_sInterfacedId, m_sInterfacedNotPaidId, m_sCancellationId ;
//BM0198
var m_sRebateFee ;

<% /* PSC 26/02/2007 EP2_1347 */ %>
var m_sValuationRefundId;

function RetrieveContextData()
{
<%
/*	
	//scScreenFunctions.SetContextParameter(window,"idApplicationNumber", "A00009709");
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber", "A100002048");
	//scScreenFunctions.SetContextParameter(window,"idApplicationNumber", "C00078387");
	scScreenFunctions.SetContextParameter(window, "idApplicationFactFindNumber", 1);	
	scScreenFunctions.SetContextParameter(window,"idRole","90");
	scScreenFunctions.SetContextParameter(window,"idUserId","SRAO");
	scScreenFunctions.SetContextParameter(window,"idReadOnly","0");
	scScreenFunctions.SetContextParameter(window, "idProcessingIndicator", "1");
	scScreenFunctions.SetContextParameter(window, "idFreezeDataIndicator", "0");
	
	// END TEST
*/
%>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);	
	m_iUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	
	<%/* SG 24/07/02 SYS5129 See if user's already placed it in context*/%>
	m_sShortFallAmount = scScreenFunctions.GetContextParameter(window,"idShortfallPayment","0");
	if (m_sShortFallAmount == "")
		m_sShortFallAmount = "0";

}

<% /* EVENTS */  %>

function window.onload()
{	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	<%/* scScreenFunctions = new scScrFunctions.ScreenFunctionsObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>	
	
	FW030SetTitles("Payment History","PP050",scScreenFunctions);
	
	RetrieveContextData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	GetParameterValues();
	GetComboValues();
	
	if(!AcceptedQuote())
	{
		frmClickOK.submit();
	}
	else
	{
		//alert('here')
		PopulateScreen();
		<% /* BS BM0271 25/02/03 
		Moved to after the ReadOnly indicator is set */ %>
		//DisableButtons();
		m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP050");
		<% /* BS BM0271 25/02/03 */ %>
		DisableButtons();	
		scScreenFunctions.SetFocusToFirstField(frmScreen);
		//BMIDS0468
		if(scScreenFunctions.IsCancelDeclineStage(window))
		{
			frmScreen.btnAcceptShortfall.disabled=true;
			frmScreen.btnAddDisbursement.disabled=true;
			<% /*MAR49 * a new button was added */%>
			frmScreen.btnDelayCompletion.disabled=true;
			frmScreen.btnCancelBalance.disabled=true;
			frmScreen.btnCancelPayment.disabled=true;
			frmScreen.btnCopyDisbursement.disabled=true;
			frmScreen.btnPayeeDetails.disabled=true;
			frmScreen.btnReturnFunds.disabled=true;
		}
	}
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}	

function AcceptedQuote()
{
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var AppXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var AppXML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>	
	
	AppXML.CreateRequestTag(window, null)
	AppXML.CreateActiveTag("APPLICATION");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AppXML.RunASP(document,"GetApplicationData.asp");
	if(AppXML.IsResponseOK())
	{
		AppXML.SelectTag(null, "APPLICATIONFACTFIND");
		if(AppXML.GetTagText("ACCEPTEDQUOTENUMBER")== "")
		{
			alert("Unable to access Disbursements as there is no accepted quotation for this application.");
			return false;
		}
		else
		{
			if(AppXML.GetTagText("APPLICATIONAPPROVALDATE")== "")
			{
				alert('The application must be approved before disbursements can be created.');
				return false ;
			}
		}
	}
	else return false ;

	return true;
}

<%/* MAR29 */%>
function frmScreen.txtCompletionDate.onblur()
{
	<% /* MAR1249  Do not do further validation if the format of the date is invalid */ %>
	if (frmScreen.txtCompletionDate.valid() == true)
	{
		<% /* PB 07/06/2006 EP696/MAR1803 Begin 
		if ((frmScreen.txtCompletionDate.value != "") && (m_sCompletionDate != frmScreen.txtCompletionDate.value))
		{
			if (ValidateCompletionDate())
			{
				m_sCompletionDate = frmScreen.txtCompletionDate.value;

				if (frmScreen.txtRetentionAmount.value == "0"  &&  
					frmScreen.txtBalance.value == "0" && 
					frmScreen.txtOutstandingIncentives.value == "0" )  
						frmScreen.btnAddDisbursement.disabled = true;
				else
						frmScreen.btnAddDisbursement.disabled = false;
			}
			else
			{
				frmScreen.btnAddDisbursement.disabled = true;
				frmScreen.txtCompletionDate.focus();
			}
		}
	} */ %>
		if (frmScreen.txtCompletionDate.value != "") 
		{
			if (m_sCompletionDate != frmScreen.txtCompletionDate.value)
			{
				if (ValidateCompletionDate())
				{
					m_sCompletionDate = frmScreen.txtCompletionDate.value;

					<% /* PSC 26/02/2007 EP2_1347 */ %>
					if (frmScreen.txtRetentionAmount.value == "0"  &&  
						frmScreen.txtBalance.value == "0" && 
						frmScreen.txtOutstandingIncentives.value == "0" &&
						frmScreen.txtOutstandingValFeeRefund.value == "0")  
							frmScreen.btnAddDisbursement.disabled = true;
					else
							frmScreen.btnAddDisbursement.disabled = false;
				}
				else
				{
					frmScreen.btnAddDisbursement.disabled = true;
					frmScreen.txtCompletionDate.focus();
				}		
			}
		}
		else
		{
			m_sCompletionDate = "";
		}
	} <% /* PB 07/06/2006 EP696/MAR1803 End */ %>
	else
	{
		m_sCompletionDate = "";
		frmScreen.btnAddDisbursement.disabled = true;
		frmScreen.txtCompletionDate.focus();
		
		alert("The Completion Date is invalid - please enter a valid date in the format DD/MM/YYYY");

	}
}	

function ValidateCompletionDate()
{
	// if the date is not valid disable add button, set focus on date field
		m_bCompletionDateValid = false;
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, "REQUEST")	//XML.CreateRequestTag(window, null)
		XML.SetAttribute("OPERATION","ValidateCompletionDate");
		XML.CreateActiveTag("VALIDATECOMPLETIONDATE");
		XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.SetAttribute("COMPLETIONDATE", frmScreen.txtCompletionDate.value );

		//XML.RunASP(document,"ValidateCompletionDate.asp");
		XML.RunASP(document,"PaymentProcessingRequest.asp");
		
		var TagRESPONSE = XML.SelectTag(null,"RESPONSE");
		if (XML.SelectTag(TagRESPONSE,"ERROR") != null)
		{
			var sErrorMessage = XML.GetTagText("DESCRIPTION");
			alert(sErrorMessage);
		}
		else
		{
			m_bCompletionDateValid = true;
		}
		var XML=null;
		return m_bCompletionDateValid;
}
<%/* MAR29 */%>

function spnTable.onclick()
{
	var iCurrentRow = scScrollTable.getRowSelected();
	
	if(iCurrentRow != -1) 
	{
		<%/* Enable or disable buttons 'Edit' and 'Paid' */%>
		var sPaymentType = tblTable.rows(iCurrentRow).getAttribute("PaymentType") ;
		var sPaymentStatus = tblTable.rows(iCurrentRow).getAttribute("PaymentStatus");
		<% /* PSC 11/11/2002 BMIDS00599 */ %>
		var sCancellationDate = tblTable.rows(iCurrentRow).getAttribute("CancelDate");
		
		<% /* BMIDS957 Use trim function from validation.js */ %>
		sCancellationDate = sCancellationDate.trim();
		
		<% /* BS BM0271 25/02/03 
		Only enable buttons if not in ReadOnly mode */ %>
		if (!m_blnReadOnly)
		{
			frmScreen.btnCopyDisbursement.disabled = false ;
			
			EnableDisableButtons(sPaymentType, sPaymentStatus, sCancellationDate);	
			<% /* WP13 MAR49 */ %>
			frmScreen.btnDelayCompletion.disabled = false;
		}
		<% /* BS BM0271 End 25/02/03 */%>
		
		<%/* if payee details associated with the payment are null, enable PayeeDetails button */%>
		if(tblTable.rows(iCurrentRow).cells(5).innerText != '') frmScreen.btnPayeeDetails.disabled = false;
		else frmScreen.btnPayeeDetails.disabled = true;
		
<% /* SR 13/07/01	if(m_iUserRole >= m_iDisbursementRole && sPaymentStatus == m_sUnsanctioned) 
			frmScreen.btnEditDisbursement.disabled = false ;
		else frmScreen.btnEditDisbursement.disabled = true ;
*/%>
				
		<% /* SR 13/07/01 : SYS2412 - Display Fee Payment Details */ %>
		var sFeePaymentAmount = tblTable.rows(iCurrentRow).getAttribute("FeePaymentAmount");
		
		if((sPaymentType == m_sInitialAdvanceId || sPaymentType == m_sInstallmentId) && sFeePaymentAmount != "0")
		{
			frmScreen.txtTotalFeePayment.value = sFeePaymentAmount ;
			frmScreen.txtFeePaymentType.value  = "Deduction" ;
		}
		else
		{
			if(sPaymentType == m_sReturnOfFundsId && sFeePaymentAmount != "0")
			{
				frmScreen.txtTotalFeePayment.value = sFeePaymentAmount ;
				frmScreen.txtFeePaymentType.value  = "Return of Funds" ;
			}
			else
			{
				frmScreen.txtTotalFeePayment.value = "" ;
				frmScreen.txtFeePaymentType.value  = ""
			}
		}
		
		//bmids0468
		
		if(scScreenFunctions.IsCancelDeclineStage(window))
		{
			frmScreen.btnAcceptShortfall.disabled=true;
			frmScreen.btnAddDisbursement.disabled=true;
			frmScreen.btnCancelBalance.disabled=true;
			frmScreen.btnCancelPayment.disabled=true;
			frmScreen.btnCopyDisbursement.disabled=true;
			frmScreen.btnPayeeDetails.disabled=true;
			frmScreen.btnReturnFunds.disabled=true;
		}
	}
}

function EnableDisableButtons(sPaymentType, sPaymentStatus, sCanellationDate)
{
	switch (sPaymentType)
	{
		<% /* PSC 28/11/2002 BMIDS01099 */ %>
		<% /* PSC 01/12/2005 MAR728 - Start */ %>
		<% /* PSC 26/02/2007 EP2_1347 */ %>
		case m_sValuationRefundId:
		case m_sInitialAdvanceId:
		case m_sInstallmentId:
		case m_sIncentiveReleaseId:
		case m_sRetentionReleaseId:
		{
			switch (sPaymentStatus)
			{
				case m_sUnsanctionedId:
				case m_sSanctionedId:
				{
					if(sCanellationDate == "" && m_iUserRole >= m_iDisbursementRole)
						frmScreen.btnCancelPayment.disabled = false ;
					break;
				}
				case m_sPaidId:
				{
					if (m_bUseSanctionProcess == "0")
						frmScreen.btnCancelPayment.disabled = false ;
					
					if (m_bUseSanctionProcess == "1")
						frmScreen.btnReturnFunds.disabled = DisableROFButton();	
					break;
				}
				default:
				{
					frmScreen.btnReturnFunds.disabled = DisableROFButton();	
					break;
				}
			}
			break;
		}
		<% /* PSC 01/12/2005 MAR728 - End */ %>
		case m_sReturnOfFundsId:
			<% /* PJO BMID0576 frmScreen.btnEditDisbursement.disabled = false ; */ %> 
			frmScreen.btnReturnFunds.disabled = true ;
			<% /* PSC 05/11/2002 BMIDS00599 */ %>
			frmScreen.btnCancelPayment.disabled = true ;
		break;
		case m_sBalanceCanellationId:
			<%/*  frmScreen.btnEditDisbursement.disabled = true ;*/%>
			frmScreen.btnCopyDisbursement.disabled = true ;
			frmScreen.btnReturnFunds.disabled = true ;
			
			if(sPaymentStatus == m_sUnsanctionedId)
			{	
				if(sCanellationDate == "")	frmScreen.btnCancelPayment.disabled = false ;
			}
			else frmScreen.btnCancelPayment.disabled = true ;
			
		break;
	}
}

function frmScreen.btnAddDisbursement.onclick()
{
	
	if (m_bCompletionDateValid == false)
		ValidateCompletionDate();
	
	if (m_bCompletionDateValid == false)
	{
		window.alert("The completion date is not valid.");
		m_sCompletionDate="";
		frmScreen.txtCompletionDate.focus();
		return;
	}

	if(m_iUserRole < m_iDisbursementRole) 
	{
		alert('You do not have the authority to add a payment.');
		return
	}
	
	UpdateReportOnTitle();
	
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
			XML.CreateActiveTag('PAYMENTDETAILS');
			XML.ActiveTag.appendChild(ApplFeeTypeXML.XMLDocument.selectSingleNode("//CALCULATIONS"));
			XML.SetAttributeOnTag("CALCULATIONS", "COMPLETIONDATE", m_sCompletionDate)
			
			scScreenFunctions.SetContextParameter(window,"idXml", XML.XMLDocument.xml);
			scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
			scScreenFunctions.SetContextParameter(window,"idXml2", ApplFeeTypeXML.XMLDocument.xml);
	
			frmToPP060.submit();
			break;
		default: // Error
	}
	
}
<%
/* SR 07-12-2001 : SYS3390 - Hide the EDIT button
function frmScreen.btnEditDisbursement.onclick()
{
	var iCurrentRow = scScrollTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		window.alert("Select a payment for editing.") ;
		return ;
	}

	if(m_iUserRole < m_iDisbursementRole)
	{
		window.alert("You do not have the authority to edit a payment.") ;
		return ;
	}
	// Set required values in context and call PP060
	var iCount = tblTable.rows(iCurrentRow).getAttribute("ListCount");
	var AppXML = new scXMLFunctions.XMLObject();	
	XML.CreateActiveTag('PAYMENTDETAILS');
	XML.ActiveTag.appendChild(ApplFeeTypeXML.XMLDocument.selectSingleNode("//CALCULATIONS"));
	
	var sCondition = '//PAYMENTRECORD[@AMOUNT $ne$ 0][DISBURSEMENTPAYMENT]';
	var xmlNodeList = ApplFeeTypeXML.XMLDocument.selectNodes(sCondition);
	XML.ActiveTag.appendChild(xmlNodeList(iCount));
	
	scScreenFunctions.SetContextParameter(window,"idXml", XML.XMLDocument.xml);
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
	
	frmToPP060.submit();
}
*/
%>
function frmScreen.btnCopyDisbursement.onclick()
{
	var iCurrentRow = scScrollTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		window.alert("Select a payment for editing.") ;
		return ;
	}
	
	if(m_iUserRole < m_iDisbursementRole)
	{
		window.alert("You do not have the authority to copy a payment.") ;
		return ;
	}
	
	<% /* BM0194 MDC 16/12/2002 */ %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK

			// Set the values to passed from this screens, in the context
			var iCount = tblTable.rows(iCurrentRow).getAttribute("ListCount");

			<%/* SG 05/06/02 SYS4818 */%>	
			<%/* DB SYS4767 - MSMS Integration */%>
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			<%/* var XML = new scXMLFunctions.XMLObject(); */%>
			<%/* DB End */%>	
			<%/* SG 05/06/02 SYS4818 */%>	
	
			XML.CreateActiveTag('PAYMENTDETAILS');
			XML.ActiveTag.appendChild(ApplFeeTypeXML.XMLDocument.selectSingleNode("//CALCULATIONS"));

			var sCondition = '//PAYMENTRECORD[@AMOUNT $ne$ 0][DISBURSEMENTPAYMENT]';
			var xmlNodeList = ApplFeeTypeXML.XMLDocument.selectNodes(sCondition);
			XML.ActiveTag.appendChild(xmlNodeList(iCount));
			scScreenFunctions.SetContextParameter(window,"idXml", XML.XMLDocument.xml);
	
			scScreenFunctions.SetContextParameter(window,"idMetaAction","Copy");
	
			frmToPP060.submit();
			break;
		default: // Error
	}
	<% /* BM0194 MDC 16/12/2002 - End */ %>
}

function frmScreen.btnPayeeDetails.onclick()
{
	<% /* BM0194 MDC 16/12/2002 */ %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			scScreenFunctions.SetContextParameter(window,"idReturnScreenId", "PP050.asp");
			frmToPP070.submit();
			break;
		default: // Error
	}
	<% /* BM0194 MDC 16/12/2002 - End */ %>
}

function frmScreen.btnReturnFunds.onclick()
{
	var iCurrentRow = scScrollTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		window.alert("Select a payment for editing.") ;
		return ;
	}

	if(m_iUserRole < m_iReturnFundsRole)
	{
		alert("You do not have the authority to return a payment.");
		return ;
	}
	else
	{
		<% /* MV - 23/07/2002 - BMIDS0213 - Core Upgrade */ %>
		var sPaymentStatus = tblTable.rows(iCurrentRow).getAttribute("PaymentStatus");
		if(sPaymentStatus != m_sPaidId && sPaymentStatus != m_sAwaitingInterfaceResponseId &&
		   sPaymentStatus != m_sInterfacedId && sPaymentStatus != m_sInterfacedNotPaidId &&
		   sPaymentStatus != m_sInterfaceFailed)// KRW     26/10/2004  BM0083
		{
			alert('The current status of payment does not allow Return of Funds to be made.');
			return ;
		}
				
		var sPaymentType = tblTable.rows(iCurrentRow).getAttribute("PaymentType");
		<% /* PSC 28/11/2002 BMIDS01099 */ %>
		if(sPaymentType == m_sReturnOfFundsId || sPaymentType == m_sBalanceCanellationId)
		{
			alert('This type of payment does not allow Return of Funds to be made.');
			return ;
		}
	}

	<% /* BM0194 MDC 16/12/2002 */ %>
	var iRet = ScreenRules();
	if (iRet == 0 || iRet == 1)
	{	
		var iCount = tblTable.rows(iCurrentRow).getAttribute("ListCount");
	
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
		XML.CreateActiveTag('PAYMENTDETAILS');
		XML.ActiveTag.appendChild(ApplFeeTypeXML.XMLDocument.selectSingleNode("//CALCULATIONS"));

		var sCondition = '//PAYMENTRECORD[@AMOUNT $ne$ 0][DISBURSEMENTPAYMENT]';
		var xmlNodeList = ApplFeeTypeXML.XMLDocument.selectNodes(sCondition);
		XML.ActiveTag.appendChild(xmlNodeList(iCount));
	
		<% /* SR SYS2412 : Find out the maximum amount that can be returned and attach it to the XML */ %>
		var xmlPayment = xmlNodeList(iCount) ;
		var xmlFeePaymentList = xmlPayment.selectNodes("FEEPAYMENT");
		var xmlFeePayment, sFeePaymentAmount, sPaymentAmount;
		var dFeePaymentAmount = 0.00 ;
		var dPaymentAmount = 0.00 ;
	
		sPaymentAmount = getAttributeValue(xmlPayment, "AMOUNT");
		if(!isNaN(sPaymentAmount)) dPaymentAmount = parseFloat(sPaymentAmount);
	
		for(var iCount=0; iCount <= xmlFeePaymentList.length-1; iCount++)
		{
			xmlFeePayment = xmlFeePaymentList(iCount) ;
			XML.ActiveTag.appendChild (xmlFeePayment) ;
			sFeePaymentAmount = getAttributeValue(xmlFeePayment, "AMOUNTPAID"); 

			if(!isNaN(sFeePaymentAmount)) 
				dFeePaymentAmount = dFeePaymentAmount + parseFloat(sFeePaymentAmount);	
		}
	
		var xmlLCPList = xmlPayment.selectNodes("LOANCOMPONENTPAYMENT");
		var xmlLCP ;
		for(var iCount=0; iCount <= xmlLCPList.length-1; iCount++)
		{
			xmlLCP = xmlLCPList(iCount) ;
			XML.ActiveTag.appendChild(xmlLCP) ;
		}
	
		var xmlROFList, xmlROF, sROFAmount, sPaymentSeqNo ;
		var dROFAmount = 0.00 ;
		<% /* PSC 11/11/2002 BMIDS00599 */ %>
		var dGrossAmount = 0.00;
	
		sPaymentSeqNo =  getAttributeValue(xmlPayment, "PAYMENTSEQUENCENUMBER");
		sCondition = '//PAYMENTRECORD[@ASSOCPAYSEQNUMBER $eq$' + sPaymentSeqNo + ']';
	
		xmlROFList = ApplFeeTypeXML.XMLDocument.selectNodes(sCondition) ;

		for(var iCount=0; iCount <= xmlROFList.length-1; iCount++)
		{
			xmlROF = xmlROFList(iCount) ;
			sROFAmount = getAttributeValue(xmlROF, "AMOUNT"); 

			if(!isNaN(sROFAmount)) dROFAmount = dROFAmount + parseFloat(sROFAmount);
		}
	
		var dMaxROFAmount = 0.00 ;
	
		<% /* PSC 11/11/2002 BMIDS00599 - Start */ %>
		dGrossAmount = dPaymentAmount + dROFAmount;
		dMaxROFAmount = dGrossAmount - dFeePaymentAmount;
		XML.CreateTag("GROSSAMOUNT", dGrossAmount);
		<% /* PSC 11/11/2002 BMIDS00599 - Start */ %>
	
		XML.CreateTag("MAXROFAMOUNT", dMaxROFAmount);
		scScreenFunctions.SetContextParameter(window,"idXml", XML.XMLDocument.xml);
	
		frmToPP080.submit();
	}		
	<% /* BM0194 MDC 16/12/2002 - End */ %>
}

function getAttributeValue(xmlNode, sAttributeName)
{
	var iLength = xmlNode.attributes.length;

	for (var iLoop=0; iLoop < iLength; iLoop++)
	{
		var AttributeNode = xmlNode.attributes.item(iLoop);
		if (AttributeNode.nodeName == sAttributeName) return AttributeNode.nodeValue;
	}
	
	return "" ;
}

function frmScreen.btnCancelBalance.onclick()
{
	if(m_iUserRole < m_iCancelBalanceRole)
	{
		alert("You do not have the authority to cancel balance.");
		return ;
	}
	
	ApplFeeTypeXML.ActiveTag = null ;
	ApplFeeTypeXML.CreateTagList('CALCULATIONS');
	if(ApplFeeTypeXML.ActiveTagList.length > 0 )
	{
		<% /* BM0194 MDC 16/12/2002 */ %>
		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				ApplFeeTypeXML.SelectTagListItem(0);

				scScreenFunctions.SetContextParameter(window,"idXml", ApplFeeTypeXML.ActiveTag.xml);
				frmToPP090.submit();

				break;
			default: // Error
		}
		<% /* BM0194 MDC 16/12/2002 - End */ %>
	}
	else alert('Error in finding summary.');
	
}

function frmScreen.btnCancelPayment.onclick()
{
	if(m_iUserRole < m_iDisbursementRole)
	{
		window.alert("You do not have the authority to cancel payment.") ;
		return ;
	}	

	var iCurrentRow = scScrollTable.getRowSelected();	
	if(iCurrentRow == -1)
	{
		window.alert("Select a payment for cancellation.") ;
		return ;
	}	
		
	var iListCount = tblTable.rows(iCurrentRow).getAttribute("ListCount");
	var xmlNodeList = ApplFeeTypeXML.XMLDocument.selectNodes('//PAYMENTRECORD[@AMOUNT $ne$ 0][DISBURSEMENTPAYMENT]');
	var xmlNode = xmlNodeList(iListCount) ;
	var sTemp ;
	
	ApplFeeTypeXML.ActiveTag = xmlNode ;
	
	// Cancel the payment and refresh the screen
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var XML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>	
	
	var xmlRequest = XML.CreateRequestTag(window , 'UPDATEDISBURSEMENT');
	XML.CreateActiveTag('PAYMENTRECORD');
	XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	
	sTemp = ApplFeeTypeXML.GetAttribute('PAYMENTSEQUENCENUMBER');
	XML.SetAttribute('PAYMENTSEQUENCENUMBER', sTemp);
	XML.SetAttribute('AMOUNT', 0);
	
	XML.CreateActiveTag('DISBURSEMENTPAYMENT');
	XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	
	sTemp = ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYMENTSEQUENCENUMBER');
	XML.SetAttribute('PAYMENTSEQUENCENUMBER', sTemp);
	
	<% /* MO - BMIDS00807 */ %>
	<% /* sTemp = scScreenFunctions.DateToString(scScreenFunctions.GetSystemDate()) */ %>
	sTemp = scScreenFunctions.DateToString(scScreenFunctions.GetAppServerDate());
	XML.SetAttribute('CANCELLATIONDATE', sTemp);
	
	<% /* SR 07-12-2001 : SYS3391 - Append IssueDate to the request and change the PaymentStatus to 'Cancelled' */ %>
	
	sTemp = ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'ISSUEDATE');
	XML.SetAttribute('ISSUEDATE', sTemp);
	
	XML.SetAttribute('PAYMENTSTATUS', m_sCancellationId);
	<% /* End SR 07-12-2001 : SYS3391 */ %>
		
	// 	XML.RunASP(document,"PaymentProcessingRequest.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	if(!XML.IsResponseOK())
	{
		alert('Error in updating payments')
		return false ;
	}
	
	<% /* MAR1428  Reinitialise the screen correctly */ %>
	PopulateScreen();
	DisableButtons();	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function btnSubmit.onclick()
{
	UpdateReportOnTitle();
	
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			frmClickOK.submit();
			break;
		default: // Error
	}
}

<% /* FUNCTIONS */ %>
function GetParameterValues()
{
	// This is a Function to retrieve the Global Parameter Values from database
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var XML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>	

	var blnReturn = false;	
	//Preparing XML Request string 
	var XMLActiveTag = XML.CreateRequestTag(window,"SEARCH");

	XML.CreateActiveTag("GLOBALPARAMETER");
	XML.CreateTag("NAME","PPROCDISBURSEMENTROLE");
	XML.ActiveTag = XMLActiveTag;
	
	XML.CreateActiveTag("GLOBALPARAMETER");
	XML.CreateTag("NAME","PPROCCANCELBALANCEROLE");
	XML.ActiveTag = XMLActiveTag;
	
	XML.CreateActiveTag("GLOBALPARAMETER");
	XML.CreateTag("NAME","PPROCRETURNFUNDSROLE");
	XML.ActiveTag = XMLActiveTag;
	
	//DB SYS4767 - MSMS Integration
	//SG 27/03/02 MSMS009
	XML.CreateActiveTag("GLOBALPARAMETER");
	XML.CreateTag("NAME","PPROCSHORTFALLMINIMUMAMOUNT");
	XML.ActiveTag = XMLActiveTag;
	
	//SG 27/03/02 MSMS009
	XML.CreateActiveTag("GLOBALPARAMETER");
	XML.CreateTag("NAME","PPROCSHORTFALLMAXIMUMAMOUNT");
	XML.ActiveTag = XMLActiveTag;
	
	//WP13 MAR49
	XML.CreateActiveTag("GLOBALPARAMETER");
	XML.CreateTag("NAME","PPUSESANCTIONPROCESS");
	XML.ActiveTag = XMLActiveTag;
	//DB End
	
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
						case "PPROCDISBURSEMENTROLE" :
							m_iDisbursementRole = XML.GetTagText("AMOUNT");
							break;
						case "PPROCCANCELBALANCEROLE":
							m_iCancelBalanceRole = XML.GetTagText("AMOUNT");
							break;
						case "PPROCRETURNFUNDSROLE":
							m_iReturnFundsRole = XML.GetTagText("AMOUNT");
							break;
							
							//DB SYS4767 - MSMS Integration
							//SG 27/03/02 MSMS009
						case "PPROCSHORTFALLMINIMUMAMOUNT":
							m_iShortfallMinimumAmount = XML.GetTagText("AMOUNT");
							break;

						//SG 27/03/02 MSMS009
						case "PPROCSHORTFALLMAXIMUMAMOUNT":
							m_iShortfallMaximumAmount = XML.GetTagText("AMOUNT");
							break;
						// WP13 MAR49	
						case "PPUSESANCTIONPROCESS":
						{
							m_iUseSanctionProcess = XML.GetTagText("BOOLEAN");
							m_bUseSanctionProcess = XML.GetTagText("BOOLEAN");
							break;
						}
						//DB End
					}
				}
			}
		}		
	}
	else blnReturn = false;

	XML = null; 
	return blnReturn;
}

function GetComboValues()
{
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	XMLCombos = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* XMLCombos = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>	
	
	var sGroupList = new Array("PaymentStatus", "PaymentType", "PaymentEvent");  // BM0198
	XMLCombos.GetComboLists(document,sGroupList);
	
	var xmlPaymentStatusCombo		= XMLCombos.GetComboListXML("PaymentStatus");
	m_sUnsanctionedId				= XMLCombos.GetComboIdForValidation('PaymentStatus', 'U', xmlPaymentStatusCombo, document) ;
	m_sPaidId						= XMLCombos.GetComboIdForValidation('PaymentStatus', 'P', xmlPaymentStatusCombo, document) ;
	m_sSanctionedId					= XMLCombos.GetComboIdForValidation('PaymentStatus', 'S', xmlPaymentStatusCombo, document) ;
	m_sAwaitingInterfaceResponseId	= XMLCombos.GetComboIdForValidation('PaymentStatus', 'R', xmlPaymentStatusCombo, document) ;
	m_sInterfacedId					= XMLCombos.GetComboIdForValidation('PaymentStatus', 'I', xmlPaymentStatusCombo, document) ;
	m_sInterfacedNotPaidId			= XMLCombos.GetComboIdForValidation('PaymentStatus', 'INP', xmlPaymentStatusCombo, document) ;
	m_sCancellationId				= XMLCombos.GetComboIdForValidation('PaymentStatus', 'C', xmlPaymentStatusCombo, document) ;
		<% /* KRW 21/10/2004 BM0083 */ %>
	m_sInterfaceFailed              = XMLCombos.GetComboIdForValidation("PaymentStatus", 'IF', xmlPaymentStatusCombo, document);
	
	
	
	var xmlPaymentTypeCombo = XMLCombos.GetComboListXML("PaymentType");
	m_sInitialAdvanceId		= XMLCombos.GetComboIdForValidation("PaymentType", 'I', xmlPaymentTypeCombo, document);
	m_sInstallmentId		= XMLCombos.GetComboIdForValidation("PaymentType", 'A', xmlPaymentTypeCombo, document);
	m_sReturnOfFundsId		= XMLCombos.GetComboIdForValidation("PaymentType", 'N', xmlPaymentTypeCombo, document);
	<% /* PSC 05/11/2002 BMIDS00599 */ %>
	m_sRetentionReleaseId	= XMLCombos.GetComboIdForValidation("PaymentType", 'R', xmlPaymentTypeCombo, document);
	m_sBalanceCanellationId	= XMLCombos.GetComboIdForValidation("PaymentType", 'NCB', xmlPaymentTypeCombo, document);
	m_sIncentiveReleaseId	= XMLCombos.GetComboIdForValidation("PaymentType", 'C', xmlPaymentTypeCombo, document);
	
	<% /* PSC 26/02/2007 EP2_1347 */ %>
	m_sValuationRefundId = XMLCombos.GetComboIdForValidation("PaymentType", 'VALREFUND', xmlPaymentTypeCombo, document); 

	
	
	var xmlPaymentEventCombo = XMLCombos.GetComboListXML("PaymentEvent");
	m_sRebateFee             = XMLCombos.GetComboIdForValidation("PaymentEvent", 'O', xmlPaymentEventCombo, document);
}

function PopulateScreen()
{
	var sFile;
	var ErrorTypes = new Array("RECORDNOTFOUND");
	
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	ApplFeeTypeXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* ApplFeeTypeXML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>
			
	ApplFeeTypeXML.CreateRequestTag(window , "GETPAYMENTSUMMARY");
	ApplFeeTypeXML.CreateActiveTag("PAYMENTRECORD");
	ApplFeeTypeXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	ApplFeeTypeXML.SetAttribute("_COMBOLOOKUP_","1");
	ApplFeeTypeXML.RunASP(document,"PaymentProcessingRequest.asp");
	
	var ErrorReturn = ApplFeeTypeXML.CheckResponse(ErrorTypes);
		
	if (ErrorReturn[0] == true)
	{	
		<% /* Display values in the header elements */ %>
		ApplFeeTypeXML.CreateTagList('CALCULATIONS');
		if(ApplFeeTypeXML.ActiveTagList.length > 0)
		{
			if(ApplFeeTypeXML.SelectTagListItem(0))
			{
				frmScreen.txtAmountRequested.value = ApplFeeTypeXML.GetAttribute('AMOUNTREQUESTED');
				frmScreen.txtDrawDown.value = ApplFeeTypeXML.GetAttribute('DRAWDOWN');
				frmScreen.txtTotalDisbursed.value  = ApplFeeTypeXML.GetAttribute('TOTALTOBEDISBURSED'); 
				frmScreen.txtTotalRefund.value	   = ApplFeeTypeXML.GetAttribute('TOTALREFUND'); 
				frmScreen.txtRetentionAmount.value = ApplFeeTypeXML.GetAttribute('RETENTIONAMOUNT'); 
				frmScreen.txtOutstandingIncentives.value = ApplFeeTypeXML.GetAttribute('MORTGAGEINCENTIVE');  		
				frmScreen.txtFeesAddedToLoan.value = ApplFeeTypeXML.GetAttribute('FEESTOBEADDEDTOLOAN');
				<% /* PSC 26/02/2007 EP2_1347 */ %>
				frmScreen.txtOutstandingValFeeRefund.value = ApplFeeTypeXML.GetAttribute('OUTSTANDINGVALUATIONREFUND');
				
				<%/* SG 24/07/02 SYS5129 Created IF statement, existing code placed in Else */%>
				if (m_sShortFallAmount != "0")
				{	
					frmScreen.txtBalance.value = parseInt(ApplFeeTypeXML.GetAttribute('BALANCE')) + parseInt(m_sShortFallAmount);
					frmScreen.txtShortfallPayment.value = m_sShortFallAmount
				}
				else
					frmScreen.txtBalance.value = ApplFeeTypeXML.GetAttribute('BALANCE');
				
				<% /* MAR49 add button setting is blocked here due to a new condition: completion date is added */ %>
				//if (frmScreen.txtRetentionAmount.value == "0"  &&  frmScreen.txtBalance.value == "0" && frmScreen.txtOutstandingIncentives.value == "0" )  
				//	frmScreen.btnAddDisbursement.disabled = true;
				//else
				//	frmScreen.btnAddDisbursement.disabled = false;
				
				//DB SYS4767 - MSMS Integration
				//SG 20/05/02 MSMS0097
				m_sOldBalance = ApplFeeTypeXML.GetAttribute('BALANCE');
				//DB End
				frmScreen.txtFeesDeductedFromAdvance.value = ApplFeeTypeXML.GetAttribute('FEESTOBEDEDUCTEDFROMADVANCE');
			}
		//DB SYS4767 - MSMS Integration
		}		
		
		//SG 28/03/02 MSMS009 START
		//Check ApplFeeTypeXML to see whether Shortfall textbox and button should be disabled.	
			
		var bDisable = false;
		<% /* PB 07/06/2006 EP696/MAR1803 Begin */ %>			
		var bCancelled = false;
		<% /* EP696/MAR1803 End */ %>
		var sXSLPattern = "";
		
		//Check there's an initial Disbursement payment that's not been cancelled.
		ApplFeeTypeXML.ActiveTag = null;		
		sXSLPattern = "DISBURSEMENTPAYMENT[@PAYMENTTYPE='" + m_sInitialAdvanceId + "' $and$ @PAYMENTSTATUS!='" + m_sCancellationId + "']"
		if (ApplFeeTypeXML.CreateTagList(sXSLPattern).length > 0)
		<% /* PB 07/06/2006 EP696/MAR1803 Begin */ %>
		{
			bDisable = true;
		}
		else
		{
			bCancelled = true;
		}
		<% /* EP696/MAR1803 End */ %>
		//Check number of Loan Component elements		
		ApplFeeTypeXML.ActiveTag = null;
		if (ApplFeeTypeXML.CreateTagList("LOANCOMPONENT").length > 1)
			bDisable = true;

		//Enable or Disable
		if (bDisable == true)
		{
			//Disable Shortfall fields
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtShortfallPayment")
			frmScreen.btnAcceptShortfall.disabled = true;
		}
		else
		{
			//Enable Shortfall fields		
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtShortfallPayment")
			frmScreen.btnAcceptShortfall.disabled = false;
		//DB End
		}
	}
	else return ;
	
	PopulateListBox(0);
	if (tblTable.rows > 0) 	scScrollTable.setRowSelected(1);
	
	<% /* MARS49 */ %>
	//Get the ReportOnTitle data
	var RotXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	RotXML.CreateRequestTag(window,"GetReportOnTitleData");
	RotXML.CreateActiveTag("REPORTONTITLE");
	RotXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	RotXML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	RotXML.RunASP(document,"ReportOnTitle.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = RotXML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0] == true)
	{
		<% /* Get Completion Date, Titles, Solicitor Ref */ %>
		RotXML.SelectTag (null, "REPORTONTITLE");
		m_sRotGUID				= RotXML.GetAttribute("ROTGUID");

		<% /* EP953 - 07/07/2006 */ %>
		m_sCompletionDate			= RotXML.GetAttribute("COMPLETIONDATE");
		frmScreen.txtCompletionDate.value = m_sCompletionDate;
		
		<% /* EP696/MAR1823 End */ %>
		frmScreen.txtTitleNumber1.value = RotXML.GetAttribute("TITLENUMBER");
		frmScreen.txtTitleNumber2.value = RotXML.GetAttribute("TITLENUMBER2");
		frmScreen.txtTitleNumber3.value = RotXML.GetAttribute("TITLENUMBER3");
		var SolictorsRefNumber = RotXML.GetAttribute("SOLICITORSREFNUMBER");
	}	
	RotXML=null;
	
	<% /* Validate Solicitor Ref */ %>
	var SolXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	SolXML.CreateRequestTag(window,null); //SolXML.CreateRequestTag(window, "REQUEST")
	SolXML.SetAttribute("OPERATION","ValidateSolicitor");
	SolXML.CreateActiveTag("SOLICITOR");
	SolXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	SolXML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		
	//SolXML.RunASP(document,"ValidateSolicitor.asp");
	SolXML.RunASP(document,"PaymentProcessingRequest.asp");	
	
	var TagRESPONSE = SolXML.SelectTag(null,"RESPONSE");
	if (SolXML.SelectTag(TagRESPONSE,"ERROR") != null)
	{	//error message and disable add button else leave add button as is has been set by now
		var sErrorMessage = SolXML.GetTagText("DESCRIPTION");
		alert(sErrorMessage)
		frmScreen.btnAddDisbursement.disabled = true;
	}
	if (frmScreen.txtBalance.value == "0")
		 frmScreen.txtCompletionDate.disabled = true;
	else
		 frmScreen.txtCompletionDate.disabled = false;
	SolXML=null;
}


function PopulateListBox(nStart)
{
	if(!FindPayeeHistoryList()) return ;

	var sCondition = '//PAYMENTRECORD[@AMOUNT $ne$ 0][DISBURSEMENTPAYMENT]' ;
	var iNumberOfRows = ApplFeeTypeXML.XMLDocument.selectNodes(sCondition).length;
	
	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}

function ShowList(nStart)
{
	var iCount, xmlNodeList ;
	var blnInterfaced = false;
	var sIssuedate, sPaymentType, sPaymentAmount, sPaymentMethod, sPaymentStatus, sCompletionDate; 
	var sPaymentTypeDesc, sPaymentMethodDesc, sPaymentStatusDesc, sCancelDate ;
	<% /* BM0195 MDC 16/09/2002 */ %>
	//DB SYS4767 - MSMS Integration
	//SG 04/04/02 MSMS009
	// var sShortfall;  
	//DB End
	var sChequeNumber = "";
	<% /* BM0195 MDC 16/09/2002 - End */ %>
	
	var sPayeeHistorySeqNo, sPayeeName ;
	var sFeePaymentAmount ; <% /* Total adjusted against the fee in this payment */ %>
	var xmlNode ;
	
	scScrollTable.clear();
	
	var sCondition = '//PAYMENTRECORD[@AMOUNT $ne$ 0][DISBURSEMENTPAYMENT]' ;
	xmlNodeList = ApplFeeTypeXML.XMLDocument.selectNodes(sCondition);
	
	for (iCount = 0; iCount < xmlNodeList.length  && iCount < m_iTableLength; iCount++)
	{	
		xmlNode = xmlNodeList(iCount+nStart) ;
		sFeePaymentAmount = GetFeePaymentAmount(xmlNode);
		ApplFeeTypeXML.ActiveTag = xmlNode ;
		
		sIssuedate			= ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'ISSUEDATE');
		sPaymentType		= ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYMENTTYPE');
		sPaymentTypeDesc	= ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYMENTTYPE_TEXT');
		sPaymentAmount		= ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'NETPAYMENTAMOUNT'); // JLD SYS4177
		sPaymentMethod		= ApplFeeTypeXML.GetAttribute('PAYMENTMETHOD');
		sPaymentMethodDesc	= ApplFeeTypeXML.GetAttribute('PAYMENTMETHOD_TEXT');
		sPaymentStatus		= ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYMENTSTATUS');
		sPaymentStatusDesc	= ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYMENTSTATUS_TEXT');
		sPayeeHistorySeqNo	= ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYEEHISTORYSEQNO');

		if (sPaymentStatus == m_sInterfacedId) 
		  blnInterfaced = true;

		m_TotalPaymentAmount =  parseFloat(m_TotalPaymentAmount) + parseFloat(sPaymentAmount);
		
		//DB SYS4767 - MSMS Integration
		//SG 04/04/02 
		<% /* PSC 11/11/2002 BMIDS00599 */ %>
		sCancelDate		= ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'CANCELLATIONDATE');
		
		<% /* BM0195 MDC 16/12/2002
		//CL 21/05/02	
		if(sPaymentType == '10') //Initial Advance
		{
			sShortfall = ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'SHORTFALL');
		}
		else
		{
			sShortfall = 0;
		}
		//End CL		
		*/ %>
		sChequeNumber = ApplFeeTypeXML.GetAttribute('CHEQUENUMBER');
		<% /* BM0195 MDC 16/12/2002 - End  */ %>
		
		//sShortfall			= ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'SHORTFALL');
		//DB End
		
		// Cater for Payments with no Payee
		if(parseInt(sPayeeHistorySeqNo) > 0)
			sPayeeName = GetPayeeName(sPayeeHistorySeqNo);
		else
			sPayeeName = "";
		sCompletionDate		= ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'COMPLETIONDATE');
				
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sIssuedate);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sPaymentTypeDesc);
		//DB SYS4767 - MSMS Integration
		//scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sPaymentAmount);
		//scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), sPaymentMethodDesc);
		//scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4), sPaymentStatusDesc);
		//scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5), sPayeeName);
		//scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(6), sCompletionDate);
		//SG 04/04/02 MSMS009
		<% /* BM0195 MDC 16/12/2002
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sShortfall);	 */ %>
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sChequeNumber);
		<% /* BM0195 MDC 16/12/2002 - End  */ %>
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), sPaymentAmount);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4), sPaymentMethodDesc);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5), sPaymentStatusDesc);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(6), sPayeeName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(7), sCompletionDate);
		//DB End
		
		tblTable.rows(iCount+1).setAttribute("PaymentType", sPaymentType);
		tblTable.rows(iCount+1).setAttribute("PaymentMethod", sPaymentMethod);
		tblTable.rows(iCount+1).setAttribute("PaymentStatus", sPaymentStatus);
		
		<% /* PSC 25/10/2002 BMIDS00599 */ %>
		tblTable.rows(iCount+1).setAttribute("ListCount", iCount+nStart);
		tblTable.rows(iCount+1).setAttribute("FeePaymentAmount", sFeePaymentAmount);
		tblTable.rows(iCount+1).setAttribute("CancelDate", sCancelDate);
	}
	<% /* MAR1408 can't change the Completion Date any disbursement has been interfaced */ %>
	if (blnInterfaced == true)
	   frmScreen.txtCompletionDate.disabled = true;
}

function FindPayeeHistoryList()
{
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	PayeeHistoryXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* PayeeHistoryXML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>
	
	PayeeHistoryXML.CreateRequestTag(window, "FindPayeeHistoryList");
	PayeeHistoryXML.CreateActiveTag("PAYEEHISTORY");
	PayeeHistoryXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	PayeeHistoryXML.SetAttribute("_COMBOLOOKUP_","1");
	PayeeHistoryXML.RunASP(document,"PaymentProcessingRequest.asp");
	//SYS2083
	// if(!PayeeHistoryXML.CheckResponse())
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = PayeeHistoryXML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0] == false)
	{
		if (ErrorReturn[1] = ErrorTypes[0])
		{
			alert("No Payee details have been entered. Please add a Payee before continuing.");
			return false;
		}	
		else if (ErrorReturn[1] != ErrorTypes[0])
		{
			alert("Error retrieving Payee details");
			return false;
		}	
	}
	return true;
}

function GetPayeeName(sPayeeHistorySeqNo)
{	
	var sPayeeName = ""
	var xmlNode = null ;
	
	if(sPayeeHistorySeqNo == '') return sPayeeName ;

	var xmlNode = PayeeHistoryXML.XMLDocument.selectSingleNode('//PAYEEHISTORY[@PAYEEHISTORYSEQNO=' + sPayeeHistorySeqNo + ']')
	if(xmlNode != null )
	{
		PayeeHistoryXML.ActiveTag = xmlNode ;
		sPayeeName = PayeeHistoryXML.GetTagAttribute('THIRDPARTY', 'COMPANYNAME'); ;
	}
	else alert('Error retriving Payee Name');

	return sPayeeName ; 
}

function GetFeePaymentAmount(xmlNode)
{<% /* SR 13/07/01 : SYS2412 - Find total of the amount adjusted agianst Fee for a PaymentRecord */%>
	var iPaymentAmount = 0 ;
	var sAmount, xmlNode, xmlAttribNode ;
	var sPaymentEvent, xmlEventNode;
	
	var xmlNodeList = xmlNode.selectNodes(".//FEEPAYMENT");
	for( var iCount = 0 ; iCount < xmlNodeList.length ; iCount ++)
	{
		xmlNode = xmlNodeList(iCount) ;
		if (xmlNode.attributes.getNamedItem("AMOUNTPAID") != null)
		{
			xmlAttribNode = xmlNode.attributes.getNamedItem("AMOUNTPAID") ;
			sAmount = xmlAttribNode.value ;
			if(!isNaN(sAmount))
			{
				xmlEventNode = xmlNode.attributes.getNamedItem("PAYMENTEVENT") ;
				sPaymentEvent = xmlEventNode.value;
				if ( sPaymentEvent != m_sRebateFee)
				{	
					iPaymentAmount = iPaymentAmount + parseFloat(sAmount)
				}	
			} 
		}
	}
	return iPaymentAmount ;
}

function DisableButtons()
{	// Disable the edit buttons based on the user role
	if((m_iUserRole < m_iDisbursementRole) || m_blnReadOnly) <% /* BS BM0271 25/02/03 */ %>
	{
		frmScreen.btnAddDisbursement.disabled = true ;	 
		<% /* SR 13-12-2001 : SYS3363 Disable buttons Copy, ReturnOfFunds and Cancel on initialisation  */ %>
		<% /*
		frmScreen.btnEditDisbursement.disabled = true ;
		frmScreen.btnCopyDisbursement.disabled = true ;
		frmScreen.btnCancelPayment.disabled = true ; */
		%>
		<% /* SR 13-12-2001 : SYS3363 - End  */ %>
	}
	
	<% /* SR 13-12-2001 : SYS3363 Disable buttons Copy, ReturnOfFunds and Cancel on initialisation  */ %>
	<% /* if(m_iUserRole < m_iReturnFundsRole) frmScreen.btnReturnFunds.disabled = true ; */ %>
	<% /* SR 13-12-2001 : SYS3363 - End */ %>
	
	if((m_iUserRole < m_iCancelBalanceRole) || m_blnReadOnly) <% /* BS BM0271 25/02/03 */ %> 
	frmScreen.btnCancelBalance.disabled = true ;
	
	<% /* SR 13-12-2001 : SYS3363 Disable buttons Copy, ReturnOfFunds and Cancel on initialisation  */ %>
	frmScreen.btnCopyDisbursement.disabled = true ;
	frmScreen.btnReturnFunds.disabled = true ;
	frmScreen.btnCancelPayment.disabled = true ;
	<% /* BS BM0271 25/02/03 
	As additional controls have now been disabled an error occurred in SetFocusToFirstFields
	when trying to set the focus to btnEditDisbursement. It is never displayed but passes the
	tests in scScreenFunctions.DoFocusProcessing on the tabindex, disabled and visibility 
	attributes. So set it to disabled so it fails the test */ %>
	<%/*  frmScreen.btnEditDisbursement.disabled = true ; */%>
	if (m_blnReadOnly) frmScreen.btnAcceptShortfall.disabled = true; 
	<% /* BS BM0271 25/02/03 */ %>
	
	<% /* WP13 MAR49 */ %>
	<% /* MAR378  Disable btnDelayCompletion unconditionaly here
	//if (m_blnReadOnly || frmScreen.txtCompletionDate.value=="" || iCurrentRow == -1)  */ %>
	frmScreen.btnDelayCompletion.disabled = true; 
	
	<% /* PSC 26/02/2007 EP2_1347 */ %>
	if (m_blnReadOnly || frmScreen.txtCompletionDate.value=="" ||
						 frmScreen.txtRetentionAmount.value == "0"  &&  
						 frmScreen.txtBalance.value == "0" && 
						 frmScreen.txtOutstandingIncentives.value == "0" &&
						 frmScreen.txtOutstandingValFeeRefund == "0")  
							frmScreen.btnAddDisbursement.disabled = true;
	<% /* MAR1428 Ensure that the Add button is enabled when necessary */ %>
	else
		frmScreen.btnAddDisbursement.disabled = false;

}
//DB SYS4767 - MSMS Integration
function frmScreen.btnAcceptShortfall.onclick()	//SG 03/04/02 MSMS009
{
	//Validate shortfall payment
	if (parseInt(frmScreen.txtShortfallPayment.value) < parseInt(m_iShortfallMinimumAmount))
	{
		alert("Shortfall Payment is less than the Minimum in the Global Parameter table.")
		return false;
	}

	if (parseInt(frmScreen.txtShortfallPayment.value) > parseInt(m_iShortfallMaximumAmount))
	{
		alert("Shortfall Payment is greater than the Maximum in the Global Parameter table.");
		return false;
	}	

	<%/* EP8 Start */%> 
	if (frmScreen.txtShortfallPayment.value == "")
	   {
	     alert("No Shortfall payment to add to Initial Advance")
	     frmScreen.txtShortfallPayment.focus();
	     return false;
	   }
	<%/* EP8 End */%> 


	<% /* BM0194 MDC 16/12/2002 */ %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			//SG 20/05/02 MSMS0097
			frmScreen.txtBalance.value = parseInt(m_sOldBalance) + parseInt(frmScreen.txtShortfallPayment.value)
	        if (frmScreen.txtBalance.value == "0")
				frmScreen.txtCompletionDate.disabled = true;
			else 
			    frmScreen.txtCompletionDate.disabled = false;	
			//Valid - set in context	
			<% /* MSMS0063 - Pass in the text control's value not the buttons. */ %>
			scScreenFunctions.SetContextParameter(window,"idShortfallPayment", frmScreen.txtShortfallPayment.value);

			break;
		default: // Error
	}
	<% /* BM0194 MDC 16/12/2002 - End */ %>
//DB End
}

function LaunchAdHocTask()
{	<% /* WP13 MARS49 */ %>
	var bSuccess = false;
	var bTaskUpdate	= false;
	//get params..
	var stageXML =		new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var taskXML =		new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var gParamXML =		new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTaskID =		gParamXML.GetGlobalParameterString(document,"TMDelayCompletionTaskID");				
	var tagSourceAppl	= "Omiga"
	var sUserRole =		scScreenFunctions.GetContextParameter(window,"idRole",null);
	var sUserID =		scScreenFunctions.GetContextParameter(window,"idUserID",null)	
	var sUnitID =		scScreenFunctions.GetContextParameter(window,"idUnitID",null)
	var sApplPriority = scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null)
	var sActivityID =	scScreenFunctions.GetContextParameter(window,"idActivityId",null)
	var sStageID =		scScreenFunctions.GetContextParameter(window,"idStageID",null)
	
	//check if exists
	stageXML.CreateRequestTag(window,"GetCurrentStage");
	stageXML.CreateActiveTag("CASEACTIVITY");
	stageXML.SetAttribute("SOURCEAPPLICATION", tagSourceAppl);
	stageXML.SetAttribute("CASEID", m_sApplicationNumber);
	stageXML.SetAttribute("ACTIVITYID", sActivityID);
	stageXML.SetAttribute("ACTIVITYINSTANCE", "1");
	stageXML.RunASP(document, "MsgTMBO.asp");
	
	if(stageXML.IsResponseOK() == false)
		return (bSuccess);
	<% /* Check if this task already exists as incomplete, only add if it does 
		not exist or if has been set to complete */ %>
	var sCaseStageSeqNo = stageXML.GetTagAttribute("CASESTAGE","CASESTAGESEQUENCENO");
	var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();						
	
	var tagList = stageXML.CreateTagList("CASETASK");			
	for (var nItem = 0; nItem < tagList.length && bSuccess == false; nItem++) {
		stageXML.SelectTagListItem(nItem);			
		if(stageXML.GetAttribute("TASKID")==sTaskID) {				
			var sTaskStatus = stageXML.GetAttribute("TASKSTATUS");
			if (TempXML.IsInComboValidationList(document,"TaskStatus",sTaskStatus,["I"])) {
				//incomplete then check due date, time
				var sOldTaskDateTime = stageXML.GetAttribute("TASKDUEDATEANDTIME");
				//if ((m_sCompletionDate + " 00:00:00") >= sOldTaskDateTime) {
				if (scScreenFunctions.CompareDateStringToSystemDateTime (sOldTaskDateTime ,">=")) {
					sCaseStageSeqNo = stageXML.GetAttribute("CASESTAGESEQUENCENO");
					bTaskUpdate = true;
					bSuccess = true; 
				}
			}		
		}				
	}
	bSuccess = false;
	<% /* EP529 / MAR1547 PB */ %>
	var dueDateTime = m_sCompletionDate;
	if (dueDateTime.length <= 10)
	{
		var paramXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		dueDateTime += " " + paramXML.GetGlobalParameterString(document, "TMDlyCompTaskDueTime");
	} 
	<% /* EP529 / MAR1547 End */ %>
	
	if (bTaskUpdate == false) 
	{
		// create AdHocTask
		var reqTag = taskXML.CreateRequestTag(window, "CreateAdhocCaseTask");	
		
		taskXML.SetAttribute("USERID", sUserID);
		taskXML.SetAttribute("OWNINGUSERID", sUserID);
		taskXML.SetAttribute("UNITID", sUnitID);		
		taskXML.SetAttribute("OWNINGUNITID",sUnitID );
		taskXML.SetAttribute("USERAUTHORITYLEVEL", sUserRole);				
		taskXML.CreateActiveTag("APPLICATION");		
		taskXML.SetAttribute("APPLICATIONPRIORITY", sApplPriority);
		taskXML.ActiveTag = reqTag;	
		
		taskXML.CreateActiveTag("CASETASK");
		taskXML.SetAttribute("SOURCEAPPLICATION", tagSourceAppl);
		taskXML.SetAttribute("CASEID", m_sApplicationNumber);	
		taskXML.SetAttribute("ACTIVITYID", sActivityID);
		taskXML.SetAttribute("ACTIVITYINSTANCE", "1");		
		taskXML.SetAttribute("STAGEID",sStageID);		
		taskXML.SetAttribute("CASESTAGESEQUENCENO", sCaseStageSeqNo);		
		taskXML.SetAttribute("TASKID", sTaskID);
		<% /* PB 16/05/2006 EP529 / MAR1547
		taskXML.SetAttribute("TASKDUEDATEANDTIME", m_sCompletionDate); */ %>
		taskXML.SetAttribute("TASKDUEDATEANDTIME", dueDateTime);
		<% /* EP529 / MAR1547 End */ %>

		taskXML.RunASP(document, "OmigaTmBO.asp");
	}
	else
	{
		// Update Case Task
		reqTag = taskXML.CreateRequestTag(window, "UpdateCaseTask");
			
		taskXML.CreateActiveTag("CASETASK");
		taskXML.SetAttribute("SOURCEAPPLICATION", tagSourceAppl);
		taskXML.SetAttribute("CASEID",m_sApplicationNumber);
		taskXML.SetAttribute("ACTIVITYID", sActivityID);
		taskXML.SetAttribute("STAGEID",sStageID);
		taskXML.SetAttribute("CASESTAGESEQUENCENO", sCaseStageSeqNo);
		taskXML.SetAttribute("TASKID", sTaskID);
		taskXML.SetAttribute("TASKINSTANCE", stageXML.GetAttribute("TASKINSTANCE"));
		<% /* PB 16/05/2006 EP529 / MAR1547
		taskXML.SetAttribute("TASKDUEDATEANDTIME", m_sCompletionDate); */ %>
		taskXML.SetAttribute("TASKDUEDATEANDTIME", dueDateTime);
		<% /* EP529 / MAR1547 End */ %>
		taskXML.SetAttribute("CASETASKNAME", stageXML.GetAttribute("TASKNAME"));
		taskXML.SetAttribute("OWNINGUNITID", sUnitID);
		taskXML.SetAttribute("OWNINGUSERID", sUserID);

		taskXML.RunASP(document, "msgTMBO.asp") ;
	}				
	if(taskXML.IsResponseOK()) 	bSuccess = true; 
	stageXML = null;
	taskXML = null;
	gParamXML = null;
	TempXML = null;
	tagList = null;
	return bSuccess;
}

function UpdateReportOnTitle()
{	
	var bSuccess=false;
	var RotXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	if (m_sRotGUID == "")
		RotXML.CreateRequestTag(window,"CreateReportOnTitle");
	else
		RotXML.CreateRequestTag(window,"UpdateReportOnTitle");
		
	RotXML.CreateActiveTag("REPORTONTITLE");
	RotXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	RotXML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	if (m_sRotGUID != "")
		RotXML.SetAttribute("ROTGUID",m_sRotGUID);
	RotXML.SetAttribute("COMPLETIONDATE",m_sCompletionDate);
	RotXML.SetAttribute("TITLENUMBER",frmScreen.txtTitleNumber1.value);
	RotXML.SetAttribute("TITLENUMBER2",frmScreen.txtTitleNumber2.value);
	RotXML.SetAttribute("TITLENUMBER3",frmScreen.txtTitleNumber3.value);
	RotXML.RunASP(document,"ReportOnTitle.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = RotXML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0] == true)
		bSuccess=true;
	RotXML=null;
	return bSuccess
}

function UpdateDisbursement()
{	<% /* WP13 MARS49 */ %>
	var bSuccess=false;
	var sTemp ;
	//select node upon selected table row
	var iCurrentRow	= scScrollTable.getRowSelected();
	var sCondition	= "//PAYMENTRECORD[@AMOUNT != 0][DISBURSEMENTPAYMENT]"
	var xmlNodeList = ApplFeeTypeXML.XMLDocument.selectNodes(sCondition);
	var xmlNode		= xmlNodeList(iCurrentRow -1) ;
	ApplFeeTypeXML.ActiveTag = xmlNode ;
	
	//create request
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequest = XML.CreateRequestTag(window , 'UPDATEDISBURSEMENT');
	
	XML.CreateActiveTag('PAYMENTRECORD');
	XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	
	sTemp = ApplFeeTypeXML.GetAttribute('PAYMENTSEQUENCENUMBER');
	XML.SetAttribute('PAYMENTSEQUENCENUMBER', sTemp);
	
	XML.CreateActiveTag('DISBURSEMENTPAYMENT');
	XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	sTemp = ApplFeeTypeXML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYMENTSEQUENCENUMBER');
	XML.SetAttribute('PAYMENTSEQUENCENUMBER', sTemp);
	XML.SetAttribute('DELAYEDCOMPLETION', "1");
	XML.SetAttribute('COMPLETIONDATE', m_sCompletionDate);
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}
	if(!XML.IsResponseOK())
	{
		alert('Error in delaying payment')
		return false ;
	}
}	
	
function ResetPostCompletionTasks()
{
	var bSuccess = false;
	var bTaskUpdate	= false;
	//get params..
	var XML =			new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var gParamXML =		new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var tagSourceAppl	= "Omiga"
	var sApplPriority = scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null)
	var sActivityID =	scScreenFunctions.GetContextParameter(window,"idActivityId",null)
	
	//MAR736  Use global parameter to identify Completion stage.				
	var sStageID = gParamXML.GetGlobalParameterString(document,"TMCompletionsStageId");				
	
	XML.CreateRequestTag(window, "REQUEST")
	//DS- EP2_1920 - Run this only if Validation type "PC" returns a value
	var sTaskType = XML.GetComboIdForValidation("TaskType", "PC", null, document);
	if (sTaskType != null)
	{
		XML.SetAttribute("OPERATION","ResetPostCompletionTasks");
		XML.SetAttribute("UPDATEPROPERTY","DueDate");
		XML.CreateActiveTag("FindCaseTaskList");
		XML.CreateActiveTag("CASETASK");
		XML.SetAttribute("SOURCEAPPLICATION", tagSourceAppl);
		XML.SetAttribute("CASEID", m_sApplicationNumber);
		XML.SetAttribute("ACTIVITYID", sActivityID);
		XML.SetAttribute("ACTIVITYINSTANCE", "1");
		XML.SetAttribute("TASKSTATUS","I");			//incomplete
		//DS- EP2_1920
		//var sTaskType = XML.GetComboIdForValidation("TaskType", "PC", null, document);
		XML.SetAttribute("TASKTYPE",sTaskType);
		
		XML.SetAttribute("STAGEID", sStageID );
		XML.SetAttribute("COMPLETIONDATE", m_sCompletionDate);
		// Ensure Task Status is displayed
		XML.SetAttribute("_COMBOLOOKUP_","y");
		
		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				XML.RunASP(document,"PaymentProcessingRequest.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
		}
		
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = XML.CheckResponse(ErrorTypes);
		if(ErrorReturn[0] == false)
		{
			return false ;
		}
	}
	XML = null;
	gParamXML = null;
}
function frmScreen.btnDelayCompletion.onclick()
{
	<% /* WP13 MARS49 */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sReturn = null;
	var sUserID =		scScreenFunctions.GetContextParameter(window,"idUserID",null)
	
	var ArrayArguments = new Array();
	
	ArrayArguments[0]= XML.CreateRequestAttributeArray(window);	
	ArrayArguments[1]= m_sApplicationNumber;
	ArrayArguments[2]= m_sApplicationFactFindNumber;	  
	ArrayArguments[3]= m_sCompletionDate;
	ArrayArguments[4]= sUserID;             

	sReturn = scScreenFunctions.DisplayPopup(window, document, "PP051.asp", ArrayArguments, 400, 200);
	if (m_sCompletionDate != sReturn)
	{
	    m_sCompletionDate = sReturn;
		var iCurrentRow = scScrollTable.getRowSelected();
		var iTotRecs =scScrollTable.getTotalRecords();
		if (iTotRecs==0)
		{
			window.alert("There are no payments to delay completion.") ;
			return;
		}	
	
		if(iCurrentRow == -1)
		{
			window.alert("Select a payment for delay completion.") ;
			return ;
		}	
	// CreateDelayCompletionDateTask, UpdateReportOnTitle data, UpdateDisbursement  completion date
		LaunchAdHocTask();
		UpdateReportOnTitle();
		UpdateDisbursement();
	//Reset Due Date on Post Completion Tasks,
		ResetPostCompletionTasks();
		frmScreen.txtCompletionDate.disabled = false;
		frmScreen.txtCompletionDate.value = m_sCompletionDate;
		frmScreen.txtCompletionDate.disabled = true;
	 }	
}
<% /* PSC 25/10/2002 BMIDS00599 - Start */ %>
function DisableROFButton()
{
	function PaymentIsFullyReturned(sPaymentSeqNo, sAmount)
	{
		var xmlROFPayments;
		var xmlROFPayment;
		var sTemp;
		var iROFCount = 0;
		var lROFTotal = 0;
		var bFullyReturned = false;

		var sCondition = "//PAYMENTRECORD[@ASSOCPAYSEQNUMBER=" + sPaymentSeqNo + "]" +
					 "[DISBURSEMENTPAYMENT/@PAYMENTTYPE=" + m_sReturnOfFundsId + "]";
					 
		xmlROFPayments = ApplFeeTypeXML.XMLDocument.selectNodes(sCondition);
		lROFTotal = 0 ;
		
		<% /* Calculate total funds returned against this payment*/ %>
		for(iROFCount = 0 ; iROFCount < xmlROFPayments.length; ++iROFCount)
		{	
			xmlROFPayment = xmlROFPayments.item(iROFCount);
			sTemp = xmlROFPayment.attributes.getNamedItem("AMOUNT").text;
			if(!isNaN(sTemp)) lROFTotal = lROFTotal + parseFloat(sTemp);
		}
			
		<% /* Payment is fully returned if total returns for payment equal payment amount */ %>
		if(parseFloat(sPaymentAmount) == Math.abs(lROFTotal)) bFullyReturned = true;
		return bFullyReturned;
	}

	var sPaymentAmount = 0;
	var iCurrentRow = scScrollTable.getRowSelected();	
	var iListCount = tblTable.rows(iCurrentRow).getAttribute("ListCount");
	var sCondition = "//PAYMENTRECORD[@AMOUNT != 0][DISBURSEMENTPAYMENT]"
	var xmlNodeList = ApplFeeTypeXML.XMLDocument.selectNodes(sCondition);
	var xmlPayment = xmlNodeList(iListCount);	
	
	<% /* See if this payment has already been fully returned */ %>
	var sPaymentSeqNo = xmlPayment.attributes.getNamedItem("PAYMENTSEQUENCENUMBER").text;
	sPaymentAmount = xmlPayment.attributes.getNamedItem("AMOUNT").text;
	
	<% /* Disable if all of this payment has been returned */ %>
	return PaymentIsFullyReturned(sPaymentSeqNo, sPaymentAmount)
}
<% /* PSC 25/10/2002 B0MIDS00599 - End */ %>

-->
</script>
</BODY>
</HTML>


