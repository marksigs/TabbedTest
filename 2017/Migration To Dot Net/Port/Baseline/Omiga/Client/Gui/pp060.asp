<%@ LANGUAGE="JSCRIPT" %>
<%/*
Workfile:      pp060.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Disbursement of Payments
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		31/01/01	New Screen 
CL		05/03/01	SYS1920 Read only functionality added
MV		19/03/01	SYS2071 
MC		21/03/01	SYS2071. Default Completion Date
MC		21/03/01	SYS2119. DefaultDate function for Issue date.
MC		26/03/01	SYS2132/SYS2138
SR		23/05/01	SYS2298
SR		01/08/01	SYS2545 Include ApplicationFactFindNumber in the request 
SA		08/08/01	SYS2184 Cannot tab into CompletionDate - removed tabindex of -1
SA		10/08/01	SYS2328 Default completion indicator to 1 - Not Interfaced
SR		06/09/01	SYS2412
SR		11/12/01	SYS3377
SR		12/12/01	SYS3410
SR		13/12/01	SYS3381/SYS3382/SYS2229/SYS3382/SYS3410/
JLD		05/02/02	SYS3971 Use ThirdPartyType to set PayeeType.
JLD		19/02/02	SYS4115 set the total of fees to the total of all fees, not just those visible in the listbox.
JLD		06/03/02	SYS4177 Use NetPaymentAmount
JLD		13/03/02	SYS4114 Check for no fees returned
JLD		13/03/02	SYS4114 If there are no fees, initialise total fees as 0
JLD		19/03/02	SYS4177 changed screen about to make it clearer, added Net Amount of Advance.
MEVA	23/04/02	SYS4057 Only use the default Advance amount when no previous value has been entered
LD		23/05/02	SYS4727 Use cached versions of frame functions
DB		29/05/02	SYS4767 MSMS to Core Integration
SG		05/06/02	SYS4818 Correct error in SYS4767
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MV		10/07/2002	BMIDS00157 - Core Upgrade 
MV		06/08/2002	BMIDS00294 - Core Upgrade inline functionality - Ref AQR's SYS5019,SYS4564,SYS5129
					Amended window.onload();RetrieveContextData();SaveDisbursement();
					spnToLastField.onfocus();spnToFirstField.onfocus()
MV		11/09/2002  BMIDS00380 - After ROF , When Create new payment the paymentType available is incorrect
PSC		05/11/2002	BMIDS00760 - Omit return of funds from payment method combo
PSC		07/11/2002	BMIDS00804 - Default advance amount on change of payment type
PSC		08/11/2002	BMIDS00880 - Check sROFPayMethod for null rather than ""
PSC		11/11/2002	BMIDS00897 - Anend validateBeforeSave to return false if date format incorrect
MDC		02/01/2003	BM0224 - Change Issue Date error message to a warning. Also fix typo's in alert messages.
MV		06/02/2003	BM0218	Amended DefaultDate()
MV		19/02/2003  BM0218	Amended InitialiseScreen() and DefaultDate()
MV		25/02/2003	BM0218	Amended Windows_Onload() and InitilaiseScreen() , Added GetSystemDate(); 
MV		10/03/2003	BM0435	Amended CalculateCompletionDate()
GD		09/04/2003	BM0198	TT Fee should only be calculated at disbursement
HMA     19/05/2003  BM0017  Do not allow negative disbursement.
GD		18/08/2003	BM0198(2) Only allow ONE TT Fee to be added - check for existence of standing data.
PSC     03/09/2003  BM0476  Rewrite CalculateCompletionDate
PSC     15/09/2003  BM0198  Amend messages on date validation
HMA     15/12/2003  BMIDS683 Correct TT Fees for incentive Release.
HMA     16/12/2003  BMIDS683 Correction to XML string in IsCashBack function
HMA     18/12/2003  BMIDS687 Disable OK button during OK processing
MC		16/06/2004	BMIDS763	CODE FOR Getting TTFee changed.
SR		05/08/2004	bmids813	CODE FOR Getting TTFee changed.
JD		10/08/2004	BMIDS813	On changing payment method, check fee type before removing fees from the listbox. Also ensure that m_bTTFeeAdded flag is accurate.
HMA     13/08/2004  BMIDS808    In ValidateBeforeSave, make sure that the New Amount is calculated.
SR		31/08/2004  BMIDS815
JD		15/09/2004	BMIDS876 set window status on account refresh
GHun	16/09/2004	BMIDS876 Set window status when generating disbursements
KRW     25/10/2004  BMIDS684 Change to make Auto Waive work when Method selected before payment Type
KRW     27/10/2004  BMIDS684 Adjustment to the above to cancel TTFee when cboPaymentType.value =40
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :
HM		22/09/2005	MAR49	Global parameter PPNoTTFee is added and etc
MV		04/11/2005	MAR409	Amended SaveDisbursement(); Added GetPaymentStatusForValidationType()
MV		04/11/2005	MAR380	Created frmScreen.txtNotes.onkeyup()
HM		10/11/2005	MAR388	chkWaiveFee.checked set to true: <input id="chkWaiveFee" type="checkbox" value="1" checked=true disabled>
JD		09/02/2006	MAR1220 added DisableFieldsForInitialDisbursementOnly.
JD		17/02/2006	MAR1268 removed 'Another' button
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History :
PSC		27/02/2007	EP2_1347 Valuation Fee Refund processing
PSC		09/02/2007	EP2_1347 Don't check against total outstanding amount for valuation refund
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


*/%>
<HEAD>
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>
<title>PP060 - Disbursement of Payment</title>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>

</HEAD>

<BODY>
<% /* Scriptlets */ %>
<% /* JLD removed SYS5019
   DB SYS4767 - MSMS Integration 
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions  style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1  type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions  style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1  type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
  DB End */%>
<FORM id="frmToPP050" method="post" action="PP050.asp" STYLE="DISPLAY: none"></FORM>

<span id="spnListScroll">
	<span style="LEFT: 305px; POSITION: absolute; TOP: 410px">
<OBJECT id=scScrollTable style="WIDTH: 300px; HEIGHT: 24px" tabIndex=-1 
type=text/x-scriptlet data=scTableListScroll.asp name=scScrollTable 
VIEWASTEXT></OBJECT>
	</span> 
</span>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="LEFT: 10px; WIDTH: 609px; POSITION: absolute; TOP: 60px; HEIGHT: 445px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		&nbsp;Total Advanced to Date £
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtTotalAdvanceToDate" style="WIDTH: 100px; POSITION: absolute" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
		&nbsp;Payment type 
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<select id="cboPaymentType" style="WIDTH: 150px" class="msgCombo" menusafe="true"></select>
		</span>
	</span>
	
	<span style="LEFT: 305px; POSITION: absolute; TOP: 40px" class="msgLabel">
		Gross Amount of advance
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtAdvanceAmount" style="WIDTH: 100px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
		&nbsp;Payment method 
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<select id="cboPaymentMethod" style="WIDTH: 150px" class="msgCombo" menusafe="true"></select>
		</span>
	</span>
	<% // GD BM0198 START %>
	
	<span id="spnchkWaiveFee" style="LEFT: 4px; POSITION: absolute; TOP: 100px" class="msgLabel">
		<label for="chkWaiveFee" class="msgLabel">&nbsp;Waive TT Fee?</label>
		<span style="LEFT: 127px; POSITION: absolute; TOP: -3px" class="msgLabel">
			<input id="chkWaiveFee" type="checkbox" value="1" checked=true disabled>
		</span>
	</span>
	
	<% // GD BM0198 END %>	
	<span style="LEFT: 305px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Issue date 
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtIssueDate" style="WIDTH: 70px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 130px" class="msgLabel">
		&nbsp;Payee
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<select id="cboPayee" style="WIDTH: 150px" class="msgCombo" menusafe="true"></select>
		</span>
	</span>
		
	<span style="LEFT: 305px; POSITION: absolute; TOP: 100px" class="msgLabel">
		Advance date 
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtCompDate" style="WIDTH: 70px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
	
	<span style="LEFT: 305px; POSITION: absolute; TOP: 160px" class="msgLabel">
		Fees
		<span id="spnTable" style="LEFT: 0px; POSITION: absolute; TOP: 15px">
			<table id="tblTable" width="290" border="0" cellspacing="0" cellpadding="0" class="msgTable">
				<tr id="rowTitles">	
					<td width="60%" class="TableHead">Fee Type</td>	
					<td class="TableHead">Amount Outstanding</td>
				</tr>
				<tr id="row01">		
					<td width="60%" class="TableTopLeft">&nbsp;</td>		
					<td class="TableTopRight">&nbsp;</td>
				</tr>
				<tr id="row02">		
					<td width="60%" class="TableLeft">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row03">		
					<td width="60%" class="TableLeft">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>	
				<tr id="row04">		
					<td width="60%" class="TableLeft">&nbsp;</td>
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row05">		
					<td width="60%" class="TableLeft">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row06">		
					<td width="60%" class="TableLeft">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row07">		
					<td width="60%" class="TableLeft">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row08">		
					<td width="60%" class="TableLeft">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row09">		
					<td width="60%" class="TableLeft">&nbsp;</td>			
					<td class="TableRight">&nbsp;</td>
				</tr>
				<tr id="row10">		
					<td width="60%" class="TableBottomLeft">&nbsp;</td>	
					<td class="TableBottomRight">&nbsp;</td>
				</tr> 
			</table>
		</span>	
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 160px" class="msgLabel">
		Notes
		<span style="LEFT: 0px; POSITION: absolute; TOP: 15px"><TEXTAREA class=msgTxt id=txtNotes style="WIDTH: 280px; POSITION: absolute; HEIGHT: 170px" name=Notes rows=9></TEXTAREA>
		</span>
	</span>
	
	<span style="LEFT: 305px; POSITION: absolute; TOP: 380px" class="msgLabel" id=SPAN1>
		Total to be deducted<br>from advance
		<span style="LEFT: 180px; POSITION: absolute; TOP: 8px">
			<input id="txtTotalFeeDeductable" style="WIDTH: 70px; POSITION: absolute" class="msgReadOnly" tabindex=-1 readonly>
		</span>
	</span>
	<span style="LEFT: 305px; POSITION: absolute; TOP: 420px" class="msgLabel">
		Net amount of advance
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtNetAdvanceAmount" style="WIDTH: 100px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
</form></DIV>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* Main Buttons */ %>
<div id="msgButtons" style="LEFT: 4px; WIDTH: 612px; POSITION: relative; TOP: 500px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" -->
</div><!-- #include FILE="fw030.asp" --><!-- #include FILE="attribs/PP060attribs.asp" -->

<%/* CODE */ %>
<script language="JScript">
<!--
var scScreenFunctions ;

var m_sApplicationNumber, m_sApplicationFactFindNumber, m_sUserId ;
var m_iTotalAdvancedToDate, m_sAmountOfAdvance;
var m_sTotalFeesDeductable = 0;
var m_sMetaAction, m_sPaymentSequenceNumber ;
var m_sPaymentType, m_sPaymentMethod, m_sPayeeHistorySeqNo ;
var m_sIssueDate, m_sCompletionDate, m_sNotes, m_sUnsanctioned, m_iTotalOutstanding ;
var m_sidXML, m_sidXML2 ;
var m_iTableLength = 10 ;
var m_sFeesToBeDeductedFromLoan ;
var XMLSave ; <% /* used to call the 'Create' and 'Update' methods */ %>
var  xmlContext ; <% /* used to retrieve values from context and populate Fee Type (List box) */ %>
var m_blnReadOnly = false;
var XMLFeeType = null;
var XMLPayments = null; 
var m_TypeOfMortgage;
var m_sDeductionId, m_sReturnOfFundId, m_sBalanceCancelPaymentId ;
var m_sInitialAdvanceId, m_sInstallmentId, m_sRetentionReleaseId, m_sIncentiveReleaseId ;
<% /* PSC 27/02/2007 EP2_1347 - Start */ %>
var m_sValuationRefundId;
var m_iValuationRefundAmount = 0;
<% /* PSC 27/02/2007 EP2_1347 - End */ %>

var m_iIncentiveAmount = 0; 
var m_iInitialIncentiveAmount = 0;
var m_iRetentionAmount = 0;
var m_sShortfallPayment = "";
var m_sDate = "";
<% // GD BM0198 START %>
<%// Make global %>
var XMLCombos;
var m_bTTFeeAdded = false;<%// Flag to indicate if a TT Fee exists %>
var m_bStandingData = false;<%// Flag to indicate if the TT Fee already exists, ie. is an 'old' App %>
var m_bWaived = false;	<%// Flag to indicate if a 'Standing Data' TT Fee has been waived, ie. check box checked %>
var m_sTTFeeOneOffCostId;
var m_sTTFeeDescription;
var m_sTTFee;
var m_sFurtherAdvanceId;
var m_nDaysToSubtract = 0;
<% // GD BM0198 END %>
<% /* WP13 MAR49 */ %>
var m_bPPNoTTFee = false; //TT Fee is not to be charged when the flag is true
var m_bPPInitialDisbursementOnly = false;
<% /* Events  */ %>
function window.onload()
{	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	FW030SetTitles("Disbursement","PP060",scScreenFunctions)
	
	<% /* var sButtonList = new Array("Submit", "Cancel", "Another"); MAR1268 */ %>
	var sButtonList = new Array("Submit", "Cancel");
	ShowMainButtons(sButtonList);
	
	SetMasks() ;
	Validation_Init();
	RetrieveContextData();
	GetSystemDate();
	InitialiseScreen();
	PopulateCboPaymentType(false);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP060");
	
	<% /* SR 18/07/01 : SYS2412  */ %>
	if(m_sMetaAction == 'Edit') 
		DoEditProcessing();
	else 
		DoNonEditProcessing();
	
	<% /*SYS4564 set focus */ %>
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	SetDaysToSubtract();
	<% /* JD MAR1220 set fields to disabled */ %>
	DisableFieldsForInitialDisbursementOnly();
}

function btnSubmit.onclick()
{
	if(!validateBeforeSave()) return ;
	<% /* SR 31/08/2004 : BMIDS815 - Refresh existing account data for porting */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject(); 
	var XMLTemp = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if(! XMLTemp.IsInComboValidationList(document, 'TypeOfMortgage', m_TypeOfMortgage, 'GA'))
	{
		XML.CreateRequestTag(window, null);
		XML.SetAttribute('THROWERROR', "TRUE");
		XML.CreateActiveTag("APPLICATION");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.RunASP(document, "GetAcceptedOrActiveQuoteData.asp");

		if(! XML.IsResponseOK()) return ;
		XML.SelectTag(null,"MORTGAGESUBQUOTE");
		var sMSQNumber = XML.GetTagText("MORTGAGESUBQUOTENUMBER");
		
		XML.ResetXMLDocument();
		XML.CreateRequestTag(window, null);
		var xmlMainTag = XML.CreateActiveTag("GETANDSAVEPORTEDSTEPANDPERIODFROMMORTGAGEACCOUNT");
		XML.CreateTag("PORTINGINDICATOR", '1');
		XML.CreateActiveTag("APPLICATION");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("TYPEOFAPPLICATION", scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue",null));
		XML.CreateTag("BMACCOUNTNUMBER", scScreenFunctions.GetContextParameter(window,"idOtherSystemAccountNumber",null));
		
		XML.ActiveTag = xmlMainTag;
		AddOtherSystemCustomerNumbers(XML);
		
		XML.ActiveTag = xmlMainTag ;
		XML.CreateActiveTag("MORTGAGESUBQUOTE");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("MORTGAGESUBQUOTENUMBER", sMSQNumber);	
		
		<% /* JD BMIDS876 set window status as account refresh can take a while. */ %>
		window.status = "Refreshing account data ...";		
		XML.RunASP(document, "GetAndSavePortedStepAndPeriodFromMortAcc.asp");
		window.status = "";
		if(! XML.IsResponseOK()) return ;
	}
	<% /* SR 31/08/2004 : BMIDS815 - End */ %>
	
	<% /* BMIDS687  Disable OK button while processing takes place */ %>
	DisableMainButton("Submit")
	
	if(m_sMetaAction == "Add" || m_sMetaAction == "Copy")
	{
		if(SaveDisbursement(true)) frmToPP050.submit();
	}
	else if(SaveDisbursement(false)) frmToPP050.submit();
}

function btnCancel.onclick()
{
	frmToPP050.submit();
}

<% /* JD MAR1268 remove Another button
function btnAnother.onclick()
{
	var sCurrentAdvance, sCurrentFeePayment ;
	
	sCurrentAdvance		= frmScreen.txtAdvanceAmount.value;
	sCurrentFeePayment  = frmScreen.txtTotalFeeDeductable.value ;
	
	//  Create disbursement record. Clear all values and assign new value to Total Advanced to date 
	if(validateBeforeSave())
	{
		if(SaveDisbursement(true))
		{
			// SR 12-12-2001 : SYS2229 
			switch(frmScreen.cboPaymentType.value)
			{
				case m_sInitialAdvanceId:
				case m_sInstallmentId: 
					m_iTotalOutstanding -= parseFloatSafe(sCurrentAdvance);	
					m_iTotalAdvancedToDate = m_iTotalAdvancedToDate + parseFloatSafe(sCurrentAdvance) + parseFloatSafe(sCurrentFeePayment);
					break ;
				case m_sIncentiveReleaseId:
					m_iIncentiveAmount -= parseFloatSafe(sCurrentAdvance);
					break ;
				case m_sRetentionReleaseId:
					m_iRetentionAmount -= parseFloatSafe(sCurrentAdvance);
					m_iTotalAdvancedToDate += parseFloatSafe(sCurrentAdvance);
					break ;
			}
			ClearScreen() ;
			frmScreen.txtTotalAdvanceToDate.value = m_iTotalAdvancedToDate ;	
			PopulateCboPaymentType(true);
		}
	}
}
***/ %>

function frmScreen.cboPaymentMethod.onchange()
{
	<% // WP13 PAR49 START %>
	//DefaultDate();
	if(frmScreen.txtIssueDate.value == "" && frmScreen.cboPaymentMethod.value != "" && frmScreen.cboPaymentType.value != "")
		frmScreen.txtIssueDate.value =  m_sDate;

	<% // WP13 PAR49 END %>
	
	<% // GD BM0198 START %>
	EnableDisableCheckBox();
	<% // GD BM0198 END %>
}
<% // GD BM0198 START %>
function EnableDisableCheckBox()
{
	if (m_bPPNoTTFee == false) // MAR49:  m_bPPNoTTFee is added
	{
		if ( IsTT(frmScreen.cboPaymentMethod.value) == true) 
		{
			frmScreen.chkWaiveFee.disabled = false;
		
			<% /* BMIDS683 Make sure the chkWaiveFee is checked BEFORE the TT Fee is calculated */ %>
			<% /* PSC 27/02/2007 EP2_1347 */ %>
			if (IsCashBack(frmScreen.cboPaymentType.value) || frmScreen.cboPaymentType.value == m_sValuationRefundId)
			{
				frmScreen.chkWaiveFee.checked = true;	// KRW     27/10/2004  BMIDS684
					CancelTTFee();
					m_bTTFeeAdded = false;	
			
			}
			if (m_bTTFeeAdded == false)<%//Don't allow a TT Fee to be added twice%>
			{
				m_bTTFeeAdded = CalculateTTFee();
				// JD BMIDS813 m_bTTFeeAdded = true;
			}		
		} 
		else
		{
			<%//Standing Data and Waived%>
			if ((m_bWaived) && (m_bStandingData))
			{
				frmScreen.txtTotalFeeDeductable.value = parseFloatSafe(frmScreen.txtTotalFeeDeductable.value) + parseFloatSafe(m_sTTFee)
				frmScreen.txtNetAdvanceAmount.value = parseFloatSafe(frmScreen.txtNetAdvanceAmount.value) - parseFloatSafe(m_sTTFee);
				m_bWaived = false;
			}
			if(m_bTTFeeAdded)
			{	
				//Only remove the TT Fee if is not a 'Standing Data' TT Fee
				if (!(m_bStandingData))
				{
					CancelTTFee();
					m_bTTFeeAdded = false;			
				
				}
				<%//Uncheck the 'Waive' checkbox%>
				frmScreen.chkWaiveFee.disabled = false;//in case not already enabled..
				frmScreen.chkWaiveFee.checked = false;
				frmScreen.chkWaiveFee.disabled = true;			
			}
		}
	}
	<% 	/*  MAR388 chkWaiveFee.checked set to true and disabled already
	else 
		frmScreen.chkWaiveFee.disabled = true;	//MAR49
	 */ %>
}

function CalculateTTFee()
{
<%	/* JD BMIDS813 Return TRUE if the TTFee is > 0. Return FALSE if the TTFee is 0 or ''
	   so that it is not created and saved */ %>
	var sDefaultLenderCode = "";
	var bReturn = true;

	var globalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	sDefaultLenderCode = globalXML.GetGlobalParameterString(document, "DefaultLenderCode");
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if (sDefaultLenderCode !="")
	{
		XML.CreateRequestTag(window, "GETTTFEEAMOUNT");
		XML.CreateActiveTag("GETTTFEEAMOUNT");
		<% /* GD BM0198 make element basedXML.SetAttribute("LENDERCODE",sDefaultLenderCode); */ %>
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.RunASP(document,"GetTTFeeAmount.asp");
		if(XML.IsResponseOK())
		{
			try
			{
				XML.ActiveTag = null;
				XML.SelectTag(null,"TTFEESET");
				m_sTTFee = XML.GetTagText("AMOUNT");
			}
			catch(e)
			{
				alert(e);
			}
		} else
		{
			alert("A problem was encountered trying to Get Lender Details for Lender Code " + sDefaultLenderCode);
			return false;
		}	
	} else
	{
		alert("A problem was encountered try to find the Global Parameter 'DefaultLenderCode'");
		return false;
	}
	//Now display the new fee and reset the derivable fields.

<%	//*****ADD LINE TO FEE LIST*****
	//Append to XMLFeeType
%>
<%	/* SR 05/08/2004 : BMIDS813 - Add TT Fee to list only if it is > 0   */  %>	
	if(m_sTTFee != null)
	{
		if (m_sTTFee != '0' && m_sTTFee!= '')
		{
			var xmlResponse = XMLFeeType.XMLDocument.selectSingleNode(".//RESPONSE")
			var xmlNew = XMLFeeType.XMLDocument.createElement("APPLICATIONFEETYPE");
			xmlResponse.appendChild(xmlNew);
			xmlNew.setAttribute("FEETYPE",m_sTTFeeOneOffCostId);
			xmlNew.setAttribute("FEETYPE_TEXT",m_sTTFeeDescription);
			xmlNew.setAttribute("AMOUNTOUTSTANDING",m_sTTFee);	
			xmlNew.setAttribute("AMOUNTPAID",m_sTTFee);	
		}
		else
			bReturn = false;
	}
	else
		bReturn = false;
<%	/* SR 05/08/2004 : BMIDS813 - End  */  %>	
	//Reinitialise Total to be Deducted from Advance and Net Amount of Advance.
	XMLFeeType.ActiveTag = null;
	XMLFeeType.CreateTagList("APPLICATIONFEETYPE");

	frmScreen.txtTotalFeeDeductable.value = parseFloatSafe(frmScreen.txtTotalFeeDeductable.value) + parseFloatSafe(m_sTTFee);
	frmScreen.txtNetAdvanceAmount.value = parseFloatSafe(frmScreen.txtNetAdvanceAmount.value) - parseFloatSafe(m_sTTFee);
	PopulateListbox(0);
	//BMIDS Specific  - IF Payment Type = Cashback then auto waive the TT Fee.
	<% /* PSC 27/02/2007 EP2_1347 */ %>
	if(frmScreen.cboPaymentType.value == m_sIncentiveReleaseId || frmScreen.cboPaymentType.value == m_sValuationRefundId)
	{
		WaiveFee();
	}
	return bReturn;
}

function CancelTTFee()
{
	<%//Description
	//If the payment method has changed from TT to another method then we need to cancel the application fee and reinitialise the derived fields on screen if the Fee has been waived.

	//Implementation
	////Delete the newly added ApplicationCost from the display screen.
	//Delete ApplicationCost where FeeType = TT Fee
	//*****DELETE FROM TABLE AND XML*****
	//  If TT fee has been added in remove it from the totals.
	//  If this method gets called, then we will have added the TTFee, and it will be the last item..
	//..in XMLFeeType..so remove it.
	// JD BMIDS813 Cannot assume that the TTFee is the last one added (it may have been 0)
	%>
	XMLFeeType.ActiveTag = null;
	XMLFeeType.CreateTagList("APPLICATIONFEETYPE");
	var iLength = 	XMLFeeType.ActiveTagList.length;
	<%//iLength should be at least 1, but check anyway%>
	<%/*if(iLength > 0)
	{
		XMLFeeType.SelectTagListItem(iLength-1);
		XMLFeeType.RemoveActiveTag();
		XMLFeeType.ActiveTag = null;
		XMLFeeType.CreateTagList("APPLICATIONFEETYPE");
		PopulateListbox(0);
	}*/%>
	var bTTFeeFound = false
	for (var TagItem = 0; TagItem < iLength && bTTFeeFound == false; TagItem++)
	{
		XMLFeeType.SelectTagListItem(TagItem);
		if(XMLFeeType.GetAttribute("FEETYPE") == m_sTTFeeOneOffCostId)
		{
			XMLFeeType.RemoveActiveTag();
			bTTFeeFound = true
		}
	}
	if(bTTFeeFound)
	{
		XMLFeeType.ActiveTag = null;
		XMLFeeType.CreateTagList("APPLICATIONFEETYPE");
		PopulateListbox(0);
		
		if (frmScreen.chkWaiveFee.checked == false)
		{
			<%//Reinitialise Total to be Deducted from Advance and Net Amount of Advance.
			//Total to be Deducted from Advance - TTFEE%>
			if (m_bTTFeeAdded)
			{
				//Don't adjust txtTotalFeeDeductable, for non standing data as its done in PopulateListBox() 
				frmScreen.txtNetAdvanceAmount.value = parseFloatSafe(frmScreen.txtNetAdvanceAmount.value) + parseFloatSafe(m_sTTFee);
			}
		}
	}
	XMLFeeType.ActiveTag = null; // reset
}
function WaiveFee()
{
	<%//Description
	//The onclick event of the check box will only fire when it is enabled,
	//So we will only end up in here when a TT Fee has been selected.
	//When this checkbox is checked the application fee ‘TT fee‘ will not be deducted from the loan

	//Implementation
	//IF CheckBox = True (checked) THEN %>
	if (frmScreen.chkWaiveFee.checked == true)
	{
		<%//Reinitialise Total to be Deducted from Advance and Net Amount of Advance.
		//Total to be Deducted from Advance - TTFEE%>
		frmScreen.txtTotalFeeDeductable.value = parseFloatSafe(frmScreen.txtTotalFeeDeductable.value) - parseFloatSafe(m_sTTFee);
		//Net Amount of Advance + TTFEE
		frmScreen.txtNetAdvanceAmount.value = parseFloatSafe(frmScreen.txtNetAdvanceAmount.value) + parseFloatSafe(m_sTTFee);
		m_bWaived = true;

	} else
	{
		<%//Checkbox is reselected to unchecked%>
		if (frmScreen.chkWaiveFee.checked == false)
		{
			<%//Reinitialise Total to be Deducted from Advance and Net Amount of Advance.
			//Total to be Deducted from Advance + TTFEE %>
			frmScreen.txtTotalFeeDeductable.value = parseFloatSafe(frmScreen.txtTotalFeeDeductable.value) + parseFloatSafe(m_sTTFee);
			<%//Net Amount of Advance - TTFEE%>
			frmScreen.txtNetAdvanceAmount.value = parseFloatSafe(frmScreen.txtNetAdvanceAmount.value) - parseFloatSafe(m_sTTFee);
			m_bWaived = false;
		}
		
	}
	
}

function frmScreen.chkWaiveFee.onclick()
{
		WaiveFee();
}

function IsTT(sPaymentMethod)
{
	var blnResult = false;
	var iIndex;
	var sPattern = ".//LISTENTRY[GROUPNAME = 'PaymentMethod' $and$ VALUEID = '" + sPaymentMethod + "']/VALIDATIONTYPELIST"
	var xmlPaymentList = XMLCombos.selectNodes(sPattern)
	var xmlPayment;
	for(iIndex=0 ; iIndex < xmlPaymentList.length ; iIndex++)
	{
		xmlPayment = xmlPaymentList.item(iIndex)
		if (xmlPayment.text == 'YC')
		{
			return(true)
		}	
	}

	return(false);
}
function IsCashBack(sPaymentType)
{
	var blnResult = false;
	var iIndex;
	<% /* BMIDS683 Correct pattern string */ %>
	var sPattern = ".//LISTENTRY[GROUPNAME = 'PaymentType' $and$ VALUEID = '" + sPaymentType + "']/VALIDATIONTYPELIST/VALIDATIONTYPE"
	var xmlPaymentList = XMLCombos.selectNodes(sPattern)
	var xmlPayment;
	for(iIndex=0 ; iIndex < xmlPaymentList.length ; iIndex++)
	{
		xmlPayment = xmlPaymentList.item(iIndex)
		if (xmlPayment.text == 'C')
		{
			return(true)
		}	
	}

	return(false);
}
<% // GD BM0198 END %>
function frmScreen.cboPaymentType.onchange()
{
	DefaultDate();
	<%/* SR   11-12-02 : SYS3377 Default the Advance amount to appropriate value when the combo payment type is changed */%>	
	<%/* MEVA 23-12-02 : SYS4057 but only when an Advance Amount has not already been entered                           */%>		
	<%/* PSC  07/11/02 : BMIS00804 Default advance amount regardless */ %>		
	<% /* PSC 27/02/2007 EP2_1347 */ %>
	scScreenFunctions.SetFieldToWritable(frmScreen, "txtAdvanceAmount");
	
	switch (frmScreen.cboPaymentType.value)
	{
		case m_sIncentiveReleaseId :
			frmScreen.txtAdvanceAmount.value = m_iIncentiveAmount ;
			break ;
		case m_sRetentionReleaseId :
			frmScreen.txtAdvanceAmount.value = m_iRetentionAmount ;
			break ;
		<% /* PSC 27/02/2007 EP2_1347 - Start */ %>
		case m_sValuationRefundId:
			frmScreen.txtAdvanceAmount.value = m_iValuationRefundAmount;
			scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAdvanceAmount");
			break;
		<% /* PSC 27/02/2007 EP2_1347 - End */ %>
		default :
			frmScreen.txtAdvanceAmount.value = m_iTotalOutstanding ;
	}
	EnableDisableCheckBox(); // BMIDS684 KRW 25/10/04
	frmScreen.txtNetAdvanceAmount.value = parseFloatSafe(frmScreen.txtAdvanceAmount.value) - parseFloatSafe(frmScreen.txtTotalFeeDeductable.value);
	
}

<% /* Functions */  %>
function DisableFieldsForInitialDisbursementOnly()
{
	if(m_bPPInitialDisbursementOnly)
	{
		//set payment type as initial advance and disable
		frmScreen.cboPaymentType.value = m_sInitialAdvanceId;
		frmScreen.cboPaymentType.onchange();
		frmScreen.cboPaymentType.disabled = true;
		
		//set to first payee. Disable if there is only one payee
		if(frmScreen.cboPayee.options.length > 1)
		{
			frmScreen.cboPayee.selectedIndex = 1;
			if (frmScreen.cboPayee.options.length == 2)
				frmScreen.cboPayee.disabled = true;
		}
		frmScreen.txtAdvanceAmount.disabled = true;
		frmScreen.txtIssueDate.disabled = true;
		frmScreen.txtCompDate.disabled = true;
		frmScreen.cboPaymentMethod.disabled = true;
	}
}
function RetrieveContextData()
{
	<%/* WP13 MAR49*/%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bPPNoTTFee = XML.GetGlobalParameterBoolean(document, "PPNoTTFee");
	
	<% /* JD MAR1220 Get Initial disbursement global*/ %>
	m_bPPInitialDisbursementOnly = XML.GetGlobalParameterBoolean(document, "PPInitialDisbursementOnly");
	
	XML= null;
	
	m_TypeOfMortgage= scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","");
	m_sidXML		= scScreenFunctions.GetContextParameter(window,"idXML","");	
	m_sidXML2		= scScreenFunctions.GetContextParameter(window,"idXML2","");	
	
	//DB SYS4767 - MSMS Integration
	//SG 05/04/02 MSMS009
	m_sShortfallPayment = scScreenFunctions.GetContextParameter(window,"idShortfallPayment","");
	//DB End
	
	<%/* SG 24/07/02 SYS5219 */%>
	if (m_sShortfallPayment=="")
		m_sShortfallPayment = "0";
		
	if(m_sidXML != "")
	{
		<%/* SG 05/06/02 SYS4818 */%>	
		<%/* DB SYS4767 - MSMS Integration */%>
		xmlContext = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		<%/* xmlContext = new scXMLFunctions.XMLObject(); */%>
		<%/* DB End */%>	
		<%/* SG 05/06/02 SYS4818 */%>

		xmlContext.LoadXML(m_sidXML);
		
		xmlContext.CreateTagList("CALCULATIONS");
		xmlContext.SelectTagListItem(0);

		<%/* SG 24/07/02 SYS5129 Start */%>
		m_iTotalOutstanding = parseFloatSafe(xmlContext.GetAttribute("BALANCE")) + parseFloatSafe(m_sShortfallPayment)
		<%/* SG 24/07/02 SYS5129 End */%>
		
		m_iIncentiveAmount = xmlContext.GetAttribute("MORTGAGEINCENTIVE") ;
		if(m_iIncentiveAmount != "") m_iIncentiveAmount = parseFloatSafe(m_iIncentiveAmount) ;
		else m_iIncentiveAmount = 0 ;
		
		<%/* SR 11-12-02 : SYS3377 Default the Advance amount to appropriate value when the combo payment type is changed */%>
		m_iRetentionAmount = xmlContext.GetAttribute("RETENTIONAMOUNT") ;
		if(m_iRetentionAmount != "") m_iRetentionAmount = parseFloatSafe(m_iRetentionAmount);
		else m_iRetentionAmount = 0;
		<%/* SR 11-12-02 : SYS3377  END */ %>
		
		<% /* PSC 27/02/2007 EP2_1347 - Start */ %>
		m_iValuationRefundAmount = xmlContext.GetAttribute("OUTSTANDINGVALUATIONREFUND") ;
		if(m_iValuationRefundAmount != "") m_iValuationRefundAmount = parseFloatSafe(m_iValuationRefundAmount);
		else m_iValuationRefundAmount = 0;
		<% /* PSC 27/02/2007 EP2_1347 - End */%>

		<% /* SR 11-12-01 : SYS3372  */ %>
		var iOutStandingFee = xmlContext.GetAttribute("FEESTOBEDEDUCTEDFROMADVANCE");
		if(iOutStandingFee != "") iOutStandingFee = parseFloatSafe(iOutStandingFee);
		else iOutStandingFee = 0;
				
		m_iTotalAdvancedToDate =  parseFloatSafe(xmlContext.GetAttribute("AMOUNTREQUESTED")) 
								 - parseFloatSafe(xmlContext.GetAttribute("BALANCE"))
								 //DB SYS4767 - MSMS Integration
								 /*  JLD - iOutStandingFee */ - m_iRetentionAmount ;
								 //DB End
		
		<% /* SR 11-12-01 : SYS3372 END */ %>
		<% /* W14 MAR49  */ %>
		m_sCompletionDate = xmlContext.GetAttribute("COMPLETIONDATE");
		
		xmlContext.ActiveTag = null ;
		xmlContext.CreateTagList("PAYMENTRECORD");
		if(xmlContext.SelectTagListItem(0))
		{
			m_sPaymentSequenceNumber	= xmlContext.GetAttribute("PAYMENTSEQUENCENUMBER");
			m_sAmountOfAdvance			= xmlContext.GetAttribute("AMOUNT");
			m_sPaymentType				= xmlContext.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYMENTTYPE');
			m_sPaymentMethod			= xmlContext.GetAttribute('PAYMENTMETHOD');
			m_sPayeeHistorySeqNo		= xmlContext.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYEEHISTORYSEQNO');
			m_sIssueDate				= xmlContext.GetTagAttribute('DISBURSEMENTPAYMENT', 'ISSUEDATE');
			m_sCompletionDate			= xmlContext.GetTagAttribute('DISBURSEMENTPAYMENT', 'COMPLETIONDATE');
			m_sNotes					= xmlContext.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYMENTNOTES');
		}	
	}
	else
	{
		m_iTotalAdvancedToDate = 0;
		m_sAmountOfAdvance = "";
		m_sPaymentSequenceNumber = "";
		m_sPaymentType = "" ;
		m_sPaymentMethod = "" ;
		m_sPayeeHistorySeqNo = "" ;
		m_sIssueDate = "" ;
		m_sCompletionDate = "" ;
		m_sNotes = "" ;
	}
	
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
	<% /* SR 12-12-2001 : SYS3410 */ %>
	m_sUserId = scScreenFunctions.GetContextParameter(window, "idUserId", "");
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", "");


	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	XMLPayments = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* XMLPayments = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>

	XMLPayments.LoadXML(m_sidXML2);
}	

function InitialiseScreen()
{	
	<% /* Populate combos from table 'ComboValue' */ %>
	XMLCombos = null;

	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLFeeType = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var XML = new scXMLFunctions.XMLObject(); */%>
	<%/* XMLFeeType = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>

	var iCount ;
	
	var blnSuccess = true;		
	<%//GD BM0198 Added OneOffCost, TypeOfMortgage %>
	var sGroupList = new Array("PaymentType", "PaymentMethod", "PaymentEvent", "OneOffCost", "TypeOfMortgage");
	if(XML.GetComboLists(document,sGroupList))
	{		
		<%//GD BM0198 19/08/2003 START %>
		XMLCombos = XML.GetComboListXML("TypeOfMortgage");
		m_sFurtherAdvanceId = XML.GetComboIdForValidation("PaymentType", "FA", XMLCombos);
		<%//GD BM0198 19/08/2003 END %>
		XMLCombos = XML.GetComboListXML("PaymentType");
		m_sInitialAdvanceId			= XML.GetComboIdForValidation("PaymentType", "I", XMLCombos) ;
		m_sInstallmentId			= XML.GetComboIdForValidation("PaymentType", "A", XMLCombos) ;
		m_sRetentionReleaseId		= XML.GetComboIdForValidation("PaymentType", "R", XMLCombos) ;
		m_sIncentiveReleaseId		= XML.GetComboIdForValidation("PaymentType", "C", XMLCombos) ;
		m_sReturnOfFundId			= XML.GetComboIdForValidation("PaymentType", "N", XMLCombos) ;
		m_sBalanceCancelPaymentId	= XML.GetComboIdForValidation("PaymentType", "NCB", XMLCombos) ;
		<% /* PSC 27/02/2007 EP2_1347 */ %>
		m_sValuationRefundId = XML.GetComboIdForValidation("PaymentType", "VALREFUND", XMLCombos) ;

		XMLCombos = XML.GetComboListXML("PaymentMethod");
		
		<% /* PSC 05/11/2002 BMIDS00760 - Start */ %>
		var sROFPayMethod = XML.GetComboIdForValidation("PaymentMethod", "R", XMLCombos); 
			
		<% /* PSC 08/11/2002 BMIDS00880 */ %>	
		if (sROFPayMethod != null)
		{
			var sCondition = ".//LISTENTRY[VALUEID='" + sROFPayMethod + "']" ;
			var xmlNode = XMLCombos.selectSingleNode(sCondition) ;
			XMLCombos.firstChild.removeChild(xmlNode) ;
		}
		<% /* PSC 05/11/2002 BMIDS00760 - End */ %>

		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document, 
								frmScreen.cboPaymentMethod,XMLCombos,true);	
								
		<%/* SR 10/05/01 : SYS2298 - Add Payment Event to the request for the method 'SaveDisbursement' */%>
		XMLCombos = XML.GetComboListXML("PaymentEvent");
		m_sDeductionId =  XML.GetComboIdForValidation("PaymentEvent", 'D', XMLCombos) ;
		<% // GD BM0198 START %>
		XMLCombos = XML.GetComboListXML("OneOffCost");
		m_sTTFeeOneOffCostId = XML.GetComboIdForValidation("OneOffCost", 'TTF', XMLCombos);
		m_sTTFeeDescription = XML.GetComboDescriptionForValidation("OneOffCost", "TTF")
		//reset xmlCombos
		XMLCombos = XML.SelectSingleNode("RESPONSE");
		<% // GD BM0198 END %>
	}

	<% /* Populate Payee combo  */ %>
	XML.ResetXMLDocument();
	XML.CreateRequestTag(window, "FindPayeeHistoryList");
	XML.CreateActiveTag("PAYEEHISTORY");
	XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.SetAttribute("_COMBOLOOKUP_","1");
	XML.RunASP(document,"PaymentProcessingRequest.asp");
	
	if(!XML.IsResponseOK())
	{
		alert('Error retreiving Payee details');
		return ;
	}
		
	var TagSELECT = document.createElement("OPTION");
	TagSELECT.text = "<SELECT>";
	TagSELECT.value = ""	
	frmScreen.cboPayee.options.add(TagSELECT);

	var TagOPTION, nLoop ;
	var TagListLISTENTRY = XML.CreateTagList("PAYEEHISTORY");
	for(nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
	{
		XML.SelectTagListItem(nLoop);
		
		TagOPTION = document.createElement("OPTION");
		TagOPTION.value = XML.GetAttribute("PAYEEHISTORYSEQNO");
		TagOPTION.text = XML.GetTagAttribute("THIRDPARTY", "COMPANYNAME");
		<% // Save PayeeType as attribute 
		%>
		TagOPTION.setAttribute("PayeeType", XML.GetTagAttribute("THIRDPARTY", "THIRDPARTYTYPE"));  // JLD use thirdPartyType not PayeeType
		frmScreen.cboPayee.options.add(TagOPTION);
	}
	
	<% /* Loading FeetTypes and Outstanding Amounts */ %>
	if(m_sidXML != '')
	{
		XMLFeeType.CreateRequestTag(window, "FindFeePaymentTotals");
		XMLFeeType.CreateActiveTag("APPLICATIONFEETYPE");
		XMLFeeType.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
		XMLFeeType.SetAttribute("_COMBOLOOKUP_","1");
		XMLFeeType.RunASP(document,"PaymentProcessingRequest.asp");
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = XMLFeeType.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{
			alert("No fees were found for this application.");
			frmScreen.txtTotalFeeDeductable.value = "0" ;
		}
		else if(ErrorReturn[0] == true)
		{
			XMLFeeType.ActiveTag= null;
			XMLFeeType.CreateTagList("APPLICATIONFEETYPE");
			<% /*----------- 18/08/2003 GD BM0198(v2) START Check for the existence of a TT Fee already---------*/ %>
			var sFeeType;
			var nLoop;
			var ValidationList = new Array(1);	
			var sAmountOutStanding;
			ValidationList[0] = "TTF";
			var XMLfee = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			<% /* GD BM0198(v2) END */ %>
	
			<% /* Loop through the XML (Fee Payments) and check */ 	%>
			for(var nLoop = 0 ; nLoop < XMLFeeType.ActiveTagList.length && nLoop<m_iTableLength; nLoop++)
			{
				XMLFeeType.SelectTagListItem(nLoop);
				sFeeType = XMLFeeType.GetAttribute("FEETYPE");
				sAmountOutStanding = XMLFeeType.GetAttribute("AMOUNTOUTSTANDING");
				//Check if the FEETYPE has the validation type of TTF from combo group OneOffCost
				if (m_bTTFeeAdded == false)
				{
					if( XMLfee.IsInComboValidationList(document,"OneOffCost", sFeeType, ValidationList))
					{
						m_sTTFee = sAmountOutStanding;
						if(m_sTTFee != null && m_sTTFee != '0' && m_sTTFee!= '') // JD BMIDS813
						{
							m_bTTFeeAdded = true;
							m_bStandingData = true;	
						}
					}
				}

			}
			<% /*-------------- 18/08/2003 GD BM0198(v2) END ---------------------------------------------------*/ %>			
			PopulateListbox(0);
		}
		else
		{
			alert('Error retrieving FeeType details');
			return ;
		}
	}
	
	<% // In update mode, if Fee type list was filled with some thing, disable 'OK' and 'Amount of Advance'
	%>
	if(m_sMetaAction == "Edit" && (frmScreen.txtTotalFeeDeductable.value == "" || frmScreen.txtTotalFeeDeductable.value != "0"))
	{
		scScreenFunctions.SetFieldToDisabled(frmScreen, "txtAdvanceAmount");
		DisableMainButton('Submit');
	}
	if(m_sMetaAction == "Edit" ||m_sMetaAction == "Copy") 
	{
		PopulateOldValues();
		<% /* 
		GD 19/08/2003 BM0198 Remove hardcoded value ids START
		if ( (m_sPaymentType == "50") || (m_sPaymentType == "99")) 
			scScreenFunctions.SetFieldState(frmScreen,"cboPaymentType","R");
		*/ %>
		if ( (m_sPaymentType == m_sReturnOfFundId) || (m_sPaymentType == m_sBalanceCancelPaymentId)) 
			scScreenFunctions.SetFieldState(frmScreen,"cboPaymentType","R");		
		
		<% /*
		GD 19/08/2003 BM0198 Remove hardcoded value ids END
		*/ %> 
	}
	else
	{
		<% /* WP13 MAR49 Blocked below
	
		//<% // GD 19/08/2003 BM0198 Remove hardcoded value ids %>
		//if(m_TypeOfMortgage != m_sFurtherAdvanceId)
		var tempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

		if(!(tempXML.IsInComboValidationList(document,"TypeOfMortgage", m_TypeOfMortgage, ["F", "M"])))
		{
			<%/*  Use Report on Title date %>
			if (GetCompletionDate() == "") 
				frmScreen.txtCompDate.value = m_sDate ;
			else
				frmScreen.txtCompDate.value = GetCompletionDate();
		}
		else if(ActingLegalRep())
		{
			<% /* Further AdvanceIf Legal Rep is acting in this case...
			 Use Report on Title date  %>
			if (GetCompletionDate() == "") 
				frmScreen.txtCompDate.value = m_sDate ;
			else
				frmScreen.txtCompDate.value = GetCompletionDate();
		}
		else
		{
			frmScreen.txtCompDate.value = m_sDate;
		} */ %>
		
		<% /* WP13 MAR49 Set the Payment Method and Advanced date */%>
		if (!SetPaymentMethod())
		{	
			alert("An error occurs while setting a payment method.");
			frmToPP050.submit();
			return;
		}
		else
		{	
			if (!SetAdvanceDate())
			{
				alert("An error occurs while setting an advanced date.");
				frmToPP050.submit();
				return;
			}
		}
	}
	frmScreen.txtIssueDate.value = m_sDate; //WP13 MAR49
	frmScreen.txtTotalAdvanceToDate.value = m_iTotalAdvancedToDate ;
	scScreenFunctions.SetFieldState(frmScreen, "txtNetAdvanceAmount", "R");
}

function PopulateCboPaymentType(bForAnotherButton)
{
	var XMLCombos = null;

	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var XML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>

	var sGroupList = new Array("PaymentType");
	if(XML.GetComboLists(document,sGroupList))
	{
		XMLCombos = XML.GetComboListXML("PaymentType");
		<%/* SR 17/0701 : SYS2412 -do not display the values 'ReturnOfFunds', 'BalanceCancellationPayment'  */%>
		
		var sCondition = ".//LISTENTRY[VALUEID='" + m_sReturnOfFundId + "']" ;
		var xmlNode = XMLCombos.selectSingleNode(sCondition) ;
		XMLCombos.firstChild.removeChild(xmlNode) ;
		
		<% /* SR 11-12-01 : SYS3382 & SYS3382 - do not populate Incentive Release /Retention Release if the respective amounts are zero */ %>
		if(m_iIncentiveAmount == 0)
		{
			sCondition = ".//LISTENTRY[VALUEID='" + m_sIncentiveReleaseId + "']" ;
			xmlNode = XMLCombos.selectSingleNode(sCondition) ;
			XMLCombos.firstChild.removeChild(xmlNode) ;
		}
		
		if(m_iRetentionAmount == 0 || m_iTotalOutstanding > 0)
		{
			sCondition = ".//LISTENTRY[VALUEID='" + m_sRetentionReleaseId + "']" ;
			xmlNode = XMLCombos.selectSingleNode(sCondition) ;
			XMLCombos.firstChild.removeChild(xmlNode) ;
		}
		
		<% /* PSC 27/02/2007 EP2_1347 - Start */ %>
		if (m_iValuationRefundAmount == 0)
		{
			sCondition = ".//LISTENTRY[VALUEID='" + m_sValuationRefundId + "']" ;
			xmlNode = XMLCombos.selectSingleNode(sCondition) ;
			XMLCombos.firstChild.removeChild(xmlNode) ;
		}
		<% /* PSC 27/02/2007 EP2_1347 - End */ %>

		<%/*if balance=0, then do not display InitialAdvance/Installment */%>
		if(m_iTotalOutstanding == 0)
		{
			sCondition = ".//LISTENTRY[VALUEID='" + m_sInitialAdvanceId + "']" ;
			xmlNode = XMLCombos.selectSingleNode(sCondition) ;
			XMLCombos.firstChild.removeChild(xmlNode) ;
			
			sCondition = ".//LISTENTRY[VALUEID='" + m_sInstallmentId + "']" ;
			xmlNode = XMLCombos.selectSingleNode(sCondition) ;
			XMLCombos.firstChild.removeChild(xmlNode) ;
		}
		else
		{
			if(bForAnotherButton)
			{
				sCondition = ".//LISTENTRY[VALUEID='" + m_sInitialAdvanceId + "']" ;
				xmlNode = XMLCombos.selectSingleNode(sCondition) ;
				XMLCombos.firstChild.removeChild(xmlNode) ;
			}
		}
		<% /* SR 11-12-01 : SYS3382 END */ %>
		
		<% /* SR 11-12-01 : SYS3381 - do not populate Balance Cancellation Payment */ %>
		sCondition = ".//LISTENTRY[VALUEID='" + m_sBalanceCancelPaymentId + "']" ;
		xmlNode = XMLCombos.selectSingleNode(sCondition) ;
		XMLCombos.firstChild.removeChild(xmlNode) ;		
		<% /* SR 11-12-01 : SYS3381 END */ %>
		
		XML.PopulateComboFromXML(document, frmScreen.cboPaymentType,XMLCombos,true);
	}
}

function ActingLegalRep()
{
	<% /* Is there a Legal Representative Third Party set up for this application */ %>
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	ThirdPartyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ThirdPartyXML.CreateRequestTag(window,null);
	<%/* ThirdPartyXML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>

	var tagApplication = ThirdPartyXML.CreateActiveTag("APPLICATION");
	ThirdPartyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	ThirdPartyXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	ThirdPartyXML.RunASP(document,"FindApplicationThirdPartyList.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = ThirdPartyXML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0] == false)
		return false;

	<% /* Attempt to find an LegalRep thirdparty */ %>
	var sCondition = "THIRDPARTY[THIRDPARTYTYPE='10']";
	ThirdPartyXML.CreateTagList(sCondition);
	if(ThirdPartyXML.ActiveTagList.length > 0)
		return true;
	else
		return false;
}

function SetAdvanceDate()
{
	<% /* WP13 MAR49 Set Advanced Date and update screeen with the value returned*/ %>
	var bSuccess = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"REQUEST");
	XML.SetAttribute("OPERATION","SetAdvanceDate");
	XML.CreateActiveTag("SETADVANCEDATE");
	XML.SetAttribute("COMPLETIONDATE",m_sCompletionDate);
	var sPayMethod = scScreenFunctions.GetComboValidationType(frmScreen, "cboPaymentMethod")
	XML.SetAttribute("PAYMENTMETHOD",sPayMethod);
	XML.RunASP(document,"PaymentProcessingRequest.asp");	
	
	var TagRESPONSE = XML.SelectTag(null,"RESPONSE");
	if (XML.SelectTag(TagRESPONSE,"ERROR") != null)
	{
		var sErrorMessage = XML.GetTagText("DESCRIPTION");
		alert(sErrorMessage);
	}
	else
	{
		XML.SelectSingleNode("/RESPONSE/ADVANCEDATERETURNED");
		var sCompDate = XML.GetAttribute("ADVANCEDATE");
		if (sCompDate !="")
		{
			frmScreen.txtCompDate.value=sCompDate;
			bSuccess=true;
		}
	}
	XML=null;
	return bSuccess;
}

function SetPaymentMethod()
{
	<% /* WP13 MAR49 Set Payment Method and update screeen with the value returned*/ %>
	var bSuccess = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"REQUEST");
	XML.SetAttribute("OPERATION","SetPaymentMethod");
	XML.CreateActiveTag("SETPAYMENTMETHOD");
	XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	XML.SetAttribute("COMPLETIONDATE",m_sCompletionDate);
	
	XML.RunASP(document,"PaymentProcessingRequest.asp");	
	
	var TagRESPONSE = XML.SelectTag(null,"RESPONSE");
	if (XML.SelectTag(TagRESPONSE,"ERROR") != null)
	{
		var sErrorMessage = XML.GetTagText("DESCRIPTION");
		alert(sErrorMessage);
	}
	else
	{
		XML.SelectSingleNode("/RESPONSE/PAYMENTMETHODRETURNED");
		var sPayMethod = XML.GetAttribute("PAYMENTMETHOD");
		if (sPayMethod !="")
		{
			scScreenFunctions.SetComboOnValidationType(frmScreen, "cboPaymentMethod", sPayMethod);
			bSuccess=true;
		}
	}
	XML=null;
	return bSuccess;
}

function GetCompletionDate()
{
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var RotXML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>

	RotXML.CreateRequestTag(window,"GetReportOnTitleData");
	RotXML.CreateActiveTag("REPORTONTITLE");
	RotXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	RotXML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	RotXML.RunASP(document,"ReportOnTitle.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = RotXML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0] == false)
		return "";
	else
	{
		<% /* Get Completion Date */ %>
		RotXML.SelectSingleNode("REPORTONTITLE");
		var sCompDate = RotXML.GetAttribute("COMPLETIONDATE");
		return sCompDate;
	}

}

function DoEditProcessing()
{
	<% /* Set PaymentType, PaymentMethod and Payee to read-only */ %>
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboPaymentType");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboPaymentMethod");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboPayee");
	
	var sIssueDate = frmScreen.txtIssueDate.value ;
	var dtIssueDate, dtCurrentDate ;
	if(sIssueDate != '')
	{
		if(scScreenFunctions.CompareDateStringToSystemDateTime(sIssueDate, '<=')) 
			scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtIssueDate"); 			
	}
	else
		frmScreen.txtIssueDate.value = m_sDate; //WP13 MAR49
}

function DoNonEditProcessing()
{
	function RemovePaymentTypeFromCombo(sPaymentId)
	{
		var iLoop = frmScreen.cboPaymentType.options.length - 1 ;
		while (iLoop >= 0)
		{
			if(frmScreen.cboPaymentType.options(iLoop).value == sPaymentId) 
			{
				frmScreen.cboPaymentType.options.remove(iLoop);
				break ;
			}
			--iLoop ;
		}	
	}

	function RemoveAllNonInitialAdvances()
	{
		var iLoop = frmScreen.cboPaymentType.options.length - 1 ;
		while (iLoop >= 0)
		{
			if(frmScreen.cboPaymentType.options(iLoop).value != m_sInitialAdvanceId && frmScreen.cboPaymentType.options(iLoop).value != "") 
			{
				frmScreen.cboPaymentType.options.remove(iLoop);
			}
			--iLoop ;
		}	
	}
	
	var sCondition		= "//PAYMENTRECORD[@AMOUNT $ne$ 0][DISBURSEMENTPAYMENT/@PAYMENTTYPE='" + m_sInitialAdvanceId + "']" ;
	var xmlPaymentList	= XMLPayments.XMLDocument.selectNodes(sCondition);
	var xmlPaymentNode ;
	var sPaymentSeqNo, lngROFTotal, iCount, sPaymentAmount, sTemp ;
	
	if(xmlPaymentList.length == 0) 
	{	<% /* if payment type 'Initial Advance' already exist, disable payment type 
						'InitialAdvance' and enable 'AdditionalAdvance'  */ 
		%>
		RemoveAllNonInitialAdvances() ;
	}
	else
	{
		xmlPaymentNode = xmlPaymentList.item(xmlPaymentList.length - 1);
		sPaymentSeqNo = xmlPaymentNode.attributes.getNamedItem("PAYMENTSEQUENCENUMBER").text ;
		sPaymentAmount = xmlPaymentNode.attributes.getNamedItem("AMOUNT").text ;
		
		sCondition = "//PAYMENTRECORD[@ASSOCPAYSEQNUMBER=" + sPaymentSeqNo + "]" +
					 "[DISBURSEMENTPAYMENT/@PAYMENTTYPE=" + m_sReturnOfFundId + "]"	
		
		xmlPaymentList = XMLPayments.XMLDocument.selectNodes(sCondition);
		
		<% /* if no ReturnOfFunds records are found */ %>
		if(xmlPaymentList.length == 0) RemovePaymentTypeFromCombo(m_sInitialAdvanceId)
		else
		{
			lngROFTotal = 0 ;
			for(iCount = 0 ; iCount <= xmlPaymentList.length - 1; ++iCount)
			{	
				xmlPaymentNode = xmlPaymentList.item(iCount);
				sTemp = xmlPaymentNode.attributes.getNamedItem("AMOUNT").text ;
				if(!isNaN(sTemp)) lngROFTotal = lngROFTotal + parseFloatSafe(sTemp);
			}
			
			<% /* MV - 11/09/2002 - BMIDS00380 - AFTER ROF, WHEN CREATE NEW PAYMENT THE PAYMENT TYPE AVAILABLE IS INCORRECT */ %>
			if(parseFloatSafe(sPaymentAmount) == Math.abs(lngROFTotal)) RemoveAllNonInitialAdvances();
			else RemovePaymentTypeFromCombo(m_sInitialAdvanceId) 
		}
	}
}

function PopulateListbox(nStart)
{
	<% /* Populate List box with FeeType and Amount Outstanding. This data is to be passed from  */ %>
	var iNumberOfRows = XMLFeeType.ActiveTagList.length;
	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}

function ShowList(nStart)
{
	var sFeeTypeDesc, sAmountOutStanding ;
	scScrollTable.clear();
	<% /* GD BM0198(v2) START */ %>
	var sFeeType;
	var ValidationList = new Array(1);	
	ValidationList[0] = "TTF";
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<% /* GD BM0198(v2) END */ %>
	
	m_sTotalFeesDeductable = 0 ;
	
	<% /* Loop through the XML (Fee Payments) and populate the list box */ 	%>
	for(var nLoop = 0 ; nLoop < XMLFeeType.ActiveTagList.length && nLoop<m_iTableLength; nLoop++)
	{
		XMLFeeType.SelectTagListItem(nStart+nLoop);
		sFeeTypeDesc = XMLFeeType.GetAttribute("FEETYPE_TEXT");
		sAmountOutStanding = XMLFeeType.GetAttribute("AMOUNTOUTSTANDING");
		//m_sTotalFeesDeductable += parseFloatSafe(sAmountOutStanding);
		
		<% /*----------- 18/08/2003 GD BM0198(v2) START Check for the existence of a TT Fee already---------*/ %>
		//sFeeType = XMLFeeType.GetAttribute("FEETYPE");
		//Check if the FEETYPE has the validation type of TTF from combo group OneOffCost
		//if (m_bTTFeeAdded == false)
		//{
		//	if( XML.IsInComboValidationList(document,"OneOffCost", sFeeType, ValidationList))
		//	{
		//		m_sTTFee = sAmountOutStanding;
		//		m_bTTFeeAdded = true;
				//m_bStandingData = true;	
		//	}
		//}
		<% /*-------------- 18/08/2003 GD BM0198(v2) END ---------------------------------------------------*/ %>
		
		<% /* Add to the search table */ %>
		scScreenFunctions.SizeTextToField(tblTable.rows(nLoop+1).cells(0), sFeeTypeDesc);
		scScreenFunctions.SizeTextToField(tblTable.rows(nLoop+1).cells(1), sAmountOutStanding);
	}
	<% /*JLD We want the total of all fees, not just those showing in the listbox*/ %>
	for(var nLoop = 0 ; nLoop < XMLFeeType.ActiveTagList.length; nLoop++)
	{
		XMLFeeType.SelectTagListItem(nLoop);
		sAmountOutStanding = XMLFeeType.GetAttribute("AMOUNTOUTSTANDING");
		m_sTotalFeesDeductable += parseFloatSafe(sAmountOutStanding);
	}
	frmScreen.txtTotalFeeDeductable.value = m_sTotalFeesDeductable ;
}

function PopulateOldValues()
{
	frmScreen.cboPaymentType.value = m_sPaymentType ;
	frmScreen.cboPaymentMethod.value = m_sPaymentMethod ;
	frmScreen.cboPayee.value	= m_sPayeeHistorySeqNo ;
	frmScreen.txtAdvanceAmount.value = m_sAmountOfAdvance ;
	m_iInitialIncentiveAmount = m_sAmountOfAdvance ;
	frmScreen.txtIssueDate.value	= m_sIssueDate ;
	frmScreen.txtCompDate.value		= m_sCompletionDate ;
	frmScreen.txtNotes.value		= m_sNotes ;
}
function frmScreen.txtAdvanceAmount.onblur()
{
	frmScreen.txtNetAdvanceAmount.value = parseFloatSafe(frmScreen.txtAdvanceAmount.value) - parseFloatSafe(frmScreen.txtTotalFeeDeductable.value);
}
function validateBeforeSave()
{
	if(!frmScreen.onsubmit())
	{
		return false ;
	}
	
	<% /* BMIDS808  Ensure that the Net Amount has been updated */ %>
	frmScreen.txtNetAdvanceAmount.value = parseFloatSafe(frmScreen.txtAdvanceAmount.value) - parseFloatSafe(frmScreen.txtTotalFeeDeductable.value);
	
	<% /* Check mandatory fields */ %>
	 if(frmScreen.cboPaymentType.value == '')
	 {
		alert('Select a value for payment type');
		frmScreen.cboPaymentType.focus();
		return false ;
	 }
	 
	 if(frmScreen.cboPaymentMethod.value == '')
	 {
		alert('Select a value for payment method');
		frmScreen.cboPaymentMethod.focus();
		return false ;
	 }
	
	if(frmScreen.cboPayee.value == '')
	{
		alert('Select a value for payee');
		frmScreen.cboPayee.focus();
		return false ;
	}
	
	if(frmScreen.txtAdvanceAmount.value == '')
	{
		alert('Amount of advance cannot be left empty');
		frmScreen.txtAdvanceAmount.focus();
		return false ;	
	}
	
	// BM0017  Check for negative disbursement
	if(frmScreen.txtNetAdvanceAmount.value <= 0)
	{
		alert('Amount of disbursement must be greater than 0');
		frmScreen.txtAdvanceAmount.focus();
		return false ;	
	}		
	
	if(frmScreen.cboPaymentType.value == m_sIncentiveReleaseId)
	{
		if(m_sMetaAction == 'Add')
		{
			if(parseFloatSafe(frmScreen.txtAdvanceAmount.value) > m_iIncentiveAmount)
			{
				alert('Amount of advance should not be greater than outstanding Incentive to be paid.');
				return false ;
			} 
		}
		else
		{
			if(m_sMetaAction == 'Edit')
			{
				if(parseFloatSafe(frmScreen.txtAdvanceAmount.value) > m_iInitialIncentiveAmount	+ m_iIncentiveAmount)
				{
					alert('Amount of advance should not be greater than outstanding Incentive to be paid.');			
					return false ; 
				}
			}
		}
	}
	
	<%/* In create mode, Amount of advance should be greater than Total to be deducted from advance*/%>
	<%/* SR 11-12-01 : SYS3380 - remove validation on OK that prevents advance amount < Outstanding fee to be deducted*/%>
	<% /*
	if(m_sMetaAction != 'Edit' && frmScreen.txtTotalFeeDeductable.value != '')
	{
		if(parseFloatSafe(frmScreen.txtAdvanceAmount.value) < parseFloatSafe(frmScreen.txtTotalFeeDeductable.value))
		{
			alert('Amount of advance should be greater than Total to be deducted from advance.');
			frmScreen.txtAdvanceAmount.focus();
			return false ;
		}
	}
	*/ %>
	<%/* SR 11-12-01 : SYS3380 END */ %>
	 
	<% /* Advance amount cannot be greater than TotalOutstanding  */ %>
	<% /* PSC 09/03/2007 EP2_1347 */ %>
	if(frmScreen.cboPaymentType.value != m_sRetentionReleaseId  &&
	  frmScreen.cboPaymentType.value != m_sIncentiveReleaseId &&
	  frmScreen.cboPaymentType.value != m_sValuationRefundId )     // JLD SYS5019
	{
		if(parseFloatSafe(frmScreen.txtAdvanceAmount.value) > m_iTotalOutstanding)
		{
			alert('Advance amount cannot exceed the total outstanding : ' + m_iTotalOutstanding );
			frmScreen.txtAdvanceAmount.focus();
			return false ;
		}
	}
	else if(frmScreen.cboPaymentType.value == m_sRetentionReleaseId)   // JLD SYS5019
	{
		if(parseFloatSafe(frmScreen.txtAdvanceAmount.value) > m_iRetentionAmount)
		{
			alert('Advance amount cannot exceed the remaining Retention Amount : ' + m_iRetentionAmount );
			frmScreen.txtAdvanceAmount.focus();
			return false ;
		}
	}
	
	<% /* PSC 08/09/2003 BM0198 - Start */ %>
	var dteCompDate;
	
	if(frmScreen.txtCompDate.value != '')
	{
		<% /* PSC 11/11/2002 BMIDS00897 - Start */ %>
		dteCompDate = scScreenFunctions.GetDateObject(frmScreen.txtCompDate);
		if (dteCompDate == null)
		{
			frmScreen.txtCompDate.focus();
			return false;
		}
	}
		
	if(frmScreen.txtIssueDate.value == '')
	{
		alert('Issue Date cannot be left empty');
		frmScreen.txtIssueDate.focus();
		return false ;		
	}
	else
	{
		<% /* PSC 11/11/2002 BMIDS00897 - Start */ %>
		var dteIssueDate = scScreenFunctions.GetDateObject(frmScreen.txtIssueDate);
		if (dteIssueDate == null)
		{
			frmScreen.txtIssueDate.focus();
			return false;
		}
		
		var blnCompDateInPast = scScreenFunctions.CompareDateStringToSystemDate(frmScreen.txtCompDate.value, '<');
		var blnIssueDateInPast = scScreenFunctions.CompareDateStringToSystemDate(frmScreen.txtIssueDate.value, '<');
			
		if (dteIssueDate > dteCompDate)
		{
			alert('Issue Date cannot be after the Advance Date');
			frmScreen.txtIssueDate.focus();
			return false;
		}
		else if (blnIssueDateInPast == true && blnCompDateInPast == true)
		{
			if(!confirm("The Issue and Advance Dates are in the past. Click OK to continue or Cancel to amend the dates."))
			{
				frmScreen.txtIssueDate.focus();
				return false ;
			}
		}
		else if (blnIssueDateInPast == true)
		{
			if(!confirm("The Issue Date is in the past. Click OK to continue or Cancel to amend the date."))
			{
				frmScreen.txtIssueDate.focus();
				return false ;
			}
		}
			
		<% /* PSC 11/11/2002 BMIDS00897 - End */ %>
		<% /* PSC 08/09/2003 BM0198 - End */ %>	
	}
		
	<% /* Note can only be a max 255 chars */ %>
	if(scScreenFunctions.RestrictLength(frmScreen, "txtNotes", 255, true))
	{
		frmScreen.txtNotes.focus();
		return false;
	}
		
	return true;
}

function SaveDisbursement(bCreate)
{
	var xmlRequestTag ;

	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	XMLSave = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* XMLSave = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>
	<%//GD BM0198 START%>
	<%//GD BM0198 Call SaveCostsFeesDisbursements %>
	xmlRequestTag = XMLSave.CreateRequestTag(window, "SAVECOSTSFEESDISBURSEMENTS");
	<%//GD BM0198 if(bCreate) xmlRequestTag = XMLSave.CreateRequestTag(window, "CREATEDISBURSEMENT");
	//GD BM0198 else xmlRequestTag = XMLSave.CreateRequestTag(window, "UPDATEDISBURSEMENT");%>
	<%//GD BM0198 END%>
	<%//GD BM0198 START %>

		var xmlSettings = XMLSave.CreateActiveTag("SETTINGS");
		if(bCreate)
		{
			xmlSettings.setAttribute("METHODCREATE", "1");
		} else
		{
			xmlSettings.setAttribute("METHODCREATE", "0");
		}
		xmlSettings.setAttribute("TTFEEADDED", m_bTTFeeAdded);
		var sTemp = "0";
		if(frmScreen.chkWaiveFee.checked) sTemp = "1";
		xmlSettings.setAttribute("WAIVECHECKBOX", sTemp);	
		
		if (m_bStandingData)
		{
			xmlSettings.setAttribute("STANDINGDATA", "1");
		} else
		{
			xmlSettings.setAttribute("STANDINGDATA", "0");
		}

	<%//GD BM0198 END %>
	XMLSave.SelectTag(null,"REQUEST");	
	var xmlPaymentRecord = XMLSave.CreateActiveTag("PAYMENTRECORD");
	XMLSave.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	<% /* SR 01/08/2001 : SYS2545 - Include applicationFactFindNumber in the request */ %>
	XMLSave.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	if(!bCreate) XMLSave.SetAttribute("PAYMENTSEQUENCENUMBER", m_sPaymentSequenceNumber);
	// JLD SYS5019 removed DB SYS4767 - MSMS Integration
	XMLSave.SetAttribute("AMOUNT", parseInt(frmScreen.txtAdvanceAmount.value)); //JLD + parseInt(frmScreen.txtTotalFeeDeductable.value));
	XMLSave.SetAttribute("USERID", m_sUserId);
	XMLSave.SetAttribute("PAYMENTMETHOD", frmScreen.cboPaymentMethod.value);		
	
	<% /* Set PaymentStatus 'UnSanctioned' 
	GetUnsanctionedStatus(); */ %>
	
	XMLSave.CreateActiveTag("DISBURSEMENTPAYMENT");
	if(!bCreate) XMLSave.SetAttribute("PAYMENTSEQUENCENUMBER", m_sPaymentSequenceNumber);
	XMLSave.SetAttribute("ISSUEDATE", frmScreen.txtIssueDate.value);
	XMLSave.SetAttribute("COMPLETIONDATE", frmScreen.txtCompDate.value);
	XMLSave.SetAttribute("PAYMENTTYPE", frmScreen.cboPaymentType.value);
	XMLSave.SetAttribute("PAYEEHISTORYSEQNO", frmScreen.cboPayee.value);
	<% /* Payee Type stored as attrib in cboPayee */ %>
	XMLSave.SetAttribute("PAYEETYPE", frmScreen.cboPayee.options(frmScreen.cboPayee.selectedIndex).getAttribute("PayeeType"));
	
	var globalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sPaymentStatus;
	if (globalXML.GetGlobalParameterBoolean(document, "PPUseSanctionProcess"))
	{
		sPaymentStatus = GetPaymentStatusForValidationType("U");
	}	
	else
	{	
		sPaymentStatus = GetPaymentStatusForValidationType("P");
	}
	
	XMLSave.SetAttribute("PAYMENTSTATUS", sPaymentStatus);
	
	XMLSave.SetAttribute("PAYMENTNOTES", frmScreen.txtNotes.value);
	//DB SYS4767 - MSMS Integration
	//XMLSave.SetAttribute("NETPAYMENTAMOUNT", parseInt(frmScreen.txtNetAdvanceAmount.value));  // JLD SYS4177
	//SG 05/04/02 MSMS009
	if (m_sShortfallPayment != "0")
		XMLSave.SetAttribute("SHORTFALL", parseInt(m_sShortfallPayment));
	
	XMLSave.SetAttribute("NETPAYMENTAMOUNT", parseInt(frmScreen.txtNetAdvanceAmount.value));  // JLD SYS4177
	//DB End
	
	<% /* Loop through the XML (Fee Payments) and add them to the request tag */ %>
	if((parseFloatSafe(frmScreen.txtTotalFeeDeductable.value) > 0) || (m_bTTFeeAdded == true))
	{
		XMLFeeType.ActiveTag = XMLFeeType.XMLDocument.documentElement;
		XMLFeeType.CreateTagList("APPLICATIONFEETYPE");
		//GD 
		
		for(var nLoop = 0; nLoop < XMLFeeType.ActiveTagList.length; nLoop++)
		{
			XMLFeeType.SelectTagListItem(nLoop);
			XMLSave.ActiveTag = xmlPaymentRecord ;
			XMLSave.CreateActiveTag("FEEPAYMENT");
			XMLSave.SetAttribute("FEETYPE",XMLFeeType.GetAttribute("FEETYPE"));
			XMLSave.SetAttribute("FEETYPESEQUENCENUMBER", XMLFeeType.GetAttribute("FEETYPESEQUENCENUMBER"));
			XMLSave.SetAttribute("AMOUNTPAID",XMLFeeType.GetAttribute("AMOUNTOUTSTANDING"));
			XMLSave.SetAttribute("PAYMENTEVENT", m_sDeductionId);
		
			//SYS2328 default completion indicator
			XMLSave.SetAttribute("COMPLETIONINDICATOR", "1");
		}
	}
	<% /* BMIDS876 GHun set window status as request may take a while. */ %>
	window.status = "Generating disbursements ...";
	XMLSave.RunASP(document,"PaymentProcessingRequest.asp");
	window.status = "";
	
	if(!XMLSave.IsResponseOK())	return false ;
	
	return true;
}

function GetPaymentStatusForValidationType(sValidationType)
{
	var xmlNode = null ;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var blnSuccess = true;		
	var sGroupList = new Array("PaymentStatus");
	
	if(XML.GetComboLists(document,sGroupList))
	{
	
		var sPaymentStatusValueID   = XML.GetComboIdForValidation("PaymentStatus", sValidationType , null, document) ;
		
		return sPaymentStatusValueID;
		//var sCondition = "//LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='" +  sValidationType + "']/VALUEID" ;
		
		//xmlNode = XML.XMLDocument.selectSingleNode(sCondition);
		
		//if(xmlNode != null)  m_sUnsanctioned = xmlNode.text ;
		//else m_sUnsanctioned = "" ;
		
		//blnSuccess = true ;	
	}
	else
	{
		alert('Error retrieving combo values.');
		return false;
		//blnSuccess =  false;
	}
	
	//return blnSuccess ;
}

function ClearScreen()
{
	frmScreen.txtTotalAdvanceToDate.value		= "" ;
	frmScreen.txtAdvanceAmount.value	= "" ;
	frmScreen.cboPaymentType.value		= "" ;
	frmScreen.cboPaymentMethod.value	= "" ;
	frmScreen.cboPayee.value			= "" ;
	frmScreen.txtIssueDate.value		= "" ;
	frmScreen.txtCompDate.value			= "" ;
	frmScreen.txtNotes.value			= "" ;
	
	<% /* SR 12-12-2001 : SYS2229 */ %>
	scScrollTable.clear();
	frmScreen.txtTotalFeeDeductable.value = 0 ;
}

<% /*SYS4564 Add functions to prevent tab leaving page. */ %>
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function DefaultDate()
{
	SetDaysToSubtract();
		 
	<% /* if Issue Date has not been entered, but Completion Date, Payment Method and Payment Type have... */ %>
	if(frmScreen.txtIssueDate.value == "" && frmScreen.txtCompDate.value != "" && frmScreen.cboPaymentMethod.value != "" && frmScreen.cboPaymentType.value != "")
	{
		var dtCompDate = scScreenFunctions.GetDateObject(frmScreen.txtCompDate);
			
		if (scScreenFunctions.IsValidationType(frmScreen.cboPaymentType,"I") || scScreenFunctions.IsValidationType(frmScreen.cboPaymentType,"C")) 	
			frmScreen.txtIssueDate.value = CalculateCompletionDate(dtCompDate,m_nDaysToSubtract)
		else
			frmScreen.txtIssueDate.value =  m_sDate;
	}
}
<% /* PSC 08/09/2003 BM0476 - Start */ %>
function CalculateCompletionDate(dtCompDate , nDaysToSubtract)
{
	var dtNewDate = new Date(dtCompDate);
	dtNewDate.setDate(dtNewDate.getDate() - nDaysToSubtract); 
	
	var sDay = dtNewDate.getDate();
	var sMonth = dtNewDate.getMonth() + 1;
	
	if (parseInt(sDay) < 10)
		sDay = "0" + sDay;
		
	if (parseInt(sMonth) < 10)
		sMonth = "0" + sMonth; 
	
	return sDay + "/" + sMonth + "/" + dtNewDate.getFullYear();
}
<% /* PSC 08/09/2003 BM0476 - Start */ %>


function GetSystemDate()
{
	var dtCurrentDte	= scScreenFunctions.GetAppServerDate();
	m_sDate				= scScreenFunctions.DateToString(dtCurrentDte);
}
<%/* GD BM0198 Added Start */%>
function parseFloatSafe(sText)
{
	if ((sText == "") || (isNaN(sText))) return(0)
	else
	{
		return(parseFloat(sText));
	}
	
}
<%/* GD BM0198 Added  End */%>

<% /* PSC 08/09/2003 BM0198 - Start */ %>
function SetDaysToSubtract()
{
	m_nDaysToSubtract = 0;
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	switch(scScreenFunctions.GetComboValidationType(frmScreen, "cboPaymentMethod"))
	{
		case "B": // BACS
			m_nDaysToSubtract = XML.GetGlobalParameterAmount(document,"BACSPaymentDays");
			break;
		case "YC": 
			// TT Payment
			m_nDaysToSubtract = XML.GetGlobalParameterAmount(document,"TTPaymentDays");
			break;
		case "CH":
			//Cheque
			m_nDaysToSubtract = XML.GetGlobalParameterAmount(document,"CHQPaymentDays");
			break;
		default:
			m_nDaysToSubtract = 2 ;
	}
}
<% /* PSC 08/09/2003 BM0198 - End */ %>

<% /* SR 31/08/2004 : BMIDS815  */ %>
function AddOtherSystemCustomerNumbers(XML)
{
	var sCustomerNumber, sCustomerVersionNumber, sCustomerRoleType, sOtherSystemCustomerNumber ;
	
	var ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(window,"SEARCH");
	ListXML.CreateActiveTag("APPLICATION");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	ListXML.RunASP(document,"FindCustomersForApplication.asp");

	if (ListXML.IsResponseOK())
	{
		var tagCustomerList = XML.CreateActiveTag("CUSTOMERLIST");
		ListXML.CreateTagList("CUSTOMERROLE");
		var iNoOfCustomers = ListXML.ActiveTagList.length;
		for(var nCount = 0; nCount < iNoOfCustomers; nCount++)
		{	
			ListXML.SelectTagListItem(nCount);
			sCustomerNumber = ListXML.GetTagText("CUSTOMERNUMBER");
			sCustomerVersionNumber = ListXML.GetTagText("CUSTOMERVERSIONNUMBER");
			sCustomerRoleType = ListXML.GetTagText("CUSTOMERROLETYPE");
			sOtherSystemCustomerNumber = ListXML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER");
			if(sCustomerNumber != "" && sCustomerVersionNumber != "" && sCustomerRoleType == "1")
			{
				XML.ActiveTag = tagCustomerList;
				XML.CreateActiveTag("CUSTOMER");
				XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
				XML.CreateTag("OTHERSYSTEMCUSTOMERNUMBER", sOtherSystemCustomerNumber);
			}
		}
	}
}
<% /* SR 31/08/2004 : BMIDS815  - End */ %>

function frmScreen.txtNotes.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtNotes", 255, true);
}
-->
</script>
</BODY>
</HTML>

