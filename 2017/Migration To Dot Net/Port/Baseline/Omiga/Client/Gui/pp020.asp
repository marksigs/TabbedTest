<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<%
/*
Workfile:      pp020.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Make Payments
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
IVW		01/01/2001	New Screen
IVW		25/01/2001	Fixed errors SYS1870,SYS1864.
SR		08/02/2001	SYS1891 - If amount is not input, run time error occurs while saving .
SR		22/02/2001	SYS1891 - Do not allow a payment without any amount
SR		23/05/2001	SYS2298 
SR		25/05/2001	SYS2298 - new field 'Refund Amount'
SR		08/06/01	SYS2298 - Refund date can be empty 
SR		06/09/01	SYS2412
PSC		17/01/2002	SYS3534 - Amend ValidateBeforeSave to validate fee amount correctly
DRC     19/01/02    SYS3524 - Don't use PaymentMethod or Cheque Number when Payment Type = Rebate or Addition                            
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description

TW      09/10/2002	Modified to incorporate client validation - SYS5115
MO		14/11/2002	BMIDS00807 -Made change to take the date from the app server not the client
PSC		16/11/2002	BMIDS00924 -Amend so that if payment amount is greater than the fee due then
								stop processing rather than asking user if they want to continue
MDC		17/12/2002	BM0196		Remove validation on amount if doing an Addition.
MDC		02/01/2003	BM0219		Use Validation Types when responding to change in Payment Type
MV		17/03/2003	BM0412		Amended DisableChequeNumber()
MV		19/03/2003	BM0412		Amended DisableChequeNumber()
GD		03/06/2003	BM0372		Handle Rebate or Additions correctly
MC		20/04/2004	BMIDS517	White space padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:

HMA     09/08/2005  MAR28		WP12. Add new Payment Record fields.
                                Comments tidied. Unused code removed.
HMA     06/10/2005  MAR28       Further changes for WP12.                                
HM      08/11/2005  MAR462      Error message appears when <SELECT> combo valueID used                               
PE		08/12/2005	MAR822		Modifed ValidateBeforeSave to cancel save if the notes field 
								is more than 255 characters.
PE		21/02/2006	MAR1267		Truncation errors on fields check no and savings account
PE		03/03/2006	MAR1342		SIT - OMIGA - Transaction reference field should be enable for payment method Savings Account
HMA     28/03/2006  MAR1517     Disable fields for RFV and NTR Payment Types.
HMA     29/03/2006  MAR1517     Further changes when editing RFV Payment types.
HMA     25/04/2006  MAR1652     Default Payment Type to 'awaiting payment' for cheques.
                                Disable Payment Type depending on authority level.
HMA     15/05/2006  MAR1754     Correct combo manipulation.     
HMA     16/05/2006  MAR1797     Correct error when editing a fee already paid by Credit Card.                           
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>

<title>PP020 - Add/Edit Application Cost  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>

</HEAD>

<BODY>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets */ %>

<form id="frmClickCancel" method="post" action="PP010.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 290px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Fee Type
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtFeeType" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>

	<span style="TOP: 10px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Total Amount Due
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtTotalAmountDue" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>

	<span style="TOP: 37px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Payment Method 
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<select id="cboFeePaymentMethod" style="WIDTH: 150px" class="msgCombo"></select>
		</span>
	</span>
	
	<span style="TOP: 37px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Amount Outstanding
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtAmountOutStanding" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
		
	<span style="TOP: 64px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Payment Type 
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<select id="cboPaymentEvent" style="WIDTH: 150px" class="msgCombo"></select>
		</span>
	</span>
	
	<span style="TOP: 91px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Amount of Payment 
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtPaymentAmount" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	
	<span style="TOP: 91px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Cheque Number
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtChequeNumber" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt" maxLength="20">
		</span>
	</span>

	<span style="TOP: 118px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Date of Payment
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtIssueDate" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	
	<span style="TOP: 118px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Savings A/C No.
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtSavingsAccNo" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt" maxLength="20">
		</span>
	</span>
	
	<span style="TOP: 145px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Transaction Reference
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtTransRef" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt" maxlength="36">
		</span>
	</span>

	<span style="TOP: 145px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Result Code
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtResultCode" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<span style="TOP: 172px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Date of Refund
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtRefundDate" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 172px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
		Amount of Refund
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtRefundAmount" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	
	<span style="TOP: 200px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Notes
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<TEXTAREA class=msgTxt id="txtNotes" name=Notes rows=5 style="POSITION: absolute; WIDTH: 450px"></TEXTAREA>  
		</span>	
	</span>
	
	<span style="TOP: 280px; LEFT: 4px; POSITION: ABSOLUTE">
		<input id="btnOk" value="OK" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
	<span style="TOP: 280px; LEFT: 100px; POSITION: ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
	
</div>
</form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="attribs/PP020attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sUserId ;
var m_sUserRole;     <% /* MAR1652 */ %>
var m_sMinUserRole;  <% /* MAR1652 */ %>

var m_sChequeNumber = "";
var scScreenFunctions;
var m_sApplicationNumber;
var m_sFeeType;
var m_sFeeTypeDesc ;
var m_sAmountOutStanding ;
var m_sFeePaymentSeqNo ;
var m_aArgArray = null;
var m_sPayMethod = null;
var m_sPaymentEvent = null;
var m_sPaymentEventDesc = null;
var m_sFeeTypeRecordSeq = null;
var m_sDate = "";
var m_sRefundDate = "";
var m_sRefundAmount = "" ;
var m_sRebateAdditionId = null ;
var m_sPaymentId = null ;
var m_sNotes = "" ;
var m_sAmountOfPayment = null;
var m_sPreviousPayment = null;
var m_BaseNonPopupWindow = null;
<% /* MAR28 */ %>
var m_sSavingsAccountNumber = null;  
var m_sTransactionReference = null;  
var m_sResultCode = null;            
var m_PaymentEventXML = null;        
var m_bCardPaymentMade = false;      
var m_sNTRPaymentMethod;            
var m_sNTRPaymentType;               
var m_sRFVPaymentType; 

<% /* MAR1517 */ %>     
var m_sARPaymentMethod;            
var m_sARPaymentType;    
           
<% /**** Events *****/ %>

function frmScreen.btnCancel.onclick()
{
	window.close();
}

function frmScreen.btnOk.onclick()
{
	if(frmScreen.onsubmit())
	{
		if(ValidatePaymentAmount() && ValidateBeforeSave())
		{
			if (SaveFeeTypePayment())
			{
				window.returnValue	= "Ok";
				window.close();
			} 
		}	
	}
}

function ValidateBeforeSave()
{	
	<% /* BM0196 MDC 17/12/2002 */ %>
	if(frmScreen.cboPaymentEvent.value != 'ADDITION')
	{
		<% /* if Amount paid is more than Amount Outstanding, prompt the user. */ %>
		if (parseInt(frmScreen.txtPaymentAmount.value) > 
		   parseInt(frmScreen.txtAmountOutStanding.value) + parseInt(m_sPreviousPayment))
		{
			<% /* PSC 16/11/2002 BMIDS00924 */ %>
			alert("The payment/rebate amount is greater than the fee due");
			frmScreen.txtPaymentAmount.focus();
			return false;
		}
	}
	<% /* BM0196 MDC 17/12/2002 - End */ %>
	
	<% /* if Amount of Refund is mentioned, it should not be greater than Amount paid */ %>
	if(frmScreen.txtRefundAmount.value != '')
	{
		if((parseInt(frmScreen.txtRefundAmount.value) > parseInt(frmScreen.txtPaymentAmount.value)) &&
		   (frmScreen.cboPaymentEvent.value != m_sNTRPaymentType) &&
		   (frmScreen.cboPaymentEvent.value != m_sRFVPaymentType) &&
		   (frmScreen.cboPaymentEvent.value != m_sARPaymentType))		   
		{
			alert('The amount of refund cannot be greater than the payment amount');
			frmScreen.txtRefundAmount.focus();
			return false ;
		}
	}
	
	<% /* if DateofRefund is mentioned, the 'AmountOfRefund' must also be mentioned */ %>
	if(frmScreen.txtRefundDate.value != '' && frmScreen.txtRefundAmount.value == '')
	{
		alert('Date of refund cannot be left empty, when amount of refund is mentioned.');
		frmScreen.txtRefundAmount.focus();
		return false ;
	}
	
	<% /* Date of Refund (if mentioned), should be greater than Date of Payment and not a past date */%>
	if(frmScreen.txtRefundDate.value != '')
	{
		if(! scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtRefundDate , '>='))
		{
			alert('Date of refund cannot be in past.');
			frmScreen.txtRefundDate.focus();
			return false ;
		}
		
		if(frmScreen.txtIssueDate.value != '')
			if(! scScreenFunctions.CompareDateFields(frmScreen.txtRefundDate, '>=', frmScreen.txtIssueDate))
			{
				alert('Date of refund cannot be prior to date of payment.');
				frmScreen.txtRefundDate.focus();
				return false ;
			}
	}

	<% /* MAR822 - Peter Edney - 08/12/2005  */ %>
	if(frmScreen.txtNotes.value.length > 255)	
	{
		alert('Notes must not be longer than 255 characters.');
		frmScreen.txtNotes.focus();
		return false ;
	}
	
	return true ;
}

function frmScreen.cboFeePaymentMethod.onchange()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();			

	DisableChequeNumber();
	<%/* AQR SYS 3524  */ %>
	m_sPayMethod = frmScreen.cboFeePaymentMethod.value;
	<%/* MAR462 comparison to '' added */ %>
	if ((m_sPayMethod != null) && (m_sPayMethod != " ") && (m_sPayMethod != ""))
	{

		<% // MAR1342 - Peter Edney - 03/03/2006 %>
		switch(true)
		{
		case XML.IsInComboValidationList(document,"FeePaymentMethod", m_sPayMethod, ["CD"]):
		{
			<% /* Payment Method has validation type of CD (eg Credit Card) */ %>
			<% /* Disable Savings Account No, Date of Refund, Amount of Refund */ %>
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtSavingsAccNo");
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtRefundDate");
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtRefundAmount");
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtTransRef");
			
			<% /* Repopulate the Payment Type combo with all values */ %>
			XML.PopulateComboFromXML(document, frmScreen.cboPaymentEvent,m_PaymentEventXML,true);  
						
			<% /* Remove Payment option (validation type P) from Payment Type combo */ %>						
			<% /* Correct combo test */ %>
			var iCount = 0;
			while(frmScreen.cboPaymentEvent.options.length > 0 && iCount < frmScreen.cboPaymentEvent.options.length)
			{
				if (scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, iCount, "P") == true)
					frmScreen.cboPaymentEvent.options.remove(iCount);
				else iCount++;
			}		
						
			<% /* Default Payment Type to Awaiting Payment (validation type AW) */ %>
			scScreenFunctions.SetComboOnValidationType(frmScreen, "cboPaymentEvent", "AW");
			
			<% /* MAR1652  Filter the Payment Type combo and set it writable */ %>
			FilterPaymentTypeCombo(frmScreen.cboPaymentEvent.value);
			
			scScreenFunctions.SetFieldToWritable(frmScreen,"cboPaymentEvent");
		}
		break;
		<% /* MAR1652 Add cheque processing */ %>						
		case XML.IsInComboValidationList(document,"FeePaymentMethod", m_sPayMethod, ["CH"]):
		{
			<% /* Payment Method has validation type of CH (Cheque) */ %>
			<% /* Disable Savings Account No, Date of Refund, Amount of Refund */ %>
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtSavingsAccNo");
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtRefundDate");
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtRefundAmount");
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtTransRef");
			
			<% /* Check if this user has the authority to change the payment type */ %>
			if (parseInt(m_sUserRole) >= parseInt(m_sMinUserRole))
			{
				<% /* Repopulate the Payment Type combo with all values */ %>
				XML.PopulateComboFromXML(document, frmScreen.cboPaymentEvent,m_PaymentEventXML,true); 
			
				<% /* Default Payment Type to Awaiting Payment (validation type AW) */ %>
				scScreenFunctions.SetComboOnValidationType(frmScreen, "cboPaymentEvent", "AW");
			
				<% /* Filter the Payment Type combo */ %>
				FilterPaymentTypeCombo(frmScreen.cboPaymentEvent.value);		
					
				scScreenFunctions.SetFieldToWritable(frmScreen,"cboPaymentEvent");			
			}
			else
			{
				<% /* Repopulate the Payment Type combo with all values */ %>
				XML.PopulateComboFromXML(document, frmScreen.cboPaymentEvent,m_PaymentEventXML,true); 

				<% /* Default Payment Type to Awaiting Payment (validation type AW) */ %>
				scScreenFunctions.SetComboOnValidationType(frmScreen, "cboPaymentEvent", "AW");
					
				scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboPaymentEvent");			
			}
			
		}
		break;	
		case XML.IsInComboValidationList(document,"FeePaymentMethod", m_sPayMethod, ["SAV"]):
		{
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtSavingsAccNo");
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtRefundDate");
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtRefundAmount");
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtTransRef");
			
			<% /* Repopulate the Payment Type combo with all values */ %>
			XML.PopulateComboFromXML(document, frmScreen.cboPaymentEvent,m_PaymentEventXML,true); 
			
			<% /* Filter the Payment Type combo */ %>
			FilterPaymentTypeCombo(m_sPaymentEvent);		

			<% /* Remove Awaiting Payment option (validation type AW) from Payment Type combo */ %>
			var iCount = 0;
			for (iCount = frmScreen.cboPaymentEvent.length - 1; iCount > 0; iCount--)
			{
				if (scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, iCount, "AW") )
					frmScreen.cboPaymentEvent.remove(iCount);
			}	

			<% /* MAR1652 Default the Payment Type to <SELECT and enable the combo. */ %>
			frmScreen.cboPaymentEvent.value = "";

			scScreenFunctions.SetFieldToWritable(frmScreen,"cboPaymentEvent");
			
		}
		break;
		<% /* MAR1517 */ %>
		case XML.IsInComboValidationList(document,"FeePaymentMethod", m_sPayMethod, ["NTR"]):
		{
			<% /* Set Payment Type to correspond */ %>
			frmScreen.cboPaymentEvent.value = m_sNTRPaymentType;
			frmScreen.txtPaymentAmount.value = "0";
		}
		break;
		case XML.IsInComboValidationList(document,"FeePaymentMethod", m_sPayMethod, ["AR"]):
		{
			<% /* Set Payment Type to correspond */ %>
			frmScreen.cboPaymentEvent.value = m_sARPaymentType;
			frmScreen.txtPaymentAmount.value = "0";
		}
		break;
		default:
		{
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtSavingsAccNo");
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtRefundDate");
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtRefundAmount");
			
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtTransRef");
			
			<% /* MAR1652 */ %>
			
			<% /* Repopulate the Payment Type combo with all values */ %>
			XML.PopulateComboFromXML(document, frmScreen.cboPaymentEvent,m_PaymentEventXML,true);  
			
			<% /* Filter the Payment Type combo */ %>
			FilterPaymentTypeCombo(m_sPaymentEvent);
			
			<% /* Remove Awaiting Payment option (validation type AW) from Payment Type combo */ %>
			var iCount = 0;
			for (iCount = frmScreen.cboPaymentEvent.length - 1; iCount > 0; iCount--)
			{
				if (scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, iCount, "AW") )
					frmScreen.cboPaymentEvent.remove(iCount);
			}				
						
			<% /* Default to <SELECT> */ %>
			frmScreen.cboPaymentEvent.value = "";

		}
		}		
	}
	else
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtSavingsAccNo");
			
}

function DisableChequeNumber()
{

	if (frmScreen.cboFeePaymentMethod.disabled == false )
	{
		var tagOption = frmScreen.cboFeePaymentMethod.options(frmScreen.cboFeePaymentMethod.selectedIndex) ;
		var sValidationId = tagOption.getAttribute("ValidationType0");
		if (scScreenFunctions.IsValidationType(frmScreen.cboFeePaymentMethod,"CH"))
		{
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtChequeNumber");
		}
		else
		{	
			frmScreen.txtChequeNumber.value="";
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtChequeNumber");
		}
	}
}

function DisableFeePaymentMethod()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();			
	var sPaymentEvent = frmScreen.cboPaymentEvent.value ;
	<% /* MAR462 comparison to null,' ','' added */ %>
	if ((sPaymentEvent != null) && (sPaymentEvent != " ") && (sPaymentEvent != ""))
	{
		<% /* AQR 3524 disable payment method on Addition or Rebate */ %>
		<% /* BM0219 MDC 02/01/2003 */ %>
		<% /* MAR28  Disable Payment Method for Paid On Previous Application */ %>
		<% /* MAR1517  Disable Payment Method for RFV, NTR and AR */ %>

		if (sPaymentEvent == "ADDITION" || 
			scScreenFunctions.IsValidationType(frmScreen.cboPaymentEvent,"RA") || 
			scScreenFunctions.IsValidationType(frmScreen.cboPaymentEvent,"RFV") ||
			XML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["POPA"]))

		<% /* BM0219 MDC 02/01/2003 - End */ %>
		{
			<% /* make payment method non mandatory, clear and  disable it */ %>
			frmScreen.cboFeePaymentMethod.setAttribute("required", "false");
			frmScreen.cboFeePaymentMethod.parentElement.parentElement.style.color="black";			
			scScreenFunctions.SetFieldToDisabled(frmScreen,"cboFeePaymentMethod");
			m_sPayMethod = "";
		}
		else
		{	
			frmScreen.cboFeePaymentMethod.value = parseInt(m_sPayMethod);
			frmScreen.cboFeePaymentMethod.setAttribute("required", "true");
			frmScreen.cboFeePaymentMethod.parentElement.parentElement.style.color="red";
		
			<% /* MAR28 Do not enable Payment Method if card payment has been made */ %>
			if ((scScreenFunctions.IsValidationType(frmScreen.cboPaymentEvent,"RFV")) || 
				(scScreenFunctions.IsValidationType(frmScreen.cboPaymentEvent,"NTR")) ||
				(scScreenFunctions.IsValidationType(frmScreen.cboPaymentEvent,"AR")))
			{
				scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboFeePaymentMethod");
			}
			else if (m_bCardPaymentMade == false)
			{
				scScreenFunctions.SetFieldToWritable(frmScreen,"cboFeePaymentMethod");
			}
		}
		DisableChequeNumber();	
	}
	else
	{
		DisableChequeNumber();	
	}
}

function frmScreen.txtPaymentAmount.onchange()
{
	ValidatePaymentAmount();
}

function ValidatePaymentAmount()
{
	if(isNaN(frmScreen.txtPaymentAmount.value))
	{
		alert('Payment Amount should be numeric.');
		return false ;
	}
	
	var nPaymentAmount = parseFloat(frmScreen.txtPaymentAmount.value) ;
	var nAmountOutStanding = parseFloat(frmScreen.txtAmountOutStanding.value) ;
	
	<% /* SYS1891 - Payment Amount should be greater than zero */ %>
	<% /* MAR28 Except for 'Not To Be Refunded', 'Refund For Valuation' and 'Already Refunded' events */ %>
	if((nPaymentAmount <= 0) && (frmScreen.cboPaymentEvent.value != m_sNTRPaymentType) &&
								(frmScreen.cboPaymentEvent.value != m_sRFVPaymentType) &&
							    (frmScreen.cboPaymentEvent.value != m_sARPaymentType))
	{
		frmScreen.txtPaymentAmount.focus();
		window.alert("Amount Paid should be greater than zero. ");
		return false;
	}
	<% /*
	if(m_sMetaAction == "Add")
	{	
		if(nPaymentAmount > nAmountOutStanding)
		{
			window.alert("Amount Paid cannot be more than the Amount Outstanding");
			frmScreen.txtPaymentAmount.focus();
			return false;
		}
	}
	else if (m_sMetaAction == "Edit")
	{
		if(nPaymentAmount > nAmountOutStanding + parseFloat(m_aArgArray[9]))
		{
			window.alert("Amount Paid cannot be more than the Amount Outstanding");
			frmScreen.txtPaymentAmount.focus();
			return false;
		}
	}
	*/ %>	
	return true;	
}

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sArguments = window.dialogArguments ;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];	
	m_aArgArray				= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	m_sApplicationNumber	= m_aArgArray[0];
	m_sFeeTypeDesc			= m_aArgArray[1];
	m_sAmount				= m_aArgArray[2];
	m_sAmountOutStanding	= m_aArgArray[3];
	m_sFeeTypeComboSeq		= m_aArgArray[4];
	m_sFeeTypeRecordSeq		= m_aArgArray[5];
	m_sUserId				= m_aArgArray[6][0];
	m_sUserRole             = m_aArgArray[6][5];    <% /* MAR1652 */ %>
	m_sMetaAction			= m_aArgArray[7];
		
    m_sPreviousPayment      = m_aArgArray[9];
    
    if (isNaN(m_sPreviousPayment))
		m_sPreviousPayment = 0;
    
	if(m_sMetaAction == "Edit")
	{
		m_sFeePaymentSeqNo		= m_aArgArray[8];
		m_sAmountOfPayment		= m_aArgArray[9];
		m_sPayDate				= m_aArgArray[10];
		m_sChequeNumber			= m_aArgArray[11];
		m_sPayMethod			= m_aArgArray[12];
		m_sPaymentEvent			= m_aArgArray[13];
		m_sPaymentEventDesc		= m_aArgArray[14];
		m_sRefundDate			= m_aArgArray[15];
		m_sNotes				= m_aArgArray[16];
		m_sRefundAmount			= m_aArgArray[17];
		m_sSavingsAccountNumber = m_aArgArray[18];   // MAR28
		m_sTransactionReference = m_aArgArray[19];   // MAR28
		m_sResultCode           = m_aArgArray[20];   // MAR28
	}
	else
	{
		<% /* default amount to pay */ %>
		m_sAmountOfPayment	= m_sAmountOutStanding;
		
		<% /* MO - BMIDS00807 */ %>
		<% /* var dtCurrentDte	= scScreenFunctions.GetSystemDate(); */ %>
		var dtCurrentDte	= scScreenFunctions.GetAppServerDate();
		m_sDate				= scScreenFunctions.DateToString(dtCurrentDte);
	}
	
	<% /* MAR1652  Get User Role from Global parameter */ %>
	var GlobalXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_sMinUserRole = GlobalXML.GetGlobalParameterAmount(document,'TakeChqPaymentAuthorityLevel');		
	
	SetMasks();  
	Validation_Init();
	if (InitialiseScreen())
	{
		scScreenFunctions.SetFocusToFirstField(frmScreen);
		<%- /* AQR SYS 3524 */ %>
		DisableFeePaymentMethod();
	
		<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
		ClientPopulateScreen();
	}
}

function frmScreen.cboPaymentEvent.onchange()
{
	<% /* MAR462 comparison to null,' ','' added */ %>
	<% /* MAR1517  Save new Payment event */ %>
	m_sPaymentEvent = frmScreen.cboPaymentEvent.value ;
	if ((m_sPaymentEvent != null) && (m_sPaymentEvent != " ") && (m_sPaymentEvent != ""))
	{
		var tagOption = frmScreen.cboPaymentEvent.options(frmScreen.cboPaymentEvent.selectedIndex) ;
		var sValidationId = tagOption.getAttribute("ValidationType0");
		<% /* MAR1517  Add processing for RFV/NTR/AR */ %>
		if(sValidationId  == 'RFV')
		{
			<% /* Set Payment Method */ %>
			frmScreen.cboFeePaymentMethod.value = " ";
			scScreenFunctions.SetFieldToDisabled(frmScreen,"cboFeePaymentMethod");
			m_sPayMethod = " ";
		}
		else if (sValidationId  == 'NTR')
		{
			frmScreen.cboFeePaymentMethod.value = m_sNTRPaymentMethod;
			scScreenFunctions.SetFieldToWritable(frmScreen,"cboFeePaymentMethod");

			m_sPayMethod = m_sNTRPaymentMethod;
		}
		else if (sValidationId  == 'AR')
		{
			frmScreen.cboFeePaymentMethod.value = m_sARPaymentMethod;
			scScreenFunctions.SetFieldToWritable(frmScreen,"cboFeePaymentMethod");

			m_sPayMethod = m_sARPaymentMethod;
		}
		else if(sValidationId != 'P')
		{
			frmScreen.txtRefundDate.value = '' ;
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtRefundDate");
		
			frmScreen.txtRefundAmount.value = '' ;
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtRefundAmount");
		}
		else
		{
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtRefundDate");
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtRefundAmount");
		}
		<%- /* AQR SYS 3524 */ %>
		DisableFeePaymentMethod();
	}
	else
	{
		frmScreen.txtRefundDate.value = '' ;
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtRefundDate");
		frmScreen.txtRefundAmount.value = '' ;
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtRefundAmount");
		DisableFeePaymentMethod();
	}
}  

<% /**** Functions *****/ %>
function InitialiseScreen()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();			
	var nOption = 0;

	if (PopulateCombos())
	{
		frmScreen.txtFeeType.value = m_sFeeTypeDesc;
		frmScreen.txtAmountOutStanding.value = m_sAmountOutStanding;
		frmScreen.txtTotalAmountDue.value = m_sAmount;
	
		if(m_sMetaAction == "Edit")
		{
			frmScreen.txtPaymentAmount.value = m_sAmountOfPayment;
			frmScreen.txtIssueDate.value = m_sPayDate;
			frmScreen.cboFeePaymentMethod.value = parseInt(m_sPayMethod);		
			frmScreen.txtChequeNumber.value = m_sChequeNumber;
			frmScreen.txtRefundDate.value = m_sRefundDate ;
			frmScreen.txtNotes.value = m_sNotes ;
			frmScreen.txtRefundAmount.value = m_sRefundAmount ;
			frmScreen.txtSavingsAccNo.value = m_sSavingsAccountNumber;  // MAR28
			frmScreen.txtTransRef.value = m_sTransactionReference;      // MAR28
			frmScreen.txtResultCode.value = m_sResultCode;              // MAR28
		
			<% /* Populate combo PaymentEvent  */ %>
			if(m_sPaymentEventDesc == 'Addition') 
				frmScreen.cboPaymentEvent.value = 'ADDITION' ;
			else 
				frmScreen.cboPaymentEvent.value = m_sPaymentEvent ;		
		
			<% /* MAR1797 Read Payment Event */ %>
			var sPaymentEvent = frmScreen.cboPaymentEvent.value ;

			<% /* Disable the Savings A/C No, Date of Refund and Amount of Refund fields if Payment Method = Credit Card*/ %>
			if ((m_sPayMethod != null) && (m_sPayMethod != " "))
			{
				if(XML.IsInComboValidationList(document,"FeePaymentMethod", m_sPayMethod, ["CD"]))
				{
					scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtSavingsAccNo");
					scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtRefundDate");      <%/* MAR1652 */%>
					scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtRefundAmount");    <%/* MAR1652 */%>

					<% /* MAR1652 Remove Payment option (validation type P) from Payment Type combo */ %>
					<% /* MAR1797 Only do this if the Credit Card payment has not aleady been paid */ %>
					
					if ((sPaymentEvent != null) && (sPaymentEvent != " ") && (XML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["P"]) == false))
					{
						var iCount = 0;
						while(frmScreen.cboPaymentEvent.options.length > 0 && iCount < frmScreen.cboPaymentEvent.options.length)
						{
							if (scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, iCount, "P") == true)
								frmScreen.cboPaymentEvent.options.remove(iCount);
							else iCount++;
						}
					}
				}	
				else
				{
					scScreenFunctions.SetFieldToWritable(frmScreen,"txtSavingsAccNo");
					scScreenFunctions.SetFieldToWritable(frmScreen,"txtRefundDate");      <%/* MAR1652 */%>
					scScreenFunctions.SetFieldToWritable(frmScreen,"txtRefundAmount");    <%/* MAR1652 */%>
				}

				<% // MAR1342 - Peter Edney - 03/03/2006 %>
				if(XML.IsInComboValidationList(document,"FeePaymentMethod", m_sPayMethod, ["SAV"]))
					scScreenFunctions.SetFieldToWritable(frmScreen,"txtTransRef");
				else
					scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTransRef");
					
				<% /* MAR1652 If method = Cheque, disable the Payment Type combo depending on authority */ %>					
					
				if(XML.IsInComboValidationList(document,"FeePaymentMethod", m_sPayMethod, ["CH"]))
				{
					if (parseInt(m_sUserRole) >= parseInt(m_sMinUserRole))
					{
						scScreenFunctions.SetFieldToWritable(frmScreen,"cboPaymentEvent");	
					}		
					else
					{
						scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboPaymentEvent");			
					}			
				}
			}
			else
			{
				scScreenFunctions.SetFieldToWritable(frmScreen,"txtSavingsAccNo");
				scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTransRef");
			}
		
			<% /* If Payment Event = Payment, enable Refund fields. */ %>
			if ((sPaymentEvent != null) && (sPaymentEvent != " "))
			{
				if(XML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["P"]))
				{
					scScreenFunctions.SetFieldToWritable(frmScreen,"txtRefundDate");
					scScreenFunctions.SetFieldToWritable(frmScreen,"txtRefundAmount");
				
					<% /* If Payment has already been made by card, set screen to Read Only */ %>
					if ((m_sPayMethod != null) && (m_sPayMethod != " "))
					{
						m_bCardPaymentMade = (XML.IsInComboValidationList(document,"FeePaymentMethod", m_sPayMethod, ["CD"])) &&
							(m_sTransactionReference != " ")
					
						if(m_bCardPaymentMade == true)
						{
							scScreenFunctions.SetScreenToReadOnly(frmScreen);
						}
					}
				}
				<% /* MAR28 If Payment Event = Refund For Valuation, default Method, Type and Amount fields. */ %>
				<% /* MAR1517 / MAR1652 Restrict Payment Type combo depending on payment event */ %>
				if ((XML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["RFV"])) ||
					(XML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["NTR"])) ||
					(XML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["AR"])))
				{
					<% /* Remove all options except those with validation types RFV, NTR and AR from combos */ %>
					FilterPaymentTypeCombo(sPaymentEvent);
					
					frmScreen.cboPaymentEvent.value = m_sPaymentEvent;
					
					DisableAllFields();
				}
				else
				{
					<% /* Remove options with validation types RFV, NTR and AR from combo */ %>
					
					FilterPaymentTypeCombo(m_sPaymentEvent);
					
					frmScreen.cboPaymentEvent.value = m_sPaymentEvent;
				}
			}
			else
			{
				scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtRefundDate");
				scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtRefundAmount");
			}	
		}
		else
		{
			frmScreen.txtPaymentAmount.value = m_sAmountOfPayment;
			frmScreen.txtIssueDate.value = m_sDate;	
			frmScreen.cboPaymentEvent.value = m_sPaymentId  ;
			
			FilterPaymentTypeCombo(m_sPaymentEvent);
		}
		return(true);
	} 
	else
	{
		alert('Failed to populate combos');
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnOk.disabled = true ;
		return(false);
	}
}

function PopulateCombos()
{
	var XMLCombos = null;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var sDeductionId, sReturnOfFundsId ;

	var sGroupList = new Array("FeePaymentMethod", "PaymentEvent");
	if(XML.GetComboLists(document,sGroupList))
	{
		XMLCombos = XML.GetComboListXML("FeePaymentMethod");
		
		<% /* MAR28 save the values for Not To Be Refunded, Refund For Valuation and Already Refunded */ %>
		m_sNTRPaymentMethod = XML.GetComboIdForValidation("FeePaymentMethod", "NTR", null, document) ;
		m_sARPaymentMethod = XML.GetComboIdForValidation("FeePaymentMethod", "AR", null, document) ;
		
		m_sNTRPaymentType = XML.GetComboIdForValidation("PaymentEvent", "NTR", null, document) ;
		m_sRFVPaymentType = XML.GetComboIdForValidation("PaymentEvent", "RFV", null, document) ;
		m_sARPaymentType = XML.GetComboIdForValidation("PaymentEvent", "AR", null, document) ;

		var blnSuccess = true;
		
		<% /* Fee Payment Method */ %>
		blnSuccess = XML.PopulateComboFromXML(document, frmScreen.cboFeePaymentMethod,XMLCombos,true);    
		
		
		<% /* Payment Event */ %>
		
		m_PaymentEventXML = XML.GetComboListXML("PaymentEvent");	
		<% /* GD BM0372 - Get where validation types = 'RA' and 'O' */ %>
		m_sRebateAdditionId = GetOmigaRebateAdditionValueId(m_PaymentEventXML);
		if(m_sRebateAdditionId == -1)
		{
			alert("No PaymentEvent combo entry exists with validation type of 'O' AND 'RA'");
			return(false);
		}
		
		m_sPaymentId		= XML.GetComboIdForValidation("PaymentEvent", "P", m_PaymentEventXML) ;
		<% /* SR 10-07-01 : SYS2412 Do not display 'Deduction' and 'ReturnOfFunds' in the combo */ %>
		sDeductionId		= XML.GetComboIdForValidation("PaymentEvent", "D", m_PaymentEventXML) ;
		sReturnOfFundsId	= XML.GetComboIdForValidation("PaymentEvent", "F", m_PaymentEventXML) ;
		
		
		var sCondition = ".//LISTENTRY[VALUEID='" + sDeductionId + "']" ;
		var xmlNode = m_PaymentEventXML.selectSingleNode(sCondition) ;
		if(xmlNode != null) m_PaymentEventXML.firstChild.removeChild(xmlNode) ;
		
		var sCondition = ".//LISTENTRY[VALUEID='" + sReturnOfFundsId + "']" ;
		var xmlNode = m_PaymentEventXML.selectSingleNode(sCondition) ;
		if(xmlNode != null) m_PaymentEventXML.firstChild.removeChild(xmlNode) ;
				
		<%/* Display the option Rebate/Addition as two different options */%>	

		var sCondition = ".//LISTENTRY[VALUEID='" + m_sRebateAdditionId + "']" ;
		var xmlNode = m_PaymentEventXML.selectSingleNode(sCondition) ;
		var xmlValueNameNode = xmlNode.selectSingleNode('.//VALUENAME') ;
		xmlValueNameNode.text = 'Rebate' ;
		
		
		var xmlNewNode = xmlNode.cloneNode(true);
		var xmlValueNameNode = xmlNewNode.selectSingleNode('.//VALUENAME') ;
		xmlValueNameNode.text = 'Addition' ;
		var xmlValueNode = xmlNewNode.selectSingleNode('.//VALUEID') ;
		xmlValueNode.text = 'ADDITION' ;
		
		m_PaymentEventXML.firstChild.appendChild(xmlNewNode) ;
		
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document, frmScreen.cboPaymentEvent,m_PaymentEventXML,true);

		return(blnSuccess);
	} 
	else 
		return(false);
}

function SaveFeeTypePayment()
{
	var bContinue, sPaymentEvent, sAmountPaid ;
	var xmlRequestNode;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	if(m_sMetaAction == "Edit")
		xmlRequestNode = XML.CreateRequestTagFromArray(m_aArgArray[6], "UPDATEFEETYPEPAYMENT");
	else  
		xmlRequestNode = XML.CreateRequestTagFromArray(m_aArgArray[6], "CREATEFEETYPEPAYMENT");
	
	XML.CreateActiveTag("FEEPAYMENT");
	
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("FEETYPE", m_sFeeTypeComboSeq);
		
	if (m_sMetaAction == "Edit")	
		XML.SetAttribute("PAYMENTSEQUENCENUMBER", m_sFeePaymentSeqNo);
	
	sPaymentEvent = frmScreen.cboPaymentEvent.value ;
	sAmountPaid	  = frmScreen.txtPaymentAmount.value ;
	if(sPaymentEvent == 'ADDITION')
	{
		sPaymentEvent = m_sRebateAdditionId ;
		sAmountPaid = (-1) * parseInt(frmScreen.txtPaymentAmount.value) ;
	} 
	
	XML.SetAttribute("PAYMENTEVENT", sPaymentEvent);	
	XML.SetAttribute("AMOUNTPAID", sAmountPaid);
	XML.SetAttribute("DATEOFPAYMENT", frmScreen.txtIssueDate.value);
	XML.SetAttribute("REFUNDDATE", frmScreen.txtRefundDate.value);
	XML.SetAttribute("REFUNDAMOUNT", frmScreen.txtRefundAmount.value);
	XML.SetAttribute("NOTES", frmScreen.txtNotes.value);
	XML.SetAttribute("FEETYPESEQUENCENUMBER", m_sFeeTypeRecordSeq);
	
	XML.ActiveTag = xmlRequestNode ;
	XML.CreateActiveTag("PAYMENTRECORD");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	
	if(m_sMetaAction == "Edit")	
		XML.SetAttribute("PAYMENTSEQUENCENUMBER", m_sFeePaymentSeqNo);
		
	XML.SetAttribute("AMOUNT", sAmountPaid);
	XML.SetAttribute("CHEQUENUMBER",frmScreen.txtChequeNumber.value);
	XML.SetAttribute("USERID", m_sUserId);
	<%- /* AQR SYS 3524 */ %>
	XML.SetAttribute("PAYMENTMETHOD", m_sPayMethod);
	XML.SetAttribute("SAVINGSACCOUNTNUMBER", frmScreen.txtSavingsAccNo.value);  // MAR28

	<% // MAR1342 - Peter Edney - 03/03/06 %>
	XML.SetAttribute("TRANSACTIONREFERENCE", frmScreen.txtTransRef.value);  
		
	<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	bContinue = XML.IsResponseOK();
	
	return bContinue;
}
<% /* GD BM0372 START */ %>

function GetOmigaRebateAdditionValueId(XMLCombos)
{
	var TagListLISTENTRY = XMLCombos.selectNodes(".//LISTENTRY");
	var xmlElem;
	var TagListVALIDATIONTYPE;
	var blnValidationFound1, blnValidationFound2;
	var sComboValueId;
	var sValidation1 = "RA";
	var sValidation2 = "O";
	var sValidationType;
	for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
	{	
		xmlElem = TagListLISTENTRY.item(nLoop);
	
		sComboValueId = xmlElem.selectSingleNode("VALUEID").text;
		TagListVALIDATIONTYPE = xmlElem.selectNodes(".//VALIDATIONTYPE")
		blnValidationFound1 = false;
		blnValidationFound2 = false;
		for(var nListLoop = 0; nListLoop < TagListVALIDATIONTYPE.length; nListLoop++)
		{
			sValidationType = TagListVALIDATIONTYPE.item(nListLoop).text;
			if(sValidationType == sValidation1) 
			{
				blnValidationFound1 = true;
			}
			if(sValidationType == sValidation2) 
			{
				blnValidationFound2 = true;
			}
		}
		if (blnValidationFound1 && blnValidationFound2)
		{
			return(sComboValueId);
		}
		
	}
	return(-1); //Error occured - not found	
}
<% /* GD BM0372 END */ %>

<% /* MAR1517 Add function to disable all fields */ %>
function DisableAllFields()
{
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPaymentAmount");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtChequeNumber");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtIssueDate");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtSavingsAccNo");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTransRef");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtResultCode");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtRefundDate");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtRefundAmount");
}

<% /* MAR1652  Add function to filter the entries in Payment Type combo correctly */ %>
function FilterPaymentTypeCombo(sPaymentEvent)
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();			

	if(m_sMetaAction == "Edit")
	{
		if ((sPaymentEvent != null) && (sPaymentEvent != " "))
		{
			<% /* Restrict Payment Type combo depending on payment event */ %>
			if ((XML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["RFV"])) ||
				(XML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["NTR"])) ||
				(XML.IsInComboValidationList(document,"PaymentEvent", sPaymentEvent, ["AR"])))
			{
				<% /* Remove all options except those with validation types RFV, NTR and AR from combos */ %>
					
				nOption = 0;
				while(frmScreen.cboPaymentEvent.options.length > 0 && nOption < frmScreen.cboPaymentEvent.options.length)
				{
					if ((scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, nOption, "RFV") == false) &&
						(scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, nOption, "NTR") == false) &&
						(scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, nOption, "AR") == false))
							
						frmScreen.cboPaymentEvent.options.remove(nOption);
					else nOption++;
				}
					
				frmScreen.cboPaymentEvent.value = sPaymentEvent;
				
				nOption = 0;
				while(frmScreen.cboFeePaymentMethod.options.length > 0 && nOption < frmScreen.cboFeePaymentMethod.options.length)
				{
					if ((scScreenFunctions.IsOptionValidationType(frmScreen.cboFeePaymentMethod, nOption, "NTR") == false) &&
						(scScreenFunctions.IsOptionValidationType(frmScreen.cboFeePaymentMethod, nOption, "AR") == false))
							
						frmScreen.cboFeePaymentMethod.options.remove(nOption);
					else nOption++;
				}
			}
			else
			{
				<% /* Remove options with validation types RFV, NTR and AR from combos */ %>
					
				nOption = 0;
				while(frmScreen.cboPaymentEvent.options.length > 0 && nOption < frmScreen.cboPaymentEvent.options.length)
				{
					if ((scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, nOption, "RFV") == true) ||
						(scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, nOption, "NTR") == true) ||
						(scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, nOption, "AR") == true))							
							
						frmScreen.cboPaymentEvent.options.remove(nOption);
					else nOption++;
				}
					
				frmScreen.cboPaymentEvent.value = sPaymentEvent;
				
				nOption = 0;
				while(frmScreen.cboFeePaymentMethod.options.length > 0 && nOption < frmScreen.cboFeePaymentMethod.options.length)
				{
					if ((scScreenFunctions.IsOptionValidationType(frmScreen.cboFeePaymentMethod, nOption, "NTR") == true) ||
						(scScreenFunctions.IsOptionValidationType(frmScreen.cboFeePaymentMethod, nOption, "AR") == true))							
							
						frmScreen.cboFeePaymentMethod.options.remove(nOption);
					else nOption++;
				}			
			}
		}
			
	}
	else
	{
		<% /* In Add mode, remove options with validation types RFV, NTR, and AR from combos */ %>
		nOption = 0;
			
		while(frmScreen.cboPaymentEvent.options.length > 0 && nOption < frmScreen.cboPaymentEvent.options.length)
		{
			if ((scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, nOption, "RFV") == true) ||
				(scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, nOption, "NTR") == true) ||
				(scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentEvent, nOption, "AR") == true))
				
				frmScreen.cboPaymentEvent.options.remove(nOption);
			else nOption++;
		}
		
		nOption = 0;
			
		while(frmScreen.cboFeePaymentMethod.options.length > 0 && nOption < frmScreen.cboFeePaymentMethod.options.length)
		{
			if ((scScreenFunctions.IsOptionValidationType(frmScreen.cboFeePaymentMethod, nOption, "NTR") == true) ||
				(scScreenFunctions.IsOptionValidationType(frmScreen.cboFeePaymentMethod, nOption, "AR") == true))
				
				frmScreen.cboFeePaymentMethod.options.remove(nOption);
			else nOption++;
		}		
	}
	return(true);
}
-->
</script>
</BODY>
</HTML>

