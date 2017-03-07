<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<% /*
Workfile:      pp120.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Fee Payment Interface  THIS IS A POP-UP SCREEN
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:

Prog	Date		Description
HMA		16/08/05	Created (MAR28)
HMA     05/10/05    MAR28  Use CardPaymentType combo for Card Type. Remove Cancel button.
                    MAR75  Changes for Gateway call.
SD		25/10/05	MAR227 Warning message if Ok button is clicked before Send Details
MV		27/10/2005	MAR287 frmScreen.btnSendDetails.onclick()
HMA     01/11/2005  MAR287 Further changes to Date Validation
SD		10/11/2005	MAR227 Reopened, removing duplicate frmScreen.btnOK.onclick
HMA     16/03/2006  MAR1421 Send amount in pence.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM History:

Prog	Date		Description
SAB		06/04/2006	EP323 Disable the 'Send Details' button
JD		17/08/2006	EP1083/MAR1844 Change the way result code F is handled so that is is treated the same was as D
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Fee Payment Interface   <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark    validate ="onchange"><!--style="VISIBILITY: hidden"> -->

<div id = "divAmount" style="HEIGHT: 33px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 40px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Payment Amount
		<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
			<input id="txtPaymentAmount" maxlength="10" style="POSITION: absolute; WIDTH: 80px" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	<span style="LEFT: 246px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Card Type
		<span style="LEFT: 64px; POSITION: absolute; TOP: -3px">
			<select id="cboCardType" style="WIDTH: 100px" class="msgCombo"></select>
		</span>
	</span>
</div>
<div id = "divDetails" style="HEIGHT: 230px; LEFT: 10px; POSITION: absolute; TOP: 50px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 40px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Card Account Number
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtCardAccountNumber" maxlength="19" style="POSITION: absolute; WIDTH: 200px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 40px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Card Verification Number
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtCardVerNumber" maxlength="4" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 40px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Card Verification Method
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<select id="cboCardVerificationMethod" style="WIDTH: 200px" class="msgCombo"></select>
		</span>
	</span>

	<span style="LEFT: 40px; POSITION: absolute; TOP: 85px" class="msgLabel">
		Card Start Date
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtCardStartDate" maxlength="5" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 40px; POSITION: absolute; TOP: 110px" class="msgLabel">
		Card Expiry Date
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtCardExpiryDate" maxlength="5" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 40px; POSITION: absolute; TOP: 135px" class="msgLabel">
		Card Issue
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtCardIssue" maxlength="2" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 40px; POSITION: absolute; TOP: 160px" class="msgLabel">
		Card Billing Address
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtCardBillAddress" maxlength="5" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 40px; POSITION: absolute; TOP: 185px" class="msgLabel">
		Card Billing Post Code
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtCardBillPostCode" maxlength="5" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 340px; POSITION: absolute; TOP: 200px">
		<% /* EP323 */ %>
		<input id="btnSendDetails" value="Send Details" type="button" style="WIDTH: 70px" class="msgButton" disabled>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 235px">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
	</span>		
</div>

</form>

<!-- File containing field attributes -->
<!--  #include FILE="attribs/pp120attribs.asp" -->

<!-- Specify Code Here -->
<script language="JScript">
<!--
var m_sPaymentAmount = null;
var m_sApplicationNumber = null;
var m_sRequestAttributes = null;
var m_sUserID = null;

var scClientScreenFunctions;
function window.onload()
{
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	scScreenFunctions.ShowCollection(frmScreen);
	m_sRequestAttributes = sArgArray[0];
	m_sApplicationNumber = sArgArray[1];
	m_sPaymentAmount = sArgArray[2];
	m_sUserID = sArgArray[3];
	
	SetMasks();
	Validation_Init();

	PopulateScreen();
		
	window.returnValue = null;   // return null if no action has been taken
	
	ClientPopulateScreen();
}

function PopulateScreen()
{
	<% /* Set the Payment Amount to the value passed in from PP010 */ %>
	frmScreen.txtPaymentAmount.value = m_sPaymentAmount;
	
	PopulateCombos();
}

function PopulateCombos()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	var XMLCardVerificationMethod = null;
	var XMLCardType = null;
		
	var sGroupList = new Array("CardPaymentType", "CardVerificationMethod");
	if(XML.GetComboLists(document, sGroupList) == true)
	{
		<% /* Populate Card Verification Method with all values from CardVerificationMethod combo */ %>
		XMLCardVerificationMethod = XML.GetComboListXML("CardVerificationMethod");
		XML.PopulateComboFromXML(document, frmScreen.cboCardVerificationMethod, XMLCardVerificationMethod, true);

		<% /* Populate Card Type with all values from CardPaymentType combo */ %>
		XMLCardType = XML.GetComboListXML("CardPaymentType");
		XML.PopulateComboFromXML(document, frmScreen.cboCardType, XMLCardType, true);
	}
	
	XML = null;
	XMLCardType = null;	
	XMLCardVerificationMethod = null;
	sGroupList = null;
}

function frmScreen.btnSendDetails.onclick()
{	
	var sCardStartDate;
	var sCardExpiryDate;
	var bDateOK;
	var sPaymentResult;
	var sReturn = new Array();
	var sTransactionMode;
	
	if (frmScreen.onsubmit())
	{
		<% /* All mandatory fields are present */ %>
		
		<% /* Validate dates */ %>
		sCardStartDate = frmScreen.txtCardStartDate.value;
		<% /* If Start Date is filled in it must be in the form MM/YY */ %>
		bDateOK = (sCardStartDate == "") || (ValidateCardDate(sCardStartDate, ""));
		
		if (bDateOK == false)
		{
			alert("Card Start Date is invalid - please re-enter in the format (MM/YY)");
			frmScreen.txtCardStartDate.focus();
			return;
		}
		
		
		<% // Validate Card Expiry Date %>
		bDateOK = ValidateCardDate(frmScreen.txtCardExpiryDate.value, ">");
		if (bDateOK == false)
		{
			alert("Card Expiry Date is invalid - please re-enter a future date (MM/YY)");
			frmScreen.txtCardExpiryDate.focus();
			return;
		}
		
		<% /* Send details to interface */ %>
	
		<% /* Retrieve the Transaction Mode from Global Parameters */ %>
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		sTransactionMode = XML.GetGlobalParameterString(document, "TransactionMode");

		XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTagFromArray(m_sRequestAttributes, "PAYFEES");
		XML.CreateTag("ApplicationNumber", m_sApplicationNumber);
		XML.CreateTag("UserID", m_sUserID);
		XML.CreateTag("CardBillingAddress", frmScreen.txtCardBillAddress.value);
		XML.CreateTag("CardBillingPostCode", frmScreen.txtCardBillPostCode.value);
		XML.CreateTag("TransactionMode", sTransactionMode);
		XML.CreateActiveTag("CreditCardDetails");
		XML.CreateTag("CardAccountNumber", frmScreen.txtCardAccountNumber.value);
		XML.CreateTag("CardExpiryDate", frmScreen.txtCardExpiryDate.value);
		XML.CreateTag("CardVerificationNumber", frmScreen.txtCardVerNumber.value);
		XML.CreateTag("CardVerificationIndicator", frmScreen.cboCardVerificationMethod.value);
		XML.CreateTag("CardType", frmScreen.cboCardType.value);
		XML.CreateTag("CardIssueNumber", frmScreen.txtCardIssue.value);
		XML.CreateTag("CardStartDate", frmScreen.txtCardStartDate.value);
		XML.CreateTag("PaymentAmount", (m_sPaymentAmount * 100.0));
		XML.CreateActiveTag("Address");
				
		switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}
					
		if(XML.IsResponseOK())
		{
			//var TagNode = XML.SelectTag(null, "TakeCardPaymentResponseDetails")
			sPaymentResult = XML.GetTagText("PaymentResult")
			<% /* RCl MAR1844 If result code A is returned then treat this as sucessful */ %>

			if (sPaymentResult == "A")  <% /* Successful */ %>
			{
				frmScreen.btnSendDetails.disabled = true;
				sReturn[0] = true;
				window.returnValue	= sReturn;
						
				alert("Card Payment successful");
				//scScreenFunctions.MSGAlert("Card Payment successful");
			}
			//else if (sPaymentResult == "D") <% /* Not authorised */ %>
			else if (sPaymentResult == "D" || sPaymentResult == "F" ) <% /* Not authorised */ %> 
			{
				sReturn[0] = false;
				window.returnValue	= sReturn;
						
				alert("Your Payment has not been successfully authorised. Please try again using a different card.");
				//scScreenFunctions.MSGAlert("Your Payment has not been successfully authorised. Please try again using a different card.");
						
			}
			else if (sPaymentResult == "E") <% /* Card details wrong */ %>
			{
				sReturn[0] = false;
				window.returnValue	= sReturn;

				alert("The card details provided are either invalid or incomplete. Please check the data you have entered and try again.");
				//scScreenFunctions.MSGAlert("The card details provided are either invalid or incomplete. Please check the data you have entered and try again.");
			}
			else  <% /* Payment Result is not present - error has occurred */ %>
			{
				sReturn[0] = false;
				window.returnValue	= sReturn;

				var sErrorString = "Error " + XML.GetTagText("NUMBER") + ". " + 
					XML.GetTagText("DESCRIPTION") + "  Source " +  XML.GetTagText("SOURCE")
				alert(sErrorString);
				
					
			}
		}
		else
		{
			sReturn[0] = false;
			window.returnValue	= sReturn;
				
			<% /* Disable Send Details button */ %>
			frmScreen.btnSendDetails.disabled = true;
		}
	}
}

function ValidateCardDate(sDate, sCompare)
{
	<% /* This function returns true if a valid string MM/YY has been entered and 
	      the Date compares correctly to the system date */ %>
	      
	var sMonth;
	var iMonth;
	var sDivider;
	var sYear;
	var iYear;
	var bFormatOK = false;

	<% /* Date will be in the form MM/YY */ %>
	
	sMonth = sDate.substr(0,2);
	iMonth = parseInt(sMonth, 10);
	
	sDivider = sDate.substr(2,1);
	
	sYear = sDate.substr(3,2);
	iYear = parseInt(sYear, 10);

	if ((iMonth > 0) && (iMonth < 13) && (iYear > 0) && (sDivider == "/"))
		bFormatOK = true;

	if (bFormatOK == true)
	{
		<% /* Check that the Expiry Date is later than today		
			  Create a date string using the first of the month */ %>
		if (sCompare != "")
		{
			sDate = "01/" + sMonth + "/20" + sYear;			  

			if (scScreenFunctions.CompareDateStringToSystemDate(sDate,sCompare))
				return true;
			else
				return false;
		}
	}
	else
		return false;
}

<% /* When we return to PP010, the return value will be:
		 NULL if no details have been sent 
		 TRUE if a successful card payment has been made 
		 FALSE if a card payment has failed */ %>

function frmScreen.btnOK.onclick()
{
	var returnVal = 0;
	returnVal = CheckDataSent();
	
	if(returnVal==2)
		return;
	else
		//if(returnVal == 1)
			window.close();
}

function CheckDataSent()
{
	var returnVal=0;
	
	if(IsChanged() && frmScreen.btnSendDetails.disabled == false)
		returnVal = scScreenFunctions.DisplayClientWarning("Pressing Ok before Send Details will delete this information and take you back to the Application Cost screen. Are you sure you want to press OK before you Send details?", "images/MSGBOX03.ICO","Yes","No");

	return returnVal;
}

-->
</script>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/Jscript"></script>
</body>
</html>
