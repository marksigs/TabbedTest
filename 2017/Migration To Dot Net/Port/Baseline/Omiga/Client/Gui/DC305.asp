<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc305.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Add/Edit Application Credit Card Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
IW		20/02/00	Initial
IW		21/02/00	Removed Cheque Guarantee Card from the list of credit card types
					(This is a known anomally ref. Steve Brunt)
AY		03/04/00	New top menu/scScreenFunctions change
SR		16/05/00	SYS0708 
BG		17/05/00	SYS0752 Removed Tooltips
SR		30/05/00	SYS0708 Enable the field IssueNumber only when CreditCardType is 'Switch' 
		31/05/00			Changed the class of txtExpiryDate to MsgTxt
BG		16/08/00	SYS1280 stop trip to database to validate combo selection and make txtIssueNumber
							a mandatory field if Switch is selected.
CL		05/03/01	SYS1920 Read only functionality added
JLD		10/12/01	SYS2806 completeness check routing							
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
ASu		10/09/2002	BMIDS00420 Reduce Expiry Date field size to 5 characters and add mask validation
HMA     17/09/2003  BM0063     Amend HTML text for radio buttons
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<% /* FORMS */ %>
<form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC300" method="post" action="DC300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC305" method="post" action="DC305.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange">
<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 170px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Are you paying the fees by Credit/Debit Card?
		<span style="TOP: -3px; LEFT: 226px; POSITION: ABSOLUTE">
			<input id="optFeePaymentMethodYes" name="grpFeePaymentMethod" type="radio" value="1"><label for="optFeePaymentMethodYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: -3px; LEFT: 280px; POSITION: ABSOLUTE">
			<input id="optFeePaymentMethodNo" name="grpFeePaymentMethod" type="radio" value="0"><label for="optFeePaymentMethodNo" class="msgLabel">No</label>
		</span> 
	</span>	
	<div id=grpMandatory>
	<span style="TOP: 36px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Card Holder's Name
		<span style="TOP: -3px; LEFT: 230px; POSITION: ABSOLUTE">
			<input id="txtNameOfCardHolders" name="NameOfCardHolders" maxlength="255" style="WIDTH: 300px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 62px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Card Type
		<span style="TOP: -3px; LEFT: 230px; POSITION: ABSOLUTE">
			<select id="cboCardType" name="CardType" style="WIDTH: 100px" class="msgCombo">
			</select>
		</span>
	</span>
	<span style="TOP: 88px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Card Number
		<span style="TOP: -3px; LEFT: 230px; POSITION: ABSOLUTE">
			<input id="txtCardNumber" name="CardNumber" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 114px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Issue Number
		<span style="TOP: -3px; LEFT: 230px; POSITION: ABSOLUTE">
			<input id="txtIssueNumber" name="IssueNumber" maxlength="5" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 140px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Expiry Date
		<span style="TOP: -3px; LEFT: 230px; POSITION: ABSOLUTE">
			<input id="txtExpiryDate" name="ExpiryDate" maxlength="5" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>		
	</div>
</div>
</form>

<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/dc305attribs.asp" -->

<script language="JScript">
<!--		
var m_sMetaAction = null;		
var m_sXMLOnEntry = null;
var m_sSequenceNumber = null;		
var scScreenFunctions;
var m_blnReadOnly = false;
var m_sReadOnly = "";


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	RetrieveContextData();
	sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	FW030SetTitles("Application Credit/Debit Card Details","DC305",scScreenFunctions);

	GetComboLists();
	PopulateScreen();
	SetMasks();
	Validation_Init();
	frmScreen.optFeePaymentMethodNo.onclick();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC305");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	scScreenFunctions.SetFocusToFirstField(frmScreen);			
	
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
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","1325");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	m_sSequenceNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","1")
}

<% /* Populates all combos with their options */ %>
function GetComboLists()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();			
	var sGroups = new Array("CreditCardType");
						
	if (XML.GetComboLists(document, sGroups) == true)
	{
		XML.PopulateCombo(document, frmScreen.cboCardType, "CreditCardType", true);
		// Remove the Cheque Guarantee Card from the Credit Card Types:
		for (iCount = frmScreen.cboCardType.length - 1; iCount > -1; iCount--)
		{
			if (frmScreen.cboCardType.item(iCount).text == "Cheque Guarantee")
			{
				frmScreen.cboCardType.remove(iCount);
			}	
		}
	}
	XML = null;
}


<% /* Retrieves the data and sets the screen accordingly */ %>
function PopulateScreen()
{
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("APPLICATIONCREDITCARD");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("APPSEQUENCENUMBER",m_sSequenceNumber );
	XML.RunASP(document,"GetApplicationCreditCard.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);

	if(ErrorReturn[1] == ErrorTypes[0])
	{
		<%/* Record not found */%>
		m_sMetaAction = "Add";
		<% /* SR 16/05/00 - SYS0708 - Set the Radio button to No */	%>		
		frmScreen.optFeePaymentMethodNo.checked = true
	}			
	else if (ErrorReturn[0] == true)
		<%/* Record found */%>
	{
		m_sMetaAction = "Edit";

		<%/* Populate the screen with the details held in the XML */%>
		if(XML.SelectTag(null,"APPLICATIONCREDITCARD") != null)
		{
			scScreenFunctions.SetRadioGroupValue(frmScreen, "grpFeePaymentMethod", XML.GetTagText("FEEPAYMENTMETHOD"));
			frmScreen.txtNameOfCardHolders.value	= XML.GetTagText("NAMEOFCARDHOLDERS");
			frmScreen.cboCardType.value	= XML.GetTagText("CARDTYPE");
			frmScreen.txtCardNumber.value	= XML.GetTagText("CARDNUMBER");
			frmScreen.txtIssueNumber.value = XML.GetTagText("ISSUENUMBER");
			frmScreen.txtExpiryDate.value = XML.GetTagText("EXPIRYDATE");
			m_sSequenceNumber = XML.GetTagText("APPSEQUENCENUMBER");
		}
	}
	
	if (scScreenFunctions.IsValidationType(frmScreen.cboCardType,"S"))
		{
			frmScreen.txtIssueNumber.disabled = false;
		}	
		else
		{
			scScreenFunctions.SetFieldToDisabled(frmScreen, "txtIssueNumber");
		}
		
	ErrorTypes = null;
	ErrorReturn = null;
}

function IsSwitchCard()
{
	var ValidationList = new Array(1);
	var bReturnVal ;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	ValidationList[0] = "S";
	bReturnVal = XML.IsInComboValidationList(document,"CreditCardType", frmScreen.cboCardType.value, ValidationList);
	ValidationList = null;
	XML = null;
	
	return bReturnVal ;
}

<% /* Events */  %>
function frmScreen.optFeePaymentMethodYes.onclick()
{
	if(frmScreen.optFeePaymentMethodYes.checked == true)
	{
		scScreenFunctions.EnableCollection(grpMandatory)
		frmScreen.txtNameOfCardHolders.focus() ;
	}
}

function frmScreen.optFeePaymentMethodNo.onclick()
{
	if(frmScreen.optFeePaymentMethodNo.checked == true)
	{
		<% /* SR 16/05/00 - SYS0708 - Clear all the other fields and set them to
							read only (do not hide them)  */
		%>	
		scScreenFunctions.ClearCollection(grpMandatory)
		scScreenFunctions.DisableCollection(grpMandatory)
	}
}

<% /* SR : 30/05/00 - SYS0708 : If the card type selected is not 'Switch', 
					  empty the field Issue Number and disable it  */ 
%>
function frmScreen.cboCardType.onchange()
<%/* BG 16/08/00 SYS1280 stop trip to database to validate combo selection and make txtIssueNumber
				 a mandatory field if Switch is selected*/%>
{	
	if (scScreenFunctions.IsValidationType(frmScreen.cboCardType,"S"))
	<%/* BG SYS1280 This function previously called IsSwitch function above*/%>
	{
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtIssueNumber");
		frmScreen.txtIssueNumber.setAttribute("required", "true");
		frmScreen.txtIssueNumber.parentElement.parentElement.style.color = "red";	
	}
	else
	{
		frmScreen.txtIssueNumber.value = '';
		scScreenFunctions.SetFieldToDisabled(frmScreen, "txtIssueNumber");
		frmScreen.txtIssueNumber.parentElement.parentElement.style.color = "black";	
	}
}

function btnSubmit.onclick()
{
	if (frmScreen.onsubmit() == true)
	{
		if (CommitScreen() == true)
		{
			if(scScreenFunctions.CompletenessCheckRouting(window))
				frmToGN300.submit();
			else
				frmToMN060.submit();
		}
	}
}
		
function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmToDC300.submit();
}

<% /* Commits the screen data to the database, either by a create or update */ %>
function CommitScreen()
{
	var bOK = false;			
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var TagRequestType = null;
						
	<% /* <REQUEST USERID=? USERTYPE=? UNIT=?> */ %>
	if(m_sMetaAction == "Add")
	{
		TagRequestType = XML.CreateRequestTag(window, "CREATE");
	}
	else
	{
		TagRequestType = XML.CreateRequestTag(window, "UPDATE");
	}

	if(TagRequestType != null)
	{
		XML.CreateActiveTag("APPLICATIONCREDITCARD");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("APPSEQUENCENUMBER",m_sSequenceNumber );
		XML.CreateTag("FEEPAYMENTMETHOD",scScreenFunctions.GetRadioGroupValue(frmScreen,"grpFeePaymentMethod"));
		XML.CreateTag("NAMEOFCARDHOLDERS",frmScreen.txtNameOfCardHolders.value);
		XML.CreateTag("CARDTYPE",frmScreen.cboCardType.value);
		XML.CreateTag("CARDNUMBER",frmScreen.txtCardNumber.value);
		XML.CreateTag("ISSUENUMBER",frmScreen.txtIssueNumber.value);
		XML.CreateTag("EXPIRYDATE",frmScreen.txtExpiryDate.value);

		if(m_sMetaAction == "Add")
		{
			// 			XML.RunASP(document, "CreateApplicationCreditCard.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "CreateApplicationCreditCard.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

		}
		else
		{
			// 			XML.RunASP(document, "UpdateApplicationCreditCard.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "UpdateApplicationCreditCard.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

		}
		bOK = XML.IsResponseOK();
	}
	else
	{
		alert("CommitScreen - Invalid MetaAction");
	}
	return bOK;
}

-->
</script>
</body>
</html>


