<%@ LANGUAGE="JSCRIPT" %>
<%
/*
Workfile:      PP080.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Return of funds
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		16/02/01	New Screen 
CL		05/03/01	SYS1920 Read only functionality added
MV		21/03/01	SYS2072 Corrected the Return Of Funds Screen Title
SR		06/09/01	SYS2412
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		30/08/2002	BMIDS00373	Modified HTML LayOut and CreateReturnFunds()
PSC		05/11/2002	BMIDS00760  Set payment method to Return Of Funds and Set Payment Status to
                                'Returned'	
PSC		11/11/2002	BMIDS00599	Set txtGrossAdvanceAmount base on the payment not on the whole amount                         
SA		18/11/2002	BMIDS00977  Screen Rules Added
PSC		28/11/2002	BMIDS01099  Amend to send down original payment type to enable processing of 
								incentive returns
PSC		22/08/2003	BM0198		Amend calculation of fee payment amount
HMA     26/09/2003  BM0198      Amend Amount Of Funds Returned field.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :
HM		21/09/2005	MAR49		Consideration of Incomlete status was added at OK button
JD		04/11/2005	MAR414		syntax error correction
HMA     30/11/2005  MAR736      Use TMCompletionsStageID global parameter
LDM		28/06/2006  EP886		Fix isuue of not having any post completion tasks.  
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>

</HEAD>

<BODY>
<% /* Scriptlets */ %>
<% /* BMIDS00977  */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<FORM id="frmToPP050" method="post" action="PP050.asp" STYLE="DISPLAY: none"></FORM>


<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 250px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 615px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Original Gross Advance
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtTotalAdvance" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<span style="TOP: 10px; LEFT: 310px; POSITION: ABSOLUTE" class="msgLabel">
		Remaining Gross Advance Amount
		<span style="TOP: -3px; LEFT: 190px; POSITION: ABSOLUTE">
			<input id="txtGrossAdvanceAmount" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<% /* <span style="TOP: 10px; LEFT: 310px; POSITION: ABSOLUTE" class="msgLabel">
		Amount of Advance 
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtAdvanceAmount" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span> */ %>
	
	<span style="TOP: 44px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Payment Type 
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtPaymentType" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<span style="TOP: 44px; LEFT: 310px; POSITION: ABSOLUTE" class="msgLabel">
		Fees Deducted
		<span style="TOP: -3px; LEFT: 190px; POSITION: ABSOLUTE">
			<input id="txtFeesDeducted" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<% /* <span style="TOP: 44px; LEFT: 310px; POSITION: ABSOLUTE" class="msgLabel">
		Issue Date 
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtIssueDate" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span> */ %>
	
	<span style="TOP: 78px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Payment Method 
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtPaymentMethod" style="WIDTH: 150px" class="msgReadOnly" readonly tabindex="-1"></select>
		</span>
	</span>
	
	<span style="TOP: 78px; LEFT: 310px; POSITION: ABSOLUTE" class="msgLabel">
		Remaining Net Advance Amount
		<span style="TOP: -3px; LEFT: 190px; POSITION: ABSOLUTE">
			<input id="txtNetAdvanceAmount" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	<% /* <span style="TOP: 78px; LEFT: 310px; POSITION: ABSOLUTE" class="msgLabel">
		Completion Date 
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtCompletionDate" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span> */ %>
	
	<span style="TOP: 112px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Payee  
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtPayeeName" style="WIDTH: 150px" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<span style="TOP: 112px; LEFT: 310px; POSITION: ABSOLUTE" class="msgLabel">
		Amount of funds returned  
		<span style="TOP: -3px; LEFT: 190px; POSITION: ABSOLUTE">
			<input id="txtAmountReturned" style="WIDTH: 100px" class="msgTxt">
		</span>
	</span>
	
	<span style="TOP: 146px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Issue Date 
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtIssueDate" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<span style="TOP: 146px; LEFT: 310; POSITION: ABSOLUTE" class="msgLabel">
		Completion Date 
		<span style="TOP: -3px; LEFT: 190px; POSITION: ABSOLUTE">
			<input id="txtCompletionDate" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<span style="TOP: 175px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Notes
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<textarea id="txtNotes" name="Notes" rows="4" style="WIDTH: 467px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>
</div>
</form>

<% /* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 330px; WIDTH: 600px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<!-- #include FILE="attribs/PP080attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;

var m_sApplicationNumber, m_sApplicationFactFindNumber, m_sTotalAdvanceToDate ;
var m_sMaxLeftToPay, m_sPaymentSeqNo, m_sPaymentAmount, m_sPaymentMethod, m_sPaymentType ;
var m_sPayee, m_sPayeeType, m_sActivityId, m_sStageId ;
var m_sIssueDate, m_sPayeeHistorySeqNo, m_sNotes, m_sidXml ;
var m_sUserId, m_sReturnedId ;
var xmlFeePaymentList, xmlLCPList ;
var PayeeHistoryXML = null;
var m_sBalance;
var m_blnReadOnly = false;
var m_iInitialAmountOfRefund ; 
<% /* BMIDS00977 */ %>
var scClientScreenFunctions;

<% /* PSC 28/11/2002 BMIDS01099 */ %>
var m_sPaymentType;
<% /* WP13 MAR19 */ %>
var m_bUpdatePostCompletionTasks = false;

<%/* LDM 28/06/2006 EP886 */%>
var m_sTaskType;

function RetrieveContextData()
{
<%	// TEST
	/* 
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber", "C00078387");
	scScreenFunctions.SetContextParameter(window, "idApplicationFactFindNumber", 1);	
	scScreenFunctions.SetContextParameter(window, "idUserId", "MVenkat");
	scScreenFunctions.SetContextParameter(window,"idProcessingIndicator","1");
	*/
	// END TEST
%>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
	m_sidXml = scScreenFunctions.GetContextParameter(window, "idXml", "");
	m_sUserId = scScreenFunctions.GetContextParameter(window, "idUserId", "");
	m_sStageId = scScreenFunctions.GetContextParameter(window, "idStageId", null);
	m_sActivityId = scScreenFunctions.GetContextParameter(window, "idActivityId", null);
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", "");
	
	if(m_sidXml == "")
	{
		alert('Error : Calcuations data is not not available in the context.');
		return false
	}
	return true ;
}

<% /*** EVENTS ****/  %>
function window.onload()
{
	// BMIDS00977 - Add Screen Rules Processing {
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	//BMIDS00977 }
	var sButtonList = new Array("Submit", "Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Return Of Funds","PP080",scScreenFunctions)
	
	SetMasks(); 
	Validation_Init();
	
	if(!RetrieveContextData()) return ;
	
	InitialiseScreen();	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP080");
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	// BMIDS00977 - Add Screen Rules processing
	ClientPopulateScreen();

}

function btnCancel.onclick()
{
	frmToPP050.submit();
}

function btnSubmit.onclick()
{
	var blnConfirmed ;
	if(ValidateBeforeSave())
	{
		blnConfirmed = confirm('Are you sure?');
		if(!blnConfirmed)
		{
			frmToPP050.submit();
			return ;
		}
		if(CreateReturnOfFunds()) 
		{
			if (m_bUpdatePostCompletionTasks == true)
				<% /* WP13 MAR49 */ %>
				UpdatePostCompletionTasks ();
			frmToPP050.submit();
		}
	}
	frmToPP050.submit();
}

function frmScreen.txtAmountReturned.onchange()
{
	var sAmountReturned = frmScreen.txtAmountReturned.value ;
	
	if(sAmountReturned == "")
	{
		alert('Amount of funds returned cannot be empty.')
		frmScreen.txtAmountReturned.focus();
		return ;
	}
	
	if(isNaN(sAmountReturned))
	{
		alert('This should be a number');
		frmScreen.txtAmountReturned.focus();
		return false;
	}
	
	if(parseFloat(frmScreen.txtAmountReturned.value) > parseFloat(frmScreen.txtNetAdvanceAmount.value))
	{
		alert('Amount of funds returned cannot be greater than the Amount of advance.');	
		frmScreen.txtAmountReturned.focus();
		return false;
	} 
}

<% /**** FUNCTIONS ****/  %>
function InitialiseScreen()
{
	<% // Populate values from context
	%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	XML.LoadXML(m_sidXml);
	xmlFeePaymentList = XML.XMLDocument.selectNodes("//FEEPAYMENT");
	xmlLCPList = XML.XMLDocument.selectNodes("//LOANCOMPONENTPAYMENT");
	
	XML.CreateTagList("CALCULATIONS")
	if(XML.SelectTagListItem(0))
	{
		m_sBalance =  parseFloat(XML.GetAttribute('BALANCE'));
	}
	
	<% /* BM0198 Initialise Amount Of Funds returned to be = Net Advance Amount 
	
	----- Initialise the Amount to the maximum that can be returned, passed from the previous screen PP050 
	var sROFAmount = "";
	XML.ActiveTag = null ;
	var xmlNode = XML.SelectSingleNode("//MAXROFAMOUNT") ;
	sROFAmount =  (xmlNode == null) ? "" : xmlNode.text ;
	frmScreen.txtAmountReturned.value = sROFAmount ; */ %>
	
	<% /* PSC 11/11/2002 BMIDS00599 - Start */ %>
	var xmlNode = XML.SelectSingleNode("//GROSSAMOUNT") ;
	m_sTotalAdvanceToDate = (xmlNode == null) ? "0" : xmlNode.text ;
	<% /* PSC 11/11/2002 BMIDS00599 - End */ %>

	var iTotalPaymentAmount, iFeePaymentAmount ;
	iTotalPaymentAmount = 0 ; 
	iFeePaymentAmount = 0;
		
	XML.ActiveTag = null ;
	XML.CreateTagList("PAYMENTRECORD");
	if(XML.SelectTagListItem(0))
	{
		iTotalPaymentAmount = XML.GetAttribute('AMOUNT');
		if (! isNaN(iTotalPaymentAmount)) iTotalPaymentAmount = parseFloat(iTotalPaymentAmount)
	
		<% /* PSC 22/08/2003 BM0198 - Start */ %>
		var strDeductionValue = XML.GetComboIdForValidation("PaymentEvent","D",null,document);

		XML.ActiveTag = null ;
		XML.CreateTagList("FEEPAYMENT[@PAYMENTEVENT=\"" + strDeductionValue + "\"]");
		<% /* PSC 22/08/2003 BM0198 - End */ %>
		
		for(var iCount = 0; iCount < XML.ActiveTagList.length ; ++iCount)
		{
			XML.SelectTagListItem(iCount);
			iFeePaymentAmount = iFeePaymentAmount + parseFloat(XML.GetAttribute('AMOUNTPAID'));
		}
	}
	else
	{
		alert('Payment details not available in context');
		m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP080");
	}
	
	XML.ActiveTag = null ;
	XML.CreateTagList("PAYMENTRECORD");
	if(XML.SelectTagListItem(0))
	{
		<%/* MV - 30/08/2002 - BMIDS00373 
		Start
		*/%>
		frmScreen.txtTotalAdvance.value			= iTotalPaymentAmount ;
		frmScreen.txtGrossAdvanceAmount.value	= m_sTotalAdvanceToDate;
		frmScreen.txtFeesDeducted.value			= iFeePaymentAmount;
		frmScreen.txtNetAdvanceAmount.value		= m_sTotalAdvanceToDate - iFeePaymentAmount ;
		<%/* End */%> 
		frmScreen.txtAmountReturned.value       = m_sTotalAdvanceToDate - iFeePaymentAmount ;  // BM0198
		m_sPaymentSeqNo							= XML.GetAttribute('PAYMENTSEQUENCENUMBER');
		frmScreen.txtPaymentType.value			= XML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYMENTTYPE_TEXT');
		m_sPaymentType							= XML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYMENTTYPE');
		frmScreen.txtIssueDate.value			= XML.GetTagAttribute('DISBURSEMENTPAYMENT', 'ISSUEDATE');
		frmScreen.txtPaymentMethod.value		= XML.GetAttribute('PAYMENTMETHOD_TEXT');

		<% /* PSC 05/11/2002 BMIDS00760 - Start */ %>
		m_sPaymentMethod = XML.GetComboIdForValidation("PaymentMethod","R",null,document);
		
		if (m_sPaymentMethod == "")
		{
			alert ("Return of Funds payment method combo validation value not found. Contact system administrator");
			frmToPP050.submit();
		}
		<% /* PSC 05/11/2002 BMIDS00760 - End */ %>
		
		frmScreen.txtCompletionDate.value		= XML.GetTagAttribute('DISBURSEMENTPAYMENT', 'COMPLETIONDATE');
		m_sPayeeHistorySeqNo					= XML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYEEHISTORYSEQNO');
		frmScreen.txtPayeeName.value			= GetPayeeName(m_sPayeeHistorySeqNo);
		m_sPayeeType							= XML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYEETYPE');
		<% /* PSC 28/11/2002 BMIDS01099 */ %>
		m_sPaymentType							= XML.GetTagAttribute('DISBURSEMENTPAYMENT', 'PAYMENTTYPE');
	}
	
	if(m_sMetaAction == "Edit") m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP080");
}

function CreateReturnOfFunds()
{
	var xmlRequestTag, xmlROFPayment, xmlROFDisbPayment, xmlOrigPayment, xmlPayment, xmlDisbPayment ;
	var xmlCaseStage, xmlApplication ;
	var iCount ;
	XMLSave = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	xmlRequestTag = XMLSave.CreateRequestTag(window, "CREATERETURNOFFUNDS");
	
	xmlROFPayment = XMLSave.CreateActiveTag("ROFPAYMENTRECORD");
	XMLSave.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	
	<%/* MV - 30/08/2002 - BMIDS00373 */%>		
	if(parseFloat(frmScreen.txtAmountReturned.value) < frmScreen.txtNetAdvanceAmount.value)
		XMLSave.SetAttribute("AMOUNT", frmScreen.txtAmountReturned.value);
	else
	{
		XMLSave.SetAttribute("AMOUNT", parseFloat(frmScreen.txtAmountReturned.value) + parseFloat(frmScreen.txtFeesDeducted.value));
		
		<%/* LDM 28/06/2006 EP886 
		Only update the post completion tasks if post completion (pc validation type)
		has been set up in the task type combo */%>		
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_sTaskType = XML.GetComboIdForValidation("TaskType", "PC", null, document);
		if (m_sTaskType)
		{
			<%/* WP13 MAR49 */%>		
			m_bUpdatePostCompletionTasks = true;
		}
		XML = null;
	}
	XMLSave.SetAttribute("PAYMENTMETHOD", m_sPaymentMethod);
	XMLSave.SetAttribute("ASSOCPAYSEQNUMBER", m_sPaymentSeqNo);
	<% /* PSC 05/11/2002 BMIDS00760 */ %>
	if(!GetReturnedStatus()) 
	{
		alert('Error finding the id for Returned payment status');
		return false ;
	}
	
	xmlROFDisbPayment = XMLSave.CreateActiveTag("ROFDISBURSEMENTPAYMENT");
	XMLSave.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	<% /* PSC 05/11/2002 BMIDS00760 */ %>
	XMLSave.SetAttribute("PAYMENTSTATUS", m_sReturnedId);
	XMLSave.SetAttribute("PAYMENTNOTES", frmScreen.txtNotes.value);
	<%/* MV - 30/08/2002 - BMIDS00373 */%>
	XMLSave.SetAttribute("NETPAYMENTAMOUNT", frmScreen.txtAmountReturned.value)
	
	XMLSave.ActiveTag = xmlROFPayment ;	 				
	xmlOrigPayment = XMLSave.CreateActiveTag("ORIGPAYMENTRECORD");
	XMLSave.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XMLSave.SetAttribute("PAYMENTSEQUENCENUMBER", m_sPaymentSeqNo);
	
	XMLSave.ActiveTag = xmlOrigPayment ;
	xmlPayment = XMLSave.CreateActiveTag("PAYMENTRECORD");
	XMLSave.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XMLSave.SetAttribute("PAYMENTSEQUENCENUMBER", m_sPaymentSeqNo);
	<%/* MV - 30/08/2002 - BMIDS00373 */%>
	XMLSave.SetAttribute("AMOUNT", frmScreen.txtGrossAdvanceAmount.value); 
	XMLSave.SetAttribute("PAYMENTMETHOD", m_sPaymentMethod);
	
	XMLSave.ActiveTag = xmlPayment ;
	xmlDisbPayment =  XMLSave.CreateActiveTag("DISBURSEMENTPAYMENT");
	XMLSave.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XMLSave.SetAttribute("PAYEEHISTORYSEQNO", m_sPayeeHistorySeqNo);
	XMLSave.SetAttribute("PAYMENTSEQUENCENUMBER", m_sPaymentSeqNo);
	XMLSave.SetAttribute("PAYEETYPE", m_sPayeeType);
	<% /* PSC 28/11/2002 BMIDS01099 */ %>
	XMLSave.SetAttribute("PAYMENTTYPE", m_sPaymentType);
	
	<% /* Append all FeePayments and LoanComponentPayment to the request */ %>
	XMLSave.ActiveTag = xmlPayment ;
	
	for(iCount = 0 ; iCount < xmlFeePaymentList.length ; iCount++)
	{
		XMLSave.ActiveTag.appendChild(xmlFeePaymentList(iCount));
	}
	
	for(iCount = 0 ; iCount < xmlLCPList.length ; iCount++)
	{
		XMLSave.ActiveTag.appendChild(xmlLCPList(iCount));
	}
	
	<% /* Append Case Stage details */ %>
	XMLSave.ActiveTag = xmlRequestTag ;
	xmlCaseStage = XMLSave.CreateActiveTag("CASESTAGE");
	XMLSave.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XMLSave.SetAttribute("CASEID", m_sApplicationNumber);
	XMLSave.SetAttribute("ACTIVITYID", m_sActivityId);
	XMLSave.SetAttribute("ACTIVITYINSTANCE", 1);
	XMLSave.SetAttribute("STAGEID", m_sStageId);
	
	XMLSave.ActiveTag = xmlRequestTag ;
	xmlApplication = XMLSave.CreateActiveTag("APPLICATION");
	XMLSave.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XMLSave.SetAttribute("APPLICATIONFACTINDNUMBER", m_sApplicationFactFindNumber);
	
	//XMLSave.RunASP(document,"PaymentProcessingRequest.asp");
	// BMIDS00977 Replace above call with screen rules processing
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XMLSave.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			XMLSave.SetErrorResponse();
		}

	if(!XMLSave.IsResponseOK())
	{
		alert('Error saving disbursement data')
		return false ;
	}
	
	return true ;
}

function ValidateBeforeSave()
{
	if(frmScreen.txtAmountReturned.value == "")
	{
		alert('Amount of funds returned cannot be empty.');
		frmScreen.txtAmountReturned.focus();
		return false ;
	}

	if(parseFloat(frmScreen.txtAmountReturned.value) == 0)
	{
		alert('Amount of funds returned should be grater than zero.');
		return false ;
	}
	
	// Amount of Retuned funds cannot be greater than Amount of Advance
	if(parseFloat(frmScreen.txtAmountReturned.value) > parseFloat(frmScreen.txtNetAdvanceAmount.value))
	{
		alert('Amount of funds returned cannot be greater than the Amount of advance.');	
		return false ;
	}
	
	// Notes can be a maximum of 255 chars
	if(scScreenFunctions.RestrictLength(frmScreen, "txtNotes", 255, true))
	{
		frmScreen.txtNotes.focus();
		return false;
	}
		
	return true ;
}

function GetPayeeName(sPayeeHistorySeqNo)
{	
	var sPayeeName = ""
	var xmlNode = null ;
	
	if(sPayeeHistorySeqNo == '') return sPayeeName ;

	PayeeHistoryXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	PayeeHistoryXML.CreateRequestTag(window, "FindPayeeHistoryList");
	PayeeHistoryXML.CreateActiveTag("PAYEEHISTORY");
	PayeeHistoryXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	PayeeHistoryXML.SetAttribute("_COMBOLOOKUP_","1");
	//BMIDS00977 Screen Rules processing added
	//PayeeHistoryXML.RunASP(document,"PaymentProcessingRequest.asp");
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			PayeeHistoryXML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			PayeeHistoryXML.SetErrorResponse();
		}

	if(!PayeeHistoryXML.IsResponseOK())
	{
		alert('Error retreiving Payee detais');
		return sPayeeName ;
	}

	xmlNode = PayeeHistoryXML.XMLDocument.selectSingleNode('//PAYEEHISTORY[@PAYEEHISTORYSEQNO=' + sPayeeHistorySeqNo + ']')
	if(xmlNode != null )
	{
		PayeeHistoryXML.ActiveTag = xmlNode ;
		sPayeeName = PayeeHistoryXML.GetTagAttribute('THIRDPARTY', 'COMPANYNAME'); ;
	}
	else alert('Error retriving Payee Name');

	return sPayeeName ; 
}

<% /* PSC 05/11/2002 BMIDS00760 - Start */ %>
function GetReturnedStatus()
{
	var xmlNode = null ;	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	var blnSuccess = true;		
	var sGroupList = new Array("PaymentStatus");
	
	if(XML.GetComboLists(document,sGroupList))
	{
		var sCondition = "//LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='ROF']/VALUEID" ;
		xmlNode = XML.XMLDocument.selectSingleNode(sCondition);
		
		if(xmlNode != null)  m_sReturnedId = xmlNode.text ;
		else m_sReturnedId = "" ;
		
		blnSuccess = true ;	
	}
	else
	{
		alert('Error retrieving combo valies.');
		blnSuccess =  false;
	}
	
	return blnSuccess ;
}
<% /* PSC 05/11/2002 BMIDS00760 - End */ %>

<% /* WP13 MAR49 */ %>
function UpdatePostCompletionTasks()
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
	XML.SetAttribute("OPERATION","ResetPostCompletionTasks");
	XML.SetAttribute("UPDATEPROPERTY","TaskStatus");
	XML.CreateActiveTag("FindCaseTaskList");
	XML.CreateActiveTag("CASETASK");
	XML.SetAttribute("SOURCEAPPLICATION", tagSourceAppl);
	XML.SetAttribute("CASEID", m_sApplicationNumber);
	XML.SetAttribute("ACTIVITYID", sActivityID);
	XML.SetAttribute("ACTIVITYINSTANCE", "1");
	XML.SetAttribute("TASKSTATUS","I");			//incomplete
	<%/* LDM 28/06/2006 EP886 code taken out to be a global. we dont want the hit of 2 checks to the db
	var sTaskType = XML.GetComboIdForValidation("TaskType", "PC", null, document);
	*/%>
	XML.SetAttribute("TASKTYPE",m_sTaskType);
	XML.SetAttribute("STAGEID", sStageID );
	XML.SetAttribute("TASKSTATUSID", "30" );	//Not Applicable
	//XML.SetAttribute("COMPLETIONDATE", frmScreen.txtCompletionDate.value );
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
	XML = null;
	gParamXML = null;

}
-->
</script>
</BODY>
</HTML>