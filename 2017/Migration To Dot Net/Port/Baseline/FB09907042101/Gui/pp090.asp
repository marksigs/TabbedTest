<%@ LANGUAGE="JSCRIPT" %>
<% /*
Workfile:      pp090.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Cancellation Reason
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		16/02/01	New
CL		05/03/01	SYS1920 Read only functionality added
LD		23/05/02	SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<HEAD>
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>

<script src="validation.js" language="JScript"></script>
</HEAD>

<BODY>
<% /* Scriptlets */ %>

<%/* FORMS  */ %>
<FORM id="frmToPP050" method="post" action="PP050.asp" STYLE="DISPLAY: none"></FORM>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 170px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Total loan amount 
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtTotalLoanAmount" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<span style="TOP: 44px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Total advance to date
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtTotalAdvancedToDate" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<span style="TOP: 74px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Balance
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtBalance" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<span style="TOP: 104px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Reason for cancellation
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<textarea id="txtNotes" name="Notes" rows="4" style="WIDTH: 400px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>
	
</div>
</form>

<% /* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 270px; WIDTH: 600px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>
<!-- #include FILE="fw030.asp" -->

<%/* CODE */ %>
<script language="JScript">
<!--
var m_sXML ;
var m_sTotalLoanAmount, m_sBalance, m_sTotalTobeDisbursed ;
var m_sApplicationNumber, m_sUserId ;
var m_sPaymentType ;
var m_SubmitEnabled = false ;
var m_blnReadOnly = false;


function RetrieveContextData()
{

	// TEST
	/*
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber", "C00078387");
	scScreenFunctions.SetContextParameter(window,"idUserId","Srini");
	*/
	// END TEST
	
	m_sXML = scScreenFunctions.GetContextParameter(window,"idXml","");
	if(m_sXML == '')
	{
		alert('Required data not found in the context.');
		return false ;
	}
	
	var xmlContext = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	xmlContext.LoadXML(m_sXML);
	
	xmlContext.CreateTagList('CALCULATIONS')
	if(xmlContext.ActiveTagList.length > 0)
	{
		xmlContext.SelectTagListItem(0);
		m_sTotalLoanAmount		= xmlContext.GetAttribute('AMOUNTREQUESTED');
		m_sTotalTobeDisbursed	= xmlContext.GetAttribute('TOTALTOBEDISBURSED');
		m_sBalance				= xmlContext.GetAttribute('BALANCE');
	}
	else
	{
		alert('Required data not found in the context.');
		return false ;	
	}
	
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sUserId			 = scScreenFunctions.GetContextParameter(window,"idUserId","");
	
	return true;
}

/* Events  */
function window.onload()
{	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Cancellation Reason","PP090",scScreenFunctions)
	
	var sButtonList = new Array("Submit", "Cancel");
	ShowMainButtons(sButtonList);
	
	Validation_Init();
	if(!RetrieveContextData()) return ;
	InitialiseScreen();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP090");
}

function frmScreen.txtNotes.onkeypress()
{
	if(!m_SubmitEnabled)
	{
		m_SubmitEnabled = true ;
		EnableMainButton('Submit');
	}
}

function btnSubmit.onclick()
{
	if(ValidateBeforeSave())
	{
		if(CancelBalance()) frmToPP050.submit();
		else return ;
	}
}

function btnCancel.onclick()
{
	frmToPP050.submit();
}

/* FUNCTIONS */
function InitialiseScreen()
{
	frmScreen.txtTotalLoanAmount.value		= m_sTotalLoanAmount ;
	frmScreen.txtBalance.value				= m_sBalance ;
	frmScreen.txtTotalAdvancedToDate.value	= parseFloat(m_sTotalTobeDisbursed) - parseFloat(m_sBalance);
	//m_sPaymentType							=  
	DisableMainButton('Submit')
}

function ValidateBeforeSave()
{	
	// Notes can be a maximum of 255 chars
	if(scScreenFunctions.RestrictLength(frmScreen, "txtNotes", 255, true))
	{
		frmScreen.txtNotes.focus();
		return false;
	}
	
	return true;
}

function CancelBalance()
{	
	 // get code for PaymentType - 'Balance Cancellation Payment'
	if(!GetBalanceCanellationCode()) return ;
	
	// Build request and call method 'CancelBalance'
	var xmlRequestTag ;
	XMLSave = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	xmlRequestTag = XMLSave.CreateRequestTag(window, "CANCELBALANCE");
	
	XMLSave.CreateActiveTag("APPLICATIONFEETYPE");
	XMLSave.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	
	XMLSave.ActiveTag = xmlRequestTag ;
	XMLSave.CreateActiveTag("PAYMENTRECORD");
	XMLSave.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XMLSave.SetAttribute("AMOUNT", frmScreen.txtBalance.value);
	XMLSave.SetAttribute("USERID", m_sUserId);
	
	XMLSave.CreateActiveTag("DISBURSEMENTPAYMENT");
	XMLSave.SetAttribute("PAYMENTTYPE", m_sPaymentType);
	
	XMLSave.RunASP(document,"PaymentProcessingRequest.asp");
	
	if(!XMLSave.IsResponseOK())	return false ;
	
	return true ;
}

function GetBalanceCanellationCode()
{
	var xmlNode = null ;	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	var blnSuccess = true;		
	var sGroupList = new Array("PaymentType");
	
	if(XML.GetComboLists(document,sGroupList))
	{
		var sCondition = "//LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='NCB']/VALUEID" ;
		xmlNode = XML.XMLDocument.selectSingleNode(sCondition);
		
		if(xmlNode != null) m_sPaymentType = xmlNode.text ;
		else m_sPaymentType = "" ;
		
		blnSuccess = true ;	
	}
	else
	{
		alert('Error retrieving combo valies.');
		blnSuccess =  false;
	}
	
	return blnSuccess ;
}
-->
</script>
</BODY>
</HTML>