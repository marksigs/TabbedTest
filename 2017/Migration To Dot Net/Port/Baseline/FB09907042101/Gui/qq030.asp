<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      qq020.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: Quick Quote Personal Debts
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		03/04/00	scScreenFunctions change
AY		13/04/00	SYS0328 - Dynamic currency display
					spnToFirstField/spnToLastField removed
BG		17/05/00	SYS0752 Removed Tooltips
APS		23/05/00	SYS0780 Made the frmScreen invisible - when scScreenFunctions.DisplayPopup is
					called it will make the frmScreen visible
LD		23/05/02	SYS4727 Use cached versions of frame functions
MC		20/04/2004	White space char padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Quick Quote Personal Debts  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>
<script src="validation.js" language="JScript"></script>
<span id=spnListScroll>
	<span style="LEFT: 145px; POSITION: absolute; TOP: 200px">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; height:24; width:304" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>	</span> 
</span>
<!-- Specify Screen Layout Here -->
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange" year4>
<div style="HEIGHT: 260px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 443px" class="msgGroup">
	<span id="txtApplicantName" name = "ApplicantName" 
		style ="LEFT: 4px; POSITION: absolute; TOP: 10px" tabindex = "-1" class="msglabel">
	</span>
	<span id="spnPersonalDebtsTable" style="LEFT: 4px; POSITION: absolute; TOP: 30px">
		<table id="tblPersonalDebtsTable" width="435" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="45%" class="TableHead">Personal Debt Type</td>	<td id="tdMPay" width="35%" class="TableHead">Monthly Payment</td>	<td id="tdOSBal" width="20%" class="TableHead">O/S Balance</td></tr>
			<tr id="row01">		<td width="45%" class="TableTopLeft">&nbsp</td>				<td width="35%" class="TableTopCenter">&nbsp</td>					<td width="20%" class="TableTopRight">&nbsp</td>	</tr>
			<tr id="row02">		<td width="45%" class="TableLeft">&nbsp</td>				<td width="35%" class="TableCenter">&nbsp</td>						<td width="20%" class="TableRight">&nbsp</td>		</tr>
			<tr id="row03">		<td width="45%" class="TableLeft">&nbsp</td>				<td width="35%" class="TableCenter">&nbsp</td>						<td width="20%" class="TableRight">&nbsp</td>		</tr>
			<tr id="row04">		<td width="45%" class="TableLeft">&nbsp</td>				<td width="35%" class="TableCenter">&nbsp</td>						<td width="20%" class="TableRight">&nbsp</td>		</tr>
			<tr id="row05">		<td width="45%" class="TableLeft">&nbsp</td>				<td width="35%" class="TableCenter">&nbsp</td>						<td width="20%" class="TableRight">&nbsp</td>		</tr>
			<tr id="row06">		<td width="45%" class="TableLeft">&nbsp</td>				<td width="35%" class="TableCenter">&nbsp</td>						<td width="20%" class="TableRight">&nbsp</td>		</tr>
			<tr id="row07">		<td width="45%" class="TableLeft">&nbsp</td>				<td width="35%" class="TableCenter">&nbsp</td>						<td width="20%" class="TableRight">&nbsp</td>		</tr>
			<tr id="row08">		<td width="45%" class="TableLeft">&nbsp</td>				<td width="35%" class="TableCenter">&nbsp</td>						<td width="20%" class="TableRight">&nbsp</td>		</tr>
			<tr id="row09">		<td width="45%" class="TableLeft">&nbsp</td>				<td width="35%" class="TableCenter">&nbsp</td>						<td width="20%" class="TableRight">&nbsp</td>		</tr>
			<tr id="row10">		<td width="45%" class="TableBottomLeft">&nbsp</td>			<td width="35%" class="TableBottomCenter">&nbsp</td>				<td width="20%" class="TableBottomRight">&nbsp</td>	</tr>
		</table>
	</span>
	<span style="LEFT: 4px; HEIGHT: 30px; POSITION: absolute; TOP: 220px; WIDTH: 435px" class="msgGroupLight">
		<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Monthly Payment
			<span style="LEFT: 100px; POSITION: absolute; TOP: 0px" class="msgLabel">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtMonthlyPayment" maxlength="10" name="MonthlyPayment" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 225px; POSITION: absolute; TOP: 10px" class="msgLabel">
			O/S Balance
			<span style="LEFT: 100px; POSITION: absolute; TOP: 0px" class="msgLabel">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtOutstandingBalance" maxlength="10" name="OutstandingBalance" style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt">
			</span>
		</span>
	</span>
</div>
<div style="LEFT: 10px; POSITION: absolute; TOP: 275px; WIDTH: 443px; HEIGHT: 100px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Total Monthly Payments
		<span style="TOP: 0px; LEFT: 150px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtTotalMonthlyPayments" name="TotalMonthlyPayments"  
						maxlength="10" style="WIDTH: 100px; TOP: -3px; POSITION: ABSOLUTE" class="msgReadOnly" tabindex="-1">
		</span>
	</span>
	<span style="TOP: 40px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Total O/S Balance
		<span style="TOP: 0px; LEFT: 150px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtTotalOSBalance" name="TotalOSBalance"  
						maxlength="10" style="WIDTH: 100px; TOP: -3px; POSITION: ABSOLUTE" class="msgReadOnly" tabindex="-1">
		</span>
	</span>
	<span style="TOP: 70px; LEFT: 4px; POSITION: ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" disabled="true" class="msgButton">
		<span style="TOP: 0px; LEFT: 64px; POSITION: ABSOLUTE">
			<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="TOP: 0px; LEFT: 128px; POSITION: ABSOLUTE">
			<input id="btnCalc" value="Calc" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>
</div>
</form>
<!-- File containing field attributes (remove if not required) -->
<!--  #include FILE="attribs/qq030attribs.asp" -->
<script language="JScript">
<!--
var m_sCustomerName				= null;
var m_sCustomerNumber			= null;
var m_sCustomerVersionNumber	= null;
var m_sApplicationNumber		= null;
var m_sApplicationFactFindNumber= null;
var m_sReadOnly					= null;
var m_sXMLPersonalDebts			= null;
var m_XMLPersonalDebts			= null;
var m_XMLDebtType				= null;
var m_iTotalPersonalDebts		= null;
var m_sTotalOSBalance			= null;
var m_sTotalMonthlyPayment		= null;
var m_iTableLength				= 10;
var m_tagQUICKQUOTEPERSONALDEBTSLIST = null;
var scScreenfunctions;
var m_BaseNonPopupWindow = null;

function window.onload()
{
	RetrieveData();
	SetMasks();
	Validation_Init();
	SetApplicantName();
	PopulateDebtTable();
	SetScreenOnReadOnly();
	window.returnValue = null;
}
function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters		= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sXMLPersonalDebts				= sParameters[0];
	m_sCustomerName					= sParameters[1];
	m_sCustomerNumber				= sParameters[2];
	m_sCustomerVersionNumber		= sParameters[3];
	m_sReadOnly						= sParameters[4];
	m_sApplicationNumber			= sParameters[5];
	m_sApplicationFactFindNumber	= sParameters[6];

	var sCurrencyText = scScreenFunctions.SetThisCurrency(frmScreen,sParameters[7]);
	tdMPay.innerText = tdMPay.innerText + " " + sCurrencyText;
	tdOSBal.innerText = tdOSBal.innerText + " " + sCurrencyText;

	m_XMLPersonalDebts = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_XMLPersonalDebts.LoadXML(m_sXMLPersonalDebts);
	m_XMLPersonalDebts.CreateTagList("QUICKQUOTEPERSONALDEBTSLIST");
	if (m_XMLPersonalDebts.SelectTagListItem(0) == true)m_tagQUICKQUOTEPERSONALDEBTSLIST = m_XMLPersonalDebts.ActiveTag;
	else
	{
		m_XMLPersonalDebts.CreateActiveTag("QUICKQUOTEPERSONALDEBTSLIST");
		m_tagQUICKQUOTEPERSONALDEBTSLIST = m_XMLPersonalDebts.ActiveTag;
	}
}
function PopulateDebtTable()
{
	var sGroups = new Array("PersonalDebtsType");
	var blnSuccess = false;
	m_XMLDebtType = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	if (m_XMLDebtType.GetComboLists(document, sGroups) == true)
	{
		blnSuccess = true;
		m_XMLDebtType.CreateTagList("LISTENTRY");
		m_iTotalPersonalDebts = m_XMLDebtType.ActiveTagList.length;
		scScrollTable.initialiseTable(tblPersonalDebtsTable, 0, "", ReadPersonalDebts, m_iTableLength, m_iTotalPersonalDebts);
		ReadPersonalDebts(0);
		scScrollTable.setRowSelected(1);
		UpdateRowControls();
	}
	if (blnSuccess == false)
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnOK.disabled = true;
	}
	return blnSuccess;
}
function InitialiseDebtTable(iStart)
{
	m_XMLDebtType.ActiveTag = null;
	var tagLISTENTRYLIST = m_XMLDebtType.CreateTagList("LISTENTRY");
	var sDebtTypeDescription, sDebtTypeValue;
	for (var iLoop = 0; iLoop < m_XMLDebtType.ActiveTagList.length && iLoop < m_iTableLength; iLoop++)
	{
		m_XMLDebtType.SelectTagListItem(iLoop + iStart);
		sDebtTypeDescription	= m_XMLDebtType.GetTagText("VALUENAME");
		sDebtTypeValue			= m_XMLDebtType.GetTagText("VALUEID");
		scScreenFunctions.SizeTextToField(tblPersonalDebtsTable.rows(iLoop+1).cells(0),sDebtTypeDescription);
		scScreenFunctions.SizeTextToField(tblPersonalDebtsTable.rows(iLoop+1).cells(1), "");
		scScreenFunctions.SizeTextToField(tblPersonalDebtsTable.rows(iLoop+1).cells(2), "");
		tblPersonalDebtsTable.rows(iLoop+1).setAttribute("DebtType", sDebtTypeValue);
	}
}
function ReadPersonalDebts(iStart)
{
	var sDebtType, sMonthlyPayment, sOSBalance;
	m_XMLPersonalDebts.ActiveTag = null;
	var tagQUICKQUOTEPERSONALDEBTS = m_XMLPersonalDebts.CreateTagList("QUICKQUOTEPERSONALDEBTS");
	InitialiseDebtTable(iStart);
	for (var iLoop = 0; iLoop < m_XMLPersonalDebts.ActiveTagList.length; iLoop++)
	{
		m_XMLPersonalDebts.SelectTagListItem(iLoop);
		sDebtType		= m_XMLPersonalDebts.GetTagText("QUICKQUOTEPERSONALDEBTSTYPE");
		sOSBalance		= m_XMLPersonalDebts.GetTagText("OUTSTANDINGBALANCE");
		sMonthlyPayment = m_XMLPersonalDebts.GetTagText("MONTHLYPAYMENT");
		AddDebtToTable(sDebtType, sOSBalance, sMonthlyPayment);
	}
	UpdateRowControls();
}
function AddDebtToTable(sDebtType, sOSBalance, sMonthlyPayment)
{
	var sRowDebtType;
	var iIndex = 1;
	var blnMatch = false;
	do 
	{
		sRowDebtType = tblPersonalDebtsTable.rows(iIndex).getAttribute("DebtType");
		if (sRowDebtType == sDebtType)
		{
			scScreenFunctions.SizeTextToField(tblPersonalDebtsTable.rows(iIndex).cells(1), sMonthlyPayment);
			scScreenFunctions.SizeTextToField(tblPersonalDebtsTable.rows(iIndex).cells(2), sOSBalance);
			blnMatch = true;
		}
		iIndex++;
	}
	while ((iIndex <= m_iTableLength) && (blnMatch == false))
}
function spnPersonalDebtsTable.onclick() 
{
	UpdateRowControls();
	frmScreen.txtMonthlyPayment.focus();  //Set focus to monthly payment   JLD aqr11
}
function frmScreen.txtOutstandingBalance.onchange()
{
	frmScreen.btnOK.disabled = true;
	frmScreen.txtTotalMonthlyPayments.value = "";
	frmScreen.txtTotalOSBalance.value = "";
}
function frmScreen.txtMonthlyPayment.onchange()
{
	frmScreen.btnOK.disabled = true;
	frmScreen.txtTotalMonthlyPayments.value = "";
	frmScreen.txtTotalOSBalance.value = "";
}
function frmScreen.txtOutstandingBalance.onblur()
{
	var iSelectedRow = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	if (iSelectedRow != -1)
	{
		scScreenFunctions.SizeTextToField(tblPersonalDebtsTable.rows(scScrollTable.getRowSelected()).cells(1), frmScreen.txtMonthlyPayment.value);
		scScreenFunctions.SizeTextToField(tblPersonalDebtsTable.rows(scScrollTable.getRowSelected()).cells(2), frmScreen.txtOutstandingBalance.value);
		SaveQuickQuotePersonalDebts();
		iSelectedRow++;
		if (iSelectedRow > m_iTableLength)
		{
			scScrollTable.oneDown();
			iSelectedRow = m_iTableLength;
		}
		scScrollTable.setRowSelected(iSelectedRow);
		UpdateRowControls();
		frmScreen.txtMonthlyPayment.focus();
	}
}
function UpdateRowControls()
{
	var iSelectedRow = scScrollTable.getRowSelected();
	if (iSelectedRow != -1)
	{
		frmScreen.txtMonthlyPayment.value		= tblPersonalDebtsTable.rows(iSelectedRow).cells(1).innerText;
		frmScreen.txtOutstandingBalance.value	= tblPersonalDebtsTable.rows(iSelectedRow).cells(2).innerText;
	}
}
function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnCalc.disabled = true;
		frmScreen.btnOK.disabled = true;
		frmScreen.btnCancel.focus();
	}
	else frmScreen.txtMonthlyPayment.focus();
}
function SetApplicantName()
{
	txtApplicantName.innerHTML = "<strong>" + m_sCustomerName + "</strong>";
}
function frmScreen.btnCancel.onclick()
{
	m_XMLPersonalDebts	= null;
	m_XMLDebtType		= null;
	window.returnValue	= null;
	window.close();
}
function frmScreen.btnOK.onclick()
{
	var sReturn = new Array();
	m_sXMLPersonalDebts = m_XMLPersonalDebts.XMLDocument.xml;
	sReturn[0] = IsChanged();
	sReturn[1] = m_sXMLPersonalDebts;
	sReturn[2] = m_sTotalMonthlyPayment;
	sReturn[3] = m_sTotalOSBalance;
	m_XMLPersonalDebts	= null;
	m_XMLDebtType		= null;
	window.returnValue	= sReturn;
	window.close();
}
function SaveQuickQuotePersonalDebts()
{
	var sMonthlyPayment, sOSBalance, sDebtType;
	var blnMatch = false;
	m_XMLPersonalDebts.ActiveTag = null;
	m_XMLPersonalDebts.CreateTagList("QUICKQUOTEPERSONALDEBTS");
	sDebtType		= tblPersonalDebtsTable.rows(scScrollTable.getRowSelected()).getAttribute("DebtType");
	sMonthlyPayment = tblPersonalDebtsTable.rows(scScrollTable.getRowSelected()).cells(1).innerText;
	sOSBalance		= tblPersonalDebtsTable.rows(scScrollTable.getRowSelected()).cells(2).innerText;
	for (var iLoop = 0; iLoop < m_XMLPersonalDebts.ActiveTagList.length && blnMatch == false; iLoop++)
	{
		m_XMLPersonalDebts.SelectTagListItem(iLoop);
		sXMLDebtType = m_XMLPersonalDebts.GetTagText("QUICKQUOTEPERSONALDEBTSTYPE");
		if (sDebtType == sXMLDebtType)
		{
			blnMatch = true;
			if ((sMonthlyPayment == "") && (sOSBalance == ""))m_XMLPersonalDebts.RemoveActiveTag();
			else
			{
				m_XMLPersonalDebts.SetTagText("MONTHLYPAYMENT",		sMonthlyPayment);
				m_XMLPersonalDebts.SetTagText("OUTSTANDINGBALANCE", sOSBalance);
			}
		}
	}
	if (blnMatch != true)AddPersonalDebtToXMLList(iLoop, sDebtType, sMonthlyPayment, sOSBalance);
}
function AddPersonalDebtToXMLList(iSelectedRow, sDebtType, sMonthlyPayment, sOSBalance)
{
	m_XMLPersonalDebts.ActiveTag = m_tagQUICKQUOTEPERSONALDEBTSLIST;
	m_XMLPersonalDebts.CreateActiveTag("QUICKQUOTEPERSONALDEBTS");
	m_XMLPersonalDebts.CreateTag("APPLICATIONNUMBER",			m_sApplicationNumber);
	m_XMLPersonalDebts.CreateTag("APPLICATIONFACTFINDNUMBER",	m_sApplicationFactFindNumber);
	m_XMLPersonalDebts.CreateTag("CUSTOMERNUMBER",				m_sCustomerNumber);
	m_XMLPersonalDebts.CreateTag("CUSTOMERVERSIONNUMBER",		m_sCustomerVersionNumber);
	m_XMLPersonalDebts.CreateTag("QUICKQUOTEPERSONALDEBTSTYPE",	sDebtType);
	m_XMLPersonalDebts.CreateTag("MONTHLYPAYMENT",				sMonthlyPayment);
	m_XMLPersonalDebts.CreateTag("OUTSTANDINGBALANCE",			sOSBalance);
}
function frmScreen.btnCalc.onclick()
{
	Calculate();
	frmScreen.btnOK.disabled = false;
}
function Calculate()
{
	var iTotalMonthlyPayment	= 0;
	var	iTotalOSBalance			= 0;
	var iMonthlyPayment			= 0;
	var iOSBalance				= 0;
	m_XMLPersonalDebts.ActiveTag = null;
	m_XMLPersonalDebts.CreateTagList("QUICKQUOTEPERSONALDEBTS");
	for (var iLoop = 0; iLoop < m_XMLPersonalDebts.ActiveTagList.length; iLoop++)
	{
		m_XMLPersonalDebts.SelectTagListItem(iLoop);
		iMonthlyPayment = parseInt(m_XMLPersonalDebts.GetTagText("MONTHLYPAYMENT"));
		iOSBalance		= parseInt(m_XMLPersonalDebts.GetTagText("OUTSTANDINGBALANCE"));
		if (isNaN(iMonthlyPayment) != true)iTotalMonthlyPayment += iMonthlyPayment;
		if (isNaN(iOSBalance) != true)iTotalOSBalance += iOSBalance;
	}
	m_sTotalMonthlyPayment	= iTotalMonthlyPayment;
	m_sTotalOSBalance		= iTotalOSBalance;
	frmScreen.txtTotalMonthlyPayments.value = iTotalMonthlyPayment;
	frmScreen.txtTotalOSBalance.value		= iTotalOSBalance;
}
-->
</script>
</body>
</html>

