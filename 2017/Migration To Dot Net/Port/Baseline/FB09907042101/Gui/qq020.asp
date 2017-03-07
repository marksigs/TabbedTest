<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      qq020.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: Quick Quote Outgoings
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		03/04/00	scScreenFunctions change
AY		13/04/00	SYS0328 - Dynamic currency display
					spnToFirstField/spnToLastField removed
MH		10/05/00	SYS0408 Validation of fields
BG		17/05/00	SYS0752 Removed Tooltips
APS		23/05/00	SYS0780 Made the frmScreen invisible - when scScreenFunctions.DisplayPopup is
					called it will make the frmScreen visible
SA		06/08/01	SYS0408 If freq input first - make sure it saves amount as well.
SA		23/08/01	SYS2569	If amount input is zero, save as empty string cboFrequexncy.onBlur altered.
					SYS0408 Rolled back above changes as they didn't work quite how they should!!
LD		23/05/02	SYS4727 Use cached versions of frame functions
MC		20/04/2004	White space char padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Quick Quote Outgoings  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>
<script src="validation.js" language="JScript"></script>
<span id=spnListScroll>
	<span style="LEFT: 145px; POSITION: absolute; TOP: 200px">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; height:24; width:304" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>	</span> 
</span>	
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange" year4>
<div style="HEIGHT: 260px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 443px" class="msgGroup">
	<span id="txtApplicantName" name = "ApplicantName" style ="LEFT: 4px; POSITION: absolute; TOP: 10px" tabindex = "-1" class="msglabel"></span>
	<span id="spnOutgoingsTable" style="LEFT: 4px; POSITION: absolute; TOP: 30px">
		<table id="tblOutgoingsTable" width="435" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="45%" class="TableHead">Outgoing Type</td><td id="tdAmount" width="35%" class="TableHead">Amount</td>	<td width="20%" class="TableHead">Frequency</td></tr>
			<tr id="row01">     <td width="45%" class="TableTopLeft">&nbsp</td><td width="35%" class="TableTopCenter">&nbsp</td><td width="20%" class="TableTopRight">&nbsp</td></tr>
			<tr id="row02">     <td width="45%" class="TableLeft">&nbsp</td>   <td width="35%" class="TableCenter">&nbsp</td><td width="20%" class="TableRight">&nbsp</td></tr>
			<tr id="row03">     <td width="45%" class="TableLeft">&nbsp</td>   <td width="35%" class="TableCenter">&nbsp</td><td width="20%" class="TableRight">&nbsp</td></tr>
			<tr id="row04">     <td width="45%" class="TableLeft">&nbsp</td>   <td width="35%" class="TableCenter">&nbsp</td><td width="20%" class="TableRight">&nbsp</td></tr>
			<tr id="row05">     <td width="45%" class="TableLeft">&nbsp</td>   <td width="35%" class="TableCenter">&nbsp</td><td width="20%" class="TableRight">&nbsp</td></tr>
			<tr id="row06">     <td width="45%" class="TableLeft">&nbsp</td>   <td width="35%" class="TableCenter">&nbsp</td><td width="20%" class="TableRight">&nbsp</td></tr>
			<tr id="row07">     <td width="45%" class="TableLeft">&nbsp</td>   <td width="35%" class="TableCenter">&nbsp</td><td width="20%" class="TableRight">&nbsp</td></tr>
			<tr id="row08">     <td width="45%" class="TableLeft">&nbsp</td>   <td width="35%" class="TableCenter">&nbsp</td><td width="20%" class="TableRight">&nbsp</td></tr>
			<tr id="row09">     <td width="45%" class="TableLeft">&nbsp</td>   <td width="35%" class="TableCenter">&nbsp</td><td width="20%" class="TableRight">&nbsp</td></tr>
			<tr id="row10">     <td width="45%" class="TableBottomLeft">&nbsp</td><td width="35%" class="TableBottomCenter">&nbsp</td><td width="20%" class="TableBottomRight">&nbsp</td></tr>
		</table>
	</span>
	<span id="spnRowInput" style="LEFT: 4px; HEIGHT: 30px; POSITION: absolute; TOP: 220px; WIDTH: 435px" class="msgGroupLight">
		<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Amount
			<span style="LEFT: 100px; POSITION: absolute; TOP: 0px" class="msgLabel">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtAmount" maxlength="10" name="Amount"  
					style="POSITION: absolute; WIDTH: 75px; TOP: -3px" class="msgTxt"> 
			</span>
		</span>
		<span id="spnFrequency" style="LEFT: 225px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Frequency
			<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
				<select id="cboFrequency" name="Frequency"  
					style="WIDTH: 100px" class="msgCombo">
				</select>
			</span>
		</span>
	</span>
</div>
<div style="LEFT: 10px; POSITION: absolute; TOP: 275px; WIDTH: 443px; HEIGHT: 70px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Total Regular Outgoings
		<span style="TOP: 0px; LEFT: 150px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtTotalOutgoings" name="TotalOutgoings"  
					maxlength="10" style="WIDTH: 100px; TOP: -3px; POSITION: ABSOLUTE" class="msgReadOnly" tabindex="-1">
		</span>
	</span>
	<span style="TOP: 40px; LEFT: 4px; POSITION: ABSOLUTE">
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
<!--  #include FILE="attribs/qq020attribs.asp" -->
<script language="JScript">
<!--
var m_sCustomerName				= null;
var m_sCustomerNumber			= null;
var m_sCustomerVersionNumber	= null;
var m_sApplicationNumber		= null;
var m_sApplicationFactFindNumber= null;
var m_sReadOnly					= null;
var m_sXMLOutgoings				= null;
var m_XMLOutgoings				= null;
var m_XMLOutgoingTypes			= null;
var m_iOutgoingTypes			= 0;
var m_sTotalAmount				= null;
var m_iTableLength				= 10;
var m_tagQUICKQUOTEOUTGOINGSLIST= null;
var scScreenFunctions;
var m_BaseNonPopupWindow = null;

function window.onload()
{
	RetrieveData();
	SetMasks();	
	Validation_Init();
	SetApplicantName();
	PopulateOutgoingsTable();
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
	m_sXMLOutgoings					= sParameters[0];
	m_sCustomerName					= sParameters[1];
	m_sCustomerNumber				= sParameters[2];
	m_sCustomerVersionNumber		= sParameters[3];
	m_sReadOnly						= sParameters[4];
	m_sApplicationNumber			= sParameters[5];
	m_sApplicationFactFindNumber	= sParameters[6];

	var sCurrencyText = scScreenFunctions.SetThisCurrency(frmScreen,sParameters[7]);
	tdAmount.innerText = tdAmount.innerText + " " + sCurrencyText;

	m_XMLOutgoings = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_XMLOutgoings.LoadXML(m_sXMLOutgoings);
	m_XMLOutgoings.CreateTagList("QUICKQUOTEOUTGOINGSLIST");
	if (m_XMLOutgoings.SelectTagListItem(0) == true)m_tagQUICKQUOTEOUTGOINGSLIST = m_XMLOutgoings.ActiveTag;
	else
	{
		m_XMLOutgoings.CreateActiveTag("QUICKQUOTEOUTGOINGSLIST");
		m_tagQUICKQUOTEOUTGOINGSLIST = m_XMLOutgoings.ActiveTag;
	}
}
function PopulateOutgoingsTable()
{
	var sGroups = new Array("RegularOutgoingsType", "RegularOutgoingsPaymentFreq");
	var blnSuccess = false;
	m_XMLOutgoingTypes = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	if (m_XMLOutgoingTypes.GetComboLists(document, sGroups) == true)
	{
		m_XMLOutgoingTypes.PopulateCombo(document, frmScreen.cboFrequency, "RegularOutgoingsPaymentFreq", true);
		FilterOutgoingTypes();
		m_XMLOutgoingTypes.ActiveTag = null;
		m_XMLOutgoingTypes.CreateTagList("LISTENTRY");
		m_iOutgoingTypes = m_XMLOutgoingTypes.ActiveTagList.length;
		scScrollTable.initialiseTable(tblOutgoingsTable, 0, "", ReadOutgoings, m_iTableLength, m_iOutgoingTypes);
		ReadOutgoings(0);
		scScrollTable.setRowSelected(1)
		UpdateRowControls();
		blnSuccess = true;
	}
	if (blnSuccess == false)
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnOK.disabled = true;
	}
	return blnSuccess;
}
function FilterOutgoingTypes()
{
	var ValidationList = new Array("Q", "B");
	var TagListLISTENTRY = m_XMLOutgoingTypes.CreateTagList("LISTENTRY");
<%	// Add all combo entries brought back from the database
%>	for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
	{
		m_XMLOutgoingTypes.ActiveTagList = TagListLISTENTRY;
		m_XMLOutgoingTypes.SelectTagListItem(nLoop);
		if (m_XMLOutgoingTypes.IsInComboValidationXML(ValidationList) != true)m_XMLOutgoingTypes.RemoveActiveTag();
	}
}
function InitialiseOutgoingTable(iStart)
{
	m_XMLOutgoingTypes.ActiveTag = null;
	var tagLISTENTRYLIST = m_XMLOutgoingTypes.CreateTagList("LISTENTRY");
	for (var iLoop = 0; iLoop < m_XMLOutgoingTypes.ActiveTagList.length && iLoop < m_iTableLength; iLoop++)
	{
		m_XMLOutgoingTypes.SelectTagListItem(iLoop + iStart);
		scScreenFunctions.SizeTextToField(tblOutgoingsTable.rows(iLoop+1).cells(0),	m_XMLOutgoingTypes.GetTagText("VALUENAME"));
		scScreenFunctions.SizeTextToField(tblOutgoingsTable.rows(iLoop+1).cells(1), "");
		scScreenFunctions.SizeTextToField(tblOutgoingsTable.rows(iLoop+1).cells(2), "");
		tblOutgoingsTable.rows(iLoop+1).setAttribute("OutgoingType", m_XMLOutgoingTypes.GetTagText("VALUEID"));
	}
}
function ReadOutgoings(iStart)
{
	m_XMLOutgoings.ActiveTag = null;
	m_XMLOutgoings.CreateTagList("QUICKQUOTEOUTGOINGS");
	InitialiseOutgoingTable(iStart);
	for (var iLoop = 0; iLoop < m_XMLOutgoings.ActiveTagList.length; iLoop++)
	{
		m_XMLOutgoings.SelectTagListItem(iLoop);
		AddOutgoingToTable(m_XMLOutgoings.GetTagText("QUICKQUOTEOUTGOINGSTYPE"), m_XMLOutgoings.GetTagText("AMOUNT"), 
		                     m_XMLOutgoings.GetTagAttribute("FREQUENCYSEQUENCENUMBER", "TEXT"), m_XMLOutgoings.GetTagText("FREQUENCYSEQUENCENUMBER"));
	}
	UpdateRowControls();
}
function AddOutgoingToTable(sOutgoingType, sAmount, sFrequencyText, sFrequencySequenceNumber)
{
	var sRowOutgoingType;
	var iIndex		= 1;
	var blnMatch	= false;
	do 
	{
		sRowOutgoingType = tblOutgoingsTable.rows(iIndex).getAttribute("OutgoingType");
		if (sRowOutgoingType == sOutgoingType)
		{
			scScreenFunctions.SizeTextToField(tblOutgoingsTable.rows(iIndex).cells(1), sAmount);
			scScreenFunctions.SizeTextToField(tblOutgoingsTable.rows(iIndex).cells(2), sFrequencyText);
			tblOutgoingsTable.rows(iIndex).setAttribute("FrequencySequenceNumber", sFrequencySequenceNumber);
			blnMatch = true;
		}
		iIndex++;
	}
	while ((iIndex <= m_iTableLength) && (blnMatch == false))
}
function spnOutgoingsTable.onclick() 
{
	if (ValidateScreen())
	{
		UpdateRowControls();
		frmScreen.txtAmount.focus();
	}
}
function frmScreen.cboFrequency.onchange()
{
	frmScreen.btnOK.disabled = true;
	frmScreen.txtTotalOutgoings.value = ""; 
}
function frmScreen.txtAmount.onchange ()
{
	frmScreen.btnOK.disabled = true;
	frmScreen.txtTotalOutgoings.value = ""; 
}
//SYS0408 Need this in case frequency is input before amount.
//function frmScreen.txtAmount.onblur()
//{
//	if (frmScreen.cboFrequency.selectedIndex > 0)
//	{
//		frmScreen.cboFrequency.onblur();
//	}
//}

function frmScreen.cboFrequency.onblur()
{
	var iIndex = 0;
	var sText = "";
	var sAmountInput="";	//SYS2569
	var iSelectedRow = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	if (iSelectedRow != -1 && ValidateScreen())
	{
		iIndex = frmScreen.cboFrequency.selectedIndex;
		
		if (iIndex != 0)sText = frmScreen.cboFrequency.item(iIndex).text;
		if (frmScreen.txtAmount.value != 0)sAmountInput = frmScreen.txtAmount.value;	//SYS2569
		//scScreenFunctions.SizeTextToField(tblOutgoingsTable.rows(scScrollTable.getRowSelected()).cells(1), frmScreen.txtAmount.value);
		scScreenFunctions.SizeTextToField(tblOutgoingsTable.rows(scScrollTable.getRowSelected()).cells(1), sAmountInput);	
		scScreenFunctions.SizeTextToField(tblOutgoingsTable.rows(scScrollTable.getRowSelected()).cells(2), sText);
		tblOutgoingsTable.rows(scScrollTable.getRowSelected()).setAttribute("FrequencySequenceNumber",GetFrequencySequenceNumber(iIndex));
		SaveQuickQuoteOutgoings();
		iSelectedRow++;
		if (iSelectedRow <= m_iOutgoingTypes)
		{
			if (iSelectedRow > m_iTableLength)
			{
				scScrollTable.oneDown();
				iSelectedRow = m_iTableLength;
			}
			scScrollTable.setRowSelected(iSelectedRow);
			UpdateRowControls();
		}
		frmScreen.txtAmount.focus();
	}
}
function GetComboOption(iSequenceNumber)
{
	var optOption = null;
	var blnFound = false;
	var iNumberOfOptions = frmScreen.cboFrequency.options.length;
	for (var iLoop = 0; iLoop < iNumberOfOptions && blnFound==false; iLoop++)
	{
		if (frmScreen.cboFrequency.options(iLoop).value == iSequenceNumber)
		{
			optOption = frmScreen.cboFrequency.options(iLoop);
			blnFound = true;
		}
	}
	return optOption;
}
function GetFrequencyCalcValue(iSequenceNumber)
{
	var iFrequency = 1;
	var optOption = GetComboOption(iSequenceNumber);
	if (optOption != null)
	{
		iFrequency = optOption.value;
		if (isNaN(iFrequency) == true)iFrequency = 1;
	}
	return iFrequency;
}
function GetFrequencySequenceNumber(iIndex)
{
	return (frmScreen.cboFrequency.options(iIndex).value);
}
function GetFrequencyText(iSequenceNumber)
{
	var sText = "";
	var optOption = GetComboOption(iSequenceNumber);
	if (optOption != null)sText = optOption.text
	return sText;
}
function SetComboText(iSequenceNumber)
{
	var iIndex = 0;
	var optOption = GetComboOption(iSequenceNumber);
	if (optOption != null)iIndex = optOption.index;
	frmScreen.cboFrequency.selectedIndex = iIndex;
}
function UpdateRowControls()
{
	var iSelectedRow = scScrollTable.getRowSelected();
	if (iSelectedRow != -1)
	{
		frmScreen.txtAmount.value = tblOutgoingsTable.rows(iSelectedRow).cells(1).innerText;
		var iFrequencySequenceNumber = tblOutgoingsTable.rows(iSelectedRow).getAttribute("FrequencySequenceNumber");	
		SetComboText(iFrequencySequenceNumber);
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
	else frmScreen.txtAmount.focus();
}
function SetApplicantName()
{
	txtApplicantName.innerHTML = "<strong>" + m_sCustomerName + "</strong>";
}
function frmScreen.btnCancel.onclick()
{
	m_XMLOutgoings		= null;
	m_XMLOutgoingTypes	= null;
	window.returnValue	= null;
	window.close();
}
function frmScreen.btnOK.onclick()
{
	if (ValidateScreen())
	{
		var sReturn = new Array();
		m_sXMLOutgoings = m_XMLOutgoings.XMLDocument.xml;
		sReturn[0] = IsChanged();
		sReturn[1] = m_sXMLOutgoings;
		sReturn[2] = m_sTotalAmount;
		m_XMLOutgoings		= null;
		m_XMLOutgoingTypes	= null;
		window.returnValue	= sReturn;
		window.close();
	}
}
function SaveQuickQuoteOutgoings()
{
	var sOutgoingType, sAmount, sFrequencyValue, sFrequencyText;
	var blnMatch = false;
	m_XMLOutgoings.ActiveTag = null;
	m_XMLOutgoings.CreateTagList("QUICKQUOTEOUTGOINGS");
	sOutgoingType	= tblOutgoingsTable.rows(scScrollTable.getRowSelected()).getAttribute("OutgoingType");
	sAmount			= tblOutgoingsTable.rows(scScrollTable.getRowSelected()).cells(1).innerText;
	sFrequencyText 	= tblOutgoingsTable.rows(scScrollTable.getRowSelected()).cells(2).innerText;
	sFrequencyValue = tblOutgoingsTable.rows(scScrollTable.getRowSelected()).getAttribute("FrequencySequenceNumber");
	for (var iLoop = 0; iLoop < m_XMLOutgoings.ActiveTagList.length && blnMatch == false; iLoop++)
	{
		m_XMLOutgoings.SelectTagListItem(iLoop);
		sXMLOutgoingType = m_XMLOutgoings.GetTagText("QUICKQUOTEOUTGOINGSTYPE");
		if (sOutgoingType == sXMLOutgoingType)
		{
			blnMatch = true;
			if ((sAmount == "") && (sFrequencyValue == ""))	m_XMLOutgoings.RemoveActiveTag();
			else
			{
				m_XMLOutgoings.SetTagText("AMOUNT",	sAmount);
				m_XMLOutgoings.SetTagText("FREQUENCYSEQUENCENUMBER", sFrequencyValue);
				m_XMLOutgoings.SetAttributeOnTag("FREQUENCYSEQUENCENUMBER", "TEXT", sFrequencyText);
			}
		}
	}
	if (blnMatch != true)AddOutgoingToXMLList(iLoop, sOutgoingType, sAmount, sFrequencyValue, sFrequencyText);
}
function AddOutgoingToXMLList(iSelectedRow, sOutgoingType, sAmount, sFrequencyValue, sFrequencyText)
{
	m_XMLOutgoings.ActiveTag = m_tagQUICKQUOTEOUTGOINGSLIST;
	m_XMLOutgoings.CreateActiveTag("QUICKQUOTEOUTGOINGS");
	m_XMLOutgoings.CreateTag("APPLICATIONNUMBER",			m_sApplicationNumber);
	m_XMLOutgoings.CreateTag("APPLICATIONFACTFINDNUMBER",	m_sApplicationFactFindNumber);
	m_XMLOutgoings.CreateTag("CUSTOMERNUMBER",				m_sCustomerNumber);
	m_XMLOutgoings.CreateTag("CUSTOMERVERSIONNUMBER",		m_sCustomerVersionNumber);
	m_XMLOutgoings.CreateTag("QUICKQUOTEOUTGOINGSTYPE",		sOutgoingType);
	m_XMLOutgoings.CreateTag("AMOUNT",						sAmount);
	m_XMLOutgoings.CreateTag("FREQUENCYSEQUENCENUMBER",		sFrequencyValue);
	m_XMLOutgoings.SetAttributeOnTag("FREQUENCYSEQUENCENUMBER", "TEXT", sFrequencyText);
}
function frmScreen.btnCalc.onclick()
{
	if (ValidateScreen())
	{
		Calculate();
		frmScreen.btnOK.disabled = false;
	}
}
function Calculate()
{
	var iTotalMonthlyAmount	 = 0; 
	var iAmount				 = 0;
	var iFrequency			 = 0;
	var iFrequencyCalculator = 0;
	m_XMLOutgoings.ActiveTag = null;
	m_XMLOutgoings.CreateTagList("QUICKQUOTEOUTGOINGS");
	for (var iLoop = 0; iLoop < m_XMLOutgoings.ActiveTagList.length; iLoop++)
	{
		m_XMLOutgoings.SelectTagListItem(iLoop);
		iAmount		= parseInt(m_XMLOutgoings.GetTagText("AMOUNT"));
		iFrequency	= parseInt(m_XMLOutgoings.GetTagText("FREQUENCYSEQUENCENUMBER"));
		iFrequencyCalculator = GetFrequencyCalcValue(iFrequency);
		if (isNaN(iAmount) != true)iTotalMonthlyAmount += ((iAmount * iFrequencyCalculator) / 12);
	}
	iTotalMonthlyAmount = m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundValue(iTotalMonthlyAmount,0);
	m_sTotalAmount = iTotalMonthlyAmount;
	frmScreen.txtTotalOutgoings.value = iTotalMonthlyAmount;
}

function ValidateScreen()
{
	var iSelectedRow = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	var iIndex;
	if (iSelectedRow != -1)
	{
		iIndex = frmScreen.cboFrequency.selectedIndex;
		
		if (iIndex>0 && frmScreen.txtAmount.value <= 0 ) 
		{
			GotoRequiredField(frmScreen.txtAmount);
			return false;
		}
		
		if (iIndex==0 && frmScreen.txtAmount.value > 0 ) 
		{
			GotoRequiredField(frmScreen.cboFrequency);
			return false;
		}

	}
	return true;

}

function GotoRequiredField(refField)
{
	alert ("Please enter a value");
	refField.focus ();
}
-->
</script>
</body>
</html>

