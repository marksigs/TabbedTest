<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      cm035.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Add/Edit screen for the Valuables  THIS IS A POP-UP SCREEN
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		26/11/99	Initial creation
JLD		19/01/2000	Altered to make it a pop-up
APS		15/03/00	SYS0490, SYS0488, SYS0471
APS		04/05/00	SYS0652
BG		17/05/00	SYS0752 Removed Tooltips
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head><meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css"><title>Valuables Over Limit</title>
</head>
<body>
<!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<script src="validation.js" language="JScript"></script>
<span id="spnListScroll">
<span style="LEFT: 94px; POSITION: absolute; TOP: 200px">
<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; height:24; width:304" 
	type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span> 
</span>
<!-- Specify Screen Layout Here-->
<form id="frmScreen" mark validate="onchange" style="VISIBILITY: hidden">
<div id="divBackground" style="TOP: 10px; LEFT: 0px; HEIGHT: 330px; WIDTH: 604px; POSITION: ABSOLUTE">
<div id="divValuablesList" style="TOP: 0px; LEFT: 10px; HEIGHT: 230px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE; WIDTH: 600px" class="msgLabel">
	Valuables Over 
	<label class="msgCurrency"></label>
	<span id="spnValuablesOverLimit" class="msgLabel">
	</span>
	<span id="spnTable" style="LEFT: 0px; POSITION: absolute; TOP: 22px">
	<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
		<tr id="rowTitles"><td width="75%" class="TableHead">Description</td><td width="25%" class="TableHead">Item Value</td>
		<tr id="row01"><td width="75%" class="TableTopLeft">&nbsp;</td><td width="25%" class="TableTopRight">&nbsp;</td></tr>
		<tr id="row02"><td width="75%" class="TableLeft">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
		<tr id="row03"><td width="75%" class="TableLeft">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
		<tr id="row04"><td width="75%" class="TableLeft">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
		<tr id="row05"><td width="75%" class="TableLeft">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
		<tr id="row06"><td width="75%" class="TableLeft">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
		<tr id="row07"><td width="75%" class="TableLeft">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
		<tr id="row08"><td width="75%" class="TableLeft">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
		<tr id="row09"><td width="75%" class="TableLeft">&nbsp;</td><td width="25%" class="TableRight">&nbsp;</td></tr>
		<tr id="row10"><td width="75%" class="TableBottomLeft">&nbsp;</td><td width="25%" class="TableBottomRight">&nbsp;</td></tr>
	</table>
	</span>
</span>
<span style="TOP: 200px; LEFT: 400px; POSITION: ABSOLUTE" class="msgLabel">
	Total
	<span style="TOP: 0px; LEFT: 50px; POSITION: ABSOLUTE">
		<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtTotal" name="Total" maxlength="10" style="TOP: -3px; WIDTH: 75px; POSITION: ABSOLUTE" class="msgTxt">
	</span>
</span>
</div>
<div id="divAdminControls" style="TOP: 235px; LEFT: 10px; HEIGHT: 65px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
	Description
	<span style="TOP: 0px; LEFT: 80px; POSITION: ABSOLUTE">
		<input id="txtDescription" name="Description" maxlength="50" style="WIDTH: 400px; POSITION: absolute" class="msgTxt">
	</span>	
</span>
<span style="TOP: 40px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
	Value
	<span style="TOP: 0px; LEFT: 80px; POSITION: ABSOLUTE">
		<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtValue" name="Value" maxlength="10" style="TOP: -3px; WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
	</span>
	<span style="TOP: -3px; LEFT: 406px; POSITION: ABSOLUTE">
		<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	<span style="TOP: -3px; LEFT: 471px; POSITION: ABSOLUTE">
		<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	<span style="TOP: -3px; LEFT: 536px; POSITION: ABSOLUTE">
		<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
</span>
<span style="TOP: 70px; LEFT: 4px; POSITION: ABSOLUTE">
	<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
	<span style="TOP: 0px; LEFT: 64px; POSITION: ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
</span>	
</div>
</div>
</form>
<!-- File containing field attributes (remove if not required) -->
<!--  #include FILE="attribs/cm035attribs.asp" -->
<!-- Specify Code Here -->
<script language="JScript">
<!--
var m_iValuablesItemLimit = 0;
var m_iMaxTableLength = 10;
var m_sXMLIn = "";
var m_sReadOnly = "";
var m_ValXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sBCSubQuoteNumber = "";
var scScreenFunctions;
var m_sCurrency = "";
var m_BaseNonPopupWindow = null;

function window.onload()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];	
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions.ShowCollection(frmScreen);
	m_sXMLIn = sArgArray[0]; //CustomerXML
	SetMasks();
	Validation_Init();
	SetValuablesOverLimit(sArgArray[3]);
	m_sCurrency = scScreenFunctions.SetThisCurrency(frmScreen, sArgArray[2]);
	Initialise();		
	m_sReadOnly = sArgArray[1];
	if (m_sReadOnly == "1")
	{
		frmScreen.btnOK.disabled = true;
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
	}
	window.returnValue = null;   // return null if OK is not pressed
}
function Initialise()
{
	m_ValXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_ValXML.LoadXML(m_sXMLIn);
	if (m_ValXML.SelectTag(null, "BUILDINGSANDCONTENTSSUBQUOTE") != null)
	{
		m_sApplicationNumber = m_ValXML.GetTagText("APPLICATIONNUMBER");
		m_sApplicationFactFindNumber = m_ValXML.GetTagText("APPLICATIONFACTFINDNUMBER");
		m_sBCSubQuoteNumber = m_ValXML.GetTagText("BCSUBQUOTENUMBER");
	}
	PopulateScreen();
	if(m_sReadOnly == "1")
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	else
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTotal");
}
function PopulateScreen()
{
	PopulateValuablesList(0, "");
	tblTable.rows("rowTitles").cells(1).innerHTML = "Item Value " + m_sCurrency;
	if(m_ValXML.ActiveTagList.length > 0)
	{
<%		// Select the first valuable and display it in the admin controls
%>	
		scScrollTable.setRowSelected(1);
		UpdateRowControls();
	}
	else
	{
		frmScreen.txtDescription.focus();
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
	}
}
function spnTable.onclick() 
{
	if (m_sReadOnly != "1")
	{	
		UpdateRowControls();
		frmScreen.txtDescription.focus();
	}
}
function frmScreen.btnDelete.onclick()
{
<% /* If both fields are empty, the table row is deleted. This is achieved by 2 methods. If the item
to be deleted is existing on the database (ie it has a valuableslimitsequencenumber) then just 
delete the description and insurablevalue fields from the xml. This way, the item will be deleted
from the database when the xml is saved in CM020. 
If the item to be deleted does not have a sequence number then it does not exist
on the database so we can remove it completely from the XML */
%>
var iSelectedRow = scScrollTable.getOffset() + scScrollTable.getRowSelected();	
if (iSelectedRow >= 0)
{
	m_ValXML.ActiveTag = null;
	m_ValXML.CreateTagList("VALUABLESOVERLIMIT");
	var iIndex = tblTable.rows(scScrollTable.getRowSelected()).getAttribute("TagListItemCount");
	m_ValXML.SelectTagListItem(iIndex);
	m_ValXML.RemoveActiveTag();
	PopulateValuablesList(scScrollTable.getOffset(), "D");
	scScrollTable.setRowSelected(-1);	
	frmScreen.txtDescription.value = "";
	frmScreen.txtValue.value = "";
	frmScreen.btnDelete.disabled = true;
	frmScreen.btnEdit.disabled = true;
}
}
function ValidateControls()
{
var blnValid = false;
if(frmScreen.txtDescription.value != "" && frmScreen.txtValue.value != "")
{
	if(parseInt(frmScreen.txtValue.value) <= m_iValuablesItemLimit)
	{
		alert("Item value must be greater than " + m_sCurrency + m_iValuablesItemLimit);
		frmScreen.txtValue.focus();
	}
	else blnValid = true;
}
else alert("Description and Insurable Value expected");	
return blnValid;
}
function frmScreen.btnAdd.onclick()
{
if (ValidateControls())
{
	var xmlValuablesListNode = m_ValXML.SelectTag(null, "VALUABLESOVERLIMITLIST");
	m_ValXML.CreateActiveTag("VALUABLESOVERLIMIT");
	m_ValXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	m_ValXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	m_ValXML.CreateTag("BCSUBQUOTENUMBER", m_sBCSubQuoteNumber);
	m_ValXML.CreateTag("VALUABLESLIMITSEQUENCENUMBER", "");
	m_ValXML.CreateTag("DESCRIPTION", frmScreen.txtDescription.value);
	m_ValXML.CreateTag("INSURABLEVALUE", frmScreen.txtValue.value);
	var iOffSet = scScrollTable.getOffset();
	if (iOffSet == -1) iOffSet = 0;
	PopulateValuablesList(iOffSet, "A");
	frmScreen.txtDescription.value = "";
	frmScreen.txtValue.value = "";
	frmScreen.txtDescription.focus();
}
}
function frmScreen.btnEdit.onclick()
{
<%  /*This function updates the table and the xml string with the new values in Description and Value.
If there are contents in either field, the table row and xml values are replaced. 
APS 07/03/00 - This must be >= than 0 */
%>
var iSelectedRow = scScrollTable.getOffset() + scScrollTable.getRowSelected();	
if (iSelectedRow >= 0)
{	
	if (ValidateControls())	
	{
		m_ValXML.ActiveTag = null;
		m_ValXML.CreateTagList("VALUABLESOVERLIMIT");
		var iIndex = tblTable.rows(scScrollTable.getRowSelected()).getAttribute("TagListItemCount");
		m_ValXML.SelectTagListItem(iIndex);	
		m_ValXML.SetTagText("DESCRIPTION", frmScreen.txtDescription.value);
		m_ValXML.SetTagText("INSURABLEVALUE", frmScreen.txtValue.value);
		PopulateValuablesList(scScrollTable.getOffset(), "E");
		scScrollTable.setRowSelected(iSelectedRow - scScrollTable.getOffset());
	}
}
}
function SetValuablesOverLimit(sValuablesItemLimit)
{
	if (sValuablesItemLimit == "")sValuablesOverLimit = " <Value Not Found>";
	spnValuablesOverLimit.innerHTML = sValuablesItemLimit;
	if (isNaN(sValuablesItemLimit) != true)
		m_iValuablesItemLimit = parseInt(sValuablesItemLimit);
	else
		m_iValuablesItemLimit = 0;
	XML = null;
}
function frmScreen.btnOK.onclick()
{
	var sReturn = new Array();
	sReturn[0] = IsChanged();	// Has there been a change made
	sReturn[1] = m_ValXML.XMLDocument.xml;  // The XML string
	window.returnValue = sReturn;
	window.close();
}
function frmScreen.btnCancel.onclick()
{
	window.close();
}
function PopulateValuablesList(nStart, sOp)
{
	m_ValXML.ActiveTag = null;			
	m_ValXML.CreateTagList("VALUABLESOVERLIMIT");
	
	if (sOp != "E")  // Edit
		scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iMaxTableLength, m_ValXML.ActiveTagList.length);
	
	if (sOp == "A") // Add
	{
		if (m_ValXML.ActiveTagList.length <= m_iMaxTableLength)
			ShowList(nStart);
		else
			scScrollTable.toEnd();
	}
	else if (sOp == "D")  // Delete
	{
		if (nStart==0)
			ShowList(nStart);
		else
			scScrollTable.setRecordSelected(nStart--);
	}
	else
		ShowList(nStart);
	
	if (m_ValXML.SelectTag(null, "BUILDINGSANDCONTENTSDETAILS") != null)
		m_ValXML.SetTagText("TOTALVALUABLESEXCEEDINGLIMIT", frmScreen.txtTotal.value);
}
function ShowList(nStart)
{
	scScrollTable.clear();
	m_ValXML.ActiveTag = null;
	m_ValXML.CreateTagList("VALUABLESOVERLIMIT");
	var iRowIdx = 0;
	for (var iCount = 0; iCount < m_ValXML.ActiveTagList.length && iRowIdx < m_iMaxTableLength; iCount++)
	{
<%		// Add to the Valuables list table if the values aren't null
%>		if (m_ValXML.SelectTagListItem(iCount + nStart) == true)
		{
			if(m_ValXML.GetTagText("DESCRIPTION") != "" && m_ValXML.GetTagText("INSURABLEVALUE") != "")
			{
				scScreenFunctions.SizeTextToField(tblTable.rows(iRowIdx+1).cells(0),m_ValXML.GetTagText("DESCRIPTION"));
				scScreenFunctions.SizeTextToField(tblTable.rows(iRowIdx+1).cells(1),m_ValXML.GetTagText("INSURABLEVALUE"));
				tblTable.rows(iRowIdx+1).setAttribute("TagListItemCount", iCount + nStart);
				iRowIdx++;
			}
		}
	}
	frmScreen.txtTotal.value = "0";
	for(iCount = 0; iCount < m_ValXML.ActiveTagList.length; iCount++)
	{
		m_ValXML.SelectTagListItem(iCount);
		if(m_ValXML.GetTagText("INSURABLEVALUE") != "")
			frmScreen.txtTotal.value = parseInt(frmScreen.txtTotal.value) + parseInt(m_ValXML.GetTagText("INSURABLEVALUE"));
	}
}
function UpdateRowControls()
{
	var iSelectedRow = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	// APS 07/03/00 - This must be >= than 0
	if (iSelectedRow >= 0)
	{
		m_ValXML.ActiveTag = null;			
		m_ValXML.CreateTagList("VALUABLESOVERLIMIT");
		if(iSelectedRow <= m_ValXML.ActiveTagList.length)
		{
			frmScreen.btnEdit.disabled = false;
			frmScreen.btnDelete.disabled = false;
			var iIndex = tblTable.rows(scScrollTable.getRowSelected()).getAttribute("TagListItemCount");
			m_ValXML.SelectTagListItem(iIndex);
			frmScreen.txtDescription.value = m_ValXML.GetTagText("DESCRIPTION");
			frmScreen.txtValue.value = m_ValXML.GetTagText("INSURABLEVALUE");
		}
	}
	else 
	{	
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
	}						
}
-->
</script>
</body>
</html>
