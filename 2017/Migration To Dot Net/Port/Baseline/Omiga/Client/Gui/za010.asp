<%@ LANGUAGE="JSCRIPT" %>
<% /*
Workfile:      za010.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   PAF Search Results screen.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		01/02/00	Reworked for IE5 and performance.
AY		03/04/00	scScreenFunctions change
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		29/05/02	SYS4767 MSMS to Core integration
SAB		08/11/02	Rolled back from MARS version to allow local postcode validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<html>

<head>
<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>PAF Search Results  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<script src="validation.js" language="JScript"></script>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>

<% // Specify Screen Layout Here %>
<form id="frmScreen" mark validate="onchange" year4 style="VISIBILITY: hidden">
	<div style="TOP: 10px; LEFT: 10px; HEIGHT: 194px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
		<span id="spnTable" style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE">
			<table id="tblTable" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
				<tr id="rowTitles">	<td width="80px" class="TableHead">Postcode</td><td class="TableHead">Address</td></tr>
				<tr id="row01">		<td width="80px" class="TableTopLeft">&nbsp</td>		<td class="TableTopRight">&nbsp</td></tr>
				<tr id="row02">		<td width="80px" class="TableLeft">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row03">		<td width="80px" class="TableLeft">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row04">		<td width="80px" class="TableLeft">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row05">		<td width="80px" class="TableLeft">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row06">		<td width="80px" class="TableLeft">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row07">		<td width="80px" class="TableLeft">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row08">		<td width="80px" class="TableLeft">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row09">		<td width="80px" class="TableLeft">&nbsp</td>		<td class="TableRight">&nbsp</td></tr>
				<tr id="row10">		<td width="80px" class="TableBottomLeft">&nbsp</td>	<td class="TableBottomRight">&nbsp</td></tr>
			</table>
		</span>

		<span style="TOP: 166px; LEFT: 4px; POSITION: ABSOLUTE">
			<input id="btnSelect" value="Select" type="button" style="WIDTH: 60px" class="msgButton">

			<span style="TOP: 0px; LEFT: 70px; POSITION: ABSOLUTE">
				<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
			</span>
		</span>
	</div>
</form>

<% // List Scroll object %>
<span style="LEFT: 310px; POSITION: absolute; TOP: 175px">
	<object data=scListScroll.htm id=scScrollPlus style="LEFT: 0px; TOP: 0px; height:24; width:304" type=text/x-scriptlet VIEWASTEXT></object>
</span> 

<% // Specify Code Here %>
<script language="JScript">
<!--	// JScript Code
var m_xmlIn;
var m_xmlAddressList;

//SG 29/05/02 SYS4767
//var m_xmlPostCodeList;

var scScreenFunctions;

//SG 29/05/02 SYS4767
//var m_bIsPostCodeList = false;

function window.onload()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	m_BaseNonPopupWindow = sArguments[5];
			
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_xmlIn = new ActiveXObject("microsoft.XMLDOM");
			
	m_xmlIn.loadXML(sArguments[4]);
			
	//SG 29/05/02 SYS4767 START
	//m_xmlPostCodeList = m_xmlIn.selectNodes("RESPONSE/POSTCODELIST/POSTCODE");
	//m_xmlAddressList = m_xmlIn.selectNodes("RESPONSE/PAFADDRESSLIST/PAFADDRESS");
	m_xmlAddressList = m_xmlIn.selectNodes("//RESPONSE/ADDRESSDATA/ITEM");
	//SG 29/05/02 SYS4767 END	
	
	scScreenFunctions.ShowCollection(frmScreen);

	Validation_Init();		
	
	scTable.initialise(tblTable, 0, "");
	
	//SG 29/05/02 SYS4767 START			
	//if(m_xmlPostCodeList.length > 0)
	//{
	//	m_bIsPostCodeList = true;
	//	scScrollPlus.Initialise(ShowList,10,m_xmlPostCodeList.length);
	//}
	//else
	//	scScrollPlus.Initialise(ShowList,10,m_xmlAddressList.length);
	//		
	//ShowList(0);
	scScrollPlus.Initialise(ShowQASList,10,m_xmlAddressList.length);

	ShowQASList(0);
	//SG 29/05/02 SYS4767 END			
}

function frmScreen.btnSelect.onclick()
{
	if(scTable.getRowSelected() < 0)
	{
		//SG 29/05/02 SYS4767 START
		//if(m_bIsPostCodeList)
		//	alert("No Post Code Selected");
		//else
		//	alert("No Address Selected");
		//return;

		alert("No Address Selected");
		return;	
		//SG 29/05/02 SYS4767 END								
	}

	//SG 29/05/02 SYS4767
	//if(m_bIsPostCodeList)
	//{
	//	m_bIsPostCodeList = false;
	//	buildSelectedList();
	//	return;
	//}

	window.returnValue = m_xmlAddressList.item(scScrollPlus.getOffset() + scTable.getRowSelected() - 1).xml;
			
	window.close();
}

function frmScreen.btnCancel.onclick()
{
	//SG 29/05/02 SYS4767
	//if(m_xmlPostCodeList.length > 0 && !m_bIsPostCodeList)
	//{
	//	m_bIsPostCodeList = true;
	//	scTable.clear();
	//	scScrollPlus.Initialise(ShowList,10,m_xmlPostCodeList.length);
	//	ShowList(0);
	//	return;
	//}

	window.close();
}
//SG 29/05/02 SYS4767 Function removed		
/*function buildSelectedList()
{
	var sPostCode = tblTable.rows(scTable.getRowSelected()).cells(0).innerText;
	var xslPattern = "RESPONSE/PAFADDRESSLIST/PAFADDRESS[ $not$ POSTCODE='" + sPostCode + "']";
			
	var xslSource = "<?xml version=\"1.0\"?>";
	xslSource += "<xsl:stylesheet xmlns:xsl=\"http://www.w3.org/TR/WD-xsl\">";
	xslSource += "<xsl:template><xsl:copy><xsl:apply-templates select=\"@*\"/><xsl:apply-templates /></xsl:copy></xsl:template>"
	xslSource += "<xsl:template match=\"";
	xslSource += xslPattern;
	xslSource += "\" />";
	xslSource += "<xsl:template match=\"RESPONSE/POSTCODELIST\" />";
	xslSource += "</xsl:stylesheet>";
			
	var xslIn = new ActiveXObject("microsoft.XMLDOM");
	xslIn.loadXML(xslSource);

	<% 
		//			if(!xslIn.loadXML(xslSource))
		//			{
		//				xmlError(xslIn.parseError);
		//				return;
		//			}
	 %>
			
	var xmlAddressList = m_xmlIn.transformNode(xslIn);
			
	var xmlIn = new ActiveXObject("microsoft.XMLDOM");
	xmlIn.loadXML(xmlAddressList);
			
	m_xmlAddressList = xmlIn.selectNodes("RESPONSE/PAFADDRESSLIST/PAFADDRESS");
			
	scTable.clear();
	scScrollPlus.Initialise(ShowList,10,m_xmlAddressList.length);
			
	ShowList(0);
}
*/
function ShowList(nStart)
{
	var i;
			
	if(m_bIsPostCodeList)
	{
		tblTable.rows(0).cells(1).innerText = "Count";

		for(i = 0; i < 10; i++)
		{
			var xmlNode = m_xmlPostCodeList.item(i + nStart);
			if(xmlNode == null)
				break;
			ShowPostCodeRow(i,xmlNode);
		}
	}
	else
	{
		tblTable.rows(0).cells(1).innerText = "Address";

		for(i = 0; i < 10; i++)
		{
			var xmlNode = m_xmlAddressList.item(i + nStart);
			if(xmlNode == null)
				break;
			ShowAddressRow(i,xmlNode);
		}
	}
}
//SG 29/05/02 SYS4767 Function removed
/*function ShowPostCodeRow(nIndex,xmlNode)
{
	<% // AY 09/09/1999 - Set the table fields %>
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex + 1).cells(0),xmlNode.text);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex + 1).cells(1),xmlNode.getAttribute("COUNT"));
}
*/
//SG 29/05/02 SYS4767 Function removed		
/*function ShowAddressRow(nIndex,xmlNode)
{
	<% // AY 09/09/1999 - Set the table fields %>
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex + 1).cells(0),xmlNode.selectSingleNode("//POSTCODE").text);

	var sAddress = null;
	var sLine = null;
			
	if(xmlNode.selectSingleNode("FLATNUMBER") != null)
		sAddress = xmlNode.selectSingleNode("FLATNUMBER").text;
			
	if(xmlNode.selectSingleNode("BUILDINGORHOUSENAME") != null)
	{
		sLine = xmlNode.selectSingleNode("BUILDINGORHOUSENAME").text;
		if(sLine.length > 0)
		{
			if(sAddress != null && sAddress.length > 0)
				sAddress += ", ";
			sAddress += sLine;
		}
	}
			
	if(xmlNode.selectSingleNode("BUILDINGORHOUSENUMBER") != null)
	{
		sLine = xmlNode.selectSingleNode("BUILDINGORHOUSENUMBER").text;
		if(sLine.length > 0)
		{
			if(sAddress != null && sAddress.length > 0)
				sAddress += ", ";
			sAddress += sLine;
		}
	}
			
	if(xmlNode.selectSingleNode("STREET") != null)
	{
		sLine = xmlNode.selectSingleNode("STREET").text;
		if(sLine.length > 0)
		{
			if(sAddress != null && sAddress.length > 0)
				sAddress += ", ";
			sAddress += sLine;
		}
	}
			
	if(xmlNode.selectSingleNode("DISTRICT") != null)
	{
		sLine = xmlNode.selectSingleNode("DISTRICT").text;
		if(sLine.length > 0)
		{
			if(sAddress != null && sAddress.length > 0)
				sAddress += ", ";
			sAddress += sLine;
		}
	}
			
	if(xmlNode.selectSingleNode("TOWN") != null)
	{
		sLine = xmlNode.selectSingleNode("TOWN").text;
		if(sLine.length > 0)
		{
			if(sAddress != null && sAddress.length > 0)
				sAddress += ", ";
			sAddress += sLine;
		}
	}
			
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex + 1).cells(1),sAddress);
}
*/		
function xmlError(objParseError)
{
	var sMsg = "XML parser error - \n\r" +
				"Reason: " + objParseError.reason + "\r\n" +
				"Error code: " + objParseError.errorCode + "\r\n" +
				"Line no.: " + objParseError.Line + "\r\n" +
				"Character: " + objParseError.linepos + "\r\n" +
				"Source text: " + objParseError.srcText;
						
	alert(sMsg);
}
//SG 29/05/02 SYS4767 Function added
function ShowQASList(nStart)
{
	var i;

	tblTable.rows(0).cells(1).innerText = "Address";

	for(i = 0; i < 10; i++)
	{
		var xmlNode = m_xmlAddressList.item(i + nStart);
		if(xmlNode == null)
			break;
		ShowQASRow(i,xmlNode);
	}
}
//SG 29/05/02 SYS4767 Function added
function ShowQASRow(nIndex,xmlNode)
{
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex + 1).cells(0),xmlNode.getAttribute("POSTCODE"));
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex + 1).cells(1),xmlNode.getAttribute("TEXT"));
}
-->
</script>
</body>
</html>
