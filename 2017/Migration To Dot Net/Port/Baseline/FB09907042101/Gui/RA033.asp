<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      RA033.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Notices of Correction
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		05/12/00	Screen Design
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
HMA     08/12/2003  BMIDS675    Display all NOC lines.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<head>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Notices of Correction  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>

<script src="validation.js" language="JScript"></script>

<span id="spnListScroll">
	<span style="LEFT: 100px; POSITION: absolute; TOP: 445px">
		<OBJECT data=scTableListScroll.asp id=scScrollTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 300px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>


<% //Span to keep tabbing within this screen %> 
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" validate ="onchange" mark>
<div id="divCCData" style="HEIGHT: 56px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 400px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		CCN Ref. No.
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtCreditCheckReferenceNumber" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Name
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtFullBureauName" maxlength="50" style="POSITION: absolute; WIDTH: 250px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
</div>

<div id="divAddress" style="HEIGHT: 197px; LEFT: 10px; POSITION: absolute; TOP: 70px; WIDTH: 400px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Current Address</strong>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
		Flat No. 
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtFlatNo" maxlength="50" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 54px" class="msgLabel">
		House Name & No. 
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtHouseName" maxlength="50" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
		<span style="LEFT: 271px; POSITION: absolute; TOP: -3px">
			<input id="txtHouseNo" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 78px" class="msgLabel">
		Street 
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtStreet" maxlength="50" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 102px" class="msgLabel">
		District 
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtDistrict" maxlength="50" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 126px" class="msgLabel">
		Town
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtTown" maxlength="50" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 150px" class="msgLabel">
		County
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtCounty" maxlength="50" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>	
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 174px" class="msgLabel">
		Post Code
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtPostCode" maxlength="50" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>	
</div>

<div id="divCorrectionNotices" style="HEIGHT: 205px; LEFT: 10px; POSITION: absolute; TOP: 271px; WIDTH: 400px" class="msgGroup">
	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 10px">
		<table id="tblTable" width="392" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="100%" class="TableHead">Notices of Correction&nbsp</td> <td width="0%" class="TableHead"></td> </tr>
			<tr id="row01">		<td width="100%" class="TableTopLeft">&nbsp;</td> <td width="0%" class="TableTopRight">&nbsp;</td> </tr>
			<tr id="row02">		<td width="100%" class="TableLeft">&nbsp;</td> <td width="0%" class="TableRight">&nbsp;</td> </tr>
			<tr id="row03">		<td width="100%" class="TableLeft">&nbsp;</td> <td width="0%" class="TableRight">&nbsp;</td> </tr>
			<tr id="row04">		<td width="100%" class="TableLeft">&nbsp;</td> <td width="0%" class="TableRight">&nbsp;</td> </tr>
			<tr id="row05">		<td width="100%" class="TableLeft">&nbsp;</td> <td width="0%" class="TableRight">&nbsp;</td> </tr>
			<tr id="row06">		<td width="100%" class="TableLeft">&nbsp;</td> <td width="0%" class="TableRight">&nbsp;</td> </tr>
			<tr id="row07">		<td width="100%" class="TableLeft">&nbsp;</td> <td width="0%" class="TableRight">&nbsp;</td> </tr>
			<tr id="row08">		<td width="100%" class="TableLeft">&nbsp;</td> <td width="0%" class="TableRight">&nbsp;</td> </tr>
			<tr id="row09">		<td width="100%" class="TableLeft">&nbsp;</td> <td width="0%" class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="100%" class="TableBottomLeft">&nbsp;</td> <td width="0%" class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 500px; WIDTH: 400px">
<!-- #include FILE="msgButtons.asp" -->
</div>

<script language="JScript">
<!--
var m_aArgArray = null;
var m_aFBAddress = null;
var m_iTableLength = 10;
var scScreenFunctions ;

var XML ;
var m_BaseNonPopupWindow = null;

/* EVENTS */
function window.onload()
{

	var sArguments		= window.dialogArguments ;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];	
	m_aArgArray			= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
	
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	PopulateScreen();
}

function btnSubmit.onclick()
{
	window.close();	
}


/* FUNCTIONS */
function PopulateScreen()
{
	frmScreen.txtCreditCheckReferenceNumber.value = m_aArgArray[0];
	frmScreen.txtFullBureauName.value = m_aArgArray[1];
	
	m_aFBAddress = m_aArgArray[2];
	
	frmScreen.txtFlatNo.value		= m_aFBAddress[0] ;
	frmScreen.txtHouseName.value	= m_aFBAddress[1] ;
	frmScreen.txtHouseNo.value		= m_aFBAddress[2] ;
	frmScreen.txtStreet.value		= m_aFBAddress[3] ;
	frmScreen.txtTown.value			= m_aFBAddress[4] ;
	frmScreen.txtDistrict.value		= m_aFBAddress[5] ;
	frmScreen.txtCounty.value		= m_aFBAddress[6] ;
	frmScreen.txtPostCode.value		= m_aFBAddress[7] ;
	
	// Populate the list box 
	var sFullBureauDataXML	= m_aArgArray[3];
	
	XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(sFullBureauDataXML);
	
	XML.CreateTagList("FULLBUREAUCORRECTIONLINES");	
	var iNumberOfBlocks = XML.ActiveTagList.length;
	
	if(iNumberOfBlocks > 0)
	{
		scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfBlocks);
		/* BMIDS675 Take into account the starting point when displaying data */
		ShowList(0);    
	}
}

/* BMIDS675 Take into account the starting point when displaying data */
function ShowList(nStart)
{
	var iCount;
	var sFBCorrectionDetails ;
	
	for (iCount=0; iCount < XML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		/* BMIDS675 Take into account the starting point when displaying data */
		XML.SelectTagListItem(iCount + nStart);
		
		sFBCorrectionDetails = XML.GetTagText("FBNOTICEOFCORRECTIONDETAILS");
		
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sFBCorrectionDetails);
	}
}
-->
</script>
</body>
</html>