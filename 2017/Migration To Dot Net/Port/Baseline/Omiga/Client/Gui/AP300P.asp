<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<%
/*
Workfile:      AP300.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Conditions Applied
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		27/02/01	New
SR		14/03/01	SYS2064
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History : 

Prog	Date		AQR			Description
MV		07/08/2002	BMIDS0302	Core Ref : SYS4728 remove non-style sheet styles
TW      09/10/2002	Modified to incorporate client validation - SYS5115
ASu		16/10/2002	BMIDS00647	Correct spelling Error.
MV		09/04/2003  BM0507		Amended spnTable.onclick()
MC		20/04/2004	BMIDS517	White space padded to the title text inorder to hide std IE popup message 'Web Page Dialog'
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<Title>AP300P - Conditions Applied  <!-- #include file="includes/TitleWhiteSpace.asp" --></Title>
</HEAD>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Scriptlets */ %>

<span id="spnListScroll">
	<span style="LEFT: 306px; POSITION: absolute; TOP: 210px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 380px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Conditions Applied</strong>
	</span>
	
	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 24px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="15%" class="TableHead">Condition Number</td>	
				<td width="70%" class="TableHead">Description</td>	
				<td width="15%" class="TableHead">Satisfied?</td><% /* ASu BMIDS00647 - Correct Spelling */ %>
			</tr>
			<tr id="row01">		
				<td width="15%" class="TableTopLeft">&nbsp;</td>		
				<td width="70%" class="TableTopCenter">&nbsp;</td>		
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">		
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">		
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row05">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row06">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row07">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row08">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row09">
				<td width="15%" class="TableLeft">&nbsp;</td>			
				<td width="70%" class="TableMiddleCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row10">
				<td width="15%" class="TableBottomLeft">&nbsp;</td>	
				<td width="70%" class="TableBottomCenter">&nbsp;</td>		
				<td class="TableBottomRight">&nbsp;</td>
			</tr> 
		</table>
	</span>
		
	<span style="TOP: 240px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Details
		<span style="TOP: 20px; LEFT: 0px; POSITION: ABSOLUTE">
			<TEXTAREA class=msgTxt id="txtDetails" name=Notes rows=5 style="POSITION: absolute; WIDTH: 595px" readonly></TEXTAREA>  
		</span>	
	</span>
	
	<span style="TOP: 350px; LEFT: 4px; POSITION: ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 75px" class="msgButton">
	</span>	
	
</div>
</form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/AP300PAttribs.asp" -->

<script language="JScript">
<!--

var scScreenFunctions ;
var m_iTableLength = 10 ;
var ApplConditionsXML ;

var m_sApplicationNumber;
var m_sRequestArray = null;
var m_BaseNonPopupWindow = null;

<%/* Events */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	m_sApplicationNumber = sArgArray[0];
	m_sRequestArray = sArgArray[1];
	
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function frmScreen.btnOK.onclick()
{
	window.close();
}

<% /* Functions */ %>
function PopulateScreen()
{
	<% /* Call the appropriate method and fetch XML */ %>
	ApplConditionsXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	ApplConditionsXML.CreateRequestTagFromArray(m_sRequestArray , "FINDAPPLICATIONCONDITIONSLIST");
	ApplConditionsXML.CreateActiveTag("APPLICATIONCONDITIONS");
	ApplConditionsXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	
	ApplConditionsXML.RunASP(document,"omAppProc.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = ApplConditionsXML.CheckResponse(ErrorTypes);
	
	if (ErrorReturn[0] == true)
	{	<% /* Populate the List Box */ %>
		ApplConditionsXML.CreateTagList("APPLICATIONCONDITIONS");
		
		var iNumberOfRows = ApplConditionsXML.ActiveTagList.length;
		scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows);
		ShowList(0);
		if(iNumberOfRows > 0) scTable.setRowSelected(1) ;
	}
}

function ShowList(nStart)
{
	var iCount ;
	var sConditionReference, sConditionName, sSatisfyStatus, sConditionDesc ;
	var sEditable, sConditionSeq ;
	
	for(iCount = 0 ; iCount < ApplConditionsXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		ApplConditionsXML.SelectTagListItem(iCount + nStart);
		sConditionReference = ApplConditionsXML.GetAttribute("CONDITIONREFERENCE");
		sConditionName		= ApplConditionsXML.GetAttribute("CONDITIONNAME");
		sConditionDesc		= ApplConditionsXML.GetAttribute("CONDITIONDESCRIPTION");
		sSatisfyStatus		= ApplConditionsXML.GetAttribute("SATISFYSTATUS");
		sEditable			= ApplConditionsXML.GetAttribute("EDIT");
		sConditionSeq		= ApplConditionsXML.GetAttribute("APPLNCONDITIONSSEQ");
		
		if(sSatisfyStatus == '1') sSatisfyStatus = 'Yes' ;
		else if(sSatisfyStatus == '0')sSatisfyStatus = 'No' ;
		
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sConditionReference);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sConditionName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sSatisfyStatus);
		
		<% /* Add the required attributes	*/ %>
		tblTable.rows(iCount+1).setAttribute("CONDITIONDESC", sConditionDesc);
		tblTable.rows(iCount+1).setAttribute("EDITABLE", sEditable);
		tblTable.rows(iCount+1).setAttribute("CONDITIONSEQ", sConditionSeq);
		tblTable.rows(iCount+1).setAttribute("SATISFYCONDITIONCHANGED", "0");
		tblTable.rows(iCount+1).setAttribute("SATISFYSTATUSCODE", sSatisfyStatus);
		tblTable.rows(iCount+1).setAttribute("LISTSEQ", iCount + nStart);
	}	
}

function spnTable.onclick()
{
	var iIndex = scScrollTable.getRowSelected();  // returns the table index 
	
	if (iIndex != -1)
		frmScreen.txtDetails.value = tblTable.rows(iIndex).getAttribute("CONDITIONDESC");
	else
		frmScreen.txtDetails.value = ''; 

}
-->
</script>
</BODY>
</HTML>


 
