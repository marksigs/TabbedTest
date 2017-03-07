<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<%
/*
Workfile:      pp040.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Payee Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		08/11/00	New Screen - just screen design
CL		05/03/01	SYS1920 Read only functionality added
LD		23/05/02	SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
</HEAD>

<BODY>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets */ %>
<span id="spnListScroll">
	<span style="LEFT: 310px; POSITION: absolute; TOP: 260px">
		<OBJECT data=scTableListScroll.asp id=scScrollTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 300px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>


<% /* FORMS */ %>
<form id="frmToPP050" method="post" action="PP050.asp" STYLE="DISPLAY: none"></form>
<form id="frmToPP070" method="post" action="PP070.asp" STYLE="DISPLAY: none"></form>
<form id="frmToPP200" method="post" action="PP200.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 230px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 24px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="30%" class="TableHead">Payee Name</td>	
				<td width="20%" class="TableHead">Payee Type</td>	
				<td width="30%" class="TableHead">Bank Name</td>
				<td class="TableHead">Account Number</td></tr>
			<tr id="row01">		
				<td width="30%" class="TableTopLeft">&nbsp;</td>		
				<td width="20%" class="TableTopCenter">&nbsp;</td>		
				<td width="30%" class="TableTopCenter">&nbsp;</td>		
				<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		
				<td width="30%" class="TableBottomLeft">&nbsp;</td>		
				<td width="20%" class="TableBottomCenter">&nbsp;</td>		
				<td width="30%" class="TableBottomCenter">&nbsp;</td>		
				<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>
	
	<span id="spnButtons" style="LEFT: 4px; POSITION: absolute; TOP: 200px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnAddPayee" value="Add Payee" type="button" style="WIDTH: 90px" class="msgButton">
		</span>

		<span style="LEFT: 95px; POSITION: absolute; TOP: 0px">
			<input id="btnEditPayee" value="Edit Payee" type="button" style="WIDTH: 90px" class="msgButton">
		</span> 
		
		<span style="LEFT: 190px; POSITION: absolute; TOP: 0px">
			<input id="btnAddPayment" value="Add Payment" type="button" style="WIDTH: 90px" class="msgButton">
		</span> 
		
	</span>
</div>
</form>

<% /* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 340px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/pp040Attribs.asp" -->

<%/* CODE */ %>
<script language="JScript">
<!--

var m_sMetaAction = "";
var m_sApplicationNumber = "";

var scScreenFunctions ;
var m_iTableLength = 10;
var PayeeHistoryXML = null;
var m_blnReadOnly = false;


function RetrieveContextData()
{
	//TEST
	/*
	scScreenFunctions.SetContextParameter(window,"idUserId","SR");
	scScreenFunctions.SetContextParameter(window,"idUnitId","Unit1");
	scScreenFunctions.SetContextParameter(window,"idMachineId","MC1");
	scScreenFunctions.SetContextParameter(window,"idDistributionChannelId","DIST1");
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","00042080");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber",1);
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
	scScreenFunctions.SetContextParameter(window,"idApplicationStage", Array(50, "Mortgage Application Complete"));
	*/
	// END TEST
	
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");

}

/*** EVENTS *****/


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Payee Details","PP040",scScreenFunctions)
	
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	RetrieveContextData();
	PopulateScreen();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP040");		
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}	

function frmScreen.btnEditPayee.onclick()
{	
	var iCurrentRow = scScrollTable.getRowSelected();
	
	if(iCurrentRow == -1)
	{
		window.alert("Select a fee type for making payment") ;
		return ;
	}
	
	var sPayeeHistorySeqNo = tblTable.rows(iCurrentRow).getAttribute("PayeeHistorySeqNo")
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateActiveTag("PAYEE_HISTORY")
	XML.CreateTag("PAYEEHISTORYSEQNO", sPayeeHistorySeqNo);
	
	scScreenFunctions.SetContextParameter(window,"idXML", XML.XMLDocument.xml)

	frmToPP050.submit();
}

function frmScreen.btnAddPayee.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idXML", "")
	frmToPP050.submit();
}

function frmScreen.btnAddPayment.onclick()
{
	frmToPP070.submit();
}

function btnSubmit.onclick()
{
	frmToPP200.submit();
}


/*** FUNCTIONS *****/
function PopulateScreen()
{
	PayeeHistoryXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	PayeeHistoryXML.CreateRequestTag(window,null)
	
	var tagPayeeHistory = PayeeHistoryXML.CreateActiveTag("PAYEE_HISTORY");
	PayeeHistoryXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	
	PayeeHistoryXML.RunASP(document,"FindPayeeHistoryList.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = PayeeHistoryXML.CheckResponse(ErrorTypes);

	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[1] != ErrorTypes[0]) PopulateListBox(0);	
	}
	
	if (tblTable.rows > 0)scScrollTable.setRowSelected(1);
}

function PopulateListBox(nStart)
{
	PayeeHistoryXML.CreateTagList("PAYEE_HISTORY");
	var iNumberOfRows = PayeeHistoryXML.ActiveTagList.length;

	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}

function ShowList(nStart)
{
	var iCount;
	scScrollTable.clear();		
	for (iCount = 0; iCount < PayeeHistoryXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		PayeeHistoryXML.SelectTagListItem(iCount + nStart);

		var sPayeeName			= PayeeHistoryXML.GetTagText("COMPANYNAME");
		var sPayeeType			= PayeeHistoryXML.GetTagAttribute("THIRDPARTYTYPE", "TEXT");	
		var sBankName			= PayeeHistoryXML.GetTagText("BANKNAME");	 	
		var sAccountNumber		= PayeeHistoryXML.GetTagText("ACCOUNTNUMBER");	 
		var sPayeeHistorySeqNo	= PayeeHistoryXML.GetTagText("PAYEEHISTORYSEQNO");
															  
		// Add to the search table
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sPayeeName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sPayeeType);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sBankName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), sAccountNumber);
		
		tblTable.rows(iCount+1).setAttribute("PayeeHistorySeqNo", sPayeeHistorySeqNo);
		
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}
}

-->
</script>
</BODY>
</HTML>



