<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      SP090.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Payment Sanction 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
CL		14/02/2001	SYS1982 Sceen first created
CL		05/03/01	SYS1920 Read only functionality added
CL		07/03/01	SYS2002 Added to array for SP093 
CL		07/03/01	SYS2002 Change to Sanction detail
CL		07/03/01	SYS2002 Adjustment for Popup SP093/4 functionality 
CL		08/03/01	SYS2002 Adjutment for combo population 
CL		08/03/01	SYS2002 Close routing to MN010
CL		15/03/01	SYS2059 Change to date checking 
CL		15/03/01	SYS2068	Change to update table after updates
CL		15/03/01	SYS2069 Change to reset counts
CL		20/03/01	SYS2069 Remove title
CL		21/03/01	SYS2130	Change to screen requery functionality
APS		22/03/01	SYS2146
MC		29/03/01	SYS2130 Use IssueDate not CreationDate
MC		30/03/01	SYS2081 Add Applicant Name, amend method called to improve speed and
					remove splash screen as no longer required.
SR		06/09/01	SYS2412
DRC     11/10/01    SYS2797 Changed PaymentMethod selection to combo from check boxes
SR		11/12/01	SYS3261 Compare approval/recommended userid with the one logged on but 
					not the one created payment
PSC		17/01/02	SYS3522 Amend CheckRowSelected to use scTableList.getRowSelected rather
                            than scTableList.getRowSelectedIndex
PSC		18/01/02	SYS3521 Amend UpdateDisbursementStatus to select the  correct application
							number
JLD		06/03/02	SYS4177 use the NetPaymentAmount
PSC		23/04/02    SYS4427 Return Application Fact Find data in FindSanctioningList rather
                            than separately  
STB		24/04/02	SYS4231 Ensure the correct payment is selected for Manual Payment.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

AW		24/09/02	BMIDS00072	Print Cheque Processing (BM029)
AW		24/09/02	BMIDS00513	Test 'ProcessChequepaymentsFlag' for false
GD		08/10/02	BMIDS00569	Typo in CheckRowSelected()
TW      09/10/2002	SYS5115		Modified to incorporate client validation
MV		09/10/2002	BMIDS00556	Modified SendData()
GD		14/10/2002	BMIDS00572  Resize sp091 popup slightly
GD		24/10/2002	BMIDS00703	Check combos against Validation Types, not Value IDs.
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
GD		18/11/2002	BMIDS00979  Cater for extra cheque type (cardiff and penderford cheques)
GD		21/11/2002	BMIDS01045  make sure payment amount is NETPaymentAmount, and NETPAYMNETAMOUNT passed to sp095.asp
MV		31/01/2003	BM0254 ApplicationField is not long enough - made length = 12
MV		26/03/2003	BM0088 - Amended SendData()
MV		27/03/2003	BM0474		Amended UpdateDisbursementStatus() ;frmScreen.btnManualPayment.onclick ; Created CheckLockForApplication() ; 
MV		27/03/2003	BM0474		Amended UpdateDisbursementStatus() ;frmScreen.btnManualPayment.onclick ; Created GetApplicationLockData
GD		08/04/2003	BM0477		Rework of list box multi-select functionality.
GD		08/04/2003	BM0477		Rework of list box multi-select functionality - further fixes
HMA     06/11/2003  BMIDS659    Ensure string comparason is irrespective of case.
MC		19/04/2004	CC057		Dialog resize due to Display Big Font fix.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS	Specific History:

GHun	22/07/2005	MAR10		Adjust layout to remove scrollbar
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM	Specific History:
IK		08/05/2006	EP453		modify layout
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets - remove any which are not required */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>

<object data="scTable.htm" height="1" id="scTable" 
style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex="-1" type="text/x-scriptlet" 
viewastext></object>

<% /* Specify Forms Here */ %>
<form id="frmToSP093" method="post" action="SP093.asp" style="DISPLAY: none"></form>
<form id="frmToSP094" method="post" action="SP094.asp" style="DISPLAY: none"></form>
<form id="frmToMN010" method="post" action="MN010.asp" style="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<% /* Start of form */ %>

<form id="frmScreen"  validate ="onchange" mark year4>
	<div id="divBackground" style="HEIGHT: 440px; margin-TOP: 50px; WIDTH: 604px" class="msgGroup">
		<div id="divPaymentSanctionList" style="margin-TOP: 10px">
			<table id="tblPaymentSanction" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">				
				<tr id="rowTitles">
					<td width="5%" class="TableHead">&nbsp;</td>
					<td width="10%" class="TableHead">Application Number</td>
					<td width="20%" class="TableHead">Lead Applicant Name</td>
					<td width="19%" class="TableHead">Payee</td>
					<td width="19%" class="TableHead">Payment Type</td>
					<td width="5%" class="TableHead">Amount</td>
					<td width="5%" class="TableHead">Payment Method</td>
					<td width="15%" class="TableHead">Issue Date</td>
					<td width="1%" class="TableHead">&nbsp;</td>
					<!--<td width="3%" class="TableHead">&nbsp;</td>-->
				</tr>
		
				<tr id="row01"><td width="5%" class="TableTopLeft">&nbsp;&nbsp;</td><td width="10%" class="TableTopCenter">&nbsp;</td><td width="20%" class="TableTopCenter">&nbsp;</td><td width="19%" class="TableTopCenter">&nbsp;</td><td width="19%" class="TableTopCenter">&nbsp;</td><td width="5%" class="TableTopCenter">&nbsp;</td><td width="5%" class="TableTopCenter">&nbsp;</td><td width="15%" class="TableTopCenter">&nbsp;</td><td width="1%" class="TableTopRight">&nbsp;</td></tr>
				<tr id="row02"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row03"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row04"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row05"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row06"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row07"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row08"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row09"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row10"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row11"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row12"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row13"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row14"><td width="5%" class="Tableleft">&nbsp;&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td width="20%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="19%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td><td width="15%" class="TableCenter">&nbsp;</td><td width="1%" class="TableRight">&nbsp;</td></tr>
				<tr id="row15"><td width="5%" class="TableBottomLeft">&nbsp;&nbsp;</td><td width="10%" class="TableBottomCenter">&nbsp;</td><td width="20%" class="TableBottomCenter">&nbsp;</td><td width="19%" class="TableBottomCenter">&nbsp;</td><td width="19%" class="TableBottomCenter">&nbsp;</td><td width="5%" class= "TableBottomCenter">&nbsp;</td><td width="5%" class= "TableBottomCenter">&nbsp;</td><td width="15%" class= "TableBottomCenter">&nbsp;</td><td width="1%" class="TableBottomRight">&nbsp;</td></tr>
			</table>
		</div>
		<!--GD Added btnSelectAll and btnDeSelectAll-->
		<div id="spnSelectionOptions" style="position:relative;top:0;left:0;FONT-WEIGHT: normal; margin-TOP: 8px" class="msgLabel">
			<!--
			<span style="FONT-WEIGHT: bold; LEFT: 0px; POSITION: absolute; TOP: -23px">
				Selection Options
			</span>
			-->
			<span style="margin-LEFT: 1px; ">
				<input id="btnSelectAll" value="Select All" type="button" style="WIDTH: 80px" class="msgButton">
			</span>
			
			<span style="margin-LEFT:4px">
				<input id="btnDeSelectAll" value="De-Select All" type="button" style="WIDTH: 80px" class="msgButton">
			</span>

			<span id="spnTableListScroll" style="right:8px; POSITION: absolute; TOP: 0px; display:none">
				<object data="scTableListScroll.asp" id="scTableList" style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex="-1" type="text/x-scriptlet" viewastext></object>
			</span>
		</div>		
		
		<!-- View Options  Title  -->
		<div id="spnInput" style="FONT-WEIGHT: normal; margin-TOP: 16px" class="msgLabel">
			<!-- View Options  Line 1  -->
			<span style="FONT-WEIGHT: bold">
				View Options
			</span>
		</div>
		
		<div style="FONT-WEIGHT: normal; position:relative; LEFT: 0px; TOP: 0px; margin-top:8px" class="msgLabel">
			Payment status
			<span style="LEFT: 82px; POSITION: absolute; TOP: -3px">
				<select id="cboPaymentStatus" style="WIDTH: 160px" class="msgCombo"></select>
			</span>
			<span style="HEIGHT: 20px; LEFT: 250px; POSITION: absolute; TOP: 0px; WIDTH: 93px" class="msgLabel">
				Application number
				<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
					<input id="txtApplicationNumber" maxlength="12" style="WIDTH: 100px" class="msgTxt">
				</span>
			</span>
			<span style="LEFT: 460px; POSITION: absolute; TOP: 0px" class="msgLabel">
				Issue date
			</span>
			<span style="right:8px; POSITION: absolute; TOP: -3px">
				<input id="txtIssueDate" maxlength="10" style="WIDTH: 70px" class="msgTxt">
			</span>
		</div>
		<!-- View Options  Line 2  -->
		<div style="FONT-WEIGHT: normal; position:relative; LEFT: 0px; TOP: 0px; margin-top:12px" class="msgLabel">
						
			Payment method
<% /* AQR SYS2797 DRC */%>
			<span style="LEFT: 82px; POSITION: absolute; TOP: -3px">
				<select id="cboPaymentMethod" style="HEIGHT: 20px; WIDTH: 161px" class="msgCombo"></select>
			</span>
			
			<span style="right: 8px; POSITION: absolute; TOP: 0px">
				<input id="btnSearch" value="Search" type="button" 
					style="WIDTH: 80px"  class="msgButton">
			</span>
		</div>
		<!-- Buttons -->
		<div id="spnButtons" style="position:relative;top:0px;left:0px;margin-TOP:24px">
			<span style="LEFT: 1px; POSITION: absolute; TOP: 0px">
				<input id="btnSanction" value="Sanction" type="button" 
					style="WIDTH: 80px" class="msgButton">
			</span>
			<span style="LEFT: 104px; POSITION: absolute; TOP: 0px">
				<input id="btnUnSanction" value="Un sanction" type="button" 
					style="WIDTH: 80px" class="msgButton">
			</span>
			<span style="right:8px; POSITION: absolute; TOP: 0px">
				<input id="btnManualPayment" value="Manual Payment" type="button" 
					style="WIDTH: 100px" class="msgButton">
			</span>
		</div>
	</div>
</form>


<% /* End of form */ %>
<!-- #include FILE="attribs/sp090attribs.asp" -->


<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 0px; POSITION: relative; TOP: 4px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript" type="text/javascript">
<!--

var m_sMetaAction = null;
var m_iTableLength = 15;
var m_iTotalRecords = 0;
var XML = null;
var XMLApplFactFind = null ;
var scScreenFunctions;
var m_sReadOnly = null;
var m_sContext;
var m_SelectedPaymentMethods;
var m_PProcManualPaymentRole;
var m_PProcSanctionRole;
var m_ProcessChequesFlag;
var m_SanctionedYN;
var m_SelectedRows = null;
var m_Old_ArrayLength = null;
var m_LastRowSelected = null;
var m_nAmount = 0;
var m_nRows = 0;
var m_sIdUserLoggedOn = "" ;
var m_sUnsanctionedId;
var m_sSanctionedId;
var m_sPaidId;
var m_sChequeId;
var XMLCombos;
<%//GD BM0477 START%>
var m_aCheckBoxArray;
var m_bSancUnsancMixFound = false; <%//flag indicating if the resultset after a search contains a mix of 'U' and 'S' payments %>
<%
/*
var m_sChecked = "X";
var m_sUnChecked = "  ";
*/
%>

<%//GD BM0477 END %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	<%// Added by automated update TW 09 Oct 2002 SYS5115 %>
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<%
	//GD BM0477 START
	//Initialise CheckBox array - this will need to be done after a search , too.
	%>
	m_aCheckBoxArray = new Array();
	<%//GD BM0477 END %>
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)	
%>	
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Payment Sanction","SP090",scScreenFunctions);	
	InitialiseScreen();
	SetMasks();
	Validation_Init();
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLApplFactFind = new top.frames[1].document.all.scXMLFunctions.XMLObject();
<%	//Populate the combo box
	// check response if ok then continue
%>
	if(GetComboList()) 
	{
		<% /* Load Global Parameter PProcManualPaymentRole  */ %>
		m_PProcManualPaymentRole = XML.GetGlobalParameterAmount(document, "PProcManualPaymentRole");
		<% /* Load Global Parameter PProcSanctionRole  */  %>
		m_PProcSanctionRole = XML.GetGlobalParameterAmount(document, "PProcSanctionRole");
		m_ProcessChequesFlag = XML.GetGlobalParameterBoolean(document, "ProcessChequepaymentsFlag");
		// Show Blank Screen on entry 
	}
	
	<% /* scScreenFunctions.HideCollection(divStatus);  */ %>
	scScreenFunctions.ShowCollection(spnTableListScroll);
	scScreenFunctions.ShowCollection(frmScreen);
	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();

	
}	

function InitialiseScreen()
{
	<% /* MO - BMIDS00807 */ %>
	<% /* frmScreen.txtIssueDate.value = scScreenFunctions.DateToString(scScreenFunctions.GetSystemDateTime()); */ %>
	frmScreen.txtIssueDate.value = scScreenFunctions.DateToString(scScreenFunctions.GetAppServerDate(true));
	m_sIdUserLoggedOn = scScreenFunctions.GetContextParameter(window,"idUserId","");
	
<% /*  AQR SYS2797
	frmScreen.chkPaymentMethod1.checked = true ;
	frmScreen.chkPaymentMethod2.checked = true ;
	frmScreen.chkPaymentMethod3.checked = true ;
	frmScreen.chkPaymentMethod4.checked = true ;
	
	PaymentMethods();
*/	%>
}

function Search()
{
	<%
	// Don't do the table clear here.
	// Wait until we know how many rows the table has then call the local clear() method
	//scTableList.clear();		
	//clear(); 
	%>
	
	<%//IK EP453 %>		
	spnTableListScroll.style.display = "none";
	
	XML.ResetXMLDocument();	
	<% /*   AQR SYS2797 PaymentMethods(); */ %>
	<% // GD BM0477 START %>
	m_bSancUnsancMixFound = false;
	var bSancFound = false;
	var bUnSancFound = false;
	<% // GD BM0477 END %>	
	var XMLRequestTag = XML.CreateRequestTag(window,"FindSanctioningList"); 
	XML.CreateActiveTag("FINDSANCTIONINGLIST");
	if(frmScreen.txtApplicationNumber.value != '') XML.SetAttribute("APPLICATIONNUMBER", frmScreen.txtApplicationNumber.value);
	XML.SetAttribute("_COMBOLOOKUP_","1");
	
	if(frmScreen.txtApplicationNumber.value == '')
	{						
		if (frmScreen.cboPaymentMethod.value != "")
		{
		<% /*   AQR SYS2797 */ %>
			XML.SetAttribute("PAYMENTMETHOD", frmScreen.cboPaymentMethod.value);
		}

		if (frmScreen.cboPaymentStatus.value != "")
		{
			XML.SetAttribute("PAYMENTSTATUS", frmScreen.cboPaymentStatus.value);
		}	
		if (frmScreen.txtIssueDate.value !="")
		{
			XML.SetAttribute("ISSUEDATE", frmScreen.txtIssueDate.value);
		}	
	}
	// 	XML.RunASP(document,"PaymentProcessingRequest.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
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
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[0] == true)
		{
			<%/* SR 06/07/01 : SYS2412 - Get FactFindData for all the applications returned from the 
							   above method  */
			%>	
			<%// GD BM0477 START %>
			//Check if any of the items in the list have a PaymentStatus of Unsanctioned
			bUnSancFound = (XML.SelectNodes("//PAYMENT[@PAYMENTSTATUS = '" + m_sUnsanctionedId + "']").length > 0)
			
			//Check if any of the items in the list have a PaymentStatus of Sanctioned
			bSancFound = (XML.SelectNodes("//PAYMENT[@PAYMENTSTATUS = '" + m_sSanctionedId + "']").length > 0)		
			<%// GD BM0477 END %>			
			if (XML.CreateTagList("PAYMENT") != null) m_iTotalRecords = XML.ActiveTagList.length;
			else m_iTotalRecords = 0;
			
			<% //GD BM0477 START - reset array after each search %>
			m_aCheckBoxArray = new Array();
			clear(m_iTotalRecords);				
			<% //GD BM0477 END %>
		
			frmScreen.scTableList.initialiseTable(tblPaymentSanction, 0, "", Showlist, m_iTableLength, m_iTotalRecords)
			<%//GD BM)477 not applicable scTableList.EnableMultiSelectTable();%>
			Showlist(0);
		}
		else	
		{
			m_iTotalRecords = 0;
			clear(m_iTotalRecords);	
			frmScreen.scTableList.initialiseTable(tblPaymentSanction, 0, "", Showlist, m_iTableLength, m_iTotalRecords)			
			ResetBlankRows(0);
			alert("No records have been found. Please amend your search criteria");
		}
	}	
	<%// BM0477 Set a global flag is resultset contains a mix of 'S' and 'U' payments %>
	m_bSancUnsancMixFound = (bSancFound && bUnSancFound);
	
	frmScreen.btnSanction.disabled = true;
	frmScreen.btnUnSanction.disabled = true;
	frmScreen.btnManualPayment.disabled = true;
	ErrorTypes = null;
	ErrorReturn = null;
	<%//GD BM0477 START - reset array after each search %>
	m_aCheckBoxArray = new Array();
	<%//GD BM0477 END %>
	
}
	
function Showlist(iOffset)
{
	var iCount ;
	var iRowCount = 1;
	var varApplicationNumber, varAmount, varPaymentMethod, varPaymentMethod_Text
	var varUserID, varCompanyName, varPaymentType, varPaymentStatus
	var varIssueDate, varLeadApplicantName
	var sApplApprovalUserId, sApplRecommendedUserId, sCondition ;
	<%//GD BM0477 START %>
	var varPaymentStatusID;
	<%//GD BM0477 END	%>	
	for (iCount=0; iCount < m_iTotalRecords && iCount < m_iTableLength;iCount++)
	{			
		if (XML.SelectTagListItem(iCount+iOffset) == true)
		{
			XML.SetAttribute("TABLEROW", iRowCount + iOffset);
			varApplicationNumber = XML.GetAttribute("APPLICATIONNUMBER");
			varAmount = XML.GetAttribute("NETPAYMENTAMOUNT"); // JLD SYS4177
			varPaymentMethod = XML.GetAttribute("PAYMENTMETHOD");
			varPaymentMethod_Text = XML.GetAttribute("PAYMENTMETHOD_TEXT");
			varUserID = XML.GetAttribute("USERID");
			varCompanyName = XML.GetAttribute("COMPANYNAME"); 
			varPaymentType = XML.GetAttribute("PAYMENTTYPE_TEXT");
			<%//GD BM0477 START %>
			<%//varPaymentStatus = XML.GetAttribute("PAYMENTSTATUS_TEXT");%>
			varPaymentStatusID = XML.GetAttribute("PAYMENTSTATUS");

			if (varPaymentStatusID == m_sSanctionedId)
			{
				varPaymentStatus = "S";
			} else if(varPaymentStatusID == m_sUnsanctionedId)
			{
				varPaymentStatus = "U";
			} else
			{
				varPaymentStatus = XML.GetAttribute("PAYMENTSTATUS_TEXT");
			}
 
			<%//GD BM0477 END %>
					
			varIssueDate = XML.GetAttribute("ISSUEDATE");			
			varLeadApplicantName = XML.GetAttribute("FIRSTFORENAME") + " " + XML.GetAttribute("SURNAME");
			sApplApprovalUserId = XML.GetAttribute("APPLICATIONAPPROVALUSERID")
			sApplRecommendedUserId = XML.GetAttribute("APPLICATIONRECOMMENDEDUSERID")			

			if(m_aCheckBoxArray[GetArrayIndex(iRowCount,iOffset + 1)] == "1")
			{
				<%//tblPaymentSanction.rows(iRowCount).cells(0).innerText = m_sChecked;%>
				SetChecked(tblPaymentSanction.rows(iRowCount).cells(0),iRowCount);
				
			} else
			{
				<%//tblPaymentSanction.rows(iRowCount).cells(0).innerText = m_sUnChecked;%>
				SetUnChecked(tblPaymentSanction.rows(iRowCount).cells(0), iRowCount);
			}

			<%// TABLEROW is in range : 1..15%>
			
			<%//GD BM0477 END			%>
			scScreenFunctions.SizeTextToField(tblPaymentSanction.rows(iRowCount).cells(1), varApplicationNumber);
			scScreenFunctions.SizeTextToField(tblPaymentSanction.rows(iRowCount).cells(2), varLeadApplicantName);
			scScreenFunctions.SizeTextToField(tblPaymentSanction.rows(iRowCount).cells(3), varCompanyName);
			scScreenFunctions.SizeTextToField(tblPaymentSanction.rows(iRowCount).cells(4), varPaymentType);
			scScreenFunctions.SizeTextToField(tblPaymentSanction.rows(iRowCount).cells(5), varAmount);
			scScreenFunctions.SizeTextToField(tblPaymentSanction.rows(iRowCount).cells(6), varPaymentMethod_Text);
			scScreenFunctions.SizeTextToField(tblPaymentSanction.rows(iRowCount).cells(7), varIssueDate);
			scScreenFunctions.SizeTextToField(tblPaymentSanction.rows(iRowCount).cells(8), varPaymentStatus);			

			
				
			iRowCount++;
			
			<%/* Add attribute USERID (value = varUserID) to row 	*/ %>
			tblPaymentSanction.rows(iCount+1).setAttribute("UserID", varUserID);
			tblPaymentSanction.rows(iCount+1).setAttribute("SelectedRow");
			<%/* Add attribute PAYMENTMETHOD (value = varPaymentMethod to row) 	*/ %>
			tblPaymentSanction.rows(iCount+1).setAttribute("PaymentMethod", varPaymentMethod);
			tblPaymentSanction.rows(iCount+1).setAttribute("ApplicationApprovalUserId", sApplApprovalUserId);
			tblPaymentSanction.rows(iCount+1).setAttribute("ApplicationRecommendedUserId", sApplRecommendedUserId);
		}
	}	
	if(iCount < (m_iTableLength))
	{
		ResetBlankRows(iCount);
	}

	<%//IK EP453 %>		
	if(m_iTotalRecords > m_iTableLength) spnTableListScroll.style.display = "block";
		
}	
<% /*  AQR SYS2797function PaymentMethods()
{
	m_SelectedPaymentMethods = "";
	if(frmScreen.chkPaymentMethod1.checked)
	{
		m_SelectedPaymentMethods = "CH";									
	}
	if(frmScreen.chkPaymentMethod2.checked )				
	{
		if (m_SelectedPaymentMethods.length != 0)
		{
			m_SelectedPaymentMethods = m_SelectedPaymentMethods +",";
		}
		m_SelectedPaymentMethods = m_SelectedPaymentMethods +"YC";
	}
	if(frmScreen.chkPaymentMethod3.checked )
	{	
		if (m_SelectedPaymentMethods.length != 0)
		{
			m_SelectedPaymentMethods = m_SelectedPaymentMethods +",";
		}
		m_SelectedPaymentMethods = m_SelectedPaymentMethods +"B";
	}
	if(frmScreen.chkPaymentMethod4.checked)
	{
		if (m_SelectedPaymentMethods.length != 0)
		{  
			m_SelectedPaymentMethods = m_SelectedPaymentMethods +",";
		}
		m_SelectedPaymentMethods = m_SelectedPaymentMethods +"TR";
	}
}
*/%>
function frmScreen.btnManualPayment.onclick()
{
	<% /* PProcManualPaymentRole  */ %>
	<%//GD BM0477 aSelectedRows = scTableList.getArrayofRowsSelected() ;%>
	aSelectedRows = GetArrayOfRowsSelected();	
	var nAmount = 0;
	var nApplicant = 0;
	<% /* Check that only one row has been selected  */  %>
	if(aSelectedRows.length != "1")
	{
		alert("One row at a time may be selected for manual payment");
		return;
	}
	else	  
	{
		<% /* Create the XML REQUEST tag  */  %>
		XMLRows = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var XMLRequestTag = XMLRows.CreateRequestTag(window,"UpdateDisbursement"); 
	
		for(var iRowIndex = 0; iRowIndex < aSelectedRows.length; ++iRowIndex)
		{
			<% /* SYS4231 - Ensure the correct payment is selected */ %>
			var iIndex = aSelectedRows[iRowIndex] + 1;<% //GD BM0477 ADDED %>
			var sPattern = "//PAYMENT[@TABLEROW='" + iIndex + "']";
			var xmlTag = XML.SelectSingleNode(sPattern);
			<% /* SYS4231 - End */ %>
			
			if (xmlTag != null)
			{		
				
				<% /* Select the value of the APPLICATIONNUMBER attribute  */  %>
				var varApplicationNumber = XML.GetAttribute("APPLICATIONNUMBER")
				var bOK =  GetApplicationLockData(varApplicationNumber); 
				if (bOK  == true)
				{
					alert('The Application : ' + varApplicationNumber + ' is locked by another user. Please try later');
					return;
				}
					
				<% /* Select the value of the PAYMENTAMOUNT attribute  */  %>
				var varPaymentAmount = XML.GetAttribute("AMOUNT")
				<% /* Select the value of the PAYMENTSTATUS attribute   */  %>
				var varPaymentStatus = XML.GetAttribute("PAYMENTSTATUS");
				if(varPaymentStatus == m_sUnsanctionedId) // Unsanctioned
				{
					alert("Payment must be sanctioned before it can be paid");
					<%//UncheckRow();%> <% /* Deselect the current record */ %>
					
					return;
				}
				else
				{
					if(varPaymentStatus != m_sPaidId)//Paid
					{
						<% //Get the value of the PAYMENTSEQUENCENUMBER attribute
						%>
						var varPaymentSequence = XML.GetAttribute("PAYMENTSEQUENCENUMBER");
						<%//GD BMIDS01045 nAmount =  parseFloat(XML.GetAttribute("AMOUNT"));%>
						nAmount =  ParseFloatSafe(XML.GetAttribute("NETPAYMENTAMOUNT"));
						nApplicant = varApplicationNumber;
						<% //Now append the new rows to XMLRows 
						%>
						XMLRows.ActiveTag = XMLRequestTag;
						XMLRows.CreateActiveTag("PAYMENTRECORD");
						XMLRows.SetAttribute("APPLICATIONNUMBER", varApplicationNumber); 
						XMLRows.SetAttribute("CHEQUENUMBER",""); 
						XMLRows.SetAttribute("PAYMENTSEQUENCENUMBER",varPaymentSequence); 
						XMLRows.CreateActiveTag("DISBURSEMENTPAYMENT");
						XMLRows.SetAttribute("PAYMENTSTATUS",m_sPaidId); 
						XMLRows.SetAttribute("TTREFNUMBER1",""); 
						XMLRows.SetAttribute("TTREFNUMBER2",""); 
							
							
						<%//UncheckRow();%> <% /* Deselect the current record */ %>
						if (SendManualPaymentData(nAmount,nApplicant,XMLRows) == true) {
							Search();
						}
						return;
					}
					else
					{
						alert("Disbursement has already been paid");
						<% //Move back to the PAYMENTRECORD tag
						   //XML.ActiveTag = ActiveTag;
						%>
						return;
					}
				}	
			}
		}
	}
}	
function ParseFloatSafe(nNumber)
{
	<%//Handle the case where nNumber = "" %>
	if(nNumber =="") return(0)
	else
	return(parseFloat(nNumber));
}
<%
//function UncheckRow()
//{	
	//GD START REMOVED
	//var iIndex = scTableList.getRowSelectedIndex();  
	//var sPattern = "//PAYMENT[@TABLEROW='" + iIndex + "']";
	//var xmlTag = XML.SelectSingleNode(sPattern);  
	//if (xmlTag != null)
	//{
		//GD scTableList.setMultiRowUnselected(scTableList.getRowSelected());
	//	return;
	//}
	//GD END REMOVED
	//GD ADDED START
	//document.all(sCheckBoxID).checked = false;
	
	//GD ADDED END
//}		
%>
function CheckRowSelected()
{	
	<% /*	if (CheckUserRole()==false) return;  */  %>
		
	var iIndex = frmScreen.scTableList.getRowSelected();  // returns the table index 
	
	<% /* SR 11-12-01 : SYS3261 Compare approval/recommended userId with user logged on 
	var sPattern = "//PAYMENT[@TABLEROW='" + iIndex + "']";	
	var xmlTag = XML.SelectSingleNode(sPattern);  // select the correct tag  
	*/ %>
	
	if (iIndex != -1)
	{
		<% /* BMIDS659  Compare strings irrespective of case */ %>
		
		var sApplApprovalUserID    = (tblPaymentSanction.rows(iIndex).getAttribute("ApplicationApprovalUserId")).toUpperCase();
		var sApplRecommendedUserId = (tblPaymentSanction.rows(iIndex).getAttribute("ApplicationRecommendedUserId")).toUpperCase();
		var sUserLoggedOn          =  m_sIdUserLoggedOn.toUpperCase();
			
		<% /* var sUserId	= XML.GetAttribute("USERID"); */ %>
		
		if((sUserLoggedOn == sApplApprovalUserID) || (sUserLoggedOn == sApplRecommendedUserId)) 
		{
			//GD BMIDS00569 - Typo error
			alert("You cannot sanction/unsanction a case you approved/recommended.");
			<%//Current_ArrayLength = m_Old_ArrayLength;%>
			<%//GD BM0477 scTableList.setMultiRowUnselected(scTableList.getRowSelected());%>
			return(false);
		}

		var iIndex = frmScreen.scTableList.getRowSelectedIndex();  // returns the table index 
		if (iIndex != null)
		{
			frmScreen.btnSanction.disabled = false;
			frmScreen.btnUnSanction.disabled = false;
			frmScreen.btnManualPayment.disabled = false;
		}
		else
		{
			frmScreen.btnSanction.disabled = true;
			frmScreen.btnUnSanction.disabled = true;
			frmScreen.btnManualPayment.disabled = true;
		}
	}
	<%//GD BM0477 ADDED %>
	return(true);
}	
function divPaymentSanctionList.ondblclick()
{

	var iIndex = frmScreen.scTableList.getRowSelected();  // returns the table index 
	var iFirstVisible = frmScreen.scTableList.getFirstVisibleRecord();	

	if (iIndex != -1)
	{
			iArrayIndex = GetArrayIndex(iIndex,iFirstVisible);
			if(m_aCheckBoxArray[iArrayIndex] == "1")
			{
				<%//Allow an Uncheck without a CheckRowSelected%>
				m_aCheckBoxArray[iArrayIndex] = "";
				<%//tblPaymentSanction.rows(iIndex).cells(0).innerText = m_sUnChecked;%>
				SetUnChecked(tblPaymentSanction.rows(iIndex).cells(0), iIndex);
			} else
			{
				if (CheckRowSelected() == true)
				{
					m_aCheckBoxArray[iArrayIndex] = "1";
					<%//tblPaymentSanction.rows(iIndex).cells(0).innerText = m_sChecked;%>
					SetChecked(tblPaymentSanction.rows(iIndex).cells(0), iIndex);
				}
			}
	}
	document.selection.empty();
}
	
<%//GD BM0477 START%>
function GetArrayIndex(iRowSelected,iFirstVisible)
{
	return(iRowSelected + iFirstVisible - 2)
}

function SelectDeSelectAll(sFlag)
{
	var iIndex;
	var iTotalRecords = frmScreen.scTableList.getTotalRecords()
	for(iIndex = 0;iIndex < iTotalRecords;iIndex++)
	{
		m_aCheckBoxArray[iIndex] = sFlag;
		if(sFlag == "1")//Doing 'Select All'
		{
			//Set the TABLEROW attribute on all PAYMENT elements in the XML
			if (XML.SelectTagListItem(iIndex) == true)
			{
				XML.SetAttribute("TABLEROW", iIndex + 1);
			}
		}
	}
}
<%//GD BM0477 END%>
function frmScreen.btnSanction.onclick()
{
		UpdateDisbursementStatus(m_sSanctionedId);
}		

function frmScreen.btnUnSanction.onclick()
{
		UpdateDisbursementStatus(m_sUnsanctionedId);
}		

function frmScreen.btnSearch.onclick()
{
	Search();
}

function UpdateDisbursementStatus(sNewStatus)
{
	var	 nFirstPayment = 0;
	var  bIsCheque = false;
	m_nRows = 0; 
	var nAmount = 0;
	var aSelectedRows = 0;
	var bOK;
	var sPattern;
	var xmlTag;
	var varApplicationNumber;
	
	if(sNewStatus == m_sUnsanctionedId)
	{
		<% /* set Pass MetaAction UnSanction */ %>
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Unsanction");
	}
	else if(sNewStatus == m_sSanctionedId)
	{
		<% /* set Pass MetaAction UnSanction */ %>
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Sanctioned");
	}
	else
		return;
		
	<%// GD BM0477 aSelectedRows = scTableList.getArrayofRowsSelected() ;	 %>
	aSelectedRows = GetArrayOfRowsSelected();
	<% /* MV - 27/03/2003 - BM0474  */%>
	<%//GD BM0477 START %>
	if (aSelectedRows.length == 0)
	{
		//Nothing to be processed
		alert("You have not 'checked' any rows to be processed.");
		return;
	}
	//Check the result set to ensure user allowed to Sanction/Unsanction all the payments selected.
	var bCanSancUnsanc = true;
	var iPosn;
	var xmlElem, xmlTemp;
	var sApplicationApprovalUserId,sApplicationRecommendedUserId;
	for(var iIndex = 0;((iIndex < aSelectedRows.length) && (bCanSancUnsanc == true));iIndex++)
	{
		iPosn = aSelectedRows[iIndex];

		XML.SelectTagListItem(iPosn);
			
		<% /* BMIDS659  Compare strings irrespective of case */ %>
					
		sApplicationApprovalUserId    = (XML.GetAttribute("APPLICATIONAPPROVALUSERID")).toUpperCase();
		sApplicationRecommendedUserId = (XML.GetAttribute("APPLICATIONRECOMMENDEDUSERID")).toUpperCase();
		var sIdUserLoggedOn = m_sIdUserLoggedOn.toUpperCase();

		if((sIdUserLoggedOn == sApplicationApprovalUserId) || (sIdUserLoggedOn == sApplicationRecommendedUserId)) 
		{
			bCanSancUnsanc = false;
		}
	}
	if(bCanSancUnsanc==false)
	{
		alert("You cannot sanction/unsanction a case you approved/recommended.");
		return;	
	}
	
	<%//GD BM0477 END %>
	for(var iRowIndex = 0; iRowIndex < aSelectedRows.length ; ++iRowIndex)
	{
		iIndex = aSelectedRows[iRowIndex] + 1;<%//GD BM0477 ADDED + 1 %>
		sPattern = "//PAYMENT[@TABLEROW='" + iIndex + "']";
		xmlTag = XML.SelectSingleNode(sPattern);
		if (xmlTag != null)
		{
			<%// GD BM0477 Check that User is allowed to Sanction the payment %>
	
			varApplicationNumber = XML.GetAttribute("APPLICATIONNUMBER");
			bOK =  GetApplicationLockData(varApplicationNumber)
			if (bOK  == true)
			{
				alert('The Application : ' + varApplicationNumber + ' is locked by another user. Pleaes try later');
				return;
			}
		}
	
	}
	
	
		
	if((m_ProcessChequesFlag < 1) && (sNewStatus == m_sSanctionedId))
	{
		for(var iRowIndex = 0; iRowIndex < aSelectedRows.length ; ++iRowIndex)
		{
		
			iIndex = aSelectedRows[iRowIndex] + 1; <%//GD BM0477 ADDED + 1 %>
			var sPattern = "//PAYMENT[@TABLEROW='" + iIndex + "']";
			var xmlTag = XML.SelectSingleNode(sPattern);
			if (xmlTag != null)
			{
				if(iRowIndex == 0 ) 
				{
					nFirstPayment = XML.GetAttribute("PAYMENTMETHOD");
					//GD BMIDS00979							
					//if(nFirstPayment == m_sChequeId) {bIsCheque = true;}
					if (IsCheque(nFirstPayment)) {bIsCheque = true;}
					
				}
				else
				{
					<% /* Select the value of the PAYMENTMETHOD attribute  */ %>
					var varPayMethod = XML.GetAttribute("PAYMENTMETHOD");
					//if((varPayMethod != nFirstPayment) && (varPayMethod == m_sChequeId || nFirstPayment == m_sChequeId))// Cheque
					//GD BMIDS00979	
					if((varPayMethod != nFirstPayment) && (IsCheque(varPayMethod) || IsCheque(nFirstPayment)))// Cheque
					{
						alert("One or more payment(s) selected has a payment method of cheque. Cannot sanction cheque payment with other method of payments");
						return false;
					}
				}
			}
		}
	}
	//	AW	05/08/02	BMID029  - End	
	<% /* Create the XML REQUEST tag  */ %>
	XMLRows = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var XMLRequestTag = XMLRows.CreateRequestTag(window,"UpdateDisbursement");
	
	for(var iRowIndex = 0; iRowIndex < aSelectedRows.length ; ++iRowIndex)
	{
		
		iIndex = aSelectedRows[iRowIndex] + 1;//GD ADDED + 1
		var sPattern = "//PAYMENT[@TABLEROW='" + iIndex + "']";
		var xmlTag = XML.SelectSingleNode(sPattern);
		if (xmlTag != null)
		{
			<% /* Select the value of the APPLICATIONNUMBER attribute  */ %>
			var varApplicationNumber = XML.GetAttribute("APPLICATIONNUMBER");

			<% /* Select the value of the PAYMENTSTATUS attribute  */ %>
			var varPaymentStatus = XML.GetAttribute("PAYMENTSTATUS");
			<% /* Select the value of the PAYMENTSEQUENCENUMBER attribute  */ %>
			var varPaymentSequence = XML.GetAttribute("PAYMENTSEQUENCENUMBER");
			<% /* Check that all the rows are correct  */ %>
				
			if(CheckPaymentStatus(sNewStatus,varPaymentStatus)== true)
			{	
				//nAmount +=  parseFloat(XML.GetAttribute("AMOUNT"));
				//GD BMIDS01045
				nAmount +=  ParseFloatSafe(XML.GetAttribute("NETPAYMENTAMOUNT"));
				m_nRows++; 
				<% /* Now append the new rows to XMLRows  */ %>
				XMLRows.ActiveTag = XMLRequestTag;
				XMLRows.CreateActiveTag("PAYMENTRECORD");
				XMLRows.SetAttribute("APPLICATIONNUMBER", varApplicationNumber); 
				XMLRows.SetAttribute("PAYMENTSEQUENCENUMBER",varPaymentSequence);
				//	AW	05/08/02	BMID029
				if (bIsCheque == true)
				{
					XMLRows.SetAttribute("AMOUNT", ParseFloatSafe(XML.GetAttribute("AMOUNT")));
					//GD BMIDS01045
					XMLRows.SetAttribute("NETPAYMENTAMOUNT", ParseFloatSafe(XML.GetAttribute("NETPAYMENTAMOUNT")));
					XMLRows.SetAttribute("REPRINTSTATUS", "0");
				} 
				XMLRows.CreateActiveTag("DISBURSEMENTPAYMENT");
				XMLRows.SetAttribute("PAYMENTSTATUS",sNewStatus); 
			}
		}			
	}
	//	AW	05/08/02	BMID029, extra parameter 'bIsCheque'
	if (m_nRows > 0)SendData(nAmount,m_nRows,XMLRows, bIsCheque);

}

function GetComboList()
{

	XML.ResetXMLDocument();	
	var ComboLists = new Array("PaymentStatus", "PaymentMethod");
	var bSuccess = false;

	if(XML.GetComboLists(document,ComboLists) == true)
	{
		//bSuccess = true;
		bSuccess = XML.PopulateCombo(document,frmScreen.cboPaymentStatus,"PaymentStatus", false);
		bSuccess = bSuccess && XML.PopulateCombo(document,frmScreen.cboPaymentMethod,"PaymentMethod", false);
	}
	else
	{ 
		bSuccess = false;
	}
	if(!bSuccess)
	{
		alert('Failed to populate combos');
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	} else
	{
		//Gte ValidationTypes required for further processing
		XMLCombos = XML.GetComboListXML("PaymentStatus");
		m_sUnsanctionedId	= XML.GetComboIdForValidation("PaymentStatus", "U", XMLCombos)
		m_sSanctionedId	    = XML.GetComboIdForValidation("PaymentStatus", "S", XMLCombos)
		m_sPaidId	        = XML.GetComboIdForValidation("PaymentStatus", "P", XMLCombos)
		
		XMLCombos    = XML.GetComboListXML("PaymentMethod");
		m_sChequeId         = XML.GetComboIdForValidation("PaymentMethod", "CH", XMLCombos)
		
		//Adjust Combo Contents For PaymentStatus
		//REMOVE All OPTIONs apart from Unsanctioned and Sanctioned

		var nIndex = 0;
		var sValueID;

		for(nIndex = frmScreen.cboPaymentStatus.length-1; nIndex > 0 ; nIndex--)
		{
			
			if (!(  (  scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentStatus, nIndex, "U")  ) || (  scScreenFunctions.IsOptionValidationType(frmScreen.cboPaymentStatus, nIndex, "S")  )  ))
			
			{
				//Remove
				frmScreen.cboPaymentStatus.remove(nIndex);
			}
		}
		
		//ADD new OPTION 'Sanctioned / Unsanctioned' to PaymentStatus Combo
		var TagOPTION	= document.createElement("OPTION");
		TagOPTION.value	= "";
		TagOPTION.text	= "Sanctioned / Unsanctioned";
		frmScreen.cboPaymentStatus.add(TagOPTION);
		//Default to Sanctioned / Unsanctioned;
		frmScreen.cboPaymentStatus.value="";
		
		
		//ADD new OPTION 'All Types' to PaymentMethod Combo		
		TagOPTION	= document.createElement("OPTION");
		TagOPTION.value	= "";
		TagOPTION.text	= "All Types";
		frmScreen.cboPaymentMethod.add(TagOPTION);
		//Default to 'All Types';
		frmScreen.cboPaymentMethod.value="";	
	}
	
	
	
	return(bSuccess);
}

function SendManualPaymentData(nAmount,nApplicant,SelectedRows)
{
	<% /* now set the extracted data to the array  */  %>
	var ArrayArguments = new Array();
	ArrayArguments[0] = nAmount;		//Value of the selected row
	ArrayArguments[1] = nApplicant;		//Application number of the selected row
	ArrayArguments[2] = SelectedRows.XMLDocument.xml;//XML of row selected
	var sReturn = scScreenFunctions.DisplayPopup(window, document, "SP094.asp", ArrayArguments, 378, 248);
	return (sReturn[0] == true);  // has the user pressed OK?
}

function SendData(nAmount,m_nRows,SelectedRows, bIsCheque) 
{
	<% /* now set the extracted data to the array  */ %>
	var ArrayArguments = new Array();
	ArrayArguments[0] = nAmount;		//Value of rows selected
	ArrayArguments[1] = m_nRows;		//Number of rows selected
	ArrayArguments[2] = SelectedRows.XMLDocument.xml;//Array of rows selected
	ArrayArguments[3] = scScreenFunctions.GetContextParameter(window,"idUserId");
	ArrayArguments[4] = scScreenFunctions.GetContextParameter(window,"idUnitId");
	
	var Return = null;
	var bConfirmed = false;
	var bReprint = false;
	
	//	AW	05/08/02	BMID029
	if(bIsCheque == true)
	{
	
		var XMLSelRows = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XMLSelRows.LoadXML(SelectedRows.XMLDocument.xml);
				
		do
		{
			ArrayArguments[2] = XMLSelRows.xml //SelectedRows.XMLDocument.xml;
			<%//Return = scScreenFunctions.DisplayPopup(window, document, "SP091.asp", ArrayArguments, 340, 250);%>
			<%//GD BMIDS00572 START %>
			Return = scScreenFunctions.DisplayPopup(window, document, "SP091.asp", ArrayArguments, 393, 285);
			<%//GD BMIDS00572 END %>			
			if (Return[0] == true)
			{
				var XMLPrintCheques = null;
				var XMLPrintCheques = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
				var XMLRequestTag = XMLPrintCheques.CreateRequestTag(window , "SanctionPrintCheques");
				XMLPrintCheques.ActiveTag = XMLRequestTag;
				XMLPrintCheques.SetAttribute("OPERATION", "SanctionPrintCheques");
				XMLPrintCheques.SetAttribute("ACTION", "SanctionPrintCheques");
				XMLPrintCheques.SetAttribute("CHEQUENUMBER", Return[1]);
				
				if (bReprint == true) XMLPrintCheques.SetAttribute("REPRINTMODE", "1");
				else XMLPrintCheques.SetAttribute("REPRINTMODE", "0");
	
				UpdateChequeNumber( Return[1]);
				
				//Append payment list								
				var PaymentList = XMLSelRows.CreateTagList("PAYMENTRECORD");		
				var iCount = 0;
				
				if (PaymentList != null)
				{
					for (iCount=0; iCount < PaymentList.length; ++iCount)
					{
						XMLPrintCheques.ActiveTag.appendChild(PaymentList(iCount));
					}
				}
	
				
				XMLPrintCheques.RunASP(document,"SanctionPrintCheques.asp");
				
				if (XMLPrintCheques.IsResponseOK() == true) 	
				{
					var ArraySP095Arguments = new Array();
					ArraySP095Arguments[0] = XMLPrintCheques.CreateRequestAttributeArray(window);
					ArraySP095Arguments[1] = XMLPrintCheques.XMLDocument.xml;//Array of rows selected
					
					ArraySP095Arguments[2] = scScreenFunctions.GetContextParameter(window,"idUserId");
					ArraySP095Arguments[3] = scScreenFunctions.GetContextParameter(window,"idUnitId");
					
					Return = scScreenFunctions.DisplayPopup(window, document, "SP095.asp", ArraySP095Arguments, 650, 500);
					if (Return[0] == true)
					{	
						var XMLPaidCheques = new top.frames[1].document.all.scXMLFunctions.XMLObject();
						var XMLPaidRequestTag = XMLPaidCheques.CreateRequestTag(window , "SetPaidChequePayments");
						
						PaymentList = XMLPrintCheques.CreateTagList("PAYMENTRECORD");		
						iCount = 0;
				
						if (PaymentList != null)
						{
							for (iCount=0; iCount < PaymentList.length; ++iCount)
							{
								if(PaymentList(iCount).getAttribute("CHEQUENUMBER") != "Locked") {
									XMLPaidCheques.ActiveTag.appendChild(PaymentList(iCount));
								}
							}
						}
						
						//Update payment status to PAID 
						var xmlNodeList = XMLPaidCheques.XMLDocument.getElementsByTagName("DISBURSEMENTPAYMENT");
						for (var i=0; i < xmlNodeList.length ; i++ ) 
						{
							xmlNodeList.item(i).setAttribute("PAYMENTSTATUS", m_sPaidId);
						}
						
						XMLPaidCheques.RunASP(document,"SetPaidChequePayments.asp");
						
						if (XMLPaidCheques.IsResponseOK() == true) 	
						{
							bConfirmed = true
							Search();
						}		
					}
					else if(Return[0] == false)
					{
						//Reprint
						XMLSelRows.LoadXML(Return[1]);
						bReprint = true;
						
						PaymentList = XMLSelRows.CreateTagList("PAYMENTRECORD");
								
						var nAmount = 0; var nRecords = 0;
				
						if (PaymentList != null)
						{
							for (iCount=0; iCount < PaymentList.length; ++iCount)
							{
								if(PaymentList(iCount).getAttribute("REPRINTSTATUS") == 1)
								{
									//nAmount = PaymentList(iCount).getAttribute("AMOUNT");
									//GD BMIDS01045
									//MV - 26/03/2003 - BM0088 
									nAmount = nAmount + ParseFloatSafe(PaymentList(iCount).getAttribute("NETPAYMENTAMOUNT"));
									nRecords++;
								}
							}
						}
						
						ArrayArguments[0] = nAmount;		//Value of the selected rows
						ArrayArguments[1] = nRecords;				
					}
				}
			}
			else //cancel from SP091
			{
				//Must not exit process by cancelling from SP091 in reprint situation
				if(bReprint == true)
				{
					alert("Cancelling from this screen is not allowed if reprinting");
				}
				else
				{
					return;
				}
			}				
		}
		while(bConfirmed == false)
	}
	else
	{
		Return = scScreenFunctions.DisplayPopup(window, document, "SP093.asp", ArrayArguments, 395, 220);		
	
		if (Return[0] == true){
			Search();
		}
	}
}

function UpdateChequeNumber(nCurrentChequeNum)
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateActiveTag("REQUEST");
	XML.CreateActiveTag("UNIT");
	XML.CreateTag("UNITID", scScreenFunctions.GetContextParameter(window,"idUnitId"));
	XML.CreateTag("UNHIGHCHEQUENUMBER", parseInt(nCurrentChequeNum) +  parseInt(m_nRows) - 1);
				
	XML.RunASP(document, "UpdateUnitHighChequeNum.asp");
}

function CheckUserRole()
{
	<% /*Get the Context UserRole */ %>
	var varUserRole = scScreenFunctions.GetContextParameter(window,"idRole");
	<%/* Compare this with the Global parameter PProcSanctionRole */ %>
	if (varUserRole < m_PProcSanctionRole)
	{
		alert("You do not have authority to use this function");
		<%//Current_ArrayLength = m_Old_ArrayLength;%>
		return false;
	}
	return true;
}
<%/*
function CheckUserID(sApplApprovalUserID, sApplRecommendedUserId) 
{
	//Get the Context UserID
	var varContextUserID = scScreenFunctions.GetContextParameter(window,"idUserName");
	//compare the two user id's 
	if(varContextUserID == UserID)
	{
		alert("You cannot sanction/unsanction a case you already own");
		Current_ArrayLength = m_Old_ArrayLength;
		return false;
	}
	return true;
} */
%>
function CheckPaymentStatus(Status,atStatus) // feed in status and status attribute 
{
	if(Status == atStatus)
	{
		alert("You have selected rows that are already Sanctioned/Unsanctioned");
		return false;
	}
	return true;
}

function btnSubmit.onclick()
{
	frmToMN010.submit();
}
//GD BMIDS00979	
function IsCheque(sPaymentMethod)
{
	var blnResult = false;
	var iIndex;
	var sPattern = ".//LISTENTRY[VALUEID = '" + sPaymentMethod + "']/VALIDATIONTYPELIST"
	var xmlPaymentList = XMLCombos.selectNodes(sPattern)
	var xmlPayment;
	for(iIndex=0 ; iIndex < xmlPaymentList.length ; iIndex++)
	{
		xmlPayment = xmlPaymentList.item(iIndex)
		if (xmlPayment.text == 'CH')
		{
			return(true)
		}	
	}

	return(false);

}


function GetApplicationLockData(sApplicationNumber)
{

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("APPLICATIONLOCK");
	XML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	XML.RunASP(document,"GetApplicationLockData.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);

	if(ErrorReturn[1] == ErrorTypes[0])
		return false ;
	else
		return true;
		
	
}
//GD START
function GetArrayOfRowsSelected()
{
	var aRetVal = new Array();
	var iRetValIndex = 0;
	var iIndex;

	for(iIndex = 0;iIndex < m_aCheckBoxArray.length;iIndex++)
	{
		if(m_aCheckBoxArray[iIndex] =="1")
		{
			aRetVal[iRetValIndex] = iIndex;
			iRetValIndex++;
		}
	}
	return(aRetVal);
}
<%//GD BM0477 START %>
function frmScreen.btnDeSelectAll.onclick()
{
	if(m_aCheckBoxArray.length > 0)
	{
		SelectDeSelectAll("");
		//Show list 'as is', but with all deselected. Don't show list from record 1, by default
		Showlist(frmScreen.scTableList.getFirstVisibleRecord() - 1);
	}
}
function frmScreen.btnSelectAll.onclick()
{
	if(!m_bSancUnsancMixFound)
	{
		SelectDeSelectAll("1");
		//Show list 'as is', but with all selected. Don't show list from record 1, by default
		Showlist(frmScreen.scTableList.getFirstVisibleRecord() - 1);
		if(m_aCheckBoxArray.length > 0)
		{
			frmScreen.btnSanction.disabled = false;
			frmScreen.btnUnSanction.disabled = false;
			frmScreen.btnManualPayment.disabled = false;
		}
	} else
	{
		alert("You cannot 'Select All' on a list which has a mixture of Sanctioned and Unsanctioned payments");
	}
}

<%//GD BM0477 END %>
<%//GD BM0477 START %>
<%/*
These 2 functions alter the style on the cell 0 (oCell)
We need to know which cell is having its style altered (iWhichCell) because
the very top and very bottom cell have slightly different styles which include more borders
*/%>

function SetChecked(oCell,iWhichCell)
{
	//oCell.className = "msgYes";
	if (iWhichCell==1) //First Row
	{
		oCell.className = "msgYesTop";
	} else
	{
		if (iWhichCell == m_iTableLength) //Last Row
		{
			oCell.className = "msgYesBottom";
		
		} else
		{
			oCell.className = "msgYes";		
		}
	}
}
function SetUnChecked(oCell,iWhichCell)
{
	if (iWhichCell==1) //First Row
	{
		oCell.className = "msgNoTop";
	} else
	{
		if (iWhichCell == m_iTableLength) //Last Row
		{
			oCell.className = "msgNoBottom";
		
		} else
		{
			oCell.className = "msgNo";		
		}
	}
}
<%//GD BM0477
//This resets the style on the rows that aren't populated, to their original style, as per HTML definition
%>
function ResetBlankRows(iCount)
{
	for(iIndex = (iCount + 1); iIndex <= m_iTableLength; iIndex++)
	{
		if (iIndex == 1)
		{
			tblPaymentSanction.rows(iIndex).cells(0).className = "TableTopLeft";
		} else
		{
			if (iIndex == m_iTableLength)
			{
				tblPaymentSanction.rows(iIndex).cells(0).className = "TableBottomLeft";
			} else
			{
				tblPaymentSanction.rows(iIndex).cells(0).className = "TableLeft";
			}
		}
	}
}
<%/*
This function clears the contents of the HTML table
It sets the innerText of column 0,of any rows that have data, to be 2 spaces,
instead of the 1 space.
This is because the Table manager scriptlet will not allow selection of a row if the data in cell 0 is a single space.
*/%>
function clear(iNumRows)
{
	if( tblPaymentSanction != null )
	{	
		for(var i = 0; i < tblPaymentSanction.rows.length; i++)
		{
			if(tblPaymentSanction.rows(i).id != "rowTitles")
			{
				
				tblPaymentSanction.rows(i).style.background = "FFFFFF";		
				tblPaymentSanction.rows(i).cells(0).title = "";
				tblPaymentSanction.rows(i).cells(0).style.color = "616161";
				if (i <= iNumRows)
				{
					tblPaymentSanction.rows(i).cells(0).innerText = "  ";
				} else
				{
					tblPaymentSanction.rows(i).cells(0).innerText = " ";			
				}
				
				for(var j = 1; j < tblPaymentSanction.rows(i).cells.length; j++)
				{
					tblPaymentSanction.rows(i).cells(j).innerText = " ";
					tblPaymentSanction.rows(i).cells(j).title = "";
					tblPaymentSanction.rows(i).cells(j).style.color = "616161";
				}
			}
		}
		//m_indexRowSelected = null;
		frmScreen.scTableList.setRowSelectedIndex(null);
		
		//m_aidActualRowsSelected = new Array();
	}
}

<%//GD BM0477 END %>
-->
</script>
</body>
</html>
