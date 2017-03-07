<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      TM030.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Application history
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		24/10/00	Created (screen paint)
JLD		10/11/00	Some implementation added
JLD		22/11/00	BO calls added
JLD		24/11/00	route back to pp215
JLD		15/12/00	SYS1709 Route back to MN060 instead of PP215
					Added fnc to check for no more stages when moving to next stage.
JLD		03/01/01	SYS1761 check for NOSTAGEAUTHORITY error when moving stage.
					Check for NOTASKAUTHORITY when adding a task.
JLD		03/01/01	SYS1778 Use time in date comparisons.
JLD		09/01/01	SYS1789 use error RECORDNOTFOUND instead of NOMORESTAGES
JLD		12/01/01	SYS1810 Don't allow edit for tasks which have been 'finished'.
JLD		12/01/01	SYS1807 Depending on current stage, set the 'move to another stage' functionality disabled.
					Also stopped edit etc buttons enabling if click on an empty row in the table.
JLD		15/01/01	SYS1808 update correct context params for stage change.
ADP		15/02/01	SYS1952 amended OK routing to go to MN070 rather than MN060
JLD		20/02/01	SYS1832 use CASETASKNAME in preference to TASKNAME in list box.
APS		02/03/01	SYS1920 If the re-issues offer task has been added then unfreeze data
DRC     07/03/01    SYS1787 Changes for invoking Task History instead of just 'Status' History 
DRC     08/03/01    SYS1786 Added collection of reason for exception for confirm in new stage
CL		12/03/01	SYS1920 Read only functionality added
SR		13/03/01	SYS2034 - Data freezing
APS		15/01/01	SYS2076
SR		15/03/01	SYS1840
CL		15/03/01    SYS2037 Change to pass values to NON Popup TM037
ADP/GD	20/03/01	SYS2078 Clear idXml onload.
IK		21/03/01	SYS1924 use idTaskXML not idXML (idXML gets over-written in task processing screens)
CL		03/04/01	SYS2213 Change screen back so that it calls TM037 as a full screen.
BG		17/04/01	SYS2096 Added additional attributes to call to MovetoStage and MovetoNextStage calls
DRC     12/07/01    SYS2211 Changed order by to "Date" on entry and as default
DRC     16/07/01    SYS2088 Added function to check for all tasks complete on entry QueryStageComplete & MoveToNextStage
JLD		08/10/01	SYS2736 Don't download omPC again
SG		06/12/01	SYS3357 Make screen read only if application is Cancelled or Declined.
SG		11/12/01	SYS3387 Display DueDate/Completed date depending on task status.
							Amended ShowList, Added GetComboValidationType
DRC     02/05/02    SYS4530 Check whether a new AdminSys account number was added in an Automatic task							
DRC     08/05/02    SYS4533 - must not be able to process tasks if in Read Only mode
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		06/06/02	SYS4822 Error caused by SYS4727
TW      09/10/2002  SYS5115 Modified to incorporate client validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
PSC     29/06/2005  MAR5		Add Delivery Type to Template attributes.
GHun	19/07/2005	MAR7		Integerate local printing
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
MV		12/09/2005	MAR35		Amended ActionProcessing();makeInterfaceCall()
PSC		22/09/2005	MAR32		WP08 Add funtionality for TAS tasks
SAB		27/10/2005	MAR245		Updated to allow revised address targeting
PSC		27/10/2005	MAR300		Update makeInterfaceCall to pass customer data down
PSC		27/10/2005	MAR283		Update makeInterfaceCall so it doesn't try to complete RunAutoApplicationExpiry task 
TW		01/11/2005	MAR211		Add Pack handling in ActionProcessing()
MV		03/11/2005	MAR402		Amended makeInterfaceCall()
TW		16/11/2005	MAR578		Add printer destination type to request for omPrint. Also update Case Task on Pack
INR		30/11/2005	MAR725		No document, don't call updatecasetask
TW 		01/12/2005 	MAR764		Fix for case task printing
GHun	05/12/2005	MAR796		Turn off document compression
HMA     07/12/2005  MAR750      Set TASKMANAGER attribute for Risk Assessment
HMA     12/12/2005  MAR667      Add extra parameter to call to TM031
PSC		18/01/2006	MAR1081		Cater for Product Switch End and Transfer of Equity End stages
JD		26/01/2006	MAR1040		For tasks with Packnumbers and screens, do the pack first, then set to pending to go to the screen next.
PSC		26/01/2006	MAR1133		Send OTHERSYSTEMCUSTOMERNUMBER in MoveToNextStage
PJO     09/03/2006  MAR1359		Add progress message while interfacing
PJO     09/03/2006  MAR1359     Parameterise progress message width
GHun	10/03/2006	MAR1300		Call ValidateProcessTaskAuthority when clicking action
GHun	20/03/2006	MAR1453		Hide progress window before popping up an error
GHun	05/04/2006	MAR1300		Changes for ChangeOfProperty
GHun	21/04/2006	MAR1658		Further changes for ChangeOfProperty
PSC		19/05/2006	MAR1627     Only show TOE end stage and PSW end stage in next stage combo if relevant to application type
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History :

Prog	Date		AQR			Description
PE		16/06/2006	EP710		Modified CreateTemplateDataBlock to ignore null nodes.
SAB		18/07/2006	EP989		Now checks the CancelDeclineFreezeDataIndicator flag before
								configuring the options in the print menu
AW		14/09/2006	EP1103		CC78 New mode parameter for calling TM031
AW		14/09/2006	EP1150      CC78 New arguments passed to TM036PTM036
AW		14/09/2006	EP1159      CC78 Increased height co-ordinate of TM036P
HMA     02/01/2007  EP2_576     E2CR67 Pass Customer Role to TM031.
SR		05/03/2007  EP2_1596	TMTOEEndStageID and TMPSWStageID are set to 90 (Completion_Exception). Do not exclude
								them explicitly.
GHun	09/03/2007	EP2_2026	Changed ProcessPrint to prevent printing twice in some cases
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 295px; LEFT: 314px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Forms Here */ %>
<% /* BM0262 no longer used
<form id="frmToTM032" method="post" action="TM032.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM033" method="post" action="TM033.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM036" method="post" action="TM036.asp" STYLE="DISPLAY: none"></form>
*/ %>
<form id="frmToTM037" method="post" action="TM037.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN070" method="post" action="MN070.asp" STYLE="DISPLAY: none"></form>
<form id="frmInputProcess" method="post" action="" STYLE="DISPLAY: none"></form>
<% /* BMIDS682 */ %>
<form id="frmToRA040" method="post" action="ra040.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>

<div id="divReorder" style="TOP: 60px; LEFT: 10px; HEIGHT: 38px; WIDTH: 608px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:11px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Current Stage
	<span style="TOP:-3px; LEFT:75px; POSITION:ABSOLUTE">
		<input id="txtCurrentStage" maxlength="50" style="WIDTH:200px" class="msgTxt">
	</span>
</span>
<span style="TOP:7px; LEFT:283px; POSITION:ABSOLUTE">
	<input id="btnStageHistory" value="Stage History" type="button" style="WIDTH:80px" class="msgButton">
</span>
<span style="TOP:10px; LEFT:381px; POSITION:ABSOLUTE" class="msgLabel">
	Sort Tasks by
	<span style="TOP:-3px; LEFT:74px; POSITION:ABSOLUTE">
		<select id="cboOrderBy" style="WIDTH:150px" class="msgCombo"></select>
	</span>
</span>
</div>

<div id="divTaskList" style="TOP: 104px; LEFT: 10px; HEIGHT: 220px; WIDTH: 608px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:0px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
	Task List
</span>
<div id="spnTable" style="TOP: 17px; LEFT: 4px; POSITION: ABSOLUTE">
	<table id="tblTable" width="600" border="0" cellspacing="0" cellpadding="0" class="msgTable">
	<tr id="rowTitles">
		<td width="35%" class="TableHead">Task</td>
		<td width="5%" class="TableHead">TAS</td>
		<td width="5%" class="TableHead">!</td>
		<td width="17%" class="TableHead">Status</td>
		<td width="10%" class="TableHead">Mandatory</td>
		<td width="23%" class="TableHead">Due/Completed<BR>Date</td>
		<td width="5%" class="TableHead">Owner</td>
	</tr>
	<tr id="row01">		
		<td class="TableTopLeft">&nbsp;</td>		
		<td class="TableTopCenter" align="center">&nbsp;</td>
		<td class="TableTopCenter" align="right">&nbsp;</td>		
		<td class="TableTopCenter">&nbsp;</td>
		<td class="TableTopCenter">&nbsp;</td>
		<td class="TableTopCenter">&nbsp;</td>		
		<td class="TableTopRight" >&nbsp;</td>
	</tr>
	<tr id="row02">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter" align="center">&nbsp;</td>
		<td class="TableCenter" align="right">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row03">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter" align="center">&nbsp;</td>
		<td class="TableCenter" align="right">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row04">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter" align="center">&nbsp;</td>
		<td class="TableCenter" align="right">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row05">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter" align="center">&nbsp;</td>
		<td class="TableCenter" align="right">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row06">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter" align="center">&nbsp;</td>
		<td class="TableCenter" align="right">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row07">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter" align="center">&nbsp;</td>
		<td class="TableCenter" align="right">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row08">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter" align="center">&nbsp;</td>
		<td class="TableCenter" align="right">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row09">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter" align="center">&nbsp;</td>
		<td class="TableCenter" align="right">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row10">		
		<td class="TableBottomLeft">&nbsp;</td>	
		<td class="TableBottomCenter" align="center">&nbsp;</td>
		<td class="TableBottomCenter" align="right">&nbsp;</td>	
		<td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomRight">&nbsp;</td>
	</tr>
	</table>
</div>

	<span style="TOP:192px; LEFT:4px; POSITION:ABSOLUTE">
		<input id="btnAdd" value="Add Task" type="button" style="WIDTH:80px" class="msgButton">
	</span>

	<% /* BM0493 MDC 03/04/2003 */ %>
	<span style="TOP:192px; LEFT:88px; POSITION:ABSOLUTE">
		<input id="btnEdit" value="Edit Task" type="button" style="WIDTH:80px" class="msgButton">
	</span>

	<span style="TOP:192px; LEFT:172px; POSITION:ABSOLUTE">
		<input id="btnTaskHistory" value="Task History" type="button" style="WIDTH:80px" class="msgButton">
	</span>

	<% /* BM0493 MDC 03/04/2003 - End */ %>

</div>
<div id="divProcessTask" style="TOP: 330px; LEFT: 10px; HEIGHT: 90px; WIDTH: 608px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP:0px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
		Process Task
	</span>
	<div class="msgInput" style="top: 15px; LEFT: 4px; POSITION: ABSOLUTE;">
		<table id="tblTaskDetail" width="594" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr>
				<td width="80"><b>Task:</b></td><td width="168" id="tblTDName"></td>
				<td width="120"><b>Status:</b></td><td id="tblTDStatus"></td>
			</tr>
			<tr>
				<td><b>Due Date:</b></td><td id="tblTDDueDate"></td>
				<td><b>Last Actioned Date:</b></td><td id="tblTDActionedDate"></td>
			</tr>
			<tr>
				<td><b>Owner:</b></td><td id="tblTDOwner"></td>
				<td><b>Last Updated By:</b></td><td id="tblTDUpdatedBy"></td>
			</tr>
		</table>
	</div>

	<span style="TOP:63px; LEFT:4px; POSITION:ABSOLUTE">
		<input id="btnAction" value="Action" type="button" style="WIDTH:80px" class="msgButton" disabled>
	</span>
	<span style="TOP:63px; LEFT:88px; POSITION:ABSOLUTE">
		<input id="btnDetails" value="Details" type="button" style="WIDTH:80px" class="msgButton" disabled>
	</span>
	<span style="TOP:63px; LEFT:172px; POSITION:ABSOLUTE">
		<input id="btnReprint" value="Reprint" type="button" style="WIDTH:80px" class="msgButton" disabled>
	</span>
	<span style="TOP:63px; LEFT:256px; POSITION:ABSOLUTE">
		<input id="btnChaseUp" value="Chase Up" type="button" style="WIDTH:80px" class="msgButton" disabled>
	</span>
	<span style="TOP:63px; LEFT:344px; POSITION:ABSOLUTE">
		<input id="btnNotApplicable" value="Not Applicable" type="button" style="WIDTH:80px" class="msgButton" disabled>
	</span>
	<span style="TOP:63px; LEFT:432px; POSITION:ABSOLUTE">
		<input id="btnMemo" value="Memo" type="button" style="WIDTH:80px" class="msgButton" disabled>
	</span>
	<span style="TOP:63px; LEFT:520px; POSITION:ABSOLUTE">
		<input id="btnContactDetails" value="Contact Details" type="button" style="WIDTH:80px" class="msgButton" disabled>
	</span>

	<!--
	<span style="TOP:272px; LEFT:4px; POSITION:ABSOLUTE">
		<input id="btnEdit" value="Edit Task" type="button" style="WIDTH:80px" class="msgButton">
	</span>

	<span style="TOP:228px; LEFT:172px; POSITION:ABSOLUTE">
		<input id="btnProcess" value="Process Task" type="button" style="WIDTH:80px" class="msgButton">
	</span>

	<span style="TOP:272px; LEFT:88px; POSITION:ABSOLUTE">
		<input id="btnTaskHistory" value="Task History" type="button" style="WIDTH:80px" class="msgButton">
	</span>
	-->
</div>

<div id="divStageStatus" style="TOP: 426px; LEFT: 10px; HEIGHT: 75px; WIDTH: 608px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:2px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
	Stage Status
</span>
<% /*
<span style="TOP:8px; LEFT:524px; POSITION:ABSOLUTE">
	<input id="btnStageHistory" value="Stage History" type="button" style="WIDTH:80px" class="msgButton">
</span> 
*/ %>
<span style="TOP:25px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Total Tasks
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtTotalTasks" maxlength="10" style="WIDTH:60px" class="msgTxt">
	</span>
</span>
<span style="TOP:25px; LEFT:170px; POSITION:ABSOLUTE" class="msgLabel">
	Total Incomplete
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtTotalIncomplete" maxlength="10" style="WIDTH:60px" class="msgTxt">
	</span>
</span>
<span style="TOP:25px; LEFT:336px; POSITION:ABSOLUTE" class="msgLabel">
	Incomplete Mandatory
	<span style="TOP:-3px; LEFT:110px; POSITION:ABSOLUTE">
		<input id="txtIncompMandatory" maxlength="10" style="WIDTH:60px" class="msgTxt">
	</span>
</span>
<span style="TOP:51px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Move to New Stage
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<select id="cboNewStage" style="WIDTH:230px" class="msgCombo" menusafe="true"></select>
	</span>
</span>
<span style="TOP:46px; LEFT:520px; POSITION:ABSOLUTE">
	<input id="btnConfirm" value="Confirm" type="button" style="WIDTH:80px" class="msgButton">
</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 503px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/TM030Attribs.asp" -->
<% /* Added BMIDS692 DRC  */ %>
<!-- #include FILE="attribs/OverrideAttribs.asp" -->

<script src="includes/Documents.js" language="javascript" type="text/javascript"></script> <% /* MAR7 GHun */ %>

<% /* Specify Code Here */ %>
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = "";
var m_sApplicationNumber = "";
var m_sApplicationFFNumber = "";
var m_sApplicationPriority = "" ;
var m_sActivityId = "";
var m_sReadOnly = "";
var m_iUserRole = 0;
var m_sCurrentStageNum = "0";
var m_iTableLength = 10;
var m_sCurrStageId = "";
var scScreenFunctions;
var stageXML = null;
var TaskStatusXML = null;
var m_sTMReIssueOffer = "";
var m_blnReadOnly = false;
var m_sLastStageId = "";
var m_sPreCompStageId="";
var m_sCancelStageId = "";
var m_sDeclineStageId = "";
<% /* BM0493 MDC 02/04/2003 */ %>
var m_sCompStageId = "";
var m_ParamXML = null;
<% /* BM0493 MDC 02/04/2003 - End */ %>
var m_iActionCounter = 0; <% /* BM0262 */ %>
<% /* BS BM0271 16/04/03 */ %>
var m_sReadOnlyAsCaseLocked = "";
var m_sProcessInd = "";

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;

// BM0493 ik 20030328
var m_taskXML = null;
// BM0493 ik 20030328 ends

<% /* MAR7 GHun */ %>
var m_sDeliveryType = "";
var m_sCompressionMethod = "";
var sDocumentID = "";
var sPrintType  = "";
var AttribsXML = null;
var sInterfaceName = "";
var node = null;
var sPrinterDestTypeID = "";
var m_UserId = "";
var m_UnitId = ""; 
var m_MachineId  = "";
var m_DistributionChannelId = "";
var xmlControlDataNode = null;
var xmlPrintDataNode = null;
var xmlTemplateDataNode = null;
var sPrinterValidationType ="";
var bAutomaticTask = false;
var sTaskName = "";
var sInputProcess = "";
<% /* MAR7 End */ %>

<% /* MAR578 TW 16/11/2005 */ %>
var sPrinterType = "";
var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
<% /* MAR578 TW 16/11/2005 End */ %>


// BMIDS682
var m_AddressTargetXML = null;

<% /* PSC 18/01/2006 MAR1081 - Start */ %>
var m_sTOEEndStageId = "";
var m_sPSWStageId = "";
<% /* PSC 18/01/2006 MAR1081 - End */ %>

var m_sTmCOPRemodelTaskId = ""; <% /* MAR1658 GHun */ %>

function window.onload()
{
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Application Stage","TM030",scScreenFunctions);

<%  // SYS2078 Clear out idTaskXML in context %>
	scScreenFunctions.SetContextParameter(window,"idTaskXML", "");

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
<%	//This function is contained in the field attributes file (remove if not required)
%>	//SetMasks();
	Validation_Init();
	PopulateScreen();
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "TM030");
	<% /* BM0670  START  This is set to catch input processes*/ %>
	CheckFreezeUnFreeze(m_sApplicationNumber);
	<% /* BM0670  END */ %>
	<% /* BM0567  START */ %>
	CheckDataFreezeFlag();
	<% /* BM0567  END */ %>
	if(m_blnReadOnly)
	{
		m_sReadOnly = "1" ;
		<% /* BS BM0271 16/04/03 Re-enable table so processed task details can be viewed
		//SYS4533 - must not be able to process tasks if in Read Only mode
		scScrollTable.DisableTable(); */ %>
		frmScreen.btnAdd.disabled = true; //JR BM0271
	}
	else
	{
		QueryStageComplete();
	}
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();

//	BM0493 ik 20030328
	m_taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
//	BM0493 ik 20030328 ends
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","E00005525");
	m_sApplicationFFNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	m_sActivityId = scScreenFunctions.GetContextParameter(window,"idActivityId","1");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	<% /* BS BM0271 16/04/03 */ %>
	m_sReadOnlyAsCaseLocked = m_sReadOnly;
	m_sProcessInd = scScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null); 
	<% /* BS BM0271 End 16/04/03 */ %>
	m_sApplicationPriority = scScreenFunctions.GetContextParameter(window,"idApplicationPriority","0");
	m_iUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	<% /* MAR7 GHun */ %>
	m_UserId = scScreenFunctions.GetContextParameter(window,"idUserId",null);
	m_UnitId = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	m_DistributionChannelId = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
	m_MachineId = scScreenFunctions.GetContextParameter(window,"idMachineId",null);
	<% /* MAR7 End */ %>
	
	GetAllGlobalParameters()
}

function PopulateScreen()
{
	frmScreen.btnEdit.disabled = true;
	frmScreen.btnTaskHistory.disabled = true;
	scScreenFunctions.SetFieldState(frmScreen, "txtCurrentStage", "R");
	frmScreen.btnConfirm.disabled = true;
	scScreenFunctions.SetFieldState(frmScreen, "txtTotalTasks", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtTotalIncomplete", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtIncompMandatory", "R");

	PopulateCombos();
	sActivitySequenceNo = GetStageInfo();
	PopulateNewStageCombo(sActivitySequenceNo);
// SYS2211 DRC	 - Order by Date when populating table
	frmScreen.cboOrderBy.onchange()
// SYS2211 DRC	
	GetReIssueOfferParameter();
}

function GetStageInfo()
{
	stageXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	stageXML.CreateRequestTag(window , "GetCurrentStage");
	stageXML.CreateActiveTag("CASEACTIVITY");
	stageXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	stageXML.SetAttribute("CASEID", m_sApplicationNumber);
	stageXML.SetAttribute("ACTIVITYID", m_sActivityId);
	stageXML.SetAttribute("ACTIVITYINSTANCE", "1");

	// 	stageXML.RunASP(document, "MsgTMBO.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	<%//GD BM0574 START
	//switch (ScreenRules())
		//{
		//case 1: // Warning
		//case 0: // OK %>
	stageXML.RunASP(document, "MsgTMBO.asp");
			<%//break;
		//default: // Error
			//stageXML.SetErrorResponse();
		//} GD BM0574 END %>
	

	var sActivitySequenceNo = "";
	if(stageXML.IsResponseOK())
	{
		stageXML.SelectTag(null, "CASESTAGE");
		m_sCurrStageId = stageXML.GetAttribute("STAGEID");
		var saOrigStageInfo = FW030GetStageInfo(m_sCurrStageId);
		var sCompletionDate = stageXML.GetAttribute("STAGECOMPLETIONDATETIME");
		//var sArray = new Array(m_sCurrStageId, null);
		scScreenFunctions.SetContextParameter(window,"idStageId", m_sCurrStageId);
		scScreenFunctions.SetContextParameter(window,"idStageName", saOrigStageInfo[0]);
		scScreenFunctions.SetContextParameter(window,"idCurrentStageOrigSeqNo", saOrigStageInfo[1]);
		// Having reset the stage, update the screen
		FW030SetTitles("Application Stage","TM030",scScreenFunctions);
		<% /* BM0567  START */ %>
		CheckDataFreezeFlag();
		<% /* BM0567  END */ %>
		frmScreen.txtCurrentStage.value = stageXML.GetAttribute("STAGENAME");
		sActivitySequenceNo = stageXML.GetAttribute("ACTIVITYSEQUENCENO");
		frmScreen.txtTotalTasks.value = stageXML.GetAttribute("TOTALTASKS");
		frmScreen.txtTotalIncomplete.value = stageXML.GetAttribute("INCOMPLETETASKS");
		frmScreen.txtIncompMandatory.value = stageXML.GetAttribute("INCOMPLETEMANDATORYTASKS");
		//Add an ORDER tag to each CASETASK for populating the table
		//Add a SYMBOL tag also to store a symbol representing the tasks lateness.
		stageXML.CreateTagList("CASETASK");
		var iNoOfTasks = stageXML.ActiveTagList.length;
		
		<% /* PSC 22/09/2005 MAR32 - Start */ %>
		var blnTASTaskInProgress = false;
		var strTaskStatus = "";
		var comboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		<% /* PSC 22/09/2005 MAR32 - End */ %>

		if (iNoOfTasks > 0)
		{
			var iLoop;

			for(iLoop = 0; iLoop < iNoOfTasks; iLoop++)
			{
				stageXML.SelectTagListItem(iLoop);
				var sSymbol = GetSymbolForDate(stageXML.GetAttribute("TASKSTATUS"), stageXML.GetAttribute("TASKDUEDATEANDTIME"));
				stageXML.CreateTag("ORDER", iLoop+1);
				stageXML.CreateTag("SYMBOL", sSymbol);
				
				<% /* PSC 22/09/2005 MAR32 - Start */ %>
				strTaskStatus = stageXML.GetAttribute("TASKSTATUS");
				if(comboXML.IsInComboValidationList(document,"TaskStatus", strTaskStatus, ["TR"]))
					blnTASTaskInProgress = true
				<% /* PSC 22/09/2005 MAR32 - End */ %>
			}			
		}
		//Depending on the stage, allow facility to move to a new stage
		<% /* PSC 22/09/2005 MAR32 */ %>
		<% /* PSC 18/01/2006 MAR1081 */ %>
		if( (blnTASTaskInProgress || m_sCurrStageId == m_sCancelStageId || m_sCurrStageId == m_sDeclineStageId ||
		     m_sCurrStageId == m_sTOEEndStageId || m_sCurrStageId == m_sPSWStageId) ||
		    (sCompletionDate != null && sCompletionDate != "") )
		{
			frmScreen.btnConfirm.disabled = true;
			frmScreen.cboNewStage.disabled = true;
		}
	}
	return(sActivitySequenceNo);
}

function GetSymbolForDate(sTaskStatus, sTaskDueDate)
{
<%	//take the time off the date/time due date to find out if it is required today
%>	var sSymbol = "";
	if(sTaskDueDate != null)
	{
		var sDate = sTaskDueDate.substr(0,10);
		var sDateTime = sTaskDueDate;
		if(GetComboText(sTaskStatus) != "Complete" && GetComboText(sTaskStatus) != "Not Applicable") // THESE MIGHT ALTER
		{
			if(scScreenFunctions.CompareDateStringToSystemDateTime(sDateTime,"<"))
				sSymbol = "!!";
			<% /* MO - BMIDS00807 */ %>
			<% /* else if(sDate == scScreenFunctions.DateToString(scScreenFunctions.GetSystemDate())) */ %>
			else if(sDate == scScreenFunctions.DateToString(scScreenFunctions.GetAppServerDate()))
				sSymbol = "*";
		}
	}
	return sSymbol;
}

function ResetCombos()
{
	var nLength = frmScreen.cboOrderBy.options.length;
	for(var iCount = nLength -1; iCount >= 0 ; iCount--)
	{
		frmScreen.cboOrderBy.options.remove(iCount);
	}
	nLength = frmScreen.cboNewStage.options.length;
	for(iCount = nLength -1; iCount >= 0 ; iCount--)
	{
		frmScreen.cboNewStage.options.remove(iCount);
	}
}

function OrderByAddOption(sName)
{
	TagOPTION = document.createElement("OPTION");
	TagOPTION.value	= sName;
	TagOPTION.text = sName;
	frmScreen.cboOrderBy.add(TagOPTION);
}

function PopulateCombos()
{
	ResetCombos();
	// Get the TaskStatus combo information
	TaskStatusXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("TaskStatus");
	TaskStatusXML.GetComboLists(document, sGroups);

    // OrderBy
    // BM0226 - MV - 28/01/2003 - First STATUS
	OrderByAddOption("Status");
	OrderByAddOption("Date");
	OrderByAddOption("Task");
    OrderByAddOption("Outstanding");
	OrderByAddOption("Mandatory");
	OrderByAddOption("Owner");
}

function PopulateNewStageCombo(sActivitySeqNo)
{	
	// 'Move to new Stage' combo
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var XMLGlobal = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	<% /* PSC 19/05/2006 MAR1627 - Start */ %>
	var sApplicationType = scScreenFunctions.GetContextParameter(window, "idMortgageApplicationValue", null);
	var comboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	var bIsTOE = comboXML.IsInComboValidationList(document,"TypeOfMortgage", sApplicationType, ["TOE"])
	var bIsPSW = comboXML.IsInComboValidationList(document,"TypeOfMortgage", sApplicationType, ["PSW"])
	<% /* SR 05/03/2007 : EP2_1596 - both TMTOEEndStageID and TMPSWStageID are set to 90 - Completion_exception. 
	      Do not exclude these two stages */ %>
	<% /*
	var sTOEEndStageId = XMLGlobal.GetGlobalParameterString(document, "TMTOEEndStageID");
	var sPSWEndStageId = XMLGlobal.GetGlobalParameterString(document, "TMPSWStageID"); 
	var sStageExclusion = "";
		
	if (!bIsTOE) sStageExclusion += " and @STAGEID!='" + sTOEEndStageId + "'";		
	if (!bIsPSW) sStageExclusion += " and @STAGEID!='" + sPSWEndStageId + "'"; */ %>
	<% /* SR 05/03/2007 : EP2_1596 - End */ %>
	<% /* PSC 19/05/2006 MAR1627 - End */ %>
		
	XML.CreateRequestTag(window , "GetStageList");
	XML.CreateActiveTag("STAGE");
	XML.SetAttribute("ACTIVITYID", m_sActivityId);
	XML.SetAttribute("DELETEFLAG", "0");
	
	// 	XML.RunASP(document, "MsgTMBO.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	<%//GD BM0574 START
	//switch (ScreenRules())
		//{
		//case 1: // Warning
		//case 0: // OK %>
	XML.RunASP(document, "MsgTMBO.asp");
			<%//break;
		//default: // Error
			//XML.SetErrorResponse();
		//}
	// GD BM0574 END %>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		// Populate the Combo from the result xml
		// create a tag list based on STAGE
		<% /* PSC 12/11/2002 BMIDS00898 - Start */ %>
		
		<% /* BM0493 MDC 09/04/2003 */ %>
		// sCompStageId = XMLGlobal.GetGlobalParameterString(document, "TMCompletionsStageId");
		if(m_sCompStageId == "")
			sCompStageId = XMLGlobal.GetGlobalParameterString(document, "TMCompletionsStageId");
		else
			sCompStageId = m_sCompStageId;
		<% /* BM0493 MDC 09/04/2003 - End */ %>
			
		var xslPattern = "<?xml version='1.0'?>" +
						 "<xsl:stylesheet version='1.0' " +
						 "xmlns:xsl='http://www.w3.org/1999/XSL/Transform\'>" +
							"<xsl:template match='RESPONSE'>" +
								"<STAGELIST>" +
									"<xsl:for-each select='STAGE'> " +
										"<xsl:sort select='STAGESEQUENCENO'/>" +
										"<xsl:copy-of select='.'/>" +
									"</xsl:for-each>" +
								"</STAGELIST>" +
							"</xsl:template>" +
						 "</xsl:stylesheet>";
							 
		var xslDoc = new ActiveXObject("Microsoft.XMLDOM");
		xslDoc.loadXML(xslPattern);
		var sSorted = XML.XMLDocument.transformNode(xslDoc);
		XML.LoadXML(sSorted);		
			
		var xmlCompStage = XML.SelectSingleNode("STAGELIST/STAGE[@STAGEID='" + sCompStageId + "']"); 
				
		var xmlPreCompStage = null;
		if (xmlCompStage != null)
		{
			xmlPreCompStage = xmlCompStage.previousSibling;
			
			if (xmlPreCompStage != null)
				m_sPreCompStageId = xmlPreCompStage.getAttribute("STAGEID"); 	
		}
		XML.ActiveTag = null;
		var xmlLastStage = XML.SelectSingleNode("STAGELIST/STAGE[end()]");
		m_sLastStageId = xmlLastStage.getAttribute("STAGEID");
		
		<% /* PSC 12/11/2002 BMIDS00898 - End */ %>
		
		XML.ActiveTag = null;
		<% /* PSC 19/05/2006 MAR1627 */ %>
		<% /* SR 05/03/2007 : EP2_1596 :  
		XML.SelectNodes("STAGELIST/STAGE[@EXCEPTIONSTAGEINDICATOR='1'" + sStageExclusion + "]"); */ %>
		XML.SelectNodes("STAGELIST/STAGE[@EXCEPTIONSTAGEINDICATOR='1']");
		<% /* SR 05/03/2007 : EP2_1596 - End */ %>
		var iNoOfStages = XML.ActiveTagList.length;

		if (iNoOfStages > 0)
		{
			var iLoop;
			
			//Add first options of "Next Stage"
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= "";
			TagOPTION.text = "<SELECT>";
			frmScreen.cboNewStage.add(TagOPTION);
			<% /* PSC 12/11/2002 BMIDS00898 - Start */ %>
			if (m_sCurrStageId != m_sPreCompStageId && m_sCurrStageId != m_sLastStageId) 
			{
				TagOPTION = document.createElement("OPTION");
				TagOPTION.value	= "Next Stage";
				TagOPTION.text = "Next Stage";
				frmScreen.cboNewStage.add(TagOPTION);
			}
			<% /* PSC 12/11/2002 BMIDS00898  - End */ %>			
			for(iLoop = 0; iLoop < iNoOfStages; iLoop++)
			{
				XML.SelectTagListItem(iLoop);
				var sStageName = XML.GetAttribute("STAGENAME");

				TagOPTION = document.createElement("OPTION");
				TagOPTION.value	= sStageName;
				TagOPTION.text = sStageName;
				TagOPTION.setAttribute("StageId", XML.GetAttribute("STAGEID"));

				frmScreen.cboNewStage.add(TagOPTION);
			}
			frmScreen.cboNewStage.selectedIndex = 0;
		}
	}
}

function GetComboText(sTaskStatusValue)
{
<%	// return the valuename from the TaskStatus combo for the valueid sTaskStatusValue
%>	TaskStatusXML.SelectTag(null, "LISTNAME");
	TaskStatusXML.CreateTagList("LISTENTRY");	
	var sValueName = "";
	for(var iCount = 0; iCount < TaskStatusXML.ActiveTagList.length && sValueName == ""; iCount++)
	{
		TaskStatusXML.SelectTagListItem(iCount);
		if(TaskStatusXML.GetTagText("VALUEID") == sTaskStatusValue)	
			sValueName = TaskStatusXML.GetTagText("VALUENAME");
	}
	return sValueName;
}

function PopulateTable()
{
	stageXML.ActiveTag = null;
	stageXML.CreateTagList("CASETASK");
	var iNumberOfTasks = stageXML.ActiveTagList.length;

	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfTasks);
	scScrollTable.EnableMultiSelectTable();
	ShowList(0);
}

function GetComboValidationType(sTaskStatusValue)
{
<%	// SG 11/12/01 SYS3387
	// Returns the ComboValidation code for a given Task Status
%>	TaskStatusXML.SelectTag(null, "LISTNAME");
	TaskStatusXML.CreateTagList("LISTENTRY");
	var sValueName = "";
	for(var iCount = 0; iCount < TaskStatusXML.ActiveTagList.length && sValueName == ""; iCount++)
	{
		TaskStatusXML.SelectTagListItem(iCount);
		if(TaskStatusXML.GetTagText("VALUEID") == sTaskStatusValue)
			sValueName = TaskStatusXML.GetTagText("VALIDATIONTYPE");
	}
	return sValueName;
}

function ShowList(nStart)
{
<%  //This method populates the listbox based on the ORDER tag on each CASETASK. We must cater
	//for scrolling of the listbox too.
%>	var iCount;
	scScrollTable.clear();	
	
	for (iCount = 0; iCount < stageXML.ActiveTagList.length; iCount++)
	{
		stageXML.SelectTagListItem(iCount);
		var nOrder = parseInt(stageXML.GetTagText("ORDER"));
		if((nOrder - nStart) <= 0 | (nOrder - nStart) > m_iTableLength)
		{
			//ignore it
		}
		else
		{
			<% /* PSC 22/09/2005 MAR32 - Start */ %>
			sTaskName = stageXML.GetAttribute("CASETASKNAME");
			if(sTaskName.length == 0) sTaskName = stageXML.GetAttribute("TASKNAME");
			scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(0),sTaskName);
			
			if(stageXML.GetAttribute("TASENABLED") != null && stageXML.GetAttribute("TASENABLED") == "1")
				tblTable.rows(nOrder - nStart).cells(1).innerHTML="<img src=\"images/heart.gif\" height=\"11px\" width=\"11px\">";
			
			scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(2),stageXML.GetTagText("SYMBOL"));
			if(stageXML.GetTagText("SYMBOL") == "!!")
				tblTable.rows(nOrder - nStart).cells(2).style.color = "magenta";
			else tblTable.rows(nOrder - nStart).cells(2).style.color = "mediumorchid";
			tblTable.rows(nOrder - nStart).cells(2).style.font = "bold 10 arial";
			scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(3),GetComboText(stageXML.GetAttribute("TASKSTATUS")) );
			scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(4),GetMandatory(stageXML.GetAttribute("MANDATORYINDICATOR")) );

		<%	//SG 11/12/01 SYS3387
			//OLD CODE:
			//if(stageXML.GetAttribute("TASKDUEDATEANDTIME") != null && stageXML.GetAttribute("TASKDUEDATEANDTIME") != "")
			//	scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(4),stageXML.GetAttribute("TASKDUEDATEANDTIME"));
			//NEW CODE:
		%>
			var Status = stageXML.GetAttribute("TASKSTATUS");
			var ValidationType = GetComboValidationType(Status);
			if (ValidationType == "I")
				{
					if(stageXML.GetAttribute("TASKDUEDATEANDTIME") != null && stageXML.GetAttribute("TASKDUEDATEANDTIME") != "")
						scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(5),stageXML.GetAttribute("TASKDUEDATEANDTIME"));
				}
			else
				{
					if(stageXML.GetAttribute("TASKSTATUSSETDATETIME") != null && stageXML.GetAttribute("TASKSTATUSSETDATETIME") != "")
						scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(5),stageXML.GetAttribute("TASKSTATUSSETDATETIME"));			
				}
		<%	//END SG
		%>	
			scScreenFunctions.SizeTextToField(tblTable.rows(nOrder - nStart).cells(6),stageXML.GetAttribute("OWNINGUSERID"));
			
			//tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
			tblTable.rows(nOrder - nStart).setAttribute("TaskId",stageXML.GetAttribute("TASKID"));
			<% /* PSC 22/09/2005 MAR32 - End */ %>
		}
	}
}

function GetMandatory(sIndicator)
{
	var sMandatory = "";
	if(sIndicator == "1") sMandatory = "Yes";
	return sMandatory;
}

function frmScreen.cboOrderBy.onchange()
{
<%	//re-order the listbox entries by re-setting the ORDER tag text on each of the casetask entries.
	//To do this, use the Array object Sort method. Fill the array with the attribute chosen in the
	//combo. Sort it alphabetically (or by ascending date order). Loop the stageXML again and reset
	//the order of each casetask based on it's value and position in the Array. Call PopulateTable.
%>	var aSortArray = new Array();
	var sSortBy = "";
	stageXML.ActiveTag = null;
	stageXML.CreateTagList("CASETASK");
	switch(frmScreen.cboOrderBy.value)
	{
		case "Task" : sSortBy = "TASKNAME"; break;
		case "Outstanding" : sSortBy = "SYMBOL"; break;
		case "Status" : sSortBy = "TASKSTATUS"; break;
		case "Mandatory" : sSortBy = "MANDATORYINDICATOR"; break;
		case "Date" : sSortBy = "TASKDUEDATEANDTIME"; break;
		case "Owner": sSortBy = "OWNINGUSERID"; break;
		default: sSortBy = "TASKSTATUS";
	}
	
	for (var iCount = 0; iCount < stageXML.ActiveTagList.length; iCount++)
	{
		stageXML.SelectTagListItem(iCount);
		if(sSortBy == "SYMBOL")
			aSortArray[iCount] = stageXML.GetTagText("SYMBOL");
		else
		{
			if(stageXML.GetAttribute(sSortBy) == null)
				aSortArray[iCount] = " ";
			else
				aSortArray[iCount] = stageXML.GetAttribute(sSortBy);
		}
	}
	
	if(sSortBy == "TASKDUEDATEANDTIME")
		aSortArray.sort(sortByDateAsc);
	else if(sSortBy == "MANDATORYINDICATOR") 
		aSortArray.sort(sortByAlphaDesc);
	else if(sSortBy == "SYMBOL")
		aSortArray.sort(sortBySymbol);
	else
		aSortArray.sort();
	
	for (var iCount = 0; iCount < stageXML.ActiveTagList.length; iCount++)
	{
		stageXML.SelectTagListItem(iCount);
		var sAttrValue;
		if(sSortBy == "SYMBOL")
			sAttrValue = stageXML.GetTagText("SYMBOL");
		else
		{
			sAttrValue = stageXML.GetAttribute(sSortBy);
			if( sAttrValue == null) sAttrValue = " ";
		}
		stageXML.SetTagText("ORDER", GetOrderFromArray(sAttrValue, aSortArray));
	}
	
	PopulateTable();
	frmScreen.btnEdit.disabled = true;
//	BM0493 ik 20030328	
//	frmScreen.btnProcess.disabled = true;
	disableProcessButtons();
//	BM0493 ik 20030328 ends
	frmScreen.btnTaskHistory.disabled = true;
}

function GetOrderFromArray(sValue, sArray)
{
<%	//loops the array to match the value passed in. Return the array position.
	//If there are identical values in the array we want to get each array position and not just
	//the first one we hit, so when we've found a value, alter the array value for that position
	//so that the sValue will not match that position again and will match the next one.
%>	var sOrder = "";
	for(var iCount = 0; iCount < sArray.length && sOrder == ""; iCount++)
	{
		if(sValue == sArray[iCount]) 
		{
			sOrder = (iCount+1);
			sArray[iCount] = sValue + "  ZZZ";
		}
	}
	return sOrder;
}

function sortByDateAsc(sArg1, sArg2)
{
	if(sArg1 == "" || sArg2 == "")
	{
		return -1;
	}
	var sDate1 = new Date(sArg1.substr(6,4), (sArg1.substr(3,2) -1), sArg1.substr(0,2), sArg1.substr(11,2), sArg1.substr(14,2), sArg1.substr(17,2));
	var sDate2 = new Date(sArg2.substr(6,4), (sArg2.substr(3,2) -1), sArg2.substr(0,2), sArg2.substr(11,2), sArg2.substr(14,2), sArg2.substr(17,2));
	if(sDate1 < sDate2)
	{
		return -1;
	}
	else if(sDate1 > sDate2) 
	{
		return 1;
	}
	return 0;
}

function sortByAlphaDesc(sArg1, sArg2)
{
	if(sArg2 < sArg1) 
	{
		return -1;
	}
	else if(sArg2 > sArg1) 
	{
		return 1;
	}
	return 0;
}

function sortBySymbol(sArg1, sArg2)
{
<%	// order as "!!", "*", ""
%>	var nNumArg1;
	var nNumArg2;
	switch(sArg1)
	{
		case "!!": nNumArg1 = 1; break;
		case "*": nNumArg1 = 2; break;
		default: nNumArg1 = 3; break;
	}
	switch(sArg2)
	{
		case "!!": nNumArg2 = 1; break;
		case "*": nNumArg2 = 2; break;
		default: nNumArg2 = 3; break;
	}
	
	if(nNumArg1 < nNumArg2) 
	{
		return -1;
	}
	else if(nNumArg1 > nNumArg2) 
	{
		return 1;
	}
	return 0;
}

function spnTable.ondblclick()
{
	if (scScrollTable.getRowSelectedIndex() != null) 
//	BM0493 ik 20030328	
//		frmScreen.btnProcess.onclick();
	{
		var sXML = GetTasksXML();
		m_taskXML.LoadXML(sXML);
		scScreenFunctions.SetContextParameter(window,"idTaskXML",sXML);
		initTaskDetail();
		<% /* BS BM0271 16/04/03 
		If ReadOnly because it's a non-processing unit or the case is locked then enable 
		the Details and Memo buttons, otherwise call enableProcessButtons */ %>
		if((m_sReadOnlyAsCaseLocked == "1") || (m_sProcessInd == "0"))
		{
			frmScreen.btnDetails.disabled = false;
			frmScreen.btnMemo.disabled = false;
		}
		else enableProcessButtons();
	}
//	BM0493 ik 20030328 ends
}

function spnTable.onclick()
{
	var saRowArray = scScrollTable.getArrayofRowsSelected();
	if(saRowArray.length > 0)
	{
		<% /* BS BM0271 16/04/03
		Only enable the Edit button if it's a processing unit and the case is not locked */ %>
		if((m_sReadOnlyAsCaseLocked != "1") && (m_sProcessInd == "1"))
			frmScreen.btnEdit.disabled = false;
		// frmScreen.btnProcess.disabled = false;
		frmScreen.btnTaskHistory.disabled = false;
//	BM0493 ik 20030328	
		var sXML = GetTasksXML();
		m_taskXML.LoadXML(sXML);
		scScreenFunctions.SetContextParameter(window,"idTaskXML",sXML);
		initTaskDetail();
		<% /* BS BM0271 16/04/03 
		If ReadOnly because it's a non-processing unit or the case is locked then enable 
		the Details and Memo buttons, otherwise call enableProcessButtons */ %>
		if((m_sReadOnlyAsCaseLocked == "1") || (m_sProcessInd == "0"))
		{
			frmScreen.btnDetails.disabled = false;
			frmScreen.btnMemo.disabled = false;
		}
		else enableProcessButtons();
//	BM0493 ik 20030328 ends
	}
	if(saRowArray.length > 1)
		frmScreen.btnTaskHistory.disabled = true;
}

function frmScreen.btnAdd.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>
	var sReturn = null;
	stageXML.SelectTag(null, "CASESTAGE");
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	<% /* MAR667  Pass UnderwritingTasksOnly flag */ %>
	var bUnderwritingTasksOnly = false;
	<% /*AW 14/09/06	EP1103 */ %>
	var bProgressTaskMode = false;
	
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sReadOnly;
	ArrayArguments[1] = m_sCurrStageId;
	ArrayArguments[2] = m_sApplicationNumber;
	ArrayArguments[3] = m_sApplicationPriority ;
	ArrayArguments[4] = m_iUserRole ;
	ArrayArguments[5] = m_sActivityId;
	ArrayArguments[6] = GetCustomerXML();
	ArrayArguments[7] = XML.CreateRequestAttributeArray(window);
	ArrayArguments[8] = stageXML.GetAttribute("CASESTAGESEQUENCENO");
	ArrayArguments[9] = bUnderwritingTasksOnly;
	<% /*AW 14/09/06	EP1103 */ %>
	ArrayArguments[10] = bProgressTaskMode;
	sReturn = scScreenFunctions.DisplayPopup(window, document, "TM031.asp", ArrayArguments, 538, 418);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		var sTaskXML = sReturn[1];
		if(sTaskXML != "")
		{
			// we have some tasks to add to the current stage.
			AddTasksToStage(sTaskXML);

		}
	}
}

function GetCustomerXML()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "");
	var tagCustomers = XML.CreateActiveTag("CUSTOMERS");
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName			= scScreenFunctions.GetContextParameter(window,"idCustomerName" + nLoop, "Jane" + nLoop);
		var sCustomerNumber			= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop, nLoop);
		var sCustomerVersionNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop, "1");
		<% /* EP2_576  Pass Customer Role */ %>
		var sCustomerRole	= scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop, "1");

		if(sCustomerName != "" && sCustomerNumber != "")
		{
			XML.ActiveTag = tagCustomers;
			XML.CreateActiveTag("CUSTOMER");
			XML.SetAttribute("NUMBER", sCustomerNumber);
			XML.SetAttribute("NAME", sCustomerName);
			XML.SetAttribute("VERSION", sCustomerVersionNumber);
			XML.SetAttribute("ROLE", sCustomerRole);
		}
	}
	return(XML.XMLDocument.xml);
}

function AddTasksToStage(sTaskXML)
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(sTaskXML);
	
	<% /*
	//APS SYS1920
	var sTaskId = new Array();
	
	XML.SelectTag(null,"REQUEST");
	XML.CreateTagList("CASETASK");
	
	for (var i=0; i<XML.ActiveTagList.length; i++){
		XML.SelectTagListItem(i);
		sTaskId[i] = XML.GetAttribute("TASKID");
	}
	*/ %>
	
	<% // 	XML.RunASP(document, "OmigaTMBO.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115 %>
	<%//GD BM0574 START
	//switch (ScreenRules())
		//{
		//case 1: // Warning
		//case 0: // OK %>
	XML.RunASP(document, "OmigaTMBO.asp");
			<%//break;
		//default: // Error
			//XML.SetErrorResponse();
		//}
	//GD BM0574 END %>
	
	var ErrorTypes = new Array("NOTASKAUTHORITY");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{
		alert("User does not have the authority to create this task");
	}
	else if(ErrorReturn[0] == true)
	{		
	   <% //DRC BMIDS670 
		//for (var i=0; i<sTaskId.length; i++) {
		   
			//if (sTaskId[i].toUpperCase() == m_sTMReIssueOffer){
			//	scScreenFunctions.SetContextParameter(window, "idFreezeDataIndicator", "0");
			//	break ;
			//}
		
		//}
		%>
		//re-populate the listbox for this stage
		<% /* BM0493 MDC 15/04/2003 - Refresh display
		GetStageInfo();
		PopulateTable(); */ %>
		CheckFreezeUnFreeze(m_sApplicationNumber);
		updateList();
		<% /* BM0493 MDC 15/04/2003 - End */ %>

	}
}

function GetTasksXML()
{
<%	//For each row of the table selected get the associated CASETASK xml and return them all
%>	var saRowArray = scScrollTable.getArrayofRowsSelected(); //THIS MIGHT ONLY BE THE VISIBLE ONES
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var tagRequest = XML.CreateRequestTag(window , "");
	for(var iCount = 0; iCount < saRowArray.length; iCount++)
	{
		var nOrderNum = saRowArray[iCount];
		var newNode = SetActiveTask(nOrderNum);
		var node = newNode.cloneNode(true);
		XML.ActiveTag.appendChild(node);
	}
	return(XML.XMLDocument.xml);
}

function SetActiveTask(nOrderNum)
{
	stageXML.ActiveTag = null;
	stageXML.CreateTagList("CASETASK");
	var node = null;
	for (iCount = 0; iCount < stageXML.ActiveTagList.length && node == null; iCount++)
	{
		stageXML.SelectTagListItem(iCount); //sets this tag as the active tag
		var nThisOrder = parseInt(stageXML.GetTagText("ORDER"));
		if(nThisOrder == nOrderNum)
			node = stageXML.GetTagListItem(iCount);
	}
	return node;
}

function frmScreen.btnTaskHistory.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>
	var sReturn = null;
	var ArrayArguments = new Array();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ArrayArguments[0] = m_sReadOnly;
	ArrayArguments[1] = GetTasksXML();
	ArrayArguments[2] = XML.CreateRequestAttributeArray(window);
	sReturn = scScreenFunctions.DisplayPopup(window, document, "TM035.asp", ArrayArguments, 550, 412);
}

/*
function frmScreen.btnProcess.onclick()
{
<%	//get the CASETASK xml element for the selected rows
%>	scScreenFunctions.SetContextParameter(window,"idTaskXML",GetTasksXML());
	frmToTM032.submit();
}
*/

function frmScreen.btnEdit.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>
<%	//get the CASETASK xml element for the selected rows. Also set idMetaAction to tell TM033 where to return to
	//Check the selected tasks - cannot edit any completed ones.
%>	var sTasksXML = GetTasksXML();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bCompletedTask = false;
	XML.LoadXML(sTasksXML);
	XML.SelectTag(null, "REQUEST");
	XML.CreateTagList("CASETASK");
	for(var iCount = 0; iCount < XML.ActiveTagList.length && bCompletedTask == false; iCount++)
	{
		XML.SelectTagListItem(iCount);
		var sTaskStatus = XML.GetAttribute("TASKSTATUS");
		var comboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if(comboXML.IsInComboValidationList(document,"TaskStatus", sTaskStatus, ["C"]))
			bCompletedTask = true;			
	}
	if(bCompletedTask == false)
	{
		<% /* BM0493 MDC 04/04/2003 */ %>
		<% /*
		scScreenFunctions.SetContextParameter(window,"idTaskXML",sTasksXML);
		scScreenFunctions.SetContextParameter(window,"idMetaAction","TM030");
		frmToTM033.submit();
		*/ %>
		
		var sReturn = null;
		var ArrayArguments = new Array();
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		ArrayArguments[0] = m_sReadOnly;
		ArrayArguments[1] = m_taskXML.XMLDocument.xml;
		ArrayArguments[2] = XML.CreateRequestAttributeArray(window);
		ArrayArguments[3] = m_iUserRole
		sReturn = scScreenFunctions.DisplayPopup(window, document, "TM033P.asp", ArrayArguments, 398, 370);
		if(sReturn != null)
			updateList();
		<% /* BM0493 MDC 04/04/2003 - End */ %>
			
	}
	else alert("Cannot Edit tasks which are in a completed state.");
}

function frmScreen.btnStageHistory.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>
	frmToTM037.submit();		
}

function frmScreen.cboNewStage.onchange()
{
	if(frmScreen.cboNewStage.value != "")
		frmScreen.btnConfirm.disabled = false;
	else frmScreen.btnConfirm.disabled = true;
}
function frmScreen.btnConfirm.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	
	<% /* BMIDS681 GHun Disable the button and display an hourglass until processing is complete */ %>
	frmScreen.btnConfirm.blur();
	frmScreen.btnConfirm.disabled = true;
	frmScreen.btnConfirm.style.cursor = "wait";
	<% /* Call the ConfirmProcessing after a timeout to allow the cursor time to change. */ %>
	window.setTimeout("ConfirmProcessing()", 0)
	<% /* BMIDS681 End */ %>
}

function ConfirmProcessing()
{
	<% // GD BM0574 END %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var iSelectedIndex = frmScreen.cboNewStage.selectedIndex;
	var ReqTag;
	if (frmScreen.cboNewStage.value == "Next Stage")
	   
	{	
		MoveToNextStage();
	} 
	else
	{
	  // SYS1786 -  Exception stage
		var sReason = "";
    	var ArrayArguments = new Array();
		sReturn = scScreenFunctions.DisplayPopup(window, document, "TM038.asp", ArrayArguments, 550, 200);  
		
		if (sReturn != null)
		  sReason = sReturn[1];
					
		if(sReason != "")
		{	
			var sLocalPrinters = GetLocalPrinters();
			LocalPrintersXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			LocalPrintersXML.LoadXML(sLocalPrinters);
			LocalPrintersXML.SelectTag(null, "RESPONSE");	
		
			var sPrinter = LocalPrintersXML.GetTagText("PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
									
			if(sPrinter == "")
			{					
				alert("You do not have a default printer set on your PC.");			
			}		
			else
			{
				ReqTag = XML.CreateRequestTag(window , "MoveToStage");
			
				XML.CreateActiveTag("CASESTAGE")
				XML.SetAttribute("STAGEID", frmScreen.cboNewStage.options(iSelectedIndex).getAttribute("StageId"));
				XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
				XML.SetAttribute("CASEID", m_sApplicationNumber);
				XML.SetAttribute("ACTIVITYID", m_sActivityId);
				XML.SetAttribute("ACTIVITYINSTANCE", "1");
				XML.SetAttribute("EXCEPTIONREASON",sReason);
				<% /* BMIDS681 No need to pass CASESTAGESEQUENCENO as it is misspelt as contains the wrong value
				XML.SetAttribute("CASESTAGESEQUENCENUMBER", scScreenFunctions.GetContextParameter(window,"idCurrentStageOrigSeqNo",""));
				*/ %>
			
				XML.ActiveTag = ReqTag;

				<% /* BM0340 MDC 08/05/2003 */ %>
				// XML.CreateActiveTag("APPLICATION");
				var AppXML = XML.CreateActiveTag("APPLICATION");
				<% /* BM0340 MDC 08/05/2003 - End */ %>

				XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
				XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
				XML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window,"idApplicationPriority",""));
			
				var iCount = 0;
				var sCustomerVersionNumber = "";
				var sCustomerNumber = "";
				var sOtherSystemCustomerNumber = "";	<% /* PSC 26/01/2006 MAR1133 */ %>
				
				for (iCount = 1; iCount <= 5; iCount++)
				{								
					sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
					if (sCustomerNumber != "")
					{	
						sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);	
						<% /* PSC 26/01/2006 MAR1133 */ %>
						sOtherSystemCustomerNumber = scScreenFunctions.GetContextParameter(window,"idOtherSystemCustomerNumber" + iCount,null);	

						XML.SelectTag(null, "APPLICATION");
						XML.CreateActiveTag("CUSTOMER");				
						XML.SetAttribute("CUSTOMERNUMBER", sCustomerNumber);
						XML.SetAttribute("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);																		
						<% /* PSC 26/01/2006 MAR1133 */ %>
						XML.SetAttribute("OTHERSYSTEMCUSTOMERNUMBER", sOtherSystemCustomerNumber);
					}
				}
			
				XML.ActiveTag = ReqTag;
				XML.CreateActiveTag("PRINTER");
				XML.SetAttribute("PRINTERNAME", sPrinter);
				XML.SetAttribute("DEFAULTIND", "1");
			
				// 				XML.RunASP(document, "OmigaTMBO.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				<%//GD BM0574 START
				//switch (ScreenRules())
					//{
					//case 1: // Warning
					//case 0: // OK
						 // BM0340 MDC 08/05/2003  %>
				window.status = "Moving to stage...";	
						<%// XML.RunASP(document, "OmigaTMBO.asp"); %>
				XML.RunASP(document, "omTmNoTxBO.asp");
						<%//window.status = "";	
						// BM0340 MDC 08/05/2003 - End 
						//break;
					//default: // Error
						//XML.SetErrorResponse();
					//}
				//GD BM0574 END %>
				var ErrorTypes = new Array("RECORDNOTFOUND", "NOSTAGEAUTHORITY");
				var ErrorReturn = XML.CheckResponse(ErrorTypes);
				if (ErrorReturn[1] == ErrorTypes[1])
				    alert("User does not have authority to create an exception stage");
				else if(ErrorReturn[1] == ErrorTypes[0])
					alert('Database Record cannot be created');
				else  if (ErrorReturn[0] == true )
				{
				    m_sReadOnly = 1;
				    scScreenFunctions.SetContextParameter(window,"idCancelDeclineFreezeDataIndicator", "1");
				}
				
				<% /* BMIDS649 GHun 13/01/2004 ProcessAutomaticTasks is called within MoveToStage in
				      the middle tier so that it will still work when called by ingestion 
				// BM0340 MDC 08/05/2003
				{
					window.status = "Stage change completed - actioning automatic tasks. Please wait...";
					ProcessAutomaticTasks(sPrinter, AppXML);

   				    m_sReadOnly = 1 ;
				}
				BMIDS649 End */ %>
				
				// BMIDS468 KRW 19/04/04
				CheckFreezeUnFreeze(m_sApplicationNumber); 
				
				window.status = "";
				
				<% /* BM0340 MDC 08/05/2003 - End */ %>
				
			}			
		}     			
	}
	
	<% /* BMIDS681 Restore the mouse cursor for the button */ %>
	frmScreen.btnConfirm.style.cursor = "hand";
	
	PopulateScreen();
	// AQR SYS4530 - Check whether a new AdminSys account number was added in an Automatic task
	<%/* SG 06/06/02 SYS4822 */%>
	<%/* scScrFunctions.SetOtherSystemACNoInContext(m_sApplicationNumber); */%>
	scScreenFunctions.SetOtherSystemACNoInContext(m_sApplicationNumber);
	
}

function btnSubmit.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>
<% // An example submit function showing the use of the validation functions 
%>	if(frmScreen.onsubmit())
	{
		//if(IsChanged())
<%			// Do some processing
%>
		frmToMN070.submit();
	}
}

function GetIssueOfferParameter()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var sTMIssueOffer = XML.GetGlobalParameterString(document, "TMIssueOffer");

	XML = null;

	return sTMIssueOffer.toUpperCase();
}

function GetReIssueOfferParameter()
{
	// SYS1920
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	<% /* BM0493 MDC 09/04/2003 */ %>
	// m_sTMReIssueOffer = XML.GetGlobalParameterString(document, "TMReIssueOffer");
	if(m_sTMReIssueOffer == "")
		m_sTMReIssueOffer = XML.GetGlobalParameterString(document, "TMReIssueOffer");
	<% /* BM0493 MDC 09/04/2003 - End */ %>

	m_sTMReIssueOffer = m_sTMReIssueOffer.toUpperCase();
	<%//GD BM0566 START %>
	return(m_sTMReIssueOffer);
	<%//GD BM0566 END   %>
}

function QueryStageComplete()
{
	<% /* PSC 13/11/2002 BMIDS00898 */ %>
	<% /* PSC 18/01/2006 MAR1081 */ %>
	if((frmScreen.txtTotalIncomplete.value == "0") &&
	   (m_sCurrStageId != m_sCancelStageId && m_sCurrStageId != m_sDeclineStageId &&
	    m_sCurrStageId != m_sLastStageId && m_sCurrStageId != m_sPreCompStageId &&
	    m_sCurrStageId != m_sTOEEndStageId && m_sCurrStageId != m_sPSWStageId))
		{
            if (confirm("The current stage is complete. Do you wish to move this application to the next stage? (OK = Yes, Cancel = No)"))
            {
				MoveToNextStage();
            }
			else
			{
			    var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XML.CreateRequestTag(window , "CreateCaseStageTrigger");
				XML.CreateActiveTag("CASETASK");
				XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
				XML.SetAttribute("CASESTAGESEQUENCENO", stageXML.GetAttribute("CASESTAGESEQUENCENO"));
				XML.SetAttribute("CASEID", m_sApplicationNumber);
				XML.SetAttribute("ACTIVITYID", m_sActivityId);
				XML.SetAttribute("ACTIVITYINSTANCE", "1");
				XML.SetAttribute("STAGEID", m_sCurrStageId);
				XML.RunASP(document, "OmigaTMBO.asp");
				XML.IsResponseOK();
			}
			PopulateScreen();
		}
}
function MoveToNextStage()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var iSelectedIndex = frmScreen.cboNewStage.selectedIndex;
	var ReqTag;
	var sLocalPrinters = GetLocalPrinters();
	LocalPrintersXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	LocalPrintersXML.LoadXML(sLocalPrinters);
	LocalPrintersXML.SelectTag(null, "RESPONSE");	
	
	var sPrinter = LocalPrintersXML.GetTagText("PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
							
	if(sPrinter == "")
	{					
		alert("You do not have a default printer set on your PC.");			
	}		
	else
	{
		
		ReqTag = XML.CreateRequestTag(window , "MoveToNextStage");
		
		XML.CreateActiveTag("CURRENTSTAGE")
		XML.SetAttribute("STAGEID", m_sCurrStageId);
		XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		XML.SetAttribute("CASEID", m_sApplicationNumber);
		XML.SetAttribute("ACTIVITYID", m_sActivityId);
		XML.SetAttribute("ACTIVITYINSTANCE", "1");
		<% /* BMIDS681 No need to pass CASESTAGESEQUENCENO as it is misspelt as contains the wrong value
		XML.SetAttribute("CASESTAGESEQUENCENUMBER", scScreenFunctions.GetContextParameter(window,"idCurrentStageOrigSeqNo",""));
		*/ %>
		
		XML.ActiveTag = ReqTag;
		<% /* BM0340 MDC 22/04/2003 */ %>
		// XML.CreateActiveTag("APPLICATION");
		var AppXML = XML.CreateActiveTag("APPLICATION");
		<% /* BM0340 MDC 22/04/2003 - End */ %>
		XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		XML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window,"idApplicationPriority",""));
		
		var iCount = 0;
		var sCustomerVersionNumber = "";
		var sCustomerNumber = "";
		var sOtherSystemCustomerNumber = "";	<% /* PSC 26/01/2006 MAR1133 */ %>

		for (iCount = 1; iCount <= 5; iCount++)
		{								
			sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
			<% /* PSC 26/01/2006 MAR1133 */ %>
			sOtherSystemCustomerNumber = scScreenFunctions.GetContextParameter(window,"idOtherSystemCustomerNumber" + iCount,null);	

			if (sCustomerNumber != "")
			{	
				sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);	
				XML.SelectTag(null, "APPLICATION");
				XML.CreateActiveTag("CUSTOMER");				
				XML.SetAttribute("CUSTOMERNUMBER", sCustomerNumber);
				XML.SetAttribute("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);																		
				<% /* PSC 26/01/2006 MAR1133 */ %>
				XML.SetAttribute("OTHERSYSTEMCUSTOMERNUMBER", sOtherSystemCustomerNumber);
			}
		}
			
		XML.ActiveTag = ReqTag;
		XML.CreateActiveTag("PRINTER");
		XML.SetAttribute("PRINTERNAME", sPrinter);
		XML.SetAttribute("DEFAULTIND", "1");
				
		// 		XML.RunASP(document, "OmigaTMBO.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		<%//GD BM0574 START
		//switch (ScreenRules())
			//{
			//case 1: // Warning
			//case 0: // OK
				// BM0340 MDC 22/04/2003 %>
		window.status = "Moving to next stage...";	
				<%// XML.RunASP(document, "OmigaTMBO.asp");%>
		XML.RunASP(document, "omTmNoTxBO.asp");
				<%//window.status = "";	
				// BM0340 MDC 22/04/2003 - End 
				//break;
			//default: // Error
				//XML.SetErrorResponse();
			//} GD BM0574 END%>

		var ErrorTypes = new Array("RECORDNOTFOUND", "NOSTAGEAUTHORITY");
		var ErrorReturn = XML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[1])
			alert("User does not have authority to create the next stage");
		else if (ErrorReturn[1] == ErrorTypes[0]) // error 292
		{
				if(confirm("There are no further stages.  Do you wish to complete this Activity?"))

				{
					XML.ResetXMLDocument();
					XML.CreateRequestTag(window , "SetCurrentCaseStageComplete");		//JLD SYS1807
					XML.CreateActiveTag("CASESTAGE");
					XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
					XML.SetAttribute("CASEID", m_sApplicationNumber);
					XML.SetAttribute("CASESTAGESEQUENCENO", stageXML.GetAttribute("CASESTAGESEQUENCENO"));
					XML.SetAttribute("ACTIVITYID", m_sActivityId);
					XML.SetAttribute("STAGEID", m_sCurrStageId);
					// 					XML.RunASP(document, "MsgTMBO.asp");
					// Added by automated update TW 09 Oct 2002 SYS5115
					<%//GD BM0574 START
					//switch (ScreenRules())
						//{
						//case 1: // Warning
						//case 0: // OK %>
					XML.RunASP(document, "MsgTMBO.asp");
						<%	//break;
						//default: // Error
							//XML.SetErrorResponse();
						//} GD BM0574 END %>

					XML.IsResponseOK();
					
					XML.ResetXMLDocument();
					XML.CreateRequestTag(window , "CompleteCaseActivity");
					XML.CreateActiveTag("CASEACTIVITY");
					XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
					XML.SetAttribute("CASEID", m_sApplicationNumber);
					XML.SetAttribute("ACTIVITYID", m_sActivityId);
					// 					XML.RunASP(document, "MsgTMBO.asp");
					// Added by automated update TW 09 Oct 2002 SYS5115
					<%//GD BM0574 START
					//switch (ScreenRules())
						//{
						//case 1: // Warning
						//case 0: // OK %>
					XML.RunASP(document, "MsgTMBO.asp");
						<%	//break;
						//default: // Error
						//	XML.SetErrorResponse();
						//} GD BM0574 END %>
					XML.IsResponseOK();
				}  
				else
				{
					XML.ResetXMLDocument();
					XML.CreateRequestTag(window , "CreateCaseStageTrigger");
					XML.CreateActiveTag("CASETASK");
					XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
					XML.SetAttribute("CASESTAGESEQUENCENO", stageXML.GetAttribute("CASESTAGESEQUENCENO"));
					XML.SetAttribute("CASEID", m_sApplicationNumber);
					XML.SetAttribute("ACTIVITYID", m_sActivityId);
					XML.SetAttribute("ACTIVITYINSTANCE", "1");
					XML.SetAttribute("STAGEID", m_sCurrStageId);
					// 					XML.RunASP(document, "OmigaTMBO.asp");
					// Added by automated update TW 09 Oct 2002 SYS5115
					<%//GD BM0574 START
					//switch (ScreenRules())
						//{
						//case 1: // Warning
						//case 0: // OK %>
					XML.RunASP(document, "OmigaTMBO.asp");
						<% //	break;
						//default: // Error
						//	XML.SetErrorResponse();
						//}
					//GD BM0574 END %>
					XML.IsResponseOK();
				}  
		    }
		    <% /* MO - 25/11/2002 - BMIDS01076 - Removed as this line of code is displayed if any error 
												 occurs , including "The following automatic task failed!"
		    else if (ErrorReturn[0] == false) 
		    {
		  	  alert('Database Record cannot be created')
			}  */ %>
			
			<% /* BMIDS649 GHun 13/01/2004 ProcessAutomaticTasks is called within MoveToNextStage
			      in the middle tier, so that is still works when called by ingestion
			// BM0340 MDC 22/04/2003
			else if(ErrorReturn[0] == true)
			{
				window.status = "Stage change completed - actioning automatic tasks. Please wait...";
				ProcessAutomaticTasks(sPrinter, AppXML);
			}
			// BM0340 MDC 22/04/2003 - End
			// BMIDS649 End */ %>
			
		}
	<% /* BM0567 START */ %>
	<% /* BMIDS670 - revisted */ %> 
	 //var sResult = IsApplicationAtOfferStage(m_sApplicationNumber);
	CheckFreezeUnFreeze(m_sApplicationNumber); 
	//if (sResult != -1) //Call has succeeded
	//{
	//	scScreenFunctions.SetContextParameter(window,"idFreezeDataIndicator", sResult);
	//}


	<% /* BM0670END   */ %>
	<% /* BM0567 END   */ %>
	
	<% /* BM0340 MDC 22/04/2003 */ %>
	window.status = "";
	<% /* BM0340 MDC 22/04/200 End */ %>

}

<% /* BMIDS649 GHun 13/01/2004 The call to ProcessAutomaticTasks has been moved back
to the middle tier, otherwise they don't run with ingestion
// BM0340 MDC 22/04/2003
function ProcessAutomaticTasks(sPrinter, xmlApplication)
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var iSelectedIndex = frmScreen.cboNewStage.selectedIndex;
	var ReqTag;

	ReqTag = XML.CreateRequestTag(window , "ProcessAutomaticTasks");
			
	XML.CreateActiveTag("CURRENTSTAGE")
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("CASEID", m_sApplicationNumber);
	XML.SetAttribute("ACTIVITYID", m_sActivityId);
	XML.SetAttribute("ACTIVITYINSTANCE", "1");
			
	XML.ActiveTag = ReqTag;
	XML.ActiveTag.appendChild(xmlApplication);
				
	XML.ActiveTag = ReqTag;
	XML.CreateActiveTag("PRINTER");
	XML.SetAttribute("PRINTERNAME", sPrinter);
	XML.SetAttribute("DEFAULTIND", "1");

	XML.RunASP(document, "omTmNoTxBO.asp");
	XML.IsResponseOK();

}
// BM0340 MDC 22/04/2003 - End
// BMIDS649 End */ %>

function enableProcessButtons()
{
	var bActionButtonState = false;
	var bDetailsButtonState = false;
	var bReprintButtonState = false;
	var bChaseUpButtonState = false;
	var bNotApplicableButtonState = false;
	var bMemoButtonState = false;
	var bContactDetailsButtonState = false;
	var bMixIncludesCompleted = false;
	var bMultiInput = false;
	<% /* PSC 22/09/2005 MAR32 - Start */ %>
	var XMLCombos = TaskStatusXML.GetComboListXML("TaskStatus");
	var sTASFailedValue = TaskStatusXML.GetComboIdForValidation("TaskStatus", "TF", XMLCombos, document); 
	var sTARFailedValue = TaskStatusXML.GetComboIdForValidation("TaskStatus", "TRF", XMLCombos, document); 
	<% /* PSC 22/09/2005 MAR32 - End */ %>

	m_taskXML.SelectTag(null, "REQUEST");
	m_taskXML.CreateTagList("CASETASK");

	if(m_taskXML.ActiveTagList.length > 1) 
	{
		bMultiInput = true;
		
		m_taskXML.SelectTag(null, "REQUEST");
		m_taskXML.CreateTagList("CASETASK[$not$ @OUTPUTDOCUMENT or @OUTPUTDOCUMENT='']");
		//if outputdocument has a value in every case
		if(m_taskXML.ActiveTagList.length == 0)
		{	
			//if status is not 10 (Not Actioned) in all cases			
			m_taskXML.SelectTag(null, "REQUEST");
			<% /* PSC 22/09/2005 MAR32 */ %>
			m_taskXML.CreateTagList("CASETASK[@TASKSTATUS!='10' or @TASKSTATUS!='" + sTASFailedValue + "' or @TASKSTATUS!='" + sTARFailedValue + "']");
			if(m_taskXML.ActiveTagList.length > 0)			
			{	
				m_taskXML.SelectTag(null, "REQUEST");
				m_taskXML.CreateTagList("CASETASK[@INPUTPROCESS!='']");
				if(m_taskXML.ActiveTagList.length > 0)
				{
					bActionButtonState = true;
				}
			}
		}
		else 	
		{	//BG 06/04/01 if there are any inputprocess attributes with values in, disable action button
			m_taskXML.SelectTag(null, "REQUEST");
			m_taskXML.CreateTagList("CASETASK[@INPUTPROCESS!='']");
			if(m_taskXML.ActiveTagList.length > 0)
			{
				bActionButtonState = true;
			}
			
			//BG 06/04/01 if there are any inputprocess INTERFACE with values in, disable action button
			m_taskXML.SelectTag(null, "REQUEST");
			m_taskXML.CreateTagList("CASETASK[@INTERFACE!='']");
		
			if(m_taskXML.ActiveTagList.length > 0)	
			{
				bActionButtonState = true;	
			}			
		}		
	}
	m_taskXML.SelectTag(null, "REQUEST");
	m_taskXML.CreateTagList("CASETASK");
	for(var iCount = 0; iCount < m_taskXML.ActiveTagList.length; iCount++)
	{
		m_taskXML.SelectTagListItem(iCount);
		
		<% /* PSC 22/09/2005 MAR32 - Start */ %>
		var comboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var bTASInProgress = false;
		var bTASInFailed = false;
		
		if(comboXML.IsInComboValidationList(document,"TaskStatus", m_taskXML.GetAttribute("TASKSTATUS"), ["TR"]))
			bTASInProgress = true;

		if(comboXML.IsInComboValidationList(document,"TaskStatus", m_taskXML.GetAttribute("TASKSTATUS"), ["TF", "TRF"]))
			bTASInFailed = true;
		<% /* PSC 22/09/2005 MAR32 - End */ %>

		if( bMultiInput && 
		    (m_taskXML.GetAttribute("TASKSTATUS") == "40" || m_taskXML.GetAttribute("TASKSTATUS") == "70"))
				bMixIncludesCompleted = true; //SYS1762 
		//if( bMultiInput &&
		//	((m_taskXML.GetAttribute("INPUTPROCESS") != null && m_taskXML.GetAttribute("INPUTPROCESS") != "") ||
		//	 (m_taskXML.GetAttribute("OUTPUTDOCUMENT") != null && m_taskXML.GetAttribute("OUTPUTDOCUMENT") != "") ||
		//	 (m_taskXML.GetAttribute("INTERFACE") != null && m_taskXML.GetAttribute("INTERFACE") != ""))        )
		//		bActionButtonState = true;
		<% /* PSC 22/09/2005 MAR32 */ %>
		if( m_taskXML.GetAttribute("TASKSTATUS") != "10" && !bTASInFailed) // "Not Actioned"
				bActionButtonState = true;
		<% /* PSC 22/09/2005 MAR32 */ %>
		if( bMultiInput ||
		    bTASInProgress ||
		    (m_taskXML.GetAttribute("TASKSTATUS") == "10" || //"Not Actioned"
		     m_taskXML.GetAttribute("TASKSTATUS") == "30" || //"Not Applicable"
		     (m_taskXML.GetAttribute("INPUTPROCESS") == null || m_taskXML.GetAttribute("INPUTPROCESS") == "")) )
				bDetailsButtonState = true;
		if( bMultiInput ||
		    ((m_taskXML.GetAttribute("OUTPUTDOCUMENT") == null || m_taskXML.GetAttribute("OUTPUTDOCUMENT") == "") ||
		     m_taskXML.GetAttribute("TASKSTATUS") != "20" ) )  //"Pending"
				bReprintButtonState = true;
		if( bMultiInput ||
		    (((m_taskXML.GetAttribute("CHASINGTASK") == null || m_taskXML.GetAttribute("CHASINGTASK") == "")
		      && 
		      (m_taskXML.GetAttribute("CONTACTTYPE") == null || m_taskXML.GetAttribute("CONTACTTYPE") == "")) ||
		    m_taskXML.GetAttribute("TASKSTATUS") != "20") )     //"Pending"
				bChaseUpButtonState = true;
		<% /* PSC 22/09/2005 MAR32 */ %>				
		if( m_taskXML.GetAttribute("NOTAPPLICABLEFLAG") == "0" ||
		    m_taskXML.GetAttribute("TASKSTATUS") != "10" &&  //"Not Actioned"
		    m_taskXML.GetAttribute("TASKSTATUS") != "20" &&  //"Pending"
		    !bTASInFailed)  
				bNotApplicableButtonState = true;
		if( bMultiInput ||
		    ((m_taskXML.GetAttribute("CONTACTTYPE") == null || m_taskXML.GetAttribute("CONTACTTYPE") == "") ||
		     m_taskXML.GetAttribute("TASKSTATUS") != "20" ) )  //"Pending"
				bContactDetailsButtonState = true;
		if( bMultiInput ) bMemoButtonState = true;
		
	}
	frmScreen.btnAction.disabled = bActionButtonState;
	frmScreen.btnChaseUp.disabled = bChaseUpButtonState;
	frmScreen.btnDetails.disabled = bDetailsButtonState;
	frmScreen.btnNotApplicable.disabled = bNotApplicableButtonState;
	frmScreen.btnReprint.disabled = bReprintButtonState;
	frmScreen.btnContactDetails.disabled = bContactDetailsButtonState;
	frmScreen.btnMemo.disabled = bMemoButtonState;
	if(bMixIncludesCompleted)  alert("Selected tasks include ones already Completed. These cannot be processed further.");
	
	m_iActionCounter = 0;	<% /* BM0262 */ %>
}

function disableProcessButtons()
{
	frmScreen.btnAction.disabled = true;
	frmScreen.btnChaseUp.disabled = true;
	frmScreen.btnDetails.disabled = true;
	frmScreen.btnNotApplicable.disabled = true;
	frmScreen.btnReprint.disabled = true;
	frmScreen.btnContactDetails.disabled = true;
	frmScreen.btnMemo.disabled = true;
}

function updateList()
{
	stageXML = null;
	GetStageInfo();

	<% /* Sort by Status */ %>
	frmScreen.cboOrderBy.selectedIndex = 0;
	frmScreen.cboOrderBy.onchange();
	
	<% /* Check if all tasks completed */ %>
	if(!m_blnReadOnly)
		QueryStageComplete();
}

function initTaskDetail()
{
	if (scScrollTable.getRowSelectedIndex() == null) 
	{
		tblTDName.innerText="";
		tblTDActionedDate.innerText="";
		tblTDDueDate.innerText="";
		tblTDOwner.innerText="";
		tblTDStatus.innerText="";
		tblTDUpdatedBy.innerText="";
	}
	else if(scScrollTable.getArrayofRowsSelected().length > 1)
	{
		// Multiple tasks selected
		tblTDName.innerText="Multiple Selection";
		tblTDActionedDate.innerText="";
		tblTDDueDate.innerText="";
		tblTDOwner.innerText="";
		tblTDStatus.innerText="";
		tblTDUpdatedBy.innerText="";
	}
	else
	{
		SetActiveTask(scScrollTable.getRowSelectedIndex());
		
		<% /* BMIDS857  Allow for a longer task name (up to 50 characters) */ %>
		<% /* AW   14/09/06 EP1150 */ %>
		if( stageXML.GetAttribute("EDITABLETASKIND") == "1" )
		{
			scScreenFunctions.SizeTextToField(tblTDName,stageXML.GetAttribute("CASETASKNAME"));
		}
		else
		{
			scScreenFunctions.SizeTextToField(tblTDName,stageXML.GetAttribute("TASKNAME"));
		}
		
		tblTDStatus.innerText = GetComboText(stageXML.GetAttribute("TASKSTATUS"));
		tblTDDueDate.innerText = stageXML.GetAttribute("TASKDUEDATEANDTIME");
		tblTDActionedDate.innerText = stageXML.GetAttribute("TASKSTATUSSETDATETIME");
		tblTDOwner.innerText = stageXML.GetAttribute("OWNINGUSERID");
		tblTDUpdatedBy.innerText = stageXML.GetAttribute("TASKSTATUSSETBYUSERID");
	}
}
<% // BMIDS579  Display hourglass cursor while Action processing is taking place %>

function frmScreen.btnAction.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>

	<% /* BM0262 */ %>
	m_iActionCounter++;
	<% /* Remove focus from Action so that Enter cannot be pressed multiple times */ %>
	frmScreen.btnAction.blur();
	
	if (m_iActionCounter > 1)
		return false;
		
	frmScreen.btnAction.style.cursor = "wait";
	disableProcessButtons();
	<% /* BM0262 End */ %>

	<% /* Call the Action Processing after a timeout to allow the cursor time to change. */ %>
	window.setTimeout(ActionProcessing, 0);
}

function ActionProcessing()
{
<%	/* 
	First, if the selected task has an interface attribute then we need to call the operation
	on the OmTmBO component

	Next, if the selected task has an associated Output Document and is currently at status
	'Not Actioned' then print the document and set to 'Pending'.

	Next, if the selected task has an associated input process (screen) then route to the
	screen.
	
	Otherwise, just set the task status as 'Complete'.
	We can safely check just the first CASETASK element for both OUTPUTDOCUMENT and INPUTPROCESS
	as the Action button will not be enabled if multiple tasks have been selected where some
	of those tasks have print or screen functions.
	
	*/
%>	

	m_taskXML.SelectTag(null, "REQUEST");
	m_taskXML.CreateTagList("CASETASK");
	//GD SYS2560 06/08/01
	var bUpdateCaseTask;
	var bInputProcess; //variable to store if an input process exists
	
	<% /* MAR1300 GHun Validate task process authority */ %>
	for(var iCount = 0; iCount < m_taskXML.ActiveTagList.length; iCount++)
	{
		var tempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		tempXML.CreateRequestTag(window, "ValidateProcessTaskAuthority");
		tempXML.ActiveTag.appendChild(m_taskXML.GetTagListItem(iCount).cloneNode(true));
		
		tempXML.RunASP(document, "MsgTMBO.asp");
		tempXML.SelectTag(null, "RESPONSE");
		if (!tempXML.IsResponseOK())
		{
			frmScreen.btnAction.style.cursor = "hand";
			return false;
		}
	}
	<% /* MAR1300 End */ %>
	
	if(m_taskXML.ActiveTagList.length == "1")
	{
		//GD SYS2560 06/08/01
		bUpdateCaseTask = true;
		
		m_taskXML.SelectTagListItem(0);  // select the first CASETASK item!
		
		<% /* MAR7 GHun */ %>
		sInterfaceName =  m_taskXML.GetAttribute("INTERFACE");
		sDocumentID = m_taskXML.GetAttribute("OUTPUTDOCUMENT");
		sInputProcess = m_taskXML.GetAttribute("INPUTPROCESS");
		<% /* MAR7 End */ %>
		
		var sTaskId = m_taskXML.GetAttribute("TASKID");

		<% /* MAR1658 GHun */ %>
		if (sTaskId == m_sTmCOPRemodelTaskId)
		{
			scScreenFunctions.SetContextParameter(window, "idRemodelCOP", "1");
		}
		<% /* MAR1658 End */ %>

		// APS SYS1920
		if (sInterfaceName != "" && sInterfaceName != null)
		{
			
			//GD SYS2560 06/08/01
			bUpdateCaseTask = false;
			
			if (makeInterfaceCall(sInterfaceName, m_taskXML.ActiveTag, sDocumentID))
			{			
				 <% /* BMIDS670 - Check for data freeze 
				if ((sTaskId.toUpperCase() == GetIssueOfferParameter()) || 
					(sTaskId.toUpperCase() == GetReIssueOfferParameter()))
				{
					scScreenFunctions.SetContextParameter(window, "idFreezeDataIndicator", "1");
				}
				*/ %>
				CheckFreezeUnFreeze(m_sApplicationNumber);
				// AQR SYS4530 - Check for Admin sys account no & put it into side bar
				if (sInterfaceName == "GetNewNumbers")
				{
					var sApplicationNo = scScreenFunctions.GetContextParameter(window,"idApplicationNumber", "");
					scScreenFunctions.SetOtherSystemACNoInContext(sApplicationNo);
				}			
				if (m_taskXML.GetAttribute("INPUTPROCESS") != "")
				{
					RouteToInputProcess(m_taskXML.GetAttribute("INPUTPROCESS"));
				}
				else
				{
					<% /* BM0493 MDC 03/04/2003 */ %>
					updateList();
					<% /* BM0493 MDC 03/04/2003 - End */ %>
//					frmToTM030.submit();
				}
			}
		} 	
		else if (sDocumentID != "" &&
		   m_taskXML.GetAttribute("TASKSTATUS") == "10")
		{
			//GD SYS2560 06/08/01
			bUpdateCaseTask = false;
			bInputProcess = PrintTask();
			//DPF 6/9/02 - APWP3 - If there's an input process re-direct to it not TM030
			if (bInputProcess != true)
			{
				<% /* BM0493 MDC 15/04/2003 */ %>
				updateList();
				<% /* BM0493 MDC 15/04/2003 - End */ %>
				
				//AW BM0201 19/12/02
				//bUpdateCaseTask = true; //AW BM0137 04/12/02
//				frmToTM030.submit(); // PSC 03/12/01 SYS3288
			}
		}
		else if(m_taskXML.GetAttribute("INPUTPROCESS") != "" &&
		        m_taskXML.GetAttribute("PACKCONTROLNUMBER") == "") //JD MAR1040
		{
			//GD SYS2560 06/08/01
			bUpdateCaseTask = false;			
			RouteToInputProcess(m_taskXML.GetAttribute("INPUTPROCESS"));
		}
<% /* MAR211 TW */ %>

		// Fullfilment packs
		else if (m_taskXML.GetAttribute("PACKCONTROLNUMBER") != "")
		{
			bUpdateCaseTask = false;
	
			var xmlRequestObj = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var xmlRequestDocument = xmlRequestObj.XMLDocument;
	
			var xmlRequest = xmlRequestDocument.createElement("REQUEST");
			xmlRequestDocument.appendChild(xmlRequest);

			xmlRequest.setAttribute("OPERATION", "SENDPACK");
			xmlRequest.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
			xmlRequest.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
			xmlRequest.setAttribute("UNITID", m_UnitId);
			xmlRequest.setAttribute("USERID", m_UserId);
			xmlRequest.setAttribute("PACKCONTROLNUMBER", m_taskXML.GetAttribute("PACKCONTROLNUMBER"));
			xmlRequest.setAttribute("MACHINEID", m_MachineId);
			xmlRequest.setAttribute("CHANNELID", m_DistributionChannelId);
			xmlRequest.setAttribute("USERAUTHORITYLEVEL", scScreenFunctions.GetContextParameter(window, "idAccessType", "0"));

			xmlRequestObj.RunASP(document, "omPackRequest.asp");

			xmlRequestObj.SelectTag(null, "RESPONSE");
			if (!xmlRequestObj.IsResponseOK()) 
			{
				frmScreen.btnAction.style.cursor = "hand";
				return false;
			}
	<% /* MAR578 TW 16/11/2005 */ %>
			bUpdateCaseTask = true;
	<% /* MAR578 TW 16/11/2005 End */ %>
			
		}

<% /* MAR211 End */ %>
		//GD SYS2560 Update the single task 
		//Update case task if it hasn't been done in the above code
		if (bUpdateCaseTask == true)
		{
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTag(window , "UpdateCaseTask");
			var node = m_taskXML.SelectTag(null, "CASETASK");
			var newNode = node.cloneNode(true);
			if(sInputProcess == "")  //JD MAR1040
				newNode.setAttribute("TASKSTATUS", "40"); // 40 = Complete
			else
				newNode.setAttribute("TASKSTATUS", "20"); // 20 = pending
			XML.ActiveTag.appendChild(newNode);
			UpdateCaseTask(XML.XMLDocument.xml);
			//DRC BMIDS670 Do we need to Freeze or UnFreeze the Data ?
			CheckFreezeUnFreeze(m_sApplicationNumber);
		}
		updateList();
	}
	else //Multiple Tasks
	{	
		//BG 30/03/01 find any nodes which either have no outputdocument attribute, or one with no value set
		m_taskXML.CreateTagList("CASETASK[$not$ @OUTPUTDOCUMENT or @OUTPUTDOCUMENT='']");
		if(m_taskXML.ActiveTagList.length == 0)
		{
			PrintTask();
//			frmToTM030.submit();
			//GD SYS2560 06/08/01 EXIT PROCESSING
		}
		else //GD SYS2560
		{
			m_taskXML.CreateTagList("CASETASK[$not$(@INPUTPROCESS='') or $not$(@OUTPUTDOCUMENT='') or $not$(@INTERFACE='')]");
			if(m_taskXML.ActiveTagList.length == 0)
			{
				//Perform UpdateCaseTask on each task if no Inputprocess,OutputDocument or Interface attribs found
				m_taskXML.CreateTagList("CASETASK");
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XML.CreateRequestTag(window , "UpdateCaseTask");
				for(var iCount = 0; iCount < m_taskXML.ActiveTagList.length; iCount++)
				{
					var node = m_taskXML.GetTagListItem(iCount);
					var newNode = node.cloneNode(true);
					newNode.setAttribute("TASKSTATUS", "40"); // 40 = Complete
					XML.ActiveTag.appendChild(newNode);
				}
				UpdateCaseTask(XML.XMLDocument.xml);
				//DRC BMIDS670 - Does any task affect the Freeze Unfreeze Data parameter?
				CheckFreezeUnFreeze(m_sApplicationNumber);	
				
			} else
			{
				alert("Cannot process tasks of different types.");
//				frmToTM030.submit();
			}
		}
	}
	
	<% /* BM0262 */ %>
	frmScreen.btnAction.style.cursor = "hand";
	<% /* BM0262 End */ %>

}

function frmScreen.btnDetails.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>
	m_taskXML.SelectTag(null, "REQUEST");
	m_taskXML.CreateTagList("CASETASK");
	
	m_taskXML.SelectTagListItem(0);
	if(m_taskXML.GetAttribute("INPUTPROCESS") != "")
	{
		RouteToInputProcess(m_taskXML.GetAttribute("INPUTPROCESS"));
	}	
}

function frmScreen.btnReprint.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>	
	m_taskXML.SelectTag(null, "REQUEST");
	m_taskXML.CreateTagList("CASETASK");
	
	m_taskXML.SelectTagListItem(0);
	if(m_taskXML.GetAttribute("OUTPUTDOCUMENT") != "")
	{
		<% /* BM0262 Prevent Reprint from being pressed multiple times */ %>
		m_iActionCounter++;
		<% /* Remove focus from Reprint button so that Enter cannot be pressed multiple times */ %>
		frmScreen.btnReprint.blur();
		
		if (m_iActionCounter > 1)
			return false;
			
		frmScreen.btnReprint.style.cursor = "wait";
		frmScreen.btnReprint.disabled = true;
		<% /* BM0262 End */ %>

		PrintTask();
		
		<% /* BM0262 */ %>
		frmScreen.btnReprint.style.cursor = "hand";
		<% /* BM0262 End */ %>
	}
}

function frmScreen.btnChaseUp.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window , "ChaseUpTask");
	XML.CreateActiveTag("CASETASK");
	m_taskXML.SelectTag(null, "CASETASK");
	XML.SetAttribute("SOURCEAPPLICATION", m_taskXML.GetAttribute("SOURCEAPPLICATION"));
	XML.SetAttribute("CASEID", m_taskXML.GetAttribute("CASEID"));
	XML.SetAttribute("ACTIVITYID", m_taskXML.GetAttribute("ACTIVITYID"));
	XML.SetAttribute("STAGEID", m_taskXML.GetAttribute("STAGEID"));
	XML.SetAttribute("TASKID", m_taskXML.GetAttribute("TASKID"));
	XML.SetAttribute("CHASINGTASK", m_taskXML.GetAttribute("CHASINGTASK"));
	XML.SetAttribute("CONTACTTYPE", m_taskXML.GetAttribute("CONTACTTYPE"));
	//BG 04/11/01 SYS333 Added TASKINSTANCE and CASESTAGESEQUENCENO attributes
	XML.SetAttribute("TASKINSTANCE", m_taskXML.GetAttribute("TASKINSTANCE"));
	XML.SetAttribute("CASESTAGESEQUENCENO", m_taskXML.GetAttribute("CASESTAGESEQUENCENO"));
	
	// 	XML.RunASP(document, "OmigaTMBO.asp");
	// Added by automated update TW 01 Oct 2002 SYS5115
	<%//GD BM0574 START
	//switch (ScreenRules())
	//{
		//case 1: // Warning
		//case 0: // OK %>
	XML.RunASP(document, "OmigaTMBO.asp");
	<%		//break;
		//default: // Error
			//XML.SetErrorResponse();
	//}GD BM0574 END %>
	if(XML.IsResponseOK())
		updateList();
}

<% // BMIDS579  Display hourglass cursor while Not Applicable processing is taking place %>

function frmScreen.btnNotApplicable.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>	
	
	<% /* Remove focus from Not Applicable so that Enter cannot be pressed multiple times */ %>
	frmScreen.btnNotApplicable.blur();
	
	frmScreen.btnNotApplicable.style.cursor = "wait";
	disableProcessButtons();

	<% /* Call the Not Applicable Processing after a timeout to allow the cursor time to change. */ %>
	window.setTimeout("NotApplicableProcessing()", 0)
}

function NotApplicableProcessing()
{
	m_taskXML.SelectTag(null, "REQUEST");
	m_taskXML.CreateTagList("CASETASK");
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window , "UpdateCaseTask");
	for(var iCount = 0; iCount < m_taskXML.ActiveTagList.length; iCount++)
	{
		var node = m_taskXML.GetTagListItem(iCount);
		if(node.getAttribute("NOTAPPLICABLEFLAG") == "1")
		{
			var newNode = node.cloneNode(true);
			newNode.setAttribute("TASKSTATUS", "30"); // 30 = Not Applicable
			XML.ActiveTag.appendChild(newNode);
		}
	}
	UpdateCaseTask(XML.XMLDocument.xml);
	//alert("At Offer Stage is " + IsApplicationAtOfferStage(m_sApplicationNumber))
	
	//UpdateCaseTask(XML);
	//GD BMIDS00037
	//DRC BMIDS670 - See if any of the Tasks made non Applicable affect the freeze data parameter
	
	CheckFreezeUnFreeze(m_sApplicationNumber);
	
	//JD BMIDS789 reset cursor to hand
	frmScreen.btnNotApplicable.style.cursor = "hand";
	
	
}

function frmScreen.btnMemo.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>	
	
	<% /* BM0493 MDC 04/04/2003 */ %>
	// tasks XML is still availlable in the context under idTaskXML
//	scScreenFunctions.SetContextParameter(window,"idReturnScreenId","TM030");
//	frmToTM036.submit();
	
	var sReturn = null;
	var ArrayArguments = new Array();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ArrayArguments[0] = m_sReadOnly;
	ArrayArguments[1] = m_taskXML.XMLDocument.xml;
	ArrayArguments[2] = XML.CreateRequestAttributeArray(window);
	<% /* BS BM0521 10/06/03 */ %>
	ArrayArguments[3] = m_sReadOnlyAsCaseLocked;
	ArrayArguments[4] = m_sProcessInd;
	<% /* BS BM0521 End 10/06/03 */ %>
	
	<% /*AW 19/09/06	EP1150 */ %>
	var bProgressTaskMode = false;
	ArrayArguments[5] = m_sCurrStageId;
	ArrayArguments[6] = bProgressTaskMode;
		
	sReturn = scScreenFunctions.DisplayPopup(window, document, "TM036P.asp", ArrayArguments, 540, 500);
	//updateList();
	<% /* BM0493 MDC 04/04/2003 - End */ %>
}

function frmScreen.btnContactDetails.onclick()
{
	<% // GD BM0574 START %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	<% // GD BM0574 END %>
	var sReturn = null;
	var ArrayArguments = new Array();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ArrayArguments[0] = m_sReadOnly;
	ArrayArguments[1] = m_taskXML.XMLDocument.xml;
	ArrayArguments[2] = XML.CreateRequestAttributeArray(window);
	sReturn = scScreenFunctions.DisplayPopup(window, document, "TM034.asp", ArrayArguments, 460, 280);
	updateList();
}

function UpdateCaseTask(sXMLString)
{	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(sXMLString);
	XML.SelectTag(null, "REQUEST");
	var sUSERID = XML.GetAttribute("USERID");
	XML.SelectTag(null, "CASETASK");
	XML.SetAttribute("TASKSTATUSSETBYUSERID", sUSERID);
	// 	XML.RunASP(document, "MsgTMBO.asp");	
	// Added by automated update TW 09 Oct 2002 SYS5115
	<%//GD BM0574 START
	//switch (ScreenRules())
		//{
		//case 1: // Warning
		//case 0: // OK %>
	XML.RunASP(document, "MsgTMBO.asp");	
			<%//break;
		//default: // Error
			//XML.SetErrorResponse();
		//} BM0574 END %>

	var ErrorTypes = new Array("NOTASKAUTHORITY");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{
		alert("User does not have the authority to update this task");
	}
	else if(ErrorReturn[0] == true)
	{
		//	only do if no error	
		updateList();
	}
}

function RouteToInputProcess(sInputProcess)
{
	// tasks XML is still available in the context under idTaskXML
	scScreenFunctions.SetContextParameter(window,"idReturnScreenId","TM030");

	// PSC 06/12/01 SYS3358 - Start
	switch(sInputProcess)
	{
		case "GN300" :
			scScreenFunctions.SetContextParameter(window, "idProcessContext", "CompletenessCheck");
			break;
		case "CM010" :
			scScreenFunctions.SetContextParameter(window, "idApplicationMode", "Cost Modelling");
			break;
		default: break;
	}
	frmInputProcess.action = sInputProcess + ".asp";
	
	<% /* BM0134 MDC 03/12/2002 */ %>
	// frmInputProcess.submit();
	<%//GD BM0574 START
	//switch (ScreenRules())
	//{
		//case 1: // Warning
		//case 0: // OK %>
	frmInputProcess.submit();
			<%//break;
		//default: // Error
	//} GD BM0574 END %>
	<% /* BM0134 MDC 03/12/2002 - End */ %>
	 
	// PSC 06/12/01 SYS3358 - End
}

//DPF 6/9/02 - APWP3 - Have amended PrintTask to include user choice over whether a print is done before
//either routing to the Input screen required or updating the task.  Printing part has been moved to 
//ProcessPrint function

function PrintTask()
{

	var bPrintReturn; //variable to store outcome of ProcessPrint
	var bInputProcess;
	
	bPrintReturn = false; //preset to successful
	bInputProcess = false;  //variable returned to specify whether the user should be routed to another screen
	
	m_taskXML.SelectTag(null, "REQUEST");
	m_taskXML.CreateTagList("CASETASK");
	if(m_taskXML.ActiveTagList.length > 0)
	{  
		for(var iCount = 0; iCount < m_taskXML.ActiveTagList.length; iCount++)
		{	
			//AW	19/12/02	BM0201
			
			if(m_taskXML.GetTagAttribute("CASETASK", "CONFIRMPRINTIND") == "1")
			{
				if(confirm("Click OK to Print the associated document")) //user interaction
				{	
					bPrintReturn = ProcessPrint(m_taskXML, iCount);	//call print function
					if (bPrintReturn == true) //if an error is returned break the loop
					{	break; }
				}
				else //user does not wish to print so check if we have an input process to route to
				{
					if(m_taskXML.GetTagAttribute("CASETASK", "INPUTPROCESS") == "")
					{
						<% /* BM0201 GHun */ %>
						//no input process so update the task status
						var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
						XML.CreateRequestTag(window , "UpdateCaseTask");
						var node = m_taskXML.SelectTag(null, "CASETASK");
						var newNode = node.cloneNode(true);
						newNode.setAttribute("TASKSTATUS", "40"); // 40 = Complete
						XML.ActiveTag.appendChild(newNode);
						UpdateCaseTask(XML.XMLDocument.xml);
						<% /* BM0201 GHun End */ %>
					}
					else //input process exists so route to the correct screen
					{
						RouteToInputProcess(m_taskXML.GetTagAttribute("CASETASK", "INPUTPROCESS"));
						bInputProcess = true;
					}			
				}						
			}			
			else
			//ConfirmPrintInd in XML is false so print as per normal
			{	
				<% /* BM0201 GHun not required
				if(m_taskXML.GetTagAttribute("CASETASK", "INPUTPROCESS") == "")
				{
					bUpdateCaseTask = true; //no input process so just update task
				}
				*/ %>
				
				bPrintReturn = ProcessPrint(m_taskXML, iCount);	//call print function			
				if (bPrintReturn == true) //if an error is returned break the loop
					{ break; }
			}
		}
	}
	return bInputProcess //return value to btnAction
}

<% /*
function IsApplicationAtOfferStage(sApplicationNumber)
{
	var sSuccess = -1;

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	// get the Activity ID from the global partameter table
	var sActivityId = scScreenFunctions.GetContextParameter(window, "idActivityId", "1");
	
	if (sActivityId == "") {
		sActivityId = m_sActivityId
		// sActivityId = XML.GetGlobalParameterAmount(document, "TMOmigaActivity");	
		
		//scScreenFunctions.SetContextParameter(window, "idActivityId", sActivityId);
		XML.ResetXMLDocument();
	}
	
	XML.CreateRequestTag(window, "ISAPPLICATIONATOFFER");
	XML.CreateActiveTag("CASESTAGE");
	
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("CASEID", sApplicationNumber);
	XML.SetAttribute("ACTIVITYID", sActivityId);
	
	// XML.RunASP(document, "OmigaTMBO.asp");
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "OmigaTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
	}
	
	//GD BMIDS00037 reset datafreeze indicator if appropriate.
	if (XML.IsResponseOK()) {
		var sDataFreezeIndicator = XML.GetTagAttribute("APPLICATION", "FREEZEDATAINDICATOR");
		//scScreenFunctions.SetContextParameter(window,"idFreezeDataIndicator", sDataFreezeIndicator);
		//alert("Data Freeze Indicator is " + sDataFreezeIndicator);
		sSuccess = sDataFreezeIndicator;
	}
	XML = null;

	return sSuccess; //-1 implies failed, else contains freeze indicator
}
*/ %>

function IsApplicationAtOfferStage(sApplicationNumber)
{
	var sSuccess = -1;

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	// get the Activity ID from the global partameter table
	var sActivityId = scScreenFunctions.GetContextParameter(window, "idActivityId", "1");
	
	if (sActivityId == "") 
	{
		<% /* BM0493 MDC 03/04/2003 */ %>
		sActivityId = m_sActivityId
		// sActivityId = XML.GetGlobalParameterAmount(document, "TMOmigaActivity");	
		<% /* BM0493 MDC 03/04/2003 - End */ %>

		//scScreenFunctions.SetContextParameter(window, "idActivityId", sActivityId);
		XML.ResetXMLDocument();
	}
	
	XML.CreateRequestTag(window, "ISAPPLICATIONATOFFER");
	XML.CreateActiveTag("CASESTAGE");
	
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("CASEID", sApplicationNumber);
	XML.SetAttribute("ACTIVITYID", sActivityId);
	
	<% /* MV - 08/04/2003 - BM0512 */ %>
	XML.RunASP(document, "OmigaTMBO.asp");
	
	//GD BMIDS00037 reset datafreeze indicator if appropriate.
	if (XML.IsResponseOK()) {
		var sDataFreezeIndicator = XML.GetTagAttribute("APPLICATION", "FREEZEDATAINDICATOR");
		//scScreenFunctions.SetContextParameter(window,"idFreezeDataIndicator", sDataFreezeIndicator);
		//alert("Data Freeze Indicator is " + sDataFreezeIndicator);
		sSuccess = sDataFreezeIndicator;
	}
	XML = null;
	
	return sSuccess; //-1 implies failed, else contains freeze indicator
}
function CheckFreezeUnFreeze(sApplicationNumber)
{
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	// get the Activity ID from the global partameter table
	var sActivityId = scScreenFunctions.GetContextParameter(window, "idActivityId", "1");
	
	if (sActivityId == "") 
	{
		sActivityId = m_sActivityId
		
		XML.ResetXMLDocument();
	}
	
	XML.CreateRequestTag(window, "FREEZEUNFREEZEAPPLICATION");
	XML.CreateActiveTag("CASESTAGE");
	
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("CASEID", sApplicationNumber);
	XML.SetAttribute("ACTIVITYID", sActivityId);
		
	XML.RunASP(document, "OmigaTMBO.asp");
	
	if (XML.IsResponseOK()) {
		var sDataFreezeIndicator = XML.GetTagAttribute("APPLICATION", "FREEZEDATAINDICATOR");
		//scScreenFunctions.SetContextParameter(window,"idFreezeDataIndicator", sDataFreezeIndicator);
		//alert("Data Freeze Indicator is " + sDataFreezeIndicator);
		if (sDataFreezeIndicator.length > 0)
		  {
			scScreenFunctions.SetContextParameter(window,"idFreezeDataIndicator", sDataFreezeIndicator);
		  // DRC BMIDS692
		    var sFreezeTaskID = ""
			if (sDataFreezeIndicator=="1") 
			  sFreezeTaskID = XML.GetTagAttribute("APPLICATION", "FREEZETASKID");
			if (ClientFreezeOverRide(sFreezeTaskID,sDataFreezeIndicator))
			 {
			   // alert("Data Freeze Over ride for " + sFreezeTaskID);
			 }              
		  // DRC BMIDS692 END		
		     
          }
	}
	XML = null;

}

function ProcessPrint(m_taskXML, iCount)
{
	var sCaseCustomerNumber = null;
	var sApplicationNumber = null;
	var bpopped = "0";
	var bPrintError = false;
	var bSuccess = null;	<% /* MAR7 GHun */ %>
	
	node = m_taskXML.GetTagListItem(iCount);
	<% /* MAR7 GHun
	AttribsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AttribsXML.CreateRequestTag(window , "GetPrintAttributes");
	AttribsXML.CreateActiveTag("FINDATTRIBUTES");			
	AttribsXML.SetAttribute("HOSTTEMPLATEID", node.getAttribute("OUTPUTDOCUMENT"));
									
	AttribsXML.RunASP(document, "PrintManager.asp");
			
	if(AttribsXML.IsResponseOK())
	*/ %>
	var sDocumentID = node.getAttribute("OUTPUTDOCUMENT")
	AttribsXML = getPrintDocumentAttributes(sDocumentID);
	if (AttribsXML != null )
	<% /* MAR7 End */ %>
	{					
		var sLocalPrinters = GetLocalPrinters();
		LocalPrintersXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		LocalPrintersXML.LoadXML(sLocalPrinters);
				
		if(AttribsXML.GetTagAttribute("ATTRIBUTES", "INACTIVEINDICATOR") == "1")
		{
			alert("The document template " + AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME") + " is inactive, please see your System Administrator");
			bPrintError = True	
		}
				
		//Check to see if there are any printer locations
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var sGroupList = new Array("PrinterDestination");
		if(XML.GetComboLists(document,sGroupList))
		{		
			XML.CreateTagList("LISTENTRY");
			if(XML.ActiveTagList.length == "0")
			{
				alert("There are currently no locations defined to print the document.");
			}	
		}

		// ik_bmids00730
		// if edit or view before printing an option, route to pm010
		if( (AttribsXML.GetTagAttribute("ATTRIBUTES", "EDITBEFOREPRINT") == 1) ||
			(AttribsXML.GetTagAttribute("ATTRIBUTES", "VIEWBEFOREPRINT") == 1) )
		{
			sCaseCustomerNumber = node.getAttribute("CUSTOMERIDENTIFIER");
			sApplicationNumber = node.getAttribute("CASEID");
			CallPM010(sCaseCustomerNumber,sApplicationNumber,m_taskXML.GetTagListItem(iCount));
			bpopped = "1";
		}
				
	<% /* MAR578 TW 16/11/2005 */ %>

//		var sPrinterType = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
		sPrinterType = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
		sPrinterValidationType = getPrinterType(window,sPrinterType);
//		var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	<% /* MAR578 TW 16/11/2005 End */ %>
		if(bpopped != "1")
		{
			if (TempXML.IsInComboValidationList(document,"PrinterDestination",sPrinterType,["L"]))
			{							
				LocalPrintersXML.CreateTagList("PRINTER[DEFAULTINDICATOR='1']");
							
				if(LocalPrintersXML.ActiveTagList.length == 0)
				{					
					alert("You do not have a default printer set on your PC.");
					bPrintError = True
				}							
				
				if (TempXML.IsInComboValidationXML(["R"]))
				{
					// Call PM010 Popup as choice is required
					// PSC 03/12/01 SYS3286						
					sCaseCustomerNumber = node.getAttribute("CUSTOMERIDENTIFIER");
					sApplicationNumber = node.getAttribute("CASEID");
					CallPM010(sCaseCustomerNumber,sApplicationNumber,m_taskXML.GetTagListItem(iCount));
					bpopped = "1";
				}						
			}	
		}
		// ik_bmids00730_ends

		//logic to know it's been in a popup
		if(bpopped != "1")
		{
			if(AttribsXML.GetTagAttribute("ATTRIBUTES", "DEFAULTCOPIES") == "" || AttribsXML.GetTagAttribute("ATTRIBUTES", "MAXCOPIES") != "")
			{	
				// PSC 03/12/01 SYS3286					
				sCaseCustomerNumber = node.getAttribute("CUSTOMERIDENTIFIER");
				sApplicationNumber = node.getAttribute("CASEID");						
				CallPM010(sCaseCustomerNumber, sApplicationNumber, m_taskXML.GetTagListItem(iCount));
				bpopped = "1";
			}
			
	<% /* MAR7 GHun */ %>
			if(bpopped != "1")	<% /* EP2_2026 GHun Only call PrintDocumentForTask when bpopped is not 1*/ %>
			{
				bPrintError = PrintDocumentForTask(node);
			}
		}
	}
	return bPrintError;
}

function CallPM010(sCaseCustomerNumber, sAppNo, CaseTaskXML)
{	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sReturn = null;
	var TMRequestArray = XML.CreateRequestAttributeArray(window);
	var sCustomerVersionNumber = "";
	var ArrayArguments = new Array();
	//var sCaseCustomerNumber = m_taskXML.GetAttribute("CUSTOMERNUMBER");
	if(sCaseCustomerNumber != "")
	{
		var iCount = 0;
		var sCustomerName = "";
		var sCustomerNumber = "";		
		var aCustomerNameArray = new Array();
		var aCustomerNumberArray = new Array();
		var aCustomerVersionNumberArray = new Array();			
		var sCustomerVersionNumber = "";
			
		
		for (iCount = 1; iCount <= 5; iCount++)
		{								
			sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
			sCustomerName = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerName" + iCount,null);
			sCustomerVersionNumber = m_FW030ScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);
			aCustomerNameArray[iCount-1] = sCustomerName;
			aCustomerNumberArray[iCount-1] = sCustomerNumber;
			aCustomerVersionNumberArray[iCount-1] = sCustomerVersionNumber;	
			if (sCustomerNumber == sCaseCustomerNumber) 
			{					
				sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);					
			}
		}			
	}
	
	<% /* MO - 07/11/2002 - BMIDS00101 - Uncommented args 2,3,4, and 7	*/ %>
	ArrayArguments[0] = sAppNo;
	ArrayArguments[1] = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	ArrayArguments[2] = scScreenFunctions.GetContextParameter(window,"idUserId",null);
	ArrayArguments[3] = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	ArrayArguments[4] = scScreenFunctions.GetContextParameter(window,"idStageid",null);
	ArrayArguments[5] = sCaseCustomerNumber;
	ArrayArguments[6] = sCustomerVersionNumber;
	ArrayArguments[7] = scScreenFunctions.GetContextParameter(window,"idRole",null);
	ArrayArguments[8] = TMRequestArray;
	ArrayArguments[9] = aCustomerNameArray
	ArrayArguments[10] = aCustomerNumberArray
	ArrayArguments[11] = aCustomerVersionNumberArray
	ArrayArguments[12] = AttribsXML.XMLDocument.xml
	ArrayArguments[13] = CaseTaskXML.xml
	ArrayArguments[14] = LocalPrintersXML.XMLDocument.xml
	ArrayArguments[15] = CaseTaskXML.getAttribute("CONTEXT");	<% /* EP2_2026 GHun */ %>
	
	<% /* MAR7 GHun */ %>
	ArrayArguments[16] = XML.GetGlobalParameterString(document,"EmailAdministrator",null);
	ArrayArguments[17] = scScreenFunctions.GetContextParameter(window,"idMachineId",null);
	ArrayArguments[18] = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
	ArrayArguments[19] = scScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null);
	<% /* MAR7 End */ %>

	// ik_bmids00730
	nWidth = 540;
	nHeight = 580;
						
	sReturn = scScreenFunctions.DisplayPopup(window, document, "PM010.asp", ArrayArguments, nWidth, nHeight);

	if(sReturn != null)
	{
		m_taskXML.SelectTag(null, "REQUEST");
		m_taskXML.CreateTagList("CASETASK");
	
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window , "UpdateCaseTask");
		for(var iCount = 0; iCount < m_taskXML.ActiveTagList.length; iCount++)
		{
			var node = m_taskXML.GetTagListItem(iCount);
			<% /* BM0201 GHun InputProcess could also be null */ %> 
			var sInputProc = node.getAttribute("INPUTPROCESS");		
			if((sInputProc == null) || (sInputProc == ""))
			{
			<% /* BM0201 GHun End */ %>
				var newNode = node.cloneNode(true);
				newNode.setAttribute("TASKSTATUS", "40");// 40 = Complete 
				XML.ActiveTag.appendChild(newNode);
			}
			else
			{
				var newNode = node.cloneNode(true);
				newNode.setAttribute("TASKSTATUS", "20"); // 30 = Pending
				XML.ActiveTag.appendChild(newNode);
			}
			
		}
		UpdateCaseTask(XML.XMLDocument.xml);
	}
}

function makeInterfaceCall(sOperation, xmlCaseTask, sDocumentID)	<% /* MAR7 GHun */ %>
{
	<% /* PJO 09/03/2006 MAR1359 Add progress message */ %>
	scScreenFunctions.progressOn("Please Wait ... Processing Task", 400);
	
	var blnSuccess = false;
	var PrintXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	<% /* MAR7 GHun */ %>
	var ReqTag;
		
	<% /* MAR7 GHun */ %>
	ReqTag = PrintXML.CreateRequestTag(window,sOperation);
	//BMIDS682 Are we running a credit check
	if (sOperation == "RunCreditCheck")
	{
		PrintXML.SetAttribute("CREDITCHECKINVOKEDFROMACTIONBUTTON", "TRUE");
		PrintXML.SetAttribute("ADDRESSTARGETREQ", "FALSE");
	}	
	
	<% /* MAR750 Set TASKMANAGER attribute for Risk Assessment */ %>
	if (sOperation == "RunRiskAssessment")
	{
		PrintXML.SetAttribute("TASKMANAGER", "T");
	}	
	
	PrintXML.ActiveTag.appendChild(xmlCaseTask.cloneNode(true));
	PrintXML.CreateActiveTag("APPLICATION");
	PrintXML.SetAttribute("APPLICATIONNUMBER", scScreenFunctions.GetContextParameter(window, "idApplicationNumber", ""));
	PrintXML.SetAttribute("APPLICATIONFACTFINDNUMBER", scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", ""));
	//SYS3268 BG Added APPLICATIONPRIORITY
	PrintXML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window, "idApplicationPriority", ""));
	// SYS3295 - add print data	if there's a doc in the task
	
	<% /* PSC 27/10/2005 MAR300 - Start */ %>
    var iCount = 0;
	var sCustomerVersionNumber = "";
	var sCustomerNumber = "";
	var sOtherSystemCustomerNumber = "";
	for (iCount = 1; iCount <= 5; iCount++)
	{								
		sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
		if (sCustomerNumber != "")
		{	
			sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);	
			sOtherSystemCustomerNumber = scScreenFunctions.GetContextParameter(window,"idOtherSystemCustomerNumber" + iCount,null);	
			PrintXML.SelectTag(null, "APPLICATION");
			PrintXML.CreateActiveTag("CUSTOMER");
			PrintXML.SetAttribute("CUSTOMERNUMBER", sCustomerNumber);
			PrintXML.SetAttribute("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
			PrintXML.SetAttribute("OTHERSYSTEMCUSTOMERNUMBER", sOtherSystemCustomerNumber);
		}
	}
	<% /* PSC 27/10/2005 MAR300 - End */ %>

    if (sDocumentID != "") 
    <% /* MAR7 End */ %>
	{
		var sLocalPrinters = GetLocalPrinters();
		LocalPrintersXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		LocalPrintersXML.LoadXML(sLocalPrinters);
		LocalPrintersXML.SelectTag(null, "RESPONSE");	
		
		var sPrinter = LocalPrintersXML.GetTagText("PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
									
		if(sPrinter == "")
		{	
			scScreenFunctions.progressOff();
			alert("You do not have a default printer set on your PC.");			
		}		
		else
		{			
			PrintXML.ActiveTag = ReqTag;
			PrintXML.CreateActiveTag("PRINTER");
			PrintXML.SetAttribute("PRINTERNAME", sPrinter);
			PrintXML.SetAttribute("DEFAULTIND", "1");
			PrintXML.SetAttribute("COMPRESSIONMETHOD", ""); <% /* MAR796 GHun turn off compression */ %>
		}	
	}
	
	<% /* MAR7 GHun */ %>
	PrintXML.RunASP(document, "OmigaTMBO.asp");

	scScreenFunctions.progressOff();

	<% /* JD BMIDS841 Check for error code CREDITCHECKOKIMPORTFAILED*/ %>
	var ErrorTypes = new Array("NOTIMPLEMENTED", "CREDITCHECKOKIMPORTFAILED","LEGALREPINACTIVE");
	var ErrorReturn = PrintXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{				
		alert("Your task has been set with an incorrect Interface name. Please contact your System Administrator.");
		return(false);
	}
	<% /* PSC 27/10/2005 MAR300 - End */ %>
	else if (sOperation=="IssueOffer" && ErrorReturn[0] == true && ErrorReturn[3] == true)
		return(false);
	else if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[1]))
	{
		<% /* MAR1300 GHun */ %>
		if (sOperation == "SetChangeOfProperty")
		{
			scScreenFunctions.SetContextParameter(window, "idPropertyChange", "1");
		}
		<% /* MAR1300 End */ %>
		
		<% /* JD BMIDS841 If error CREDITCHECKOKIMPORTFAILED show error message */ %>
		if(ErrorReturn[1] == ErrorTypes[1])
		{
			var TagRESPONSE = PrintXML.SelectTag(null,"RESPONSE");
			if(PrintXML.SelectTag(TagRESPONSE,"ERROR"))
			{
				var sErrorMessage = PrintXML.GetTagText("DESCRIPTION");
				var sErrorSource = PrintXML.GetTagText("SOURCE");
				var sVersion = PrintXML.GetTagText("VERSION");
				alert(sErrorMessage + "\nSource: " + sErrorSource + "\nVersion: " + sVersion);
				return(false);
			}
		}
						
		<% /* BMIDS682 Are we AddressTargeting*/ %>
		PrintXML.SelectTag(null, "RESPONSE");
		PrintXML.SelectTag(PrintXML.ActiveTag, "TARGETINGDATA"); //MAR245 Revised XML structure in MARS
		
		<% /* PSC 27/10/2005 MAR300 - Start */ %>		
		var sIsAddrTarget = "";
		if (PrintXML.ActiveTag != null)
			sIsAddrTarget = PrintXML.GetTagText("ADDRESSTARGETING");
		<% /* PSC 27/10/2005 MAR300 - End */ %>		
		
		<% /* MV 03/11/2005 MAR402 */ %>		
		PrintXML.SelectTag(null, "RESPONSE");
		
		if (sIsAddrTarget == "YES")
		{
			//BMIDS730 Address Targeting
			// tasks XML is still available in the context under idTaskXML
			scScreenFunctions.SetContextParameter(window,"idReturnScreenId","TM030");
			<% /* MV 03/11/2005 MAR402
			PrintXML.ActiveTag = null;
			PrintXML.SelectTag(null, "RESPONSE"); */%>
			scScreenFunctions.SetContextParameter(window,"idAddressTarget", PrintXML.ActiveTag.xml);
			frmToRA040.submit();
			return(false);
		}
		else
		{
			//BMIDS730 Not Address Targeting
			<% /* MAR7 GHun */ %>
			if (CallAxwordClass(PrintXML))
			{
				<% /* If there is a routing screen then do not update the task */%>
				<% /* PSC 27/10/2005 MAR300 */ %>
				if (sInputProcess == "" && sOperation != "RunAutoApplicationExpiry") 
				{
					var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					XML.CreateRequestTag(window , "UpdateCaseTask");
					XML.ActiveTag.appendChild(xmlCaseTask);
					XML.SelectTag(null, "CASETASK");
					XML.ActiveTag.setAttribute("TASKSTATUS", "40"); 
					UpdateCaseTask(XML.XMLDocument.xml);
					return(true); 
				}
			}
			<% /* MAR7 End */ %>
			
			return(true);
		}
	}	
	else if (ErrorReturn[1] == ErrorTypes[2])
	{
		var TagRESPONSE = PrintXML.SelectTag(null,"RESPONSE");
		if(PrintXML.SelectTag(TagRESPONSE,"ERROR"))
		{
			var sErrorMessage = PrintXML.GetTagText("DESCRIPTION");
			var sErrorSource = PrintXML.GetTagText("SOURCE");
			var sVersion = PrintXML.GetTagText("VERSION");
			alert(sErrorMessage);
			return(false);
		}
	}
}

function GetAllGlobalParameters()
{
	m_ParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ReqTag = m_ParamXML.CreateRequestTag(window,"");		
	var ListTag = m_ParamXML.CreateActiveTag("GLOBALPARAMETERLIST");
	m_ParamXML.CreateActiveTag("GLOBALPARAMETER");
	m_ParamXML.CreateTag("NAME", "DeclinedStageValue");
	m_ParamXML.ActiveTag = ListTag;
	m_ParamXML.CreateActiveTag("GLOBALPARAMETER");
	m_ParamXML.CreateTag("NAME", "CancelledStageValue");
	m_ParamXML.ActiveTag = ListTag;
	m_ParamXML.CreateActiveTag("GLOBALPARAMETER");
	m_ParamXML.CreateTag("NAME", "TMCompletionsStageId");
	m_ParamXML.ActiveTag = ListTag;
	m_ParamXML.CreateActiveTag("GLOBALPARAMETER");
	m_ParamXML.CreateTag("NAME", "TMReIssueOffer");
	m_ParamXML.ActiveTag = ListTag;
	m_ParamXML.CreateActiveTag("GLOBALPARAMETER");
	m_ParamXML.CreateTag("NAME", "TMOmigaActivity");
	
	<% /* PSC 18/01/2006 MAR1081 - Start */ %>
	m_ParamXML.ActiveTag = ListTag;
	m_ParamXML.CreateActiveTag("GLOBALPARAMETER");
	m_ParamXML.CreateTag("NAME", "TMTOEEndStageID");
	m_ParamXML.ActiveTag = ListTag;
	m_ParamXML.CreateActiveTag("GLOBALPARAMETER");
	m_ParamXML.CreateTag("NAME", "TMPSWStageID");
	<% /* PSC 18/01/2006 MAR1081 - End */ %>

	<% /* MAR1658 GHun */ %>
	m_ParamXML.ActiveTag = ListTag;
	m_ParamXML.CreateActiveTag("GLOBALPARAMETER");
	m_ParamXML.CreateTag("NAME", "TMChgOfPropRemodelQuoteTaskID");
	<% /* MAR1658 End */ %>
	
	<%//GD BM0574 START
	//switch (ScreenRules())
	//{
		//case 1: // Warning
		//case 0: // OK %>
	m_ParamXML.RunASP(document, "GetCurrentParameterListEx.asp");
			<%//break;
		//default: // Error
			//m_ParamXML.SetErrorResponse();
	//} GD BM0574 END %>

	if (m_ParamXML.IsResponseOK())
	{
		// Get param values	
		m_sDeclineStageId = TM030GetGlobalParameterValue("DeclinedStageValue", "AMOUNT");
		m_sCancelStageId = TM030GetGlobalParameterValue("CancelledStageValue", "AMOUNT");
		m_sCompStageId = TM030GetGlobalParameterValue("TMCompletionsStageId", "STRING");
		m_sTMReIssueOffer = TM030GetGlobalParameterValue("TMReIssueOffer", "STRING");
		m_sActivityId = TM030GetGlobalParameterValue("TMOmigaActivity", "STRING");
		
		<% /* PSC 18/01/2006 MAR1081 - Start */ %>
		m_sTOEEndStageId = TM030GetGlobalParameterValue("TMTOEEndStageID", "STRING");	
		m_sPSWStageId = TM030GetGlobalParameterValue("TMPSWStageID", "STRING");	
		<% /* PSC 18/01/2006 MAR1081 - End */ %>
		
		m_sTmCOPRemodelTaskId = TM030GetGlobalParameterValue("TMChgOfPropRemodelQuoteTaskID", "STRING") <% /* MAR1658 GHun */ %>
	}	
}

function TM030GetGlobalParameterValue(sParameterName, sParameterType)
{
	var sRet = "";
	
	<% /* Reset the ActiveTag so that the whole document is searched */ %>
	m_ParamXML.ActiveTag = null;
	
	<% /* Find the parameter matching the name and type provided */ %>
	m_ParamXML.CreateTagList("GLOBALPARAMETER[NAME='" + sParameterName + "']");

	if(m_ParamXML.SelectTagListItem(0)) 
		sRet = m_ParamXML.GetTagText(sParameterType);

	return sRet
}

<% /* BM0262 Make sure double clicks are ignored */ %>
function frmScreen.btnAction.ondblclick()
{
	return false;
}

function frmScreen.btnReprint.ondblclick()
{
	return false;
}
<% /* BM0262 End */ %>

<% /* BM0262 Function converted from VBScript to JScript to improve performance */ %>
function GetLocalPrinters()
{
	var strOut = "";
	var objOmPC = new ActiveXObject("omPC.PCAttributesBO");
	if (objOmPC != null)
	{
		var strXML = "<?xml version='1.0'?><REQUEST ACTION='CREATE'></REQUEST>";
		strOut = objOmPC.FindLocalPrinterList(strXML);
		objOmPC = null;
	}
	return strOut;
}
<% /* BM0262 End */ %>
<% /* BM0567 START */ %>
function CheckDataFreezeFlag()
{
    var sDataFreezeFlag       = scScreenFunctions.GetContextParameter(window,"idFreezeDataIndicator",null);
    var sReadOnly             = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
    var sProcessingIndicator  = scScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null);
	var sCDDataFreezeFlag	  = scScreenFunctions.GetContextParameter(window,"idCancelDeclineFreezeDataIndicator",null);

    //if not readonly due to ReadOnly and ProcessingIndicator context variables, and datafreeze is set in context...

	<%/* EP989 - Check if either data freeze flag has been set */%>

	if (((sReadOnly != "1") && (sProcessingIndicator == "1")) && ((sDataFreezeFlag == "1") || (sCDDataFreezeFlag == "1")))
    {
        <% // Remove R/O %>
		lblFW030Status.style.visibility = "hidden"; <%//BM0567-Remove R/O (if displayed)%>        
        <% // Enable Printing %>
		divFW030PrintingOptionDisabled.style.display = "none";
		divFW030PrintingOptionEnabled.style.display = "block";<% //BM0567-ensure printing enabled%>
    }
}
<% /* BM0567 END */ %>

<% /* MAR7 GHun */ %>
function PrintDocumentForTask(node)
{
	var sQuotationNumber = "";
	var sMortgageSubQuoteNo = "";
	var bPrintError = false;
	
	var PrintXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	PrintXML.CreateRequestTag(window , "PRINTDOCUMENTFORTASK");
	PrintXML.SetAttribute("OPERATION","PRINTDOCUMENT");
	PrintXML.SetAttribute("PRINTINDICATOR", "1");
	
	CreatePrintDataBlock(PrintXML,"REQUEST");
	
	CreateTemplateDataBlock(PrintXML,"REQUEST");
	
	CreateControlDataBlock(PrintXML,"REQUEST");

	
	switch (ScreenRules())
	{
		case 1:
		case 0:
			PrintXML.RunASP(document, "PrintManager.asp");
			if(PrintXML.IsResponseOK())
			{
				if (TempXML.IsInComboValidationList(document,"PrinterDestination",sPrinterType,["W"]))
				{
// TW 01/12/2005 MAR764	
//					CallAxwordPrintDocument(PrintXML);
					CallAxwordClass(PrintXML);
// TW 01/12/2005 MAR764	
				}	
				else
				{
<% /* MAR578 TW 16/11/2005 */ %>
//					scScreenFunctions.MSGAlert("Document has been sent to print.");
					alert("Document has been sent to print.");
<% /* MAR578 TW 16/11/2005 End */ %>
				}
				
				<% /* UpdateCaseTask */ %>
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XML.CreateRequestTag(window , "UpdateCaseTask");
				var newNode = node.cloneNode(true);
				newNode.setAttribute("TASKSTATUS", "40"); 
				XML.ActiveTag.appendChild(newNode);
				UpdateCaseTask(XML.XMLDocument.xml);
				return false ; 
			}
			break;
		default:
			{	
				PrintXML.SetErrorResponse();
				return true ;
			}
	}
	
	return false ;
}


function ProcessAutomaticTasks(sPrinter, xmlApplication)
{
	
	sTaskName = "" ;
		
	<% /* Get the Current Stage Node */ %>
	var sActivitySequenceNo = "";
	var stageXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	stageXML.CreateRequestTag(window , "GetCurrentStage");
	stageXML.CreateActiveTag("CASEACTIVITY");
	stageXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	stageXML.SetAttribute("CASEID", m_sApplicationNumber);
	stageXML.SetAttribute("ACTIVITYID", m_sActivityId);
	stageXML.SetAttribute("ACTIVITYINSTANCE", "1");

	stageXML.RunASP(document, "MsgTMBO.asp");
	
	<% /* If Response OK  */ %>
	if(stageXML.IsResponseOK())	
	{
		<% /* Select CaseStage Node from the Response  */ %>
		var CaseStageXML = stageXML.SelectTag(null, "CASESTAGE");	
		
		if (stageXML.SelectTag(null, "CASESTAGE")!= null)
		{
			<% /* List all nodes with AutoMaticTaskInd = '1' from the CaseStage Node  */ %>
			stageXML.CreateTagList("CASETASK[@AUTOMATICTASKIND='1']");
			var iNoOfTasks = stageXML.ActiveTagList.length;
			if (iNoOfTasks > 0)
			{
				var iLoop;
				for(iLoop = 0; iLoop < iNoOfTasks; iLoop++)
				{
					stageXML.SelectTagListItem(iLoop);
					
					sTaskName  =  stageXML.GetAttribute("TASKNAME");
					sInterfaceName = stageXML.GetAttribute("INTERFACE");
					sDocumentID = stageXML.GetAttribute("OUTPUTDOCUMENT");
					bAutomaticTask = true;
					
					if (sInterfaceName != "" && sInterfaceName != null)
					{
						if (makeInterfaceCall(sInterfaceName, stageXML.ActiveTag ,sDocumentID))
						{			
							var sApplicationNo = scScreenFunctions.GetContextParameter(window,"idApplicationNumber", "");
							scScreenFunctions.SetOtherSystemACNoInContext(sApplicationNo);
						}
					}	
				}
			}
		}
	}	
	else
	{
<% /* MAR578 TW 16/11/2005 */ %>
//		scScreenFunctions.MSGAlert("No current CASESTAGE detail");
		alert("No current CASESTAGE detail");
<% /* MAR578 TW 16/11/2005 End */ %>
		return false;
	}
	bAutomaticTask = false;
}

function CreateTemplateDataBlock(PrintXML,sRootElement)
{
	//GetApplicationData
	var AppXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	AppXML.CreateRequestTag(window, "GetAcceptedQuoteData");
	AppXML.CreateActiveTag("APPLICATION");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AppXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
			
	AppXML.RunASP(document,"GetAcceptedOrActiveQuoteData.asp");
	if(AppXML.IsResponseOK())
	{
		// EP710 - Ignore NULL quotation node
		if(AppXML.SelectTag(null, "QUOTATION"))
			sQuotationNumber = AppXML.GetTagText("QUOTATIONNUMBER");
		else
			sQuotationNumber = "";
		
		// EP710 - Ignore NULL mortgagesubquote node	
		if(AppXML.SelectTag(null, "MORTGAGESUBQUOTE"))
			sMortgageSubQuoteNo = AppXML.GetTagText("MORTGAGESUBQUOTENUMBER");
		else
			sMortgageSubQuoteNo = ""
		
		PrintXML.SelectTag(null, sRootElement);
		xmlTemplateDataNode = PrintXML.CreateActiveTag("TEMPLATEDATA");
		PrintXML.SetAttribute("QUOTATIONNUMBER", sQuotationNumber);
		PrintXML.SetAttribute("MORTGAGESUBQUOTENUMBER",sMortgageSubQuoteNo);
	}	
}

function CreatePrintDataBlock(PrintXML,sRootElement)
{
	PrintXML.SelectTag(null,sRootElement);
	xmlPrintDataNode = PrintXML.CreateActiveTag("PRINTDATA");
	PrintXML.SetAttribute("METHODNAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "PDMMETHOD"));	
														
	PrintXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	PrintXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	
}

function CreateControlDataBlock(PrintXML,sRootElement)
{
	PrintXML.SelectTag(null, sRootElement);
	
	xmlControlDataNode = PrintXML.CreateActiveTag("CONTROLDATA");
	
	//DocumentID
	PrintXML.SetAttribute("DOCUMENTID",sDocumentID);
	
	//Printer
	if (sPrinterValidationType == "L")
	{	
		LocalPrintersXML.SelectTag(null, "RESPONSE");			
		var sPrinter = LocalPrintersXML.GetTagText("PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
		PrintXML.SetAttribute("PRINTER", sPrinter);
	}	
	else 
	{									
		PrintXML.SetAttribute("PRINTER", AttribsXML.GetTagAttribute("ATTRIBUTES", "REMOTEPRINTERLOCATION"));
	}	
	
	//Max Copies
	PrintXML.SetAttribute("COPIES",AttribsXML.GetTagAttribute("ATTRIBUTES", "DEFAULTCOPIES"));
	
	//DestinationType

	PrintXML.SetAttribute("DESTINATIONTYPE", sPrinterValidationType);
	
	//DPSDocumentID
	PrintXML.SetAttribute("DPSDOCUMENTID",AttribsXML.GetTagAttribute("ATTRIBUTES", "DPSTEMPLATEID"));
	
	//HostTemplateName
	PrintXML.SetAttribute("HOSTTEMPLATENAME",AttribsXML.GetTagAttribute("ATTRIBUTES","HOSTTEMPLATENAME"));
	
	//HostTemplateDescription
	PrintXML.SetAttribute("HOSTTEMPLATEDESCRIPTION",AttribsXML.GetTagAttribute("ATTRIBUTES","HOSTTEMPLATEDESCRIPTION"));
	
	//PostToWeb
	PrintXML.SetAttribute("POSTTOWEB","0");

// TW 01/12/2005 MAR764	
//	PrintXML.SetAttribute("DELIVERYTYPE",m_sDeliveryType);
	PrintXML.SetAttribute("DELIVERYTYPE",AttribsXML.GetTagAttribute("ATTRIBUTES", "DELIVERYTYPE"));
// TW 01/12/2005 MAR764	End
	
	PrintXML.SetAttribute("COMPRESSIONMETHOD",m_sCompressionMethod);
	
	PrintXML.SetAttribute("EDITBEFOREPRINT",AttribsXML.GetTagAttribute("ATTRIBUTES", "EDITBEFOREPRINT"));
	
	PrintXML.SetAttribute("VIEWBEFOREPRINT",AttribsXML.GetTagAttribute("ATTRIBUTES", "VIEWBEFOREPRINT"));
	
	//WebDocumentType
	PrintXML.SetAttribute("WEBDOCUMENTTYPE",AttribsXML.GetTagAttribute("ATTRIBUTES","WEBDOCUMENTTYPE"));
	
}

function CallAxwordClass(PrintXML)
{
	var bSuccess = true;
	var nCopies ;
	var fileContents ="";
	var bViewBeforePrint;
	var bShowPrintDialog ;
	var bShowProgressbar ;
	
	if (sDocumentID != "")
	{
		AttribsXML = getPrintDocumentAttributes(sDocumentID);
			
		m_sDeliveryType = AttribsXML.GetTagAttribute("ATTRIBUTES", "DELIVERYTYPE")
// TW 01/12/2005 MAR764	
//		m_sCompressionMethod = "ZLIB";
		m_sCompressionMethod = "";
// TW 01/12/2005 MAR764	End
		nCopies = 	AttribsXML.GetTagAttribute("ATTRIBUTES", "MAXCOPIES");
		sPrinterDestTypeID = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
		sPrinterValidationType = getPrinterType(window,sPrinterDestTypeID);
	
		fileContents = PrintXML.GetTagAttribute("DOCUMENTCONTENTS", "FILECONTENTS");
		if (fileContents == "")
		{
			fileContents = PrintXML.GetTagAttribute("PRINTDOCUMENTDETAILS", "PRINTDOCUMENT");
		}
				
		if (AttribsXML.GetTagAttribute("ATTRIBUTES", "VIEWBEFOREPRINT") == "0")
			bViewBeforePrint = false;
		else
			bViewBeforePrint = true;
	
		if (bAutomaticTask)
		{
			bShowPrintDialog = false;
			bViewBeforePrint = false;
			bShowProgressbar = true;
		}
		else
		{
			if (sPrinterValidationType == "W")
			{
				bShowPrintDialog = true;
				bShowProgressbar = true;
			}
		}
		
		if (sPrinterValidationType == "W")		
		{
			var printDocumentData = 
								printDocument(
											fileContents, 
											sDocumentID, 
											sTaskName,
											m_sDeliveryType, 
											m_sCompressionMethod, 
											sPrinterValidationType,
											nCopies,
											bShowPrintDialog, 
											bShowProgressbar,
											true, 
											!bViewBeforePrint)
	
			if (printDocumentData != null && printDocumentData.get_success())
			{
				if (printDocumentData.get_fileContents() != null)
				{	
				
					CreateControlDataBlock(PrintXML,"RESPONSE");
					CreatePrintDataBlock(PrintXML,"RESPONSE");	
					CreateTemplateDataBlock(PrintXML,"RESPONSE");	
						
					bSuccess = 
								savePrintedDocument(
												window, m_UserId, m_UnitId, m_MachineId, m_DistributionChannelId, "10", 
												xmlControlDataNode, xmlPrintDataNode, xmlTemplateDataNode, 
												printDocumentData.get_fileSize(), 
												printDocumentData.get_fileContents(), 
												m_sCompressionMethod);
				}
				else
				{
					bSuccess = true;
				}
			}
			else		
			{
				bSuccess = false;
			}
		}	
		else
		{
			bSuccess = true;
		}
	}
	//MAR725 No document, don't call updatecasetask
	else
	{
		bSuccess = false;
	}

	return bSuccess;
}
<% /* MAR7 End */ %>

-->
</script>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
</body>
</html>
