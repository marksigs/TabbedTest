<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      TM032.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Task Processing
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		26/10/00	Created (screen paint)
JLD		22/11/00	added BO calls
JLD		23/11/00	Unit test fixing
JLD		29/11/00	Deal with attributes not being in the XML (returning null rather than "")
JLD		03/01/01	SYS1761 added check for NOTASKAUTHORITY.
JLD		03/01/01	SYS1762 show error message if both completed and not-completed tasks are selected for action
JLD		09/01/01	SYS1789 use RECORDNOTFOUND instead of NOMORESTAGES
JLD		12/01/01	SYS1807 Handle exception stages.
JLD		14/02/01	SYS1832 added functionality for inputprocess and outputdocument tasks
APS		02/03/01	SYS1920 Interface processing for tasks
JLD		06/03/01	SYS1879 Add routing to AP010
JLD		07/03/01	SYS1879 When routing to an INPUTPROCESS screen, set the context idReturnScreenId
					to TM030 to tell the routed-to screen where to route back to.
DJP		07/03/01	SYS1839 Added call to AP205.
SR		08/03/01	Added call to AP300
APS		08/03/01	SYS1923 Issue Offer Functionality
CL		12/03/01	SYS1920 Read only functionality added
GD		12/03/01	SYS2039 allow routing to AP250
APS		14/03/01    SYS2064 Added in AP400.asp routing
APS		08/03/01	SYS1923 Amended Re/Issue Offer Functionality
IK		17/03/01	SYS1924 Completeness Check routing etc.
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
BG		30/03/01	SYS2096 Added functionality to handle multiple tasks on btnAction.onclick
DRC     16/07/01    SYS2390 - Check for null on CustomerIdentifier & Context Attributes
DRC     16/07/01    SYS2088 - Code to advance to next stage if last task has been moved to TM30
GD		06/08/01	SYS2560 - Rework of btnAction.onclick to allow 'simple'tasks
					(ie. non-printing, non-interface and non-screen tasks) to be updated.
					AQR is still open due to MS parser issue.
JLD		08/10/01	SYS2736 Don't download omPC again
BG		21/11/01	SYS2984 Conditionally added customerversionnumber element in PrintTask function.
PSC		03/12/01	SYS3286 Get CUSTOMERIDENTIFIER rather than CUSTOMERNUMBER
PSC		03/12/01	SYS3288 Route to TM030 after printing task
PSC		06/12/01	SYS3358 Make RouteToInputProcess dynamic rather than being hardcoded
BG		04/11/01	SYS333 Added TASKINSTANCE and CASESTAGESEQUENCENO attributes on btnchaseup.onclick
DRC     10/12/01    SYS3295 - added for Interface call to IssueOffer print data if there's a doc in the task
BG		21/12/01	SYS3268 - added APPLICATIONPRIORITY to makeinterfacecall function
JLD		15/01/02	SYS3517 - when routing to CM010, set up context param.
DRC     03/05/02  	SYS4530 - Check for Admin sys account no & put it into side bar
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

DPF		04/09/02	BMIDS00344 - BM092 Amend PrintTasks() function and create new function ProcessPrint()
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		09/10/2002	BMIDS00556	 
SA		29/10/2002	BMIDS00719 ProcessPrint changed - InputProcess not always available
IK		05/11/2002	BMIDS00730 edit or view before printing options
MO		07/11/2002	BMIDS00101 Made changes to pass the correct details to PM010
GD		18/11/2002	BMIDS00037
MDC		03/12/2002	BM0134 Screen Rules in MakeInterfaceCall and IsApplicationAtOfferStage
AW		04/12/02	BM0137		When printing a task with no Input Process, set task to complete
AW		19/12/02	BM0201		Restructured PrintTask()
GHun	14/01/2003	BM0201		Cancel from PM010 completes task even though letter not printed
MV		10/02/2003	BM0337		Amended btnAction.Onclick() and btnReprint.Onclick();
MV		17/03/2003	BM0337		Amended frmScreen.btnReprint.onclick();frmScreen.btnAction.onclick()
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<% /* Specify Forms Here */ %>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM036" method="post" action="TM036.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM037" method="post" action="TM037.asp" STYLE="DISPLAY: none"></form>
<form id="frmInputProcess" method="post" action="" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 200px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Task
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtTask" maxlength="10" style="WIDTH:200px" class="msgTxt">
	</span>
</span>
<span style="TOP:7px; LEFT:350px; POSITION:ABSOLUTE">
	<input id="btnAction" value="Action" type="button" style="WIDTH:80px" class="msgButton">
</span>
<span style="TOP:40px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Status
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtStatus" maxlength="10" style="WIDTH:200px" class="msgTxt">
	</span>
</span>
<span style="TOP:31px; LEFT:350px; POSITION:ABSOLUTE">
	<input id="btnDetails" value="Details" type="button" style="WIDTH:80px" class="msgButton">
</span>
<span style="TOP:70px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Due Date
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtDueDate" maxlength="10" style="WIDTH:200px" class="msgTxt">
	</span>
</span>
<span style="TOP:55px; LEFT:350px; POSITION:ABSOLUTE">
	<input id="btnReprint" value="Reprint" type="button" style="WIDTH:80px" class="msgButton">
</span>
<span style="TOP:100px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Last Actioned Date
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtLastActionedDate" maxlength="10" style="WIDTH:200px" class="msgTxt">
	</span>
</span>
<span style="TOP:79px; LEFT:350px; POSITION:ABSOLUTE">
	<input id="btnChaseUp" value="Chase Up" type="button" style="WIDTH:80px" class="msgButton">
</span>
<span style="TOP:130px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Owner
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtOwner" maxlength="10" style="WIDTH:200px" class="msgTxt">
	</span>
</span>
<span style="TOP:103px; LEFT:350px; POSITION:ABSOLUTE">
	<input id="btnNotApplicable" value="Not Applicable" type="button" style="WIDTH:80px" class="msgButton">
</span>
<span style="TOP:160px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Last Updated By
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtLastUpdatedBy" maxlength="10" style="WIDTH:200px" class="msgTxt">
	</span>
</span>
<span style="TOP:127px; LEFT:350px; POSITION:ABSOLUTE">
	<input id="btnMemo" value="Memo" type="button" style="WIDTH:80px" class="msgButton">
</span>
<span style="TOP:151px; LEFT:350px; POSITION:ABSOLUTE">
	<input id="btnContactDetails" value="Contact Details" type="button" style="WIDTH:80px" class="msgButton">
</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 270px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/TM032attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sTasksXML = "";
var m_taskXML = null;
var AttribsXML = null;
var TaskStatusXML = null;
var scScreenFunctions;
var m_sApplicationNumber;
var m_sApplicationFFNumber;
var m_blnReadOnly = false;
var LocalPrintersXML = "";



<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Task Processing","TM032",scScreenFunctions);

	RetrieveContextData();
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
	PopulateScreen();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "TM032");	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
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
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sTasksXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFFNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
}

function PopulateScreen()
{
	scScreenFunctions.SetFieldState(frmScreen, "txtTask", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtStatus", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtDueDate", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtLastActionedDate", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtOwner", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtLastUpdatedBy", "R");
	
	// Get the TaskStatus combo information
	TaskStatusXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("TaskStatus");
	TaskStatusXML.GetComboLists(document, sGroups);
	
	m_taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_taskXML.LoadXML(m_sTasksXML);
	m_taskXML.SelectTag(null, "REQUEST");
	m_taskXML.CreateTagList("CASETASK");
	if( m_taskXML.ActiveTagList.length > 1)
	{
		frmScreen.txtTask.value = "Multiple Selection";
	}
	else
	{
		m_taskXML.SelectTagListItem(0);
		frmScreen.txtTask.value = m_taskXML.GetAttribute("TASKNAME");
		frmScreen.txtStatus.value = GetComboText(m_taskXML.GetAttribute("TASKSTATUS"));
		frmScreen.txtDueDate.value = m_taskXML.GetAttribute("TASKDUEDATEANDTIME");
		frmScreen.txtLastActionedDate.value = m_taskXML.GetAttribute("TASKSTATUSSETDATETIME");
		frmScreen.txtOwner.value = m_taskXML.GetAttribute("OWNINGUSERID");
		frmScreen.txtLastUpdatedBy.value = m_taskXML.GetAttribute("TASKSTATUSSETBYUSERID");
	}
	SetButtonStates();
}
function SetButtonStates()
{	
	var bActionButtonState = false;
	var bDetailsButtonState = false;
	var bReprintButtonState = false;
	var bChaseUpButtonState = false;
	var bNotApplicableButtonState = false;
	var bMemoButtonState = false;
	var bContactDetailsButtonState = false;
	var bMixIncludesCompleted = false;

	m_taskXML.SelectTag(null, "REQUEST");
	m_taskXML.CreateTagList("CASETASK");
	var bMultiInput = false;
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
			m_taskXML.CreateTagList("CASETASK[@TASKSTATUS!='10']");
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
		if( bMultiInput && 
		    (m_taskXML.GetAttribute("TASKSTATUS") == "40" || m_taskXML.GetAttribute("TASKSTATUS") == "70"))
				bMixIncludesCompleted = true; //SYS1762 
		//if( bMultiInput &&
		//	((m_taskXML.GetAttribute("INPUTPROCESS") != null && m_taskXML.GetAttribute("INPUTPROCESS") != "") ||
		//	 (m_taskXML.GetAttribute("OUTPUTDOCUMENT") != null && m_taskXML.GetAttribute("OUTPUTDOCUMENT") != "") ||
		//	 (m_taskXML.GetAttribute("INTERFACE") != null && m_taskXML.GetAttribute("INTERFACE") != ""))        )
		//		bActionButtonState = true;
		if( m_taskXML.GetAttribute("TASKSTATUS") != "10") // "Not Actioned"
				bActionButtonState = true;
		if( bMultiInput || 
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
		if( m_taskXML.GetAttribute("NOTAPPLICABLEFLAG") == "0" ||
		    m_taskXML.GetAttribute("TASKSTATUS") != "10" &&  //"Not Actioned"
		    m_taskXML.GetAttribute("TASKSTATUS") != "20"    )  //"Pending"
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
	if(bMixIncludesCompleted)  Be("Selected tasks include ones already Completed. These cannot be processed further.");
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
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			frmInputProcess.submit();
			break;
		default: // Error
	}
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


//DPF 6/9/02 - APWP3 -  new function to handle printing of tasks - code extracted from PrintTask function.
function ProcessPrint(m_taskXML, iCount)
{
	var node = null;
	var sCaseCustomerNumber = null;
	var sApplicationNumber = null;
	var bpopped = "0";
	var bPrintError = null;
	
	bPrintError = false;
	
	node = m_taskXML.GetTagListItem(iCount);
	AttribsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AttribsXML.CreateRequestTag(window , "GetPrintAttributes");
	AttribsXML.CreateActiveTag("FINDATTRIBUTES");			
	AttribsXML.SetAttribute("HOSTTEMPLATEID", node.getAttribute("OUTPUTDOCUMENT"));
									
	// 	AttribsXML.RunASP(document, "PrintManager.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			AttribsXML.RunASP(document, "PrintManager.asp");
			break;
		default: // Error
			AttribsXML.SetErrorResponse();
		}

				
	if(AttribsXML.IsResponseOK())
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
				
		var sPrinterType = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
		var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
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
			if(bpopped != "1")
			{		
				//call OmTMBO print method 
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
						
				// PSC 03/12/01 SYS3286
				sCaseCustomerNumber = node.getAttribute("CUSTOMERIDENTIFIER");
				var sContext = node.getAttribute("CONTEXT");
				XML.CreateRequestTag(window , "PrintDocumentForTask");
				XML.CreateActiveTag("CASETASK");
				XML.SetAttribute("SOURCEAPPLICATION", node.getAttribute("SOURCEAPPLICATION"));
				XML.SetAttribute("CASEID", node.getAttribute("CASEID"));
				XML.SetAttribute("ACTIVITYID", node.getAttribute("ACTIVITYID"));
				XML.SetAttribute("ACTIVITYINSTANCE", node.getAttribute("ACTIVITYINSTANCE"));
				XML.SetAttribute("STAGEID", node.getAttribute("STAGEID"));
				XML.SetAttribute("CASESTAGESEQUENCENO", node.getAttribute("CASESTAGESEQUENCENO"));
				XML.SetAttribute("TASKID", node.getAttribute("TASKID"));
				XML.SetAttribute("TASKINSTANCE", node.getAttribute("TASKINSTANCE"));
				XML.SetAttribute("TASKSTATUS", node.getAttribute("TASKSTATUS"));
				XML.SetAttribute("OUTPUTDOCUMENT", node.getAttribute("OUTPUTDOCUMENT"));
				//BMIDS00719 There won't always necessarily be an input process
				if(node.getAttribute("INPUTPROCESS")!= null)
				{				
					XML.SetAttribute("INPUTPROCESS", node.getAttribute("INPUTPROCESS"));
				}
				if(sContext != null)
				{
					XML.SetAttribute("CONTEXT", sContext);
				}
                if(sCaseCustomerNumber != null)
				{
					XML.SetAttribute("CUSTOMERIDENTIFIER", sCaseCustomerNumber);
				}
						
				XML.SelectTag(null, "REQUEST");
				XML.CreateActiveTag("APPLICATION");
				XML.SetAttribute("APPLICATIONFACTFINDNUMBER", scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null));
						
				<%/*BG 21/11/01 SYS2984 moved "var sCustomerVersionNumber = "";" outside of if statement  */%>
				var sCustomerVersionNumber = "";
						
				if(sCaseCustomerNumber != "")
				{
					var iCount = 0;
					var sCustomerNumber = "";	
					<% /* var sCustomerVersionNumber = ""; */%>	
												
					for (iCount = 1; iCount <= 5; iCount++)
					{								
						sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
						if (sCustomerNumber == sCaseCustomerNumber) 
						{					
							sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);					
						}
					}			
				} 	
				<%/*BG 21/11/01 SYS2984 added if statement to conditionally add customerversionnumber element */%>
						
				if(sCustomerVersionNumber !="")
				{
					XML.SetAttribute("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
				}						
				XML.SelectTag(null, "REQUEST");
				
//	ik_debug
//	stopIT();
//	var bBogus=true;
//	ik_debug_ends
				
				// ik_bmids00831
				// copy ALL PRINTATTRIBUTES
				var xmlSrceNode = AttribsXML.XMLDocument.selectSingleNode("RESPONSE/ATTRIBUTES");
				var xmlTargetNode = XML.CreateActiveTag("PRINTATTRIBUTES");
				for(var i=0; i < xmlSrceNode.attributes.length;i++)
					xmlTargetNode.attributes.setNamedItem(xmlSrceNode.attributes.item(i).cloneNode(true));

				// still needed				
				XML.SetAttribute("COPIES", AttribsXML.GetTagAttribute("ATTRIBUTES", "DEFAULTCOPIES"));
//				XML.SetAttribute("RECIPIENTTYPE", AttribsXML.GetTagAttribute("ATTRIBUTES", "RECIPIENTTYPE"));
				XML.SetAttribute("DPSDOCUMENTID", AttribsXML.GetTagAttribute("ATTRIBUTES", "DPSTEMPLATEID"));
				XML.SetAttribute("METHODNAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "PDMMETHOD"));
						
				TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				var sGroups = new Array("PrinterDestination");
				var sList = TempXML.GetComboLists(document,sGroups);
				var sValueID = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
				var sPrintType = TempXML.GetTagText("LIST/LISTNAME/LISTENTRY[VALUEID='" + sValueID + "']/VALIDATIONTYPELIST/VALIDATIONTYPE");
						
				XML.SetAttribute("DESTINATIONTYPE", sPrintType);
						
				var sPrinterType = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
				var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				if (TempXML.IsInComboValidationList(document,"PrinterDestination",sPrinterType,["L"]))
				{	
					LocalPrintersXML.SelectTag(null, "RESPONSE");			
					var sPrinter = LocalPrintersXML.GetTagText("PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
					XML.SetAttribute("PRINTER", sPrinter);
				}	
				else 
				{									
					XML.SetAttribute("PRINTER", AttribsXML.GetTagAttribute("ATTRIBUTES", "REMOTEPRINTERLOCATION"));
				}	
													
				TempXML = null;
														
				// 				XML.RunASP(document, "OmigaTMBO.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									XML.RunASP(document, "OmigaTMBO.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
					}

				if(XML.IsResponseOK())
				{
					//alert("Printed document " + node.getAttribute("OUTPUTDOCUMENT") );
				}
				else 
				{
					alert("Failed to print document " + node.getAttribute("OUTPUTDOCUMENT") );
					//AW	19/12/02	BM0201
					bPrintError = true;
				}
			}
		}
	}
	//AW	19/12/02	BM0201
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
	ArrayArguments[15] = m_taskXML.GetAttribute("CONTEXT");

	// ik_bmids00730
	nWidth = 540;
	nHeight = 560;
						
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

function frmScreen.btnAction.onclick()
{
	frmScreen.style.cursor = "wait";
	frmScreen.btnAction.disabled = true;
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
%>	m_taskXML.SelectTag(null, "REQUEST");
	m_taskXML.CreateTagList("CASETASK");
	//GD SYS2560 06/08/01
	var bUpdateCaseTask;
	var bInputProcess; //variable to store if an input process exists
		
	if(m_taskXML.ActiveTagList.length == "1")
	{
		//GD SYS2560 06/08/01
		bUpdateCaseTask = true;
		bInputProcess
		m_taskXML.SelectTagListItem(0);  // select the first CASETASK item!

		// APS SYS1920
		if (m_taskXML.GetAttribute("INTERFACE") != "")
		{
			//GD SYS2560 06/08/01
			bUpdateCaseTask = false;
			var sTaskId = m_taskXML.GetAttribute("TASKID");
			
			if (makeInterfaceCall(m_taskXML.GetAttribute("INTERFACE"), m_taskXML.ActiveTag) == true)
			{			
				if ((sTaskId.toUpperCase() == GetIssueOfferParameter()) || 
					(sTaskId.toUpperCase() == GetReIssueOfferParameter()))
				{
					scScreenFunctions.SetContextParameter(window, "idFreezeDataIndicator", "1");
				}
				// AQR SYS4530 - Check for Admin sys account no & put it into side bar
				if (m_taskXML.GetAttribute("INTERFACE") == "GetNewNumbers")
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
					frmToTM030.submit();
				}
			}
		} 	
		else if (m_taskXML.GetAttribute("OUTPUTDOCUMENT") != "" &&
		   m_taskXML.GetAttribute("TASKSTATUS") == "10"      )
		{
			//GD SYS2560 06/08/01
			bUpdateCaseTask = false;
			bInputProcess = PrintTask();
			//DPF 6/9/02 - APWP3 - If there's an input process re-direct to it not TM030
			if (bInputProcess != true)
			{
				//AW BM0201 19/12/02
				//bUpdateCaseTask = true; //AW BM0137 04/12/02
				frmToTM030.submit(); // PSC 03/12/01 SYS3288
			}
		}
		else if(m_taskXML.GetAttribute("INPUTPROCESS") != "")
		{
			//GD SYS2560 06/08/01
			bUpdateCaseTask = false;			
			RouteToInputProcess(m_taskXML.GetAttribute("INPUTPROCESS"));
		}
		//GD SYS2560 Update the single task 
		//Update case task if it hasn't been done in the above code
		if (bUpdateCaseTask == true)
		{
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTag(window , "UpdateCaseTask");
			var node = m_taskXML.SelectTag(null, "CASETASK");
			var newNode = node.cloneNode(true);
			newNode.setAttribute("TASKSTATUS", "40"); // 40 = Complete
			XML.ActiveTag.appendChild(newNode);
			UpdateCaseTask(XML.XMLDocument.xml);		
		}
	}
	else //Multiple Tasks
	{	
		//BG 30/03/01 find any nodes which either have no outputdocument attribute, or one with no value set
		m_taskXML.CreateTagList("CASETASK[$not$ @OUTPUTDOCUMENT or @OUTPUTDOCUMENT='']");
		if(m_taskXML.ActiveTagList.length == 0)
		{
			PrintTask();
			frmToTM030.submit();
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
			} else
			{
				alert("Cannot process tasks of different types.");
				frmToTM030.submit();
			}
		}
	}
	frmScreen.btnAction.disabled = false;
	frmScreen.style.cursor = "default";
}

function frmScreen.btnNotApplicable.onclick()
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
	var sResult = IsApplicationAtOfferStage(m_sApplicationNumber);
	if (sResult != -1) //Call has succeeded
	{
		scScreenFunctions.SetContextParameter(window,"idFreezeDataIndicator", sResult);
	}
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
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "MsgTMBO.asp");	
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	var ErrorTypes = new Array("NOTASKAUTHORITY");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{
		alert("User does not have the authority to update this task");
	}
	// SYS2088 - Code for Moving to next stage moved to start of TM030
	frmToTM030.submit();
}

function frmScreen.btnDetails.onclick()
{
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
	
	frmScreen.style.cursor = "wait";	
	frmScreen.btnReprint.disabled = true;
	
	m_taskXML.SelectTag(null, "REQUEST");
	m_taskXML.CreateTagList("CASETASK");
	
	m_taskXML.SelectTagListItem(0);
	if(m_taskXML.GetAttribute("OUTPUTDOCUMENT") != "")
		PrintTask();
	
	frmScreen.btnReprint.disabled = false;
	frmScreen.style.cursor = "default";
	
}
function frmScreen.btnChaseUp.onclick()
{
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
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "OmigaTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
	}
	if(XML.IsResponseOK())
		frmToTM030.submit();
}
function frmScreen.btnMemo.onclick()
{
	// tasks XML is still availlable in the context under idTaskXML
	scScreenFunctions.SetContextParameter(window,"idReturnScreenId","TM030");
	frmToTM036.submit();
}
function frmScreen.btnContactDetails.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ArrayArguments[0] = m_sReadOnly;
	ArrayArguments[1] = m_taskXML.XMLDocument.xml;
	ArrayArguments[2] = XML.CreateRequestAttributeArray(window);
	sReturn = scScreenFunctions.DisplayPopup(window, document, "TM034.asp", ArrayArguments, 460, 280);
}
function btnCancel.onclick()
{
	frmToTM030.submit();
}
function btnSubmit.onclick()
{
<% // An example submit function showing the use of the validation functions 
%>	if(frmScreen.onsubmit())
	{
		//if(IsChanged())
<%			// Do some processing
%>
		//frmToXX999.submit();
	}
}

// SYS1920
function makeInterfaceCall(sOperation, xmlCaseTask)
{
	var blnSuccess = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ReqTag;
	ReqTag = XML.CreateRequestTag(window,sOperation);		
	XML.ActiveTag.appendChild(xmlCaseTask.cloneNode(true));
	XML.CreateActiveTag("APPLICATION");
	XML.SetAttribute("APPLICATIONNUMBER", scScreenFunctions.GetContextParameter(window, "idApplicationNumber", ""));
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", ""));
	//SYS3268 BG Added APPLICATIONPRIORITY
	XML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window, "idApplicationPriority", ""));
// SYS3295 - add print data	if there's a doc in the task
    if (m_taskXML.GetAttribute("OUTPUTDOCUMENT") != "") 
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
			var iCount = 0;
			var sCustomerVersionNumber = "";
			var sCustomerNumber = "";
			for (iCount = 1; iCount <= 5; iCount++)
			{								
				sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
				if (sCustomerNumber != "")
				{	
					sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);	
					XML.SelectTag(null, "APPLICATION");
					XML.CreateActiveTag("CUSTOMER");				
					XML.SetAttribute("CUSTOMERNUMBER", sCustomerNumber);
					XML.SetAttribute("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);																		
				}
			}
			XML.ActiveTag = ReqTag;
			XML.CreateActiveTag("PRINTER");
			XML.SetAttribute("PRINTERNAME", sPrinter);
			XML.SetAttribute("DEFAULTIND", "1");
		}	
	}

	<% /* BM0134 MDC 03/12/2002 */ %>
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
	<% /* BM0134 MDC 03/12/2002 - End */ %>

	if (XML.IsResponseOK()) {
		
		blnSuccess = true;		
	}	
	XML = null;
	return blnSuccess;
}
// SYS1920
function GetIssueOfferParameter()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var sTMIssueOffer = XML.GetGlobalParameterString(document, "TMIssueOffer");

	XML = null;

	return sTMIssueOffer.toUpperCase();
}
function GetReIssueOfferParameter()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var sTMReIssueOffer = XML.GetGlobalParameterString(document, "TMReIssueOffer");

	XML = null;

	return sTMReIssueOffer.toUpperCase();
}
//GD BMIDS00037
function IsApplicationAtOfferStage(sApplicationNumber)
{
	var sSuccess = -1;

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	// get the Activity ID from the global partameter table
	var sActivityId = scScreenFunctions.GetContextParameter(window, "idActivityId", "1");
	
	if (sActivityId == "") {
		sActivityId = XML.GetGlobalParameterAmount(document, "TMOmigaActivity");	
		//scScreenFunctions.SetContextParameter(window, "idActivityId", sActivityId);
		XML.ResetXMLDocument();
	}
	
	XML.CreateRequestTag(window, "ISAPPLICATIONATOFFER");
	XML.CreateActiveTag("CASESTAGE");
	
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("CASEID", sApplicationNumber);
	XML.SetAttribute("ACTIVITYID", sActivityId);
	
	<% /* BM0134 MDC 03/12/2002 */ %>
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
	<% /* BM0134 MDC 03/12/2002 - End */ %>
	
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
-->
</script>
<SCRIPT TYPE="text/vbscript" LANGUAGE="VBScript">
Function GetLocalPrinters()

	dim obj
	set obj  = createobject("omPC.PCAttributesBO")
	if not obj is nothing then
		strXML = "<?xml version=""1.0""?>"
		strXML = strXML & "<REQUEST ACTION = ""CREATE"">"
		strXML = strXML & "</REQUEST>"
		strOUT = obj.FindLocalPrinterList(strXML)
		set obj = nothing
	end if
	GetLocalPrinters = strOUT
	
End Function 
</SCRIPT>
</body>
</html>


