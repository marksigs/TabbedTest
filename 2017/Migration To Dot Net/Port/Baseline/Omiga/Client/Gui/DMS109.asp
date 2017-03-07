<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DMS109.asp
Copyright:     Copyright © 2005 Marlborough Stirling

Description:   Document Recategorisation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:
Prog	Date		AQR		Description
AM		13/09/2005  MAR50	Creating a new Screen to Recategorize the Document
GHun	27/02/2006	MAR1332 Get screen working
PE		05/04/2006	MAR1578	Increased size of field and screen so all items can be read.
PE		15/05/2006  MAR1776 Change taskstatus from 20 (pending) to 10 (incomplete)- LINE 405.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<head>
	<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title>Document Recategorisation<!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<object id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex="-1" type="text/x-scriptlet" height="1" width="1" data="scClientFunctions.asp" viewastext></object>

<% /* Span to keep tabbing within this screen */ %>

<span id="spnToLastField" tabindex="0"></span>
<% /* Specify Screen Layout Here */ %>

<form id="frmToDMS105" method="post" action="DMS105.asp" style="DISPLAY: none"></form>
<form id="frmScreen" validate ="onchange" mark>
<div id="divBackground" style="LEFT: 10px; WIDTH: 450px; POSITION: absolute; TOP: 10px; HEIGHT: 181px" class="msgGroup">

	<span style="LEFT: 20px; POSITION: absolute; TOP: 15px" class="msgLabel">
		Application Number
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtApplicationNumber" maxlength="12" style="WIDTH: 270px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 20px; POSITION: absolute; TOP: 45px" class="msgLabel">
		Applicant
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">   
			<select id="cboApplicant" style="WIDTH: 270px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 20px; POSITION: absolute; TOP: 75px" class="msgLabel">
		Document Purpose
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">   
			<select id="cboDocumentPurpose" style="WIDTH: 270px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 20px; POSITION: absolute; TOP: 105px" class="msgLabel">
		Document Group
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">   
			<select id="cboDocumentGroup" style="WIDTH: 270px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 20px; POSITION: absolute; TOP: 135px" class="msgLabel">
		Document Name
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">   
			<select id="cboDocumentName" style="WIDTH: 270px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 20px; POSITION: absolute; TOP: 190px">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
		<span style="LEFT: 80px; POSITION: absolute"> 
			<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>
	
</div>
</form>	
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<!-- #include FILE="attribs/DMS109Attribs.asp" -->
<% /* Specify Code Here */ %>
<script src="includes/Documents.js" language="jscript" type="text/javascript"></script>
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = null;
var m_blnReadOnly = false;
var m_sReadOnly = "";
var scScreenFunctions = null;
var m_XML = null;
var m_DocumentGUID = null;

var m_aArgArray = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sUserId = null;
var m_sXML = null;
var m_BaseNonPopupWindow = null;
var sCustomerName = "";

var scClientScreenFunctions;
var m_aCustomerVersionNumberArray = null;
var m_sCustomerName = new Array();
var m_sCustomerNumber = new Array();
var m_sCustomerVersion = new Array();

var m_sDocumentGroup = null;
var m_sDocumentName = null;
var m_sDocumentPurpose = null;
var m_comboXML = null;	<% /* MAR1332 GHun */ %>

function window.onload()
{
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();	
	RetrieveContextData();	
	m_XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	m_XML.LoadXML(m_sXML);
	m_XML.CreateTagList("REQUEST");

	if(m_XML.SelectTagListItem(0))
	{		
		m_DocumentGUID = m_XML.GetAttribute("DOCUMENTGUID");
		
		frmScreen.txtApplicationNumber.value = m_sApplicationNumber;
	}
	
	<% /* MAR1332 GHun */ %>
	GetCombos();
	PopulateCustomerCombo();

	if (m_sDocumentPurpose.length > 0)
	{
		frmScreen.cboDocumentPurpose.value = m_sDocumentPurpose;
		frmScreen.cboDocumentPurpose.onchange();
	}
	
	if (m_sDocumentGroup.length > 0)
		frmScreen.cboDocumentGroup.value = m_sDocumentGroup;
	
	for(var i=0; i < frmScreen.cboDocumentName.options.length; i++)
	{
		if(frmScreen.cboDocumentName.options[i].text == m_sDocumentName)
			frmScreen.cboDocumentName.selectedIndex = i;
	}
		
	SetMasks();
	Validation_Init();
	
	ClientPopulateScreen();
	<% /* MAR1332 End */ %>
}

function RetrieveContextData()
{	
	var sArguments = window.dialogArguments;
	window.dialogTop	= sArguments[0]	+ "px";
	window.dialogLeft	= sArguments[1] + "px";
	window.dialogWidth	= sArguments[2] + "px";
	window.dialogHeight = sArguments[3] + "px";	
	m_aArgArray = window.dialogArguments[4];
	m_BaseNonPopupWindow = window.dialogArguments[5];
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sApplicationNumber = m_aArgArray[0];
	m_sApplicationFactFindNumber = m_aArgArray[1];
	m_sUserId = m_aArgArray[2];
	m_sUnitId = m_aArgArray[3];
	m_sStageId = m_aArgArray[4];
	m_sStageName = m_aArgArray[5];
	m_sXML = m_aArgArray[11];//Any XML passed through
	
	m_sDocumentName = m_aArgArray[15];
	m_sDocumentGroup = m_aArgArray[16];
	m_sDocumentPurpose = m_aArgArray[17];
	m_aCustomerVersionNumberArray = m_aArgArray[14];
	
} 

<% /* MAR1332 GHun */ %>
function GetCombos()
{
	m_comboXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("DocumentPurpose", "DocumentGroup", "DocumentName");

	if (m_comboXML.GetComboLists(document, sGroupList))
	{
		var bSuccess = m_comboXML.PopulateCombo(document, frmScreen.cboDocumentPurpose, "DocumentPurpose", true);
	}
}
<% /* MAR1332 End */ %>

<%/* Populating DocumentNames Combo with values */%> 
function PopulateCustomerCombo()
{
	<% /* populate customers in context. Add a <SELECT> option */ %>
	var TagOPTION	= document.createElement("OPTION");
	TagOPTION.value	= "";
	TagOPTION.text	= "<SELECT>";

	frmScreen.cboApplicant.add(TagOPTION);
	m_nMaxCustomers = 0;
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerName" + nLoop);
		m_sCustomerName[nLoop] = sCustomerName;
		
		var sCustomerNumber = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerNumber" + nLoop);
		m_sCustomerNumber[nLoop] = sCustomerNumber;
		
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerVersionNumber" + nLoop);
		m_sCustomerVersion[nLoop] = sCustomerVersionNumber;
		
		if(sCustomerName != "" && sCustomerNumber != "")
		{
			m_nMaxCustomers++;
			TagOPTION		= document.createElement("OPTION");
			TagOPTION.value	= sCustomerNumber;
			TagOPTION.text	= sCustomerName;
			TagOPTION.setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
			frmScreen.cboApplicant.add(TagOPTION);
		}
	}	
	if (m_nMaxCustomers == 1) 
		frmScreen.cboApplicant.selectedIndex = 1;
	else 
		frmScreen.cboApplicant.selectedIndex = 0;
}

<% /* MAR1332 GHun */ %>
function frmScreen.cboDocumentPurpose.onchange()
{
	var valType = scScreenFunctions.GetComboValidationType(frmScreen, "cboDocumentPurpose");
	
	PopulateComboByValType(frmScreen.cboDocumentGroup, "DocumentGroup", valType); 
	PopulateComboByValType(frmScreen.cboDocumentName, "DocumentName", valType);
}

function PopulateComboByValType(cboCombo, strCombo, strValidationType)
{
	var tempXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var root = tempXML.CreateActiveTag("LISTNAME");
	var options = m_comboXML.XMLDocument.selectNodes("/RESPONSE/LIST/LISTNAME[@NAME='" + strCombo + "']/LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE[.='" + strValidationType + "']]");
				
	for(var nOpt=0; nOpt < options.length; nOpt++)
	{
		root.appendChild(options.item(nOpt).cloneNode(true));
	}

	m_comboXML.PopulateComboFromXML(document, cboCombo, tempXML.XMLDocument, true);
}

function CommitData()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var xmlRoot = XML.CreateRequestTag(m_BaseNonPopupWindow, "UPDATEDOCUMENTAUDITDETAILS");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("DOCUMENTGUID", m_DocumentGUID);
	
	var eventDetail = m_XML.SelectTag(null, "EVENTDETAIL");
	XML.ActiveTag.appendChild(eventDetail.cloneNode(true));

	XML.ActiveTag = xmlRoot;

	XML.CreateActiveTag("DOCUMENTDETAILS");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("DOCUMENTGUID", m_DocumentGUID);
	XML.SetAttribute("DOCUMENTGROUP", frmScreen.cboDocumentGroup.value);
	XML.SetAttribute("DOCUMENTPURPOSE", frmScreen.cboDocumentPurpose.value);
	XML.SetAttribute("DOCUMENTNAME", frmScreen.cboDocumentName.options.item(frmScreen.cboDocumentName.selectedIndex).text);

	switch (ScreenRules())
	{
		case 1: <% // Warning %>
		case 0: <% // OK %>
			XML.RunASP(document, "omPMRequest.asp");
			break;
		default: <% // Error %>
			XML.SetErrorResponse();
	}

	XML.SelectTag(null, "RESPONSE");
	if (!XML.IsResponseOK()) 
	{
		alert("Error saving document audit details.");
		return false;
	}
	else
	{
		ProcessTask();
		return true;
	}
}

function frmScreen.btnOK.onclick()
{
	if(frmScreen.onsubmit())
	{
		if(IsChanged())
		{
			if (CommitData())
				window.returnValue = 1;
			else
				window.returnValue = null;
			
			window.close();
		}
		else
			window.close();
	}
}

function frmScreen.btnCancel.onclick()
{
	window.returnValue = null;
	window.close();
}

function ProcessTask()
{
	var taskId = scScreenFunctions.GetComboValidationType(frmScreen, "cboDocumentPurpose");

	var bTaskUpdate	= false;
	
	//get params..
	var stageXML =		new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var taskXML =		new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sUserRole =		scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idRole", null);
	var sUserId =		scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idUserID", null)	
	var sUnitId =		scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idUnitID", null)
	var sApplPriority = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idApplicationPriority", null)
	var sActivityId =	scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idActivityId", null)
	var sStageId =		scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idStageID", null)
	
	//check if exists
	stageXML.CreateRequestTag(m_BaseNonPopupWindow, "GetCurrentStage");
	stageXML.CreateActiveTag("CASEACTIVITY");
	stageXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	stageXML.SetAttribute("CASEID", m_sApplicationNumber);
	stageXML.SetAttribute("ACTIVITYID", sActivityId);
	stageXML.SetAttribute("ACTIVITYINSTANCE", "1");
	stageXML.RunASP(document, "MsgTMBO.asp");
	
	if(!stageXML.IsResponseOK())
	{
		alert("Error retrieving current stage details.");
		return;
	}
		
	var sTaskSeqNo = stageXML.GetTagAttribute("CASESTAGE", "CASESTAGESEQUENCENO");
	
	<% /* Check if this task already exists */ %>
	var tagList = stageXML.CreateTagList("CASETASK");			
	
	for (var nItem = 0; nItem < tagList.length; nItem++)
	{
		stageXML.SelectTagListItem(nItem);			
		if(stageXML.GetAttribute("TASKID")==taskId) {				
			bTaskUpdate = true;
			break;
		}				
	}

	var currentDate = scScreenFunctions.DateTimeToString(scScreenFunctions.GetAppServerDate(true));

	if (bTaskUpdate == false) 
	{
		// create AdHocTask
		var reqTag = taskXML.CreateRequestTag(m_BaseNonPopupWindow, "CreateAdhocCaseTask");	
		taskXML.SetAttribute("USERID", sUserId);
		taskXML.SetAttribute("UNITID", sUnitId);		
		taskXML.SetAttribute("USERAUTHORITYLEVEL", sUserRole);				
		taskXML.CreateActiveTag("APPLICATION");		
		taskXML.SetAttribute("APPLICATIONPRIORITY", sApplPriority);
		taskXML.ActiveTag = reqTag;	
		taskXML.CreateActiveTag("CASETASK");
		taskXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		taskXML.SetAttribute("CASEID", m_sApplicationNumber);	
		taskXML.SetAttribute("ACTIVITYID", sActivityId);
		taskXML.SetAttribute("ACTIVITYINSTANCE", "1");		
		taskXML.SetAttribute("CASESTAGESEQUENCENO", sTaskSeqNo);
		taskXML.SetAttribute("STAGEID", sStageId);		
		taskXML.SetAttribute("TASKID", taskId);
		taskXML.SetAttribute("TASKDUEDATEANDTIME", currentDate);
		
		taskXML.RunASP(document, "OmigaTmBO.asp");
	}
	else
	{
		// Update Case Task
		reqTag = taskXML.CreateRequestTag(m_BaseNonPopupWindow, "UpdateCaseTask");
			
		taskXML.CreateActiveTag("CASETASK");
		taskXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		taskXML.SetAttribute("CASEID", m_sApplicationNumber);
		taskXML.SetAttribute("ACTIVITYID", sActivityId);
		taskXML.SetAttribute("STAGEID", sStageId);
		taskXML.SetAttribute("CASESTAGESEQUENCENO", sTaskSeqNo);
		taskXML.SetAttribute("TASKID", taskId);
		taskXML.SetAttribute("TASKINSTANCE", stageXML.GetAttribute("TASKINSTANCE"));
		taskXML.SetAttribute("TASKDUEDATEANDTIME", currentDate);
		taskXML.SetAttribute("TASKSTATUS", "10");	<% /* Set to incomplete */ %>
		taskXML.SetAttribute("TASKSTATUSSETDATETIME", currentDate);
		taskXML.SetAttribute("TASKSTATUSSETBYUSERID", sUserId);
		taskXML.SetAttribute("TASKSTATUSSETBYUNITID", sUnitId);
			
		taskXML.RunASP(document, "msgTMBO.asp");
	}				
	taskXML.IsResponseOK();
	
	stageXML = null;
	taskXML = null;
}
<% /* MAR1332 End */ %>
-->
</script>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
</body>
</html>
