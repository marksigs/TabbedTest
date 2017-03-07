<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
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
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
SR		26/03/01	SYS2111: Add Condition Ref to the request for deletion 
SR		27/03/01	SYS2111
SR		24/04/01	SYS2169
JLD		20/11/01	SYS3052 bug fix in UpdateTaskAndApplicationCondtions
STB		23/04/02	SYS2175 Delete no-longer enabled if no items are selected.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		07/08/2002	BMIDS00302	Remove non-style sheet styles
GD		26/09/2002	BMIDS00313	APWP2 - Allow for embedded DB values in ConditionDescription
TW      09/10/2002    Modified to incorporate client validation - SYS5115
ASu		11/10/2002	BMIDS00587	Correct spelling Error on Gui
GD		11/11/2002	BMIDS00766	Ensure buttons disabled if entered screen from Conditions Review task.
GD		12/11/2002	BMIDS00746	Allow double-click to toggle Satisfy status on/off
MV		10/04/2003  BM0520	Amended   CreateDefaultConditions()
BS		11/06/2003	BM0521	Disable screen when user in a non-processing unit or the case is locked
HMA     13/01/2004  BMIDS635  Correct Edit processing
KRW     01/12/2004  BM0520  Review Special Conditions Status now setting correctly
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ING Specific History : 

Prog	Date		AQR			Description
MV		12/09/2005	MAR35		Added new attribute in UpdateApplicationConditions() and 
								UpdateTaskAndApplicationCondtions()
PSC		16/06/2006	MAR1873		Set up correct request to CreateDefaultApplicationConditions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM History

Prog	date		AQR			Description
PB		19/06/2006	EP683		Bug with checking security level
PE		17/07/2006	EP974		MAR1873 - Mars Merge (Creating default application conditions fail)		
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<link href="stylesheet.css" rel="stylesheet" type="text/css">

	<% /* Validation script - Controls Soft Coded Field Attributes */ %>
	<script src="validation.js" language="JScript" type="text/javascript"></script>
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Scriptlets */ %>

<% /* FORMS */ %>
<form id="frmToAP310" method="post" action="AP310.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM032" method="post" action="TM032.asp" STYLE="DISPLAY: none"></form>

<span id="spnListScroll">
	<span style="LEFT: 306px; POSITION: absolute; TOP: 260px">
		<object data="scTableListScroll.asp" id="scTable" name="scScrollTable" style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 380px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Conditions Applied</strong>
	</span>
	
	<div id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 24px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="15%" class="TableHead">Condition Number</td>	
				<td width="70%" class="TableHead">Description</td>	
				<td width="15%" class="TableHead">Satisfied?</td> <% /* ASu Bmids00587 Change */%>
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
	</div>
	
	<span style="TOP: 200px; LEFT: 4px; POSITION: ABSOLUTE">
		<input id="btnSatisfy" value="Satisfy/Unsatisfy" type="button" style="WIDTH: 120px" class="msgButton">
	</span>
		
	<span style="TOP: 240px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Details
		<span style="TOP: 20px; LEFT: 0px; POSITION: ABSOLUTE">
			<TEXTAREA class=msgTxt id="txtDetails" name=Notes rows=5 style="POSITION: absolute; WIDTH: 595px" readonly></TEXTAREA>  
		</span>	
	</span>
	
	<span style="TOP: 350px; LEFT: 4px; POSITION: ABSOLUTE">
		<input id="btnAdd" value="Add" type="button" style="WIDTH: 75px" class="msgButton">
	</span>	
	
	<span style="TOP: 350px; LEFT: 90px; POSITION: ABSOLUTE">
		<input id="btnEdit" value="Edit" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
	
	<span style="TOP: 350px; LEFT: 176px; POSITION: ABSOLUTE">
		<input id="btnDelete" value="Delete" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
	
</div>
</form>

<%/* Main Buttons  */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 470px; WIDTH: 500px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/AP300Attribs.asp" -->

<script language="JScript" type="text/javascript">
<!--

var scScreenFunctions ;
var m_iTableLength = 10 ;
var ApplConditionsXML ;
var m_iUserRole = 0;
var m_iConditionsAuthorityLevel = 0 ;
var m_bIsUserAuthorityOk = false ;

var m_sUserId, m_sUnitId ;
var m_sApplicationNumber, m_sReadOnly, m_sIdXml,m_sApplicationFFNumber ;
var m_sTaskId, m_sTaskStatus, m_sTaskInstance, m_sStageId, m_sCaseActivityGuid, m_iCaseStageSequenceNo ;
var m_sTMConditions,m_sTMConditionsReview, m_sActivityId ;
<% //GD BMIDS00766 START %>
var m_blnConditionsReview = false;
<% //GD BMIDS00766 END %>
<% /* BS BM0521 11/06/03 */ %>
var m_sProcessInd = "";

function RetrieveContextData()
{
	<%/* TEST */ %>
	<% /*
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","A01201018");
	scScreenFunctions.SetContextParameter(window,"idReadOnly","0");
	scScreenFunctions.SetContextParameter(window,"idRole","40");
	scScreenFunctions.SetContextParameter(window,"idUserId","SRINI");
	scScreenFunctions.SetContextParameter(window,"idUnitId","UNIT1");
	scScreenFunctions.SetContextParameter(window,"idTaskXML", getCaseTaskFromContext());
	scScreenFunctions.SetContextParameter(window,"idDistributionChannelId","1"); */
	%>
	
	<% /*END TEST  */ %>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFFNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","");
	m_sReadOnly			 = scScreenFunctions.GetContextParameter(window,"idReadOnly","");
	m_iUserRole			 = scScreenFunctions.GetContextParameter(window,"idRole","0");
	m_sUserId			 = scScreenFunctions.GetContextParameter(window,"idUserId","");
	m_sUnitId			 = scScreenFunctions.GetContextParameter(window,"idUnitId","");
	m_sActivityId		 = scScreenFunctions.GetContextParameter(window,"idActivityId","");
	m_sIdXml			 = scScreenFunctions.GetContextParameter(window,"idTaskXML","");
	<% /* BS BM0521 11/06/03 */ %>
	m_sProcessInd = scScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null); 
}

<%/* Events */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{	
	<% // Added by automated update TW 09 Oct 2002 SYS5115 %>
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	<% // Added by automated update TW 09 Oct 2002 SYS5115 %>
	SetMasks();
	Validation_Init();
	var sButtonList = new Array("Submit", "Cancel", "Confirm");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Conditions Applied","AP300",scScreenFunctions);
	
	RetrieveContextData();
	InitialiseScreen() ;
	
	<% // Added by automated update TW 09 Oct 2002 SYS5115 %>
	ClientPopulateScreen();
}

function spnTable.onclick()
{
	var aSelectedRows = scScrollTable.getArrayofRowsSelected();
	
	if(aSelectedRows.length == 1) <%/* Single row has been selected */ %>
	{
		<%/* SR 23/04/01 : SYS2169 - get the ConditionDesc from the XML, not from the table.
		  if 11th (or greater) row is selected, referring to the respective row in the table might give an error  */
		%>
		aSelectedRows[0] = scScrollTable.getRowSelected();

		ApplConditionsXML.SelectTagListItem(tblTable.rows(aSelectedRows[0]).getAttribute("LISTSEQ"));
		frmScreen.txtDetails.value =  ApplConditionsXML.GetAttribute("CONDITIONDESCRIPTION");
		<% /* BS BM0521 11/06/03 
		If case is not locked and user is in a processing unit then enable buttons */ %>
		if(m_sReadOnly != "1" && m_sProcessInd == "1")
		{
			frmScreen.btnSatisfy.disabled = false ;
			if(m_bIsUserAuthorityOk)
			{
				frmScreen.btnDelete.disabled = false ;
				<%/* if condition selected is Editable, enable Edit button */%>
				if(tblTable.rows(aSelectedRows[0]).getAttribute("EDITABLE") == '1') frmScreen.btnEdit.disabled = false ;
				else frmScreen.btnEdit.disabled = true ;
			}
		}
	}
	else <% //No row(s) selected %>
	{
		frmScreen.txtDetails.value = '';
		frmScreen.btnEdit.disabled = true ;
		
		<% /* BS BM0521 11/06/03 
		If case is not locked and user is in a processing unit then enable buttons */ %>
		if(m_sReadOnly != "1" && m_sProcessInd == "1")
		{
			<% /* SYS2175 - Only enable delete if an item is selected. */ %>
			if((m_bIsUserAuthorityOk) && (aSelectedRows.length > 0))
			{
				frmScreen.btnDelete.disabled = false;
			}
			<% /* SYS2175 - End. */ %>
		
			<%/* if all Satisfy status of all the rows selectes is same, enable Satisfy/Unsatisfy button */ %>
			if(IsSatisfyStatusSame(aSelectedRows)) frmScreen.btnSatisfy.disabled = false ;
			else frmScreen.btnSatisfy.disabled = true ;
		}
	}
	
	<% //GD BMIDS00766 START %>
	if (m_blnConditionsReview==true)
	{
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
	}
	<% //GD BMIDS00766 END	 %>
}

function frmScreen.btnSatisfy.onclick()
{
	var iRowIndex, sNewStatusDesc, iNewStatusCode ;
	var aSelectedRows = scScrollTable.getArrayofRowsSelected();
	
	if(aSelectedRows.length == 0)
	{
		alert('Select a row.');
		return ;
	}

	<% /* on multi-row selection, this would be enabled only status in all the rows is same. */ %>
	if(tblTable.rows(aSelectedRows[0]).cells(2).innerText == 'Yes')
	{
		sNewStatusDesc = 'No' ;
		iNewStatusCode = 0 ;
	}
	else 
	{
		sNewStatusDesc = 'Yes' ;
		iNewStatusCode = 1 ;
	}
	
	for(iRowIndex = 0 ; iRowIndex < aSelectedRows.length ; ++iRowIndex)
	{
		scScreenFunctions.SizeTextToField(tblTable.rows(aSelectedRows[iRowIndex]).cells(2), sNewStatusDesc);
		ApplConditionsXML.SelectTagListItem(tblTable.rows(aSelectedRows[iRowIndex]).getAttribute("LISTSEQ"));
		ApplConditionsXML.SetAttribute("SATISFYSTATUS", iNewStatusCode);
		
		if(ApplConditionsXML.GetAttribute("SATISFYCONDITIONCHANGED") == '0')
			ApplConditionsXML.SetAttribute("SATISFYCONDITIONCHANGED", '1');
		else
			ApplConditionsXML.SetAttribute("SATISFYCONDITIONCHANGED", '0');			
	}
}

function frmScreen.btnAdd.onclick()
{
	if(m_iUserRole < m_iConditionsAuthorityLevel)
	{
		alert('You do not have authority to Add conditions to this application.');
		return ;
	}
	
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToAP310.submit();
}

function frmScreen.btnEdit.onclick()
{
	if(m_iUserRole < m_iConditionsAuthorityLevel)
	{
		alert('You do not have authority to Edit conditions on this application.');
		return ;
	}
	
	<% /* Set MetaAction and the current record data in the context and call AP310 */ %>
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");

	<% /* BMIDS635  Use getRowSelected to get correct row from scScrollTable  */ %>
	var iRowIndex = scScrollTable.getRowSelected() ;
	var sIdXml = ApplConditionsXML.ActiveTagList.item(tblTable.rows(iRowIndex).getAttribute("LISTSEQ")).xml ;
	scScreenFunctions.SetContextParameter(window,"idXml2", sIdXml);
	frmToAP310.submit();
}

function frmScreen.btnDelete.onclick()
{
	var aSelectedRows = scScrollTable.getArrayofRowsSelected();
	
	if(aSelectedRows.length == 0)
	{
		alert('Select a row.');
		return ;
	}

	if(m_iUserRole < m_iConditionsAuthorityLevel)
	{
		alert('You do not have authority to Delete conditions from this application.');
		return ;
	}

	if (confirm('Are you sure you want to delete the selected condition(s) from the application?'))
		if(DeleteConditions(aSelectedRows))		
		{
			scTable.clear();
			PopulateScreen();
			InitialiseButtonsState();
		}
}

function btnSubmit.onclick()
{
	<% /* Update all the modified application conditions, and route back to calling screen */ %>
	if(UpdateApplicationConditions()) frmToTM030.submit();
	else return ;
}

function btnCancel.onclick()
{
	frmToTM030.submit();
}

function btnConfirm.onclick()
{
	<% /* Call method omTmBO.UpdateApplicationConditions */ %>
	if(UpdateTaskAndApplicationCondtions()) frmToTM030.submit()
	else return ;
}

<% /* Functions */ %>
function InitialiseScreen()
{
	if(!GetParameterValues()) 
	{
		alert('Error retrieving User Authority Level');
		SetToReadOnlyProcessing();
	}
	
	<% /* PB 19/06/2006 EP683 Begin
	if(m_iConditionsAuthorityLevel < m_iUserRole) m_bIsUserAuthorityOk = true ; */ %>
	if(m_iConditionsAuthorityLevel <= m_iUserRole) m_bIsUserAuthorityOk = true ;
	<% /* EP683 End */ %>
	else m_bIsUserAuthorityOk = false ;
	
	if(!CreateDefaultConditions())
	{
		SetToReadOnlyProcessing();
		return ;
	}
	
	PopulateScreen();

	<% /* For read-only context, Disable Add, Edit, OK and Confirm buttons */ %>
	<% /* BS BM0521 11/06/03 */ %>
	//if(m_sReadOnly == "1") SetToReadOnlyProcessing();
	if(m_sReadOnly == "1" || m_sProcessInd != "1") SetToReadOnlyProcessing();
		
	scScrollTable.EnableMultiSelectTable();
	InitialiseButtonsState();
	
	<% /* if the task is complete, disable all the buttons except Cancel */ %>
	if(m_sTaskStatus == "40") SetToReadOnlyProcessing();
}

function CreateDefaultConditions()
{
	<%/* if the screen has been accessed by a task (i.e., idTaskXML has data peratining to case task),
	     create default conditions */ %> 
	
	if(m_sIdXml == '') return true ;
	
	var TaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var XML		= new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	m_sTMConditions = TaskXML.GetGlobalParameterString(document, 'TMConditions');
	m_sTMConditionsReview = TaskXML.GetGlobalParameterString(document, 'TMConditionsReview');	
	
	TaskXML.LoadXML(m_sIdXml)
	TaskXML.CreateTagList('CASETASK');
	
	<%/* If no case tasks, Do not create DefaultConditions */%>
	if(TaskXML.ActiveTagList.length == 0) return true ;
	
	TaskXML.SelectTagListItem(0);
	m_sTaskId			= TaskXML.GetAttribute("TASKID");
	m_sTaskStatus		= TaskXML.GetAttribute("TASKSTATUS");
	m_sTaskInstance		= TaskXML.GetAttribute("TASKINSTANCE");
	m_sCaseActivityGuid	= TaskXML.GetAttribute("CASEACTIVITYGUID");
	m_sStageId			= TaskXML.GetAttribute("STAGEID");
	m_iCaseStageSequenceNo = TaskXML.GetAttribute("CASESTAGESEQUENCENO");
	
	<% /* BMIDS00313 additional conditions review task */ %>
	if ((m_sTaskId != '') && (m_sTaskId == m_sTMConditionsReview) && (m_sTaskStatus == 10)) // BM0520 KRW 01/12/04
	{
		<% //GD BMIDS00766 START %>
		m_blnConditionsReview = true;
		<% //GD BMIDS00766 END %>
		
		<% //Disable Add, Edit and Delete Button. %>
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
		
		<% /* MV - 10/04/2003 - BM0520*/ %>
		var SetToPendingXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		SetToPendingXML.CreateRequestTag(window , "UpdateCaseTask");
		SetToPendingXML.CreateActiveTag("CASETASK");
		SetToPendingXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		SetToPendingXML.SetAttribute("ACTIVITYID", m_sActivityId);
		SetToPendingXML.SetAttribute("STAGEID", m_sStageId);
		SetToPendingXML.SetAttribute("CASESTAGESEQUENCENO", m_iCaseStageSequenceNo);
		SetToPendingXML.SetAttribute("CASEID", m_sApplicationNumber);
		SetToPendingXML.SetAttribute("TASKID", m_sTaskId);
		SetToPendingXML.SetAttribute("TASKINSTANCE", m_sTaskInstance);
		SetToPendingXML.SetAttribute("TASKSTATUS", "20"); <% /* Pending */ %>
			
		SetToPendingXML.RunASP(document, "msgTMBO.asp") ;
			
		if (!SetToPendingXML.IsResponseOK()) 
			return false;
		
		TaskXML.SetAttribute("TASKSTATUS", "20");
		scScreenFunctions.SetContextParameter(window,"idTaskXML", TaskXML.XMLDocument.xml);
		<% //NB. KEEP 'Satisfy' button enabled,Enable OK and Cancel and Confirm buttons %>
	} 
	else
	{
		<% /* for the tasks not started, delete any existing conditions and create new ones  */ %>
		if(m_sTaskId != '' && m_sTaskId == m_sTMConditions  && m_sTaskStatus == 10)
		{
			var xmlRequest = XML.CreateRequestTag(window , "CREATEDEFAULTAPPLICATIONCONDITIONS");
			<%/* Set UserAuthorityLevel as UserRole from Context. This is required for MsgTmBO.UpdateCaseTask */%>
			XML.SetAttribute("USERAUTHORITYLEVEL", m_iUserRole);
			
			XML.CreateActiveTag("APPLICATION");
			XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
					
			XML.ActiveTag = xmlRequest ;
			XML.CreateActiveTag("CASETASK");
			<% /* PSC 16/06/2006 MAR1873 - Start */ %>
			XML.SetAttribute('SOURCEAPPLICATION', 'Omiga');
			XML.SetAttribute('ACTIVITYID', m_sActivityId);
			<% /* PSC 16/06/2006 MAR1873 - End */ %>
			XML.SetAttribute('CASEACTIVITYGUID', m_sCaseActivityGuid);
			XML.SetAttribute('CASEID', m_sApplicationNumber);
			XML.SetAttribute('STAGEID', m_sStageId);
			XML.SetAttribute('TASKID', m_sTaskId);
			XML.SetAttribute('TASKINSTANCE', m_sTaskInstance); 
			XML.SetAttribute('CASESTAGESEQUENCENO', m_iCaseStageSequenceNo); 
			
			XML.RunASP(document,"OmigaTMBO.asp");
			if(!XML.IsResponseOK()) return false
			
			<% /* SR 27/03/01 SYS2111 : Update the Task Status in the context as Pending */ %>
			TaskXML.SetAttribute("TASKSTATUS", "20");
			scScreenFunctions.SetContextParameter(window,"idTaskXML", TaskXML.XMLDocument.xml);
			
		}
	}
	return true
}

function PopulateScreen()
{
	<% /* Call the appropriate method and fetch XML */ %>
	ApplConditionsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ApplConditionsXML.CreateRequestTag(window , "FINDAPPLICATIONCONDITIONSLIST");
	ApplConditionsXML.CreateActiveTag("APPLICATIONCONDITIONS");
	ApplConditionsXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	
	<% // 	ApplConditionsXML.RunASP(document,"omAppProc.asp"); %>
	<% // Added by automated update TW 09 Oct 2002 SYS5115 %>
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			ApplConditionsXML.RunASP(document,"omAppProc.asp");
			break;
		default: // Error
			ApplConditionsXML.SetErrorResponse();
		}

	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = ApplConditionsXML.CheckResponse(ErrorTypes);
	
	if (ErrorReturn[0] == true)
	{	<% /* Populate the List Box */ %>
		ApplConditionsXML.CreateTagList("APPLICATIONCONDITIONS");

		var iNumberOfRows = ApplConditionsXML.ActiveTagList.length;
		scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows);
		ShowList(0);
		if(iNumberOfRows > 0) 
		{
			scTable.setRowSelected(1) ;
			spnTable.onclick();
		}
	} else <% //No rows to populate %>
	{
		frmScreen.txtDetails.value = "";
	}
}

function ShowList(nStart)
{
	var iCount ;
	var sConditionReference, sConditionName, sSatisfyStatus, sConditionDesc ;
	var sEditable, sConditionSeq, sSatisfyCondChanged ;
	
	for(iCount = 0 ; iCount < ApplConditionsXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		ApplConditionsXML.SelectTagListItem(iCount + nStart);
		sConditionReference = ApplConditionsXML.GetAttribute("CONDITIONREFERENCE");
		sConditionName		= ApplConditionsXML.GetAttribute("CONDITIONNAME");
		sSatisfyStatus		= ApplConditionsXML.GetAttribute("SATISFYSTATUS");
		sEditable			= ApplConditionsXML.GetAttribute("EDIT");
		sConditionSeq		= ApplConditionsXML.GetAttribute("APPLNCONDITIONSSEQ");
		
		if(sSatisfyStatus == '1') sSatisfyStatus = 'Yes' ;
		else if(sSatisfyStatus == '0')sSatisfyStatus = 'No' ;
		
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sConditionReference);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sConditionName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sSatisfyStatus);
		
		<%/* Add a new attrib 'SATISFYCONDITIONCHANGED' to each element in XML. Initialise this to '0' and 
			 update it when the user changes the status. Possible values are 0,1 */ 
		%>
		sSatisfyCondChanged = ApplConditionsXML.GetAttribute("SATISFYCONDITIONCHANGED");
		if(sSatisfyCondChanged == '') ApplConditionsXML.SetAttribute("SATISFYCONDITIONCHANGED", '0');
		
		<% /* Add the required attributes	*/ %>
		tblTable.rows(iCount+1).setAttribute("EDITABLE", sEditable);
		tblTable.rows(iCount+1).setAttribute("CONDITIONSEQ", sConditionSeq);
		tblTable.rows(iCount+1).setAttribute("SATISFYCONDITIONCHANGED", "0");
		tblTable.rows(iCount+1).setAttribute("SATISFYSTATUSCODE", ApplConditionsXML.GetAttribute("SATISFYSTATUS"));
		tblTable.rows(iCount+1).setAttribute("LISTSEQ", iCount + nStart);
	}	
}

function DeleteConditions(aSelectedRows)
{
	if(aSelectedRows.length == 0) 
	{	
		alert('No rows selected for deletion');
		return false ;
	}
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequest = XML.CreateRequestTag(window , "DELETEAPPLICATIONCONDITIONS");
	
	for (var iCount = 0; iCount < aSelectedRows.length ; iCount++)
	{
		XML.CreateActiveTag("APPLICATIONCONDITIONS");
		XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.SetAttribute("APPLNCONDITIONSSEQ", tblTable.rows(aSelectedRows[iCount]).getAttribute("CONDITIONSEQ"));
		<%  /* SR 26/03/01 SYS2111: Add Condition Ref to the request for deletion   */ %>
		XML.SetAttribute("CONDITIONREFERENCE", tblTable.rows(aSelectedRows[iCount]).cells(0).innerText);

		XML.ActiveTag = xmlRequest;
	}
	
	<% // 	XML.RunASP(document,"omAppProc.asp"); %>
	<% // Added by automated update TW 09 Oct 2002 SYS5115 %>
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"omAppProc.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	if(!XML.IsResponseOK()) return false
	
	return true ;
}

function UpdateApplicationConditions()
{
	var iChangeCount = 0 ;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequest = XML.CreateRequestTag(window , "UPDATEAPPLICATIONCONDITIONS");
	if (ApplConditionsXML.ActiveTagList != null)
	{
		for(var iCount = 0 ; iCount < ApplConditionsXML.ActiveTagList.length ; ++iCount)
		{
			ApplConditionsXML.SelectTagListItem(iCount);
			
			if(ApplConditionsXML.GetAttribute("SATISFYCONDITIONCHANGED") == "1")
			{
				++iChangeCount ;
				XML.CreateActiveTag("APPLICATIONCONDITIONS");
				XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
				XML.SetAttribute("APPLNCONDITIONSSEQ", ApplConditionsXML.GetAttribute("APPLNCONDITIONSSEQ"));
				XML.SetAttribute("CONDITIONREFERENCE", ApplConditionsXML.GetAttribute("CONDITIONREFERENCE"));
				XML.SetAttribute("SATISFYSTATUS", ApplConditionsXML.GetAttribute("SATISFYSTATUS"));
				XML.SetAttribute("USERID", m_sUserId);
				XML.SetAttribute("UNITID", m_sUnitId);
				XML.SetAttribute("USERMODIFIED","1");
				XML.ActiveTag = xmlRequest ;
			}
		}
	}
	if(iChangeCount > 0) 
	{
		<% // 		XML.RunASP(document,"omAppProc.asp"); %>
		<% // Added by automated update TW 09 Oct 2002 SYS5115 %>
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"omAppProc.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		
		if(!XML.IsResponseOK()) return false ;
		else return true ;
	}
	else return true ;
}

function UpdateTaskAndApplicationCondtions()
{
	var iChangeCount = 0 ;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequest = XML.CreateRequestTag(window , "UPDATEAPPLICATIONCONDITIONS");

	<% /* Add CaseTask element */ %>
	XML.CreateActiveTag("CASETASK");
	XML.SetAttribute("CASEACTIVITYGUID", m_sCaseActivityGuid);
	XML.SetAttribute("STAGEID", m_sStageId);
	XML.SetAttribute("TASKID", m_sTaskId);
	XML.SetAttribute("TASKINSTANCE", m_sTaskInstance);
	XML.SetAttribute("CASESTAGESEQUENCENO", m_iCaseStageSequenceNo);
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("CASEID", m_sApplicationNumber);
	XML.SetAttribute("ACTIVITYID", m_sActivityId);
	XML.SetAttribute("ACTIVITYINSTANCE", "1");
	
	XML.ActiveTag = xmlRequest ;
	
	for(var iCount = 0 ; ApplConditionsXML.ActiveTagList != null && iCount < ApplConditionsXML.ActiveTagList.length ; ++iCount)
	{
		ApplConditionsXML.SelectTagListItem(iCount);
		
		if(ApplConditionsXML.GetAttribute("SATISFYCONDITIONCHANGED") == "1")
		{
			++iChangeCount;
			XML.CreateActiveTag("APPLICATIONCONDITIONS");
			XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.SetAttribute("APPLNCONDITIONSSEQ", ApplConditionsXML.GetAttribute("APPLNCONDITIONSSEQ"));
			XML.SetAttribute("CONDITIONREFERENCE", ApplConditionsXML.GetAttribute("CONDITIONREFERENCE"));
			XML.SetAttribute("SATISFYSTATUS", ApplConditionsXML.GetAttribute("SATISFYSTATUS"));
			XML.SetAttribute("USERID", m_sUserId);
			XML.SetAttribute("UNITID", m_sUnitId);
			XML.SetAttribute("USERMODIFIED","1");
			XML.ActiveTag = xmlRequest ;
		}
	}	
	
	<% // 	XML.RunASP(document,"OmigaTMBO.asp"); %>
	<% // Added by automated update TW 09 Oct 2002 SYS5115 %>
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"OmigaTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

		
	if(!XML.IsResponseOK()) return false ;
	else return true ;
}

function InitialiseButtonsState()
{
	frmScreen.btnEdit.disabled = true ;
	frmScreen.btnDelete.disabled = true ;
	
	<% /* if user does not have sufficient authority, disable Add  button also */ %>
	if(!m_bIsUserAuthorityOk) frmScreen.btnAdd.disabled = true ;
}

function IsSatisfyStatusSame(aSelectedRows)
{
	<% /* SR 23/04/01 : SYS2169 - take the Satisfy status from XML. if the 11th (or greater) row is selected,
					    referring to a respective row in table might raise an error  */
	%>
	var sSatisfy, iCount ;
	ApplConditionsXML.SelectTagListItem(aSelectedRows[0]-1);
	sSatisfy = 	ApplConditionsXML.GetAttribute("SATISFYSTATUS") ; 
	for(iCount=1; iCount < aSelectedRows.length ; iCount++)
	{
		ApplConditionsXML.SelectTagListItem(aSelectedRows[iCount]-1);
		if(sSatisfy != ApplConditionsXML.GetAttribute('SATISFYSTATUS')) return false
	}	
	
	return true ;
}

function GetParameterValues()
{
	var blnReturn = false;	
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var XMLActiveTag = XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("GLOBALPARAMETER");
	XML.CreateTag("NAME","CONDITIONSAUTHORITYLEVEL");
	XML.ActiveTag = XMLActiveTag;
	
	XML.RunASP(document,"FindCurrentParameterList.asp");

	if (XML.IsResponseOK())
	{	
		blnReturn = true;
		if(XML.SelectTag(null,"GLOBALPARAMETERLIST") != null)
		{
			XML.CreateTagList("GLOBALPARAMETER");
			XML.SelectTagListItem(0);
			m_iConditionsAuthorityLevel = XML.GetTagText("AMOUNT");
		}		
	}			
	else blnReturn = false;

	XML = null; 
	return blnReturn;
}

function SetToReadOnlyProcessing()
{
	frmScreen.btnAdd.disabled = true ;
	frmScreen.btnEdit.disabled = true ;
	frmScreen.btnDelete.disabled = true ;
	frmScreen.btnSatisfy.disabled = true ;
	DisableMainButton('Submit') ;
	DisableMainButton('Confirm');
}
<%
/*
function getCaseTaskFromContext()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	XML.CreateActiveTag("CASETASK");
	XML.SetAttribute("STAGEID", "70");
	XML.SetAttribute("TASKID", "Apply_Spec_Cond");
	XML.SetAttribute("CASESTAGESEQUENCENO", "7");
	XML.SetAttribute("TASKINSTANCE", "1");
	XML.SetAttribute("CASEACTIVITYGUID", "BA3BE202E60F11D48270001020C01877");
	XML.SetAttribute("TASKSTATUS", "10");
	
	return XML.XMLDocument.xml ;
}
*/
%>
<% // BMIDS00746 START %>
function spnTable.ondblclick()
{
	if (tblTable.rows(scScrollTable.getRowSelected()).getAttribute("LISTSEQ") != null)
	{
		frmScreen.btnSatisfy.onclick();
	}
}
<% // BMIDS00746 END %>



-->
</script>
</body>
</html>


