<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      TM031.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Select Task to add (POP-UP SCREEN)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		26/10/00	Created (screen paint)
JLD		22/11/00	Added calls to BO
JLD		27/11/00	removed debug
JLD		14/02/01	Change to GetAttribute() method.
SR		15/03/01	SYS1840
SR		25/04/01	SYS1840
JLD		19/12/01	SYS3362 when generating a task name use the 
                    existing taskid not the taskname. This then matches the auto-generated name
LD		23/05/02	SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDSHistory:

Prog	Date		Description
ASu		30/09/2002	BMIDS00469 - Where the confirm button has not been selected before comitting the screen test for null
SA		24/10/2002  BMIDS00687 - prevent clicking anywhere in white space to cause error - spnTable.onClick()
SA		28/10/2002	BMIDS00737 - Use new method to populate screen.
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
DB		31/03/2003	BM0248 - Send in mandatoryflag when adding ad-hoc task.
HMA     02/09/2004  BMIDS863   - Enable Screen Rules.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:

Prog	Date		Description
HMA     12/12/2005  MAR667  Implement UnderwritingTasksOnly flag.
JD		27/01/2006	MAR891	If there are no additional tasks to add don't show error
PSC		01/03/2006	MAR1341	Add CHASINGPERIODMINUTES
DRC		17/03/2006  MAR1350 Due date for Post Completion tasks
HMA     17/03/2006  MAR1449 Correct filtering of tasks.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ESPOM History:

AW		27/09/06	EP1160	Further amendments for editable tasks
HMA     02/01/07    EP2_576 E2CR67 Identify applicants and guarantors on tasks.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
TL		07/09/2005	   Added Customer Task Processing. - MAR23
PSC		03/10/2005	   MAR32 Added print funtionality for ALWAYSAUTOMATICONCREATION
AW		14/09/2006	   EP1103   CC78 Changes when called from TM040
AW		14/09/2006	   EP1103	Further amendments	
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Select Tasks to Add  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 205px; LEFT: 210px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange">
<div id="divTasks" style="TOP: 10px; LEFT: 10px; HEIGHT: 230px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
	Double click the task to add
</span>
<span id="spnTable" style="TOP: 33px; LEFT: 4px; POSITION: ABSOLUTE">
	<table id="tblTable" width="495px" border="0" cellspacing="0" cellpadding="0" class="msgTable">
	<tr id="rowTitles">
		<td width="90%" class="TableHead">Tasks</td>	
		<td width="10%" class="TableHead">Add?</td>
	</tr>
	<tr id="row01">		
		<td class="TableTopLeft">&nbsp;</td>		
		<td class="TableTopRight">&nbsp;</td>
	</tr>
	<tr id="row02">		
		<td class="TableLeft">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row03">		
		<td class="TableLeft">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row04">		
		<td class="TableLeft">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row05">		
		<td class="TableLeft">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row06">		
		<td class="TableLeft">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row07">		
		<td class="TableLeft">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row08">		
		<td class="TableLeft">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row09">		
		<td class="TableLeft">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row10">		
		<td class="TableBottomLeft">&nbsp;</td>
		<td class="TableBottomRight">&nbsp;</td>
	</tr>
	</table>
</span>
</div>
<div id="divTaskname" style="TOP: 230px; LEFT: 10px; HEIGHT: 65px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
		Task Name
		<span style="TOP:-3px; LEFT:65px; POSITION:ABSOLUTE">
			<input id="txtTaskName" maxlength="100" style="WIDTH:430px" class="msgTxt">
		</span>
	</span>
</div>
<div id="divApplicant" style="TOP: 258px; LEFT: 10px; HEIGHT: 65px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Applicant
	<span style="TOP:-3px; LEFT:65px; POSITION:ABSOLUTE">
		<select id="cboApplicant" style="WIDTH:250px" class="msgCombo"></select>
	</span>
</span>
<span style="TOP:33px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Organisation
	<span style="TOP:-3px; LEFT:65px; POSITION:ABSOLUTE">
		<select id="cboOrganisation" style="WIDTH:250px" class="msgCombo"></select>
	</span>
</span>
<span style="TOP:30px; LEFT:375px; POSITION:ABSOLUTE">
	<input id="btnConfirm" value="Confirm" type="button" style="WIDTH:60px" class="msgButton">
</span>
</div>
<div id="divButtons" style="TOP: 325px; LEFT: 10px; HEIGHT: 35px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:5px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnOK" value="OK" type="button" style="WIDTH:60px" class="msgButton">
</span>
<span style="TOP:5px; LEFT:70px; POSITION:ABSOLUTE">
	<input id="btnCancel" value="Cancel" type="button" style="WIDTH:60px" class="msgButton">
</span>
</div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/TM031Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_sStageId = "";
var m_sApplicationNo = "";
var m_sApplicationPriority ;
var m_iUserRole = 0 ;
var m_sActivityId = "";
var taskXML = null;
var m_iTableLength = 10;
var m_asFlagArray = null;
var m_asApplicantArray = null;
var m_asOrganisationArray = null ;
var m_asFlagArray = null ;
var m_sRequestAttributes = "";
var m_sSeqNo = "";
var m_bIsChanged = false ;
var m_sTMReIssueOffer = "";
var m_BaseNonPopupWindow = null;
var m_bUnderwritingTasksOnly = false;    // MAR667
<% /* AW   14/09/06 EP1103 */ %>
var m_bProgressTaskMode = false;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	RetrieveData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	SetScreenOnReadOnly();
	//if(!InitialiseScreen()) return ; JD MAR891
	if(InitialiseScreen())
	{
		PopulateScreen();
		window.returnValue = null;
	
		// Added by automated update TW 09 Oct 2002 SYS5115
		ClientPopulateScreen();
	}
	else
	{
		frmScreen.btnOK.disabled = true;
		frmScreen.btnConfirm.disabled = true;
		frmScreen.cboApplicant.disabled = true;
		frmScreen.cboOrganisation.disabled = true;
	}
}

function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnOK.disabled = true;
		frmScreen.btnCancel.focus();
	}
}

function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop		= sArguments[0];
	window.dialogLeft		= sArguments[1];
	window.dialogWidth		= sArguments[2];
	window.dialogHeight		= sArguments[3];
	var sParameters			= sArguments[4];
	m_BaseNonPopupWindow	= sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sReadOnly				= sParameters[0];
	m_sStageId				= sParameters[1];
	m_sApplicationNo		= sParameters[2];
	m_sApplicationPriority	= sParameters[3];
	m_iUserRole				= sParameters[4];
	m_sActivityId			= sParameters[5];
	m_sAppXML				= sParameters[6];
	m_sRequestAttributes	= sParameters[7];
	m_sSeqNo				= sParameters[8];
	m_bUnderwritingTasksOnly = sParameters[9];
	<% /* AW   14/09/06 EP1103 */ %>
	m_bProgressTaskMode		= sParameters[10];
}

function PopulateScreen()
{
	scScreenFunctions.SetCollectionState(divApplicant, "D");
	frmScreen.btnConfirm.disabled = true;
	<% /* AW   27/09/06 EP1160 */ %>
	scScreenFunctions.HideCollection(divTaskname);
	PopulateListBox() ;
	
	frmScreen.btnOK.disabled = true;
}

function InitialiseScreen()
{
	var bReturn = true;
	PopulateApplicantCombo();
	
	taskXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	//BMIDS00737 SA 28/10/02 Use new method to populate screen,
	//taskXML.CreateRequestTagFromArray(m_sRequestAttributes, "GetStageTaskDetail");
	<% /* AW   14/09/06 EP1103 */ %>
	if (!m_bProgressTaskMode)
	{
		taskXML.CreateRequestTagFromArray(m_sRequestAttributes, "GetStageAddtlTaskDetail");
	}
	else
	{
		taskXML.CreateRequestTagFromArray(m_sRequestAttributes, "GetEditableTaskDetailList");
	}
	
	taskXML.CreateActiveTag("STAGETASK");
	taskXML.SetAttribute("STAGEID", m_sStageId);
	taskXML.SetAttribute("ADDITIONALFLAG", "1");
	taskXML.SetAttribute("DELETEFLAG", "0");
	taskXML.RunASP(document, "MsgTMBO.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = taskXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{
		<% /* JD MAR891 warn if there are no tasks to add */ %>
		alert("No additional tasks have been identified for this stage.");
		bReturn = false;
	}
	else
	{
		taskXML.ActiveTag = null;
		taskXML.CreateTagList("STAGETASK");
		
		FilterTasks();
		
		InitialiseArrays();
	}
	
	return bReturn;
}

function PopulateApplicantCombo()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ComboXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.LoadXML(m_sAppXML);
	XML.SelectTag(null, "REQUEST");
	XML.CreateTagList("CUSTOMER");
	
	var TagOPTION	= document.createElement("OPTION");
	TagOPTION.value	= "" ;
	TagOPTION.text	= "<SELECT>" ;
	frmScreen.cboApplicant.add(TagOPTION);
	
	for(var nLoop = 0; nLoop < XML.ActiveTagList.length; nLoop++)
	{
		XML.SelectTagListItem(nLoop);
		var sCustomerName = XML.GetAttribute("NAME");
		var sCustomerNumber = XML.GetAttribute("NUMBER");
		var sCustomerVersionNumber = XML.GetAttribute("VERSION");
		var sCustomerRole = XML.GetAttribute("ROLE");
		
		<% /* EP2_576 Add the customer role to the customer name text */ %>
		if(sCustomerName != "" && sCustomerNumber != "")
		{
			TagOPTION		= document.createElement("OPTION");
			TagOPTION.value	= sCustomerNumber;
			
			if (ComboXML.IsInComboValidationList(document,"CustomerRoleType",sCustomerRole,["G"]))
			{
				sCustomerName = sCustomerName + " (Guarantor)";
			}
			else
			{
				sCustomerName = sCustomerName + " (Applicant)";
			}
			TagOPTION.text	= sCustomerName;
			
			TagOPTION.setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
			frmScreen.cboApplicant.add(TagOPTION);
		}
	}
}

function PopulateListBox()
{
	taskXML.ActiveTag = null;
	taskXML.CreateTagList("STAGETASK");

	var iNumberOfTasks = taskXML.ActiveTagList.length;
	if(iNumberOfTasks == 0)alert("No additional tasks have been identified for this stage.");
	else
	{
		scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfTasks);
		ShowList(0);	
	}
}

function ShowList(nStart)
{
	var iCount, sTaskName, sCaseTaskName;
	
	scScrollTable.clear();

	for (iCount = 0; iCount < taskXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		taskXML.SelectTagListItem(iCount + nStart);
		
		<%/* If the task was already added, show the description that is specific to this case */%>
		sTaskName = taskXML.GetAttribute("TASKNAME") ;
		sCaseTaskName = taskXML.GetAttribute("CASETASKNAME") ;
		if(sCaseTaskName != '') scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sCaseTaskName);
		else scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sTaskName);
		
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),GetAddedFlag(iCount+nStart));
		tblTable.rows(iCount+1).setAttribute("TaskId", taskXML.GetAttribute("TASKID"));
		tblTable.rows(iCount+1).setAttribute("ContactType", taskXML.GetAttribute("CONTACTTYPE"));
		tblTable.rows(iCount+1).setAttribute("Applicant", taskXML.GetAttribute("APPLICANT"));
		tblTable.rows(iCount+1).setAttribute("ApplicantDetailsAdded", taskXML.GetAttribute("APPLICANTDETAILSADDED"));
		tblTable.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));
	}
}

var bNeedsCustomer; <%/* MAR23 - TL 07/09/2005*/%>

function spnTable.ondblclick()
{
	<% /*On double click, toggle the 'Add to Stage?' cell between 'Yes' and 'No' */ %>
	taskXML.ActiveTag = null;
	taskXML.CreateTagList("STAGETASK");
	if(taskXML.ActiveTagList.length > 0)
	{
		var iRowSelected = scScrollTable.getRowSelected();
		<%/*BMIDS00687 SA Check that the row is actually selected.*/%>
		if (iRowSelected >= 0)
		{
			var nTaskItem = parseInt(tblTable.rows(iRowSelected).getAttribute("TagListItem"));
		
			taskXML.SelectTagListItem(nTaskItem);

			var bNeedsApplicant;
			var bEditableTask = false;
			
			if(taskXML.GetAttribute("APPLICANT") == "1") bNeedsApplicant = true;
			else bNeedsApplicant = false;
	
			<%/* MAR23 - TL 07/09/2005*/%>
			if(!bNeedsApplicant && taskXML.GetAttribute("CUSTOMERTASK") == "1") bNeedsCustomer = true;
			else bNeedsCustomer = false;
	
			if(taskXML.GetAttribute("EDITABLETASKIND") == "1") bEditableTask = true;
			
			if(GetAddedFlag(nTaskItem) == "No")
			{
				scScreenFunctions.SizeTextToField(tblTable.rows(iRowSelected).cells(1),"Yes");
				frmScreen.btnOK.disabled = false ;
				<% /* if the applicant reference has already been added, do not prompt user  */ %>
				if(parseInt(tblTable.rows(iRowSelected).getAttribute("ApplicantDetailsAdded")) != 1)
				{
					<%/* When Yes is selected, if the task has an applicant reference, the applicant must be chosen */ %>
					if(bNeedsApplicant)
					{
						alert("Please select an Applicant and Organisation to associate with this task");
						ChooseApplicant();
					}
					else if(bNeedsCustomer) <%/* MAR23 - TL 07/09/2005*/%>
					{
						alert("Please select an Applicant to associate with this task");
						ChooseApplicant();
					}
					<% /* AW   14/09/06 EP1103 */ %>
					else if(bEditableTask && frmScreen.txtTaskName.value == '')<% /* AW   27/09/06 EP1160 */ %>
					{
						alert('Please enter a task name.');
						frmScreen.btnConfirm.disabled = false;
						frmScreen.btnOK.disabled = true ;
						frmScreen.txtTaskName.focus();
					}
					else SetAddedFlag(nTaskItem, "Yes") <%/* Set the Added Flag in the array */ %>
				}
				else SetAddedFlag(nTaskItem, "Yes")
			}
			else
			{
				scScreenFunctions.SizeTextToField(tblTable.rows(iRowSelected).cells(1),"No");
				SetAddedFlag(nTaskItem, "No");
			}
		}	
	}
}

function spnTable.onclick()
{
<%	//on selecting a row, if the task is to be added to the stage and the task needs
	//to be associated with an applicant, enable the Set Applicant btn
%>	taskXML.ActiveTag = null;
	taskXML.CreateTagList("STAGETASK");
	if(taskXML.ActiveTagList.length > 0)
	{
		var iRowSelected = scScrollTable.getRowSelected();
		<%/*BMIDS00687 SA Check that the row is actually selected.*/%>
		if (iRowSelected >= 0)
		{
			var nTaskItem = parseInt(tblTable.rows(iRowSelected).getAttribute("TagListItem"));
			taskXML.SelectTagListItem(nTaskItem);
			
				<% /* AW   27/09/06 EP1160 */ %>
				if( taskXML.GetAttribute("EDITABLETASKIND") == "1" )
				{
					scScreenFunctions.ShowCollection(divTaskname);
				}
				else
				{
					scScreenFunctions.HideCollection(divTaskname);
				}
		}
	}
	<%/*BMIDS00687 SA Check that the row is actually selected.*/%>
	if (iRowSelected >= 0)
	{
		if(parseInt(tblTable.rows(iRowSelected).getAttribute("Applicant")) != "1")
			scScreenFunctions.SetCollectionState(divApplicant, "D");
	}
}

function ChooseApplicant()
{
	ClearOrganisationCombo();
	frmScreen.cboApplicant.value = '';
	
	scScreenFunctions.SetCollectionState(divApplicant, "W");
	if (bNeedsCustomer)
		frmScreen.cboOrganisation.disabled = true ;
	frmScreen.btnConfirm.disabled = false;
	scScreenFunctions.SetCollectionState(divTasks, "D");
	scScreenFunctions.SetCollectionState(divButtons, "D");
	frmScreen.cboApplicant.focus();
}

function frmScreen.cboApplicant.onchange()
{
	if (bNeedsCustomer)
		return;
		
	if(frmScreen.cboApplicant.value != "")
	{
		ClearOrganisationCombo();
		if(PopulateOrganisationCombo())
		{
			frmScreen.cboOrganisation.disabled = false ;
			frmScreen.cboOrganisation.focus();
		}
		else  <% /* SR: SYS1840 - Change the status in list box to NO */ %>
		{
			var iRowSelected = scScrollTable.getRowSelected();
			var nTaskItem = parseInt(tblTable.rows(iRowSelected).getAttribute("TagListItem"));
			scScreenFunctions.SizeTextToField(tblTable.rows(iRowSelected).cells(1),"No");
			SetAddedFlag(nTaskItem, "No");
			frmScreen.cboOrganisation.disabled = true ;
		}	
	}
}

function frmScreen.btnConfirm.onclick()
{
	<% // HMA  BMIDS863 Add Screen Rules %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	
	<% /* AW   14/09/06 EP1103 */ %>
	var iRowSelected = scScrollTable.getRowSelected();
	var sTaskItem = tblTable.rows(iRowSelected).getAttribute("TagListItem")
	
	taskXML.ActiveTag = null;
	taskXML.CreateTagList("STAGETASK");
	taskXML.SelectTagListItem(sTaskItem);
	
		<% /* AW   27/09/06 EP1160  - Start */ %>
	var bEditableTask = false;
	if (taskXML.GetAttribute("EDITABLETASKIND") == "1") bEditableTask = true;
	
	if (bEditableTask)
	{
		if(frmScreen.txtTaskName.value == '')
		{
			alert('Please enter a task name.');
			frmScreen.txtTaskName.focus();
			return ;
		}
	}
	else
	{
		if(!bNeedsCustomer && (frmScreen.cboApplicant.value == '' || frmScreen.cboOrganisation.value == ''))
		{
			alert('Choose applicant and an Organisation and confirm.');
			return ;
		}	
	}
	<% /* AW   27/09/06 EP1160 - End*/ %>
		
	var iAppSelected = frmScreen.cboApplicant.selectedIndex;
	var iOrgSelected;
	if(!bNeedsCustomer)
		iOrgSelected = frmScreen.cboOrganisation.selectedIndex;
	
	var xmlTempNode = taskXML.ActiveTag.cloneNode(true) ;
	
	var sOldTaskName = taskXML.GetAttribute("TASKNAME");
	
	var sNewTaskName;
	<% /* AW   14/09/06 EP1103 */ %>
	<% /* AW   27/09/06 EP1160 -  Start*/ %>
	if(bEditableTask)
	{
		sNewTaskName = frmScreen.txtTaskName.value;
	}
	else 
	{
		if(!bNeedsCustomer)
		{
			sNewTaskName = frmScreen.cboOrganisation.options(iOrgSelected).text + " " +
						   sOldTaskName + " for " + frmScreen.cboApplicant.options(iAppSelected).text;
		}
		else
		{
			sNewTaskName = sOldTaskName + " for " + frmScreen.cboApplicant.options(iAppSelected).text;
		}
	}
	<% /* AW   27/09/06 EP1160  - End */ %>
	if (bEditableTask && sNewTaskName.length > 100)
	{
		alert("New task name: "  + "'" + sNewTaskName + "'" + " exceeds 100 characters.");
		return ;
	}
	
	taskXML.SetAttribute("CASETASKNAME", sNewTaskName);
	taskXML.SetAttribute("APPLICANTDETAILSADDED", "1");
	
	taskXML.XMLDocument.documentElement.insertBefore(xmlTempNode, taskXML.ActiveTag);
	
	<%/* Add a new element in the arrays that maintain the other details of each row   */%>
	m_asApplicantArray = AddToArray(m_asApplicantArray, parseInt(tblTable.rows(iRowSelected).getAttribute("TagListItem"))+1, frmScreen.cboApplicant.value);
	if(!bNeedsCustomer)
		m_asOrganisationArray = AddToArray(m_asOrganisationArray, parseInt(tblTable.rows(iRowSelected).getAttribute("TagListItem"))+1, frmScreen.cboOrganisation.value);
	m_asFlagArray = AddToArray(m_asFlagArray, parseInt(tblTable.rows(iRowSelected).getAttribute("TagListItem"))+1, "Yes");
	
	PopulateScreen() ;
	ClearOrganisationCombo();
	scScreenFunctions.SetCollectionState(divApplicant, "D");
	frmScreen.btnConfirm.disabled = true;
	scScreenFunctions.SetCollectionState(divTasks, "W");
	scScreenFunctions.SetCollectionState(divButtons, "W");
	frmScreen.btnOK.disabled = false ;
}

function PopulateOrganisationCombo()
{
	var iRowSelected = scScrollTable.getRowSelected();
	var nTaskItem = parseInt(tblTable.rows(iRowSelected).getAttribute("TagListItem"));
	taskXML.ActiveTag = null;
	taskXML.CreateTagList("STAGETASK");
	taskXML.SelectTagListItem(nTaskItem);
	
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_sRequestAttributes, "");
	XML.CreateActiveTag("THIRDPARTY");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", "1");  
	XML.CreateTag("CUSTOMERNUMBER", frmScreen.cboApplicant.value); 
	var TagOption = frmScreen.cboApplicant.options(frmScreen.cboApplicant.selectedIndex) ;	
	XML.CreateTag("CUSTOMERVERSIONNUMBER", TagOption.getAttribute('CustomerVersionNumber')); 
	XML.CreateTag("THIRDPARTYTYPE", taskXML.GetAttribute("CONTACTTYPE"));
	
	XML.RunASP(document,"FindThirdPartyForCustomer.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{
		alert('No Third Parties could be found for this customer');
		return false;
	}

	if(ErrorReturn[0] == true)
	{
		XML.ActiveTag = null;
		XML.CreateTagList("THIRDPARTY");
		
		<% /* When no ThirdParty records are returned (this might not have been returned as Success by the methods),
		      return fasle */ %>
		if(XML.ActiveTagList.length == 0)
		{
			alert('No Organisations could be found for this Applicant');
			return false ;
		}
		
		for(var iCount = 0; iCount < XML.ActiveTagList.length; iCount++)
		{
			<% /* add COMPANYNAME to the combo */ %>
			XML.SelectTagListItem(iCount);
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= XML.GetTagText("CONTEXT");
			TagOPTION.text = XML.GetTagText("NAME");
			frmScreen.cboOrganisation.add(TagOPTION);
		}
	}
	
	return true ;
}

function ClearOrganisationCombo()
{
	if(frmScreen.cboOrganisation.length >= 1)
		for(var iCount = frmScreen.cboOrganisation.length-1 ; iCount >= 0; --iCount)
		{
			frmScreen.cboOrganisation.remove(iCount) ;						
		}
}

function InitialiseArrays()
{
	m_asFlagArray = new Array();
	m_asApplicantArray = new Array();
	m_asOrganisationArray = new Array();
	for(var iTask = 0; iTask < taskXML.ActiveTagList.length; iTask++)
	{
		m_asFlagArray[iTask] = "No";
		m_asApplicantArray[iTask] = "";
		m_asOrganisationArray[iTask] = "";
	}
}

function SetAddedFlag(nTaskItem, sValue)
{
<%	//record for each task in the list whether it is to be added to the stage.
	//Also if any task is selected, enable the OK button. If no tasks are selected, disable it
%>	m_asFlagArray[nTaskItem] = sValue;
	var sDisableOK = true;
	for(var iTask = 0; iTask < m_asFlagArray.length; iTask++)
	{
		if(m_asFlagArray[iTask] == "Yes") sDisableOK = false;
	}
	frmScreen.btnOK.disabled = sDisableOK;
}

function GetAddedFlag(nTaskItem)
{
	return (m_asFlagArray[nTaskItem])
}

function GetApplicant(nTaskItem)
{
	return (m_asApplicantArray[nTaskItem])
}

function SetApplicant(nTaskItem, sValue)
{
	m_asApplicantArray[nTaskItem] = sValue;
}

function AddToArray(saArray, nPosition, sValue)
{ <%/* Add a new element to the array passed in, at the position 'nPosition' with a value 'sValue' */ %>
	
	var saFirstArray = saArray.slice(0, nPosition);
	var saSecondArray = saArray.slice(nPosition);
	
	saArray = saFirstArray.concat(sValue, saSecondArray);
	
	return saArray ;
}

function GetOrganisation(nTaskItem)
{
	return (m_asOrganisationArray[nTaskItem])
}

function SetOrganisation(nTaskItem, sValue)
{
	m_asOrganisationArray[nTaskItem] = sValue;
}
<% /* PSC 01/03/2006 MAR1341 */ %>				
function GetNewDueDate(nDays,nHours, nMinutes, strTaskType)
{
	<% /* MO - BMIDS00807 */ %>
	var dtBase = scScreenFunctions.GetAppServerDate(true);
	<% /* DRC MAR1350 - Check for Post Completion task */ %>
	var strCompletionDate = "";
	var TempXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	if (TempXML.IsInComboValidationList(document,"TaskType",strTaskType,["PC"]))
	  	strCompletionDate = GetCompletionDate();
    if 	(strCompletionDate.length > 0)
         dtBase = scScreenFunctions.GetDateObjectFromString(strCompletionDate);
    <% /* DRC MAR1350 - End */ %>

	<% /* var dtToday = new Date(); */ %>
	<% /* PSC 01/03/2006 MAR1341 */ %>				
	var dtNewDate = new Date(dtBase.getYear(), dtBase.getMonth(), 
	                         dtBase.getDate() + nDays, 
	                         dtBase.getHours() + nHours,
	                         dtBase.getMinutes() + nMinutes, 
	                         dtBase.getSeconds());
	return(dtNewDate.getDate() + "/" + (dtNewDate.getMonth() + 1) + "/" + dtNewDate.getYear() + " "
	       + dtNewDate.getHours() + ":" + dtNewDate.getMinutes() + ":" + dtNewDate.getSeconds());
}

function GetTaskXML()
{
	taskXML.ActiveTag = null;
	taskXML.CreateTagList("STAGETASK");
	
	<% /* ASu BMIDS00469 - Where the confirm button has not been selected before comitting the screen Task XML is null -Start*/%>
	var xmlfound = false;
	<% /* ASu - End */ %>	
	var returnXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequest = returnXML.CreateRequestTagFromArray(m_sRequestAttributes, "CreateAdhocCaseTask");
	returnXML.SetAttribute("USERAUTHORITYLEVEL", m_iUserRole);
	
	for(var iTask = 0; iTask < m_asFlagArray.length; iTask++)
	{
		if(m_asFlagArray[iTask] == "Yes")
		{
			<% /* ASu BMIDS00469 - Where the confirm button has not been selected before comitting the screen test for null -Start*/%>
			xmlfound = true;
			<% /* ASu - End */ %>				
			m_bIsChanged = true ;
			taskXML.SelectTagListItem(iTask);
			
			<%/* calculate the task due date based on chasingperioddays and chasingperiodhours (if present) */ %>
			var nDays = 0;
			var nHours = 0;
			<% /* PSC 01/03/2006 MAR1341 */ %>				
			var nMinutes = 0;
			if(taskXML.GetAttribute("CHASINGPERIODDAYS") != "") nDays = parseInt(taskXML.GetAttribute("CHASINGPERIODDAYS"));
			if(taskXML.GetAttribute("CHASINGPERIODHOURS") != "") nHours = parseInt(taskXML.GetAttribute("CHASINGPERIODHOURS"));
			<% /* PSC 01/03/2006 MAR1341 */ %>
			if(taskXML.GetAttribute("CHASINGPERIODMINUTES") != "") nMinutes = parseInt(taskXML.GetAttribute("CHASINGPERIODMINUTES"));
			<% /* DRC 17/03/2006 MAR1350 - nedd to see if TASK is PC */  %>								
			var strTaskType = "";
			strTaskType = taskXML.GetAttribute("TASKTYPE");
			<% /* DRC 17/03/2006 MAR1350 */ %>
			returnXML.ActiveTag = xmlRequest;
			returnXML.CreateActiveTag("CASETASK");
			returnXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
			returnXML.SetAttribute("CASEID", m_sApplicationNo);
			returnXML.SetAttribute("CASESTAGESEQUENCENO", m_sSeqNo);
			returnXML.SetAttribute("ACTIVITYID", m_sActivityId);
			returnXML.SetAttribute("ACTIVITYINSTANCE", "1");
			returnXML.SetAttribute("STAGEID", m_sStageId);
			returnXML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
			<% /* PSC 03/10/2005 MAR32 - Start */ %>
			returnXML.SetAttribute("OUTPUTDOCUMENT", taskXML.GetAttribute("OUTPUTDOCUMENT"));
			returnXML.SetAttribute("ALWAYSAUTOMATICONCREATION", taskXML.GetAttribute("ALWAYSAUTOMATICONCREATION"));
			<% /* PSC 03/10/2005 MAR32 - End */ %>
			
			<% /* PSC 01/03/2006 MAR1341 - Start */ %>				
			if(nDays != 0 || nHours != 0 || nMinutes != 0) 
				returnXML.SetAttribute("TASKDUEDATEANDTIME", GetNewDueDate(nDays, nHours, nMinutes, strTaskType));
			<% /* PSC 01/03/2006 MAR1341 - End */ %>				
			
			if(taskXML.GetAttribute("CASETASKNAME") == "")
				returnXML.SetAttribute("CASETASKNAME", taskXML.GetAttribute("TASKNAME"));
			else
				returnXML.SetAttribute("CASETASKNAME", taskXML.GetAttribute("CASETASKNAME"));
				
			returnXML.SetAttribute("CUSTOMERIDENTIFIER", GetApplicant(iTask));
			returnXML.SetAttribute("CONTEXT", GetOrganisation(iTask));
			
			<% /* DB BM0248 - Pass in mandatory flag when adding ad-hoc task. */ %>
			returnXML.SetAttribute("MANDATORYFLAG", taskXML.GetAttribute("MANDATORYFLAG"));
			<% /* DB End */ %>
			
			returnXML.ActiveTag = xmlRequest;
			returnXML.CreateActiveTag("APPLICATION");
			returnXML.SetAttribute("APPLICATIONPRIORITY", m_sApplicationPriority);
		}
	}
	
	<% /* PSC 03/10/2005 MAR32 - Start */ %>
	if (xmlfound)
	{
		returnXML.SelectTag(null, "REQUEST");
		returnXML.SelectNodes("CASETASK[@OUTPUTDOCUMENT!='' and @ALWAYSAUTOMATICONCREATION='1']");
		
		if (returnXML.ActiveTagList.length > 0)
		{
			var sPrinter = GetLocalPrinters();
			
			if(sPrinter == "")
			{					
				alert("One or more tasks selected require a default print destination. You do not have a default printer set on your PC.");
				return "No Printer"		
			}
			else
			{
				returnXML.CreateActiveTag("PRINTER");
				returnXML.SetAttribute("PRINTERNAME", sPrinter);
			}			
		}
	}
	<% /* PSC 03/10/2005 MAR32 - End/ %>
	
		
	<% /* ASu BMIDS00469 - Where the confirm button has not been selected before comitting the screen test for null -Start*/%>
	if(xmlfound != true) // The Adhoc Task has not been found and this will raise a not found error due to "Confirm" button not having been selected first.
		return false;
	else			
		return returnXML.XMLDocument.xml ;
	<% /* ASu - End */ %>	
}

function frmScreen.btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}

function frmScreen.btnOK.onclick()
{
	<% // HMA  BMIDS863 Add Screen Rules %>
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}
	
	var sReturn = new Array();
	sReturn[0] = m_bIsChanged ; 
	sReturn[1] = GetTaskXML();
	<% /* ASu BMIDS00469 - Where the confirm button has not been selected before comitting the screen test for null -Start*/%>
	if (sReturn[1] != false)
	{
		<% /* PSC 03/10/2005 MAR32 - Start */ %>
		if (sReturn[1] != "No Printer")
		{
			window.returnValue	= sReturn;
			window.close();
		}
		<% /* PSC 03/10/2005 MAR32 - End */ %>

	}
	else
	{
		alert('Please "Confirm" the details before selecting the "Ok" button');
	}	
	<% /* ASu - End */ %>
}

<% /* PSC 03/10/2005 MAR32 - Start */ %>
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
<% /* PSC 03/10/2005 MAR32 - End */ %>
<% /* DRC 17/03/2006 MARS1350 - Start */ %>
function GetCompletionDate()
{
var strCompletionDate = "";
var RotXML =  new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	RotXML.CreateRequestTagFromArray(m_sRequestAttributes,"GetReportOnTitleData"); 
	RotXML.CreateActiveTag("REPORTONTITLE");
	RotXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNo);
	RotXML.SetAttribute("APPLICATIONFACTFINDNUMBER", "1");
	RotXML.RunASP(document,"ReportOnTitle.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = RotXML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0] == true)
	{
		<% /* Get Completion  */ %>
		RotXML.SelectTag (null, "REPORTONTITLE");
    	strCompletionDate = RotXML.GetAttribute("COMPLETIONDATE");	
	}	
	RotXML=null;
	return strCompletionDate;

}
<% /* DRC 17/03/2006 MAR1350 - End */ %>

<% /* MAR1449 Filter tasks */ %>

function FilterTasks()
{
	var iCount;
	var bUnderwriter;
	var sTaskType;
	
	for (iCount = 0; iCount < taskXML.ActiveTagList.length; iCount++)
	{
		taskXML.SelectTagListItem(iCount);

		<% /* Only show Underwriting tasks if the flag is set otherwise show all tasks */ %>
		bUnderwriter = false;
		sTaskType = taskXML.GetAttribute("TASKTYPE");		

		var TempXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		if (TempXML.IsInComboValidationList(document,"TaskType",sTaskType,["UW"]))
			bUnderwriter=true;		

		if ((m_bUnderwritingTasksOnly == true) && (bUnderwriter == false))
		{
			<%/* Remove this task from the list */%>
			taskXML.RemoveActiveTag();
		}
	}
		
	return;
}

-->
</script>
</body>
</html>




