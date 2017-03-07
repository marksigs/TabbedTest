<%@ LANGUAGE="JSCRIPT" %>
<%
/*
Workfile:      AP250.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Valuers Instruction Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
		16.01.2001	Created  By  GD. SYS2039
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
SYS2078	19/03/01	GD : Change to btnEditInstructions - Readonly if status == 30 (completed)
SYS2092 11/05/01	GD : added to btnConfirm Update ValuationStatus to 20 (assign)
AT		30/04/02	SYS4359 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions

BMIDS History

BS		11/06/2003	BM0521 Disable Add button when screen is read only
HMA     08/10/2003  BMIDS640  Highlight 'New' instruction on entry.
                              Only allow Confirm when a New instruction is selected.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<html>
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
<OBJECT data=scTable.htm height=1 id=scTable style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet 
VIEWASTEXT></OBJECT>
<span style="TOP: 240px; LEFT: 310px; POSITION: absolute">
	<OBJECT data=scListScroll.htm id=scScrollPlus 
style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 
type=text/x-scriptlet VIEWASTEXT></OBJECT>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToAP030" method="post" action="AP030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN070" method="post" action="MN070.asp" STYLE="DISPLAY: none"></form>

<form id="frmToAP260" method="post" action="AP260.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 220px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">

	<span id="spnValuerList" style="LEFT: 4px; POSITION: absolute; TOP: 4px">
		<table id="tblValuerList" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width='15%' class='TableHead' id='idValuationNo'>&nbsp;</td><td width='17%' class='TableHead' id='idValuationType'>&nbsp;</td><td width='19%' class='TableHead'>Appointment Date</td><td width='19%' class='TableHead'>Instruction Date</td><td width='12%' class='TableHead'>Status</td><td class='TableHead'>Panel No.</td></tr>
			<tr id="row01">		<td width="15%" class="TableTopLeft">&nbsp;</td><td width="17%" class="TableTopCenter">&nbsp;   </td><td width="19%" class="TableTopCenter">&nbsp;</td>		<td width="19%" class="TableTopCenter">&nbsp;</td>		<td  width="12%" class="TableTopCenter">&nbsp;  </td>  <td   class="TableTopRight">&nbsp;  </td> </tr>
			<tr id="row02">		<td width="15%" class="TableLeft">&nbsp;</td><td width="17%" class="TableCenter">&nbsp;      </td><td width="19%" class="TableCenter">&nbsp;</td>		<td width="19%" class="TableCenter">&nbsp;</td>			<td  width="12%" class="TableCenter">&nbsp;      </td> <td   class="TableRight">&nbsp;  </td> </tr>
			<tr id="row03">		<td width="15%" class="TableLeft">&nbsp;</td><td width="17%" class="TableCenter">&nbsp;      </td><td width="19%" class="TableCenter">&nbsp;</td>		<td width="19%" class="TableCenter">&nbsp;</td>			<td  width="12%" class="TableCenter">&nbsp;      </td> <td   class="TableRight">&nbsp;  </td> </tr>
			<tr id="row04">		<td width="15%" class="TableLeft">&nbsp;</td><td width="17%" class="TableCenter">&nbsp;      </td><td width="19%" class="TableCenter">&nbsp;</td>		<td width="19%" class="TableCenter">&nbsp;</td>			<td  width="12%" class="TableCenter">&nbsp;      </td> <td   class="TableRight">&nbsp;  </td> </tr>
			<tr id="row05">		<td width="15%" class="TableLeft">&nbsp;</td><td width="17%" class="TableCenter">&nbsp;      </td><td width="19%" class="TableCenter">&nbsp;</td>		<td width="19%" class="TableCenter">&nbsp;</td>			<td  width="12%" class="TableCenter">&nbsp;      </td> <td   class="TableRight">&nbsp;  </td> </tr>
			<tr id="row06">		<td width="15%" class="TableLeft">&nbsp;</td><td width="17%" class="TableCenter">&nbsp;      </td><td width="19%" class="TableCenter">&nbsp;</td>		<td width="19%" class="TableCenter">&nbsp;</td>			<td  width="12%" class="TableCenter">&nbsp;      </td> <td   class="TableRight">&nbsp;  </td> </tr>
			<tr id="row07">		<td width="15%" class="TableLeft">&nbsp;</td><td width="17%" class="TableCenter">&nbsp;      </td><td width="19%" class="TableCenter">&nbsp;</td>		<td width="19%" class="TableCenter">&nbsp;</td>			<td  width="12%" class="TableCenter">&nbsp;      </td> <td   class="TableRight">&nbsp;  </td> </tr>
			<tr id="row08">		<td width="15%" class="TableLeft">&nbsp;</td><td width="17%" class="TableCenter">&nbsp;      </td><td width="19%" class="TableCenter">&nbsp;</td>		<td width="19%" class="TableCenter">&nbsp;</td>			<td  width="12%" class="TableCenter">&nbsp;      </td> <td   class="TableRight">&nbsp;  </td> </tr>
			<tr id="row09">		<td width="15%" class="TableLeft">&nbsp;</td><td width="17%" class="TableCenter">&nbsp;      </td><td width="19%" class="TableCenter">&nbsp;</td>		<td width="19%" class="TableCenter">&nbsp;</td>			<td  width="12%" class="TableCenter">&nbsp;      </td> <td   class="TableRight">&nbsp;  </td> </tr>
			<tr id="row10">		<td width="15%" class="TableBottomLeft">&nbsp;</td><td width="17%" class="TableBottomCenter">&nbsp;</td><td width="19%" class="TableBottomCenter">&nbsp;</td>	<td width="19%" class="TableBottomCenter">&nbsp;</td>	<td  width="12%" class="TableBottomCenter">&nbsp;</td> <td   class="TableBottomRight">&nbsp;  </td> </tr>
		</table>
	</span>

	<span id="spnButtons" style="LEFT: 4px; POSITION: absolute; TOP: 180px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnAddInstructions" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
			
		<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">
			<input id="btnEditInstructions" value="Edit" type="button" disabled = "true" style="WIDTH: 60px" class="msgButton">
		</span>
		<!--SYS2092
		<span style="LEFT: 128px; POSITION: absolute; TOP: 0px">
			<input id="btnInstruct" value="Instruct" type="button"  disabled = true style="WIDTH: 60px" class="msgButton">
		</span>
		-->
		
		
	</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 320px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>

<!-- #include FILE="Customise/AP250Customise.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/AP250Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction				= null;
var m_sApplicationNumber		= null;
var m_sApplicationFactFindNumber= null;
var m_sInstructionSequenceNo	= null;
var m_sValuationStatus			= null;

var m_sReadOnly					= null;
var XML							= null;
var iTableSize					= 10;
var scScreenFunctions;


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
	var sButtonList = new Array("Submit","Confirm");
	ShowMainButtons(sButtonList);
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Valuers Instruction Summary","AP250",scScreenFunctions);
	RetrieveContextData();
	Customise();
	scTable.initialise(tblValuerList, 0, "");	
	Initialise();
	PopulateScreen();
	
	<% /* BS BM0521 11/06/03 */ %>
	if (scScreenFunctions.IsMainScreenReadOnly(window, "") == true) 
		frmScreen.btnAddInstructions.disabled = true;
		
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}



function Initialise()
{
	
	//FOR DEBUG PURPOSES
	
	frmScreen.btnEditInstructions.disabled=false;
}

function PopulateScreen()
{
<%
	//Call omAppProcBO
	//with "FindValuerInstructionList" , ApplicationNumber, ApplicationFactFindNumber
%>
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, "FindValuerInstructionList");
		XML.CreateActiveTag("VALUERINSTRUCTION");
		XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.SetAttribute("_COMBOLOOKUP_", "1");		
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		//FindValuerInstructionList
		XML.RunASP(document,"omAppProc.asp");
		
		var sErrorArray = new Array("RECORDNOTFOUND");
		var sResponseArray = XML.CheckResponse(sErrorArray);
			
		//CHECK RESPONSE IS OK
		if (sResponseArray[0]==true)
		{
			var iIsRecords = PopulateTable(0);          // BMIDS640
			XML.ActiveTag = null;
			XML.CreateTagList("VALUERINSTRUCTION");
			scScrollPlus.Initialise(PopulateTable, iTableSize, XML.ActiveTagList.length);
			if(iIsRecords > 0)                         // BMIDS640
			{
				scTable.setRowSelected(iIsRecords);    // BMIDS640
				spnValuerList.onclick();
			}


		} else //Error found
		{
			if (sResponseArray[1]=="RECORDNOTFOUND")
			{
				//alert("There are no records to display");
				//Disable Add and Edit buttons
				
				frmScreen.btnEditInstructions.disabled = true;
			} 
		}

}

function spnValuerList.onclick()
{

	if (scTable.getRowSelectedId() != null)
	{
		m_sInstructionSequenceNo = tblValuerList.rows(scTable.getRowSelectedId()).getAttribute("InstructionSequenceNo");
		m_sValuationStatus       = tblValuerList.rows(scTable.getRowSelectedId()).getAttribute("ValuationStatus");
		<%//SYS2092
		//if (m_sValuationStatus == "20")
		//{
		//	frmScreen.btnInstruct.disabled = false;
		//} else
		//{
			//frmScreen.btnInstruct.disabled = true;
		//}
		%>
	}
		
}

		


function spnValuerList.ondblclick()
{
	if (scTable.getRowSelectedId() != null)
	{
		frmScreen.btnEditInstructions.onclick();

	}
}

function PopulateTable(nStart)
{	
	XML.ActiveTag = null;
	XML.CreateTagList("VALUERINSTRUCTION");
	var iCount;
	var iNumberOfValuers = 0;
	iNumberOfValuers = XML.ActiveTagList.length;
	var bDisableAddButton = false;
	var bIsRecords = false;
	var iIsRecords = 0;               // BMIDS640
	for (iCount = 0; (iCount < iNumberOfValuers) && (iCount < iTableSize); iCount++)
	{
		
		bIsRecords=true;
		XML.SelectTagListItem(iCount + nStart);

		var sInstructionSequenceNo	= XML.GetAttribute("INSTRUCTIONSEQUENCENO");
		var sValuationTypeText		= XML.GetAttribute("VALUATIONTYPE_TEXT");
		var sAppointmentDate		= XML.GetAttribute("APPOINTMENTDATE");
		var sDateOfInstruction		= XML.GetAttribute("DATEOFINSTRUCTION");
		var sValuationStatusText	= XML.GetAttribute("VALUATIONSTATUS_TEXT");
		var sValuerPanelNo			= XML.GetAttribute("VALUERPANELNO");
		var sValuationStatus		= XML.GetAttribute("VALUATIONSTATUS");
		//If New or Assigned
		if ((sValuationStatus=="10") || (sValuationStatus=="20"))
		{
			bDisableAddButton=true;
		}

		<% /* BMIDS640 Select the row if it is a New instruction. */ %>		
		if (sValuationStatus=="10") iIsRecords = iCount+1; 

		// Add to the search table
		scScreenFunctions.SizeTextToField(tblValuerList.rows(iCount+1).cells(0),sInstructionSequenceNo);
		scScreenFunctions.SizeTextToField(tblValuerList.rows(iCount+1).cells(1),sValuationTypeText);
		scScreenFunctions.SizeTextToField(tblValuerList.rows(iCount+1).cells(2),sAppointmentDate);	
		scScreenFunctions.SizeTextToField(tblValuerList.rows(iCount+1).cells(3),sDateOfInstruction);
		scScreenFunctions.SizeTextToField(tblValuerList.rows(iCount+1).cells(4),sValuationStatusText);					
		scScreenFunctions.SizeTextToField(tblValuerList.rows(iCount+1).cells(5),sValuerPanelNo);

		tblValuerList.rows(iCount+1).setAttribute("TagListItemCount", iCount);
		//Add InstructionSequenceNo to table row attributes
		tblValuerList.rows(iCount+1).setAttribute("InstructionSequenceNo", sInstructionSequenceNo);	
		tblValuerList.rows(iCount+1).setAttribute("ValuationStatus", sValuationStatus);	
			
	}
	//Disable Add button if necessary
	
	if (bDisableAddButton==true) frmScreen.btnAddInstructions.disabled=true;
 
	//there are some valuers
	if (iNumberOfValuers==0)
	{
		frmScreen.btnEditInstructions.disabled = false;
	} else
	{	
		spnValuerList.onclick()
	}
	
	return(iIsRecords);               // BMIDS640
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
	m_sApplicationNumber		 = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);	
	m_sReadOnly					 = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
}

function frmScreen.btnAddInstructions.onclick()
{
	//frmScreen.btnEditInstructions.disabled=true;
	//Reset temp Read Only flag
	scScreenFunctions.SetContextParameter(window,"idMetaAction",null);
	//Ensure that instruction sequence number is NOT passed in context to allow an add..
	scScreenFunctions.SetContextParameter(window,"idXML2",null);
	frmToAP260.submit();
}

function frmScreen.btnEditInstructions.onclick()
{
	//Pass in Instruction Sequence No.
	scScreenFunctions.SetContextParameter(window,"idXML2", m_sInstructionSequenceNo);
	//Reset the Read only flag
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
	//Application Number, and Application Fact Find Number will already be there
	// GD SYS2078 10/03/2001 removed OR ==20 (assigned).
	if (m_sValuationStatus=="30")
	{
		//Make AP260 READONLY
		scScreenFunctions.SetContextParameter(window,"idMetaAction","READONLY");
	}
	
	frmToAP260.submit();
}
<%//sys2092
/*
function frmScreen.btnInstruct.onclick()
{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var sTMAssignValuer = XML.GetGlobalParameterString(document,"TMAssignValuer");
		if (XML.IsResponseOK() != true)
		{
			alert("Error getting global parameter TMAssignValuer.");
			return;
		}
		var XMLCaseTask = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XMLCaseTask.CreateRequestTag(window, "CompleteSimpleCaseTask");
		XMLCaseTask.CreateActiveTag("CASETASK");
		XMLCaseTask.SetAttribute("SOURCEAPPLICATION", "Omiga");
		XMLCaseTask.SetAttribute("CASEID", m_sApplicationNumber);	
		XMLCaseTask.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",null));
		XMLCaseTask.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
		XMLCaseTask.SetAttribute("TASKID", sTMAssignValuer );		
		XMLCaseTask.RunASP(document,"MsgTMBO.asp");
		
		if(XMLCaseTask.IsResponseOK() != true)
		{
			alert("Error Completing Simple Case Task record.");
			return;
		}
		alert("Instruction Successful.");
}
*/
%>
function btnSubmit.onclick()
{
	// Check which screen was the previous screen. If MN070, go back there
	// else go to TM030
	var sOriginScreen = scScreenFunctions.GetContextParameter(window,"idReturnScreenId",null)
	
	if ((sOriginScreen =="MN070.asp") || (sOriginScreen =="MN070") ||(sOriginScreen =="mn070.asp") || (sOriginScreen =="mn070"))
	{
		//Go to mn070
		frmToMN070.submit();	
	} else
	{
		frmToTM030.submit();
	}
}
//SYS2092 added btnConfirm
function btnConfirm.onclick()
{
	<% /* BMIDS640  Check that a New instruction has been selected before Confirming. */ %>

	if ((scTable.getRowSelectedId() <= 0) || (m_sValuationStatus != "10"))
	{
		alert("Please select a 'New' Instruction");
		return;
	}
	else
	{
		//call complete valuer instructions
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var ReqTag = XML.CreateRequestTag(window, "CompleteValuerInstructions");
		XML.CreateActiveTag("CASETASK");
		XML.SetAttribute("SOURCEAPPLICATION","Omiga" );
		XML.SetAttribute("CASEID", m_sApplicationNumber);		
		XML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",null));	
		XML.SetAttribute("STAGEID", scScreenFunctions.GetContextParameter(window,"idStageId",null));

		
		//Get idTaskXML from context, if there
		var sXMLCaseInfo = scScreenFunctions.GetContextParameter(window,"idTaskXML",null)
		//alert("sXMLCaseInfo is " + sXMLCaseInfo);
		if (sXMLCaseInfo !="")
		{
			var XMLcontext = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XMLcontext.LoadXML(sXMLCaseInfo);
			XMLcontext.SelectTag(null,"CASETASK");
			var sTaskInstance = XMLcontext.GetAttribute("TASKINSTANCE");
			var sActivityInstance = XMLcontext.GetAttribute("ACTIVITYINSTANCE");
			var sCaseStageSequenceNo = XMLcontext.GetAttribute("CASESTAGESEQUENCENO");	
			XML.SetAttribute("TASKINSTANCE", sTaskInstance );
			XML.SetAttribute("ACTIVITYINSTANCE", sActivityInstance);
			XML.SetAttribute("CASESTAGESEQUENCENO", sCaseStageSequenceNo);			
		}
		//End of Get idTaskXML from context	
		
		
		XML.ActiveTag=ReqTag;
		XML.CreateActiveTag("APPLICATION");
		XML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null) );
		XML.SetAttribute("APPLICATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null) );
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null) );
		//alert("XML b4 call to CompleteValuerInstructions is " + XML.XMLDocument.xml);
		// 		XML.RunASP(document,"OmigaTmBO.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"OmigaTmBO.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		
		if (XML.IsResponseOK() == false)
		{
			alert("Error Completing Valuer Instructions");
			return;
		} else //Complete Valuer Instruction was a success
		//GD ADDED 11/05/01 - SYS2092
		{
			if (m_sInstructionSequenceNo !="") 
			{
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				
				XML.CreateRequestTag(window, "UpdateValuerInstructions");
				XML.CreateActiveTag("VALUERINSTRUCTION");
				XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
				XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
				XML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInstructionSequenceNo);
				
				XML.SetAttribute("VALUATIONSTATUS", 20);//Assign
				
				// 				XML.RunASP(document,"omAppProc.asp");		
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									XML.RunASP(document,"omAppProc.asp");		
						break;
					default: // Error
						XML.SetErrorResponse();
					}

				if(XML.IsResponseOK() != true)
				{
					alert("Error updating Valuation Instructions record.");
					return;
				}
			}		
		
		}
		
	}
	frmToTM030.submit()
}

-->
</script>
</body>
</html>




