<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      gn300.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Completeness Check screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		17/05/00	Initial revision
MH      31/05/00    Added Memopad screen and finished off.
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
IK		05/04/01	SYS2244 Add call to SyncCustomerIndex function
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
GHun	04/12/2002	BM0159		CustomerVersionNumber should be set in context in btnCompletionSelect.onclick
MV		03/03/2002	BM0404		Amended RunCompletenessCheck();btnSubmit.onclick();RetreiveContextData()
GD		31/03/2003	BM0495		Amended RunCompletenessCheck() to use Context variable if present, else call middle tier
GD		31/03/2003	BM0495		Amended RunCompletenessCheck() to ALWAYS call midlle tier on entry
HMA     22/11/2004  BMIDS600    Use Global parameter for maximum number of customers.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Completeness Check</title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<span id="spnRulesScroll" style="TOP: 448px; LEFT: 310px; POSITION: absolute; VISIBILITY: hidden">
	<object data="scTableListScroll.asp" id="scRulesScroll" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Screen Layout Here */ %>
<div id="divStatus" style="TOP: 60px; LEFT: 10px; WIDTH: 614px; HEIGHT: 100px; POSITION: absolute" class="msgGroup">
	<table id="tblStatus" width="604px" height="100px" border="0" cellspacing="0" cellpadding="0">
		<tr align="center"><td id="colStatus" align="center" class="msgLabelWait" >Please Wait...<br/>Running Completeness Check Rules</td></tr>
	</table>
</div>

<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 434px; WIDTH: 604px; POSITION: ABSOLUTE; VISIBILITY: HIDDEN" class="msgGroup">
	<form id="frmScreen" style="VISIBILITY: hidden">
	<div id="divRules" style="TOP: 4px; LEFT: 4px; WIDTH: 604px; POSITION: ABSOLUTE">
		<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Completion Rules</strong>
		</span>

		<span id="spnRulesTable" style="TOP: 16px; LEFT: 0px; POSITION: ABSOLUTE">
			<table id="tblRules" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
				<tr id="rowTitles">	<td width="12%" class="TableHead">ID</td>			<td class="TableHead">Description</td> <td width="12%" class="TableHead">Cost Modelling?</td></tr>
				<tr id="row01">		<td width="12%" class="TableTopLeft">&nbsp</td>		<td class="TableTopCenter">&nbsp</td> <td width="12%" class="TableTopRight">&nbsp</td></tr>
				<tr id="row02">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row03">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row04">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row05">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row06">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row07">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row08">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row09">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row10">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row11">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row12">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row13">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row14">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row15">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row16">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row17">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row18">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row19">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row20">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row21">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row22">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row23">		<td width="12%" class="TableLeft">&nbsp</td>		<td class="TableCenter">&nbsp</td> <td width="12%" class="TableRight">&nbsp</td></tr>
				<tr id="row24">		<td width="12%" class="TableBottomLeft">&nbsp</td>	<td class="TableBottomCenter">&nbsp</td> <td width="12%" class="TableBottomRight">&nbsp</td></tr>
			</table>
		</span>
		
		<span style="TOP: 388px; LEFT: 0px; POSITION: ABSOLUTE">
			<input id="btnCompletionSelect" value="Select" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		
	</div>
	</div>

<% /* Not used in cost modelling. Not yet supported. Maybe one day
		<span style="TOP: 0px; LEFT: 132px; POSITION: ABSOLUTE">
			<input id="btnPrint" value="Print" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="TOP: 0px; LEFT: 198px; POSITION: ABSOLUTE">
			<input id="btnDetails" value="Details" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
 */ %>
	</div>
</form>
</div>

<form id="frmGo" method="post" style="VISIBILITY: hidden">
</form>
<form id="frmGoBack" method="post" action="mn060.asp" style="VISIBILITY: hidden">
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 500px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/gn300Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_bIsCostModelling = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sApplicationStageArray = null;
var m_sUserId = null;
var m_sActivityId = null;
var m_bReadOnly = null;
var m_bApplicationReadOnly = null;
var m_sCompletenessTaskId = null;
var m_sStageId = null;
var m_sRequestArray = null;
var scScreenFunctions;
var ListXML = null;
var CondXML = null;
var TaskXML = null;
var m_sListXML = null ;
var m_iMaxCustomers = 0;    // BMIDS600

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
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Completeness Check","GN300",scScreenFunctions);
	
	window.setTimeout("CompleteInitialisation()",100);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}	
	
function CompleteInitialisation()	
{
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions.ShowCollection(frmScreen);
	RetrieveContextData();

	// Get data and populate the tables
	RunCompletenessCheck();
	
	divStatus.style.visibility = "hidden";
	divBackground.style.visibility = "visible";
	frmScreen.style.visibility = "visible";
	spnRulesScroll.style.visibility = "visible";
	
	<% /* BMIDS600 Get maximum customers from Global Parameter */ %>
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_iMaxCustomers = XML.GetGlobalParameterAmount(document, "MaximumCustomers") ;
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
	m_sActivityId = scScreenFunctions.GetContextParameter(window, "idStageId", null);
	m_sListXML =  scScreenFunctions.GetContextParameter(window, "idXML", null);
}

function RunCompletenessCheck()
{
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<% // GD BM0495 START 
	/*
	if (m_sListXML != "")
	{
		ListXML.LoadXML(m_sListXML);
		scRulesScroll.initialiseTable(tblRules,0,"",PopulateRulesTable,24,PopulateRulesTable(0));
	} else
	{
		ListXML.CreateRequestTag(window,null);
		ListXML.CreateActiveTag("COMPLETENESSCHECK");
		ListXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
		ListXML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		ListXML.RunASP(document,"RunCompletenessCheck.asp");
		if(ListXML.IsResponseOK())
		{
			scRulesScroll.initialiseTable(tblRules,0,"",PopulateRulesTable,24,PopulateRulesTable(0));
		} 
	}
	*/
	 // BM0495 END %>
	
	ListXML.CreateRequestTag(window,null);
	ListXML.CreateActiveTag("COMPLETENESSCHECK");
	ListXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	ListXML.RunASP(document,"RunCompletenessCheck.asp");
	if(ListXML.IsResponseOK())
	{
		scRulesScroll.initialiseTable(tblRules,0,"",PopulateRulesTable,24,PopulateRulesTable(0));
	}
	
	
}

function PopulateRulesTable(nStart)
{
	var strText = "";
	
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("COMPLETIONRULE");
			
	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) && nLoop < 24; nLoop++)
	{								
		scScreenFunctions.SizeTextToField(tblRules.rows(nLoop + 1).cells(0), ListXML.GetAttribute("RULEID"));
		scScreenFunctions.SizeTextToField(tblRules.rows(nLoop + 1).cells(1), ListXML.GetAttribute("DESCRIPTION"));
		if(ListXML.GetAttribute("COSTMODELLINGIND") == "Y")
			strText = "Yes";
		else
			strText = "";
		scScreenFunctions.SizeTextToField(tblRules.rows(nLoop + 1).cells(2), strText);
	}

	return ListXML.ActiveTagList.length;
}

function btnSubmit.onclick()
{
	DoTaskStatus();
	frmGoBack.action = scScreenFunctions.GetContextParameter(window, "idReturnScreenId", "MN060") + ".asp"
	scScreenFunctions.SetContextParameter(window,"idProcessContext",null);
	scScreenFunctions.SetContextParameter(window,"idXML",null);
	frmGoBack.submit();
}

function frmScreen.btnCompletionSelect.onclick()
{
//	alert("getRowSelected() = " + scRulesScroll.getRowSelected());
//	alert("getRowSelectedIndex() = " + scRulesScroll.getRowSelectedIndex());
	// close window and return selected screen data
	var iCurrentRow = scRulesScroll.getRowSelected();
	if(iCurrentRow == -1)
	{
		window.alert("Please select a Completion Rule") ;
		return ;
	}

	// Get the ScreenId corresponding to the currently selected rule
	var strRuleId = tblRules.rows(iCurrentRow).cells(0).innerText;
	var strPattern = "RESPONSE/COMPLETIONRULE";
	ListXML.ActiveTag = null;
	ListXML.ActiveTagList = ListXML.XMLDocument.selectNodes(strPattern);
	// ListXML.SelectTagListItem(iCurrentRow -1);
	ListXML.SelectTagListItem(scRulesScroll.getRowSelectedIndex() -1);
	
	var sScreen = ListXML.GetAttribute("SCREENID");
	var sCustNo = ListXML.GetAttribute("CUSTOMERNUMBER");
	
	<% /* BM0159 Get the CustomerVersionNumber */ %>
	var sCustVerNo = "";
	
	if (sCustNo.length > 0)
	{
		<% /* Get the CustomerVersionNumber for the Customer matching sCustNo from Context */ %>
		var nLoop = 1;
		var blnCustomerFound = false;
		
		while(!blnCustomerFound && (nLoop <= m_iMaxCustomers))   // BMIDS600
		{
			if (sCustNo == scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop))
			{
				sCustVerNo = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
				blnCustomerFound = true;
			}
			nLoop++;
		}
	}
	else
	{
		<% /* No customer is specified in this completion rule so default to the first customer in context */ %>
		sCustNo = scScreenFunctions.GetContextParameter(window, "idCustomerNumber1", null);
		sCustVerNo = scScreenFunctions.GetContextParameter(window, "idCustomerVersionNumber1", null);
	}
	<% /* BM0159 End */ %>
	
	if(sScreen.length > 0)
	{
		if(sCustNo.length > 0)
		{
			scScreenFunctions.SetContextParameter(window,"idCustomerNumber",sCustNo);
			<% /* BM0159 
			scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber","1");
			*/ %>
			scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",sCustVerNo);
			scScreenFunctions.SyncCustomerIndex(window);
		}
		frmGo.action = sScreen+".asp";
		frmGo.submit();
	}
}

function DoTaskStatus()
{
//	alert(ListXML.XMLDocument.xml);
//	alert(scScreenFunctions.GetContextParameter(window, "idTaskXML",null));

	if(m_bReadOnly || m_bApplicationReadOnly) return;
	
	if(scScreenFunctions.GetContextParameter(window, "idTaskXML",null).length == 0) return;

	var sTaskId = "";
	var xmlTask = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	xmlTask.LoadXML(scScreenFunctions.GetContextParameter(window, "idTaskXML",null));
	xmlTask.SelectTag(null,"CASETASK");
	if(xmlTask.ActiveTag) sTaskId = xmlTask.GetAttribute("TASKID");
	
	if(sTaskId.length == 0) 
	{
		xmlTask = null;
		return;
	}

	if(ListXML.XMLDocument.selectNodes("RESPONSE/COMPLETIONRULE").length != 0)
	{
		if(confirm("Complete Task") == false)
			return;
	}
	
	var xmlReq = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	xmlReq.CreateRequestTag(window,"UpdateCaseTask");
	xmlReq.CreateActiveTag("CASETASK");
	xmlReq.SetAttribute("SOURCEAPPLICATION","Omiga");
	xmlReq.SetAttribute("CASEID",m_sApplicationNumber);
	xmlReq.SetAttribute("ACTIVITYID",xmlTask.GetAttribute("ACTIVITYID"));
	xmlReq.SetAttribute("ACTIVITYINSTANCE",xmlTask.GetAttribute("ACTIVITYINSTANCE"));
	xmlReq.SetAttribute("STAGEID",xmlTask.GetAttribute("STAGEID"));
	xmlReq.SetAttribute("CASESTAGESEQUENCENO",xmlTask.GetAttribute("CASESTAGESEQUENCENO"));
	xmlReq.SetAttribute("TASKID",sTaskId);
	xmlReq.SetAttribute("TASKINSTANCE",xmlTask.GetAttribute("TASKINSTANCE"));
	xmlReq.SetAttribute("TASKSTATUS",scScreenFunctions.GetContextParameter(window, "idconstTaskStatusComplete",null));

	xmlTask = null;
	xmlReq.RunASP(document,"MsgTmBO.asp");
	xmlReq.IsResponseOK();
	xmlReq = null;
}
-->
</script>
</body>
</html>




