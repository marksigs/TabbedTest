<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<%/*
Workfile:      AP310.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Add/ Edit Conditions 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		02/03/01	New
SR		14/03/01	SYS2064
SR		26/03/01	SYS2111
BG		12/11/01	SYS3456 Clear details on btnAnother.onclick()
LD		23/05/02	SYS4727 Use cached versions of frame functions
BG		28/06/02	SYS4767 MSMS/Core integration
STB		08/03/02	SYS4157 Restrict condition details field to 1000 chars.
BG		28/06/02	SYS4767 MSMS/Core integration - END
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History : 

Prog	Date		AQR			Description
MV		07/08/2002	BMIDS0302	Core Ref : SYS4728 remove non-style sheet styles
GD		26/09/2002	BMIDS00313	APWP2 - Allow for embedded DB values in ConditionDescription
GD		17/10/2002	BMIDS00650	Remove CHANNELID in call to FindConditionsList()
GD		21/01/2003	BM0263		Don't restrict length of txtDetails to 1000 chars
HA      08/05/03    BM0518      Default Satisy Status to 'Yes' on Conditions.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
ING Specific History : 

Prog	Date		AQR			Description
MV		12/09/2005	MAR35		Added new attribute in UpdateApplicationConditions()
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>

<HEAD>
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
</HEAD>

<% /* Scriptlets */ %>

<% /* FORMS */ %>
<form id="frmToAP300" method="post" action="AP300.asp" STYLE="DISPLAY: none"></form>

<span id="spnListScroll">
	<span style="LEFT: 306px; POSITION: absolute; TOP: 260px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 350px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Add/Edit Conditions </strong>
	</span>
	
	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 24px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="20%" class="TableHead">Condition Number</td>	
				<td width="80%" class="TableHead">Condition Name</td>	
			</tr>
			<tr id="row01">		
				<td width="20%" class="TableTopLeft">&nbsp;</td>		
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">		
				<td width="20%" class="TableLeft">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">		
				<td width="20%" class="TableLeft">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">
				<td width="20%" class="TableLeft">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row05">
				<td width="20%" class="TableLeft">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row06">
				<td width="20%" class="TableLeft">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row07">
				<td width="20%" class="TableLeft">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row08">
				<td width="20%" class="TableLeft">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row09">
				<td width="20%" class="TableLeft">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row10">
				<td width="20%" class="TableBottomLeft">&nbsp;</td>	
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
		
</div>
</form>

<%/* Main Buttons  */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 440px; WIDTH: 500px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<script language="JScript">
<!--

var scScreenFunctions ;
var m_iTableLength = 10 ;

var m_sApplicationNumber, m_sChannelId, m_sMetaAction, m_sApplnConditionSeq ;
var m_sUserId, m_sUnitId ;
var ConditionsXML = null ;


function RetrieveContextData()
{


	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber", "");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber", "");
	m_sChannelId		 = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId", "");	
	m_sMetaAction		 = scScreenFunctions.GetContextParameter(window,"idMetaAction", "");
	m_sUserId			 = scScreenFunctions.GetContextParameter(window,"idUserId", "");
	m_sUnitId			 = scScreenFunctions.GetContextParameter(window,"idUnitId", "");

}

<%/* Events */ %>
function window.onload()
{	
	var sButtonList = new Array("Submit", "Cancel", "Another");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Conditions Applied","AP310",scScreenFunctions);
	
	RetrieveContextData();
	InitialiseScreen();
}

function spnTable.onclick()
{
	var iRowIndex = scTable.getRowSelected() ;
	
	if(iRowIndex != -1) 
	{
		var sFreeFormat = tblTable.rows(iRowIndex).getAttribute("FREEFORMAT");
		var sEditable   = tblTable.rows(iRowIndex).getAttribute("EDITABLE");
		
		frmScreen.txtDetails.value = tblTable.rows(iRowIndex).getAttribute("CONDITIONDESC");
		
		if(sFreeFormat == '1' || sEditable == '1') frmScreen.txtDetails.readOnly = false ;
		else frmScreen.txtDetails.readOnly = true ;
	}
}

function btnSubmit.onclick()
{
	if(m_sMetaAction == "Edit")
	{
		if(UpdateApplicationConditions()) frmToAP300.submit();
		else return ;
	}
	else
	{
		if(m_sMetaAction == "Add")
		{
			if(CreateApplicationConditions()) frmToAP300.submit();
			else return ;
		}
	}
}

function btnCancel.onclick()
{
	frmToAP300.submit();
}

function btnAnother.onclick()
{
	if(CreateApplicationConditions()) 
	{
		alert('This condition has been added to the list') ;
		//BG 12/11/01 SYS3456 Clear details
		frmScreen.txtDetails.value = "";
	}
}


<% /* Functions */ %>

function InitialiseScreen()
{
	
	PopulateScreen();
}

function PopulateScreen()
{
	//alert("populate screen");
	<% /*In Add mode, populate the list box with all conditions that can be applied  */ %>
	if(m_sMetaAction == "Add")
	{
		//BMIDS00313 Check for a cached version of CONDITIONSLIST - use if there
		ConditionsXML =  new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
		ConditionsXML.CreateRequestTag(window, "FINDCONDITIONSLIST");
		ConditionsXML.CreateActiveTag("CONDITIONS");
		<%//GD BMIDS00650 Remove ChannelID as criteria for find
		//ConditionsXML.SetAttribute("CHANNELID", m_sChannelId);
		%>
		ConditionsXML.SetAttribute("DELETEFLAG", "0");
		
		<%//GD BMIDS00313 Start%>
		//Pass in ApplicationNumber and ApplicationFactFindNumber
		ConditionsXML.CreateActiveTag("APPLICATION");
		ConditionsXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		ConditionsXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		
		<%//GD BMIDS00313 End%>
		
		ConditionsXML.RunASP(document,"omAppProc.asp");
		
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = ConditionsXML.CheckResponse(ErrorTypes);
		
	
		if (ErrorReturn[0] == true)
		{	<% /* Populate the List Box */ %>
		
			ConditionsXML.CreateTagList("CONDITIONS");
		
			var iNumberOfRows = ConditionsXML.ActiveTagList.length;
			scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows);
			ShowList(0);
		}
	}
	else
	{ 
		if(m_sMetaAction == "Edit")
		{
			var sIdXml = scScreenFunctions.GetContextParameter(window,"idXml2", "");	
			ConditionsXML =  new top.frames[1].document.all.scXMLFunctions.XMLObject();
			ConditionsXML.LoadXML(sIdXml);
			
			ConditionsXML.CreateTagList("APPLICATIONCONDITIONS")
		
			scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, 1);
			ShowEditList();
			
			DisableMainButton('Another');
		}	
	}
}	

function ShowList(nStart)
{
	var iCount ;
	var sConditionReference, sConditionAbbrText, sConditionDesc ;
	var sEditable, sFreeFormat, sConditionSeq, sConditionType ;
	
	for(iCount = 0 ; iCount < ConditionsXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		ConditionsXML.SelectTagListItem(iCount + nStart);
		sConditionReference = ConditionsXML.GetAttribute("CONDITIONREFERENCE");
		sConditionName		= ConditionsXML.GetAttribute("CONDITIONNAME");
		sConditionDesc		= ConditionsXML.GetAttribute("CONDITIONDESCRIPTION");
		sEditable			= ConditionsXML.GetAttribute("EDITABLEIND");
		sFreeFormat			= ConditionsXML.GetAttribute("FREEFORMATIND");
		sConditionType		= ConditionsXML.GetAttribute("CONDITIONTYPE");
		
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sConditionReference);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sConditionName);
		
		<% /* Add the required attributes	*/ %>
		tblTable.rows(iCount+1).setAttribute("CONDITIONDESC", sConditionDesc);
		tblTable.rows(iCount+1).setAttribute("EDITABLE", sEditable);
		tblTable.rows(iCount+1).setAttribute("FREEFORMAT", sFreeFormat);
		tblTable.rows(iCount+1).setAttribute("CONDITIONTYPE", sConditionType);
	}	
}

function ShowEditList()
{
	ConditionsXML.SelectTagListItem(0);
	var sConditionReference = ConditionsXML.GetAttribute("CONDITIONREFERENCE");
	var sConditionName		= ConditionsXML.GetAttribute("CONDITIONNAME");
	var sConditionDesc		= ConditionsXML.GetAttribute("CONDITIONDESCRIPTION");
	var sEditable			= ConditionsXML.GetAttribute("EDIT");
	var sFreeFormat			= ConditionsXML.GetAttribute("FREEFORMAT");
	m_sApplnConditionSeq	= ConditionsXML.GetAttribute("APPLNCONDITIONSSEQ");
		
	scScreenFunctions.SizeTextToField(tblTable.rows(1).cells(0), sConditionReference);
	scScreenFunctions.SizeTextToField(tblTable.rows(+1).cells(1), sConditionName);
		
	<% /* Add the required attributes	*/ %>
	tblTable.rows(1).setAttribute("CONDITIONDESC", sConditionDesc);
	tblTable.rows(1).setAttribute("EDITABLE", sEditable);
	tblTable.rows(1).setAttribute("FREEFORMAT", sFreeFormat);
}

function UpdateApplicationConditions()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	XML.CreateRequestTag(window, "UPDATEAPPLICATIONCONDITIONS");
	XML.CreateActiveTag("APPLICATIONCONDITIONS");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("APPLNCONDITIONSSEQ", m_sApplnConditionSeq);
	XML.SetAttribute("CONDITIONREFERENCE", ConditionsXML.GetAttribute("CONDITIONREFERENCE"));
	
	var sDetails = frmScreen.txtDetails.value ;
	XML.SetAttribute("CONDITIONDESCRIPTION", sDetails);
	XML.SetAttribute("CONDITIONNAME", sDetails.substr(0, 50));
	XML.SetAttribute("USERMODIFIED","1");
	XML.RunASP(document,"omAppProc.asp");
	
	if(!XML.IsResponseOK()) return false
	
	return true ;
}

function CreateApplicationConditions()
{
	var iRowIndex = scScrollTable.getRowSelected()
	
	if(iRowIndex == -1)
	{
		alert('Select a condition to add to the application.');
		return false ;
	}
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	XML.CreateRequestTag(window, "CREATEAPPLICATIONCONDITIONS");
	XML.CreateActiveTag("APPLICATIONCONDITIONS");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("CONDITIONREFERENCE", tblTable.rows(iRowIndex).cells(0).innerText);
	XML.SetAttribute("CONDITIONDESCRIPTION", frmScreen.txtDetails.value);
	XML.SetAttribute("CONDITIONNAME", tblTable.rows(iRowIndex).cells(1).innerText);
	XML.SetAttribute("CONDITIONTYPE", tblTable.rows(iRowIndex).getAttribute("CONDITIONTYPE"));
	XML.SetAttribute("EDIT", tblTable.rows(iRowIndex).getAttribute("EDITABLE"));
	XML.SetAttribute("FREEFORMAT", tblTable.rows(iRowIndex).getAttribute("FREEFORMAT"));
	//HA BM0518  Change Satisfy Status default to 'Yes'
	XML.SetAttribute("SATISFYSTATUS", 1);
	XML.SetAttribute("USERID", m_sUserId);
	XML.SetAttribute("UNITID", m_sUnitId);
	
	XML.RunASP(document,"omAppProc.asp");
	
	if(!XML.IsResponseOK()) return false
	
	return true ;
}

//GD 21/01/2003	BM0263 START
//BG		28/06/02	SYS4767 MSMS/Core integration
//function frmScreen.txtDetails.onkeyup()
//{	
//	scScreenFunctions.RestrictLength(frmScreen, "txtDetails", 1000, true);
//}
//BG		28/06/02	SYS4767 MSMS/Core integration - END
//GD 21/01/2003	BM0263 END

-->
</script>
</BODY>
</HTML>
<% /* OMIGA BUILD VERSION 009.01.03.22.01 */ %>
