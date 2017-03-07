<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /* 
Workfile:      DC230.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Other Residents Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		08/03/2000	Created
AD		04/02/2000	Rework
AD		02/03/2000	Fixed SYS0336
AY		31/03/00	New top menu/scScreenFunctions change
IW		26/04/00	Cancel Navigates to DC225
BG		17/05/00	SYS0752 Removed Tooltips
CL		05/03/01	SYS1920 Read only functionality added
JLD		10/12/01	SYS2806 completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
ASu		18/09/2002	BMIDS00396 Correct setting of Delete and Edit buttons
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History

Prog	Date		AQR		Description
MF		04/08/2005	MARS20	Routes back to DC201
MF		08/08/2005	MAR20	Parameterised routing for Global Parameter "ThirdPartySummary"
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM 2 Specific History

Prog	Date		AQR		Description
PE		14/02/2007	EP2_745	New column "Relationship to Applicant"
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>
<span id="spnListScroll">
	<span style="LEFT: 306px; POSITION: absolute; TOP: 230px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>


<%/* FORMS */%>
<form id="frmToDC225" method="post" action="DC225.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC235" method="post" action="DC235.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC240" method="post" action="DC240.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC201" method="post" action="DC201.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC280" method="post" action="DC280.asp" STYLE="DISPLAY: none"></form>
<form id="frmGeneric" method="post" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="HEIGHT: 200px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 10px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">
				<td width="25%" class="TableHead">First Forename&nbsp</td>
				<td width="25%" class="TableHead">Surname&nbsp</td>
				<td width="10%" class="TableHead">Sex&nbsp</td>
				<td width="12%" class="TableHead">DOB&nbsp</td>
				<td width="17%" class="TableHead">Relationship</td>
			</tr>
			<tr id="row01">
				<td width="25%" class="TableTopLeft">&nbsp</td>
				<td width="25%" class="TableTopCenter">&nbsp</td>
				<td width="10%" class="TableTopCenter">&nbsp</td>
				<td width="12%" class="TableTopCenter">&nbsp</td>
				<td class="TableTopRight">&nbsp</td>
			</tr>
			<tr id="row02">
				<td width="25%" class="TableLeft">&nbsp</td>
				<td width="25%" class="TableCenter">&nbsp</td>
				<td width="10%" class="TableCenter">&nbsp</td>
				<td width="12%" class="TableCenter">&nbsp</td>
				<td class="TableRight">&nbsp</td>
			</tr>
			<tr id="row03">
				<td width="25%" class="TableLeft">&nbsp</td>
				<td width="25%" class="TableCenter">&nbsp</td>
				<td width="10%" class="TableCenter">&nbsp</td>
				<td width="12%" class="TableCenter">&nbsp</td>
				<td class="TableRight">&nbsp</td>
			</tr>
			<tr id="row04">
				<td width="25%" class="TableLeft">&nbsp</td>
				<td width="25%" class="TableCenter">&nbsp</td>
				<td width="10%" class="TableCenter">&nbsp</td>
				<td width="12%" class="TableCenter">&nbsp</td>
				<td class="TableRight">&nbsp</td>
			</tr>
			<tr id="row05">
				<td width="25%" class="TableLeft">&nbsp</td>
				<td width="25%" class="TableCenter">&nbsp</td>
				<td width="10%" class="TableCenter">&nbsp</td>
				<td width="12%" class="TableCenter">&nbsp</td>
				<td class="TableRight">&nbsp</td>
			</tr>
			<tr id="row06">
				<td width="25%" class="TableLeft">&nbsp</td>
				<td width="25%" class="TableCenter">&nbsp</td>
				<td width="10%" class="TableCenter">&nbsp</td>
				<td width="12%" class="TableCenter">&nbsp</td>
				<td class="TableRight">&nbsp</td>
			</tr>
			<tr id="row07">
				<td width="25%" class="TableLeft">&nbsp</td>
				<td width="25%" class="TableCenter">&nbsp</td>
				<td width="10%" class="TableCenter">&nbsp</td>
				<td width="12%" class="TableCenter">&nbsp</td>
				<td class="TableRight">&nbsp</td>
			</tr>
			<tr id="row08">
				<td width="25%" class="TableLeft">&nbsp</td>
				<td width="25%" class="TableCenter">&nbsp</td>
				<td width="10%" class="TableCenter">&nbsp</td>
				<td width="12%" class="TableCenter">&nbsp</td>
				<td class="TableRight">&nbsp</td>
			</tr>
			<tr id="row09">
				<td width="25%" class="TableLeft">&nbsp</td>
				<td width="25%" class="TableCenter">&nbsp</td>
				<td width="10%" class="TableCenter">&nbsp</td>
				<td width="12%" class="TableCenter">&nbsp</td>
				<td class="TableRight">&nbsp</td>
			</tr>
			<tr id="row10">
				<td width="25%" class="TableBottomLeft">&nbsp</td>
				<td width="25%" class="TableBottomCenter">&nbsp</td>
				<td width="10%" class="TableBottomCenter">&nbsp</td>
				<td width="12%" class="TableBottomCenter">&nbsp</td>
				<td class="TableBottomRight">&nbsp</td>
			</tr>
		</table>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 172px">
		<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnAdd.onClick()"> 
	</span>

	<span style="LEFT: 68px; POSITION: absolute; TOP: 172px">
		<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnEdit.onClick()"> 
	</span>

	<span style="LEFT: 132px; POSITION: absolute; TOP: 172px">
		<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnDelete.onClick()"> 
	</span> 
</div> 
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 470px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/DC230Attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var OtherResidentXML = null;
var scScreenFunctions;
var m_blnReadOnly = false;


/* EVENTS */

function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber", m_sApplicationNumber);
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber", m_sApplicationFactFindNumber);
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Add");
	frmToDC235.submit();
}

function btnCancel.onclick()
{	
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else {
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var bNewPropertySummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");			
		if(bNewPropertySummary)
			frmToDC201.submit();
		else
			<% /* MF Existing default route before MARS20 */ %>
			frmToDC225.submit();		
	}
}

function frmScreen.btnDelete.onclick()
{
	//Get the XML that just contains the GroupConnection chosen in the listbox
	var XML = GetXMLBlock(true, scTable.getRowSelected());
	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();

	// 	XML.RunASP(document,"DeleteOtherResident.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"DeleteOtherResident.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if (XML.IsResponseOK())
	{
		OtherResidentXML.RemoveActiveTag();
		var nRowSelected = scTable.getRowSelected();
		scTable.RowDeleted();

		if (GetXMLBlock(false, nRowSelected) == null)
			<%/* Current row is now blank, so select the previous row */%>
			nRowSelected--;
		if (nRowSelected > 0)
			scTable.setRowSelected(nRowSelected);
		<% /* ASu BMIDS00396 - Start */ %>
		else
			frmScreen.btnAdd.disabled = false;
			frmScreen.btnDelete.disabled = true;
			frmScreen.btnEdit.disabled = true;
			btnSubmit.focus();
		<% /* ASu - End	*/ %>
	}
}

function frmScreen.btnEdit.onclick()
{
	var XML = GetXMLBlock(true,scTable.getRowSelected());
	XML.SelectTag(null,"OTHERRESIDENT");
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber", m_sApplicationNumber);
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber", m_sApplicationFactFindNumber);
	scScreenFunctions.SetContextParameter(window,"idOtherResidentSequenceNumber", XML.GetTagText("OTHERRESIDENTSEQUENCENUMBER"));
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Edit");
	frmToDC235.submit();
}

function btnSubmit.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else {						
		<% /* MF Read Global Parameters to decide route */ %>
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var bNewPropertySummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");			
		if(bNewPropertySummary)
			frmToDC201.submit();
		else{
			var bThirdPartySummary = XML.GetGlobalParameterBoolean(document,"ThirdPartySummary");	
			if(bThirdPartySummary)			
				<% /* MF Existing default route before MARS20 */ %>
				frmToDC240.submit();
			else
				frmToDC280.submit();			
		}
	}		
}		

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function spnTable.ondblclick()
{
	//ASu BMIDS00396 Changes
	if ((scTable.getRowSelected() != null) && (scTable.getRowSelected() != -1)) frmScreen.btnEdit.onclick();
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Other Residents Summary","DC230",scScreenFunctions);

	RetrieveContextData();
	Initialise();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC230");
	if (m_blnReadOnly == true) 
	{
		m_sReadOnly = "1";
		<% /* MO - 18/11/2002 - BMIDS00376 */ %>
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnDelete.disabled = true;
	}
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

function GetXMLBlock(bForEdit, nTableRowSelected)
{
	OtherResidentXML.ActiveTag = null;
	OtherResidentXML.CreateTagList("OTHERRESIDENT");

	<% /* Get the index of the selected row */ %>
	var nRowSelected = scTable.getOffset() + nTableRowSelected;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	if(OtherResidentXML.SelectTagListItem(nRowSelected-1) == true)
	{
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("OTHERRESIDENT");

		XML.CreateTag("APPLICATIONNUMBER", OtherResidentXML.GetTagText("APPLICATIONNUMBER"));
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", OtherResidentXML.GetTagText("APPLICATIONFACTFINDNUMBER"));
		if(bForEdit)
			XML.CreateTag("OTHERRESIDENTSEQUENCENUMBER", OtherResidentXML.GetTagText("OTHERRESIDENTSEQUENCENUMBER"));
	}
	else
		XML = null;

	return(XML);
}

function Initialise()
{
	PopulateScreen();
}

function PopulateScreen()
{
	OtherResidentXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	OtherResidentXML.CreateRequestTag(window,null);
	OtherResidentXML.CreateActiveTag("OTHERRESIDENT");
	OtherResidentXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	OtherResidentXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);

	OtherResidentXML.RunASP(document,"FindOtherResidentList.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = OtherResidentXML.CheckResponse(ErrorTypes);
	var nListLength = 0;
	if ((ErrorReturn[1] == ErrorTypes[0]) | (OtherResidentXML.XMLDocument.text == ""))
	{
		//Error: record not found
		if(m_sReadOnly == "1")
			frmScreen.btnAdd.disabled = true;
		else
			frmScreen.btnAdd.disabled = false;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
		btnSubmit.focus();
	}
	else if (ErrorReturn[0] == true)
	{
		// No error
		PopulateTable(0);
		OtherResidentXML.ActiveTag = null;
		OtherResidentXML.CreateTagList("OTHERRESIDENT");
		nListLength = OtherResidentXML.ActiveTagList.length;

		if(m_sReadOnly == "1")
		{
			frmScreen.btnAdd.disabled = true;
			frmScreen.btnDelete.disabled = true;
			frmScreen.btnEdit.focus();
		}
		else
		{
			frmScreen.btnAdd.disabled = false;
			frmScreen.btnDelete.disabled = false;
			frmScreen.btnEdit.disabled = false;
			frmScreen.btnAdd.focus();
		}
	}

	ErrorTypes = null;
	ErrorReturn = null;

	scTable.initialiseTable(tblTable,0,"",PopulateTable,10,nListLength);
	if (nListLength > 0) scTable.setRowSelected(1);
	
	if (frmScreen.btnAdd.disabled == false) frmScreen.btnAdd.focus();
}

function PopulateTable(nStart)
{
	OtherResidentXML.ActiveTag = null;
	OtherResidentXML.CreateTagList("OTHERRESIDENT");

	var sFirstForename = "";
	var sSurname = "";
	var sGender = "";
	var sDOB = "";
	var sRelationship = "";

	for (var iCount = 0; (iCount < OtherResidentXML.ActiveTagList.length) && (iCount < 10); iCount++)
	{
		OtherResidentXML.SelectTagListItem(iCount + nStart);

		sFirstForename = OtherResidentXML.GetTagText("FIRSTFORENAME");
		sSurname = OtherResidentXML.GetTagText("SURNAME");
		sGender = OtherResidentXML.GetTagAttribute("GENDER","TEXT");
		sDOB = OtherResidentXML.GetTagText("DATEOFBIRTH");
		sRelationship = OtherResidentXML.GetTagAttribute("RELATIONSHIPTOAPPLICANT", "TEXT");

		// Add to the search table
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sFirstForename);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),sSurname);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),sGender);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),sDOB);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),sRelationship);
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount);
	}
}
		
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","1325");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
}
-->
</script>
</body>
</html>




