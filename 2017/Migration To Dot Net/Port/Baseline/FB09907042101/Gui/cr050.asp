<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<%
/*
Workfile:      cr050.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Applicants/Guarantors screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		04/02/00	Optimised version
AY		11/02/00	Change to msgButtons button types
AY		29/03/00	New top menu/scScreenFunctions change
IVW		08/05/00	SYS0616 - set innerText to space for a removed row otherwise -
					wierd things happen when swapping between apps/gauarantors.
CL		05/03/01	SYS1920 Read only functionality added
SA		07/08/01	SYS0778 Correct enabling/disabling of buttons
LD		23/05/02	SYS4727 Use cached versions of frame functions
TW      09/10/02    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
GHun	02/07/2002  BMIDS00114  Runtime error assigning applicants & guarantors
GHun	05/07/2002	BMIDS00186	Runtime error assigning applicants & guarantors
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles					
HMA     21/11/2004  BMIDS600    Use Global Parameter for maximum applicants.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom Specific History:

Prog	Date		AQR			Description
GHun	19/12/2006	EP2_56		Changes for TOE
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<head>
	<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scApplicantTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<object data="scTable.htm" height="1" id="scGuarantorTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>

<% /* Specify Forms Here */ %>
<form id="frmSubmit" method="post" action="cr030.asp" STYLE="DISPLAY: none"></form>
<form id="frmCancel" method="post" action="cr030.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen">
<div style="TOP: 60px; LEFT: 10px; HEIGHT: 262px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 4px; LEFT: 120px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Applicants...</strong>
	</span>

	<div id="spnApplicants" style="TOP: 30px; LEFT: 120px; POSITION: ABSOLUTE">
		<table id="tblApplicants" width="250" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="row01"><td class="TableOneColumnTop">&nbsp;</td></tr>
			<tr id="row02"><td class="TableOneColumnCenter">&nbsp;</td></tr>
			<tr id="row03"><td class="TableOneColumnCenter">&nbsp;</td></tr>
			<tr id="row04"><td class="TableOneColumnCenter">&nbsp;</td></tr>
			<tr id="row05"><td class="TableOneColumnCenter">&nbsp;</td></tr>
			<tr id="row06"><td class="TableOneColumnBottom">&nbsp;</td></tr>
		</table>

		<span style="TOP: 0px; LEFT: 260px; POSITION: ABSOLUTE">
			<input id="btnAppMoveUp" value="Move Up" type="button" style="WIDTH: 100px" class="msgButton" disabled>
		</span>

		<span style="TOP: 30px; LEFT: 260px; POSITION: ABSOLUTE">
			<input id="btnAppMoveDown" value="Move Down" type="button" style="WIDTH: 100px" class="msgButton" disabled>
		</span>

		<span style="TOP: 60px; LEFT: 260px; POSITION: ABSOLUTE">
			<input id="btnToGuarantor" value="Guarantor &gt;&gt;" type="button" style="WIDTH: 100px" class="msgButton" disabled>
		</span>
	</div>

	<span style="TOP: 144px; LEFT: 120px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Guarantors...</strong>
	</span>

	<div id="spnGuarantors" style="TOP: 170px; LEFT: 120px; POSITION: ABSOLUTE">
		<table id="tblGuarantors" width="250" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="row01"><td class="TableOneColumnTop">&nbsp;</td></tr>
			<tr id="row02"><td class="TableOneColumnCenter">&nbsp;</td></tr>
			<tr id="row03"><td class="TableOneColumnCenter">&nbsp;</td></tr>
			<tr id="row04"><td class="TableOneColumnCenter">&nbsp;</td></tr>
			<tr id="row05"><td class="TableOneColumnCenter">&nbsp;</td></tr>
			<tr id="row06"><td class="TableOneColumnBottom">&nbsp;</td></tr>
		</table>

		<span style="TOP: 0px; LEFT: 260px; POSITION: ABSOLUTE">
			<input id="btnGuarMoveUp" value="Move Up" type="button" style="WIDTH: 100px" class="msgButton" disabled>
		</span>

		<span style="TOP: 30px; LEFT: 260px; POSITION: ABSOLUTE">
			<input id="btnGuarMoveDown" value="Move Down" type="button" style="WIDTH: 100px" class="msgButton" disabled>
		</span>

		<span style="TOP: 60px; LEFT: 260px; POSITION: ABSOLUTE">
			<input id="btnToApplicant" value="Applicant &gt;&gt;" type="button" style="WIDTH: 100px" class="msgButton" disabled>
		</span>
	</div>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cr050Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript" type="text/javascript">
<!--
var XML = null;
var XMLOut = null;
var scScreenFunctions;

var m_iApplicantRows = 0;
var m_iGuarantorRows = 0;
var m_iMaxApplicants = 0;     // BMIDS600

var m_sTypeOfApplication = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sXML = null;
var m_blnReadOnly = false;


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
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Assign Applicants/Guarantors","CR050",scScreenFunctions);
	RetrieveContextData();
	
	<% /* BMIDS600 Get maximum number of applicants */ %>
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_iMaxApplicants = XML.GetGlobalParameterAmount(document, "MaximumApplicants") ;
		
	<% /* BMIDS00114 GHun 02/07/2002 Removed duplicate new
	XML	= new new top.frames[1].document.all.scXMLFunctions.XMLObject();	*/ %>
	XML	= new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(m_sXML);
	XML.CreateTagList("CUSTOMERROLE");
	if (XML.ActiveTagList != null)
	{
		XML.SelectTagListItem(0);
		if (XML.ActiveTag != null)
		{
			m_sApplicationNumber = XML.GetTagText("APPLICATIONNUMBER");
			m_sApplicationFactFindNumber = XML.GetTagText("APPLICATIONFACTFINDNUMBER");
		}
	}

	ShowList();

	scApplicantTable.initialise (tblApplicants, 0, "");
	scGuarantorTable.initialise (tblGuarantors, 0, "");

	EnableButtons();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CR050");
	
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

function ShowList()
{
	var iTableIndex;
	var tblTable;

	XML.ActiveTag = null;
	var TagListCUSTOMERROLE = XML.CreateTagList("CUSTOMERROLE");
	var iNoOfCustomers = XML.ActiveTagList.length;

<%	// loop for all the customers retrieved
%>	for (var i0=0; i0<iNoOfCustomers; i0++)
	{
		XML.ActiveTagList = TagListCUSTOMERROLE;
		XML.SelectTagListItem(i0);

		var sCustomerNumber = XML.GetTagText("CUSTOMERNUMBER");
		var sCustomerVersionNumber = XML.GetTagText("CUSTOMERVERSIONNUMBER");
		var sCustomerRoleType = XML.GetTagText("CUSTOMERROLETYPE");

<%		// We need to know where this customer is an Applicant or Guarantor
%>		if (sCustomerRoleType == "1")
		{
			m_iApplicantRows += 1;
			iTableIndex = m_iApplicantRows;
			tblTable = tblApplicants;
		}
		else
		{
			m_iGuarantorRows += 1;
			iTableIndex = m_iGuarantorRows;
			tblTable = tblGuarantors;
		}

		XML.CreateTagList("CUSTOMERVERSION");
		XML.SelectTagListItem(0);

		var sSurname = XML.GetTagText("SURNAME");
		var sFirstForename = XML.GetTagText("FIRSTFORENAME");
		var sSecondForename = XML.GetTagText("SECONDFORENAME");
		var sOtherForenames = XML.GetTagText("OTHERFORENAMES");

<%		// APS 06/09/99 - UNIT TEST REF 5
%>		ShowRow(iTableIndex-1,sFirstForename,sSecondForename,sOtherForenames,sSurname,sCustomerNumber,sCustomerVersionNumber,tblTable);
	}
}

function ShowRow(iRowIndex,sFirstForename,sSecondForename,sOtherForenames,sSurname,sCustomerNumber,sCustomerVersionNumber,tblTable)
{
<%	// APS 06/09/99 - UNIT TEST REF 5, 88
%>	var sRow = sFirstForename;

	if (sSecondForename.length > 0)
	{
		sRow += " ";
		sRow += sSecondForename;
	}

	if (sOtherForenames.length > 0)
	{
		sRow += " ";
		sRow += sOtherForenames;
	}

	sRow += " ";
	sRow += sSurname;

<%	// AY 09/09/1999 - Set the table fields
%>	scScreenFunctions.SizeTextToField(tblTable.rows(iRowIndex).cells(0),sRow);
			
	tblTable.rows(iRowIndex).setAttribute("CustomerNumber", sCustomerNumber);
	tblTable.rows(iRowIndex).setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
}

function RetrieveContextData()
{
	m_sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	m_sTypeOfApplication = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","1");
}

function btnSubmit.onclick()
{
<%	// TODO: Call applicationBO.UpdateCustomerRoles
	// AY 15/09/99 - only commit data if changed
%>	if(IsChanged())
	{
		UpdateCustomerRoles();
		// 		XMLOut.RunASP(document,"UpdateCustomerRoles.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XMLOut.RunASP(document,"UpdateCustomerRoles.asp");
				break;
			default: // Error
				XMLOut.SetErrorResponse();
			}

		if (XMLOut.IsResponseOK()) frmSubmit.submit();
	}
	else frmSubmit.submit();
}

function UpdateCustomerRoles()
{
	<% /* BMIDS00186 GHun 05/07/2002 Removed duplicate new
	XMLOut = new new top.frames[1].document.all.scXMLFunctions.XMLObject();	*/ %>
	XMLOut = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLOut.CreateRequestTag(window,"UPDATE");
//	XMLOut.CreateActiveTag("UPDATE");
	WriteAfterImage();
}

function WriteAfterImage()
{
	var tagCUSTOMERROLELIST = XMLOut.CreateActiveTag("CUSTOMERROLELIST");

	for (var iRow=0; iRow < m_iApplicantRows; iRow++)
	{
		XMLOut.ActiveTag = tagCUSTOMERROLELIST;
		XMLOut.CreateActiveTag("CUSTOMERROLE");

		var iOrder = iRow+1;

		XMLOut.CreateTag("CUSTOMERNUMBER", tblApplicants.rows(iRow).getAttribute("CustomerNumber"));
		XMLOut.CreateTag("CUSTOMERVERSIONNUMBER", tblApplicants.rows(iRow).getAttribute("CustomerVersionNumber"));
		XMLOut.CreateTag("CUSTOMERORDER", iOrder.toString());
		XMLOut.CreateTag("CUSTOMERROLETYPE", "1");
		XMLOut.CreateTag("TYPEOFAPPLICATION", m_sTypeOfApplication);
		XMLOut.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XMLOut.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	}

	for (var iRow=0; iRow < m_iGuarantorRows; iRow++)
	{
		XMLOut.ActiveTag = tagCUSTOMERROLELIST;
		XMLOut.CreateActiveTag("CUSTOMERROLE");

		var iOrder = iRow+1;

		XMLOut.CreateTag("CUSTOMERNUMBER", tblGuarantors.rows(iRow).getAttribute("CustomerNumber"));
		XMLOut.CreateTag("CUSTOMERVERSIONNUMBER", tblGuarantors.rows(iRow).getAttribute("CustomerVersionNumber"));
		XMLOut.CreateTag("CUSTOMERORDER", iOrder.toString());
		XMLOut.CreateTag("CUSTOMERROLETYPE", "2");
		XMLOut.CreateTag("TYPEOFAPPLICATION", m_sTypeOfApplication);
		XMLOut.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XMLOut.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	}
	
	<% /* EP2_56 GHun For TOE check if any customers have change RoleType */ %>
	CheckTOECustomers();
}

function btnCancel.onclick()
{
	frmCancel.submit();
}
		
function EnableButtons()
{
	var iGuarantorSelectedRow	= -1;
	var iApplicantSelectedRow	= -1;
	var blnDisabled				= false;
			
	iGuarantorSelectedRow = scGuarantorTable.getRowSelected();
	iApplicantSelectedRow = scApplicantTable.getRowSelected();
			
<%	// Do we have any applicants
%>	if ((m_iApplicantRows > 1) && (iApplicantSelectedRow != -1)) blnDisabled = false;
	else blnDisabled = true;

	frmScreen.btnToGuarantor.disabled = blnDisabled;
	frmScreen.btnAppMoveDown.disabled = blnDisabled;
	frmScreen.btnAppMoveUp.disabled = blnDisabled;

<%	// Do we have any guarantors 
%>	if ((m_iGuarantorRows > 0) && (iGuarantorSelectedRow != -1)) blnDisabled = false;
	else blnDisabled = true;
			
	frmScreen.btnToApplicant.disabled = blnDisabled;
	frmScreen.btnAppMoveDown.disabled = blnDisabled;
	frmScreen.btnAppMoveUp.disabled = blnDisabled;
}

function frmScreen.btnToApplicant.onclick()
{
	var iRowSelected = scGuarantorTable.getRowSelected();

<%	// This should be redundant as the Applicant button is only enable on
	// valid selection of Guarantor table
%>	if (iRowSelected == -1) alert("Please select an entry from the list");
	else
	{
<%		// add the row to the applicant table
		// move the customer number and version number over
%>		tblApplicants.rows(m_iApplicantRows).cells(0).innerText = tblGuarantors.rows(iRowSelected).cells(0).innerText;
		var sCustomerNumber = tblGuarantors.rows(iRowSelected).getAttribute("CustomerNumber");
		var sCustomerVersionNumber = tblGuarantors.rows(iRowSelected).getAttribute("CustomerVersionNumber");
		tblApplicants.rows(m_iApplicantRows).setAttribute("CustomerNumber", sCustomerNumber);
		tblApplicants.rows(m_iApplicantRows).setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
		scApplicantTable.setRowSelected(m_iApplicantRows);
		m_iApplicantRows += 1;

<%		// move all the guarantor rows down one row in the table
		// remove the last rows for the guarantor table
		// decrement the number of guarantors
		// De-select the table
%>		for (var iLoop = iRowSelected; iLoop < m_iGuarantorRows; iLoop++)
		{
			MoveLine(tblGuarantors, iLoop+1, iLoop);
<%			//tblGuarantors.rows(iLoop).cells(0).innerText = tblGuarantors.rows(iLoop+1).cells(0).innerText;
%>		}

		// IVW SYS0616 - set innerText to space otherwise wierd things happen, it thinks it has
		// has something to display but it has not, this causes the box to re-size, lose its border , etc
		tblGuarantors.rows(m_iGuarantorRows).cells(0).innerText = " ";
		m_iGuarantorRows -= 1;
		scGuarantorTable.setRowSelected(-1);
		EnableButtons();
		FlagChange(true); <% /* AY 15/09/99 */ %>
		spnApplicants.onclick();	//SYS0778 Call to correctly refresh buttons
	}
}

function MoveLine(tblTable, iNewLineNumber, iOldLineNumber)
{
<%	// save the new line
	// copy the old line to the new line
	// replace the original selected line with the saved line
	// save the Customer No. from the new line
	// set the Customer No. for the new line
	// set the Customer No. for the old line
%>	var sTempLineText = tblTable.rows(iNewLineNumber).cells(0).innerText;
	tblTable.rows(iNewLineNumber).cells(0).innerText = tblTable.rows(iOldLineNumber).cells(0).innerText;
	tblTable.rows(iOldLineNumber).cells(0).innerText = sTempLineText;

	var sTempCustomerNo = tblTable.rows(iNewLineNumber).getAttribute("CustomerNumber");
	var sTempCustomerVerNo = tblTable.rows(iNewLineNumber).getAttribute("CustomerVersionNumber");

	tblTable.rows(iNewLineNumber).setAttribute("CustomerNumber",tblTable.rows(iOldLineNumber).getAttribute("CustomerNumber"));
	tblTable.rows(iNewLineNumber).setAttribute("CustomerVersionNumber", tblTable.rows(iOldLineNumber).getAttribute("CustomerVersionNumber"));
	tblTable.rows(iOldLineNumber).setAttribute("CustomerNumber", sTempCustomerNo);
	tblTable.rows(iOldLineNumber).setAttribute("CustomerVersionNumber", sTempCustomerVerNo);
}

function frmScreen.btnGuarMoveUp.onclick()
{
	var iRowSelected = scGuarantorTable.getRowSelected();

<%	// you cannot move the top row up!
%>	if (iRowSelected > 0)
	{
		MoveLine(tblGuarantors, iRowSelected-1, iRowSelected);
<%		// keep the original selection
%>		scGuarantorTable.setRowSelected(iRowSelected-1);
		FlagChange(true); <% /* AY 15/09/99 */ %>
	}
}

function frmScreen.btnGuarMoveDown.onclick()
{
	var iRowSelected = scGuarantorTable.getRowSelected();

<%	// if the table is selected and the selected row is not the bottom row
%>	if ((iRowSelected != -1) && (iRowSelected < m_iGuarantorRows-1))
	{
		MoveLine(tblGuarantors, iRowSelected+1, iRowSelected);
<%		// keep the original row selected
%>		scGuarantorTable.setRowSelected(iRowSelected+1);
		FlagChange(true); <% /* AY 15/09/99 */ %>
	}
}

function frmScreen.btnToGuarantor.onclick()
{
	var iRowSelected = scApplicantTable.getRowSelected();

	if (iRowSelected == -1) alert("Please select an entry from the list");
	else
	{
<%		// must be more than one applicant!
%>		if (m_iApplicantRows > 1)
		{
<%			// add applicant as a guarantor					
			// move the customer number and version number over
			// make the newly added row the selected														
			// move the applicants up
			// remove the selected applicant
			// set no row selected
%>			tblGuarantors.rows(m_iGuarantorRows).cells(0).innerText = tblApplicants.rows(iRowSelected).cells(0).innerText;
			var sCustomerNumber = tblApplicants.rows(iRowSelected).getAttribute("CustomerNumber");
			var sCustomerVersionNumber = tblApplicants.rows(iRowSelected).getAttribute("CustomerVersionNumber");
			tblGuarantors.rows(m_iGuarantorRows).setAttribute("CustomerNumber", sCustomerNumber);
			tblGuarantors.rows(m_iGuarantorRows).setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
			scGuarantorTable.setRowSelected(m_iGuarantorRows);
			m_iGuarantorRows += 1;					

			for (var iLoop = iRowSelected; iLoop < m_iApplicantRows; iLoop++)
			{						
				MoveLine(tblApplicants, iLoop+1, iLoop);
<%				//tblApplicants.rows(iLoop).cells(0).innerText = tblApplicants.rows(iLoop+1).cells(0).innerText;
%>			}

			// IVW SYS0616 - set innerText to space otherwise wierd things happen, it thinks it has
			// has something to display but it has not, this causes the box to re-size, lose its border , etc
			tblApplicants.rows(m_iApplicantRows).cells(0).innerText = " ";
			m_iApplicantRows -= 1;
			scApplicantTable.setRowSelected(-1);
			EnableButtons();
			FlagChange(true); <% /* AY 15/09/99 */ %>
			spnGuarantors.onclick();	//SYS0778 Call to correctly refresh buttons
		}
	}
}

function frmScreen.btnAppMoveUp.onclick()
{
	var iRowSelected = scApplicantTable.getRowSelected();

	if (iRowSelected > 0)
	{
		MoveLine(tblApplicants,iRowSelected-1, iRowSelected);
<%		// set the selected row back to the original row selected
%>		scApplicantTable.setRowSelected(iRowSelected-1);
		FlagChange(true); <% /* AY 15/09/99 */ %>
	}
}

function frmScreen.btnAppMoveDown.onclick()
{
	var iRowSelected = scApplicantTable.getRowSelected();

<%	// make sure we have a selected row and its not the bottom row
%>	if ((iRowSelected != -1) && (iRowSelected < m_iApplicantRows-1))
	{
		MoveLine(tblApplicants,iRowSelected+1, iRowSelected);
<%		// keep the original selection
%>		scApplicantTable.setRowSelected(iRowSelected+1);
		FlagChange(true); <% /* AY 15/09/99 */ %>
	}
}

function spnApplicants.onclick()
{
	var iSelectedRow = -1;
	var blnTopRow = false;
	var blnBottomRow = false;
	var blnDisableMoveUp = false;
	var blnDisableMoveDown = false;

	iSelectedRow = scApplicantTable.getRowSelected();

	if (iSelectedRow != -1)
	{
		blnTopRow = (iSelectedRow == 0);
		blnBottomRow = (iSelectedRow == (m_iApplicantRows-1));

		if ((blnTopRow) && (!blnBottomRow))
		{
			blnDisableMoveUp = true;
			blnDisableMoveDown = false;
		}
		else if ((!blnTopRow) && (blnBottomRow))
		{
			blnDisableMoveUp = false;
			blnDisableMoveDown = true;
		}
		else if ((blnTopRow) && (blnBottomRow))
		{
			blnDisableMoveUp = true;
			blnDisableMoveDown = true;
		}

		frmScreen.btnAppMoveUp.disabled = blnDisableMoveUp;
		frmScreen.btnAppMoveDown.disabled = blnDisableMoveDown;

		if (m_iApplicantRows > 1) frmScreen.btnToGuarantor.disabled = false;
	//SYS0778
	}
	else
	{
		//no row is selcted so disable all btns
		frmScreen.btnAppMoveUp.disabled = true;
		frmScreen.btnAppMoveDown.disabled = true;
		frmScreen.btnToGuarantor.disabled = true;
		
	}
	//}
}

function spnGuarantors.onclick()
{
	var iSelectedRow = -1;
	var blnTopRow = false;
	var blnBottomRow = false;			
	var blnDisableMoveUp = false;
	var blnDisableMoveDown = false;

	iSelectedRow = scGuarantorTable.getRowSelected();

	if (iSelectedRow != -1)
	{
		blnTopRow = (iSelectedRow == 0);
		blnBottomRow = (iSelectedRow == (m_iGuarantorRows-1));

		if ((blnTopRow) && (!blnBottomRow))
		{
			blnDisableMoveUp = true;
			blnDisableMoveDown = false;
		}
		else if ((!blnTopRow) && (blnBottomRow))
		{
			blnDisableMoveUp = false;
			blnDisableMoveDown = true;
		}
		else if ((blnTopRow) && (blnBottomRow))
		{
			blnDisableMoveUp = true;
			blnDisableMoveDown = true;
		}

		frmScreen.btnGuarMoveUp.disabled = blnDisableMoveUp;
		frmScreen.btnGuarMoveDown.disabled = blnDisableMoveDown;
		//SYS0778 Only enable button if no more than MAX Applicants already.
		if (m_iApplicantRows < m_iMaxApplicants)
		{
			frmScreen.btnToApplicant.disabled = false;
		}
	//SYS0778
	}
	else
	{
		//no row is selcted so disable all btns
		frmScreen.btnGuarMoveUp.disabled = true;
		frmScreen.btnGuarMoveDown.disabled = true;
		frmScreen.btnToApplicant.disabled = true;
		
	}
	//}
}

<% /* EP2_56 GHun */ %>
function CheckTOECustomers()
{
	<% /* Check if the application is transfer of equity */ %>
	var combXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if(combXML.IsInComboValidationList(document,"TypeOfMortgage", m_sTypeOfApplication, ["CC"])) 
	{
		XMLOut.ActiveTag = XMLOut.XMLDocument.documentElement;
		var tagXmlOutParent = XMLOut.CreateActiveTag("REMOVEDTOECUSTOMERLIST");
	
		XML.ActiveTag = null;
		var TagListCUSTOMERROLE = XML.CreateTagList("CUSTOMERROLE");
		var iNoOfCustomers = XML.ActiveTagList.length;

		<%	// loop through all the customers passed in	%>
		for (var i0=0; i0<iNoOfCustomers; i0++)
		{
			XML.ActiveTagList = TagListCUSTOMERROLE;
			XML.SelectTagListItem(i0);

			var sCustomerNumber = XML.GetTagText("CUSTOMERNUMBER");
			var sCustomerVersionNumber = XML.GetTagText("CUSTOMERVERSIONNUMBER");
			var sCustomerRoleType = XML.GetTagText("CUSTOMERROLETYPE");

			if (sCustomerRoleType == "1") <% /* Check if an applicant has changed to a guarantor */ %>
			{
				for (var iRow=0; iRow < m_iGuarantorRows; iRow++)
				{
					if ((tblGuarantors.rows(iRow).getAttribute("CustomerNumber") == sCustomerNumber)
						&& (tblGuarantors.rows(iRow).getAttribute("CustomerVersionNumber") == sCustomerVersionNumber))
					{
						XMLOut.ActiveTag = tagXmlOutParent;
						XMLOut.CreateActiveTag("REMOVEDTOECUSTOMER");
					
						XMLOut.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
						XMLOut.CreateTag("OMIGACUSTOMERNUMBER", sCustomerNumber);
						XMLOut.CreateTag("CIFNUMBER", XML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER"));
						XMLOut.CreateTag("TITLE", XML.GetTagText("TITLE"));
						XMLOut.CreateTag("FIRSTFORENAME", XML.GetTagText("FIRSTFORENAME"));
						XMLOut.CreateTag("SECONDFORENAME", XML.GetTagText("SECONDFORENAME"));
						XMLOut.CreateTag("OTHERFORENAMES", XML.GetTagText("OTHERFORENAMES"));
						XMLOut.CreateTag("SURNAME", XML.GetTagText("SURNAME"));
						XMLOut.CreateTag("CUSTOMERROLE", "2");
						XMLOut.CreateTag("TYPE", "M");
						
					}
				}
			}
			else	<% /* Check if a guarantor has changed to an applicant */ %>
			{
				for (var iRow=0; iRow < m_iApplicantRows; iRow++)
				{
					if ((tblApplicants.rows(iRow).getAttribute("CustomerNumber") == sCustomerNumber)
						&& (tblApplicants.rows(iRow).getAttribute("CustomerVersionNumber") == sCustomerVersionNumber))
					{
						XMLOut.ActiveTag = tagXmlOutParent;
						XMLOut.CreateActiveTag("REMOVEDTOECUSTOMER");
					
						XMLOut.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
						XMLOut.CreateTag("OMIGACUSTOMERNUMBER", sCustomerNumber);
						XMLOut.CreateTag("CIFNUMBER", XML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER"));
						XMLOut.CreateTag("TITLE", XML.GetTagText("TITLE"));
						XMLOut.CreateTag("FIRSTFORENAME", XML.GetTagText("FIRSTFORENAME"));
						XMLOut.CreateTag("SECONDFORENAME", XML.GetTagText("SECONDFORENAME"));
						XMLOut.CreateTag("OTHERFORENAMES", XML.GetTagText("OTHERFORENAMES"));
						XMLOut.CreateTag("SURNAME", XML.GetTagText("SURNAME"));
						XMLOut.CreateTag("CUSTOMERROLE", "1");
						XMLOut.CreateTag("TYPE", "M");
						
					}
				}
			}
		}
	}
}
<% /* EP2_56 GHun */ %>
-->
</script>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
</body>
</html>
