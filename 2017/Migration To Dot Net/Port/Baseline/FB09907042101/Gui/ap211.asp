<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP211.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:	Valuation Report - Buy To Let Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AShaw	22/11/2006	EP2_2 Created screen Links to AP210.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/

%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>AP211 - Buy To Let  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" mark validate ="onchange">
	<div id="divBackground" style="LEFT: 10px; WIDTH: 400px; POSITION: absolute; TOP: 10px; HEIGHT: 270px" class="msgGroup">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 25px" class="msgLabel">
		Rental Demand
			<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
				<select id="cboRentalDemand" style="WIDTH: 120px" class="msgCombo" NAME="cboTypeOfProperty"></select>
			</span>
		</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Rental Income
		<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
			<input id="txtRentalIncome" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt" NAME="txtRentalIncome">
		</span>
	</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 95px" class="msgLabel">	
			Suitable Condition for Letting?
			<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
				<input id="Letting_Yes" name="Letting" type="radio" value="1"><label for="Letting_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
				<input id="Letting_No" name="Letting" type="radio" value="0"><label for="Letting_No" class="msgLabel">No</label>
			</span>
		</span>
		
	<span style="TOP: 200px; LEFT: 50px; POSITION: ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton" NAME="btnOK">
	</span>
	<span style="TOP: 200px; LEFT: 115px; POSITION: ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton" NAME="btnCancel">
	</span>
		
	</div>
</form>

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP211attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;
	
var BTLXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_BaseNonPopupWindow = null;

/* EVENTS */

function frmScreen.btnCancel.onclick()
{
	BTLXML = null;
	window.close();
}

function frmScreen.btnOK.onclick()
{
	if (CommitChanges())
	{
		BTLXML = null;
		window.close();
	}
}

function window.onload()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters	= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	BTLXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	BTLXML.LoadXML(sParameters[0]);
	m_sReadOnly	= sParameters[1];
	m_sApplicationNumber = sParameters[2];
	m_sApplicationFactFindNumber = sParameters[3];

	SetMasks();
	Validation_Init();	
	Initialise();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

/* FUNCTIONS */

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
			{
				bSuccess = SaveBTL();
			}
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function Initialise()
// Initialises the screen
{
	PopulateCombo()
	PopulateScreen();
	if(m_sReadOnly == "1")
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateScreen()
// Populates the screen with details of the item selected in 
{
	BTLXML.SelectTag(null,"GETVALUATIONREPORT");
	if (BTLXML.ActiveTag != null)
		{
			frmScreen.cboRentalDemand.value = BTLXML.GetAttribute("RENTALDEMAND");
			frmScreen.txtRentalIncome.value = BTLXML.GetAttribute("RENTALINCOME"); 
			SetOptionValue(frmScreen.Letting_Yes, frmScreen.Letting_No, BTLXML.GetAttribute("RENTABLECONDITION")); 
		}
}

function SaveBTL()
{
	var sReturn = new Array();
	if (IsChanged())
	{
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		var roottag = XML.CreateActiveTag("HOME");
		XML.ActiveTag = roottag
		XML.CreateActiveTag("BTL");
		XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.SetAttribute("RENTALDEMAND", frmScreen.cboRentalDemand.value);
		XML.SetAttribute("RENTALINCOME", frmScreen.txtRentalIncome.value);
		XML.SetAttribute("RENTABLECONDITION", GetOptionValue(frmScreen.Letting_Yes));

		BTLXML = XML.XMLDocument.xml;
	}
	
	sReturn[0] = IsChanged();
	sReturn[1] = BTLXML;
	window.returnValue	= sReturn;
	window.close();
}

function GetOptionValue( objOption )
{
	var sVal = "0";
	
	if(objOption.checked == true)
	{
		sVal = "1";
	}
	return(sVal)
}

function SetOptionValue( objOptionYes, objOptionNo , sVal )
{
	if( sVal == "1" )
	{
		objOptionYes.checked = true;
	}
	else
	{
		objOptionNo.checked = true;
	}
}

function PopulateCombo()
{
	
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("RentalDemand");

	if(XML.GetComboLists(document,sGroupList))
	{
		XML.PopulateCombo(document,frmScreen.cboRentalDemand,"RentalDemand",true);
	}
}

-->
</script>
</body>
</html>

