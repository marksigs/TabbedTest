<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      ap215.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Flats Screen (a popup).
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Speccific History:

Prog	Date		Description
DPF		02/09/2002	Created new pop window for flat details
HMA     16/09/2003  BM0063  Amend HTML text for radio buttons
MC		05/05/2004	BMIDS751	- White space added to the title
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>AP225 - Flat Details<!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 15px; LEFT: 25px; HEIGHT: 175px; WIDTH: 200px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 5px" class="msgLabel">
		&nbsp;Storeys In Block
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtStoreysInBlock" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
		&nbsp;Floor In Block
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtFloorInBlock" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt" NAME="txtFloorInBlock">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 55px" class="msgLabel">
		&nbsp;Number of units<br>in the development
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtNumberOfUnits" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt" NAME="txtNumberOfUnits">
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 90" class="msgLabel">
		Businesses In Block
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="BusinessInBlock_Yes" name="BusinessInBlock" type="radio" value="1"><label for="BusinessInBlock_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="BusinessInBlock_No" name="BusinessInBlock" type="radio" value="0"><label for="BusinessInBlock_No" class="msgLabel">No</label>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 115px" class="msgLabel">
		&nbsp;Type Of Business
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtTypeOfBusiness" style="POSITION: absolute; WIDTH: 90px" class="msgTxt" NAME="txtTypeOfBusiness">
		</span>
	</span>
	<span style="TOP: 140px; LEFT: 0px; POSITION: ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton" NAME="btnOK">
	</span>
	<span style="TOP: 140px; LEFT: 65px; POSITION: ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	

</div>
</form>

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP225attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;
	
var FlatXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_BaseNonPopupWindow = null;

/* EVENTS */

function frmScreen.btnCancel.onclick()
{
	FlatXML = null;
	window.close();
}

function frmScreen.btnOK.onclick()
{
	if (CommitChanges())
	{
		FlatXML = null;
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
	FlatXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	FlatXML.LoadXML(sParameters[0]);
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
				bSuccess = SaveFlat();
			}
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function Initialise()
// Initialises the screen
{
	PopulateScreen();
	if(m_sReadOnly == "1")
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateScreen()
// Populates the screen with details of the item selected in 
{
	FlatXML.SelectTag(null,"GETVALUATIONREPORT");
	if (FlatXML.ActiveTag != null)
		{
			frmScreen.txtStoreysInBlock.value = FlatXML.GetAttribute("STOREYSINBLOCK");
			frmScreen.txtFloorInBlock.value = FlatXML.GetAttribute("FLOORINBLOCK");
			SetOptionValue(frmScreen.BusinessInBlock_Yes, frmScreen.BusinessInBlock_No, FlatXML.GetAttribute("BUSINESSINBLOCK"));	
			//EP2_2 New fields
			frmScreen.txtNumberOfUnits.value = FlatXML.GetAttribute("NUMBEROFUNITS");
			frmScreen.txtTypeOfBusiness.value = FlatXML.GetAttribute("BUSINESSINBLOCKTYPEOFBUSINESS");
			// Enable txtTypeOfBusiness if BusinessInBlock_Yes
			var bEnableTofBusiness = frmScreen.BusinessInBlock_Yes.checked;
			OnChangeTypeBusiness(bEnableTofBusiness);
		}
}

function SaveFlat()
{
	var sReturn = new Array();
	if (IsChanged())
	{
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		var roottag = XML.CreateActiveTag("HOME");
		XML.ActiveTag = roottag
		XML.CreateActiveTag("FLAT");
		XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.SetAttribute("STOREYSINBLOCK", frmScreen.txtStoreysInBlock.value);
		XML.SetAttribute("FLOORINBLOCK", frmScreen.txtFloorInBlock.value);
		XML.SetAttribute("BUSINESSINBLOCK", GetOptionValue(frmScreen.BusinessInBlock_Yes));
		//EP2_2 New fields
		XML.SetAttribute("NUMBEROFUNITS", frmScreen.txtNumberOfUnits.value);
		XML.SetAttribute("BUSINESSINBLOCKTYPEOFBUSINESS", frmScreen.txtTypeOfBusiness.value);

		FlatXML = XML.XMLDocument.xml;
	}
	
	sReturn[0] = IsChanged();
	sReturn[1] = FlatXML;
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

//EP2_2 - New methods
function frmScreen.BusinessInBlock_Yes.onclick()
{
	OnChangeTypeBusiness(true);
}	

function frmScreen.BusinessInBlock_No.onclick()
{
	OnChangeTypeBusiness(false);
}
	
function OnChangeTypeBusiness(bFlag)
{
	// Enable if true.
	if (bFlag == true)
		frmScreen.txtTypeOfBusiness.disabled = false;	
	else
	{	
		frmScreen.txtTypeOfBusiness.value = "";
		frmScreen.txtTypeOfBusiness.disabled = true;
	}	
}	

//EP2_2 - END New methods


-->
</script>
</body>
</html>
