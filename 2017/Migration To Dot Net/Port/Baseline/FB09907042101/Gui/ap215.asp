<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      ap215.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Rooms Screen (a popup).
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
DPF		30/08/2002	Created new pop window for rooms in property
HMA     16/09/2003  BM0063  Amended HTML text for rdio buttons
MC		20/04/2004	BMIDS517	white space padded to the right of title text. (to hide std. title text 'Web Page Dialog...')
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
INGDUK Specific History:
JD		23/09/2005	MAR40 Added number of kitchens to screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>AP215 - Rooms  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 15px; LEFT: 25px; HEIGHT: 310px; WIDTH: 200px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 5px" class="msgLabel">
		&nbsp;Living Rooms
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtLivingRooms" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
		&nbsp;Bedrooms
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtBedrooms" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 55px" class="msgLabel">
		&nbsp;Bathrooms
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtBathrooms" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 80px" class="msgLabel">
		&nbsp;WCs
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtWCs" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 105px" class="msgLabel">
		&nbsp;Habitable Rooms
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtHabitableRooms" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 130px" class="msgLabel">
		&nbsp;Kitchens
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtKitchens" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 155px" class="msgLabel">
		&nbsp;Garages
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtGarages" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 180px" class="msgLabel">
		&nbsp;Parking Spaces
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtParkingSpaces" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt" NAME="txtParkingSpaces">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 205px" class="msgLabel">
		&nbsp;Floors
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtFloors" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt" NAME="txtFloors">
		</span>
	</span>
	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 230px" class="msgLabel">
		Conservatory
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="Conservatory_Yes" name="Conservatory" type="radio" value="1"><label for="Conservatory_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="Conservatory_No" name="Conservatory" type="radio" value="0"><label for="Conservatory_No" class="msgLabel">No</label>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 255px" class="msgLabel">
		Garden Included
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="GardenIncluded_Yes" name="GardenIncluded" type="radio" value="1"><label for="GardenIncluded_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="GardenIncluded_No" name="GardenIncluded" type="radio" value="0"><label for="GardenIncluded_No" class="msgLabel">No</label>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 280px" class="msgLabel">
		Property accessed<br>by balcony?
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="Balcony_Yes" name="Balcony" type="radio" value="1"><label for="Balcony_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="Balcony_No" name="Balcony" type="radio" value="0"><label for="Balcony_No" class="msgLabel">No</label>
		</span>
	</span>
	<span style="TOP: 315px; LEFT: 0px; POSITION: ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	<span style="TOP: 315px; LEFT: 65px; POSITION: ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	
</div>
</form>

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP215attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;
	
var RoomsXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_BaseNonPopupWindow = null;

/* EVENTS */

function frmScreen.btnCancel.onclick()
{
	RoomsXML = null;
	window.close();
}

function frmScreen.btnOK.onclick()
{
	if (CommitChanges())
	{
		RoomsXML = null;
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
	RoomsXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	RoomsXML.LoadXML(sParameters[0]);
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
				bSuccess = SaveRooms();
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
	RoomsXML.SelectTag(null,"GETVALUATIONREPORT");
	if (RoomsXML.ActiveTag != null)
		{
			frmScreen.txtLivingRooms.value = RoomsXML.GetAttribute("LIVINGROOMS");
			frmScreen.txtBedrooms.value = RoomsXML.GetAttribute("NUMBEROFBEDROOMS");
			frmScreen.txtBathrooms.value = RoomsXML.GetAttribute("BATHROOMS");
			frmScreen.txtWCs.value = RoomsXML.GetAttribute("SEPERATEWCS");
			frmScreen.txtHabitableRooms.value = RoomsXML.GetAttribute("HABITABLEROOMS");
			frmScreen.txtKitchens.value = RoomsXML.GetAttribute("NUMBEROFKITCHENS"); //JD MAR40
			frmScreen.txtGarages.value = RoomsXML.GetAttribute("GARAGES");
			frmScreen.txtParkingSpaces.value = RoomsXML.GetAttribute("PARKINGSPACES");
			SetOptionValue(frmScreen.Conservatory_Yes, frmScreen.Conservatory_No, RoomsXML.GetAttribute("CONSERVATORY"));
			SetOptionValue(frmScreen.GardenIncluded_Yes, frmScreen.GardenIncluded_No, RoomsXML.GetAttribute("GARDENINCLUDED"));	
			// EP2_2 - Add new fields
			frmScreen.txtFloors.value = RoomsXML.GetAttribute("FLOORS");
			SetOptionValue(frmScreen.Balcony_Yes, frmScreen.Balcony_No, RoomsXML.GetAttribute("BALCONYACCESS"));
			
		}
}

function SaveRooms()
{
	var sReturn = new Array();
	if (IsChanged())
	{
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		var roottag = XML.CreateActiveTag("HOME");
		XML.ActiveTag = roottag
		XML.CreateActiveTag("ROOMS");
		XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.SetAttribute("LIVINGROOMS", frmScreen.txtLivingRooms.value);
		XML.SetAttribute("BEDROOMS", frmScreen.txtBedrooms.value);
		XML.SetAttribute("BATHROOMS", frmScreen.txtBathrooms.value);
		XML.SetAttribute("SEPERATEWCS", frmScreen.txtWCs.value);
		XML.SetAttribute("HABITABLEROOMS", frmScreen.txtHabitableRooms.value);
		XML.SetAttribute("NUMBEROFKITCHENS", frmScreen.txtKitchens.value); //JD MAR40
		XML.SetAttribute("GARAGES", frmScreen.txtGarages.value);
		XML.SetAttribute("PARKINGSPACES", frmScreen.txtParkingSpaces.value);
		XML.SetAttribute("CONSERVATORY", GetOptionValue(frmScreen.Conservatory_Yes));
		XML.SetAttribute("GARDENINCLUDED", GetOptionValue(frmScreen.GardenIncluded_Yes));
		// EP2_2 - Add new fields
		XML.SetAttribute("FLOORS", frmScreen.txtFloors.value);
		XML.SetAttribute("BALCONYACCESS", GetOptionValue(frmScreen.Balcony_Yes));
	
		RoomsXML = XML.XMLDocument.xml;
	}
	
	sReturn[0] = IsChanged();
	sReturn[1] = RoomsXML;
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
-->
</script>
</body>
</html>

