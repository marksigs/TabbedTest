<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      dc221.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Rooms Screen (a popup).
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		22/02/00	Created.
AY		31/03/00	scScreenFunctions change
IW		03/05/00	SYS0623
SA		31/07/01	SYS1025 Move OK & Cancel to outside yellow area.
MC		01/10/01	SYS2757 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
MC		19/04/2004	BMIDS517	Meta Tag added (not to cache)
								White space padded to the title.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<head>
<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>DC221 - Rooms <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 15px; LEFT: 35px; HEIGHT: 193px; WIDTH: 180px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
		&nbsp;Living Rooms
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtLivingRooms" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
		&nbsp;Bedrooms
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtBedrooms" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 90px" class="msgLabel">
		&nbsp;Bathrooms
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtBathrooms" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 120px" class="msgLabel">
		&nbsp;<LABEL id=idWCs></LABEL>
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtWCs" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 150px" class="msgLabel">
		&nbsp;Kitchens
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtKitchens" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 200px; LEFT: 0px; POSITION: ABSOLUTE">
		<input id="btnSubmit" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	<span style="TOP: 200px; LEFT: 65px; POSITION: ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	
</div>
</form>

<% /* File containing field attributes */ %>
<!--  #include FILE="attribs/dc221attribs.asp" -->
<!--  #include FILE="Customise/DC221Customise.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_xmlRoomTypesComboList = null;
var m_sRoomTypesXML = null;
var m_XMLRoomTypes = null;
var m_nComboRoomTypes = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sReadOnly = null;
var scScreenFunctions;
var m_BaseNonPopupWindow = null;

function window.onload()
{
	RetrieveData();
	SetMasks();
	// MC SYS2564/SYS2757 for client customisation
	Customise();

	Validation_Init();
	PopulateScreen();
	Initialise();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	window.returnValue = null;
}
function Initialise()
{
	if(m_sReadOnly == "1")
			scScreenFunctions.SetScreenToReadOnly(frmScreen);

}
function PopulateScreen()
{
	<% /* var sGroups = new Array("NewPropertyRoomType");
	m_xmlRoomTypesComboList = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	if (!m_xmlRoomTypesComboList.GetComboLists(document, sGroups))
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnOK.disabled = true;
	} */ %>
	
	m_XMLRoomTypes.ActiveTag = null;
	m_XMLRoomTypes.CreateTagList("NEWPROPERTYROOMTYPELIST");
	if (m_XMLRoomTypes.SelectTagListItem(0))
	{
		var bContinue = true;
		var sRoomType = null;
		var sNumRooms = null;

		var tagList = m_XMLRoomTypes.CreateTagList("NEWPROPERTYROOMTYPE");
		
		for (var nItem = 0; nItem < tagList.length; nItem++)
		{
			m_XMLRoomTypes.SelectTagListItem(nItem);
			sRoomType = m_XMLRoomTypes.GetTagText("ROOMTYPE");
			sNumRooms = m_XMLRoomTypes.GetTagText("NUMBEROFROOMS");
			PopulateField(sRoomType,sNumRooms);
		}		
	}
}

function PopulateField(vsRoomType,vsNumRooms)
{
	switch (vsRoomType)
	{
	case "":
		break;
	case "1":
	{
		frmScreen.txtLivingRooms.value = vsNumRooms;
		break;
	}
	case "2":
	{
		frmScreen.txtBedrooms.value = vsNumRooms;
		break;
	}
	case "3":
	{
		frmScreen.txtBathrooms.value = vsNumRooms;
		break;
	}
	case "4":
	{
		frmScreen.txtWCs.value = vsNumRooms;
		break;
	}
	case "5":
	{
		frmScreen.txtKitchens.value = vsNumRooms;
		break;
	}
	default:
	{
		alert('Invalid room type: ' + vsRoomType);
		break;
	}
	}
}

function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters	= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sRoomTypesXML	= sParameters[0];
	m_sReadOnly	= sParameters[1];
	m_sApplicationNumber = sParameters[2];
	m_sApplicationFactFindNumber = sParameters[3];
	
	m_XMLRoomTypes = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_XMLRoomTypes.LoadXML(m_sRoomTypesXML);
}

function frmScreen.btnSubmit.onclick()
{
	var sReturn = new Array();
	if (IsChanged())
	{
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		
		var tagNEWPROPERTYROOMTYPELIST = XML.CreateActiveTag("NEWPROPERTYROOMTYPELIST");
		
		XML.CreateActiveTag("NEWPROPERTYROOMTYPE");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("ROOMTYPE","1");
		XML.CreateTag("NUMBEROFROOMS",frmScreen.txtLivingRooms.value);
		
		XML.ActiveTag = tagNEWPROPERTYROOMTYPELIST;

		XML.CreateActiveTag("NEWPROPERTYROOMTYPE");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("ROOMTYPE","2");
		XML.CreateTag("NUMBEROFROOMS",frmScreen.txtBedrooms.value);

		XML.ActiveTag = tagNEWPROPERTYROOMTYPELIST;

		XML.CreateActiveTag("NEWPROPERTYROOMTYPE");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("ROOMTYPE","3");
		XML.CreateTag("NUMBEROFROOMS",frmScreen.txtBathrooms.value);

		XML.ActiveTag = tagNEWPROPERTYROOMTYPELIST;

		XML.CreateActiveTag("NEWPROPERTYROOMTYPE");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("ROOMTYPE","4");
		XML.CreateTag("NUMBEROFROOMS",frmScreen.txtWCs.value);

		XML.ActiveTag = tagNEWPROPERTYROOMTYPELIST;

		XML.CreateActiveTag("NEWPROPERTYROOMTYPE");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("ROOMTYPE","5");
		XML.CreateTag("NUMBEROFROOMS",frmScreen.txtKitchens.value);
		
		m_sRoomTypesXML = XML.XMLDocument.xml;
	}
	
	sReturn[0] = IsChanged();
	sReturn[1] = m_sRoomTypesXML;
	m_XMLRoomTypes = null;
	window.returnValue	= sReturn;
	window.close();
}

function frmScreen.btnCancel.onclick()
{
	m_XMLRoomTypes = null;
	window.returnValue	= null;
	window.close();
}
-->
</script>
</body>
</html>