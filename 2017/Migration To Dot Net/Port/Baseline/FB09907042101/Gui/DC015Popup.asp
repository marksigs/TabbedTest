<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/* 
Workfile:      DC015.asp
Copyright:     Copyright © 1999 Marlborough Stirling
Description:   Intermediary Search Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog  Date     Description
JLD		04/11/99	Created
JLD		10/12/1999	DC/013 - Added scrolling to listbox
					DC/015 - Changed the look of the gui
					DC/021 - set address format ok
AD		30/01/2000	Rework
AY		14/02/00	Change to msgButtons button types
AY		29/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MO		29/09/2002	BMIDS00502 - INWP1, BM061, Change to introducer search to show level introducer 
						id's and modify search to call BM's system.
MO		14/10/2002	BMIDS00663 - Fixed bug where the value of direct/indirect was passed and not the 
						validation type, put in place the 'proper' introducer services.  Also made
						the change so that omAdmin was called rather than omInt
MO		24/10/2002	BMIDS00713 - Fixed bug which ment the attribs werent loading properly.
MO		14/11/2002	BMIDS00936 - Changed the town label to Town / Postcode
HMA     18/09/2003  BM0063     - Amend HTML text for radio buttons.
MC		20/04/2004	BMIDS517	white space padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
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
	<title>DC015P - Find Intermediary  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>
<span id="spnListScroll">
	<span style="LEFT: 230px; POSITION: absolute; TOP: 295px">
		<object data="scTableListScroll.asp" id="scScrollTable" 
			style="LEFT: 0px; TOP: 0px; height:24; width:304" 
			type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
	</span> 
</span>	
<form id="frmToDC010" method="post" action="DC010.asp" STYLE="DISPLAY: none">
</form>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 310px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Search</strong>
	</span>

	<span style="TOP: 30px; LEFT: 8px; POSITION: ABSOLUTE" class="msgLabel">
		MCCB/BMId
		<span style="TOP: -3px; LEFT: 60px; POSITION: ABSOLUTE">
			<input id="txtBMId" name="BMId" maxlength="12" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 30px; LEFT: 250px; POSITION: ABSOLUTE" class="msgLabel">
		Town / Postcode
		<span style="TOP: -3px; LEFT: 105px; POSITION: ABSOLUTE">
			<input id="txtTown" name="Town" maxlength="40" style="WIDTH: 225px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 55px; LEFT: 8px; POSITION: ABSOLUTE" class="msgLabel">
		Name 1
		<span style="TOP: -3px; LEFT: 60px; POSITION: ABSOLUTE">
			<input id="txtName1" name="Name1" maxlength="40" style="WIDTH: 225px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 55px; LEFT: 295px; POSITION: ABSOLUTE" class="msgLabel">
		Name 2
		<span style="TOP: -3px; LEFT: 60px; POSITION: ABSOLUTE">
			<input id="txtName2" name="Name2" maxlength="40" style="WIDTH: 225px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span style="TOP: 80px; LEFT: 8px; POSITION: ABSOLUTE" class="msgLabel">
		Packager?
		<span style="TOP: -3px; LEFT: 60px; POSITION: ABSOLUTE">
			<input id="optPackagerYes" name="Packager" type="radio" value="1"><label for="optPackagerYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="optPackagerNo" name="Packager" type="radio" value="0" checked><label for="optPackagerNo" class="msgLabel">No</label>
		</span> 
	</span>
	
	<span style="TOP: 80px; LEFT: 520px; POSITION: ABSOLUTE">
		<input id="btnSearch" value="Search" type="button" style="WIDTH: 60px" class="msgButton">
	</span>

	<span style="TOP: 105px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Intermediary List</strong>
	</span>
	<span id="spnIntermedListTable" style="TOP: 125px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblIntermedList" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="20%" class="TableHead">Introducer Id</td>		<td width="30%" class="TableHead">Name</td>					<td width="50%" class="TableHead">Address</td></tr>
			<tr id="row01">		<td width="20%" class="TableTopLeft">&nbsp</td>				<td width="30%" class="TableTopCenter">&nbsp</td>			<td class="TableTopRight">&nbsp</td></tr>
			<tr id="row02">		<td width="20%" class="TableLeft">&nbsp</td>				<td width="30%" class="TableCenter">&nbsp</td>				<td class="TableRight">&nbsp</td></tr>
			<tr id="row03">		<td width="20%" class="TableLeft">&nbsp</td>				<td width="30%" class="TableCenter">&nbsp</td>				<td class="TableRight">&nbsp</td></tr>
			<tr id="row04">		<td width="20%" class="TableLeft">&nbsp</td>				<td width="30%" class="TableCenter">&nbsp</td>				<td class="TableRight">&nbsp</td></tr>
			<tr id="row05">		<td width="20%" class="TableLeft">&nbsp</td>				<td width="30%" class="TableCenter">&nbsp</td>				<td class="TableRight">&nbsp</td></tr>
			<tr id="row06">		<td width="20%" class="TableLeft">&nbsp</td>				<td width="30%" class="TableCenter">&nbsp</td>				<td class="TableRight">&nbsp</td></tr>
			<tr id="row07">		<td width="20%" class="TableLeft">&nbsp</td>				<td width="30%" class="TableCenter">&nbsp</td>				<td class="TableRight">&nbsp</td></tr>
			<tr id="row08">		<td width="20%" class="TableLeft">&nbsp</td>				<td width="30%" class="TableCenter">&nbsp</td>				<td class="TableRight">&nbsp</td></tr>
			<tr id="row09">		<td width="20%" class="TableLeft">&nbsp</td>				<td width="30%" class="TableCenter">&nbsp</td>				<td class="TableRight">&nbsp</td></tr>
			<tr id="row10">		<td width="20%" class="TableBottomLeft">&nbsp</td>			<td width="30%" class="TableBottomCenter">&nbsp</td>		<td class="TableBottomRight">&nbsp</td></tr>
		</table>
	</span>
</div>
</form>

<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 330px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- Specify Code Here -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/DC015PopupAttribs.asp" -->
<script language="JScript">
<!--
// JScript Code
var m_sMetaAction = null;
var IntermedXML = null;
var m_iTableLength = 10;
var scScreenFunctions;
var m_sGN210Data = "";
var m_GN210XML = null;
var strDirectIndirectBusiness = "";
var strDirectIndirectBusinessValType = "";
var sParameters;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	
	RetrieveData();
	
	//This function is contained in the field attributes file (remove if not required)
	SetMasks();
	Validation_Init();
	
	frmScreen.txtBMId.focus();
	DisableMainButton("Submit");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

<% // MO - 29/09/2002 - BMIDS00502 %>
function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	sParameters		= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	<% // MO - 29/09/2002 - BMIDS00502 - Get the direct/indirect status from the parameters passed in %>
	<% // MO - 29/09/2002 - BMIDS00663 - fix bug get the direct/indirect validation type %>
	strDirectIndirectBusinessValType = sParameters[4];
	
}

<% // MO - 29/09/2002 - BMIDS00502 %>
function SetFocusToEmptyField()
{
	if( IsStringEmpty(frmScreen.txtBMId.value) )
	{
		frmScreen.txtBMId.value = "";
		frmScreen.txtBMId.focus();
	}
	else if( IsStringEmpty(frmScreen.txtName1.value))
	{
		frmScreen.txtName1.value = "";
		frmScreen.txtName1.focus();
	}
	else if( IsStringEmpty(frmScreen.txtName2.value))
	{
		frmScreen.txtName2.value = "";
		frmScreen.txtName2.focus();
	}
	else if( IsStringEmpty(frmScreen.txtTown.value))
	{
		frmScreen.txtTown.value = "";
		frmScreen.txtTown.focus();
	}
	else
	{
		<% // default to bmid %>
		frmScreen.txtBMId.value = "";
		frmScreen.txtBMId.focus();
	}	
}

function IsStringEmpty(strString)
{
	var bStringIsEmpty = false;
	var ssArray = strString.split(" ");
	var nLength = ssArray.length;
	var nIndex = 0;
	while(ssArray[nIndex] == "" && nIndex < nLength)
		nIndex++;

	if(ssArray[nIndex] == null || ssArray[nIndex] == "*")
		bStringIsEmpty = true;

	return(bStringIsEmpty);
}

function frmScreen.btnSearch.onclick()
{
	<% //check that all fields are occupied %>
	if( IsStringEmpty(frmScreen.txtBMId.value) &&
		IsStringEmpty(frmScreen.txtName1.value) &&
		IsStringEmpty(frmScreen.txtName2.value) &&
		IsStringEmpty(frmScreen.txtTown.value))
	{
		alert("You must enter search criteria");
		SetFocusToEmptyField();
	}
	else
	{
		<% //all present and correct, populate the listbox %>
		if(FindIndividualIntermediary())
		{
			scScrollTable.setRowSelected(1);
			EnableMainButton("Submit");
			btnSubmit.focus();
		}
		<% /* MO - 14/10/2002 - BMIDS00663 No longer needed as the error is returned by the BM Introducer system else
		{
			scScrollTable.clear();
			alert("Unable to find any intermediaries for your search criteria. Please amend the search criteria and/or wildcard your search");
			DisableMainButton("Submit");
			frmScreen.txtBMId.focus();
		} */ %>
	}
}

<% // MO - 29/09/2002 - BMIDS00502 %>
function PopulateListBox(nStart)
{
	IntermedXML.ActiveTag = null;
	IntermedXML.CreateTagList("INTRODUCER");
	var iNumberOfIntermediaries = IntermedXML.ActiveTagList.length;

	if(iNumberOfIntermediaries > 0)
	{
		scScrollTable.initialiseTable(tblIntermedList, 0, "", ShowList, m_iTableLength, iNumberOfIntermediaries);
		ShowList(nStart);
		return true;
	} else {
		return false;
	}
}

<% // MO - 29/09/2002 - BMIDS00502 %>
function ShowList(nStart)
{
	var iCount;
	scScrollTable.clear();
	for (iCount = 0; iCount < IntermedXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		IntermedXML.ActiveTag = null;
		IntermedXML.CreateTagList("INTRODUCER");  
		IntermedXML.SelectTagListItem(iCount + nStart);

		var sBMReference = IntermedXML.GetAttribute("BMREFERENCE");
		var sName =  IntermedXML.GetAttribute("NAME");
		var sPostcode = IntermedXML.GetAttribute("POSTCODE");
		var sFlatNumber = IntermedXML.GetAttribute("FLATNUMBER");
		var sHouseName = IntermedXML.GetAttribute("BUILDINGNAME");
		var sHouseNumber = IntermedXML.GetAttribute("BUILDINGNO");
		var sStreet = IntermedXML.GetAttribute("STREET");
		var sDistrict = IntermedXML.GetAttribute("DISTRICT");
		var sTown = IntermedXML.GetAttribute("TOWN");
		var sCounty = IntermedXML.GetAttribute("COUNTY");
		var sCountry = IntermedXML.GetAttribute("COUNTRY");

		<% //create an intelligent address line %>
		var sAddress = "";
		if(sPostcode != "") sAddress += sPostcode + ",";
		if(sFlatNumber != "") sAddress += sFlatNumber + ",";
		if(sHouseName != "") sAddress += sHouseName + ",";
		if(sHouseNumber != "") sAddress += sHouseNumber + ",";
		if(sStreet != "") sAddress += sStreet + ",";
		if(sDistrict != "") sAddress += sDistrict + ",";
		if(sTown != "") sAddress += sTown + ",";
		if(sCounty != "") sAddress += sCounty + ",";
		if(sCountry != "") sAddress += sCountry;

		<% // Add to the search table %>
		scScreenFunctions.SizeTextToField(tblIntermedList.rows(iCount+1).cells(0),sBMReference);
		scScreenFunctions.SizeTextToField(tblIntermedList.rows(iCount+1).cells(1),sName);
		scScreenFunctions.SizeTextToField(tblIntermedList.rows(iCount+1).cells(2),sAddress);
		tblIntermedList.rows(iCount+1).setAttribute("IntroducerId", sBMReference);
	}
}

<% // MO - 29/09/2002 - BMIDS00502 %>
function FindIndividualIntermediary()
{
	var bSuccess = true;

	IntermedXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	with (IntermedXML)
	{			
		CreateRequestTagFromArray(sParameters, "FindIntroducer")
		SetAttribute("BMMessageType", "FindIntroducer");
		CreateActiveTag("INTRODUCER");
		if (frmScreen.txtBMId.value.length > 0) SetAttribute("BMREFERENCE", frmScreen.txtBMId.value);
		if (frmScreen.txtName1.value.length > 0) SetAttribute("NAME1", frmScreen.txtName1.value);
		if (frmScreen.txtName2.value.length > 0) SetAttribute("NAME2", frmScreen.txtName2.value);
		if (frmScreen.txtTown.value.length > 0) SetAttribute("TOWNPOSTCODE", frmScreen.txtTown.value);
		SetAttribute("DIRECTINDIRECT", strDirectIndirectBusinessValType);
		if (frmScreen.optPackagerYes.checked == true) {
			SetAttribute("PACKAGER", "1");
		} else {
			SetAttribute("PACKAGER", "0");
		}
		// 		RunASP(document, "omAdmin.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					RunASP(document, "omAdmin.asp");
				break;
			default: // Error
				SetErrorResponse();
			}

	}
	
	var bSuccess = IntermedXML.IsResponseOK();
	if (bSuccess ==  true) {
		<% /* Populate Introducer details */ %>
		if (PopulateListBox(0) == false) {
			<% /* No Introducers found */ %>
			bSuccess = false;
		}
	} else { 
		<% /* disable buttons, so that nothing can be done with Introducers.*/ %>
		scScrollTable.clear();
		DisableMainButton("Submit");
		frmScreen.txtBMId.focus();
	}
	
	
	return(bSuccess);
}

function btnSubmit.onclick()
{
	// Store Intermediary guid from the attribute in the <tr> in XML as idXML2 context
	var nRowSelected =  scScrollTable.getRowSelected();
	if (nRowSelected > 0 )
	{
		var sReturn = new Array();
		sReturn[0] = tblIntermedList.rows(nRowSelected).cells(0).innerText;
		sReturn[1] = tblIntermedList.rows(nRowSelected).cells(1).innerText;
		sReturn[2] = tblIntermedList.rows(nRowSelected).getAttribute("IntroducerId");
		window.returnValue = sReturn;
	}
	window.close();
}

function btnCancel.onclick()
{
	window.close();
}

-->
</script>
</body>

</html>



