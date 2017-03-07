<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      ZA030.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   This frame is used to store the new name & address in the
				name & address directory.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DPol	26/11/1999	Created.
DPol	06/12/1999	Added code now that BO completed.
AD		02/02/2000	Rework
AY		03/04/00	scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client

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
	<title>ZA030 - Add Name And Address To Directory <!-- #include file="includes/TitleWhiteSpace.asp" --> </title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp 
VIEWASTEXT></OBJECT>

<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>

<%/* FORMS */%>
<form id="frmScreen" >
<div id="divBackground" style="LEFT: 10px; WIDTH: 491px; POSITION: absolute; TOP: 10px; HEIGHT: 220px" class="msgGroup">
	<span style="LEFT: 31px; WIDTH: 400px; POSITION: absolute; TOP: 8px" class="msgLabel">
		Head Office?
		<span style="LEFT: 200px; POSITION: absolute; TOP: -3px">
			<input id="optHeadOfficeYes" name="HeadOfficeRadioGroup" type="radio" value="1">
			<label for="optHeadOfficeYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 254px; POSITION: absolute; TOP: -3px">
			<input id="optHeadOfficeNo" name="HeadOfficeRadioGroup" type="radio" value="0">
			<label for="optHeadOfficeNo" class="msgLabel">No</label>
		</span> 
	</span>		

	<span style="LEFT: 31px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Notes
		<span style="LEFT: 0px; POSITION: absolute; TOP: 16px"><TEXTAREA class=msgTxt id=txtNotes style="WIDTH: 300px; POSITION: absolute" name=Notes rows=5 maxlength="500"></TEXTAREA> 
		</span>
	</span> 

	<span style="LEFT: 31px; POSITION: absolute; TOP: 140px" class="msgLabel" id="spnOrganisationType">
		Organisation Type
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<select id="cboOrganisationType" name="Organisation Type" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="LEFT: 31px; WIDTH: 400px; POSITION: absolute; TOP: 164px" class="msgLabel" id="spnPartOfPanel">
		Part of Mortgage Lender Panel?
		<span style="LEFT: 200px; POSITION: absolute; TOP: -3px">
			<input id="optPartOfPanelYes" name="PartOfPanelRadioGroup" type="radio" value="1">
			<label for="optPartOfPanelYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 254px; POSITION: absolute; TOP: -3px">
			<input id="optPartOfPanelNo" name="PartOfPanelRadioGroup" type="radio" value="0">
			<label for="optPartOfPanelNo" class="msgLabel">No</label>
		</span> 
	</span>		

	<span style="LEFT: 31px; POSITION: absolute; TOP: 188px" class="msgLabel" id="spnMortgageLender">
		Mortgage Lender
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<select id="cboMortgageLender" name="Mortgage Lender" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="LEFT: 31px; POSITION: absolute; TOP: 224px">
		<input id="btnOK" value="OK" type="button" style="LEFT: 31px; WIDTH: 60px" class="msgButton">
		<span style="LEFT: 70px; POSITION: absolute; TOP: 0px">
			<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>
</div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/za030Attribs.asp" -->

<%/* SCRIPT */%>
<script language="JScript">
<!--
var m_sMetaAction = null;
var DirectoryXML = null;
var AttributeArray = null;
var m_bPartOfPanelInd = null;
var m_bHeadOfficeInd = null;
var m_bComboResultsFound = null;
var m_iNameAndAddressType = null;
var m_iBANK = 3;	// hard coded this so that should it change it changes here in one place.
var scScreenFunctions;
var m_BaseNonPopupWindow = null;

function frmScreen.btnCancel.onclick()
{
	window.close();
}

function frmScreen.btnOK.onclick()
{
	var bSuccess = true;

	// Perform Validation
	if (m_bPartOfPanelInd & (frmScreen.cboMortgageLender.selectedIndex == 0))
	{
		alert("Panel Lender has not been selected");
		return;
	}

	with (DirectoryXML)
	{
		SelectTag(null,"NAMEANDADDRESSDIRECTORY");
		CreateTag("HEADOFFICEINDICATOR", m_bHeadOfficeInd ? "1" : "0");
		CreateTag("NOTES", frmScreen.txtNotes.value);
		SetTagText("ORGANISATIONTYPE", frmScreen.cboOrganisationType.value);
		<% /* MO - BMIDS00807 */ %>
		<% /* CreateTag("NAMEANDADDRESSACTIVEFROM", scScreenFunctions.DateToString(scScreenFunctions.GetSystemDate()));  */ %>
		CreateTag("NAMEANDADDRESSACTIVEFROM", scScreenFunctions.DateToString(scScreenFunctions.GetAppServerDate()));  
		// 		RunASP(document,"CreateDirectory.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					RunASP(document,"CreateDirectory.asp");
				break;
			default: // Error
				SetErrorResponse();
			}

	}

	// Check the response and retrieve the created record's Directory GUID
	if (!DirectoryXML.IsResponseOK())
		return;
	sDirectoryGUID = DirectoryXML.GetTagText("DIRECTORYGUID")
	var bDirectoryAddressInd = true;
	var sOrganisationId = "";

	if ((m_bPartOfPanelInd) & (m_iNameAndAddressType == m_iBANK))
	{
		// Create the Mortgage Directory record
		with (frmScreen.cboMortgageLender)
			sOrganisationId = options(selectedIndex).getAttribute("OrganisationId");

		XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

		with (XML)
		{
			CreateRequestTagFromArray(AttributeArray,null)
			CreateActiveTag("MORTGAGELENDERDIRECTORY");
			CreateTag("ORGANISATIONID", sOrganisationId);
			CreateTag("DIRECTORYGUID", sDirectoryGUID);
			CreateTag("MAINMORTGAGELENDERIND", "0");
			// 			RunASP(document,"CreateMortgageLenderDirectory.asp")
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							RunASP(document,"CreateMortgageLenderDirectory.asp")
					break;
				default: // Error
					SetErrorResponse();
				}

			if (!IsResponseOK())
				return;
		}
	}

	FlagChange(true);

	// Return the new DirectoryGUID and OrganisationID to the calling screen
	var sReturn = new Array();
	sReturn[0] = IsChanged();
	sReturn[1] = sDirectoryGUID;
	sReturn[2] = frmScreen.cboOrganisationType.value;

	XML = null;
	window.returnValue = sReturn;
	window.close();
}

function frmScreen.optHeadOfficeNo.onclick()
{
	m_bHeadOfficeInd = false;
}

function frmScreen.optHeadOfficeYes.onclick()
{
	m_bHeadOfficeInd = true;
}

function frmScreen.optPartOfPanelNo.onclick()
{
	m_bPartOfPanelInd = false;
	ShowOrHideMortgageLender();
}

function frmScreen.optPartOfPanelYes.onclick()
{
	m_bPartOfPanelInd = true;
	ShowOrHideMortgageLender();
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	RetrieveData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	scScreenFunctions.ShowCollection(frmScreen);
	scScreenFunctions.SetRadioGroupValue(frmScreen, "PartOfPanelRadioGroup", frmScreen.optPartOfPanelNo.value);
	scScreenFunctions.SetRadioGroupValue(frmScreen, "HeadOfficeRadioGroup", frmScreen.optHeadOfficeNo.value);
	frmScreen.optPartOfPanelNo.onclick();

	// populate the combo with the Mortgage Lenders names.
	m_bComboResultsFound = false;
	if (m_iNameAndAddressType == m_iBANK)
	{
		PopulateCombos();
		var sOrganisationType = DirectoryXML.GetTagText("ORGANISATIONTYPE");
		if (sOrganisationType == "")
			frmScreen.cboOrganisationType.selectedIndex = 0;
		else
			frmScreen.cboOrganisationType.value = sOrganisationType;
		m_bComboResultsFound = PopulateMortgageLenderList();
	}

	// Hide/show the Mortgage Lender combo and Panel option buttons as appropriate
	ShowOrHideMortgageLender();
	if (!m_bComboResultsFound) scScreenFunctions.HideCollection(spnPartOfPanel);

	// set focus on the HeadOffice Yes
	frmScreen.optHeadOfficeYes.focus();
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function PopulateCombos()
// Populates all combos on the screen
{
	var blnSuccess = true;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("OrganisationType");

	if(XML.GetComboLists(document,sGroupList))
		// Populate Country combo
		blnSuccess = XML.PopulateCombo(document,frmScreen.cboOrganisationType,"OrganisationType",true);

	XML = null;
}

function PopulateMortgageLenderList()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bResultsFound = false;
	var sASPFile;
	XML.CreateRequestTagFromArray(AttributeArray,null)
	XML.RunASP(document,"FindMainMortgageLenderList.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		// Populate the Combo from the result xml
		// create a tag list based on NAMEANDADDRESSDIRECTORY
		var TagListMORTGAGELENDERDIRECTORY = XML.CreateTagList("MORTGAGELENDERDIRECTORY");
		var iNoOfMortgageLenderDirectories = XML.ActiveTagList.length;

		//check there are some mortgage lender directories to retrieve
		if (iNoOfMortgageLenderDirectories > 0)
		{
			bResultsFound = true;
			// Add a <SELECT> option
			var TagOPTION	= document.createElement("OPTION");
			TagOPTION.value	= "";
			TagOPTION.text	= "<SELECT>";
			frmScreen.cboMortgageLender.add(TagOPTION);

			// Loop through all customer context entries
			for(var iLoop = 0; iLoop < iNoOfMortgageLenderDirectories; iLoop++)
			{
				XML.SelectTagListItem(iLoop);
				var sOrganisationId	= XML.GetTagText("ORGANISATIONID");
				var sDirectoryGuid = XML.GetTagText("DIRECTORYGUID");
				var sCompanyName = XML.GetTagText("COMPANYNAME");

				TagOPTION = document.createElement("OPTION");
				TagOPTION.value	= sCompanyName;
				TagOPTION.text = sCompanyName;
				TagOPTION.setAttribute("DirectoryGuid", sDirectoryGuid);
				TagOPTION.setAttribute("OrganisationId", sOrganisationId);

				frmScreen.cboMortgageLender.add(TagOPTION);
			}
			frmScreen.cboMortgageLender.selectedIndex = 0;
		}
	}
	return (bResultsFound);
}

function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	AttributeArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	var sDirectoryXML = AttributeArray[4];

	DirectoryXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	DirectoryXML.LoadXML(sDirectoryXML);
	DirectoryXML.SelectTag(null, "REQUEST");

	// Extract the Name And Address Type from the XML string passed over
	m_iNameAndAddressType = DirectoryXML.GetTagText("NAMEANDADDRESSTYPE");
}

function ShowOrHideMortgageLender()
{
	if ((m_iNameAndAddressType != m_iBANK) | !m_bPartOfPanelInd | !m_bComboResultsFound)
		scScreenFunctions.HideCollection(spnMortgageLender);
	else
		scScreenFunctions.ShowCollection(spnMortgageLender);

	if (m_iNameAndAddressType != m_iBANK)
		scScreenFunctions.HideCollection(spnOrganisationType);
	else
		scScreenFunctions.ShowCollection(spnOrganisationType);
}
-->
</script>
</body>
</html>




