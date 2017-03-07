<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      DC235.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Add/Edit Other Resident
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		08/03/2000	Created
AY		31/03/00	New top menu/scScreenFunctions change
MH      10/05/00    SYS0629 DOB & Tooltiptext.
BG		17/02/2000	SYS0744 - Changed label "Gender" to "Sex"
BG		17/05/00	SYS0752 Removed Tooltips
CL		05/03/01	SYS1920 Read only functionality added
LD		23/05/02	SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
MO		20/11/2002	BMIDS00376	Data freezing errors
*/ %>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM 2 Specific History

Prog	Date		AQR		Description
PE		14/02/2007	EP2_745	New field "Relationship to Applicant"
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
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<% /* FORMS */ %>
<form id="frmToDC230" method="post" action="DC230.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange" year4>
<div style="HEIGHT: 108px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Forenames
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
			<input id="txtFirstForename" maxlength="30" style="POSITION: absolute; WIDTH: 150px" class="msgTxt">
		</span> 

		<span style="LEFT: 240px; POSITION: absolute; TOP: -3px">
			<input id="txtSecondForename" maxlength="30" style="POSITION: absolute; WIDTH: 150px" class="msgTxt">
		</span> 

		<span style="LEFT:400px; POSITION: absolute; TOP: -3px">
			<input id="txtOtherForenames" maxlength="30" style="POSITION: absolute; WIDTH: 150px" class="msgTxt">
		</span> 
	</span>

	<span style="TOP: 34px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Surname
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
			<input id="txtSurname" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgTxt">
		</span> 
	</span>

	<span style="TOP: 58px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Date of Birth
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
			<input id="txtDateOfBirth" maxlength="10" style="POSITION: absolute; WIDTH: 80px" class="msgTxt">
		</span> 
	</span>

	<span style="TOP: 82px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Sex
		<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
			<select id="cboGender" style="WIDTH: 80px" class="msgCombo"></select>
		</span>
	</span>

	<span style="TOP: 106px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<span style="WIDTH: 60px">Relationship to Applicant</span>		
		<span style="TOP: -3px; LEFT: 80px; POSITION: ABSOLUTE">
			<select id="cboRelationship" style="WIDTH: 150px" class="msgCombo"></select>
		</span>
	</span>

</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/DC235attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sOtherResidentSequenceNumber = "";
var m_sPersonGUID = "";
var OtherResidentXML = null;
var scScreenFunctions;
var m_blnReadOnly = false;


/* EVENTS */

function btnAnother.onclick()
{
	if (CommitChanges())
		Initialise(false);
}

function btnCancel.onclick()
{
	frmToDC230.submit();
}

function btnSubmit.onclick()
{
	if (CommitChanges())
		frmToDC230.submit();
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function IsDOBValid ()
{
	var dteDOB;
	var dteToday;
	var dteVeryOld;
	
	dteDOB=scScreenFunctions.GetDateObject(frmScreen.txtDateOfBirth);
	<% /* MO - BMIDS00807 */ %>
	dteToday = scScreenFunctions.GetAppServerDate(true);
	<% /* dteToday = new Date(); */ %>
	
	if (dteDOB > dteToday) 
	{
		alert ("Date of Birth must be in the past");
		frmScreen.txtDateOfBirth.focus();
		return false;
	}
	dteVeryOld=new Date(dteToday.getYear()-100,dteToday.getMonth(),dteToday.getDate())	
	
	if (dteDOB < dteVeryOld) 
	{
		alert ("Date of Birth must be less than 100 years in the past");
		frmScreen.txtDateOfBirth.focus();
		return false;
	}
	return true;
}

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel","Another");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Add/Edit Other Resident","DC235",scScreenFunctions);

	RetrieveContextData();
	SetMasks();

	Validation_Init();	
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC235");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	Initialise(true);
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

function ClearFields()
// Clears all fields of data
{
	frmScreen.txtDateOfBirth.value = "";
	frmScreen.txtFirstForename.value = "";
	frmScreen.txtOtherForenames.value = "";
	frmScreen.txtSecondForename.value = "";
	frmScreen.txtSurname.value = "";
	frmScreen.cboGender.selectedIndex = 0;
	frmScreen.cboRelationship.selectedIndex = 0;
	m_sPersonGUID = "";
	m_sOtherResidentSequenceNumber = "";
}

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				if (IsDOBValid())
					bSuccess = SaveOtherResident();
				else
					bSuccess = false;
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function DefaultFields()
// Inserts default values into all fields
{
	ClearFields();
}

function Initialise(bOnLoad)
// Initialises the screen
{
	if(bOnLoad == true)
		PopulateCombos();

	if(m_sMetaAction == "Edit")
	{
		PopulateScreen();
		DisableMainButton("Another");
	}
	else
		DefaultFields();

	if(m_sReadOnly == "1")
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			DisableMainButton("Another");
			<% /* MO 20/11/2002 BMIDS00376
			DisableMainButton("Cancel"); */ %>
		}	

	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateCombos()
// Populates all combos on the screen
{
	var blnSuccess = true;
	ComboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("Sex", "ApplicantRelationship");

	if(ComboXML.GetComboLists(document,sGroupList))
	{
		blnSuccess = ComboXML.PopulateCombo(document,frmScreen.cboGender,"Sex",true);
		blnSuccess = ComboXML.PopulateCombo(document,frmScreen.cboRelationship,"ApplicantRelationship",true);

		if(!blnSuccess) scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
}

function PopulateScreen()
// Populates the screen with details of the item selected in dc160
{
	OtherResidentXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	OtherResidentXML.CreateRequestTag(window,null);
	OtherResidentXML.CreateActiveTag("OTHERRESIDENT");
	OtherResidentXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	OtherResidentXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	OtherResidentXML.CreateTag("OTHERRESIDENTSEQUENCENUMBER", m_sOtherResidentSequenceNumber);

	// 	OtherResidentXML.RunASP(document,"GetOtherResident.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			OtherResidentXML.RunASP(document,"GetOtherResident.asp");
			break;
		default: // Error
			OtherResidentXML.SetErrorResponse();
		}


	if(OtherResidentXML.SelectTag(null, "OTHERRESIDENT") != null)
	{
		with (frmScreen)
		{
			m_sPersonGUID = OtherResidentXML.GetTagText("PERSONGUID");
			txtDateOfBirth.value = OtherResidentXML.GetTagText("DATEOFBIRTH");
			txtFirstForename.value = OtherResidentXML.GetTagText("FIRSTFORENAME");
			txtOtherForenames.value = OtherResidentXML.GetTagText("OTHERFORENAMES");
			txtSecondForename.value = OtherResidentXML.GetTagText("SECONDFORENAME");
			txtSurname.value = OtherResidentXML.GetTagText("SURNAME");
			cboGender.value = OtherResidentXML.GetTagText("GENDER");
			cboRelationship.value = OtherResidentXML.GetTagText("RELATIONSHIPTOAPPLICANT");
		}
	}
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","1325");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	if (m_sMetaAction == "Edit")
		m_sOtherResidentSequenceNumber = scScreenFunctions.GetContextParameter(window,"idOtherResidentSequenceNumber","1");    
}

function SaveOtherResident()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,null)
	XML.CreateActiveTag("OTHERRESIDENT");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	if (m_sMetaAction == "Edit")
		XML.CreateTag("OTHERRESIDENTSEQUENCENUMBER", m_sOtherResidentSequenceNumber);
	XML.CreateTag("PERSONGUID",m_sPersonGUID);
	XML.CreateActiveTag("PERSON");
	XML.CreateTag("PERSONGUID",m_sPersonGUID);
	XML.CreateTag("FIRSTFORENAME",frmScreen.txtFirstForename.value);
	XML.CreateTag("SECONDFORENAME",frmScreen.txtSecondForename.value);
	XML.CreateTag("OTHERFORENAMES",frmScreen.txtOtherForenames.value);
	XML.CreateTag("SURNAME",frmScreen.txtSurname.value);
	XML.CreateTag("DATEOFBIRTH",frmScreen.txtDateOfBirth.value);
	XML.CreateTag("GENDER",frmScreen.cboGender.value);
	XML.CreateTag("RELATIONSHIPTOAPPLICANT",frmScreen.cboRelationship.value);
	if (m_sMetaAction == "Edit")
		// 		XML.RunASP(document,"UpdateOtherResident.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateOtherResident.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	else
		// 		XML.RunASP(document,"CreateOtherResident.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateOtherResident.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}


	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}
-->
</script>
</body>
</html>


