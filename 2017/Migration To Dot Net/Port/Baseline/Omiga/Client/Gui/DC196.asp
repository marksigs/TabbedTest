<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/*
Workfile:      DC196.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Add/Edit Other Income
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		21/02/2000	Created
AY		31/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
MC		09/06/00	SYS0866 Standardise format of balance and payment fields
CL		05/03/01	SYS1920 Read only functionality added
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MV		22/08/2002	BMIDS00355	IE 5.5 Upgrade - Modified the Msgbuttons Position 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :

Prog	Date		Description
SD		14/11/2005	MAR258	Critical data check
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
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
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<% /* FORMS */ %>
<form id="frmToDC195" method="post" action="DC195.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange" year4>
<div style="HEIGHT: 110px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Add/Edit Other Income</strong> 
	</span>

	<span style="TOP: 36px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Description
		<span style="TOP: -3px; LEFT: 90px; POSITION: ABSOLUTE">
			<select id="cboDescription" style="WIDTH: 200px" class="msgCombo"></select>
		</span>
	</span>

	<span style="TOP: 60px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Amount
		<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
			<input id="txtAmount" maxlength="6" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span> 
	</span>

	<span style="TOP: 84px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Frequency
		<span style="TOP: -3px; LEFT: 90px; POSITION: ABSOLUTE">
			<select id="cboFrequency" style="WIDTH: 200px" class="msgCombo"></select>
		</span>
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 200px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/DC196attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sUnearnedIncomeSequenceNumber = "";
var UnearnedIncomeXML = null;
var scScreenFunctions;
var m_blnReadOnly = false;


/* EVENTS */

function btnAnother.onclick()
{
	if (CommitChanges())
		Initialise();
}

function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
	scScreenFunctions.SetContextParameter(window,"idEmploymentStatusType", "");
	frmToDC195.submit();
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		//clear the contexts
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		frmToDC195.submit();
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
	FW030SetTitles("Add/Edit Other Income","DC196",scScreenFunctions);

	RetrieveContextData();
	SetMasks();

	if (m_sMetaAction == "Edit")
		DisableMainButton("Another");

	Validation_Init();	
	Initialise(true);

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC196");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				bSuccess = SaveUnearnedIncome();
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function DefaultFields()
// Inserts default values into all fields
{
	frmScreen.cboDescription.selectedIndex = 0;
	frmScreen.cboFrequency.selectedIndex = 0;
	frmScreen.txtAmount.value = 0;
}

function DisableOrEnableAddToDirectory()
{
	frmScreen.btnAddToDirectory.disabled = (m_bDirectoryAddress | AllFieldsEmpty())
}

function Initialise(bOnLoad)
// Initialises the screen
{
	if(bOnLoad == true)
		PopulateCombos();

	if(m_sMetaAction == "Edit")
		PopulateScreen();
	else
		DefaultFields();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateCombos()
// Populates all combos on the screen
{
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("UnEarnedIncomePaymentFreq","UnEarnedIncomeDescription");

	if(XML.GetComboLists(document,sGroupList))
	{
		// Populate Country combo
		blnSuccess = XML.PopulateCombo(document,frmScreen.cboFrequency,"UnEarnedIncomePaymentFreq",true);
		blnSuccess = blnSuccess & XML.PopulateCombo(document,frmScreen.cboDescription,"UnEarnedIncomeDescription",true);

		if(!blnSuccess) scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}

	XML = null;
}

function PopulateScreen()
// Populates the screen with details of the item selected in dc160
{
	UnearnedIncomeXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	UnearnedIncomeXML.LoadXML(m_sXML);

	if(UnearnedIncomeXML.SelectTag(null, "UNEARNEDINCOME") != null)
		with (frmScreen)
		{
			cboDescription.value = UnearnedIncomeXML.GetTagText("UNEARNEDINCOMETYPE");
			cboFrequency.value = UnearnedIncomeXML.GetTagText("PAYMENTFREQUENCY");
			txtAmount.value = UnearnedIncomeXML.GetTagText("UNEARNEDINCOMEAMOUNT");
			m_sUnearnedIncomeSequenceNumber = UnearnedIncomeXML.GetTagText("UNEARNEDINCOMESEQUENCENUMBER");
		}
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sXML = scScreenFunctions.GetContextParameter(window,"idXML","");
	m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","1325");
	m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","1");
}

function SaveUnearnedIncome()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	var TagRequestType = XML.CreateRequestTag(window, null);
	XML.SetAttribute("OPERATION","CriticalDataCheck");
	
	XML.CreateRequestTag(window,null)
	XML.CreateActiveTag("UNEARNEDINCOME");
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	XML.CreateTag("UNEARNEDINCOMEAMOUNT", frmScreen.txtAmount.value);
	XML.CreateTag("UNEARNEDINCOMETYPE", frmScreen.cboDescription.value);
	XML.CreateTag("PAYMENTFREQUENCY", frmScreen.cboFrequency.value);

	if (m_sMetaAction == "Edit")
		XML.CreateTag("UNEARNEDINCOMESEQUENCENUMBER", m_sUnearnedIncomeSequenceNumber);

	// Save the details
	XML.SelectTag(null,"REQUEST");
	XML.CreateActiveTag("CRITICALDATACONTEXT");
	XML.SetAttribute("APPLICATIONNUMBER",scScreenFunctions.GetContextParameter(window, "idApplicationNumber",null));
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER",scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber",null));
	XML.SetAttribute("SOURCEAPPLICATION","Omiga");
	XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
	XML.SetAttribute("ACTIVITYINSTANCE","1");
	XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
	XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
	XML.SetAttribute("COMPONENT","omCE.CustomerEmploymentBO");
		
	XML.SetAttribute("METHOD","SaveOtherIncome");	
	
	window.status = "Critical Data Check - please wait";
		
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"OmigaTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	window.status = ""
	 
	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}
-->
</script>
</body>
</html>


