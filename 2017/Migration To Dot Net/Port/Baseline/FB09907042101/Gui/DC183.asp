<%@ LANGUAGE="JSCRIPT" %>
<html>
<comment>
Workfile:      DC183.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Contract Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		16/02/2000	Created
AY		30/03/00	New top menu/scScreenFunctions change
IVW		11/04/2000	Changed Prev/Next/Cancel to Ok/Cancel
BG		17/05/00	SYS0752 Removed Tooltips
MC		23/05/00	SYS0756 If Read-only mode, disabled fields
MH      14/06/00    SYS0791 DC191 Popup width
MC		21/06/00	SYS0956 Accountant not mandatory. Enable popup
					when in read only mode
BG		22/08/00	SYS0790 Changed title, default yes on radiobutton, validation.
BG		08/09/00	SYS0790	Further validation - added checklength function.
CL		05/03/01	SYS1920 Read only functionality added
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
</comment>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom Specific changes
Prog    Date           Description
MAH		20/11/2006		E2CR35 Add numberOfRenewals
*/

%>
<HEAD>
	<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
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
<form id="frmToDC160" method="post" action="DC160.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC170" method="post" action="DC170.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC192" method="post" action="DC192.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 156px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Employer's Name
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtEmployerName" maxlength="50" style="POSITION: absolute; WIDTH: 200px" class="msgReadOnly" readonly tabindex=-1>
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Length of Contract

		<span style="LEFT: 160px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Years
			<span style="LEFT: 40px; POSITION: absolute; TOP: -3px">
				<input id="txtLengthOfContractYears" maxlength="2" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
			</span> 
		</span>

		<span style="LEFT: 260px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Months
			<span style="LEFT: 40px; POSITION: absolute; TOP: -3px">
				<input id="txtLengthOfContractMonths" maxlength="2" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
			</span> 
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 58px" class="msgLabel">
		Length of Time Employed

		<span style="LEFT: 160px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Years
			<span style="LEFT: 40px; POSITION: absolute; TOP: -3px">
				<input id="txtNumberOfYearsEmployed" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
			</span> 
		</span>

		<span style="LEFT: 260px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Months
			<span style="LEFT: 40px; POSITION: absolute; TOP: -3px">
				<input id="txtNumberOfMonthsEmployed" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
			</span> 
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 82px" class="msgLabel">How many times has the contract been renewed?
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="txtNumberOfRenewals" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgTxt" NAME="txtNumberOfRenewals">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 106px" class="msgLabel">
		Contract likely to be renewed?
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="optLikelyToBeRenewedYes" name="LikelyToBeRenewedGroup" type="radio" value="1">
			<label for="optLikelyToBeRenewedYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px">
			<input id="optLikelyToBeRenewedNo" name="LikelyToBeRenewedGroup" type="radio" value="0">
			<label for="optLikelyToBeRenewedNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 130px" class="msgLabel">
		Accountant
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountantName" name="Accountant Name" maxlength="30" style="POSITION: absolute; WIDTH: 200px" class="msgReadOnly" readonly tabindex=-1>
		</span>
		<span style="LEFT: 365px; POSITION: absolute; TOP: -9px">				
			<input id="btnAccountant" type="button" style="HEIGHT: 32px; WIDTH: 32px" class ="msgDDButton">				
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
<!-- #include FILE="attribs/DC183attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sEmploymentSequenceNumber = "";
var m_sEmployerName = "";
var ContractDetailsXML = null;
var m_sAccountantGUID = "";
var m_bCanAddToDirectory = false;
var scScreenFunctions;
var m_blnReadOnly = false;

/* EVENTS */

function btnCancel.onclick()
{
	//clear the contexts
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Edit");
	frmToDC170.submit();
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		//clear the contexts
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		frmToDC192.submit();
	}
}

function frmScreen.btnAccountant.onclick()
{
	<%/* Interface to Accountant Details popup */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ArrayArguments = XML.CreateRequestAttributeArray(window);
	ArrayArguments[4] = m_sAccountantGUID;    
	ArrayArguments[5] = m_bCanAddToDirectory;		
	ArrayArguments[6] = m_sReadOnly;
	var sReturn = scScreenFunctions.DisplayPopup(window, document, "DC191.asp", ArrayArguments, 628, 480);

	if(sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sAccountantGUID = sReturn[1];
		frmScreen.txtAccountantName.value = sReturn[2];
	}
	XML = null;
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
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Contract Details","DC183",scScreenFunctions);

	RetrieveContextData();
	SetMasks();

	Validation_Init();	
	Initialise(true);

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC183");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

function SetScreenToReadOnly()
{
	scScreenFunctions.SetScreenToReadOnly(frmScreen);
	// frmScreen.btnAccountant.disabled = true;
}

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
			if (IsChanged())
				if (!checklength())
				{				
					bSuccess = false;
				}
				else
				<% /* SYS0956 - Accountant not mandatory 
				if (m_sAccountantGUID == "")
				{
					alert("An accountant must be specified.");
					bSuccess = false;
				}
				else	*/ %>
					bSuccess = SaveContractDetails();
	return(bSuccess);
}

function checklength()
{
	if(parseInt(frmScreen.txtLengthOfContractYears.value) > "50") 
	{	
		alert("Number of years must be less than 51");
		frmScreen.txtLengthOfContractYears.focus();
		return false;
	}	
	
	else if(parseInt(frmScreen.txtNumberOfYearsEmployed.value) > "50")
	{	
		alert("Number of years must be less than 51");
		frmScreen.txtNumberOfYearsEmployed.focus();
		return false;
	}
	
	else if(parseInt(frmScreen.txtLengthOfContractMonths.value) > "11") 
	{
		alert("Number of months must be less than 12");	
		frmScreen.txtLengthOfContractMonths.focus();
		return false;
	}
	else if(parseInt(frmScreen.txtNumberOfMonthsEmployed.value) > "11")
	{
		alert("Number of months must be less than 12");	
		frmScreen.txtLengthOfContractMonths.focus();
		return false;
	}
	else
		return true;
}

function DefaultFields()
// Inserts default values into all fields
{
	with (frmScreen)
	{
		txtLengthOfContractYears.value = "";
		txtLengthOfContractMonths.value = "";
		txtNumberOfYearsEmployed.value = "";
		txtNumberOfMonthsEmployed.value = "";
		txtNumberOfRenewals.value = ""; <%/*MAH 20/11/2006 E2CR35 */%>
		scScreenFunctions.SetRadioGroupValue(frmScreen,"LikelyToBeRenewedGroup","1");
	}
}

function Initialise(bOnLoad)
// Initialises the screen
{	
	frmScreen.txtEmployerName.value = m_sEmployerName;

	if (!PopulateScreen())
		DefaultFields();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateScreen()
// Populates the screen with details of the item selected in dc160
{
	ContractDetailsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	ContractDetailsXML.CreateRequestTag(window,null);
	ContractDetailsXML.CreateActiveTag("CONTRACTDETAILS");
	ContractDetailsXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	ContractDetailsXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	ContractDetailsXML.CreateTag("EMPLOYMENTSEQUENCENUMBER", m_sEmploymentSequenceNumber);
	ContractDetailsXML.RunASP(document,"GetContractDetails.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = ContractDetailsXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[1] == ErrorTypes[0]))
	{
		//Error: record not found
		m_sMetaAction = "Add";
	}
	else if (ErrorReturn[0] == true)
	{
		// No error
		m_sMetaAction = "Edit";

		if(ContractDetailsXML.SelectTag(null, "CONTRACTDETAILS") != null)
			with (frmScreen)
			{
				txtLengthOfContractYears.value = ContractDetailsXML.GetTagText("LENGTHOFCONTRACTYEARS");
				txtLengthOfContractMonths.value = ContractDetailsXML.GetTagText("LENGTHOFCONTRACTMONTHS");
				txtNumberOfYearsEmployed.value = ContractDetailsXML.GetTagText("NUMBEROFYEARSEMPLOYED");
				txtNumberOfMonthsEmployed.value = ContractDetailsXML.GetTagText("NUMBEROFMONTHSEMPLOYED");
				scScreenFunctions.SetRadioGroupValue(frmScreen,"LikelyToBeRenewedGroup",ContractDetailsXML.GetTagText("LIKELYTOBERENEWED"));
				txtNumberOfRenewals.value = ContractDetailsXML.GetTagText("NUMBEROFRENEWALS"); <%/*MAH 20/11/2006 E2CR35 */%>
				if (ContractDetailsXML.SelectTag(null, "ACCOUNTANT") != null)
				{
					m_sAccountantGUID = ContractDetailsXML.GetTagText("ACCOUNTANTGUID");
					frmScreen.txtAccountantName.value = ContractDetailsXML.GetTagText("COMPANYNAME");
				}
			}
	}

	ErrorTypes = null;
	ErrorReturn = null;

	return(m_sMetaAction == "Edit");
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","1325");
	m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","1");
	m_sEmploymentSequenceNumber = scScreenFunctions.GetContextParameter(window,"idEmploymentSequenceNumber","1");
	m_sEmployerName = scScreenFunctions.GetContextParameter(window,"idEmployerName","");

	var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bCanAddToDirectory = (sUserRole >= XML.GetGlobalParameterAmount(document,"AddToDirectoryRole"));

	XML = null;
}

function SaveContractDetails()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,null)

	XML.CreateActiveTag("CONTRACTDETAILS");
	XML.CreateTag("ACCOUNTANTGUID", m_sAccountantGUID);
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	XML.CreateTag("EMPLOYMENTSEQUENCENUMBER",m_sEmploymentSequenceNumber);
	XML.CreateTag("LENGTHOFCONTRACTYEARS",frmScreen.txtLengthOfContractYears.value);
	XML.CreateTag("LENGTHOFCONTRACTMONTHS",frmScreen.txtLengthOfContractMonths.value);
	XML.CreateTag("NUMBEROFYEARSEMPLOYED",frmScreen.txtNumberOfYearsEmployed.value);
	XML.CreateTag("NUMBEROFMONTHSEMPLOYED",frmScreen.txtNumberOfMonthsEmployed.value);
	XML.CreateTag("NUMBEROFRENEWALS",frmScreen.txtNumberOfRenewals.value); <%/*MAH 20/11/2006 E2CR35 */%>
	XML.CreateTag("LIKELYTOBERENEWED",scScreenFunctions.GetRadioGroupValue(frmScreen,"LikelyToBeRenewedGroup"));

	<%/* Save the details */%>
	// 	XML.RunASP(document,"SaveContractDetails.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"SaveContractDetails.asp");
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


