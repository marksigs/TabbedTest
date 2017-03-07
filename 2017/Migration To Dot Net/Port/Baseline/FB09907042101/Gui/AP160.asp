<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP160.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Previous Landlords Reference
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		23/01/01	Created
JLD		02/02/01	Added some BO calls
CL		05/03/01	SYS1920 Read only functionality added
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
SA		02/05/01	SYS2275 Removed Complete Task button and added Confirm Button
					Existing click code moved from Complete Task button to Confirm button					
SA		03/12/01	SYS3285/3280 If Application set to "under Review" - don't make it read only. 
PSC		12/12/01	SYS3388 Prompt before running confirm process
LD		23/05/02	SYS4727 Use cached versions of frame functions

BMIDS History:

Prog	Date		Description
MO		21/11/2002	BMIDS01037 - Creating and Updating of references and casetasks result in 
					duplicate key violations
HMA     16/09/2003  BM0063  Amend HTML text for radio buttons					
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
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
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Forms Here */ %>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divAppInfo" style="TOP: 60px; LEFT: 10px; HEIGHT: 40px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Resident From Date
	<span style="TOP:-3px; LEFT:110px; POSITION:ABSOLUTE">
		<input id="txtCustStartDate" maxlength="10" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<span style="TOP:10px; LEFT:200px; POSITION:ABSOLUTE" class="msgLabel">
	Resident To Date
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtCustEndDate" maxlength="10" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<span style="TOP:10px; LEFT:390px; POSITION:ABSOLUTE" class="msgLabel">
	Monthly Rental
	<span style="TOP:0px; LEFT:100px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtCustRent" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
</div>
<div id="divRefInfo" style="TOP: 110px; LEFT: 10px; HEIGHT: 100px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Tenancy Start Date
	<span style="TOP:-3px; LEFT:110px; POSITION:ABSOLUTE">
		<input id="txtStartDate" maxlength="10" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<span style="TOP:10px; LEFT:200px; POSITION:ABSOLUTE" class="msgLabel">
	Date of Leaving
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtEndDate" maxlength="10" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<span style="TOP:10px; LEFT:390px; POSITION:ABSOLUTE" class="msgLabel">
	Monthly Rental
	<span style="TOP:0px; LEFT:100px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtRent" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
<span style="TOP:40px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Satisfactory Conduct?
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="optSatisfactoryConductYes" name="RadioGroup" type="radio" value="1"><label for="optSatisfactoryConductYes" class="msgLabel">Yes</label>
	</span>
	<span style="TOP:-3px; LEFT:180px; POSITION:ABSOLUTE">
		<input id="optSatisfactoryConductNo" name="RadioGroup" type="radio" value="0"><label for="optSatisfactoryConductNo" class="msgLabel">No</label>
	</span>
</span>
<% /* SYS2275 SA Complete Task button replaced with confirm */ %>
<% /* <span style="TOP:70px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnConfirm" value="Confirm" type="button" style="WIDTH:70px" class="msgButton">
</span> */ %>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 220px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP160attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sTaskXML = "";
var m_taskXML = null;
var m_prevLandlordXML = null;
var scScreenFunctions;
var m_bCreatePrevLandlordsRef = null;
var m_blnReadOnly = false;
//SYS2275 SA 1/5/01
var m_sAppFactFindNo = "";
<% /* MO - 21/11/2002 - BMIDS01037 */ %>
var m_bSetTaskToPending = null;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
	// SYS2275 SA Add Confirm Button
%>	var sButtonList = new Array("Submit","Cancel", "Confirm");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Previous Landlord's Reference","AP160",scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
	PopulateScreen();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP160");
	
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

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	//SYS2275 SA 2/5/01
	m_sAppFactFindNo = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
//DEBUG
//scScreenFunctions.SetContextParameter(window,"idCustomerNumber1","00073962");
//scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber1","1");
//m_sTaskXML = "<CASETASK SOURCEAPPLICATION=\"Omiga\" CASEID=\"C00071021\" ACTIVITYID=\"10\" TASKID=\"Second_Charge\" CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"5\" TASKSTATUS=\"20\" STAGEID=\"60\" CASESTAGESEQUENCENO=\"6\" TASKINSTANCE=\"1\"/>";
}
function GetCustomerVersion(sCustomerNumber)
{
	// find the customerversion number in context for this customernumber
	for(iCount = 1; iCount < 6; iCount++)
	{
		if(scScreenFunctions.GetContextParameter(window, "idCustomerNumber" + iCount.toString(), "") == sCustomerNumber)
			return(scScreenFunctions.GetContextParameter(window, "idCustomerVersionNumber" + iCount.toString(), ""));
	}
	alert("CustomerVersionNumber not found for customer " + sCustomerNumber);
}
function PopulateScreen()
{
	if(m_sTaskXML.length > 0)
	{
		m_taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_taskXML.LoadXML(m_sTaskXML);
		m_taskXML.SelectTag(null,"CASETASK");
		if(m_taskXML.GetAttribute("TASKSTATUS") == "40") m_sReadOnly = "1";
		//Populate applicant entered landlord info
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window , "GetTenancyDetailsForRef");
		XML.CreateActiveTag("TENANCYDETAIL");
		XML.SetAttribute("CUSTOMERNUMBER", m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
		XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
		XML.SetAttribute("CUSTOMERADDRESSSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
		XML.RunASP(document, "omAppProc.asp");
		if(XML.IsResponseOK())
		{
			XML.SelectTag(null, "TENANCYDETAILS");
			frmScreen.txtCustStartDate.value = XML.GetAttribute("DATEMOVEDIN");
			frmScreen.txtCustEndDate.value = XML.GetAttribute("DATEMOVEDOUT");
			frmScreen.txtCustRent.value = XML.GetAttribute("MONTHLYRENTAMOUNT");
		}
<%		//Populate the previous landlords reference info if some already stored. The task status 
		//may be 'Pending' but still have no reference details input yet if the task
		//has an OutputDocument associated with it. We need to cater for this.
%>		<% /* MO - 21/11/2002 - BMIDS01037 */ %>
		<% /* Rewritten if(m_taskXML.GetAttribute("TASKSTATUS") != "10")
		{
			m_prevLandlordXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			m_prevLandlordXML.CreateRequestTag(window , "GetPrevLandlordsRef");
			m_prevLandlordXML.CreateActiveTag("PREVIOUSLANDLORDSREF");
			m_prevLandlordXML.SetAttribute("CUSTOMERNUMBER", m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
			m_prevLandlordXML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
			m_prevLandlordXML.SetAttribute("CUSTOMERADDRESSSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
			m_prevLandlordXML.RunASP(document, "omAppProc.asp");
			var ErrorTypes = new Array("RECORDNOTFOUND");
			var ErrorReturn = m_prevLandlordXML.CheckResponse(ErrorTypes);
			if (ErrorReturn[1] == ErrorTypes[0])
			{
				if(m_taskXML.GetAttribute("OUTPUTDOCUMENT") == "")
				{
					alert("Previous Landlords reference cannot be found");
					m_bCreatePrevLandlordsRef = false;
				}
				else m_bCreatePrevLandlordsRef = true;
			}
			else if(ErrorReturn[0] == true)
			{
				m_bCreatePrevLandlordsRef = false;
				m_prevLandlordXML.SelectTag(null, "PREVIOUSLANDLORDSREF");
				frmScreen.txtStartDate.value = m_prevLandlordXML.GetAttribute("TENANCYSTARTDATE");
				frmScreen.txtEndDate.value = m_prevLandlordXML.GetAttribute("DATEOFLEAVING");
				frmScreen.txtRent.value = m_prevLandlordXML.GetAttribute("MONTHLYRENTAL");
				if(m_prevLandlordXML.GetAttribute("SATISFACTORYCONDUCT") == "1")
					frmScreen.optSatisfactoryConductYes.checked = true;
				else if(m_prevLandlordXML.GetAttribute("SATISFACTORYCONDUCT") == "0")
					frmScreen.optSatisfactoryConductNo.checked = true;
			}
		}
		else m_bCreatePrevLandlordsRef = true; */ %>
		m_prevLandlordXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_prevLandlordXML.CreateRequestTag(window , "GetPrevLandlordsRef");
		m_prevLandlordXML.CreateActiveTag("PREVIOUSLANDLORDSREF");
		m_prevLandlordXML.SetAttribute("CUSTOMERNUMBER", m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
		m_prevLandlordXML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
		m_prevLandlordXML.SetAttribute("CUSTOMERADDRESSSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
		m_prevLandlordXML.RunASP(document, "omAppProc.asp");
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = m_prevLandlordXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{
			m_bCreatePrevLandlordsRef = true;
		}
		else if(ErrorReturn[0] == true)
		{
			m_bCreatePrevLandlordsRef = false;
			m_prevLandlordXML.SelectTag(null, "PREVIOUSLANDLORDSREF");
			frmScreen.txtStartDate.value = m_prevLandlordXML.GetAttribute("TENANCYSTARTDATE");
			frmScreen.txtEndDate.value = m_prevLandlordXML.GetAttribute("DATEOFLEAVING");
			frmScreen.txtRent.value = m_prevLandlordXML.GetAttribute("MONTHLYRENTAL");
			if(m_prevLandlordXML.GetAttribute("SATISFACTORYCONDUCT") == "1")
				frmScreen.optSatisfactoryConductYes.checked = true;
			else if(m_prevLandlordXML.GetAttribute("SATISFACTORYCONDUCT") == "0")
				frmScreen.optSatisfactoryConductNo.checked = true;
		}
		if (m_taskXML.GetAttribute("TASKSTATUS") != "10") {
			m_bSetTaskToPending = false;
		} else {
			m_bSetTaskToPending = true;
		}
	}
	if(m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		<% /* SYS2275 SA Complete Task button replaced with confirm */ %>
		<% /* frmScreen.btnConfirm.disabled = true; */ %>
		DisableMainButton("Confirm");
		DisableMainButton("Submit");
	}
	else
	{
		scScreenFunctions.SetFieldState(frmScreen, "txtCustEndDate", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtCustRent", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtCustStartDate", "R");
	}
}
function UpdateLandlordRef()
{
	//Check the entered dates (scScreenFunctions comment suggests checking for the failure criteria)
	if(scScreenFunctions.CompareDateFields(frmScreen.txtEndDate,"<=",frmScreen.txtStartDate) == true)
	{
		alert("Date of leaving is before tenancy start date, please amend.");
		return false;
	}
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sASPFile = "";
	if(m_bCreatePrevLandlordsRef == true)
	{
		//Create new prev landlord record
		var reqTag = XML.CreateRequestTag(window , "CreatePrevLandlordsRef");
		XML.CreateActiveTag("CASETASK");
		XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		XML.SetAttribute("ACTIVITYID", m_taskXML.GetAttribute("ACTIVITYID"));
		XML.SetAttribute("CASEID", m_taskXML.GetAttribute("CASEID"));
		XML.SetAttribute("TASKID", m_taskXML.GetAttribute("TASKID"));
		XML.SetAttribute("TASKINSTANCE", m_taskXML.GetAttribute("TASKINSTANCE"));
		XML.SetAttribute("STAGEID", m_taskXML.GetAttribute("STAGEID"));
		XML.SetAttribute("CASESTAGESEQUENCENO", m_taskXML.GetAttribute("CASESTAGESEQUENCENO"));
		XML.ActiveTag = reqTag;
		XML.CreateActiveTag("PREVIOUSLANDLORDSREF");
		sASPFile = "OmigaTMBO.asp";
	}
	else
	{
		//Update existing prev landlord record
		XML.CreateRequestTag(window , "UpdatePrevLandlordsRef");
		m_prevLandlordXML.SelectTag(null, "RESPONSE");
		XML.AddXMLBlock(m_prevLandlordXML.CreateFragment());
		XML.SelectTag(null, "PREVIOUSLANDLORDSREF");
		sASPFile = "omAppProc.asp";
	}
	XML.SetAttribute("CUSTOMERNUMBER", m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
	XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
	XML.SetAttribute("CUSTOMERADDRESSSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
	XML.SetAttribute("TENANCYSTARTDATE", frmScreen.txtStartDate.value);
	XML.SetAttribute("DATEOFLEAVING", frmScreen.txtEndDate.value);
	XML.SetAttribute("MONTHLYRENTAL", frmScreen.txtRent.value);
	if(frmScreen.optSatisfactoryConductYes.checked == true)
		XML.SetAttribute("SATISFACTORYCONDUCT", "1");
	else if(frmScreen.optSatisfactoryConductNo.checked == true)
		XML.SetAttribute("SATISFACTORYCONDUCT", "0");
	else XML.SetAttribute("SATISFACTORYCONDUCT", "");
	// 	XML.RunASP(document, sASPFile);
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, sASPFile);
			break;
		default: // Error
			XML.SetErrorResponse();
		}
		
	<% /* MO - 21/11/2002 - BMIDS01037 - If everything worked ok and we have so far updated the lenders ref but
	still need to set the task to pending, do it now */ %>
	if(XML.IsResponseOK()) {
		if ((m_bCreatePrevLandlordsRef == false) && (m_bSetTaskToPending == true)) {
			
			SetToPendingXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			SetToPendingXML.CreateRequestTag(window , "UpdateCaseTask");
			SetToPendingXML.CreateActiveTag("CASETASK");
			SetToPendingXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
			SetToPendingXML.SetAttribute("ACTIVITYID", m_taskXML.GetAttribute("ACTIVITYID"));
			SetToPendingXML.SetAttribute("STAGEID", m_taskXML.GetAttribute("STAGEID"));
			SetToPendingXML.SetAttribute("CASESTAGESEQUENCENO", m_taskXML.GetAttribute("CASESTAGESEQUENCENO"));
			SetToPendingXML.SetAttribute("CASEID", m_taskXML.GetAttribute("CASEID"));
			SetToPendingXML.SetAttribute("TASKID", m_taskXML.GetAttribute("TASKID"));
			SetToPendingXML.SetAttribute("TASKINSTANCE", m_taskXML.GetAttribute("TASKINSTANCE"));
			SetToPendingXML.SetAttribute("TASKSTATUS", "20"); <% /* Pending */ %>
			
			SetToPendingXML.RunASP(document, "msgTMBO.asp") ;
			
			if (SetToPendingXML.IsResponseOK()) {
				return true;
			} else {
				return false;
			}
			
		} else {
			return true;
		}
	} else {
		return false;
	}
}
function ValidateLandlordRef()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = XML.CreateRequestTag(window , "ValidatePrevLandlordsRef");
	XML.CreateActiveTag("CASETASK");
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("ACTIVITYID", m_taskXML.GetAttribute("ACTIVITYID"));
	XML.SetAttribute("CASEID", m_taskXML.GetAttribute("CASEID"));
	XML.SetAttribute("TASKID", m_taskXML.GetAttribute("TASKID"));
	XML.SetAttribute("TASKINSTANCE", m_taskXML.GetAttribute("TASKINSTANCE"));
	XML.SetAttribute("STAGEID", m_taskXML.GetAttribute("STAGEID"));
	XML.SetAttribute("CASESTAGESEQUENCENO", m_taskXML.GetAttribute("CASESTAGESEQUENCENO"));
	// SYS2275 SA 1/5/01 Add missing attributes {
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
	XML.SetAttribute("CASETASKNAME", m_taskXML.GetAttribute("CASETASKNAME"));
	// SYS2275 }
	XML.ActiveTag = reqTag;
	//XML.CreateActiveTag("VALIDATELANDLORDREF");
	XML.CreateActiveTag("VALIDATEPREVLANDREF");
	XML.SetAttribute("CUSTOMERNUMBER",m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
	XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
	//SYS2275 SA 2/5/01 Changed attribute below from CustomerAddresSequenceNumber to CustomerAddressSequenceNumber
	XML.SetAttribute("CUSTOMERADDRESSSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
	// 	XML.RunASP(document, "OmigaTMBO.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "OmigaTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if(XML.IsResponseOK())
	{
// PUT IN WHEN METHODS AVAILABLE
//SYS2275 SA Methods now available...
		XML.SelectTag(null, "APPSTATUS");
		if(XML.GetAttribute("UNDERREVIEWIND") == "1")
		{
			alert("The Application has been placed under review");
			//SYS3285/3280 don't make it read only. 
			//scScreenFunctions.SetContextParameter(window,"idReadOnly","1");
			scScreenFunctions.SetContextParameter(window,"idAppUnderReview","1");
		}
		return true;
	}
	return false;
}
//SYS2275 Changed Complete Task button to Confirm Button
// function frmScreen.btnConfirm.onclick()
function btnConfirm.onclick()
{
	<% /* PSC 12/12/01 SYS3388 */ %>
	if (confirm("Please ensure all data has been entered correctly before continuing"))
	{	

		//run rules and set task status.
		if(UpdateLandlordRef())
		{
			if(ValidateLandlordRef())
			{
				scScreenFunctions.SetContextParameter(window,"idTaskXML","");
				frmToTM030.submit();
			}
		}
	}
}
function btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
	{
		if(UpdateLandlordRef())
		{
			scScreenFunctions.SetContextParameter(window,"idTaskXML","");
			frmToTM030.submit();
		}
	}
}
function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idTaskXML","");
	frmToTM030.submit();
}
-->
</script>
</body>
</html>


