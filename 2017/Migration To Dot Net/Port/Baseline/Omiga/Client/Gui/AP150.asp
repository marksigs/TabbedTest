<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP150.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Landlords Reference
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		22/01/01	SYS1832 Created
JLD		23/01/01	Route to TM030, don't set radio buttons if not set in xml.
JLD		01/02/01	Added BO calls.
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
	Date Moved In
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtCustDate" maxlength="10" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<span style="TOP:10px; LEFT:280px; POSITION:ABSOLUTE" class="msgLabel">
	Monthly Rental
	<span style="TOP:0px; LEFT:100px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtCustMonthlyRental" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
</div>
<div id="divRefInfo" style="TOP: 110px; LEFT: 10px; HEIGHT: 140px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Start Date of Tenancy
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtDate" maxlength="10" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<span style="TOP:10px; LEFT:280px; POSITION:ABSOLUTE" class="msgLabel">
	Monthly Rental
	<span style="TOP:0px; LEFT:100px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtMonthlyRental" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<span style="TOP:40px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Current Arrears
	<span style="TOP:0px; LEFT:120px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtArrears" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
<span style="TOP:40px; LEFT:280px; POSITION:ABSOLUTE" class="msgLabel">
	Satisfactory Conduct?
	<span style="TOP:-3px; LEFT:130px; POSITION:ABSOLUTE">
		<input id="optSatisfactoryConductYes" name="RadioGroup1" type="radio" value="1"><label for="optSatisfactoryConductYes" class="msgLabel">Yes</label>
	</span>
	<span style="TOP:-3px; LEFT:190px; POSITION:ABSOLUTE">
		<input id="optSatisfactoryConductNo" name="RadioGroup1" type="radio" value="0"><label for="optSatisfactoryConductNo" class="msgLabel">No</label>
	</span>
</span>
<span style="TOP:70px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Has the rent been in<BR>arrears in the last 3 years?
	<span style="TOP:-3px; LEFT:150px; POSITION:ABSOLUTE">
		<input id="optArrearsYes" name="RadioGroup2" type="radio" value="1"><label for="optArrearsYes" class="msgLabel">Yes</label>
	</span>
	<span style="TOP:-3px; LEFT:210px; POSITION:ABSOLUTE">
		<input id="optArrearsNo" name="RadioGroup2" type="radio" value="0"><label for="optArrearsNo" class="msgLabel">No</label>
	</span>
</span>
<span style="TOP:70px; LEFT:280px; POSITION:ABSOLUTE" class="msgLabel">
	For how long (months)?
	<span style="TOP:-3px; LEFT:130px; POSITION:ABSOLUTE">
		<input id="txtMonths" maxlength="3" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<% /* SYS2275 SA Complete Task button replaced with confirm */ %>
<% /*<span style="TOP:110px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnConfirm" value="Confirm" type="button" style="WIDTH:100px" class="msgButton">
</span> */ %>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 260px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP150attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sTaskXML = "";
var m_taskXML = null;
var m_currLandlordXML = null;
var scScreenFunctions;
var m_bCreateLandlordsRef = null;
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
	FW030SetTitles("Landlord's Reference","AP150",scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
	SetMasks();
	Validation_Init();
	PopulateScreen();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP150");
	
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
	//SYS2275 SA 1/5/01
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
		m_taskXML.SelectTag(null, "CASETASK");
		if(m_taskXML.GetAttribute("TASKSTATUS") == "40") m_sReadOnly = "1";
		//Populate Applicant entered info
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
			frmScreen.txtCustDate.value = XML.GetAttribute("DATEMOVEDIN");
			frmScreen.txtCustMonthlyRental.value = XML.GetAttribute("MONTHLYRENTAMOUNT");
		}
<%		//Populate the landlords reference info if some already stored. The task status 
		//may be 'Pending' but still have no reference details input yet if the task
		//has an OutputDocument associated with it. We need to cater for this.
%>		<% /* MO - 21/11/2002 - BMIDS01037 */ %>
		<% /* Rewritten if(m_taskXML.GetAttribute("TASKSTATUS") != "10")
		{
			m_currLandlordXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			m_currLandlordXML.CreateRequestTag(window , "GetCurrLandlordsRef");
			m_currLandlordXML.CreateActiveTag("CURRENTLANDLORDSREF");
			m_currLandlordXML.SetAttribute("CUSTOMERNUMBER", m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
			m_currLandlordXML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
			m_currLandlordXML.SetAttribute("CUSTOMERADDRESSSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
			m_currLandlordXML.RunASP(document, "omAppProc.asp")
			var ErrorTypes = new Array("RECORDNOTFOUND");
			var ErrorReturn = m_currLandlordXML.CheckResponse(ErrorTypes);
			if (ErrorReturn[1] == ErrorTypes[0])
			{
				if(m_taskXML.GetAttribute("OUTPUTDOCUMENT") == "")
				{
					alert("Landlords reference cannot be found");
					m_bCreateLandlordsRef = false;
				}
				else m_bCreateLandlordsRef = true;
			}
			else if(ErrorReturn[0] == true)
			{
				m_bCreateLandlordsRef = false;
				m_currLandlordXML.SelectTag(null,"CURRENTLANDLORDSREF");
				frmScreen.txtDate.value = m_currLandlordXML.GetAttribute("TENANCYSTARTDATE");
				frmScreen.txtMonthlyRental.value = m_currLandlordXML.GetAttribute("MONTHLYRENTAL");
				frmScreen.txtArrears.value = m_currLandlordXML.GetAttribute("CURRENTARREARS");
				if(m_currLandlordXML.GetAttribute("SATISFACTORYCONDUCT") == "1")
					frmScreen.optSatisfactoryConductYes.checked = true;
				else if(m_currLandlordXML.GetAttribute("SATISFACTORYCONDUCT") == "0")
					frmScreen.optSatisfactoryConductNo.checked = true;
				if(m_currLandlordXML.GetAttribute("RENTINARREARS") == "1")
					frmScreen.optArrearsYes.checked = true;
				else if(m_currLandlordXML.GetAttribute("RENTINARREARS") == "0")
				{
					frmScreen.optArrearsNo.checked = true;
					scScreenFunctions.SetFieldState(frmScreen, "txtMonths", "R");
				}
				frmScreen.txtMonths.value = m_currLandlordXML.GetAttribute("MONTHSRENTINARREARS");
			}
		}
		else m_bCreateLandlordsRef = true; */ %>
		m_currLandlordXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_currLandlordXML.CreateRequestTag(window , "GetCurrLandlordsRef");
		m_currLandlordXML.CreateActiveTag("CURRENTLANDLORDSREF");
		m_currLandlordXML.SetAttribute("CUSTOMERNUMBER", m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
		m_currLandlordXML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
		m_currLandlordXML.SetAttribute("CUSTOMERADDRESSSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
		m_currLandlordXML.RunASP(document, "omAppProc.asp")
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = m_currLandlordXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{
			m_bCreateLandlordsRef = true;
		}
		else if(ErrorReturn[0] == true)
		{
			m_bCreateLandlordsRef = false;
			m_currLandlordXML.SelectTag(null,"CURRENTLANDLORDSREF");
			frmScreen.txtDate.value = m_currLandlordXML.GetAttribute("TENANCYSTARTDATE");
			frmScreen.txtMonthlyRental.value = m_currLandlordXML.GetAttribute("MONTHLYRENTAL");
			frmScreen.txtArrears.value = m_currLandlordXML.GetAttribute("CURRENTARREARS");
			if(m_currLandlordXML.GetAttribute("SATISFACTORYCONDUCT") == "1")
				frmScreen.optSatisfactoryConductYes.checked = true;
			else if(m_currLandlordXML.GetAttribute("SATISFACTORYCONDUCT") == "0")
				frmScreen.optSatisfactoryConductNo.checked = true;
			if(m_currLandlordXML.GetAttribute("RENTINARREARS") == "1")
				frmScreen.optArrearsYes.checked = true;
			else if(m_currLandlordXML.GetAttribute("RENTINARREARS") == "0")
			{
				frmScreen.optArrearsNo.checked = true;
				scScreenFunctions.SetFieldState(frmScreen, "txtMonths", "R");
			}
			frmScreen.txtMonths.value = m_currLandlordXML.GetAttribute("MONTHSRENTINARREARS");
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
		scScreenFunctions.SetFieldState(frmScreen, "txtCustDate", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtCustMonthlyRental", "R");
	}
}
function frmScreen.optArrearsYes.onclick()
{
	scScreenFunctions.SetFieldState(frmScreen, "txtMonths", "W");
}
function frmScreen.optArrearsNo.onclick()
{
	frmScreen.txtMonths.value = "";
	scScreenFunctions.SetFieldState(frmScreen, "txtMonths", "R");
}
function UpdateLandlordRef()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sASPFile = "";
	if(m_bCreateLandlordsRef == true)
	{
		//Create a landlord ref
		var reqTag = XML.CreateRequestTag(window , "CreateCurrLandlordsRef");
		XML.CreateActiveTag("CASETASK");
		XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		XML.SetAttribute("ACTIVITYID", m_taskXML.GetAttribute("ACTIVITYID"));
		XML.SetAttribute("CASEID", m_taskXML.GetAttribute("CASEID"));
		XML.SetAttribute("TASKID", m_taskXML.GetAttribute("TASKID"));
		XML.SetAttribute("TASKINSTANCE", m_taskXML.GetAttribute("TASKINSTANCE"));
		XML.SetAttribute("STAGEID", m_taskXML.GetAttribute("STAGEID"));
		XML.SetAttribute("CASESTAGESEQUENCENO", m_taskXML.GetAttribute("CASESTAGESEQUENCENO"));
		XML.ActiveTag = reqTag;
		XML.CreateActiveTag("CURRENTLANDLORDSREF");
		sASPFile = "OmigaTMBO.asp";
	}
	else
	{
		//UPDATE the landlord ref
		XML.CreateRequestTag(window , "UpdateCurrLandlordsRef");
		m_currLandlordXML.SelectTag(null, "RESPONSE");
		XML.AddXMLBlock(m_currLandlordXML.CreateFragment());
		XML.SelectTag(null, "CURRENTLANDLORDSREF");
		sASPFile = "omAppProc.asp";
	}
	XML.SetAttribute("CUSTOMERNUMBER", m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
	XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
	XML.SetAttribute("CUSTOMERADDRESSSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
	XML.SetAttribute("TENANCYSTARTDATE",frmScreen.txtDate.value);
	XML.SetAttribute("MONTHLYRENTAL",frmScreen.txtMonthlyRental.value);
	XML.SetAttribute("CURRENTARREARS",frmScreen.txtArrears.value);
	if(frmScreen.optSatisfactoryConductYes.checked == true)
		XML.SetAttribute("SATISFACTORYCONDUCT","1");
	else if(frmScreen.optSatisfactoryConductNo.checked == true)
		XML.SetAttribute("SATISFACTORYCONDUCT","0");
	else XML.SetAttribute("SATISFACTORYCONDUCT","");
	if(frmScreen.optArrearsYes.checked == true)
		XML.SetAttribute("RENTINARREARS","1");
	else if(frmScreen.optArrearsNo.checked == true)
		XML.SetAttribute("RENTINARREARS","0");
	else XML.SetAttribute("RENTINARREARS","");
	XML.SetAttribute("MONTHSRENTINARREARS",frmScreen.txtMonths.value);
	// 	XML.RunASP(document, sASPFile)
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, sASPFile)
			break;
		default: // Error
			XML.SetErrorResponse();
		}
	
	<% /* MO - 21/11/2002 - BMIDS01037 - If everything worked ok and we have so far updated the lenders ref but
	still need to set the task to pending, do it now */ %>
	if(XML.IsResponseOK()) {
		if ((m_bCreateLandlordsRef == false) && (m_bSetTaskToPending == true)) {
			
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
	var reqTag = XML.CreateRequestTag(window , "ValidateCurrLandlordsRef");
	XML.CreateActiveTag("CASETASK");
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("ACTIVITYID", m_taskXML.GetAttribute("ACTIVITYID"));
	XML.SetAttribute("CASEID", m_taskXML.GetAttribute("CASEID"));
	XML.SetAttribute("TASKID", m_taskXML.GetAttribute("TASKID"));
	XML.SetAttribute("STAGEID", m_taskXML.GetAttribute("STAGEID"));
	XML.SetAttribute("CASESTAGESEQUENCENO", m_taskXML.GetAttribute("CASESTAGESEQUENCENO"));
	XML.SetAttribute("TASKINSTANCE", m_taskXML.GetAttribute("TASKINSTANCE"));
	// SYS2275 SA 1/5/01 Add missing attributes {
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
	XML.SetAttribute("CASETASKNAME", m_taskXML.GetAttribute("CASETASKNAME"));
	// SYS2275 }
	XML.ActiveTag = reqTag;
	XML.CreateActiveTag("VALIDATELANDLORDREF");
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
			//SYS3285/3280 Don't make it read only. 
			//scScreenFunctions.SetContextParameter(window,"idReadOnly","1");
			scScreenFunctions.SetContextParameter(window,"idAppUnderReview","1");
		}
		return true;
	}
	return false;
}
//SYS2275 Changed Complete Task button to Confirm Button
//function frmScreen.btnConfirm.onclick()
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


