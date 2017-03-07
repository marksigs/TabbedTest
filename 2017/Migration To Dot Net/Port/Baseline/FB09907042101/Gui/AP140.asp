<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP140.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Previous Employer's Reference
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		18/01/01	Created
JLD		22/01/01	Added some screen processing
JLD		23/01/01	Route to TM030.
JLD		25/01/01	change EMPLOYMENTSEQUENCENO  to EMPLOYMENTSEQUENCENUMBER in line with db
					Get the correct customernumber from context
CL		05/03/01	SYS1920 Read only functionality added
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
SA		02/05/01	SYS2275 Removed Complete Task button and added Confirm Button
					Existing click code moved from Complete Task button to Confirm button
SA		03/12/01	SYS3285/3280 If Application set to "under Review" - don't make it read only. 
PSC		12/12/01	SYS3388 Prompt before running confirm process
AT		23/04/02	SYS4359 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions

BMIDS History:

Prog	Date		Description
SA		17/10/02	BMIDS00608 Change PopulateScreen GetEmpDetailsForRef call needs a request element of <EMPDETAILSFORREF> not <EMPLOYMENTDETAILS>
MO		21/11/2002	BMIDS01037 - Creating and Updating of references and casetasks result in 
					duplicate key violations
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
<div id="divAppInfo" style="TOP: 60px; LEFT: 10px; HEIGHT: 220px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Job Title
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtCustJobTitle" maxlength="50" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
<span style="TOP:10px; LEFT:270px; POSITION:ABSOLUTE" class="msgLabel">
	<LABEL id="idPostcode"></LABEL>
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtPostcode" maxlength="8" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:40px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Employer
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtEmployer" maxlength="45" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
<span style="TOP:40px; LEFT:270px; POSITION:ABSOLUTE" class="msgLabel">
	<LABEL id="idBuildingName"></LABEL>
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtBuildingName" maxlength="40" style="WIDTH:100px" class="msgTxt">
	</span>
	<span style="TOP:-3px; LEFT:210px; POSITION:ABSOLUTE">
		<input id="txtBuildingNo" maxlength="10" style="WIDTH:50px" class="msgTxt">
	</span>
</span>
<span style="TOP:70px; LEFT:270px; POSITION:ABSOLUTE" class="msgLabel">
	Street
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtStreet" maxlength="40" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
<span style="TOP:100px; LEFT:270px; POSITION:ABSOLUTE" class="msgLabel">
	<LABEL id="idDistrict"></LABEL>
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtDistrict" maxlength="40" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
<span style="TOP:130px; LEFT:270px; POSITION:ABSOLUTE" class="msgLabel">
	<LABEL id="idTown"></LABEL>
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtTown" maxlength="40" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
<span style="TOP:160px; LEFT:270px; POSITION:ABSOLUTE" class="msgLabel">
	<LABEL id="idCounty"></LABEL>
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtCounty" maxlength="40" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
<span style="TOP:190px; LEFT:270px; POSITION:ABSOLUTE" class="msgLabel">
	Country
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtCountry" maxlength="50" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
</div>
<div id="divRefInfo" style="TOP: 285px; LEFT: 10px; HEIGHT: 135px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Job Title
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtJobTitle" maxlength="50" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
<span style="TOP:10px; LEFT:270px; POSITION:ABSOLUTE" class="msgLabel">
	Annual Gross Wage
	<span style="TOP:0px; LEFT:120px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtGrossWage" maxlength="10" style="TOP:-3px; WIDTH:200px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<span style="TOP:40px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Date Started
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtDateStarted" maxlength="10" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<span style="TOP:40px; LEFT:270px; POSITION:ABSOLUTE" class="msgLabel">
	Reason for Leaving
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<textarea id="txtReason" rows="5" style="WIDTH:200px" class="msgTxt"></textarea>
	</span>
</span>
<span style="TOP:70px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Date Left
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtDateLeft" maxlength="10" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<% /* SYS2275 SA Complete Task button replaced with confirm */ %>
<% /* <span style="TOP:100px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnConfirm" value="Complete Task" type="button" style="WIDTH:100px" class="msgButton">
</span> */ %>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 430px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP140attribs.asp" -->
<!-- #include FILE="Customise/AP140Customise.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sTaskXML = "";
var m_taskXML = null;
var m_PrevEmpXML = null;
var scScreenFunctions;
var m_bCreatePrevEmployersRef = null;
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
	FW030SetTitles("Previous Employer's Reference","AP140",scScreenFunctions);

	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
	// AT SYS4359 for client customisation
	Customise();
	SetMasks();
	Validation_Init();
	PopulateScreen();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP140");
	
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
//scScreenFunctions.SetContextParameter(window,"idCustomerNumber1","00068012");
//scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber1","1");
//m_sTaskXML = "<CASETASK SOURCEAPPLICATION=\"Omiga\" CASEID=\"C00071021\" ACTIVITYID=\"10\" TASKID=\"Second_Charge\" CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"1\" TASKSTATUS=\"20\" TASKINSTANCE=\"1\" STAGEID=\"60\" CASESTAGESEQUENCENO=\"6\"/>";

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
		m_taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject;
		m_taskXML.LoadXML(m_sTaskXML);
		m_taskXML.SelectTag(null, "CASETASK");
		if(m_taskXML.GetAttribute("TASKSTATUS") == "40") m_sReadOnly = "1";
		//Populate applicant entered prev emp details
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window , "GetEmpDetailsForRef");
		//BMIDS00608 Need to pass in EmpDetailsForRef tag name -- XML.CreateActiveTag("EMPLOYMENTDETAIL");
		XML.CreateActiveTag("EMPDETAILSFORREF");
		XML.SetAttribute("CUSTOMERNUMBER", m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
		XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
		XML.SetAttribute("EMPLOYMENTSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
		XML.SetAttribute("_COMBOLOOKUP_","y");
		XML.RunASP(document, "omAppProc.asp");  //ReferencesBO
		if(XML.IsResponseOK())
		{
			XML.SelectTag(null,"EMPLOYMENTDETAIL");
			frmScreen.txtCustJobTitle.value = XML.GetAttribute("JOBTITLE");
			frmScreen.txtEmployer.value = XML.GetAttribute("COMPANYNAME");
			frmScreen.txtPostcode.value = XML.GetAttribute("POSTCODE");
			frmScreen.txtBuildingName.value = XML.GetAttribute("BUILDINGORHOUSENAME");
			frmScreen.txtBuildingNo.value = XML.GetAttribute("BUILDINGORHOUSENUMBER");
			frmScreen.txtStreet.value = XML.GetAttribute("STREET");
			frmScreen.txtDistrict.value = XML.GetAttribute("DISTRICT");
			frmScreen.txtTown.value = XML.GetAttribute("TOWN");
			frmScreen.txtCounty.value = XML.GetAttribute("COUNTY");
			frmScreen.txtCountry.value = XML.GetAttribute("COUNTRY_TEXT");
		}
<%		//Populate the previous employers reference info if some already stored. The task status 
		//may be 'Pending' but still have no reference details input yet if the task
		//has an OutputDocument associated with it. We need to cater for this.
%>		<% /* MO - 21/11/2002 - BMIDS01037 */ %>
		<% /* if(m_taskXML.GetAttribute("TASKSTATUS") != "10") - Rewritten
		{
			m_PrevEmpXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			m_PrevEmpXML.CreateRequestTag(window , "GetPrevEmployersRef");
			m_PrevEmpXML.CreateActiveTag("PREVIOUSEMPLOYERSREF");
			m_PrevEmpXML.SetAttribute("CUSTOMERNUMBER", m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
			m_PrevEmpXML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
			m_PrevEmpXML.SetAttribute("EMPLOYMENTSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
			m_PrevEmpXML.RunASP(document, "omAppProc.asp");  //ReferencesBO
			var ErrorTypes = new Array("RECORDNOTFOUND");
			var ErrorReturn = m_PrevEmpXML.CheckResponse(ErrorTypes);
			if (ErrorReturn[1] == ErrorTypes[0])
			{
				if(m_taskXML.GetAttribute("OUTPUTDOCUMENT") == "")
				{
					alert("Previous employers reference cannot be found");
					m_bCreatePrevEmployersRef = false;
				}
				else m_bCreatePrevEmployersRef = true;
			}
			else if(ErrorReturn[0] == true)
			{
				m_bCreatePrevEmployersRef = false;
				m_PrevEmpXML.SelectTag(null, "PREVIOUSEMPLOYERSREF");
				frmScreen.txtJobTitle.value = m_PrevEmpXML.GetAttribute("JOBTITLE");
				frmScreen.txtGrossWage.value = m_PrevEmpXML.GetAttribute("ANNUALGROSSSALARY");
				frmScreen.txtDateStarted.value = m_PrevEmpXML.GetAttribute("DATESTARTED");
				frmScreen.txtDateLeft.value = m_PrevEmpXML.GetAttribute("DATELEFT");
				frmScreen.txtReason.value = m_PrevEmpXML.GetAttribute("REASONFORLEAVING");
			}
		}
		else m_bCreatePrevEmployersRef = true; */ %>
		m_PrevEmpXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_PrevEmpXML.CreateRequestTag(window , "GetPrevEmployersRef");
		m_PrevEmpXML.CreateActiveTag("PREVIOUSEMPLOYERSREF");
		m_PrevEmpXML.SetAttribute("CUSTOMERNUMBER", m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
		m_PrevEmpXML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
		m_PrevEmpXML.SetAttribute("EMPLOYMENTSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
		m_PrevEmpXML.RunASP(document, "omAppProc.asp");  //ReferencesBO
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = m_PrevEmpXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{
			m_bCreatePrevEmployersRef = true;
		}
		else if(ErrorReturn[0] == true)
		{
			m_bCreatePrevEmployersRef = false;
			m_PrevEmpXML.SelectTag(null, "PREVIOUSEMPLOYERSREF");
			frmScreen.txtJobTitle.value = m_PrevEmpXML.GetAttribute("JOBTITLE");
			frmScreen.txtGrossWage.value = m_PrevEmpXML.GetAttribute("ANNUALGROSSSALARY");
			frmScreen.txtDateStarted.value = m_PrevEmpXML.GetAttribute("DATESTARTED");
			frmScreen.txtDateLeft.value = m_PrevEmpXML.GetAttribute("DATELEFT");
			frmScreen.txtReason.value = m_PrevEmpXML.GetAttribute("REASONFORLEAVING");
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
		scScreenFunctions.SetFieldState(frmScreen, "txtCustJobTitle", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtEmployer", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtPostcode", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtBuildingName", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtBuildingNo", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtStreet", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtDistrict", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtTown", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtCounty", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtCountry", "R");
	}
}
function UpdatePrevEmpRef()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sASPFile = "";
	if(m_bCreatePrevEmployersRef == true)
	{
		//Create a prev emp ref
		var reqTag = XML.CreateRequestTag(window , "CreatePrevEmployersRef");
		XML.CreateActiveTag("CASETASK");
		XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		XML.SetAttribute("ACTIVITYID", m_taskXML.GetAttribute("ACTIVITYID"));
		XML.SetAttribute("STAGEID", m_taskXML.GetAttribute("STAGEID"));
		XML.SetAttribute("CASESTAGESEQUENCENO", m_taskXML.GetAttribute("CASESTAGESEQUENCENO"));
		XML.SetAttribute("CASEID", m_taskXML.GetAttribute("CASEID"));
		XML.SetAttribute("TASKID", m_taskXML.GetAttribute("TASKID"));
		XML.SetAttribute("TASKINSTANCE", m_taskXML.GetAttribute("TASKINSTANCE"));
		XML.ActiveTag = reqTag;
		XML.CreateActiveTag("PREVIOUSEMPLOYERSREF");
		sASPFile = "OmigaTMBO.asp";
	}
	else
	{
		//UPDATE the prev emp ref
		XML.CreateRequestTag(window , "UpdatePrevEmployersRef");
		m_PrevEmpXML.SelectTag(null, "RESPONSE");
		XML.AddXMLBlock(m_PrevEmpXML.CreateFragment());
		XML.SelectTag(null, "PREVIOUSEMPLOYERSREF");
		sASPFile = "omAppProc.asp";
	}
	XML.SetAttribute("CUSTOMERNUMBER",m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
	XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
	XML.SetAttribute("EMPLOYMENTSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
	XML.SetAttribute("JOBTITLE",frmScreen.txtJobTitle.value);
	XML.SetAttribute("ANNUALGROSSSALARY",frmScreen.txtGrossWage.value);
	XML.SetAttribute("DATESTARTED",frmScreen.txtDateStarted.value);
	XML.SetAttribute("DATELEFT",frmScreen.txtDateLeft.value);
	XML.SetAttribute("REASONFORLEAVING",frmScreen.txtReason.value);
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
		if ((m_bCreatePrevEmployersRef == false) && (m_bSetTaskToPending == true)) {
			
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
function ValidatePrevEmpRef()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = XML.CreateRequestTag(window , "ValidatePrevEmployersRef");
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
	XML.CreateActiveTag("VALPREVEMP");
	XML.SetAttribute("CUSTOMERNUMBER",m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
	XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
	XML.SetAttribute("EMPLOYMENTSEQUENCENUMBER",m_taskXML.GetAttribute("CONTEXT"));
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
		if(UpdatePrevEmpRef())
		{
			if(ValidatePrevEmpRef())
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
		if(UpdatePrevEmpRef())
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


