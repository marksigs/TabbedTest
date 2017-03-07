<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP130.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Previous Lender's reference
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		17/01/01	Created
JLD		22/01/01	Added some screen processing
JLD		23/01/01	Don't set radios if value not set, Route to TM030.
JLD		31/01/01	Added some BO methods and change some screen text
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
<div id="divAppInfo" style="TOP: 60px; LEFT: 10px; HEIGHT: 75px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Lender's Name
	<span style="TOP:-3px; LEFT:135px; POSITION:ABSOLUTE">
		<input id="txtLendersName" maxlength="45" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
<span style="TOP:10px; LEFT:305px; POSITION:ABSOLUTE" class="msgLabel">
	Account Number
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtCustLoanNumber" maxlength="30" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
<span style="TOP:40px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Total Monthly Repayment
	<span style="TOP:0px; LEFT:135px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency">£</label>
		<input id="txtCustRepayment" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
</div>
<div id="divLenderInfo" style="TOP: 140px; LEFT: 10px; HEIGHT: 200px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Monthly Repayment
	<span style="TOP:0px; LEFT:140px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtRepayment" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<span style="TOP:10px; LEFT:290px; POSITION:ABSOLUTE" class="msgLabel">
	Balance of House Purchase
	<span style="TOP:0px; LEFT:170px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtHousePurchase" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<span style="TOP:40px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Redemption Date
	<span style="TOP:-3px; LEFT:140px; POSITION:ABSOLUTE">
		<input id="txtRedemptionDate" maxlength="10" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<span style="TOP:40px; LEFT:290px; POSITION:ABSOLUTE" class="msgLabel">
	Balance of Home improvements
	<span style="TOP:0px; LEFT:170px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtHomeImpr" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<span style="TOP:70px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Date of Loan
	<span style="TOP:-3px; LEFT:140px; POSITION:ABSOLUTE">
		<input id="txtLoanDate" maxlength="10" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<span style="TOP:70px; LEFT:290px; POSITION:ABSOLUTE" class="msgLabel">
	Balance of Capital Raising
	<span style="TOP:0px; LEFT:170px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtCapitalRaising" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<span style="TOP:100px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Loan Number
	<span style="TOP:-3px; LEFT:140px; POSITION:ABSOLUTE">
		<input id="txtLoanNumber" maxlength="30" style="WIDTH:70px" class="msgTxt">
	</span>
</span>
<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
<span style="TOP:100px; LEFT:290px; POSITION:ABSOLUTE" class="msgLabel">
	Good Account Conduct?
	<span style="TOP:-3px; LEFT:170px; POSITION:ABSOLUTE">
		<input id="optAccountConductYes" name="RadioGroup1" type="radio" value="1"><label for="optAccountConductYes" class="msgLabel">Yes</label>
	</span>
	<span style="TOP:-3px; LEFT:230px; POSITION:ABSOLUTE">
		<input id="optAccountConductNo" name="RadioGroup1" type="radio" value="0"><label for="optAccountConductNo" class="msgLabel">No</label>
	</span>
</span>
<span style="TOP:130px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Original House Purchase
	<span style="TOP:0px; LEFT:140px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency"></label>
		<input id="txtOrigHousePurchase" maxlength="10" style="TOP:-3px; WIDTH:100px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<span style="TOP:130px; LEFT:290px; POSITION:ABSOLUTE" class="msgLabel">
	Previous Arrears?
	<span style="TOP:-3px; LEFT:170px; POSITION:ABSOLUTE">
		<input id="optPreviousArrearsYes" name="RadioGroup2" type="radio" value="1"><label for="optPreviousArrearsYes" class="msgLabel">Yes</label>
	</span>
	<span style="TOP:-3px; LEFT:230px; POSITION:ABSOLUTE">
		<input id="optPreviousArrearsNo" name="RadioGroup2" type="radio" value="0"><label for="optPreviousArrearsNo" class="msgLabel">No</label>
	</span>
</span>
<% /* SYS2275 SA Complete Task button replaced with confirm */ %>
<% /* <span style="TOP:160px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnConfirm" value="Complete Task" type="button" style="WIDTH:100px" class="msgButton">
</span> */ %>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 350px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP130attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sTaskXML = "";
var m_taskXML = null;
var m_lendersRefXML = null;
var scScreenFunctions;
var m_bCreatePrevLendersRef = null;
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
	FW030SetTitles("Previous Lender's Reference","AP130",scScreenFunctions);

	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
	SetMasks();
	Validation_Init();
	PopulateScreen();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP130");
	
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
//m_sTaskXML = "<CASETASK SOURCEAPPLICATION=\"Omiga\" CASEID=\"C00071021\" ACTIVITYID=\"10\" TASKID=\"Second_Charge\" CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"A7768CA8984D11D4B62200105ABB1680\" TASKSTATUS=\"20\" STAGEID=\"60\" CASESTAGESEQUENCENO=\"6\" TASKINSTANCE=\"1\"/>";	
}
function PopulateScreen()
{
	if(m_sTaskXML.length > 0)
	{
		m_taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_taskXML.LoadXML(m_sTaskXML);
		m_taskXML.SelectTag(null,"CASETASK");
		if(m_taskXML.GetAttribute("TASKSTATUS") == "40")m_sReadOnly = "1";
		//Populate applicant info
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window , "GetLoanDetailsForRef");
		XML.CreateActiveTag("LOANDETAIL");
		XML.SetAttribute("ACCOUNTGUID", m_taskXML.GetAttribute("CONTEXT"));
		XML.RunASP(document, "omAppProc.asp");
		if(XML.IsResponseOK())
		{
			XML.SelectTag(null, "LOANDETAIL");
			frmScreen.txtCustLoanNumber.value = XML.GetAttribute("ACCOUNTNUMBER");
			frmScreen.txtCustRepayment.value = XML.GetAttribute("TOTALMONTHLYREPAYMENT");
			frmScreen.txtLendersName.value = XML.GetAttribute("COMPANYNAME");
		}
<%		//Populate the previous lenders reference info if some already stored. The task status 
		//may be 'Pending' but still have no reference details input yet if the task
		//has an OutputDocument associated with it. We need to cater for this.
%>		<% /* MO - 21/11/2002 - BMIDS01037 - Rewritten */ %>
		<% /* if(m_taskXML.GetAttribute("TASKSTATUS") != "10")
		{
			m_lendersRefXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			m_lendersRefXML.CreateRequestTag(window , "GetPrevLendersRef");
			m_lendersRefXML.CreateActiveTag("PREVIOUSLENDERSREF");
			m_lendersRefXML.SetAttribute("ACCOUNTGUID", m_taskXML.GetAttribute("CONTEXT"));
			m_lendersRefXML.RunASP(document, "omAppProc.asp");
			var ErrorTypes = new Array("RECORDNOTFOUND");
			var ErrorReturn = m_lendersRefXML.CheckResponse(ErrorTypes);
			if (ErrorReturn[1] == ErrorTypes[0])
			{
				if(m_taskXML.GetAttribute("OUTPUTDOCUMENT") == "")
				{
					alert("Lenders reference cannot be found");
					m_bCreatePrevLendersRef = false;
				}
				else m_bCreatePrevLendersRef = true;
			}
			else if(ErrorReturn[0] == true)
			{
				m_bCreatePrevLendersRef = false;
				m_lendersRefXML.SelectTag(null, "PREVIOUSLENDERSREF");
				frmScreen.txtRepayment.value = m_lendersRefXML.GetAttribute("MONTHLYREPAYMENT");
				frmScreen.txtHousePurchase.value = m_lendersRefXML.GetAttribute("BALHOUSEPURCHASE");
				frmScreen.txtRedemptionDate.value = m_lendersRefXML.GetAttribute("REDEMPTIONDATE");
				frmScreen.txtHomeImpr.value = m_lendersRefXML.GetAttribute("BALHOUSEIMPROVEMENT");
				frmScreen.txtLoanDate.value = m_lendersRefXML.GetAttribute("DATEOFLOAN");
				frmScreen.txtCapitalRaising.value = m_lendersRefXML.GetAttribute("BALCAPITALRAISING");
				frmScreen.txtLoanNumber.value = m_lendersRefXML.GetAttribute("LOANNUMBER");
				frmScreen.txtOrigHousePurchase.value = m_lendersRefXML.GetAttribute("ORIGINALHOUSEPURCHASE");
				if(m_lendersRefXML.GetAttribute("GOODACCOUNTCONDUCT") == "1")
					frmScreen.optAccountConductYes.checked = true;
				else if(m_lendersRefXML.GetAttribute("GOODACCOUNTCONDUCT") == "0")
					frmScreen.optAccountConductNo.checked = true;
				if(m_lendersRefXML.GetAttribute("PREVIOUSARREARS") == "1")
					frmScreen.optPreviousArrearsYes.checked = true;
				else if(m_lendersRefXML.GetAttribute("PREVIOUSARREARS") == "0")
					frmScreen.optPreviousArrearsNo.checked = true;
			}
		}
		else m_bCreatePrevLendersRef = true; */ %>
		m_lendersRefXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_lendersRefXML.CreateRequestTag(window , "GetPrevLendersRef");
		m_lendersRefXML.CreateActiveTag("PREVIOUSLENDERSREF");
		m_lendersRefXML.SetAttribute("ACCOUNTGUID", m_taskXML.GetAttribute("CONTEXT"));
		m_lendersRefXML.RunASP(document, "omAppProc.asp");
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = m_lendersRefXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{
			m_bCreatePrevLendersRef = true;
		}
		else if(ErrorReturn[0] == true)
		{
			m_bCreatePrevLendersRef = false;
			m_lendersRefXML.SelectTag(null, "PREVIOUSLENDERSREF");
			frmScreen.txtRepayment.value = m_lendersRefXML.GetAttribute("MONTHLYREPAYMENT");
			frmScreen.txtHousePurchase.value = m_lendersRefXML.GetAttribute("BALHOUSEPURCHASE");
			frmScreen.txtRedemptionDate.value = m_lendersRefXML.GetAttribute("REDEMPTIONDATE");
			frmScreen.txtHomeImpr.value = m_lendersRefXML.GetAttribute("BALHOUSEIMPROVEMENT");
			frmScreen.txtLoanDate.value = m_lendersRefXML.GetAttribute("DATEOFLOAN");
			frmScreen.txtCapitalRaising.value = m_lendersRefXML.GetAttribute("BALCAPITALRAISING");
			frmScreen.txtLoanNumber.value = m_lendersRefXML.GetAttribute("LOANNUMBER");
			frmScreen.txtOrigHousePurchase.value = m_lendersRefXML.GetAttribute("ORIGINALHOUSEPURCHASE");
			if(m_lendersRefXML.GetAttribute("GOODACCOUNTCONDUCT") == "1")
				frmScreen.optAccountConductYes.checked = true;
			else if(m_lendersRefXML.GetAttribute("GOODACCOUNTCONDUCT") == "0")
				frmScreen.optAccountConductNo.checked = true;
			if(m_lendersRefXML.GetAttribute("PREVIOUSARREARS") == "1")
				frmScreen.optPreviousArrearsYes.checked = true;
			else if(m_lendersRefXML.GetAttribute("PREVIOUSARREARS") == "0")
				frmScreen.optPreviousArrearsNo.checked = true;
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
		scScreenFunctions.SetFieldState(frmScreen, "txtLendersName", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtCustLoanNumber", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtCustRepayment", "R");
	}
}
function UpdateLenderRef()
{
	var XML = null;
	var sASPFile = "";
	if(m_bCreatePrevLendersRef == true)
	{
		// CREATE a new prev lenders ref
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var reqTag = XML.CreateRequestTag(window , "CreatePrevLendersRef");
		XML.CreateActiveTag("CASETASK");
		XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		XML.SetAttribute("ACTIVITYID", m_taskXML.GetAttribute("ACTIVITYID"));
		XML.SetAttribute("CASEID", m_taskXML.GetAttribute("CASEID"));
		XML.SetAttribute("TASKID", m_taskXML.GetAttribute("TASKID"));
		XML.SetAttribute("TASKINSTANCE", m_taskXML.GetAttribute("TASKINSTANCE"));
		XML.SetAttribute("STAGEID", m_taskXML.GetAttribute("STAGEID"));
		XML.SetAttribute("CASESTAGESEQUENCENO", m_taskXML.GetAttribute("CASESTAGESEQUENCENO"));
		XML.ActiveTag = reqTag;
		XML.CreateActiveTag("PREVIOUSLENDERSREF");
		XML.SetAttribute("ACCOUNTGUID", m_taskXML.GetAttribute("CONTEXT"))
		sASPFile = "OmigaTMBO.asp";
	}
	else
	{
		// UPDATE the existing prev lenders ref
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window , "UpdatePrevLendersRef");
		m_lendersRefXML.SelectTag(null, "RESPONSE");
		XML.AddXMLBlock(m_lendersRefXML.CreateFragment());
		XML.SelectTag(null, "PREVIOUSLENDERSREF");
		sASPFile = "omAppProc.asp";
	}
	XML.SetAttribute("MONTHLYREPAYMENT",frmScreen.txtRepayment.value);
	XML.SetAttribute("BALHOUSEPURCHASE",frmScreen.txtHousePurchase.value);
	XML.SetAttribute("REDEMPTIONDATE",frmScreen.txtRedemptionDate.value);
	XML.SetAttribute("BALHOUSEIMPROVEMENT",frmScreen.txtHomeImpr.value);
	XML.SetAttribute("DATEOFLOAN",frmScreen.txtLoanDate.value);
	XML.SetAttribute("BALCAPITALRAISING",frmScreen.txtCapitalRaising.value);
	XML.SetAttribute("LOANNUMBER",frmScreen.txtLoanNumber.value);
	XML.SetAttribute("ORIGINALHOUSEPURCHASE",frmScreen.txtOrigHousePurchase.value);
	if(frmScreen.optAccountConductYes.checked == true)
		XML.SetAttribute("GOODACCOUNTCONDUCT","1");
	else if(frmScreen.optAccountConductNo.checked == true)
		XML.SetAttribute("GOODACCOUNTCONDUCT","0");
	else XML.SetAttribute("GOODACCOUNTCONDUCT","");
	if(frmScreen.optPreviousArrearsYes.checked == true)
		XML.SetAttribute("PREVIOUSARREARS","1");
	else if(frmScreen.optPreviousArrearsNo.checked == true)
		XML.SetAttribute("PREVIOUSARREARS","0");
	else XML.SetAttribute("PREVIOUSARREARS","");
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
		if ((m_bCreatePrevLendersRef == false) && (m_bSetTaskToPending == true)) {
			
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
function ValidateLenderRef()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = XML.CreateRequestTag(window , "ValidatePrevLendersRef");
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
	XML.CreateActiveTag("VALPREVLENDREF");
	XML.SetAttribute("ACCOUNTGUID", m_taskXML.GetAttribute("CONTEXT"));
	//SYS2275 SA 2/5/01 Rules Component needs these too {
	XML.SetAttribute("CUSTOMERNUMBER",m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
	XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
	// SYS2275 2/5/01 SA }
	// 	XML.RunASP(document, "OmigaTMBO.asp")
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "OmigaTMBO.asp")
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
//SYS2275 SA 2/5/01 Now we need to get the customer version number attribute for the
// ValidateLendersRef function, new function GetCustomerVersion added below 
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
//SYS2275 Changed Complete Task button to Confirm Button
//function frmScreen.btnConfirm.onclick()
function btnConfirm.onclick()
{
	<% /* PSC 12/12/01 SYS3388 */ %>
	if (confirm("Please ensure all data has been entered correctly before continuing"))
	{	
		//run rules and set task status.
		if(UpdateLenderRef())
		{
			if(ValidateLenderRef())
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
		if(UpdateLenderRef())
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


