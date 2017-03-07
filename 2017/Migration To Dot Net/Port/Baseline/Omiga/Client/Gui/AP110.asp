<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP110.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Accountant's reference.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		17/01/01	Created
JLD		19/01/01	Added some screen processing
JLD		23/01/01	On exit of screen go to TM030, Deal with underReview.
JLD		25/01/01	change EMPLOYMENTSEQUENCENO  to EMPLOYMENTSEQUENCENUMBER in line with db
					Add some methods.
JLD		14/02/01	Modifications to get it running with Task Management.
CL		05/03/01	SYS1920 Read only functionality added
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
SA		02/05/01	SYS2275 Removed Complete Task button and added Confirm Button
					Existing click code moved from Complete Task button to Confirm button
SA		03/12/01	SYS3285/3280 If Application set to "under Review" - don't make it read only. 
PSC		12/12/01	SYS3388 Prompt before running confirm process
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
TW      09/10/2002  SYS5115		Modified to incorporate client validation - 
AW		24/10/2002	BMIDS00653	BM029 Call to Allowable Income
MDC		04/11/2002	BMIDS00654	BM088 Maximum Borrowing
MV		11/11/2002	BMIDS00916	Self Employed Accountant ref error in btnSubmit.onclick()
MDC		16/11/2002	BMIDS00917  Handle record not found in GetSelfEmployedDetails
MDC		21/11/2002	BMIDS01034  Return true from RunIncomeCalculation allowing user to progress
MO		21/11/2002	BMIDS01037 - Creating and Updating of references and casetasks result in 
					duplicate key violations
GHun	13/03/2003	BM0457		Include guarantors in allowable income calculation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
MF		15/09/2005	MAR30		IncomeCalcs modified to send ActivityID into calculation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Forms Here */ %>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div id="divClient" style="HEIGHT: 75px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Date Established
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtCustDateEstablished" maxlength="10" style="WIDTH: 70px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 290px; POSITION: absolute; TOP: 7px" class="msgLabel">
	Date Financial<BR>Interest Held
	<span style="LEFT: 110px; POSITION: absolute; TOP: 0px">
		<input id="txtCustDateFinInterest" maxlength="10" style="WIDTH: 70px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Registration number
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtCustRegNumber" maxlength="50" style="WIDTH: 150px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 290px; POSITION: absolute; TOP: 37px" class="msgLabel">
	Number of Years<BR>Acting for Client
	<span style="LEFT: 110px; POSITION: absolute; TOP: 0px">
		<input id="txtCustYearsActing" maxlength="2" style="WIDTH: 50px" class="msgTxt">
	</span>
</span>
</div>
<div id="divRef1" style="HEIGHT: 105px; LEFT: 10px; POSITION: absolute; TOP: 140px; WIDTH: 604px" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Date Established
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtDateEstablished" maxlength="10" style="WIDTH: 70px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 290px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Tax Office
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtTaxOffice" maxlength="60" style="WIDTH: 150px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Position Held
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtPositionHeld" maxlength="40" style="WIDTH: 150px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 290px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Tax Reference
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtTaxRef" maxlength="50" style="WIDTH: 150px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 67px" class="msgLabel">
	Number of years<BR>Accounts held
	<span style="LEFT: 110px; POSITION: absolute; TOP: 0px">
		<input id="txtYearsActing" maxlength="2" style="WIDTH: 50px" class="msgTxt">
	</span>
</span>
</div>
<div id="divRef2" style="HEIGHT: 285px; LEFT: 10px; POSITION: absolute; TOP: 250px; WIDTH: 604px" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Year End
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtYearEnd1" maxlength="4" style="WIDTH: 70px" class="msgTxt">
	</span>
	<span style="LEFT: 250px; POSITION: absolute; TOP: -3px">
		<input id="txtYearEnd2" maxlength="4" style="WIDTH: 70px" class="msgTxt">
	</span>
	<span style="LEFT: 390px; POSITION: absolute; TOP: -3px">
		<input id="txtYearEnd3" maxlength="4" style="WIDTH: 70px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Annual Turnover
	<span style="LEFT: 110px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtAnnualTurnover1" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 250px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtAnnualTurnover2" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 390px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtAnnualTurnover3" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>		
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
	Gross Profit
	<span style="LEFT: 110px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtGrossProfit1" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 250px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtGrossProfit2" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 390px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtGrossProfit3" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 100px" class="msgLabel">
	Net Profit Entitlement
	<span style="LEFT: 110px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtNetProfit1" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 250px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtNetProfit2" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 390px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtNetProfit3" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 130px" class="msgLabel">
	Salary
	<span style="LEFT: 110px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtSalary1" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 250px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtSalary2" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 390px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtSalary3" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 160px" class="msgLabel">
	Dividend
	<span style="LEFT: 110px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtDividend1" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 250px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtDividend2" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 390px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtDividend3" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 190px" class="msgLabel">
	Capital &amp; Reserve
	<span style="LEFT: 110px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtCapital1" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 250px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtCapital2" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
	<span style="LEFT: 390px; POSITION: absolute; TOP: 0px">
		<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency">£</label>
		<input id="txtCapital3" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 220px" class="msgLabel">
	Current Ratio
	<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
		<input id="txtRatio1" maxlength="3" style="WIDTH: 70px" class="msgTxt">
		<label style="LEFT: 75px; POSITION: absolute; TOP: 3px" class="msgLabel">%</label>
	</span>
	<span style="LEFT: 250px; POSITION: absolute; TOP: -3px">
		<input id="txtRatio2" maxlength="3" style="WIDTH: 70px" class="msgTxt">
		<label style="LEFT: 75px; POSITION: absolute; TOP: 3px" class="msgLabel">%</label>
	</span>
	<span style="LEFT: 390px; POSITION: absolute; TOP: -3px">
		<input id="txtRatio3" maxlength="3" style="WIDTH: 70px" class="msgTxt">
		<label style="LEFT: 75px; POSITION: absolute; TOP: 3px" class="msgLabel">%</label>
	</span>
</span>
<% /* SYS2275 SA Complete Task button replaced with confirm */ %>
<% /* <span style="TOP:250px; LEFT:4px; POSITION:ABSOLUTE">
	<input id="btnComplete" value="Complete Task" type="button" style="WIDTH:100px" class="msgButton">
</span> */ %>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 540px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP110attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sTaskXML = "";
var m_taskXML = null;
var m_accountXML = null;
var scScreenFunctions;
var m_bCreateFirstAccountantsRef = null;
var m_blnReadOnly = false;
//SYS2275 SA 1/5/01
var m_sAppFactFindNo = "";
var m_sApplicationNumber = "";
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
	FW030SetTitles("Accountant's Reference","AP110",scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
	SetMasks();
	Validation_Init();
	PopulateScreen();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP110");
	
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
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
//DEBUG
//scScreenFunctions.SetContextParameter(window,"idCustomerNumber1","00001546");
//scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber1","1");
//m_sTaskXML = "<CASETASK SOURCEAPPLICATION=\"Omiga\" CASEID=\"C00071021\" ACTIVITYID=\"10\" TASKID=\"Second_Charge\" CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"2\" TASKSTATUS=\"20\" TASKINSTANCE=\"1\" STAGEID=\"60\" CASESTAGESEQUENCENO=\"6\"/>";
}
function PopulateScreen()
{
	if(m_sTaskXML.length > 0)
	{
		m_taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_taskXML.LoadXML(m_sTaskXML);
		m_taskXML.SelectTag(null, "CASETASK");
		if(m_taskXML.GetAttribute("TASKSTATUS") == "40") m_sReadOnly = "1";
		//Get Applicant entered info
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window , "");
		XML.CreateActiveTag("SELFEMPLOYEDDETAILS");
		XML.CreateTag("CUSTOMERNUMBER",m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
		XML.CreateTag("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
		XML.CreateTag("EMPLOYMENTSEQUENCENUMBER", m_taskXML.GetAttribute("CONTEXT"));
		XML.RunASP(document,"GetSelfEmployedDetails.asp");
		<% /* BMIDS00917 MDC 16/11/2002 - Handle record not found */ %>
		// if(XML.IsResponseOK())
		var SelfEmpErrorTypes = new Array("RECORDNOTFOUND");
		var SelfEmpErrorReturn = XML.CheckResponse(SelfEmpErrorTypes);
		if(SelfEmpErrorReturn[0] == true)
		{
			XML.SelectTag(null, "SELFEMPLOYEDDETAILS");
			frmScreen.txtCustDateEstablished.value = XML.GetTagText("DATESTARTEDORESTABLISHED");
			frmScreen.txtCustDateFinInterest.value = XML.GetTagText("DATEFINANCIALINTERESTHELD");
			frmScreen.txtCustRegNumber.value = XML.GetTagText("REGISTRATIONNUMBER");
			frmScreen.txtCustYearsActing.value = XML.GetTagText("YEARSACTINGFORCUSTOMER");
			frmScreen.txtYearEnd1.value = XML.GetTagText("YEAR1");
			frmScreen.txtYearEnd2.value = XML.GetTagText("YEAR2");
			frmScreen.txtYearEnd3.value = XML.GetTagText("YEAR3");
		}
<%		//Populate the Accountants reference info if some already stored. The task status 
		//may be 'Pending' but still have no reference details input yet if the task
		//has an OutputDocument associated with it. We need to cater for this.
%>		<% /* MO - 21/11/2002 - BMIDS01037 - Rewritten */ %>
		<% /* if(m_taskXML.GetAttribute("TASKSTATUS") != "10")
		{
			//Get previously saved ref data
			m_accountXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			m_accountXML.CreateRequestTag(window , "GetAccountantsRef");
			m_accountXML.CreateActiveTag("ACCOUNTANT");
			m_accountXML.SetAttribute("CUSTOMERNUMBER",m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
			m_accountXML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
			m_accountXML.SetAttribute("EMPLOYMENTSEQUENCENUMBER", m_taskXML.GetAttribute("CONTEXT"));
			m_accountXML.RunASP(document, "omAppProc.asp");
			var ErrorTypes = new Array("RECORDNOTFOUND");
			var ErrorReturn = m_accountXML.CheckResponse(ErrorTypes);
			if (ErrorReturn[1] == ErrorTypes[0])
			{
				if(m_taskXML.GetAttribute("OUTPUTDOCUMENT") == "")
				{
					alert("Accountants reference cannot be found");
					m_bCreateFirstAccountantsRef = false;
				}
				else m_bCreateFirstAccountantsRef = true;
			}
			else if(ErrorReturn[0] == true)
			{
				m_bCreateFirstAccountantsRef = false;
				m_accountXML.SelectTag(null, "ACCOUNTANTREF");
				frmScreen.txtDateEstablished.value = m_accountXML.GetAttribute("DATEESTABLISHEDACCOUNTANT");
				frmScreen.txtTaxOffice.value = m_accountXML.GetAttribute("TAXOFFICE");
				frmScreen.txtTaxRef.value = m_accountXML.GetAttribute("TAXREFERENCE");
				frmScreen.txtPositionHeld.value = m_accountXML.GetAttribute("POSITIONHELD");
				frmScreen.txtYearsActing.value = m_accountXML.GetAttribute("YEARSACCOUNTHELD");
				frmScreen.txtYearEnd1.value = m_accountXML.GetAttribute("YEAREND1");
				frmScreen.txtYearEnd2.value = m_accountXML.GetAttribute("YEAREND2");
				frmScreen.txtYearEnd3.value = m_accountXML.GetAttribute("YEAREND3");
				frmScreen.txtAnnualTurnover1.value = m_accountXML.GetAttribute("ANNUALTURNOVER1");
				frmScreen.txtAnnualTurnover2.value = m_accountXML.GetAttribute("ANNUALTURNOVER2");
				frmScreen.txtAnnualTurnover3.value = m_accountXML.GetAttribute("ANNUALTURNOVER3");
				frmScreen.txtGrossProfit1.value = m_accountXML.GetAttribute("GROSSPROFIT1");
				frmScreen.txtGrossProfit2.value = m_accountXML.GetAttribute("GROSSPROFIT2");
				frmScreen.txtGrossProfit3.value = m_accountXML.GetAttribute("GROSSPROFIT3");
				frmScreen.txtNetProfit1.value = m_accountXML.GetAttribute("NETPROFITENTITLEMENT1");
				frmScreen.txtNetProfit2.value = m_accountXML.GetAttribute("NETPROFITENTITLEMENT2");
				frmScreen.txtNetProfit3.value = m_accountXML.GetAttribute("NETPROFITENTITLEMENT3");
				frmScreen.txtSalary1.value = m_accountXML.GetAttribute("SALARY1");
				frmScreen.txtSalary2.value = m_accountXML.GetAttribute("SALARY2");
				frmScreen.txtSalary3.value = m_accountXML.GetAttribute("SALARY3");
				frmScreen.txtDividend1.value = m_accountXML.GetAttribute("DIVIDEND1");
				frmScreen.txtDividend2.value = m_accountXML.GetAttribute("DIVIDEND2");
				frmScreen.txtDividend3.value = m_accountXML.GetAttribute("DIVIDEND3");
				frmScreen.txtCapital1.value = m_accountXML.GetAttribute("CAPITALRESERVE1");
				frmScreen.txtCapital2.value = m_accountXML.GetAttribute("CAPITALRESERVE2");
				frmScreen.txtCapital3.value = m_accountXML.GetAttribute("CAPITALRESERVE3");
				frmScreen.txtRatio1.value = m_accountXML.GetAttribute("CURRENTRATIO1");
				frmScreen.txtRatio2.value = m_accountXML.GetAttribute("CURRENTRATIO2");
				frmScreen.txtRatio3.value = m_accountXML.GetAttribute("CURRENTRATIO3");
			}
		}
		else m_bCreateFirstAccountantsRef = true; */ %>
		//Get previously saved ref data
		m_accountXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_accountXML.CreateRequestTag(window , "GetAccountantsRef");
		m_accountXML.CreateActiveTag("ACCOUNTANT");
		m_accountXML.SetAttribute("CUSTOMERNUMBER",m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
		m_accountXML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
		m_accountXML.SetAttribute("EMPLOYMENTSEQUENCENUMBER", m_taskXML.GetAttribute("CONTEXT"));
		m_accountXML.RunASP(document, "omAppProc.asp");
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = m_accountXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{
			m_bCreateFirstAccountantsRef = true;
		}
		else if(ErrorReturn[0] == true)
		{
			m_bCreateFirstAccountantsRef = false;
			m_accountXML.SelectTag(null, "ACCOUNTANTREF");
			frmScreen.txtDateEstablished.value = m_accountXML.GetAttribute("DATEESTABLISHEDACCOUNTANT");
			frmScreen.txtTaxOffice.value = m_accountXML.GetAttribute("TAXOFFICE");
			frmScreen.txtTaxRef.value = m_accountXML.GetAttribute("TAXREFERENCE");
			frmScreen.txtPositionHeld.value = m_accountXML.GetAttribute("POSITIONHELD");
			frmScreen.txtYearsActing.value = m_accountXML.GetAttribute("YEARSACCOUNTHELD");
			frmScreen.txtYearEnd1.value = m_accountXML.GetAttribute("YEAREND1");
			frmScreen.txtYearEnd2.value = m_accountXML.GetAttribute("YEAREND2");
			frmScreen.txtYearEnd3.value = m_accountXML.GetAttribute("YEAREND3");
			frmScreen.txtAnnualTurnover1.value = m_accountXML.GetAttribute("ANNUALTURNOVER1");
			frmScreen.txtAnnualTurnover2.value = m_accountXML.GetAttribute("ANNUALTURNOVER2");
			frmScreen.txtAnnualTurnover3.value = m_accountXML.GetAttribute("ANNUALTURNOVER3");
			frmScreen.txtGrossProfit1.value = m_accountXML.GetAttribute("GROSSPROFIT1");
			frmScreen.txtGrossProfit2.value = m_accountXML.GetAttribute("GROSSPROFIT2");
			frmScreen.txtGrossProfit3.value = m_accountXML.GetAttribute("GROSSPROFIT3");
			frmScreen.txtNetProfit1.value = m_accountXML.GetAttribute("NETPROFITENTITLEMENT1");
			frmScreen.txtNetProfit2.value = m_accountXML.GetAttribute("NETPROFITENTITLEMENT2");
			frmScreen.txtNetProfit3.value = m_accountXML.GetAttribute("NETPROFITENTITLEMENT3");
			frmScreen.txtSalary1.value = m_accountXML.GetAttribute("SALARY1");
			frmScreen.txtSalary2.value = m_accountXML.GetAttribute("SALARY2");
			frmScreen.txtSalary3.value = m_accountXML.GetAttribute("SALARY3");
			frmScreen.txtDividend1.value = m_accountXML.GetAttribute("DIVIDEND1");
			frmScreen.txtDividend2.value = m_accountXML.GetAttribute("DIVIDEND2");
			frmScreen.txtDividend3.value = m_accountXML.GetAttribute("DIVIDEND3");
			frmScreen.txtCapital1.value = m_accountXML.GetAttribute("CAPITALRESERVE1");
			frmScreen.txtCapital2.value = m_accountXML.GetAttribute("CAPITALRESERVE2");
			frmScreen.txtCapital3.value = m_accountXML.GetAttribute("CAPITALRESERVE3");
			frmScreen.txtRatio1.value = m_accountXML.GetAttribute("CURRENTRATIO1");
			frmScreen.txtRatio2.value = m_accountXML.GetAttribute("CURRENTRATIO2");
			frmScreen.txtRatio3.value = m_accountXML.GetAttribute("CURRENTRATIO3");
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
		<% /* frmScreen.btnComplete.disabled = true; */ %>
		DisableMainButton("Confirm");
		DisableMainButton("Submit");
	}
	else
	{
		scScreenFunctions.SetFieldState(frmScreen, "txtCustDateEstablished", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtCustDateFinInterest", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtCustRegNumber", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtCustYearsActing", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtYearEnd1", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtYearEnd2", "R");
		scScreenFunctions.SetFieldState(frmScreen, "txtYearEnd3", "R");
	}
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
function UpdateAccountantRef()
{
	var XML = null;
	var sASPFile = "";
	if(m_bCreateFirstAccountantsRef == true)
	{
		// CREATE a new accountants ref
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var reqTag = XML.CreateRequestTag(window , "CreateAccountantsRef");
		XML.CreateActiveTag("CASETASK");
		XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		XML.SetAttribute("ACTIVITYID", m_taskXML.GetAttribute("ACTIVITYID"));
		XML.SetAttribute("STAGEID", m_taskXML.GetAttribute("STAGEID"));
		XML.SetAttribute("CASESTAGESEQUENCENO", m_taskXML.GetAttribute("CASESTAGESEQUENCENO"));
		XML.SetAttribute("CASEID", m_taskXML.GetAttribute("CASEID"));
		XML.SetAttribute("TASKID", m_taskXML.GetAttribute("TASKID"));
		XML.SetAttribute("TASKINSTANCE", m_taskXML.GetAttribute("TASKINSTANCE"));
		XML.ActiveTag = reqTag;
		XML.CreateActiveTag("ACCOUNTANTREF");
		sASPFile = "OmigaTMBO.asp";
	}
	else
	{
		// UPDATE the existing accountants ref
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window , "UpdateAccountantsRef");
		m_accountXML.SelectTag(null, "RESPONSE");
		XML.AddXMLBlock(m_accountXML.CreateFragment());
		XML.SelectTag(null, "ACCOUNTANTREF");
		sASPFile = "omAppProc.asp";
	}
	XML.SetAttribute("CUSTOMERNUMBER",m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
	XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
	XML.SetAttribute("EMPLOYMENTSEQUENCENUMBER", m_taskXML.GetAttribute("CONTEXT"));
	XML.SetAttribute("DATEESTABLISHEDACCOUNTANT", frmScreen.txtDateEstablished.value);
	XML.SetAttribute("YEARSACCOUNTHELD", frmScreen.txtYearsActing.value);
	XML.SetAttribute("TAXOFFICE", frmScreen.txtTaxOffice.value);
	XML.SetAttribute("TAXREFERENCE", frmScreen.txtTaxRef.value);
	XML.SetAttribute("POSITIONHELD", frmScreen.txtPositionHeld.value);
	XML.SetAttribute("YEAREND1", frmScreen.txtYearEnd1.value);
	XML.SetAttribute("ANNUALTURNOVER1", frmScreen.txtAnnualTurnover1.value);
	XML.SetAttribute("GROSSPROFIT1", frmScreen.txtGrossProfit1.value);
	XML.SetAttribute("NETPROFITENTITLEMENT1", frmScreen.txtNetProfit1.value);
	XML.SetAttribute("SALARY1", frmScreen.txtSalary1.value);
	XML.SetAttribute("DIVIDEND1", frmScreen.txtDividend1.value);
	XML.SetAttribute("CAPITALRESERVE1", frmScreen.txtCapital1.value);
	XML.SetAttribute("CURRENTRATIO1", frmScreen.txtRatio1.value);
	XML.SetAttribute("YEAREND2", frmScreen.txtYearEnd2.value);
	XML.SetAttribute("ANNUALTURNOVER2", frmScreen.txtAnnualTurnover2.value);
	XML.SetAttribute("GROSSPROFIT2", frmScreen.txtGrossProfit2.value);
	XML.SetAttribute("NETPROFITENTITLEMENT2", frmScreen.txtNetProfit2.value);
	XML.SetAttribute("SALARY2", frmScreen.txtSalary2.value);
	XML.SetAttribute("DIVIDEND2", frmScreen.txtDividend2.value);
	XML.SetAttribute("CAPITALRESERVE2", frmScreen.txtCapital2.value);
	XML.SetAttribute("CURRENTRATIO2", frmScreen.txtRatio2.value);
	XML.SetAttribute("YEAREND3", frmScreen.txtYearEnd3.value);
	XML.SetAttribute("ANNUALTURNOVER3", frmScreen.txtAnnualTurnover3.value);
	XML.SetAttribute("GROSSPROFIT3", frmScreen.txtGrossProfit3.value);
	XML.SetAttribute("NETPROFITENTITLEMENT3", frmScreen.txtNetProfit3.value);
	XML.SetAttribute("SALARY3", frmScreen.txtSalary3.value);
	XML.SetAttribute("DIVIDEND3", frmScreen.txtDividend3.value);
	XML.SetAttribute("CAPITALRESERVE3", frmScreen.txtCapital3.value);
	XML.SetAttribute("CURRENTRATIO3", frmScreen.txtRatio3.value);
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
		if ((m_bCreateFirstAccountantsRef == false) && (m_bSetTaskToPending == true)) {
			
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
function ValidateAccountantRef()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = XML.CreateRequestTag(window , "ValidateAccountantsRef");
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
	XML.CreateActiveTag("VALACCTREF");
	XML.SetAttribute("CUSTOMERNUMBER",m_taskXML.GetAttribute("CUSTOMERIDENTIFIER"));
	XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_taskXML.GetAttribute("CUSTOMERIDENTIFIER")));
	XML.SetAttribute("EMPLOYMENTSEQUENCENUMBER", m_taskXML.GetAttribute("CONTEXT"));
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
//function frmScreen.btnComplete.onclick()
function btnConfirm.onclick()
{
	<% /* PSC 12/12/01 SYS3388 */ %>
	if (confirm("Please ensure all data has been entered correctly before continuing"))
	{	
		//run rules and set task status.
		if(UpdateAccountantRef())
		{
			if(ValidateAccountantRef())
			{
				<% /* MF MAR30 Only run income calcs if data has changed */ %>
				var bContinue = true;
				if(IsChanged()){
					bContinue = RunIncomeCalculations();
				}
				if (bContinue)	//AW	24/10/2002	BMIDS00653
				{
					scScreenFunctions.SetContextParameter(window,"idTaskXML","");
					frmToTM030.submit();
				}
			}
		}
	} 
}
function btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
	{
		if(UpdateAccountantRef())
		{
			<% /* MF MAR30 Only run income calcs if data has changed */ %>
			var bContinue = true;
			if(IsChanged()){
				bContinue = RunIncomeCalculations();
			}
			if (bContinue)	//AW	24/10/2002	BMIDS00653
			{
				scScreenFunctions.SetContextParameter(window,"idTaskXML","");
				frmToTM030.submit();
			}
		}
	}
}
function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idTaskXML","");
	frmToTM030.submit();
}

<% /* BMIDS00654 MDC 04/11/2002 */ %>
<% /* AW	24/10/2002	BMIDS00653 */ %>
function RunIncomeCalculations()
{
	var AllowableIncXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AllowableIncXML.CreateRequestTag(window,null);
	<% /* MAR30 */ %> 
	AllowableIncXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));
		
	// Set up request for Income Calculation
	var TagRequest = AllowableIncXML.CreateActiveTag("INCOMECALCULATION");
	AllowableIncXML.SetAttribute("CALCULATEMAXBORROWING","1");
	AllowableIncXML.SetAttribute("QUICKQUOTEMAXBORROWING","0");
	AllowableIncXML.CreateActiveTag("APPLICATION");
	AllowableIncXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AllowableIncXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
	AllowableIncXML.ActiveTag = TagRequest;
	<% /* MAR30 */ %> 
	AllowableIncXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));
	
	var TagCUSTOMERLIST = AllowableIncXML.CreateActiveTag("CUSTOMERLIST");

	for(var nLoop = 1; nLoop <= 2; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		<% /* BM0457 The customer must exist */ %>
		if (sCustomerNumber.trim().length > 0)
		{
		<% /* BM0457 End */ %>
			var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
			var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);
			var sCustomerOrder = scScreenFunctions.GetContextParameter(window,"idCustomerOrder" + nLoop);

			<% /* BM0457
			if(sCustomerRoleType == "1")
			{ */ %>
			AllowableIncXML.CreateActiveTag("CUSTOMER");
			AllowableIncXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			AllowableIncXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
			AllowableIncXML.CreateTag("CUSTOMERROLETYPE", sCustomerRoleType);
			AllowableIncXML.CreateTag("CUSTOMERORDER", sCustomerOrder);
			AllowableIncXML.ActiveTag = TagCUSTOMERLIST;
		}
	}
	
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
					AllowableIncXML.RunASP(document,"RunIncomeCalculations.asp");
			break;
		default: // Error
			AllowableIncXML.SetErrorResponse();
	}
	
	<% /* BMIDS01034 MDC 21/11/2002 - Always return true to allow user to progress
	if(AllowableIncXML.IsResponseOK())
	{
		return(true);
	}

	return(false); */ %>
	AllowableIncXML.IsResponseOK()
	return(true);
	<% /* BMIDS01034 MDC 21/11/2002 - End */ %>
}
<%	/*BMIDS00653  - End */ %>
<% /* BMIDS00654 MDC 04/11/2002 - End */ %>

-->
</script>
</body>
</html>


