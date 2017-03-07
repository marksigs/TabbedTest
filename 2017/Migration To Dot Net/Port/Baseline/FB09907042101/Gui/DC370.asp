<%@ Language=JScript %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%/*
Description:	This screen summarises the customers Income details and gets the 
				Credit Check Result.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AM		19/08/2005	New Decision Screen Created
MV		21/09/2005	MAR29 -  Removed CreditCheck DecisionResult
JD		28/09/2005	MAR30 - Get verified income from IncomeSummary table.
MV		17/10/2005  MAR188 - Amended
MV		21/10/2005  MAR228	- AMended btnsubmit.onclick()
MV		27/10/2005	MAR305	Amended Window.onload()
PJO     09/11/2005  MAR475  Disable Case Assessment Decision Result text field
HMA     14/11/2005  MAR507  Increase size of Affordability text box 
Maha T	15/11/2005	MAR272	(Point 5)Do not show "No details are updated" warning message.
INR		16/11/2005	MAR295	DecisionButton needs to run RunCreditCheck & sort out Outstanding tasks.
INR		18/11/2005	MAR295	Needs to run RunCreditCheck even if no Outstanding credit check tasks.
Maha T	28/11/2005	MAR656	Call RunIncomeCalculations() before calling PopulateDeclaredIncome().
Maha T	29/11/2005	MAR639	1. Display Decision result when Get Decision button clicked.
							2. On page load GetLatestRiskAssessment into Case Assessment Decision Result.
							3. Disable credit check radio buttons if credit check already done.
INR		02/12/2005  MAR791	When Credit Check success disable btnGetDecision.
HMA     06/12/2005  MAR756  Display cost to 2 decimal places.
PJO     13/12/2005  MAR791  Decision radios and push button disabling following credit check
PJO     14/12/2005  MAR800  Message during Get Decision
RF		13/02/2006  MAR1243 Address targeting implemented 
HMA     23/02/2006  MAR1315 Get latest Risk Assessment for application regardless of stage
JD		24/02/2006	MAR1314 Expand the text displayed for the income multiple types.
PJO     06/03/2006  MAR1359 Improve status message
PJO     09/03/2006  MAR1359 Parameterise width of progress message
GHun	20/03/2006	MAR1453	Hide progress window before popping up an error
JD		29/03/2006	MAR1542 display only one maximum borrowing figure.
JD		05/07/2006	MAR1890 remove call to RunIncomeCalculation. Call new stored proc to get data to display
PSC		03/08/2006	MAR1928	Amend CreditCheck so that it doesn't set the first credit check to complete
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>

<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title>DC370 - Decision</title>
</head>

<body>

<object id="scClientFunctions" 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex="-1" 
type="text/x-scriptlet" height="1" width="1" data="scClientFunctions.asp" 
viewastext></object>

<form id="frmToDC085" method="post" action="DC085.asp" style="DISPLAY: none"></form>
<form id="frmToDC210" method="post" action="dc210.asp" style="DISPLAY: none"></form>
<form id="frmToDC200" method="post" action="dc200.asp" style="DISPLAY: none"></form>
<form id="frmToRA040" method="post" action="ra040.asp" style="DISPLAY: none"></form>
<form id="frmScreen" validate="onchange" mark>

<% // PJO 14/12/2005 MAR800 Added %>
<% /* PJO 14/12/2005 MAR1359 Removed 
<div id="divStatus" style="TOP: 60px; LEFT: 10px; WIDTH: 614px; HEIGHT: 100px; POSITION: absolute" class="msgGroup">
	<table id="tblStatus" width="604px" height="100px" border="0" cellspacing="0" cellpadding="0">
		<tr align="center"><td id="colStatus" align="center" class="msgLabelWait">Please Wait ... Getting Decision</td></tr>
	</table>
</div> */ %>  
<div id="divScreen" style="LEFT: 8px; WIDTH: 604px; POSITION: absolute; TOP: 58px; HEIGHT: 471px" class="msgGroup">
		
	<span style="LEFT: 4px;  POSITION: absolute; TOP: 5px" class="msgLabel">
		<strong>Affordability</strong>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 20px"><textarea class="msgReadOnly" id="txtAffordability" style="LEFT: 14px; WIDTH: 569px; TOP: 14px; HEIGHT: 158px" rows="11" readOnly cols="19">		</textarea>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 185px" class="msgLabel">
		<strong>Quote Details</strong> 
	</span>
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 205px" class="msgLabel">
		The amount you request to borrow
		<span style="LEFT: 160px; POSITION: absolute; TOP: 0px">
			<input id="txtAmountRequested" maxlength="9" style="LEFT: 21px; WIDTH: 100px; POSITION: absolute; TOP: -3px" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 310px; POSITION: absolute; TOP: 205px" class="msgLabel">
		at a cost of
	</span>
	<span style="LEFT: 160px; POSITION: absolute; TOP: 0px">
		<input id="txtTotalMonthlyCost" maxlength="9" style="LEFT: 220px; WIDTH: 100px; POSITION: absolute; TOP: 203px" class="msgReadOnly" readonly tabindex=-1>
	</span>
	
	<span style="LEFT: 11px; POSITION: absolute; TOP: 233px" class="msgLabel">
		Your Allowable Income
	</span>
	<span style="LEFT: 283px; POSITION: absolute; TOP: 233px" class="msgLabel">
		Declared
	</span>
	<span style="LEFT: 431px; POSITION: absolute; TOP: 233px" class="msgLabel">
		Verified
	</span>

	<span id="lblCustomerName1" style="LEFT: 13px; POSITION: absolute; TOP: 256px" class="msgLabel">
	</span>
	<span id="lblCustomerName2" style="LEFT: 12px; POSITION: absolute; TOP: 280px" class="msgLabel">
	</span>	
	<span  style="LEFT: 11px; POSITION: absolute; TOP: 304px" class="msgLabel">
		Calculated income multiple
	</span>	
	<span style="LEFT: 12px; POSITION: absolute; TOP: 328px" class="msgLabel">   
		Income multiple Type Applied
	</span>
	<span style="LEFT: 12px; WIDTH: 251px; POSITION: absolute; TOP: 352px; HEIGHT: 28px" class="msgLabelHeadInOrange">
		<strong>Based on what you have told me the maximum you can apply for is £</strong>
	</span>

	<input class="msgReadOnly" id="txtDeclaredIncome1" style="LEFT: 282px; WIDTH: 115px; POSITION: absolute; TOP: 257px" tabIndex=-1 readOnly maxLength="9">
	<input class="msgReadOnly" id="txtVerifiedIncome1" style="LEFT: 432px; WIDTH: 115px; POSITION: absolute; TOP: 257px" tabIndex=-1 readOnly maxLength="9">
	<input class="msgReadOnly" id="txtDeclaredIncome2" style="LEFT: 282px; VISIBILITY: visible; WIDTH: 115px; POSITION: absolute; TOP: 281px" tabIndex=-1 readOnly maxLength="9">
	<input class="msgReadOnly" id="txtVerifiedIncome2" style="LEFT: 432px; VISIBILITY: visible; WIDTH: 115px; POSITION: absolute; TOP: 281px" tabIndex=-1 readOnly maxLength="9">
	
	<input class="msgReadOnly" id="txtDeclaredIncomeMultiple" style="LEFT: 282px; WIDTH: 115px; POSITION: absolute; TOP: 305px" tabIndex=-1 readOnly maxLength="4">
	<input class="msgReadOnly" id="txtVerifiedIncomeMultiple" style="LEFT: 432px; WIDTH: 115px; POSITION: absolute; TOP: 305px" tabIndex=-1 readOnly maxLength="5">
	<input class="msgReadOnly" id="txtDeclaredMultiplierType" style="LEFT: 282px; WIDTH: 115px; POSITION: absolute; TOP: 329px" tabIndex=-1 readOnly maxLength="1">
	<input class="msgReadOnly" id="txtVerifiedMultiplierType" style="LEFT: 432px; WIDTH: 115px; POSITION: absolute; TOP: 329px" tabIndex=-1 readOnly maxLength="1">
	<input class="msgReadOnly" id="txtMaximumBorrowing" style="LEFT: 380px; WIDTH: 70px; POSITION: absolute; TOP: 356px; HEIGHT: 18px" tabIndex=-1 readOnly maxLength="9" size="10">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 392px" class="msgLabelHeadInOrange">
		<strong>This is subject to credit checks</strong>
	</span>
	
	<span class=msgLabel style="LEFT: 11px; POSITION: absolute; TOP: 418px">
		Are you willing to proceed with the credit check process 
	</span>
	
	<span style="LEFT: 310px; POSITION: absolute; TOP: 412px">
		<input id="optAnswer_Yes" type="radio" value ="1" name="Answer">
		<label class="msgLabel" for="optAnswer_Yes">Yes</label> 
	</span>
	<span style="LEFT: 356px; POSITION: absolute; TOP: 412px">
		<input id="optAnswer_No" style="LEFT: 0px; TOP: 2px" type="radio" checked value="0" name="Answer">
		<label class="msgLabel" for="optAnswer_No">No</label> 
	</span>
	<span style="LEFT: 434px; POSITION: absolute; TOP: 409px">
		<input class="msgButton" id="btnGetDecision" style="WIDTH: 71px; HEIGHT: 25px" type="button" value="GetDecision">
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 443px" class="msgLabel">
		Case Assessment Decision Result
		<input id="txtCaseAssessmentDecision" style="LEFT: 178px; WIDTH: 100px; POSITION: absolute" maxLength="45" class="msgReadOnly"> 
	</span>
</div>	
</form><!-- Main Buttons -->
	
<% /* span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0">
</span>
<span id="spnToLastField" tabindex="0">
</span>
<!-- Main Buttons -->	
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 529px; HEIGHT: 19px">
<!-- #include FILE="msgButtons.asp" --> 
	</div> 
<% /* File containing field attributes */ %>
<!-- #include FILE="attribs/DC370Attribs.asp" --><!-- #include FILE="fw030.asp" -->

<% /* Code */ %>
<script language="JScript" type="text/javascript">

var m_sMetaAction					= null;
var m_sApplicationNumber			= null;

var m_sApplicationFactFindNumber	= null;
////var m_sReadOnly	;
var m_sCustomerName1;
var m_sCurrentCustomerNumber;
var m_sCustomer1Number = "";
var m_sCustomerName2;

var m_sCustomerNumber = "";
var m_sCustomerNumber1 = "";
var m_sCustomerVersionNumber1 = "";
var m_sCustomerVersionNumber2 = "";
var m_CurrEmpXML = null;
var m_sCustomerRoleType = null;
var m_nMaxCustomers = 0;
var scScreenFunctions;
var scClientScreenFunctions;
var sCreditWillilng =0;
var m_sStageId = "";
var m_TaskXML = null;
var m_sReadOnly;
var m_blnReadOnly;
var m_sCurTaskCaseStageSeqNo = "";
var m_sActivityID = "";
var m_sUserID =	"";	
var m_sUnitID =	"";

/** EVENTS **/
function window.onload()
{
	/* Get ScreenFunctions object from frame */
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	//GetRulesData();
	/*Setup Main Buttons */
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Decision","DC370",scScreenFunctions);
	Initialise();
	
	RetrieveContextData();	
	scScreenFunctions.ShowCollection(frmScreen);
	
	ClientPopulateScreen();
	PopulateData();
	//PopulateDeclaredIncome(); JD MAR1890
	//PopulateVerifiedIncome(); JD MAR30
	PopulateScreen();
	<% /* MAR639 - Maha T */ %>
	CheckAndShowRiskAssessment();
	Validation_Init();
	
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	if (m_sReadOnly == "1")
		m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC370");	
			
	<% // PJO 14/12/2005 MAR800 %>
	<% // PJO 14/12/2005 MAR1359 - removed	divStatus.style.visibility = "hidden";%>
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
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
	
	//m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", "1");
	m_sCustomerName1 = scScreenFunctions.GetContextParameter(window,"idCustomerName1",null);
	m_sCustomerName2 = scScreenFunctions.GetContextParameter(window,"idCustomerName2",null);
	m_sCustomerVersionNumber1 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1",null);
	m_sCustomerVersionNumber2 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber2",null);
	m_sCustomerNumber1 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber1",null);
	m_sCustomerNumber2 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber2",null);
	//m_sCustomerRoleType = scScreenFunctions.SetContextParameter(window,"idCustomerRoleType1",null);
	//m_sCustomerOrder1 = scScreenFunctions.SetContextParameter(window,"idCustomerOrder1",null);
	//m_sCustomerOrder2 = scScreenFunctions.SetContextParameter(window,"idCustomerOrder2",null);

	//MAR295
	m_sActivityID = scScreenFunctions.GetContextParameter(window,"idActivityId",null);
	m_sStageId = scScreenFunctions.GetContextParameter(window,"idStageId",null);
	m_sUserID =	scScreenFunctions.GetContextParameter(window,"idUserId",null);
	m_sUnitID =	scScreenFunctions.GetContextParameter(window,"idUnitId",null);

	<%/* update number of applicants */ %>
	if (m_sCustomerNumber2 == "") 
		m_nMaxCustomers = 1;
	else
		m_nMaxCustomers = 2;
		
	////scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	var TaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	if(TaskXML != "")
	{
		m_TaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_TaskXML.LoadXML(m_sTaskXML);
		m_TaskXML.SelectTag(null, "CASETASK");
		////m_sInsSeqNo = m_TaskXML.GetAttribute("CONTEXT");
		////m_sTaskID = m_TaskXML.GetAttribute("TASKID")	
	}	
		
}
function Initialise()
{
	frmScreen.btnGetDecision.disabled = true;
	// PJO 09/11/2005 MAR475 Disable Case Assessment Decision Result text field
	frmScreen.txtCaseAssessmentDecision.readOnly = true;
	frmScreen.optAnswer_Yes.focus();	
}

<%/* To hide the text boxes if only one Customer exist */%>
function PopulateScreen()
{
	lblCustomerName1.innerText = m_sCustomerName1;
	
	if (m_sCustomerName2.replace(/\s+/, "") != "")
	{
		lblCustomerName2.innerText = m_sCustomerName2;
	}

}


<% /* Populating AMOUNTREQUESTED,TOTALNETMONTHLYCOST,INCOMEMULTIPLE,CONFIRMEDCALCULATEDINCMULTIPLE,
INCOMEMULTIPLIERTYPE,CONFIRMEDCALCULATEDINCMULTIPLIERTYPE,MAXIMUMBORROWINGAMOUNT,CONFIRMEDMAXBORROWING */%>
function PopulateData()
{
	
	////var sAmountRequested = ""; 
	////var sTotalMonthlyCost = ""; 
	
	<% /* JD MAR1890 call GetDecisionDetails */ %>
	var AppXML = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
	AppXML.CreateRequestTag(window, "GetDecisionDetails");
	AppXML.CreateActiveTag("DECISIONDETAILS");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AppXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	AppXML.RunASP(document,"GetDecisionDetails.asp");
	
	if(AppXML.IsResponseOK())
	{
		
		var RespTag = AppXML.SelectTag(null, "RESPONSE");
	
		if (AppXML.ActiveTag)
		{
			frmScreen.txtAmountRequested.value  = AppXML.GetTagAttribute("DECISIONDETAILS", "AMOUNTREQUESTED");			
			
			frmScreen.txtTotalMonthlyCost.value =    
						top.frames[1].document.all.scMathFunctions.RoundValue(AppXML.GetTagAttribute("DECISIONDETAILS","TOTALNETMONTHLYCOST"), 2);			
			
			frmScreen.txtDeclaredIncomeMultiple.value = AppXML.GetTagAttribute("DECISIONDETAILS","INCOMEMULTIPLE");
			frmScreen.txtVerifiedIncomeMultiple.value = AppXML.GetTagAttribute("DECISIONDETAILS","CONFIRMEDCALCULATEDINCMULTIPLE");
			<% /* JD MAR1314 expand income multiple types text */ %>
			frmScreen.txtDeclaredMultiplierType.value = GetIncMultTypeText(AppXML.GetTagAttribute("DECISIONDETAILS","INCOMEMULTIPLIERTYPE")); 
			frmScreen.txtVerifiedMultiplierType.value = GetIncMultTypeText(AppXML.GetTagAttribute("DECISIONDETAILS","CONFIRMEDCALCULATEDINCMULTIPLIERTYPE"));
			<% /* JD MAR1542 display only one maximum borrowing figure.*/ %>
			if(AppXML.GetTagAttribute("DECISIONDETAILS","CONFIRMEDMAXBORROWING")!= "" && AppXML.GetTagAttribute("DECISIONDETAILS","CONFIRMEDMAXBORROWING") != "0")
				frmScreen.txtMaximumBorrowing.value = AppXML.GetTagAttribute("DECISIONDETAILS","CONFIRMEDMAXBORROWING"); 
			else
				frmScreen.txtMaximumBorrowing.value = AppXML.GetTagAttribute("DECISIONDETAILS","MAXIMUMBORROWINGAMOUNT"); 
				
			//now populate the customerdata (customers 1 and 2)
			AppXML.SelectSingleNode("DECISIONDETAILS[@CUSTOMERORDER='1']");
			if(AppXML.ActiveTag != null)
			{
				frmScreen.txtDeclaredIncome1.value = AppXML.GetAttribute("NETALLOWABLEANNUALINCOME");
				frmScreen.txtVerifiedIncome1.value = AppXML.GetAttribute("NETCONFIRMEDALLOWABLEINCOME");
			}
			AppXML.ActiveTag = RespTag;
			AppXML.SelectSingleNode("DECISIONDETAILS[@CUSTOMERORDER='2']");
			if(AppXML.ActiveTag != null)
			{
				frmScreen.txtDeclaredIncome2.value = AppXML.GetAttribute("NETALLOWABLEANNUALINCOME");
				frmScreen.txtVerifiedIncome2.value = AppXML.GetAttribute("NETCONFIRMEDALLOWABLEINCOME");
			}
			
		}
		<% /* START: MAR272 Maha T */ %>
		<% /*
		else
			alert("No Details are Updated");
		*/%>
		<% /* END: MAR272 */ %>
		
	}	
}
<% /* JD MAR1314 added function */ %>
function GetIncMultTypeText(sIncMultType)
{
	var sIncMultTypeText;
	
	switch(sIncMultType)
	{
		case "S": sIncMultTypeText = "S - Single"; break;
		case "J": sIncMultTypeText = "J - Joint"; break;
		case "H": sIncMultTypeText = "HL - Highest Lowest"; break;
		case "U": sIncMultTypeText = "U - Unknown"; break;
		default : sIncMultTypeText = sIncMultType; break;
	}
	return sIncMultTypeText;
}
<%/**/%>
	<%/*Populating the VerifiedIncome fields for Two Customers */%>
function PopulateVerifiedIncome()

{
	for(var iCount = 1; iCount <= m_nMaxCustomers; iCount++)
	{
		if (iCount == 1)
		{
			var sCustomerNumber = m_sCustomerNumber1;
			var sCustomerVersionNumber = m_sCustomerVersionNumber1;
		}
		else
		{
			sCustomerNumber = m_sCustomerNumber2;
			sCustomerVersionNumber = m_sCustomerVersionNumber2;
		}
				
		m_CurrEmpXML = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_CurrEmpXML.CreateRequestTag(window, "GetCurrEmployersRef");
		m_CurrEmpXML.CreateActiveTag("CURRENTEMPLOYERSREF");
		m_CurrEmpXML.SetAttribute("CUSTOMERNUMBER", sCustomerNumber);
		m_CurrEmpXML.SetAttribute("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
		
		m_CurrEmpXML.RunASP(document,"omAppProc.asp");
		
		if(m_CurrEmpXML.IsResponseOK())
		{
			m_CurrEmpXML.SelectTag(null, "CURRENTEMPLOYERSREF");
			if (m_CurrEmpXML.ActiveTag)			
			{				
				if (iCount==1)
				{
					frmScreen.txtVerifiedIncome1.value = m_CurrEmpXML.GetAttribute("NETALLOWABLEINCOME");
				}
				else
				{
					frmScreen.txtVerifiedIncome2.value = m_CurrEmpXML.GetAttribute("NETALLOWABLEINCOME");
				}  
			}
		}
		else
		alert("No Details of Cutomers exist");
		
	}
	
}

	<% /* Populating the Declared Income for the  Two Customers*/%>
function PopulateDeclaredIncome()
{
for(var iCount = 1; iCount <= m_nMaxCustomers; iCount++)
	{
		if (iCount == 1)
		{
			var sCustomerNumber = m_sCustomerNumber1;
			var sCustomerVersionNumber = m_sCustomerVersionNumber1;
		}
		else
		{
			sCustomerNumber = m_sCustomerNumber2;
			sCustomerVersionNumber = m_sCustomerVersionNumber2;
		}
			
		m_DecEmpXml = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_DecEmpXml.CreateRequestTag(window, "GetIncomeSummary");
		m_DecEmpXml.CreateActiveTag("INCOMESUMMARY");
	
		m_DecEmpXml.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
		m_DecEmpXml.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
		m_DecEmpXml.RunASP(document,"GetIncomeSummary.asp");
	
		if(m_DecEmpXml.IsResponseOK())
		{
			m_DecEmpXml.SelectTag(null, "INCOMESUMMARY");
			if (m_DecEmpXml.ActiveTag)
			{
				var sAmountIncomeAmount= "";
			
				if (iCount==1)
				{
					frmScreen.txtDeclaredIncome1.value = m_DecEmpXml.GetTagText("NETALLOWABLEANNUALINCOME");
					frmScreen.txtVerifiedIncome1.value = m_DecEmpXml.GetTagText("NETCONFIRMEDALLOWABLEINCOME"); //JD MAR30
				}
				else
				{
					frmScreen.txtDeclaredIncome2.value = m_DecEmpXml.GetTagText("NETALLOWABLEANNUALINCOME");
					frmScreen.txtVerifiedIncome2.value = m_DecEmpXml.GetTagText("NETCONFIRMEDALLOWABLEINCOME"); //JD MAR30
				}
			}
			else
			alert("No Details of Customers exist");
		}
	}
}

function frmScreen.optAnswer_Yes.onclick()
{
	frmScreen.btnGetDecision.disabled = false;
	frmScreen.btnGetDecision.focus();
}
function frmScreen.optAnswer_No.onclick()
{
	frmScreen.btnGetDecision.disabled = true;
}

function frmScreen.btnGetDecision.onclick()
{
	<% /* PJO 14/12/2005  MAR800
	      PJO 06/03/2006 MAR1359 - change to popup
	divScreen.style.display = "none";
	divStatus.style.display = "block"; */ %>
	scScreenFunctions.progressOn("Please Wait ... Getting Decision", 400);

	<% // 14/12/2005 PJO MAR800 - To force screen redisplay	%>
	window.setTimeout(DecisionHandler, 0);
}


function DecisionHandler()
{
	var bSuccess = true;
	var sCreditCheck;
	
	if(frmScreen.optAnswer_Yes.checked)
		sCreditCheck = "1";

	var ApplicationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ApplicationXML.CreateRequestTag(window, "UPDATE");
	ApplicationXML.CreateActiveTag("APPLICATIONFACTFIND");
	ApplicationXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	ApplicationXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
			
	ApplicationXML.CreateTag("CREDITCHECKWILLINGTOPROCEED", sCreditCheck);
	//ApplicationXML.RunASP(document,"UpdateApplicationFactFind.asp");
	///////////if(ApplicationXML.IsResponseOK())
	//{
	switch (ScreenRules())
	{
	case 1: // Warning
	case 0: // OK
		ApplicationXML.RunASP(document,"UpdateApplicationFactFind.asp");
		break;
	default: // Error
		ApplicationXML.SetErrorResponse();
	}
	
	scScreenFunctions.progressOff();
	bSuccess = ApplicationXML.IsResponseOK();
	//}

	//MAR295 Make the call to CreditCheck
	if(bSuccess)
	{
		CreditCheck(); 
	}
	<%// 14/12/2005 PJO MAR800 - Screen redisplay %>
	<% /* PJO 06/03/2006 MAR1359 - change to a popup 
	divStatus.style.display = "none"; 
	divScreen.style.display = "block"; */ %>

	<% /* MAR1243 */ %>	
	var sAddressTargetXml = scScreenFunctions.GetContextParameter(window,"idAddressTarget");
	if (sAddressTargetXml.length > 0)
		frmToRA040.submit();
}

<%/*Populate the Credit Check Decision Result,Case Assessment Decision Result    */%>
function CreditCheck()
{
	scScreenFunctions.progressOn("Please Wait ... Getting Decision", 400);

	var bSuccess = false;
	var ccCallSuccess = false;
	var	initialTask = 0;
	var XML = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTMIntialCreditCheck = XML.GetGlobalParameterString(document, "TMInitialCreditCheck", null);

	var CreditXML = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
	CreditXML.CreateRequestTag(window, "FindCaseTaskList");
	CreditXML.CreateActiveTag("CASETASK");
	CreditXML.CreateTag("SOURCEAPPLICATIONNUMBER", "Omiga");
	CreditXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	CreditXML.CreateTag("STAGEID", m_sStageId);
	CreditXML.SetAttribute("CASEID",m_sApplicationNumber);

	CreditXML.CreateTag("TASKID", sTMIntialCreditCheck);
	
	CreditXML.RunASP(document,"MsgTMBO.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = CreditXML.CheckResponse(ErrorTypes);
	
	if(ErrorReturn[1] == ErrorTypes[0])
	{
		<%/* No outstanding tasks*/%>
	}
	else if (ErrorReturn[0] == true)
	{
		CreditXML.ActiveTag = null;
		CreditXML.CreateTagList("CASETASK");
		CreditXML.SelectTagListItem(0);
		m_sCurTaskCaseStageSeqNo = CreditXML.GetAttribute('CASESTAGESEQUENCENO');
		
		<% /* Have we any outstanding tasks which will have to be completed */ %>	
		CreditXML.ActiveTag = null;
		CreditXML.CreateTagList("CASETASK[@TASKID=\"" + sTMIntialCreditCheck + "\" and (@TASKSTATUS=\"10\" or @TASKSTATUS=\"20\")]");

		if (CreditXML.ActiveTagList.length == 0)
		{
			<%/* No outstanding creditcheck Tasks */%>
			if(! GetCreditCheckStatus())
			{
				<%/* Credit Check not been done  */%>
				CreditXML = CreateCCinTaskManager(sTMIntialCreditCheck);
				<%/* No point calling if we don't have a task, won't get through TM */%>
				if(CreditXML)
				{
					if(RunTheCreditCheck(CreditXML))
					{
						ccCallSuccess = true;
						//all outstanding tasks are complete
						initialTask = CreditXML.ActiveTagList.length;
						<% /* MAR639: Maha T - Call GetLatestRiskAssessment to display the decision result*/ %>
						GetLatestRiskAssessment();
						<% // PJO 13/12/2005 MAR791 - Disable buttons %>
						frmScreen.btnGetDecision.disabled = true;
						frmScreen.optAnswer_Yes.checked = true;	
						frmScreen.optAnswer_Yes.disabled = true;
						frmScreen.optAnswer_No.disabled = true;
					}
				}
				else
				{
					scScreenFunctions.progressOff();
					alert("Failed to create and get Credit Check Task.");
				}
			}
		}
		else
		{
			<% /* Call New Function, have we already done a Creditcheck, as
				we don't want to do more than one from here  */ %>
			if(! GetCreditCheckStatus())
			{
				if(RunTheCreditCheck(CreditXML))
				{
					ccCallSuccess = true;
					//First outstanding task should be complete
					initialTask = 1;
					
					<% /* MAR639: Maha T - Call GetLatestRiskAssessment to display the decision result*/ %>
					GetLatestRiskAssessment();
				}
			}
			else
			{
				<%/* Credit Check has been done
				but we have an outstanding task to complete  */%>
				ccCallSuccess = true;
			}
	
			<% /* If we have any outstanding creditcheck tasks - complete them */ %>
			if((CreditXML.ActiveTagList.length > 0) && (ccCallSuccess))
			{
				<% /* 1st one has been completed in creditcheck call, initialtask 
				will then be set to start at the second one if there is one*/ %>
				<% /* PSC 03/08/2006 MAR1928 */ %>
				for(var iCount = initialTask; iCount < CreditXML.ActiveTagList.length; ++iCount)
				{
					CreditXML.SelectTagListItem(iCount);
					//TaskStatus to Completed
					CreditXML.SetAttribute('TASKSTATUS', '40');

					var taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					var ReqTag;
					
					ReqTag = taskXML.CreateRequestTag(window, "UpdateCaseTask");
					taskXML.SetAttribute("ACTION", "UpdateCaseTask");

					// Pass the CASETASK tag to get through validation etc.
					taskXML.ActiveTag.appendChild(CreditXML.ActiveTag.cloneNode(true));

					taskXML.RunASP(document, "msgTMBO.asp");

					if(taskXML.IsResponseOK())
					{
						//First outstanding task should be complete
						blnSuccess = true;	
							
					}
					else
					{
						scScreenFunctions.progressOff();
						alert("Failed to update Credit Check Task.");
					}
				}
			}
		}

		scScreenFunctions.progressOff();
	}
}
//MAR295
function CreateCCinTaskManager(sTMIntialCreditCheck)
{
	var globalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sActivityId = globalXML.GetGlobalParameterString(document,"TMOmigaActivity",null);
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	var reqTag = XML.CreateRequestTag(window , "GetTaskDetail");
	XML.SetAttribute("ACTION","CreateAdhocCaseTask");
	XML.CreateActiveTag("TASK");
	XML.SetAttribute("CASEPRIORITY",  scScreenFunctions.GetContextParameter(window,"idApplicationPriority",""));
	XML.SetAttribute("TASKID",  sTMIntialCreditCheck);
	XML.SetAttribute("STAGEID", m_sStageId);

	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "msgTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}
		
	if(XML.IsResponseOK()) 
	{		
		<% /* Use Info to Create the task */ %>
		XML.SelectTag(null,"RESPONSE/TASK");
		
		var dtTaskDue =  scScreenFunctions.DateTimeToString(scScreenFunctions.GetAppServerDate(true));
		createXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		createXML.CreateRequestTag(window , "CreateCaseTask");
		createXML.SetAttribute("ACTION","CreateAdhocCaseTask");
		createXML.CreateActiveTag("CASETASK");
		createXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		createXML.SetAttribute("CASEID", m_sApplicationNumber);
	
		createXML.SetAttribute("CASESTAGESEQUENCENO", m_sCurTaskCaseStageSeqNo);
		createXML.SetAttribute("ACTIVITYID", m_sActivityID);
		createXML.SetAttribute("ACTIVITYINSTANCE", "1");
		createXML.SetAttribute("STAGEID", m_sStageId);
		createXML.SetAttribute("TASKID",  sTMIntialCreditCheck);
		createXML.SetAttribute("OUPUTDOCUMENT", "");
		createXML.SetAttribute("ALWAYSAUTOMATICCONCREATION", XML.GetAttribute("ALWAYSAUTOMATICONCREATION"));
		createXML.SetAttribute("CASETASKNAME", XML.GetAttribute("TASKNAME"));
		
		createXML.SetAttribute("CUSTOMERIDENTIFIER", "");
		createXML.SetAttribute("CONTEXT", "");
		
		createXML.SetAttribute("MANDATORYFLAG",  XML.GetAttribute("MANDATORYFLAG"));
		createXML.SetAttribute("TASKDUEDATEANDTIME", dtTaskDue);
		if(XML.GetAttribute("OWNINGUSERID"))
			createXML.SetAttribute("OWNINGUSERID", XML.GetAttribute("OWNINGUSERID"));
		else
			//ApplicationOwner if not set up against the task
			createXML.SetAttribute("OWNINGUSERID", m_sUserID);
		if(XML.GetAttribute("OWNINGUNITID"))
			createXML.SetAttribute("OWNINGUNITID", XML.GetAttribute("OWNINGUNITID"));
		else
			//ApplicationOwner if not set up against the task
			createXML.SetAttribute("OWNINGUNITID", m_sUnitID);
		createXML.SetAttribute("ORIGINATINGSTAGEID", m_sStageId);
		createXML.SetAttribute("MANDATORYINDICATOR", "0");
		
		createXML.RunASP(document, "msgTMBO.asp");
		
		if (createXML.IsResponseOK()) 
		{
			return createXML;
		}
		else 
		{
			return null;
		}			
	} 
	else 
	{
		return null;
	}
}

function RunTheCreditCheck(CreditXML)
{
	scScreenFunctions.SetContextParameter(window,"idAddressTarget", null); <% /* MAR1243 */ %>
	var XML = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTMIntialCreditCheck = XML.GetGlobalParameterString(document,"TMInitialCreditCheck",null);
	var CreditCheckXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ReqTag;
	
	ReqTag = CreditCheckXML.CreateRequestTag(window, "RUNXMLCREDITCHECK");
	CreditCheckXML.SetAttribute("CREDITCHECKINVOKEDFROMACTIONBUTTON", "TRUE");
	CreditCheckXML.CreateActiveTag("APPLICATION");
	CreditCheckXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	CreditCheckXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	CreditCheckXML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window, "idApplicationPriority", ""));
	CreditCheckXML.SelectTag(null,"REQUEST");

	CreditXML.CreateTagList("CASETASK[@TASKID=\"" + sTMIntialCreditCheck + "\"]");
	if (CreditXML.ActiveTagList.length > 0)
	{
		// Pass the CASETASK tag to get through validation etc.
		CreditXML.SelectTagListItem(0);
		CreditCheckXML.ActiveTag.appendChild(CreditXML.ActiveTag.cloneNode(true));
	}

	CreditCheckXML.RunASP(document, "OmigaTMBO.asp");
	
	if(CreditCheckXML.IsResponseOK())
	{
		<% /* MAR791 - Credit Check success disable the button*/ %>
		frmScreen.btnGetDecision.disabled = true;
		
		<% /* MAR1243 Address targeting implemented */ %> 		
		CreditCheckXML.SelectTag(null, "RESPONSE");
		CreditCheckXML.SelectTag(CreditCheckXML.ActiveTag, "TARGETINGDATA"); //Revised XML structure in MARS
		
		var sIsAddrTarget = "";
		if (CreditCheckXML.ActiveTag != null)
			sIsAddrTarget = CreditCheckXML.GetTagText("ADDRESSTARGETING");
		
		CreditCheckXML.SelectTag(null, "RESPONSE");
		
		if (sIsAddrTarget == "YES")
		{
			scScreenFunctions.SetContextParameter(window,"idReturnScreenId","DC370");
			scScreenFunctions.SetContextParameter(window,"idAddressTarget", CreditCheckXML.ActiveTag.xml);
		}
		
		return true;
	}
	else
	{
		alert("Run Credit Check failed.");
		return false;
	}
}

<% /*Updating Willing to CreditCheckWillingToProceed and Route to New Property Address */%>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
function btnSubmit.onclick()
{
	var bSuccess = true;
	var sCreditCheck;
	
	if(frmScreen.optAnswer_Yes.checked)
		sCreditCheck = "1";

	var ApplicationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ApplicationXML.CreateRequestTag(window,"UPDATE");
	ApplicationXML.CreateActiveTag("APPLICATIONFACTFIND");
	ApplicationXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ApplicationXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	ApplicationXML.CreateTag("CREDITCHECKWILLINGTOPROCEED",sCreditCheck);
	
	switch (ScreenRules())
	{
		case 1: <%/* Warning*/%>
		case 0: <%/* OK*/%>
			ApplicationXML.RunASP(document,"UpdateApplicationFactFind.asp");
			break;
		default:  <%/*Error*/%>
			ApplicationXML.SetErrorResponse();
	}
	bSuccess = ApplicationXML.IsResponseOK();
	
	if (bSuccess == true)
	{
		if(ApplicationXML.GetGlobalParameterBoolean(document,"NewPropertySummary"))
			frmToDC210.submit();
		else
			frmToDC200.submit();
	}		
}
	<% /*Route to Financial Summary */%>
function btnCancel.onclick()
{
	frmToDC085.submit();
}

<% /* INR MAR295 New Function  */ %>
function GetCreditCheckStatus()
{
	var bSuccess = false;
	var nNumCreditCheck;

	CreditCheckXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	with (CreditCheckXML)
	{
		<% /* Retrieve Credit Check status */ %>
		CreateRequestTag(window, null)

		<% /* Create request block */ %>
		CreateActiveTag("SEARCH");
		CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);

		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					RunASP(document, "GetCreditCheckStatus.asp");
				break;
			default: // Error
				SetErrorResponse();
			}

		if (!IsResponseOK())
		{
			alert("ERROR : Call to ApplicationCreditCheck failed.");
		}
		else
		{
			SelectSingleNode("//APPLICATIONCREDITCHECK");
			if (ActiveTag != null) 
			{
				nNumCreditCheck = GetAttribute("SEQUENCENUMBER");
				if (nNumCreditCheck > 0) 
					bSuccess = true;
			}
		}
	}

	return bSuccess;
}

<% /* MAR639: Maha T - New Function  */ %>
function CheckAndShowRiskAssessment()
{
	// 1st check if already decision is made, if so disable the radio credit check radio buttons
	// and also get decision button
	if (GetCreditCheckStatus())
	{
		frmScreen.btnGetDecision.disabled = true;
		frmScreen.optAnswer_Yes.checked = true;	
		frmScreen.optAnswer_Yes.disabled = true;
		frmScreen.optAnswer_No.disabled = true;
	}
	
	// Now display the Latest RiskAssessment
	GetLatestRiskAssessment();
}

<% /* MAR639: Maha T - New Function  */ %>
function GetLatestRiskAssessment()
{
	<% /* MAR1315 Get the latest Risk Assessment for this application (regardless of stage) */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"Search");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);

	XML.RunASP(document,"GetLatestRANoStage.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{	<% /* It worked */ %>
		frmScreen.txtCaseAssessmentDecision.value = XML.GetTagAttribute("RISKASSESSMENTDECISION","TEXT");
		return true;
	}
	else 
	{ 	<% /* It failed */ %>
		frmScreen.txtCaseAssessmentDecision.value = "";
		return false;
	}
}

</script>
<script src="validation.js" language="JScript" type="text/javascript"></script>
</body>
</html>
