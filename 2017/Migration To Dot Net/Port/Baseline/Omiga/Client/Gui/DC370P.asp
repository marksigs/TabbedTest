<%@ Language=JScript %>
<HTML>
<%
/*
	Despcripttion:	This screen summarise back  the customers Income details and get the 
	Credit Check Result
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
----	----		-----------
AM		19/08/2005	New Decision Screen Created
JD		28/09/2005	MAR30 - Get verified income from IncomeSummary table.
HMA     14/11/2005  MAR507  Increase size of Affordability text area.
Maha T	29/11/2005	MAR639	1. On page load GetLatestRiskAssessment into Case Assessment Decision Result.
							2. Disable credit check radio buttons if credit check already done.
							3. Ammended screen attributes.
							4. Remove "Credit Check Decision Result" label and input field.
							5. Do not show "No details are updated" warning message.
HMA     23/02/2006  MAR1315 Get latest Risk Assessment for application regardless of stage
JD		24/02/2006	MAR1314 Expand the text displayed for the income multiple types.
JD		29/03/2006	MAR1542 display only one maximum borrowing figure.
JD		07/07/2006	MAR1890 Get display data from new stored procedure.
JD		12/07/2006	MAR1856 always set credit check radio disabled
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<HEAD>
	<meta name=2vs_targetSchema2 content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title> DC370P- Decision  </title>
</HEAD>
<body>
<script src="validation.js" language="JScript"></script>

<OBJECT id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scScreenFunctions.asp 
viewastext></OBJECT>
<OBJECT id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp 
viewastext></OBJECT>
<OBJECT id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scXMLFunctions.asp 
viewastext></OBJECT>

<form id="frmToDC085" method="post" action="DC085.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC210" method="post" action="dc210.asp" style="DISPLAY: none"></form>
<form id="frmScreen" validate    ="onchange" mark>

<div style="LEFT: 8px; WIDTH: 604px; POSITION: absolute; TOP: 10px; HEIGHT: 409px" class="msgGroup">
		
	<span style="LEFT: 4px;  POSITION: absolute; TOP: 5px" class="msgLabel">
		<strong>Affordability</strong>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 20px"><TEXTAREA class=msgReadOnly id=txtAffordability style="LEFT: 10px; WIDTH: 573px; TOP: 13px" rows=5 readOnly></TEXTAREA>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 100px" class="msgLabel">
		<strong>Quote Details</strong> 
	</span>
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 123px" class="msgLabel">
		The amount you request to borrow
		<span style="LEFT: 160px; POSITION: absolute; TOP: 3px">
			<input id="txtAmountRequested" name="txtAmountRequested" maxlength="9" style="LEFT: 21px; WIDTH: 100px; POSITION: absolute; TOP: -3px" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 310px; POSITION: absolute; TOP: 123px" class="msgLabel">
		at a cost of
	</span>
	<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
		<input id="txtTotalMonthlyCost" maxlength="9" style="LEFT: 220px; WIDTH: 100px; POSITION: absolute; TOP: 127px" class="msgReadOnly" readonly tabindex=-1>
	</span>
	
	<span style="LEFT: 11px; POSITION: absolute; TOP: 147px" class="msgLabel">
		Your Allowable Income
	</span></SPAN>
		<span style="LEFT: 294px; POSITION: absolute; TOP: 147px" class="msgLabel">
		Declared
		</span>
		<span style="LEFT: 444px; POSITION: absolute; TOP: 147px" class="msgLabel">
		Verified
		</span></SPAN>
	<span id="lblCustomerName1" style="LEFT: 13px;  POSITION: absolute; TOP: 176px" class="msgLabel">
	</span>
	
	<span id="lblCustomerName2" style="LEFT: 12px; POSITION: absolute; TOP: 203px" class="msgLabel">
	</span>	
	<span style="LEFT: 12px; POSITION: absolute; TOP: 228px" class="msgLabel">
		Calculated income multiple
	</span>	
	<span style="LEFT: 12px; POSITION: absolute; TOP: 252px" class="msgLabel">
		Income multiple Type Applied
	</span>
	<span style="LEFT: 13px; WIDTH: 272px; POSITION: absolute; TOP: 275px; HEIGHT: 28px" class="msgLabelHeadInOrange">
		Based on what you have told me the maximum you can apply for is £
	</span>
	
	
	<input class="msgReadOnly" id="txtDeclaredIncome1" style="LEFT: 290px; WIDTH: 115px; POSITION: absolute; TOP: 176px" tabIndex =-1 readOnly maxLength=9>
	<input class="msgReadOnly" id="txtVerifiedIncome1" style="LEFT: 440px; WIDTH: 115px; POSITION: absolute; TOP: 176px" tabIndex =-1 readOnly maxLength=9>
	<input class="msgReadOnly" id="txtDeclaredIncome2" style="LEFT: 290px; VISIBILITY: visible; WIDTH: 115px; POSITION: absolute; TOP: 203px" tabIndex=-1 readOnly maxLength=9>
	<input class="msgReadOnly" id="txtVerifiedIncome2" style="LEFT: 440px; VISIBILITY: visible; WIDTH: 115px; POSITION: absolute; TOP: 203px" tabIndex =-1 readOnly maxLength=9></SPAN>
	
	<input class="msgReadOnly" id="txtDeclaredIncomeMultiple" style="LEFT: 290px; WIDTH: 115px; POSITION: absolute; TOP: 228px" tabIndex=-1 readOnly maxLength=4>
	<input class="msgReadOnly" id="txtVerifiedIncomeMultiple" style="LEFT: 440px; WIDTH: 115px; POSITION: absolute; TOP: 228px" tabIndex=-1 readOnly maxLength=5>
	<input class="msgReadOnly" id="txtDeclaredMultiplierType" style="LEFT: 290px; WIDTH: 115px; POSITION: absolute; TOP: 252px" tabIndex=-1 readOnly maxLength=1>
	<input class="msgReadOnly" id="txtVerifiedMultiplierType" style="LEFT: 440px; WIDTH: 115px; POSITION: absolute; TOP: 252px" tabIndex=-1 readOnly maxLength=1>
	<input class="msgReadOnly" id="txtMaximumBorrowing" style="LEFT: 380px; WIDTH: 80px; POSITION: absolute; TOP: 278px; HEIGHT: 18px" tabIndex=-1 readOnly maxLength=9 size=10>
	<span style="LEFT: 12px; POSITION: absolute; TOP: 315px" class="msgLabelHeadInOrange">
		<strong>This is subject to credit checks</strong>
	</span>
	
	<span class=msgLabel style="LEFT: 12px; POSITION: absolute; TOP: 339px">
		Are you willing to proceed with the credit check process 
	</span>
	
	<span style="LEFT: 318px; POSITION: absolute; TOP: 336px" id=SPAN1>
		<input id="optAnswer_Yes" type="radio" value="1" name="Answer" style="LEFT: 0px; TOP: 1px">
		<label class="msgLabel" for="optAnswer_Yes">Yes</label>
	</span>
	<span style="LEFT: 367px; POSITION: absolute; TOP: 336px">
		<input id="optAnswer_No" type="radio" checked value="0" name="Answer" style="LEFT: 0px; TOP: 1px">
		<label class="msgLabel" for="optAnswer_No">No</label> 
	</span>
	
	<span style="LEFT: 12px; POSITION: absolute; TOP: 365" class="msgLabel">
		Case Assessment Decision Result
	</span>
	
	<INPUT class="msgReadOnly" id="txtCaseAssessmentDecision" style="LEFT: 184px; WIDTH: 100px; POSITION: absolute; TOP: 365px" tabIndex=-1 readOnly maxLength=45>
	
	<% /* span to keep tabbing within this screen */ %>
	<span id="spnToFirstField" tabindex="0"></span>
	
	<span id="spnToLastField" tabindex="0"></span>
</div>
</form><!-- Main Buttons -->
<div id="msgButtons" style="LEFT: 8px; WIDTH: 597px; POSITION: absolute; TOP: 410px; HEIGHT: 19px">
<!-- #include FILE="msgButtons.asp" --> 
</div>
<% /* File containing field attributes */ %>
<!-- #include FILE="attribs/DC370PAttribs.asp" -->


<% /* Code */ %>
<script language="JScript">

var m_sMetaAction					= null;
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
var m_sReadOnly;
var m_sCustomerName1 ;
var m_sCurrentCustomerNumber ;
var m_sCustomer1Number = "";
var m_sCustomerName2 ;
var m_BaseNonPopupWindow = null ;
var m_sCustomerNumber = "" ;
var m_sCustomerNumber1 = "" ;
var m_sCustomerVersionNumber1 = "" ;
var m_sCustomerVersionNumber2 = "" ;
var AppXML = null ;
var m_CurrEmpXML = null ;
var m_DecEmpXml = null ;
var m_nMaxCustomers = 0 ;
var m_sCustomerNumber2 = null ;
var scScreenFunctions;
var scClientScreenFunctions;

	<%/*  EVENTS  */ %>
function window.onload()
{
	<%/* Setup Main Buttons */%>
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	<%/* Get ScreenFunctions object from frame */%>
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	
	
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	m_BaseNonPopupWindow = sArguments[5];
	
	<% /* Array contains the basic config including userid */ %>
	var sArgArray = sArguments[4];
	
	m_sUserId=sArgArray[0];
	m_sUnitId=sArgArray[1];
	m_sApplicationNumber = sArgArray[2];
	m_sApplicationFactFindNumber = sArgArray[3];
	m_sActiveApplicationFactFindNumber = m_sApplicationFactFindNumber;
	m_sStageNumber = sArgArray[4];
	m_sActiveStageNumber  = m_sStageNumber;
	m_arrRequestAttributes = sArgArray[5];

	scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();			
	scScreenFunctions.ShowCollection(frmScreen);
	ClientPopulateScreen();
	RetrieveContextData();
	PopulateData();
	PopulateScreen();
	<% /* MAR639 - Maha T */ %>
	CheckAndShowRiskAssessment();
	Validation_Init();
	//PopulateVerifiedIncome() ; JD MAR30
	//PopulateDeclaredIncome() ; JD MAR1890
	
}	

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idMetaAction", null);
	//m_sApplicationNumber = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idApplicationNumber", "1627");
	//m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idApplicationFactFindNumber", "1");
	m_sCustomerName1 = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerName1",null);
	m_sCustomerName2 = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerName2",null);
	m_sCustomerVersionNumber1 = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerVersionNumber1",null);
	m_sCustomerVersionNumber2 = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerVersionNumber2",null);
	m_sCustomerNumber1 = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerNumber1",null);
	m_sCustomerNumber2 = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerNumber2",null);

	<%/* update number of applicants */ %>
	if (m_sCustomerNumber2 == "") 
		m_nMaxCustomers = 1;
	else
		m_nMaxCustomers = 2;
}

<% /* Displaying the customer names */%>
function PopulateScreen()
{
	lblCustomerName1.innerText = m_sCustomerName1 ;
	if (m_sCustomerName2.replace(/\s+/, "") != "")
	{
		lblCustomerName2.innerText = m_sCustomerName2 ;
	}
}

function Initialise()
{
	scScreenFunctions.SetScreenToReadOnly(frmScreen);
}

<% /* Populating AMOUNTREQUESTED,TOTALNETMONTHLYCOST,INCOMEMULTIPLE,CONFIRMEDCALCULATEDINCMULTIPLE,
INCOMEMULTIPLIERTYPE,CONFIRMEDCALCULATEDINCMULTIPLIERTYPE,MAXIMUMBORROWINGAMOUNT,CONFIRMEDMAXBORROWING */%>
function PopulateData()
{
	var AppXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();	
	AppXML.CreateRequestTag(m_BaseNonPopupWindow, "GetDecisionDetails");
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
			frmScreen.txtTotalMonthlyCost.value = AppXML.GetTagAttribute("DECISIONDETAILS","TOTALNETMONTHLYCOST");   
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

		<% /* START: MAR639 Maha T */ %>
		<% /*
		else
			alert("No Details are Updated") ;
		*/%>
		<% /* END: MAR639 */ %>
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
	<%/* Populating the Verified Income for two Customers*/%>
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
	
		m_CurrEmpXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_CurrEmpXML.CreateRequestTag(m_BaseNonPopupWindow, "GetCurrEmployersRef");
		m_CurrEmpXML.CreateActiveTag("CURRENTEMPLOYERSREF");
		m_CurrEmpXML.SetAttribute("CUSTOMERNUMBER", sCustomerNumber);
		m_CurrEmpXML.SetAttribute("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
			
		m_CurrEmpXML.RunASP(document,"omAppProc.asp");
	
		if(m_CurrEmpXML.IsResponseOK())
		{
		m_CurrEmpXML.SelectTag(null, "CURRENTEMPLOYERSREF");	
		if (m_CurrEmpXML.ActiveTag)
		{
			var sEarnedIncomeAmount = "";
			var sAllowableIncome = "";
		
			
			sEarnedIncomeAmount = m_CurrEmpXML.GetAttribute("EARNEDINCOMEAMOUNT") ;
			sAllowableIncome = m_CurrEmpXML.GetAttribute("GROSSALLOWABLEINCOME") ;
			
			if (iCount==1)
				{
				frmScreen.txtVerifiedIncome1.value = m_CurrEmpXML.GetAttribute("NETALLOWABLEINCOME") ;
				}
			else
				{
				frmScreen.txtVerifiedIncome2.value = m_CurrEmpXML.GetAttribute("NETALLOWABLEINCOME") ;
				}
		}
		else
		alert("No Details of Customers exist");
		}
	}
	
}
	<%/* Populating the Declared Income for two Customers*/%>
<% /* JD MAR1890 This Method No Longer Called */ %>
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
		m_DecEmpXml = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject() ;
		m_DecEmpXml.CreateRequestTag(m_BaseNonPopupWindow, "GetIncomeSummary") ;
		m_DecEmpXml.CreateActiveTag("INCOMESUMMARY");
		m_DecEmpXml.CreateTag("CUSTOMERNUMBER", sCustomerNumber) ;
		m_DecEmpXml.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber) ;
		m_DecEmpXml.RunASP(document,"GetIncomeSummary.asp");
		if(m_DecEmpXml.IsResponseOK())
		{	
			m_DecEmpXml.SelectTag(null, "INCOMESUMMARY");					
			if (m_DecEmpXml.ActiveTag)
			{
				var sAmountIncomeAmount= "" ;
			
				
				//sAmountIncomeAmount = m_DecEmpXml.GetTagText("ALLOWABLEANNUALINCOME") ;
				if (iCount==1)
				{
				frmScreen.txtDeclaredIncome1.value = m_DecEmpXml.GetTagText("NETALLOWABLEANNUALINCOME") ;
				frmScreen.txtVerifiedIncome1.value = m_DecEmpXml.GetTagText("NETCONFIRMEDALLOWABLEINCOME") ;
				}
				else
				{
				frmScreen.txtDeclaredIncome2.value = m_DecEmpXml.GetTagText("NETALLOWABLEANNUALINCOME") ;
				frmScreen.txtVerifiedIncome2.value = m_DecEmpXml.GetTagText("NETCONFIRMEDALLOWABLEINCOME") ;
				}
			}
			else
			alert("No Details of Customers exist");
		}
	}
}

<% /* MAR639: Maha T -  New Function  */ %>
function GetCreditCheckStatus()
{
	var bSuccess = false;
	var nNumCreditCheck;

	CreditCheckXML = new scXMLFunctions.XMLObject();
	with (CreditCheckXML)
	{
		<% /* Retrieve Credit Check status */ %>
		CreateRequestTagFromArray(m_arrRequestAttributes,null);
		
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
				nNumCreditCheck = GetAttribute("SEQUENCENUMBER")
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
	if (GetCreditCheckStatus())
	{
		frmScreen.optAnswer_Yes.checked = true;	
	}
	
	// Now display the Latest RiskAssessment
	GetLatestRiskAssessment();
	
	<% /* JD MAR1856 always set disabled */ %>
	frmScreen.optAnswer_Yes.disabled = true;
	frmScreen.optAnswer_No.disabled = true;
}

<% /* MAR639: Maha T - New Function  */ %>
function GetLatestRiskAssessment()
{
	<% /* MAR1315 Get the latest Risk Assessment regardless of stage */ %>
	var XML = new scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Search");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sActiveApplicationFactFindNumber);

	XML.RunASP(document,"GetLatestRANoStage.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{	<% /* It worked */ %>
		frmScreen.txtCaseAssessmentDecision.value = XML.GetTagAttribute("RISKASSESSMENTDECISION","TEXT");;
		return true ;
	}
	else 
	{ 	<% /* It failed */ %>
		frmScreen.txtCaseAssessmentDecision.value = "";
		return false ;
	}
}

function btnSubmit.onclick()
{
	window.close();				
}

</script>
</body>
</HTML>