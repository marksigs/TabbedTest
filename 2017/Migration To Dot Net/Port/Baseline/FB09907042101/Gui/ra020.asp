<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      ra020.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Credit Check Summary Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MDC		13/03/2000	Original
AY		03/04/00	New top menu/scScreenFunctions change
					Errors in screen format
MDC		13/04/2000	Amended to run as a popup
MDC		14/04/2000	Enable buttons after successful credit check
BG		17/05/00	SYS0752 Removed Tooltips
MDC		18/05/00	SYS0755 Remove UnitID and UserID nodes
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MDC		05/09/2002	BMIDS00336	CCWP1 Credit Check & Bureau Download
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		09/10/2002	BMIDS00556	Removed Validate_Int()
MC		19/04/2004	BMIDS517	ra025,ra030 popup dialog height/sizes change
								white space padded to the title text.
INR		21/05/2004  BMIDS744	Third party data changes. Add Third Party Data Outcome Code.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Credit Check Summary  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4 style="VISIBILITY: hidden">

<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 235px; WIDTH: 580px; POSITION: ABSOLUTE" class="msgGroup">

	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Application Data</strong>
	</span>

	<div style="TOP: 24px; LEFT: 38px; HEIGHT: 160px; WIDTH: 528px; POSITION: ABSOLUTE" class="msgGroupLight">

		<span style="TOP: 15px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Application Number
			<span style="TOP: -3px; LEFT: 300px; POSITION: ABSOLUTE">
				<input id="txtAppNumber" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 40px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Customer Name
			<span style="TOP: -3px; LEFT: 300px; POSITION: ABSOLUTE">
				<input id="txtName" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 65px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Owner of Current Application Credit Check
			<span style="TOP: -3px; LEFT: 300px; POSITION: ABSOLUTE">
				<input id="txtUserID" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 90px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Date and Time of Current Application Credit Check
			<span style="TOP: -3px; LEFT: 300px; POSITION: ABSOLUTE">
				<input id="txtDateTime" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>
	</div>

	<div style="TOP: 150px; LEFT: 38px; HEIGHT: 70px; WIDTH: 528px; POSITION: ABSOLUTE" class="msgGroupLight">

		<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Score</strong>
		</span>

		<span style="TOP: 19px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Scorecard ID
			<span style="TOP: -3px; LEFT: 155px; POSITION: ABSOLUTE">
				<input id="txtScorecardID" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 19px; LEFT: 270px; POSITION: ABSOLUTE" class="msgLabel">
			Score
			<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
				<input id="txtScore" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 44px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Third Party Data Outcome Code
			<span style="TOP: -3px; LEFT: 155px; POSITION: ABSOLUTE">
				<input id="txtTPDOutcomeCode" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>
	</div>
</div>

<div id="divBackground2" style="TOP: 255px; LEFT: 10px; HEIGHT: 490px; WIDTH: 580px; POSITION: ABSOLUTE" class="msgGroup">

	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Credit Check Summary</strong>
	</span>

	<div style="TOP: 24px; LEFT: 38px; HEIGHT: 125px; WIDTH: 528px; POSITION: ABSOLUTE" class="msgGroupLight">

		<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>General Information</strong>
		</span>

		<span style="TOP: 19px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Number of CCJs
			<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
				<input id="txtNumCCJs" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 19px; LEFT: 240px; POSITION: ABSOLUTE" class="msgLabel">
			Number of Non-M/O CAIS 8/9s
			<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
				<input id="txtNumNonMOCAIS" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 44px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Total Value (Fine) of CCJs
			<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
				<input id="txtValCCJs" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 44px; LEFT: 240px; POSITION: ABSOLUTE" class="msgLabel">
			Total Value (Fine) of Non-M/O CAIS 8/9s
			<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
				<input id="txtValNonMOCAIS" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 69px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Age of Most Recent CCJ
			<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
				<input id="txtAgeRecentCCJ" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 69px; LEFT: 240px; POSITION: ABSOLUTE" class="msgLabel">
			Age of Most Recent CAIS 8/9
			<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
				<input id="txtAgeRecentCAIS" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 94px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Number of Outstanding CCJs
			<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
				<input id="txtNumOSCCJ" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>
	</div>

	<div style="TOP: 155px; LEFT: 38px; HEIGHT: 100px; WIDTH: 528px; POSITION: ABSOLUTE" class="msgGroupLight">

		<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>All Active CAIS Accounts</strong>
		</span>

		<span style="TOP: 19px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Balance (Fine) Mortgages
			<span style="TOP: -3px; LEFT: 430px; POSITION: ABSOLUTE">
				<input id="txtBalanceMort" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 44px; LEFT: 8px; POSITION: ABSOLUTE" class="msgLabel">
			&nbspWorst Status (Fine) in the Last 6 Months on Mortgage Accounts
			<span style="TOP: -3px; LEFT: 432px; POSITION: ABSOLUTE">
				<input id="txtWorstStatusFine" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 69px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Time SInce Most Recent Mortgage Account Opened
			<span style="TOP: -3px; LEFT: 430px; POSITION: ABSOLUTE">
				<input id="txtTimeMortAC" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>
	</div>

	<div style="TOP: 260px; LEFT: 38px; HEIGHT: 75px; WIDTH: 528px; POSITION: ABSOLUTE" class="msgGroupLight">

		<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Own Company CAIS Account</strong>
		</span>

		<span style="TOP: 19px; LEFT: 8px; POSITION: ABSOLUTE" class="msgLabel">
			&nbspWorst Status (Course) in the Last 6 Months of Active Accounts
			<span style="TOP: -3px; LEFT: 432px; POSITION: ABSOLUTE">
				<input id="txtWorstStatusCourse" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 44px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Balance (Fine) on Active Accounts
			<span style="TOP: -3px; LEFT: 430px; POSITION: ABSOLUTE">
				<input id="txtBalanceAC" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>
	</div>

	<div style="TOP: 340px; LEFT: 38px; HEIGHT: 60px; WIDTH: 528px; POSITION: ABSOLUTE" class="msgGroupLight">

		<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Name Confirmed</strong>
		</span>

		<span style="TOP: 25px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Electoral Roll & CAIS
			<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
				<input id="txtNamesConfirmed" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>

		<span style="TOP: 25px; LEFT: 240px; POSITION: ABSOLUTE" class="msgLabel">
			Electoral Roll/PAF
			<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
				<input id="txtVotersRollPAF" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>
	</div>

	<div style="TOP: 405px; LEFT: 38px; HEIGHT: 45px; WIDTH: 528px; POSITION: ABSOLUTE" class="msgGroupLight">

		<span style="TOP: 15px; LEFT: 13px; POSITION: ABSOLUTE" class="msgLabel">
			CAIS Summary
			<span style="TOP: -8px; LEFT: 100px; POSITION: ABSOLUTE">
				<input id="btnCAISButton" value="" type="button" style="WIDTH: 26px; HEIGHT: 26px" class="msgDDButton">
			</span>
		</span>

		<% /* BMIDS00336 MDC 05/09/2002 */ %>
		<% /*			
		<span style="TOP: 15px; LEFT: 240px; POSITION: ABSOLUTE" class="msgLabel">
			New Credit Check
			<span style="TOP: -5px; LEFT: 100px; POSITION: ABSOLUTE">
				<input id="btnCreditCheck" value="Run" type="button" class="msgButton" style="WIDTH: 60px">
			</span>
		</span>
		*/ %>
		<span style="TOP: 15px; LEFT: 240px; POSITION: ABSOLUTE" class="msgLabel">
			View Full Bureau Results
			<span style="TOP: -5px; LEFT: 130px; POSITION: ABSOLUTE">
				<input id="btnView" value="View" type="button" class="msgButton" style="WIDTH: 60px">
			</span>
		</span>
		<% /* BMIDS00336 MDC 05/09/2002 - End */ %>

		<span style="TOP: 55px; LEFT: 10px; POSITION: ABSOLUTE">
			<input id="btnSubmit" value="OK" type="button" class="msgButton" style="WIDTH: 60px">
		</span>
		<span style="TOP: 55px; LEFT: 80px; POSITION: ABSOLUTE">
			<input id="btnCancel" value="Cancel" type="button" class="msgButton" style="WIDTH: 60px">
		</span>

	</div>

</div>


</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/ra020Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sUserID = "";
var m_sUnitID = "";

var m_nCurrentCustomerOrder = 1;
var CreditCheckXML = null;
var m_nPersonType = 0;
var m_sCreditCheckGUID = "";
var m_sCustomerName = "";
var m_sCustomerNumber = "";
var m_sCustomerVersion = "";
var scScreenFunctions;
var m_sRequestAttributes = "";
var m_nMaxCustomers = 1;
var m_BaseNonPopupWindow = null;
var m_sArgArray = null; <% /* BMIDS00336 MDC 05/09/2002 */ %>


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];	
	<% /* BMIDS00336 MDC 05/09/2002 
	var sArgArray = sArguments[4]; */ %>
	m_sArgArray = sArguments[4];
	<% /* BMIDS00336 MDC 05/09/2002 - End */ %>
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	<% /* BMIDS00336 MDC 05/09/2002 
	RetrieveContextData(sArgArray); */ %>
	RetrieveContextData();
	<% /* BMIDS00336 MDC 05/09/2002 - End */ %>
	
	Initialise();
	SetAllFieldsToReadOnly();
	
	window.returnValue = null;

	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function SetAllFieldsToReadOnly()
{
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtAppNumber");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtName");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtUserID");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtDateTime");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtScorecardID");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtNumCCJs");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtNumNonMOCAIS");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtValCCJs");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtValNonMOCAIS");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtAgeRecentCCJ");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtAgeRecentCAIS");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtNumOSCCJ");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtBalanceMort");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtWorstStatusFine");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTimeMortAC");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtWorstStatusCourse");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtBalanceAC");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtNamesConfirmed");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtVotersRollPAF");

}

<% /* BMIDS00336 MDC 05/09/2002 - Various amendments to this function */ %>
function RetrieveContextData()
{
var sCustomerNumber2 = "";
var sCustomerRoleType2 = "";

	m_sApplicationNumber = m_sArgArray[0];
	m_sApplicationFactFindNumber = m_sArgArray[1];
	m_sUserID = m_sArgArray[8];
	m_sUnitID = m_sArgArray[9];
	sCustomerNumber2 = m_sArgArray[4];
	sCustomerRoleType2 = m_sArgArray[11];
	m_sRequestAttributes = m_sArgArray[10];
	
	if (sCustomerNumber2 == "" || sCustomerRoleType2 == "2")
		m_nMaxCustomers = 1;
	else
		m_nMaxCustomers = 2;
		
}

function GetCreditCheckData()
{
var bSuccess = false;

	<% /* Retrieve Credit Check data */ %>
	CreditCheckXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	with (CreditCheckXML)
	{
		<% /* Create request block */ %>
  		CreateRequestTagFromArray(m_sRequestAttributes, "SEARCH");

		CreateActiveTag("SEARCH");
		CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		CreateTag("CUSTOMERORDER", m_nCurrentCustomerOrder);
		// 		RunASP(document, "GetCreditCheckDetails.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					RunASP(document, "GetCreditCheckDetails.asp");
				break;
			default: // Error
				SetErrorResponse();
			}

		if (!IsResponseOK())
		{
			frmScreen.btnCAISButton.disabled = true;
			frmScreen.btnSubmit.disabled = true;
		}
		else
		{
			frmScreen.btnCAISButton.disabled = false;
			frmScreen.btnSubmit.disabled = false;
			bSuccess = true;
		}
		
		m_nPersonType = GetTagText("PERSONTYPE");
		m_sCreditCheckGUID = GetTagText("CREDITCHECKGUID");
		m_sCustomerNumber = GetTagText("CUSTOMERNUMBER");
		m_sCustomerVersion = GetTagText("CUSTOMERVERSIONNUMBER");
	}

	return bSuccess;
}

function PopulateScreen()
{
	with (frmScreen)
	{
		<% /* Application Number */ %>
		txtAppNumber.value = m_sApplicationNumber;
		<% /* Customer Name */ %>
		m_sCustomerName = CreditCheckXML.GetTagText("CUSTOMERNAME");
		txtName.value = m_sCustomerName;
		<% /* Application Data */ %>
		txtUserID.value = CreditCheckXML.GetTagText("USERID");
		txtDateTime.value = CreditCheckXML.GetTagText("DATETIME");
		txtScorecardID.value = CreditCheckXML.GetTagText("SCORECARDID1");
		txtScore.value = CreditCheckXML.GetTagText("SCORE1");
		<% /* BMIDS744 */ %>
		txtTPDOutcomeCode.value = CreditCheckXML.GetTagText("TPDOUTCOMECODE");
		<% /* Customer Specific Data */ %>
		txtNumCCJs.value = CreditCheckXML.GetTagText("TOTALCCJS");
		txtValCCJs.value = CreditCheckXML.GetTagText("TOTALVALUECCJS");
		txtAgeRecentCCJ.value = CreditCheckXML.GetTagText("AGEOFLASTCCJ");
		txtNumOSCCJ.value = CreditCheckXML.GetTagText("TOTALOUTSTANDINGCCJ");
		txtNumNonMOCAIS.value = CreditCheckXML.GetTagText("TOTALCAIS89");
		txtValNonMOCAIS.value = CreditCheckXML.GetTagText("TOTALVALUECAIS89");
		txtAgeRecentCAIS.value = CreditCheckXML.GetTagText("AGEOFLASTCAIS89");
		txtBalanceMort.value = CreditCheckXML.GetTagText("TOTALVALUEACTIVECAISINCMORT");
		txtWorstStatusFine.value = CreditCheckXML.GetTagText("WORSTSTATUSMORTGAGE6M");
		txtTimeMortAC.value = CreditCheckXML.GetTagText("MORTGAGEACCOUNTTIME");
		txtWorstStatusCourse.value = CreditCheckXML.GetTagText("WORSTSTATUSOC");
		txtBalanceAC.value = CreditCheckXML.GetTagText("TOTALVALUEOC");
		txtNamesConfirmed.value = CreditCheckXML.GetTagText("NAMESCONFIRMED");
		txtVotersRollPAF.value = CreditCheckXML.GetTagText("VOTERSROLLPAF");
	}

}


function Initialise()
{
	if (GetCreditCheckData())
	{
		scScreenFunctions.ShowCollection(frmScreen);
		PopulateScreen();
	}
	else
		scScreenFunctions.ShowCollection(frmScreen);
	
}

function frmScreen.btnSubmit.onclick()
{
	CreditCheckXML = null;
	
	if (m_nCurrentCustomerOrder < m_nMaxCustomers)
	{
		m_nCurrentCustomerOrder = m_nCurrentCustomerOrder + 1;
		Initialise();
	}
	else
		window.close();
}

function frmScreen.btnCancel.onclick()
{
	CreditCheckXML = null;
	
	if (m_nCurrentCustomerOrder > 1)
	{
		m_nCurrentCustomerOrder = m_nCurrentCustomerOrder - 1;
		Initialise();
	}
	else
		window.close();
}

function frmScreen.btnCAISButton.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array(4);

	ArrayArguments[0] = m_sCustomerNumber;
	ArrayArguments[1] = m_sCustomerVersion;
	ArrayArguments[2] = m_nPersonType;
	ArrayArguments[3] = m_sCreditCheckGUID;	
	ArrayArguments[4] = m_sCustomerName;	

	sReturn = scScreenFunctions.DisplayPopup(window, document, "ra025.asp", ArrayArguments, 378, 336);

}

<% /* BMIDS00336 MDC 05/09/2002 */ %>
function frmScreen.btnView.onclick()
{
	scScreenFunctions.DisplayPopup(window, document, "RA030.asp", m_sArgArray, 630, 480) ;
}

<% /*
function frmScreen.btnCreditCheck.onclick()
{

var bContinue;

	bContinue = window.confirm("A new credit check will now be run. Please note that this process may take up to 30 seconds to complete. Click OK to continue. Click Cancel to stop.");

	if (bContinue)
	{
		CreditCheckXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		with (CreditCheckXML)
		{
  			CreateRequestTagFromArray(m_sRequestAttributes, "UPDATE");
			CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
			RunASP(document, "RunCreditCheck.asp");
			if (IsResponseOK())
			{
				Initialise();
				alert("The credit check completed successfully.");
			}
		}	
	}
	else
		alert("Credit Check cancelled.");
}
*/ %>
<% /* BMIDS00336 MDC 05/09/2002 - End */ %>

-->
</script>
</body>
</html>




