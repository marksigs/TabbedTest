<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      cr040.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Customer Existing Business screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		03/12/99	SC017: Validate working hours on locking an application.
					Renamed IsApplicationLocked to LockApplication.
AY		04/02/00	Optimised version	
AY		11/02/00	Change to msgButtons button types
AY		15/02/00	SYS0179 - Record Not Found errors handled within the GUI
					SYS0058 - Stay on this screen in the event of a lock failure
SR		10/03/00	modified method btnSubmit.OnClick - Added Tags CUSTOMERROLETYPE
					and CUSTOMERORDER to Request.
AY		10/03/00	SYS0242 - Introduction of StageNumber context field
AY		14/03/00	SYS0050 - Clear customer name from frame on cancel
AY		24/03/00	SYS0558 - Flags introduced to prevent processing being run twice
SR		29/03/00	SYS0015 - In case of errors due to lock already existing, alert
					the user and navigate to CR010 or CR030 based on user's response  
AY		29/03/00	New top menu/scScreenFunctions change
AY		05/04/00	SYS0597 - Select not working in read only mode
AY		06/04/00	SYS0569 - New context field for stage information - names are
					no longer hard coded
IW		05/05/00	SYS0320 - Enabling/Disabling of Combo box fixed				
MC		16/05/00	SYS0020 - List Box to include the scrolling control	
IW		23/05/00	SYS0774 - DistributionChannelId Now ChannelID to match back end.
JLD		14/12/00	SYS1709 - create the case activity along with the application.
JLD		22/12/00	SYS1712 - Call OmigaTmBO.CreateApplication instead of omAPP.CreateApplication.
JLD		15/01/01	SYS1808 - Change to context variables to do with application stage.
BG		22/01/01	SYS1860 - Added Summary button to display cr042.asp, passing StageName through.
CL		25/01/01	SYS1756 - Adjusted to pass back to Contact History (CR025)
APS		02/01/03	SYS1993 - Save the ApplicationPriority selection		
CL		05/03/01	SYS1920 Read only functionality added
GD		23/04/01	SYS2165 Set Combo Application Priority to NORMAL - type 20, by default	
SR		13/06/01	SYS2362 added two new columns 'Type' and 'Relationship' to the listbox
JLD		25/03/02	SYS4318 Route to CR044
JLD		26/03/02	SYS4320 set context on route to cr044
SA		17/05/02	SYS4636 Default TypeOfBuyer to Concurrent if it's a Further Advance case
LD		23/05/02	SYS4727 Use cached versions of frame functions
CL		29/05/02	SYS4725 - Set select button to disabled on first click.
STB		31/05/02	SYS4451 Summary button is initially disabled and if no existing business is selected.
JLD		10/06/02	sys4838	Summary button should always be enabled if there is any existing business, regardless of business type.
DS		18/06/02	SYS4769 Do not allow a further advance to be created on an account where the customer is guarantor.
SG		19/06/02	SYS4870 Remove 'First Time' from Type of Buyer combo.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMids History:

Prog	Date		Description
MDC		16/05/2002	BMIDS00004 - BM073 Versions of employment
MDC		07/06/2002	BMIDS00036 - Error creating new application
GHun	06/08/2002	BMIDS00006 - CAWP1 BM054 Customer Account Download
GHun	10/09/2002	BMIDS00425 - Set context parameter OtherSystemAccountNumber before routing to cr030
GHun	11/09/2002	BMIDS00434 - Add a warning message for Further Advances
GHun	16/09/2002	BMIDS00459 - Updated further advance validation
DPF		02/10/2002	BMIDS00046 - CPWP1 - BM037, add two new fields to list box for Drawdown & Overpayments
								 and check on 'TypeOfMortgage' Combo, on the onchange event.
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		09/10/2002	BMIDS00556	Removed Validation_Init()
DPF		04/11/2002  BMIDS00809  Amended message on alert box called upon checking 'TypeOfMortgage' combo value
GHun	05/11/2002	BMIDS00823	Add LoanClassType to table and change cboTypeOfMortgage.onchange
MO		13/11/2002	BMIDS00723	Made changes after the font sizes were changed
GHun	18/11/2002	BMIDS00980  Get both AccountNumber and ApplicationNumber so FurtherAdvances can work
GHun	18/11/2002	BMIDS00982	ImportCustomersIntoApplication should only be called if an account is selected
GHun	20/11/2002	BMIDS01014	Set OtherSystemAccountNumber in context
GHun	05/12/2002	BM0091		Clear account number for new loans
GHun	06/12/2002	BM0091		Account number should only be set for validation types 'M' and either 'F' or 'T'
GHun	23/10/2003	BMIDS624	Check if rates have changed when selecting an app
JD		28/07/2004	BMIDS749	Removed check for changed rates.
HMA     12/08/2004  BMIDS836    Set the NewToeCustomerInd when creating a new customer version.
HMA     17/08/2004  BMIDS845    Ensure that only one new customer version is created when 'Account' is selected.
GHun	05/10/2004	BMIDS907	Display a status message when calling admin system
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History :

Prog	Date		Description
AM		23/09/2005	MAR57 Added Regulation Indicator Combo 
PSC		11/10/2005	MAR57 Changes to disable summary button and ImportAccountsIntoApplication
PSC		28/10/2005	MAR300 Amend to set idExistingApplication on selecting existing application
KRW     11/11/2005  MAR438 Added check on m_sOtherSystemCustomerNumber.length before populating OTHERSYSTEMCUSTOMERNUMBER
					avoiding error on null value being passed to search (causing no cases being displayed)
PJO     13/12/2005  Add Direct/Indirect Combo		
JD		14/03/2006	MAR1426 reset typeofbuyer combo on change of typeofmortgage combo			
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM History :

Prog	Date		Description
PE		26/06/2006	EP608 - Modify the population of Type of Buyer.
PE		04/07/2006	EP608 - Modify the population of Type of Buyer.
PE		07/07/2006	EP608 - Modify the population of Type of Buyer.
AShaw	26/10/2006	EP2_8 - Modify the cboTypeOfMortgage.onchange method.
AShaw	15/11/2006	EP2_8 - Comment out Alert code until Optimus issue sorted.
		11/12/2006		  - Reinstated.
AShaw	15/11/2006	EP2_55 - Alter GetTypeOfMortgageList and btnSubmit.OnClick
		12/12/2006		  - Reinstated.
PE		30/11/2006	EP2_52 - Modified re-entry variable m_bIsSubmit
PE		01/12/2006	EP2_272 - Modified GetTypeOfBuyerList
AShaw	31/01/2007	EP2_1116 - Save Direct/Indirect flag if Account.
PSC		31/01/2007	EP2_1146 - Set account number in context correctly
AShaw	01/02/2007	EP2_213 - Fix logic for Add Borrowing alert.
AShaw	02/02/2007	EP2_1104 - Alter Menu for Type of Mortgage App to show longer values.
PSC		06/02/2007	EP2_1242 Amend population of type of buyer combo based on application type
AShaw	08/02/2007	EP2_1116 - Save Direct/Indirect flag if Account, alternate path.
AShaw	08/02/2007	EP2_213 - Logic change for Additional Borrowing alert.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/
%>

<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>

</SCRIPT>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets */ %>
<% /* CORE UPGRADE 702 <OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT> */ %>
<OBJECT data=scTable.htm height=1 id=scTable 
style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet 
VIEWASTEXT></OBJECT>
	
<!-- List Scroll object -->
<span style="LEFT: 310px; POSITION: absolute; TOP: 250px">
<OBJECT data=scListScroll.htm id=scScrollPlus 
style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 
type=text/x-scriptlet VIEWASTEXT></OBJECT>
</span> 

<% /* Specify Forms Here */ %>
<form id="frmSelect" method="post" action="cr030.asp" STYLE="DISPLAY: none"></form>
<!--<form id="frmCancel" method="post" action="cr025.asp" STYLE="DISPLAY: none"></form>-->
<!--GD SYS1752 21/02/2001 route to cr020 on cancel, and not cr025-->
<form id="frmCancel" method="post" action="cr020.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCR010" method="post" action="cr010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCR042" method="post" action="cr042.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCR044" method="post" action="cr044.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen">
<div style="HEIGHT: 222px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
<% /* DPF 03/10/2002 - CPWP1 - Add Drawdown and Overpayments columns */ %>
	<span id="spnExistingBusiness" style="LEFT: 4px; POSITION: absolute; TOP: 4px">
		<table id="tblExistingBusiness" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="10%" class="TableHead" >Type</td>			<td width="11%" class="TableHead">Business Type</td>	<td width="11%" class="TableHead">App./Acc Number</td>	<td width="11%" class="TableHead">Date Created</td>		<td width="11%" class="TableHead">Amount</td>			<td width="10%" class="TableHead">Relation-<br>ship</td><td width="10%" class="TableHead">Current Stage</td>	<td width="9%" class="TableHead">Draw-<br>down</td>		<td width="9%" class="TableHead">Over-<br>payment</td>	<td width="8%" class="TableHead">Loan Class Type</td></tr>
			<tr id="row01">		<td width="10%" class="TableTopLeft">&nbsp;</td>		<td width="11%" class="TableTopCenter">&nbsp;</td>		<td width="11%" class="TableTopCenter">&nbsp;</td>		<td width="11%" class="TableTopCenter">&nbsp;</td>		<td width="11%" class="TableTopCenter">&nbsp;</td>		<td width="10%" class="TableTopCenter">&nbsp;</td>		<td width="10%" class="TableTopCenter">&nbsp;</td>		<td width="9%" class="TableTopCenter">&nbsp;</td>		<td width="9%" class="TableTopCenter">&nbsp;</td>		<td width="8%" class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="10%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="8%" class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="10%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="8%" class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="10%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="8%" class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="10%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="8%" class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="10%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="8%" class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="10%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="8%" class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="10%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="8%" class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="10%" class="TableLeft">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="11%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="9%" class="TableCenter">&nbsp;</td>			<td width="8%" class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="10%" class="TableBottomLeft">&nbsp;</td>		<td width="11%" class="TableBottomCenter">&nbsp;</td>	<td width="11%" class="TableBottomCenter">&nbsp;</td>	<td width="11%" class="TableBottomCenter">&nbsp;</td>	<td width="11%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="9%" class="TableBottomCenter">&nbsp;</td>	<td width="9%" class="TableBottomCenter">&nbsp;</td>	<td width="8%" class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>

	<span id="spnButtons" style="LEFT: 4px; POSITION: absolute; TOP: 190px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
			<input id="btnSelect" value="Select" type="button" style="WIDTH: 60px" class="msgButton" disabled LANGUAGE=javascript onclick="return btnSelect_onclick()">
		</span>
			
		<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">
			<input id="btnNewApplication" value="New Application" type="button" style="WIDTH: 100px" class="msgButton">
		</span>
		
		<span style="LEFT: 167px; POSITION: absolute; TOP: 0px">
			<input id="btnSummary" value="Summary" type="button" style="WIDTH: 70px" class="msgButton">
		</span>
	</span>
</div>

<div id="divApplicationType" style="HEIGHT: 60px; LEFT: 10px; POSITION: absolute; TOP: 272px; VISIBILITY: hidden; WIDTH: 604px" class="msgGroup">
<!--GD DEBUG<div id="divApplicationType" style="TOP: 262px; LEFT: 10px; HEIGHT: 60px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">-->

	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Type of Mortgage Application
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<select id="cboTypeOfMortgage" style="WIDTH: 180px" class="msgCombo" ></select>
		</span>
	</span>
<%	/* JLD Added Type of buyer */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Type of Buyer
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<select id="cboTypeOfBuyer" style="WIDTH: 180px" class="msgCombo" ></select>
		</span>
	</span>
<%	/* PJO 13/12/2005 MAR850 Added Direct / Indirect Combo */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 58px" class="msgLabel">
		Direct/Indirect
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<select id="cboDirectIndirect" style="WIDTH: 180px" class="msgCombo" ></select>
		</span>
	</span>
<%	/* GD Added Application Priority Combo*/ %>
	<span style="LEFT: 350px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Application Priority
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<select id="cboApplicationPriority" style="WIDTH: 150px" class="msgCombo" ></select>
		</span>
	</span>
	<%/* AM MARS Added Regulation Indicator Combo */%>
	<span style="LEFT: 350px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Regulation Indicator
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<select id="cboRegulationIndicator" style="WIDTH: 150px" class="msgCombo" ></select>
		</span>
	</span>
</div>

</form>
	
<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"> 
	<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->


<% /*Include the file to lock the application */ %>
<!-- #include FILE="Includes/LockApp.asp" --> 

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cr040Attribs.asp" -->

<script language="JScript">
<!--
var m_sCustomerNumber = null;
var m_sCustomerVersionNumber = null;
var m_sCompanyId = null;
var m_sDistributionChannelId = null;
<%//GD Added 05.02.01%>
var m_sOtherSystemCustomerNumber = null;

var m_sUserId = null;
var m_sUnitId = null;
var m_sReadOnly = null;
var m_sApplicationNumber = null;
var	m_sApplicationFactFindNumber = null;
var	m_sPackageNumber = null;
var	m_sTypeOfApplicationText = null;
var	m_sTypeOfApplicationValue = null;
var m_sStageName = null;
var m_sApplicationPriority = null;
var m_sStageId = null;
var m_sStageSeqNo = null;
var m_nRows;
var XMLTypeOfMortgage = null;
var XMLTypeOfBuyerNewLoan = null;
var XMLTypeOfBuyerRemortgage = null;
var XMLRedemptionStatus = null;	<% /* BMIDS00006 */ %>
var XMLDirectIndirect = null ; <% /* PJO MAR850 13/12/2005 */ %>
var m_bLegacyCustomer = null;
var ListXML;
var scScreenFunctions;
var m_bIsSubmit = false;
var m_blnReadOnly = false;
<% /* BMIDS00004 MDC 17/05/2002 */ %>
var m_intTotalApps = 0;
<% /* BMIDS00004 MDC 17/05/2002 - End */ %>
<%/* SG 19/06/02 SYS4870 */%>
var m_sMetaAction = null;
var m_intMortgageAccounts = 0;	<% /* BMIDS00459 */ %>
var m_sFlexibleLoanClassType = "";	<% /* BMIDS00823 */ %>
<%/*AM MAR57 added*/%>
var m_sRegulationIndicator = null ;
<% /* PSC 11/10/2005 MAR57 - Start */ %>
var m_bUseAdminGetAccountSummary = false;
var m_bUseAdminGetAccountCustomer = false;
<% /* PSC 11/10/2005 MAR57 - End */ %>
var m_sAccountNumber = null; //

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Cancel","Undo");
	// SYS4451 - Initialse the summary button to disabled.
	frmScreen.btnSummary.disabled = true;
	//frmScreen.btnSummary.disabled = true;    JLD SYS4320
	ShowMainButtons(sButtonList);
	DisableMainButton("Submit");
	DisableMainButton("Undo");

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	<% /* CORE UPGRADE 702 scScreenFunctions = new scScrFunctions.ScreenFunctionsObject(); */ %>
	FW030SetTitles("Customer Existing Business","CR040",scScreenFunctions);
	RetrieveContextData();
	scTable.initialise(tblExistingBusiness, 0, "");
	PopulateScreen();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CR040");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	SetScreenReadOnly();
	
	scScreenFunctions.HideCollection(divApplicationType);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function SetScreenReadOnly()
{
	if (m_sReadOnly == "1") frmScreen.btnNewApplication.disabled = true;	
}

function DeleteCustomerLock()
{
<%	// APS UNIT TEST REF 72 - Delete the Customer lock by calling
	// CustomerlockBO.Delete
%>	var blnReturn = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<% /* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject(); */ %>

	XML.CreateRequestTag(window,"DELETE");
	XML.CreateActiveTag("CUSTOMERLOCK");
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	// 	XML.RunASP(document,"DeleteCustomerLock.asp")
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"DeleteCustomerLock.asp")
			break;
		default: // Error
			XML.SetErrorResponse();
		}


	if(XML.IsResponseOK()) blnReturn = true;

	XML = null;

	return blnReturn;
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function frmScreen.btnNewApplication.onclick()
{
	scScreenFunctions.ShowCollection(divApplicationType);
<%	//divApplicationType.style.visibility = "visible";
%>	frmScreen.btnSelect.disabled = true;
	scTable.DisableTable();
	GetTypeOfMortgageList();
	GetTypeOfBuyerList();
	<%//GD added%>
	GetApplicationPriorityList();
	<%/*AM MARS added*/%>
	GetRegulationIndicator();
	<% /* PJO 13/12/2005 MAR850 */ %>
	GetDirectIndirect () ;
	EnableMainButton("Submit");
	EnableMainButton("Undo");
	frmScreen.btnNewApplication.disabled = true;
	frmScreen.cboTypeOfMortgage.focus();
}

function frmScreen.btnSelect.onclick()
{
	// CL AQR SYS4725
	frmScreen.btnSelect.disabled=true;
	//END SYS4725
	//GD PUT IN LOCK APP STUFF.....
	SetContextDataOnExistingBusiness();
	SetAccountNumberInContext();	<% /* BMIDS01014 */ %>
	
	// APS SYS1993 Centralise the LockApplication processing
	// see includes\LockApp.asp for details about the return values from the below call
	var iReturn = LockApplication(m_sReadOnly, m_sApplicationNumber, m_sApplicationFactFindNumber,
									m_sCustomerNumber, m_sDistributionChannelId);
	
	var blnSelectSubmit = false;	<% /* BMIDS624 GHun 23/10/2003 */ %>
	
	switch (iReturn) {
	
	//AQR SYS2005 Start: CP 16/04/2001 13:38:00
	
	// Remove all code that prevents cancelling of option to return to original screen
	// This also required customer lock deletion to be removed on the cancel option.
	
		case 0: //DeleteCustomerLock();
				//frmToCR010.submit();
	
				ResetContextData();
				break;
				
	//AQR SYS2005 Start: CP 16/04/2001 13:38:00								
	
		case 1:	if (m_sReadOnly != "1") DeleteCustomerLock();
				blnSelectSubmit = true;
				break;
		case 2: blnSelectSubmit = true;
				break;
		default:blnSelectSubmit = true;
	}
	
	<% /* BMIDS624 GHun 23/10/2003 */ %>
	if (blnSelectSubmit)
	{
		<% /* PSC 28/10/2005 MAR300 */ %>
		scScreenFunctions.SetContextParameter(window,"idExistingApplication","1");

		// JD BMIDS749  removed    CheckForChangedRates(m_sApplicationNumber, m_sApplicationFactFindNumber);
		frmSelect.submit();
	}
	<% /* BMIDS624 End */ %>
}

function ResetContextData()
{
<%	/*	AY 15/02/00 SYS0058 - do not clear customer number and customer version number */
%>	m_sApplicationNumber = null;
	m_sApplicationFactFindNumber = null;
	m_sPackageNumber = null;
	m_sTypeOfApplicationText = null;
	m_sTypeOfApplicationValue = null;
	m_sStageName = null;
	m_sStageId = null;
	m_sStageSeqNo = null;
<% /* SR 28/03/00 SYS0015 - Clear Customer data also from the context. CustomerName is not reset 
							in UpdateContext, so do it explicitly here	*/
%>
	m_sCustomerNumber = null;
	m_sCustomerVersionNumber = null;
	scScreenFunctions.SetContextParameter(window,"idCustomerName1",null);
	<% /* PSC 13/10/2005 MAR57 */ %>
	scScreenFunctions.SetContextParameter(window,"idCustomerCategory1",null);
	UpdateContext();
}

function spnExistingBusiness.onclick()
{	
	if (scTable.getRowSelectedId() != null)
	{
		
		<% /* PSC 11/10/2005 MAR57 */ %>
		var sBusinessType = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("BusinessTypeIndicator");
		
		<% /* PSC 11/10/2005 MAR57 - Start */ %>
		if (sBusinessType == "M" && !m_bUseAdminGetAccountSummary)
			frmScreen.btnSummary.disabled = true;
		else
			frmScreen.btnSummary.disabled = false; //JLD SYS4838 enable for any business type.
		<% /* PSC 11/10/2005 MAR57 - End */ %>

		<% /* PSC 11/10/2005 MAR57 */ %>
		if(sBusinessType == "A")
		{
			frmScreen.btnSelect.disabled = false;
			/* CORE UPGRADE 702 frmScreen.btnSummary.disabled = false; */
		}
		else 
		{
			frmScreen.btnSelect.disabled = true;
			// JLD SYS4320 frmScreen.btnSummary.disabled = true;
		}		
	}
	<% /* SYS4451 - Disable the summary button if nothing was selected.*/ %>	
	else
	{
		frmScreen.btnSummary.disabled = true;
	}
}

<%
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Function:		spnExistingBusiness.ondblclick()
	Description:	Handles the double click event from the span surrounding 
					the table. Using the principle of event bubbling we pick 
					up the onclick event after the table_onclick event in 
					the scTable.htm scriptlet. The double click event calls
					the Select process
					APS 06/09/99 - UNIT TEST REF 9
	Args Passed:	N/a
	Returns:		N/a
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
%>
function spnExistingBusiness.ondblclick()
{
	if (scTable.getRowSelectedId() != null)
		if (tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("BusinessTypeIndicator") == "A")
			frmScreen.btnSelect.onclick();
}

function btnSubmit.onclick()
{
	if(m_bIsSubmit) return;
	m_bIsSubmit = true;
	
	var bExit = false;
	var bFurtherAdvance = false;
		
	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject(); */
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	<% /* BMIDS00004 MDC 16/05/2002 - BM073 Versions of Employment */ %>
	<% /* If existing Omiga business, create new customer version (copying employment data) */ %>
	var bContinue = true;
	
	<% /* BMIDS836  Examine the Type of Mortgage and set the Transfer of Equity Indicator accordingly */ %>
	var bToEInd = 0;  
	var ValidationList = new Array(1);
	ValidationList[0] = "CC";   // Transfer Of Equity   
	
	var sTypeOfMortgage = frmScreen.cboTypeOfMortgage.value;
	if (sTypeOfMortgage != "" )
	{
		var combXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if(combXML.IsInComboValidationList(document,"TypeOfMortgage", sTypeOfMortgage , ValidationList)) 
			bToEInd = 1;
	}		
	
	<% /* BMIDS845 Check the type and do not create a new version here if the type is 'Account'
	               The new version is created during ImportCustomersIntoApplication */ %>
	               
	var sType = "";
	if (scTable.getRowSelectedId() != null)
		sType = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("Type");
		
	if(m_intTotalApps - m_intMortgageAccounts > 0)
	{
		if (sType != "Account")
		{
			<% /* Create new customer version */ %>
			/* CORE UPGRADE 702 var CustXML = new scXMLFunctions.XMLObject(); */
			var CustXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			CustXML.CreateRequestTag(window,null);
			CustXML.CreateActiveTag("SEARCH");
			CustXML.CreateActiveTag("CUSTOMER");
			CustXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
			CustXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
			CustXML.CreateTag("NEWTOECUSTOMERIND", bToEInd);  // BMIDS836
		
			// 		CustXML.RunASP(document, "CreateNewCustomerVersion.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
						CustXML.RunASP(document, "CreateNewCustomerVersion.asp");
					break;
				default: // Error
					CustXML.SetErrorResponse();
				}

			if(CustXML.IsResponseOK())
			{
				// Update customer version number
				<% /* BMIDS00036 MDC 07/06/2002 - Correct typo */ %>
				var CustVerTag = CustXML.SelectTag(null,"CUSTOMERKEY")
				<% /* BMIDS00036 MDC 07/06/2002 - End */ %>
				m_sCustomerVersionNumber = CustXML.GetTagText("CUSTOMERVERSIONNUMBER");
			}
			else
			{
				alert("Unable to create new customer version. Please contact Help Desk.");
				bContinue = false;
			}
		}
	}
	
	if(bContinue)
	{	
		<% /* BMIDS00434 Moved validation here from ApplicationManagerBO.ValidateCustomerToApplication
			  It is now only a warning and excludes Secure Personal Loans and Equity Purchase */ %>
		var iIndex = frmScreen.cboTypeOfMortgage.selectedIndex;
		if (m_bLegacyCustomer && scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,iIndex,"F") && scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,iIndex,"M"))
			if (!(scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,iIndex,"P") || scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,iIndex,"S")))
				alert("A customer may not be reordered within, assigned to, or removed from this type of mortgage application.");
		<% /* BMIDS00434 End */ %>
		
		//If listbox does not contain mortgage accounts
		<% /* BMIDS00982 */ %>
		
		if (sType != "Account")
		<% /* if ((m_intMortgageAccounts == 0) || (scTable.getRowSelectedId() == null)) */ %>
		<% /* BMIDS00982 End */ %>
		<% /* BMIDS00004 MDC 16/05/2002 - End */ %>
		{
			// EP2_55 - Add new Error messages for PSW/NP and TOE.
			if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"PSW"))
			{
				alert("You must select the existing Mortgage Account that requires the Product Switch.");
				m_bIsSubmit = false;
				return false;
			}

			if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"NP"))
			{
				alert("You must select the existing Mortgage Account that requires Porting.");
				m_bIsSubmit = false;
				return false;
			}

			if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"TOE"))
			{
				alert("You must select the existing Mortgage Account that requires the Transfer of Equity.");
				m_bIsSubmit = false;
				return false;
			}
			
			if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"CLI"))
			{
				alert("You must select the existing Mortgage Account that requires the Credit Limit Increase.");
				m_bIsSubmit = false;
				return false;
			}
			// EP2_55 end.
			<% /* BMIDS00006 */ %>
			<% /* PSC 11/10/2005 MAR57 */ %>
			if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"F"))
			{
				<% /* PSC 11/10/2005 MAR57 */ %>
				alert("You must select the existing Mortgage Account for Additional Borrowing applications.");
				m_bIsSubmit = false;
				return false;
			}
			<% /* BMIDS00006 End */ %>

			
			
			var reqTag = XML.CreateRequestTag(window,"CreateApplication");
			XML.CreateActiveTag("APPLICATION");
			XML.SetAttribute("TYPEOFAPPLICATION", frmScreen.cboTypeOfMortgage.value);
			//THIS WILL BE ABSENT ON FURTHER ADVANCE...
			//SYS4636 Further Advance is now defaulted to "concurrent"
			XML.SetAttribute("TYPEOFBUYER", frmScreen.cboTypeOfBuyer.value);
			
			// APS SYS1993 Save the ApplicationPriority selection
			XML.SetAttribute("APPLICATIONPRIORITYVALUE", frmScreen.cboApplicationPriority.value);
			
			<%/*AM MARS added*/%>
			XML.SetAttribute("REGULATIONINDICATOR", frmScreen.cboRegulationIndicator.value);
			
			<% // PJO 13/12/2005 MAR850 %>
			XML.ActiveTag = reqTag;
			XML.CreateActiveTag("APPLICATIONFACTFIND");
			XML.SetAttribute("DIRECTINDIRECTBUSINESS", frmScreen.cboDirectIndirect.value);
						
			XML.ActiveTag = reqTag;
			XML.CreateActiveTag("CUSTOMER");
			XML.SetAttribute("CUSTOMERROLETYPE", "1");
			XML.SetAttribute("CUSTOMERORDER", "1");
			XML.SetAttribute("CUSTOMERNUMBER", m_sCustomerNumber);
			XML.SetAttribute("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
					
			XML.ActiveTag = reqTag;
			XML.CreateActiveTag("CASEACTIVITY");
			XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
			XML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId","1"));

			// 			XML.RunASP(document, "OmigaTMBO.asp");
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
		<%	/*	XML.CreateTagList("CUSTOMERPACKAGEKEYS"); */
		%>
				XML.CreateTagList("APPLICATION");
				if(XML.SelectTagListItem(0))
				{
		<%			// APS UNITTEST REF 2 - Adding returned key fields into the context					
		%>			SetContextDataOnNewBusiness(XML);
					bExit = true;
				}
			} else
			{
				alert("Unable to create New Application. Please contact Help Desk");
			}

			XML = null;
			if(bExit) frmSelect.submit();
			else m_bIsSubmit = false;	
		
		} else // Account selected
		{
			//if this is a further advance and the customer is guarantor, display warning message
			if (tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("CustomerRoleType")!= 1 &&
				scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"F")) 
			{
				alert("The customer is a Guarantor on the selected account. A Further Advance application cannot be created.");
				m_bIsSubmit = false;
				return false;
			} else {
				<% /* BMIDS00006 No longer required
				var bAddAllCustomers = true;
				//If type of mortgage is not a further advance...
				
				if (!(scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"F")))
				{
						
					//Get first item in list (ie. most recent)
					var iNumApplicants = NumCustsInGetMostRecentAccount();
					if (iNumApplicants > 1)
					{
						if (!confirm("Add the customers from the mortgage account to the new application? OK = Yes, Cancel = No"))
						{
							bAddAllCustomers=false;
						}
					}
					
					bAddAllCustomers = true;
				} 
				//BMIDS00006 End */ %>
				
				//set the acct no. for the ImportAccountIntoApplication call to the selected account
				bFurtherAdvance=true;

				// BMIDS00006 GHun 
				var sAmount = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("Amount")
				var sStatus = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("Status")
				var ValidationList = new Array(2);
				ValidationList[0] = "A";
				ValidationList[1] = "R";	// BMIDS00459 
				var RedeemXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				var blnRedeemed = RedeemXML.IsInComboValidationList(document, "RedemptionStatus", sStatus, ValidationList)
				var lDisable = false;  //EP2_55
				// BMIDS00459 if (blnRedeemed || (sAmount == 0) || (sAmount == "")) 
				if (blnRedeemed || (sAmount <= 0) || (sAmount == ""))	// BMIDS00459 
				{
					// EP2_55 - Add new Error messages for PSW,NP,CLI and TOE.
					if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"PSW"))
					{
						alert("Product Switch is not allowed on Selected Mortgage Account.");
						lDisable = true;
					}

					if ((lDisable == false) && scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"NP"))
					{
						alert("Product Porting is not allowed on Selected Mortgage Account.");
						lDisable = true;
					}

					if ((lDisable == false) && scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"TOE"))
					{
						alert("Transfer of Equity is not allowed on Selected Mortgage Account.");
						lDisable = true;
					}
					
					if ((lDisable == false) && scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"CLI"))
					{
						alert("Credit Limit Increase is not allowed on Selected Mortgage Account.");
						lDisable = true;
					}
					
					if ((lDisable == false) && scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"F"))
					{
						alert("Further Advance is not allowed on Selected Mortgage Account.");
						lDisable = true;
					}
					
					if (lDisable == true)
					{			 
						m_bIsSubmit = false;
						DisableMainButton("Submit");
						EnableMainButton("Cancel");
						EnableMainButton("Undo");
						return false;
					}
					// EP2_55 end.
				}
				<% /* BMIDS00006 Call ImportCustomersIntoApplication */ %>
				//var reqTag = XML.CreateRequestTag(window,"CUSTOMER");
				//XML.SetAttribute("OPERATION", "ImportAccountsIntoApplication");
				<% /* PSC 11/10/2005 MAR57 - Start */ %>
				if (m_bUseAdminGetAccountCustomer)
				{
					var reqTag = XML.CreateRequestTag(window,"ImportCustomersIntoApplication");
					XML.CreateActiveTag("IMPORTCUSTOMERSINTOAPPLICATION");
					<% /* BMIDS00006 End */ %>
				
					XML.CreateActiveTag("CUSTOMER");
					XML.SetAttribute("CUSTOMERNUMBER", m_sCustomerNumber);
					XML.SetAttribute("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
					XML.SetAttribute("OTHERSYSTEMCUSTOMERNUMBER", m_sOtherSystemCustomerNumber);
				
					<% /* BMIDS00006 */ %>
					XML.ActiveTag = XML.ActiveTag.parentNode;
					XML.CreateActiveTag("ACCOUNT");
					<% /* BMIDS00006 End */ %>
				
					<% /* BMIDS00459 ACCOUNTNUMBER should always be sent
					if (bFurtherAdvance)
					{
					*/ %>
					<% /* BMIDS00980 AccountNumber is now kept under its own attribute
					XML.SetAttribute("ACCOUNTNUMBER", tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("ApplicationNumber"));
					*/ %>
					XML.SetAttribute("ACCOUNTNUMBER", tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("AccountNumber"));
					//}

					<% /* BMIDS00006 */ %>
	
					XML.ActiveTag = XML.ActiveTag.parentNode;
					XML.CreateActiveTag("APPLICATION");
					//XML.SetAttribute("ADDALLCUSTOMERS", bAddAllCustomers);
					<% /* BMIDS00006 End */ %>
				
					XML.SetAttribute("TYPEOFBUYER", frmScreen.cboTypeOfBuyer.value);
					XML.SetAttribute("TYPEOFAPPLICATION", frmScreen.cboTypeOfMortgage.value);
			
					// APS SYS1993 Save the ApplicationPriority selection
					XML.SetAttribute("APPLICATIONPRIORITYVALUE", frmScreen.cboApplicationPriority.value);
			
					<%/*AM MARS added*/%>
					XML.SetAttribute("REGULATIONINDICATOR", frmScreen.cboRegulationIndicator.value);
				
					XML.ActiveTag = reqTag;	
					XML.CreateActiveTag("CASEACTIVITY");
					XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
					XML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId","1"));
								
					// 				XML.RunASP(document, "OmigaTMBO.asp");
					// Added by automated update TW 09 Oct 2002 SYS5115
					switch (ScreenRules())
						{
						case 1: // Warning
						case 0: // OK
								<% /* BMIDS907 GHun Display a status message */ %>
								window.status = "Retrieving account customers ...";
								XML.RunASP(document, "OmigaTMBO.asp");
								window.status = "";
								<% /* BMIDS907 End */ %>
							break;
						default: // Error
							XML.SetErrorResponse();
						}
			
					//CHECK RESPONSE and set exit flag bExit = true
					if(XML.IsResponseOK())
					{
						XML.CreateTagList("APPLICATION");
						if(XML.SelectTagListItem(0))
						{
							SetContextDataOnNewBusiness(XML);
							bExit = true;
						}
					} else
					{
						alert("Unable to create New Application. Please contact Help Desk");
					}
					
					// EP2_1116 - Save DIRECTINDIRECTBUSINESS flag
					var affXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					affXML.CreateRequestTag(window);
					var xn = affXML.XMLDocument.documentElement;
					xn.setAttribute("CRUD_OP","UPDATE");
					xn.setAttribute("SCHEMA_NAME","omCRUD");
					xn.setAttribute("ENTITY_REF","APPLICATIONFACTFIND");
					var xe = affXML.XMLDocument.createElement("APPLICATIONFACTFIND");
					xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
					xe.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
					xe.setAttribute("DIRECTINDIRECTBUSINESS", frmScreen.cboDirectIndirect.value);
					xn.appendChild(xe);
					affXML.RunASP(document, "omCRUDIf.asp");
					
				}
				else
				{
					var reqTag = XML.CreateRequestTag(window,"CreateApplication");
					XML.CreateActiveTag("APPLICATION");
					XML.SetAttribute("TYPEOFAPPLICATION", frmScreen.cboTypeOfMortgage.value);
					XML.SetAttribute("TYPEOFBUYER", frmScreen.cboTypeOfBuyer.value);					
					XML.SetAttribute("APPLICATIONPRIORITYVALUE", frmScreen.cboApplicationPriority.value);
					XML.SetAttribute("REGULATIONINDICATOR", frmScreen.cboRegulationIndicator.value);
					// EP2_1116 - Save DIRECTINDIRECTBUSINESS flag
					XML.ActiveTag = reqTag;
					XML.CreateActiveTag("APPLICATIONFACTFIND");
					XML.SetAttribute("DIRECTINDIRECTBUSINESS", frmScreen.cboDirectIndirect.value);
					XML.ActiveTag = reqTag;
					XML.CreateActiveTag("CUSTOMER");
					XML.SetAttribute("CUSTOMERROLETYPE", "1");
					XML.SetAttribute("CUSTOMERORDER", "1");
					XML.SetAttribute("CUSTOMERNUMBER", m_sCustomerNumber);
					XML.SetAttribute("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
					XML.ActiveTag = reqTag;
					XML.CreateActiveTag("CASEACTIVITY");
					XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
					XML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId","1"));

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
						XML.CreateTagList("APPLICATION");
						if(XML.SelectTagListItem(0))
						{
							SetContextDataOnNewBusiness(XML);
							bExit = true;
						}
					} 
					else
					{
						alert("Unable to create New Application. Please contact Help Desk");
					}

					XML = null;
					if(bExit) frmSelect.submit();
					else m_bIsSubmit = false;	
				}
				<% /* PSC 11/10/2005 MAR57 - Start */ %>
			}
		}
	}
	
	XML = null;
	if(bExit) frmSelect.submit();

	m_bIsSubmit = false;
}

//GD Added SYS1752
function NumberOfMortgageAccountsInList()
{
	var iCount = 0;
	ListXML.ActiveTag = null;
	
	ListXML.CreateTagList("EXISTINGBUSINESS");

	for(var nLoop = 0;ListXML.SelectTagListItem(nLoop) != false;nLoop++)
	{
		if (ListXML.GetTagText("BUSINESSTYPEINDICATOR")=="M")
		{
			iCount++;
		}
	}

	return(iCount);
}

function NumCustsInGetMostRecentAccount()
{
	var sMax = "01/01/1800"; // Set low date to compare against
	var iNumApplicants = 0;
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("EXISTINGBUSINESS");
	for(var nLoop = 0;ListXML.SelectTagListItem(nLoop) != false; nLoop++)
	{
		
		if (ListXML.GetTagText("BUSINESSTYPEINDICATOR")=="M")
		{
			var sCurrent = ListXML.GetTagText("DATECREATED");
			var dtMax=scScreenFunctions.GetDateObjectFromString(sMax);
			var dtCurrent=scScreenFunctions.GetDateObjectFromString(sCurrent);
			
			if (dtCurrent > dtMax)
			{
				sMax = sCurrent;
				var iNumApplicants = ListXML.GetTagText("NUMBEROFAPPLICANTS");
			}
		}
	}
	
	return(iNumApplicants);	
}


function btnCancel.onclick()
{
<%	// APS UNIT TEST REF 22 - Possiblly change the metaction if it was previously 
	// CreateNewCustomerForNewApplication because we want to see the customer when we
	// route back to CR025
	// APS UNIT TEST REF 22 - Now routes back to CR025
	// AY 14/03/00 SYS0050 - Clear CustomerName1
%>	scScreenFunctions.SetContextParameter(window,"idMetaAction", "CreateExistingCustomerForNewApplication");
	scScreenFunctions.SetContextParameter(window,"idCustomerName1",null);
	<% /* PSC 13/10/2005 MAR57 */ %>
	scScreenFunctions.SetContextParameter(window,"idCustomerCategory1",null);

	frmCancel.submit();
}

function btnUndo.onclick()
{
	scScreenFunctions.HideCollection(divApplicationType);
<%	//divApplicationType.style.visibility = "hidden";
%>	frmScreen.btnNewApplication.disabled = false;
	scTable.EnableTable();
	DisableMainButton("Submit");
	DisableMainButton("Undo");
	spnExistingBusiness.onclick();
	frmScreen.btnNewApplication.focus();
}

function RetrieveContextData()
{
	m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber",null);
	m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber",null);
	m_sCompanyId = scScreenFunctions.GetContextParameter(window,"idCompanyId",null);
	m_sDistributionChannelId = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
	m_sUserId = scScreenFunctions.GetContextParameter(window,"idUserId",null);
	m_sUnitId = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	<%//GD Added 05.02.01 %>
	m_sOtherSystemCustomerNumber = scScreenFunctions.GetContextParameter(window,"idOtherSystemCustomerNumber",null);
	<%/* SG 19/06/02 SYS4870 */%>
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","");
	
}

// APS SYS1993 Lock Application processing
function SetContextDataOnExistingBusiness()
{
<%	// APS UNIT TEST REF 2 - Changed this function name to signify a difference between
	// setting the context up from an existing business
%>	m_sApplicationNumber = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("ApplicationNumber");
	m_sApplicationFactFindNumber = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("ApplicationFactFindNumber");
}

function SetContextDataOnNewBusiness(XML)
{
<%	// APS UNIT TEST REF 2 - Set up the context from the returned XML from 
	// CreatePackageAndApplication
%>	m_sApplicationNumber = XML.GetTagText("APPLICATIONNUMBER");
	m_sApplicationFactFindNumber = XML.GetTagText("APPLICATIONFACTFINDNUMBER");			
	m_sPackageNumber = XML.GetTagText("PACKAGENUMBER");
	m_sTypeOfApplicationText = frmScreen.cboTypeOfMortgage.options(frmScreen.cboTypeOfMortgage.selectedIndex).text;
	m_sTypeOfApplicationValue = frmScreen.cboTypeOfMortgage.value;
	m_sStageName = XML.GetTagText("STAGENAME");
	m_sStageId = XML.GetTagText("STAGENUMBER");
	m_sStageSeqNo = XML.GetTagText("STAGESEQUENCENO");
	// APS SYS1993
	m_sApplicationPriority = frmScreen.cboApplicationPriority.value;
	<%/* AM MARS 23/09/2005 */%>
	m_sRegulationIndicator = frmScreen.cboRegulationIndicator.value ;
	UpdateContext();
}

function UpdateContext()
{
<%	// APS UNIT TEST REF 2 - Update the context which is done only from this function
%>	scScreenFunctions.SetContextParameter(window,"idApplicationNumber",m_sApplicationNumber);
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber",m_sApplicationFactFindNumber);
	scScreenFunctions.SetContextParameter(window,"idPackageNumber",m_sPackageNumber);
	scScreenFunctions.SetContextParameter(window,"idMortgageApplicationDescription",m_sTypeOfApplicationText);
	scScreenFunctions.SetContextParameter(window,"idMortgageApplicationValue",m_sTypeOfApplicationValue);
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber",m_sCustomerNumber);		
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",m_sCustomerVersionNumber);
	scScreenFunctions.SetContextParameter(window,"idStageId",m_sStageId);
	scScreenFunctions.SetContextParameter(window,"idStageName",m_sStageName);
	scScreenFunctions.SetContextParameter(window,"idCurrentStageOrigSeqNo",m_sStageSeqNo);
	//JLD add activityId context
	scScreenFunctions.SetContextParameter(window,"idActivityId","10");
	// APS SYS1993
	scScreenFunctions.SetContextParameter(window,"idApplicationPriority",m_sApplicationPriority);
	<%/*AM MARS added*/%>
	scScreenFunctions.SetContextParameter(window,"idRegulationIndicator",m_sRegulationIndicator);
	
	<% /* BMIDS001014 */ %>
	SetAccountNumberInContext();
}

function PopulateScreen()
{
<%	/*	AY 15/02/00 SYS0179 - Handle ERRORNOTFOUND within the screen */
%>	frmScreen.btnNewApplication.focus();

	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject(); */
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<% // PJO 13/12/2005 MAR850 Added Direct/Indirect %>
	var sGroupList = new Array("TypeOfMortgage", "TypeOfBuyerNewLoan", "TypeOfBuyerRemortgage", "RedemptionStatus", "Direct/Indirect");

	if(XML.GetComboLists(document,sGroupList))
	{
		
		XMLTypeOfMortgage = XML.GetComboListXML("TypeOfMortgage");
		XMLTypeOfBuyerNewLoan = XML.GetComboListXML("TypeOfBuyerNewLoan");
		XMLTypeOfBuyerRemortgage = XML.GetComboListXML("TypeOfBuyerRemortgage");
		XMLRedemptionStatus = XML.GetComboListXML("RedemptionStatus");	<% /* BMIDS00006 */ %>
		XMLDirectIndect =  XML.GetComboListXML("Direct/Indirect");	<% /* PJO 13/12/2005 MAR850 */ %>
	}

	
	<%//GD Added 02/02/01
	//	scScreenFunctions.SetComboOnValidationType(frmScreen, "cboApplicationPriority", "D");
	//	Leave as is until PC and AP resolve the issue in CustomerRegGui.doc
	%>

	<% /* BMIDS00823 */ %>
	var GlobalParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
 	m_sFlexibleLoanClassType = GlobalParamXML.GetGlobalParameterString(document,"FlexibleLoanClassType");
	<% /* BMIDS00823 End */ %>
	
	<% /* PSC 11/10/2005 MAR57 - Start */ %>
	m_bUseAdminGetAccountCustomer = GlobalParamXML.GetGlobalParameterBoolean(document,"UseAdminGetAccountCustomer");
	m_bUseAdminGetAccountSummary  = GlobalParamXML.GetGlobalParameterBoolean(document,"UseAdminGetAccountSummary");	
	<% /* PSC 11/10/2005 MAR57 - End */ %>
	
	m_bLegacyCustomer = XML.GetGlobalParameterBoolean(document,"FindLegacyCustomer");
	XML = null;

	/* CORE UPGRADE 702 ListXML = new scXMLFunctions.XMLObject(); */
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	if(scScreenFunctions.GetContextParameter(window,"idMetaAction","CreateNewCustomerForNewApplication") != "CreateNewCustomerForNewApplication")
	{
		ListXML.CreateRequestTag(window,"SEARCH");
		ListXML.CreateActiveTag("CUSTOMER");
		ListXML.CreateTag("CUSTOMERNUMBER",m_sCustomerNumber);
		<% /* BMIDS00004 MDC 21/05/2002 - Search by customer, not customer version 
		ListXML.CreateTag("CUSTOMERVERSIONNUMBER",m_sCustomerVersionNumber);
		*/ %>
		
		<%//GD Added 
		// Added "OTHERSYSTEMCUSTOMERNUMBER"
		
		// 
		%>
		//GD
		if (m_sOtherSystemCustomerNumber.length > 0) // MAR438 KRW 11/11/2005 
			ListXML.CreateTag("OTHERSYSTEMCUSTOMERNUMBER", m_sOtherSystemCustomerNumber);
		
	
		<% /* BMIDS907 GHun Display a status message */ %>
		window.status = "Retrieving mortgage accounts and applications ...";
		ListXML.RunASP(document,"FindBusinessForCustomer.asp");
		window.status = "";
		<% /* BMIDS907 End */ %>
				
		var sErrorArray = new Array("RECORDNOTFOUND");
		var sResponseArray = ListXML.CheckResponse(sErrorArray);
		if(sResponseArray[0])
		{
			var bIsRecords = ShowList(0)
			ListXML.ActiveTag = null;
			//GD ListXML.CreateTagList("OMIGABUSINESS");
			ListXML.CreateTagList("EXISTINGBUSINESS");

			<% /* BMIDS00004 MDC 17/05/2002 - Get total number of apps/mort acc's */ %>
			m_intTotalApps = ListXML.ActiveTagList.length;
			<% /* BMIDS00004 MDC 17/05/2002 - End */ %>

			scScrollPlus.Initialise(ShowList,10,ListXML.ActiveTagList.length);
			
			m_intMortgageAccounts = NumberOfMortgageAccountsInList();	<% /* BMIDS00459 */ %>
			
			if(bIsRecords)
			{
				scTable.setRowSelected(1);
				spnExistingBusiness.onclick();
				if(!frmScreen.btnSelect.disabled) frmScreen.btnSelect.focus();
			}
		}
		
		<%//GD Added 05.02.01 %>
		else if(sResponseArray[1] == "RECORDNOTFOUND") alert("No Existing Business Found");	
	}
}

function GetTypeOfMortgageList()
{
	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject(); */
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	



<%
//GD SYS 1752 Added to allow validation attributes to be set...
%>
		var sGroups = new Array("TypeOfMortgage");
		var bSuccess = false;
		if(XML.GetComboLists(document,sGroups))
		{
			bSuccess = XML.PopulateCombo(document,frmScreen.cboTypeOfMortgage,"TypeOfMortgage",false);	
		}



<%
	// GD SYS 1752 REMOVED for above reason 
	//XML.PopulateComboFromXML(document,frmScreen.cboTypeOfMortgage,XMLTypeOfMortgage,false);
%>
	XML = null;
	
	if (bSuccess==true)
	{
		
		var sBusinessTypeIndicator = "A";
		if(scTable.getRowSelectedId() != null)
			sBusinessTypeIndicator = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("BusinessTypeIndicator");
			
		if(m_bLegacyCustomer && sBusinessTypeIndicator == "A")
		{
			
			var nOption = 0;
			while(frmScreen.cboTypeOfMortgage.options.length > 0 && nOption < frmScreen.cboTypeOfMortgage.options.length)
			{   // EP2_55 - Test for new values
				if(scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,nOption,"F")
					|| scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,nOption,"M")
					|| scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,nOption,"ABXA")
					|| scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,nOption,"CLIXA")
					|| scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,nOption,"PSWXA")
					|| scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,nOption,"NPXA")
					|| scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,nOption,"TOEXA")
					)
					frmScreen.cboTypeOfMortgage.options.remove(nOption);
				else nOption++;
			}
		}
	}
}


<%
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Function:		GetApplicationPriorityList()
	Description:	Populates the Application Priority combo. 
					GD 02/02/2001
	Args Passed:	N/a
	Returns:		N/a
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
%>
function GetApplicationPriorityList()
{

		/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject(); */
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var sGroups = new Array("ApplicationPriority");
		var bSuccess = false;
		if(XML.GetComboLists(document,sGroups))
		{
			bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboApplicationPriority,"ApplicationPriority",false);	
		}
		<%//GD added 23/04/2001
		//Set Combo to NORMAL - type 20 SYS2165
		%>
		frmScreen.cboApplicationPriority.value = 20;
}
<%
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Function:		GetRegulationIndicator()
	Description:	Populates the Regulation Indicator combo. 
					AM 23/09/2005
	Args Passed:	N/a
	Returns:		N/a
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
%>
function GetRegulationIndicator()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("RegulationIndicator");
	var bSuccess = false;
	if(XML.GetComboLists(document,sGroups))
	{
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboRegulationIndicator,"RegulationIndicator",false);	
		<% /* PSC 11/10/2005 MAR57 */ %>
		frmScreen.cboRegulationIndicator.value = XML.GetComboIdForValidation("RegulationIndicator", "R", null, document);
	}
}

<%
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Function:		GetDirectIndirect()
	Description:	Populates the DirectIndirect combo. PJO MAR850 13/12/2005
	Args Passed:	N/a
	Returns:		N/a
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
%>
function GetDirectIndirect()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("Direct/Indirect");
	var bSuccess = false;
	if(XML.GetComboLists(document,sGroups))
	{
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboDirectIndirect,"Direct/Indirect",false);	
		frmScreen.cboDirectIndirect.value = XML.GetComboIdForValidation("Direct/Indirect", "P", null, document);
	}
}


<%
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Function:		GetTypeOfBuyerList()
	Description:	Populates the Type Of Buyer combo based on the 
					choice of mortgage type.
					JLD 02/11/1999
	Args Passed:	N/a
	Returns:		N/a
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
%>
function GetTypeOfBuyerList()
{
	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject();*/
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

<%	// select options for TypeOfBuyer based on choice in TypeOfMortgage
%>	var selIndex = frmScreen.cboTypeOfMortgage.selectedIndex;
	if(scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,selIndex,"N") == true)
	{
<%		// New Loan
%>		scScreenFunctions.SetFieldState(frmScreen,"cboTypeOfBuyer","W")

<%/* Peter Edney - 04/07/2006 - EP608 */%>
//		/* CORE UPGRADE 702 XML.PopulateComboFromXML(document,frmScreen.cboTypeOfBuyer,XMLTypeOfBuyerNewLoan,false); */
//		<%/* SG 19/06/02 SYS4870 START */%>
//
//		//See if we have existing business items
//		ListXML.ActiveTag = null;
//		ListXML.CreateTagList("EXISTINGBUSINESS");
//		if (ListXML.SelectTagListItem(0) != false)
//		{
//			//There are existing business items.
//			
//			<%/* Need to do two things */%>
//			<%/* 1. Remove 'First Time' from TypeOfBuyer combo */%>
//			<%/* 2. Set default TypeOfBuyer to 'Subsequent' */%>
//			
//			<%/* Create XMLTypeOfBuyerNewLoan2, a copy of the original XML */%>
//			var XMLTypeOfBuyerNewLoan2 = new top.frames[1].document.all.scXMLFunctions.XMLObject();								
//			XMLTypeOfBuyerNewLoan2.LoadXML(XMLTypeOfBuyerNewLoan.xml);
//
//			<%/* Create variable to hold default combo value*/%>
//			var iDefault = 0;
//			
//			<%/* Iterate through XMLTypeOfBuyerNewLoan2...*/%>
//			XMLTypeOfBuyerNewLoan2.CreateTagList("LISTENTRY");
//			for(var nLoop = 0;XMLTypeOfBuyerNewLoan2.SelectTagListItem(nLoop) != false;nLoop++)
//			{
//				if (XMLTypeOfBuyerNewLoan2.GetTagText("VALIDATIONTYPE")=="F")
//				{
//					<%/* Remove 'First Time'*/%>
//					XMLTypeOfBuyerNewLoan2.RemoveActiveTag();
//				}
//				else if(XMLTypeOfBuyerNewLoan2.GetTagText("VALIDATIONTYPE")=="S")
//				{
//					<%/* Capture where 'Subsequent' is in the list*/%>
//					iDefault = nLoop - 1
//				}
//			}
//			
//			<%/* Populate combo using XMLTypeOfBuyerNewLoan2*/%>
//			XML.PopulateComboFromXML(document,frmScreen.cboTypeOfBuyer,XMLTypeOfBuyerNewLoan2.XMLDocument,false);
//			
//			<%/* Set the default combo value*/%>
//			frmScreen.cboTypeOfBuyer.selectedIndex = iDefault;
//						
//		}
//		else
//		{
//
//			//There are no existing business items

			<%/* Peter Edney - 26/06/2006 - EP608 */%>
			var sMortgageType = ""
						
			<% /* PSC 06/02/2007 EP2_1242 - Start */ %>
			if(scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,selIndex,"PRCH") == true)
				sMortgageType = "PRCH";
			else if(scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,selIndex,"FT") == true)
				sMortgageType = "FT";
			else if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,selIndex,"HM") == true)
				sMortgageType = "HM";
			<% /* PSC 06/02/2007 EP2_1242 - End */ %>

			<%/* Create XMLTypeOfBuyerNewLoan2, a copy of the original XML */%>
			var XMLTypeOfBuyerNewLoan2 = new top.frames[1].document.all.scXMLFunctions.XMLObject();								
			XMLTypeOfBuyerNewLoan2.LoadXML(XMLTypeOfBuyerNewLoan.xml);

			<%/* Iterate through XMLTypeOfBuyerNewLoan2...*/%>
			XMLTypeOfBuyerNewLoan2.CreateTagList("LISTENTRY");
			for(var nLoop = 0;XMLTypeOfBuyerNewLoan2.SelectTagListItem(nLoop) != false;nLoop++)
			{
				var bFound = false;
				var XMLTypeOfBuyerValidation = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XMLTypeOfBuyerValidation.LoadXML(XMLTypeOfBuyerNewLoan2.ActiveTagList[nLoop].xml);
				XMLTypeOfBuyerValidation.CreateTagList("VALIDATIONTYPE");

				switch(sMortgageType)				
				{
				<% /* PSC 06/02/2007 EP2_1242 - Start */ %>
				case "PRCH":
					for(var ValidationLoop = 0;XMLTypeOfBuyerValidation.SelectTagListItem(ValidationLoop) != false;ValidationLoop++)
					{
						if(XMLTypeOfBuyerValidation.ActiveTag.text=="F"  || XMLTypeOfBuyerValidation.ActiveTag.text=="S")
							bFound = true;
					}					
					if(!bFound)
					{
						XMLTypeOfBuyerNewLoan2.RemoveActiveTag();					
					}
					break						
				<% /* PSC 06/02/2007 EP2_1242 - End */ %>

				<%/* If Mortgage type is first time buyer, remove entries from type of buyer that arent of type F */%>
				case "FT":
					//if(XMLTypeOfBuyerNewLoan2.GetTagText("VALIDATIONTYPE")!="F")
					//	XMLTypeOfBuyerNewLoan2.RemoveActiveTag();
					for(var ValidationLoop = 0;XMLTypeOfBuyerValidation.SelectTagListItem(ValidationLoop) != false;ValidationLoop++)
					{
						if(XMLTypeOfBuyerValidation.ActiveTag.text=="F")
							bFound = true;
					}					
					if(!bFound)
					{
						XMLTypeOfBuyerNewLoan2.RemoveActiveTag();					
					}
					break						

				<%/* If Mortgage type is subsequent buyer, remove entries from type of buyer that aren't of type S */%>					
				case "HM":
					//if(XMLTypeOfBuyerNewLoan2.GetTagText("VALIDATIONTYPE")!="S")
					//	XMLTypeOfBuyerNewLoan2.RemoveActiveTag();
					for(var ValidationLoop = 0;XMLTypeOfBuyerValidation.SelectTagListItem(ValidationLoop) != false;ValidationLoop++)
					{
						if(XMLTypeOfBuyerValidation.ActiveTag.text=="S")
							bFound = true;
					}					
					if(!bFound)
					{
						XMLTypeOfBuyerNewLoan2.RemoveActiveTag();					
					}
					break						
				}				
			}
			
			XML.PopulateComboFromXML(document,frmScreen.cboTypeOfBuyer,XMLTypeOfBuyerNewLoan2.XMLDocument,false);		
		
//		}
//		<%/* SG 19/06/02 SYS4870 END */%>
	}
	else if(scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,selIndex,"F") == true)
	{
<%		//SYS4636 Further Advance
%>		scScreenFunctions.SetFieldState(frmScreen,"cboTypeOfBuyer","W")
		XML.PopulateComboFromXML(document,frmScreen.cboTypeOfBuyer,XMLTypeOfBuyerRemortgage,false);
	}
	else if(scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,selIndex,"R") == true)
	{
<%		//Remortgage
%>		scScreenFunctions.SetFieldState(frmScreen,"cboTypeOfBuyer","W")
		XML.PopulateComboFromXML(document,frmScreen.cboTypeOfBuyer,XMLTypeOfBuyerRemortgage,false);
	}
	else
	{
		frmScreen.cboTypeOfBuyer.selectedIndex = -1;
		scScreenFunctions.SetFieldState(frmScreen,"cboTypeOfBuyer","D")
	}

	XML = null;
}
<%
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Function:		frmScreen.cboTypeOfMortgage.onchange()
	Description:	on selection of a mortgage type, the options
					may change in the type of buyer.
					JLD 02/11/1999
	Args Passed:	N/a
	Returns:		N/a
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
%>

<% /* MAR1426 only have one onchange method
function frmScreen.cboTypeOfMortgage.onchange()
{
	GetTypeOfBuyerList();
}
*/ %>

function ShowList(nStart)
{
	var bIsRecords = false;
	var sTypeOfApplicationText;
	ListXML.ActiveTag = null;
	<%
	// GD Changed 22.02.01 SYS1752 
	// COM return tag name has changed from OMIGABUSINESS to EXISTINGBUSINESS
	// ListXML.CreateTagList("OMIGABUSINESS");
	%>
	ListXML.CreateTagList("EXISTINGBUSINESS");
	for(var nLoop = 0;ListXML.SelectTagListItem(nStart + nLoop) != false && nLoop < 10;nLoop++)
	{
		bIsRecords = true;
		<%
		// GD Changed 22.02.01 SYS1752 
		// COM return tag name has changed from APPLICATIONNUMBER
		%>
		<% /* BMIDS00980 */ %>
		var sDisplayNumber = ListXML.GetTagText("APPLICATIONNUMBERORACCOUNTNUMBER");
		var sApplicationNumber = ListXML.GetTagText("APPLICATIONNUMBER");
		var sAccountNumber  = ListXML.GetTagText("ACCOUNTNUMBER");
		<% /* BMIDS00980 End */ %>
		var sBusinessTypeIndicator = ListXML.GetTagText("BUSINESSTYPEINDICATOR");
		var sApplicationFactFindNumber = ListXML.GetTagText("APPLICATIONFACTFINDNUMBER");
		var sDateCreated = ListXML.GetTagText("DATECREATED");
		var sPackageNumber = ListXML.GetTagText("PACKAGENUMBER");
		var sAmount = ListXML.GetTagText("AMOUNT");
		var sStatus = ListXML.GetTagText("STATUS"); //stageName
		var sStageId = ListXML.GetTagText("STAGENUMBER");
		var sStageSeqNo = ListXML.GetTagText("CASESEQUENCENO"); //original stage seq no
		var sTypeOfApplicationValue = ListXML.GetTagText("TYPEOFAPPLICATION");
		var sCustomerRoleTypeDesc	= ListXML.GetTagAttribute("CUSTOMERROLETYPE", "TEXT");
		var sCustomerRoleType = ListXML.GetTagText("CUSTOMERROLETYPE");
		var sType	= ListXML.GetTagText("TYPE");
		//DPF 03/10/2002 - CPWP1 - Add Drawdown and Overpayments
		var sDrawDown = ListXML.GetTagText("DRAWDOWN");
		var sOverpayment = ListXML.GetTagText("OVERPAYMENTS");
		var sLoanClassType = ListXML.GetTagText("LOANCLASSTYPE"); <% /* BMIDS00823 */ %>
		var sTypeOfApplicationText;
		if (sBusinessTypeIndicator=="M")
		{
			sTypeOfApplicationText=ListXML.GetTagText("BUSINESSTYPE");
		} else
		{
			sTypeOfApplicationText = ListXML.GetTagAttribute("TYPEOFAPPLICATION","TEXT");
		}
		<% /* BMIDS00006 Get Status description for accounts */ %>
		if (sType == "Account")
		{
			var XMLStatus = XMLRedemptionStatus.selectSingleNode("LISTNAME/LISTENTRY[VALUEID[.='" + sStatus + "']]/VALUENAME");
			if (XMLStatus != null)
				var sStatusDesc = XMLStatus.text;
			else
				var sStatusDesc = "<unknown>";
		}
		else
			var sStatusDesc = sStatus;
		<% /* BMIDS00006 End */ %>
		
		<% /* BMIDS00823 Add LoanClassType */ %>
		//DPF 03/10/2002 - CPWP1 - Add Drawdown and Overpayments
		ShowRow(nLoop + 1,sType,sTypeOfApplicationText,sTypeOfApplicationValue,sDisplayNumber,sApplicationNumber,sAccountNumber,sDateCreated,
				sAmount, sCustomerRoleType, sCustomerRoleTypeDesc,sStatus,sStageId, sStageSeqNo,sBusinessTypeIndicator,sPackageNumber,sApplicationFactFindNumber,sStatusDesc, sDrawDown, sOverpayment, sLoanClassType);
	}
	return bIsRecords;
}

<% /* BMIDS00823 Add LoanClassType */ %>
//DPF 03/10/2002 - CPWP1 - Add Drawdown and Overpayments
function ShowRow(nIndex,sType,sTypeOfApplicationText,sTypeOfApplicationValue,sDisplayNumber,sApplicationNumber,sAccountNumber,sDateCreated,
					sAmount, sCustomerRoleType, sCustomerRoleTypeDesc,sStatus,sStageId,sStageSeqNo,sBusinessTypeIndicator,sPackageNumber,sApplicationFactFindNumber,sStatusDesc, sDrawDown, sOverpayment, sLoanClassType)
{
<%	// AY 09/09/1999 - Set the table fields
%>	scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(nIndex).cells(0),sType);
	scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(nIndex).cells(1),sTypeOfApplicationText);
	scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(nIndex).cells(2),sDisplayNumber);
	scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(nIndex).cells(3),sDateCreated);
	scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(nIndex).cells(4),sAmount);
	scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(nIndex).cells(5),sCustomerRoleTypeDesc);
	scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(nIndex).cells(6),sStatusDesc);
	//DPF 03/10/2002 - CPWP1 - Add Drawdown and Overpayments
	scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(nIndex).cells(7),sDrawDown);
	scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(nIndex).cells(8),sOverpayment);
	<% /* BMIDS00823 */ %>
	scScreenFunctions.SizeTextToField(tblExistingBusiness.rows(nIndex).cells(9),sLoanClassType);

	tblExistingBusiness.rows(nIndex).setAttribute("BusinessTypeIndicator", sBusinessTypeIndicator);
	tblExistingBusiness.rows(nIndex).setAttribute("PackageNumber", sPackageNumber);
	tblExistingBusiness.rows(nIndex).setAttribute("TypeOfApplicationText", sTypeOfApplicationText);
	tblExistingBusiness.rows(nIndex).setAttribute("TypeOfApplicationValue", sTypeOfApplicationValue);
	tblExistingBusiness.rows(nIndex).setAttribute("ApplicationNumber", sApplicationNumber);
	tblExistingBusiness.rows(nIndex).setAttribute("ApplicationFactFindNumber", sApplicationFactFindNumber);
	tblExistingBusiness.rows(nIndex).setAttribute("StageId", sStageId);
	tblExistingBusiness.rows(nIndex).setAttribute("StageSeqNo", sStageSeqNo);
	tblExistingBusiness.rows(nIndex).setAttribute("Status", sStatus);
	tblExistingBusiness.rows(nIndex).setAttribute("Type", sType);
	tblExistingBusiness.rows(nIndex).setAttribute("CustomerRoleType", sCustomerRoleType);
	//DPF 03/10/2002 - CPWP1 - Add Drawdown and Overpayments
	tblExistingBusiness.rows(nIndex).setAttribute("DrawDown", sDrawDown);
	tblExistingBusiness.rows(nIndex).setAttribute("Overpayment", sOverpayment);
	<% /* BMIDS00006 */ %>
	tblExistingBusiness.rows(nIndex).setAttribute("Amount", sAmount);
	<% /* BMIDS00823 */ %>
	tblExistingBusiness.rows(nIndex).setAttribute("LoanClassType", sLoanClassType);
	<% /* BMIDS00980 */ %>
	tblExistingBusiness.rows(nIndex).setAttribute("AccountNumber", sAccountNumber);
	//EP2_213 - Wasn't storing DateCreated
	tblExistingBusiness.rows(nIndex).setAttribute("DateCreated", sDateCreated);

}

function frmScreen.btnSummary.onclick()
{
	if(tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("Type") == "Account")  // JLD
	{
		//set the account number in context
		scScreenFunctions.SetContextParameter(window,"idOtherSystemAccountNumber",tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("AccountNumber"));
		frmToCR044.submit();
	}
	else
	{
		SetContextForCR042();
		frmToCR042.submit();
	}
	//GD 05.02.01
	//Code needs to be implemented here that allow routing to DisplayAccountSummary Screen
	//If an Account is selected, as opposed to an Application. - DONE for Optimus
}	
function SetContextForCR042()
{
	var iApplicationNumber = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("ApplicationNumber");
	var iApplicationFactFindNumber = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("ApplicationFactFindNumber");
	var iStageName = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("Status");
	
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber",iApplicationNumber);
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber",iApplicationFactFindNumber);
	scScreenFunctions.SetContextParameter(window,"idStageName",iStageName);
}

<% /* BMIDS01014 No longer required
// BMIDS00459
function SetIndicatorForCR030()
{
	var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	TempXML.CreateActiveTag("CUSTOMER");	

	if (m_intMortgageAccounts > 0)
	{
		TempXML.CreateTag("HASEXISTINGACCOUNTS", "true");
		
		// BMIDS00425 set idOtherSystemAccountNumber context paramter
		var iSelected = scTable.getRowSelectedId();
		if (iSelected != null)
			// BMIDS00980 Also need to pass AccountNumber for applications based on an account
			// if (tblExistingBusiness.rows(iSelected).getAttribute("BusinessTypeIndicator") == "M")
			//
			var sAccountNumber = tblExistingBusiness.rows(iSelected).getAttribute("AccountNumber");
			if (sAccountNumber.length > 0)
				TempXML.CreateTag("OTHERSYSTEMACCOUNTNUMBER", sAccountNumber);
		// BMIDS00425 End
	}
	else
		TempXML.CreateTag("HASEXISTINGACCOUNTS", "false");

	scScreenFunctions.SetContextParameter(window,"idXML",TempXML.XMLDocument.xml);
}
// BMIDS00459 End
// BMIDS01014 End */ %>

<% /* BMIDS01014 */ %>
function SetAccountNumberInContext()
{
	var iSelected = scTable.getRowSelectedId();
	if (iSelected != null)
	{
		var sAccountNumber = tblExistingBusiness.rows(iSelected).getAttribute("AccountNumber");
		if (sAccountNumber.length > 0)
		{
			<% /* BM0091 Only set the account number for validation types ('M' and ('F' or 'T')) */ %>	
			var iComboIndex = frmScreen.cboTypeOfMortgage.selectedIndex;
			if (iComboIndex > -1)
				<% /* PSC 31/01/2007 EP2_1146 - Start */ %>
				if ((scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,iComboIndex,"F") && 
				    scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,iComboIndex,"M")) ||
				    scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,iComboIndex,"PSW") ||
				    scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,iComboIndex,"TOE") ||
				    scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,iComboIndex,"NP"))
						scScreenFunctions.SetContextParameter(window,"idOtherSystemAccountNumber", sAccountNumber);
				<% /* PSC 31/01/2007 EP2_1146 - Start */ %>
			<% /* BM0091 End */ %>	
		}
	}
}
<% /* BMIDS01014 End */ %>

//DPF 03/10/2002 - CPWP1 (BM037)
//Function added to stop further advances being done when customer has
//drawdown or overpayments on an existing account
<% /* BMIDS00823  Only applies if LoanClassType is flexible */ %>
function frmScreen.cboTypeOfMortgage.onchange()
{
	if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"F"))
	{	
		var iRow = scTable.getRowSelectedId();
		if (iRow != null)
		{
			var sLoanClassType = tblExistingBusiness.rows(iRow).getAttribute("LoanClassType");
			var sDrawDown = tblExistingBusiness.rows(iRow).getAttribute("DrawDown");
			var sOverPayment = tblExistingBusiness.rows(iRow).getAttribute("OverPayment");
			// EP2_8 - Add new Global calls.
			var sRetention = CalculateRetainedAmount();

			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var sIsRestrictAddBorrowWhenDrawDown = XML.GetGlobalParameterBoolean(document, "RestrictAddBorrowWhenDrawdown");
			var sIsRestrictAddBorrowWhenOverPay = XML.GetGlobalParameterBoolean(document, "RestrictAddBorrowWhenOverPay");
			var sIsRestrictAddBorrowWhenRetention = XML.GetGlobalParameterBoolean(document, "RestrictAddBorrowWhenRetention");
			var sIsRestrictAddBorrowByExistTerm = XML.GetGlobalParameterBoolean(document, "RestrictAddBorrowByExistTerm");
			var sAccMinExistMortHeldInMonths = XML.GetGlobalParameterAmount(document, "AccMinExistMortHeldInMonths");
			var bcontinue = true; // Continue with next test (saves using nested IFs).

			if (sDrawDown.length == 0)
				sDrawDown = 0;
			if (sOverPayment.length == 0)
				sOverPayment = 0;
			if (sRetention.length == 0)
				sRetention = 0;

			//EP2_8 - New logic
			// Drawdown
			if ((sIsRestrictAddBorrowWhenDrawDown == 1)&& (sDrawDown > 0))
			{
				alert("You cannot take out Additional Borrowing as you have a drawdown outstanding.");
				frmScreen.cboTypeOfMortgage.value = 10;
				bcontinue = false;
			}
			// Overpay - // EP2_213 - bcontinue logic was incorrect.
			if ((bcontinue == true) && (sIsRestrictAddBorrowWhenOverPay == 1) && (sOverPayment > 0))
			{
				alert("You cannot take out Additional Borrowing as you have an overpayment outstanding.");
				frmScreen.cboTypeOfMortgage.value = 10;
				bcontinue = false;
			}
			// Retention - // EP2_213 - bcontinue logic was incorrect.
			if ((bcontinue == true) && (sIsRestrictAddBorrowWhenRetention == 1)&& (sRetention > 0))
			{
				alert("You cannot take out Additional Borrowing as you have an outstanding retention.");
				frmScreen.cboTypeOfMortgage.value = 10;
				bcontinue = false;
			}
			// Existing Term - // EP2_213 - bcontinue logic was incorrect.
			if ((bcontinue == true) && (sIsRestrictAddBorrowByExistTerm == 1))
			{
				// ONLY test for ACCOUNTS rows with F&M validation.		
				var sType = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("Type");
							
				if ((sType == "Account") && scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"F") && scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfMortgage,frmScreen.cboTypeOfMortgage.selectedIndex,"M"))
				{
				var sStartdate = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("DateCreated")
									
				// Calculate Months Account held for.
					var iMonthsAccountHeldFor = CalculateMonthsOnCover(sStartdate);
					
					// If les than min time allowed, raise alert.
					// EP2_213 - parseInt the Months on cover.
					if ((iMonthsAccountHeldFor != -1) && (iMonthsAccountHeldFor < parseInt(sAccMinExistMortHeldInMonths)))
					{
						var sAlertText = "You cannot take out Additional Borrowing as the account is less than " + sAccMinExistMortHeldInMonths + " months old.";
						alert(sAlertText);
						frmScreen.cboTypeOfMortgage.value = 10;
					}
	
				
				} // End (sType == "Account")
			
			} // End (bcontinue == true) && (sIsRestrictAddBorrowByExistTerm == 1)

		}
	}
	GetTypeOfBuyerList(); //MAR1426	
}

<% /* BMIDS624 GHun 22/10/2003 */ %>
function CheckForChangedRates(sApplicationNumber, sApplicationFactFindNumber)
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"");
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	XML.RunASP(document, "HaveRatesChanged.asp");
	if (XML.IsResponseOK())
	{
		if (XML.GetTagAttribute("QUOTATION", "RATESINCONSISTENT") == "1")
			alert("Rates are inconsistent on your current accepted or active quotation.  Please remodel.");	
	}

}

<%/* EP2_8 - New function.*/%>
function CalculateMonthsOnCover(lStartDate)
{
	if (lStartDate != "")
	{
		strDate = lStartDate.split("/"); 
		strYear = strDate[2].split(" ");
		dStartDate = new Date(strYear[0],strDate[1]-1,strDate[0])

		// Now compare and get differences
		var CurrentDate = scScreenFunctions.GetAppServerDate(true);
		NoOfMonths = GetDiffInMonths(dStartDate, CurrentDate);
	}
	return NoOfMonths;
	
}

<%/* EP2_8 - New function.*/%>
function CalculateRetainedAmount()
{
	// Get the AccountGUID
	var AccountGUID = "";
	var TotalRetention = 0;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	// Is Application row
	if (tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("TypeOfApplicationText") == "Application")
	{
		m_sApplicationNumber = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("ApplicationNumber");
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("APPLICATION");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.RunASP(document,"GetApplicationData.asp");
		if(XML.IsResponseOK())
		{
			XML.SelectTag(null,"APPLICATION");
			var AccountGUID = XML.GetTagText("ACCOUNTGUID");
		}
	}
	else // Is an Account Type
		// Awaiting spec change from RobinC 21Dec06
		AccountGUID = "";
/*		m_sAccountNumber = tblExistingBusiness.rows(scTable.getRowSelectedId()).getAttribute("AccountNumber");
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("ACCOUNT");
		XML.CreateTag("ACCOUNTNUMBER", m_sAccountNumber);
		XML.RunASP(document,"GetAccountData.asp");
		if(XML.IsResponseOK())
		{
			XML.SelectTag(null,"ACCOUNT");
			var AccountGUID = XML.GetTagText("ACCOUNTGUID");
		}
*/
	
	if (AccountGUID != "")
	{
		// Use AccountGUID from Application to retrieve the STARTDATE(S) from MortgageLoan table.
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window);
		var xn = XML.XMLDocument.documentElement;
		xn.setAttribute("CRUD_OP","READ");
		xn.setAttribute("SCHEMA_NAME","epsomCRUD");
		xn.setAttribute("ENTITY_REF","MORTGAGELOAN");
		var xe = XML.XMLDocument.createElement("MORTGAGELOAN");
		xe.setAttribute("ACCOUNTGUID", AccountGUID);
		xn.appendChild(xe);

		XML.RunASP(document, "omCRUDIf.asp");
		if (XML.IsResponseOK())
		{
			<% /* Loop round results, and set TotalRetention to longest period */%>
			XML.ActiveTag = null;
			XML.CreateTagList("MORTGAGELOAN");
			var m_iNumOfLoanComponents = XML.ActiveTagList.length;
			var tagActiveLoan = null;
			for (var iCount = 0; iCount < m_iNumOfLoanComponents; iCount++)
			{
				XML.SelectTagListItem(iCount);
				tagActiveLoan = XML.ActiveTag;
				var lRetention = XML.GetAttribute("EXISTINGRETENTIONS");

				// Add to total.
				if (lRetention == "") lRetention = 0;
				lRetention = parseInt(lRetention);
				TotalRetention = TotalRetention + lRetention;
			}

		} // End - XML.IsResponseOK()
	}
	return TotalRetention;
	
}
<%/* EP2_8 - New function.*/%>

function GetDiffInMonths(FirstDate, SecondDate)
{
// Use these if passing string in.
//	strDate1 = FirstDate.split("/"); 
//	strYear1 = strDate1[2].split(" ");
//	var date1 = new Date(strYear1[0],strDate1[1]-1,strDate1[0])
//	strDate2 = SecondDate.split("/"); 
//	strYear2 = strDate2[2].split(" ");
//	var date2 = new Date(strYear2[0],strDate2[1]-1,strDate2[0])

	var date1 = FirstDate;
	var date2 = SecondDate;
	
	// Use miliseconds difference from baseline
	var iDiffMS = date2.valueOf() - date1.valueOf();
	var dtDiff = new Date(iDiffMS);

	// Calc Difference (Also includes day adjustment - Courtesy Mark Chivers Inc.)
	var nYears  = date2.getUTCFullYear() - date1.getUTCFullYear();
	var TotalMonths = date2.getUTCMonth() - date1.getUTCMonth() + (nYears!=0 ? nYears*12 : 0);
    // Day of month adjustment. If previous DAY > latest DAY.
    if ((date1.getDate() - date2.getDate()) > 0)
    TotalMonths = TotalMonths - 1


	// Return no of months.
	return TotalMonths;
}


<% /* BMIDS624 End */ %>
-->
</script>
</body>
</html>
