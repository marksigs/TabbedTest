<%@  Language=JavaScript %>
<%
	/* PSC 16/10/2005 MAR57 - Start */
	var sStartScreen = Request.QueryString("Screen");
	/* PSC 16/10/2005 MAR57 - End */	
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
Workfile:      modal_omigamenu.asp (fw010 in the spec)
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Menu screen.  Also holds the context fields
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		12/11/99	Addition of XML2 context field
JLD		15/11/99	Addition of GroupConnection context for DC010
APS		03/09/99	Added UserType 
RF		16/11/99	Added more context info; removed CompanyId and UserType from Session object
JLD		22/11/99	Added idHasDependants for routing to/from DC040
RF		23/11/99	Added debug context info
MCS		16/11/99	Added idQuotationNumber and removed idHasDependants
JLD		17/01/00	Added idBCSubQuoteNumber and idBusinessType
AY		08/02/00	Optimised version
JLD		14/02/00	Added idMortgageSubQuoteNumber, idLifeSubquoteNumber, idEmploymentSequenceNumber
AY		16/02/00	SYS0040 - All appropriate context fields cleared when application/customer unlocked
AY		25/02/00	Hide the navigation menu when the application number is cleared
AY		29/02/00	Menu/field styles changed to match navigation bar, Application Number now displayed
AY		01/03/00	Addition of ApplicationMode context field
AY		01/03/00	Body background image set to none
					(and the date on the above ApplicationMode comment was wrong!)
APS		08/03/00	Added idApplicant1PPEligible and idApplicant2PPEligible
AD		08/03/00	Added idOtherResidentSequenceNumber
AY		10/03/00	SYS0242 - addition of StageNumber
AY		04/03/00	APPLICATION MENU title moved from FW020
					Co-ordinates changed to bring in line with FW020
AY		06/04/00	SYS0569 - Stage Number context field replaced with combo context field 
					which is populated with stage info on system start
AY		11/04/00	SYS0328 - introduction of Currency context field
MH		12/05/00    SYS0735 - IE5.1/IE4 behaviour difference. FUDGE REQUIRED
					Hitting the cancel button within a msgbox seems to generate an unneeded
					onbeforeunload event for the page hosting the cancel button. In this 
					form it means that a customer is unlocked when it shouldn't be.
MC		15/05/00	SYS0694 - Amended idXML and idXML2 to type=hidden therefore allowing linefeed
							  characters to be saved in textarea fields.
MH      02/06/00    SYS0735 Again.
AP		16/08/00	SYS0735 Yet Again.
MV		14/11/00	CORE000007 - Addition of sUnitName Variable and Input
JLD		29/11/00	Added idActivityID
MV		05/12/00	CORE00015 -  Added sDepartmentName and sChannelName 
GD		20/12/00	AAA18 (previously SYS1687) Added Task Management link to mn015.asp
GD		21/12/00	SYS1742 changes made to allow use of hidden frame
					and eliminate session variables and use Opener.
JLD		15/01/01	SYS1808 replaced context idApplicationStageId with idStageName, idStageId and idCurrentStageOrigSeqNo
JLD		22/01/01	SYS1832 Added context idAppUnderReview
MV		15/01/01	SYS1948 Added Context idAccessAuditGUID
MV		20/01/01	SYS1967 Added Context idApplicationPriority
APS		02/03/01	SYS1920 Added into the context 'FreezeDataIndicator'
MV		05/03/01	SYS2001 Commented AccessAuditGUID
JLD		07/03/01	SYS1879 Added idReturnScreenId context
CL		07/03/01	SYS2015 Added Payment Sanction
CL		15/03/01	SYS2044 Added Customer Name
IK		17/03/01	SYS1924 Add idProcessContext as idMetaAction is mis-used
IK		17/03/01	SYS1924 Add CONSTANT value for TaskStatus = Complete as idconstTaskStatusComplete
IK		21/03/01	SYS1924 Add idTaskXML for casetask XML context
IK		22/03/01	SYS2145 Add idNoCompleteCheck
SR		24/03/01	SYS2325 Change the layout for displaying Application No, 
					Display OtherSystemAccountNo as 'A/CNO'
RF		13/11/01	SYS2927 Added AdminSystemState to context.
GF		26/11/01	SYS2927 Commented out AdminSystemState call
RF		21/01/02	SYS2927 Reversed GF's changes
DP		19/01/02	SYS2564 (parent), SYS3829(child) allow client projects to set
					their own currency. Customise routine added.
MV		31/01/2002	SYS3961 Customised SystemMenu Div for Clients to disable Mortgage Calculator menu Option				
TW      09/10/2002	SYS5115	Modified to incorporate client validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		22/05/2002	BMIDS00005	Added SearchFlag,SearchSurname,SearchForeName,SearchDOB,SearchPostCode to Context
AW		13/06/02	CMWP1		Removed Mortgage Calculator from system menu
GD		19/06/2002	BMIDS00077  Upgrade to Core 7.0.2
MO		02/07/2002	BMIDS00090	Removed Mortgage Calculator from system menu, and changed 'customer registration' to 'customer search'
MV		08/07/2002  BMIDS00188  Core Upgrade Error - Ameded Window.onLoad()
PSC		05/08/2002	MBIDS00062	Add Account Guid
MV		09/10/2002	BMIDS00556	Removed Validate_Init()
LD		13/03/2002  BMIDS0432	Retry unlocking customer/app after an error
BS		17/03/2003	BM0324		Amend unlock cust/app failure message
DRC     03/02/2004	BMIDS692    New include file OverrideContextAttribs.asp for clent configurable freeze data overides 
HMA     29/01/2004  BMIDS678    Add CreditCheckAccess
INR		20/02/2004	BMIDS682	Added idAddressTarget context
MC		30/04/2004	BM0468		Cancel and Decline stage freeze screens functionality
MC		24/05/2004	BMIDS747	idVerificationSeqNo ADDED
MC		06/07/2004	BMIDS763	idApplicationDate Added to context
GHun	29/08/2004	BMIDS821	Clear ApplicationDate
SR		27/08/2004  BMIDS815	idLTVChanged, idCallingScreenID
HMA     04/10/2004  BMIDS903    Clear idLTVChanged on exit from an application 
HMA     23/11/2004  BMIDS948    Add Type of Application.
HMA     08/12/2004  BMIDS957    Remove unused VBScript.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
PSC		06/10/2005	MAR57		Add Customer CategoriesLaunch, LaunchCustomerXML and ExistingCustomerLaunch and LaunchCustomerNumber
PSC		25/10/2005	MAR300		Add OtherSystemCustomerNumber for each customer and change idExistingCustomerLaunch to idExistingApplication
HMA     28/11/2005  MAR324      Add AllowExitFromWrap context variable.
HMA     07/12/2005  MAR797      Disable Customer Search option depending on Global Parameters.
PE		27/02/2006	MAR1310		Allow testing option to be switched off.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
SAB		13/04/2006	EP388		MAR1300
SAB		13/04/2006	EP419		MAR1587
AW		26/05/2006	EP623		Clear Shortfall on application exit	
AW		12/12/2006	EP1240		Added contexts idIsUserQualityChecker, idIsUserFulfillApproved, idCorrespondenceSalutation	

*/
%>
<% //GD session variables no longer reqd (deleted by RF 13/11/01) %>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<title>Omiga 4</title>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/modal_omigamenuAttribs.asp" -->

<% /* BMIDS957  Remove unused vbScript */ %>


<script language="javascript" type="text/javascript">
<!--

function window_onload()
{
	var sUName = null;
	var sUGuid = null;
	
	frmContext.IE5BugDisableUnloadEvent.value ="";
	
}
-->
</script>
</head>
<body topmargin="15" leftmargin="15" class="sysMenuBackground" bgcolor="#616161" alink="black" vlink="gray" link="black" onload="window_onload()">
<% /* XML Functions Object - see comments within scXMLFunctions.asp for details */ %>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" viewastext tabindex="-1"></object>
<% /*  MAR1587 M Heys 13/04/2006*/ %>
<object data="scScreenFunctions.asp" height="1" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabindex="-1" type="text/x-scriptlet" width="1" viewastext></object>

<form id="frmContext" style="display:none">
	<input id="IE5BugDisableUnloadEvent" name="bIE5BugDisableUnloadEvent" tabindex="-1">
	<input id="idDebug" name="Debug"  tabindex="-1">
	<input id="idUserId" name="UserId"  tabindex="-1">
	<input id="idUserName" name="UserName" tabindex="-1">
	<input id="idUserType" name="UserType" tabindex="-1">
	<input id="idAccessType" name="AccessType" tabindex="-1">
	<input id="idUserCompetency" name="UserCompetency" tabindex="-1">
	<input id="idQualificationsListXml" name="QualificationsListXml" tabindex="-1">
	<input id="idRole" name="Role" tabindex="-1">
	<input id="idUnitId" name="UnitId" tabindex="-1">
	<% // MV - 13/11/00 */ %>
	<input id="idUnitName" name="UnitName"  tabindex="-1">
	<input id="idDepartmentId" name="DepartmentId" tabindex="-1">
	<% // MV - 05/12/00 */ %>
	<input id="idDepartmentName" name="DepartmentName" tabindex="-1">
	<input id="idCompanyId" name="CompanyId" tabindex="-1">
	<input id="idDistributionChannelId" name="DistributionChannelId"  tabindex="-1">
	<% // MV - 05/12/00 */ %>
	<input id="idDistributionChannelName" name="DistributionChannelName"  tabindex="-1">
	<input id="idProcessingIndicator" name="ProcessingIndicator"  tabindex="-1">
	<input id="idMachineId" name="MachineId"  tabindex="-1">
	<input id="idCurrency" name="Currency" tabindex="-1">
	<input id="idMetaAction" name="MetaAction" tabindex="-1">
	<input id="idProcessContext" name="ProcessContext" tabindex="-1">
	<input id="idAmountRequested" name="AmountRequested" tabindex="-1">
	<input id="idPurchasePrice" name="PurchasePrice" tabindex="-1">
	<input id="idTermYears" name="TermYears" tabindex="-1">
	<input id="idTermMonths" name="TermMonths" tabindex="-1">
	<input id="idMortgageProductId" name="MortgageProductId" tabindex="-1">
	<input id="idMortgageProductStartDate" name="MortgageProductStartDate" tabindex="-1">
	<input id="idLendersCode" name="LendersCode" tabindex="-1">
	<input id="idPackageNumber" name="PackageNumber" tabindex="-1">
	<input id="idApplicationNumber" name="ApplicationNumber" tabindex="-1">
	<input id="idApplicationFactFindNumber" name="ApplicationFactFindNumber" tabindex="-1">
	<input id="idOtherSystemCustomerNumber" name="OtherSystemCustomerNumber" tabindex="-1">
	<input id="idMortgageApplicationDescription" name="MortgageApplicationDescription" tabindex="-1">
	<input id="idMortgageApplicationValue" name="MortgageApplicationValue" tabindex="-1">
	<input id="idCustomerNumber" name="CustomerNumber" tabindex="-1">
	<input id="idCustomerVersionNumber" name="CustomerVersionNumber" tabindex="-1">
	<input id="idCustomerName" name="CustomerName" tabindex="-1">
	<input id="idCustomerName1" name="CustomerName1" tabindex="-1">
	<input id="idCustomerCategory1" name="CustomerName1" tabindex="-1">
	<input id="idOtherSystemCustomerNumber1" name="CustomerName1" tabindex="-1">
	<input id="idCustomerNumber1" name="CustomerNumber1" tabindex="-1">
	<input id="idCustomerVersionNumber1" name="CustomerVersionNumber1" tabindex="-1">
	<input id="idCustomerRoleType1" name="CustomerRoleType1" tabindex="-1">
	<input id="idCustomerOrder1" name="CustomerOrder1" tabindex="-1">
	<input id="idCustomerName2" name="CustomerName2" tabindex="-1">
	<input id="idCustomerCategory2" name="CustomerName1" tabindex="-1">
	<input id="idOtherSystemCustomerNumber2" name="CustomerName1" tabindex="-1">
	<input id="idCustomerNumber2" name="CustomerNumber2" tabindex="-1">
	<input id="idCustomerVersionNumber2" name="CustomerVersionNumber2" tabindex="-1">
	<input id="idCustomerRoleType2" name="CustomerRoleType2" tabindex="-1">
	<input id="idCustomerOrder2" name="CustomerOrder2" tabindex="-1">
	<input id="idCustomerName3" name="CustomerName3" tabindex="-1">
	<input id="idCustomerCategory3" name="CustomerName1" tabindex="-1">
	<input id="idOtherSystemCustomerNumber3" name="CustomerName1" tabindex="-1">
	<input id="idCustomerNumber3" name="CustomerNumber3" tabindex="-1">
	<input id="idCustomerVersionNumber3" name="CustomerVersionNumber3" tabindex="-1">
	<input id="idCustomerRoleType3" name="CustomerRoleType3" tabindex="-1">
	<input id="idCustomerOrder3" name="CustomerOrder3" tabindex="-1">
	<input id="idCustomerName4" name="CustomerName4" tabindex="-1">
	<input id="idCustomerCategory4" name="CustomerName1" tabindex="-1">
	<input id="idOtherSystemCustomerNumber4" name="CustomerName1" tabindex="-1">
	<input id="idCustomerNumber4" name="CustomerNumber4" tabindex="-1">
	<input id="idCustomerVersionNumber4" name="CustomerVersionNumber4" tabindex="-1">
	<input id="idCustomerRoleType4" name="CustomerRoleType4" tabindex="-1">
	<input id="idCustomerOrder4" name="CustomerOrder4" tabindex="-1">
	<input id="idCustomerName5" name="CustomerName5" tabindex="-1">
	<input id="idCustomerCategory5" name="CustomerName1" tabindex="-1">
	<input id="idOtherSystemCustomerNumber5" name="CustomerName1" tabindex="-1">
	<input id="idCustomerNumber5" name="CustomerNumber5" tabindex="-1">
	<input id="idCustomerVersionNumber5" name="CustomerVersionNumber5" tabindex="-1">
	<input id="idCustomerRoleType5" name="CustomerRoleType5" tabindex="-1">
	<input id="idCustomerOrder5" name="CustomerOrder5" tabindex="-1">
	<input type=hidden id="idXML" name="XML" tabindex="-1">
	<input type=hidden id="idXML2" name="XML2" tabindex="-1">
	<input id="idTaskXML" name="TaskXML" tabindex="-1">
	<input id="idCustomerReadOnly" name="CustomerReadOnly" tabindex="-1">
	<input id="idApplicationReadOnly" name="ApplicationReadOnly" tabindex="-1">
	<input id="idReadOnly" name="ReadOnly" tabindex="-1">
	<input id="idCustomerLockIndicator" name="CustomerLockIndicator" tabindex="-1">
	<input id="idMCIllustrationNumber" name="MCIllustrationNumber" tabindex="-1">
	<input id="idMCDisplayNumber" name="MCDisplayNumber" tabindex="-1">
	<input id="idMultipleLender" name="MultipleLender" tabindex="-1">
	<input id="idQuotationNumber" name="QuotationNumber" tabindex="-1">
	<input id="idMortgageSubQuoteNumber" name="MortgageSubQuoteNumber" tabindex="-1">
	<input id="idLifeSubQuoteNumber" name="LifeSubQuoteNumber" tabindex="-1">
	<input id="idBCSubQuoteNumber" name="BCSubQuoteNumber" tabindex="-1">
	<input id="idPPSubQuoteNumber" name="PPSubQuoteNumber" tabindex="-1">
	<input id="idBusinessType" name="BusinessType" tabindex="-1">
	<input id="idEmploymentSequenceNumber" name="EmploymentSequenceNumber" tabindex="-1">
	<input id="idCustomerIndex" name="CustomerIndex" tabindex="-1">
	<input id="idEmploymentStatusType" name="EmploymentStatusType" tabindex="-1">
	<input id="idCustomerAddressType" name="CustomerAddressType" tabindex="-1">
	<input id="idApplicationMode" name="ApplicationMode" tabindex="-1">
	<input id="idEmployerName" name="EmployerName" tabindex="-1">
	<input id="idMainEmploymentCount" name="MainEmploymentCount" tabindex="-1">
	<input id="idEmploymentCount" name="EmploymentCount" tabindex="-1">
	<input id="idApplicant1PPEligible" name="Applicant1PPEligible" tabindex="-1">
	<input id="idApplicant2PPEligible" name="Applicant2PPEligible" tabindex="-1">
	<input id="idOtherResidentSequenceNumber" name="OtherResidentSequenceNumber" tabindex="-1">
	<input id="idActivityId" name="ActivityId" value="10" tabindex="-1">
	<input id="idStageName" name="StageName" tabindex="-1">
	<input id="idStageId" name="StageId" tabindex="-1">
	<input id="idCurrentStageOrigSeqNo" name="StageSequenceNumber" tabindex="-1">
	<% // input id="idAccessAuditGUID" name="AccessAuditGUID" tabindex="-1"> %>
	<input id="idAppUnderReview" name="AppUnderReview" tabindex="-1">
	<input id="idApplicationPriority" name="ApplicationPriority" tabindex="-1">
	<input id="idFreezeDataIndicator" name="FreezeDataIndicator" tabindex="-1">
	<input id="idCancelDeclineFreezeDataIndicator" name="CancelDeclineFreezeDataIndicator" tabindex="-1">
	<!-- #include FILE="attribs/OverrideContextAttribs.asp" -->
	<input id="idReturnScreenId" name="ReturnScreenId" tabindex="-1">
	<input id="idNoCompleteCheck" name="NoCompleteCheck" tabindex="-1">
	<%	// IK 17/03/2001, constant values here as GLOBALPARAMETER values 	%>
	<input id="idconstTaskStatusComplete" name="TaskStatusComplete" value="40" tabindex="-1">	
	<select id="idDataFreezeScreens" name="DataFreezeScreens" tabindex="-1"></select>
	<select id="idCancelDeclineDataFreezeScreens" name="CancelDeclineDataFreezeScreens" tabindex="-1"></select>
	<input id="idOtherSystemAccountNumber" name="OtherSystemAccountNumber" tabindex="-1">
	<% // RF 13/11/01 SYS2927 %>
	<input id="idAdminSystemState" name="AdminSystemState" tabindex="-1">
	<% // AQR SYS3048 DC 22/11/2001 - Default Instruction Sequence Number %>
	<input id="idInstructionSequenceNo" name="InstructionSequenceNo" tabindex="-1">
	<% /* MV : 22/05/2002 :  BMIDS00005 - BM052 Customer Searches 
	Start */ %>
	<input id="idSearchFlag" name="SearchFlag" tabindex="-1">
	<input id="idSearchSurname" name="SearchSurname" tabindex="-1">
	<input id="idSearchForename" name="SearchForename" tabindex="-1">
	<input id="idSearchDateOfBirth" name="SearchDateOfBirth" tabindex="-1">
	<input id="idSearchPostCode" name="SearchPostCode" tabindex="-1">
	<% //End %>
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	<% // DB SYS4767 - MSMS Integration %>
	<% // SG 05/04/02 MSMS009 %>
	<input id="idShortfallPayment" name="ShortfallPayment" tabindex="-1">
	<% // DB End %>
	<% /* SYS4536 - If payment details exist for a payee, set PP100 to read-only. */ %>
	<input id="idPayeeReadOnly" name="PayeeReadOnly" tabindex="-1">
	<% // PSC 05/08/2002 BMIDS00006 %>
	<input id="idAccountGuid" name="AccountGuid" tabindex="-1">
	<% // BMIDS678 - 29/01/04 */ %>
	<input id="idCreditCheckAccess" name="CreditCheckAccess" tabindex="-1">
		<% // INR 20/02/2004 BMIDS682 %>
	<input id="idAddressTarget" name="AddressTarget" tabindex="-1">

	<input id="idCancelledStageValue" name="CancelledStageValue" tabindex="-1">
	<input id="idDeclinedStageValue" name="DeclinedStageValue" tabindex="-1">
	<input id="idVerificationSeqNo" name="VerificationSeqNo" tabindex="-1">
	
	<input id="idApplicationDate" name="ApplicationDate" tabindex="-1">
	<%/* SR 27/08/2004 : BMIDS815 */%>
	<input id="idLTVChanged" name="LTVChanged" tabindex="-1"> 
	<input id="idCallingScreenID" name="CallingScreenID" tabindex="-1"> 
	<%/* SR 27/08/2004 : BMIDS815 - End */%>
	<%/* PSC 06/10/2005 MAR57 - Start */%>
	<input id="idLaunchCustomerXML" name="CustomerLaunchXML" tabindex="-1">
	<%/* PSC 25/10/2005 MAR300 */%>
	<input id="idExistingApplication" name="ExistingApplication" tabindex="-1"> 
	<input id="idLaunchCustomerNumber" name="LaunchCustomerNumber" tabindex="-1">
	<input id="idRegulationIndicator" name="RegulationIndicator" tabindex="-1"> 
	<%/* PSC 06/10/2005 MAR57 - End */%>
	<%/* MAR324 */%>
	<input id="idAllowExitFromWrap" name="AllowExitFromWrap" tabindex="-1"> 
	<%/* MAR324 - End */%>
	
	<input id="idAllowTesting" name="AllowTesting" tabindex="-1"> <% // PE 27/02/2006 MAR1310 %>
	<input id="idPropertyChange" name="PropertyChange" tabindex="-1"> <% // MAR1300 GHun %>
	<%/* AW 30/10/2006 : EP1240 */%>
	<input id="idIsUserQualityChecker" name="IsUserQualityChecker" tabindex="-1"> 
	<input id="idIsUserFulfillApproved" name="IsUserFulfillmentApproved" tabindex="-1">
	<input id="idCorrespondenceSalutation" name="CorrespondenceSalutation" tabindex="-1">  
	<%/* AW 30/10/2006 : EP1240 - End */%>		
</form>

<div style="TOP: 0px; LEFT: 4px; POSITION: ABSOLUTE">
	<img src="images/mainback.gif" width="151" height="265" style="LEFT: -4px; POSITION: absolute" alt="">
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	<% // DB SYS4767 - MSMS Integration %>
	<% //<div style="TOP: 65px; LEFT: 0px; POSITION: absolute" > %>
	<div id="spnVersion" style="POSITION: absolute; TOP: 60px;">
		<div class="navText" id="divVersionNo" style="LEFT: -3px; POSITION: absolute; TOP: 0px">
			<label id="lblVersionNo" style="WIDTH: 130px"></label>
		</div>
	</div>
	<% // DB End %>

	<div style="TOP: 72px; LEFT: 0px; WIDTH: 145px; POSITION: absolute" >
		<div class="navTitle" id="systemmenu">SYSTEM MENU</div>
		<div class="navLink" id="aMortgageCalculator"></div>
		<!--<div class="navLink"><a id="aCustomerRegistration" href="#" tabindex="-1">Customer Registration</a></div>-->
		<div class="navLink" id="idCustomerSearch"><a id="aCustomerSearch" href="#" tabindex="-1">Customer Search</a></div>
		<div class="navLink" id="idPipelineEnquiry"><a id="aPipelineEnquiry" href="#" tabindex="-1">Application Enquiry</a></div>
		<div class="navLink" id="idTaskManagement"><a id="aTaskManagement" href="#" tabindex="-1">Task Management</a></div>
		<div class="navLink" id="idPaymentSanction"><a id="aPaymentSanction" href="#" tabindex="-1">Payment Sanction</a></div>
	</div>

<!-- Commented Out by MV - 31/01/2002
	
	<div class="navTitle" style="TOP: 35px; LEFT: 0px; POSITION: absolute">SYSTEM MENU</div>
	<div class="navLink" style="TOP: 49px; LEFT: 0px; POSITION: absolute"><a id="aMortgageCalculator" href="#" tabindex="-1">Mortgage Calculator</a></div>
	<div class="navLink" style="TOP: 63px; LEFT: 0px; POSITION: absolute"><a id="aCustomerRegistration" href="#" tabindex="-1">Customer Registration</a></div>
	<div class="navLink" style="TOP: 77px; LEFT: 0px; POSITION: absolute"><a id="aPipelineEnquiry" href="#" tabindex="-1">Application Enquiry</a></div>
	<div class="navLink" style="TOP: 91px; LEFT: 0px; POSITION: absolute"><a id="aTaskManagement" href="#" tabindex="-1">Task Management</a></div>
	<div class="navLink" style="TOP: 105px; LEFT: 0px; POSITION: absolute"><a id="aPaymentSanction" href="#" tabindex="-1">Payment Sanction</a></div>
-->

<!--
	<div class="navLink" style="TOP: 72px; LEFT: 0px; POSITION: absolute"><a id="aPipelineEnquiry" href="#" tabindex="-1">Pipeline Enquiry</a></div>
	<div class="navLink" style="TOP: 86px; LEFT: 0px; POSITION: absolute"><a id="aMultipleApplicationEnquiry" href="#" tabindex="-1">Multi-Application Enq.</a></div>
	<div class="navLink" style="TOP: 100px; LEFT: 0px; POSITION: absolute"><a id="aSupervisor" href="#" tabindex="-1">Supervisor</a></div>
	<div class="navLink" style="TOP: 114px; LEFT: 0px; POSITION: absolute"><a id="aWorkflow" href="#" tabindex="-1">Workflow</a></div>
-->		
</div>

<div id="spnApplication" style="TOP: 141px; LEFT: 2px; POSITION: ABSOLUTE; VISIBILITY: hidden">
	<div class="navOutline" style="TOP: 0px; LEFT: 0px; WIDTH: 145px; HEIGHT: 50px">
		<div class="navTitle" style="TOP: 4px; LEFT: 4px; POSITION: absolute">APP NO:</div>
		<div class="navText" style="TOP: 4px; LEFT: 54px; POSITION: absolute">
			<label id="lblApplicationNo" style="WIDTH: 95px"></label>
		</div>
	</div>
</div>
<div id="spnAccount" style="TOP: 160px; LEFT: 2px; POSITION: ABSOLUTE; VISIBILITY: hidden">
	<div class="navTitle" style="TOP: 0px; LEFT: 4px; POSITION: absolute">A/C NO:</div>
	<div class="navText" style="TOP: 0px; LEFT: 54px; POSITION: absolute">
		<label id="lblAccountNo" style="WIDTH: 95px"></label>
	</div>
</div>
<div id="divPropertyChange" style="TOP: 160px; LEFT: 2px; POSITION: ABSOLUTE; VISIBILITY: hidden">
	<div class="navTitle" style="TOP: 0px; LEFT: 4px; POSITION: absolute">Property Change</div>
	<div class="navText" style="TOP: 0px; LEFT: 54px; POSITION: absolute; VISIBILITY: hidden">
		<label id="lblPropertyChange" style="VISIBILITY: hidden"></label>
	</div>
</div>
<div id="spnApplicationType" style="TOP: 175px; LEFT: 2px; POSITION: ABSOLUTE; VISIBILITY: hidden">
	<div class="navText" style="TOP: 0px; LEFT: 4px; POSITION: absolute">
		<label id="lblApplicationType" style="WIDTH: 130px"></label>
	</div>
</div>

<div id="spnCustomers" style="TOP: 197px; LEFT: 2px; COLOR: white; POSITION: ABSOLUTE; VISIBILITY: hidden">
	<div class="navOutline" style="LEFT: 0px; POSITION: ABSOLUTE; WIDTH: 145px; HEIGHT: 132px">
		<div class="navTitle" id="customermenu" style="WIDTH: 144px; TOP: 2px; LEFT: 4px; POSITION: absolute">CUSTOMERS</div>
		<div id="divCustomer1"class="navText" style="TOP: 14px; LEFT: 4px; POSITION: absolute">
			<label id="lblCustomer1" style="WIDTH: 120px"></label>
			<div id="divCustomerCategory1"class="navText" style="TOP: 10px; LEFT: 10px; POSITION: absolute">
				<label id="lblCustomerCategory1" style="WIDTH: 110px"></label>
			</div>
		</div>
		<div id="divCustomer2"class="navText" style="TOP: 37px; LEFT: 4px; POSITION: absolute">
			<label id="lblCustomer2" style="WIDTH: 120px"></label>
			<div id="divCustomerCategory2"class="navText" style="TOP: 10px; LEFT: 10px; POSITION: absolute">
				<label id="lblCustomerCategory2" style="WIDTH: 110px"></label>
			</div>
		</div>
		<div id="divCustomer3"class="navText" style="TOP: 60px; LEFT: 4px; POSITION: absolute">
			<label id="lblCustomer3" style="WIDTH: 120px"></label>
			<div id="divCustomerCategory3"class="navText" style="TOP: 10px; LEFT: 10px; POSITION: absolute">
				<label id="lblCustomerCategory3" style="WIDTH: 110px"></label>
			</div>
		</div>
		<div id="divCustomer4"class="navText" style="TOP: 83px; LEFT: 4px; POSITION: absolute">
			<label id="lblCustomer4" style="WIDTH: 120px"></label>
			<div id="divCustomerCategory4"class="navText" style="TOP: 10px; LEFT: 10px; POSITION: absolute">
				<label id="lblCustomerCategory4" style="WIDTH: 110px"></label>
			</div>
		</div>
		<div id="divCustomer5"class="navText" style="TOP: 106px; LEFT: 4px; POSITION: absolute">
			<label id="lblCustomer5" style="WIDTH: 120px"></label>
			<div id="divCustomerCategory5"class="navText" style="TOP: 10px; LEFT: 10px; POSITION: absolute">
				<label id="lblCustomerCategory5" style="WIDTH: 110px"></label>
			</div>
		</div>
	</div>
</div>

<% /* HMA Navigation menu title is moved back to FW020 */ %>
<% /* <div id="divNavigation" align="center" class="navTitle" style="TOP: 300px; LEFT: 8px; WIDTH: 130px; POSITION: absolute; VISIBILITY: hidden">APPLICATION MENU</div> */ %>
<!--GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2-->
<!--  #include FILE="Customise/modal_omigamenuCustomise.asp" -->
<script language="JScript" type="text/javascript">

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	
	var objOpener, objHiddenFrameForm;
	
	objOpener=top.window.opener
	if (objOpener==null)
	{
		<%//Opener is null - redirect to error page		%>
	}
	else
	<%//opener is not null	%>
	{
		<%
		//Locate opener's frameset
		%>
		var objParent = objOpener.parent;
		if (objParent != null)
		{
			<%	//Grab every field required from opener's lhs %>
			objHiddenFrameForm=objParent.document.frames("fraDetails").document.forms("frmDetails");
			<%	//GD populate the 15 fields from logon	%>
			frmContext.idDebug.value=objHiddenFrameForm.txtDebug.value;
			frmContext.idUserId.value=objHiddenFrameForm.txtUserId.value;
			frmContext.idUserName.value=objHiddenFrameForm.txtUserName.value;
			frmContext.idAccessType.value=objHiddenFrameForm.txtAccessType.value;
			frmContext.idUserCompetency.value=objHiddenFrameForm.txtUserCompetency.value;
			frmContext.idQualificationsListXml.value=objHiddenFrameForm.txtQualificationsListXml.value;
			frmContext.idRole.value=objHiddenFrameForm.txtRole.value;
			frmContext.idUnitId.value=objHiddenFrameForm.txtUnitId.value;
			frmContext.idCreditCheckAccess.value=objHiddenFrameForm.txtCreditCheckAccess.value;
			frmContext.idUnitName.value=objHiddenFrameForm.txtUnitName.value;
			frmContext.idDepartmentId.value=objHiddenFrameForm.txtDepartmentId.value;
			frmContext.idDepartmentName.value=objHiddenFrameForm.txtDepartmentName.value;
			frmContext.idDistributionChannelId.value=objHiddenFrameForm.txtDistributionChannelId.value;
			frmContext.idDistributionChannelName.value=objHiddenFrameForm.txtDistributionChannelName.value;
			frmContext.idProcessingIndicator.value=objHiddenFrameForm.txtProcessingIndicator.value;
			frmContext.idMachineId.value=objHiddenFrameForm.txtMachineId.value;
			<% //frmContext.idAccessAuditGUID.value =objHiddenFrameForm.txtAccessAuditGUID.value; %>
			<% // RF 13/11/01 SYS2927 %>
			frmContext.idAdminSystemState.value=objHiddenFrameForm.txtAdminSystemState.value;
			<% /* PSC 17/10/2005 MAR57 */ %>
			frmContext.idLaunchCustomerXML.value = objHiddenFrameForm.txtLaunchXML.value;	
			frmContext.idAllowExitFromWrap.value = objHiddenFrameForm.txtAllowExitFromWrap.value;  // MAR324
			<% // AW 30/10/06 EP1240 %>
			frmContext.idIsUserQualityChecker.value=objHiddenFrameForm.txtIsUserQualityChecker.value;
			frmContext.idIsUserFulfillApproved.value=objHiddenFrameForm.txtIsUserFulfillApproved.value;
		}
	}
	
	<% /* MAR324  Hide Pipeline Enquiry, Task Management and Payment Sanction if AllowOmigaExitFromWrapUp 
	              is set for this unit */ %>
	
	var sAllowOmigaExitFromWrapUp = objHiddenFrameForm.txtAllowExitFromWrap.value;
	
	if (sAllowOmigaExitFromWrapUp == "1")
	{
		document.all("idPipelineEnquiry").style.visibility="hidden";
		document.all("idTaskManagement").style.visibility="hidden";
		document.all("idPaymentSanction").style.visibility="hidden";	
	}		
	
	<% /* MAR797  Hide Customer Search depending on Global Parameters */ %>
	
	var globalXML = new scXMLFunctions.XMLObject();
	var bDisableCustomerSearch = globalXML.GetGlobalParameterBoolean(document, "NAVDisableCustSearch");

	if ((sAllowOmigaExitFromWrapUp == "1") || (bDisableCustomerSearch))
	{
		document.all("idCustomerSearch").style.visibility="hidden";
	}		
	
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	// DB SYS4767 - MSMS Integration
	// STB: SYS4143 - Display the version.
	var sVersionNo = "";			
	sVersionNo = globalXML.GetGlobalParameterString(document, "OmigaVersion");
	lblVersionNo.innerText = " " + sVersionNo;
	// DB End	
	// APS SYS1920	
	SetMultipleLenderInContext();
	
	var XML = new scXMLFunctions.XMLObject();
	
	//Get Cancel and Decline Stage Values
	try
	{
		frmContext.idCancelledStageValue.value = XML.GetGlobalParameterAmount(document,"CancelledStageValue");
		frmContext.idDeclinedStageValue.value = XML.GetGlobalParameterAmount(document,"DeclinedStageValue");
	}
	catch(e)
	{
	}
	
	var sGroups = new Array("DataFreezeScreen","CancelDeclineDataFreezeScreen");
	if (XML.GetComboLists(document,sGroups)) 
	{
		XML.PopulateCombo(document,frmContext.idDataFreezeScreens,"DataFreezeScreen",false);
		//New CancelFreeze list is combination of both items.
		try
		{
		
		XML.PopulateCombo(document,frmContext.idCancelDeclineDataFreezeScreens,"DataFreezeScreen",false);
		XML.PopulateCombo(document,frmContext.idCancelDeclineDataFreezeScreens,"CancelDeclineDataFreezeScreen",false,1);
		}
		catch(ex)
		{
			alert(ex);
		}
	}
	frmContext.idCancelDeclineFreezeDataIndicator.value="";
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	// DJP Customisation SYS2564
	Customise();
	
	<% /* MV - 08/07/2002- BMIDS00188 - Core Upgrade Error */ %>
	<% /* PSC 16/10/2005 MAR57 */ %>
	top.frames[2].navigate("<%=sStartScreen%>");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function window.onbeforeunload()
{
<%	/* This event fires before this page is unloaded.  It allows the closing
	 down of the framework to be halted if there is an error within
	 ValidateApplicationOrCustomer
	 If the unlocking was unsuccessful, setting event.returnValue causes a message to be
	 displayed which requires user input
	
	IE5BugDisableUnloadEvent is a fudge. The confirm box seems to fire this event needlessly.
	This flag is set to false after a confirm box is used if Cancel was pressed. 
	If the flag is not true the unlock event will not fire.
	Closing the frameset sets this flag to true so this event should then fire.
	*/	
%>	
	if (frmContext.IE5BugDisableUnloadEvent.value=="")
	{	
		<% /* MAR1587 M Heys 13/04/2006 */ %>
		if(frmContext.idApplicationNumber.value != "")
			{
			if (frmContext.idAllowExitFromWrap.value=="1")
				{
				wrapUponbeforeunload();
				}
			}
		<% /* MAR1587 M Heys 13/04/2006 end */  %>
		if(!ValidateApplicationOrCustomer(false)) 
			<% /* BS BM0324 17/03/03 
			event.returnValue = "An error occurred while trying to remove locks"; */ %>
			event.returnValue = "WARNING - Your Customer/Application locks have not been removed, please contact your systems administrator now. - WARNING";
	}
}

function aCustomerSearch.onclick()
{
	if (ValidateApplicationOrCustomer(true)) parent.main.location = "cr010.asp";
}

<% /* MO 02/07/2002 BMIDS00090 - Rename Customer Registration
//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2
function aCustomerRegistration.onclick()
{
	if (ValidateApplicationOrCustomer(true)) parent.main.location = "cr010.asp";
}
//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2*/ %>

function aMortgageCalculator.onclick()
{
	if (ValidateApplicationOrCustomer(true)) parent.main.location = "mc010.asp";
}
function aPipelineEnquiry.onclick()
{
	if (ValidateApplicationOrCustomer(true)) parent.main.location = "gn200.asp";
}

function aTaskManagement.onclick()
{
	if (ValidateApplicationOrCustomer(true)) parent.main.location  = "mn015.asp";
}

function aPaymentSanction.onclick()
{
	if (ValidateApplicationOrCustomer(true)) parent.main.location  = "SP090.asp";
}

function SetMultipleLenderInContext()
{
<%	// if the multiple lender flag is blank then we need to retrieve the
	// parameter from the database
%>	var sMultipleLender	= frmContext.idMultipleLender.value;

	if (sMultipleLender == "")
	{
		var XML = new scXMLFunctions.XMLObject();
		XML.RunASP(document,"GetMultipleLender.asp");
		if (XML.IsResponseOK())
		{
			sMultipleLender = XML.GetTagText("MULTIPLELENDER");
			frmContext.idMultipleLender.value = sMultipleLender;
		}
		XML = null;
	}
}

function ValidateApplicationOrCustomer(bShowPrompts)
{
<%	// Removes Application and/or Customer locks and clears context fields.
	// If called from a menu option, it asks the user whether to continue.
	// If called from the dialog close button, no prompts are displayed.
	// bShowPrompts - true if from a menu option, false if from dialog close button.
	// Returns: True is successful, false if an error occurs.
%>	var bOK = false;
	var bClearContext = false;

<%	// If readonly then there are no locks so return OK,
	// otherwise we need to run lock processing
%>	if(frmContext.idReadOnly.value == "1") 
		bOK = true;
	else
	{
<%		// Is there an application number
%>		if(frmContext.idApplicationNumber.value != "")
		{
<%			// Show the question if it is required.
			// If not the locks are removed automatically
%>			var bRemoveLocks = true;
					
			if(bShowPrompts)
				if(!confirm("Press OK to exit the application.")) bRemoveLocks = false;
					frmContext.IE5BugDisableUnloadEvent.value = "Do not unload";
					
			<% // If the locks are to be removed, generate the XML and call the data access .asp file %>	
			if(bRemoveLocks)
			{
				frmContext.IE5BugDisableUnloadEvent.value ="";
				var XML = new scXMLFunctions.XMLObject();
				XML.CreateRequestTag(window,"DELETE");
				XML.CreateTag("APPLICATIONNUMBER",frmContext.idApplicationNumber.value);
				// 				XML.RunASP(document,"UnlockApplicationAndCustomers.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									XML.RunASP(document,"UnlockApplicationAndCustomers.asp");
									<% // try again on an error %>	
									if (XML.IsResponseOK() == false)
									{
										XML.RunASP(document,"UnlockApplicationAndCustomers.asp");
									}
						break;
					default: // Error
						XML.SetErrorResponse();
					}

				if(XML.IsResponseOK()) bOK = true;
				XML = null;
			}
		}
		else
		{
			frmContext.IE5BugDisableUnloadEvent.value ="Do not unload";
			if(frmContext.idCustomerNumber.value != "")
			{
<%				// Show the question if it is required.
				// If not the locks are removed automatically
%>				var bRemoveLocks = true;

				if(bShowPrompts)
					if(!confirm("Press OK to exit the customer.")) bRemoveLocks = false;
				
				<%	// If the locks are to be removed, generate the XML and call the data access .asp file %>
				if(bRemoveLocks)
				{
					frmContext.IE5BugDisableUnloadEvent.value ="";
					var XML = new scXMLFunctions.XMLObject();
					XML.CreateRequestTag(window,"DELETE");
					XML.CreateActiveTag("CUSTOMERLOCK");
					XML.CreateTag("CUSTOMERNUMBER", frmContext.idCustomerNumber.value);
					// 					XML.RunASP(document,"DeleteCustomerLock.asp");
					// Added by automated update TW 09 Oct 2002 SYS5115
					switch (ScreenRules())
						{
						case 1: // Warning
						case 0: // OK
											XML.RunASP(document,"DeleteCustomerLock.asp");
							break;
						default: // Error
							XML.SetErrorResponse();
						}

					if(XML.IsResponseOK()) bOK = true;
					XML = null;
				}
			}
			else {
				frmContext.IE5BugDisableUnloadEvent.value ="Do not unload";
				bOK = true; <% /* There are no locks to clear, so return OK */ %>
			}
		}
	}

<%	// If we don't want to stay where we are, clear the required context fields.
%>	if(bOK)
	{
		frmContext.idMetaAction.value = "";
		frmContext.idProcessContext.value = "";
		frmContext.idAmountRequested.value = "";
		frmContext.idPurchasePrice.value = "";
		frmContext.idTermYears.value = "";
		frmContext.idTermMonths.value = "";
		frmContext.idMortgageProductId.value = "";
		frmContext.idMortgageProductStartDate.value = "";
		frmContext.idLendersCode.value = "";
		frmContext.idPackageNumber.value = "";
		frmContext.idApplicationNumber.value = "";
		frmContext.idApplicationFactFindNumber.value = "";
		frmContext.idOtherSystemCustomerNumber.value = "";
		frmContext.idMortgageApplicationDescription.value = "";
		frmContext.idMortgageApplicationValue.value = "";
		frmContext.idCustomerNumber.value = "";
		frmContext.idCustomerName.value = "";
		frmContext.idCustomerVersionNumber.value = "";
		frmContext.idCustomerName1.value = "";
		<% /* PSC 12/10/2005 MAR57 */ %>
		frmContext.idCustomerCategory1.value = "";
		<% /* PSC 25/10/2005 MAR300 */ %>
		frmContext.idOtherSystemCustomerNumber1.value = "";
		frmContext.idCustomerNumber1.value = "";
		frmContext.idCustomerVersionNumber1.value = "";
		frmContext.idCustomerRoleType1.value = "";
		frmContext.idCustomerOrder1.value = "";
		frmContext.idCustomerName2.value = "";
		<% /* PSC 12/10/2005 MAR57 */ %>
		frmContext.idCustomerCategory2.value = "";
		<% /* PSC 25/10/2005 MAR300 */ %>
		frmContext.idOtherSystemCustomerNumber2.value = "";
		frmContext.idCustomerNumber2.value = "";
		frmContext.idCustomerVersionNumber2.value = "";
		frmContext.idCustomerRoleType2.value = "";
		frmContext.idCustomerOrder2.value = "";
		frmContext.idCustomerName3.value = "";
		<% /* PSC 12/10/2005 MAR57 */ %>
		frmContext.idCustomerCategory3.value = "";
		<% /* PSC 25/10/2005 MAR300 */ %>
		frmContext.idOtherSystemCustomerNumber3.value = "";
		frmContext.idCustomerNumber3.value = "";
		frmContext.idCustomerVersionNumber3.value = "";
		frmContext.idCustomerRoleType3.value = "";
		frmContext.idCustomerOrder3.value = "";
		frmContext.idCustomerName4.value = "";
		<% /* PSC 12/10/2005 MAR57 */ %>
		frmContext.idCustomerCategory4.value = "";
		<% /* PSC 25/10/2005 MAR300 */ %>
		frmContext.idOtherSystemCustomerNumber4.value = "";
		frmContext.idCustomerNumber4.value = "";
		frmContext.idCustomerVersionNumber4.value = "";
		frmContext.idCustomerRoleType4.value = "";
		frmContext.idCustomerOrder4.value = "";
		frmContext.idCustomerName5.value = "";
		<% /* PSC 12/10/2005 MAR57 */ %>
		frmContext.idCustomerCategory5.value = "";
		<% /* PSC 25/10/2005 MAR300 */ %>
		frmContext.idOtherSystemCustomerNumber1.value = "";
		frmContext.idCustomerNumber5.value = "";
		frmContext.idCustomerVersionNumber5.value = "";
		frmContext.idCustomerRoleType5.value = "";
		frmContext.idCustomerOrder5.value = "";
		frmContext.idXML.value = "";
		frmContext.idXML2.value = "";
		frmContext.idTaskXML.value = "";
		frmContext.idCustomerReadOnly.value = "";
		frmContext.idApplicationReadOnly.value = "";
		frmContext.idReadOnly.value = "";
		frmContext.idCustomerLockIndicator.value = "";
		frmContext.idMCIllustrationNumber.value = "";
		frmContext.idMCDisplayNumber.value = "";
		frmContext.idQuotationNumber.value = "";
		frmContext.idMortgageSubQuoteNumber.value = "";
		frmContext.idLifeSubQuoteNumber.value = "";
		frmContext.idBCSubQuoteNumber.value = "";
		frmContext.idPPSubQuoteNumber.value = "";
		frmContext.idBusinessType.value = "";
		frmContext.idEmploymentSequenceNumber.value = "";
		frmContext.idCustomerIndex.value = "";
		frmContext.idEmploymentStatusType.value = "";
		frmContext.idCustomerAddressType.value = "";
		frmContext.idApplicationMode.value = "";
		frmContext.idEmployerName.value = "";
		frmContext.idMainEmploymentCount.value = "";
		frmContext.idOtherResidentSequenceNumber.value = "";
		frmContext.idNoCompleteCheck.value = "";
		frmContext.idStageId.value = "";
		frmContext.idStageName.value = "";
		frmContext.idCurrentStageOrigSeqNo.value = "";
		frmContext.idApplicationPriority.value = "";
		frmContext.idAppUnderReview.value = "";
		<% /* ASu BMIDS00037 - Start */ %>
		frmContext.idFreezeDataIndicator.value = "";
		<% /* ASu End */ %>
		<% /* MV : 22/05/2002 :  BMIDS00005 - BM052 Customer Searches 
		Start*/ %>
		frmContext.idCancelDeclineFreezeDataIndicator.value="";
		frmContext.idSearchFlag.value = "";
		frmContext.idSearchForename.value = "";
		frmContext.idSearchSurname.value = "";
		frmContext.idSearchDateOfBirth.value = "";
		frmContext.idSearchPostCode.value = "";
		<% // End %> 
		<% // INR 20/02/2004 BMIDS682 %>
		frmContext.idAddressTarget.value = "";
		<% // AW 31/10/2006 EP1240 %>
		frmContext.idCorrespondenceSalutation.value = "";
		
		<% /* BMIDS821 GHun Clear ApplicationDate */ %>
		frmContext.idApplicationDate.value = "";
		
		<% /* BMIDS903 Clear idLTVChanged */ %>
		frmContext.idLTVChanged.value = "";
		
		<% /* EP623 AW Clear Shortfall */ %>
		frmContext.idShortfallPayment.value = "";
		
<%		// We also need to hide and reset the application number and customer name fields
%>		spnApplication.style.visibility = "hidden";
		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
		//SYS4207 DRC 02/05/2002
        spnAccount.style.visibility = "hidden";
        divPropertyChange.style.visibility = "hidden";	<% //MAR1300 GHun %>
        spnApplicationType.style.visibility = "hidden";     // BMIDS948
		spnCustomers.style.visibility = "hidden";

		divCustomer1.style.visibility = "hidden";
		divCustomer1.title = "";
		lblCustomer1.innerHTML = "";
		<% /* PSC 12/10/2005 MAR57 - Start */ %>
		divCustomerCategory1.title = "";
		lblCustomerCategory1.innerHTML = "";
		<% /* PSC 12/10/2005 MAR57 - End */ %>

		divCustomer2.style.visibility = "hidden";
		divCustomer2.title = "";
		lblCustomer2.innerHTML = "";
		<% /* PSC 12/10/2005 MAR57 - Start */ %>
		divCustomerCategory2.title = "";
		lblCustomerCategory2.innerHTML = "";
		<% /* PSC 12/10/2005 MAR57 - End */ %>
		
		divCustomer3.style.visibility = "hidden";
		divCustomer3.title = "";
		lblCustomer3.innerHTML = "";
		<% /* PSC 12/10/2005 MAR57 - Start */ %>
		divCustomerCategory3.title = "";
		lblCustomerCategory3.innerHTML = "";
		<% /* PSC 12/10/2005 MAR57 - End */ %>

		divCustomer4.style.visibility = "hidden";
		divCustomer4.title = "";
		lblCustomer4.innerHTML = "";
		<% /* PSC 12/10/2005 MAR57 - Start */ %>
		divCustomerCategory4.title = "";
		lblCustomerCategory4.innerHTML = "";
		<% /* PSC 12/10/2005 MAR57 - End */ %>

		divCustomer5.style.visibility = "hidden";
		divCustomer5.title = "";
		lblCustomer5.innerHTML = "";
		<% /* PSC 12/10/2005 MAR57 - Start */ %>
		divCustomerCategory5.title = "";
		lblCustomerCategory5.innerHTML = "";
		
		frmContext.idLaunchCustomerXML.value = "";
		<% /* PSC 25/10/2005 MAR300 */ %>
		frmContext.idExistingApplication.value = "";
		frmContext.idLaunchCustomerNumber.value = "";
		<% /* PSC 12/10/2005 MAR57 - End */ %>
		
		<% /* MAR324 */ %>
		frmContext.idAllowExitFromWrap.value = "";
		<% /* MAR324 - End */ %>

<%		// AY 25/02/00 - Hide the navigation menu
		// AY 04/04/00 - The navigation menu title is now in this file
%>		window.parent.frames("navigation").document.all("divNavigation").style.visibility = "hidden";
		<% /* divNavigation.style.visibility = "hidden"; */ %>
	}

	return bOK;
}
<% /* MAR1587 M Heys 13/04/2006 */ %>
function wrapUponbeforeunload()
{
	var ArrayArguments = new Array();
	var nWidth;
	var nHeight;
	var wrapUpRequestArray;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	wrapUpRequestArray = XML.CreateRequestAttributeArray(window);
	XML = null;
	
	ArrayArguments[0] = frmContext.idApplicationNumber.value;
	ArrayArguments[1] = frmContext.idApplicationFactFindNumber.value;
	ArrayArguments[2] = wrapUpRequestArray;
	ArrayArguments[3] = frmContext.idCustomerNumber1.value;
	ArrayArguments[4] = frmContext.idCustomerVersionNumber1.value;
	ArrayArguments[5] = frmContext.idCustomerName1.value;
	nWidth = 626;
	nHeight = 590;
	
	var rtnValue = scScrFunctions.DisplayPopup(window, document, "DC360.asp", ArrayArguments, nWidth, nHeight); 
	if (rtnValue == "exitOmiga")
	{
		// close Omiga 
		window.top.close();
	}
}
<% /* MAR1587 M Heys 13/04/2006  End*/ %>
</script>
</body>
</html>


