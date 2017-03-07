<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      cr030.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Customer Application Summary screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		04/02/00	Optimised version
AY		11/02/00	Change to msgButtons button types
AY		01/03/00	SYS0019 - All forenames now displayed
					SYS0243 - Address column and associated processing removed
					SYS0266 - CustomerOrder context field was being set
					to CustomerVersionNumber
AY		02/03/00	SYS0243 (revisited) table cell reference for role in delete function was incorrect
SR		10/03/00	modified method btnDelete.onclick()-- Added CustomerRoleCode to 
					CUSTOMERROLE Tag of the REQUEST 
AY		29/03/00	New top menu/scScreenFunctions change
SR		17/05/00	SYS0748 (Points [1],[4],[5]
MC		31/05/00	SYS0026 Disable buttons on major error in context
SR		01/06/00	SYS0748 Points [4],[5] Disable ADD button, if maximum number of Applicants
					already exist.
CL		05/03/01	SYS1920 Read only functionality added
PSC		13/03/01	SYS1930 Disable the ability to add remove and reorder customers if the 
                    Application Stage is on or after the stage in global parameters 
DJP		06/08/01	SYS2564/SYS2738 (child) Client specific cosmetic customisation					
SG		06/12/01	SYS3357 Make screen read only if application is Cancelled or Declined.
JLD		10/12/01	SYS2806 Routing for completeness check
LD		23/05/02	SYS4727 Use cached versions of frame functions
STB		29/05/02	SYS2217 Only progress if one or more applicants are setup.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		AQR			Description
MV		09/05/2002	BMIDS00004	Modfied SetButtonStates() and Added CheckTypeOfMortgageValidationList()
MV		14/05/2002	BMIDS00004	Code Review Errors - Modfied SetButtonStates() 
								and Renamed CheckTypeOfMortgageValidationList() to IsPurchaseOFEquityOrSecuredLoan()
MV		15/05/2002	BMIDS00004  Code Review Errors - Modfied SetButtonStates()
GHun	08/08/2002	BMIDS00006	CAWP1 BM054 Customer Account Download
GHun	09/09/2002	BMIDS00425	ImportAccountsIntoApplication requires extra values passed in and error checking
GHun	16/09/2002	BMIDS00459	Only call ImportAccountsIntoApplication if there is existing business
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
PSC		15/10/2002	BMIDS00464	Disable Delete button if further advance
PSC		23/10/2002	BMIDS00465	Amend when account data is imported
SA		29/10/2002	BMIDS00686	Make screen editable if application stage is Cancelled/Declined.
GD		15/11/2002	BMIDS00376	make sure buttons disabled if readonly
GHun	20/11/2002	BMIDS01014	ImportAccountsIntoApplication not always called when it should be
GHun	04/12/2002	BM0125		ImportAccountsIntoApplication should be called regardless of MetaAction
MV		10/02/2003	BM0337		Amended btnSubmit.onclick();
MV		17/03/2003	BM0337		Amended btnSubmit.onclick()
LDM		24/06/2003  BMIDS586    Make sure btns enabled/disabled on status properly
PJO     03/07/2003  BM0006      Allow guarantor to be added after 4 applicants
KRW     25/03/2004  BMIDS586    Changed processing for enabling screen buttons when sequence number is 10 + 
MC		30/04/2004	BM0468		Cancel and Decline stage freeze screens functionality
HMA     30/06/2004  BMIDS758    CC069  On delete, update RemovedTOECustomer table.
GHun	05/10/2004	BMIDS907	Display a status message when importing accounts
HMA     22/11/2004  BMIDS600    Use Global Parameter for maximum number of customers.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:

Prog	Date		AQR			Description
PSC		18/10/2005	MAR57     	Added processing to synchronise customers
PSC		25/10/2005	MAR300		Correct call and response from GetAndSyncCustomerDetails
PSC		09/01/2006	MAR1001		Add OtherSystemCustomerNumber into context
PSC		09/03/2006	MAR1353		Don't synchronise customer data if the application is in an 
								exception stage or at completion stage
LDM		19/02/2006	EP2_1394	Make the type of mortgage a drop down combo
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<% /* Scriptlets */ %>
<% /* CORE UPGRADE 702 <object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object> */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>

<% /* Specify Forms Here */ %>
<form id="frmSubmit" method="post" action="" STYLE="DISPLAY: none"></form>
<form id="frmAddCustomer" method="post" action="cr010.asp" STYLE="DISPLAY: none"></form>
<form id="frmEditCustomer" method="post" action="cr020.asp" STYLE="DISPLAY: none"></form>
<form id="frmAssignCustomers" method="post" action="cr050.asp" STYLE="DISPLAY: none"></form>
<form id="frmApplicationMenu" method="post" action="mn060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen">
<div style="TOP: 60px; LEFT: 10px; HEIGHT: 250px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Type of Mortgage Application:
		<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
			<select id="cboTypeOfMortgage" style="WIDTH: 180px" class="msgCombo" ></select>
		</span>
	</span>
<!-- #include FILE="Customise/cr030CustomiseTable.asp" -->	

	<span id="spnButtons" style="TOP: 220px; LEFT: 4px; POSITION: ABSOLUTE">
		<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
			<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="TOP: 0px; LEFT: 64px; POSITION: ABSOLUTE">
			<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="TOP: 0px; LEFT: 128px; POSITION: ABSOLUTE">
			<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="TOP: 0px; LEFT: 416px; POSITION: ABSOLUTE">
			<input id="btnAssign" value="Assign Applicants & Guarantors" type="button" style="WIDTH: 180px" class="msgButton">
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
<!-- #include FILE="Customise/cr030Customise.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cr030Attribs.asp" -->
<script language="JScript">
<!--
<% /* main XML object */ %>
var ListXML;
var scScreenFunctions;

var m_bFurtherAdvance = false;
var m_bLegacyCustomer = false;

<% /* form frmContext information */ %>
var m_sDistributionChannelId = null;
var m_sMetaAction = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sMortgageApplicationValue = null;
var m_sMortgageApplicationDescription = null;
var m_sReadOnly = null;
var m_sPackageNumber = null;

var m_iNoOfCustomers = null ;
var m_iNoOfApplicants = null;
var m_iMaxCustomers = null ;
var m_iMaxApplicants = null ;
var m_iStageNumber = null;
var m_iDisableApplicantsStageNo = null;
var m_iDisableReorderStage = null;
var m_blnReadOnly = false;
<% /* BMIDS00459 */ %>
var m_blnHasExistingAccounts = false;
var m_sOtherSystemAccountNumber = "";
<% /* BMIDS00459 End */ %>
<% /* PSC 18/10/2005 MAR57 - Start */ %>
var m_blnUseAdminGetAccountDetails = false
var m_blnUseAdminGetCustDetailAppAccess = false;
<% /* PSC 25/10/2005 MAR300 */ %>
var m_sExistingApplication = "";
var m_sActivityId = "";
var m_sApplicationPriority = "";
var m_sLaunchCustomerNumber = "";
<% /* PSC 18/10/2005 MAR57 - End */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
var m_TypeMortCmboIndex = -1;

function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
<%	/*	Make the required buttons available on the bottom of the screen
		(see msgButtons.asp for details)
		APS UNIT TEST REF 2 - Added ApplicationFactFindNumber to logic test

		access the list of validation types from the combogroup entry and
		set the further advance flag
	*/
%>	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	/* CORE UPGRADE scScreenFunctions = new scScrFunctions.ScreenFunctionsObject(); */
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	FW030SetTitles("Customer Application Summary","CR030",scScreenFunctions);

	Customise();
	scTable.initialise(tblTable, 0, "");
	RetrieveContextData();
	if ((m_sApplicationNumber == "") || (m_sApplicationFactFindNumber == ""))
	{
		alert("Application/FactFind number is not known. Please report the error to the helpdesk.");
		DisableButtonsOnError();
	}	
	else 
	{
		GetMaximumApplicantsData() ;
		FindCustomersForApplication();
	}
	
<%	/* SR 17/05/00 - SYS0748: In Read-only mode, set the focus to OK button	*/ %>
	if(m_sReadOnly == "1") document.all("btnSubmit").focus() ;
	
<%	// APS UNIT TEST REF 102 10/09/99
%>	scScreenFunctions.SetContextParameter(window, "idCustomerLockIndicator", "0");

	scScreenFunctions.SetContextParameter(window, "idCustomerLockIndicator", "0");

	CheckMortgageTypeValidationList();
	FindLegacyCustomer();
	<% /* PSC 18/10/2005 MAR57 */ %>
	GetAdminSystemSwitches();

<%	// PSC 13/03/01 - SYS1930 Get the stage at which amending applicants is disabled  
%>	GetDisableApplicantsStage();
	
<% /* SR 17/05/00 - SR0748: On initialisation, set focus to first line */ %>
	if ( parseIntSafe(m_iNoOfCustomers) >= 1)
	   scTable.setRowSelected(1);

	SetButtonStates();
	//if((scScreenFunctions.GetContextParameter(window,"idStageId",0) == 910)||(scScreenFunctions.GetContextParameter(window,"idStageId",0) == 920))
	if((scScreenFunctions.GetContextParameter(window,"idStageId",0) == scScreenFunctions.getCancelledStageValue(window))||(scScreenFunctions.GetContextParameter(window,"idStageId",0) == scScreenFunctions.getDeclinedStageValue(window)))
	{
		scScreenFunctions.SetContextParameter(window,"idCancelDeclineFreezeDataIndicator","1")
	}
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CR030");
	<% //GD BMIDS00376 START %>
	if (m_blnReadOnly == true)
	{
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnAssign.disabled = true;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
	}
	<% //GD BMIDS00376 END %>	
	// Added by automated update TW 09 Oct 2002 SYS5115
	
	GetTypeOfMortgageList();
	
	ClientPopulateScreen();
}

function DisableButtonsOnError()
{
	<% /* SYS0026 Disable buttons on major error in context */ %>
	frmScreen.btnAdd.disabled = true;
	frmScreen.btnAssign.disabled = true;
	frmScreen.btnDelete.disabled = true;
	frmScreen.btnEdit.disabled = true;
	DisableMainButton("Submit");
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function spnTable.onclick()
{
<%	// APS UNIT TEST REF 72 - ReadOnly processing
%>	
<%// GD BMIDS00376  %>
if (m_blnReadOnly == false)
{
	if (scTable.getRowSelectedId() != null)
		{
			frmScreen.btnEdit.disabled = false;
			<% /* PSC 15/10/2002 BMIDS00464 */ %>
			if ((m_sReadOnly != "1") && !IsFurtherAdvance() 
			 && parseIntSafe(m_iStageNumber) < parseIntSafe(m_iDisableApplicantsStageNo)) // BMIDS586 KRW 25/03/04
			     frmScreen.btnDelete.disabled = false;
   		}
}
}

function spnTable.ondblclick()
{
	<% //GD BMIDS00376 START %>	
	if (m_blnReadOnly == false)
	{
		if (scTable.getRowSelectedId() != null)
		{
			frmScreen.btnEdit.onclick();
		}
		<% //GD BMIDS00376 END %>		
	}	
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","CreateNewCustomerForNewApplication");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","rfapp01");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	m_sMortgageApplicationDescription = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationDescription","Further Advance");
	m_sMortgageApplicationValue = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","3");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sPackageNumber = scScreenFunctions.GetContextParameter(window,"idPackageNumber","0");
	m_iStageNumber = scScreenFunctions.GetContextParameter(window,"idCurrentStageOrigSeqNo","0");
	
	<% /* BMIDS01014 */ %>
	m_sOtherSystemAccountNumber = scScreenFunctions.GetContextParameter(window,"idOtherSystemAccountNumber",null);
	<% /* PSC 18/10/2005 MAR57 - Start */ %>
	<% /* PSC 25/10/2005 MAR300 */ %>
	m_sExistingApplication = scScreenFunctions.GetContextParameter(window,"idExistingApplication","0");
	m_sLaunchCustomerNumber = scScreenFunctions.GetContextParameter(window,"idLaunchCustomerNumber","");
	m_sActivityId = scScreenFunctions.GetContextParameter(window,"idActivityId","")
	m_sApplicationPriority = scScreenFunctions.GetContextParameter(window,"idApplicationPriority","")
	<% /* PSC 18/10/2005 MAR57 - End */ %>
	
	
	<% /* BMIDS01014 No longer required
	// BMIDS00459
	var strTempXML = scScreenFunctions.GetContextParameter(window,"idXML","");
	if (strTempXML.length > 0)
	{
		var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		TempXML.LoadXML(strTempXML);
		TempXML.SelectTag(null, "CUSTOMER");
		if (TempXML.GetTagText("HASEXISTINGACCOUNTS") == "true")
			m_blnHasExistingAccounts = true;
		else
			m_blnHasExistingAccounts = false;
		m_sOtherSystemAccountNumber = TempXML.GetTagText("OTHERSYSTEMACCOUNTNUMBER");
	}
	else
		m_blnHasExistingAccounts = false;
	// BMIDS00459 End
	// BMIDS01014 End */ %>

	if (m_sReadOnly == "0")
	{
		<% /* If the Screen is not already read only, check the application status.
		Cancelled or Declined applications should be read only.*/ %>
		<%/*BMIDS00686 - screen should be editable if application status is declined/cancelled!!*/%>
		//XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		/* CORE UPGRADE 702 XML = new scXMLFunctions.XMLObject(); */
		//var CancelledStageValue = XML.GetGlobalParameterAmount(document, "CancelledStageValue");
		//var DeclinedStageValue = XML.GetGlobalParameterAmount(document, "DeclinedStageValue");
		//var IdStageID = scScreenFunctions.GetContextParameter(window,"idStageID",null);

		//if (IdStageID == CancelledStageValue)
		//{
			<% /* Application is Cancelled. */%>
		<%/*	m_sReadOnly = "1";
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			scScreenFunctions.SetContextParameter(window,"idReadOnly","1");		
		}
		*/%>	
		//if (IdStageID == DeclinedStageValue)
		//{
			<% /* Application is Declined. */ %>
		//	m_sReadOnly = "1";
    	//	scScreenFunctions.SetScreenToReadOnly(frmScreen);
		//	scScreenFunctions.SetContextParameter(window,"idReadOnly","1");		
		//}
	}
}

function SetButtonStates()
{
	<% /* Disables the Add, Edit, Delete and Assign buttons dependent on specified criteria */ %>	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_iDisableReorderStage =XML.GetGlobalParameterAmount(document, "DisableReorderStage");
	
	if (parseIntSafe(m_iStageNumber) >= parseIntSafe(m_iDisableApplicantsStageNo))
	{
		frmScreen.btnAdd.disabled = true;
//		frmScreen.btnAssign.disabled = true;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
	}
	else if (((IsFurtherAdvance()) && (IsLegacyCustomer())) || (m_sReadOnly == "1"))
	{
		if (IsPurchaseOfEquityOrSecuredLoan()) 
		{
			frmScreen.btnAdd.disabled = false;
			frmScreen.btnAssign.disabled = true;
			frmScreen.btnDelete.disabled = false;
		}
		else
		{
			frmScreen.btnAdd.disabled = true;
			frmScreen.btnAssign.disabled = true;
			frmScreen.btnDelete.disabled = true;
		}
	}

	 // BMIDS586 KRW 25/03/04
	if (parseIntSafe(m_iStageNumber) >= parseIntSafe(m_iDisableReorderStage))
	{
		frmScreen.btnAssign.disabled = true;
	}	
	     
	if (!frmScreen.btnAdd.disabled) frmScreen.btnAdd.focus();

	if (scTable.getRowSelectedId() == null)
	{
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
	}
}

function IsPurchaseOfEquityOrSecuredLoan()
{
	var bSuccess = false;
	var ValidationList = new Array(2);
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	ValidationList[0] = "S";
	ValidationList[1] = "P";
	bSuccess = XML.IsInComboValidationList(document,"TypeOfMortgage",m_sMortgageApplicationValue, ValidationList);
	
	ValidationList = null;
	XML = null;	
	
	return bSuccess;
}

function DisableOrEnableAddButton()
{
	<% /* Disbale the ADD button, if the number of applicants >= Max Allowed */ %>
	<% /* PJO BM0006 Allow 5 adds always - the fifth will be a guarantor if the first 4 are applicants */ %>
	<% /* if ( parseIntSafe(m_iNoOfCustomers) >= parseIntSafe(m_iMaxCustomers) || parseIntSafe(m_iNoOfApplicants) >= parseIntSafe(m_iMaxApplicants)) */ %>
	if ( parseIntSafe(m_iNoOfCustomers) >= parseIntSafe(m_iMaxCustomers))
	{
		frmScreen.btnAdd.disabled = true ;
	}
	else
		if (!((IsFurtherAdvance()) && (IsLegacyCustomer())) || (m_sReadOnly == "1"))
			frmScreen.btnAdd.disabled = false ;
}

function GetMaximumApplicantsData()
{
	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject();*/
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	m_iMaxCustomers = XML.GetGlobalParameterAmount(document, "MaximumCustomers") ;
	m_iMaxApplicants = XML.GetGlobalParameterAmount(document, "MaximumApplicants") ;
	
	XML = null ;
}

function CheckMortgageTypeValidationList()
{
<%	/*	Checks the TypeOfMortgage for the entry specified by the Valueid. The Validation entries are then
		checked against the validation list supplied. Here we flag if the application is a FurtherAdvance(F).

		Set the further advance flag so we don't have to redo the call
	*/
%>	var ValidationList = new Array(1);
	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject();*/
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	ValidationList[0] = "F";
	SetFurtherAdvance(XML.IsInComboValidationList(document,"TypeOfMortgage",m_sMortgageApplicationValue, ValidationList));
	ValidationList = null;
	XML = null;
}

function SetFurtherAdvance(bFurtherAdvance)
{
	m_bFurtherAdvance = bFurtherAdvance;
}

function IsFurtherAdvance()
{
	return (m_bFurtherAdvance == true);
}
		
function SetLegacyCustomer(bLegacyCustomer)
{
	m_bLegacyCustomer = bLegacyCustomer;
}

function IsLegacyCustomer()
{
	return (m_bLegacyCustomer == true);
}

function FindLegacyCustomer()
{
	<% /* Checks the Legacy Customer global parameter */ %>
	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject();*/
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	SetLegacyCustomer(XML.GetGlobalParameterBoolean(document,"FindLegacyCustomer"));
	XML = null;
}

function frmScreen.btnAdd.onclick()
{
	<% /* BMIDS01014 Clear OtherSystemCustomerNumber before adding a new customer */ %>
	scScreenFunctions.SetContextParameter(window, "idOtherSystemCustomerNumber", null);
	<% /* BMIDS01014 End */ %>
	frmAddCustomer.submit();
}

function frmScreen.btnEdit.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","UpdateExistingCustomer");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber",GetCustomerNumber());
	scScreenFunctions.SetContextParameter(window,"idCustomerName",GetCustomerName());
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",GetCustomerVersionNumber());
	frmEditCustomer.submit();
}

function frmScreen.btnDelete.onclick()
{	
	var bAllowDelete = false;
	var sRoleCode ;
	var sRole = tblTable.rows(scTable.getRowSelectedId()).cells(3).innerText;
	var sSurname = tblTable.rows(scTable.getRowSelectedId()).cells(0).innerText;	
	
<%	// UNIT TEST REF 4
%>	if (sRole == "Applicant") sRoleCode = "1";
	else if (sRole == "Guarantor") sRoleCode = "2";
	else sRoleCode = "" ;
	
	if (sRole == "Applicant")
	{
		for(var nLoop = 1;nLoop <= 10 && !bAllowDelete;nLoop++)
			if(nLoop != scTable.getRowSelected())
<%				// UNIT TEST REF 4
%>				if(tblTable.rows(nLoop).cells(3).innerText == "Applicant") bAllowDelete = true;
	}
	else bAllowDelete = true;

	if(bAllowDelete)
	{
		if (confirm("Are you sure?"))
		{
			/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject();*/
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTag(window,"DELETE");
			XML.CreateActiveTag("CUSTOMERPACKAGE"); 
			XML.CreateTag("CUSTOMERNUMBER",GetCustomerNumber());
			XML.CreateTag("CUSTOMERVERSIONNUMBER",GetCustomerVersionNumber());
			XML.CreateTag("CUSTOMERROLETYPE",sRoleCode);
			XML.CreateTag("CUSTOMERORDER",null);
			XML.CreateTag("PACKAGENUMBER",m_sPackageNumber);
			XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
			XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
			XML.CreateTag("TYPEOFAPPLICATION",m_sMortgageApplicationValue);
			/* BMIDS758 Add tags for creating an entry in the RemovedToECustomer table */
			XML.CreateTag("CIFNUMBER",GetOtherSystemCustomerNumber());
			XML.CreateTag("OMIGACUSTOMERNUMBER",GetCustomerNumber());
			XML.CreateTag("TITLE",GetTitle());
			XML.CreateTag("FIRSTFORENAME",GetFirstForename());
			XML.CreateTag("SECONDFORENAME",GetSecondForename());
			XML.CreateTag("OTHERFORENAMES",GetOtherForenames()   );
			XML.CreateTag("SURNAME",sSurname);
			
			// 			XML.RunASP(document, "DeletecustomerFromApplication.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "DeletecustomerFromApplication.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

			if(XML.IsResponseOK())
			{
				ListXML.ActiveTag = null;
				ListXML.CreateTagList("CUSTOMERROLE");

				var nRowSelected = scTable.getRowSelected();
				ListXML.SelectTagListItem(nRowSelected-1);
				if(ListXML.ActiveTagList.length == nRowSelected) nRowSelected = nRowSelected - 1;

				ListXML.RemoveActiveTag();
				PopulateTable();
				
				<%/* Set the focus to first row of the table, after deletion of a row */%>
				if ( parseIntSafe(m_iNoOfCustomers) >= 1)
				scTable.setRowSelected(1);
			}
			XML = null;
		}
	}
	else 
		alert("Unable to Delete. There must be at least one applicant on the application");
}

function InitialiseContext()
{
	for (var iCustomerIndex=1; iCustomerIndex<=m_iMaxCustomers; iCustomerIndex++)  // BMIDS600
	{
		scScreenFunctions.SetContextParameter(window,"idCustomerName" + iCustomerIndex, "");
		scScreenFunctions.SetContextParameter(window,"idCustomerNumber" + iCustomerIndex, "");
		scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber" + iCustomerIndex, "");
		scScreenFunctions.SetContextParameter(window,"idCustomerRoleType" + iCustomerIndex, "");
		scScreenFunctions.SetContextParameter(window,"idCustomerOrder" + iCustomerIndex, "");
		<% /* PSC 09/01/2006 MAR1001 */ %>
		scScreenFunctions.SetContextParameter(window,"idOtherSystemCustomerNumber" + iCustomerIndex, "");
	}
}

function GetCustomerVersionNumber()
{
	<%/* the customer version number is stored as 'hidden' attribute of each row */ %>
	return tblTable.rows(scTable.getRowSelectedId()).getAttribute("CustomerVersionNumber");
}

function GetCustomerNumber()
{
	<%/* the customer number is stored as 'hidden' attribute of each row */ %>	
	return tblTable.rows(scTable.getRowSelectedId()).getAttribute("CustomerNumber");
}

function GetCustomerName()
{
	<%/* the customer name is stored as 'hidden' attribute of each row */%>
	return tblTable.rows(scTable.getRowSelectedId()).getAttribute("CustomerName");
}

<% /* BMIDS758 Start 
      Add functions to get Title, Names and OtherSystemCustomerNumber which are all stored as 'hidden' attributes of each row */ %>
function GetTitle()
{
	return tblTable.rows(scTable.getRowSelectedId()).getAttribute("Titlex");
}
function GetFirstForename()
{
	return tblTable.rows(scTable.getRowSelectedId()).getAttribute("FirstForename");
}
function GetSecondForename()
{
	return tblTable.rows(scTable.getRowSelectedId()).getAttribute("SecondForename");
}
function GetOtherForenames()
{
	return tblTable.rows(scTable.getRowSelectedId()).getAttribute("OtherForenames");
}
function GetOtherSystemCustomerNumber()
{
	return tblTable.rows(scTable.getRowSelectedId()).getAttribute("OtherSystemCustomerNumber");
}
<% /* BMIDS758 End */ %>


function frmScreen.btnAssign.onclick()
{
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("RESPONSE");
	ListXML.SelectTagListItem(0);

	<%/* put the response data into the context form */%>	
	scScreenFunctions.SetContextParameter(window,"XML",ListXML.ActiveTag.xml);
	frmAssignCustomers.submit();
}

function btnSubmit.onclick()
{	
	frmScreen.style.cursor = "wait";
	DisableMainButton("Submit");
	/* CORE UPGRADE 702 if ((m_sMetaAction == "CreateNewCustomerForNewApplication") || (m_sMetaAction == "CreateExistingCustomerForNewApplication"))
	{
		// FIXME: call ApplicationBO.UpdateCorrespondenceSalutaion
	}
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmApplicationMenu.submit(); */
	<% /* SYS2217 - Only progress if one or more applicants are setup. */ %>
	if (m_iNoOfApplicants > 0)
	{
		if (m_sReadOnly != "1")
		{

			<% /* PSC 23/10/2002 BMIDS00465 */ %>
			<% /* BM0125 ImportAccountsIntoApplication should be called no matter what MetaAction is set
			if ((m_sMetaAction == "CreateNewCustomerForNewApplication") || 
				(m_sMetaAction == "CreateExistingCustomerForNewApplication") ||
				(m_sMetaAction == "UpdateExistingCustomer") ||
				(m_sMetaAction == "CreateExistingCustomerForExistingApplication") ||
				(m_sMetaAction == "CreateNewCustomerForExistingApplication"))
			{
			*/ %>
				<% // FIXME: call ApplicationBO.UpdateCorrespondenceSalutaion %>

				<% /* BMIDS01014 ImportAccountsIntoApplication should be called for all admin system 
				customers on the application, not just if the first customer is an admin system customer */ %>
				<% /* BMIDS00006 CAWP1 BM054 */ %>
				<% /* var sOtherSystemCustomerNumber = scScreenFunctions.GetContextParameter(window,"idOtherSystemCustomerNumber",""); */ %>
				<% /* BMIDS00459  if (sOtherSystemCustomerNumber != "") */ %>
				<% /* if ((sOtherSystemCustomerNumber != "") && m_blnHasExistingAccounts) */ %>
			<% /* PSC 18/10/2005 MAR57 - Start */ %>
		
			<% /* PSC 25/10/2005 MAR300 */ %>
			if (m_sExistingApplication == "1" && m_blnUseAdminGetCustDetailAppAccess)
			{
				<% /* PSC 09/03/2006 MAR1353 - Start */ %>
				var blnSynchCustomerData = true;
				
				var sActivityId = scScreenFunctions.GetContextParameter(window,"idActivityId", ""); 
		        var sStageId = scScreenFunctions.GetContextParameter(window,"idStageId", "");
				var ParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				var sCompletionStageID = ParamXML.GetGlobalParameterString(document,"TMCompletionsStageId");				

				if (sStageId != sCompletionStageID)
				{		
					var XMLStages = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					XMLStages.CreateRequestTag(window,"FINDSTAGELIST");
					XMLStages.CreateActiveTag("STAGE");
					XMLStages.SetAttribute("ACTIVITYID",sActivityId);
					XMLStages.SetAttribute("EXCEPTIONSTAGEINDICATOR","1");
					switch (ScreenRules())
					{
						case 1: // Warning
						case 0: // OK
							XMLStages.RunASP(document,"FindStageList.asp");
							break;
						default: // Error
							XMLStages.SetErrorResponse();
					}

					var ErrorTypes = new Array("RECORDNOTFOUND");
					var ErrorReturn = XMLStages.CheckResponse(ErrorTypes);
					if(ErrorReturn[0] && ErrorReturn[1] != ErrorTypes[0])
					{
						XMLStages.SelectSingleNode ("//RESPONSE/STAGE[@STAGEID='" + sStageId + "']");
						
						if (XMLStages.ActiveTag != null)
							blnSynchCustomerData = false;
					}
				}
				else
				{
					blnSynchCustomerData = false;
				}
				
				if (blnSynchCustomerData)
					SynchroniseCustomerData();
				<% /* PSC 09/03/2006 MAR1353 - End */ %>		
			}
		
			if (m_blnUseAdminGetAccountDetails)
				ImportAccountsIntoApplication();
			<% /* PSC 18/10/2005 MAR57 - End */ %>			
				<% /* BMIDS00459 End */ %>
				<% /* BMIDS00006 End */ %>
				<% /* BMIDS01014 End */ %>
			<% /* BM0125 } */%>
		}		
		ClearLaunchContext();
		
		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else
			frmApplicationMenu.submit();
	}
	else
	{
		alert("This application must have a customer assigned as an applicant.  Please go to screen CR050 to reassign the customer roles.");
	}
	
	EnableMainButton("Submit");
	frmScreen.style.cursor = "default";
}

function FindCustomersForApplication()
{
	/* CORE UPGRADE 702 ListXML = new scXMLFunctions.XMLObject(); */
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(window,"SEARCH");
	ListXML.CreateActiveTag("APPLICATION");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	ListXML.RunASP(document,"FindCustomersForApplication.asp");

	if (ListXML.IsResponseOK()) PopulateTable();
}

function PopulateTable()
{
	scTable.clear();
	ListXML.ActiveTag = null;
	var TagListCUSTOMERROLE = ListXML.CreateTagList("CUSTOMERROLE");
	
	<% /* Store Number of customers in Public variable. */ %>
	m_iNoOfCustomers = ListXML.ActiveTagList.length;
	m_iNoOfApplicants = 0 ;

	InitialiseContext();

	for (var i0=0; i0<m_iNoOfCustomers; i0++)
	{
		ListXML.ActiveTagList = TagListCUSTOMERROLE;
		ListXML.SelectTagListItem(i0);

		var sCustomerNumber = ListXML.GetTagText("CUSTOMERNUMBER");
		var sCustomerVersionNumber = ListXML.GetTagText("CUSTOMERVERSIONNUMBER");
		var sCustomerRoleType = ListXML.GetTagText("CUSTOMERROLETYPE");
		var sCustomerOrder = ListXML.GetTagText("CUSTOMERORDER");
		var sOtherSystemCustomerNumber = ListXML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER");

		<% /* Change from role numbers to a meaningful description of the customers role */ %>	
		if (sCustomerRoleType == "1")
		{
			sRoleTypeDescription = "Applicant";
			m_iNoOfApplicants++ ;
		}
		else if (sCustomerRoleType == "2") sRoleTypeDescription = "Guarantor";
		else sRoleTypeDescription = "Error";

		ListXML.CreateTagList("CUSTOMERVERSION")
		ListXML.SelectTagListItem(0);

		var sSurname = ListXML.GetTagText("SURNAME");
		var sFirstForename = ListXML.GetTagText("FIRSTFORENAME");
		var sSecondForename = ListXML.GetTagText("SECONDFORENAME");                        // BMIDS758
		var sOtherForenames = ListXML.GetTagText("OTHERFORENAMES");                        // BMIDS758
		
		var sDOB = ListXML.GetTagText("DATEOFBIRTH");

		var sForenames = sFirstForename + " " + sSecondForename + " " + sOtherForenames;   // BMIDS758
		
		var sTitle = ListXML.GetTagText("TITLE");                                          // BMIDS758

		<% /* Populate the Customer Name, Customer Number and CustomerVersion Number
		in the context where the maximum number of Customers is currently set to 5  */ %>		
		if (i0 < m_iMaxCustomers)                                                          // BMIDS600
		{
			var nCustomerIndex = i0 + 1;
			var sCustomerName = sFirstForename + " " + sSurname;

			scScreenFunctions.SetContextParameter(window,"idCustomerName" + nCustomerIndex, sCustomerName);
			scScreenFunctions.SetContextParameter(window,"idCustomerNumber" + nCustomerIndex, sCustomerNumber);
			scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber" + nCustomerIndex, sCustomerVersionNumber);
			scScreenFunctions.SetContextParameter(window,"idCustomerRoleType" + nCustomerIndex, sCustomerRoleType);
			scScreenFunctions.SetContextParameter(window,"idCustomerOrder" + nCustomerIndex, sCustomerOrder);
			<% /* PSC 09/01/2006 MAR1001 */ %>
			scScreenFunctions.SetContextParameter(window,"idOtherSystemCustomerNumber" + nCustomerIndex, sOtherSystemCustomerNumber);
		}

		ShowRow(i0+1,sSurname,sForenames,sDOB,sRoleTypeDescription,sCustomerNumber,sCustomerVersionNumber,sCustomerRoleType,sOtherSystemCustomerNumber,sTitle,sFirstForename,sSecondForename,sOtherForenames); // BMIDS758
	}

	<%/* Disable ADD button, if number of customers >= maximum allowed */ %>
	DisableOrEnableAddButton();
}

function ShowRow(nIndex,sSurname,sForenames,sDOB,sCustomerRoleTypeDesc,sCustomerNumber,sCustomerVersionNumber,sCustomerRoleType,sOtherSystemCustomerNumber,sTitle,sFirstForename,sSecondForename,sOtherForenames)
{
	<%	// Set the table fields %>	
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(0),sSurname);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(1),sForenames);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(2),sDOB);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(3),sCustomerRoleTypeDesc);
	<%	// Set the invisible context for each row %>	
	tblTable.rows(nIndex).setAttribute("CustomerNumber", sCustomerNumber);
	tblTable.rows(nIndex).setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
	tblTable.rows(nIndex).setAttribute("CustomerName", sForenames + " " + sSurname);
	<% /* BMIDS00006 CAWP1 */ %>
	tblTable.rows(nIndex).setAttribute("OtherSystemCustomerNumber", sOtherSystemCustomerNumber);
	tblTable.rows(nIndex).setAttribute("CustomerRoleType", sCustomerRoleType);
	<% /* BMIDS758 */ %>
	tblTable.rows(nIndex).setAttribute("Titlex", sTitle);
	tblTable.rows(nIndex).setAttribute("FirstForename", sFirstForename);
	tblTable.rows(nIndex).setAttribute("SecondForename", sSecondForename);
	tblTable.rows(nIndex).setAttribute("OtherForenames", sOtherForenames);
}

function GetDisableApplicantsStage()
{
	<%	//	Checks the Legacy Customer global parameter %>	
	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject(); */
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	SetDisableApplicantsStage(XML.GetGlobalParameterAmount(document,"DisableApplicantsStage"));
	XML = null;
}

function SetDisableApplicantsStage(iStageNumber)
{
	m_iDisableApplicantsStageNo = iStageNumber;
}

<% /* BMIDS00006 CAWP1 BM054 */ %>
function ImportAccountsIntoApplication()
{
	var blnHaveAdminCustomer = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "IMPORTACCOUNTSINTOAPPLICATION");
	XML.CreateActiveTag("IMPORTACCOUNTSINTOAPPLICATION");
	XML.CreateActiveTag("CUSTOMERLIST");
		
	for (var i=1; i <= m_iNoOfCustomers; i++)
	{	
		<% /* BMIDS01014 Only add admin system customers to the request */ %>
		var sOtherSystemCustomerNumber = tblTable.rows(i).getAttribute("OtherSystemCustomerNumber")
		if (sOtherSystemCustomerNumber.length > 0)
		{
			blnHaveAdminCustomer = true;
		<% /* BMIDS01014 End */ %>
			XML.CreateActiveTag("CUSTOMER");
			XML.CreateTag("CUSTOMERNUMBER", tblTable.rows(i).getAttribute("CustomerNumber"));
			XML.CreateTag("CUSTOMERVERSIONNUMBER", tblTable.rows(i).getAttribute("CustomerVersionNumber"));
			XML.CreateTag("OTHERSYSTEMCUSTOMERNUMBER", sOtherSystemCustomerNumber);
			XML.CreateTag("CUSTOMERROLETYPE", tblTable.rows(i).getAttribute("CustomerRoleType"));
			XML.ActiveTag = XML.ActiveTag.parentNode;
		}
	}
	
	<% /* BMIDS01014 Only call ImportAccountsIntoApplication if there is at least one admin system customer */ %>
	if (blnHaveAdminCustomer)
	{
	<% /* BMIDS01014 End */ %>
		XML.ActiveTag = XML.ActiveTag.parentNode;
		XML.CreateActiveTag("APPLICATION");
		XML.CreateTag("ACCOUNTNUMBER", m_sOtherSystemAccountNumber);
		XML.CreateTag("TYPEOFAPPLICATION", scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue",""));
		<% /* BMIDS00425 */ %>
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		<% /* BMIDS00425 End*/ %>
		<% /* BMIDS907 GHun Display a status message */ %>
		window.status = "Retrieving full account details ...";
		XML.RunASP(document, "ImportAccountsIntoApplication.asp")
		window.status = "";
		<% /* BMIDS907 End */ %>
		XML.IsResponseOK();	<% /* BMIDS00425 */ %>
	}
}
<% /* BMIDS00006 End */ %>

//LDM 24/06/03 BMIDS586 make sure int comparisons are correctly applied
function parseIntSafe(sText)
{
	if (sText=="")
	   return(0);
	else
	   return(parseInt(sText));
}

<% /* PSC 18/10/2005 MAR57 - Start */ %>
function GetAdminSystemSwitches()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_blnUseAdminGetAccountDetails = XML.GetGlobalParameterBoolean(document,"UseAdminGetAccountDetails");
	m_blnUseAdminGetCustDetailAppAccess =  XML.GetGlobalParameterBoolean(document,"UseAdminGetCustDetailAppAccess");
}
function SynchroniseCustomerData()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, null);
	XML.CreateActiveTag("SEARCH");
	XML.SetAttribute("CUSTOMERDATACHECK", "1");
	XML.CreateActiveTag("CUSTOMERLIST");
		
	for (var i=1; i <= m_iNoOfCustomers; i++)
	{	
		var sOtherSystemCustomerNumber = tblTable.rows(i).getAttribute("OtherSystemCustomerNumber")
		
		XML.CreateActiveTag("CUSTOMER");
		XML.SetAttribute("NOLOCK", "1");		
		
		XML.CreateTag("CUSTOMERNUMBER", tblTable.rows(i).getAttribute("CustomerNumber"));
		XML.CreateTag("CUSTOMERVERSIONNUMBER", tblTable.rows(i).getAttribute("CustomerVersionNumber"));
		XML.CreateTag("OTHERSYSTEMCUSTOMERNUMBER", sOtherSystemCustomerNumber);
	}
	XML.SelectTag(null, "SEARCH");;
	XML.CreateActiveTag("CRITICALDATACONTEXT");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.SetAttribute("APPLICATIONPRIORITY", m_sApplicationPriority);
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("ACTIVITYID", m_sActivityId);
	<% /* PSC 25/10/2005 MAR300 */ %>
	XML.SetAttribute("ACTIVITYINSTANCE", "1");
	XML.SetAttribute("STAGEID", m_iStageNumber);
	window.status = "Synchronising customer details ...";
	XML.RunASP(document, "GetAndSyncCustomerDetails.asp");
	<% /* PSC 25/10/2005 MAR300 - Start */ %>
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null, "RESPONSE")
		
		sReadOnly = XML.GetAttribute("READONLY");
		scScreenFunctions.SetContextParameter(window,"idCustomerReadOnly",sReadOnly);
		scScreenFunctions.SetContextParameter(window,"idReadOnly",sReadOnly);
	}
	<% /* PSC 25/10/2005 MAR300 - End */ %>

	window.status = "";
}
function ClearLaunchContext()
{
	scScreenFunctions.SetContextParameter(window, "idLaunchCustomerXML", "");
	<% /* PSC 25/10/2005 MAR300 */ %>
	scScreenFunctions.SetContextParameter(window, "idExistingApplication", "");
	scScreenFunctions.SetContextParameter(window, "idLaunchCustomerNumber", "");
}
<% /* PSC 18/10/2005 MAR57 - End */ %>


<% /* LDM 19/02/2006 EP2_1394 Populate the type of mortgage combo
 Only put in the combo the current type and the allowed types that you can move to.
 This is signified in the validation type for the current mortgage type. Any validation type 
 starting with a @ is a valid valueid you can move to */ %>
function GetTypeOfMortgageList()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	var sGroupList = new Array("TypeOfMortgage");
	var arrValIds = null;
	var nAddedCount = 0;
	var TagListNew = null;

	frmScreen.cboTypeOfMortgage.selectedIndex = -1
	frmScreen.cboTypeOfMortgage.disabled = true;

	if(m_sMortgageApplicationValue != "" && m_sMortgageApplicationDescription != "")
	{
		frmScreen.cboTypeOfMortgage.value = m_sMortgageApplicationValue;

		if(XML.GetComboLists(document,sGroupList))
		{		
			<% /* get all the @ prefixed validation codes for the current type of mortgage and put in list */ %>
			XML.SelectSingleNode("//LISTENTRY[VALUEID='" +  m_sMortgageApplicationValue + "']");
			TagListNew = XML.ActiveTag.getElementsByTagName("VALIDATIONTYPE");
			arrValIds = new Array(TagListNew.length);

			for(var nListLoop = 0;nListLoop < TagListNew.length;nListLoop++)
			{
				if(TagListNew.item(nListLoop).text.substr(0,1) == "@")
				{
					arrValIds[nAddedCount] = new Array (TagListNew.item(nListLoop).text.substr(1));
					nAddedCount++;
				}
			}			

			<% /* reset to the parent */ %>
			XML.ActiveTag = XML.ActiveTag.parentNode;
			XML.CreateTagList("LISTENTRY");

			<% /* Loop through all entries in the comboXML; only add to the combo controll
			      the value ids that match the numbers in the list created above and the current value*/ %>
			nAddedCount = 0;
			for(var iLoop = 0; iLoop < XML.ActiveTagList.length; iLoop++)
			{
				XML.SelectTagListItem(iLoop);
				var sGroupName	= XML.GetTagText("VALUENAME");
				var sValueID = XML.GetTagText("VALUEID");
				TagOPTION = document.createElement("OPTION");
				TagOPTION.value	= sValueID;
				TagOPTION.text = sGroupName;
					
				if(sValueID == m_sMortgageApplicationValue) <% /* Current type */ %>
				{
					frmScreen.cboTypeOfMortgage.add(TagOPTION);
					m_TypeMortCmboIndex = nAddedCount;
					frmScreen.cboTypeOfMortgage.selectedIndex = m_TypeMortCmboIndex;
					nAddedCount ++;
				}
				else
				{
					for(var nListLoop = 0; nListLoop < arrValIds.length && arrValIds[nListLoop]!= null; nListLoop++)
					{
						if (arrValIds[nListLoop] == sValueID) <% /* allowable @ types */ %>
						{
							frmScreen.cboTypeOfMortgage.add(TagOPTION);
							nAddedCount ++; 
						}
					}
				}
			}

			<% /* If not in Read-only mode, and the global parameter AllowTypeOfApplicationChange is true.
				  and the current stage is <= the global parameter TMLastStageTypeOfMtgeChange.	 
			      Only let the user change the combo if there is something there; and they are allowed to change it */ %>
			var bAllowAppChange = false;
			var iLastStageType = -1;
			
			bAllowAppChange = XML.GetGlobalParameterBoolean(document, "AllowTypeOfApplicationChange") ;
			iLastStageType  = XML.GetGlobalParameterAmount(document, "TMLastStageTypeOfMtgeChange") ;

			if(nAddedCount > 0 && 
			   m_sReadOnly == "0" && 
			   bAllowAppChange == true && 
			   m_iStageNumber <= iLastStageType)
			{
				frmScreen.cboTypeOfMortgage.disabled = false;
			}
		}
	}

	XML = null;
	sGroupList = null;
	arrValIds = null;
	TagListNew = null;
}

<% /* LDM 19/02/2006 EP2_1394 */ %>
function frmScreen.cboTypeOfMortgage.onchange()
{
	//GetApplicationData
	var AppXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sQuotationNumber = null;
	
	AppXML.CreateRequestTag(window, "GetAcceptedQuoteData");
	AppXML.CreateActiveTag("APPLICATION");
	AppXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AppXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
			
	AppXML.RunASP(document,"GetAcceptedOrActiveQuoteData.asp");
	if(AppXML.IsResponseOK())
	{
		// EP710 - Ignore NULL quotation node
		if(AppXML.SelectTag(null, "QUOTATION"))
			sQuotationNumber = AppXML.GetTagText("QUOTATIONNUMBER");
		else
			sQuotationNumber = "";
	}
	
	if (sQuotationNumber != null && sQuotationNumber != "")
	{
		if (confirm("Changing the Type of Mortgage Application will require the quote to be remodelled. Do you wish to Continue?"))
		{
			if(CommitChanges())
			{
				m_TypeMortCmboIndex = frmScreen.cboTypeOfMortgage.selectedIndex;
				m_sMortgageApplicationValue = frmScreen.cboTypeOfMortgage.value;
				m_sMortgageApplicationDescription = frmScreen.cboTypeOfMortgage.options(m_TypeMortCmboIndex).text

				ProcessRemodelTask();
				
				<% /* context values for type of mortgage value and description */ %>
				scScreenFunctions.SetContextParameter(window,"idMortgageApplicationDescription", m_sMortgageApplicationDescription);
				scScreenFunctions.SetContextParameter(window,"idMortgageApplicationValue",m_sMortgageApplicationValue);
			}
			else
			{
				frmScreen.cboTypeOfMortgage.selectedIndex = m_TypeMortCmboIndex;
			}
		}
		else
		{
			frmScreen.cboTypeOfMortgage.selectedIndex = m_TypeMortCmboIndex;
		}
	}
	else <% /* no quote so save away */ %>
	{
		if(CommitChanges())
		{
			m_TypeMortCmboIndex = frmScreen.cboTypeOfMortgage.selectedIndex;
			m_sMortgageApplicationValue = frmScreen.cboTypeOfMortgage.value;
			m_sMortgageApplicationDescription = frmScreen.cboTypeOfMortgage.options(m_TypeMortCmboIndex).text
			<% /* context values for type of mortgage value and description */ %>
			scScreenFunctions.SetContextParameter(window,"idMortgageApplicationDescription", m_sMortgageApplicationDescription);
			scScreenFunctions.SetContextParameter(window,"idMortgageApplicationValue",m_sMortgageApplicationValue);
		}
		else
		{
			frmScreen.cboTypeOfMortgage.selectedIndex = m_TypeMortCmboIndex;
		}
	}
	
}

<% /* LDM 19/02/2006 EP2_1394 */ %>
function CommitChanges()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,"UPDATE")
	XML.CreateActiveTag("APPLICATIONFACTFIND");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("TYPEOFAPPLICATION", frmScreen.cboTypeOfMortgage.value); 

	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"UpdateApplicationFactFind.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}

<% /* LDM 19/02/2006 EP2_1394 */ %>
function ProcessRemodelTask()
{
	var ParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var taskId = ParamXML.GetGlobalParameterString(document,"TMRemodelMortgage",null);
	var bTaskCreate	= true;
	
	<% /* get params.. */ %>
	var stageXML =		new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var taskXML =		new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sUserRole =		scScreenFunctions.GetContextParameter(window, "idRole", null);
	var sUserId =		scScreenFunctions.GetContextParameter(window, "idUserID", null)	
	var sUnitId =		scScreenFunctions.GetContextParameter(window, "idUnitID", null)
	var sApplPriority = scScreenFunctions.GetContextParameter(window, "idApplicationPriority", null)
	var sActivityId =	scScreenFunctions.GetContextParameter(window, "idActivityId", null)
	var sStageId =		scScreenFunctions.GetContextParameter(window, "idStageID", null)
	
	<% /* check if exists */ %>
	stageXML.CreateRequestTag(window, "GetCurrentStage");
	stageXML.CreateActiveTag("CASEACTIVITY");
	stageXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	stageXML.SetAttribute("CASEID", m_sApplicationNumber);
	stageXML.SetAttribute("ACTIVITYID", sActivityId);
	stageXML.SetAttribute("ACTIVITYINSTANCE", "1");
	stageXML.RunASP(document, "MsgTMBO.asp");
	
	if(!stageXML.IsResponseOK())
	{
		alert("Error retrieving current stage details.");
		return;
	}
		
	var sTaskSeqNo = stageXML.GetTagAttribute("CASESTAGE", "CASESTAGESEQUENCENO");
	
	<% /* Check if this task already exists and is incomplete, 
		  Only create this task if there is not a incomplete task already there */ %>
	var tagList = stageXML.CreateTagList("CASETASK");			
	
	for (var nItem = 0; nItem < tagList.length; nItem++)
	{
		stageXML.SelectTagListItem(nItem);			
		if(stageXML.GetAttribute("TASKID")==taskId)
		{
			if(stageXML.GetAttribute("TASKSTATUS") == "10" || stageXML.GetAttribute("TASKSTATUS") == "20")
			{
				bTaskCreate = false;
				break;
			}
		}				
	}

	var currentDate = scScreenFunctions.DateTimeToString(scScreenFunctions.GetAppServerDate(true));

	if (bTaskCreate == true) 
	{
		var reqTag = taskXML.CreateRequestTag(window, "CreateAdhocCaseTask");	

		taskXML.SetAttribute("USERID", sUserId);
		taskXML.SetAttribute("UNITID", sUnitId);		
		taskXML.SetAttribute("USERAUTHORITYLEVEL", sUserRole);				
		taskXML.CreateActiveTag("APPLICATION");		
		taskXML.SetAttribute("APPLICATIONPRIORITY", sApplPriority);
		taskXML.ActiveTag = reqTag;	
		taskXML.CreateActiveTag("CASETASK");
		taskXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		taskXML.SetAttribute("CASEID", m_sApplicationNumber);	
		taskXML.SetAttribute("ACTIVITYID", sActivityId);
		taskXML.SetAttribute("ACTIVITYINSTANCE", "1");		
		taskXML.SetAttribute("CASESTAGESEQUENCENO", sTaskSeqNo);
		taskXML.SetAttribute("STAGEID", sStageId);		
		taskXML.SetAttribute("TASKID", taskId);
		taskXML.SetAttribute("TASKDUEDATEANDTIME", currentDate);
		
		taskXML.RunASP(document, "OmigaTmBO.asp");
		taskXML.IsResponseOK();
	}
	
	stageXML = null;
	taskXML = null;
	ParamXML = null;
}

-->
</script>
</body>
</html>

