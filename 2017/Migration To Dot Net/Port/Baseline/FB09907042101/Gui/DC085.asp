<%@ LANGUAGE="JSCRIPT" %>
<html>
	<%
/*
Workfile:      DC085.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Financial Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date			Description
SR 		14/06/2004 		BMIDS772 - New

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog    Date			Description
MF		22/07/2005		IA_WP01 process flow changes
HMA     10/11/2005      MAR172 Correct setting focus on load.
PJO     20/12/2005      MAR825  Take DC195 out of routing on Global Parameter setting 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog    Date			Description

SAB		03/04/2006		EP8 - Routing change
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
	<HEAD>
		<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4" />
		<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
		<style type="text/css">
		<!--
			.hidden {display:none;}
		-->
		</style>
		<title></title>
	</HEAD>
	<body>
		<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1">
		</object>
		<% // PJO 20/12/2005 MAR825 %>
		<object data="scXMLFunctions.asp" height="1" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" 
			tabIndex="-1" type="text/x-scriptlet" width="1" viewastext>
		</object>
		<script src="validation.js" language="JScript"></script>
		<form id="frmToDC080" method="post" action="DC080.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToDC090" method="post" action="DC090.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToDC155" method="post" action="DC155.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToDC070" method="post" action="DC070.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToDC060" method="post" action="DC060.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToDC110" method="post" action="DC110.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToDC120" method="post" action="DC120.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToDC130" method="post" action="DC130.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToDC140" method="post" action="DC140.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToDC150" method="post" action="DC150.asp" STYLE="DISPLAY: none"></form>
		<% // PJO 20/12/2005 MAR825 %>
		<form id="frmToDC160" method="post" action="DC160.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToDC195" method="post" action="DC195.asp" STYLE="DISPLAY: none"></form>		
		<% // SAB 03/04/2006 EP8  %>
		<form id="frmToDC200" method="post" action="DC200.asp" STYLE="DISPLAY: none"></form>
		<form id="frmToDC370" method="post" action="DC370.asp" STYLE="DISPLAY: none"></form>		
		<% /* span field to keep tabbing within this screen */ %>
		<span id="spnToLastField" tabindex="0"></span>
		<form id="frmScreen" mark validate="onchange">
			<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 360px; WIDTH: 604px; POSITION: ABSOLUTE"
				class="msgGroup">				
				<span style="LEFT: 4px; POSITION: absolute; TOP: 0px" class="msgLabel"><strong>Financial 
					Summary</strong>
				</span>
				<div id="rowRegOutgoings" class="hidden">	
					<span style="TOP: 30px; LEFT: 10px; WIDTH: 400px; POSITION: ABSOLUTE" class="msgLabel">
						Do you have any Regular Outgoings?
						<span id="spnRegularOutgoingsInd">
							<span style="TOP: -3px; LEFT: 300px; POSITION: ABSOLUTE">
								<input id="optRegularOutgoingsYes" name="RegularOutgoingsInd" type="radio" value="1">
								<label for="optRegularOutgoingsYes" class="msgLabel">Yes</label>
							</span>
							<span style="TOP: -3px; LEFT: 350px; POSITION: ABSOLUTE">
								<input id="optRegularOutgoingsNo" name="RegularOutgoingsInd" type="radio" value="0">
								<label for="optRegularOutgoingsNo" class="msgLabel">No</label>
							</span>
						</span>
						<span style="LEFT: 400px; POSITION: absolute; TOP: -3px"><input id="btnRegoutgoings" value="Regular Outgoings" type="button" style="WIDTH: 120px" class="msgButton">
						</span>
					</span>
				</div>			
				<div id="rowAnyAccounts" class="hidden">				
					<span style="TOP: 70px; LEFT: 10px; WIDTH: 400px; POSITION: ABSOLUTE" class="msgLabel">
						Have you held any accounts in last 3 years?
						<span id="spnAccountsLast3YearsInd">
							<span style="TOP: -3px; LEFT: 300px; POSITION: ABSOLUTE">
								<input id="optAccountsLast3YearsYes" name="AccountsLast3YearsInd" type="radio" value="1">
								<label for="optAccountsLast3YearsYes" class="msgLabel">Yes</label>
							</span>
							<span style="TOP: -3px; LEFT: 350px; POSITION: ABSOLUTE">
								<input id="optAccountsLast3YearsNo" name="AccountsLast3YearsInd" type="radio" value="0">
								<label for="optAccountsLast3YearsNo" class="msgLabel">No</label>
							</span>
						</span>
						<span style="LEFT: 400px; POSITION: absolute; TOP: -3px"><input id="btnAccountLast3Years" value="Existing Accounts" type="button" style="WIDTH: 120px"
							class="msgButton">
						</span>
					</span>
				</div>
				<div id="rowLiabilities" class="hidden">
					<span style="TOP: 110px; LEFT: 10px; WIDTH: 400px; POSITION: ABSOLUTE" class="msgLabel">
						Do you have loans/liabilities?
						<span id="spnLoansLiabilitiesInd">
							<span style="TOP: -3px; LEFT: 300px; POSITION: ABSOLUTE">
								<input id="optLoansLiabilitiesYes" name="LoansLiabilitiesInd" type="radio" value="1">
								<label for="optLoansLiabilitiesYes" class="msgLabel">Yes</label>
							</span>
							<span style="TOP: -3px; LEFT: 350px; POSITION: ABSOLUTE">
								<input id="optLoansLiabilitiesNo" name="LoansLiabilitiesInd" type="radio" value="0">
								<label for="optLoansLiabilitiesNo" class="msgLabel">No</label>
							</span>
						</span>
						<span style="LEFT: 400px; POSITION: absolute; TOP: -3px"><input id="btnLoansLiabilities" value="Loans Liabilities" type="button" style="WIDTH: 120px"
							class="msgButton">
						</span>
					</span>
				</div>
				<div id="rowCreditHistory" class="hidden">
					<span style="TOP: 150px; LEFT: 10px; WIDTH: 400px; POSITION: ABSOLUTE" class="msgLabel">
						Do you have bank/credit cards?
						<span id="spnBankCreditCardsInd">
							<span style="TOP: -3px; LEFT: 300px; POSITION: ABSOLUTE">
								<input id="optBankCreditCardsYes" name="BankCreditCardsInd" type="radio" value="1">
								<label for="optBankCreditCardsYes" class="msgLabel">Yes</label>
							</span>
							<span style="TOP: -3px; LEFT: 350px; POSITION: ABSOLUTE">
								<input id="optBankCreditCardsNo" name="BankCreditCardsInd" type="radio" value="0">
								<label for="optBankCreditCardsNo" class="msgLabel">No</label>
							</span>
						</span>
						<span style="LEFT: 400px; POSITION: absolute; TOP: -3px"><input id="btnBankCreditCards" value="Bank Credit Cards" type="button" style="WIDTH: 120px"
								class="msgButton">
						</span>
					</span>
				</div>
				<div id="rowAppsDeclined" class="hidden">
					<span style="LEFT: 10px; POSITION: absolute; TOP: 190px; WIDTH: 400px" class="msgLabel">					
						Have you previously had any mortgage applications declined?
						<span id="spnMortgagesDeclinedInd">
							<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
								<input id="optMortgagesDeclinedYes" name="MortgagesDeclinedInd" type="radio" value="1">
								<label for="optMortgagesDeclinedYes" class="msgLabel">Yes</label>
							</span> 
							<span style="LEFT: 350px; POSITION: absolute; TOP: -3px">
								<input id="optMortgagesDeclinedNo" name="MortgagesDeclinedInd" type="radio" value="0">
								<label for="optMortgagesDeclinedNo" class="msgLabel">No</label>
							</span> 
						</span>							
						<span style="LEFT: 400px; POSITION: absolute; TOP: -3px">
							<input id="btnMortgagesDeclined" value="Mortgages Declined" type="button" style="WIDTH: 120px" class="msgButton" NAME="btnMortgagesDeclined">
						</span>							
					</span>						
				</div>	
				<div id="rowArrears" class="hidden">
					<span style="LEFT: 10px; POSITION: absolute; TOP: 230px; WIDTH: 400px" class="msgLabel">
						Do you have any arrears History?
						<span id="spnArrearsHistoryInd">
							<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
								<input id="optArrearsHistoryYes" name="ArrearsHistoryInd" type="radio" value="1">
								<label for="optArrearsHistoryYes" class="msgLabel">Yes</label>
							</span> 
							<span style="LEFT: 350px; POSITION: absolute; TOP: -3px">
								<input id="optArrearsHistoryNo" name="ArrearsHistoryInd" type="radio" value="0">
								<label for="optArrearsHistoryNo" class="msgLabel">No</label>
							</span>
						</span>
						<span style="LEFT: 400px; POSITION: absolute; TOP: -3px">
							<input id="btnArrearsHistory" value="Arrears History" type="button" style="WIDTH: 120px" class="msgButton" NAME="btnArrearsHistory">
						</span> 
					</span>	
				</div>	
				<div id="rowBankruptcy" class="hidden">
					<span style="LEFT: 10px; POSITION: absolute; TOP: 270px; WIDTH: 400px" class="msgLabel">
						Do you have any history of bankruptcy?
						<span id="spnBankruptcyHistoryInd">
							<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
								<input id="optBankruptcyHistoryYes" name="BankruptcyHistoryInd" type="radio" value="1">
								<label for="optBankruptcyHistoryYes" class="msgLabel">Yes</label>
							</span> 
							<span style="LEFT: 350px; POSITION: absolute; TOP: -3px">
								<input id="optBankruptcyHistoryNo" name="BankruptcyHistoryInd" type="radio" value="0">
								<label for="optBankruptcyHistoryNo" class="msgLabel">No</label>
							</span>
						</span>						
						<span style="LEFT: 400px; POSITION: absolute; TOP: -3px">
							<input id="btnBankruptcyHistory" value="Bankruptcy History" type="button" style="WIDTH: 120px" class="msgButton" NAME="btnBankruptcyHistory">
						</span> 						
					</span>	
				</div>	
				<div id="rowCCJ" class="hidden">
					<span style="LEFT: 10px; POSITION: absolute; TOP: 310px; WIDTH: 400px" class="msgLabel">
						Have you ever had any court judgements?
						<span id="spnCCJHistoryInd">
							<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
								<input id="optCCJHistoryYes" name="CCJHistoryInd" type="radio" value="1">
								<label for="optCCJHistoryYes" class="msgLabel">Yes</label>
							</span> 
							<span style="LEFT: 350px; POSITION: absolute; TOP: -3px">
								<input id="optCCJHistoryNo" name="CCJHistoryInd" type="radio" value="0">
								<label for="optCCJHistoryNo" class="msgLabel">No</label>
							</span>
						</span>
						
						<span style="LEFT: 400px; POSITION: absolute; TOP: -3px">
							<input id="btnCCJHistory" value="CCJ History" type="button" style="WIDTH: 120px" class="msgButton" NAME="btnCCJHistory">
						</span> 
					</span>	
				</div>
			</div>
		</form>
		<!-- Main Buttons -->
		<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 280px; WIDTH: 612px">
			<!-- #include FILE="msgButtons.asp" -->
		</div>
		<% /* span to keep tabbing within this screen */ %>
		<span id="spnToFirstField" tabindex="0"></span>
		<!-- #include FILE="fw030.asp" -->
		<!-- #include FILE="attribs/dc085Attribs.asp" -->
		<script language="JScript">
<!--
var m_sMetaAction					= null;
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
var m_sReadOnly						= null;

var ListXML	= null ;
var m_sROExistInd="" ;
var m_sMAExistInd="" ;
var m_sLLExistInd="" ;
var m_sBCExistInd="" ;
var m_sRORecordsExist="";
var m_sMARecordsExist="";
var m_sLLRecordsExist="";
var m_sBCRecordsExist="";
var m_sDMExistInd="" ;
var m_sAHExistInd="" ;
var m_sBHExistInd="" ;
var m_sCCJExistInd="" ;
var m_sDMRecordsExist="";
var m_sAHRecordsExist="";
var m_sBHRecordsExist="";
var m_sCCJRecordsExist="";

var m_bNewFSRecord = false ;

var scScreenFunctions;
var scClientScreenFunctions;

function window.onload()
{
	SetRowVisibility();

	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Financial Summary","DC085",scScreenFunctions);
	
	SetMasks();
	Validation_Init();
	
	RetrieveContextData();
	
	PopulateScreen();
	
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC085");
	if (m_sReadOnly=="1") scScreenFunctions.SetScreenToReadOnly(frmScreen);

	scScreenFunctions.SetFocusToFirstField(frmScreen);

	ClientPopulateScreen();
}

<% /* MF 25/07/2005 Parameterise visibility of Financial rows 
Rows are absolute positioned 40px apart so gaps will appear unless
TOP & HEIGHT style attributes are dynamically modified
*/ %>
function SetRowVisibility()
{	
	var iNextTop=30; 
	var iSpacing=40;
	var bShow=true;
	var aRows=new Array("rowRegOutgoings","rowAnyAccounts",
						"rowLiabilities","rowCreditHistory",
						"rowAppsDeclined","rowArrears",
						"rowBankruptcy","rowCCJ");	
	var aFSParams=new Array("FSRegularOutgoings","FSExistingAccounts",
							"FSLoanLiabilities","FSBankCards", 
							"FSMortgages", "FSArrearsHistory",
							"FSBankruptcy","FSCCJHistory");
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	for(var i=0;i<aRows.length;i++){		
		bShow=XML.GetGlobalParameterBoolean(document,aFSParams[i]);
		if(bShow)
		{
			document.all.item(aRows[i]).className = "";
			document.all.item(aRows[i]).children[0].style.top=iNextTop;
			iNextTop+=iSpacing;
		}
		else
		{
			document.all.item(aRows[i]).className = "hidden";
			
			<% /* MAR172 Hide all the Input fields so that the SetFocus will not error */ %>
			for(var nLoop = 0;nLoop < document.all.item(aRows[i]).getElementsByTagName("INPUT").length;nLoop++)
			{
				document.all.item(aRows[i]).getElementsByTagName("INPUT").item(nLoop).style.visibility = "hidden";
			}
		}
	}
	document.all.item("msgButtons").style.top = iNextTop + 90;
	document.all.item("divBackground").style.height = iNextTop;	
} 

<% /* keep the focus within	this screen when using the tab key */ %>
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<% /* keep the focus within	this screen when using the tab key */ %>
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", "1627");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", "1");
}

function PopulateScreen()
{
	var sCustomerNumber, sCustomerVersionNumber, sCustomerRoleType;
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	var TagRequest = ListXML.CreateRequestTag(window, "SEARCH");
	ListXML.CreateActiveTag("APPLICATION");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	
	ListXML.ActiveTag = TagRequest;
	var TagCustomerList = ListXML.CreateActiveTag("CUSTOMERLIST");
	
	<% /* Loop through all customer context entries */ %>
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		sCustomerNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		<% /* If the customer is an applicant, add him/her to the search */ %>
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			ListXML.CreateActiveTag("CUSTOMER");
			ListXML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			ListXML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			ListXML.ActiveTag = TagCustomerList;
		}
	}
	ListXML.RunASP(document, "GetFinancialSummaryView.asp");
	
	<% /* A record not found error is valid */ %>
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);
		
	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		if(ListXML.SelectTag(null,"FINANCIALSUMMARY") != null)
		{<% /* Find indicators set on FinancialSummary table */ %>
			m_sROExistInd = ListXML.GetTagText("REGULAROUTGOINGSINDICATOR");
			m_sMAExistInd = ListXML.GetTagText("EXISTINGMORTGAGEINDICATOR");
			m_sLLExistInd = ListXML.GetTagText("LOANLIABILITYINDICATOR");
			m_sBCExistInd = ListXML.GetTagText("BANKCARDINDICATOR");
			m_sDMExistInd = (ListXML.GetTagText("DECLINEDMORTGAGEINDICATOR")) ;
			m_sAHExistInd = ListXML.GetTagText("ARREARSHISTORYINDICATOR");
			m_sBHExistInd = ListXML.GetTagText("BANKRUPTCYHISTORYINDICATOR");
			m_sCCJExistInd = ListXML.GetTagText("CCJHISTORYINDICATOR");
			m_bNewFSRecord = false ; 
		}
		else m_bNewFSRecord = true ; <% /* Corresponseing Financial Summary recod exists. So create one while saving */ %>
		
		<% /* Find whether individual records exist */ %>
		m_sRORecordsExist = m_sMARecordsExist = m_sLLRecordsExist = m_sBCRecordsExist = "0";
		
		ListXML.SelectNodes("//FINANCIALSUMMARYVIEW")
		for(var iLoopCounter=0 ; ListXML.SelectTagListItem(iLoopCounter) != false && iLoopCounter < ListXML.ActiveTagList.length; iLoopCounter++)
		{
			if(ListXML.GetTagText("REGULAROUTGOINGSEXISTINGDATA")=="1") m_sRORecordsExist="1";
			if(ListXML.GetTagText("MORTGAGEACCOUNTEXISTINGDATA")=="1") m_sMARecordsExist="1";
			if(ListXML.GetTagText("LOANSLIABILITIESEXISTINGDATA")=="1") m_sLLRecordsExist="1";
			if(ListXML.GetTagText("BANKCREDITCARDEXISTINGDATA")=="1") m_sBCRecordsExist="1";
		}
		
		SetAnswersToQuestions();
		SetQuestionsSettings();
	}
}

function SetAnswersToQuestions()
{
	if(m_sROExistInd=="1" || m_sRORecordsExist=="1")
	{ 
		scScreenFunctions.SetRadioGroupValue(frmScreen, "RegularOutgoingsInd","1");
		m_sROExistInd="1" ; <%/* need to be set to 1 when m_sRORecordsExist=="1" and m_sROExistInd="0" */%>
	}
	else
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "RegularOutgoingsInd","0");
		m_sROExistInd="0" ; <%/* need to be set to 0 when m_sRORecordsExist=="0" and m_sROExistInd="" */%>
	}
	
	if(m_sMAExistInd=="1" || m_sMARecordsExist=="1")
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "AccountsLast3YearsInd","1");
		m_sMAExistInd="1" ;
	}
	else
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "AccountsLast3YearsInd","0");
		m_sMAExistInd="0" ;
	}
	
	if(m_sLLExistInd=="1" || m_sLLRecordsExist=="1")
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "LoansLiabilitiesInd","1");
		m_sLLExistInd="1" ;
	}
	else
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "LoansLiabilitiesInd","0");
		m_sLLExistInd="0" ;
	}
	
	if(m_sBCExistInd=="1" || m_sBCRecordsExist=="1")
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "BankCreditCardsInd","1");
		m_sBCExistInd="1" ;
	}
	else
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "BankCreditCardsInd","0");
		m_sBCExistInd="0" ;
	}
	
	if(m_sDMExistInd=="1" || m_sDMRecordsExist=="1")
	{ 
		scScreenFunctions.SetRadioGroupValue(frmScreen, "MortgagesDeclinedInd","1");
		m_sDMExistInd="1" ; <%/* need to be set to 1 when m_sDMRecordsExist="1" and m_sDMExistInd="0" */%>
	}
	else
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "MortgagesDeclinedInd","0");
		m_sDMExistInd="0" ; <%/* need to be set to 0 when m_sDMRecordsExist="0" and m_sDMExistInd="" */%>
	}
	
	if(m_sAHExistInd=="1" || m_sAHRecordsExist=="1")
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "ArrearsHistoryInd","1");
		m_sAHExistInd ="1";
	}
	else
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "ArrearsHistoryInd","0");
		m_sAHExistInd ="0";
	}
	
	if(m_sBHExistInd=="1" || m_sBHRecordsExist=="1")
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "BankruptcyHistoryInd","1");
		m_sBHExistInd=="1";
	}
	else
	{	
		scScreenFunctions.SetRadioGroupValue(frmScreen, "BankruptcyHistoryInd","0");
		m_sBHExistInd=="0"
	}
	
	
	if(m_sCCJExistInd=="1" || m_sCCJRecordsExist=="1")
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "CCJHistoryInd","1");
		m_sCCJExistInd=="1";
	}
	else
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen, "CCJHistoryInd","0");				
		m_sCCJExistInd=="0";
	}
}

function SetQuestionsSettings()
{
	if(m_sRORecordsExist=="1") scScreenFunctions.SetCollectionState(spnRegularOutgoingsInd, "R");	
	else
	{
		if(m_sROExistInd=="0") frmScreen.btnRegoutgoings.disabled = true ;
	}
	
	if(m_sMARecordsExist=="1")scScreenFunctions.SetCollectionState(spnAccountsLast3YearsInd, "R");
	else
	{
		if(m_sMAExistInd=="0") frmScreen.btnAccountLast3Years.disabled = true ;	
	}
	
	if(m_sLLRecordsExist=="1") scScreenFunctions.SetCollectionState(spnLoansLiabilitiesInd, "R");
	else
	{
		 if( m_sLLExistInd=="0") frmScreen.btnLoansLiabilities.disabled = true ;
	}
	
	if(m_sBCRecordsExist=="1") scScreenFunctions.SetCollectionState(spnBankCreditCardsInd, "R");
	else
	{
		if(m_sBCExistInd=="0") frmScreen.btnBankCreditCards.disabled = true;
	}
	
	if(m_sDMRecordsExist==  	
	"1") 
	scScreenFunctions.SetCollectionState(spnMortgagesDeclinedInd,		
		"R");else{if(m_sDMExistInd=="0"||m_sDMExistInd==0) frmScreen.btnMortgagesDeclined.disabled = true ;
	}

	if(m_sAHRecordsExist==  	
	"1") 
	scScreenFunctions.SetCollectionState(spnArrearsHistoryInd,
		"R");else{if(m_sAHExistInd=="0"||m_sAHExistInd==0) frmScreen.btnArrearsHistory.disabled = true ;
	}
	
	if(m_sBHRecordsExist=="1") scScreenFunctions.SetCollectionState(spnBankruptcyHistoryInd, "R");
	else 
	{	
		<%/*=MC - variant type, sometimes treats as numeric -weird stuff*/%>
		if(m_sBHExistInd=="0"||m_sBHExistInd==0) 
		{
			frmScreen.btnBankruptcyHistory.disabled = true ;
		}
	}
	
	if(m_sCCJRecordsExist=="1") scScreenFunctions.SetCollectionState(spnCCJHistoryInd, "R");
	else
	{
		<%/*=MC - variant type, sometimes treats as numeric -weird stuff*/%>
		if(m_sCCJExistInd=="0"||m_sCCJExistInd==0) 
		{
			frmScreen.btnCCJHistory.disabled = true ;
		}
	}
}

function frmScreen.optRegularOutgoingsYes.onclick()
{
	m_sROExistInd = (frmScreen.optRegularOutgoingsYes.checked) ? "1" : "0";
	if(frmScreen.optRegularOutgoingsYes.checked) 
		frmScreen.btnRegoutgoings.disabled = false ;
}

function frmScreen.optRegularOutgoingsNo.onclick()
{
	m_sROExistInd = (frmScreen.optRegularOutgoingsNo.checked) ? "0" : "1";
	if(frmScreen.optRegularOutgoingsNo.checked)
		frmScreen.btnRegoutgoings.disabled = true ;
}

function frmScreen.optAccountsLast3YearsYes.onclick()
{
	m_sMAExistInd = (frmScreen.optAccountsLast3YearsYes.checked) ? "1" : "0" ;
	if(frmScreen.optAccountsLast3YearsYes.checked)
		frmScreen.btnAccountLast3Years.disabled = false ;
}

function frmScreen.optAccountsLast3YearsNo.onclick()
{
	m_sMAExistInd = (frmScreen.optAccountsLast3YearsNo.checked) ? "0" : "1" ;
	if(frmScreen.optAccountsLast3YearsNo.checked)
		frmScreen.btnAccountLast3Years.disabled = true ;	
}

function frmScreen.optLoansLiabilitiesYes.onclick()
{
	m_sLLExistInd = (frmScreen.optLoansLiabilitiesYes.checked) ? "1" : "0" ;
	if(frmScreen.optLoansLiabilitiesYes.checked)
		 frmScreen.btnLoansLiabilities.disabled = false ;
}

function frmScreen.optLoansLiabilitiesNo.onclick()
{
	m_sLLExistInd = (frmScreen.optLoansLiabilitiesNo.checked) ? "0" : "1" ;
	if(frmScreen.optLoansLiabilitiesNo.checked) 
		frmScreen.btnLoansLiabilities.disabled = true ;
}

function frmScreen.optBankCreditCardsYes.onclick()
{
	m_sBCExistInd = (frmScreen.optBankCreditCardsYes.checked) ? "1" : "0" ; 
	if(frmScreen.optBankCreditCardsYes.checked) 
		frmScreen.btnBankCreditCards.disabled = false ;
}

function frmScreen.optBankCreditCardsNo.onclick()
{
	m_sBCExistInd = (frmScreen.optBankCreditCardsNo.checked) ? "0" : "1" ;
	if(frmScreen.optBankCreditCardsNo.checked)
		frmScreen.btnBankCreditCards.disabled = true ;
}

function frmScreen.optMortgagesDeclinedYes.onclick()
{
	m_sDMExistInd = (frmScreen.optMortgagesDeclinedYes.checked) ? "1" : "0";
	if(frmScreen.optMortgagesDeclinedYes.checked)  
		frmScreen.btnMortgagesDeclined.disabled = false ;
}

function frmScreen.optMortgagesDeclinedNo.onclick()
{
	m_sDMExistInd = (frmScreen.optMortgagesDeclinedNo.checked) ? "0" : "1";
	if(frmScreen.optMortgagesDeclinedNo.checked)
		frmScreen.btnMortgagesDeclined.disabled = true ;
}

function frmScreen.optArrearsHistoryYes.onclick()
{
	m_sAHExistInd = (frmScreen.optArrearsHistoryYes.checked) ? "1" : "0" ;
	if(frmScreen.optArrearsHistoryYes.checked)
		frmScreen.btnArrearsHistory.disabled = false ;
}

function frmScreen.optArrearsHistoryNo.onclick()
{
	m_sAHExistInd = (frmScreen.optArrearsHistoryNo.checked) ? "0" : "1" ;
	if(frmScreen.optArrearsHistoryNo.checked)
		frmScreen.btnArrearsHistory.disabled = true ;	
}

function frmScreen.optBankruptcyHistoryYes.onclick()
{
	m_sBHExistInd = (frmScreen.optBankruptcyHistoryYes.checked) ? "1" : "0" ;
	if(frmScreen.optBankruptcyHistoryYes.checked)
		frmScreen.btnBankruptcyHistory.disabled = false ;	
}

function frmScreen.optBankruptcyHistoryNo.onclick()
{
	m_sBHExistInd = (frmScreen.optBankruptcyHistoryNo.checked) ? "0" : "1" ;
	if(frmScreen.optBankruptcyHistoryNo.checked)
		frmScreen.btnBankruptcyHistory.disabled = true ;
}

function frmScreen.optCCJHistoryYes.onclick()
{
	m_sCCJExistInd = (frmScreen.optCCJHistoryYes.checked) ? "1" : "0" ;
	if(frmScreen.optCCJHistoryYes.checked)
		frmScreen.btnCCJHistory.disabled = false ;	
}

function frmScreen.optCCJHistoryNo.onclick()
{
	m_sCCJExistInd = (frmScreen.optCCJHistoryNo.checked) ? "0" : "1" ;
	if(frmScreen.optCCJHistoryNo.checked)
		frmScreen.btnCCJHistory.disabled = true ;
}

function frmScreen.btnMortgagesDeclined.onclick()
{
	if (SaveFinancialSummary()) frmToDC120.submit();
}

function frmScreen.btnArrearsHistory.onclick()
{
	if (SaveFinancialSummary()) frmToDC130.submit();
}

function frmScreen.btnBankruptcyHistory.onclick()
{
	if (SaveFinancialSummary()) frmToDC140.submit();
}

function frmScreen.btnCCJHistory.onclick()
{
	if (SaveFinancialSummary()) frmToDC150.submit();
}

function frmScreen.btnRegoutgoings.onclick()
{
	if (SaveFinancialSummary()) frmToDC155.submit();
}

function frmScreen.btnAccountLast3Years.onclick()
{
	if (SaveFinancialSummary()) frmToDC070.submit();
}

function frmScreen.btnLoansLiabilities.onclick()
{
	if (SaveFinancialSummary()) frmToDC090.submit();
}

function frmScreen.btnBankCreditCards.onclick()
{
	if (SaveFinancialSummary()) frmToDC080.submit();
}

function btnSubmit.onclick()
{
	if (! SaveFinancialSummary()) return ;
	
	if(scScreenFunctions.CompletenessCheckRouting(window)) 
		frmToGN300.submit();
	else 
	{
		<% // SAB 03/04/2006 EP8 - Change to process flow %>
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if (XML.GetGlobalParameterBoolean(document, "HideDecision"))
			frmToDC200.submit();
		else
			frmToDC370.submit();
	}
}

function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
	{ 
		frmToGN300.submit();
	}
	else
	{
	<% /* changed process flow. Cancel to Other income screen 
		frmToDC060.submit(); */ %>
		<% // PJO 20/12/2005 MAR825 - Show / hide other income on global parameter %>
		var GlobalParamXML = new scXMLFunctions.XMLObject()
		if (GlobalParamXML.GetGlobalParameterBoolean(document,"OtherIncomeSummary"))
		{
			frmToDC195.submit();
		}
		else
		{	
			frmToDC160.submit();
		}
	}
}


function SaveFinancialSummary()
{
	var bReturnVal = false ;
	var TagRequest
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
	if (m_bNewFSRecord) TagRequest = XML.CreateRequestTag(window, "CREATEFINANCIALSUMMARY");
	else TagRequest = XML.CreateRequestTag(window, "UPDATEFINANCIALSUMMARY");
		
	XML.CreateActiveTag("FINANCIALSUMMARY");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	XML.CreateTag("EXISTINGMORTGAGEINDICATOR" , m_sMAExistInd);
	XML.CreateTag("LOANLIABILITYINDICATOR" , m_sLLExistInd);
	XML.CreateTag("BANKCARDINDICATOR" , m_sBCExistInd);
	XML.CreateTag("REGULAROUTGOINGSINDICATOR", m_sROExistInd);
	XML.CreateTag("DECLINEDMORTGAGEINDICATOR", m_sDMExistInd );
	XML.CreateTag("ARREARSHISTORYINDICATOR", m_sAHExistInd );
	XML.CreateTag("BANKRUPTCYHISTORYINDICATOR", m_sBHExistInd );
	XML.CreateTag("CCJHISTORYINDICATOR", m_sCCJExistInd );	
	switch (ScreenRules())
	{
		case 1: 
		case 0: 
				if (m_bNewFSRecord) XML.RunASP(document, "CreateFinancialSummary.asp");
				else XML.RunASP(document, "UpdateFinancialSummary.asp");
				
				break;
		default: 
				XML.SetErrorResponse();
	}	
	
	bReturnVal = XML.IsResponseOK()		
	return bReturnVal ;
}

-->
		</script>
	</body>
</html>
