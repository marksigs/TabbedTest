<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      DC110.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Credit History Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR 		21/05/2004 	BMIDS772 - New
MC		22/06/2004	BMIDS772	DEFECT FIXES FOR BMIDS772 - Disable buttons.
KRW     14/07/2004  BMIDS772   Change to mortgages declined text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</HEAD>

<body>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<script src="validation.js" language="JScript"></script>

<form id="frmToDC160" method="post" action="DC160.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC085" method="post" action="DC085.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC120" method="post" action="DC120.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC130" method="post" action="DC130.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC140" method="post" action="DC140.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC150" method="post" action="DC150.asp" STYLE="DISPLAY: none"></form>


<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" validate  ="onchange" mark>
<div id="divBackground" style="HEIGHT: 190px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
		<strong>Credit History Summary</strong>
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 30px; WIDTH: 440px" class="msgLabel">
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
			<input id="btnMortgagesDeclined" value="Mortgages Declined" type="button" style="WIDTH: 120px" class="msgButton">
		</span>	
		
	</span>	
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 70px; WIDTH: 400px" class="msgLabel">
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
			<input id="btnArrearsHistory" value="Arrears History" type="button" style="WIDTH: 120px" class="msgButton">
		</span> 
	</span>	
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 110px; WIDTH: 400px" class="msgLabel">
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
			<input id="btnBankruptcyHistory" value="Bankruptcy History" type="button" style="WIDTH: 120px" class="msgButton">
		</span> 
		
	</span>	
	
		<span style="LEFT: 10px; POSITION: absolute; TOP: 150px; WIDTH: 400px" class="msgLabel">
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
			<input id="btnCCJHistory" value="CCJ History" type="button" style="WIDTH: 120px" class="msgButton">
		</span> 
	</span>	
</div>
</form><!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 280px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div>


<% /* span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- #include FILE="attribs/dc110Attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction					= null;
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
var m_sReadOnly						= null;

var ListXML	= null ;
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
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Credit History Summary","DC110",scScreenFunctions);
	
	SetMasks();
	Validation_Init();
	
	RetrieveContextData();
	
	PopulateScreen();
	
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC085");
	if (m_sReadOnly=="1") scScreenFunctions.SetScreenToReadOnly(frmScreen);

	scScreenFunctions.SetFocusToFirstField(frmScreen);

	ClientPopulateScreen();
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
	var sCustomerNumber, sCustomerVersionNumber, sCustomerRoleType
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
	ListXML.RunASP(document, "GetCreditHistorySummaryView.asp");
	
	<% /* A record not found error is valid */ %>
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);
		
	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		if(ListXML.SelectTag(null,"FINANCIALSUMMARY") != null)
		{<% /* Find indicators set on FinancialSummary table */ %>
			m_sDMExistInd = (ListXML.GetTagText("DECLINEDMORTGAGEINDICATOR")) 
			m_sAHExistInd = ListXML.GetTagText("ARREARSHISTORYINDICATOR");
			m_sBHExistInd = ListXML.GetTagText("BANKRUPTCYHISTORYINDICATOR")
			m_sCCJExistInd = ListXML.GetTagText("CCJHISTORYINDICATOR")
			m_bNewFSRecord = false ; 
		}
		else m_bNewFSRecord = true ; <% /* Corresponseing Financial Summary recod exists. So create one while saving */ %>
		
		<% /* Find whether individual records exist */ %>
		m_sDMRecordsExist = m_sAHRecordsExist = m_sBHRecordsExist = m_sCCJRecordsExist = "0";
		
		ListXML.SelectNodes("//CREDITHISTORYSUMMARYVIEW")
		for(var iLoopCounter=0 ; ListXML.SelectTagListItem(iLoopCounter) != false && iLoopCounter < ListXML.ActiveTagList.length; iLoopCounter++)
		{
			if(ListXML.GetTagText("DECLINEDMORTGAGEEXISTINGDATA")=="1") m_sDMRecordsExist="1";
			if(ListXML.GetTagText("ARREARSHISTORYEXISTINGDATA")=="1") m_sAHRecordsExist="1";
			if(ListXML.GetTagText("BANKRUPTCYHISTORYEXISTINGDATA")=="1") m_sBHRecordsExist="1";
			if(ListXML.GetTagText("CCJHISTORYEXISTINGDATA")=="1") m_sCCJRecordsExist="1";
		}
		
		SetAnswersToQuestions();
		SetQuestionsSettings();
	}
}

function SetAnswersToQuestions()
{
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

function frmScreen.optMortgagesDeclinedYes.onclick()
{
	m_sDMExistInd = (frmScreen.optMortgagesDeclinedYes.checked) ? "1" : "0"
	if(frmScreen.optMortgagesDeclinedYes.checked)  
		frmScreen.btnMortgagesDeclined.disabled = false ;
}

function frmScreen.optMortgagesDeclinedNo.onclick()
{
	m_sDMExistInd = (frmScreen.optMortgagesDeclinedNo.checked) ? "0" : "1"
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
	if (SaveFinancialSummary()) frmToDC120.submit()
}

function frmScreen.btnArrearsHistory.onclick()
{
	if (SaveFinancialSummary()) frmToDC130.submit()
}

function frmScreen.btnBankruptcyHistory.onclick()
{
	if (SaveFinancialSummary()) frmToDC140.submit()
}

function frmScreen.btnCCJHistory.onclick()
{
	if (SaveFinancialSummary()) frmToDC150.submit()
}

function btnSubmit.onclick()
{
	if (! SaveFinancialSummary()) return ;

	if(scScreenFunctions.CompletenessCheckRouting(window)) frmToGN300.submit();
	else frmToDC160.submit();
}

function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window)) frmToGN300.submit();
	else frmToDC085.submit();
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
	
	XML.CreateTag("DECLINEDMORTGAGEINDICATOR", m_sDMExistInd );
	XML.CreateTag("ARREARSHISTORYINDICATOR", m_sAHExistInd );
	XML.CreateTag("BANKRUPTCYHISTORYINDICATOR", m_sBHExistInd );
	XML.CreateTag("CCJHISTORYINDICATOR", m_sCCJExistInd );
	
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
				if (m_bNewFSRecord) XML.RunASP(document, "CreateFinancialSummary.asp");
				else XML.RunASP(document, "UpdateFinancialSummary.asp");
				
				break;
		default: // Error
				XML.SetErrorResponse();
	}	
	
	bReturnVal = XML.IsResponseOK()		
	return bReturnVal ;
}

-->
</script>
</body>
</html>