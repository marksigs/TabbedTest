<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      cm065P.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Stored Quotes Summary (popup)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
CL		27/03/01	Creation of screen
LD		23/05/02	SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
GD		15/08/02	BMIDS00222 - Removed 'Life?' column
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		09/10/2002	BMIDS00556	Removed Validation_Init()
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:
JD		10/03/2006	MAR1061 added amt req, prop price, ltv
HMA     28/03/2006  MAR1500 Check rates on entry to screen.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Stored Quotes Summary <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets - remove any which are not required */ %>
<span style="TOP: 215px; LEFT: 310px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scTable" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToCM010" method="post" action="CM010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM060" method="post" action="CM060.asp" STYLE="DISPLAY: none"></form>


<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 225px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span id="spnStoredQuotes" style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblStoredQuotes" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable" style="PADDING-LEFT: 5px;PADDING-RIGHT: 5px">
			<COLGROUP>
				<COL width="2%"/>
				<COL width="10%"/>
				<COL width="5%"/>
				<COL width="8%"/>
				<COL width="8%"/>
				<COL width="3%"/>
				<COL width="9%"/>
				<COL width="5%"/>
				<COL width="5%"/>
				<COL width="5%"/>
				<COL width="4%"/>
				<COL width="5%"/>
				<COL width="5%"/>
				<COL width="5%"/>
				<COL width="9%"/>
			</COLGROUP>
			<tr id="rowTitles">	
				<td class="TableHead">No.</td>
				<td class="TableHead">Product Details</td>
				<td class="TableHead">Repay Type</td>
				<td class="TableHead">Amt Req</td>	
				<td class="TableHead">Prop price</td>
				<td class="TableHead">LTV</td>
				<td class="TableHead">Mtg Cost</td>		
				<td class="TableHead">B&amp;C Cost</td>			
				<td class="TableHead">PP Cost</td>		
				<td class="TableHead">Total Cost</td>		
				<td class="TableHead">APR</td>
				<td class="TableHead">Act?</td>
				<td class="TableHead">Rec?</td>
				<td class="TableHead">Acc?</td>
				<td class="TableHead">Cost at Final Rate</td>
			</tr>
			<!-- row data -->
			<tr id="row01">
				<td class="TableTopLeft">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopCenter">&nbsp;</td>
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr id="row03">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr id="row04">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr id="row05">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<tr id="row06">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row07">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row08">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row09">
				<td class="TableLeft">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row10">
				<td class="TableBottomLeft">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomCenter">&nbsp;</td>
				<td class="TableBottomRight">&nbsp;</td>
			</tr>
		</table>
	</span>
</div>
</form>


<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 250px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cm065PAttribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var ListXML = null;
var m_sApplicationMode = null;
var m_sReadOnly = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sActiveQuotationNumber = null;
var m_iActiveLocation = null;
var scScreenFunctions;
var m_blnReadOnly = false;
var m_idMetaAction = null;
var m_sRequestArray = null;
var m_BaseNonPopupWindow = null;
var m_XMLRepay = null;	

var COLUMN_ACTIVE_QUOTE = 11;
var COLUMN_REC_QUOTE = 12;
var COLUMN_ACC_QUOTE = 13;

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
%>	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	RetrieveData();
	
	var sGroups = new Array("RepaymentType");
	var XMLTemp = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLTemp.GetComboLists(document,sGroups);
	m_XMLRepay = XMLTemp.GetComboListXML("RepaymentType")
	
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveData()
{
	//var ArraySelectedRows = new Array();
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters	= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sRequestArray = sParameters[0];
	m_sApplicationMode = sParameters[1]; //m_sApplicationMode
	m_sReadOnly = sParameters[2]; //m_sReadOnly
	m_sApplicationNumber = sParameters[3]; //m_sApplicationNumber
	m_sApplicationFactFindNumber = sParameters[4]; //m_sApplicationFactFindNumber
	m_sActiveQuotationNumber = sParameters[5]; //m_sActiveQuotationNumber
	m_sMetaAction = sParameters[6]; //m_sMetaAction 
	m_sidXML =  sParameters[7]; //m_sidXML
}

function PopulateScreen()
{
	ListXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTagFromArray(m_sRequestArray,"SEARCH");
	ListXML.CreateActiveTag("QUOTATION");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);//zzzz????
	
	if(m_sApplicationMode == "Quick Quote") 
	{
		ListXML.RunASP(document,"QQFindStoredQuoteDetails.asp");
	}
	else 
	{
		ListXML.RunASP(document,"AQFindStoredQuoteDetails.asp");
	}
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);
	if(sResponseArray[0])
	{
		PopulateTable(0);
		ListXML.ActiveTag = null;
		ListXML.CreateTagList("STOREDQUOTATION");
		scTable.initialiseTable(tblStoredQuotes,0,"",PopulateTable,10,ListXML.ActiveTagList.length);
		if(m_sMetaAction == "fromCM060")
		{
			var sQuotationNumber = m_sidXML;
			var nTotal = 0;
			var bFound = false;
			while(nTotal <= ListXML.ActiveTagList.length && !bFound)
			{
				for(var nLoop = 1;nLoop <= 10;nLoop++)
				{
					if(tblStoredQuotes.rows(nLoop).cells(0).innerText == sQuotationNumber)
					{
						bFound = true;
						scTable.setRowSelected(nLoop);
					}
					nTotal++;
				}
				
				if(!bFound && nTotal < ListXML.ActiveTagList.length) scTable.pageDown();
			}
			spnStoredQuotes.onclick();
		}
		else
		{
			<% /* MAR1500 Check if the Active Quote has consistent rates */ %>
			if (CheckForChangedRates(m_sApplicationNumber, m_sApplicationFactFindNumber) == true)
			{
				alert("The Rates have changed since the accepted/active quote was created and the details held may be inconsisent. Please create a quote ");
			}
		}
	}
}

function PopulateTable(nStart)
{
	ListXML.ActiveTag = null;
	var xmlAPPLICATIONFACTFIND = ListXML.SelectTag(null,"APPLICATIONFACTFIND");

	if(xmlAPPLICATIONFACTFIND != null)
	{	
		sActiveQuoteNumber =  ListXML.GetTagText("ACTIVEQUOTENUMBER");
		
		if(sActiveQuoteNumber != m_sActiveQuotationNumber)
		{
			sActiveQuoteNumber = m_sActiveQuotationNumber;
		}
		
		sRecommendedQuoteNumber =  ListXML.GetTagText("RECOMMENDEDQUOTENUMBER");
		sAcceptedQuoteNumber =  ListXML.GetTagText("ACCEPTEDQUOTENUMBER");
		m_sRegulationIndicator = ListXML.GetTagText("REGULATIONINDICATOR");
	}
	ListXML.ActiveTag = null;
	var TagListSTOREDQUOTATION = ListXML.CreateTagList("STOREDQUOTATION");
	//var sQuotationNumber;		MAR7
	var sActiveQuoteNumber;
	var sRecommendedQuoteNumber;
	var sAcceptedQuoteNumber;
	var sMortgageSubQuoteNumber;	<% /* MAR7 GHun */ %>
		
	//Added for testing only !!
	if (m_sActiveQuotationNumber != sActiveQuoteNumber)
	alert ("The Active Quotation Number in Context is not the same as the Active Quotation Number returnedf rom the search")
	
	<% /* Populate the table with a set of records, starting with 
		record nStart */ %>
	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) && nLoop < 10; nLoop++)
	{			
		<% /* MAR7 GHun */ %>
		sMortgageSubQuoteNumber = ListXML.GetTagText("MORTGAGESUBQUOTENUMBER");
		tblStoredQuotes.rows(nLoop+1).setAttribute("MortgageSubQuoteNumber", sMortgageSubQuoteNumber);
		<% /* MAR7 End */ %>
		
		sQuotationNumber = ListXML.GetTagText("QUOTATIONNUMBER");					
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(0),sQuotationNumber );
		//BMIDS00209 - removed column (moved cells to the right across one).  DPF 29/07/2002 
		//scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(3), ListXML.GetTagText("LIFECOVER"));
		
		<% /* MAR1061 amt req, Prop Price, LTV*/ %>
		//Amt Req
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(3),ListXML.GetTagText("AMOUNTREQUESTED") );
		
		//Prop Price
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(4),ListXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE") );
		
		//ltv
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(5),m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("LTV"), 2) );
		
		//Mtg Cost
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(6),m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("TOTALNETMONTHLYCOST"), 2) );
		
		//B&C Cost
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(7),m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("TOTALBCMONTHLYCOST"), 2) );

		//PP Cost
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(8),m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("TOTALPPMONTHLYCOST"), 2) );
				
		//TOTAL COST				
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(9),m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("TOTALQUOTATIONCOST"), 2) );

		//APR
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(10),m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundValue(ListXML.GetTagText("APR"), 2) );

		var sAQN = "";		
		if(sQuotationNumber == sActiveQuoteNumber) 
		{
			sAQN = "Yes";
			m_iActiveLocation = nLoop+1;
		}
		
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(COLUMN_ACTIVE_QUOTE),sAQN);
			
		var sRQN = "";
		if(sQuotationNumber == sRecommendedQuoteNumber) sRQN = "Yes";

		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(COLUMN_REC_QUOTE),sRQN);
		
		var sAccQN = "";
		if(sQuotationNumber == sAcceptedQuoteNumber)
		{
			sAccQN = "Yes";
			m_iAcceptedLocation = nLoop + 1; 
		}	
		
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(COLUMN_ACC_QUOTE),sAccQN);


		var sRepayTypeText = "Not applicable";
		var sRepayTypeValidation = "";
		
		var sProductDetail = "";
		var nComponents = "0";
		var dblFinalRMC;
		var nFinalRateCost = 0;  <% /* BMIDS00224 - DPF 30/07/02 - added FINALRATEMONTHLYCOST */ %>
		var sPPBreakdown = "";
		var dblInterestOnly;
		var nTotalInterestOnly = 0;
		var dblCapitalInterest;
		var nTotalCapitalInterest = 0;
		
		ListXML.CreateTagList("LOANCOMPONENT");
		for(var nComponentLoop = 0;ListXML.SelectTagListItem(nComponentLoop);nComponentLoop++)
		{
			nComponents++;
			var sThisRepayTypeText = ListXML.GetTagAttribute("REPAYMENTMETHOD","TEXT");
			var sThisRepayTypeValidation = ListXML.GetTagText("REPAYMENTMETHOD");
			
			var sSearch = "LISTNAME/LISTENTRY[VALUEID='" + sThisRepayTypeValidation + "']/VALIDATIONTYPELIST/VALIDATIONTYPE";
			sThisRepayTypeValidation = m_XMLRepay.selectSingleNode(sSearch).text;
		
			var sThisProductDetail = ListXML.GetTagText("PRODUCTNAME");

			dblFinalRMC = parseFloat(ListXML.GetTagText("FINALRATEMONTHLYCOST"));
			if ( ! isNaN(dblFinalRMC) )
			{
				nFinalRateCost = nFinalRateCost + dblFinalRMC;
			}
			
			// Repayment Type can be Interest Only(I), Capital & Interest (C) or Part and Part(P)
			// If Repayment Type Validation is different for different Loan Components, set it to "Multiple" overall (MARS1061 - was set to P&P previously)
			
			// If the Validation Type is the same for different Loan Components but the text is different,
			// set the text to be a general value.
			
			if(sThisRepayTypeValidation != "" && sRepayTypeText != "Part and Part") // Already decided this is P&P
			{
				if(sRepayTypeValidation != "")
				{
					if (sThisRepayTypeValidation != sRepayTypeValidation) // Validation Types differ - set to P&P
					{
						sRepayTypeText = "Multiple"; //MAR1061
					}
					else if (sThisRepayTypeValidation == sRepayTypeValidation) // Validation Types the same - check text
					{
						if (sThisRepayTypeText != sRepayTypeText)
						{	
							if (sThisRepayTypeValidation == "I") sRepayTypeText = "I/O"; //MAR1061
							else if (sThisRepayTypeValidation == "C") sRepayTypeText = "C&I"; //MAR1061
							else sRepayTypeText = "P&P"; //MAR1061
						}
					}	
				}
				else
				{
					sRepayTypeValidation = sThisRepayTypeValidation;
					sRepayTypeText = sThisRepayTypeText;
				}									
			}	
			
			//Keep a running total of Interest Only and Capital & Interest amounts.
			dblLoanAmount = parseFloat(ListXML.GetTagText("LOANAMOUNT"));
			dblInterestOnly = parseFloat(ListXML.GetTagText("INTERESTONLYELEMENT"));
			dblCapitalInterest = parseFloat(ListXML.GetTagText("CAPITALANDINTERESTELEMENT"));
			
			if (sThisRepayTypeValidation == "I")
			{	
				// Loan Amount holds the Interest Only element
				if (!isNaN(dblLoanAmount)) nTotalInterestOnly = nTotalInterestOnly + dblLoanAmount;
			}
			else if (sThisRepayTypeValidation == "C")
			{
				// Loan Amount holds the Capital & Interest element
				if (!isNaN(dblLoanAmount)) nTotalCapitalInterest = nTotalCapitalInterest + dblLoanAmount; 	
			}
			else
			{
				// Part & Part : Interest Only and Capital&Interest amounts are held separately
				if (!isNaN(dblInterestOnly)) nTotalInterestOnly = nTotalInterestOnly + dblInterestOnly;
				if (!isNaN(dblCapitalInterest)) nTotalCapitalInterest = nTotalCapitalInterest + dblCapitalInterest;
			}
						
			// If Product Name is different for different Loan Components, set it to "Not Applicable"
			if (sProductDetail == "" && sThisProductDetail != "")
			{
				sProductDetail = sThisProductDetail;
			}
			else
			{
				if (sProductDetail != "Not Applicable" && sThisProductDetail != sProductDetail) 
				{
					sProductDetail = "Multiple Component";
				}
			}		
		}

		ListXML.ActiveTagList = TagListSTOREDQUOTATION;		
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(1),sProductDetail);
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(2),GetShortVersion(sRepayTypeText));
		scScreenFunctions.SizeTextToField(tblStoredQuotes.rows(nLoop+1).cells(14),m_BaseNonPopupWindow.top.frames[1].document.all.scMathFunctions.RoundValue(nFinalRateCost, 2));
		
		//Tool Tip 
		if (sRepayTypeText == "P&P") //MAR1061
		{
			sPPBreakdown = "I/O £" + nTotalInterestOnly + " - C/I £" + nTotalCapitalInterest;
			tblStoredQuotes.rows(nLoop+1).cells(2).title = sPPBreakdown ;
		}	
	}
	//if(nLoop < (m_iTableLength))
	//{
//		ResetBlankRows(nLoop);
	//}	
}
<%/* PSC 29/06/2005 MAR5 - End */%>

function GetShortVersion(sRepayTypeText)
{
	var sShortVersion = "";
	switch (sRepayTypeText)
	{
		case "Interest Only": sShortVersion = "I/O"; break;
		case "Capital and Interest": sShortVersion = "C&I"; break;
		case "Part and Part": sShortVersion = "P&P"; break;
		default : sShortVersion = sRepayTypeText;
	}
	return sShortVersion;
}
function ResetBlankRows(iCount)
{
	for(iIndex = (iCount + 1); iIndex <= m_iTableLength; iIndex++)
	{
		if (iIndex == 1)
		{
			tblStoredQuotes.rows(iIndex).cells(12).className = "TableTopCenter";
			tblStoredQuotes.rows(iIndex).cells(13).className = "TableTopCenter";
			tblStoredQuotes.rows(iIndex).cells(14).className = "TableTopRight";
		} else
		{
			if (iIndex == m_iTableLength)
			{
				tblStoredQuotes.rows(iIndex).cells(12).className = "TableBottomCenter";
				tblStoredQuotes.rows(iIndex).cells(13).className = "TableBottomCenter";
				tblStoredQuotes.rows(iIndex).cells(14).className = "TableBottomRight";
			} else
			{
				tblStoredQuotes.rows(iIndex).cells(12).className = "TableCenter";
				tblStoredQuotes.rows(iIndex).cells(13).className = "TableCenter";
				tblStoredQuotes.rows(iIndex).cells(14).className = "TableRight";
			}
		}
	}
}

function btnSubmit.onclick()
{
	window.close();	
}

<% /* MAR1500 Check if rates have changed for the active quote */ %>
function CheckForChangedRates(sApplicationNumber, sApplicationFactFindNumber)
{	
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTagFromArray(m_sRequestArray,"");
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	
	XML.RunASP(document, "HaveRatesChanged.asp");
	XML.IsResponseOK()
	
	if (XML.IsResponseOK())
	{
		if (XML.GetTagAttribute("QUOTATION", "RATESINCONSISTENT") == "1")
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
}
-->
</script>
</body>
</html>




