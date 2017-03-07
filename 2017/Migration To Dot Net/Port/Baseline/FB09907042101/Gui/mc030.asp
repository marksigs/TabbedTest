<%@ LANGUAGE="JSCRIPT" %>

<% /*
Workfile:      mc030.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Mortgage Calculator Illustration Details screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		24/11/99	AQR MC1: Extend width of Product Name field.
RF		02/02/00	Reworked for IE5 and performance.
AY		14/02/00	Change to msgButtons button types
JLD		29/02/00	SYS0287 increased size of CM130
JLD		01/03/00	SYS0379 various minor tweaks
JLD		07/03/00	SYS0402 if rate period is -1, set it to "Remaining Period"
JLD		14/03/00	SYS0407 Add an incentives button.
AY		17/03/00	MetaAction changed to ApplicationMode
AY		03/04/00	New top menu/scScreenFunctions change
JLD		11/04/00	AQR SYS0534 - the interest rate should always be rounded.
AY		13/04/00	SYS0328 - Dynamic currency display
JLD		17/04/00	Change to interface with CM150
DM		27/06/00	Removed the MIRAS Chk box and txt box.
CL		05/03/01	SYS1920 Read only functionality added
APS		13/03/01	SYS1920 Screen Title changes
MC		05/10/01	SYS2776 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		06/06/02	SYS4800 - Error following SYS4727
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>
<script src="validation.js" language="JScript"></script>

<OBJECT data=scTable.htm height=1 id=scTable 
style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet 
VIEWASTEXT></OBJECT>

<% // Specify Forms Here %>
<form id="frmStoredIllustrations" method="post" action="mc020.asp" 	STYLE="DISPLAY: none"></form>
		
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% // Specify Screen Layout Here %>
<form id="frmScreen" year4 validate="onchange" mark>
	<div id="divBackground" 
		style="HEIGHT: 440px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px">
												
		<div style="HEIGHT: 100px; LEFT: 0px; POSITION: absolute; TOP: 10px; WIDTH: 604px" 
			class="msgGroup">

			<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
				<strong>Loan Details...</strong>
			</span>
			<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
				Purchase Price/<br>Estimated Value
				<span style="LEFT: 100px; POSITION: absolute; TOP: 5px">
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtPurchasePrice" name="PurchasePrice" maxlength="7" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 65px" class="msgTxt"> 
				</span> 
			</span>
			<span style="LEFT: 200px; POSITION: absolute; TOP: 30px" class="msgLabel">
				Amount <br>Requested
				<span style="LEFT: 100px; POSITION: absolute; TOP: 5px">
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtAmountRequested" name="AmountRequested" maxlength="7" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 65px" class="msgTxt"> 
				</span> 
			</span>
			<span style="LEFT: 420px; POSITION: absolute; TOP: 35px" class="msgLabel">
				LTV
				<span style="LEFT: 30px; POSITION: absolute; TOP: 0px">
					<input id="txtLTV" name="LTV" maxlength="10" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 60px" class="msgTxt">
					<span style="LEFT: 70px; POSITION: absolute; TOP: 0px" class="msgLabel">
						%
					</span>
				</span> 
			</span>
			<span style="LEFT: 4px; POSITION: absolute; TOP: 65px" class="msgLabel">
				Interest Only<br>Amount
				<span style="LEFT: 100px; POSITION: absolute; TOP: 5px">
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtInterestOnlyAmount" name="InterestOnlyAmount" maxlength="7" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 65px" class="msgTxt"> 
				</span> 
			</span>
			<span style="LEFT: 200px; POSITION: absolute; TOP: 65px" class="msgLabel">
				<LABEL id=idCIAmount></LABEL> 
				<span style="LEFT: 100px; POSITION: absolute; TOP: 5px">					
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtCapitalAndInterest" name="CapitalAndInterest" maxlength="7" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 65px" class="msgTxt"> 
				</span> 
			</span>
			<span style="LEFT: 380px; POSITION: absolute; TOP: 70px" class="msgLabel">
				<LABEL id="idTerm"></LABEL> 
				<span style="LEFT: 70px; POSITION: absolute; TOP: 0px">						
					<input id="txtTermYears" name="TermYears" maxlength="2" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 28px" class="msgTxt"> 
					<span style="LEFT: 34px; POSITION: absolute; TOP: 0px" class="msgLabel">
						Years 
					</span> 
				</span>				
				<span style="LEFT: 145px; POSITION: absolute; TOP: 0px">						
					<input id="txtTermMonths" name="TermMonths" maxlength="2" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 28px" class="msgTxt"> 
					<span style="LEFT: 34px; POSITION: absolute; TOP: 0px" class="msgLabel">
						Months
					</span> 
				</span>					
			</span>

		</div>
			
		<div style="HEIGHT: 200px; LEFT: 0px; POSITION: absolute; TOP: 115px; WIDTH: 604px" class="msgGroup">

			<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
				<strong>Product Details...</strong> 
			</span>									
			<span style="LEFT: 4px; POSITION: absolute; TOP: 35px" class="msgLabel">
				Product Code
				<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
					<input id="txtProductCode" name="ProductCode" maxlength="10" 
						style="POSITION: absolute; WIDTH: 60px" class="msgTxt"> 
				</span> 
			</span>
			<span style="LEFT: 200px; POSITION: absolute; TOP: 30px" class="msgLabel">
				Product <br>Name
				<span style="LEFT: 100px; POSITION: absolute; TOP: 3px">
					<input id="txtProductName" name="ProductName" maxlength="80" 
						style="POSITION: absolute; WIDTH: 250px" class="msgTxt"> 
				</span> 
			</span>
			<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
				Discount
				<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
					<span style="LEFT: 65px; POSITION: absolute; TOP: 0px" class="msgLabel">
						%
					</span>
					<input id="txtDiscount" name="Discount" maxlength="10" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 60px" class="msgTxt"> 
				</span> 
			</span>
			<span id="spnLender" style="LEFT: 200px; POSITION: absolute; TOP: 65px" class="msgLabel">
				Lender <br>Name
				<span style="LEFT: 100px; POSITION: absolute; TOP: 3px">
					<input id="txtLenderName" name="LenderName" maxlength="10" 
						style="POSITION: absolute; WIDTH: 250px" class="msgTxt"> 
				</span> 
			</span>
			<span id="spnInterestRateTable" 
					style="LEFT: 4px; POSITION: absolute; TOP: 100px">
				<table id="tblInterestRateTable" width="594" border="0" 
					cellspacing="0" cellpadding="0" class="msgTable">
						
					<tr id="rowTitles">	
						<td width="40%" class="TableHead">Type</td>
						<td width="20%" class="TableHead">Rate %</td>
						<td width="20%" class="TableHead">Period</td>
						<td class="TableHead">End Date</td></tr>
					<tr id="row01">		
						<td width="40%" class="TableTopLeft">&nbsp;</td>	
						<td width="20%" class="TableTopCenter">&nbsp;</td>
						<td width="20%" class="TableTopCenter">&nbsp;</td>
						<td class="TableTopRight">&nbsp;</td></tr>
					<tr id="row02">
						<td width="40%" class="TableLeft">&nbsp;</td>	
						<td width="20%" class="TableCenter">&nbsp;</td>
						<td width="20%" class="TableCenter">&nbsp;</td>
						<td class="TableRight">&nbsp;</td></tr>
					<tr id="row03">		
						<td width="40%" class="TableLeft">&nbsp;</td>
						<td width="20%" class="TableCenter">&nbsp;</td>
						<td width="20%" class="TableCenter">&nbsp;</td>
						<td class="TableRight">&nbsp;</td></tr>
					<tr id="row04">		
						<td width="40%" class="TableLeft">&nbsp;</td>
						<td width="20%" class="TableCenter">&nbsp;</td>
						<td width="20%" class="TableCenter">&nbsp;</td>
						<td class="TableRight">&nbsp;</td></tr>
					<tr id="row05">		
						<td width="40%" class="TableBottomLeft">&nbsp;</td>
						<td width="20%" class="TableBottomCenter">&nbsp;</td>
						<td width="20%" class="TableBottomCenter">&nbsp;</td>	
						<td class="TableBottomRight">&nbsp;</td></tr>
				</table>
			</span>
		</div>
						
		<div style="HEIGHT: 100px; LEFT: 0px; POSITION: absolute; TOP: 320px; WIDTH: 604px" class="msgGroup">
				
			<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
				<strong>Cost Details...</strong> 
			</span>		
			<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
				Interest Only <br>Monthly Cost
				<span style="LEFT: 100px; POSITION: absolute; TOP: 5px">
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtMonthlyCostInterestOnly" name="MonthlyCostInterestOnly" maxlength="7" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 65px" class="msgTxt"> 
				</span> 
			</span>
			<span style="LEFT: 200px; POSITION: absolute; TOP: 30px" class="msgLabel">
				<LABEL id="idCIMonthlyCost"></LABEL> 
				<span style="LEFT: 100px; POSITION: absolute; TOP: 5px">
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtMonthlyCostCapitalAndInterest" 
						name="MonthlyCostCapitalAndInterest" maxlength="7" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 65px" class="msgTxt"> 
				</span> 
			</span>
			<span id="spnPartAndPartCost" style="LEFT: 400px; POSITION: absolute; TOP: 30px" class="msgLabel">
				Part &amp; Part <br>Monthly Cost 
				<span style="LEFT: 100px; POSITION: absolute; TOP: 5px">
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtMonthlyCostPartAndPart" 
						name="MonthlyCostPartAndPart" maxlength="7" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 65px" class="msgTxt"> 
				</span> 
			</span> 
			<span style="LEFT: 4px; POSITION: absolute; TOP: 70px">
					<input id="btnOneOffCosts" value="One-Off Costs" type="button" 
						style="WIDTH: 100px" class="msgButton"> 					
			</span>
			<span style="LEFT: 110px; POSITION: absolute; TOP: 70px">
					<input id="btnIncentives" value="Incentives" type="button" 
						style="WIDTH: 100px" class="msgButton"> 					
			</span>			
		</div>	
	</div> 
</form>

<% // Main Buttons %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 485px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!--  #include FILE="Customise/MC030Customise.asp" -->

<% // Specify Code Here %>
<script language="JScript">
<!--	// JScript Code
	
var m_sApplicationMode			= "";
var m_IllustrationListXML	= "";
var m_iDisplayNumber		= 0;
var m_sMultipleLender		= "";
var m_sOneOffCosts			= "";
var m_sTotalIncentives		= "";
var m_sUnitId				= "";
var m_sUserId				= "";
var m_sUserType				= "";
var m_sStartDate			= "";
var m_sCurrency = "";
var scScreenFunctions;

function window.onload()
{
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Mortgage Calculator","MC030",scScreenFunctions);

	scScreenFunctions.SetCurrency(window,frmScreen);
	RetrieveContextData();
	PopulateScreen();
	// MC SYS2564/SYS2775 for client customisation
	Customise();

	Validation_Init();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
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
	m_sApplicationMode		= scScreenFunctions.GetContextParameter(window,
									"idApplicationMode",null);
	m_sIllustrationList	= scScreenFunctions.GetContextParameter(window,
									"idXML",null);
	m_iDisplayNumber	= scScreenFunctions.GetContextParameter(window,
									"idMCDisplayNumber","0");
	m_sMultipleLender	= scScreenFunctions.GetContextParameter(window,
									"idMultipleLender","0");
	m_sUnitId			= scScreenFunctions.GetContextParameter(window,
									"idUnitId",null);
	m_sUserType			= scScreenFunctions.GetContextParameter(window,
									"idUserType",null);
	m_sUserId			= scScreenFunctions.GetContextParameter(window,
									"idUserId",null);
	m_sCurrency = scScreenFunctions.GetContextParameter(window,"idCurrency",null);
}

function PopulateScreen()
{
	<% /*
		Displays the illustration data from the XML to the screen controls
	*/ %>

	var blnValidIllustration = false;
			
	m_IllustrationListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();						
	m_IllustrationListXML.LoadXML(m_sIllustrationList);

	if (m_IllustrationListXML.CreateTagList("ILLUSTRATION") != null)
	{
		if (m_IllustrationListXML.SelectTagListItem(m_iDisplayNumber-1) == true)
		{
			var tagILLUSTRATION = m_IllustrationListXML.ActiveTag;
					
			frmScreen.txtPurchasePrice.value	= m_IllustrationListXML.GetTagText("PURCHASEPRICE");
			frmScreen.txtAmountRequested.value 	= m_IllustrationListXML.GetTagText("AMOUNTREQUESTED");
			frmScreen.txtLTV.value 				= m_IllustrationListXML.GetTagText("LTV");
			frmScreen.txtTermYears.value 		= m_IllustrationListXML.GetTagText("TERMINYEARS");
			frmScreen.txtTermMonths.value 		= m_IllustrationListXML.GetTagText("TERMINMONTHS");				
			frmScreen.txtInterestOnlyAmount.value	= m_IllustrationListXML.GetTagText("INTERESTONLYELEMENT");
			frmScreen.txtCapitalAndInterest.value	= m_IllustrationListXML.GetTagText("CAPITALANDINTERESTELEMENT");
			m_sStartDate = m_IllustrationListXML.GetTagText("STARTDATE");
																								
			frmScreen.txtProductCode.value		= m_IllustrationListXML.GetTagText("MORTGAGEPRODUCTCODE");
			frmScreen.txtProductName.value 		= m_IllustrationListXML.GetTagText("PRODUCTNAME");
					
			<% /* total incentives for one-off costs screen */%>
			m_sTotalIncentives						= m_IllustrationListXML.GetTagText("TOTALINCENTIVES");
					
			if (m_sMultipleLender == "0")
				scScreenFunctions.HideCollection(spnLender);
			else
				frmScreen.txtLenderName.value = m_IllustrationListXML.GetTagText("LENDERNAME");	
															
			frmScreen.txtMonthlyCostInterestOnly.value	= m_IllustrationListXML.GetTagText("MONTHLYINTERESTONLYCOST");
			frmScreen.txtMonthlyCostCapitalAndInterest.value	= m_IllustrationListXML.GetTagText("MONTHLYCAPITALANDINTERESTCOST");
			frmScreen.txtMonthlyCostPartAndPart.value	= m_IllustrationListXML.GetTagText("MONTHLYPARTANDPARTCOST");
					
			<% /* one off costs screen */%>
			if (m_IllustrationListXML.SelectTag(tagILLUSTRATION, "ONEOFFCOSTLIST") != null)
				m_sOneOffCosts = m_IllustrationListXML.ActiveTag.xml;
					
					
			ReadInterestRateBands(tagILLUSTRATION);
				
			blnValidIllustration = true;
		}												
	}
			
	scScreenFunctions.SetScreenToReadOnly(frmScreen);
			
	if (blnValidIllustration == false)
		alert("Illustration not found");
}
		
		
function ReadInterestRateBands(tagILLUSTRATION)
{
	<% /* 
		Description:	Reads the interest rate bands associated with the
						product which in turn is associated with the 
						illustration required to be displayed
		Args Passed:	tagILLUSTRATION	- XML ActiveTag for the current illustration
	*/ %>

	m_IllustrationListXML.ActiveTag = tagILLUSTRATION;
	if (m_IllustrationListXML.CreateTagList("INTERESTRATETYPE") != null)
	{			
		var tagINTERESTRATETYPELIST = m_IllustrationListXML.ActiveTagList;
				
		<% /* Up to a maximum of 5 interest rate bands */ %>
		for (var iLoop=0; iLoop < m_IllustrationListXML.ActiveTagList.length; iLoop++)
		{
			m_IllustrationListXML.ActiveTagList = tagINTERESTRATETYPELIST;
					
			if (m_IllustrationListXML.SelectTagListItem(iLoop) == true)
			{											
				DisplayInterestRate(iLoop, tagILLUSTRATION);
			}
		}
	}
}
		
function DisplayInterestRate(iLoop, tagILLUSTRATION)
{
	<% /*
		Args Passed:	iLoop			- Pointer to the interest rate band
						tagILLUSTRATION	- XML ActiveTag for the current illustration
	*/ %>
	
	var sSequenceNumber			= null;
	var sInterestRateType		= null;
	var sInterestRatePeriod		= null;
	var sInterestRateEndDate	= null;
	var sInterestRateTypeText	= "Fixed";
	var dInterestRate			= 0.00;
	var dBaseInterestRate		= 0.00;
			
	sSequenceNumber		= m_IllustrationListXML.GetTagText("INTERESTRATETYPESEQUENCENUMBER")
	sInterestRateType	= m_IllustrationListXML.GetTagText("RATETYPE");
	sInterestRatePeriod = m_IllustrationListXML.GetTagText("INTERESTRATEPERIOD");
	if(sInterestRatePeriod == "-1") sInterestRatePeriod = "Remaining Term";
	sInterestRateEndDate= m_IllustrationListXML.GetTagText("INTERESTRATEENDDATE");
	dInterestRate		= m_IllustrationListXML.GetTagFloat("RATE");
			
	if (sInterestRateType != "F")			// Fixed rate 
	{
		var dCeilingRate = 0.00;
		var dFlooredRate = 0.00;
				
		dCeilingRate	= m_IllustrationListXML.GetTagFloat("CEILINGRATE");
		dFlooredRate	= m_IllustrationListXML.GetTagFloat("FLOOREDRATE");
				
		m_IllustrationListXML.SelectTag(tagILLUSTRATION, "BASERATEBAND");
		dBaseInterestRate = m_IllustrationListXML.GetTagFloat("RATE");
					
		switch (sInterestRateType)
		{
		case "B":						// Base Rate
			sInterestRateTypeText = "Base";
			dInterestRate = dBaseInterestRate;
			break;
						
		case "D":						// Discount Rate
			if (sSequenceNumber == "1")
			{
				<%/* SG 06/06/02 SYS4800 */%>
				<%/* frmScreen.txtDiscount.value = scMath.RoundValue(dInterestRate,2);*/%>
				frmScreen.txtDiscount.value = top.frames[1].document.all.scMathFunctions.RoundValue(dInterestRate,2);			
			}
			sInterestRateTypeText = "Discounted";
			dInterestRate = dBaseInterestRate - dInterestRate;
			break;
						
		case "C":						// Capped/Floored
			if (sSequenceNumber == "1")
			{
				<%/* SG 06/06/02 SYS4800 */%>
				<%/* frmScreen.txtDiscount.value = scMath.RoundValue(dInterestRate,2);*/%>
				frmScreen.txtDiscount.value = top.frames[1].document.all.scMathFunctions.RoundValue(dInterestRate,2);
			}
			sInterestRateTypeText = "Capped/Floored";
			dInterestRate = dBaseInterestRate - dInterestRate;
			if (dInterestRate > dCeilingRate)
				dInterestRate = dCeilingRate;
			else if (dInterestRate < dFlooredRate)
				dInterestRate = dFlooredRate;
			break;
						
		default:
			dInterestRate = "0.00";
			break;	
		}
	}
	<%/* SG 06/06/02 SYS4800 */%>
	<%/* dInterestRate = scMath.RoundValue(dInterestRate, 2);*/%>									
	dInterestRate = top.frames[1].document.all.scMathFunctions.RoundValue(dInterestRate, 2);
	
	ShowRow(iLoop+1, sInterestRateTypeText, dInterestRate, 
				sInterestRatePeriod, sInterestRateEndDate);
}	
		
function ShowRow(iRow, sInterestRateTypeText, dInterestRate, 
				sInterestRatePeriod, sInterestRateEndDate)
{
	<% /*
		Description:	Displays a row in the illustration table
		Args Passed:	iRow					
						sInterestRateTypeText
						dInterestRate
						sInterestRatePeriod
						sInterestRateEndDate
	*/ %>

	scScreenFunctions.SizeTextToField(tblInterestRateTable.rows(iRow).cells(0),
										sInterestRateTypeText);
	scScreenFunctions.SizeTextToField(tblInterestRateTable.rows(iRow).cells(1),
										dInterestRate);
	scScreenFunctions.SizeTextToField(tblInterestRateTable.rows(iRow).cells(2),
										sInterestRatePeriod);
	scScreenFunctions.SizeTextToField(tblInterestRateTable.rows(iRow).cells(3),
										sInterestRateEndDate);	
}
		
function frmScreen.btnOneOffCosts.onclick()
{
	var sReturn				= null;			
	var ArrayArguments		= new Array();
			
	ArrayArguments[0] = m_sApplicationMode;
	ArrayArguments[1] = m_sOneOffCosts;
	ArrayArguments[2] = m_sTotalIncentives;
	ArrayArguments[3] = frmScreen.txtPurchasePrice.value;
	ArrayArguments[4] = frmScreen.txtAmountRequested.value;
	ArrayArguments[5] = m_sUserId;
	ArrayArguments[6] = m_sUserType;
	ArrayArguments[7] = m_sUnitId;
	ArrayArguments[8] = m_sCurrency;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "cm130.asp", 
												ArrayArguments, 430, 510);
}
function frmScreen.btnIncentives.onclick()
{
	var sReturn = null;			
	var ArrayArguments = new Array();
	var bContinue = true;

	if(bContinue)
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		ArrayArguments[0] = m_sApplicationMode;
		ArrayArguments[1] = "0";  // no read only status for Mortgage Calculator
		ArrayArguments[2] = XML.CreateRequestAttributeArray(window);
		ArrayArguments[3] = frmScreen.txtProductCode.value;
		ArrayArguments[4] = m_sStartDate;
		ArrayArguments[5] = frmScreen.txtAmountRequested.value;
		ArrayArguments[6] = m_sCurrency;

		sReturn = scScreenFunctions.DisplayPopup(window, document, "cm150.asp", ArrayArguments, 500, 410);
		if (sReturn != null)
		{
			m_sNEWINCENTIVESXML = sReturn[2];
		}
		XML = null;
	}
}		
function btnSubmit.onclick()
{
	<% /* RF 24/11/99 */%>
	m_IllustrationListXML = null;
			
	frmStoredIllustrations.submit();
}

-->
</script>
</body>
</html>
