<%@ LANGUAGE="JSCRIPT" %>
<html>

<% /*
Workfile:      mc020.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Stored illustrations
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		24/11/99	AQR 14: Centre "Y" in "Current?" column of list box
					AQR 15: Remove "Cancel" button
RF		02/02/00	Reworked for IE5 and performance.
AY		14/02/00	Change to msgButtons button types
AY		03/04/00	New top menu/scScreenFunctions change
					OK button was overlapping div
AY		13/04/00	SYS0328 - Dynamic currency display
CL		05/03/01	SYS1920 Read only functionality added
APS		05/03/01	SYS1920 Screen Title changes
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>
<script src="validation.js" language="JScript"></script>
<object data="scTable.htm" height="1" id="scIllustrationTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>

<% /* List Scroll object */%>
<span style="LEFT: 310px; POSITION: absolute; TOP: 280px">
	<OBJECT data=scListScroll.htm id=scScrollPlus 
	style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 
	type=text/x-scriptlet VIEWASTEXT></OBJECT>
</span>
	
<% /* Specify Forms Here */%>
<form id="frmMortgageCalculatorDetails" method="post" action="mc010.asp" STYLE="DISPLAY: none"></form>
<form id="frmIllustrationDetails" method="post" action="mc030.asp" STYLE="DISPLAY: none"></form>	
	
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */%>
<form id="frmScreen" mark validate="onchange" year4>
	<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 255px; 
		WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
			
		<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Illustration List...</strong>
		</span>
			
		<span id="spnStoredIllustrations" style="TOP: 30px; LEFT: 4px; POSITION: ABSOLUTE">
			<table id="tblStoredIllustrations" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
				<tr id="rowTitles">	<td width="5%" class="TableHead">No.</td>	
									<td width="30%" class="TableHead">Product Name</td>	
									<td width="10%" class="TableHead">Interest Type</td>	
									<td width="10%" class="TableHead">Purchase Price/Est.<br>Value</td>
									<td width="10%" class="TableHead">Amount Requested</td>
									<td id="tdIntOnly" width="10%" class="TableHead">Int Only</td>
									<td id="tdCapInt" width="10%" class="TableHead">Cap & Int</td>
									<td id="tdPart" width="10%" class="TableHead">Part & Part</td>										
									<td class="TableHead">Current?</td>
				</tr>
				<tr id="row01">		<td width="5%" class="TableTopLeft">&nbsp</td><td width="30%" class="TableTopCenter">&nbsp</td>
									<td width="10%" class="TableTopCenter">&nbsp</td><td width="10%" class="TableTopCenter">&nbsp</td>
									<td width="10%" class="TableTopCenter">&nbsp</td><td width="10%" class="TableTopCenter">&nbsp</td>
									<td width="10%" class="TableTopCenter">&nbsp</td><td width="10%" class="TableTopCenter">&nbsp</td>
									<td class="TableTopRight" align="center">&nbsp</td>
				</tr>
				<tr id="row02">		<td width="5%" class="TableLeft">&nbsp</td><td width="30%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td class="TableRight" align="center">&nbsp</td>
				</tr>
				<tr id="row03">		<td width="5%" class="TableLeft">&nbsp</td><td width="30%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td class="TableRight" align="center">&nbsp</td>
				</tr>
				<tr id="row04">		<td width="5%" class="TableLeft">&nbsp</td><td width="30%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td class="TableRight" align="center">&nbsp</td>
				</tr>
				<tr id="row05">		<td width="5%" class="TableLeft">&nbsp</td><td width="30%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td class="TableRight" align="center">&nbsp</td>
				</tr>
				<tr id="row06">		<td width="5%" class="TableLeft">&nbsp</td><td width="30%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td class="TableRight" align="center">&nbsp</td>
				</tr>
				<tr id="row07">		<td width="5%" class="TableLeft">&nbsp</td><td width="30%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td class="TableRight" align="center">&nbsp</td>
				</tr>
				<tr id="row08">		<td width="5%" class="TableLeft">&nbsp</td><td width="30%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td class="TableRight" align="center">&nbsp</td>
				</tr>
				<tr id="row09">		<td width="5%" class="TableLeft">&nbsp</td><td width="30%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td>
									<td class="TableRight" align="center">&nbsp</td>
				</tr>					
				<tr id="row10">		<td width="5%" class="TableBottomLeft">&nbsp</td><td width="30%" class="TableBottomCenter">&nbsp</td>
									<td width="10%" class="TableBottomCenter">&nbsp</td><td width="10%" class="TableBottomCenter">&nbsp</td>
									<td width="10%" class="TableBottomCenter">&nbsp</td><td width="10%" class="TableBottomCenter">&nbsp</td>
									<td width="10%" class="TableBottomCenter">&nbsp</td><td width="10%" class="TableBottomCenter">&nbsp</td>
									<td class="TableBottomRight" align="center">&nbsp</td>
				</tr>
			</table>
		</span>
		
		<span style="TOP: 220px; LEFT: 4px; POSITION: ABSOLUTE">
			<input id="btnDetails" value="Details" type="button" 
				style="WIDTH: 60px" class="msgButton" disabled >
		</span>
			
		<span style="TOP: 220px; LEFT: 65px; POSITION: ABSOLUTE">
			<input id="btnReinstate" value="Reinstate" type="button" 
				style="WIDTH: 60px" class="msgButton" disabled >
		</span>
			
	</div>
</form>

<% /* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--	// JScript Code
var m_sMetaAction				= null;
var m_sIllustrationList			= null;		
var m_IllustrationListXML		= null;
var m_iTableLength				= 10;
var m_iIllustrationNumber		= 0;
var scScreenFunctions;

function window.onload()
{
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Mortgage Calculator","MC020",scScreenFunctions);

	SetCurrency();
	RetrieveContextData();
	PopulateScreen();
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

function SetCurrency()
{
	var sCurrencyText = scScreenFunctions.SetCurrency(window,frmScreen);
	tdIntOnly.innerText = tdIntOnly.innerText + " " + sCurrencyText;
	tdCapInt.innerText = tdCapInt.innerText + " " + sCurrencyText;
	tdPart.innerText = tdPart.innerText + " " + sCurrencyText;
}

function RetrieveContextData()
{
	m_sMetaAction			= scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sIllustrationList		= scScreenFunctions.GetContextParameter(window,"idXML",null);
	m_iIllustrationNumber	= scScreenFunctions.GetContextParameter(window,"idMCIllustrationNumber","0");
}
		
function PopulateScreen()
{
<%	/* Display the illustration data held in XML to the screen controls */
%>
	var iNumberOfIllustrations = 0;
			
	scIllustrationTable.initialise(tblStoredIllustrations,0,"");
						
	m_IllustrationListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();			
	m_IllustrationListXML.LoadXML(m_sIllustrationList);			
			
<% /* Display the illustrations starting from the beginning */
%>
	DisplayIllustrations(0);			
			
	iNumberOfIllustrations = m_IllustrationListXML.ActiveTagList.length;
			
	scScrollPlus.Initialise(DisplayIllustrations, m_iTableLength, 
								iNumberOfIllustrations);
}

/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
function DisplayIllustrations(iStart)
{
	<% /*
		Description:	Display the illustration data held in XML to the screen
						controls
		Args Passed:	iStart	- The offset used to display the XML in the table
	*/ %>

	var tagILLUSTRATIONLIST		= null;			
	var sPurchasePrice			= null;
	var sAmountRequested		= null;
	var sInterestOnlyCost		= null;
	var sCapitalAndInterestCost	= null;
	var sPartAndPartCost		= null;
	var sCurrentIllustration	= null;
	var iIllustrationNumber		= 0;
	var iNumberOfIllustrations	= 0;
			
	m_IllustrationListXML.ActiveTag = null;
	tagILLUSTRATIONLIST = m_IllustrationListXML.CreateTagList("ILLUSTRATION");

	iNumberOfIllustrations = m_IllustrationListXML.ActiveTagList.length;
			
	for (var iLoop = 0; iLoop < iNumberOfIllustrations && 
			iLoop < m_iTableLength; iLoop++)
	{
		m_IllustrationListXML.ActiveTagList = tagILLUSTRATIONLIST;
				
		iIllustrationNumber = iLoop + iStart;
		if (m_IllustrationListXML.SelectTagListItem(iIllustrationNumber) == true)
		{
			sIllustrationNumber		= iIllustrationNumber;
			sPurchasePrice			= m_IllustrationListXML.GetTagText("PURCHASEPRICE");
			sAmountRequested		= m_IllustrationListXML.GetTagText("AMOUNTREQUESTED");
			sInterestOnlyCost		= m_IllustrationListXML.GetTagText("MONTHLYINTERESTONLYCOST");
			sCapitalAndInterestCost = m_IllustrationListXML.GetTagText("MONTHLYCAPITALANDINTERESTCOST");
			sPartAndPartCost 		= m_IllustrationListXML.GetTagText("MONTHLYPARTANDPARTCOST");
			sProductName			= m_IllustrationListXML.GetTagText("PRODUCTNAME");
					
			sInterestRateType		= GetInterestRateType();
					
			if (iIllustrationNumber + 1 == m_iIllustrationNumber)
				sCurrentIllustration	= "Y";
			else
				sCurrentIllustration	= " ";
															
			ShowRow(iLoop + 1, sIllustrationNumber + 1, sProductName, sPurchasePrice, 
					sAmountRequested, sInterestOnlyCost, sCapitalAndInterestCost,
					sPartAndPartCost, sInterestRateType, sCurrentIllustration);
		}				
	}			
}
		
function GetInterestRateType()
{					
	<% /* Description:	Retrieves the interest rate description */%>

	var sRateTypeDescription = "";
			
	<% /* for the first interest rate only */ %>
	var sInterestRateType = m_IllustrationListXML.GetTagText("RATETYPE");
			
	switch (sInterestRateType)
	{
	case "F":
		sRateTypeDescription = "Fixed";
		break;
	case "B":
		sRateTypeDescription = "Base";
		break;
	case "D":
		sRateTypeDescription = "Discounted";
		break;
	case "C":
		sRateTypeDescription = "Capped/Floored";
		break;
	default:
		sRateTypeDescription = "Error - not specified";			
	}
			
	return sRateTypeDescription;			
}
		
function ShowRow(iRow, sIllustrationNumber, sProductName, sPurchasePrice, 
				sAmountRequested, sInterestOnlyCost, sCapitalAndInterestCost, 
				sPartAndPartCost, sInterestRateType, sCurrentIllustration)
{
	<% /*
		Description:	Displays a row in the illustration table
		Args Passed:	iRow					
						sIllustrationNumber		
						sProductName
						sPurchasePrice
						sAmountRequested
						sInterestOnlyCost
						sCapitalAndInterestCost
						sPartAndPartCost
						sInterestRateType		
						sCurrentIllustration	(Y or blank)
	*/ %>
	
	scScreenFunctions.SizeTextToField(tblStoredIllustrations.rows(iRow).cells(0),
										sIllustrationNumber);
	scScreenFunctions.SizeTextToField(tblStoredIllustrations.rows(iRow).cells(1),
										sProductName);
	scScreenFunctions.SizeTextToField(tblStoredIllustrations.rows(iRow).cells(2),
										sInterestRateType);
	scScreenFunctions.SizeTextToField(tblStoredIllustrations.rows(iRow).cells(3),
										sPurchasePrice);
	scScreenFunctions.SizeTextToField(tblStoredIllustrations.rows(iRow).cells(4),
										sAmountRequested);
	scScreenFunctions.SizeTextToField(tblStoredIllustrations.rows(iRow).cells(5),
										sInterestOnlyCost);
	scScreenFunctions.SizeTextToField(tblStoredIllustrations.rows(iRow).cells(6),
										sCapitalAndInterestCost);
	scScreenFunctions.SizeTextToField(tblStoredIllustrations.rows(iRow).cells(7),
										sPartAndPartCost);
	scScreenFunctions.SizeTextToField(tblStoredIllustrations.rows(iRow).cells(8),
										sCurrentIllustration);
			
}
		
		
function IsCurrentIllustration()
{
	<% /* 
		Description:	Do we have the current illustration selected?
		Returns:		TRUE if the selected illustration is
						the current illustration otherwise FALSE
	*/	%>
	
	var blnReturn = false;
			
	var iSelectedRow = scIllustrationTable.getRowSelected();
			
	if ("Y" == tblStoredIllustrations.rows(iSelectedRow).cells(8).innerText)
		blnReturn = true;
			
	return blnReturn;
}
		
function spnStoredIllustrations.onclick()
{
	<% /* Description:	Event handler for the table click event */%>

	if (scIllustrationTable.getRowSelected() != -1)
	{
		frmScreen.btnDetails.disabled = false;
				
		if (!IsCurrentIllustration())
			frmScreen.btnReinstate.disabled = false;
		else
			frmScreen.btnReinstate.disabled = true;
	}
}		
		
function frmScreen.btnDetails.onclick()
{
	var iOffset = scScrollPlus.getOffset() + scIllustrationTable.getRowSelected();

	<% /* set the display number for mc030 */%>
	scScreenFunctions.SetContextParameter(window,"idMCDisplayNumber", iOffset);
			
	frmIllustrationDetails.submit();
}
		
function frmScreen.btnReinstate.onclick()
{
	var iSelectedRow = scIllustrationTable.getRowSelected();
			
	if (iSelectedRow != -1)
	{
		<% /* clear the eighth column for the table; this is easier than
			  finding the previous selected illustration */ %>
		
		for (var iLoop=1; iLoop < m_iTableLength; iLoop++)
			tblStoredIllustrations.rows(iLoop).cells(8).innerText = " ";
			
		tblStoredIllustrations.rows(iSelectedRow).cells(8).innerText = "Y";
							
		m_iIllustrationNumber = scScrollPlus.getOffset() + iSelectedRow;
		
		scScreenFunctions.SetContextParameter(window,"idMCIllustrationNumber",m_iIllustrationNumber);

		
	}
	else
		alert("No row selected");
}
		
function btnSubmit.onclick()
{
	mIllustrationListXML = null;
			
	<% /* set the illustration number */ %>
	scScreenFunctions.SetContextParameter(window,"idMCIllustrationNumber",
												m_iIllustrationNumber);
			
	frmMortgageCalculatorDetails.submit();
}
		
-->
</script>
</body>
</html>
