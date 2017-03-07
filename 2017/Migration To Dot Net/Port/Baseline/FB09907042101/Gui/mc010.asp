<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      mc010.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Mortgage Calculator screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		23/11/99	Write xml to file if debugging.
					AQR MC18: Ensure correct LTV value is used in search.
RF		24/11/99	AQR MC6: Move "Amount Requested" label.
RF		24/11/99	AQR MC9: Ensure split amounts initialised correctly on entry
					to screen cm160.
RF		25/11/99	AQR MC4: List box scrolling incorrect so change to using 
					scTableListScroll.
					AQR MC24: Clear out old calc values in case calc fails.
RF		02/02/00	Reworked for IE5 and performance.
AY		14/02/00	Change to msgButtons button types
APS		16/02/00	Change xml manipulation as a result of the output xml change from 
					FindMortgageProducts.asp
JLD		23/02/2000	AQR SYS00284
JLD		25/02/2000	AQR SYS0292, SYS0315
JLD		29/02/00	SYS0287 increased size of CM130
					SYS0299 various minor tweaks.
JLD		01/03/00	SYS0376 minor tweaks
					SYS0372 Calculate the total incentives after hitting calc button
JLD		02/03/00	SYS0312 Order the products returned.
					SYS0299 Make listbox header only 2 lines thick. Move Split and LTV buttons slightly.
JLD		10/03/00	SYS0372 Cater for INCENTIVESLIST being null.
JLD		10/03/00	SYS0293 Remove LTV button and calc LTV on entering purchase price and amt Requested.
					Also, tidy screen up a bit.
JLD		14/03/00	SYS0407 Make sure total incentives gets passed to CM150 via stored illustrations
AY		17/03/00	MetaAction changed to ApplicationMode
AY		03/04/00	New top menu/scScreenFunctions change
AY		13/04/00	SYS0328 - Dynamic currency display
JLD		17/04/00	Change to interface with CM150
IW		24/05/00	SYS0774 DISTRIBUTIONCHANNELID S/B CHANNELID
PSC		08/06/00	SYS0859 Backout SYS0774
PSC     20/06/00    SYS0763 - Page products correctly
ADP		30/06/00	SYS0861 - OK to route back to MN010.  Cancel removed.
CL		05/03/01	SYS1920 Read only functionality added
APS		13/03/01	SYS1920 Screen Title changes
MC		05/10/01	SYS2775 Enable client versions to override labels
JLD		10/12/01	SYS2806 Completeness Check routing
STB		11/04/02	SYS4139 Correct error message for missing part-and-part split information.
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		06/06/02	SYS4800 - Error following SYS4727
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom History:

Prog    Date           Description
SR		04/03/2007	EP2_1644	increased the width of scScrollPlus control
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<script src="validation.js" language="JScript"></script>

<% /* List Box object */%>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="LEFT: 302px; POSITION: absolute; TOP: 365px" id="spnPageScroll">
	<object data="scPageScroll.htm" id="scScrollPlus" style="LEFT: 0px; TOP: 0px; height:24; width:304" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>
	
<% /* Specify Forms Here */ %>
<form id="frmWelcomeMenu" method="post" action="mn010.asp" STYLE="DISPLAY:none"></form>
<form id="frmCustomerSearch" method="post" action="cr010.asp" STYLE="DISPLAY:none"></form>
<form id="frmStoredIllustrations" method="post" action="mc020.asp" STYLE="DISPLAY:none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY:none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */%>
<form id="frmScreen" mark validate="onchange" year4>
	<div id="divBackground" 
		style="HEIGHT:400px; LEFT:10px; POSITION:absolute; TOP:60px; WIDTH:604px">
						
		<div id="divLoanDetails" 
			style="HEIGHT: 110px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 604px"
			 class="msgGroup">
				
			<span style="LEFT: 4px; POSITION: absolute; TOP: 6px" class="msgLabel">
				Purchase Price/<br>Estimated Value
				<span style="LEFT: 90px; POSITION: absolute; TOP: 5px">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
					<input id="txtPurchasePrice" name="PurchasePrice" maxlength="7" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 65px" class="msgTxt"> 
				</span> 
			</span>
				
			<span style="LEFT: 175px; POSITION: absolute; TOP: 11px" class="msgLabel">
				Amount Requested
				<span style="LEFT: 110px; POSITION: absolute; TOP: 0px">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
					<input id="txtAmountRequested" name="AmountRequested" maxlength="7" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 65px" class="msgTxt"> 
				</span> 
			</span>
				
			<span style="LEFT: 375px; POSITION: absolute; TOP: 11px" class="msgLabel">
				<LABEL id=idTerm></LABEL> 
				<span id="spnTerm" style="LEFT: 70px; POSITION: absolute; TOP: 0px">						
					<input id="txtTermYears" name="TermYears" maxlength="2" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 35px" class="msgTxt"> 
					<span style="LEFT: 37px; POSITION: absolute; TOP: 0px" class="msgLabel">
						Years 
					</span> 
				</span> 					
			</span>
				
			<span style="LEFT: 515px; POSITION: absolute; TOP: 6px" class="msgLabel">					
				<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">						
					<input id="txtTermMonths" name="TermMonths" maxlength="2" 
						style="POSITION: absolute; TOP: 2px; WIDTH: 35px" class="msgTxt"> 
					<span style="LEFT: 40px; POSITION: absolute; TOP: 5px" class="msgLabel">
						Months
					</span> 
				</span> 					
			</span>
								
			<span id="spnPartAndPart" style="LEFT: 4px; POSITION: absolute; TOP: 43px" class="msgLabel">
				Do you want Part &amp;<BR> Part costs to be <BR>illustrated?
				<span style="LEFT: 110px; POSITION: absolute; TOP: 3px">
					<input id="optPartAndPartYes" name="PartAndPart" type="radio" value="1">
					<label for="optPartAndPartYes" class="msgLabel">Yes</label> 
				</span> 
				<span style="LEFT: 160px; POSITION: absolute; TOP: 3px">
					<input id="optPartAndPartNo" name="PartAndPart" type="radio" checked value="0">
					<label for="optPartAndPartNo" class="msgLabel">No</label> 
				</span> 
				<span style="LEFT: 205px; POSITION: absolute; TOP: 3px">
					<input id="btnPartAndPartSplit" value="Split" type="button"
						 style ="WIDTH: 60px" class="msgButton"> 
				</span> 					
			</span>
				
			<span style="LEFT: 4px; POSITION: absolute; TOP: 91px" class="msgLabel">
				LTV
				<span style="LEFT: 112px; POSITION: absolute; TOP: 0px">													
					<input id="txtLTV" name="LTV" maxlength="3" 
						style="POSITION: absolute; TOP: -3px; WIDTH: 40px" 
						class="msgReadOnly" tabindex="-1"> 
					<span style="LEFT: 45px; POSITION: absolute; TOP: 0px" class="msgLabel">
						%
					</span> 
				</span>
			</span>								
								 
			<span style="LEFT: 325px; POSITION: absolute; TOP: 91px">
				<span style="LEFT: 0px; POSITION: absolute; TOP: -6px">
					<input id="btnSearch" value="Search" type="button" 
						style="WIDTH: 60px" class="msgButton"> 					
				</span>
				<span style="LEFT: 65px; POSITION: absolute; TOP: -6px">
					<input id="btnFurtherFiltering" value="FurtherFiltering" type="button" 
						style ="WIDTH: 100px" class="msgButton"> 					
				</span>
				<span style="LEFT: 170px; POSITION: absolute; TOP: -6px">
					<input id="btnStoredIllustrations" value="Stored Illustrations" type="button" 
						style="WIDTH: 100px" class="msgButton" disabled>
				</span> 
			</span>				
		</div>
			
		<div id="divAvailableProducts" 
			style="HEIGHT: 220px; LEFT: 0px; POSITION: absolute; TOP: 115px; WIDTH: 604px"
			 class="msgGroup">
				
			<span id="spnAvailableProducts" style="LEFT: 4px; POSITION: absolute; TOP: 10px">
				<% /* This is the multiple lender table version */%>
				<table id="tblMultipleLenderTable" 
					width="596" border="0" cellspacing="0" cellpadding="0" 
					style="LEFT: 0px; POSITION: absolute; TOP: 0px; z-Index:0"
					class="msgTable">
						
					<tr id="rowTitles">								
						<td width="20%" class="TableHead">Lender</td>
						<td width="20%" class="TableHead">Product Name</td>							
						<td width="20%" class="TableHead">First Interest Rate %</td>	
						<td width="20%" class="TableHead">Period/End Date</td>	
						<td class="TableHead">Flexible Mortgage?</td></tr>
					<tr id="row01">		
						<td width="35%" class="TableTopLeft">&nbsp</td>
						<td width="35%" class="TableTopCenter">&nbsp</td>
						<td width="10%" class="TableTopCenter">&nbsp</td>
						<td width="15%" class="TableTopCenter">&nbsp</td>
						<td class="TableTopRight">&nbsp</td></tr>
					<tr id="row02">		
						<td width="35%" class="TableLeft">&nbsp</td>
						<td width="35%" class="TableCenter">&nbsp</td>
						<td width="10%" class="TableCenter">&nbsp</td>
						<td width="15%" class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row03">		
						<td width="35%" class="TableLeft">&nbsp</td>
						<td width="35%" class="TableCenter">&nbsp</td>
						<td width="10%" class="TableCenter">&nbsp</td>
						<td width="15%" class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row04">		
						<td width="35%" class="TableLeft">&nbsp</td>				
						<td width="35%" class="TableCenter">&nbsp</td>			
						<td width="10%" class="TableCenter">&nbsp</td>
						<td width="15%" class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row05">		
						<td width="35%" class="TableLeft">&nbsp</td>
						<td width="35%" class="TableCenter">&nbsp</td>
						<td width="10%" class="TableCenter">&nbsp</td>
						<td width="15%" class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row06">		
						<td width="35%" class="TableLeft">&nbsp</td>
						<td width="35%" class="TableCenter">&nbsp</td>
						<td width="10%" class="TableCenter">&nbsp</td>
						<td width="15%" class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row07">		
						<td width="35%" class="TableLeft">&nbsp</td>
						<td width="35%" class="TableCenter">&nbsp</td>
						<td width="10%" class="TableCenter">&nbsp</td>
						<td width="15%" class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row08">		
						<td width="35%" class="TableLeft">&nbsp</td>
						<td width="35%" class="TableCenter">&nbsp</td>
						<td width="10%" class="TableCenter">&nbsp</td>
						<td width="15%" class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row09">		
						<td width="35%" class="TableLeft">&nbsp</td>
						<td width="35%" class="TableCenter">&nbsp</td>
						<td width="10%" class="TableCenter">&nbsp</td>
						<td width="15%" class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row10">		
						<td width="35%" class="TableBottomLeft">&nbsp</td>
						<td width="35%" class="TableBottomCenter">&nbsp</td>
						<td width="10%" class="TableBottomCenter">&nbsp</td>	
						<td width="15%" class="TableBottomCenter">&nbsp</td>	
						<td class="TableBottomRight">&nbsp</td></tr>
				</table>
				<% /* This is the single lender table version */%>
				<table id="tblSingleLenderTable" 
					width="596" border="0" cellspacing="0" cellpadding="0" 
					style="LEFT: 0px; POSITION: absolute; TOP: 0px; z-Index:1"
					class="msgTable">
						
					<tr id="rowTitles">															
						<td width="50%" class="TableHead">Product Name</td>							
						<td width="20%" class="TableHead">First Interest Rate %</td>	
						<td width="25%" class="TableHead">Period/<br>End Date</td>	
						<td class="TableHead">Flexible Mortgage?</td></tr>
					<tr id="row01">		
						<td class="TableTopLeft">&nbsp</td>							
						<td class="TableTopCenter">&nbsp</td>
						<td class="TableTopCenter">&nbsp</td>
						<td class="TableTopRight">&nbsp</td></tr>
					<tr id="row02">		
						<td class="TableLeft">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row03">									
						<td class="TableLeft">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row04">		
						<td class="TableLeft">&nbsp</td>			
						<td class="TableCenter">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row05">		
						<td class="TableLeft">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row06">		
						<td class="TableLeft">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row07">		
						<td class="TableLeft">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row08">		
						<td class="TableLeft">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row09">		
						<td class="TableLeft">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableCenter">&nbsp</td>
						<td class="TableRight">&nbsp</td></tr>
					<tr id="row10">		
						<td class="TableBottomLeft">&nbsp</td>
						<td class="TableBottomCenter">&nbsp</td>	
						<td class="TableBottomCenter">&nbsp</td>	
						<td class="TableBottomRight">&nbsp</td></tr>
				</table>
			</span> 
				
			<span id=spnTableButtons style="LEFT: 4px; POSITION: absolute; TOP: 190px">
				<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
					<input id="btnCalculate" value="Calculate" 
						type="button" style="WIDTH: 60px" class="msgButton"> 					
				</span> 
				<span style="LEFT: 65px; POSITION: absolute; TOP: 0px">
					<input id="btnOneOffCosts" value="One-Off Costs" type="button" 
						style="WIDTH: 100px" class="msgButton"> 					
				</span>
				<span style="LEFT: 170px; POSITION: absolute; TOP: 0px">
					<input id="btnIncentives" value="Incentives" type="button" 
						style="WIDTH: 60px" class="msgButton"> 					
				</span>
				<span style="LEFT: 235px; POSITION: absolute; TOP: 0px">
					<input id="btnDetails" value="Details" 
						type="button" style="WIDTH: 60px" class="msgButton"> 					
				</span> 
			</span> 
		</div>
			
		<div id="divCosts" 
			style="HEIGHT: 40px; LEFT: 0px; POSITION: absolute; TOP: 340px; WIDTH: 604px"
			 class="msgGroup">
				
			<span id=spnCosts style="LEFT: 4px; POSITION: absolute; TOP: 0px; WIDTH: 604px">										
				<span style="LEFT: 0px; POSITION: absolute; TOP: 5px" class="msgLabel">
					Monthly<br>Interest Only Cost
					<span style="LEFT: 110px; POSITION: absolute; TOP: 5px">						
						<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
						<input id="txtInterestOnlyCost" name="InterestOnlyCost" maxlength="7" 
							style="POSITION: absolute; TOP: -3px; WIDTH: 65px" 
							class="msgReadOnly" tabindex="-1"> 
					</span> 
				</span>
					
				<span style="LEFT: 210px; POSITION: absolute; TOP: 5px" class="msgLabel">
					<LABEL id="idMonthlyCICost"></LABEL> 
					<span style="LEFT: 100px; POSITION: absolute; TOP: 5px">						
						<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
						<input id="txtCapitalAndInterestCost" name="CapitalAndInterestCost" maxlength="7" 
							style="POSITION: absolute; TOP: -3px; WIDTH: 65px" 
							class="msgReadOnly" tabindex="-1"> 
					</span> 
				</span>
					
				<span id="spnPartAndPartCost" style="LEFT: 415px; POSITION: absolute; TOP: 5px" class="msgLabel">
					Monthly<br>Part &amp; Part Cost
					<span style="LEFT: 105px; POSITION: absolute; TOP: 5px">						
						<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
						<input id="txtPartAndPartCost" name="PartAndPartCost" maxlength="7" 
							style="POSITION: absolute; TOP: -3px; WIDTH: 65px" 
							class="msgReadOnly" tabindex="-1"> 
					</span> 
				</span> 	
			</span> 
		</div> 	
	</div> 
</form>

<% /* Main Buttons */%>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 445px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */%>
<!-- #include FILE="attribs/mc010attribs.asp" -->
<!--  #include FILE="Customise/MC010Customise.asp" -->

<% /* Specify Code Here */%>
<script language="JScript">
<!--	// JScript Code

var m_sDebug	= "";
var m_sApplicationMode	= "";		
var m_sDistributionChannelId		= "";
var m_sUnitId	= "";
var m_sUserId	= "";
var m_sUserType	= "";		
var m_sCurrency = "";
var	m_sFurtherFiltering	= "";
var m_sIllustrationList	= "";		
var m_sInterestOnlyPart	= "";
var m_sCapitalOnlyPart	= "";
var m_sIllustrationProductCode      = "";
var m_sIllustrationProductStartDate = "";
var m_sMultipleLender	= "";
var m_sTotalIncentives	= "";
var m_sOneOffCosts		= "";
var m_iIllustrationNumber	= 0;
var m_iNumberOfIllustrations	= 0;
var m_iTableLength			= 10;
var m_ProductXML			= null;
var m_IllustrationListXML	= null;
var m_tblProductTable		= null;
var m_iPageNo = 0;	
var m_blnSetButtonsOnNewQuote = true;
var scScreenFunctions;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit"); //ADP: AQR SYS0861
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Mortgage Calculator","MC010",scScreenFunctions);

	scScreenFunctions.SetCurrency(window,frmScreen);
	RetrieveContextData();
	SetMasks();
	// MC SYS2564/SYS2775 for client customisation
	Customise();

	Validation_Init();

	PopulateScreen();
			
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
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
	m_sApplicationMode				= "Mortgage Calculator";
	scScreenFunctions.SetContextParameter(window,"idApplicationMode",m_sApplicationMode);
			
	m_sDebug					= scScreenFunctions.GetContextParameter(window,"idDebug",null);
	m_sUnitId					= scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	m_sUserType					= scScreenFunctions.GetContextParameter(window,"idUserType",null);
	m_sUserId					= scScreenFunctions.GetContextParameter(window,"idUserId",null);
	m_sDistributionChannelId	= scScreenFunctions.GetContextParameter(window,"idDistributionChannelId","1");
	m_sIllustrationList			= scScreenFunctions.GetContextParameter(window,"idXML",null);
	m_iIllustrationNumber		= scScreenFunctions.GetContextParameter(window,"idMCIllustrationNumber","0");
	m_sMultipleLender			= scScreenFunctions.GetContextParameter(window,"idMultipleLender","0");
	m_sCurrency = scScreenFunctions.GetContextParameter(window,"idCurrency",null);
}

function PopulateScreen()
{
	m_IllustrationListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	if (m_sIllustrationList != "")
		m_IllustrationListXML.LoadXML(m_sIllustrationList);
	else
		m_IllustrationListXML.CreateActiveTag("ILLUSTRATIONLIST");
						
	if (m_sMultipleLender == "1")
		m_tblProductTable = tblMultipleLenderTable;				
	else
		m_tblProductTable = tblSingleLenderTable;
			
	m_tblProductTable.style.zIndex="2";
	
	InitialiseTable();
	scScrollPlus.Clear();

	<% /*scAvailableProductsTable.initialise(m_tblProductTable,0,""); */%>
			
	if (ReadIllustration() == true)
	{
		m_blnSetButtonsOnNewQuote = false;
		scScrollPlus.InitialiseAtSection(DoProductSearch,DisplayProducts,m_tblProductTable.rows.length - 1,1,m_iPageNo);
		
		if (SelectIllustratedProduct() != true)
			alert("Unable to reselect illustrated product");
		else
			SetButtonsOnExistingQuote();
			
		m_blnSetButtonsOnNewQuote = true;			
	}
	else
		SetButtonsOnNewQuote();

	SetPartAndPartSplitEnabling();
	SetFieldsToReadOnly();
}

function SetFieldsToReadOnly()
{
	scScreenFunctions.SetCollectionToReadOnly(spnCosts);			
	scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.txtLTV.id);			
}

function ReadIllustration()
{			
	var blnFound = false;
						
	m_IllustrationListXML.CreateTagList("ILLUSTRATION");
			
	m_iNumberOfIllustrations = m_IllustrationListXML.ActiveTagList.length;
			
	if (m_IllustrationListXML.SelectTagListItem(m_iIllustrationNumber-1) == true)
	{
		var tagILLUSTRATION = m_IllustrationListXML.ActiveTag;
				
		m_iPageNo = m_IllustrationListXML.GetTagText("PRODUCTPAGE"); 

		frmScreen.txtAmountRequested.value = m_IllustrationListXML.GetTagText("AMOUNTREQUESTED");
		frmScreen.txtPurchasePrice.value = m_IllustrationListXML.GetTagText("PURCHASEPRICE");
		frmScreen.txtTermYears.value = m_IllustrationListXML.GetTagText("TERMINYEARS");
		frmScreen.txtTermMonths.value = m_IllustrationListXML.GetTagText("TERMINMONTHS");
		
		<%/* SG 06/06/02 SYS4800 */%>
		<%/* frmScreen.txtLTV.value	= scMath.RoundValue(m_IllustrationListXML.GetTagText("LTV"),2); */%>
		frmScreen.txtLTV.value	= top.frames[1].document.all.scMathFunctions.RoundValue(m_IllustrationListXML.GetTagText("LTV"),2);
				
		scScreenFunctions.SetRadioGroupValue(frmScreen,"PartAndPart", 
								m_IllustrationListXML.GetTagText("PARTANDPARTREQUIRED"));
				
		m_sInterestOnlyPart = m_IllustrationListXML.GetTagText("INTERESTONLYELEMENT");
		m_sCapitalOnlyPart	= m_IllustrationListXML.GetTagText("CAPITALANDINTERESTELEMENT");
				
		if (m_IllustrationListXML.SelectTag(tagILLUSTRATION, "REGULARCOSTS") != null)
		{
			frmScreen.txtInterestOnlyCost.value			= m_IllustrationListXML.GetTagText("MONTHLYINTERESTONLYCOST");
			frmScreen.txtCapitalAndInterestCost.value	= m_IllustrationListXML.GetTagText("MONTHLYCAPITALANDINTERESTCOST");
			frmScreen.txtPartAndPartCost.value			= m_IllustrationListXML.GetTagText("MONTHLYPARTANDPARTCOST");
		}
								
		if (m_IllustrationListXML.SelectTag(tagILLUSTRATION, "FURTHERFILTERING") != null)
			m_sFurtherFiltering = m_IllustrationListXML.ActiveTag.xml;
				
		if (m_IllustrationListXML.SelectTag(tagILLUSTRATION, "MORTGAGEPRODUCT") != null)
		{
			m_sIllustrationProductCode		= m_IllustrationListXML.GetTagText("MORTGAGEPRODUCTCODE");
			m_sIllustrationProductStartDate = m_IllustrationListXML.GetTagText("STARTDATE");					
		}
				
		m_sTotalIncentives = m_IllustrationListXML.GetTagText("TOTALINCENTIVES");
				
		if (m_IllustrationListXML.SelectTag(tagILLUSTRATION, "ONEOFFCOSTLIST") != null)
			m_sOneOffCosts = m_IllustrationListXML.ActiveTag.xml;
				
		blnFound = true;
	}			
			
	return blnFound;
}		
		
function frmScreen.txtPurchasePrice.onchange()
{
	InitialiseTable();
	scTable.clear();
	SetButtonsOnNewQuote();
	frmScreen.txtLTV.value = "";
	CalcLTV();
}
		
function frmScreen.txtAmountRequested.onchange()
{
	InitialiseTable();
	scTable.clear();
	SetButtonsOnNewQuote();
	frmScreen.txtLTV.value = "";
	m_sInterestOnlyPart = "";
	m_sCapitalOnlyPart = "";
	frmScreen.optPartAndPartNo.checked = true
	frmScreen.optPartAndPartYes.checked = false;
	frmScreen.btnPartAndPartSplit.disabled = true;
	CalcLTV();
}
		
function frmScreen.txtTermYears.onchange()
{
	InitialiseTable();
	scTable.clear();
	SetButtonsOnNewQuote();
}
		
function frmScreen.txtTermMonths.onchange()
{
	InitialiseTable();
	scTable.clear();
	SetButtonsOnNewQuote();
}
		
function frmScreen.optPartAndPartNo.onclick()
{
	SetButtonsOnNewQuote();
	//InitialiseTable();
			
	<% /*
		// RF 24/11/99 AQR MC8
		//frmScreen.btnPartAndPartSplit.disabled = true;
	*/ %>
	SetPartAndPartSplitEnabling();
}
		
function frmScreen.optPartAndPartYes.onclick()
{
	SetButtonsOnNewQuote();
	//InitialiseTable();
			
	<% /*
		// RF 24/11/99 AQR MC8
		//frmScreen.btnPartAndPartSplit.disabled = false;
	*/ %>
	SetPartAndPartSplitEnabling();
}
		
function frmScreen.btnPartAndPartSplit.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
			
	ArrayArguments[0] = frmScreen.txtAmountRequested.value;
			
	<% /* 
		RF 24/11/99 AQR MC9: Ensure split amounts initialised correctly on entry
		to screen cm160
	*/ %>
	ArrayArguments[1] = m_sInterestOnlyPart;
	ArrayArguments[2] = m_sCapitalOnlyPart;

	ArrayArguments[3] = "0";		<% /* read only flag */%>
	ArrayArguments[4] = m_sCurrency;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "cm160.asp", 
												ArrayArguments, 500, 200);
	if (sReturn != null)
	{						
		m_sInterestOnlyPart = sReturn[1];
		m_sCapitalOnlyPart	= sReturn[2];
				
		if (sReturn[0] == true)
		{
			FlagChange(sReturn[0]);
			SetButtonsOnNewQuote();
			//InitialiseTable();
		}				
	}	
}
		
function frmScreen.btnFurtherFiltering.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
			
	ArrayArguments[0] = m_sApplicationMode;
	ArrayArguments[1] = m_sUserId;
	ArrayArguments[2] = m_sUnitId;
	ArrayArguments[3] = m_sUserType;
	ArrayArguments[4] = m_sFurtherFiltering;
						
	sReturn = scScreenFunctions.DisplayPopup(window, document, "cm120.asp", 
												ArrayArguments, 490, 440);
	if (sReturn != null)
	{
		m_sFurtherFiltering = sReturn[1];
				
		if (sReturn[0] == true)
		{
			FlagChange(sReturn[0]);
			SetButtonsOnNewQuote();
			//InitialiseTable();
		}
	}			
}
		
function InitialiseTable()
{
	<% /* RF 25/11/99 AQR MC4 */%>
	scTable.initialise(m_tblProductTable, 0, "");
}
		
function SetButtonsOnNewQuote()
{
	frmScreen.btnOneOffCosts.disabled	= true;
	frmScreen.btnIncentives.disabled	= true;
	if(scTable.getRowSelected() == -1)
	{
<%		/*If we haven't got a product selected, disable the calculate and details
          buttons, otherwise we want to leave them enabled*/	
%>		frmScreen.btnCalculate.disabled		= true;
		frmScreen.btnDetails.disabled		= true;
	}
			
	<% /*
		// RF 25/11/99 AQR MC24 
		//frmScreen.txtCapitalAndInterestCost.value	= "";
		//frmScreen.txtPartAndPartCost.value			= "";
		//frmScreen.txtInterestOnlyCost.value			= "";
	*/ %>
	InitialiseCalculatedCosts();
						

	//ADP SYS0861 DisableMainButton("Submit");
}
		
function InitialiseCalculatedCosts()
{
	frmScreen.txtCapitalAndInterestCost.value	= "";
	frmScreen.txtPartAndPartCost.value			= "";
	frmScreen.txtInterestOnlyCost.value			= "";
}
		
function SetButtonsOnExistingQuote()
{			
	frmScreen.btnOneOffCosts.disabled			= false;
	frmScreen.btnIncentives.disabled			= false;
	frmScreen.btnStoredIllustrations.disabled	= false;
			
	//ADP: SYS0861 EnableMainButton("Submit");
}
		
function SetPartAndPartSplitEnabling()
{
	if (frmScreen.optPartAndPartNo.checked == true)
	{
		frmScreen.btnPartAndPartSplit.disabled = true;
				
		<% /* RF 24/11/99 AQR MC8 */%>
		m_sInterestOnlyPart = "";
		m_sCapitalOnlyPart = "";
	}
	else
		frmScreen.btnPartAndPartSplit.disabled = false;
}
		
function frmScreen.btnSearch.onclick()
{
	if(checkData())
	{
		scTable.clear();
		InitialiseTable();
		scScrollPlus.Clear();
		m_blnSetButtonsOnNewQuote = true;
		scScrollPlus.Initialise(DoProductSearch,DisplayProducts,m_tblProductTable.rows.length - 1,1);		
	}
}
function checkData()
{
<%	/* until we have an attrib for max/min, check input values */
%>	var bSuccess = true;
	if(parseInt(frmScreen.txtTermMonths.value) < 0 || parseInt(frmScreen.txtTermMonths.value) > 11)
	{
		bSuccess = false;
		alert("Term in Months out of range");
		frmScreen.txtTermMonths.focus();
	}
	return bSuccess;	
}
function DoProductSearch(iStart)
{
	scTable.setRowSelected(-1);
				
	m_ProductXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	m_ProductXML.CreateRequestTag(window, "SEARCH");
						
	m_ProductXML.CreateActiveTag("MORTGAGEPRODUCT");
			
	m_ProductXML.CreateTag("SEARCHCONTEXT",			"Mortgage Calculator");
	m_ProductXML.CreateTag("DISTRIBUTIONCHANNELID",	m_sDistributionChannelId);
	m_ProductXML.CreateTag("TERMINYEARS",			frmScreen.txtTermYears.value);
	m_ProductXML.CreateTag("TERMINMONTHS",			frmScreen.txtTermMonths.value);
	m_ProductXML.CreateTag("AMOUNTREQUESTED",		frmScreen.txtAmountRequested.value);
	m_ProductXML.CreateTag("PURCHASEPRICE",			frmScreen.txtPurchasePrice.value);
	m_ProductXML.CreateTag("LTV",					frmScreen.txtLTV.value);
	m_ProductXML.CreateTag("ORDERBY",				" FIRSTMONTHLYINTERESTRATE ");
	var nTableLength = m_tblProductTable.rows.length - 1;	
	m_ProductXML.CreateTag("RECORDCOUNT", m_tblProductTable.rows.length - 1);
	var nStartRecord = ((nTableLength * iStart) + 1) - nTableLength; 
	m_ProductXML.CreateTag("STARTRECORD", nStartRecord);
				
	var FurtherFilteringXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	FurtherFilteringXML.LoadXML(m_sFurtherFiltering);
			
	m_ProductXML.AddXMLBlock(FurtherFilteringXML.XMLDocument);
			
	FurtherFilteringXML = null;
			
	m_ProductXML.RunASP(document, "FindMortgageProducts.asp");
			
	if (m_sDebug == "True")
		m_ProductXML.WriteXMLToFile("C:\\temp\\mc010FindMortgageProducts.xml");
			
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = m_ProductXML.CheckResponse(sErrorArray);

	if(sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		alert("No suitable products could be found that match the search criteria");
		scTable.clear();
	}
	else if(sResponseArray[0] == true)
	{				
		if(m_ProductXML.SelectTag(null,"MORTGAGEPRODUCTLIST") != null)
		{
			m_iPageNo = iStart;
			
			if (m_blnSetButtonsOnNewQuote) 	
				SetButtonsOnNewQuote();
					
			return m_ProductXML.GetAttribute("TOTAL");
		}
	}	
	sErrorArray = null;	
}

function GetNumberOfProducts()
{
	m_ProductXML.ActiveTag = null;
	m_ProductXML.CreateTagList("MORTGAGEPRODUCT");
			
	return m_ProductXML.ActiveTagList.length;
}

function SelectIllustratedProduct()
{
	var sProductCode		= null;
	var sProductStartDate	= null;
	var iNumberOfProducts	= 0;
	var blnProductFound		= false;
						
	iNumberOfProducts = GetNumberOfProducts();
			
	for (var nProduct = 1; 
		nProduct <= iNumberOfProducts && blnProductFound == false; 
		nProduct++)
	{
		if (m_ProductXML.SelectTagListItem(nProduct-1) == true)
		{
			sProductCode		= m_ProductXML.GetTagText("MORTGAGEPRODUCTCODE");
			sProductStartDate	= m_ProductXML.GetTagText("STARTDATE");
				
			if ((sProductCode == m_sIllustrationProductCode) &&
				(sProductStartDate == m_sIllustrationProductStartDate))
			{
				blnProductFound = true;
						
				<% /* RF 25/11/99 AQR MC4 : select the record */%>
				scTable.setRowSelected(nProduct);
			}
		}
	}
			
	return blnProductFound;
}		

function DisplayProducts(iStart)
{			
	<% /* iStart is zero based */%>
	
	var sProductName			= null;
	var sLenderName				= null;
	var sFlexibleMortgage		= null;
	var blnFlexibleMortgage		= false;
	var sInterestRateEndDate	= null;
	var sRateEndDateOrPeriod	= null;
	var tagPRODUCTLIST			= null;
	var tagPRODUCTDETAILS		= null;
	var iNumberOfProducts		= 0;
	
	scTable.clear();
	
	m_ProductXML.ActiveTag = null;
	tagPRODUCTLIST = m_ProductXML.CreateTagList("MORTGAGEPRODUCT");
	iNumberOfProducts = m_ProductXML.ActiveTagList.length;
			
	for (var nProduct = 0; nProduct < iNumberOfProducts && 
			nProduct < m_iTableLength; nProduct++)
	{
		m_ProductXML.ActiveTagList = tagPRODUCTLIST;
		if (m_ProductXML.SelectTagListItem(nProduct + iStart) == true)
		{
			sProductName = m_ProductXML.GetTagText("PRODUCTNAME");					
			blnFlexibleMortgage = m_ProductXML.GetTagBoolean("FLEXIBLEMORTGAGEPRODUCT");
			if (blnFlexibleMortgage)
				sFlexibleMortgage = "Yes";
			else
				sFlexibleMortgage = "";
										
			tagPRODUCTDETAILS = m_ProductXML.ActiveTag;					
					
			if (m_sMultipleLender == "1")
			{
				<% 
				/* APS 16/02/00 - Changes made to XML format from Xeneration Stored Procedure
				m_ProductXML.SelectTag(tagPRODUCTDETAILS, "MORTGAGELENDER"); */ %>
				sLenderName = m_ProductXML.GetTagText("LENDERSNAME");
			}
			<%  
			/* APS 16/02/00 - Changes made to XML format from Xeneration Stored Procedure
			m_ProductXML.SelectTag(tagPRODUCTDETAILS, "INTERESTRATETYPE");																				
			sInterestRate = GetInterestRate(tagPRODUCTDETAILS); */ %>
			sInterestRate = m_ProductXML.GetTagText("FIRSTMONTHLYINTERESTRATE");
					
			<% 
			/* APS 16/02/00 - Changes made to XML format from Xeneration Stored Procedure
			Interest rate end date or period 
			m_ProductXML.SelectTag(tagPRODUCTDETAILS, "INTERESTRATETYPE");*/ %>
			sInterestRateEndDate = m_ProductXML.GetTagText("INTERESTRATEENDDATE"); 
					
			<% /* PSC 22/02/00 - Check for empty string as well as null */ %>
			if (sInterestRateEndDate != null && sInterestRateEndDate != "")
				sRateEndDateOrPeriod = sInterestRateEndDate;
			else
			{
				sRateEndDateOrPeriod = m_ProductXML.GetTagText("INTERESTRATEPERIOD");
				
				<%
				/* PSC 22/02/00 - If term is -1 then it is for the rest of term 
				therefore set display to n/a */
				/* PSC 28/02/00 - Amend n/a to Whole Term */ %>

				if (sRateEndDateOrPeriod == "-1")
					sRateEndDateOrPeriod = "Whole Term";
				else
					sRateEndDateOrPeriod = sRateEndDateOrPeriod + " Months"
			}
					
			ShowRow(nProduct+1, sLenderName, sProductName, sInterestRate, sRateEndDateOrPeriod, sFlexibleMortgage);
		}				
	}
			
	return iNumberOfProducts;
}
		
<% 
/* APS 16/02/00 - Function commented out due to changes made to XML 
format from Xeneration Stored Procedure */
function GetInterestRate(tagPRODUCTDETAILS)
{			
	var sInterestRateType	= null;
	var dInterestRate		= 0.00;
	var dBaseInterestRate	= 0.00;
			
	sInterestRateType	= m_ProductXML.GetTagText("RATETYPE");
	dInterestRate		= m_ProductXML.GetTagFloat("RATE");
			
	if (sInterestRateType != "F")
	{
		var dCeilingRate = 0.00;
		var dFlooredRate = 0.00;
				
		dCeilingRate	= m_ProductXML.GetTagFloat("CEILINGRATE");
		dFlooredRate	= m_ProductXML.GetTagFloat("FLOOREDRATE");
				
		m_ProductXML.SelectTag(tagPRODUCTDETAILS, "BASERATEBAND");
		dBaseInterestRate = m_ProductXML.GetTagFloat("RATE");
					
		switch (sInterestRateType)
		{
			case "B":
				dInterestRate = dBaseInterestRate;
				break;
						
			case "D":
				dInterestRate = dBaseInterestRate - dInterestRate;
				break;
						
			case "C":
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
	dInterestRate = top.frames[1].document.all.scMathFunctions	.RoundValue(dInterestRate, 2);		

	return dInterestRate;
}
%>
		
function spnAvailableProducts.onclick()
{
	<% /* RF 25/11/99 AQR MC4 */ %>
	if (scTable.getRowSelected() != -1)
	{
		InitialiseCalculatedCosts();
		
		frmScreen.btnCalculate.disabled = false;
		frmScreen.btnDetails.disabled = false;
		frmScreen.btnOneOffCosts.disabled = true;
		frmScreen.btnIncentives.disabled = true;
	}
	else
	{
		frmScreen.btnCalculate.disabled = true;
		frmScreen.btnDetails.disabled = true;
	}
}
		
function GetProductFromTable()
{
	var blnReturn = false;
			
	var iOffset = scTable.getRowSelected() - 1;					
			
	m_ProductXML.ActiveTag = null;
			
	if (m_ProductXML.CreateTagList("MORTGAGEPRODUCT") != null)
	{
		if (m_ProductXML.SelectTagListItem(iOffset) == true)
			blnReturn = true;
	}
			
	return blnReturn;
}
		
function GetProductCode()
{
	var sProductCode = "";
			
	if (GetProductFromTable())
		sProductCode = m_ProductXML.GetTagText("MORTGAGEPRODUCTCODE");
			
	return sProductCode;
}

function GetProductStartDate()
{
	var sStartDate = "";
			
	if (GetProductFromTable() == true)
		sStartDate = m_ProductXML.GetTagText("STARTDATE");
			
	return sStartDate;
}
		
function GetProductTextDetails()
{
	var sProductTextDetails = "";
			
	if (GetProductFromTable())
		sProductTextDetails = m_ProductXML.GetTagText("PRODUCTTEXTDETAILS");
			
	return sProductTextDetails;
}
		
function GetSelectedProductXML()
{
	var sSelectedProductXML = "";
			
	if (GetProductFromTable())
		sSelectedProductXML = m_ProductXML.ActiveTag.xml;
			
	return sSelectedProductXML;
}
		
function ShowRow(iRow, sLenderName, sProductName, sRate, sPeriod, sFlexible)
{
	var iCellIndex = 0;
						
	if (m_sMultipleLender == "1")
		scScreenFunctions.SizeTextToField(m_tblProductTable.rows(iRow).cells(iCellIndex++),
										sLenderName);
												
	scScreenFunctions.SizeTextToField(m_tblProductTable.rows(iRow).cells(iCellIndex++),
										sProductName);
	scScreenFunctions.SizeTextToField(m_tblProductTable.rows(iRow).cells(iCellIndex++),
										sRate);
	scScreenFunctions.SizeTextToField(m_tblProductTable.rows(iRow).cells(iCellIndex++),
										sPeriod);
	scScreenFunctions.SizeTextToField(m_tblProductTable.rows(iRow).cells(iCellIndex++),
										sFlexible);												
}
		
function frmScreen.btnCalculate.onclick()
{
	<% /* RF 25/11/99 AQR MC24: Clear out old calc values in case calc fails */%>
	InitialiseCalculatedCosts();
									
	var iPartAndPart = scScreenFunctions.GetRadioGroupValue(frmScreen, "PartAndPart")
			
	if ((iPartAndPart == frmScreen.optPartAndPartYes.value) && 
		((m_sCapitalOnlyPart == "") ||
	 	 (m_sInterestOnlyPart == "")))
		<% /* SYS4139 - Display correct error message */ %>
		alert("Loan split must be entered for part and part.");			
	else
	{
		var sProductCode	= GetProductCode();
		var sStartDate		= GetProductStartDate();
		var CalculateXML	= new top.frames[1].document.all.scXMLFunctions.XMLObject();			
			
		CalculateXML.CreateRequestTag(window);

		CalculateXML.CreateActiveTag("CALCS");
			
		CalculateXML.CreateTag("MORTGAGEPRODUCTCODE",	sProductCode);
		CalculateXML.CreateTag("STARTDATE",				sStartDate);
			
		CalculateXML.CreateTag("AMOUNTREQUESTED",	frmScreen.txtAmountRequested.value);
		CalculateXML.CreateTag("PURCHASEPRICE",		frmScreen.txtPurchasePrice.value);
		CalculateXML.CreateTag("TERMINYEARS",		frmScreen.txtTermYears.value);
		CalculateXML.CreateTag("TERMINMONTHS",		frmScreen.txtTermMonths.value);
		CalculateXML.CreateTag("LTV",				frmScreen.txtLTV.value);

		CalculateXML.CreateTag("INTERESTONLYELEMENT",		m_sInterestOnlyPart);
		CalculateXML.CreateTag("CAPITALANDINTERESTELEMENT",	m_sCapitalOnlyPart);
			
		// 		CalculateXML.RunASP(document, "CalculateCosts.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					CalculateXML.RunASP(document, "CalculateCosts.asp");
				break;
			default: // Error
				CalculateXML.SetErrorResponse();
			}

			
		var ErrorTypes = new Array("ONEOFFCOSTS");
		var ErrorReturn = CalculateXML.CheckResponse(ErrorTypes);
				
		<% /* continue on if we get error no. 167 */%>
		if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
		{									
			var sCost = "";
			
			if (CalculateXML.SelectTag(null, "ONEOFFCOSTLIST") != null)
				m_sOneOffCosts = CalculateXML.ActiveTag.xml;
			
			if (CalculateXML.SelectTag(null, "REGULARCOSTS") != null)
			{
				sCost = CalculateXML.GetTagText("MONTHLYINTERESTONLYCOST");

				<%/* SG 06/06/02 SYS4800 */%>
				<%/* sCost = scMath.RoundValue(sCost,2); */%>
				sCost = top.frames[1].document.all.scMathFunctions.RoundValue(sCost,2);

				CalculateXML.SetTagText("MONTHLYINTERESTONLYCOST", sCost);
				frmScreen.txtInterestOnlyCost.value = sCost;
					
				sCost = CalculateXML.GetTagText("MONTHLYCAPITALANDINTERESTCOST");

				<%/* SG 06/06/02 SYS4800 */%>
				<%/* sCost = scMath.RoundValue(sCost,2); */%>
				sCost = top.frames[1].document.all.scMathFunctions.RoundValue(sCost,2);

				CalculateXML.SetTagText("MONTHLYCAPITALANDINTERESTCOST", sCost);
				frmScreen.txtCapitalAndInterestCost.value = sCost;
					
				sCost = CalculateXML.GetTagText("MONTHLYPARTANDPARTCOST");

				<%/* SG 06/06/02 SYS4800 */%>
				<%/* sCost = scMath.RoundValue(sCost,2); */%>
				sCost = top.frames[1].document.all.scMathFunctions.RoundValue(sCost,2);
				
				CalculateXML.SetTagText("MONTHLYPARTANDPARTCOST", sCost);					
				frmScreen.txtPartAndPartCost.value = sCost;
			}																	
			SetButtonsOnExistingQuote();
												
			CalculateXML.ActiveTag = null;
			CalculateTotalIncentives(CalculateXML);
			CreateIllustration(CalculateXML);
		}				
	}									
	CalculateXML = null;			
}
function CalculateTotalIncentives(CalculateXML)
{
<%	/* set the total of all inclusive incentives and the first exclusive incentive */
%>	var iTotal = 0;
	if(CalculateXML.SelectTag(null, "INCLUSIVEINCENTIVELIST") != null)
	{
		CalculateXML.CreateTagList("INCENTIVE");
		var iNumberOfIncentives = CalculateXML.ActiveTagList.length;
		for (var nIncentive = 0; nIncentive < iNumberOfIncentives ; nIncentive++)
		{
			CalculateXML.SelectTagListItem(nIncentive);
			iTotal += parseFloat(CalculateXML.GetTagText("AMOUNT"));
		}
	}
	if(CalculateXML.SelectTag(null, "EXCLUSIVEINCENTIVELIST") != null)
	{
		CalculateXML.CreateTagList("INCENTIVE");
		var iNumberOfIncentives = CalculateXML.ActiveTagList.length;
		if(iNumberOfIncentives > 0)
		{
			CalculateXML.SelectTagListItem(0);
			iTotal += parseFloat(CalculateXML.GetTagText("AMOUNT"));
		}
	}
	m_sTotalIncentives = iTotal.toString();
}
function CreateIllustration(CalculateXML)
{
	var lllustrationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	lllustrationXML.CreateActiveTag("ILLUSTRATION");
						
	lllustrationXML.CreateTag("PURCHASEPRICE",		frmScreen.txtPurchasePrice.value);
	lllustrationXML.CreateTag("AMOUNTREQUESTED",	frmScreen.txtAmountRequested.value);			
	lllustrationXML.CreateTag("TERMINYEARS",		frmScreen.txtTermYears.value);
	lllustrationXML.CreateTag("TERMINMONTHS",		frmScreen.txtTermMonths.value);
	lllustrationXML.CreateTag("LTV",				frmScreen.txtLTV.value);
	lllustrationXML.CreateTag("PARTANDPARTREQUIRED",scScreenFunctions.GetRadioGroupValue(frmScreen, "PartAndPart"));
			
	lllustrationXML.CreateTag("INTERESTONLYELEMENT",		m_sInterestOnlyPart);
	lllustrationXML.CreateTag("CAPITALANDINTERESTELEMENT",	m_sCapitalOnlyPart);
			
	lllustrationXML.CreateTag("TOTALINCENTIVES", m_sTotalIncentives);
	lllustrationXML.CreateTag("PRODUCTPAGE", m_iPageNo);
						
	var FurtherFilteringXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	FurtherFilteringXML.LoadXML(m_sFurtherFiltering);						
	lllustrationXML.AddXMLBlock(FurtherFilteringXML.XMLDocument);			
	FurtherFilteringXML = null;
						
	var SelectedProductXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	SelectedProductXML.LoadXML(GetSelectedProductXML());
	lllustrationXML.AddXMLBlock(SelectedProductXML.XMLDocument);
	SelectedProductXML = null;			
			
	CalculateXML.SelectTag(null,"RESPONSE");
	lllustrationXML.AddXMLBlock(CalculateXML.ActiveTag);			
						
	m_IllustrationListXML.SelectTag(null, "ILLUSTRATIONLIST");
	m_IllustrationListXML.AddXMLBlock(lllustrationXML.XMLDocument);
									
	m_iNumberOfIllustrations++;
	m_iIllustrationNumber = m_iNumberOfIllustrations;						
			
	lllustrationXML = null;
}
		
function frmScreen.btnIncentives.onclick()
{
	var sReturn = null;			
	var ArrayArguments = new Array();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	ArrayArguments[0] = m_sApplicationMode;
	ArrayArguments[1] = "0";  // no read only status for Mortgage Calculator
	ArrayArguments[2] = XML.CreateRequestAttributeArray(window);
	ArrayArguments[3] = GetProductCode();
	ArrayArguments[4] = GetProductStartDate();
	ArrayArguments[5] = frmScreen.txtAmountRequested.value;
	ArrayArguments[6] = m_sCurrency;
						
	sReturn = scScreenFunctions.DisplayPopup(window, document, "cm150.asp", 
												ArrayArguments, 500, 410);
	if (sReturn != null)
	{
		m_sTotalIncentives = sReturn[1];

		m_IllustrationListXML.ActiveTag = null;

		if (m_IllustrationListXML.CreateTagList("ILLUSTRATION") != null)
		{
			if (m_IllustrationListXML.SelectTagListItem(m_iIllustrationNumber-1) == true)
				m_IllustrationListXML.SetTagText("TOTALINCENTIVES", m_sTotalIncentives);
		}
	}
	XML = null;
}
		
function frmScreen.btnOneOffCosts.onclick()
{
	var sReturn = null;			
	var ArrayArguments = new Array();
			
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
		
function frmScreen.btnDetails.onclick()
{
	var sReturn = null;			
	var ArrayArguments = new Array();
			
	ArrayArguments[0] = GetProductTextDetails();
						
	sReturn = scScreenFunctions.DisplayPopup(window, document, "cm115.asp", 
												ArrayArguments, 350, 290);
}		
		
function CalcLTV()
{
	if( (frmScreen.txtPurchasePrice.value != "" && frmScreen.txtPurchasePrice.value != "0") &&
	    (frmScreen.txtAmountRequested.value != "" && frmScreen.txtAmountRequested.value != "0") )
	{
		var LTVXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				
		LTVXML.CreateRequestTag(window);
		LTVXML.CreateActiveTag("LTV");
		LTVXML.CreateTag("AMOUNTREQUESTED",	frmScreen.txtAmountRequested.value);
		LTVXML.CreateTag("PURCHASEPRICE",	frmScreen.txtPurchasePrice.value);
		LTVXML.RunASP(document, "CalculateLTV.asp");
		if (LTVXML.IsResponseOK() == true)
		{
			<%/* SG 06/06/02 SYS4800 */%>
			<%/* frmScreen.txtLTV.value = scMath.RoundValue(LTVXML.GetTagText("LTV"),2); */%>
			frmScreen.txtLTV.value = top.frames[1].document.all.scMathFunctions.RoundValue(LTVXML.GetTagText("LTV"),2);		
		}
		LTVXML = null;
	}
}

function CleanUp()
{
	m_ProductXML			= null;
	m_IllustrationListXML	= null;
}
		
function btnSubmit.onclick()
{			
	<% 
	/* AP SYS0861 
	scScreenFunctions.SetContextParameter(window,"idAmountRequested", frmScreen.txtAmountRequested.value);
	scScreenFunctions.SetContextParameter(window,"idPurchasePrice", frmScreen.txtPurchasePrice.value);
	scScreenFunctions.SetContextParameter(window,"idMortgageProductId", GetProductCode());
	scScreenFunctions.SetContextParameter(window,"MortgageProductStartDate",GetProductStartDate());
	scScreenFunctions.SetContextParameter(window,"idTermYears",	 frmScreen.txtTermYears.value);
	scScreenFunctions.SetContextParameter(window,"idTermMonths", frmScreen.txtTermMonths.value);
	CleanUp();
	frmCustomerSearch.submit(); */
	%>
	
	scScreenFunctions.SetContextParameter(window,"idXML","");
	CleanUp();
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmWelcomeMenu.submit();
}
		
function frmScreen.btnStoredIllustrations.onclick()
{						
	scScreenFunctions.SetContextParameter(window,"idXML", m_IllustrationListXML.XMLDocument.xml);
	scScreenFunctions.SetContextParameter(window,"idMCIllustrationNumber", m_iIllustrationNumber);
	CleanUp();
	frmStoredIllustrations.submit();
}
		
-->
</script>
</body>
</html>


