<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
Workfile:      CM100.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Loan Composition
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		14/02/2000	Created
JLD		28/03/00	SYS0462 - screen alterations
					SYS0502 - reserve button error. Also added view subquote stuff
AY		04/04/00	New top menu/scScreenFunctions change
JLD		10/04/00	SYS0570 - Change to confirm message text
JLD		10/04/00	SYS0566 - replace Product Number with Product Name in listbox
AY		13/04/00	SYS0328 - Dynamic currency display
JLD		20/04/00	SYS0639 - make sure edit btn is enabled after NewSubquote 
							  if listbox entry is selected
JLD		28/04/00	Addition of CM130 functionality
PSC     02/05/00    SYS0593 - Change Reset to call QuickQuote/ApplicationQuote
                              ResetMortgageSubQuote rather than MortgageSubQuote
JLD		05/05/00	SYS0642	- If we already have an amount requested, recalc the LTV on entry
							  in case the purchase price has altered.
JLD		12/05/00	SYS0666 - Cope with a product with no base rate set
JLD		12/05/00	SYS0715 - Call QQ and AQ specific UpdateMortgageSubquote methods
PSC		23/05/00	SYS0467 - Add product code into list box
APS		31/05/00	SYS0328 - Added Dynamic currency display
MS		07/06/00	SYS0823 - Disable/enable calculate button when populating list box
PSC		30/06/00	SYS1012 - Remove LTV button
BG		19/10/00	SYS1074 - Added CustomerList function to only get the first two APPLICANTS (not guarantors)
BG		05/12/00	SYS1653 - Commented out offending code in btnDelete.onclick and added initialise()
CL		05/03/01	SYS1920 Read only functionality added
IK		22/03/01	SYS2145 Add idNoCompleteCheck
BG		11/04/01	SYS2096 Added Printing functionality 
JLD		08/10/01	SYS2736 Don't download omPC again
JLD		4/12/01		SYS2806 Completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
BG		28/06/02	SYS4767 MSMS/Core integration
STB		28/02/02	SYS4144 Display APR column
JLD		26/03/02	MSMS0029	Addition of manual rate adjustment
JLD		24/04/02	MSMS0034    add combos for type of app and type of buyer.
JLD		14/05/02	MSMS0054	Display Resolved Rate to 2dp always
JLD		14/05/02	MSMS0082	Fixes for CR01
JLD		20/05/02	MSMS0092	clear Type of Buyer if TxEquity.
BG		28/06/02	SYS4767 MSMS/Core integration - END
SG		06/06/02	SYS4821 Error from SYS4767
JLD		11/06/02	SYS4846 corrected commenting error  
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description 
DPF		10/07/02	CMWP3 - have amended the height of pop up window for CM110.asp
					due to increase in number of fields.
PSC		10/07/02	BMIDS00062	Add Cost at Final Rate onto list box
MV		18/07/02	BMIDS00179	Core Upgrade Rollback - Modified GUI HTML , PopulateScreen() ;InitialiseCombos();
DPF		10/07/02	CMWP3 - have added two new arugements to the array passed across to CM130.asp
					for capital & interest and interest only amounts to be used in part-part calculations
ASu		11/09/02	BMIDS00431 - Add cancel button to asp - always route to CM010.
DPF		01/10/02	CPWP1 - have added field to record the drawdown amount for flexible mortgages
					Have also added a SaveMortgageSubQuoteDetails function (called on btnSubmit.onClick)
DPF		08/10/02	CPWP1 - have added field to record the monthly cost of the mortgage (less drawdown)
TW      09/10/02	Modified to incorporate client validation - SYS5115
DPF		10/10/02	Manually amended some of the ScreenRule changes.
BG		10/10/02	BMIDS00612  Add new print indicator to ensure archiving is done.
DPF		14/10/02	CPWP1 - Have added an alert box on Initialize event if manual incentive
					amount has been entered previously on CM130.asp.
MV		01/11/2002	BMIDS00723	- Modified the HTML Text Positions
MDC		04/11/2002	BMIDS00654 - ICWP BM088 Income Calculations
MDC		08/11/2002	BMIDS00878 - Disable/Enable buttons post calc, pre-Income Calc.
DPF		08/11/2002  BMIDS00808 - Add Drawdown amount into the parameters passed to CM130.asp
PSC		15/11/2002	BMIDS00964 - Disable drawdown if part and part
MDC		22/11/2002	BMIDS01070 - Repayment Type abbreviation showing as ??
GHun	28/11/2002	BMIDS01093 - fixed word spacing and alignment
MDC		04/12/2002	BM0135 - Screen Rules fixes.
MV		29/01/2003	BM0300	- Modified PopulateScreen();
MV		10/02/2003	BM0337 - Amended btnCalculate.Onclick()
GHun	13/03/2003	BM0457 - Include guarantors in allowable income calculation
DB		20/03/2003	BM0443 - When user clicks cancel stop completeness check being run again in cm010.
HMA     15/05/2003  BM0133 - Pass context parameter Application Type to CM130
KRW     23/02/2004  BM0562 - Resets Amt Requested to previous valueif user cancels on message displayed when 
						     Amount has been changed , Removed ref to Type of Mortgage and Type of Buyer in message
MC		19/04/2004	bmids517	OneoffCost dialog (CM130) and Manual Adjustment dialog (CM101) height incr. by 15px
DC      10/06/2004  BMIDS767 Change size of popup window in call to CM130
JD		19/07/04	BMIDS763 ApplicationDate required for call to CalculateMortgageCosts.
JD		26/07/04	BMIDS816 ApplicationDate is required in CM110.
SR		31/08/2004  BMIDS815 
sR		03/09/2004  BMIDS815 
GHun	07/09/2004	BMIDS815
GHun	13/09/2004	BMIDS815 Changed InitialRate to InterestRateStep
JD		15/09/2004	BMIDS876 set window status on account refresh
GHun	16/09/2004	BMIDS876 Set window status on calculation
HMA     23/09/2004  BMIDS878 Use Initial Interest Rate in call to CM101
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description 
MV		19/09/2002	MAR35		Amended Reset()
MF		15/09/2005	MAR30		IncomeCalcs modified to send ActivityID into calculation
SD		10/10/2005	MAR86		If Requested amount is greater than counter offer, display a warning.
GHun	12/10/2005	MAR46		Made CM110 window larger
DRC     09/03/2006  MAR1380     Pass extra paramters to CM110 for Critical Data
JD		09/03/2006	MAR1061		display msq.purchaseprice
JD		22/03/2006	MAR1061		Fixing errors
PK		04/05/2006	MAR1707		Change LTV from 2dp to 1dp.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
SW		20/06/2006	EP771		Increased the size of the CM110 popup
SAB		21/06/2006	EP837		Override MAR1707 Format LTV to 2 significant numbers
AW		23/06/2006	EP852		Added fact find number to ReserveProduct call
PE		30/06/2006	EP974		MAR1888 - Mars Merge/Changed SaveMortgageSubQuoteDetails to only save updated values
AShaw	07/11/2006	EP2_8		Add extra param when loading CM110 (Add/Edit).
								Added default value for new param.
AShaw	16/11/2006	EP2_55		New code for PSW / TOE apps.
AShaw	28/12/2006	EP2_56		New method for Adding components.
******	NB. DRAWDOWN needs uncommenting when col in table. ******
AShaw	26/01/2007	EP2_771		DisableAmountRequested call moved.
AShaw	31/01/2007	EP2_1023	The ReturnIndicator method MUST return a non empty 
								(nor a single space) string.
PSC		02/02/2007	EP2_1113	Correct screen functionality for further borrowing, switches etc	
PSC		07/02/2007	EP2_1271	Set up start date on loan component when generating from account							
PSC		12/02/2007	EP2_1314	Fixes for ported loans
DS		30/03/2007	EP2_1943	Disabled 'Add' and 'New Subquote' buttons for application types TOE,PSW or NP 
AShaw	02/04/2007	EP2_2168	Add new method to prevent Drawdown > Requested amount.
DS		10/04/2007	EP2_1943	On calculate button click disabled 'Add' and 'New Subquote' buttons for application types TOE,PSW or NP 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<span style="TOP: 271px; LEFT: 310px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scScrollTable" style="DISPLAY: none; HEIGHT: 24px; WIDTH: 304px; VISIBILITY: hidden" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToCM010" method="post" action="cm010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM075" method="post" action="cm075.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange">

<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 342px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Purchase Price
		<span style="TOP: 0px; LEFT: 90px; POSITION: ABSOLUTE">
			<label style="LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtPurchasePrice" maxlength="10" style="TOP: -3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 10px; LEFT: 170px; POSITION: ABSOLUTE" class="msgLabel">
		Amount Requested
		<span style="TOP: 0px; LEFT: 110px; POSITION: ABSOLUTE">
			<label style="LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtAmtRequested" maxlength="10" style="TOP: -3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 10px; LEFT: 360px; POSITION: ABSOLUTE" class="msgLabel">
		LTV 
		<span style="TOP: -3px; LEFT: 25px; POSITION: ABSOLUTE">
			<input id="txtLTVPercent" maxlength="10" style="WIDTH: 38px; POSITION: ABSOLUTE" class="msgTxt">
			<span style="LEFT: 41px; POSITION: absolute; TOP: 3px" class="msgLabel">
				%
			</span>
		</span>
	</span>	
<% /* DPF 09/10/2002 - CPWP1 - new field */ %>
	<span style="TOP: 35px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Draw Down
		<span style="TOP: 0px; LEFT: 90px; POSITION: ABSOLUTE">
			<label style="LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtDrawDown" maxlength="10" style="TOP: -3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>	
	<span style="TOP:35px; LEFT: 170px; POSITION: ABSOLUTE" class="msgLabel">
		Total Loan Amount
		<span style="TOP: 0px; LEFT: 110px; POSITION: ABSOLUTE">
			<label style="LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtTotalLoanAmt" maxlength="10" style="TOP: -3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>	
	<div id="spnTable" style="TOP: 60px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblTable" width="596" border="0" cellspacing="0"cellpadding="0" class="msgTable">
			<tr id="rowTitles">
				<td width="5%" class="TableHead">Ind</td>	
				<td width="10%" class="TableHead">Product Number</td>	
				<td width="8%" class="TableHead">Product Name</td>	
				<td width="8%" class="TableHead">Initial Type</td>		
				<td width="10%" class="TableHead">Interest Rate&nbsp;Step</td>	
				<td width="10%" class="TableHead">Resolved Rate</td>
				<td width="10%" class="TableHead">Loan Amt</td>  
				<td width="8%" class="TableHead">Repay Type</td>  
				<td width="10%" class="TableHead">Monthly Cost</td>
				<!--PSC		10/07/02	BMIDS00062 - Start -->
				<td width="8%" class="TableHead">APR</td>
				<td width="10%" class="TableHead">Cost at Final Rate</td></tr>						
			<tr id="row01">		<td width="5%" class="TableTopLeft">&nbsp;</td>	<td width="10%" class="TableTopCenter">&nbsp;</td>	<td width="8%" class="TableTopCenter">&nbsp;</td>	<td width="8%" class="TableTopCenter">&nbsp;</td> <td width="10%" class="TableTopCenter">&nbsp;</td>		<td width="10%" class="TableTopCenter">&nbsp;</td>		<td width="10%" class="TableTopCenter">&nbsp;</td>  <td width="10%" class="TableTopCenter">&nbsp;</td>  <td width="10%" class="TableTopCenter">&nbsp;</td>		<td width="8%" class="TableTopCenter">&nbsp;</td>		<td width="10%" class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="5%" class="TableLeft">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>	<td width="10%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>     <td width="10%" class="TableCenter">&nbsp;</td>		<td width="8%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="5%" class="TableBottomLeft">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="8%" class="TableBottomCenter">&nbsp;</td>	<td width="8%" class="TableBottomCenter">&nbsp;</td> <td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>   <td width="10%" class="TableBottomCenter">&nbsp;</td>   <td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="8%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomRight">&nbsp;</td></tr>
			<!--PSC		10/07/02	BMIDS00062 - End-->
		</table>
	</div>
	<span style="TOP: 264px; LEFT: 4px; POSITION: ABSOLUTE">
		<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	<span style="TOP: 264px; LEFT: 70px; POSITION: ABSOLUTE">
		<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	<span style="TOP: 264px; LEFT: 140px; POSITION: ABSOLUTE">
		<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	<span style="TOP: 264px; LEFT: 210px; POSITION: ABSOLUTE">
		<input id="btnReserve" value="Reserve" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	<span style="TOP: 267px; LEFT: 334px; POSITION: ABSOLUTE" class="msgLabel">
		Total Monthly Cost
		<span style="TOP: 0px; LEFT: 106px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtTotalMnthCost" maxlength="10" style="TOP:-3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
<% /* DPF 09/10/2002 - CPWP1 - new field */ %>
	<span style="TOP: 292px; LEFT: 275px; POSITION: ABSOLUTE" class="msgLabel">
		Monthly Cost Less Drawdown
		<span style="TOP: 0px; LEFT: 165px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtMnthCostLessDD" maxlength="10" style="TOP:-3px; WIDTH: 60px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 264px; LEFT: 506px; POSITION: ABSOLUTE">
		<input id="btnCalculate" value="Calculate" type="button" style="WIDTH: 90px" class="msgButton">
	</span>	
	<span style="TOP: 316px; LEFT: 4px; POSITION: ABSOLUTE">
		<input id="btnViewSubquotes" value="View Subquotes" type="button" style="WIDTH: 100px" class="msgButton" NAME="btnViewSubquotes">
	</span>
	<span style="TOP: 316px; LEFT: 112px; POSITION: ABSOLUTE">
		<input id="btnGenerateQuote" value="Generate Quote from Mtg Acc" type="button" style="WIDTH: 160px" class="msgButton" NAME="btnGenerateQuote">
	</span>
	<span style="TOP: 316px; LEFT: 279px; POSITION: ABSOLUTE">
		<input id="btnNewSubquote" value="New Subquote" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	<span style="TOP: 316px; LEFT: 387px; POSITION: ABSOLUTE">
		<input id="btnOneOffCosts" value="One Off Costs" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	<span style="TOP: 316px; LEFT: 496px; POSITION: ABSOLUTE; DISPLAY: none">
		<input id="btnPrint" value="Print" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	<span style="TOP: 290px; LEFT: 4px; POSITION: ABSOLUTE">
		<input id="btnManAdjust" value="Interest Rate Adjustment" type="button" style="WIDTH: 127px" class="msgButton">
	</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 413px; WIDTH: 612px">
<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!--  #include FILE="attribs/cm100attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript" type="text/javascript">
<!--
var m_sApplicationMode = "";
var m_sReadOnly = "";
var m_sMortgageSubquoteNumber = "";
var m_sLifeSubquoteNumber = "";
var m_sApplicationNumber = "";
var m_sApplicationFFNumber = "";
var m_sCurrency = "";
var subQuoteXML = null;
var CM100ParamsXML = null;   // EP2_8 - New Params for Product search on CM110.

<% /* BMIDS00654 MDC 04/11/2002 */ %>
var m_sAmountRequested = "";
<% /* BMIDS00654 MDC 04/11/2002 - End */ %>

var m_iTableLength = 10;
var m_iNumOfLoanComponents = 0;
var m_sAmtRequested = "";
var m_sPurchasePrice = ""; //MAR1061
var m_sDrawDown = ""; // DPF 09/10/02 - (CPWP1) Variable to hold drawdown amount
var m_iManualIncentive = 0;  // DPF 14/10/02 - (CPWP1) Variable to hold manual incentive amount
var m_PopWinInd = "No";  // DPF 14/10/02 - (CPWP1) Variable to hold marker for initialisation
var scScreenFunctions;
var m_blnReadOnly = false;
var m_iINTERESTONLYAMOUNT = 0;
var m_iCAPITALINTERESTAMOUNT = 0;
<% /* PSC 15/11/2002 BMIDS00964 - Start */ %>
var m_XMLRepay = null;
var m_bPPPresent = false;
<% /* PSC 15/11/2002 BMIDS00964 - End */ %>	
<% /* BM0135 MDC 04/12/2002 */ %>
var m_bScreenEntry = false;
<% /* BM0135 MDC 04/12/2002 - End */ %>
var m_sLTVChanged = "";

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;

<%/*MAR86 Start*/%>
var clientVar_CounterOffer="";
<%/*MAR86 End*/%>

<% /* PSC 12/02/2007 EP2_1314 - Start */ %>
var m_blnIsPort = false;
<% /* PSC 12/02/2007 EP2_1314 - End */ %>

function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();

	var sButtonList = new Array("Submit", "Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Loan Composition","CM100",scScreenFunctions);
	m_sCurrency = scScreenFunctions.SetCurrency(window,frmScreen);
	RetrieveContextData();
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
	
	m_bScreenEntry = true;	<% /* BM0135 MDC 04/12/2002 */ %>

	Initialise();	

	m_bScreenEntry = false;	<% /* BM0135 MDC 04/12/2002 */ %>
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CM100");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();

	<%/*MAR86 Start*/%>
	GetCounterOffer();
	if (CheckCounterOffer()==2)
		return;
	<%/*MAR86 End*/%>

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
	m_sApplicationMode = scScreenFunctions.GetContextParameter(window,"idApplicationMode","Quick Quote");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFFNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sMortgageSubquoteNumber = scScreenFunctions.GetContextParameter(window,"idMortgageSubQuoteNumber",null);
	m_sLifeSubquoteNumber = scScreenFunctions.GetContextParameter(window,"idLifeSubquoteNumber",null);	
	<% /* BMIDS00654 MDC 04/11/2002 */ %>
	m_sAmountRequested = scScreenFunctions.GetContextParameter(window,"idAmountRequested","0");
	<% /* BMIDS00654 MDC 04/11/2002 - End */ %>
}

function Initialise()
{
	var bSuccess = true;
	
	<% /* PSC 12/02/2007 EP2_1314 - Start */ %>
	frmScreen.btnAdd.disabled = false;
	
	var AppTypeXML = new top.frames[1].document.all.scXMLFunctions.XMLObject(); 
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	m_blnIsPort = AppTypeXML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP'));
	<% /* PSC 12/02/2007 EP2_1314 - End */ %>

	<% /* PSC 15/11/2002 BMIDS00964 - Start */ %>
	var sGroups = new Array("RepaymentType");
	var XMLTemp = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLTemp.GetComboLists(document,sGroups);
	m_XMLRepay = XMLTemp.GetComboListXML("RepaymentType")
	<% /* PSC 15/11/2002 BMIDS00964 - End */ %>
	
	// EP2_55 (EP2_771 Moved to here) - Alter logic for Amount requested enabling.
	<% /* PSC 12/02/2007 EP2_1314 */ %>
	if (DisableAmountRequested() == true)
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAmtRequested");
	else
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtAmtRequested");
	
	if(m_sReadOnly != "1" && (m_sMortgageSubquoteNumber == "" || m_sMortgageSubquoteNumber == "0"))
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window,"CREATE");
		if(m_sApplicationMode == "Quick Quote")
		{
			XML.CreateActiveTag("QUOTATION");
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
			XML.CreateTag("QUOTATIONTYPE", "1");
			XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idQuotationNumber","1"));
			<% /* BM0135 MDC 04/12/2002 */ %>
			if(m_bScreenEntry)
				XML.RunASP(document,"QQCreateFirstMortgageLifeSubQuote.asp");
			else
			{
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
				{
					case 1: // Warning
					case 0: // OK
						XML.RunASP(document,"QQCreateFirstMortgageLifeSubQuote.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
				}
			}
			<% /* BM0135 MDC 04/12/2002 - End */ %>
		}
		else 
		{
			XML.CreateActiveTag("LIFESUBQUOTES");
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
			XML.CreateTag("QUOTATIONTYPE", "2");
			XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idQuotationNumber","1"));
			<% /* BM0135 MDC 04/12/2002 */ %>
			if(m_bScreenEntry)
				XML.RunASP(document, "AQCreateFirstMortgageLifeSubquote.asp");
			else
			{
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
				{
					case 1: // Warning
					case 0: // OK
						XML.RunASP(document, "AQCreateFirstMortgageLifeSubquote.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
				}
			}
			<% /* BM0135 MDC 04/12/2002 - End */ %>
		}
		if(XML.IsResponseOK())
		{
			XML.SelectTag(null, "RESPONSE");
			m_sMortgageSubquoteNumber = XML.GetTagText("MORTGAGESUBQUOTENUMBER");
			m_sLifeSubquoteNumber = XML.GetTagText("LIFESUBQUOTENUMBER");
			scScreenFunctions.SetContextParameter(window,"idMortgageSubQuoteNumber",m_sMortgageSubquoteNumber);
			scScreenFunctions.SetContextParameter(window,"idLifeSubquoteNumber",m_sLifeSubquoteNumber);
		}
		else
		{
			bSuccess = false;
			alert("Failed to create first subquote");
			scScreenFunctions.SetCollectionToReadOnly(divBackground);
			DisableButtons();
		}
	}
	if(bSuccess == true)
	{
		subQuoteXML = null;
		subQuoteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		subQuoteXML.CreateRequestTag(window,null);
		subQuoteXML.CreateActiveTag("LOANCOMPOSITION");
		subQuoteXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		subQuoteXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		subQuoteXML.CreateTag("APPLICATIONDATE", scScreenFunctions.GetContextParameter(window,"idApplicationDate","")); //JD BMIDS763
		subQuoteXML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
		if(m_sApplicationMode == "Quick Quote")
		{
			subQuoteXML.CreateTag("QUOTATIONTYPE", "1");
			<% /* BM0135 MDC 04/12/2002 */ %>
			if(m_bScreenEntry)
				subQuoteXML.RunASP(document,"QQGetLoanCompositionDetails.asp");
			else
			{
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
				{
					case 1: // Warning
					case 0: // OK
						subQuoteXML.RunASP(document,"QQGetLoanCompositionDetails.asp");
						break;
					default: // Error
						subQuoteXML.SetErrorResponse();
				}
			}
			<% /* BM0135 MDC 04/12/2002 - End */ %>
		}
		else 
		{
			subQuoteXML.CreateTag("QUOTATIONTYPE", "2");
			<% /* BM0135 MDC 04/12/2002 */ %>
			if(m_bScreenEntry)
				subQuoteXML.RunASP(document, "AQGetLoanCompositionDetails.asp");
			else
			{
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
				{
					case 1: // Warning
					case 0: // OK
						subQuoteXML.RunASP(document, "AQGetLoanCompositionDetails.asp");
						break;
					default: // Error
						subQuoteXML.SetErrorResponse();
				}
			}
			<% /* BM0135 MDC 04/12/2002 - End */ %>
		}
		if(subQuoteXML.IsResponseOK()) 
		{
			PopulateScreen();
			
			//DPF 14/10/2002 - CPWP1 (BM020) - check for manual incentive amount
			m_iManualIncentive = subQuoteXML.GetTagText("MANUALINCENTIVEAMOUNT");
			
			if (m_iManualIncentive != "" && m_iManualIncentive > 0 && m_PopWinInd == 'No')
			{	
				alert('Please review the manual incentive amount in one-off costs');
			}
			
			//CMWP3 - assign part-part values to variables to pass through to CM130.asp			
			m_iINTERESTONLYAMOUNT = subQuoteXML.GetTagText("INTERESTONLYELEMENT");
			m_iCAPITALINTERESTAMOUNT = subQuoteXML.GetTagText("CAPITALANDINTERESTELEMENT");
		
			<% /* PSC 02/02/2007 EP2_1113 */ %> 
			ResetGenerateQuote();
		}
		else
		{
			scScreenFunctions.SetCollectionToReadOnly(divBackground);
			DisableButtons();
		}
	}
}

function DisableButtons()
{
<%	/* When something has gone wrong, disable all buttons */
%>  frmScreen.btnAdd.disabled = true;
	frmScreen.btnCalculate.disabled = true;
	frmScreen.btnDelete.disabled = true;
	frmScreen.btnEdit.disabled = true;
	frmScreen.btnNewSubquote.disabled = true;
	frmScreen.btnOneOffCosts.disabled = true;
	frmScreen.btnPrint.disabled = true;
	frmScreen.btnReserve.disabled = true;
	frmScreen.btnViewSubquotes.disabled = true;
	frmScreen.btnGenerateQuote.disabled = true;  // EP2_55
	

<%/*BG		28/06/02	SYS4767 MSMS/Core integration*/%>
	frmScreen.btnManAdjust.disabled = true;
<%/*BG		28/06/02	SYS4767 MSMS/Core integration*/%>

}
function PopulateScreen()
{
	subQuoteXML.SelectTag(null,"MORTGAGESUBQUOTE");

	m_sAmtRequested = subQuoteXML.GetTagText("AMOUNTREQUESTED");
	// DPF 09/10/02 - (CPWP1) - assign Drawdown to variable
	m_sDrawDown = subQuoteXML.GetTagText("DRAWDOWN");
	
	if(m_sAmtRequested != "0")
		frmScreen.txtAmtRequested.value = m_sAmtRequested;
	<% /* BMIDS00654 MDC 04/11/2002 */ %>
	else if(m_sAmountRequested != "0")
	{
		frmScreen.txtAmtRequested.value = m_sAmountRequested;
		m_sAmtRequested = m_sAmountRequested;
	}
	<% /* BMIDS00654 MDC 04/11/2002 - End */ %>
	
	frmScreen.txtTotalLoanAmt.value = subQuoteXML.GetTagText("TOTALLOANAMOUNT");
	<% /* MAR1061 purchaseprice*/ %>
	if(subQuoteXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE")!= "" &&
	   subQuoteXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE")!= "0")
		frmScreen.txtPurchasePrice.value = subQuoteXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE");
	else
	{
		//Get the purchaseprice from the ApplicationFactFind
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("APPLICATIONFACTFIND");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		XML.RunASP(document,"GetApplicationFFData.asp");
		if(XML.IsResponseOK())
		{
			XML.SelectTag(null,"APPLICATIONFACTFIND");
			frmScreen.txtPurchasePrice.value = XML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE");
		}
	}
	m_sPurchasePrice = frmScreen.txtPurchasePrice.value;
	
	<%/* SAB - 21/06/2006 - EP837 */%>
	if(subQuoteXML.GetTagText("LTV") != "0")
		frmScreen.txtLTVPercent.value = top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("LTV"),2);
	
	//MV - 29/01/2003 - BM0300 	
	//frmScreen.txtTotalMnthCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("TOTALNETMONTHLYCOST"), 2);
	frmScreen.txtTotalMnthCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("TOTALGROSSMONTHLYCOST"), 2);
	// DPF 09/10/02 - (CPWP1) - assign value from MSQ table to Cost less drawdown field
	frmScreen.txtMnthCostLessDD.value = top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("MONTHLYCOSTLESSDRAWDOWN"), 2);
	
	<% /* PSC 15/11/2002 BMIDS00964 */ %>
	m_bPPPresent = false;
	PopulateListBox(0);
	
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTotalLoanAmt");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtLTVPercent");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTotalMnthCost");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtMnthCostLessDD");
	if(scScrollTable.getRowSelected() == -1)
	{
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnReserve.disabled = true;

		
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration*/%>
		frmScreen.btnManAdjust.disabled = true;
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
	}
	else
	{
		frmScreen.btnEdit.disabled = false;
		frmScreen.btnDelete.disabled = false;
		frmScreen.btnReserve.disabled = false;

		<%/*BG		28/06/02	SYS4767 MSMS/Core integration*/%>
		frmScreen.btnManAdjust.disabled = false;
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
		
	}
	subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE");
	
	//DPF 01/10/2002 - CPWP1 - set DrawDown Field read/write property
	<% /* PSC 15/11/2002 BMIDS00964 - Disable if Part and part present and default to zero */ %>
	if ((subQuoteXML.GetTagText("FLEXIBLEMORTGAGEPRODUCT").length == 0) || 
		(subQuoteXML.GetTagText("FLEXIBLEMORTGAGEPRODUCT") == "0") ||
		m_bPPPresent)
	{
		<% /* PSC 15/11/2002 BMIDS00964 */ %>
		frmScreen.txtDrawDown.value = "0";
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDrawDown");
	}
	else
	{
		frmScreen.txtDrawDown.value = m_sDrawDown;
		scScreenFunctions.SetFieldState(frmScreen, "txtDrawDown", "W");
	}
	
	if(subQuoteXML.GetTagText("TOTALNETMONTHLYCOST") == "" || subQuoteXML.GetTagText("TOTALNETMONTHLYCOST") == "0")
	{
		frmScreen.btnOneOffCosts.disabled = true;
		frmScreen.btnPrint.disabled = true;
		frmScreen.btnNewSubquote.disabled = true;
		if(subQuoteXML.GetTagText("TOTALLOANAMOUNT") != subQuoteXML.GetTagText("AMOUNTREQUESTED"))frmScreen.btnCalculate.disabled = true;
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAmtRequested");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtPurchasePrice"); //MAR1061
		// DPF 09/10/02 - (CPWP1) - set drawdown field to read only
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDrawDown");
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnCalculate.disabled = true;

		
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration*/%>
		frmScreen.btnManAdjust.disabled = true;
	}	
		
	if(m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnReserve.disabled = true;
		frmScreen.btnCalculate.disabled = true;
		frmScreen.btnNewSubquote.disabled = true;
		frmScreen.btnGenerateQuote.disabled = true;  // EP2_55
	
	
		<%/*BG - 28/06/02	SYS4767 MSMS/Core integration*/%>
		frmScreen.btnManAdjust.disabled = true;
	}
	
	<%/*Begin: DS - EP2_1943*/%>
	var sMortgType ;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	
	if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW')))
		sMortgType = "PSW";
	else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE')))
		sMortgType = "TOE";
	else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP')))
		sMortgType = "NP";
	if ((sMortgType != null)&&(sMortgType == 'TOE')||(sMortgType == 'PSW')||(sMortgType == 'NP'))
	{
		frmScreen.btnNewSubquote.disabled = true;
		if(m_iNumOfLoanComponents == 0)
		{
			frmScreen.btnAdd.disabled = true;
		}
	}
	<%/* End: DS - EP2_1943*/%>
}


function PopulateListBox(nStart)
{
<%  /* Populate the listbox with values from subQuoteXML */
%>	subQuoteXML.ActiveTag = null;
	subQuoteXML.CreateTagList("LOANCOMPONENT");
	m_iNumOfLoanComponents = subQuoteXML.ActiveTagList.length;
	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, subQuoteXML.ActiveTagList.length);
	ShowList(nStart);
	if(m_iNumOfLoanComponents > 0)
	{
		scScrollTable.setRowSelected(1);	
		frmScreen.btnCalculate.disabled = false;
	}
	else
	frmScreen.btnCalculate.disabled = true;
}

function ShowList(nStart)
{
	scScrollTable.clear();
	for (var iCount = 0; iCount < subQuoteXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		subQuoteXML.SelectTagListItem(iCount + nStart);
		
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),subQuoteXML.GetTagText("PRODUCTNAME"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("RESOLVEDRATE"), 2));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(6),subQuoteXML.GetTagText("TOTALLOANCOMPONENTAMOUNT"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(7),GetAbbrevRepayMethod(subQuoteXML.GetTagText("REPAYMENTMETHOD")));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(8),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("NETMONTHLYCOST"), 2));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(9),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("APR"), 2));
		
		<%/* PSC 10/07/2002 BMIDS00062 */ %>
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(10),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("FINALRATEMONTHLYCOST"), 2));

		// EP2_55 - Add Indicator
		var sIndicator = ReturnIndicator(subQuoteXML.GetTagBoolean("PRODUCTSWITCHRETAINPRODUCTIND"),subQuoteXML.GetTagBoolean("MANUALPORTEDLOANIND"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sIndicator);

		<%/*scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),subQuoteXML.GetTagText("PRODUCTNAME"));				
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),subQuoteXML.GetTagText("TOTALLOANCOMPONENTAMOUNT"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5),subQuoteXML.GetTagAttribute("REPAYMENTMETHOD", "TEXT"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(6),top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("NETMONTHLYCOST"), 2));*/%>
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
		
		if (subQuoteXML.GetTagText("PORTEDLOAN") == "0" || subQuoteXML.GetTagText("PORTEDLOAN") == "")
		{
			var arrayReturn = GetProductAndRate();
			if( arrayReturn != null)
			{
				<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
				scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),arrayReturn[0]);
				scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),arrayReturn[1]);
				<% /* PSC 02/02/2007 EP2_1113 - End */ %> 
			}
		}
		
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}	
}
<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
function GetAbbrevRepayMethod(sRepayMethod)
{
<%	/* to condense the information in the listbox, output an abbreviated string */
%>	var sAbbrev;
	
	<% /* PSC 15/11/2002 BMIDS00964  - Start */ %>
	var xmlNode = null;
	var sSearch = "LISTNAME/LISTENTRY[VALUEID='" + sRepayMethod + "' and VALIDATIONTYPELIST/VALIDATIONTYPE='I']";	<% /* BMIDS01070 MDC 22/11/2002 */ %>
	xmlNode = m_XMLRepay.selectSingleNode(sSearch)
	
	<% /* Interest Only */ %>
	if (xmlNode != null)
		sAbbrev = "I/O";	
	else
	{
		sSearch = "LISTNAME/LISTENTRY[VALUEID='" + sRepayMethod + "' and VALIDATIONTYPELIST/VALIDATIONTYPE='C']";	<% /* BMIDS01070 MDC 22/11/2002 */ %>
		xmlNode = m_XMLRepay.selectSingleNode(sSearch)
		<% /* Capital And Interest */ %>
		if (xmlNode != null)
			sAbbrev = "C&I";
		else
		{
			sSearch = "LISTNAME/LISTENTRY[VALUEID='" + sRepayMethod + "' and VALIDATIONTYPELIST/VALIDATIONTYPE='P']";	<% /* BMIDS01070 MDC 22/11/2002 */ %>
			xmlNode = m_XMLRepay.selectSingleNode(sSearch)
			<% /* Part and Part */ %>
			if (xmlNode != null)
			{
				sAbbrev = "P&P";
				m_bPPPresent = true;
			}
			else
				sAbbrev = "??";
		}
	}
	<% /* PSC 15/11/2002 BMIDS00964  - End */ %>
	
	return sAbbrev;
}
<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>

function GetProductAndRate()
{
<%  /* All the information for this loancomponent is listed under the currently selected 
       loancomponent in subQuoteXML so there is no need to select a tag. We remember which
       taglist was active when we came in so we can reset it on the way out. */
%>  var arrayReturn = new Array(3);           // BMIDS878
	var bSuccess = false;
	var currentActiveTagList = subQuoteXML.ActiveTagList;
	var currentActiveTag = subQuoteXML.ActiveTag;
	var strProductType = "";
	 
	<% /* BMIDS878  Retrieve the Rate as well as the Step so that it can be used in CM101 */ %> 
	var strRate = "";                         // BMIDS878
	var strStep = "";
	
	if(subQuoteXML.GetTagText("NONPANELLENDEROPTION") == "0")
	{
<%		// get the baserate rate for future reference
%>		var strBaseRate = "0";
		if(subQuoteXML.SelectTag(currentActiveTag, "BASERATEBAND") != null)
			strBaseRate = subQuoteXML.GetTagText("RATE");
		
		subQuoteXML.ActiveTag = currentActiveTag;
		subQuoteXML.CreateTagList("INTERESTRATETYPE");
		<% /* BMIDS815 GHun
		for (var iCount = 0; iCount < subQuoteXML.ActiveTagList.length && bSuccess == false; iCount++)
		{
			subQuoteXML.SelectTagListItem(iCount);
			if(subQuoteXML.GetTagText("INTERESTRATETYPESEQUENCENUMBER") == "1"){
		*/ %>
			subQuoteXML.SelectTagListItem(0);
			strStep = subQuoteXML.GetTagText("INTERESTRATETYPESEQUENCENUMBER");
			<% /* BMIDS815 End */ %>
			
				var strRateType = subQuoteXML.GetTagText("RATETYPE");
				if(strRateType == "F"){
					strProductType = "Fixed";
					strRate = subQuoteXML.GetTagText("RATE");   // BMIDS878
					<% /* BMIDS815 GHun
					bSuccess = true;
					*/ %>
				}
				else if(strRateType == "B"){
					strProductType = "Base";
					strRate = strBaseRate;                      // BMIDS878
					<% /* BMIDS815 GHun
					bSuccess = true;
					*/ %>
				}
				else if(strRateType == "D"){
					strProductType = "Discounted";
					strRate = parseFloat(strBaseRate) - parseFloat(subQuoteXML.GetTagText("RATE"));  // BMIDS878					
					<% /* BMIDS815 GHun
					bSuccess = true;
					*/ %>
				}
				else if(strRateType == "C"){
					strProductType = "Capped/Floored";
					strRate = parseFloat(strBaseRate) - parseFloat(subQuoteXML.GetTagText("RATE"));  // BMIDS878
					if(parseFloat(strRate) < parseFloat(subQuoteXML.GetTagText("FLOOREDRATE")))
						strRate = subQuoteXML.GetTagText("FLOOREDRATE");
					if(parseFloat(strRate) > parseFloat(subQuoteXML.GetTagText("CEILINGRATE")))
						strRate = subQuoteXML.GetTagText("CEILINGRATE");					
					<% /* BMIDS815 GHun
					bSuccess = true;
					*/ %>
				}
				bSuccess = true;
		<%/*	}
		}	
		*/ %>
	}
	else
	{
		strProductType = subQuoteXML.GetTagText("MORTGAGEPRODUCTTYPE");
		strRate = subQuoteXML.GetTagText("INTERESTRATE1");
		bSuccess = true;
	}
	subQuoteXML.ActiveTagList = currentActiveTagList;
	if(bSuccess)
	{
		arrayReturn[0] = strProductType;
		arrayReturn[1] = strStep;
		arrayReturn[2] = top.frames[1].document.all.scMathFunctions.RoundValue(strRate, 2);
		
		return(arrayReturn);
	}
	else return(null);
}

function frmScreen.txtAmtRequested.onblur()
{
  // If the amount requested changes then reset 
	if(frmScreen.txtAmtRequested.value != m_sAmtRequested && 
	   frmScreen.txtAmtRequested.value != "")
	{
		<%/*MAR86 Start*/%>
		if (CheckCounterOffer()==2)
			return;
		<%/*MAR86 End*/%>

		<% /* PSC 12/02/2007 EP2_1314 - Start */ %>
		if(m_sAmtRequested != "" && m_sAmtRequested != "0" && !m_blnIsPort)
			Reset();
		else
		{
			if (frmScreen.txtAmtRequested.value != "" && frmScreen.txtTotalLoanAmt.value != "")
			{
				if (parseInt(frmScreen.txtAmtRequested.value) < parseInt(frmScreen.txtTotalLoanAmt.value))
				{
					alert("The amount requested cannot be less than the existing mortgage balance"); 
					frmScreen.txtAmtRequested.value = m_sAmtRequested;
					return;
				}
			}
					 
			CalcAndPopulateLTV();
			
			if (m_blnIsPort)
			{
				SaveMortgageSubQuoteDetails();
				Initialise();
			}
		}
		<% /* PSC 12/02/2007 EP2_1314 - End */ %>

		m_sAmtRequested = frmScreen.txtAmtRequested.value;
		
	}
	
	// EP2_2168 - Check Drawdown < AmtReq
	CheckDrawDownValid();
}

function frmScreen.txtPurchasePrice.onblur()
{
	// If the purchaseprice changes then reset 
	if(frmScreen.txtPurchasePrice.value != m_sPurchasePrice && 
	   frmScreen.txtPurchasePrice.value != "")
	{
		//First save it to AFF so the LTV will be calculated correctly in ResetMortgageSubquote
		SaveAFFPurchasePrice();
		if(m_sPurchasePrice != "" && m_sPurchasePrice != "0")Reset();
		else 
		{
			//if(SaveAFFPurchasePrice())
			<% /* PSC 02/02/2007 EP2_1113 */ %>
			if(frmScreen.txtAmtRequested.value != "" && parseInt(frmScreen.txtAmtRequested.value) != 0)
				CalcAndPopulateLTV();
		}
		m_sPurchasePrice = frmScreen.txtPurchasePrice.value;	
	}
	
	<% /* PSC 02/02/2007 EP2_1113 */ %> 
	ResetGenerateQuote();

}
function SaveAFFPurchasePrice()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("APPLICATIONFACTFIND");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	XML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value);
	XML.RunASP(document, "UpdateApplicationFactFind.asp");
	return(XML.IsResponseOK())
}
function Reset()
{
<%  /*Resets the screen after a change to AmountRequested/purchase price so that the loan composition
      for the mortgage subquote may be re-established. 
      Returns TRUE if the MSQ was reset and FALSE if it wasn't.
      
      First check for any loancomponents in the listbox. If there aren't any, we needn't
      reset.
      
      If we need to reset, check that the Amount Requested is present. If not, force the 
      user to enter one before trying to reset.*/
%>	var bContinue = true;

	if(bContinue)
	{
//		if( confirm("Changing Amount Requested, Type of Mortgage or Type of Buyer will reset the screen and will clear any loan components already established. Do you wish to continue?") == true)
//  BM0562 KRW 
		if( confirm("Changing Amount Requested or purchase price will reset the screen and will clear any loan components already established. Do you wish to continue?") == true)

		{
			
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
			XML.CreateRequestTag(window,null);
			              
			if(m_sApplicationMode == "Quick Quote")	
				XML.CreateActiveTag("QUICKQUOTE");
			else 
				XML.CreateActiveTag("APPLICATIONQUOTE");
				
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
			XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
			XML.CreateTag("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
			XML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value); //MAR1061
			XML.CreateTag("DRAWDOWN", frmScreen.txtDrawDown.value);
			
			AddCustomerList(XML);
			
			XML.SelectTag(null,"REQUEST");
			XML.SetAttribute("OPERATION","CriticalDataCheck"); 
			XML.CreateActiveTag("CRITICALDATACONTEXT");
			XML.SetAttribute("APPLICATIONNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null));
			XML.SetAttribute("APPLICATIONFACTFINDNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null));
			XML.SetAttribute("SOURCEAPPLICATION","Omiga");
			XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
			XML.SetAttribute("ACTIVITYINSTANCE","1");
			XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
			XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
			
			if(m_sApplicationMode == "Quick Quote")
				XML.SetAttribute("COMPONENT","omQQ.QuickQuoteBO");
			else 
				XML.SetAttribute("COMPONENT","omAQ.ApplicationQuoteBO");
			
			XML.SetAttribute("METHOD","ResetMortgageSubQuote");
			
			window.status = "Critical Data Check - please wait";
					
			switch (ScreenRules())
			{
				case 1: // Warning
				case 0: // OK
					XML.RunASP(document,"OmigaTMBO.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
			}

			window.status = "";
			
			if(XML.IsResponseOK()) 
				Initialise();
			
			XML = null;
		}
		else  //BM0562 KRW 
		{
			frmScreen.txtAmtRequested.value = m_sAmtRequested;
			frmScreen.txtPurchasePrice.value = m_sPurchasePrice; //MAR1061
		}
	}
	
	return false;
}

function AddCustomerList(XML)
{
	var tagCustomerList = XML.CreateActiveTag("CUSTOMERLIST");
	for(var nCount = 1; nCount < 6; nCount++)
	{
		var sCustomerNumber			= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nCount, "rf1111");
		var sCustomerVersionNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nCount, "1");
		if(sCustomerNumber != "" && sCustomerVersionNumber != "")
		{
			XML.ActiveTag = tagCustomerList;
			XML.CreateActiveTag("CUSTOMER");
			XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
		}
	}
}
function spnTable.onclick()
{
	if (scScrollTable.getRowSelected() != -1)
	{
		if(frmScreen.txtAmtRequested.readOnly == false)
		{
<%			/* Only enable editing if subquote has been calculated */
%>			frmScreen.btnEdit.disabled = false;
			frmScreen.btnDelete.disabled = false;
			frmScreen.btnReserve.disabled = false;

		}
	}
	else
	{
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnReserve.disabled = true;

	}
}

function CalcAndPopulateLTV()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("LTV");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	XML.CreateTag("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
		
	AddCustomerList(XML);
	// 	if(m_sApplicationMode == "Quick Quote")	XML.RunASP(document, "CalcCostModelLTV.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	//DPF 10/10/2002 - Applied manual fix to Screen Rules
	if(m_sApplicationMode == "Quick Quote")	
		// XML.RunASP(document, "CalcCostModelLTV.asp");
		// Added by automated update TW 01 Oct 2002 SYS5115
		{
			switch (ScreenRules())
			{
				case 1: // Warning
				case 0: // OK
					XML.RunASP(document, "CalcCostModelLTV.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
			}
		}
	// 	else XML.RunASP(document, "AQCalcCostModelLTV.asp");
	// Added by automated update TW 01 Oct 2002 SYS5115
	else
	{
		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				XML.RunASP(document, "AQCalcCostModelLTV.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
		}
	}

	if(XML.IsResponseOK())
	{
		XML.SelectTag(null, "RESPONSE"); 
		frmScreen.txtLTVPercent.value = XML.GetTagText("LTV");
	}
}

function GetFloat(sString)
{
<% /* This function returns a float representation of the sString passed in.
      It does full checks for NaN and 0 */
%>
	if( sString.length == 0 ) return 0.0;
	if( isNaN(parseFloat(sString))) return 0.0;
	else
	{	
		return( top.frames[1].document.all.scMathFunctions.RoundValue(sString, 2));
	}
}

function frmScreen.btnManAdjust.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	subQuoteXML.ActiveTag = null;			
	subQuoteXML.CreateTagList("LOANCOMPONENT");
	subQuoteXML.SelectTagListItem(nRowSelected -1);
	
	ArrayArguments[0] = m_sReadOnly;
	var arrayRateReturn = GetProductAndRate();
	subQuoteXML.ActiveTag = null;			
	subQuoteXML.CreateTagList("LOANCOMPONENT");
	subQuoteXML.SelectTagListItem(nRowSelected -1);
	
	<% /* BMIDS878  Retrieve Initial Rate from GetProductAndRate function */ %>
	ArrayArguments[1] = arrayRateReturn[2];
	var dOrigAdjustment;
	if(subQuoteXML.GetTagText("RESOLVEDRATE") == "")
		dOrigAdjustment = 0.0
	else
		// BMIDS878 Calculate Adjustment from Resolved Rate - Initial Rate
		dOrigAdjustment = top.frames[1].document.all.scMathFunctions.RoundValue((GetFloat(subQuoteXML.GetTagText("RESOLVEDRATE")) - GetFloat(arrayRateReturn[2])), 2);
	ArrayArguments[2] = dOrigAdjustment;
	ArrayArguments[3] = subQuoteXML.ActiveTag.xml; // correct loancomponent and product
	ArrayArguments[4] = subQuoteXML.XMLDocument.xml; // application information
	ArrayArguments[5] = m_sApplicationMode;
	
	<%/* SG 06/06/02 SYS4821 */%>
	<%/* var XML = new scXMLFunctions.XMLObject(); */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	ArrayArguments[6] = XML.CreateRequestAttributeArray(window);
	sReturn = scScreenFunctions.DisplayPopup(window, document, "CM101.asp", ArrayArguments, 335, 230);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		var dAdjustment = GetFloat(sReturn[1]);
		var sAdjustment = sReturn[1];
		if(dAdjustment != dOrigAdjustment)
		{
			//save the new adjustment.
			<%/* SG 06/06/02 SYS4821 */%>
			<%/* var LCXML = new scXMLFunctions.XMLObject(); */%>
			var LCXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
			LCXML.CreateRequestTag(window, null);
			LCXML.CreateActiveTag("LOANCOMPONENT");
			LCXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			LCXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
			LCXML.CreateTag("MORTGAGESUBQUOTENUMBER", subQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER"));
			LCXML.CreateTag("LOANCOMPONENTSEQUENCENUMBER", subQuoteXML.GetTagText("LOANCOMPONENTSEQUENCENUMBER"));
			LCXML.CreateTag("MANUALADJUSTMENTPERCENT", sAdjustment);
			LCXML.CreateTag("MANUALADJUSTUSERID", scScreenFunctions.GetContextParameter(window,"idUserId",null));
			// BMIDS878  Use Initial Rate instead of Step
			var dResolvedRate = parseFloat(GetFloat(arrayRateReturn[2])) + parseFloat(dAdjustment);
			LCXML.CreateTag("RESOLVEDRATE", dResolvedRate.toString());
			// 			LCXML.RunASP(document, "UpdateLoanComponent.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							LCXML.RunASP(document, "UpdateLoanComponent.asp");
					break;
				default: // Error
					LCXML.SetErrorResponse();
				}

			if(LCXML.IsResponseOK())
			{
				m_PopWinInd = 'Yes';
				Initialise();
			}
		}
	}
}


function frmScreen.btnAdd.onclick()
{
	var bContinue = true;
	if(m_sApplicationMode == "Quick Quote")
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var lMaxQQComponents = parseInt(XML.GetGlobalParameterAmount(document,"MaximumQQComponents"));
		if (m_iNumOfLoanComponents == lMaxQQComponents )
		{
			alert("Maximum number of loan components allowed already created on this mortgage sub-quote");
			bContinue = false;
		}
	}
	else
	{
		if(subQuoteXML != null)
		{
			if( subQuoteXML.SelectTag(null, "MORTGAGELENDER") != null)
			{
				if(m_iNumOfLoanComponents == parseInt(subQuoteXML.GetTagText("MAXNOLOANS")))
				{
					alert("Maximum number of loan components allowed already created on this mortgage sub-quote");
					bContinue = false;
				}
			}
			if( bContinue && (parseInt(frmScreen.txtTotalLoanAmt.value) == parseInt(frmScreen.txtAmtRequested.value)) )
			{
				alert("Total of loans already equals amount requested; it is not possible to add another loan component");
				bContinue = false;
			}
		}
	}
	if(bContinue == true)
	{
		if(frmScreen.txtAmtRequested.value == "" || frmScreen.txtAmtRequested.value == "0")
		{
			alert("Please enter amount requested before adding loan components");
			frmScreen.txtAmtRequested.focus();
		}
		else if(frmScreen.txtPurchasePrice.value == "" || frmScreen.txtPurchasePrice.value == "0")
		{
			alert("Please enter a purchase price before adding loan components"); //MAR1061
			frmScreen.txtPurchasePrice.focus();
		}
		else
		{
<%			/* First make sure there is an LTV present */
%>			if(frmScreen.txtLTVPercent.value == "") CalcAndPopulateLTV();

<%			/* Update the mortgagesubquote on the database */
%>			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			if(subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE") != null)
			{
				subQuoteXML.SetTagText("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
				subQuoteXML.SetTagText("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value); //MAR1061
				subQuoteXML.SetTagText("LTV", frmScreen.txtLTVPercent.value);
				// DPF 09/10/02 - (CPWP1) - save drawdown amount
				subQuoteXML.SetTagText("DRAWDOWN", frmScreen.txtDrawDown.value);

				XML.CreateRequestTag(window,null);
				XML.CreateActiveTag("MORTGAGESUBQUOTE");
				XML.AddXMLBlock(subQuoteXML.CreateFragment());
				// 				if(m_sApplicationMode == "Quick Quote")XML.RunASP(document, "QQUpdateMortgageSubQuote.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				//DPF 10/10/2002 - manual adjustment to screen rules
				if(m_sApplicationMode == "Quick Quote")
				{
					switch (ScreenRules())
					{
						case 1: // Warning
						case 0: // OK
							XML.RunASP(document, "QQUpdateMortgageSubQuote.asp");
							break;
						default: // Error
							XML.SetErrorResponse();
					}
				}
				// 				else XML.RunASP(document, "AQUpdateMortgageSubQuote.asp");
				// Added by automated update TW 01 Oct 2002 SYS5115
				else
				{
					switch (ScreenRules())
					{
						case 1: // Warning
						case 0: // OK
							XML.RunASP(document, "AQUpdateMortgageSubQuote.asp");
							break;
						default: // Error
							XML.SetErrorResponse();
					}
				}

				if(XML.IsResponseOK() != true)
				{
					bContinue = false;
					btnSubmit.focus();
				}
			}
			if(bContinue)
			{
				CreateCM100ParamsXML(); // EP2_8 - Create XML with new params
				var sReturn = null;
				var ArrayArguments = new Array(17);  //DRC MAR1380  // AS EP2_8 Extra Params.
				ArrayArguments[0] = m_sApplicationMode;
				ArrayArguments[1] = scScreenFunctions.GetContextParameter(window,"idUserId",null);
				ArrayArguments[2] = scScreenFunctions.GetContextParameter(window,"idUserType",null);
				ArrayArguments[3] = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
				ArrayArguments[4] = m_sReadOnly;
				if(subQuoteXML != null)	ArrayArguments[5] = subQuoteXML.XMLDocument.xml;
				else ArrayArguments[5] = null;
				ArrayArguments[6] = null; //Loan Component sequence number
				ArrayArguments[7] = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","30");
				ArrayArguments[8] = m_iNumOfLoanComponents;
				ArrayArguments[9] = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
				ArrayArguments[10] = XML.CreateRequestAttributeArray(window);
				ArrayArguments[11] = m_sCurrency;
				ArrayArguments[12] = scScreenFunctions.GetContextParameter(window,"idApplicationDate","");  //JD BMIDS816
				ArrayArguments[13] = scScreenFunctions.GetContextParameter(window,"idActivityId","");  //DRC MAR1380
				ArrayArguments[14] = scScreenFunctions.GetContextParameter(window,"idStageId","");  //DRC MAR1380
				ArrayArguments[15] = scScreenFunctions.GetContextParameter(window,"idApplicationPriority","");  //DRC MAR1380
				ArrayArguments[16] = "Add"  // Add or Edit - Need mode on CM110.
				if(CM100ParamsXML != null)	ArrayArguments[17] = CM100ParamsXML.XMLDocument.xml;
				else ArrayArguments[17] = null; // EP2_8 - Need on CM110 

				<%// IK 24/06/2006 EP771 %>
				sReturn = scScreenFunctions.DisplayPopup(window, document, "CM110.asp", ArrayArguments, 630, 560);

				if(sReturn != null && m_sReadOnly != "1")
				{
					m_PopWinInd = 'Yes';
					FlagChange(sReturn[0]);
					Initialise();
					if(parseFloat(frmScreen.txtTotalLoanAmt.value) == parseFloat(frmScreen.txtAmtRequested.value)) 
						frmScreen.btnCalculate.disabled = false;
				}
			}
		}		
	}
}

function frmScreen.btnEdit.onclick()
{
<%  /* locate the LOANCOMPONENT XML section for the currently selected loancomponent */
%>  var bContinue = true;
	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	subQuoteXML.ActiveTag = null;			
	subQuoteXML.CreateTagList("LOANCOMPONENT");
	subQuoteXML.SelectTagListItem(nRowSelected -1);
	if(subQuoteXML.GetTagText("PORTEDLOAN") == "1")alert("Ported loan components cannot be edited or deleted.");
	else
	{
		// Update the mortgagesubquote on the database
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if(subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE") != null)
		{
			subQuoteXML.SetTagText("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
			subQuoteXML.SetTagText("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value);
			subQuoteXML.SetTagText("LTV", frmScreen.txtLTVPercent.value);
			// DPF 09/10/02 - (CPWP1) - save drawdown amount
			subQuoteXML.SetTagText("DRAWDOWN", frmScreen.txtDrawDown.value);
			
			XML.CreateRequestTag(window,null);
			XML.CreateActiveTag("MORTGAGESUBQUOTE");
			XML.AddXMLBlock(subQuoteXML.CreateFragment());
			// 			if(m_sApplicationMode == "Quick Quote")XML.RunASP(document, "QQUpdateMortgageSubquote.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			//DPF 10/10/02 - Applied manual fix to Screenrules
			if(m_sApplicationMode == "Quick Quote")
			{
				switch (ScreenRules())
				{
					case 1: // Warning
					case 0: // OK
						XML.RunASP(document, "QQUpdateMortgageSubquote.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
				}
			}
			// 			else XML.RunASP(document, "AQUpdateMortgageSubquote.asp");
			// Added by automated update TW 01 Oct 2002 SYS5115
			else
			{
				switch (ScreenRules())
				{
					case 1: // Warning
					case 0: // OK
						XML.RunASP(document, "AQUpdateMortgageSubquote.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
				}
			}

			if(XML.IsResponseOK() != true)
			{
				bContinue = false;
				btnSubmit.focus();
			}
		}
		if(bContinue)
		{
			CreateCM100ParamsXML(); // EP2_8 - Create XML with new params
			subQuoteXML.ActiveTag = null;
			subQuoteXML.CreateTagList("LOANCOMPONENT");
			subQuoteXML.SelectTagListItem(nRowSelected -1);
			var sReturn = null;
			var ArrayArguments = new Array(17);  //DRC MAR1380  // AS EP2_8 Extra Params.
			ArrayArguments[0] = m_sApplicationMode;
			ArrayArguments[1] = scScreenFunctions.GetContextParameter(window,"idUserId",null);
			ArrayArguments[2] = scScreenFunctions.GetContextParameter(window,"idUserType",null);
			ArrayArguments[3] = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
			ArrayArguments[4] = m_sReadOnly;
			ArrayArguments[5] = subQuoteXML.XMLDocument.xml;
			ArrayArguments[6] = subQuoteXML.GetTagText("LOANCOMPONENTSEQUENCENUMBER");
			ArrayArguments[7] = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","30");
			ArrayArguments[8] = m_iNumOfLoanComponents;
			ArrayArguments[9] = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
			ArrayArguments[10] = XML.CreateRequestAttributeArray(window);
			ArrayArguments[11] = m_sCurrency;
			ArrayArguments[12] = scScreenFunctions.GetContextParameter(window,"idApplicationDate","");  //JD BMIDS816
			ArrayArguments[13] = scScreenFunctions.GetContextParameter(window,"idActivityId","");  //DRC MAR1380
			ArrayArguments[14] = scScreenFunctions.GetContextParameter(window,"idStageId","");  //DRC MAR1380
			ArrayArguments[15] = scScreenFunctions.GetContextParameter(window,"idApplicationPriority","");  //DRC MAR1380
			ArrayArguments[16] = "Edit"  // Add or Edit - Need mode on CM110.
			if(CM100ParamsXML != null)	ArrayArguments[17] = CM100ParamsXML.XMLDocument.xml;
			else ArrayArguments[17] = null; // EP2_8 - Need on CM110 
			<%//SW 20/06/2006 EP771%>
			sReturn = scScreenFunctions.DisplayPopup(window, document, "CM110.asp", ArrayArguments, 630, 540);

			if(sReturn != null && m_sReadOnly != "1")
			{
				m_PopWinInd = 'Yes';
				FlagChange(sReturn[0]);
				Initialise();
				if(parseFloat(frmScreen.txtTotalLoanAmt.value) == parseFloat(frmScreen.txtAmtRequested.value)) 
					frmScreen.btnCalculate.disabled = false;
			}
		}	
	}
}

function frmScreen.btnDelete.onclick()
{
<%  /* locate the LOANCOMPONENT XML section for the currently selected loancomponent */
%>  var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	subQuoteXML.ActiveTag = null;			
	subQuoteXML.CreateTagList("LOANCOMPONENT");
	subQuoteXML.SelectTagListItem(nRowSelected -1);
	if(subQuoteXML.GetTagText("PORTEDLOAN") == "1")alert("Ported loan components cannot be edited or deleted.");
	else
	{
		if(confirm("This will delete the selected loan component. Are you sure?") == true)
		{
			var sLoanAmount = subQuoteXML.GetTagText("LOANAMOUNT");
			var sMnthlyCost = subQuoteXML.GetTagText("NETMONTHLYCOST");
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTag(window,"DELETE");
			XML.CreateActiveTag("MORTGAGESUBQUOTE");
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
			XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
			XML.CreateTag("LOANCOMPONENTSEQUENCENUMBER", subQuoteXML.GetTagText("LOANCOMPONENTSEQUENCENUMBER"));
			// 			XML.RunASP(document,"DeleteLoanComponent.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document,"DeleteLoanComponent.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

			if(XML.IsResponseOK())
			{
				/* delete this loancomponent from the subQuoteXML and repopulate the listbox.
					Also, need to correct the Total Loan Amount field as we are not re-getting
					the mortgage subquote. The DeleteLoanComponent() function should update the
					database with the new total loan amount
					
				BG 05/12/00 SYS1653 I have commented out the following code as it was not doing
				as the spec required and was causing the problem found in this AQR - I have added
				Initialise(); to reinitialise subQuoteXML.
				
				var tagNode = subQuoteXML.SelectTag(null, "LOANCOMPONENTLIST");
				tagNode.removeChild(tagNode.childNodes.item(nRowSelected -1));
				if( parseInt(frmScreen.txtTotalLoanAmt.value) > 0)frmScreen.txtTotalLoanAmt.value = (parseFloat(frmScreen.txtTotalLoanAmt.value) - parseFloat(sLoanAmount));
				if( parseInt(frmScreen.txtTotalMnthCost.value) > 0)frmScreen.txtTotalMnthCost.value = (parseFloat(frmScreen.txtTotalMnthCost.value) - parseFloat(sMnthlyCost));
				PopulateListBox(0);*/
				
				m_PopWinInd = 'Yes';
				Initialise();
				frmScreen.btnCalculate.disabled = true;
			}
		}
	}
}

function frmScreen.btnCalculate.onclick()
{
	
	<% /* BMIDS815 GHun Fix hour glass*/ %>
	frmScreen.btnCalculate.style.cursor = "wait";

	<% /* Remove focus from Calculate so that Enter cannot be pressed multiple times */ %>
	frmScreen.btnCalculate.blur();

	<% /* Call Calculate Processing after a timeout to allow the cursor time to change. */ %>
	window.setTimeout("CalculateProcessing()", 0)
	<% /* BMIDS815 End */ %>
	
}

<% /* BMIDS815 GHun Moved code from btnCalculate.onclick() into CalculateProcessing */ %>
function CalculateProcessing()
{	

	<% /* SR 31/08/2004 : BMIDS815  */ %>
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	var XMLTemp = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject(); 
	var xmlMainTag;
	
	var blnFurtherAdvance = XMLTemp.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('GA'));
	var blnTranferOfEquity =  XMLTemp.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE'));
	
	XML.CreateRequestTag(window, null);
	
	if(blnFurtherAdvance) 
		xmlMainTag = XML.CreateActiveTag("REFRESHMORTGAGEACCOUNTDATA");
	else 
		xmlMainTag = XML.CreateActiveTag("GETANDSAVEPORTEDSTEPANDPERIODFROMMORTGAGEACCOUNT"); 
		
	if(blnFurtherAdvance) XML.CreateTag("PORTINGINDICATOR", "0");
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	XML.CreateTag("TYPEOFAPPLICATION", scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue",null));
	XML.CreateTag("BMACCOUNTNUMBER", scScreenFunctions.GetContextParameter(window,"idOtherSystemAccountNumber",null));
	
	XML.ActiveTag = xmlMainTag;	
	AddOtherSystemCustomerNumbers(XML)
	
	XML.ActiveTag = xmlMainTag ;
	XML.CreateActiveTag("MORTGAGESUBQUOTE");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
	XML.CreateTag("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
	XML.CreateTag("LTV", frmScreen.txtLTVPercent.value);
		
	if(blnFurtherAdvance)
	{ <% /* Refresh MortgageAccount data from Admin system */ %>
		<% /* JD BMIDS876 set window status as account refresh can take a while. */ %>
		window.status = "Refreshing account data ...";
		XML.RunASP(document,"RefreshMortgageAccountData.asp");
		window.status = "";
		if(!XML.IsResponseOK())
		{
			frmScreen.btnCalculate.style.cursor = "default";
			return;
		}
		else
		{
			if(XML.SelectSingleNode("//LTVCHANGED"))
			{
				m_sLTVChanged = XML.ActiveTag.text ;
				scScreenFunctions.SetContextParameter(window,"idLTVChanged", m_sLTVChanged);
				
				XML.ActiveTag = null ;
				if(XML.SelectSingleNode("//LTV"))
				{
					var sLTV = XML.ActiveTag.text ;
					frmScreen.txtLTVPercent.value = sLTV;
				}				
			}			
		}
	}
	else
	{
		window.status = "Refreshing account data ..."; 
		XML.RunASP(document,"GetAndSavePortedStepAndPeriodFromMortAcc.asp");
		window.status = "";
		if(!XML.IsResponseOK())
		{
			frmScreen.btnCalculate.style.cursor = "default";
			return;
		}
	}
	<% /* SR 31/08/2004 : BMIDS815 - End */ %>
	
	SaveMortgageSubQuoteDetails();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("MORTGAGECOSTS");
	XML.CreateTag("CONTEXT", m_sApplicationMode);
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	XML.CreateTag("APPLICATIONDATE", scScreenFunctions.GetContextParameter(window,"idApplicationDate","")); //JD BMIDS763
	XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idQuotationNumber","1"));
	XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
	XML.CreateTag("LIFESUBQUOTENUMBER", m_sLifeSubquoteNumber);
	CustomerList(XML);

	// 	if(m_sApplicationMode == "Quick Quote")XML.RunASP(document, "QQCalculateMortgageCosts.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	// DPF 10/10/2002 - Applied manual fix to screen rules
	if(m_sApplicationMode == "Quick Quote")
	{
		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				<% /* BMIDS876 GHun set window status as calculation can take a while. */ %>
				window.status = "Calculating costs ...";
				XML.RunASP(document, "QQCalculateMortgageCosts.asp");
				window.status = "";
				break;
			default: // Error
				XML.SetErrorResponse();
		}
	}
	// 	else XML.RunASP(document, "AQCalculateMortgageCosts.asp");
	// Added by automated update TW 01 Oct 2002 SYS5115
	else
	{
		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				<% /* BMIDS876 GHun set window status as calculation can take a while. */ %>
				window.status = "Calculating costs ...";
				XML.RunASP(document, "AQCalculateMortgageCosts.asp");
				window.status = "";
				break;
			default: // Error
				XML.SetErrorResponse();
		}
	}

	if(XML.IsResponseOK())
	{
		m_PopWinInd = 'Yes';
		Initialise();
					
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtAmtRequested");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtPurchasePrice"); //MAR1061
		frmScreen.btnReserve.disabled = false;
		frmScreen.btnOneOffCosts.disabled = false;
		frmScreen.btnPrint.disabled = false;
		frmScreen.btnNewSubquote.disabled = false;
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnCalculate.disabled = true;


		<% /* BMIDS00878 MDC 08/11/2002 */ %>
		RunIncomeCalculations();
		<% /* BMIDS00878 MDC 08/11/2002 */ %>
	}
	
	
	<%/*Begin: DS - EP2_1943*/%>
	var sMortgType ;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	
	if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW')))
		sMortgType = "PSW";
	else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE')))
		sMortgType = "TOE";
	else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP')))
		sMortgType = "NP";
	if ((sMortgType != null)&&(sMortgType == 'TOE')||(sMortgType == 'PSW')||(sMortgType == 'NP'))
	{
		frmScreen.btnNewSubquote.disabled = true;
		if(m_iNumOfLoanComponents == 0)
		{
			frmScreen.btnAdd.disabled = true;
		}
	}
	<%/* End: DS - EP2_1943*/%>
	
	XML = null;
	frmScreen.btnCalculate.style.cursor = "default";
}
<% /* BMIDS815 End */ %>

<% /* SR 31/08/2004 : BMIDS815  */ %>
function AddOtherSystemCustomerNumbers(XML)
{
	var sCustomerNumber, sCustomerVersionNumber, sCustomerRoleType, sOtherSystemCustomerNumber ;
	
	var ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(window,"SEARCH");
	ListXML.CreateActiveTag("APPLICATION");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	ListXML.RunASP(document,"FindCustomersForApplication.asp");

	if (ListXML.IsResponseOK())
	{
		var tagCustomerList = XML.CreateActiveTag("CUSTOMERLIST");
		ListXML.CreateTagList("CUSTOMERROLE");
		var iNoOfCustomers = ListXML.ActiveTagList.length;
		for(var nCount = 0; nCount < iNoOfCustomers; nCount++)
		{	
			ListXML.SelectTagListItem(nCount);
			sCustomerNumber = ListXML.GetTagText("CUSTOMERNUMBER");
			sCustomerVersionNumber = ListXML.GetTagText("CUSTOMERVERSIONNUMBER");
			sCustomerRoleType = ListXML.GetTagText("CUSTOMERROLETYPE");
			sOtherSystemCustomerNumber = ListXML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER");
			if(sCustomerNumber != "" && sCustomerVersionNumber != "" && sCustomerRoleType == "1")
			{
				XML.ActiveTag = tagCustomerList;
				XML.CreateActiveTag("CUSTOMER");
				XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
				XML.CreateTag("OTHERSYSTEMCUSTOMERNUMBER", sOtherSystemCustomerNumber);
			}
		}
	}
}
<% /* SR 31/08/2004 : BMIDS815  - End */ %>

function frmScreen.btnNewSubquote.onclick()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idQuotationNumber","1"));
	XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
	XML.CreateTag("LIFESUBQUOTENUMBER", m_sLifeSubquoteNumber);
	if(m_sApplicationMode == "Quick Quote")XML.CreateTag("QUOTATIONTYPE", "1");
	else XML.CreateTag("QUOTATIONTYPE", "2");
	// 	XML.RunASP(document, "CreateNewMortgageLifeSubquote.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "CreateNewMortgageLifeSubquote.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if(XML.IsResponseOK())
	{
		XML.SelectTag(null, "MORTGAGESUBQUOTE");
		m_sMortgageSubquoteNumber = XML.GetTagText("MORTGAGESUBQUOTENUMBER");
		m_sLifeSubquoteNumber = XML.GetTagText("LIFESUBQUOTENUMBER");
		scScreenFunctions.SetContextParameter(window,"idMortgageSubQuoteNumber",m_sMortgageSubquoteNumber);
		scScreenFunctions.SetContextParameter(window,"idLifeSubquoteNumber",m_sLifeSubquoteNumber);
		
		m_PopWinInd = 'Yes';
		Initialise();
		
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtAmtRequested");
		scScreenFunctions.SetFieldToWritable(frmScreen, "txtPurchasePrice"); //MAR1061
		//DPF 09/10/2002 - CPWP1 (BM037) - new quote so empty monthly cost less drawdown
		frmScreen.txtMnthCostLessDD.value = "";
		frmScreen.btnAdd.disabled = false;
		frmScreen.btnCalculate.disabled = false;
		if(scScrollTable.getRowSelected() == -1)
		{
			frmScreen.btnEdit.disabled = true;
			frmScreen.btnDelete.disabled = true;
		}
		frmScreen.btnReserve.disabled = true;
		frmScreen.btnOneOffCosts.disabled = true;
		frmScreen.btnPrint.disabled = true;
		frmScreen.btnNewSubquote.disabled = true;

	}
	XML = null;
}

function frmScreen.btnOneOffCosts.onclick()
{
	var sReturn = null;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sApplicationType;

	//CMWP3 & BM0133 - extended array size.
	var ArrayArguments = new Array(12);
	ArrayArguments[0] = m_sApplicationMode;
	ArrayArguments[1] = m_sApplicationNumber;
	ArrayArguments[2] = m_sApplicationFFNumber;
	ArrayArguments[3] = m_sMortgageSubquoteNumber;
	ArrayArguments[4] = m_sLifeSubquoteNumber;
	ArrayArguments[5] = m_sReadOnly;
	ArrayArguments[6] = XML.CreateRequestAttributeArray(window);
	ArrayArguments[7] = m_sCurrency;
	XML.ResetXMLDocument();
	AddCustomerList(XML);
	ArrayArguments[8] = XML.XMLDocument.xml;
	
	//CMWP3 - two new arguments added to array.
	ArrayArguments[9] = m_iINTERESTONLYAMOUNT;
	ArrayArguments[10] = m_iCAPITALINTERESTAMOUNT;
	
	//BMIDS00808 - new argument - pass across the drawdown amount
	ArrayArguments[11] = frmScreen.txtDrawDown.value;
	
	//BM0133 - new argument - pass across the Application Type
	//Get Application type from context
	sApplicationType = scScreenFunctions.GetContextParameter(window, "idMortgageApplicationValue", null);
	
	ArrayArguments[12] = sApplicationType;
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "CM130.asp", ArrayArguments, 425, 550);

	if(sReturn != null)
	{
		m_PopWinInd = 'Yes';
		FlagChange(sReturn[0]);
		Initialise();
	}
	XML = null;
}

function frmScreen.btnReserve.onclick()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("MORTGAGESUBQUOTE");
	
<%	/* Find the section of subQuoteXML that corresponds to the selected loan component */
%>	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	subQuoteXML.ActiveTag = null;			
	subQuoteXML.CreateTagList("LOANCOMPONENT");
	subQuoteXML.SelectTagListItem(nRowSelected -1);
	
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	//EP852 AW 23/06/06
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	XML.CreateTag("MORTGAGEPRODUCTCODE", subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE"));
	XML.CreateTag("STARTDATE", subQuoteXML.GetTagText("STARTDATE"));
	if(subQuoteXML.GetTagText("PORTEDLOAN") != "")XML.CreateTag("PORTEDLOAN", subQuoteXML.GetTagText("PORTEDLOAN"));
	else XML.CreateTag("PORTEDLOAN", "0");
	// 	XML.RunASP(document, "ReserveMortgageProduct.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "ReserveMortgageProduct.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if(XML.IsResponseOK())alert("Selected mortgage product has been reserved.");

	XML = null;	
}
function frmScreen.btnViewSubquotes.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","CM100");
	scScreenFunctions.SetContextParameter(window,"idBusinessType","Mortgage"); // CHANGE TO 'M'
	frmToCM075.submit();
}
function frmScreen.btnPrint.onclick()
{	
	var GlobalXML = null;
	GlobalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var iDocumentID = parseInt(GlobalXML.GetGlobalParameterAmount(document,"PrintCostModelHostTemplateId"));

	AttribsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AttribsXML.CreateRequestTag(window, "GetPrintAttributes");
	AttribsXML.CreateActiveTag("FINDATTRIBUTES");
	AttribsXML.SetAttribute("HOSTTEMPLATEID", iDocumentID);
					
	// 	AttribsXML.RunASP(document, "PrintManager.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			AttribsXML.RunASP(document, "PrintManager.asp");
			break;
		default: // Error
			AttribsXML.SetErrorResponse();
		}

	if(AttribsXML.IsResponseOK())
	{
		if(AttribsXML.GetTagAttribute("ATTRIBUTES", "INACTIVEINDICATOR") == 1)
		{
			alert("This document template " + AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME") + " is inactive.");		
		}
		else
		{	
			var sPrinterType = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
			var ValidXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			if (ValidXML.IsInComboValidationList(document,"PrinterDestination",sPrinterType,["L"]) == false)
			{				
				alert("The document template printer destination has been defined incorrectly.  See your system administrator.");
			}
			else
			{
				//Call PrintBO.PrintDocument
			
				var sLocalPrinters = GetLocalPrinters();
				LocalPrintersXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				LocalPrintersXML.LoadXML(sLocalPrinters);
				LocalPrintersXML.SelectTag(null, "RESPONSE");	
			
				var sPrinter = LocalPrintersXML.GetTagText("PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
									
				if(sPrinter == "")
				{					
					alert("You do not have a default printer set on your PC.");			
				}		
				else
				{
					TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					var sGroups = new Array("PrinterDestination");
					var sList = TempXML.GetComboLists(document,sGroups);
					var sValueID = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
					var sPrintType = TempXML.GetTagText("LIST/LISTNAME/LISTENTRY[VALUEID='" + sValueID + "']/VALIDATIONTYPELIST/VALIDATIONTYPE");
								
					PrintXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					PrintXML.CreateRequestTag(window, "PrintDocument");
											
					PrintXML.SelectTag(null, "REQUEST");
					
					//BG 10/10/02 BMIDS00612 
					PrintXML.SetAttribute("PRINTINDICATOR", "1");
					//BG 10/10/02 BMIDS00612 - END	
							
					PrintXML.CreateActiveTag("CONTROLDATA");
					var iCopies = AttribsXML.GetTagAttribute("ATTRIBUTES", "DEFAULTCOPIES")
					if(iCopies == "")
					{
						iCopies = 1;
					}
					PrintXML.SetAttribute("COPIES", iCopies);
					PrintXML.SetAttribute("PRINTER", sPrinter);
					PrintXML.SetAttribute("DOCUMENTID", AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATEID"));
					PrintXML.SetAttribute("DPSDOCUMENTID", AttribsXML.GetTagAttribute("ATTRIBUTES", "DPSTEMPLATEID"));

					// IK_BMIDS00830 add dms attributes
					PrintXML.SetAttribute("TEMPLATEGROUPID", AttribsXML.GetTagAttribute("ATTRIBUTES", "TEMPLATEGROUPID"));
					PrintXML.SetAttribute("HOSTTEMPLATENAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME"));
					PrintXML.SetAttribute("HOSTTEMPLATEDESCRIPTION", AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATEDESCRIPTION"));
					PrintXML.SetAttribute("DOCUMENTNAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "DOCUMENTNAME"));
					// IK_BMIDS00830_ends

					PrintXML.SetAttribute("DESTINATIONTYPE", sPrintType);
						
					PrintXML.SelectTag(null, "REQUEST");						
					PrintXML.CreateActiveTag("PRINTDATA");
					PrintXML.SetAttribute("METHODNAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "PDMMETHOD"));	
					PrintXML.SetAttribute("METHODNAME", "CostModellingTemplate");	
													
					PrintXML.SetAttribute("APPLICATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null));
					PrintXML.SetAttribute("APPLICATIONFACTFINDNUMBER", scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null));
									
					PrintXML.SelectTag(null, "REQUEST");						
					PrintXML.CreateActiveTag("TEMPLATEDATA");
					PrintXML.SetAttribute("MORTGAGESUBQUOTENUMBER", scScreenFunctions.GetContextParameter(window,"idMortgageSubQuoteNumber",null));	
						
					// 					PrintXML.RunASP(document, "PrintManager.asp");
					// Added by automated update TW 09 Oct 2002 SYS5115
					switch (ScreenRules())
						{
						case 1: // Warning
						case 0: // OK
											PrintXML.RunASP(document, "PrintManager.asp");
							break;
						default: // Error
							PrintXML.SetErrorResponse();
						}

	
					if(PrintXML.IsResponseOK())
					{
						alert("Document has been sent to print.");
					}																	
				}
			}
		}
	}	
}

// DPF 09/10/02 - (CPWP1) - the save process has been moved to function
function SaveMortgageSubQuoteDetails()
{
// First make sure there is an LTV present 
if(frmScreen.txtLTVPercent.value == "") CalcAndPopulateLTV();

// Update the mortgagesubquote on the database
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if(subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE") != null)
	{
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("MORTGAGESUBQUOTE");

		<% /* EP974/MAR1888 PE/GHun Only update values that may have changed */ %>
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
		XML.CreateTag("AMOUNTREQUESTED", frmScreen.txtAmtRequested.value);
		XML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value); //MAR1061
		XML.CreateTag("LTV", frmScreen.txtLTVPercent.value);
		// DPF 09/10/02 - (CPWP1) - save drawdown amount
		XML.CreateTag("DRAWDOWN", frmScreen.txtDrawDown.value);
		<% /* MAR1888 End */ %>


		//if(m_sApplicationMode == "Quick Quote") XML.RunASP(document, "QQUpdateMortgageSubQuote.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		if(m_sApplicationMode == "Quick Quote")
		{
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
					XML.RunASP(document, "QQUpdateMortgageSubQuote.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}
		}
		else		
		// else XML.RunASP(document, "AQUpdateMortgageSubQuote.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		{
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
					XML.RunASP(document, "AQUpdateMortgageSubQuote.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}
		}
		
		if(XML.IsResponseOK() != true)
		{
			bContinue = false;
			btnSubmit.focus();
			<% /* BM0135 MDC 04/12/2002 */ %>
			return false;
			<% /* BM0135 MDC 04/12/2002 - End */ %>
		}
	}
	<% /* BM0135 MDC 04/12/2002 */ %>
	return true;
	<% /* BM0135 MDC 04/12/2002 - End */ %>

}
function btnSubmit.onclick()
{
	<%/*MAR86 Start*/%>
	var returnVal = 0;
	returnVal = CheckCounterOfferOnSubmit();
	//alert(returnVal);
	if(returnVal==2)
		btnCancel.onclick();
	else
		if(returnVal == 1)
			return;
			
	<%/*MAR86 End*/%>
	if(frmScreen.onsubmit())
	{
		<% /* BM0135 MDC 04/12/2002 */ %>
		if(SaveMortgageSubQuoteDetails())
		{
			scScreenFunctions.SetContextParameter(window,"idNoCompleteCheck","1");
			if(scScreenFunctions.CompletenessCheckRouting(window))
				frmToGN300.submit();
			else 
				frmToCM010.submit();
		}
		<% /* BM0135 MDC 04/12/2002 - End */ %>
		
		<% /* SR 08/09/2004 : BMIDS815 */ %>
		if (m_sLTVChanged == '1')
			scScreenFunctions.SetContextParameter(window, "idCallingScreenID", "CM100");
		<% /* SR 08/09/2004 : BMIDS815 - End */ %>	
	}
}

function CustomerList(XML)
{
	var tagCustomerList = XML.CreateActiveTag("CUSTOMERLIST");
	var nCustCount = 0;
	for(var nCount = 1; nCount < 6; nCount++)
	{
		if(nCustCount < "2")
		{
			var sCustomerNumber			= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nCount, "rf1111");
			var sCustomerVersionNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nCount, "1");
			var sCustomerRoleType		= scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nCount, "1");
			if(sCustomerNumber != "" && sCustomerVersionNumber != "" && sCustomerRoleType == "1")
			{
				XML.ActiveTag = tagCustomerList;
				XML.CreateActiveTag("CUSTOMER");
				XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
				nCustCount = nCustCount + 1;
			}
		}
	}
}

<% /* ASu - BMIDS00431 Add Cancel Button - Start */ %>
function btnCancel.onclick()
{
	<% /* DB BM0443 - If cancel clicked ensure that completeness check is not re-run in cm010. */ %>
	scScreenFunctions.SetContextParameter(window,"idNoCompleteCheck","1");
	<% /* DB End */ %>
	scScreenFunctions.SetContextParameter(window,"idMetaAction","CM100");
	
	<% /* SR 08/09/2004 : BMIDS815 */ %>
	if (m_sLTVChanged == '1')
		scScreenFunctions.SetContextParameter(window, "idCallingScreenID", "CM100");
	<% /* SR 08/09/2004 : BMIDS815 - End */ %>	
	
	frmToCM010.submit();
}
<% /* ASu - BMIDS00431 - End */ %>


// EP2_2168 - New Method - Check Drawdown < AmtReq
function frmScreen.txtDrawDown.onblur()
{
	CheckDrawDownValid();
}

// EP2_2168 - New Method - Check Drawdown < AmtReq
function CheckDrawDownValid()
{
	// Drawdown is present.
	if(frmScreen.txtDrawDown.value != "" && frmScreen.txtDrawDown.value != "0")
	{
		// Amount Requested is present.
		if(frmScreen.txtAmtRequested.value != "" && frmScreen.txtAmtRequested.value != "0")
		{
			// Is Drawdown > Amount Requested - This is a NoNo!
			if (parseInt(frmScreen.txtDrawDown.value) > parseInt(frmScreen.txtAmtRequested.value))
			{
				alert("The Drawdown amount cannot be greater than the amount requested");
				frmScreen.txtDrawDown.focus();
			}
		}
	}
}


<% /* BMIDS00654 MDC 04/11/2002 */ %>
function RunIncomeCalculations()
{
	var AllowableIncXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AllowableIncXML.CreateRequestTag(window,null);
	<% /* MAR30 */ %> 
	AllowableIncXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));		
		
	// Set up request for Income Calculation
	var TagRequest = AllowableIncXML.CreateActiveTag("INCOMECALCULATION");
	AllowableIncXML.SetAttribute("CALCULATEMAXBORROWING","1");
	AllowableIncXML.SetAttribute("QUICKQUOTEMAXBORROWING","0");
	AllowableIncXML.CreateActiveTag("APPLICATION");
	AllowableIncXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AllowableIncXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	AllowableIncXML.ActiveTag = TagRequest;
	
	var TagCUSTOMERLIST = AllowableIncXML.CreateActiveTag("CUSTOMERLIST");

	for(var nLoop = 1; nLoop <= 2; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		<% /* BM0457 The customer must exist */ %>
		if (sCustomerNumber.trim().length > 0)
		{
		<% /* BM0457 End */ %>
			var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
			var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);
			var sCustomerOrder = scScreenFunctions.GetContextParameter(window,"idCustomerOrder" + nLoop);

			<% /* BM0457
			if(sCustomerRoleType == "1")
			{ */ %>
			AllowableIncXML.CreateActiveTag("CUSTOMER");
			AllowableIncXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			AllowableIncXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
			AllowableIncXML.CreateTag("CUSTOMERROLETYPE", sCustomerRoleType);
			AllowableIncXML.CreateTag("CUSTOMERORDER", sCustomerOrder);
			AllowableIncXML.ActiveTag = TagCUSTOMERLIST;
		}
	}
	
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			AllowableIncXML.RunASP(document,"RunIncomeCalculations.asp");
			break;
		default: // Error
			AllowableIncXML.SetErrorResponse();
	}
	
	if(AllowableIncXML.IsResponseOK())
	{
		return(true);
	}

	return(false);
}
<% /* BMIDS00654 MDC 04/11/2002 - End */ %>

<% /* BMIDS815 Function converted from VBScript to JScript */ %>
function GetLocalPrinters()
{
	var strOut = "";
	var objOmPC = new ActiveXObject("omPC.PCAttributesBO");
	if (objOmPC != null)
	{
		var strXML = "<?xml version='1.0'?><REQUEST ACTION='CREATE'></REQUEST>";
		strOut = objOmPC.FindLocalPrinterList(strXML);
		objOmPC = null;
	}
	return strOut;
}
<% /* BMIDS815 End */ %>

<%/*MAR86 Start*/%>
function GetCounterOffer()
{
	var XML;
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateActiveTag("REQUEST");
	XML.SetAttribute("RB_TEMPLATE", "CM100Template");
	XML.CreateActiveTag("APPLICATION");
	XML.SetAttribute("_SCHEMA_","APPLICATION");
	XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFFNumber);
	
	//alert("m_sApplicationNumber " + m_sApplicationNumber);
	
	XML.RunASP(document,"GetDataFromRequestBroker.asp");

	if(XML.SelectTag(null, "APPLICATIONUNDERWRITING") != null)
		clientVar_CounterOffer = XML.GetAttribute("COUNTEROFFER");
}

function CheckCounterOffer()
{
	var returnVal=0;
	
	//alert("clientVar_CounterOffer  " + clientVar_CounterOffer);	
	//alert("frmScreen.txtAmtRequested.value  " + frmScreen.txtAmtRequested.value);
	if(clientVar_CounterOffer != "")
		if(parseInt(frmScreen.txtAmtRequested.value) > parseInt(clientVar_CounterOffer))
			returnVal = scScreenFunctions.DisplayClientError("The amount requestetd > Counter Offer amount " + clientVar_CounterOffer, "images/MSGBOX03.ICO");

	return returnVal;
}

function CheckCounterOfferOnSubmit()
{
	var returnVal=0;
	
	//alert("clientVar_CounterOffer  " + clientVar_CounterOffer);	
	//alert("frmScreen.txtAmtRequested.value  " + frmScreen.txtAmtRequested.value);
	if(clientVar_CounterOffer != "")
		if(parseInt(frmScreen.txtAmtRequested.value) > parseInt(clientVar_CounterOffer))
			returnVal = scScreenFunctions.DisplayClientWarning("The amount requestetd > Counter Offer amount " + clientVar_CounterOffer + ".\n Hit Exit button to Cancel." , "images/MSGBOX03.ICO","Ok","Exit");

	return returnVal;
}

<%/* EP2_8 - New function.*/%>
function CreateCM100ParamsXML()
{
	// Create doc basics
	CM100ParamsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	CM100ParamsXML.CreateActiveTag("PARAMS");
	
	// NATUREOFLOAN, CREDITSCHEME and APPLICATIONINCOMESTATUS flags.
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window);
	var xn = XML.XMLDocument.documentElement;
	xn.setAttribute("CRUD_OP","READ");
	xn.setAttribute("SCHEMA_NAME","omCRUD");
	xn.setAttribute("ENTITY_REF","APPLICATIONFACTFIND");
	var xe = XML.XMLDocument.createElement("APPLICATIONFACTFIND");
	xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	xe.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
	xn.appendChild(xe);
	XML.RunASP(document, "omCRUDIf.asp");
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null,"APPLICATIONFACTFIND");
		CM100ParamsXML.CreateTag("NATUREOFLOAN", XML.GetAttribute("NATUREOFLOAN"));
		CM100ParamsXML.CreateTag("APPLICATIONINCOMESTATUS", XML.GetAttribute("APPLICATIONINCOMESTATUS"));
		CM100ParamsXML.CreateTag("CREDITSCHEME", XML.GetAttribute("PRODUCTSCHEME"));
	}
	
	// CREDITSCHEME flag.
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.RunASP(document,"GetApplicationData.asp");
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null,"APPLICATION");
		var AccountGUID = XML.GetTagText("ACCOUNTGUID");
	}
	
	// GUARANTORIND flag.
	var IsThereAGuarantor = "0"
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"CustomerNumber" + nLoop);
		if (sCustomerNumber.trim().length > 0)
		{
			var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"CustomerRoleType" + nLoop);
			if(sCustomerRoleType == "2") // Guarantor
			{
				IsThereAGuarantor = "1";
				break;
			}
		}
		else
		break;   // No value => no more customers.
	}
	CM100ParamsXML.CreateTag("GUARANTORIND", IsThereAGuarantor);

	// FLEXIBLEPRODUCTS / NONFLEXIBLEPRODUCTS
	
	// Set defaults.
	CM100ParamsXML.CreateTag("FLEXIBLEPRODUCTS", "-1");
	CM100ParamsXML.CreateTag("NONFLEXIBLEPRODUCTS", "-1");

	// If we have existing components.
	if (m_iNumOfLoanComponents > 0) 
	{
		// Get first row and find ProductCode
		subQuoteXML.SelectTagListItem(0);
		var ProdCode = subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE");
		// Now get the FLEXIBLEMORTGAGEPRODUCT flag from MortgageProduct table.
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window);
		var xn = XML.XMLDocument.documentElement;
		xn.setAttribute("CRUD_OP","READ");
		xn.setAttribute("SCHEMA_NAME","epsomCRUD");
		xn.setAttribute("ENTITY_REF","MORTGAGEPRODUCT");
		var xe = XML.XMLDocument.createElement("MORTGAGEPRODUCT");
		xe.setAttribute("MORTGAGEPRODUCTCODE", ProdCode);
		xn.appendChild(xe);

		XML.RunASP(document, "omCRUDIf.asp");
		if (XML.IsResponseOK())
		{
			XML.SelectTag(null,"MORTGAGEPRODUCT");
			var FlexibleMortFlag = XML.GetAttribute("FLEXIBLEMORTGAGEPRODUCT")
			// Flexible Products flag.
			CM100ParamsXML.SetTagText("FLEXIBLEPRODUCTS", FlexibleMortFlag);
			// NONFLEXIBLEPRODUCTS Products flag.
			// Checked with Robin 25Oct06 - If Null then treat as non-flexible. 
			if (FlexibleMortFlag == "1")
			CM100ParamsXML.SetTagText("NONFLEXIBLEPRODUCTS", "0");
			else
			CM100ParamsXML.SetTagText("NONFLEXIBLEPRODUCTS", "1");
		}
	}	
	else // NO existing Loan Components
	{
		// If: Is additional lending (use Validation types)
		//	Then: Get Product code from MortgageLoan via application.AccountGUID)
		// Gather round data folks
		var sTypeOfMortgage = scScreenFunctions.GetContextParameter(window,"MORTGAGEAPPLICATIONVALUE");
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject(); 
		var blnValidationF = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('F'));
		var blnValidationM = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('M'));
		var blnValidationTOE = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE'));
		var blnValidationPSW = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW'));
	
		// Now check our conditions F&M or TOE or PSW
		if (((blnValidationF && blnValidationM) || blnValidationTOE || blnValidationPSW) && AccountGUID != "")
		{
			// Use AccountGUID from Application (retrieved above with CREDITSCHEME)
			// to retrieve the MORTGAGEPRODUCTCODE from MortgageLoan table.
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
				XML.SelectTag(null,"MORTGAGELOAN");
				var ProdCode = XML.GetAttribute("MORTGAGEPRODUCTCODE");
			}
		
			// Now get the FLEXIBLEMORTGAGEPRODUCT flag from MortgageProduct table.
			XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTag(window);
			var xn = XML.XMLDocument.documentElement;
			xn.setAttribute("CRUD_OP","READ");
			xn.setAttribute("SCHEMA_NAME","epsomCRUD");
			xn.setAttribute("ENTITY_REF","MORTGAGEPRODUCT");
			var xe = XML.XMLDocument.createElement("MORTGAGEPRODUCT");
			xe.setAttribute("MORTGAGEPRODUCTCODE", ProdCode);
			xn.appendChild(xe);

			XML.RunASP(document, "omCRUDIf.asp");
			if (XML.IsResponseOK())
			{
				XML.SelectTag(null,"MORTGAGEPRODUCT");
				var FlexibleMortFlag = XML.GetAttribute("FLEXIBLEMORTGAGEPRODUCT")
				// FLEXIBLEPRODUCTS / NONFLEXIBLEPRODUCTS flags.
				// Checked with RobinH 25Oct06 - If Null then treat as non-flexible. 
				if (FlexibleMortFlag == "1")
				{
					CM100ParamsXML.SetTagText("FLEXIBLEPRODUCTS", "1");
					CM100ParamsXML.SetTagText("NONFLEXIBLEPRODUCTS", "0");
				}
				else
				{
					CM100ParamsXML.SetTagText("FLEXIBLEPRODUCTS", "0");
					CM100ParamsXML.SetTagText("NONFLEXIBLEPRODUCTS", "1");
				}
			}	
		}
	}

}

<%/*MAR86 End*/%>

// EP2_55 - New Method.
function DisableAmountRequested()
{
	// Method to decide whether we are disabling AmountRequested field.
	var lDisable = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	
	// Check PSW
	var bDisablePSW = XML.GetGlobalParameterBoolean(document,"PSWDisableAmtRequested");
	var bIsPSWType = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW'));
	if((bDisablePSW == true) && (bIsPSWType == true))
		lDisable = true;

	// Check TOE
	var bDisableTOE = XML.GetGlobalParameterBoolean(document,"TOEDisableAmtRequested");
	var bIsTOEType = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE'));
	if((bDisableTOE == true) && (bIsTOEType == true))
		lDisable = true;
	
	return lDisable;	
	
}

// EP2_55 - New Method.
<% /* PSC 02/02/2007 EP2_1113 - Start */ %>
function ReturnIndicator(bIsPSWRetain, bIsManualPort)
{
	// Method to create indicator for each row.
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	
	// Check PSW
	var bIsPSWType = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW'));
	// Check TOE
	var bIsTOEType = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE'));
	// Check NP
	var bIsNPType = XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP'));
	
	if (bIsPSWType == true)
		if (bIsPSWRetain)
			return "R";
		else
			return "S";
	else if (bIsTOEType == true)
		return "T";
	else if (bIsNPType == true)
		if (bIsManualPort)
			return "P";
		else
			return "N";
	else
		return "N"		
}
<% /* PSC 02/02/2007 EP2_1113 - End */ %> 

// EP2_55 - New Method.
function frmScreen.btnGenerateQuote.onclick()
{
	<% /* PSC 02/02/2007 EP2_1113 - Start */ %>
	var TotalNetMonthlycost = frmScreen.txtTotalMnthCost.value
	
	if (m_iNumOfLoanComponents > 0 && (TotalNetMonthlycost == "" || TotalNetMonthlycost == 0))
	{
		alert("Delete existing loancomponents or calculate");
		return;
	}
	
	if (m_iNumOfLoanComponents > 0 && TotalNetMonthlycost != "" && TotalNetMonthlycost > 0)
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("QUOTATION");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
		XML.CreateTag("QUOTATIONNUMBER", scScreenFunctions.GetContextParameter(window,"idQuotationNumber","1"));
		XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
		XML.CreateTag("LIFESUBQUOTENUMBER", m_sLifeSubquoteNumber);
		XML.CreateTag("QUOTATIONTYPE", "2");
		XML.CreateTag("CREATESUBQUOTES", "0");

		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				XML.RunASP(document, "CreateNewMortgageLifeSubquote.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
		}

		if(XML.IsResponseOK()) 
		{	// Reset m_sMortgageSubquoteNumber before call.
				m_sMortgageSubquoteNumber = XML.GetTagText("MORTGAGESUBQUOTENUMBER")
				scScreenFunctions.SetContextParameter(window,"idMortgageSubQuoteNumber",m_sMortgageSubquoteNumber);
		}
	}
		
	CreateComponentsFromExistingAcc();
	
	<% /* PSC 12/02/2007 EP2_1314 */ %>
	scScreenFunctions.SetFieldToWritable(frmScreen, "txtPurchasePrice");

	// Now reset the form
	m_PopWinInd = 'Yes';
	Initialise();

	XML = null;	
	<% /* PSC 02/02/2007 EP2_1113 - End */ %> 
}

// EP2_55 - New Method.  
// Specced in EP2_56 Middle tier, but implemented using CRUD as per 'new code' dictat.
function CreateComponentsFromExistingAcc()
{
	//Set up variables
	var AccountGUID= "";		// Account GUID from Application table. 
	var MortList;               // List fo Mortgageloan details
	var bContinue = true;
	
	// Is our Application PSW, TOE or NP? Return if not.
	var sMortType;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
	
	if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW')))
		sMortType = "PSW";
	else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE')))
		sMortType = "TOE";
	else if(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP')))
		sMortType = "NP";
	else
	{
		alert("Mortgage Type is not valid for this operation. Operation cancelled");
		return ;
	}	
	// Get AccountGUID flag.
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.RunASP(document,"GetApplicationData.asp");
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null,"APPLICATION");
		AccountGUID = XML.GetTagText("ACCOUNTGUID");
	}

	// Now retrieve the Mortgageloan details
	// Use AccountGUID from Application to retrieve the MortgageLoan data.
	if (AccountGUID != "")
	{
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window);
		var xn = XML.XMLDocument.documentElement;
		xn.setAttribute("CRUD_OP","READ");
		xn.setAttribute("SCHEMA_NAME","epsomCRUD");
		<% /* PSC 02/02/2007 EP2_1271 */ %> 
		xn.setAttribute("ENTITY_REF","MORTGAGELOANANDPRODUCT");
		var xe = XML.XMLDocument.createElement("MORTGAGELOAN");
		xe.setAttribute("ACCOUNTGUID", AccountGUID);
		xn.appendChild(xe);

		XML.RunASP(document, "omCRUDIf.asp");
		if (XML.IsResponseOK())
		{
				<% /* PSC 07/02/2007 EP2_1271 - Start */ %> 
				var invalidLoan = XML.XMLDocument.selectSingleNode("RESPONSE/MORTGAGELOAN[not(MORTGAGEPRODUCT)]");
	
				if (invalidLoan != null)
				{
					var productCode = invalidLoan.getAttribute("MORTGAGEPRODUCTCODE");
					
					alert("Unable to create loan components. Mortgage loan has an invalid product code " + productCode);
					return;
				}
				<% /* PSC 07/02/2007 EP2_1271 - End */ %> 

				if (sMortType == 'PSW')
				{
					var sPSRValue = XML.GetComboIdForValidation("RedemptionStatus", "PSR", null, document);
					var sPSValue = XML.GetComboIdForValidation("RedemptionStatus", "PS", null, document);
					var strPattern = "RESPONSE/MORTGAGELOAN[@REDEMPTIONSTATUS='" + sPSRValue + "'or @REDEMPTIONSTATUS='" + sPSValue + "']";
					MortList = XML.XMLDocument.selectNodes(strPattern);
				}
				else if (sMortType == 'NP')
				{
					var sTBPValue = XML.GetComboIdForValidation("RedemptionStatus", "TBP", null, document);
					var strPattern = "RESPONSE/MORTGAGELOAN[@REDEMPTIONSTATUS='" + sTBPValue + "']";
					MortList = XML.XMLDocument.selectNodes(strPattern);
				}
				else if (sMortType == 'TOE')
				{
					var sTOEValue = XML.GetComboIdForValidation("RedemptionStatus", "TOE", null, document);
					var strPattern = "RESPONSE/MORTGAGELOAN[@REDEMPTIONSTATUS='" + sTOEValue + "']";
					MortList = XML.XMLDocument.selectNodes(strPattern);
				}
				
		}

		// Only continue if we've found Mortgage loan details
		if (MortList != null)
		{
			// Set variables
			var iMorts = MortList.length
			var TotalAmountRequested = 0;
			
			// Create new Loancomponents for each Mortgageloan.
			for(var nLoop = 0; nLoop < iMorts; nLoop++)
			{
				if(bContinue == true)
				{
					var CurrMortgageLoan = MortList.item(nLoop);
					var sRedemptionStatus = (CurrMortgageLoan.getAttribute("REDEMPTIONSTATUS")? CurrMortgageLoan.getAttribute("REDEMPTIONSTATUS"): 0) ;
					var gMortgageLoanGUID = (CurrMortgageLoan.getAttribute("MORTGAGELOANGUID")? CurrMortgageLoan.getAttribute("MORTGAGELOANGUID"): "") ;
					var OSBalance = parseFloat(CurrMortgageLoan.getAttribute("OUTSTANDINGBALANCE")? CurrMortgageLoan.getAttribute("OUTSTANDINGBALANCE"): 0) ;
					var AvailDisburse = parseFloat(CurrMortgageLoan.getAttribute("AVAILABLEFORDISBURSEMENT")? CurrMortgageLoan.getAttribute("AVAILABLEFORDISBURSEMENT"): 0) ;
					var OriginalLTV = (CurrMortgageLoan.getAttribute("ORIGINALLTV")? CurrMortgageLoan.getAttribute("ORIGINALLTV"): 0) ;
					var RepayType = (CurrMortgageLoan.getAttribute("REPAYMENTTYPE")? CurrMortgageLoan.getAttribute("REPAYMENTTYPE"): 0) ;
					var nOrigPandPIntOnly = (CurrMortgageLoan.getAttribute("ORIGINALPARTANDPARTINTONLYAMT")? CurrMortgageLoan.getAttribute("ORIGINALPARTANDPARTINTONLYAMT"): 0) ;
					var sMortgageProdCode = CurrMortgageLoan.getAttribute("MORTGAGEPRODUCTCODE") ;
					var dStartDate = CurrMortgageLoan.getAttribute("STARTDATE");
					
					<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
					var sLoanStartDate = CurrMortgageLoan.getAttribute("STARTDATE");
					var lYears = CurrMortgageLoan.getAttribute("ORIGINALTERMYEARS");
					if (lYears == "")
						lYears = "0";
					var lMonths = CurrMortgageLoan.getAttribute("ORIGINALTERMMONTHS");
					if (lMonths == "") 
						lMonths = "0";
					
					var l_sOutStandingTerm = CalculateOutstandingTerm(sLoanStartDate, lYears, lMonths);
					
					lYears = Math.floor(l_sOutStandingTerm /12);
					lMonths = l_sOutStandingTerm % 12;
					
					<% /* PSC 07/02/2007 EP2_1271 - Start */ %> 
					var mortgageProductNode = CurrMortgageLoan.selectSingleNode("MORTGAGEPRODUCT");
					var dProductStartDate = mortgageProductNode.getAttribute("STARTDATE");
					<% /* PSC 02/02/2007 EP2_1113 - End */ %>
					<% /* PSC 07/02/2007 EP2_1271 - End */ %> 

					var lXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					lXML.CreateRequestTag(window);
					var xn = lXML.XMLDocument.documentElement;
					xn.setAttribute("CRUD_OP","CREATE");
					xn.setAttribute("SCHEMA_NAME","omCRUD");
					xn.setAttribute("ENTITY_REF","LOANCOMPONENT");
					var xe = lXML.XMLDocument.createElement("LOANCOMPONENT");
					xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
					xe.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
					xe.setAttribute("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
					xe.setAttribute("MORTGAGELOANGUID", gMortgageLoanGUID);
	//				"LOANCOMPONENTSEQUENCENUMBER" is a Generated field. 
					xe.setAttribute("LOANAMOUNT",OSBalance );
					var TotalBalance = OSBalance + AvailDisburse;
					xe.setAttribute("TOTALLOANCOMPONENTAMOUNT", TotalBalance );
					// Add to Total amount
					TotalAmountRequested = TotalAmountRequested + TotalBalance;

	//				THIS NEEDS TO BE ADDED AFTER Column added to table.
	//				xe.setAttribute("DRAWDOWN",CurrMortgageLoan.getAttribute(""));
					xe.setAttribute("ORIGINALLTV", OriginalLTV);
					xe.setAttribute("REPAYMENTMETHOD", RepayType );
					
					<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
					xe.setAttribute("TERMINYEARS", lYears);
					xe.setAttribute("TERMINMONTHS", lMonths);
					<% /* PSC 02/02/2007 EP2_1113 - End */ %> 

					xe.setAttribute("PRODUCTCODESEARCHIND", 1 );

					// If Part and Part payment
					if (XML.IsInComboValidationList(document, 'RepaymentType', RepayType, Array('P')))
					{
						xe.setAttribute("INTERESTONLYELEMENT", nOrigPandPIntOnly );
						var CapLessInt = OSBalance - nOrigPandPIntOnly;
						xe.setAttribute("CAPITALANDINTERESTELEMENT", CapLessInt);
						xe.setAttribute("NETINTONLYELEMENT", nOrigPandPIntOnly );
						xe.setAttribute("NETCAPANDINTELEMENT", CapLessInt );
					}
					// Product Switch Retained
					if (XML.IsInComboValidationList(document, 'RedemptionStatus', sRedemptionStatus, Array('PSR')))
						xe.setAttribute("PRODUCTSWITCHRETAINPRODUCTIND", 1 );
					else 
						xe.setAttribute("PRODUCTSWITCHRETAINPRODUCTIND", 0 );

					// To be ported
					if (XML.IsInComboValidationList(document, 'RedemptionStatus', sRedemptionStatus, Array('TBP')))
						xe.setAttribute("MANUALPORTEDLOANIND", 1 );
					else 
						xe.setAttribute("MANUALPORTEDLOANIND", 0 );

					// Application Type
					if (sMortType == 'PSW')
					{
						var sPurposeOfLoan = XML.GetComboIdForValidation("PurposeOfLoanNew", "PSWD", null, document);
						xe.setAttribute("PURPOSEOFLOAN", sPurposeOfLoan);
					}
					else if (sMortType == 'NP')
					{
						var sPurposeOfLoan = XML.GetComboIdForValidation("PurposeOfLoanNew", "NPD", null, document);
						xe.setAttribute("PURPOSEOFLOAN", sPurposeOfLoan);
					}
					else if (sMortType == 'TOE')
					{
						var sPurposeOfLoan = XML.GetComboIdForValidation("PurposeOfLoanNew", "TOED", null, document);
						xe.setAttribute("PURPOSEOFLOAN", sPurposeOfLoan);
					}
					if (sMortgageProdCode != null)
						xe.setAttribute("MORTGAGEPRODUCTCODE", sMortgageProdCode );
					if (dStartDate != null)
						xe.setAttribute("STARTDATE", dStartDate );
						
					<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
					if (dProductStartDate != null)
						xe.setAttribute("STARTDATE", dProductStartDate);
					<% /* PSC 02/02/2007 EP2_1113 - End */ %> 

						
					// Append to Request
					xn.appendChild(xe);

					lXML.RunASP(document, "omCRUDIf.asp");
					if (!lXML.IsResponseOK())
					{	
						alert ("Error creating LoanComponent entries in table");
						bContinue = false;
					}

				} // End If bContinue == true

			}  // End nLoop < iMorts
			
			// Lastly Update the mortgage subquote Total
			if (bContinue == true)
			{
				<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
				frmScreen.txtAmtRequested.value = TotalAmountRequested;
				CalcAndPopulateLTV();
				<% /* PSC 02/02/2007 EP2_1113 - End */ %> 
				
				XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				XML.CreateRequestTag(window);
				var xn = XML.XMLDocument.documentElement;
				xn.setAttribute("CRUD_OP","UPDATE");
				xn.setAttribute("SCHEMA_NAME","omCRUD");
				xn.setAttribute("ENTITY_REF","MORTGAGESUBQUOTE");
				var xe = XML.XMLDocument.createElement("MORTGAGESUBQUOTE");
				xe.setAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
				xe.setAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber);
				xe.setAttribute("MORTGAGESUBQUOTENUMBER", m_sMortgageSubquoteNumber);
				xe.setAttribute("AMOUNTREQUESTED", TotalAmountRequested);
				
				<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
				xe.setAttribute("TOTALLOANAMOUNT", TotalAmountRequested);
				xe.setAttribute("LTV", frmScreen.txtLTVPercent.value);
				<% /* PSC 02/02/2007 EP2_1113 - End */ %> 

				xn.appendChild(xe);

				XML.RunASP(document, "omCRUDIf.asp");
				if (!XML.IsResponseOK())
					Alert ("Failed to update AmountRequested in MortgageSubquote table");
			} // End update subquote
		  
		} // End (MortList != null)
	} // End (AccountGUID != "")
	else
		alert("No existing Account details found. Operation cancelled");
	
	// Clear up the XML object.		
	XML = null;
}

<% /* PSC 02/02/2007 EP2_1113 - Start */ %> 
function ResetGenerateQuote()
{
	var tempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTypeOfMortgage =  scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", null);
		
	if (tempXML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW'))
			|| tempXML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE'))
			|| tempXML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP')))
			
	{
			var purchasePrice = frmScreen.txtPurchasePrice.value;
			var totalMonthlyCost = frmScreen.txtTotalMnthCost.value;
			
			if (m_iNumOfLoanComponents == 0 && purchasePrice != "" && purchasePrice > 0 &&
			    (totalMonthlyCost == "" || totalMonthlyCost == 0))
				frmScreen.btnGenerateQuote.disabled = false;
			else if (m_iNumOfLoanComponents > 0 && purchasePrice != "" && purchasePrice > 0 &&
			         (totalMonthlyCost == "" || totalMonthlyCost > 0))
				frmScreen.btnGenerateQuote.disabled = false;
			else
				frmScreen.btnGenerateQuote.disabled = true;
	}
	else
		frmScreen.btnGenerateQuote.disabled = true;
}

function CalculateOutstandingTerm(strStartDate, strOriginalTermYears, strOriginalTermMonths)
{
	var lOutstandingMonths = 0;
		
	var dtStartDate = scScreenFunctions.GetDateObjectFromString(strStartDate);
	var sOriginalTermYears = strOriginalTermYears;
	var sOriginalTermMonths = strOriginalTermMonths;
		
	if(dtStartDate != null)
	{
		var nStartYear = dtStartDate.getFullYear();
		var nStartMonth = dtStartDate.getMonth();
			
		<% /* Get the current date (but strip out the time info) */ %>
		<% /* MO - BMIDS00807 */ %>
		var dtTempCurrentDate = scScreenFunctions.GetAppServerDate(true);
		<% /* var dtTempCurrentDate	= new Date(); */ %>
		var dtCurrentDate		= new Date(dtTempCurrentDate.getFullYear(),
								           dtTempCurrentDate.getMonth(),
								           dtTempCurrentDate.getDate());

		var nCurrentYear = dtCurrentDate.getFullYear();
		var nCurrentMonth = dtCurrentDate.getMonth();

		if(dtCurrentDate > dtStartDate)
		{
			var nYearsPassed = nCurrentYear - nStartYear;
			var nMonthsPassed = 0;
				
			if(nStartMonth > nCurrentMonth)
			{
				nMonthsPassed = 12 - (nStartMonth - nCurrentMonth);
				nYearsPassed--;
			}
			else
				nMonthsPassed = nCurrentMonth - nStartMonth;

			if(sOriginalTermYears != "")
			{
				var nYearsToGo = sOriginalTermYears - nYearsPassed;
				var nMonthsToGo = 0;
					
				if(nMonthsPassed > sOriginalTermMonths)
				{
					nMonthsToGo = 12 - (nMonthsPassed - sOriginalTermMonths);
					nYearsToGo--;
				}
				else
					nMonthsToGo = sOriginalTermMonths - nMonthsPassed;
					
				if((nYearsToGo == 0 && nMonthsToGo > 0) || nYearsToGo > 0)
				{
					lOutstandingMonths = nYearsToGo * 12 + nMonthsToGo;
				}
			}
		}
	}
	
	return lOutstandingMonths;

}
<% /* PSC 02/02/2007 EP2_1113 - End */ %> 
-->
</script>
</body>
</html>


