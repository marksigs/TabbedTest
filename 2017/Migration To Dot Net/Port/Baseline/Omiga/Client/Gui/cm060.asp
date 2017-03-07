<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      cm060.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Quotation Summary Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		15/03/00	Initial revision
AY		29/03/00	New top menu/scScreenFunctions change
AY		12/04/00	SYS0328 - Dynamic currency display
MS		05/06/00	SYS0814 - modified MORTGAGECOVERFORAPPLICANT1 to MORTGAGECOVERFORAPPLICANT2
DP		01/09/00	SYS0814 - Changed DisplayPP to call GetTagFloat instead of GetTagText
CL		05/03/01	SYS1920 Read only functionality added
IK		22/03/01	SYS2145 Add idNoCompleteCheck
LD		23/05/02	SYS4727 Use cached versions of frame functions

BG		28/06/02	SYS4767 MSMS/Core integration
STB		28/02/02	SYS4144 Added APR column to Mortgage Details list.
JLD		16/04/02	MSMS0029	include ResolvedRate
JLD		14/05/02	MSMS0055	display rates as 2dp always.
BG		28/06/02	SYS4767 MSMS/Core integration - END	
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

DPF		30/07/02	BMIDS00224 - Add 'Cost at Final Rate' field
GD		14/08/02	BMIDS00320 - Rework to change displayed info for PP and BC summary details
GD		16/08/02	BMIDS00222 - Remove Life Cover displays
TW      09/10/02    Modified to incorporate client validation - SYS5115
MV		09/10/02	BMIDS00556	Removed Validation_Init()
DPF		15/10/02	BMIDS00046 - CPWP1 (BM020) Amended Mortgage list box to show only 4 loan components at 
					once and to be scrollable. Added one off cost analysis to screen (read only).
DPF		11/11/02	Applied fix to stop 'Record Not Found' error being displayed.
MO		13/11/2002	BMIDS00723	Made changes after the font sizes were changed
DPF		15/11/2002  BMIDS00516  Have added a tooltip to the Repayment method cell if P&P to display amounts
DPF		19/11/2002	BMIDS00777	Fixed bug to ensure the correct details for one off costs are being displayed
GD		22/11/2002	BMIDS00974	Ensure the table headings are displayed in Mortgage Table, and rework layout to accomodate.
GHun	26/11/2002	BMIDS00974	Adjusted table headings and layout so everything is visible
MDC		19/12/2002	BM0176 Use correct MortgageSubQuoteNumber to retrieve OneOffCosts
GD		16/05/2003	BM0368		Eliminate 'NaN' problems by using new method parseIntSafe()
INR		06/08/2003	BMIDS624 Change interest rate to resolved rate & resolved rate to Manual Adj%
GHun	27/10/2003	BMIDS624 Check if rates have changed on entry
DRC     17/05/2004  BMIDS767 Put AdHoc indicator in table
HMA     23/06/2004  BMIDS776 Change heading from 'Monthly Cost' to 'Cost' for Buildings and Contents.
DRC     25/06/2004  BMIDS763    Allow ProductSwitchFee Display
MC		07/07/2004	BMIDS763	isTIDType() function added inorder to check TID type in an array.
								ApplicationDate is read and passed with oneoffcalc calculation request.
SR		12/07/2004  BMIDS767
JD		28/07/2004	BMIDS749 Removed check for changed rates.
HMA     04/10/2004  BMIDS903    Set the CallingScreenID contex parameter when LTV has changed
HMA     21/04/2005  BMIDS977    Use correct mortgage subquote number for One Off Costs
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

PSC		15/02/2006	MAR1276 Use correct mortgage subquote number for One Off Costs
JD		09/03/2006	MAR1061 Add display of PurchasePrice, totalLoanAmount, AmountRequested and LTV
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:
DS		20/03/2007	EP2_1916	Added 'Refund' column to 'One Off Costs' list		
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
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>

<% /* Specify Forms Here */ %>

<form id="frmToPP010" method="post" action="PP010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM010" method="post" action="CM010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM065" method="post" action="CM065.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* DPF 15/10/2002 - Spans to hold navigation buttons for list boxes */ %>
<span id ="spnListScroll">
	<!--GD BMIDS0974<span style="LEFT:300px; POSITION: absolute; TOP:163px">-->
	<span style="LEFT:300px; POSITION: absolute; TOP:225px">
		<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT:0px; TOP: 0px; height:24; width:304" TYPE="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>
<span id ="spnListScroll">
	<!--GD BMIDS0974<span style="LEFT:100px; POSITION: absolute; TOP:363px">-->
	<span style="LEFT:30px; POSITION: absolute; TOP:425px">
		<object data="scTableListScroll.asp" id="scScrollTable2" style="LEFT:0px; TOP: 0px; height:24; width:304" TYPE="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen">
<div id="divStatus" style="TOP: 60px; LEFT: 10px; WIDTH: 604px; HEIGHT: 100px; POSITION: absolute " class="msgGroup">
	<table id="tblStatus" width="604px" height="100px" border="0" cellspacing="0" cellpadding="0">
		<tr align="center"><td id="colStatus" align="center" class="msgLabel" style="FONT-SIZE: 20px">Please Wait...</td></tr>
	</table>
</div>

<div id="divBackground" style="TOP: 60px; LEFT: 10px; WIDTH: 604px; POSITION: absolute">
	<div id="divMort" style="HEIGHT: 118px; WIDTH: 604px; DISPLAY: none">
		<div style="TOP: 0px; LEFT: 0px; HEIGHT: 45px; WIDTH: 604px" class="msgGroup">
			<span style="TOP: 4px; LEFT: 4px; POSITION: relative" class="msgLabel">
				<font face="MS Sans Serif" size="1">
					<strong>Mortgage Details</strong>
				</font>
			</span>
			<span style="TOP: 25px; LEFT: 4px; POSITION: absolute" class="msgLabel">
				Total Loan Amount
				<span style="TOP: 0px; LEFT: 100px; POSITION: absolute">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgLabel"></label>
					<input id="txtTotalLoanAmount" name="TotalLoanAmount" maxlength="10" style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
				</span>
			</span>
			<span style="TOP: 25px; LEFT: 180px; POSITION: absolute" class="msgLabel">
				Amount Requested
				<span style="TOP: 0px; LEFT: 100px; POSITION: absolute">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgLabel"></label>
					<input id="txtAmountRequested" name="AmountRequested" maxlength="10" style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
				</span>
			</span>
			<span style="TOP: 25px; LEFT: 360px; POSITION: absolute" class="msgLabel">
				Purchase Price
				<span style="TOP: 0px; LEFT: 80px; POSITION: absolute">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgLabel"></label>
					<input id="txtPurchasePrice" name="PurchasePrice" maxlength="10" style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
				</span>
			</span>
			<span style="TOP: 25px; LEFT: 520px; POSITION: absolute" class="msgLabel">
				LTV
				<span style="TOP: 0px; LEFT: 30px; POSITION: absolute">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgLabel"></label>
					<input id="txtLTV" name="LTV" maxlength="10" style="TOP: -3px; WIDTH: 50px; POSITION: ABSOLUTE" class="msgTxt">
				</span>
			</span>
		</div>
		<div style="TOP: 0px; LEFT: 0px; HEIGHT: 118px; WIDTH: 604px" class="msgGroup">
			<span style="TOP: 4px; LEFT: 4px; POSITION: relative" class="msgLabel">
				<font face="MS Sans Serif" size="1">
					<strong>Loan Details</strong>
				</font>
			</span>
			<!--BG		28/06/02	SYS4767 MSMS/Core integration-->
			<!--DPF		30/07/02	BMIDS00224 - add Cost at Final Rate field to table -->
			<!--DPF		15/10/02    BMIDS00046 - made table 4 rows -->
			<span style="TOP: 10px; LEFT: 4px; POSITION: relative">
				<table id="tblMortgage" width="596px" border="0" cellspacing="0" cellpadding="0" class="msgTable">
					<!--GD BMIDS00974
					<tr id="rowMortTitles">	 -->
					<tr id="rowTitles">
						<td width="9%"  class="TableHead">Prod.<br>No</td>
						<td width="8%"  class="TableHead">Prod.<br>Type</td>
						<td width="12%" class="TableHead">Product<br>Description</td>
						<td width="9%"  class="TableHead">Rate</td>
						<td width="9%"  class="TableHead">Manual<br>Adj %</td>
						<td width="9%"  class="TableHead">Loan<br>Amount</td>
						<td width="10%" class="TableHead">Term</td>		
						<td width="8%"  class="TableHead">Repay<br>Type</td>
						<td width="8%"  class="TableHead">APR</td>	
						<td width="9%"  class="TableHead">Monthly<br>Cost</td>
						<td width="9%"  class="TableHead">FinalRate<br>Cost</td>
					</tr>
					<tr id="rowMort01">		<td width="9%" class="TableTopLeft">&nbsp</td>		<td width="8%" class="TableTopCenter">&nbsp</td>	<td width="12%" class="TableTopCenter">&nbsp</td>		<td width="9%" class="TableTopCenter">&nbsp</td>	<td width="9%" class="TableTopCenter">&nbsp</td>	<td width="9%" class="TableTopCenter">&nbsp</td>		<td width="10%" class="TableTopCenter">&nbsp</td>	<td width="8%" class="TableTopCenter">&nbsp</td>	<td width="8%" class="TableTopCenter">&nbsp</td>	<td width="9%" class="TableTopCenter">&nbsp</td>	<td width="9%" class="TableTopRight">&nbsp</td></tr>
					<tr id="rowMort02">		<td width="9%" class="TableLeft">&nbsp</td>			<td width="8%" class="TableCenter">&nbsp</td>		<td width="12%" class="TableCenter">&nbsp</td>			<td width="9%" class="TableCenter">&nbsp</td>	    <td width="9%" class="TableCenter">&nbsp</td>		<td width="9%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>		<td width="8%" class="TableCenter">&nbsp</td>		<td width="8%" class="TableCenter">&nbsp</td>		<td width="9%" class="TableCenter">&nbsp</td>		<td width="9%" class="TableRight">&nbsp</td></tr>
					<tr id="rowMort03">		<td width="9%" class="TableLeft">&nbsp</td>			<td width="8%" class="TableCenter">&nbsp</td>		<td width="12%" class="TableCenter">&nbsp</td>			<td width="9%" class="TableCenter">&nbsp</td>		<td width="9%" class="TableCenter">&nbsp</td>		<td width="9%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>		<td width="8%" class="TableCenter">&nbsp</td>		<td width="8%" class="TableCenter">&nbsp</td>		<td width="9%" class="TableCenter">&nbsp</td>		<td width="9%" class="TableRight">&nbsp</td></tr>
					<tr id="rowMort04">		<td width="9%" class="TableBottomLeft">&nbsp</td>	<td width="8%" class="TableBottomCenter">&nbsp</td>	<td width="12%" class="TableBottomCenter">&nbsp</td>	<td width="9%" class="TableBottomCenter">&nbsp</td> <td width="9%" class="TableBottomCenter">&nbsp</td>	<td width="9%" class="TableBottomCenter">&nbsp</td>		<td width="10%" class="TableBottomCenter">&nbsp</td><td width="8%" class="TableBottomCenter">&nbsp</td>	<td width="8%" class="TableBottomCenter">&nbsp</td>	<td width="9%" class="TableBottomCenter">&nbsp</td>	<td width="9%" class="TableBottomRight">&nbsp</td></tr>				
				</table>
			</span>
		</div>
	</div>
	<!-- DPF 15/10/02 - BMIDS00046:  Added one off cost analysis to screen -->
	<div id="divOneOff" style="HEIGHT: 208px; WIDTH: 604px; DISPLAY: none">
		<div style="TOP: 0px; LEFT: 0px; HEIGHT: 208px; WIDTH: 604px" class="msgGroup">
			<!--GD BMIDS0974<span style="TOP: 4px; LEFT: 4px; POSITION: relative" class="msgLabel">-->
			<span style="TOP: 24px; LEFT: 4px; POSITION: relative" class="msgLabel">
				<font face="MS Sans Serif" size="1">
					<strong>One Off Costs</strong>
				</font>
			</span>
			<!--GD BMIDS0974<span id="spnOneOffCosts" style="LEFT: 4px; POSITION: relative; TOP: 10px">-->
			<span id="spnOneOffCosts" style="LEFT: 4px; POSITION: relative; TOP: 30px">
				<span id="fldOneOffCosts" style="LEFT: 260px; POSITION: relative; TOP: 10px">
					<span style="TOP: 10px; LEFT: 4px; POSITION: absolute" class="msgLabel">
						Total One-Off Costs
						<span style="TOP: 0px; LEFT: 165px; POSITION: absolute">
							<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
							<input id="txtTotalOneOffCosts" name="TotalOneOffCosts" maxlength="10" style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
						</span>
					</span>
					<span style="TOP: 35px; LEFT: 4px; POSITION: absolute" class="msgLabel">
						Deposit
						<span style="TOP: 0px; LEFT: 165px; POSITION: absolute">
							<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
							<input id="txtDeposit" name="Deposit" maxlength="10" style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
						</span>
					</span>
					<span style="TOP: 60px; LEFT: 4px; POSITION: absolute" class="msgLabel">
						One-Off Costs Added To Loan
						<span style="TOP: 0px; LEFT: 165px; POSITION: absolute">
							<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
							<input id="txtOneOffCostsAddedToLoan" name="OneOffCostsAddedToLoan" maxlength="10" style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
						</span>
					</span>
					<span style="TOP: 85px; LEFT:4px; POSITION: absolute" class="msgLabel">
						Total Incentives
						<span style="TOP: 0px; LEFT: 165px; POSITION: absolute">
							<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
							<input id="txtTotalIncentives" name="TotalIncentives" maxlength="10" style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
						</span>
					</span>
					<span style="TOP: 110px; LEFT: 4px; POSITION: absolute" class="msgLabel">
						Manual Incentives
						<span style="TOP: 0px; LEFT: 165px; POSITION: absolute">
							<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
							<input id="txtManualIncentives" name="ManualIncentives" maxlength="10" style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
						</span>
					</span>
					<span style="TOP: 135px; LEFT: 4px; POSITION: absolute" class="msgLabel">
						Total Net Cost
						<span style="TOP: 0px; LEFT: 165px; POSITION: absolute">
							<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
							<input id="txtTotalNetCost" name="TotalNetCost" maxlength="10" style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
						</span>
					</span>
				</span>
				<table id="tblOneOff" width="326" border="0" cellspacing="0" cellpadding="0" class="msgTable">
				<!--Begin: EP2_1916 - DS- Added Refund column -->
				<tr id="rowTitles">	
						<td width="40%" class="TableHead">Description</td>	
					    <td id="tdRefundAmount" width="20%" class="TableHead">Refund</td>
						<td id="tdAmount" width="20%" class="TableHead">Amount</td>
						<td id="tdAddToLoan" width="5%" class="TableHead">Add?</td>
						<td id="tdAdHocInd" width="5%" class="TableHead">AdHoc?</td>
				</tr>
				<!--End: EP2_1916 -->
					<!--<tr id="rowTitles">	
						<td width="50%" class="TableHead">Description</td>	
					<!-- DRC Reg027	<td width="20%" class="TableHead">Name</td> -->
					<!-- <td width="20%" class="TableHead"></td>
						<td id="tdAmount" width="20%" class="TableHead">Amount</td>
						<td id="tdAddToLoan" width="5%" class="TableHead">Add?</td>
						<td id="tdAdHocInd" class="TableHead">AdHoc?</td>
					</tr> -->
					<tr id="row01">	<td width="50%" class="TableTopLeft">&nbsp;</td> <td width="20%" class="TableTopCenter">&nbsp;</td>	<td width="20%" class="TableTopCenter">&nbsp;</td><td width="5%" class="TableTopCenter">&nbsp;</td><td class="TableTopRight">&nbsp;</td></tr>
					<tr id="row02">	<td width="50%" class="TableLeft">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>  <td width="20%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td>	 <td class="TableRight">&nbsp;</td></tr>
					<tr id="row03">	<td width="50%" class="TableLeft">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>  <td width="20%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td>   <td class="TableRight">&nbsp;</td></tr>
					<tr id="row04">	<td width="50%" class="TableLeft">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>  <td width="20%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td>	 <td class="TableRight">&nbsp;</td></tr>
					<tr id="row05">	<td width="50%" class="TableLeft">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>  <td width="20%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td>  <td class="TableRight">&nbsp;</td></tr>
					<tr id="row06">	<td width="50%" class="TableLeft">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>  <td width="20%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td>   <td class="TableRight">&nbsp;</td></tr>
					<tr id="row07">	<td width="50%" class="TableLeft">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td>   <td class="TableRight">&nbsp;</td></tr>
					<tr id="row08">	<td width="50%" class="TableLeft">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td>   <td class="TableRight">&nbsp;</td></tr>
					<tr id="row09">	<td width="50%" class="TableLeft">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td>	<td width="20%" class="TableCenter">&nbsp;</td><td width="5%" class="TableCenter">&nbsp;</td>   <td class="TableRight">&nbsp;</td></tr>
					<tr id="row10">	<td width="50%" class="TableBottomLeft">&nbsp;</td>	<td width="20%" class="TableBottomCenter">&nbsp;</td> <td width="20%" class="TableBottomCenter">&nbsp;</td> <td width="5%" class="TableBottomCenter">&nbsp;</td>	<td class="TableBottomRight">&nbsp;</td></tr>
					</table>
			</span>
		</div>
	</div>
<%/* END of Change for BMIDS00046 - have also commented out life cover code as not required
	
	<div id="divLife" style="HEIGHT: 138px; WIDTH: 604px; DISPLAY: none">
		<div style="TOP: 0px; LEFT: 0px; HEIGHT: 128px; WIDTH: 604px" class="msgGroup">
			<span style="TOP: 4px; LEFT: 4px; POSITION: relative" class="msgLabel">
				<font face="MS Sans Serif" size="1">
					<strong>Life Cover Details</strong>
				</font>
			</span>

			<span style="TOP: 10px; LEFT: 4px; POSITION: relative">
				<table id="tblLife" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable" style="PADDING-LEFT: 1px;PADDING-RIGHT: 1px">
					<tr id="rowLifeTitles">	<td width="50%" class="TableHead">Applicant(s)</td>	<td width="40%" class="TableHead">Benefit Type</td>		<td class="TableHead">Monthly Cost</td></tr>
					<tr id="rowLife01">		<td width="50%" class="TableTopLeft">&nbsp</td>		<td width="40%" class="TableTopCenter">&nbsp</td>		<td class="TableTopRight">&nbsp</td></tr>
					<tr id="rowLife02">		<td width="50%" class="TableLeft">&nbsp</td>		<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="rowLife03">		<td width="50%" class="TableLeft">&nbsp</td>		<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="rowLife04">		<td width="50%" class="TableLeft">&nbsp</td>		<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="rowLife05">		<td width="50%" class="TableBottomLeft">&nbsp</td>	<td width="40%" class="TableBottomCenter">&nbsp</td>	<td class="TableBottomRight">&nbsp</td></tr>
				</table>
			</span>
		</div>
	</div>
*/ %>
	<div id="divBC" style="HEIGHT: 82px; WIDTH: 604px; DISPLAY: none">
		<!--<div id="divBCGroup" style="TOP: 0px; LEFT: 0px; HEIGHT: 72px; WIDTH: 604px" class="msgGroup">-->
		<div id="divBCGroup" style="TOP: 0px; LEFT: 0px; HEIGHT: 82px; WIDTH: 604px" class="msgGroup">
			<!--GD BMIDS0974<span style="TOP: 4px; LEFT: 4px; POSITION: relative" class="msgLabel">-->
			<span style="TOP: 24px; LEFT: 4px; POSITION: relative" class="msgLabel">
				<font face="MS Sans Serif" size="1">
					<strong>Buildings &amp; Contents Details</strong>
				</font>
			</span>

			<!--GD BMIDS0974<span style="TOP: 10px; LEFT: 4px; POSITION: relative">-->
			<span style="TOP: 30px; LEFT: 4px; POSITION: relative">
				<table id="tblBC" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable" style="PADDING-LEFT: 1px;PADDING-RIGHT: 1px">
					<tr id="rowBCTitles">	<td width="40%" class="TableHead">Cover Type</td><td width="40%" class="TableHead">Reference Number</td><td class="TableHead">Cost</td></tr>
					<tr id="rowBC01">		<td width="40%" class="TableOneRowLeft">&nbsp</td><td width="40%" class="TableOneRowCenter">&nbsp</td><td class="TableOneRowRight">&nbsp</td></tr>
				</table>
			</span>

			<div id="divBCComment" style="TOP: 15px; LEFT: 4px; WIDTH: 596px; POSITION: relative; VISIBILITY: hidden" class="msgLabel">
				* Buildings &amp; Contents is paid annually but has been included as a monthly figure in the list above
			</div>
		</div>
	</div>
	<div id="divPP" style="HEIGHT: 82px; WIDTH: 604px; DISPLAY: none">
		<div style="TOP: 0px; LEFT: 0px; HEIGHT: 72px; WIDTH: 604px" class="msgGroup">
			<span style="TOP: 4px; LEFT: 4px; POSITION: relative" class="msgLabel">
				<font face="MS Sans Serif" size="1">
					<strong>Payment Protection Details</strong>
				</font>
			</span>

			<span style="TOP: 10px; LEFT: 4px; POSITION: relative">
				<table id="tblPP" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable" style="PADDING-LEFT: 1px;PADDING-RIGHT: 1px">
					<tr id="rowPPTitles">	<td width="40%" class="TableHead">Cover Type</td>	<td width="40%" class="TableHead">Reference Number</td>		<td class="TableHead">Monthly Cost</td></tr>
					<tr id="rowPP01">		<td width="40%" class="TableOneRowLeft">&nbsp</td>		<td width="40%" class="TableOneRowCenter">&nbsp</td>		<td class="TableOneRowRight">&nbsp</td></tr>
<%/*					<!--
					<tr id="rowPP02">		<td width="40%" class="TableLeft">&nbsp</td>		<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="rowPP03">		<td width="40%" class="TableLeft">&nbsp</td>		<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="rowPP04">		<td width="40%" class="TableLeft">&nbsp</td>		<td width="40%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
					<tr id="rowPP05">		<td width="40%" class="TableBottomLeft">&nbsp</td>	<td width="40%" class="TableBottomCenter">&nbsp</td>	<td class="TableBottomRight">&nbsp</td></tr>
					--> */ %>
				</table>
			</span>
		</div>
	</div>

	<div id="divTotal" style="HEIGHT: 40px; WIDTH: 604px; DISPLAY: none">
		<div style="TOP: 0px; LEFT: 0px; HEIGHT: 30px; WIDTH: 604px" class="msgGroup">
			<span style="TOP: 8px; LEFT: 420px; POSITION: relative" class="msgLabel">
				Total Cost
				<span style="TOP: 0px; LEFT: 92px; POSITION: ABSOLUTE">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
					<input id="txtTotalCost" maxlength="10" style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
				</span>
			</span>
		</div>
	</div>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 450px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cm060Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_blnReadOnly = false;
var m_iTableLength = 4; // DPF 15/10/02 - max rows for Mortgage Details table
var m_iTable2Length = 10; // DPF 15/10/02 - max rows for One Off Costs table
var m_MortgageSubQuoteNumber = 0;     // BMIDS977
var m_sApplicationNumber;             // BMIDS977
var m_sApplicationFactFindNumber;     // BMIDS977

var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
var XMLOneOffCosts = new top.frames[1].document.all.scXMLFunctions.XMLObject();
var XMLComboList = new top.frames[1].document.all.scXMLFunctions.XMLObject();		


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

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Quotation Summary","CM060",scScreenFunctions);

	scScreenFunctions.SetCurrency(window,frmScreen);
	scScreenFunctions.SetFieldState(frmScreen,"txtTotalCost","R");
	PopulateScreen();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CM060");
//	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
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

function PopulateScreen()
{
<%	// Get the Quotation Summary details
	// (The quotation number is passed in via the XML context variable)
	// Display the appropriate sections
%>	var dTotalCost = 0.0;
	<% /* BMIDS624 */ %>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","");
	<% /* BMIDS624 End */ %>
	<% /* PSC 15/02/2006 MAR1276 */ %>
	m_MortgageSubQuoteNumber = scScreenFunctions.GetContextParameter(window,"idXML2","");

	XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("QUOTATIONNUMBER",scScreenFunctions.GetContextParameter(window,"idXML",""));
	XML.RunASP(document,"GetQuotationSummary.asp");
	divStatus.style.display = "none";

	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = XML.CheckResponse(sErrorArray);
	if(sResponseArray[0])
	{
		//DPF 15/10/02 - CPWP1: Changed the way the mortgage list box works
		dTotalCost += ShowMortgageList(0);
		divOneOff.style.display = "Block";
		//GD BMIDS00222 dTotalCost += DisplayLife(XML);
		dTotalCost += DisplayBC(XML);
		dTotalCost += DisplayPP(XML);
		
		<% /*MARS1061 populate Mortgage Details*/ %>
		XML.SelectTag(null,"QUOTATIONSUMMARY");
		frmScreen.txtTotalLoanAmount.value = XML.GetTagText("TOTALLOANAMOUNT");
		frmScreen.txtAmountRequested.value = XML.GetTagText("AMOUNTREQUESTED");
		frmScreen.txtPurchasePrice.value = XML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE");
		frmScreen.txtLTV.value = XML.GetTagText("LTV");
		<% /*make the fields readonly */ %>
		scScreenFunctions.SetFieldState(frmScreen,"txtTotalLoanAmount","R");
		scScreenFunctions.SetFieldState(frmScreen,"txtAmountRequested","R");
		scScreenFunctions.SetFieldState(frmScreen,"txtPurchasePrice","R");
		scScreenFunctions.SetFieldState(frmScreen,"txtLTV","R");
	}
	
	GetOneOffCosts(); // DPF 16/10/02 - call new function to retrieve one off cost details
		
	DisplayEnd(dTotalCost);
	
	<% /* BMIDS624 */ %>
	// JD BMIDS749 removed   CheckForChangedRates(sApplicationNumber, sApplicationFactFindNumber);
}

// DPF 16/10/02 - BMIDS00046 -  Series of functions (mainly nicked from CM130) to retrieve one off cost details
function GetOneOffCosts()
{
	if (GetComboList())
	{
		<% /*BMIDS977  Refresh the MortgageSubQuoteNumber in case a remodel has been performed. */ %>
		<% /* PSC 15/02/2006 MAR1276 */ %>
		<% /* GetMortgageSubQuoteNumber(); */ %>
		
		XMLOneOffCosts.CreateRequestTag(window,"SEARCH");
		XMLOneOffCosts.CreateActiveTag("ONEOFFCOST");
		XMLOneOffCosts.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);                  // BMIDS977
		XMLOneOffCosts.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);  // BMIDS977
		XMLOneOffCosts.CreateTag("MORTGAGESUBQUOTENUMBER",m_MortgageSubQuoteNumber);         // BMIDS977
		XMLOneOffCosts.RunASP(document, "GetOneOffCostsDetails.asp");
		
		<%/* A record not found error is valid */%>
		var sErrorArray = new Array("RECORDNOTFOUND");
		var sResponseArray = XMLOneOffCosts.CheckResponse(sErrorArray);
		var sNewLoanValue;
		var sApplicationType;
		var iNumberOfValidCosts = 0 ; <% /* SR 12/07/2004 : BMIDS767 */ %>
			
		if (sResponseArray[0] == true || 
			sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
		
		//if (XMLOneOffCosts.IsResponseOK() == true)
		{
			scScreenFunctions.SetFieldState(frmScreen, "txtTotalOneOffCosts","R");
			scScreenFunctions.SetFieldState(frmScreen, "txtDeposit","R");
			scScreenFunctions.SetFieldState(frmScreen, "txtOneOffCostsAddedToLoan","R");
			scScreenFunctions.SetFieldState(frmScreen, "txtTotalIncentives","R");
			scScreenFunctions.SetFieldState(frmScreen, "txtTotalNetCost","R");
			scScreenFunctions.SetFieldState(frmScreen, "txtManualIncentives", "R");
			
			if (sResponseArray[0] == true)
			{
				XMLOneOffCosts.ActiveTag = null;
				XMLOneOffCosts.SelectTag(null,"MORTGAGESUBQUOTE");
				
				// BM0133
				// Check the Application Type held in context.
				// Only display the DEPOSIT if the type is New Loan (validation type 'N')
	
				// Get the value associated with Validation Type 'N'
				sNewLoanValue = XMLOneOffCosts.GetComboIdForValidation("TypeOfMortgage", "N", null, document);
	
				//Get Application type from context
				sApplicationType = scScreenFunctions.GetContextParameter(window, "idMortgageApplicationValue", null);
		
				if (sApplicationType == sNewLoanValue)
				{
					frmScreen.txtDeposit.value = XMLOneOffCosts.GetTagText("DEPOSIT");
				}		
				else
				{	
					frmScreen.txtDeposit.value = "0";
				}		
				// BM0133  End						
				
				frmScreen.txtManualIncentives.value = XMLOneOffCosts.GetTagText("MANUALINCENTIVEAMOUNT");
				
				XMLOneOffCosts.ActiveTag = null;
				var tagCOSTSLIST = XMLOneOffCosts.CreateTagList("MORTGAGEONEOFFCOST");
				var iNumberOfOneOffCosts = XMLOneOffCosts.ActiveTagList.length;
				var m_iTotalOneOffCosts = 0;
				var m_iOneOffCostsAddedToLoan = 0;
		
				for (var nCost = 0; nCost < iNumberOfOneOffCosts; nCost++)
				{
					XMLOneOffCosts.ActiveTagList = tagCOSTSLIST;
					if (XMLOneOffCosts.SelectTagListItem(nCost) == true)
					{
						if(XMLOneOffCosts.GetTagText("AMOUNT") != "" && XMLOneOffCosts.GetTagText("AMOUNT") != "0")
						{	
							var iAmount = parseIntSafe(XMLOneOffCosts.GetTagText("AMOUNT"));
							var sRetArray = GetComboValues(XMLOneOffCosts.GetTagText("MORTGAGEONEOFFCOSTTYPE"));
							// DRC BMIDS767 - check all validation types for valueid
							var sAddedToLoan = GetAddedToLoanString(sRetArray, XMLOneOffCosts.GetTagText("ADDTOLOAN"));
							<% /* SR 12/07/2004 : BMIDS767  */ %>
							if(!isTIDType(sRetArray,1))
							{
								m_iTotalOneOffCosts += iAmount;
								iNumberOfValidCosts += 1;
							}
							<% /* SR 12/07/2004 : BMIDS767 - End */ %>
							
							if(sAddedToLoan == "Yes")
								m_iOneOffCostsAddedToLoan += iAmount;
						}
					}
				}
		
				frmScreen.txtTotalOneOffCosts.value = m_iTotalOneOffCosts;
				frmScreen.txtOneOffCostsAddedToLoan.value = m_iOneOffCostsAddedToLoan;
				
				XMLOneOffCosts.SelectTag(null, "RESPONSE");
				frmScreen.txtTotalIncentives.value = XMLOneOffCosts.GetTagText("TOTALINCENTIVES");
				
				if(frmScreen.txtManualIncentives.value != "0" && frmScreen.txtManualIncentives.value != "")
				{
					m_iManualIncentives = frmScreen.txtManualIncentives.value;
					var iTotalNetCost = m_iTotalOneOffCosts + parseIntSafe(frmScreen.txtDeposit.value) 
										- m_iOneOffCostsAddedToLoan - parseIntSafe(m_iManualIncentives);
				}	
				else
				{
					var iTotalNetCost = m_iTotalOneOffCosts + parseIntSafe(frmScreen.txtDeposit.value) 
									    - m_iOneOffCostsAddedToLoan - parseIntSafe(frmScreen.txtTotalIncentives.value);
				}
		
				frmScreen.txtTotalNetCost.value = iTotalNetCost;
				<% /* SR 12/07/2004 : BMIDS767
				var iNumberOfValidCosts;
				iNumberOfValidCosts = GetNumberOfValidCosts(); */ %>
				scScrollTable2.initialiseTable(tblOneOff, 0, "", DisplayCosts, m_iTable2Length, iNumberOfValidCosts);
				if(iNumberOfValidCosts > 0)DisplayCosts(0);

			}
		}
	}


}
function GetNumberOfValidCosts()
{
	var iNumberOfValidCosts = 0;
	
	XMLOneOffCosts.ActiveTag = null;
	XMLOneOffCosts.CreateTagList("MORTGAGEONEOFFCOST");
	for(var iLoop = 0; iLoop < XMLOneOffCosts.ActiveTagList.length; iLoop++)
	{
		if(XMLOneOffCosts.SelectTagListItem(iLoop) == true)
		{
			if(XMLOneOffCosts.GetTagText("AMOUNT") != "" && XMLOneOffCosts.GetTagText("AMOUNT") != "0")
				iNumberOfValidCosts++;		
		}
	}

	return(iNumberOfValidCosts);
}

function DisplayCosts(iStart)
{
	scScrollTable2.clear();
	XMLOneOffCosts.ActiveTag = null;
	var tagCOSTSLIST = XMLOneOffCosts.CreateTagList("MORTGAGEONEOFFCOST");
	var iNumberOfOneOffCosts = XMLOneOffCosts.ActiveTagList.length;
	var iRowAdded = 0;
	var iValidRow = 0;
	for (var nCost = 0; nCost < iNumberOfOneOffCosts && iRowAdded < m_iTable2Length; nCost++)
	{
		XMLOneOffCosts.ActiveTagList = tagCOSTSLIST;
		if (XMLOneOffCosts.SelectTagListItem(nCost) == true)
		{
			if(XMLOneOffCosts.GetTagText("AMOUNT") != "" && XMLOneOffCosts.GetTagText("AMOUNT") != "0")
			{
				if(iValidRow == iStart)
				{
					<% /* SR 09/07/2004 : BMIDS767
					iRowAdded++;
					var sName = XMLOneOffCosts.GetTagText("DESCRIPTION");
					if (sName == "NULL")sName=""; */ %>
					var sAmount = XMLOneOffCosts.GetTagText("AMOUNT");
					<% /*Begin: EP2_1916 - DS- Added Refund column*/ %>
					var iRefundAmount = parseIntSafe(XMLOneOffCosts.GetTagText("REFUNDAMOUNT"));
					<% /*End: EP2_1916 */ %>
					var sValueId = XMLOneOffCosts.GetTagText("MORTGAGEONEOFFCOSTTYPE");
					//  DRC BMIDS 767 - start
					var sAdHocInd = XMLOneOffCosts.GetTagText("ADHOCIND");
					if (sAdHocInd == "1" ) 
					  sAdHocInd = "Yes";
					else
					  sAdHocInd = "";
					//  DRC BMIDS 767 - end  
					var sRetArray = GetComboValues(sValueId);
					//if (sRetArray[1] != "TID")
					if(!isTIDType(sRetArray,1))
					{
						iRowAdded++; <% /* SR 09/07/2004 : BMIDS767 */ %>
					    //  DRC BMIDS 767 - check all validation types
						var sAddedToLoan = GetAddedToLoanString(sRetArray, XMLOneOffCosts.GetTagText("ADDTOLOAN"));
						<% /*Begin: EP2_1916 - DS- Added Refund column*/ %>
						ShowRow(iRowAdded, sRetArray[0], iRefundAmount,sAmount ,sAddedToLoan, sAdHocInd, nCost);	
						<% /*End: EP2_1916 */ %>				
					}
				}
				else iValidRow++;
			}
		}
	}
}

function isTIDType(sArray,iStart)
{
	var bResult = new Boolean();
	bResult=false;
	
	for(i=0;i<sArray.length;i++)
	{
		if(sArray[i] == "TID")
		{
			bResult=true;
		}
	}
	
	return bResult;
}

function ShowRow(iRow, sCostDescription, sName, sAmount, sAddedToLoan, sAdHocInd, nAttribute)
{
	scScreenFunctions.SizeTextToField(tblOneOff.rows(iRow).cells(0),sCostDescription);
	scScreenFunctions.SizeTextToField(tblOneOff.rows(iRow).cells(1),sName);
	scScreenFunctions.SizeTextToField(tblOneOff.rows(iRow).cells(2),sAmount);
	scScreenFunctions.SizeTextToField(tblOneOff.rows(iRow).cells(3),sAddedToLoan);
	scScreenFunctions.SizeTextToField(tblOneOff.rows(iRow).cells(4),sAdHocInd);
	tblOneOff.rows(iRow).setAttribute("tagListItem", nAttribute);
	if(sAddedToLoan != "No" && sAddedToLoan != "Yes")
		tblOneOff.rows(iRow).setAttribute("AddAllowed", false);
	else tblOneOff.rows(iRow).setAttribute("AddAllowed", true);
}

function GetComboList()
{
	var sGroups = new Array("OneOffCost");			
	var bSuccess = true;
			
	if (XMLComboList.GetComboLists(document, sGroups) != true)
	{				
		bSuccess = false;
		alert("Unable to read the combo records for OneOffCost group");
	}
			
	return bSuccess;			
}
function GetComboValues(sValueId)
{
	var sRetArray = new Array;
	XMLComboList.SelectTag(null, "LISTNAME");
	var tagLISTENTRY = XMLComboList.CreateTagList("LISTENTRY");
	var iNumberOfEntries = XMLComboList.ActiveTagList.length;
	
	var bFound = false;
	for (var nEntry = 0; nEntry < iNumberOfEntries && bFound == false; nEntry++)
	{
		XMLComboList.ActiveTagList = tagLISTENTRY;
		if (XMLComboList.SelectTagListItem(nEntry) == true)
		{
			if(XMLComboList.GetTagText("VALUEID") == sValueId)
			{
				bFound = true;
				sRetArray[0] = XMLComboList.GetTagText("VALUENAME");
				var tagVALIDATIONLIST = XMLComboList.CreateTagList("VALIDATIONTYPE");
				for(var i = 0; i < XMLComboList.ActiveTagList.length; i++)
				{
				//DRC BMIDS767 change to get all validationtypes
					sRetArray[i+1] = XMLComboList.GetTagListItem(i).text;
				}
			}
		}
	}
	return sRetArray;
}

function GetAddedToLoanString(sValidationArray, sAddToLoan)
{
	//  DRC BMIDS 767 To check for all validationtypes for a valueid, need an array, not just the 
	//  first value 
	var sAddedToLoan = "";
	var sValidationType = "";
	var bFound = false;
	var tagMORTGAGELENDER = XMLOneOffCosts.SelectTag(null, "MORTGAGELENDER");
	for (var nEntry = 1; nEntry < sValidationArray.length && bFound == false; nEntry++)
	{
	    sValidationType = sValidationArray[nEntry];
	
		if( (sValidationType == "ADM" && XMLOneOffCosts.GetTagText("ALLOWADMINFEEADDED") == "1") ||
			(sValidationType == "ARR" && XMLOneOffCosts.GetTagText("ALLOWARRGEMTFEEADDED") == "1") ||
			(sValidationType == "DEE" && XMLOneOffCosts.GetTagText("ALLOWDEEDSRELFEEADD") == "1") ||
			(sValidationType == "MIG" && XMLOneOffCosts.GetTagText("ALLOWMIGFEEADDED") == "1") ||
			(sValidationType == "POR" && XMLOneOffCosts.GetTagText("ALLOWPORTINGFEEADDED") == "1") ||
			(sValidationType == "REI" && XMLOneOffCosts.GetTagText("ALLOWREINSPTFEEADDED") == "1") ||
			(sValidationType == "REV" && XMLOneOffCosts.GetTagText("ALLOWREVALUATIONFEEADDED") == "1") ||
			(sValidationType == "SEA" && XMLOneOffCosts.GetTagText("ALLOWSEALINGFEEADDED") == "1") ||
			(sValidationType == "TTF" && XMLOneOffCosts.GetTagText("ALLOWTTFEEADDED") == "1") ||
			(sValidationType == "VAL" && XMLOneOffCosts.GetTagText("ALLOWVALNFEEADDED") == "1") ||
			(sValidationType == "LEG" && XMLOneOffCosts.GetTagText("ALLOWLEGALFEEADDED") == "1") ||
			(sValidationType == "PSF" && XMLOneOffCosts.GetTagText("ALLOWPRODUCTSWITCHFEEADDED") == "1") ||
			(sValidationType == "OTH" && XMLOneOffCosts.GetTagText("ALLOWOTHERFEEADDED") == "1")     
		  )
		{   bFound = true;
			if(sAddToLoan == "1")
				sAddedToLoan="Yes";
			else sAddedToLoan="No";
		}
	}
		
	return(sAddedToLoan);
}
//DPF 16/10/02 - BMIDS00046 - END OF ONE OFF COSTS CHANGE


//DPF 15/10/02 - CPWP1: Changed the way the mortgage list box works
function ShowMortgageList(nStart)
{
	var dTotalCost = 0.0;
	var icount;
	XML.ActiveTag = null;
	XML.SelectTag(null, "QUOTATIONSUMMARY");
	XML.CreateTagList("LOANCOMPONENT");
	var iNumberofComponents = XML.ActiveTagList.length;
	
	
	
	for(icount = 0;icount < XML.ActiveTagList.length;icount++)
	{
		XML.SelectTagListItem(icount);
		dTotalCost += XML.GetTagFloat("NETMONTHLYCOST");
	}
	
	if(iNumberofComponents > 0)
	{
		divMort.style.display = "block";
		scScrollTable.initialiseTable(tblMortgage, 0, "", DisplayMortgage, m_iTableLength, iNumberofComponents);
		DisplayMortgage(nStart);
		return dTotalCost;
	}
}

function DisplayMortgage(nStart)
{
	var dTotal = 0.0;
	var icount;
	var sPPBreakdown = "";
	scScrollTable.clear();
	
	for(icount = 0;icount < XML.ActiveTagList.length && icount < m_iTableLength;icount++)
	{
		XML.ActiveTag = null;
		XML.SelectTag(null,"QUOTATIONSUMMARY");
		XML.CreateTagList("LOANCOMPONENT");
		XML.SelectTagListItem(icount + nStart);
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(0),XML.GetTagText("MORTGAGEPRODUCTCODE"));
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(1),XML.GetTagText("RATETYPE"));
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(2),XML.GetTagText("PRODUCTNAME"));
		
		//BG		28/06/02	SYS4767 MSMS/Core integration
		<% /* INR		06/08/03	BMIDS624  	Change interest rate to resolved rate & resolved rate to Manual Adj%
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(3),top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("RATE"), 2));
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(4),top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("RESOLVEDRATE"), 2));
		*/ %>
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(3),top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("RESOLVEDRATE"), 2));
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(4),top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("MANUALADJUSTMENTPERCENT"), 2));
		<% /* BMIDS624 End */ %>
		
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(5),XML.GetTagText("TOTALLOANCOMPONENTAMOUNT"));

		//scScreenFunctions.SizeTextToField(tblMortgage.rows(NLoop).cells(3),XML.GetTagText("RATE"));
		//scScreenFunctions.SizeTextToField(tblMortgage.rows(NLoop).cells(4),XML.GetTagText("TOTALLOANCOMPONENTAMOUNT"));
		//BG		28/06/02	SYS4767 MSMS/Core integration - END	
		
		var sTerm = XML.GetTagText("TERMINYEARS") + " yrs " + XML.GetTagText("TERMINMONTHS") + " mths";
		
		//BG		28/06/02	SYS4767 MSMS/Core integration
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(6),sTerm);
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(7),GetAbbrevRepayMethod(XML.GetTagText("REPAYMENTMETHOD")));

		//DPF 15/11/2002 - BMIDS00516 - Have added Tool Tip to repayment method (if P&P)
		if (XML.GetTagText("REPAYMENTMETHOD") == '3')
		{
			sPPBreakdown = "C&I=" + XML.GetTagText("CAPITALANDINTERESTELEMENT") + " - I/O=" + XML.GetTagText("INTERESTONLYELEMENT")
			tblMortgage.rows(icount+1).cells(7).title = sPPBreakdown ;
		}
		//END OF DPF

		//scScreenFunctions.SizeTextToField(tblMortgage.rows(NLoop).cells(5),sTerm);
		//scScreenFunctions.SizeTextToField(tblMortgage.rows(NLoop).cells(6),XML.GetTagAttribute("REPAYMENTMETHOD","TEXT"));
		//BG		28/06/02	SYS4767 MSMS/Core integration - END
		
		//BG		28/06/02	SYS4767 MSMS/Core integration
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(8),top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("APR"),2));
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(9),top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("NETMONTHLYCOST"),2));		
		//BMIDS00224	DPF - 30/07/02 add in field for Monthly Cost at Final Rate
		scScreenFunctions.SizeTextToField(tblMortgage.rows(icount+1).cells(10),top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("FINALRATEMONTHLYCOST"),2));
		//scScreenFunctions.SizeTextToField(tblMortgage.rows(NLoop).cells(7),top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("NETMONTHLYCOST"),2));
		//BG		28/06/02	SYS4767 MSMS/Core integration - END
	}
}
// END OF CHANGE FOR CPWP1

//BG		28/06/02	SYS4767 MSMS/Core integration
function GetAbbrevRepayMethod(sRepayMethod)
{
<%	/* to condense the information in the listbox, output an abbreviated string */
%>	var sAbbrev;
	switch(sRepayMethod)
	{
		case "1" : sAbbrev = "I/O"; break;
		case "2" : sAbbrev = "C&I"; break;
		case "3" : sAbbrev = "P&P"; break;
		default : sAbbrev = "??"; break;
	}
	return sAbbrev;
}
//BG		28/06/02	SYS4767 MSMS/Core integration
<%/*
function DisplayLife(XML)
{
	var dTotal = 0.0;

	XML.SelectTag(null,"QUOTATIONSUMMARY");
	XML.CreateTagList("LIFEBENEFIT");
	if(XML.ActiveTagList.length > 0)
	{
		if (XML.GetTagText("CUSTOMERNUMBER1") != "")
		{
			divLife.style.display = "block";
			for(var NLoop = 1;XML.SelectTagListItem(NLoop-1) && NLoop <= 5;NLoop++)
			{
				var sCustomerNumber1 = XML.GetTagText("CUSTOMERNUMBER1");
				var sCustomerNumber2 = XML.GetTagText("CUSTOMERNUMBER2");
				var sCustomerName = "";
				if(sCustomerNumber1 != "")
				{
					 sCustomerName = scScreenFunctions.GetContextCustomerName(window,sCustomerNumber1);
					if(sCustomerNumber2 != "")
					{
						if(sCustomerName != "") sCustomerName += " and ";
						sCustomerName += scScreenFunctions.GetContextCustomerName(window,sCustomerNumber2);
					}
					scScreenFunctions.SizeTextToField(tblLife.rows(NLoop).cells(0),sCustomerName);

					scScreenFunctions.SizeTextToField(tblLife.rows(NLoop).cells(1),XML.GetTagAttribute("BENEFITTYPE","TEXT"));

					dTotal += XML.GetTagFloat("MONTHLYCOST");
					scScreenFunctions.SizeTextToField(tblLife.rows(NLoop).cells(2),top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("MONTHLYCOST"),2));
				}
			}
		}
	}
	return dTotal
}
*/%>

function DisplayBC(XML)
{
	var dTotal = 0.0;

	XML.SelectTag(null,"QUOTATIONSUMMARY");
	var sBCCoverType = XML.GetTagAttribute("BCCOVERTYPE","TEXT");
	var sBCReferenceNumber = XML.GetTagText("BCREFERENCENUMBER");
	if(sBCCoverType != "")
	{
		divBC.style.display = "block";
		scScreenFunctions.SizeTextToField(tblBC.rows(1).cells(0),sBCCoverType);
		scScreenFunctions.SizeTextToField(tblBC.rows(1).cells(1),sBCReferenceNumber);
		//var sAccBuildCover = XML.GetTagText("ACCIDENTALBUILDCOVERREQUIRED");
		//if(sAccBuildCover == "1") sAccBuildCover = "Yes";
		//else if(sAccBuildCover == "0") sAccBuildCover = "No";
		//else sAccBuildCover = " ";
		//scScreenFunctions.SizeTextToField(tblBC.rows(1).cells(1),sAccBuildCover);

		//var sAccContCover = XML.GetTagText("ACCIDENTALCONTENTCOVERREQUIRED");
		//if(sAccContCover == "1") sAccContCover = "Yes";
		//else if(sAccContCover == "0") sAccContCover = "No";
		//else sAccContCover = " ";
		//scScreenFunctions.SizeTextToField(tblBC.rows(1).cells(2),sAccContCover);

		dTotal += XML.GetTagFloat("TOTALBCMONTHLYCOST");

<%		// Where there is a premium present, check if the BC payment is to be annual.
		// If it is a piece of text has to be displayed and we need to alter the sizes
		// of the BC divs
		// N.B. The TOTALBCMONTHLYCOST comes from the BC subquote record whereas the
		// BC payment frequency comes from the BC details record
%>		var sTotal = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALBCMONTHLYCOST"),2);
		if(XML.GetTagText("BCFREQUENCY") == "2" && sTotal != "")
		{
			sTotal += "*";
			divBCComment.style.visibility = "visible";
			var nNewHeight = divBC.style.posHeight + 20;
			divBC.style.height = nNewHeight + "px";
			nNewHeight = divBCGroup.style.posHeight + 20;
			divBCGroup.style.height = nNewHeight + "px";
		}
		//scScreenFunctions.SizeTextToField(tblBC.rows(1).cells(3),sTotal);
		scScreenFunctions.SizeTextToField(tblBC.rows(1).cells(2),sTotal);
	}

	return dTotal;
}

function DisplayPP(XML)
{
	var dTotal = 0.0;

	XML.SelectTag(null,"QUOTATIONSUMMARY");
	var sPPCoverType = XML.GetTagAttribute("PPCOVERTYPE","TEXT");
	var sPPReferenceNumber = XML.GetTagText("PPREFERENCENUMBER");
	if(sPPCoverType != "")
	{
		divPP.style.display = "block";

<%		// N.B. This processing will be incorrect if the applicant orders are ever changed
%>		//var sCustomerName = "";
		//if(XML.GetTagFloat("MORTGAGECOVERFORAPPLICANT1") != 0) 
			//sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName1","");

		//if(XML.GetTagFloat("MORTGAGECOVERFORAPPLICANT2") != 0)
		//{
		//	if(sCustomerName != "") sCustomerName += " and ";
		//	sCustomerName += scScreenFunctions.GetContextParameter(window,"idCustomerName2","");
		//}
		scScreenFunctions.SizeTextToField(tblPP.rows(1).cells(0),sPPCoverType);
		scScreenFunctions.SizeTextToField(tblPP.rows(1).cells(1),sPPReferenceNumber);

		dTotal += XML.GetTagFloat("TOTALPPMONTHLYCOST");
		scScreenFunctions.SizeTextToField(tblPP.rows(1).cells(2),top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetTagText("TOTALPPMONTHLYCOST"),2));
	}
		
	return dTotal;
}

function DisplayEnd(dTotalCost)
{
<%	// Calculate whether the OK button needs to move from its standard position by
	// checking which divs are visible and totalling their size
	// Also display the total monthly cost at this point, but only if there are details, otherwise
	// show the no details found message
%>	var nTotal = 60;
	if(divMort.style.display != "none") nTotal += divMort.style.posHeight;
	//if(divLife.style.display != "none") nTotal += divLife.style.posHeight;
	if(divOneOff.style.display !="none") nTotal += divOneOff.style.posHeight;
	if(divBC.style.display != "none") nTotal += divBC.style.posHeight;
	if(divPP.style.display != "none") nTotal += divPP.style.posHeight;

	if(nTotal > 60)
	{
		divTotal.style.display = "block";
		frmScreen.txtTotalCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(dTotalCost,2);
		nTotal += divTotal.style.posHeight;
	}

	if(nTotal > msgButtons.style.posTop) msgButtons.style.top = nTotal + "px";

	if(nTotal == 60)
	{
		divStatus.style.display = "block";
		colStatus.innerText = "No Quotation Details Found";
	}
}

function btnSubmit.onclick()
{
	var sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	if(sMetaAction == "fromCM065")
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","fromCM060");
		frmToCM065.submit();
	}
	else 
	{
	// DRC BMIDS767 - add route back to PP010
	if (sMetaAction == "fromPP010")
	  {
		scScreenFunctions.SetContextParameter(window,"idMetaAction",null);
		frmToPP010.submit();
	  }

	else
	  {
		scScreenFunctions.SetContextParameter(window,"idMetaAction",null);
		scScreenFunctions.SetContextParameter(window,"idNoCompleteCheck","1");
		frmToCM010.submit();
	  }
	}
	<% /* BMIDS903 Set the CallingScreenID context parameter for use by CM100 */ %>
	if (scScreenFunctions.GetContextParameter(window,"idLTVChanged") == '1')
	{
		scScreenFunctions.SetContextParameter(window, "idCallingScreenID", "CM060");
	}		
}
<%/* GD BM0368 Added  End */%>
function parseIntSafe(sText)
{
	if (sText == "") return(0)
	else
	return(parseInt(sText));
}
<%/* GD BM0368 Added  End */%>

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
			alert("Rates on this product are not consistent, please remodel.");
	}
}
<% /* BMIDS624 End */ %>

<% /* BMISD977  Add new function to refresh MortgageSubQuoteNumber */ %>
function GetMortgageSubQuoteNumber()
{
 var AppQuoteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
		
	// Setup Application data
	AppQuoteXML.CreateRequestTag(window , "GETACCEPTEDQUOTEDATA");
	AppQuoteXML.CreateActiveTag("APPLICATION");
	AppQuoteXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
	AppQuoteXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber );
	
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			AppQuoteXML.RunASP(document,"AQGetAcceptedQuoteData.asp");
			break;
		default: // Error
			AppQuoteXML.SetErrorResponse();
		}

   if(AppQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE") != null) 
   {
		if(AppQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER")!= "")
		{
			m_MortgageSubQuoteNumber  = AppQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER");
			return;
	    }		
   }
   else
   {
		m_MortgageSubQuoteNumber = scScreenFunctions.GetContextParameter(window,"idXML2","");
   }
}
-->
</script>
</body>
</html>
