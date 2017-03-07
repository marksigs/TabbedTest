<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
	<% /*
Workfile:      cm110.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Add/Edit loan components screen  THIS IS A POP-UP SCREEN
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		16/02/2000	Initial creation
JLD		28/03/00	SYS0465 - on Search, check mandatory fields.
JLD		31/03/00	Altered size of CM120 calling screen
AY		04/04/00	scScreenFunctions change
JLD		10/04/00	SYS0467 - Added Product Number to list box.
JLD		10/04/00	SYS0567 - Check global parameter
AY		13/04/00	SYS0328 - Dynamic currency display
JLD		17/04/00	Change to interface with CM150
JLD		19/04/00	Changes for CM150 processing
BG		18/05/00	SYS0752 Removed ToolTips
IW		24/05/00	SYS0774 DISTRIBUTIONCHANNELID S/B/ CHANNELID
PSC		31/05/00    SYS0763 Page products correctly
PSC     31/05/00	SYS0539 Disable Details, Incentives and NonPanelDetails based on
                    product selected. Remove interest rate from Non Panel products
PSC     08/06/00	Backout SYS0774 and correct search context
PSC     20/06/00	SYS0763 - Amend to use the scPageScroll scriplet
BG		20/09/00	SYS1438 - fix multiple lender problems.  
SR		18/06/01	SYS2377
DRC     06/07/01                SYS2358 - fix bug in onchange events of the loan parameters
GD		27/07/01	SYS2530 - check for Add or Edit mode in btnOK.onclick(), plus changed
					match for tag text from "Part & Part" to "Part and Part"
GHun	23/04/02	SYS1051	Buttons incorrectly enabled when the product table is blank
JLD		29/04/02	SYS4488 ensure P&P split is entered before searching for products.
JLD		29/04/02	SYS4489 ensure loan amount is > 0 before searching for products.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specifc History:

Prog	Date		AQR			Description
MV		11/06/2002	BMIDS00032	Modified PoupulateScreen - Defaults the New Loancomponents
AW		10/06/02	BMIDS00013	Display Rate Type/Rate Amount
GD		19/06/2002	BMIDS00077  Upgrade to Core 7.0.2
MV		18/07/2002	BMIDS00179  Core Upgrade Rollback Modified FindProducts();
DPF		09/07/2002	BMIDS00044  CMWP3  1.  add new combo box - Sub-purpose of loan
										2.  add yes /no question for ported loan components
										3.  As part of above changes altered position settings of various objects.
										4.  Also added lines to PopulateScreen / SaveComponentDetails / FindProducts
										5.  Added onChange event for Ported Loan field to clear products table
DPF		13/08/2002	BMIDS00307  Default 'Loan to Be Ported' to No.							
MV		23/08/2002	BMIDS00355	IE 5.5 Upgrade - Modified the display Height of the CM120.asp
DPF		30/09/2002	BMIDS00046	Modified call to FindProducts to included FlexibleProduct indicator if not first product 
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
DPF		17/10/2002	BMIDS00046	Extra array arguement to send to CM150.asp (SubQuoteXML)
SA		06/11/2002	BMIDS00656	In FindProducts, reset Flexible Products indicator if it's loan component 1 we are editing
MDC		15/11/2002	BMIDS00939  CC015 Ported Products
SA		15/11/2002	BMIDS00656 Product Search - send in Loan Amount & Loan Component Amount
MDC		22/11/2002	BMIDS01071 
MDC		17/12/2002	BM0150 - Check for empty or null value in OnChange event for RepaymentType
MV		10/02/2003	BM0337 - Amended frmScreen.btnSearch.onClick();
MV		18/02/2003	BM0235		Amended PopulateScreen()
MC		19/04/2004	BMIDS517	Popup Dialog height increased by 10px (white space padded to the title text)
								CM120 (further filtering,OneOffcosts etc.,)
JD		26/07/04	BMIDS816	ApplicationDate is required in CM150, passed thru' from CM100
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specifc History:
JJ		30/09/2005 MAR58		cboSubPurposeOfLoan textbox and label removed.
GHun	12/10/2005	MAR46	Added LoanComponentToBeRetained
Maha T	17/11/2005	MAR350	LoanComponentToBeRetained  to be displayed only if Application Type is (PSW, Product Switch)
DRC     09/03/2006  MAR1380 - Need to put in a critical data check to see if the term has changed 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
db Specifc History:
DRC		27/04/2006  EP430     MAR58 reversed cboSubPurposeOfLoan textbox and label reinstated
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specifc History:
SW		21/06/2006  EP771	  Added cboRepaymentVehicle and functionality
PB      19/06/2006  EP1089	  Merge MAR1833 - Correct searching after Further Filtering.
AShaw	24/10/2006	EP2_8	  Incorporate spec changes to Initialise and SearchForMortgageProducts code.
							  Extra load params (Add/Edit and new search params).	
AShaw	17/11/2006	EP2_55	  New methods when changing Ported/Retained options.
							  Initialise method changing.	
AShaw	13/12/2006	EP2_55	  Revisited to add more changes for EP2_55
		18/12/2006	EP2_55	  Altered Ported Option code. (Left code in place, but now does nothing onclick).
MAH		29/12/2006	EP2_444	  Made changes to Repayment vehicle and added monthly cost
IK		14/01/3007	EP2_760   Load relevant PurposeOfLoan combo
PSC		30/01/2007	EP2_1100  Use remaining term rather than outstanding term
PSC		30/01/2007	EP2_1110  Calculate remaining term in edit mode too
PSC		1302/2007	EP2_1314  Correct enabling and disabling of fields
INR		19/02/2007	EP2_1476	PurposeOfLoan combo entries can have multiple validation types
IK		27/02/2007	EP2_1694  default further filtering to 'All Products With Checks'							
LDM		27/02/2007	EP2_1567  Remove porting questionif type of app is not Porting
SR		04/03/2007	EP2_1644	increased the width of scScrollPlus control
MAH		08/03/2007	EP2_1751	Added scMathFunctions to Assist rendering decimal places
MAH		09/03/2007	EP2_1750	Disabled Repayment type/vehicle selectivity when Product Switch
AW		15/03/2007	EP2_1968	Corrected definition of undefined bShowLoanCompToBeRetainedQuestion
PSC		23/03/2007	EP2_1622	Set first interest rate correctly for fixed rates
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
	<head>
		<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4" />
		<link href="stylesheet.css" rel="stylesheet" type="text/css" />
		<title>Loan Components <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
	</head>
	<body> <!-- Form Manager Object - Controls Soft Coded Field Attributes -->
		<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
		<object data="scMathFunctions.asp" height="1px" id="scMathFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			type="text/x-scriptlet" width="1" viewastext tabIndex="-1">
		</object>
		<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			type="text/x-scriptlet" width="1" viewastext tabIndex="-1">
		</object>
		<script src="validation.js" language="JScript" type="text/javascript"></script>
		<!-- List Scroll object ( and table I believe) -->
		<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px"
			tabIndex="-1" type="text/x-scriptlet" viewastext>
		</object>
		<span style="LEFT: 302px; POSITION: absolute; TOP: 380px">
			<object data="scPageScroll.htm" id="scScrollPlus" style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px"
				tabIndex="-1" type="text/x-scriptlet" viewastext>
			</object>
		</span><!-- Specify Screen Layout Here -->
		<form id="frmScreen" mark validate="onchange"> <!--style="VISIBILITY: hidden"> -->
			<div id="divSearch" style="HEIGHT: 160px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px"
				class="msgGroup">
				<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
					Purpose of Loan
					<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
						<select id="cboPurposeOfLoan" style="WIDTH: 190px" class="msgCombo">
						</select>
					</span>
				</span>
				<% // CMWP3 - DPF Jul 02 - added subpurpose of loan %>
				<span id="SubPurposeOfLoan" style="LEFT: 295px; POSITION: absolute; TOP: 10px" class="msgLabel">
					Sub-purpose of Loan
					<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
						<select id="cboSubPurposeOfLoan" style="WIDTH: 190px" class="msgCombo">
						</select>
					</span>
				</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
					Loan Amount
					<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
						<input id="txtLoanAmount" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
					</span>
				</span>
				<span style="LEFT: 295px; POSITION: absolute; TOP: 40px" class="msgLabel">
					Repayment Type
					<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
						<select id="cboRepaymentType" style="WIDTH: 190px" class="msgCombo">
						</select>
					</span>
				</span>
				<span style="LEFT: 295px; POSITION: absolute; TOP: 70px" class="msgLabel">
					Repayment Vehicle
					<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
						<select id="cboRepaymentVehicle" style="WIDTH: 190px" class="msgCombo" NAME="cboRepaymentVehicle">
						</select>
					</span>
				</span>
				<%/*EP2_444*/%>
				<span style="LEFT: 295px; POSITION: absolute; TOP: 100px" class="msglabel">
					Repayment Vehicle<br>Monthly Cost
					<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
						<input id="txtRepaymentVehicleMonthlyCost" NAME="txtRepaymentVehicleMonthlyCost" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
					</span>
				</span>
				<span style="LEFT: 210px; POSITION: absolute; TOP: 35px">
					<input id="btnSplit" value="Split" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
					Term
					<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
						<input id="txtTermYears" maxlength="3" style="POSITION: absolute; WIDTH: 40px" class="msgTxt">
						<span style="LEFT: 48px; POSITION: absolute; TOP: 2px" class="msgLabel">Years</span>
					</span>
					<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
						<input id="txtTermMonths" maxlength="3" style="POSITION: absolute; WIDTH: 40px" class="msgTxt">
						<span style="LEFT: 48px; POSITION: absolute; TOP: 2px" class="msgLabel">Months</span>
					</span>		
				</span>
				<span style="TOP: 100px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel" id="spnLoanComponentToBePorted">
					Loan Component to be Ported?
					<span style="TOP: -3px; LEFT: 180px; WIDTH: 45px; POSITION: ABSOLUTE">
						<input id="optLoanComponentToBePortedYes" name="LoanComponentToBePortedInd" type="radio"
							value="1">
						<label for="optLoanComponentToBePortedYes" class="msgLabel">Yes</label>
					</span> 
					<span style="TOP: -3px; LEFT: 230px; WIDTH: 45px; POSITION: ABSOLUTE">
						<input id="optLoanComponentToBePortedNo" name="LoanComponentToBePortedInd" type="radio"
							value="0" checked>
						<label for="optLoanComponentToBePortedNo" class="msgLabel">No</label>
					</span> 
				</span>
				<span style="TOP: 125px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel" id="spnLoanComponentToBeRetained">
					Loan Component to be Retained?
					<span style="TOP: -3px; LEFT: 180px; WIDTH: 45px; POSITION: ABSOLUTE">
						<input id="optLoanComponentToBeRetainedYes" name="LoanComponentToBeRetainedInd" type="radio"
							value="1">
						<label for="optLoanComponentToBeRetainedYes" class="msgLabel">Yes</label>
					</span> 
					<span style="TOP: -3px; LEFT: 230px; WIDTH: 45px; POSITION: ABSOLUTE">
						<input id="optLoanComponentToBeRetainedNo" name="LoanComponentToBeRetainedInd" type="radio"
							value="0" checked>
						<label for="optLoanComponentToBeRetainedNo" class="msgLabel">No</label>
					</span> 
				</span>
				<span style="LEFT: 430px; POSITION: absolute; TOP: 130px">
					<input id="btnSearch" value="Search" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
				<span style="LEFT: 500px; POSITION: absolute; TOP: 130px">
					<input id="btnFurtherFiltering" value="Further Filtering" type="button" style="WIDTH: 100px"
						class="msgButton">
				</span>
			</div>
			<div id="divResults" style="HEIGHT: 245px; LEFT: 10px; POSITION: absolute; TOP: 175px; WIDTH: 604px"
				class="msgGroup">
				<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
					<!--GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2<font face="MS Sans Serif" size="1">-->
					<strong>Available Mortgage Products</strong> 
					<!--</font>-->
				</span>
				<div id="spnTables" style="LEFT: 4px; POSITION: absolute; TOP: 30px">
					<div id="spnMultipleTable" style="LEFT: 0px; POSITION: absolute; TOP: 0px">
						<table id="tblMultipleTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
							<tr id="rowTitles">
								<td width="20%" class="TableHead">Lender</td>
								<td width="10%" class="TableHead">Product Number</td>
								<td width="30%" class="TableHead">Product Name</td>
								<td width="20%" class="TableHead">First Interest Rate</td>
								<td width="10%" class="TableHead">Rate Type</td>
								<td class="TableHead">Rate Period</td>
							</tr>
							<tr id="row01">
								<td width="20%" class="TableTopLeft">&nbsp;</td>
								<td width="10%" class="TableTopCenter">&nbsp;</td>
								<td width="30%" class="TableTopCenter">&nbsp;</td>
								<td width="20%" class="TableTopCenter">&nbsp;</td>
								<td width="10%" class="TableTopCenter">&nbsp;</td>
								<td class="TableTopRight">&nbsp;</td>
							</tr>
							<tr id="row02">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row03">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row04">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row05">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row06">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row07">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row08">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row09">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row10">
								<td width="20%" class="TableBottomLeft">&nbsp;</td>
								<td width="10%" class="TableBottomCenter">&nbsp;</td>
								<td width="30%" class="TableBottomCenter">&nbsp;</td>
								<td width="20%" class="TableBottomCenter">&nbsp;</td>
								<td width="10%" class="TableBottomCenter">&nbsp;</td>
								<td class="TableBottomRight">&nbsp;</td>
							</tr>
						</table>
					</div>
					<div id="spnSingleTable" style="LEFT: 0px; POSITION: absolute; TOP: 0px">
						<table id="tblSingleTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
							<tr id="rowTitles">
								<td width="10%" class="TableHead">Product Number</td>
								<td width="45%" class="TableHead">Product Name</td>
								<td width="15%" class="TableHead">First<BR>
									Interest Rate</td>
								<td width="15%" class="TableHead">Rate Type</td>
								<td class="TableHead">Rate Period</td>
							</tr>
							<tr id="row01">
								<td width="10%" class="TableTopLeft">&nbsp;</td>
								<td width="45%" class="TableTopCenter">&nbsp;</td>
								<td width="15%" class="TableTopCenter">&nbsp;</td>
								<td width="15%" class="TableTopCenter">&nbsp;</td>
								<td class="TableTopRight">&nbsp;</td>
							</tr>
							<tr id="row02">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row03">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row04">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row05">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row06">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row07">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row08">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row09">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="row10">
								<td width="10%" class="TableBottomLeft">&nbsp;</td>
								<td width="45%" class="TableBottomCenter">&nbsp;</td>
								<td width="15%" class="TableBottomCenter">&nbsp;</td>
								<td width="15%" class="TableBottomCenter">&nbsp;</td>
								<td class="TableBottomRight">&nbsp;</td>
							</tr>
						</table>
					</div>
				</div>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 250px">
					<input id="btnDetails" value="Details" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
				<span style="LEFT: 70px; POSITION: absolute; TOP: 250px">
					<input id="btnIncentives" value="Incentives" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
				<!--GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2<span style="LEFT: 140px; POSITION: absolute; TOP: 215px">
		<input id="btnNonPanelDetails" value="Non-Panel Details" type="button" style="WIDTH: 100px" class="msgButton">
	</span>-->
				<span style="LEFT: 4px; POSITION: absolute; TOP: 280px">
					<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
					<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">
						<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
					</span>
				</span>
			</div>
		</form>
		<!-- File containing field attributes (remove if not required) -->
		<!--  #include FILE="attribs/cm110attribs.asp" -->
		<!-- Specify Code Here -->
		<script language="JScript" type="text/javascript">
<!--
var m_iTableLength = 10;
var m_tblProductTable = null;
var m_sSubQuoteXML = "";
var m_subQuoteXML = null;
var m_NewLoanCompXML = null;
var m_ProductXML = null;
var m_FurtherFilteringXML = null;
var m_sReadOnly = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sApplicationDate = ""; // JD BMIDS816
var m_sMortgageSubQuoteNumber = "";
var m_sApplicationMode = "";
var m_sUserId = "";
var m_sUserType = "";
var m_sUnitId = "";
var m_sActivityId = ""; // DRC MAR1380
var	m_sStageId = ""; // DRC MAR1380
var	m_sApplicationPriority = ""; // DRC MAR1380
var m_sLoanComponentSeqNum = "";
var m_sidMortgageApplicationValue = "";
var m_sNumberOfLoanComponents = "";
var m_sDistributionChannelId = "";
var m_sCurrency = "";
var m_sRequestAttributes = "";
var m_sINTERESTONLYAMOUNT = "";
var m_sCAPITALINTERESTAMOUNT = "";
var m_sORIGINALINCENTIVESXML = null;
var m_sNEWINCENTIVESXML = "";
var m_sORIGINALPRODUCTCODE = "";
var m_sNEWPRODUCTCODE = "";
var m_sORIGINALPRODUCTSTARTDATE = "";
var m_sNEWPRODUCTSTARTDATE = "";
var m_sMultipleLender = "0";
var scScreenFunctions;
var m_ListEmpty = true;
var m_BaseNonPopupWindow = null;
//CMWP3 - DPF 23/07/02 - New variable for Loan Indicator
var m_LoanIndicator = 0;
var sPurposeListName = "";
var m_bFurtherFiltering = false;    <% /* EP1089/MAR1833 */ %>
<% /* EP2_8 Oct06 4 - New variables */ %>
var m_sAdditionalTerm = "0";		
var m_AddOrEdit = "";				
var m_sCM100XML = "";				
var m_CM100XML = null;	
var m_CalculatedFiltering = "";			
var m_bDataRead = false; <%/*EP2_444*/%>
<% /* PSC 30/01/2007 EP2_1100 */ %>
var m_iNumOfLoanComponents = 0;			
var m_bShowLoanCompToBeRetainedQuestion;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();%>
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2scScreenFunctions.ShowCollection(frmScreen);%>
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	scScreenFunctions.ShowCollection(frmScreen);
	m_sApplicationMode = sArgArray[0];
	m_sUserId = sArgArray[1];
	m_sUserType = sArgArray[2];
	m_sUnitId = sArgArray[3];
	m_sReadOnly = sArgArray[4];
	m_sSubQuoteXML = sArgArray[5];
	m_sLoanComponentSeqNum = sArgArray[6];
	m_sidMortgageApplicationValue = sArgArray[7];
	m_sNumberOfLoanComponents = sArgArray[8];
	m_sDistributionChannelId = sArgArray[9];
	m_sRequestAttributes = sArgArray[10];
	m_sCurrency = sArgArray[11];
	m_sApplicationDate = sArgArray[12]; // JD BMIDS816
	m_sActivityId = sArgArray[13]; // DRC MAR1380
	m_sStageId = sArgArray[14]; // DRC MAR1380
	m_sApplicationPriority = sArgArray[15]; // DRC MAR1380
	m_AddOrEdit = sArgArray[16]; // AS EP2_8 19Oct06
	m_sCM100XML = sArgArray[17]; // AS EP2_8 24Oct06
	SetMasks();
	Validation_Init();
	Initialise();
	if (m_sReadOnly == "1")frmScreen.btnOK.disabled = true;	
	window.returnValue = null;   // return null if OK is not pressed
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
	
	<% /* MAR350 - Maha T */ %>
	ShowHideFormItems();
}

function Initialise()
{
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2m_subQuoteXML = new scXMLFunctions.XMLObject();%>
	m_subQuoteXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_subQuoteXML.LoadXML(m_sSubQuoteXML);
	<%// EP2_8 - Load Param values and set flag for further filtering %>
	m_CM100XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_CM100XML.LoadXML(m_sCM100XML);
	m_CM100XML.SelectTag(null, "PARAMS");
	m_CalculatedFiltering = m_CM100XML.GetTagText("FLEXIBLEPRODUCTS");
	
	m_subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE");
	m_sApplicationNumber = m_subQuoteXML.GetTagText("APPLICATIONNUMBER");
	m_sApplicationFactFindNumber = m_subQuoteXML.GetTagText("APPLICATIONFACTFINDNUMBER");
	m_sMortgageSubQuoteNumber = m_subQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER");
	PopulateScreen();
	if(m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnFurtherFiltering.disabled = true;
		frmScreen.btnSearch.disabled = true;
	}
	if(m_sApplicationMode == "Quick Quote")
	{
<%		/* Don't allow the user to change the amount requested if the maximum number of
			allowed components is one */	
%>		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var XML = new scXMLFunctions.XMLObject();%>
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		var lMaxQQComponents = parseInt(XML.GetGlobalParameterAmount(document,"MaximumQQComponents"));
		if(lMaxQQComponents == 1)scScreenFunctions.SetFieldState(frmScreen, "txtLoanAmount", "R");
		XML = null;
	}
		
}
function PopulateScreen()
{
	var bShowLoanCompToBePortedQuestion = true;
	var bContinue = true;
	var m_Indicator;
	
	PopulateCombos();

	m_bShowLoanCompToBeRetainedQuestion = true;
	m_sMultipleLender = "0";
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var XML = new scXMLFunctions.XMLObject();%>
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.RunASP(document, "GetMultipleLender.asp");
	if (XML.IsResponseOK() == true)
	{
		m_sMultipleLender = XML.GetTagText("MULTIPLELENDER");
	}

	<% /* LDM 27/02/07 EP2_1567 dont show the porting question if the case is not a porting one.
		  moved to the top of the function to stop question appearing underneath the hardcoded popup  
		  Check NP (validation type for "new loan - ported")*/ %>
	if(XML.IsInComboValidationList(document, 'TypeOfMortgage', m_sidMortgageApplicationValue , Array('NP')) == false)
	{
		scScreenFunctions.HideCollection(spnLoanComponentToBePorted);
		bShowLoanCompToBePortedQuestion = false;
	}
	<% /* Check PSW (validation type for "product switch") */ %>
	if(XML.IsInComboValidationList(document, 'TypeOfMortgage', m_sidMortgageApplicationValue, Array('PSW')) == false)
	{
		scScreenFunctions.HideCollection(spnLoanComponentToBeRetained);
		m_bShowLoanCompToBeRetainedQuestion = false;
	}
	XML = null;

<%	/* set the table to the correct one*/
%>	if (m_sMultipleLender == "1")
	{
		m_tblProductTable = tblMultipleTable;
		scScreenFunctions.HideCollection(tblSingleTable);
		scScreenFunctions.HideCollection(spnSingleTable);
		scScreenFunctions.ShowCollection(tblMultipleTable);
	}
	else
	{
		m_tblProductTable = tblSingleTable;
		scScreenFunctions.HideCollection(tblMultipleTable);
		scScreenFunctions.ShowCollection(tblSingleTable);
	}	
	
	scTable.initialise(m_tblProductTable, 0, "");

	if(m_sLoanComponentSeqNum == null)
	{
		m_subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE");
		
		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2m_NewLoanCompXML = new scXMLFunctions.XMLObject();%>
		m_NewLoanCompXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_NewLoanCompXML.CreateRequestTagFromArray(m_sRequestAttributes, null);
		m_NewLoanCompXML.CreateActiveTag("LOANCOMPONENTDETAILS");
		m_NewLoanCompXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		m_NewLoanCompXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		if( parseInt(m_sNumberOfLoanComponents) > 0 )m_NewLoanCompXML.CreateTag("NUMBEROFCOMPONENTS", m_sNumberOfLoanComponents);
		else m_NewLoanCompXML.CreateTag("NUMBEROFCOMPONENTS", "");
		m_NewLoanCompXML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber);
		m_NewLoanCompXML.CreateTag("AMOUNTREQUESTED", m_subQuoteXML.GetTagText("AMOUNTREQUESTED"));
		if(m_sApplicationMode == "Cost Modelling")m_NewLoanCompXML.RunASP(document,"AQGetDefaultsForNewLoanComponent.asp");
		else m_NewLoanCompXML.RunASP(document, "QQGetDefaultsForNewLoanComponent.asp");
		if(m_NewLoanCompXML.IsResponseOK() == false)
		{
			frmScreen.btnDetails.disabled = true;
			frmScreen.btnFurtherFiltering.disabled = true;
			frmScreen.btnIncentives.disabled = true;
			<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2frmScreen.btnNonPanelDetails.disabled = true;%>
			frmScreen.btnOK.disabled = true;
			frmScreen.btnSearch.disabled = true;
			frmScreen.btnSplit.disabled = true;
			frmScreen.btnCancel.disabled = false;
			frmScreen.btnCancel.focus();
			bContinue = false;
		}
	}

	<% /* AShaw EP2_8 19Oct06 - Calculate Additional term */%>
	m_sAdditionalTerm = "0";
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	<% /* PSC 30/01/2007 EP2_1110 */ %>
	if( (m_AddOrEdit == "Add" || m_AddOrEdit == "Edit") && XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("F")) && XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("M")))
	{
			<% /* Get the AccountGUID */%>
			var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTagFromArray(m_sRequestAttributes, null);
			XML.CreateActiveTag("APPLICATION");
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
			XML.RunASP(document,"GetApplicationData.asp");
			if(XML.IsResponseOK())
			{
				XML.SelectTag(null,"APPLICATION");
				var AccountGUID = XML.GetTagText("ACCOUNTGUID");
			}

			if ( AccountGUID != "") <% /* If "" was getting all rows - Officially classified as BAD! */%>
			{
				<% /* PSC 30/01/2007 EP2_1110 - Start */ %>
				<% /* Get the Loan details for given AccountGUID */%>
				XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
				XML.CreateRequestTagFromArray(m_sRequestAttributes, null);
				var xn = XML.XMLDocument.documentElement;
				xn.setAttribute("CRUD_OP","READ");
				xn.setAttribute("SCHEMA_NAME","epsomCRUD");
				xn.setAttribute("ENTITY_REF","MORTGAGELOAN");
				var xe = XML.XMLDocument.createElement("MORTGAGELOAN");
				xe.setAttribute("ACCOUNTGUID", AccountGUID);
				xn.appendChild(xe);
				<% /* PSC 30/01/2007 EP2_1110 - End */ %>

				XML.RunASP(document, "omCRUDIf.asp");
				if (XML.IsResponseOK())
				{
					<% /* Loop round results, and set m_sAdditionalTerm to longest period */%>
					XML.ActiveTag = null;
					XML.CreateTagList("MORTGAGELOAN");
					m_iNumOfLoanComponents = XML.ActiveTagList.length;
					var l_sAdditionalTerm = 0;
					var tagActiveLoan = null;
					for (var iCount = 0; iCount < m_iNumOfLoanComponents; iCount++)
					{
						XML.SelectTagListItem(iCount);
						tagActiveLoan = XML.ActiveTag;
						
						<% /* PSC 30/01/2007 EP2_1110 - Start */ %>
						var sRedemptionStatus = XML.GetAttribute("REDEMPTIONSTATUS");
						var statusXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject(); 
						if (statusXML.IsInComboValidationList(document,"RedemptionStatus", sRedemptionStatus, ["N", "TBP", "PS", "TOE"]));
						{					
							<% /* PSC 30/01/2007 EP2_1100 */ %>
							var sLoanStartDate = XML.GetAttribute("STARTDATE")
							var lYears = XML.GetAttribute("ORIGINALTERMYEARS");
							if (lYears == "") lYears = "0";
							var lMonths = XML.GetAttribute("ORIGINALTERMMONTHS");
							if (lMonths == "") lMonths = "0";
						
							<% /* PSC 30/01/2007 EP2_1100 */ %>
							l_sAdditionalTerm = CalculateOutstandingTerm(sLoanStartDate, lYears, lMonths); 
						
							<% /* Set if higher than existing highest  */%>
							if (l_sAdditionalTerm > m_sAdditionalTerm)
								m_sAdditionalTerm = l_sAdditionalTerm;
						}
						<% /* PSC 30/01/2007 EP2_1110 - End */ %>
					}

				} // End - XML.IsResponseOK()

			}	// End - AccountGUID != ""
	}	// End - m_AddOrEdit == "Add" etc...	

	if(bContinue)
	{
		<% /* MV - 11/06/2002 - BMIDS00032 - Default ComboRepayment Type */ %>
		<% /* if(parseInt(m_sNumberOfLoanComponents) == 0)
		{
			//If we currently have no loan components then we must be adding the first one
            //so we know that we will have a m_NewLoanCompXML set up. 
			m_NewLoanCompXML.SelectTag(null, "NEWLOANCOMPONENT");
			var iAttitudeToBorrowingScore = parseInt(m_NewLoanCompXML.GetTagText("ATTITUDETOBORROWINGSCORE"));
			if(iAttitudeToBorrowingScore == 0)
				scScreenFunctions.SetComboOnValidationType(frmScreen, "cboRepaymentType", "P");
			else if(iAttitudeToBorrowingScore < 0)
				scScreenFunctions.SetComboOnValidationType(frmScreen, "cboRepaymentType", "C");
			else
				scScreenFunctions.SetComboOnValidationType(frmScreen, "cboRepaymentType", "I");
		} */ %>
		if(m_sLoanComponentSeqNum != null)
		{
<%			/*Locate the loan component within subquoteXML which corresponds to the seq num*/
%>			
			m_subQuoteXML.CreateTagList("LOANCOMPONENT");
			var bFound = false;
			var tagActiveLoan = null;
			for (var iCount = 0; iCount < m_subQuoteXML.ActiveTagList.length && bFound == false; iCount++)
			{
				m_subQuoteXML.SelectTagListItem(iCount);
				tagActiveLoan = m_subQuoteXML.ActiveTag;
				if(m_subQuoteXML.GetTagText("LOANCOMPONENTSEQUENCENUMBER") == m_sLoanComponentSeqNum)
					bFound = true;
					
			}
			if(bFound)
			{
				// EP2_55 - Setup form enabling / Defaults.
				//Initialise - LoanComponent != Null

				// Check NP (validation type for "new loan - ported")
				if(bShowLoanCompToBePortedQuestion == true)
				{
					// Set the Indicator and Radio Button value
					m_Indicator = m_subQuoteXML.GetTagText("MANUALPORTEDLOANIND");
					scScreenFunctions.SetRadioGroupValue(frmScreen, "LoanComponentToBePortedInd", m_Indicator);
					// Disable the buttons.
					frmScreen.optLoanComponentToBePortedYes.disabled = true;
					frmScreen.optLoanComponentToBePortedNo.disabled = true;
					// Do the bizz.
					if (frmScreen.optLoanComponentToBePortedYes.checked == true)
					{	// Disable fields.
						frmScreen.cboPurposeOfLoan.disabled = true;
						frmScreen.txtTermYears.disabled = true;
						frmScreen.txtTermMonths.disabled = true;
						frmScreen.cboRepaymentType.disabled = true;
						frmScreen.cboRepaymentVehicle.disabled = true; <%/*EP_444*/%>
						frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true; <%/*EP_444*/%>
						frmScreen.txtLoanAmount.disabled = true;
						frmScreen.btnFurtherFiltering.disabled = true;
						frmScreen.btnSearch.disabled = true;
						<% /* PSC 13/02/2007 EP2_1314 */ %>
						frmScreen.cboSubPurposeOfLoan.disabled = true;
					}
					else  // Ind = False
					{	// Enable all fields 
						frmScreen.cboPurposeOfLoan.disabled = false;
						frmScreen.txtTermYears.disabled = false;
						frmScreen.txtTermMonths.disabled = false;
						frmScreen.cboRepaymentType.disabled = false;
						frmScreen.cboRepaymentVehicle.disabled = false; <%/*EP_444*/%>
						frmScreen.txtRepaymentVehicleMonthlyCost.disabled = false; <%/*EP_444*/%>
						frmScreen.txtLoanAmount.disabled = false;
						frmScreen.btnFurtherFiltering.disabled = false;
						frmScreen.btnSearch.disabled = false;
						<% /* PSC 13/02/2007 EP2_1314 */ %>
						frmScreen.cboSubPurposeOfLoan.disabled = false;
					}
				}

				// Check PSW (validation type for "product switch")
				if(m_bShowLoanCompToBeRetainedQuestion == true)
				{
					// Set the Indicator and Radio Button value
					m_Indicator = m_subQuoteXML.GetTagText("PRODUCTSWITCHRETAINPRODUCTIND");
					scScreenFunctions.SetRadioGroupValue(frmScreen, "LoanComponentToBeRetainedInd", m_Indicator);
					// Enable the buttons.
					frmScreen.optLoanComponentToBeRetainedYes.disabled = false;
					frmScreen.optLoanComponentToBeRetainedNo.disabled = false;
					
					<% /* PSC 13/02/2007 EP2_1314 - Start */ %>
					frmScreen.cboPurposeOfLoan.disabled = true;
					frmScreen.txtTermYears.disabled = true;
					frmScreen.txtTermMonths.disabled = true;
					frmScreen.txtLoanAmount.disabled = true;
					frmScreen.cboSubPurposeOfLoan.disabled = true;						
					<% /* PSC 13/02/2007 EP2_1314  - End */ %>
					<%/* EP2_1750 MAH 09/03/2007 start */%>
					frmScreen.cboRepaymentType.disabled = true;
					frmScreen.cboRepaymentVehicle.disabled = true;
					frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true;
					<%/* EP2_1750 End */%>
					
					if (frmScreen.optLoanComponentToBeRetainedYes.checked == true)
					{
						<% /* PSC 13/02/2007 EP2_1314 - Start */ %>
						<% /* EP2_1750 MAH 09/03/2007
						frmScreen.cboRepaymentType.disabled = true;
						frmScreen.cboRepaymentVehicle.disabled = true;
						frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true; */ %>
						frmScreen.btnFurtherFiltering.disabled = true;
						frmScreen.btnSearch.disabled = true;
						<% /* PSC 13/02/2007 EP2_1314 - Start */ %>
					}
					else
					{
						<% /* PSC 13/02/2007 EP2_1314 - Start */ %>
						<% /* EP2_1750 MAH 09/03/2007
						frmScreen.cboRepaymentType.disabled = false;
						frmScreen.cboRepaymentVehicle.disabled = false;
						frmScreen.txtRepaymentVehicleMonthlyCost.disabled = false; */ %>
						frmScreen.btnFurtherFiltering.disabled = false;
						frmScreen.btnSearch.disabled = false;
						<% /* PSC 13/02/2007 EP2_1314 - End */ %>
					}
				}
					
				var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
				// Check TOE
				if(XML.IsInComboValidationList(document, 'TypeOfMortgage', m_sidMortgageApplicationValue, Array('TOE')) == true)
				{	// Disable fields.
					frmScreen.cboPurposeOfLoan.disabled = true;
					frmScreen.txtTermYears.disabled = true;
					frmScreen.txtTermMonths.disabled = true;
					frmScreen.cboRepaymentType.disabled = true;
					frmScreen.cboRepaymentVehicle.disabled = true; <%/*EP_444*/%>
					frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true; <%/*EP_444*/%>
					frmScreen.txtLoanAmount.disabled = true;
					frmScreen.btnFurtherFiltering.disabled = true;
					frmScreen.btnSearch.disabled = true;
				}
				frmScreen.cboPurposeOfLoan.value = m_subQuoteXML.GetTagText("PURPOSEOFLOAN");
				frmScreen.txtLoanAmount.value = m_subQuoteXML.GetTagText("LOANAMOUNT");
				frmScreen.cboRepaymentType.value = m_subQuoteXML.GetTagText("REPAYMENTMETHOD");

				<%/*MAH 03/01/2007 EP2_444*/%>
				var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
				var ValidationList = new Array(2);
				ValidationList[0] = "P";
				ValidationList[1] = "I"
				
				<% // SW 20/06/2006 EP771 %>
				<%/*if( frmScreen.cboRepaymentType.value != "" && frmScreen.cboRepaymentType.value != null)--EP2_444*/%>
				if(frmScreen.cboRepaymentType.value != "" && frmScreen.cboRepaymentType.value != null
					&& XML.IsInComboValidationList(document,"RepaymentType", frmScreen.cboRepaymentType.value, ValidationList))<%/*++ MAH 29/12/2006 EP2_444*/%>
				{
					<%/* EP2_1750 
					frmScreen.cboRepaymentVehicle.disabled = false;*/%>
					if(!m_bShowLoanCompToBeRetainedQuestion){frmScreen.cboRepaymentVehicle.disabled = false;}					PopulateRepaymentVehicleCombo();
					m_bDataRead = true;<%/*EP2_444*/%>
					frmScreen.cboRepaymentVehicle.value  = m_subQuoteXML.GetTagText("REPAYMENTVEHICLE");
					frmScreen.txtRepaymentVehicleMonthlyCost.value = m_subQuoteXML.GetTagText("REPAYMENTVEHICLEMONTHLYCOST");<%/*EP2_444*/%>
				}
				else
				{
					frmScreen.cboRepaymentVehicle.value="";<%/*EP2_444*/%>
					frmScreen.cboRepaymentVehicle.disabled = true;
					frmScreen.txtRepaymentVehicleMonthlyCost.value = "";<%/*EP2_444*/%>
					frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true; <%/*EP2_444*/%>
				}
				XML = null;<%/*EP2_444*/%>
				ValidationList = null;<%/*EP2_444*/%>
				
				//CMWP3 - DPF Jul 02 - pull back two new fields from XML - assign one to variable
				<% /* EP430 Uncomment to reverse MAR58 */ %>
				//START: (MAR58) Code commented by Joyce Joseph on 30-Sep-2005
				frmScreen.cboSubPurposeOfLoan.value = m_subQuoteXML.GetTagText("SUBPURPOSEOFLOAN");
				// EP2_55 Moved code
				//m_LoanIndicator = m_subQuoteXML.GetTagText("MANUALPORTEDLOANIND");
				//scScreenFunctions.SetRadioGroupValue(frmScreen, "LoanComponentToBePortedInd", m_subQuoteXML.GetTagText("MANUALPORTEDLOANIND"));
				<% /* MAR46 GHun */ %>
				scScreenFunctions.SetRadioGroupValue(frmScreen, "LoanComponentToBeRetainedInd", m_subQuoteXML.GetTagText("PRODUCTSWITCHRETAINPRODUCTIND"));
				
				//GD SYS2530 changed "Part & Part" to "Part and Part"
				if(m_subQuoteXML.GetTagAttribute("REPAYMENTMETHOD", "TEXT") == "Part and Part")
				{
					frmScreen.btnSplit.disabled = false;
					m_sINTERESTONLYAMOUNT = m_subQuoteXML.GetTagText("INTERESTONLYELEMENT");
					m_sCAPITALINTERESTAMOUNT = m_subQuoteXML.GetTagText("CAPITALANDINTERESTELEMENT");
				}
				else
				{
					frmScreen.btnSplit.disabled = true;
					m_sINTERESTONLYAMOUNT = "0";
					m_sCAPITALINTERESTAMOUNT = "0";
				}
				frmScreen.txtTermYears.value = m_subQuoteXML.GetTagText("TERMINYEARS");
				frmScreen.txtTermMonths.value = m_subQuoteXML.GetTagText("TERMINMONTHS");
<%				/* fill in the list box with the loan component mortgage product details.
				   Create a bogus m_ProductXML with the details found, as populateListBox
				   uses m_ProductXML */
%>				m_ProductXML = null;
				<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2m_ProductXML = new scXMLFunctions.XMLObject();%>
				m_ProductXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
				m_ProductXML.CreateRequestTagFromArray(m_sRequestAttributes, null);
				m_ProductXML.CreateActiveTag("MORTGAGEPRODUCTLIST");
				m_ProductXML.CreateActiveTag("MORTGAGEPRODUCT");
				var sLenderName = "";
				var sProductName = "";
				var sProductNum = "";
				var sInterestRate = "";
				var sType = "";
				var sRatePeriod = "";
				var sDiscountAmount = "";
				<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
				<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
				var sResolvedRate = "";
				<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
				
				
				var sIsFlexible = "";
				// needed for the Details and Incentives buttons ...
				var sProdDesc = "";
				var sProdCode = "";
				var sStartDate = "";
				if(null != m_subQuoteXML.SelectTag(tagActiveLoan, "NONPANELMORTGAGEPRODUCT"))
				{
					sInterestRate = m_subQuoteXML.GetTagText("INTERESTRATE1");  //from NONPANELMORTGAGEPRODUCT XML
					m_subQuoteXML.ActiveTag = tagActiveLoan;
					sProductName = m_subQuoteXML.GetTagText("PRODUCTNAME"); // from MORTGAGEPRODUCTLANGUAGE XML
					sProductNum = m_subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE"); // from MORTGAGEPRODUCTLANGUAGE XML
					sProdDesc = m_subQuoteXML.GetTagText("PRODUCTTEXTDETAILS"); // from MORTGAGEPRODUCTLANGUAGE XML
					//	AW	10/06/02	BM013
					sType = m_subQuoteXML.GetTagText("RATETYPE"); // from MORTGAGEPRODUCT XML
					sRatePeriod = m_subQuoteXML.GetTagText("INTERESTRATEPERIOD"); // from MORTGAGEPRODUCT 
					
					sProdCode = m_subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE"); // from MORTGAGEPRODUCT
					sStartDate = m_subQuoteXML.GetTagText("STARTDATE"); // from MORTGAGEPRODUCT
					
					m_subQuoteXML.CreateTagList("INTERESTRATETYPE");
					sDiscountAmount = m_subQuoteXML.GetTagText("RATE");
					<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
					<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
					sResolvedRate = m_subQuoteXML.GetTagText("RESOLVEDRATE");
					<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
					if(m_sMultipleLender == "1")
					{
						m_subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE");
						sLenderName = m_subQuoteXML.GetTagText("NONPANELLENDERNAME"); // from MORTGAGESUBQUOTE XML
					}
					<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2frmScreen.btnNonPanelDetails.disabled = false;%>
				}
				else
				{
					m_subQuoteXML.ActiveTag = tagActiveLoan;
					if(m_sMultipleLender == "1")
					{
						sLenderName = m_subQuoteXML.GetTagText("LENDERNAME");  //from MORTGAGELENDER XML
					}
					sProductName = m_subQuoteXML.GetTagText("PRODUCTNAME"); // from MORTGAGEPRODUCTLANGUAGE XML
					sProductNum = m_subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE"); // from MORTGAGEPRODUCTLANGUAGE XML
					sProdDesc = m_subQuoteXML.GetTagText("PRODUCTTEXTDETAILS"); // from MORTGAGEPRODUCTLANGUAGE XML
					//	AW	10/06/02	BM013
					sType = m_subQuoteXML.GetTagText("RATETYPE"); // from MORTGAGEPRODUCT XML
					sRatePeriod = m_subQuoteXML.GetTagText("INTERESTRATEPERIOD"); // from MORTGAGEPRODUCT 
					<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
					sIsFlexible = m_subQuoteXML.GetTagText("FLEXIBLEMORTGAGEPRODUCT"); // from MORTGAGEPRODUCT XML
					sProdCode = m_subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE"); // from MORTGAGEPRODUCT
					sStartDate = m_subQuoteXML.GetTagText("STARTDATE"); // from MORTGAGEPRODUCT
					<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
					<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
					sResolvedRate = m_subQuoteXML.GetTagText("RESOLVEDRATE");
					<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
					sInterestRate = GetFirstInterestRate(tagActiveLoan);
					
					m_subQuoteXML.CreateTagList("INTERESTRATETYPE");
					sDiscountAmount = m_subQuoteXML.GetTagText("RATE");
					//frmScreen.btnNonPanelDetails.disabled = true;
				}
				m_ProductXML.CreateTag("MORTGAGEPRODUCTCODE", sProductNum);
				m_ProductXML.CreateTag("PRODUCTNAME", sProductName);
				<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
				m_ProductXML.CreateTag("FLEXIBLEMORTGAGEPRODUCT", sIsFlexible);				
				//	AW	10/06/02	BM013
				m_ProductXML.CreateTag("TYPE", sType);
				m_ProductXML.CreateTag("INTERESTRATEPERIOD", sRatePeriod);
				m_ProductXML.CreateTag("DISCOUNTAMOUNT", sDiscountAmount);
				m_ProductXML.CreateTag("LENDERSNAME", sLenderName);
				m_ProductXML.CreateTag("FIRSTMONTHLYINTERESTRATE", sInterestRate);
				m_ProductXML.CreateTag("PRODUCTTEXTDETAILS", sProdDesc);
				m_ProductXML.CreateTag("MORTGAGEPRODUCTCODE", sProdCode);
				m_ProductXML.CreateTag("STARTDATE", sStartDate);
				<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
				<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
				m_ProductXML.CreateTag("RESOLVEDRATE", sResolvedRate);
				<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
<%				/* Now add the details to the listbox */
%>				PopulateListBox();
				scTable.setRowSelected(1);
//EP2_55		frmScreen.btnFurtherFiltering.disabled = false;
//EP2_55		frmScreen.btnSearch.disabled = false;
				frmScreen.btnDetails.disabled = false;
				frmScreen.btnIncentives.disabled = false;
				frmScreen.btnCancel.disabled = false;
				frmScreen.btnOK.disabled = false;
				m_subQuoteXML.ActiveTag = tagActiveLoan;
				m_sORIGINALPRODUCTCODE = m_subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE");
				m_sORIGINALPRODUCTSTARTDATE = m_subQuoteXML.GetTagText("STARTDATE");
				m_sNEWPRODUCTCODE = "";
				m_sNEWPRODUCSTARTDATE = "";
				if(m_subQuoteXML.SelectTag(tagActiveLoan, "MORTGAGEINCENTIVELIST") != null)
				{
					m_sORIGINALINCENTIVESXML = m_subQuoteXML.ActiveTag.xml;
				}
				else
				{
					m_sORIGINALINCENTIVESXML = "";
					m_sNEWINCENTIVESXML = "";
				}
			}
		} 
		else
		{
			<% /* Initialise the screen to add a new loan component */ %>			
			<% /* MV - 11/06/2002 - BMIDS00032 - Default ComboRepayment Type */ %>
			m_NewLoanCompXML.SelectTag(null, "NEWLOANCOMPONENT");
			
			<% /* MV - 18/02/2003 - BM0235*/%>
			<% /* PSC 30/01/2007 EP2_1100 - Start */ %>
			if (m_iNumOfLoanComponents > 0)
			{
				var XMLTypeOfApp = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
 
			    if (XMLTypeOfApp.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("F")) && 
			        XMLTypeOfApp.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("M")))
				{
					if (m_sAdditionalTerm != 0)
					{
						frmScreen.txtTermYears.value = Math.floor(m_sAdditionalTerm /12);
						frmScreen.txtTermMonths.value = m_sAdditionalTerm % 12;
					}
					else
					{
						alert("The remaining loan term is not set on the existing account. The term has been defaulted.");
					}
				}
			}
			<% /* PSC 30/01/2007 EP2_1100 - End */ %>
				
			frmScreen.cboRepaymentType.value  = m_NewLoanCompXML.GetTagText("REPAYMENTTYPE");

				<%/*MAH 03/01/2007 EP2_444*/%>
				var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
				var ValidationList = new Array(2);
				ValidationList[0] = "P";
				ValidationList[1] = "I"
				
				<% // SW 20/06/2006 EP771 %>
				<%/*if( frmScreen.cboRepaymentType.value != "" && frmScreen.cboRepaymentType.value != null)--EP2_444*/%>
				if(frmScreen.cboRepaymentType.value != "" && frmScreen.cboRepaymentType.value != null
					&& XML.IsInComboValidationList(document,"RepaymentType", frmScreen.cboRepaymentType.value, ValidationList))<%/*++ MAH 29/12/2006 EP2_444*/%>
			{
				m_bDataRead = true;<%/*EP2_444*/%>
				<%/* EP2_1750
				frmScreen.cboRepaymentVehicle.disabled = false;*/%>
				if(!m_bShowLoanCompToBeRetainedQuestion){frmScreen.cboRepaymentVehicle.disabled = false;}
				PopulateRepaymentVehicleCombo();
				frmScreen.cboRepaymentVehicle.value  = m_NewLoanCompXML.GetTagText("REPAYMENTVEHICLE");
				frmScreen.txtRepaymentVehicleMonthlyCost.value = m_NewLoanCompXML.GetTagText("REPAYMENTVEHICLEMONTHLYCOST");<%/*EP2_444*/%>
			}
			else
			{
				frmScreen.cboRepaymentVehicle.disabled = true;	
				frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true;	<%/*EP2_444*/%>
			}
			XML = null;<%/*EP2_444*/%>
			ValidationList = null;<%/*EP2_444*/%>
			frmScreen.txtLoanAmount.value = m_NewLoanCompXML.GetTagText("DEFAULTLOANAMOUNT");
			
			<% /* PSC 30/01/2007 EP2_1100 - Start */ %>
			if (frmScreen.txtTermYears.value == "")
				frmScreen.txtTermYears.value = m_NewLoanCompXML.GetTagText("DEFAULTTERMYEARS");
			
			if (frmScreen.txtTermMonths.value == "")	
				frmScreen.txtTermMonths.value = m_NewLoanCompXML.GetTagText("DEFAULTTERMMONTHS");
			<% /* PSC 30/01/2007 EP2_1100 - End */ %>
				
			frmScreen.btnFurtherFiltering.disabled = false;
			frmScreen.btnSearch.disabled = false;
			frmScreen.btnOK.disabled = true;
			frmScreen.btnCancel.disabled = false;
			frmScreen.btnSplit.disabled = true;
			frmScreen.btnDetails.disabled = true;
			<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
			<%//frmScreen.btnNonPanelDetails.disabled = true;%>
			frmScreen.btnIncentives.disabled = true;
			m_sORIGINALPRODUCTCODE = "";
			m_sORIGINALPRODUCTSTARTDATE = "";
			m_sNEWPRODUCTCODE = "";
			m_sNEWPRODUCSTARTDATE = "";
			m_sORIGINALINCENTIVESXML = "";
			m_sNEWINCENTIVESXML = "";
		}
	}

}
function GetFirstInterestRate( currentActiveTag )
{
	var strRate = "";
	var bSuccess = false;
<%	// get the baserate rate for future reference
%>	m_subQuoteXML.SelectTag(currentActiveTag, "BASERATEBAND");
	var strBaseRate = m_subQuoteXML.GetTagText("RATE");
		
	m_subQuoteXML.ActiveTag = currentActiveTag;
	m_subQuoteXML.CreateTagList("INTERESTRATETYPE");
	for (var iCount = 0; iCount < m_subQuoteXML.ActiveTagList.length && bSuccess == false; iCount++)
	{
		m_subQuoteXML.SelectTagListItem(iCount);
		if(m_subQuoteXML.GetTagText("INTERESTRATETYPESEQUENCENUMBER") == "1")
		{
			var strRateType = m_subQuoteXML.GetTagText("RATETYPE");
			if(strRateType == "F"){
				<% /* PSC 23/03/2007 EP2_1622 */ %>
				strRate = strBaseRate;
				bSuccess = true;
			}
			else if(strRateType == "B"){
				strRate = strBaseRate;
				bSuccess = true;
			}
			else if(strRateType == "D"){
				<%/* EP2_1751 Added Rounding to force 2 decimal places*/%>
				strRate = scMathFunctions.RoundValue(parseFloat(strBaseRate) - parseFloat(m_subQuoteXML.GetTagText("RATE")),2);
				bSuccess = true;
			}
			else if(strRateType == "C"){
				strRate = parseFloat(strBaseRate) - parseFloat(m_subQuoteXML.GetTagText("RATE"));
				if(parseFloat(strRate) < parseFloat(m_subQuoteXML.GetTagText("FLOOREDRATE")))
					strRate = m_subQuoteXML.GetTagText("FLOOREDRATE");
				if(parseFloat(strRate) > parseFloat(m_subQuoteXML.GetTagText("CEILINGRATE")))
					strRate = m_subQuoteXML.GetTagText("CEILINGRATE");
				bSuccess = true;
			}
		}
	}
	return(strRate);	
}
function PopulateCombos()
{
	//	ik_EP2_760_20070114
	var XMLPurposeOfLoan = null;
	var XMLRepaymentType = null;
	<% // SW 20/06/2006 EP771 %>
	var XMLRepaymentVehicle = null;

	var XMLSubPurposeOfLoan = null;  /* CMWP3 - DPF Jul 02 - added subpurpose of loan */
	var ValidationList = new Array(1);
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var XML = new scXMLFunctions.XMLObject();%>
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	ValidationList[0] = "N";
	if( XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, ValidationList))
	{
		// EP2_55 - Default Purpose of Loan combo.
		PopulateNPurposeOfLoanCombo();
	}
	else
	{
		ValidationList[0] = "R";
		if ( XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, ValidationList))
			sPurposeListName = "PurposeOfLoanRemortgage";
		else
		{
			ValidationList[0] = "F";
			if ( XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, ValidationList))
				sPurposeListName = "PurposeOfLoanFurtherAdv";
			else
			{
				ValidationList[0] = "T";
				if ( XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, ValidationList))
					sPurposeListName = "PurposeOfLoanTrfOfEquity";
			}
		}
	}
	
	/* CMWP3 - DPF Jul 02 - added subpurpose of loan */
	<% // SW 20/06/2006 EP771 %>
	var sGroupList = new Array("RepaymentType", "SubPurposeOfLoan","RepaymentVehicle");
	// Add Purposelist if present.
	if (sPurposeListName != "") sGroupList[3] = sPurposeListName;
	if(XML.GetComboLists(document, sGroupList) == true)
	{

		XMLRepaymentType = XML.GetComboListXML("RepaymentType");
		
		<% // SW 20/06/2006 EP771 %>
		XMLRepaymentVehicle = XML.GetComboListXML("RepaymentVehicle");
		
		XMLSubPurposeOfLoan = XML.GetComboListXML("SubPurposeOfLoan")
		
		//	ik_EP2_760_20070114
		if (sPurposeListName != "")
			XMLPurposeOfLoan = XML.GetComboListXML(sPurposeListName)
		
<%		/* Need to remove part&part repayment type if there are more than one loancomps already 
           or we already have one and are trying to add another one */
%>		var iIndex = -1;
		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var RepayXML = new scXMLFunctions.XMLObject();%>
		var RepayXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		if(parseInt(m_sNumberOfLoanComponents) > 1 || (parseInt(m_sNumberOfLoanComponents) == 1 && m_sLoanComponentSeqNum == null))
		{
			RepayXML.CreateActiveTag("TOP");
			RepayXML.AddXMLBlock(XMLRepaymentType);
			RepayXML.CreateTagList("LISTENTRY");
			
			for (var iCount = 0; iCount < RepayXML.ActiveTagList.length && iIndex == -1; iCount++)
			{
				RepayXML.SelectTagListItem(iCount);
				if(RepayXML.GetTagText("VALIDATIONTYPE") == "P")iIndex = iCount;
			}
			if(iIndex != -1)
			{
				var tagNode = RepayXML.SelectTag(null, "LISTNAME");
				tagNode.removeChild(tagNode.childNodes.item(iIndex));
			}
			RepayXML.SelectTag(null, "TOP");
			XMLRepaymentType = RepayXML.CreateFragment();
		}

		// Only fill Purpose of loan if not already filled.
		if (sPurposeListName != "")
			XML.PopulateComboFromXML(document, frmScreen.cboPurposeOfLoan,XMLPurposeOfLoan,true);
		<% /* EP430 Uncomment to reverse MAR58 */ %>
		XML.PopulateComboFromXML(document, frmScreen.cboSubPurposeOfLoan,XMLSubPurposeOfLoan,true);
		XML.PopulateComboFromXML(document, frmScreen.cboRepaymentType,XMLRepaymentType,true);
		<% // SW 20/06/2006 EP771 %>
		XML.PopulateComboFromXML(document, frmScreen.cboRepaymentVehicle,XMLRepaymentVehicle,true);
	}
	XML = null;
	RepayXML = null;	
	sGroupList = null;
}
function spnTables.onclick()
{
<%	/* On selection of a product, enable the relevant buttons */
%>	
	var blnNonPanelLender = m_tblProductTable.rows(scTable.getRowSelectedId()).getAttribute("NonPanelLender")

	if (blnNonPanelLender)
	{
		frmScreen.btnDetails.disabled = true;
		frmScreen.btnIncentives.disabled = true;
		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
		<%//frmScreen.btnNonPanelDetails.disabled = false;%>
		frmScreen.btnOK.disabled = false; 	// SYS1051
	}
	else
	{
		// SYS1051 Only enable the buttons if the product list is not empty
		if (m_ListEmpty)
		{
			frmScreen.btnDetails.disabled = true;
			frmScreen.btnIncentives.disabled = true;
			frmScreen.btnOK.disabled = true;
		}
		else
		{
			frmScreen.btnDetails.disabled = false;
			frmScreen.btnIncentives.disabled = false;
			frmScreen.btnOK.disabled = false;
		}
		// SYS1051 End
		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
		<%//frmScreen.btnNonPanelDetails.disabled = true;%>
	}
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
	frmScreen.btnOK.disabled = false;
	<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
}
function frmScreen.txtLoanAmount.onchange()
{
	m_sINTERESTONLYAMOUNT = "";
	m_sCAPITALINTERESTAMOUNT = "";

	// SYS0878 TJ - 09 April 2001
	if (m_ListEmpty != true)
	{	
		alert("Loan amount has changed - please re-Search available products");
		clearProductTable();	 // SYS1051
	}
	
	//End SYS0878

}
// SYS0878 TJ - 09 April 2001 	
function frmScreen.txtTermYears.onchange()
{
	if (m_ListEmpty != true)
	{	
		alert("Term in years has changed - please re-Search available products");
		clearProductTable(); 	 // SYS1051
	} 
}
function frmScreen.txtTermMonths.onchange()
{
	if (m_ListEmpty != true)
	{	
		alert("Term in months has changed - please re-Search available products");
		clearProductTable();	 // SYS1051
	}
}

function frmScreen.cboPurposeOfLoan.onchange()
{
	if (m_ListEmpty != true)
	{	
		alert("Purpose of the loan has changed - please re-Search available products");
		clearProductTable();	 // SYS1051
	}
}
//End SYS0878 

<% /* BMIDS00939 MDC 15/11/2002
//CMWP3 - DPF 23/07/02 - have added new field so ensure products selected are actually valid
function frmScreen.optLoanComponentToBePortedNo.onclick()
{
	if (m_LoanIndicator != 0)
	{
		if (m_ListEmpty != true)
		{	
			alert("Loan to be Ported has changed - please re-Search available products");
			clearProductTable();
			m_LoanIndicator = 0
		}
	}
}
function frmScreen.optLoanComponentToBePortedYes.onclick()
{
	if (m_LoanIndicator != 1)
	{
		if (m_ListEmpty != true)
		{	
			alert("Loan to be Ported has changed - please re-Search available products");
			clearProductTable();
			m_LoanIndicator = 1
		}
	}
}
//CMWP3 END
*/ %>
 
function frmScreen.cboRepaymentType.onchange()
{
	m_bDataRead = false;<%/*EP2_444*/%>
	frmScreen.txtRepaymentVehicleMonthlyCost.value = "";<%/*EP2_444*/%>
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var XML = new scXMLFunctions.XMLObject();%>
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ValidationList = new Array(1);
	ValidationList[0] = "P";
	<% /* BM0150 MDC 17/12/2002 - Check if value is empty of null
	if( XML.IsInComboValidationList(document,"RepaymentType", frmScreen.cboRepaymentType.value, ValidationList)) */ %>
	if( frmScreen.cboRepaymentType.value != "" && frmScreen.cboRepaymentType.value != null
		&& XML.IsInComboValidationList(document,"RepaymentType", frmScreen.cboRepaymentType.value, ValidationList))
		frmScreen.btnSplit.disabled = false;
	else
	{
		frmScreen.btnSplit.disabled = true;
		m_sINTERESTONLYAMOUNT = "";
		m_sCAPITALINTERESTAMOUNT = "";
	}
	XML = null;
	ValidationList = null;
	if (m_ListEmpty != true)
	{	
		alert("Repayment type changed - please re-Search available products");
		clearProductTable();	 // SYS1051
	}

	<%/*MAH 29/12/2006 EP2_444*/%>
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ValidationList = new Array(2);
	ValidationList[0] = "P";
	ValidationList[1] = "I"
	<% // SW 20/06/2006 EP771 %>
	<%/*if(frmScreen.cboRepaymentType.value != "" && frmScreen.cboRepaymentType.value != null)--EP2_444*/%>
	if(frmScreen.cboRepaymentType.value != "" && frmScreen.cboRepaymentType.value != null
		&& XML.IsInComboValidationList(document,"RepaymentType", frmScreen.cboRepaymentType.value, ValidationList))<%/*++ MAH 29/12/2006 EP2_444*/%>
	{
		<%/* EP2_1750
		frmScreen.cboRepaymentVehicle.disabled = false;*/%>
		if(!m_bShowLoanCompToBeRetainedQuestion){frmScreen.cboRepaymentVehicle.disabled = false;}
		frmScreen.txtRepaymentVehicleMonthlyCost.disabled = false;<%/* ++ MAH 29/12/2006 EP2_444*/%>
		PopulateRepaymentVehicleCombo();
	}
	else
	{
		frmScreen.cboRepaymentVehicle.value="";<%/* ++ MAH 02/01/2007 EP2_444*/%>
		frmScreen.cboRepaymentVehicle.disabled = true;
		frmScreen.txtRepaymentVehicleMonthlyCost.value = "";<%/* ++ MAH 02/01/2007 EP2_444*/%>
		frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true;<%/* ++ MAH 29/12/2006 EP2_444*/%>	
	}
	XML = null;<%/* ++ MAH 02/01/2007 EP2_444*/%>
	ValidationList = null;<%/* ++ MAH 02/01/2007 EP2_444*/%>
}
function frmScreen.btnSearch.onclick()
{	
	frmScreen.style.cursor = "wait";
	if(frmScreen.onsubmit())
	{
		// EP2_8 - Check for invalid Borrowing period
		if (InvalidMortgageTerm() == 1) //Period entered is longer than existing period.
		{
			frmScreen.style.cursor = "default";	<% /* PSC 30/01/2007 EP2_1110 */ %>

			<% /* PSC 30/01/2007 EP2_1100 - Start */ %>
			var lYears = Math.floor(m_sAdditionalTerm / 12);
			var lMonths = m_sAdditionalTerm % 12;
			var sTerm = "";
			
			if (lYears > 0)
				sTerm = lYears + " year";
				
			if (lYears > 1)
			    sTerm += "s";
			    
			if (lMonths > 0)
			{
				if (lYears > 0)
					sTerm += " ";
			
				sTerm += lMonths + " month";
			}
			   
			if (lMonths > 1)
				sTerm += "s";
				
			var lAlertMesage = "Term entered cannot be greater than the term of the existing mortgage account " + sTerm + ".";
			<% /* PSC 30/01/2007 EP2_1100 - End */ %>				
			alert(lAlertMesage);
			return;
		}
		
		//check Part & Part values if neccessary JLD SYS4488
		<%//var XML = new scXMLFunctions.XMLObject();%>
		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
		var ValidationList = new Array(1);
		ValidationList[0] = "P";
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		if( XML.IsInComboValidationList(document,"RepaymentType", frmScreen.cboRepaymentType.value, ValidationList))
		{
			if(m_sINTERESTONLYAMOUNT == "" && m_sCAPITALINTERESTAMOUNT == "")
			{
				alert("Please enter the Part and Part split before searching for products.");
				frmScreen.btnSplit.focus();
				return;
			}
		}
		
		//check that the loan amount is > 0  JLD SYS4489
		if( parseInt(frmScreen.txtLoanAmount.value,10) <=0)
		{
			alert("Please enter a loan amount before searching for products");
			frmScreen.txtLoanAmount.focus();
			return;
		}
		
		scTable.clear();
		scScrollPlus.Clear();
		scScrollPlus.Initialise(FindProducts,DisplayProducts,m_tblProductTable.rows.length - 1,1);		
		
	}
	frmScreen.style.cursor = "default";
}
function PopulateListBox()
{
	var iNumberOfProducts;
	m_ProductXML.ActiveTag = null;
	m_ProductXML.CreateTagList("MORTGAGEPRODUCT");
	iNumberOfProducts = m_ProductXML.ActiveTagList.length;

	scTable.initialise(m_tblProductTable, 0, "");
	if(iNumberOfProducts > 0)
	 {
	   DisplayProducts(0);
	   	m_ListEmpty = false; 
	 }
	
	else
	{
<%		/* If we have no products but the XML returned a type SUCCESS 
			we must have a warning message. Alert the description */
%>		clearProductTable();	 // SYS1051
		m_ProductXML.SelectTag(null, "RESPONSE");
		alert(m_ProductXML.GetTagText("DESCRIPTION"));
	}
}
function DisplayProducts(iStart)
{
	var sProductName			= null;
	var sLenderName				= null;
	var sRateType				= null;
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	var sFlexibleMortgage		= null;
	var sRatePeriod				= null;
	var blnFlexibleMortgage		= false;
	var sInterestRateEndDate	= null;
	var blnNonPanelLender       = null;
	var sRateEndDateOrPeriod	= null;
	var tagPRODUCTLIST			= null;
	var tagPRODUCTDETAILS		= null;
	var iNumberOfProducts		= 0;
	var iHighLight = -1;

	scTable.clear();
	m_ListEmpty = true;
	m_ProductXML.ActiveTag = null;
	tagPRODUCTLIST = m_ProductXML.CreateTagList("MORTGAGEPRODUCT");
	iNumberOfProducts = m_ProductXML.ActiveTagList.length;
    if (iNumberOfProducts > 0) m_ListEmpty = false;
	for (var nProduct = 0; nProduct < iNumberOfProducts && nProduct < m_iTableLength; nProduct++)
	{
		m_ProductXML.ActiveTagList = tagPRODUCTLIST;
		if (m_ProductXML.SelectTagListItem(nProduct + iStart) == true)
		{
			sProductName = m_ProductXML.GetTagText("PRODUCTNAME");
			<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
			blnFlexibleMortgage = m_ProductXML.GetTagBoolean("FLEXIBLEMORTGAGEPRODUCT");
			if (blnFlexibleMortgage)
				sFlexibleMortgage = "Y";
			else
				sFlexibleMortgage = "";
			//	AW	10/06/02	BM013
			switch(m_ProductXML.GetTagText("TYPE"))
			{
				case "F" : sRateType = "Fixed"; break;
				case "D" :	
						if (parseFloat(m_ProductXML.GetTagText("DISCOUNTAMOUNT")) < 0.0)
						{
							sRateType = "Tracker";
						}
						else
						{
							sRateType = "Discount";
						}
						
						break;
				case "C" : sRateType = "Capped/Floor"; break;
				case "B" : sRateType = "Base Variable"; break;
				
				default: sRateType = "";
		
			}		 
			
			sRatePeriod = m_ProductXML.GetTagText("INTERESTRATEPERIOD");
			if(sRatePeriod == "-1") sRatePeriod = ""
	
			var sProductCode = m_ProductXML.GetTagText("MORTGAGEPRODUCTCODE");
			var sStartDate = m_ProductXML.GetTagText("STARTDATE");
			<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
			<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
			var sResolvedRate = m_ProductXML.GetTagText("RESOLVEDRATE");
			<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>			
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
			blnNonPanelLender = m_ProductXML.GetTagBoolean("NONPANELLENDEROPTION");
			
			if (blnNonPanelLender)sInterestRate = ""
			else sInterestRate = m_ProductXML.GetTagText("FIRSTMONTHLYINTERESTRATE");
			<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
			<%//CHECK IF WE NEED THIS LINE DURING TESTINGGD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
			//ShowRow(nProduct+1, sLenderName, sProductCode, sProductName, sInterestRate, sResolvedRate, sFlexibleMortgage, blnNonPanelLender);
			<%/*ShowRow(nProduct+1, sLenderName, sProductCode, sProductName, sInterestRate, sFlexibleMortgage, blnNonPanelLender);*/%>
			<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
			ShowRow(nProduct+1, sLenderName, sProductCode, sProductName, sInterestRate, sRateType, sRatePeriod, blnNonPanelLender);
			if( m_NewLoanCompXML != null &&
			    sProductCode == m_NewLoanCompXML.GetTagText("DEFAULTMORTGAGEPRODUCTCODE") &&
			    sStartDate == m_NewLoanCompXML.GetTagText("DEFAULTSTARTDATE") )
			{
				iHighLight = nProduct+1;
			}
		}
	}
	if(iHighLight != -1)
	{
		scTable.setRowSelected(iHighLight);
		frmScreen.btnDetails.disabled = false;
		frmScreen.btnIncentives.disabled = false;
		frmScreen.btnOK.disabled = false;
	}
	else
	{
		frmScreen.btnDetails.disabled = true;
		frmScreen.btnIncentives.disabled = true;
		frmScreen.btnOK.disabled = true;
	}
	
}
function ShowRow(iRow, sLenderName, sProductNum, sProductName, sRate, sRateType, sRatePeriod, blnNonPanelLender)
{
	var iCellIndex = 0;
	if (m_sMultipleLender == "1")
		scScreenFunctions.SizeTextToField(m_tblProductTable.rows(iRow).cells(iCellIndex++),sLenderName);
	scScreenFunctions.SizeTextToField(m_tblProductTable.rows(iRow).cells(iCellIndex++),	sProductNum);
	scScreenFunctions.SizeTextToField(m_tblProductTable.rows(iRow).cells(iCellIndex++),	sProductName);
	scScreenFunctions.SizeTextToField(m_tblProductTable.rows(iRow).cells(iCellIndex++),	sRate);
	scScreenFunctions.SizeTextToField(m_tblProductTable.rows(iRow).cells(iCellIndex++),	sRateType);
	scScreenFunctions.SizeTextToField(m_tblProductTable.rows(iRow).cells(iCellIndex++),	sRatePeriod);
	m_tblProductTable.rows(iRow).setAttribute("NonPanelLender", blnNonPanelLender);												
}
function frmScreen.btnDetails.onclick()
{
	var sReturn = null;			
	var ArrayArguments = new Array();

	var iOffset = scTable.getRowSelected() - 1;
	m_ProductXML.ActiveTag = null;
	if (m_ProductXML.CreateTagList("MORTGAGEPRODUCT") != null)
	{
		if (m_ProductXML.SelectTagListItem(iOffset) == true)
		{
			ArrayArguments[0] = m_ProductXML.GetTagText("PRODUCTTEXTDETAILS");
			sReturn = scScreenFunctions.DisplayPopup(window, document, "cm115.asp", ArrayArguments, 350, 295);
			
		}
	}
}
function frmScreen.btnSplit.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
			
	ArrayArguments[0] = frmScreen.txtLoanAmount.value;
	ArrayArguments[1] = m_sINTERESTONLYAMOUNT;
	ArrayArguments[2] = m_sCAPITALINTERESTAMOUNT;
	ArrayArguments[3] = m_sReadOnly;
	ArrayArguments[4] = m_sCurrency;
						
	sReturn = scScreenFunctions.DisplayPopup(window, document, "cm160.asp", ArrayArguments, 505, 225);
	if (sReturn != null)
	{						
		m_sINTERESTONLYAMOUNT = sReturn[1];
		m_sCAPITALINTERESTAMOUNT = sReturn[2];
				
		if (sReturn[0] == true)FlagChange(sReturn[0]);
	}
}

function frmScreen.btnIncentives.onclick()
{
	var sReturn = null;			
	var sProductCode = "";
	var sStartDate = "";
	var ArrayArguments = new Array();
	var bContinue = true;

	var iOffset = scTable.getRowSelected() - 1;
	m_ProductXML.ActiveTag = null;
	if (m_ProductXML.CreateTagList("MORTGAGEPRODUCT") != null)
	{
		if (m_ProductXML.SelectTagListItem(iOffset) == true)
		{
			if(m_ProductXML.GetTagText("NONPANELLENDEROPTION") == "1")
			{
				alert("Incentives are not applicable to ported loan components or non-panel lender mortgage products");
				bContinue = false;
			}
			else
			{
				m_ProductXML.ActiveTag = null;
				m_ProductXML.CreateTagList("MORTGAGEPRODUCT");
				if (m_ProductXML.SelectTagListItem(iOffset) == true)
				{
					sProductCode = m_ProductXML.GetTagText("MORTGAGEPRODUCTCODE");
					sStartDate = m_ProductXML.GetTagText("STARTDATE");
				}
			}
		}
	}
	if(bContinue)
	{
		if(sProductCode != m_sNEWPRODUCTCODE || sStartDate != m_sNEWPRODUCTSTARTDATE)
		{
			if(sProductCode == m_sORIGINALPRODUCTCODE && sStartDate == m_sORIGINALPRODUCTSTARTDATE)
				m_sNEWINCENTIVESXML = m_sORIGINALINCENTIVESXML;
			else m_sNEWINCENTIVESXML = "";
			m_sNEWPRODUCTCODE = sProductCode;
			m_sNEWPRODUCTSTARTDATE = sStartDate;
		}
		
		ArrayArguments[0] = m_sApplicationMode;
		ArrayArguments[1] = m_sReadOnly;
		ArrayArguments[2] = m_sRequestAttributes;
		ArrayArguments[3] = sProductCode;
		ArrayArguments[4] = sStartDate;
		ArrayArguments[5] = frmScreen.txtLoanAmount.value;
		ArrayArguments[6] = m_sCurrency;
		ArrayArguments[7] = m_sNEWINCENTIVESXML;
		ArrayArguments[8] = m_sApplicationNumber;
		ArrayArguments[9] = m_sApplicationFactFindNumber;
		ArrayArguments[10]= m_sMortgageSubQuoteNumber;
		ArrayArguments[11]= m_sLoanComponentSeqNum;
		//DPF 17/10/02 - CPWP1 - Pass sub quote data to CM150
		ArrayArguments[12]= m_subQuoteXML.XMLDocument.xml;
		ArrayArguments[13] = m_sApplicationDate; // JD BMIDS816
							
		sReturn = scScreenFunctions.DisplayPopup(window, document, "cm150.asp", ArrayArguments, 505, 465);
		if (sReturn != null)
		{
			m_sNEWINCENTIVESXML = sReturn[2];
		}
	}
}
function frmScreen.btnFurtherFiltering.onclick()
{
	if(m_FurtherFilteringXML == null)
	{

		m_FurtherFilteringXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_FurtherFilteringXML.CreateRequestTagFromArray(m_sRequestAttributes, null);
		m_FurtherFilteringXML.CreateActiveTag("FURTHERFILTERING");
		<% /* EP2_1694 */ %>
		m_FurtherFilteringXML.CreateTag("ALLPRODUCTSWITHCHECKS", "1");
		m_FurtherFilteringXML.CreateTag("ALLPRODUCTSWITHOUTCHECKS", "");
		m_FurtherFilteringXML.CreateTag("PRODUCTSBYGROUP", "");
		m_FurtherFilteringXML.CreateTag("PRODUCTGROUP", "");
		m_FurtherFilteringXML.CreateTag("DISCOUNTEDPRODUCTS", "");
		m_FurtherFilteringXML.CreateTag("DISCOUNTEDPERIOD", "");
		m_FurtherFilteringXML.CreateTag("FIXEDRATEPRODUCTS", "");
		m_FurtherFilteringXML.CreateTag("FIXEDRATEPERIOD", "");
		m_FurtherFilteringXML.CreateTag("STANDARDVARIABLEPRODUCTS", "");
		m_FurtherFilteringXML.CreateTag("CAPPEDFLOOREDPRODUCTS", "");
		m_FurtherFilteringXML.CreateTag("CAPPEDFLOOREDPERIOD", "");
		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
		m_FurtherFilteringXML.CreateTag("PRODUCTOVERRIDECODE", "");
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
		<% /* BMIDS00939 MDC 15/11/2002 */ %>
		m_FurtherFilteringXML.CreateTag("MANUALPORTEDLOANIND", "");
		<% /* BMIDS00939 MDC 15/11/2002 - End */ %>
		// EP2_8 - Set flag to disable options on Further filtering screen.
		m_FurtherFilteringXML.CreateTag("CALCULATEDFILTER", m_CalculatedFiltering );

	}
	var sReturn = null;
	var ArrayArguments = new Array(10);
	ArrayArguments[0] = m_sApplicationMode;
	ArrayArguments[1] = m_sUserId;
	ArrayArguments[2] = m_sUnitId;
	ArrayArguments[3] = m_sUserType;
	ArrayArguments[4] = m_FurtherFilteringXML.XMLDocument.xml;
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
	sReturn = scScreenFunctions.DisplayPopup(window, document, "CM120.asp", ArrayArguments, 490, 480);
	<%/*sReturn = scScreenFunctions.DisplayPopup(window, document, "CM120.asp", ArrayArguments, 490, 440);*/%>
	<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
	if(sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_FurtherFilteringXML.LoadXML(sReturn[1]);
		
		m_bFurtherFiltering = true;   <% /* EP1089/MAR1833 */ %>
	}
}
function SaveLoanComponent()
{
	var bSuccess = true;
	var sProductCode = "";
	var sStartDate = "";
	var sPortedLoan = "0";
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
	var sProductCodeSearchInd = "0";
	
	<%/* STB: SYS4219 - If a specific product was requested in the search, we'll 
	need to store this fact */%>
	if (m_FurtherFilteringXML != null)
	{
		<% /* BMIDS01071 MDC 22/11/2002 
		if (m_FurtherFilteringXML.GetTagText("PRODUCTOVERRIDECODE") == "") */ %>
		if (m_FurtherFilteringXML.ActiveTag == null || m_FurtherFilteringXML.GetTagText("PRODUCTOVERRIDECODE") == "")
		{
			sProductCodeSearchInd = "0";
		}
		else
		{
			sProductCodeSearchInd = "1";
		}
	}
	<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>	
	var iOffset = scTable.getRowSelected() - 1;
	m_ProductXML.ActiveTag = null;
	if (m_ProductXML.CreateTagList("MORTGAGEPRODUCT") != null)
	{
		if (m_ProductXML.SelectTagListItem(iOffset) == true)
		{
			sProductCode = m_ProductXML.GetTagText("MORTGAGEPRODUCTCODE");
			sStartDate = m_ProductXML.GetTagText("STARTDATE");
			if(m_ProductXML.GetTagText("NONPANELLENDEROPTION") == "1")sPortedLoan = "1";
		}
	}
	
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2var XML = new scXMLFunctions.XMLObject();%>

	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();	
	XML.CreateRequestTagFromArray(m_sRequestAttributes, "SAVE");
	XML.CreateActiveTag("MORTGAGESUBQUOTE");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber);
	if(m_sLoanComponentSeqNum != null)XML.CreateTag("LOANCOMPONENTSEQUENCENUMBER", m_sLoanComponentSeqNum);
	XML.CreateTag("PURPOSEOFLOAN", frmScreen.cboPurposeOfLoan.value);
	//CMWP3 - DPF Jul 02 -  two new fields added to the XML to be saved into the LOANCOMPONENT table
	<% /* EP430 - need to save the value selected in the combo */ %>
	XML.CreateTag("SUBPURPOSEOFLOAN",frmScreen.cboSubPurposeOfLoan.value);
	if (frmScreen.optLoanComponentToBePortedYes.checked)
		XML.CreateTag("MANUALPORTEDLOANIND", "1");
	else
		XML.CreateTag("MANUALPORTEDLOANIND", "0");
	<% /* MAR46 GHun */ %>
	if (frmScreen.optLoanComponentToBeRetainedYes.checked)
		XML.CreateTag("PRODUCTSWITCHRETAINPRODUCTIND", "1");
	else
		XML.CreateTag("PRODUCTSWITCHRETAINPRODUCTIND", "0");
	<% /* MAR46 End */ %>
	XML.CreateTag("LOANAMOUNT", frmScreen.txtLoanAmount.value);
	XML.CreateTag("REPAYMENTTYPE", frmScreen.cboRepaymentType.value);
	<% //SW 20/06/2006 EP771 %>
	XML.CreateTag("REPAYMENTVEHICLE", frmScreen.cboRepaymentVehicle.value);
	XML.CreateTag("REPAYMENTVEHICLEMONTHLYCOST", frmScreen.txtRepaymentVehicleMonthlyCost.value);<%/*EP2_444*/%>
	XML.CreateTag("TERMINYEARS", frmScreen.txtTermYears.value);
	if(frmScreen.txtTermMonths.value == "")XML.CreateTag("TERMINMONTHS", "0");
	else XML.CreateTag("TERMINMONTHS", frmScreen.txtTermMonths.value);
	XML.CreateTag("MORTGAGEPRODUCTCODE", sProductCode);
	XML.CreateTag("PORTEDLOAN", sPortedLoan);
	XML.CreateTag("STARTDATE", sStartDate);
	XML.CreateTag("CAPITALANDINTERESTELEMENT", m_sCAPITALINTERESTAMOUNT);
	XML.CreateTag("INTERESTONLYELEMENT", m_sINTERESTONLYAMOUNT);
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
	XML.CreateTag("NETCAPANDINTELEMENT", m_sCAPITALINTERESTAMOUNT);
	XML.CreateTag("NETINTONLYELEMENT", m_sINTERESTONLYAMOUNT);
	<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
	XML.CreateTag("STARTDATE", sStartDate);
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
	XML.CreateTag("PRODUCTCODESEARCHIND", sProductCodeSearchInd);
	XML.CreateTag("RESOLVEDRATE", "");
	XML.CreateTag("MANUALADJUSTMENTPERCENT", "");
	<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
	
	/* the following non-panel details inputs are not yet implemented :-
	XML.CreateTag("LENDERNAME", );
	XML.CreateTag("MORTGAGEPRODUCTNAME", );
	XML.CreateTag("MORTGAGEPRODUCTTYPE", );
	XML.CreateTag("INTERESTRATE1", );
	XML.CreateTag("INTERESTRATE1PERIOD", );
	XML.CreateTag("INTERESTRATE2", );   */

	if(m_sNEWINCENTIVESXML != "")
	{
		XML.CreateTag("INCENTIVESPRODUCTCODE", m_sNEWPRODUCTCODE);
		XML.CreateTag("INCENTIVESPRODUCTSTARTDATE", m_sNEWPRODUCTSTARTDATE);
		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
		<%//var incentivesXML = new scXMLFunctions.XMLObject();%>
		var incentivesXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		incentivesXML.LoadXML(m_sNEWINCENTIVESXML);
		incentivesXML.SelectTag(null, "MORTGAGEINCENTIVELIST");
		XML.CreateActiveTag("INCENTIVELIST");
		XML.AddXMLBlock(incentivesXML.CreateFragment());
		incentivesXML = null;
	}
	else
	{
<%		/* no new incentives so set these to equal the product */
%>		XML.CreateTag("INCENTIVESPRODUCTCODE", sProductCode);
		XML.CreateTag("INCENTIVESPRODUCTSTARTDATE", sStartDate);
	}
	<%//DRC  MAR1380 - Need to put in a critical data check to see if the term has changed %>
	XML.SelectTag(null,"REQUEST");
	XML.SetAttribute("OPERATION","CriticalDataCheck"); 
	XML.CreateActiveTag("CRITICALDATACONTEXT");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.SetAttribute("SOURCEAPPLICATION","Omiga");
	XML.SetAttribute("ACTIVITYID",m_sActivityId);
	XML.SetAttribute("ACTIVITYINSTANCE","1");
	XML.SetAttribute("STAGEID",m_sStageId);
	XML.SetAttribute("APPLICATIONPRIORITY",m_sApplicationPriority);
	XML.SetAttribute("COMPONENT","omCM.MortgageSubQuoteBO");
	XML.SetAttribute("METHOD","SaveLoanComponentDetails");

	// 	XML.RunASP(document, "SaveLoanComponentDetails.asp");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
<%//DRC  MAR1380 - Need to put in a critical data check to see if the term has changed %>		
			// XML.RunASP(document, "SaveLoanComponentDetails.asp");
			window.status = "Critical Data Check - please wait";
			XML.RunASP(document,"OmigaTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	bSuccess = XML.IsResponseOK();
	
	return(bSuccess);
}
function frmScreen.btnOK.onclick()
{
	var sReturn = new Array();
	var bSuccess = true;
	
	//SYS0953 TJ - 10 April 2001
	// Check whether this loan component is one less than the maximum
	// If this is the case the following needs to apply
	//SYS2530 GD start
	if (m_sLoanComponentSeqNum == null)
	{
		if (m_sNumberOfLoanComponents == (m_subQuoteXML.GetTagText("MAXNOLOANS")) - 1) 
			{
			alert("The maximum number of loan components has been reached ... The loan value will be defaulted to the remaining value.");
			frmScreen.txtLoanAmount.value = m_NewLoanCompXML.GetTagText("DEFAULTLOANAMOUNT");
			}
		// End SYS0953 
	}
	//SYS2530 GD end
	
	if(frmScreen.onsubmit())
	{
		if(m_sReadOnly != "1")bSuccess = SaveLoanComponent();
	
		if(bSuccess)
		{
			sReturn[0] = IsChanged();		// Has there been a change made
			window.returnValue = sReturn;
			window.close();
		}
		else
		{
			frmScreen.btnCancel.disabled = false;
			frmScreen.btnCancel.focus();
			frmScreen.btnDetails.disabled = true;
			frmScreen.btnFurtherFiltering.disabled = true;
			frmScreen.btnIncentives.disabled = true;
			<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
			<%//frmScreen.btnNonPanelDetails.disabled = true;%>
			frmScreen.btnOK.disabled = true;
			frmScreen.btnSearch.disabled = true;
			frmScreen.btnSplit.disabled = true;
		}
	}
}
function frmScreen.btnCancel.onclick()
{
	window.close();
}

function FindProducts(iPageNo)
{
	m_subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE");
	<%//m_ProductXML = new scXMLFunctions.XMLObject();%>
	m_ProductXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_ProductXML.CreateRequestTagFromArray(m_sRequestAttributes, "SEARCH");
	m_ProductXML.CreateActiveTag("MORTGAGEPRODUCTREQUEST");
	m_ProductXML.CreateTag("SEARCHCONTEXT", m_sApplicationMode)	
	m_ProductXML.CreateTag("DISTRIBUTIONCHANNELID", m_sDistributionChannelId);
	m_ProductXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	m_ProductXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	m_ProductXML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber);
	m_ProductXML.CreateTag("PURPOSEOFLOAN", frmScreen.cboPurposeOfLoan.value);
	<% /* BMIDS00939 MDC 15/11/2002 */ %>
	<% /*
	//CMWP3 - new field added to Find Products routine
	if (frmScreen.optLoanComponentToBePortedYes.checked)
		m_ProductXML.CreateTag("MANUALPORTEDLOANIND", "1");
	else
		m_ProductXML.CreateTag("MANUALPORTEDLOANIND", "0"); */ %>
	<% /* BMIDS00939 MDC 15/11/2002 - End */ %>
	m_ProductXML.CreateTag("TERMINYEARS", frmScreen.txtTermYears.value);
	if(frmScreen.txtTermMonths.value == "")m_ProductXML.CreateTag("TERMINMONTHS", "0");
	else m_ProductXML.CreateTag("TERMINMONTHS", frmScreen.txtTermMonths.value);
	m_ProductXML.CreateTag("AMOUNTREQUESTED", m_subQuoteXML.GetTagText("AMOUNTREQUESTED"));
	//BMIDS00656 Send in Loan Component Amount & Sequence number to use when checking for Lenders
	m_ProductXML.CreateTag("LOANCOMPONENTAMOUNT", frmScreen.txtLoanAmount.value);
	m_ProductXML.CreateTag("LTV", m_subQuoteXML.GetTagText("LTV"));
	//BMIDS00656 Reset FlexibleMortgageProduct if Loan Component 1 being edited.

	<% /* EP2_8 24Oct06 - Assumes ALL have values set correctly on CM100 */ %>
	m_CM100XML.SelectTag(null, "PARAMS");
	m_ProductXML.CreateTag("PORTABLEPRODUCTS", "")
	if (m_CalculatedFiltering != "-1")
	{
		m_ProductXML.CreateTag("FLEXIBLENONFLEXIBLEPRODUCTS", "1");
		m_ProductXML.CreateTag("FLEXIBLEIND", m_CalculatedFiltering);
	}
	else
	{
		m_ProductXML.CreateTag("FLEXIBLENONFLEXIBLEPRODUCTS", "0");
		m_ProductXML.CreateTag("FLEXIBLEIND", "");
	}
	m_ProductXML.CreateTag("NATUREOFLOAN", m_CM100XML.GetTagText("NATUREOFLOAN"));
	m_ProductXML.CreateTag("CREDITSCHEME", m_CM100XML.GetTagText("CREDITSCHEME"));
	m_ProductXML.CreateTag("GUARANTORIND", m_CM100XML.GetTagText("GUARANTORIND"));
	m_ProductXML.CreateTag("APPLICATIONINCOMESTATUS", m_CM100XML.GetTagText("APPLICATIONINCOMESTATUS"));
	<% /* EP2_8 24Oct06 - End */ %>


	<% /* MV - 18/07/2002 - BMIDS00179 - Core Upgrade Roll back
	if(m_sApplicationMode == "Quick Quote")
	{
		m_ProductXML.CreateTag("TYPEOFAPPLICATION", m_sidMortgageApplicationValue);
		m_ProductXML.CreateTag("TYPEOFBUYER", m_subQuoteXML.GetTagText("TYPEOFBUYER"));
	}
	*/ %>

	var nTableLength = m_tblProductTable.rows.length - 1;	
	m_ProductXML.CreateTag("RECORDCOUNT", m_tblProductTable.rows.length - 1);
	var nStartRecord = ((nTableLength * iPageNo) + 1) - nTableLength; 
	m_ProductXML.CreateTag("STARTRECORD", nStartRecord);
	if(m_bFurtherFiltering == false)
	{
		m_ProductXML.CreateTag("ALLPRODUCTSWITHOUTCHECKS", "");
		m_ProductXML.CreateTag("ALLPRODUCTSWITHCHECKS", "1");
		m_ProductXML.CreateTag("PRODUCTSBYGROUP", "");
		m_ProductXML.CreateTag("PRODUCTGROUP", "");
		m_ProductXML.CreateTag("DISCOUNTEDPRODUCTS", "");
		m_ProductXML.CreateTag("DISCOUNTEDPERIOD", "");
		m_ProductXML.CreateTag("FIXEDRATEPRODUCTS", "");
		m_ProductXML.CreateTag("FIXEDRATEPERIOD", "");
		m_ProductXML.CreateTag("STANDARDVARIABLEPRODUCTS", "");
		m_ProductXML.CreateTag("CAPPEDFLOOREDPRODUCTS", "");
		m_ProductXML.CreateTag("CAPPEDFLOOREDPERIOD", "");
		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
		m_ProductXML.CreateTag("MORTGAGEPRODUCTCODE", "");
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
		<% /* BMIDS00939 MDC 15/11/2002 */ %>
		m_ProductXML.CreateTag("MANUALPORTEDLOANIND", "");
		<% /* BMIDS00939 MDC 15/11/2002 - End */ %>
		
	}
	else
	{
		m_FurtherFilteringXML.SelectTag(null, "FURTHERFILTERING");
		m_ProductXML.CreateTag("ALLPRODUCTSWITHOUTCHECKS", m_FurtherFilteringXML.GetTagText("ALLPRODUCTSWITHOUTCHECKS"));
		m_ProductXML.CreateTag("ALLPRODUCTSWITHCHECKS", m_FurtherFilteringXML.GetTagText("ALLPRODUCTSWITHCHECKS"));
		m_ProductXML.CreateTag("PRODUCTSBYGROUP", m_FurtherFilteringXML.GetTagText("PRODUCTSBYGROUP"));
		m_ProductXML.CreateTag("PRODUCTGROUP", m_FurtherFilteringXML.GetTagText("PRODUCTGROUP"));
		m_ProductXML.CreateTag("DISCOUNTEDPRODUCTS", m_FurtherFilteringXML.GetTagText("DISCOUNTEDPRODUCTS"));
		m_ProductXML.CreateTag("DISCOUNTEDPERIOD", m_FurtherFilteringXML.GetTagText("DISCOUNTEDPERIOD"));
		m_ProductXML.CreateTag("FIXEDRATEPRODUCTS", m_FurtherFilteringXML.GetTagText("FIXEDRATEPRODUCTS"));
		m_ProductXML.CreateTag("FIXEDRATEPERIOD", m_FurtherFilteringXML.GetTagText("FIXEDRATEPERIOD"));
		m_ProductXML.CreateTag("STANDARDVARIABLEPRODUCTS", m_FurtherFilteringXML.GetTagText("STANDARDVARIABLEPRODUCTS"));
		m_ProductXML.CreateTag("CAPPEDFLOOREDPRODUCTS", m_FurtherFilteringXML.GetTagText("CAPPEDFLOOREDPRODUCTS"));
		m_ProductXML.CreateTag("CAPPEDFLOOREDPERIOD", m_FurtherFilteringXML.GetTagText("CAPPEDFLOOREDPERIOD"));
		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration */%>
		m_ProductXML.CreateTag("MORTGAGEPRODUCTCODE", m_FurtherFilteringXML.GetTagText("PRODUCTOVERRIDECODE"));
		<%/*BG		28/06/02	SYS4767 MSMS/Core integration END*/%>
		<% /* BMIDS00939 MDC 15/11/2002 */ %>
		m_ProductXML.CreateTag("MANUALPORTEDLOANIND", m_FurtherFilteringXML.GetTagText("MANUALPORTEDLOANIND"));
		<% /* BMIDS00939 MDC 15/11/2002 - End */ %>
	
		<% /* EP2_8 25Oct06 - Set Flexible/Non-Flexible flags if not already set. */ %>
		if (m_CalculatedFiltering == "-1")
		{	m_ProductXML.SetTagText("FLEXIBLENONFLEXIBLEPRODUCTS", m_FurtherFilteringXML.GetTagText("FLEXIBLENONFLEXIBLEPRODUCTS"));
			m_ProductXML.SetTagText("FLEXIBLEIND", m_FurtherFilteringXML.GetTagText("FLEXIBLEIND"));
		}
	}
	
	if(m_sApplicationMode == "Cost Modelling")m_ProductXML.RunASP(document,"AQFindMortgageProducts.asp");
	else m_ProductXML.RunASP(document, "QQFindMortgageProducts.asp");
	
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = m_ProductXML.CheckResponse(sErrorArray);

	if(sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		alert("No suitable products could be found that match the search criteria");
		scTable.clear();
		m_ListEmpty = true;
		sErrorArray = null;
	}
	else if (sResponseArray[0])
	{
		if(m_ProductXML.SelectTag(null,"MORTGAGEPRODUCTLIST") != null)
			return m_ProductXML.GetAttribute("TOTAL");
	}	
}

<% /* SYS1051 Function to clear the contents of the product table and set the button state */ %>
function clearProductTable()
{
	scTable.clear();
	m_ListEmpty = true;
	frmScreen.btnDetails.disabled = true;
	frmScreen.btnIncentives.disabled = true;
	frmScreen.btnOK.disabled = true;
}
<% /* SYS1051 End */ %>

<% /* START: MAR350 - Maha T */ %>
function ShowHideFormItems()
{
	<% /* MAR350 - MAHA T (Show/Hide LoanComponentToBeRetainedInd radio group) */ %>
	var ValidationList = new Array(1);
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	ValidationList[0] = "PSW";
		
	if(XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, ValidationList))
	{
		scScreenFunctions.ShowCollection(spnLoanComponentToBeRetained);
	}
	else
	{
		scScreenFunctions.HideCollection(spnLoanComponentToBeRetained);
	}
	XML = null;
}

<% //SW 20/06/2006 EP771 %>
function PopulateRepaymentVehicleCombo()
{
	var XMLRepaymentVehicle = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();		
	var XMLRepaymentVehicle2 = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();								
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	var sGroupList = new Array("RepaymentVehicle");

	XML.GetComboLists(document, sGroupList);

	var strValidationType = "&" + frmScreen.cboRepaymentType.options[frmScreen.cboRepaymentType.selectedIndex].ValidationType1;

	XMLRepaymentVehicle = XML.GetComboListXML("RepaymentVehicle");
	XMLRepaymentVehicle2.LoadXML(XMLRepaymentVehicle.xml);

	<%/* Create variable to hold default combo value*/%>
	var iDefault = 0;
	
	<%/* Iterate through XMLTypeOfBuyerNewLoan2...*/%>
	XMLRepaymentVehicle2.CreateTagList("LISTENTRY");
	for(var nLoop = 0;XMLRepaymentVehicle2.SelectTagListItem(nLoop) != false;nLoop++)
	{
		if (XMLRepaymentVehicle2.GetTagText("VALIDATIONTYPE") != strValidationType && XMLRepaymentVehicle2.GetTagText("VALIDATIONTYPE") !="N")
		{
			<%/* Remove 'First Time'*/%>
			XMLRepaymentVehicle2.RemoveActiveTag();
		}
	}
		
	<%/* Populate combo using XMLTypeOfBuyerNewLoan2*/%>
	XML.PopulateComboFromXML(document,frmScreen.cboRepaymentVehicle,XMLRepaymentVehicle2.XMLDocument,false);	
}

<% /* EP2_8 - New Method = AShaw - 19Oct06 */ %>
function InvalidMortgageTerm()
{
	var NewTerm = 0;
	var lYears = 0;
	var lMonths = 0;
	if (frmScreen.txtTermYears.value != "") 
		lYears = parseInt(frmScreen.txtTermYears.value);
	if (frmScreen.txtTermMonths.value != "") 
		lMonths = parseInt(frmScreen.txtTermMonths.value);
	NewTerm = (lYears * 12 ) + lMonths;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	if( NewTerm > m_sAdditionalTerm && XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("F")) && XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("M")))
		return 1;
	else
		return 0;
}
<% /* END: MAR350 */ %>

<% /* EP2_55 - New Methods = AShaw - 17Nov06 */ %>
function frmScreen.optLoanComponentToBePortedYes.onclick()
{
	ChangeLoanCompToBePortedInd (true);
}

function frmScreen.optLoanComponentToBePortedNo.onclick()
{
	ChangeLoanCompToBePortedInd (false);
}

function frmScreen.optLoanComponentToBeRetainedYes.onclick()
{
	ChangeLoanCompToBeRetaindedInd (true);
}
function frmScreen.optLoanComponentToBeRetainedNo.onclick()
{
	ChangeLoanCompToBeRetaindedInd (false);
}


function ChangeLoanCompToBePortedInd(bflag)
{
// Do nothing
}

function ChangeLoanCompToBeRetaindedInd(bToBeRetained)
{
	// Existing values.
	m_subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE");
	var ExistingRetainedInd = m_subQuoteXML.GetTagBoolean("PRODUCTSWITCHRETAINPRODUCTIND");
	
	<% /* PSC 13/02/2007 EP2_1314 - Start */ %>
	if (ExistingRetainedInd == true)
	{	
		<%/* EP2_1750 MAH 09/03/2007
		frmScreen.cboRepaymentType.disabled = bToBeRetained;
		frmScreen.cboRepaymentVehicle.disabled = bToBeRetained;
		frmScreen.txtRepaymentVehicleMonthlyCost.disabled = bToBeRetained;*/%>
		frmScreen.btnFurtherFiltering.disabled = bToBeRetained;
		frmScreen.btnSearch.disabled = bToBeRetained;
	}
	else // ExistingRetainedInd == false
	{	
		<%/* EP2_1750
		frmScreen.cboPurposeOfLoan.disabled = bToBeRetained;
		frmScreen.txtTermMonths.disabled = bToBeRetained;
		frmScreen.txtTermYears.disabled = bToBeRetained;
		frmScreen.cboRepaymentType.disabled = bToBeRetained;
		frmScreen.txtLoanAmount.disabled = bToBeRetained;*/%>
		frmScreen.btnFurtherFiltering.disabled = bToBeRetained;
		frmScreen.btnSearch.disabled = bToBeRetained;
		<%/* EP2_1750
		frmScreen.cboSubPurposeOfLoan.disabled = bToBeRetained;
		frmScreen.cboRepaymentVehicle.disabled = bToBeRetained;
		frmScreen.txtRepaymentVehicleMonthlyCost.disabled = bToBeRetained;*/%>
	}
	<% /* PSC 13/02/2007 EP2_1314 - End */ %>
}

function PopulateNPurposeOfLoanCombo()
{
	// EP2_55 Set Defaults values for Purpose of Loan.
	//Get the TypeOfApplication from the ApplicationFactFind table.
	var sTypeOfMortgage;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_sRequestAttributes, null);
	XML.CreateActiveTag("APPLICATIONFACTFIND");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.RunASP(document,"GetApplicationFFData.asp");
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null,"APPLICATIONFACTFIND");
		sTypeOfMortgage = XML.GetTagText("TYPEOFAPPLICATION");
	}

	// Create Variables
	var XMLPurposeOfLoan = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();		
	var XMLPurposeOfLoan2 = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();								
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	// Create list
	var sGroupList = new Array("PurposeOfLoanNew");

	//Get list
	XML.GetComboLists(document, sGroupList);

	// Get values from list
	XMLPurposeOfLoan = XML.GetComboListXML("PurposeOfLoanNew");
	XMLPurposeOfLoan2.LoadXML(XMLPurposeOfLoan.xml);

	// Set default value
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sValidationValue = "";          // Validation for selected value.
	var sValidationDefaultValue = "";	// Default for selected value.
	var sDefaultValue;					// Default for selected value.
	var bcontinue = true;
	// Check NP
	if( XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('NP')) == true)
	{
		sValidationValue = "NP";
		sValidationDefaultValue = "NPD";
		bcontinue = false;
	}

	// Check PSW
	if(bcontinue == true && XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW')) == true)
	{
		sValidationValue = "PSW";
		sValidationDefaultValue = "PSWD";
		bcontinue = false;
	}
	
	// Check TOE
	if(bcontinue == true &&  XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE')) == true)
	{
		sValidationValue = "TOE";
		sValidationDefaultValue = "TOED";
		bcontinue = false;
	}
	// Others, just N.
	if(bcontinue == true &&  XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('N')) == true)
	{
		sValidationValue = "N";
	}
	
	// Iterate through the list
	XMLPurposeOfLoan2.CreateTagList("LISTENTRY");
	for(var nLoop = 0;XMLPurposeOfLoan2.SelectTagListItem(nLoop)!= false; nLoop++)
	{
		//EP2_1476 Can have multiple validation types here, 
		//don't delete unless we've checked them all
		var bRequired = false;
		var svalue = XMLPurposeOfLoan2.ActiveTag.selectSingleNode("VALUEID").text
		var lstValidationType = XMLPurposeOfLoan2.ActiveTag.selectNodes("VALIDATIONTYPELIST/VALIDATIONTYPE");
		for(var nLoop2 = 0; nLoop2 < lstValidationType.length ;nLoop2++)
		{
			var sItem = lstValidationType.item(nLoop2).text;
			// Default Value?
			if (sItem == sValidationDefaultValue ||sItem == sValidationValue)
			{
				bRequired = true;
				break;
			}

		}
		if (!bRequired)
		{
			// Remove
			XMLPurposeOfLoan2.RemoveActiveTag();
		}
	}	

	<%/* Populate combo using XMLPurposeOfLoan*/%>
		XML.PopulateComboFromXML(document, frmScreen.cboPurposeOfLoan,XMLPurposeOfLoan2.XMLDocument,true);

	// Set default value.
	frmScreen.cboPurposeOfLoan.value = sDefaultValue;

}

<% /* END EP2_55 - New Methods = AShaw - 17Nov06 */ %>
<%/*EP2_444 Start*/%>
function frmScreen.cboRepaymentVehicle.onchange()
{
	if(m_bDataRead == false)
	{
		frmScreen.txtRepaymentVehicleMonthlyCost.value = "";
	}
	m_bDataRead = false;
}
<%/*EP2_444 End*/%>

<% /* PSC 30/01/2007 EP2_1100 - Start */ %>
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
<% /* PSC 30/01/2007 EP2_1100 - End */ %>


-->
		</script>
	</body>
</html>
