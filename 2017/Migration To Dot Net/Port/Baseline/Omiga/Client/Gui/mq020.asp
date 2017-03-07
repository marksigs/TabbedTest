<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
	<% /*
Workfile:      mq020.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Add/Edit loan components screen  THIS IS A POP-UP SCREEN
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
HMA		20/05/05	Created (MAR18)
HMA     07/09/05	Temporarily remove MSGAlert    

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specifc History:
Prog	Date	   AQR. NO 		Description
JJ		30/09/2005 MAR58		cboSubPurposeOfLoan textbox and label removed.
HMA		08/11/2005 MAR162       Enable Search button to allow re-search.
HMA     11/11/2005 MAR351       Initialise ProductSwitchRetainProductInd
PE		20/02/2006 MAR1185		Cursor stays as an hourglass
MAH		05/04/2006 MAR1566		OK button enabled when empty table clicked
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specifc History:
Prog	Date	   AQR. NO 		Description
SW		21/06/2006 EP771		Added cboRepaymentVehicl and functionality
PB      19/06/2006 EP1089		Merge MAR1833 - Correct searching after Further Filtering.
AShaw	08/11/2006	EP2_8		Incorporate spec changes to Initialise and SearchForMortgageProducts code.
								Extra load params (Add/Edit and new search params).					
								Passes params to and from MQ010.
GHun	21/11/2006 EP2_123		Changed PopulateScreen to read OriginalTerm* as attributes rather than elements
AShaw	17/11/2006	EP2_55		New methods when changing Ported/Retained options.
								Initialise method changing.	 
AShaw	19/12/2006	EP2_55	    Revisited to add more changes for EP2_55
						  	    Altered Ported Option code. (Left code in place, but now does nothing onclick).
INR		15/01/2007	EP2_697		Added Repayment vehicle and reinstated cboSubPurposeOfLoan
PE		20/01/2007	EP2_895		Small change for population of purpose of loan.
AShaw	24/01/2007	EP2_889		Do NOT disabled Term fields when Second/subsequent components added.
PSC		30/01/2007	EP2_1100	Use remaining term rather than outstanding term
PSC		30/01/2007	EP2_1110	Calculate remaining term in edit mode too
PSC		1302/2007	EP2_1314	Correct enabling and disabling of fields
INR		19/02/2007	EP2_1476	PurposeOfLoan combo entries can have multiple validation types
LDM		27/02/2007	EP2_1567	Remove porting questionif type of app is not Porting
SR		05/03/2007	EP2_1644	modified width and cor-ordinates of scScrollPlus
MAH		08/03/2007	EP2_1751	Added scMathFunctions to Assist rendering decimal places
MAH		09/03/2007	EP2_1750	Disabled Repayment type/vehicle selectivity when Product Switch
MAH		11/03/2007	EP2_1714	Add Manual Port indicator to logic for displaying the text, checked the product switch logic works OK
MAH		12/03/2007	EP2_1703	Checked MORTGAGELOANGUID on Ported loan
MAH		15/03/2007	EP2_1703	Set default where Ported loan has extra new loancompnonent 
PSC		23/03/2007	EP2_1622	Set first interest rate correctly for fixed rates
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
	<head>
		<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
		<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
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
		<span style="LEFT: 306px; POSITION: absolute; TOP: 365px">
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
						<select id="cboPurposeOfLoan" style="WIDTH: 190px" class="msgCombo" NAME="cboPurposeOfLoan">
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
						<input id="txtLoanAmount" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgTxt"
							NAME="txtLoanAmount">
					</span>
				</span>
				<span style="LEFT: 295px; POSITION: absolute; TOP: 40px" class="msgLabel">
					Repayment Type
					<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
						<select id="cboRepaymentType" style="WIDTH: 190px" class="msgCombo" NAME="cboRepaymentType">
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
				<%/*EP2_697*/%>
				<span style="LEFT: 295px; POSITION: absolute; TOP: 100px" class="msglabel">
					Repayment Vehicle<br>Monthly Cost
					<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
						<input id="txtRepaymentVehicleMonthlyCost" NAME="txtRepaymentVehicleMonthlyCost" maxlength="10"
							style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
					</span>
				</span>
				<span style="LEFT: 210px; POSITION: absolute; TOP: 35px">
					<input id="btnSplit" value="Split" type="button" style="WIDTH: 60px" class="msgButton"
						NAME="btnSplit">
				</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
					Term
					<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
						<input id="txtTermYears" maxlength="3" style="POSITION: absolute; WIDTH: 40px" class="msgTxt"
							NAME="txtTermYears">
						<span style="LEFT: 48px; POSITION: absolute; TOP: 2px" class="msgLabel">Years</span>
					</span>
					<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
						<input id="txtTermMonths" maxlength="3" style="POSITION: absolute; WIDTH: 40px" class="msgTxt"
							NAME="txtTermMonths">
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
					<input id="btnSearch" value="Search" type="button" style="WIDTH: 60px" class="msgButton"
						NAME="btnSearch">
				</span>
				<span style="LEFT: 500px; POSITION: absolute; TOP: 130px">
					<input id="btnFurtherFiltering" value="Further Filtering" type="button" style="WIDTH: 100px"
						class="msgButton" NAME="btnFurtherFiltering">
				</span>
			</div>
			<div id="divResults" style="HEIGHT: 245px; LEFT: 10px; POSITION: absolute; TOP: 160px; WIDTH: 604px"
				class="msgGroup">
				<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
					<strong>Available Mortgage Products</strong>
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
							<tr id="rowM1">
								<td width="20%" class="TableTopLeft">&nbsp;</td>
								<td width="10%" class="TableTopCenter">&nbsp;</td>
								<td width="30%" class="TableTopCenter">&nbsp;</td>
								<td width="20%" class="TableTopCenter">&nbsp;</td>
								<td width="10%" class="TableTopCenter">&nbsp;</td>
								<td class="TableTopRight">&nbsp;</td>
							</tr>
							<tr id="rowM2">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowM3">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowM4">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowM5">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowM6">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowM7">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowM8">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowM9">
								<td width="20%" class="TableLeft">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td width="30%" class="TableCenter">&nbsp;</td>
								<td width="20%" class="TableCenter">&nbsp;</td>
								<td width="10%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowM10">
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
							<tr id="rowS1">
								<td width="10%" class="TableTopLeft">&nbsp;</td>
								<td width="45%" class="TableTopCenter">&nbsp;</td>
								<td width="15%" class="TableTopCenter">&nbsp;</td>
								<td width="15%" class="TableTopCenter">&nbsp;</td>
								<td class="TableTopRight">&nbsp;</td>
							</tr>
							<tr id="rowS2">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowS3">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowS4">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowS5">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowS6">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowS7">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowS8">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowS9">
								<td width="10%" class="TableLeft">&nbsp;</td>
								<td width="45%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td width="15%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td>
							</tr>
							<tr id="rowS10">
								<td width="10%" class="TableBottomLeft">&nbsp;</td>
								<td width="45%" class="TableBottomCenter">&nbsp;</td>
								<td width="15%" class="TableBottomCenter">&nbsp;</td>
								<td width="15%" class="TableBottomCenter">&nbsp;</td>
								<td class="TableBottomRight">&nbsp;</td>
							</tr>
						</table>
					</div>
				</div>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 210px">
					<input id="btnDetails" value="Details" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
				<span style="LEFT: 70px; POSITION: absolute; TOP: 210px">
					<input id="btnIncentives" value="Incentives" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
				<span style="LEFT: 4px; POSITION: absolute; TOP: 240px">
					<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
					<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">
						<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
					</span>
				</span>
			</div>
		</form>
		<!-- File containing field attributes (remove if not required) -->
		<!--  #include FILE="attribs/mq020attribs.asp" -->
		<!-- Specify Code Here -->
		<script language="JScript">
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
var m_sApplicationDate = ""; 
var m_sMortgageSubQuoteNumber = "";
var m_sApplicationMode = "";
var m_sUserId = "";
var m_sUserType = "";
var m_sUnitId = "";
var m_sLoanComponentSeqNum = "";
var m_sidMortgageApplicationValue = "";
var m_sNumberOfLoanComponents = "";
var m_sDistributionChannelId = "";
var m_sCurrency = "";
var m_sRequestAttributes = "";
var m_sInterestOnlyAmount = "";
var m_sCapitalInterestAmount = "";
var m_sOriginalIncentivesXML = null;
var m_sNewIncentivesXML = "";
var m_sOriginalProductCode = "";
var m_sNewProductCode = "";
var m_sOriginalProductStartDate = "";
var m_sNewProductStartDate = "";
var m_sMultipleLender = "0";
var scScreenFunctions;
var m_ListEmpty = true;
var m_BaseNonPopupWindow = null;
var m_LoanIndicator = 0;
var sPurposeListName = "";
var m_sTermYears = "";
var m_sTermMonths = "";
var m_bNewLoanComponent = false;  //MAR162
var m_bFurtherFiltering = false;   <% /* EP1089/MAR1833 */ %>
<% /* EP2_8 26Oct06 New variables */ %>
var m_sAdditionalTerm = "0";		
var m_AddOrEdit = "";				
var m_sMQ010XML = "";				
var m_MQ010XML = null;	
var m_CalculatedFiltering = "";
<% /* PSC 30/01/2007 EP2_1100 */ %>
var m_iNumOfLoanComponents = 0;			

var scClientScreenFunctions;
function window.onload()
{
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
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
	m_sApplicationDate = sArgArray[12]; 
	m_AddOrEdit = sArgArray[13]; // EP2_8 19Oct06
	m_sMQ010XML = sArgArray[14]; // AS EP2_8 08Nov06

	SetMasks();
	Validation_Init();
	Initialise();
	if (m_sReadOnly == "1")
		frmScreen.btnOK.disabled = true;	
		
	window.returnValue = null;   // return null if OK is not pressed
	
	ClientPopulateScreen();
}
function Initialise()
{
	m_subQuoteXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_subQuoteXML.LoadXML(m_sSubQuoteXML);
	
	<%// EP2_8 - Load Param values and set flag for further filtering %>
	m_MQ010XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_MQ010XML.LoadXML(m_sMQ010XML);
	m_MQ010XML.SelectTag(null, "PARAMS");
	m_CalculatedFiltering = m_MQ010XML.GetTagText("FLEXIBLEPRODUCTS");

	m_subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE");
	m_sApplicationNumber = m_subQuoteXML.GetTagText("APPLICATIONNUMBER");
	m_sApplicationFactFindNumber = m_subQuoteXML.GetTagText("APPLICATIONFACTFINDNUMBER");
	m_sMortgageSubQuoteNumber = m_subQuoteXML.GetTagText("MORTGAGESUBQUOTENUMBER");
	
	if ((m_sLoanComponentSeqNum == null) && (parseInt(m_sNumberOfLoanComponents) > 0 ))
	{
		<% /* Add Mode and a Loan Component has already been added
		      Default the term to be the same as the first Loan Component */ %>
		      
		m_subQuoteXML.CreateTagList("LOANCOMPONENT");
		m_subQuoteXML.SelectTagListItem(0);

		m_sTermYears = m_subQuoteXML.GetTagText("TERMINYEARS");
		m_sTermMonths = m_subQuoteXML.GetTagText("TERMINMONTHS");
		
		frmScreen.txtTermYears.value = m_sTermYears;
		frmScreen.txtTermMonths.value = m_sTermMonths;
		
		<% /* Disable these fields */ %>
//		EP2_889 - Remove this code.		
//		frmScreen.txtTermMonths.disabled = true;
//		frmScreen.txtTermYears.disabled = true;
	}
	
	PopulateScreen();
	if(m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnFurtherFiltering.disabled = true;
		frmScreen.btnSearch.disabled = true;
	}
}
function PopulateScreen()
{
	var bShowLoanCompToBePortedQuestion = true;
	var bShowLoanCompToBeRetainedQuestion = true;
	var bContinue = true;
	
	PopulateCombos();
	
	m_sMultipleLender = "0";
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.RunASP(document, "GetMultipleLender.asp");
	if (XML.IsResponseOK() == true)
	{
		m_sMultipleLender = XML.GetTagText("MULTIPLELENDER");
	}

	<% /* LDM 27/02/07 EP2_1567 dont show the porting question if the case is not a porting one.
		  moved to the top of the function to stop question appearing underneath the hardcoded popup  
		  Check NP (validation type for "new loan - ported")*/ %>
	<%/* EP2_1714 MAH 11/03/2007 Add the "MANUALPORTEDLOANIND is set" validation
	if(XML.IsInComboValidationList(document, 'TypeOfMortgage', m_sidMortgageApplicationValue , Array('NP')) == false)*/%>
	var bIsNP = XML.IsInComboValidationList(document, 'TypeOfMortgage', m_sidMortgageApplicationValue , Array('NP'))
	var sManPortInd = m_subQuoteXML.GetTagText("MANUALPORTEDLOANIND")
	if((bIsNP == false) || ((bIsNP == true) && (sManPortInd != null) && (sManPortInd == "0")))
	{
		scScreenFunctions.HideCollection(spnLoanComponentToBePorted);
		bShowLoanCompToBePortedQuestion = false;
	}
	<% /* Check PSW (validation type for "product switch") */ %>
	if(XML.IsInComboValidationList(document, 'TypeOfMortgage', m_sidMortgageApplicationValue, Array('PSW')) == false)
	{
		scScreenFunctions.HideCollection(spnLoanComponentToBeRetained);
		bShowLoanCompToBeRetainedQuestion = false;
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
		m_NewLoanCompXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_NewLoanCompXML.CreateRequestTagFromArray(m_sRequestAttributes, null);
		m_NewLoanCompXML.CreateActiveTag("LOANCOMPONENTDETAILS");
		m_NewLoanCompXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		m_NewLoanCompXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		if( parseInt(m_sNumberOfLoanComponents) > 0 )m_NewLoanCompXML.CreateTag("NUMBEROFCOMPONENTS", m_sNumberOfLoanComponents);
		else m_NewLoanCompXML.CreateTag("NUMBEROFCOMPONENTS", "");
		m_NewLoanCompXML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber);
		m_NewLoanCompXML.CreateTag("AMOUNTREQUESTED", m_subQuoteXML.GetTagText("AMOUNTREQUESTED"));
		m_NewLoanCompXML.RunASP(document,"AQGetDefaultsForNewLoanComponent.asp");
		
		if(m_NewLoanCompXML.IsResponseOK() == false)
		{
			frmScreen.btnDetails.disabled = true;
			frmScreen.btnFurtherFiltering.disabled = true;
			frmScreen.btnIncentives.disabled = true;
			frmScreen.btnOK.disabled = true;
			frmScreen.btnSearch.disabled = true;
			frmScreen.btnSplit.disabled = true;
			frmScreen.btnCancel.disabled = false;
			frmScreen.btnCancel.focus();
			bContinue = false;
		}
	}
	
		<% /* AShaw EP2_8 26Oct06 - Calculate Additional term */%>
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
							var lYears = XML.GetAttribute("ORIGINALTERMYEARS");	<% /* EP2_123 GHun */ %>
							if (lYears == "") lYears = "0";
							var lMonths = XML.GetAttribute("ORIGINALTERMMONTHS"); <% /* EP2_123 GHun */ %>
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
		if(m_sLoanComponentSeqNum != null)
		{
			<%/* Locate the loan component within subquoteXML which corresponds to the seq num */%>			
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
					//m_Indicator = m_subQuoteXML.GetTagText("MANUALPORTEDLOANIND");
					<%/* EP2_1703 MAH start
					m_Indicator = true;*/%>
					if(m_subQuoteXML.GetTagText("MORTGAGELOANGUID") != null)
					{
						m_Indicator = true;
					}
					else
					{
						m_Indicator = false;
					}
					<%/* EP2_1703 MAH End */%>
					scScreenFunctions.SetRadioGroupValue(frmScreen, "LoanComponentToBePortedInd", m_Indicator);
					// Disable the buttons.
					frmScreen.optLoanComponentToBePortedYes.disabled = true;
					frmScreen.optLoanComponentToBePortedNo.disabled = true;
					// Do the bizz.
					//if (frmScreen.optLoanComponentToBePortedYes.checked == true)
					{	// Disable fields.
						frmScreen.cboPurposeOfLoan.disabled = true;
						frmScreen.txtTermYears.disabled = true;
						frmScreen.txtTermMonths.disabled = true;
						frmScreen.cboRepaymentType.disabled = true;
						frmScreen.cboRepaymentVehicle.disabled = true; <%/*EP2_697*/%>
						frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true; <%/*EP_697*/%>
						frmScreen.txtLoanAmount.disabled = true;
						frmScreen.btnFurtherFiltering.disabled = true;
						frmScreen.btnSearch.disabled = true;
						<% /* PSC 13/02/2007 EP2_1314 */ %>
						frmScreen.cboSubPurposeOfLoan.disabled = true;						
					}
					//else  // Ind = False
					{	// Enable all fields 
					
						// EP2_895 - 20/01/2007
						//frmScreen.cboPurposeOfLoan.disabled = false;
						//frmScreen.txtTermYears.disabled = false;
						//frmScreen.txtTermMonths.disabled = false;
						//frmScreen.cboRepaymentType.disabled = false;
						//frmScreen.cboRepaymentVehicle.disabled = false; <%/*EP2_697*/%>
						//frmScreen.txtRepaymentVehicleMonthlyCost.disabled = false; <%/*EP_697*/%>
						//frmScreen.txtLoanAmount.disabled = false;
						//frmScreen.btnFurtherFiltering.disabled = false;
						//frmScreen.btnSearch.disabled = false;
					}
				}

				// Check PSW (validation type for "product switch")
				if(bShowLoanCompToBeRetainedQuestion == true)
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
					frmScreen.cboRepaymentVehicle.disabled = true; <%/*EP_697*/%>
					frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true; <%/*EP_697*/%>
					frmScreen.txtLoanAmount.disabled = true;
					frmScreen.btnFurtherFiltering.disabled = true;
					frmScreen.btnSearch.disabled = true;
				}
				
				frmScreen.cboPurposeOfLoan.value = m_subQuoteXML.GetTagText("PURPOSEOFLOAN");
				frmScreen.txtLoanAmount.value = m_subQuoteXML.GetTagText("LOANAMOUNT");
				frmScreen.cboRepaymentType.value = m_subQuoteXML.GetTagText("REPAYMENTMETHOD");

				<%/*EP2_697*/%>
				var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
				var ValidationList = new Array(2);
				ValidationList[0] = "P";
				ValidationList[1] = "I"

				<% // SW 20/06/2006 EP771 %>
				if(frmScreen.cboRepaymentType.value != "" && frmScreen.cboRepaymentType.value != null
					&& XML.IsInComboValidationList(document,"RepaymentType", frmScreen.cboRepaymentType.value, ValidationList))<%/*++ EP2_697*/%>
				{
					<%/* EP2_1750 
					frmScreen.cboRepaymentVehicle.disabled = false;*/%>
					if(!bShowLoanCompToBeRetainedQuestion){frmScreen.cboRepaymentVehicle.disabled = false;}
					PopulateRepaymentVehicleCombo();
					m_bDataRead = true;<%/*EP2_697*/%>
					frmScreen.cboRepaymentVehicle.value  = m_subQuoteXML.GetTagText("REPAYMENTVEHICLE");
					frmScreen.txtRepaymentVehicleMonthlyCost.value = m_subQuoteXML.GetTagText("REPAYMENTVEHICLEMONTHLYCOST");<%/*EP2_697*/%>
				}
				else
				{
					frmScreen.cboRepaymentVehicle.value="";<%/*EP2_697*/%>
					frmScreen.cboRepaymentVehicle.disabled = true;
					frmScreen.txtRepaymentVehicleMonthlyCost.value = "";<%/*EP2_697*/%>
					frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true; <%/*EP2_697*/%>
				}
				XML = null;<%/*EP2_697*/%>
				ValidationList = null;<%/*EP2_697*/%>

				<% /* pull back two new fields from XML - assign one to variable */ %>
				//START: (MAR58) Code commented by Joyce Joseph on 30-Sep-2005				
				<% /* EP2_697 */ %>
				frmScreen.cboSubPurposeOfLoan.value = m_subQuoteXML.GetTagText("SUBPURPOSEOFLOAN");
				// EP2_55 Moved code
				//m_LoanIndicator = m_subQuoteXML.GetTagText("MANUALPORTEDLOANIND");
				//scScreenFunctions.SetRadioGroupValue(frmScreen, "LoanComponentToBePortedInd", m_subQuoteXML.GetTagText("MANUALPORTEDLOANIND"));
				
				if(m_subQuoteXML.GetTagAttribute("REPAYMENTMETHOD", "TEXT") == "Part and Part")
				{
					frmScreen.btnSplit.disabled = false;
					m_sInterestOnlyAmount = m_subQuoteXML.GetTagText("INTERESTONLYELEMENT");
					m_sCapitalInterestAmount = m_subQuoteXML.GetTagText("CAPITALANDINTERESTELEMENT");
				}
				else
				{
					frmScreen.btnSplit.disabled = true;
					m_sInterestOnlyAmount = "0";
					m_sCapitalInterestAmount = "0";
				}
				
				m_sTermYears = m_subQuoteXML.GetTagText("TERMINYEARS");
				m_sTermMonths = m_subQuoteXML.GetTagText("TERMINMONTHS");
				
				frmScreen.txtTermYears.value = m_sTermYears;
				frmScreen.txtTermMonths.value = m_sTermMonths;
				
<%				/* fill in the list box with the loan component mortgage product details.
				   Create a bogus m_ProductXML with the details found, as populateListBox
				   uses m_ProductXML */
%>				m_ProductXML = null;
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
				var sResolvedRate = "";
				
				var sIsFlexible = "";
				<% /* needed for the Details and Incentives buttons ... */ %> 
				var sProdDesc = "";
				var sProdCode = "";
				var sStartDate = "";
				if(null != m_subQuoteXML.SelectTag(tagActiveLoan, "NONPANELMORTGAGEPRODUCT"))
				{
					sInterestRate = m_subQuoteXML.GetTagText("INTERESTRATE1");		//from NONPANELMORTGAGEPRODUCT XML
					m_subQuoteXML.ActiveTag = tagActiveLoan;
					sProductName = m_subQuoteXML.GetTagText("PRODUCTNAME");			// from MORTGAGEPRODUCTLANGUAGE XML
					sProductNum = m_subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE");	// from MORTGAGEPRODUCTLANGUAGE XML
					sProdDesc = m_subQuoteXML.GetTagText("PRODUCTTEXTDETAILS");		// from MORTGAGEPRODUCTLANGUAGE XML
					sType = m_subQuoteXML.GetTagText("RATETYPE");					// from MORTGAGEPRODUCT XML
					sRatePeriod = m_subQuoteXML.GetTagText("INTERESTRATEPERIOD");	// from MORTGAGEPRODUCT 
					
					sProdCode = m_subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE");	// from MORTGAGEPRODUCT
					sStartDate = m_subQuoteXML.GetTagText("STARTDATE");				// from MORTGAGEPRODUCT
					
					m_subQuoteXML.CreateTagList("INTERESTRATETYPE");
					sDiscountAmount = m_subQuoteXML.GetTagText("RATE");
					sResolvedRate = m_subQuoteXML.GetTagText("RESOLVEDRATE");
					if(m_sMultipleLender == "1")
					{
						m_subQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE");
						sLenderName = m_subQuoteXML.GetTagText("NONPANELLENDERNAME"); // from MORTGAGESUBQUOTE XML
					}
				}
				else
				{
					m_subQuoteXML.ActiveTag = tagActiveLoan;
					if(m_sMultipleLender == "1")
					{
						sLenderName = m_subQuoteXML.GetTagText("LENDERNAME");			//from MORTGAGELENDER XML
					}
					sProductName = m_subQuoteXML.GetTagText("PRODUCTNAME");				// from MORTGAGEPRODUCTLANGUAGE XML
					sProductNum = m_subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE");		// from MORTGAGEPRODUCTLANGUAGE XML
					sProdDesc = m_subQuoteXML.GetTagText("PRODUCTTEXTDETAILS");			// from MORTGAGEPRODUCTLANGUAGE XML
					sType = m_subQuoteXML.GetTagText("RATETYPE");						// from MORTGAGEPRODUCT XML
					sRatePeriod = m_subQuoteXML.GetTagText("INTERESTRATEPERIOD");		// from MORTGAGEPRODUCT 
					sIsFlexible = m_subQuoteXML.GetTagText("FLEXIBLEMORTGAGEPRODUCT");	// from MORTGAGEPRODUCT XML
					sProdCode = m_subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE");		// from MORTGAGEPRODUCT
					sStartDate = m_subQuoteXML.GetTagText("STARTDATE");					// from MORTGAGEPRODUCT
					sResolvedRate = m_subQuoteXML.GetTagText("RESOLVEDRATE");
					sInterestRate = GetFirstInterestRate(tagActiveLoan);
					
					m_subQuoteXML.CreateTagList("INTERESTRATETYPE");
					sDiscountAmount = m_subQuoteXML.GetTagText("RATE");
				}
				m_ProductXML.CreateTag("MORTGAGEPRODUCTCODE", sProductNum);
				m_ProductXML.CreateTag("PRODUCTNAME", sProductName);
				m_ProductXML.CreateTag("FLEXIBLEMORTGAGEPRODUCT", sIsFlexible);				
				m_ProductXML.CreateTag("TYPE", sType);
				m_ProductXML.CreateTag("INTERESTRATEPERIOD", sRatePeriod);
				m_ProductXML.CreateTag("DISCOUNTAMOUNT", sDiscountAmount);
				m_ProductXML.CreateTag("LENDERSNAME", sLenderName);
				m_ProductXML.CreateTag("FIRSTMONTHLYINTERESTRATE", sInterestRate);
				m_ProductXML.CreateTag("PRODUCTTEXTDETAILS", sProdDesc);
				m_ProductXML.CreateTag("MORTGAGEPRODUCTCODE", sProdCode);
				m_ProductXML.CreateTag("STARTDATE", sStartDate);
				m_ProductXML.CreateTag("RESOLVEDRATE", sResolvedRate);
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
				m_sOriginalProductCode = m_subQuoteXML.GetTagText("MORTGAGEPRODUCTCODE");
				m_sOriginalProductStartDate = m_subQuoteXML.GetTagText("STARTDATE");
				m_sNewProductCode = "";
				m_sNewProductStartDate = "";
				if(m_subQuoteXML.SelectTag(tagActiveLoan, "MORTGAGEINCENTIVELIST") != null)
				{
					m_sOriginalIncentivesXML = m_subQuoteXML.ActiveTag.xml;
				}
				else
				{
					m_sOriginalIncentivesXML = "";
					m_sNewIncentivesXML = "";
				}
			}
		} 
		else
		{
			<% /* Initialise the screen to add a new loan component */ %>			
			m_bNewLoanComponent = true;   // MAR162
			<%/* EP2_1703  */%>
			frmScreen.optLoanComponentToBePortedNo.value = "1";
			frmScreen.optLoanComponentToBePortedYes.disabled = true;
			frmScreen.optLoanComponentToBePortedNo.disabled = true;
			
			m_NewLoanCompXML.SelectTag(null, "NEWLOANCOMPONENT");
			
			<% /* PSC 30/01/2007 EP2_1100 - Start */ %>
			if (m_iNumOfLoanComponents > 0)
			{
				var XMLTypeOfApp = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
 
			    if (XMLTypeOfApp.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("F")) && 
			        XMLTypeOfApp.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("M")))
				{
					if (m_sAdditionalTerm != 0)
					{
						m_sTermYears = Math.floor(m_sAdditionalTerm /12);
						m_sTermMonths = m_sAdditionalTerm % 12;
						frmScreen.txtTermYears.value = m_sTermYears;
						frmScreen.txtTermMonths.value = m_sTermMonths;
					}
					else
					{
						alert("The remaining loan term is not set on the existing account. The term has been defaulted.");
					}
				}
			}
			<% /* PSC 30/01/2007 EP2_1100 - End */ %>

			frmScreen.cboRepaymentType.value  = m_NewLoanCompXML.GetTagText("REPAYMENTTYPE");
			<%/* EP2_697*/%>
			
			var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			var ValidationList = new Array(2);
			ValidationList[0] = "P";
			ValidationList[1] = "I"

			<% // SW 20/06/2006 EP771 %>
			<%/*++ EP2_697*/%>
			if(frmScreen.cboRepaymentType.value != "" && frmScreen.cboRepaymentType.value != null
				&& XML.IsInComboValidationList(document,"RepaymentType", frmScreen.cboRepaymentType.value, ValidationList))
			{
				m_bDataRead = true;<%/*EP2_697*/%>
				<%/* EP2_1750 
				frmScreen.cboRepaymentVehicle.disabled = false;*/%>
				if(!bShowLoanCompToBeRetainedQuestion){frmScreen.cboRepaymentVehicle.disabled = false;}
				PopulateRepaymentVehicleCombo();
				frmScreen.cboRepaymentVehicle.value  = m_NewLoanCompXML.GetTagText("REPAYMENTVEHICLE");
				frmScreen.txtRepaymentVehicleMonthlyCost.value = m_NewLoanCompXML.GetTagText("REPAYMENTVEHICLEMONTHLYCOST");<%/*EP2_697*/%>
			}
			else
			{
				frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true;	<%/*EP2_697*/%>
			}
			XML = null;<%/*EP2_697*/%>
			ValidationList = null;<%/*EP2_697*/%>


			frmScreen.txtLoanAmount.value = m_NewLoanCompXML.GetTagText("DEFAULTLOANAMOUNT");
			
			if (m_sTermYears == "")
				frmScreen.txtTermYears.value = m_NewLoanCompXML.GetTagText("DEFAULTTERMYEARS");
				
			if (m_sTermMonths == "")
				frmScreen.txtTermMonths.value = m_NewLoanCompXML.GetTagText("DEFAULTTERMMONTHS");
				
			frmScreen.btnFurtherFiltering.disabled = false;
			frmScreen.btnSearch.disabled = false;
			frmScreen.btnOK.disabled = true;
			frmScreen.btnCancel.disabled = false;
			frmScreen.btnSplit.disabled = true;
			frmScreen.btnDetails.disabled = true;
			frmScreen.btnIncentives.disabled = true;
			m_sOriginalProductCode = "";
			m_sOriginalProductStartDate = "";
			m_sNewProductCode = "";
			m_sNewProductStartDate = "";
			m_sOriginalIncentivesXML = "";
			m_sNewIncentivesXML = "";
		}
	}
}
function GetFirstInterestRate( currentActiveTag )
{
	var strRate = "";
	var bSuccess = false;
<%	/* get the baserate rate for future reference */
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
			if(strRateType == "F")
			{
				<% /* PSC 23/03/2007 EP2_1622 */ %>
				strRate = strBaseRate;
				bSuccess = true;
			}
			else if(strRateType == "B")
			{
				strRate = strBaseRate;
				bSuccess = true;
			}
			else if(strRateType == "D")
			{
				<%/* EP2_1751 Added Rounding to force 2 decimal places*/%>
				strRate = scMathFunctions.RoundValue(parseFloat(strBaseRate) - parseFloat(m_subQuoteXML.GetTagText("RATE")),2);
				bSuccess = true;
			}
			else if(strRateType == "C")
			{
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
	var XMLPurposeofLoan = null;
	var XMLRepaymentType = null;
	var XMLSubPurposeOfLoan = null;
	<% // SW 20/06/2006 EP771 %>
	var XMLRepaymentVehicle = null;
	  
	var ValidationList = new Array(1);
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
<%		/* Need to remove part&part repayment type if there are more than one loancomps already 
           or we already have one and are trying to add another one */
%>		var iIndex = -1;
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
		{
			XMLPurposeOfLoan = XML.GetComboListXML(sPurposeListName);
			XML.PopulateComboFromXML(document, frmScreen.cboPurposeOfLoan,XMLPurposeOfLoan,true);
		}
		
		<% /* EP2_697 */ %>
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
		frmScreen.btnOK.disabled = false; 	
	}
	else
	{
		<% /* Only enable the buttons if the product list is not empty */ %>
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
	}
	<%  /* MARS1566 M Heys 05/04/2006*/  %>
	//frmScreen.btnOK.disabled = false;
}
function frmScreen.txtLoanAmount.onchange()
{
	m_sInterestOnlyAmount = "";
	m_sCapitalInterestAmount = "";

	if (m_ListEmpty != true)
	{	
		alert("Loan amount has changed - please re-Search available products");
		//scScreenFunctions.MSGAlert("Loan amount has changed - please re-Search available products");
		clearProductTable();	 
	}
}
function frmScreen.txtTermYears.onchange()
{
	if (m_ListEmpty != true)
	{	
		alert("Term in years has changed - please re-Search available products");
		//scScreenFunctions.MSGAlert("Term in years has changed - please re-Search available products");
		clearProductTable(); 	 
	} 
}
function frmScreen.txtTermMonths.onchange()
{
	if (m_ListEmpty != true)
	{	
		alert("Term in months has changed - please re-Search available products");
		//scScreenFunctions.MSGAlert("Term in months has changed - please re-Search available products");
		clearProductTable();	 
	}
}

function frmScreen.cboPurposeOfLoan.onchange()
{
	if (m_ListEmpty != true)
	{	
		alert("Purpose of the loan has changed - please re-Search available products");
		//scScreenFunctions.MSGAlert("Purpose of the loan has changed - please re-Search available products");
		clearProductTable();	 
	}
}

function frmScreen.cboRepaymentType.onchange()
{
	<%/*EP2_697*/%>
	m_bDataRead = false;
	frmScreen.txtRepaymentVehicleMonthlyCost.value = "";
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ValidationList = new Array(1);
	ValidationList[0] = "P";
	if( frmScreen.cboRepaymentType.value != "" && frmScreen.cboRepaymentType.value != null
		&& XML.IsInComboValidationList(document,"RepaymentType", frmScreen.cboRepaymentType.value, ValidationList))
		frmScreen.btnSplit.disabled = false;
	else
	{
		frmScreen.btnSplit.disabled = true;
		m_sInterestOnlyAmount = "";
		m_sCapitalInterestAmount = "";
	}
	XML = null;
	ValidationList = null;
	if (m_ListEmpty != true)
	{
		alert("Repayment type changed - please re-Search available products");
		//scScreenFunctions.MSGAlert("Repayment type changed - please re-Search available products");
		clearProductTable();	 
	}
	
	<%/*++ EP2_697*/%>
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ValidationList = new Array(2);
	ValidationList[0] = "P";
	ValidationList[1] = "I"
	<% // SW 20/06/2006 EP771 %>
	<%/*++ EP2_697*/%>
	if(frmScreen.cboRepaymentType.value != "" && frmScreen.cboRepaymentType.value != null
		&& XML.IsInComboValidationList(document,"RepaymentType", frmScreen.cboRepaymentType.value, ValidationList))
	{
		<%/* EP2_1750 
		frmScreen.cboRepaymentVehicle.disabled = false;*/%>
		if(XML.IsInComboValidationList(document, 'TypeOfMortgage', m_sidMortgageApplicationValue, Array('PSW')) == false)
		{
		frmScreen.cboRepaymentVehicle.disabled = false;
		frmScreen.txtRepaymentVehicleMonthlyCost.disabled = false;
		}
		PopulateRepaymentVehicleCombo();
	}
	else
	{
		frmScreen.cboRepaymentVehicle.value="";<%/* EP2_697*/%>
		frmScreen.cboRepaymentVehicle.disabled = true;
		frmScreen.txtRepaymentVehicleMonthlyCost.value = "";<%/* ++ EP2_697*/%>
		frmScreen.txtRepaymentVehicleMonthlyCost.disabled = true;<%/* ++ EP2_697*/%>	
	}
	XML = null;<%/* ++ EP2_697*/%>
	ValidationList = null;<%/* ++ EP2_697*/%>

}
function frmScreen.btnSearch.onclick()
{	
	frmScreen.style.cursor = "wait";
	if(frmScreen.onsubmit())
	{
		// EP2_8 - Check for invalid Borrowing period
		if (InvalidMortgageTerm() == 1) //Period entered is longer than existing period.
		{
			frmScreen.style.cursor = "default";	<% /* EP2_123 GHun */ %>
			
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
		
		<% /* check Part & Part values if neccessary */ %>
		var ValidationList = new Array(1);
		ValidationList[0] = "P";
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		if( XML.IsInComboValidationList(document,"RepaymentType", frmScreen.cboRepaymentType.value, ValidationList))
		{
			if(m_sInterestOnlyAmount == "" && m_sCapitalInterestAmount == "")
			{
				<% // MAR1185 %>
				frmScreen.style.cursor = "default";
				
				alert("Please enter the Part and Part split before searching for products.");
				//scScreenFunctions.MSGAlert("Please enter the Part and Part split before searching for products.");
				frmScreen.btnSplit.focus();
				return;
			}
		}
		
		<% /* check that the loan amount is > 0  */ %>
		if( parseInt(frmScreen.txtLoanAmount.value,10) <=0)
		{
			<% // MAR1185 %>
			frmScreen.style.cursor = "default";
			
			alert("Please enter a loan amount before searching for products");
			//scScreenFunctions.MSGAlert("Please enter a loan amount before searching for products");
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
%>		clearProductTable();	
		m_ProductXML.SelectTag(null, "RESPONSE");
		alert(m_ProductXML.GetTagText("DESCRIPTION"));
		//scScreenFunctions.MSGAlert(m_ProductXML.GetTagText("DESCRIPTION"));
	}
}
function DisplayProducts(iStart)
{
	var sProductName			= null;
	var sLenderName				= null;
	var sRateType				= null;
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
    if (iNumberOfProducts > 0) 
		m_ListEmpty = false;
		
	for (var nProduct = 0; nProduct < iNumberOfProducts && nProduct < m_iTableLength; nProduct++)
	{
		m_ProductXML.ActiveTagList = tagPRODUCTLIST;
		if (m_ProductXML.SelectTagListItem(nProduct + iStart) == true)
		{
			sProductName = m_ProductXML.GetTagText("PRODUCTNAME");
			blnFlexibleMortgage = m_ProductXML.GetTagBoolean("FLEXIBLEMORTGAGEPRODUCT");
			if (blnFlexibleMortgage)
				sFlexibleMortgage = "Y";
			else
				sFlexibleMortgage = "";
				
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
			if(sRatePeriod == "-1") 
				sRatePeriod = ""
	
			var sProductCode = m_ProductXML.GetTagText("MORTGAGEPRODUCTCODE");
			var sStartDate = m_ProductXML.GetTagText("STARTDATE");
			var sResolvedRate = m_ProductXML.GetTagText("RESOLVEDRATE");
			tagPRODUCTDETAILS = m_ProductXML.ActiveTag;

			if (m_sMultipleLender == "1")
			{
				sLenderName = m_ProductXML.GetTagText("LENDERSNAME");
			}
			
			blnNonPanelLender = m_ProductXML.GetTagBoolean("NONPANELLENDEROPTION");
			
			if (blnNonPanelLender)
				sInterestRate = ""
			else
				sInterestRate = m_ProductXML.GetTagText("FIRSTMONTHLYINTERESTRATE");
				
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
	ArrayArguments[1] = m_sInterestOnlyAmount;
	ArrayArguments[2] = m_sCapitalInterestAmount;
	ArrayArguments[3] = m_sReadOnly;
	ArrayArguments[4] = m_sCurrency;
						
	sReturn = scScreenFunctions.DisplayPopup(window, document, "cm160.asp", ArrayArguments, 505, 225);
	if (sReturn != null)
	{						
		m_sInterestOnlyAmount = sReturn[1];
		m_sCapitalInterestAmount = sReturn[2];
				
		if (sReturn[0] == true)
			FlagChange(sReturn[0]);
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
				//scScreenFunctions.MSGAlert("Incentives are not applicable to ported loan components or non-panel lender mortgage products");
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
		if(sProductCode != m_sNewProductCode || sStartDate != m_sNewProductStartDate)
		{
			if(sProductCode == m_sOriginalProductCode && sStartDate == m_sOriginalProductStartDate)
				m_sNewIncentivesXML = m_sOriginalIncentivesXML;
			else m_sNewIncentivesXML = "";
			m_sNewProductCode = sProductCode;
			m_sNewProductStartDate = sStartDate;
		}
		
		ArrayArguments[0] = m_sApplicationMode;
		ArrayArguments[1] = m_sReadOnly;
		ArrayArguments[2] = m_sRequestAttributes;
		ArrayArguments[3] = sProductCode;
		ArrayArguments[4] = sStartDate;
		ArrayArguments[5] = frmScreen.txtLoanAmount.value;
		ArrayArguments[6] = m_sCurrency;
		ArrayArguments[7] = m_sNewIncentivesXML;
		ArrayArguments[8] = m_sApplicationNumber;
		ArrayArguments[9] = m_sApplicationFactFindNumber;
		ArrayArguments[10]= m_sMortgageSubQuoteNumber;
		ArrayArguments[11]= m_sLoanComponentSeqNum;
		<% /* Pass sub quote data to CM150 */ %>
		ArrayArguments[12]= m_subQuoteXML.XMLDocument.xml;
		ArrayArguments[13] = m_sApplicationDate; 
							
		sReturn = scScreenFunctions.DisplayPopup(window, document, "cm150.asp", ArrayArguments, 505, 465);
		if (sReturn != null)
		{
			m_sNewIncentivesXML = sReturn[2];
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
		m_FurtherFilteringXML.CreateTag("ALLPRODUCTSWITHCHECKS", "");
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
		m_FurtherFilteringXML.CreateTag("PRODUCTOVERRIDECODE", "");
		m_FurtherFilteringXML.CreateTag("MANUALPORTEDLOANIND", "");
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
	sReturn = scScreenFunctions.DisplayPopup(window, document, "CM120.asp", ArrayArguments, 490, 480);
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
	var sProductCodeSearchInd = "0";
	
	<%/* If a specific product was requested in the search, we'll need to store this fact */%>
	if (m_FurtherFilteringXML != null)
	{
		if (m_FurtherFilteringXML.ActiveTag == null || m_FurtherFilteringXML.GetTagText("PRODUCTOVERRIDECODE") == "")
		{
			sProductCodeSearchInd = "0";
		}
		else
		{
			sProductCodeSearchInd = "1";
		}
	}
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
	
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();	
	XML.CreateRequestTagFromArray(m_sRequestAttributes, "SAVE");
	XML.CreateActiveTag("MORTGAGESUBQUOTE");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber);
	if(m_sLoanComponentSeqNum != null)XML.CreateTag("LOANCOMPONENTSEQUENCENUMBER", m_sLoanComponentSeqNum);
	XML.CreateTag("PURPOSEOFLOAN", frmScreen.cboPurposeOfLoan.value);
	<% // EP2_697 %>
	XML.CreateTag("SUBPURPOSEOFLOAN",frmScreen.cboSubPurposeOfLoan.value);
	//END: (MAR58)
	if (frmScreen.optLoanComponentToBePortedYes.checked)
		XML.CreateTag("MANUALPORTEDLOANIND", "1");
	else 
		XML.CreateTag("MANUALPORTEDLOANIND", "0");
	<% /* PSC 13/02/2007 EP2_1314 - Start */ %>
	if (frmScreen.optLoanComponentToBeRetainedYes.checked)
		XML.CreateTag("PRODUCTSWITCHRETAINPRODUCTIND", "1");
	else
		XML.CreateTag("PRODUCTSWITCHRETAINPRODUCTIND", "0");
	<% /* PSC 13/02/2007 EP2_1314 - End */ %>
	XML.CreateTag("LOANAMOUNT", frmScreen.txtLoanAmount.value);
	XML.CreateTag("REPAYMENTTYPE", frmScreen.cboRepaymentType.value);
	<% //SW 20/06/2006 EP771 %>
	XML.CreateTag("REPAYMENTVEHICLE", frmScreen.cboRepaymentVehicle.value);
	XML.CreateTag("REPAYMENTVEHICLEMONTHLYCOST", frmScreen.txtRepaymentVehicleMonthlyCost.value);<%/*EP2_697*/%>
	XML.CreateTag("TERMINYEARS", frmScreen.txtTermYears.value);
	if(frmScreen.txtTermMonths.value == "")XML.CreateTag("TERMINMONTHS", "0");
	else XML.CreateTag("TERMINMONTHS", frmScreen.txtTermMonths.value);
	XML.CreateTag("MORTGAGEPRODUCTCODE", sProductCode);
	XML.CreateTag("PORTEDLOAN", sPortedLoan);
	XML.CreateTag("STARTDATE", sStartDate);
	XML.CreateTag("CAPITALANDINTERESTELEMENT", m_sCapitalInterestAmount);
	XML.CreateTag("INTERESTONLYELEMENT", m_sInterestOnlyAmount);
	XML.CreateTag("NETCAPANDINTELEMENT", m_sCapitalInterestAmount);
	XML.CreateTag("NETINTONLYELEMENT", m_sInterestOnlyAmount);
	XML.CreateTag("STARTDATE", sStartDate);
	XML.CreateTag("PRODUCTCODESEARCHIND", sProductCodeSearchInd);
	XML.CreateTag("RESOLVEDRATE", "");
	XML.CreateTag("MANUALADJUSTMENTPERCENT", "");
	
	if(m_sNewIncentivesXML != "")
	{
		XML.CreateTag("INCENTIVESPRODUCTCODE", m_sNewProductCode);
		XML.CreateTag("INCENTIVESPRODUCTSTARTDATE", m_sNewProductStartDate);
		var incentivesXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		incentivesXML.LoadXML(m_sNewIncentivesXML);
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
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "SaveLoanComponentDetails.asp");
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
	var sAmountRequested;
	
	<% /* Check whether this loan component is one less than the maximum
	      If this is the case the following needs to apply  */ %>
	if (m_sLoanComponentSeqNum == null)
	{
		if (m_sNumberOfLoanComponents == (m_subQuoteXML.GetTagText("MAXNOLOANS")) - 1) 
			{
			alert("The maximum number of loan components has been reached ... The loan value will be defaulted to the remaining value.");
			//scScreenFunctions.MSGAlert("The maximum number of loan components has been reached ... The loan value will be defaulted to the remaining value.");
			frmScreen.txtLoanAmount.value = m_NewLoanCompXML.GetTagText("DEFAULTLOANAMOUNT");
			}
	}
	
	<% /* MAR162. If this is the first Loan Component, check the Loan Amount */ %>
	if (m_bNewLoanComponent == true)
	{
		sAmountRequested = m_NewLoanCompXML.GetTagText("DEFAULTLOANAMOUNT");
	
		if (parseInt(frmScreen.txtLoanAmount.value) > parseInt(sAmountRequested))
		{
			alert("Loan amount is greater than the amount requested (£" + sAmountRequested + ")");
			//scScreenFunctions.MSGAlert("Loan amount is greater than the amount requested (£" + sAmountRequested + ")");
			frmScreen.txtLoanAmount.focus();
			return;
		}	
	}
	
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
			frmScreen.btnOK.disabled = true;
			
			<% /* MAR162  Do not disable the Search button - a re-search should be allowed */ %>
			frmScreen.btnSearch.disabled = false;
			
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
	m_ProductXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_ProductXML.CreateRequestTagFromArray(m_sRequestAttributes, "SEARCH");
	m_ProductXML.CreateActiveTag("MORTGAGEPRODUCTREQUEST");
	m_ProductXML.CreateTag("SEARCHCONTEXT", m_sApplicationMode)	
	m_ProductXML.CreateTag("DISTRIBUTIONCHANNELID", m_sDistributionChannelId);
	m_ProductXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	m_ProductXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	m_ProductXML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber);
	m_ProductXML.CreateTag("PURPOSEOFLOAN", frmScreen.cboPurposeOfLoan.value);
	m_ProductXML.CreateTag("TERMINYEARS", frmScreen.txtTermYears.value);
	if(frmScreen.txtTermMonths.value == "")m_ProductXML.CreateTag("TERMINMONTHS", "0");
	else m_ProductXML.CreateTag("TERMINMONTHS", frmScreen.txtTermMonths.value);
	m_ProductXML.CreateTag("AMOUNTREQUESTED", m_subQuoteXML.GetTagText("AMOUNTREQUESTED"));
	<% /* Send in Loan Component Amount & Sequence number to use when checking for Lenders */ %>
	m_ProductXML.CreateTag("LOANCOMPONENTAMOUNT", frmScreen.txtLoanAmount.value);
	m_ProductXML.CreateTag("LTV", m_subQuoteXML.GetTagText("LTV"));

	<% /* EP2_8 08Nov06 - Assumes ALL have values set correctly on MQ010 */ %>
	m_MQ010XML.SelectTag(null, "PARAMS");
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
	m_ProductXML.CreateTag("NATUREOFLOAN", m_MQ010XML.GetTagText("NATUREOFLOAN"));
	m_ProductXML.CreateTag("CREDITSCHEME", m_MQ010XML.GetTagText("CREDITSCHEME"));
	m_ProductXML.CreateTag("GUARANTORIND", m_MQ010XML.GetTagText("GUARANTORIND"));
	m_ProductXML.CreateTag("APPLICATIONINCOMESTATUS", m_MQ010XML.GetTagText("APPLICATIONINCOMESTATUS"));
	<% /* EP2_8 24Oct06 - End */ %>

	var nTableLength = m_tblProductTable.rows.length - 1;	
	m_ProductXML.CreateTag("RECORDCOUNT", m_tblProductTable.rows.length - 1);
	var nStartRecord = ((nTableLength * iPageNo) + 1) - nTableLength; 
	m_ProductXML.CreateTag("STARTRECORD", nStartRecord);
	if(m_bFurtherFiltering == false)			<% /* EP1089/MAR1833 */ %>
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
		m_ProductXML.CreateTag("MORTGAGEPRODUCTCODE", "");
		m_ProductXML.CreateTag("MANUALPORTEDLOANIND", "");
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
		m_ProductXML.CreateTag("MORTGAGEPRODUCTCODE", m_FurtherFilteringXML.GetTagText("PRODUCTOVERRIDECODE"));
		m_ProductXML.CreateTag("MANUALPORTEDLOANIND", m_FurtherFilteringXML.GetTagText("MANUALPORTEDLOANIND"));

		<% /* EP2_8 08Nov06 - Set Flexible/Non-Flexible flags if not already set. */ %>
		if (m_CalculatedFiltering == "-1")
		{	m_ProductXML.CreateTag("FLEXIBLENONFLEXIBLEPRODUCTS", m_FurtherFilteringXML.GetTagText("FLEXIBLENONFLEXIBLEPRODUCTS"));
			m_ProductXML.CreateTag("FLEXIBLEIND", m_FurtherFilteringXML.GetTagText("FLEXIBLEIND"));
		}

	}
	
	m_ProductXML.RunASP(document,"AQFindMortgageProducts.asp");
		
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = m_ProductXML.CheckResponse(sErrorArray);

	if(sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		alert("No suitable products could be found that match the search criteria");
		//scScreenFunctions.MSGAlert("No suitable products could be found that match the search criteria");
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

<% /*Function to clear the contents of the product table and set the button state */ %>
function clearProductTable()
{
	scTable.clear();
	m_ListEmpty = true;
	frmScreen.btnDetails.disabled = true;
	frmScreen.btnIncentives.disabled = true;
	frmScreen.btnOK.disabled = true;
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

<% /* EP2_8 - New Method - 26Oct06 */ %>
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
	var sGroupName = "";
	
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
		sGroupName = "PurposeOfLoanNew";
		bcontinue = false;
	}

	// Check PSW
	if(bcontinue == true && XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('PSW')) == true)
	{
		sValidationValue = "PSW";
		sValidationDefaultValue = "PSWD";
		sGroupName = "PurposeOfLoanNew";
		bcontinue = false;
	}
	
	// Check N.
	if(bcontinue == true &&  XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('N')) == true)
	{
		sValidationValue = "N";
		sGroupName = "PurposeOfLoanNew";
	}

	// Check TOE
	if(bcontinue == true &&  (XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('TOE')) == true) ||
		(XML.IsInComboValidationList(document, 'TypeOfMortgage', sTypeOfMortgage, Array('T')) == true))
	{
		sValidationValue = "TOE";
		sValidationDefaultValue = "TOED";
		sGroupName = "PurposeOfLoanTfrOfEquity";
		bcontinue = false;
	}

	// Create list
	var sGroupList = new Array(sGroupName);

	//Get list
	XML.GetComboLists(document, sGroupList);

	// Get values from list
	XMLPurposeOfLoan = XML.GetComboListXML(sGroupName);
	XMLPurposeOfLoan2.LoadXML(XMLPurposeOfLoan.xml);
	
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
<%/*EP2_697*/%>
function frmScreen.cboRepaymentVehicle.onchange()
{
	if(m_bDataRead == false)
	{
		frmScreen.txtRepaymentVehicleMonthlyCost.value = "";
	}
	m_bDataRead = false;
}

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
		<% /* Validation script - Controls Soft Coded Field Attributes */ %>
		<script src="validation.js" language="JScript" type="text/Jscript"></script>
	</body>
</html>
