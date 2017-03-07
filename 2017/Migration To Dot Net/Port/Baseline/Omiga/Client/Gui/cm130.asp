<%@ LANGUAGE="JSCRIPT" codePage="28591"%>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<%
/*
Workfile:		cm130.asp
Copyright:		Copyright © 1999 Marlborough Stirling
 
Description:	One-Off Costs Screen.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		09/02/00	Reworked for IE5 and performance.
JLD		22/02/2000	AQR SYS0285,SYS0286
JLD		29/02/00	SYS0287 increased size of table
MCS		01/03/00	SYS0300
MCS		03/03/00	SYS0301
AY		17/03/00	MetaAction changed to ApplicationMode
AY		29/03/00	scScreenFunctions change
AY		12/04/00	SYS0328 - Dynamic currency display
AY		13/04/00	SYS0328 - Dynamic currency display
JLD		28/04/00	Functionality added for non-Mortgage Calculator mode.
JLD		02/05/00	Corrected spelling mistake on ALLOWMIGFEEADDED
JLD		17/05/00	SYS0725 - Added life benefit information to save method
MS		12/06/00	SYS0832 - Added code to trap Description = "NULL".
AP		15/08/00	SYS0898 - Added code to disable calc button on initialise
DF		19/06/02	BMIDS00077 - Altering code to make this compatible with
					with version 7.0.2 of Omiga 4, the following comment was in Core. 
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MO		06/06/2002	CMWP2		Added valuation type REV, for revaluations.
DPF		10/07/2002	CMWP3		Code alter for part - part calculations altered to include fees.
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MV		07/10/2002	BMIDS00530	Amended PopulateFields() and DisplayCosts()
DPF		09/10/2002	BMIDS00046	BM020 - Add 'Manual Incentive' amount to screen and alter calculations
TW      09/10/2002  SYS5115		Modified to incorporate client validation
DPF		10/10/2002				Applied manual fix to If statements for Screen Rules
DPF		24/10/2002  BMIDS00046  BM020 - extra set of scroll buttons were present so they have been removed
DPF		07/11/2002	BMIDS00808	Take in new parameter (DRAWDOWN) and pas them into calculation.
HMA     15/05/2003  BM0133      Only display deposit if the type of application = New Loan.
GD		16/05/2003	BM0368		Eliminate 'NaN' problems by using new method parseIntSafe()
mc		19/04/2004	BMIDS517	white space padded to the title text
DRC     18/05/2004  BM767       CC071 - Adding AdHoc fees to loan & APR
DRC     17/06/2004  BMIDS767    Fixed validation type processing
DRC     25/06/2004  BMIDS763    Allow ProductSwitchFee Display
MC		06/07/2004	BMIDS763	isTIDType() function added inorder to check TID type in an array.
								ApplicationDate is read and passed with oneoffcalc calculation request.
GHun	09/07/2004	BMIDS766	Save updated MortgageSubQuote fields
SR		12/07/2004  BMIDS767
SR		12/07/2004  BMIDS767
GHun	28/10/2004	BMIDS936	Changed GetNonAPRValidationType to GetFeeValidationType to exclude other validation types
KRW     11/11/2004  BMIDS944    Added check to spnOneOffCosts to ignore non selection error ie. iRowSelected = -1 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:
HMA		08/08/2005	MAR28		WP12 - Fees. Add Refund Amount.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom 2 Specific History:
PE		07/04/2007	 EP2_2274	Added check for PSW and CLI validation types to GetAddedToLoanString.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<html>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>One-Off Costs <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<body id="scBody">
<script src="validation.js" language="JScript"></script>

<% /* these are being removed	-	DPF 19/06/02 - BMIDS00077
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
*/%>

<!-- List Scroll object ( and table I believe) -->
<span id="spnListScroll">
<% /* DRC BMIDS767 Start*/%>
<% /*   <span style="LEFT: 100px; POSITION: absolute; TOP: 200px"> */%>
<span style="LEFT: 100px; POSITION: absolute; TOP: 220px">
<% /* DRC BMIDS767 End*/%>
<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; height:24; width:304" 
	type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Forms Here */%>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */%>
<form id="frmScreen" mark validate="onchange" year4>
	<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 340px; WIDTH: 404px; POSITION: ABSOLUTE">
<% /* DRC BMIDS767 Start*/%>	
<% /* 		<span style="TOP: 0px; LEFT: 0px; WIDTH: 404px; HEIGHT: 210px; POSITION: ABSOLUTE" class="msgGroup">*/%>
		<span style="TOP: 0px; LEFT: 0px; WIDTH: 404px; HEIGHT: 240px; POSITION: ABSOLUTE" class="msgGroup">
<% /* DRC BMIDS767 End*/%>		
			<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
				<strong>One-Off Costs...</strong>
			</span>
			<span id="spnOneOffCosts" style="TOP: 30px; LEFT: 4px; POSITION: ABSOLUTE">
				<table id="tblTable" width="396px" border="0" 
					cellspacing="0"cellpadding="0" class="msgTable">
					<tr id="rowTitles">
					<% /* DRC BMIDS767 - Rearrange leaving out Name - Start
					      MAR28 Add refund column */ %>	
						<td width="40%" class="TableHead">Description</td>	
						<td id="tdRefundAmount" width="20%" class="TableHead">Refund</td>
						<td id="tdAmount" width="20%" class="TableHead">Amount</td>
						<td id="tdAddToLoan" width="5%" class="TableHead">Add?</td>
						<td id="tdAdHoc" width="5%" class="TableHead">AdHoc?</td>
						<td id="tdAPR" width="10%" class="TableHead">Include APR</td></tr>
					<tr id="row01">	<td width="40%" class="TableTopLeft">&nbsp</td> <td width="20%" class="TableTopCenter">&nbsp</td>	<td width="20%" class="TableTopCenter">&nbsp</td><td width="5%" class="TableTopCenter">&nbsp</td><td width="5%" class="TableTopCenter">&nbsp</td><td width="10%" class="TableTopRight">&nbsp</td></tr>
					<tr id="row02">	<td width="40%" class="TableLeft">&nbsp</td>	<td width="20%" class="TableCenter">&nbsp</td>  <td width="20%" class="TableCenter">&nbsp</td>	<td width="5%" class="TableCenter">&nbsp</td>	<td width="5%" class="TableCenter">&nbsp</td> <td width="10%" class="TableRight">&nbsp</td></tr>
					<tr id="row03">	<td width="40%" class="TableLeft">&nbsp</td>	<td width="20%" class="TableCenter">&nbsp</td>  <td width="20%" class="TableCenter">&nbsp</td>	<td width="5%" class="TableCenter">&nbsp</td>   <td width="5%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
					<tr id="row04">	<td width="40%" class="TableLeft">&nbsp</td>	<td width="20%" class="TableCenter">&nbsp</td>  <td width="20%" class="TableCenter">&nbsp</td>	<td width="5%" class="TableCenter">&nbsp</td>	<td width="5%" class="TableCenter">&nbsp</td> <td width="10%" class="TableRight">&nbsp</td></tr>
					<tr id="row05">	<td width="40%" class="TableLeft">&nbsp</td>	<td width="20%" class="TableCenter">&nbsp</td>  <td width="20%" class="TableCenter">&nbsp</td>	<td width="5%" class="TableCenter">&nbsp</td>   <td width="5%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
					<tr id="row06">	<td width="40%" class="TableLeft">&nbsp</td>	<td width="20%" class="TableCenter">&nbsp</td>  <td width="20%" class="TableCenter">&nbsp</td>	<td width="5%" class="TableCenter">&nbsp</td>   <td width="5%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
					<tr id="row07">	<td width="40%" class="TableLeft">&nbsp</td>	<td width="20%" class="TableCenter">&nbsp</td>	<td width="20%" class="TableCenter">&nbsp</td>	<td width="5%" class="TableCenter">&nbsp</td>   <td width="5%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
					<tr id="row08">	<td width="40%" class="TableLeft">&nbsp</td>	<td width="20%" class="TableCenter">&nbsp</td>	<td width="20%" class="TableCenter">&nbsp</td>	<td width="5%" class="TableCenter">&nbsp</td>   <td width="5%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
					<tr id="row09">	<td width="40%" class="TableLeft">&nbsp</td>	<td width="20%" class="TableCenter">&nbsp</td>	<td width="20%" class="TableCenter">&nbsp</td>	<td width="5%" class="TableCenter">&nbsp</td>   <td width="5%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
					<tr id="row10">	<td width="40%" class="TableBottomLeft">&nbsp</td>	<td width="20%" class="TableBottomCenter">&nbsp</td> <td width="20%" class="TableBottomCenter">&nbsp</td>	<td width="5%" class="TableBottomCenter">&nbsp</td>	 <td width="5%" class="TableBottomCenter">&nbsp</td> <td width="10%" class="TableBottomRight">&nbsp</td></tr>
					<% /* DRC BMIDS767 End*/%>
			</table>
			</span>
		</span>
		<% /* DRC BMIDS767 Start*/%>
		<% /* <span style="TOP: 215px; LEFT: 0px; WIDTH: 404px; HEIGHT: 230px; 
				POSITION: ABSOLUTE" class="msgGroup"> */%>
        <span style="TOP: 250px; LEFT: 0px; WIDTH: 404px; HEIGHT: 230px; 
				POSITION: ABSOLUTE" class="msgGroup">				
        <% /* DRC BMIDS767 End*/%>				
			<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
				<strong>Totals...</strong>
			</span>
			<span style="TOP: 40px; LEFT: 14px; POSITION: ABSOLUTE" class="msgLabel">
				Total One-Off Costs
				<span style="TOP: 0px; LEFT: 175px; POSITION: ABSOLUTE">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
					<input id="txtTotalOneOffCosts" name="TotalOneOffCosts" maxlength="10" 
						style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" 
						class="msgTxt">
				</span>
			</span>
			<span style="TOP: 65px; LEFT: 14px; POSITION: ABSOLUTE" class="msgLabel">
				Deposit
				<span style="TOP: 0px; LEFT: 175px; POSITION: ABSOLUTE">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
					<input id="txtDeposit" name="Deposit" maxlength="10" 
						style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" 
						class="msgTxt">
				</span>
			</span>
			<span style="TOP: 90px; LEFT: 14px; POSITION: ABSOLUTE" class="msgLabel">
				One-Off Costs Added To Loan
				<span style="TOP: 0px; LEFT: 175px; POSITION: ABSOLUTE">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
					<input id="txtOneOffCostsAddedToLoan" name="OneOffCostsAddedToLoan" maxlength="10" 
						style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" 
						class="msgTxt">
				</span>
			</span>
			<span style="TOP: 115px; LEFT: 14px; POSITION: ABSOLUTE" class="msgLabel">
				Total Incentives
				<span style="TOP: 0px; LEFT: 175px; POSITION: ABSOLUTE">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
					<input id="txtTotalIncentives" name="TotalIncentives" maxlength="10" 
						style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" 
						class="msgTxt">
				</span>
			</span>
			<span style="TOP: 140px; LEFT: 14px; POSITION: ABSOLUTE" class="msgLabel">
				Manual Incentives
				<span style="TOP: 0px; LEFT: 175px; POSITION: ABSOLUTE">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
					<input id="txtManualIncentives" name="ManualIncentives" maxlength="10" 
						style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" 
						class="msgTxt">
				</span>
			</span>
			<span style="TOP: 165px; LEFT: 14px; POSITION: ABSOLUTE" class="msgLabel">
				Total Net Cost
				<span style="TOP: 0px; LEFT: 175px; POSITION: ABSOLUTE">
					<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
					<input id="txtTotalNetCost" name="TotalNetCost" maxlength="10" 
						style="TOP: -3px; WIDTH: 70px; POSITION: ABSOLUTE" 
						class="msgTxt">
				</span>
			</span>
			<span style="TOP: 200px; LEFT: 4px; POSITION: ABSOLUTE">
				<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
					<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
				<span style="TOP: 0px; LEFT: 65px; POSITION: ABSOLUTE">
					<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
				<span style="TOP: 0px; LEFT: 130px; POSITION: ABSOLUTE">
					<input id="btnCalculate" value="Calculate" type="button" style="WIDTH: 60px" class="msgButton">
				</span>
				</span>
		</span>						
	</div>
</form>
	
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cm130Attribs.asp" -->
	
<% /* Specify Code Here */%>
<script language="JScript">
<!--	
var m_sApplicationMode = null;
var m_sCurrency = "";
var m_XMLOneOffCosts = null;
var m_XMLComboCosts = null;
var m_xmlCustomerList = null;
var m_xmlGetOneOffCosts = null;
var m_xmlREVISEDCOSTS = null;
var m_XMLRecost = null;
var m_iPurchasePrice = 0;
var m_iAmountRequested = 0;
var	m_iTotalIncentives = 0;
var m_iManualIncentives = 0;
var m_iTableLength = 10;
var m_sApplicationNo = "";
var	m_sApplicationFFNo = "";
var	m_sMortSubQuoteNo = "";
var	m_sLifeSubQuoteNo = "";
var m_iDrawdown = 0;	//BMIDS00808 - DPF Nov 02 - new variable
var m_sAttributeArray;
var m_sReadOnly;
var m_iFeesAddedToLoanTotal = 0;	//CMWP3 - DPF Jul 02 - new variable
var m_sINTERESTONLYAMOUNT = 0;
var m_sCAPITALINTERESTAMOUNT = 0;
var m_iInterestOnlyAmount = 0;		//CMWP3 - DPF Jul 02 - new variable
var m_iCapitalInterestAmount = 0;	//CMWP3 - DPF Jul 02 - new variable
var m_iOrigCapitalInterestAmount = 0;	//CMWP3 - DPF Jul 02 - new variable
var m_iTotalOneOffCosts = 0;
var m_iOneOffCostsAddedToLoan = 0;
var scScreenFunctions;
// this line added from V7 - DPF (19/06/02) - BMIDS00077
var m_BaseNonPopupWindow = null;
var m_sApplicationType = "";        // BM0133  New parameter


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	//this line is being removed as per V7 - DPF 19/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	RetrieveData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	PopulateScreen();				
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveData()
{
	<% /* Retrieves all context data required for use within this screen 
		Parameter	MC					CM/QQ
		1			Context				Context
		2			One-Off Costs XML	Application Number
		3			Total Incentives	ApplicationFactFind Number
		4			Purchase Price		Mortgage Sub Quote Number
		5			Amount Requested	Life Sub Quote Number
		6			UserId etc			Read Only flag
		7			Currency			UserId, UnitId, MachineId, ChannelId
		8			not used			Currency
		9			not used			CustomerListXML
		
				
		//CMWP3 - DPF Jul 02 - two new parameters
		10			not used			Interest Only Amount
		11			not used			Capital & Interest Amount
		
		//BMIDS00808 - DPF Nov 02 - new parameter
		12			not used			Drawdown
		
		//BM0133 - new parameter
		13          not used            Application Type 
		*/
	%>
	
	var sArguments = window.dialogArguments;			
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];			
	var sParameters	= sArguments[4];
	//next 2 lines added as per V7 - DPF 19/06/02 - BMIDS00077
	m_BaseNonPopupWindow = sArguments[5];
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	var sCurrencyText;
			
	m_sApplicationMode	= sParameters[0];
	if (m_sApplicationMode == "Mortgage Calculator")
	{
		m_iPurchasePrice = parseIntSafe(sParameters[3]);
		m_iAmountRequested	= parseIntSafe(sParameters[4]);
		//next line replaced with line below as per V7 - DPF 19/06/02 - BMIDS00077
		//m_XMLOneOffCosts = new scXMLFunctions.XMLObject();
		m_XMLOneOffCosts = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_XMLOneOffCosts.LoadXML(sParameters[1]);			
		m_iTotalIncentives = parseIntSafe(sParameters[2]);
		m_sAttributeArray = sParameters[5];
		m_sCurrency = sParameters[6];
		sCurrencyText = scScreenFunctions.SetThisCurrency(frmScreen,sParameters[6]);
		if (isNaN(m_iTotalIncentives)) m_iTotalIncentives = 0;		
	}
	else
	{		
		m_sApplicationNo = sParameters[1];
		m_sApplicationFFNo = sParameters[2];
		m_sMortSubQuoteNo = sParameters[3];
		m_sLifeSubQuoteNo = sParameters[4];
		m_sReadOnly = sParameters[5];
		m_sAttributeArray = sParameters[6];
		m_sCurrency = sParameters[7];
		sCurrencyText = scScreenFunctions.SetThisCurrency(frmScreen,sParameters[7]);
		//next line replaced with line below as per V7 - DPF 19/06/02 - BMIDS00077
		//m_xmlCustomerList = new scXMLFunctions.XMLObject();
		m_xmlCustomerList = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_xmlCustomerList.LoadXML(sParameters[8]);
		
		//CMWP3 - DPF Jul 02 - two new variable asignments
		m_iInterestOnlyAmount = sParameters[9];
		m_iCapitalInterestAmount = sParameters[10];
		
		//BMIDS00808 - DPF Nov 02 - new variable
		m_iDrawdown = sParameters[11];
		
		// BM0133
		m_sApplicationType = sParameters[12];
	}	
	tdAmount.innerText = tdAmount.innerText + " " + sCurrencyText;
	tdRefundAmount.innerText = tdRefundAmount.innerText + " " + sCurrencyText ;   // MAR28
}

function PopulateScreen()
{
	if (GetComboList())
	{
		if (m_sApplicationMode == "Mortgage Calculator")
		{
			SetOneOffCosts();					
			var iNumberOfCosts = GetNumberOfValidCosts();					
			scScrollTable.initialiseTable(tblTable, 0, "", DisplayMCCosts, m_iTableLength, iNumberOfCosts);
			DisplayMCCosts(0);
			scScreenFunctions.SetScreenToReadOnly(frmScreen);			
			frmScreen.btnCalculate.disabled = true;
			frmScreen.btnCancel.disabled = true;
		}
		else
		{
			/* MO - BMIDS00105 m_xmlGetOneOffCosts = new scXMLFunctions.XMLObject(); */
			m_xmlGetOneOffCosts = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			m_xmlGetOneOffCosts.CreateRequestTagFromArray(m_sAttributeArray,"SEARCH");
			m_xmlGetOneOffCosts.CreateActiveTag("ONEOFFCOST");
			m_xmlGetOneOffCosts.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
			m_xmlGetOneOffCosts.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
			m_xmlGetOneOffCosts.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortSubQuoteNo);
			m_xmlGetOneOffCosts.RunASP(document, "GetOneOffCostsDetails.asp");
		
			if (m_xmlGetOneOffCosts.IsResponseOK() == true)
			{
				PopulateListBox();
				// DPF 11/10/2002 - CPWP1 - Manual Incentives

				m_xmlGetOneOffCosts.SelectTag(null, "RESPONSE");
				m_iManualIncentives = m_xmlGetOneOffCosts.GetTagText("MANUALINCENTIVEAMOUNT");
				PopulateFields();
				//CMWP3 - recalculate CapitalInterest Amount by remove the one-off costs (to be added again on Calculate)
				m_iOrigCapitalInterestAmount = parseIntSafe(m_iCapitalInterestAmount) - parseIntSafe(m_iOneOffCostsAddedToLoan);
				scScreenFunctions.SetFieldState(frmScreen, "txtTotalOneOffCosts","R");
				scScreenFunctions.SetFieldState(frmScreen, "txtDeposit","R");
				scScreenFunctions.SetFieldState(frmScreen, "txtOneOffCostsAddedToLoan","R");
				scScreenFunctions.SetFieldState(frmScreen, "txtTotalIncentives","R");
				scScreenFunctions.SetFieldState(frmScreen, "txtTotalNetCost","R");
				//DPF 11/10/2002 - CPWP1 - determine if Manual Incentives field should be enabled
				if(frmScreen.txtTotalIncentives.value == "" || frmScreen.txtTotalIncentives.value == "0")
				{	
					frmScreen.txtManualIncentives.value = "0";
					scScreenFunctions.SetFieldState(frmScreen, "txtManualIncentives","R");
				}
				
				frmScreen.btnOK.disabled = true;
				frmScreen.btnCalculate.disabled = true;
			}
			if(m_sReadOnly == "1")
			{
				scScreenFunctions.SetScreenToReadOnly(frmScreen);			
				frmScreen.btnCalculate.disabled = true;
			}
		}
	}
}
function PopulateListBox()
{
	var iNumberOfValidCosts;
	iNumberOfValidCosts = GetNumberOfValidCosts();
	scScrollTable.initialiseTable(tblTable, 0, "", DisplayCosts, m_iTableLength, iNumberOfValidCosts);
	if(iNumberOfValidCosts > 0)DisplayCosts(0);
}
function PopulateFields()
{
	m_xmlGetOneOffCosts.ActiveTag = null;
	var tagCOSTSLIST = m_xmlGetOneOffCosts.CreateTagList("MORTGAGEONEOFFCOST");
	var iNumberOfOneOffCosts = m_xmlGetOneOffCosts.ActiveTagList.length;
	var sNewLoanValue;
	
	m_iTotalOneOffCosts = 0;
	m_iOneOffCostsAddedToLoan = 0;
	
	for (var nCost = 0; nCost < iNumberOfOneOffCosts; nCost++)
	{
		m_xmlGetOneOffCosts.ActiveTagList = tagCOSTSLIST;
		if (m_xmlGetOneOffCosts.SelectTagListItem(nCost) == true)
		{
			if(m_xmlGetOneOffCosts.GetTagText("AMOUNT") != "" && m_xmlGetOneOffCosts.GetTagText("AMOUNT") != "0")
			{
				var iAmount = parseIntSafe(m_xmlGetOneOffCosts.GetTagText("AMOUNT"));
				var iRefundAmount = parseIntSafe(m_xmlGetOneOffCosts.GetTagText("REFUNDAMOUNT"));   // MAR28

				var sRetArray = GetComboValues(m_xmlGetOneOffCosts.GetTagText("MORTGAGEONEOFFCOSTTYPE"));
				// DRC BMIDS767 - check all validation types for valueid				
				var sAddedToLoan = GetAddedToLoanString(sRetArray, m_xmlGetOneOffCosts.GetTagText("ADDTOLOAN"));
				
				if(!isTIDType(sRetArray)) m_iTotalOneOffCosts += iAmount; <% /* SR 12/07/2004 : BMIDS767 */ %>

				if(sAddedToLoan == "Yes")
					m_iOneOffCostsAddedToLoan += iAmount;
			}
		}
	}
	frmScreen.txtTotalOneOffCosts.value = m_iTotalOneOffCosts;
	frmScreen.txtOneOffCostsAddedToLoan.value = m_iOneOffCostsAddedToLoan;
	m_xmlGetOneOffCosts.SelectTag(null, "MORTGAGESUBQUOTE");
	
	// BM0133
	// Check the Application Type passed through from CM100.
	// Only display the DEPOSIT if the type is New Loan (validation type 'N')
	
	// Get the value associated with Validation Type 'N'
	sNewLoanValue = m_xmlGetOneOffCosts.GetComboIdForValidation("TypeOfMortgage", "N", null, document);
	
	if (m_sApplicationType == sNewLoanValue)
	{
		frmScreen.txtDeposit.value = m_xmlGetOneOffCosts.GetTagText("DEPOSIT");
	}		
	else
	{	
		frmScreen.txtDeposit.value = "0";
	}		
	// BM0133  End		
	m_xmlGetOneOffCosts.SelectTag(null, "RESPONSE");
	frmScreen.txtTotalIncentives.value = m_xmlGetOneOffCosts.GetTagText("TOTALINCENTIVES");
	
	// DPF 11/10/2002 - CPWP1 - Manual Incentives
	frmScreen.txtManualIncentives.value = m_iManualIncentives
	
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
}
function DisplayCosts(iStart)
{
	scScrollTable.clear();
	m_xmlGetOneOffCosts.ActiveTag = null;
	var tagCOSTSLIST = m_xmlGetOneOffCosts.CreateTagList("MORTGAGEONEOFFCOST");
	var iNumberOfOneOffCosts = m_xmlGetOneOffCosts.ActiveTagList.length;
	var iRowAdded = 0;
	var iValidRow = 0;
	for (var nCost = 0; nCost < iNumberOfOneOffCosts && iRowAdded < m_iTableLength; nCost++)
	{
		m_xmlGetOneOffCosts.ActiveTagList = tagCOSTSLIST;
		if (m_xmlGetOneOffCosts.SelectTagListItem(nCost) == true)
		{
			if(m_xmlGetOneOffCosts.GetTagText("AMOUNT") != "" && m_xmlGetOneOffCosts.GetTagText("AMOUNT") != "0")
			{
				if(iValidRow == iStart)
				{
					<% /* DRC BMIDS767  Start
					iRowAdded++;
					var sName = m_xmlGetOneOffCosts.GetTagText("DESCRIPTION");
					if (sName == "NULL")sName=""; 
					  DRC BMIDS767 End */%>
					var sAmount = m_xmlGetOneOffCosts.GetTagText("AMOUNT");
					var sRefundAmount = m_xmlGetOneOffCosts.GetTagText("REFUNDAMOUNT");  // MAR28
					var sValueId = m_xmlGetOneOffCosts.GetTagText("MORTGAGEONEOFFCOSTTYPE");
					var sRetArray = GetComboValues(sValueId);
					<% /* DRC BMIDS767  Start */ %>
					var sAdHocCost = "";
					if (m_xmlGetOneOffCosts.GetTagText("ADHOCIND") == "1") sAdHocCost = "Yes";
					var sAPR = IncludeAPR(sValueId); 
                    <% /*  DRC BMIDS767 End */%>
                    
					//if (sRetArray[1] != "TID")
					//*=[MC] TID Type may exists in other than position 1 in an array, iterate through all the items in
					//the list.
					if(!isTIDType(sRetArray))
					{
					    iRowAdded++;
						var sAddedToLoan = GetAddedToLoanString(sRetArray, m_xmlGetOneOffCosts.GetTagText("ADDTOLOAN"));
						<% /* DRC BMIDS767 changed parameters in Showrow call
						ShowRow(iRowAdded, sRetArray[0], sName, sAmount, sAddedToLoan, nCost); */ %>
						ShowRow(iRowAdded, sRetArray[0], sRefundAmount, sAmount, sAddedToLoan, sAdHocCost, sAPR, nCost);   // MAR28
					}
				}
				else iValidRow++;
			}
		}
	}
}

function isTIDType(sArray)
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

function GetComboList()
{
	//next line removed to be replaced by line below as per V7 - DPF 19/06/02 - BMIDS00077
	//m_XMLComboCosts = new scXMLFunctions.XMLObject();
	m_XMLComboCosts = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	var sGroups = new Array("OneOffCost");			
	var bSuccess = true;
			
	if (m_XMLComboCosts.GetComboLists(document, sGroups) != true)
	{				
		bSuccess = false;
		alert("Unable to read the combo records for OneOffCost group");
	}
			
	return bSuccess;			
}
function GetComboValues(sValueId)
{
	var sRetArray = new Array;
	m_XMLComboCosts.SelectTag(null, "LISTNAME");
	var tagLISTENTRY = m_XMLComboCosts.CreateTagList("LISTENTRY");
	var iNumberOfEntries = m_XMLComboCosts.ActiveTagList.length;
	
	var bFound = false;
	for (var nEntry = 0; nEntry < iNumberOfEntries && bFound == false; nEntry++)
	{
		m_XMLComboCosts.ActiveTagList = tagLISTENTRY;
		if (m_XMLComboCosts.SelectTagListItem(nEntry) == true)
		{
			if(m_XMLComboCosts.GetTagText("VALUEID") == sValueId)
			{
				bFound = true;
				sRetArray[0] = m_XMLComboCosts.GetTagText("VALUENAME");
				var tagVALIDATIONLIST = m_XMLComboCosts.CreateTagList("VALIDATIONTYPE");
				for(var i = 0; i < m_XMLComboCosts.ActiveTagList.length; i++)
				{
					//DRC BMIDS767 change to get all validationtypes
					sRetArray[i+1] = m_XMLComboCosts.GetTagListItem(i).text;
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
	var tagMORTGAGELENDER = m_xmlGetOneOffCosts.SelectTag(null, "MORTGAGELENDER");
	
	for (var nEntry = 1; nEntry < sValidationArray.length && bFound == false; nEntry++)
	{
	    sValidationType = sValidationArray[nEntry];
		if( (sValidationType == "ADM" && m_xmlGetOneOffCosts.GetTagText("ALLOWADMINFEEADDED") == "1") ||
		    (sValidationType == "ARR" && m_xmlGetOneOffCosts.GetTagText("ALLOWARRGEMTFEEADDED") == "1") ||
		    (sValidationType == "DEE" && m_xmlGetOneOffCosts.GetTagText("ALLOWDEEDSRELFEEADD") == "1") ||
		    (sValidationType == "MIG" && m_xmlGetOneOffCosts.GetTagText("ALLOWMIGFEEADDED") == "1") ||
		    (sValidationType == "POR" && m_xmlGetOneOffCosts.GetTagText("ALLOWPORTINGFEEADDED") == "1") ||
		    (sValidationType == "REI" && m_xmlGetOneOffCosts.GetTagText("ALLOWREINSPTFEEADDED") == "1") ||
		    (sValidationType == "REV" && m_xmlGetOneOffCosts.GetTagText("ALLOWREVALUATIONFEEADDED") == "1") ||
		    (sValidationType == "SEA" && m_xmlGetOneOffCosts.GetTagText("ALLOWSEALINGFEEADDED") == "1") ||
		    (sValidationType == "TTF" && m_xmlGetOneOffCosts.GetTagText("ALLOWTTFEEADDED") == "1") ||
		    (sValidationType == "VAL" && m_xmlGetOneOffCosts.GetTagText("ALLOWVALNFEEADDED") == "1") ||
		    (sValidationType == "LEG" && m_xmlGetOneOffCosts.GetTagText("ALLOWLEGALFEEADDED") == "1") ||
		    <% /* DRC BMIDS763 Start */ %>
		    (sValidationType == "PSF" && m_xmlGetOneOffCosts.GetTagText("ALLOWPRODUCTSWITCHFEEADDED") == "1")  ||
			<% /* DRC BMIDS763 End */ %>
		    <% /* DRC BMIDS767 Start */ %>
		    (sValidationType == "OTH" && m_xmlGetOneOffCosts.GetTagText("ALLOWOTHERFEEADDED") == "1") ||   
		    (sValidationType == "IAF" && m_xmlGetOneOffCosts.GetTagText("ALLOWNONLENDINSADMINFEEADDED") == "1") ||
		    <% /* MSla BBG110:WP04 - Start */ %>
		    (sValidationType == "AB" && m_xmlGetOneOffCosts.GetTagText("ALLOWFURTHERLENDINGFEEADDED") == "1") ||
		    <% /* MSla BBG110:WP04 - End */ %>
		    <% /* DRC BMIDS767 End */ %>
		    (sValidationType == "CLI" && m_xmlGetOneOffCosts.GetTagText("ALLOWCREDITLIMITINCFEEADDED") == "1"))
		    {
				bFound = true;
				if(sAddToLoan == "1")
					sAddedToLoan="Yes";
				else sAddedToLoan="No";
			}
	}
		
	return(sAddedToLoan);
}
 <% /* DRC BMIDS767 Start - cloned from  MSla BBG101:WP07 */ %>
function IncludeAPR(sValueID) 
{
	var sRet = "";
	var arrValidationType = new Array;
	var ValidXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	<% /* bReturn = IsInComboValidationList(thisDocument,sListName, sValueID, ValidationList) 
		 	thisDocument	- the document object of the calling page 
		 	sListName		- The group name for the combo list required from the database. 
		 	sValueID		- The combo group entry value-identification number 
		 	ValidationList	- An array of validation strings 
	*/ %>
	
	arrValidationType[0] = "APR";
	
	if ( ValidXML.IsInComboValidationList(document, "OneOffCost", sValueID, arrValidationType ) )
	{
		sRet = "Yes";
	}
	else
	{
		sRet = "No";
	}
	
	return(sRet);
}
 <% /* DRC BMIDS767 End */ %>
function InitialiseCostXML()
{
	var iNumberOfCosts = 0;
	m_XMLOneOffCosts.ActiveTag = null;
	if (m_XMLOneOffCosts.CreateTagList("ONEOFFCOST") != null)
		iNumberOfCosts = m_XMLOneOffCosts.ActiveTagList.length;
	return iNumberOfCosts;
}
function spnOneOffCosts.ondblclick()
{
<%	/* Double click toggles the Added? text between Yes and No in non Mortgage Calculator mode only */
%>	if(m_sApplicationMode != "Mortgage Calculator")
	{
		var iRowSelected = scScrollTable.getRowSelected();
		if(iRowSelected != -1) // KRW 11/11/04 BMIDS944
		{
			var iIndex = tblTable.rows(iRowSelected).getAttribute("tagListItem");
			m_xmlGetOneOffCosts.ActiveTag = null;
			var tagCOSTSLIST = m_xmlGetOneOffCosts.CreateTagList("MORTGAGEONEOFFCOST");
			if(m_xmlGetOneOffCosts.SelectTagListItem(iIndex) == true)
			{
				if(tblTable.rows(iRowSelected).getAttribute("AddAllowed") == true)
				{
					if(m_xmlGetOneOffCosts.GetTagText("ADDTOLOAN") == "1")
					{
						m_xmlGetOneOffCosts.SetTagText("ADDTOLOAN", "0");
						scScreenFunctions.SizeTextToField(tblTable.rows(iRowSelected).cells(3),"No");
						//CMWP3 - DPF 1/8 - Ensure Ok button can't be clicked until a re-calc is done
						frmScreen.btnOK.disabled = true;
						frmScreen.btnCalculate.disabled = false;
					}
					else
					{
						m_xmlGetOneOffCosts.SetTagText("ADDTOLOAN", "1");
						scScreenFunctions.SizeTextToField(tblTable.rows(iRowSelected).cells(3),"Yes");
						//CMWP3 - DPF 1/8 - Ensure Ok button can't be clicked until a re-calc is done
						frmScreen.btnOK.disabled = true;
						frmScreen.btnCalculate.disabled = false;
					}
	<%				/* re-calculate the totals fields */
	%>				PopulateFields();
				}
			}
		}
	}
}
function SetOneOffCosts()
{
	var iAmount	= 0;
	var iTotalOneOffCost = 0;
	var iTotalNetCost	= 0;			
	var iDeposit	= 0;
	var iNumberOfCosts	= InitialiseCostXML();			
						
	<% /* Deposit */%>
	if (m_iPurchasePrice > m_iAmountRequested)
		iDeposit = m_iPurchasePrice - m_iAmountRequested;
	frmScreen.txtDeposit.value = iDeposit;
			
	<% /* Total incentives */%>
	frmScreen.txtTotalIncentives.value = m_iTotalIncentives;
				
	<% /* Total one-off costs */ %>
	for (var iLoop = 0; iLoop < iNumberOfCosts; iLoop++)
	{
		if (m_XMLOneOffCosts.SelectTagListItem(iLoop) == true)
		{	
			if (m_XMLOneOffCosts.GetTagText("IDENTIFIER")!= "TID")
			{
				iAmount	= m_XMLOneOffCosts.GetTagInt("AMOUNT");
				iTotalOneOffCost	+= iAmount;
			}
		}
	}			
	frmScreen.txtTotalOneOffCosts.value = iTotalOneOffCost;
			
	<% /* Total Net Cost */ %>
	iTotalNetCost = iTotalOneOffCost + iDeposit - m_iTotalIncentives;
	frmScreen.txtTotalNetCost.value = iTotalNetCost;
}
function GetNumberOfValidCosts()
{
	var iNumberOfValidCosts = 0;
	
	if(m_sApplicationMode == "Mortgage Calculator")
	{
		var iNumberOfCosts = InitialiseCostXML();
		for (var iLoop = 0; iLoop < iNumberOfCosts; iLoop++)
		{
			if (m_XMLOneOffCosts.SelectTagListItem(iLoop) == true)
			{
				var sAmount	= m_XMLOneOffCosts.GetTagText("AMOUNT");
				var sIdentifier = m_XMLOneOffCosts.GetTagText("IDENTIFIER");

				if ( sIdentifier != "TID")
				{
					if(sAmount != "" && sAmount != "0")
					{
						iNumberOfValidCosts++;
					}
				}
			}
		}
	}
	else
	{
		m_xmlGetOneOffCosts.ActiveTag = null;
		m_xmlGetOneOffCosts.CreateTagList("MORTGAGEONEOFFCOST");
		for(var iLoop = 0; iLoop < m_xmlGetOneOffCosts.ActiveTagList.length; iLoop++)
		{
			if(m_xmlGetOneOffCosts.SelectTagListItem(iLoop) == true)
			{
				if(m_xmlGetOneOffCosts.GetTagText("AMOUNT") != "" && m_xmlGetOneOffCosts.GetTagText("AMOUNT") != "0")
				{	<% /* SR 12/07/2004 : BMIDS767 */ %>
					var sRetArray = GetComboValues(m_xmlGetOneOffCosts.GetTagText("MORTGAGEONEOFFCOSTTYPE"));
					if (!isTIDType(sRetArray)) iNumberOfValidCosts++;		
					<% /* SR 12/07/2004 : BMIDS767 - End */ %>
				}
			}
		}
	}
	return(iNumberOfValidCosts);
}
function DisplayMCCosts(iStart)
{
	var sAmount = "";
	var sIdentifier = "";
	var sCostDescription = "";
	<% /* DRC BMIDS767 var sName = ""; */ %>

	var iTotalAmount = 0;
	var iNumberOfCosts = InitialiseCostXML();
	var iRowAdded = 0;
	var iValidRow = 0;
	scScrollTable.clear();
	for (var iLoop = 0; iLoop < iNumberOfCosts && iRowAdded < m_iTableLength; iLoop++)
	{
		if (m_XMLOneOffCosts.SelectTagListItem(iLoop) == true)
		{
			sAmount	= m_XMLOneOffCosts.GetTagInt("AMOUNT");
			if(sAmount != "" && sAmount != "0" )
			{
<%				/* Because we are missing out costs which have no amount set, we need to
				   make sure that when scrolling the table, we start from the right valid
				   cost. iValidRow is a count of all valid cost rows and iRowAdded is a
				   count of the actual rows that are added. */
%>				if(iStart == iValidRow)
				{
					sIdentifier = m_XMLOneOffCosts.GetTagText("IDENTIFIER");
					sRefundAmount = m_XMLOneOffCosts.GetTagInt("REFUNDAMOUNT");  // MAR28
					if ( sIdentifier != "TID")
					{
						iRowAdded++;
						<% /* DRC BMIDS767 if (sIdentifier == "OTH")
							sName = m_XMLOneOffCosts.GetTagText("NAME"); */ %>
						sCostDescription = GetCostDescription(sIdentifier);
						<% /* DRC BMIDS767 ShowRow(iRowAdded, sCostDescription, sName, sAmount, "", iLoop);*/ %>
						ShowRow(iRowAdded, sCostDescription, sRefundAmount, sAmount, "", "", "",iLoop);
					}
				}
				else iValidRow++;
			}
		}
	}
}

function GetCostDescription(sIdentifier)
{
	var sCostDescription = "";

	m_XMLComboCosts.ActiveTag = null;
	if (m_XMLComboCosts.CreateTagList("LISTENTRY") != null)
	{
		var iNumberOfRecords = m_XMLComboCosts.ActiveTagList.length;
		var sValidationType;
				
		for (var iLoop=0; iLoop < iNumberOfRecords && sCostDescription.length == 0; iLoop++)
		{
			if (m_XMLComboCosts.SelectTagListItem(iLoop))
			{
				sValidationType = m_XMLComboCosts.GetTagText("VALIDATIONTYPE");
				if (sIdentifier == sValidationType)
					sCostDescription = m_XMLComboCosts.GetTagText("VALUENAME");
			}
		}	
	}
	return sCostDescription;
}
<% /* DRC BMIDS767 function ShowRow(iRow, sCostDescription, sName, sAmount, sAddedToLoan, nAttribute) */ %>
function ShowRow(iRow, sCostDescription, sRefundAmount, sAmount, sAddedToLoan, sAdHocCost,  sAPR, nAttribute)		
{
	scScreenFunctions.SizeTextToField(tblTable.rows(iRow).cells(0),sCostDescription);
	scScreenFunctions.SizeTextToField(tblTable.rows(iRow).cells(1),sRefundAmount);  // MAR28
	scScreenFunctions.SizeTextToField(tblTable.rows(iRow).cells(2),sAmount);
	scScreenFunctions.SizeTextToField(tblTable.rows(iRow).cells(3),sAddedToLoan);
	<% /* DRC BMIDS767 Start */ %>
	scScreenFunctions.SizeTextToField(tblTable.rows(iRow).cells(4),sAdHocCost);
	scScreenFunctions.SizeTextToField(tblTable.rows(iRow).cells(5),sAPR);
	<% /* DRC BMIDS767 End */ %>
	tblTable.rows(iRow).setAttribute("tagListItem", nAttribute);
	if(sAddedToLoan != "No" && sAddedToLoan != "Yes")
		tblTable.rows(iRow).setAttribute("AddAllowed", false);
	else tblTable.rows(iRow).setAttribute("AddAllowed", true);
}

// DPF 11/10/2002 - CPWP1 - ensure we re-calculate
function frmScreen.txtManualIncentives.onchange()
{
	frmScreen.btnOK.disabled = true;
	frmScreen.btnCalculate.disabled = false;
	m_iManualIncentives = frmScreen.txtManualIncentives.value
	/* re-calculate the totals fields */
	PopulateFields();
}

function frmScreen.btnCalculate.onclick()
{
	
	var sApplicationDate = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idApplicationDate",null);
	//next line replaced with line below as per V7 - DPF 19/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_sAttributeArray,null);
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
	XML.CreateTag("APPLICATIONDATE", sApplicationDate);
	XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortSubQuoteNo);
	// 	XML.RunASP(document, "IsAttachedToOneQuote.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "IsAttachedToOneQuote.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if(XML.IsResponseOK())
	{
		var xmlOneOffCost = ConstructOneOffCostXML();
		var tagActiveTag;
		xmlOneOffCost.SelectTag(null, "ONEOFFCOSTLIST");
		//next line replaced with line below as per V7 - DPF 19/06/02 - BMIDS00077
		//m_xmlREVISEDCOSTS = new scXMLFunctions.XMLObject(); 
		m_xmlREVISEDCOSTS = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject(); 
		m_xmlREVISEDCOSTS.CreateRequestTagFromArray(m_sAttributeArray,null);
		if(m_sApplicationMode == "Quick Quote")tagActiveTag = m_xmlREVISEDCOSTS.CreateActiveTag("QUICKQUOTE");
		else tagActiveTag = m_xmlREVISEDCOSTS.CreateActiveTag("ADDEDONEOFFCOSTS");
		m_xmlREVISEDCOSTS.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
		m_xmlREVISEDCOSTS.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
		
		m_xmlREVISEDCOSTS.CreateTag("APPLICATIONDATE", sApplicationDate);
		
		m_xmlREVISEDCOSTS.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortSubQuoteNo);
		m_xmlREVISEDCOSTS.CreateActiveTag("ONEOFFCOSTLIST");
		m_xmlREVISEDCOSTS.AddXMLBlock(xmlOneOffCost.CreateFragment());
		m_xmlREVISEDCOSTS.ActiveTag = tagActiveTag;
		m_xmlREVISEDCOSTS.CreateActiveTag("CUSTOMERLIST");
		m_xmlCustomerList.SelectTag(null, "CUSTOMERLIST");
		m_xmlREVISEDCOSTS.AddXMLBlock(m_xmlCustomerList.CreateFragment());
		// 		if(m_sApplicationMode == "Quick Quote")m_xmlREVISEDCOSTS.RunASP(document, "QQProcessAddedOneOffCosts.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		//DPF 10/10/2002 - applied manual fix to screen rules
		if(m_sApplicationMode == "Quick Quote")
		{
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
					m_xmlREVISEDCOSTS.RunASP(document, "QQProcessAddedOneOffCosts.asp");
					break;
				default: // Error
					m_xmlREVISEDCOSTS.SetErrorResponse();
				}
		}
		else
		{
		// 		else m_xmlREVISEDCOSTS.RunASP(document, "AQProcessAddedOneOffCosts.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
					m_xmlREVISEDCOSTS.RunASP(document, "AQProcessAddedOneOffCosts.asp");
					break;
				default: // Error
					m_xmlREVISEDCOSTS.SetErrorResponse();
				}
		}
		
		if(m_xmlREVISEDCOSTS.IsResponseOK())
		{
			m_xmlREVISEDCOSTS.SelectTag(null, "ADDEDONEOFFCOSTS");
			m_sINTERESTONLYAMOUNT = 0;
			m_sCAPITALINTERESTAMOUNT = 0;
			if(m_xmlREVISEDCOSTS.GetTagText("PARTANDPART") == "True")
			{
			//CMWP3 - DPF Jul 02 - START:  section of code ripped out and replaced by two lines below - no manual intervention
				m_sINTERESTONLYAMOUNT = parseIntSafe(m_iInterestOnlyAmount);
			//CMWP3 - easiet way to calculate was to take original CapitalInt amount
			//(b4 fees)and add total of fees to be added onto it.
				m_sCAPITALINTERESTAMOUNT = parseIntSafe(m_iOrigCapitalInterestAmount) + parseIntSafe(m_iFeesAddedToLoanTotal);
			}
			//CMWP3 - END

			//next line replaced by line below as per V7 - DPF 19/06/02 - BMIDS00077
			//m_XMLRecost = new scXMLFunctions.XMLObject();
			m_XMLRecost = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			var tagActiveTag;
			m_XMLRecost.CreateRequestTagFromArray(m_sAttributeArray,null);
			if(m_sApplicationMode == "Quick Quote") tagActiveTag = m_XMLRecost.CreateActiveTag("QUICKQUOTE");
			else tagActiveTag = m_XMLRecost.CreateActiveTag("APPLICATIONQUOTE");
			m_XMLRecost.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
			m_XMLRecost.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
			
			m_XMLRecost.CreateTag("APPLICATIONDATE", sApplicationDate);
			
			<% /* PSC 15/07/2002 BMIDS00062
			 m_XMLRecost.CreateTag("LIFESUBQUOTENUMBER", m_sLifeSubQuoteNo); */ %>
			m_XMLRecost.CreateTag("CONTEXT", m_sApplicationMode);
			m_XMLRecost.CreateTag("INTERESTONLYAMOUNT", m_sINTERESTONLYAMOUNT);
			m_XMLRecost.CreateTag("CAPITALANDINTERESTAMOUNT", m_sCAPITALINTERESTAMOUNT);
			m_xmlREVISEDCOSTS.SelectTag(null, "RESPONSE");
			m_XMLRecost.CreateTag("APPLICATIONDATE", m_xmlREVISEDCOSTS.GetTagText("APPLICATIONDATE"));
			//DPF Nov 02 - BMIDS00808 - pass drawdown value into RecostMortgage function
			if( m_iDrawdown != null && m_iDrawdown != "0" && m_iDrawdown != "")
			{
				m_XMLRecost.CreateTag("DRAWDOWN", m_iDrawdown);
			}
			
			<% /* PSC 15/07/2002 BMIDS00062	
			m_XMLRecost.CreateActiveTag("LOANCOMPONENT")
			m_xmlREVISEDCOSTS.SelectTag(null, "LOANCOMPONENT");
			m_XMLRecost.AddXMLBlock(m_xmlREVISEDCOSTS.CreateFragment());
			PSC 15/07/2002 BMIDS00062 - End*/ %>
			
			m_XMLRecost.ActiveTag = tagActiveTag;
			m_XMLRecost.CreateActiveTag("LOANCOMPONENTLIST")
			m_xmlREVISEDCOSTS.SelectTag(null, "LOANCOMPONENTLIST");
			m_XMLRecost.AddXMLBlock(m_xmlREVISEDCOSTS.CreateFragment());
			m_XMLRecost.ActiveTag = tagActiveTag;
			m_XMLRecost.CreateActiveTag("ONEOFFCOSTLIST")
			xmlOneOffCost.SelectTag(null, "ONEOFFCOSTLIST");
			m_XMLRecost.AddXMLBlock(xmlOneOffCost.CreateFragment());
			m_XMLRecost.ActiveTag = tagActiveTag;
			m_XMLRecost.CreateActiveTag("CUSTOMERLIST")
			m_xmlCustomerList.SelectTag(null, "CUSTOMERLIST");
			m_XMLRecost.AddXMLBlock(m_xmlCustomerList.CreateFragment());
			//if(m_sApplicationMode == "Quick Quote") m_XMLRecost.RunASP(document, "QQRecostMortgageAndLifeCover.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			//DPF 10/10/2002 - applied manual fix to Screen Rules
			
			if(m_sApplicationMode == "Quick Quote")
			{
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
						m_XMLRecost.RunASP(document, "QQRecostMortgageAndLifeCover.asp");
						break;
					default: // Error
						m_XMLRecost.SetErrorResponse();
					}
			}
			else
			{
			// 			else m_XMLRecost.RunASP(document, "AQRecostMortgageAndLifeCover.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
						m_XMLRecost.RunASP(document, "AQRecostMortgageAndLifeCover.asp");
						break;
					default: // Error
					m_XMLRecost.SetErrorResponse();
					}
			}	
			if(m_XMLRecost.IsResponseOK())
			{
				UpdateOneOffCostsXML();
<%				/* re-calculate the totals fields */
%>				PopulateFields();			
				frmScreen.btnOK.disabled = false;
				frmScreen.btnCalculate.disabled = true;
			}
		}
		else m_xmlREVISEDCOSTS = null;
	}
	XML = null;
}
function UpdateOneOffCostsXML()
{
<%	/* take the returned oneoff costs xml from revisedCostsXml and update MIG in the listbox if necessary */
%>	
    m_xmlREVISEDCOSTS.ActiveTag = null;
	m_xmlREVISEDCOSTS.CreateTagList("ONEOFFCOST");
	var sNewMig = "";
	for (var nCost = 0; nCost < m_xmlREVISEDCOSTS.ActiveTagList.length && sNewMig == ""; nCost++)
	{
		if (m_xmlREVISEDCOSTS.SelectTagListItem(nCost) == true)
		{
			if(m_xmlREVISEDCOSTS.GetTagText("COMBOVALIDATIONTYPE") == "MIG")
			{
				sNewMig = m_xmlREVISEDCOSTS.GetTagText("AMOUNT");
			}
		}
	}
	if(sNewMig != "")
	{
		m_xmlGetOneOffCosts.ActiveTag = null;
		var bFound = false;
		for (var nCost = 0; nCost < m_xmlGetOneOffCosts.ActiveTagList.length && bFound == false; nCost++)
		{
			if (m_xmlGetOneOffCosts.SelectTagListItem(nCost) == true)
			{
				var sRetArray = GetComboValues(m_xmlGetOneOffCosts.GetTagText("MORTGAGEONEOFFCOSTTYPE"));
				if(sRetArray[1] == "MIG")
				{
					m_xmlGetOneOffCosts.SetTagText("AMOUNT", sNewMig);
					bFound = true;
					DisplayCosts(scScrollTable.getOffset());
				}
			}
		}
	}
}
function ConstructOneOffCostXML()
{
	//CMWP3 - reset Fees added to loan amount to 0 before recalculating
	m_iFeesAddedToLoanTotal = 0;
	//next line replaced with line below as per V7 - DPF 19/06/02 - BMIDS00077
	//var OneOffCostXML = new scXMLFunctions.XMLObject();
	var OneOffCostXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_xmlGetOneOffCosts.ActiveTag = null;
	var tagCOSTSLIST = m_xmlGetOneOffCosts.CreateTagList("MORTGAGEONEOFFCOST");
	var iNumberOfOneOffCosts = m_xmlGetOneOffCosts.ActiveTagList.length;
	var tagLIST = OneOffCostXML.CreateActiveTag("ONEOFFCOSTLIST");
	var sValidationType ; <% /* SR 12/07/2004 : BMIDS767 */ %>
	for (var nCost = 0; nCost < iNumberOfOneOffCosts; nCost++)
	{
		m_xmlGetOneOffCosts.ActiveTagList = tagCOSTSLIST;
		if (m_xmlGetOneOffCosts.SelectTagListItem(nCost) == true)
		{
			if(m_xmlGetOneOffCosts.GetTagText("AMOUNT") != "" && m_xmlGetOneOffCosts.GetTagText("AMOUNT") != "0")
			{
				<% /* DRC BMIDS767 var sName = m_xmlGetOneOffCosts.GetTagText("DESCRIPTION");
				if (sName == "NULL")sName=""; */ %>
				var sAmount = m_xmlGetOneOffCosts.GetTagText("AMOUNT");
				var sRetArray = GetComboValues(m_xmlGetOneOffCosts.GetTagText("MORTGAGEONEOFFCOSTTYPE"));
				<% /* SR 12/07/2004 : BMIDS767 */ %>
				if(!isTIDType(sRetArray)) 
					sValidationType = GetFeeValidationType(sRetArray, m_xmlGetOneOffCosts.GetTagText("IDENTIFIER"));
				else 
					continue;
				<% /* SR 12/07/2004 : BMIDS767 - End */ %>
				OneOffCostXML.ActiveTag = tagLIST;
				OneOffCostXML.CreateActiveTag("ONEOFFCOST");
				OneOffCostXML.CreateTag("COMBOVALIDATIONTYPE", sValidationType); <% /* SR 12/07/2004 : BMIDS767 */ %>
				<% /* DRC BMIDS767 OneOffCostXML.CreateTag("NAME", sName); */ %>	
				OneOffCostXML.CreateTag("AMOUNT", sAmount);
				OneOffCostXML.CreateTag("ADDEDTOLOAN", m_xmlGetOneOffCosts.GetTagText("ADDTOLOAN"));
				OneOffCostXML.CreateTag("ADHOCIND", m_xmlGetOneOffCosts.GetTagText("ADHOCIND"));
				<% /* MAR28 Add Refund Amount */ %>
				OneOffCostXML.CreateTag("REFUNDAMOUNT", m_xmlGetOneOffCosts.GetTagText("REFUNDAMOUNT"));
				
				//CMWP3 - DPF Jul 02 - if condition and sum added in
				if (m_xmlGetOneOffCosts.GetTagText("ADDTOLOAN") == "1")
				{
					m_iFeesAddedToLoanTotal += parseIntSafe(sAmount);
				}
			}
		}
	}
	return(OneOffCostXML);
}
<% /* BMIDS936 GHun Also exclude XAC, XDIS and END validation types */ %>
<% /* SR 12/07/2004 : BMIDS767 */ %>
function GetFeeValidationType(saValidationArray, si)
{
	var sRetValue = "";
	<% /* saValidationArray[0] contains the description of the fee */ %>
	for (var i=1; i<saValidationArray.length; ++i) 
	{
		if (!((saValidationArray[i] == "APR") || (saValidationArray[i] == "END") 
			|| (saValidationArray[i] == "XAC") || (saValidationArray[i] == "XDIS")))
		{ 
			sRetValue = saValidationArray[i];
			break ;
		}
	}
	
	return sRetValue
}
<% /* SR 12/07/2004 : BMIDS767 - End */ %>
<% /* BMIDS936 End */ %>

function frmScreen.btnOK.onclick() 
{
	if (m_sApplicationMode == "Mortgage Calculator")
	{
		m_XMLComboCosts = null;
		m_XMLOneOffCosts = null;
	}
	else
	{
		if(m_xmlREVISEDCOSTS != null)
		{
			//next line replaced with line below as per V7 - DPF 19/06/02 - BMIDS00077
			//var XML = new scXMLFunctions.XMLObject();
			var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTagFromArray(m_sAttributeArray,"UPDATE");
			
			
			var tagActiveTag = XML.CreateActiveTag("QUOTATION");
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
			XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortSubQuoteNo);
			XML.CreateTag("LIFESUBQUOTENUMBER", m_sLifeSubQuoteNo);
			m_xmlREVISEDCOSTS.SelectTag(null, "ADDEDONEOFFCOSTS");
			XML.CreateTag("LOANCOMPONENTSEQUENCENUMBER", m_xmlREVISEDCOSTS.GetTagText("LOANCOMPONENTSEQUENCENUMBER"));
			XML.CreateActiveTag("ONEOFFCOSTLIST");
			m_xmlREVISEDCOSTS.SelectTag(null, "ONEOFFCOSTLIST");
			XML.AddXMLBlock(m_xmlREVISEDCOSTS.CreateFragment());
			XML.ActiveTag = tagActiveTag;
			m_xmlREVISEDCOSTS.SelectTag(null, "ADDEDONEOFFCOSTS");
			XML.CreateTag("MIGIPT", m_xmlREVISEDCOSTS.GetTagText("MIGIPT"));
			XML.CreateTag("TOTALLOANAMOUNT", m_xmlREVISEDCOSTS.GetTagText("TOTALLOANAMOUNT"));
			XML.CreateTag("TOTALINCENTIVES", frmScreen.txtTotalIncentives.value);
			//DPF 10/10/2002 - CPWP1 (BM020) - New field
			XML.CreateTag("MANUALINCENTIVES", frmScreen.txtManualIncentives.value);
			XML.CreateTag("INTERESTONLYAMOUNT", m_sINTERESTONLYAMOUNT);
			XML.CreateTag("CAPITALANDINTERESTAMOUNT", m_sCAPITALINTERESTAMOUNT);
			<% /*
			m_XMLRecost.SelectTag(null, "RESPONSE");
			XML.CreateTag("COMPONENTNETMONTHLYCOST", m_XMLRecost.GetTagText("COMPONENTNETMONTHLYCOST"));
			XML.CreateTag("COMPONENTGROSSMONTHLYCOST", m_XMLRecost.GetTagText("COMPONENTGROSSMONTHLYCOST"));
			XML.CreateTag("OUT020_ACCRUEDINTEREST", m_XMLRecost.GetTagText("OUT020_ACCRUEDINTEREST"));
			XML.CreateTag("APR", m_XMLRecost.GetTagText("APR"));
			XML.CreateTag("OUT040_TOTALAMOUNTPAYABLE", m_XMLRecost.GetTagText("OUT040_TOTALAMOUNTPAYABLE"));
			XML.CreateTag("OUT050_TOTALMORTGAGEPAYMENTS", m_XMLRecost.GetTagText("OUT050_TOTALMORTGAGEPAYMENTS"));
			XML.CreateTag("OUT070_TOTALNETMORTGAGEPAYMENT2", m_XMLRecost.GetTagText("OUT070_TOTALNETMORTGAGEPAYMENT2"));
			XML.CreateTag("OUT080_TOTALNETMORTGAGEPAYMENT3", m_XMLRecost.GetTagText("OUT080_TOTALNETMORTGAGEPAYMENT3"));
			XML.CreateTag("OUT090_TOTALNETMORTGAGEPAYMENT4", m_XMLRecost.GetTagText("OUT090_TOTALNETMORTGAGEPAYMENT4"));
			XML.CreateTag("OUT100_TOTALNETMORTGAGEPAYMENT5", m_XMLRecost.GetTagText("OUT100_TOTALNETMORTGAGEPAYMENT5"));
			XML.CreateTag("OUT110_TOTALNETMORTGAGEPAYMENT6", m_XMLRecost.GetTagText("OUT110_TOTALNETMORTGAGEPAYMENT6"));
			XML.CreateTag("OUT120_TOTALNETMORTGAGEPAYMENT7", m_XMLRecost.GetTagText("OUT120_TOTALNETMORTGAGEPAYMENT7"));
			XML.CreateTag("OUT140_TOTALGROSSMORTGAGEPAYMENT2", m_XMLRecost.GetTagText("OUT140_TOTALGROSSMORTGAGEPAYMENT2"));
			XML.CreateTag("OUT150_TOTALGROSSMORTGAGEPAYMENT3", m_XMLRecost.GetTagText("OUT150_TOTALGROSSMORTGAGEPAYMENT3"));
			XML.CreateTag("OUT160_TOTALGROSSMORTGAGEPAYMENT4", m_XMLRecost.GetTagText("OUT160_TOTALGROSSMORTGAGEPAYMENT4"));
			XML.CreateTag("OUT170_TOTALGROSSMORTGAGEPAYMENT5", m_XMLRecost.GetTagText("OUT170_TOTALGROSSMORTGAGEPAYMENT5"));
			XML.CreateTag("OUT180_TOTALGROSSMORTGAGEPAYMENT6", m_XMLRecost.GetTagText("OUT180_TOTALGROSSMORTGAGEPAYMENT6"));
			XML.CreateTag("OUT190_TOTALGROSSMORTGAGEPAYMENT7", m_XMLRecost.GetTagText("OUT190_TOTALGROSSMORTGAGEPAYMENT7"));
			XML.CreateTag("OUT200_UNROUNDEDAPR", m_XMLRecost.GetTagText("OUT200_UNROUNDEDAPR"));
			XML.CreateTag("OUT210_FINALGROSSPAYMENT", m_XMLRecost.GetTagText("OUT210_FINALGROSSPAYMENT"));
			*/ %>
			
			XML.CreateTag("TOTALGROSSMONTHLYCOST", m_XMLRecost.GetTagText("TOTALGROSSMONTHLYCOST"));
			XML.CreateTag("TOTALNETMONTHLYCOST", m_XMLRecost.GetTagText("TOTALNETMONTHLYCOST"));
			
			//DPF 7/11/2002 - add in Monthly Cost Less Drawdown to saving routine (if there is one)			
			if( m_iDrawdown != null && m_iDrawdown != "0" && m_iDrawdown != "")	
			{
				var xmlNode = null;
				var sMCLDAmount = 0;
				xmlNode = m_XMLRecost.XMLDocument.selectSingleNode("//RESPONSE/MONTHLYCOSTLESSDRAWDOWN");
				if (xmlNode != null)
				{
					sMCLDAmount = xmlNode.text;
				}
					
				XML.CreateTag("MONTHLYCOSTLESSDRAWDOWN", sMCLDAmount);
     		}
     		//END OF DPF CHANGE
			
			<% /* BMIDS766 GHun */ %>
			XML.CreateTag("AMOUNTPERUNITBORROWED", m_XMLRecost.GetTagText("AMOUNTPERUNITBORROWED"));
			XML.CreateTag("APR", m_XMLRecost.GetTagText("APR"));
			XML.CreateTag("TOTALAMOUNTPAYABLE", m_XMLRecost.GetTagText("TOTALAMOUNTPAYABLE"));
			XML.CreateTag("TOTALMORTGAGEPAYMENTS", m_XMLRecost.GetTagText("TOTALMORTGAGEPAYMENTS"));
			<% /* BMIDS766 End */ %>
			
			XML.CreateTag("TOTALACCRUEDINTEREST", m_XMLRecost.GetTagText("TOTALACCRUEDINTEREST"));
			XML.CreateTag("CONTEXT", "UPDATE");
			XML.CreateActiveTag("LOANCOMPONENTLIST");
			m_XMLRecost.SelectTag(null, "RESPONSE/LOANCOMPONENTLIST");
			XML.AddXMLBlock(m_XMLRecost.CreateFragment());

			<% /*
			//Add the Life Benefit and Life Cover stuff
			XML.CreateTag("TOTALLIFEMONTHLYCOST", m_XMLRecost.GetTagText("TOTALLIFEMONTHLYCOST"));
			if(m_XMLRecost.SelectTag(null, "LIFEBENEFITS") != null)
			{
				XML.CreateActiveTag("LIFEBENEFITS")
				XML.AddXMLBlock(m_XMLRecost.CreateFragment())
				//XML.ActiveTag = tagActiveTag;
				var tagCustList = m_xmlCustomerList.CreateTagList("CUSTOMER");
				for (var nCust = 0; nCust < m_xmlCustomerList.ActiveTagList.length; nCust++)
				{
					m_xmlCustomerList.ActiveTagList = tagCustList;
					if (m_xmlCustomerList.SelectTagListItem(nCust) == true)
					{
						var sCust = "CUSTOMERNUMBER" + (nCust +1);
						var sCustVer = "CUSTOMERVERSIONNUMBER" + (nCust + 1);
						XML.CreateTag(sCust, m_xmlCustomerList.GetTagText("CUSTOMERNUMBER"));
						XML.CreateTag(sCustVer, m_xmlCustomerList.GetTagText("CUSTOMERVERSIONNUMBER"));
					}
				}
			}
		
			PSC 15/07/2002 BMIDS00062 - End */ %>
			
			// 			XML.RunASP(document, "SaveOneOffCostDetails.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "SaveOneOffCostDetails.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

			if(XML.IsResponseOK())
			{
				var sReturn = new Array();
				sReturn[0] = IsChanged();
				window.returnValue = sReturn;
			}
		}
	}
	window.close();
}
		
function frmScreen.btnCancel.onclick()
{
	window.close();
}

<%/* GD BM0368 Added  End */%>
function parseIntSafe(sText)
{
	if (sText == "") return(0)
	else
	return(parseInt(sText));
}
<%/* GD BM0368 Added  End */%>

-->
</script>
</body>
</html>
