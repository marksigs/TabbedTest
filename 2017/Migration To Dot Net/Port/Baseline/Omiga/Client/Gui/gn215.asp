<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<% /*
Workfile:      gn215.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Extended Search Results
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MV		31/10/2000	AQR No - CORE000002 - Created
ASm		21/11/2000	CORE000010: Various minor GUI changes
ASm		21/11/2000	CORE000010: Increased PAGESIZE parameter to match table page size parameter
JLD		04/01/2001	SYS1767 Use LockApplication method from include file LockApp.asp instead.
JLD		08/01/2001	SYS1767 Move LockApp.asp file to the includes directory
APS		08/01/2001	SYS1782 Added Back button processing
APS		08/01/2001	SYS1803 Read Only processing
MV		15/01/2001  SYS1819 Changed ShowRow and SetContextDataOnExistingbusiness Procs
					Added ApplicationNumber Attribute in ShowRow Proc 
					changed m_sApplicationNumber Value from RowInnerText to Attribute Based Value
					Changed frmScreen.btnSelect.onclick from RowInnerText to Attribute Based Value
MV		15/02/2001	SYS1940 Removed the function from SetContextDataOnExistingBusiness from frmScreen.btnSelect.onclick
CL		05/03/01	SYS1920 Read only functionality added
APS		11/03/01	SYS2030 Reinstate Cancelled/Declined processing added
BG		11/04/01	SYS2096 Added Printing functionality 
JLD		08/10/01	SYS2736 Don't download omPC again
AT		08/04/02	SYS4359 Add SetCurrency() for internationalisation
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		15/08/2002	BMIDS00122	Added spnTable.ondblclick()
DRC     25/09/2002  BMIDS00122	Added Boolean switch m_bIsSelectOk to stop > 1 click on the select button
DRC     01/10/2002  BMIDS00122	Moved  Boolean switch off  m_bIsSelectOk & Select disable logic
MO		02/10/2002	BMIDS00502	Made changes to the GUI to implement BM's Introducer system
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		09/10/2002	BMIDS00556	Removed Validation_Init()
MO		31/10/2002	BMISD00685	Modified the column widths of the search results table
GHun	05/12/2002	BM0035		CC011 Customer business search in application enquiry
GHun	03/06/2003	BM0523		Removed TotalLoanAmount, ApplicationPriority, ApplicationStageNumber and
								ApplicationStageSequenceNumber as they are not required
GHun	23/10/2003	BMIDS624	Check if rates have changed when selecting an application
MC		30/04/2004	BM0468		Cancel and Decline stage freeze screens functionality
JD		28/07/2004	BMIDS749 Removed check for changed rates.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR		Description
MV		12/10/2002	MAR22	Amended frmScreen.btnReinstate.onclick()
PSC		25/10/2005	MAR300  Amend call to LockApplication to synchronise customer details
MV		04/11/2002	MAR236	Amended frmScreen.btnReinstate.onclick()
PJO     13/12/2005  MAR815	Disable print button permanently
HMA     15/02/2006  MAR1164/MAR1204  Changes to Reinstate processing.
GHun	17/03/2006	MAR1453 Progress window
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
IK		05/05/2006	EP311		back-out MAR1204 & more, remove FirstTitle processing
SR		04/03/2007	EP2_1644	increased the width of scScrollPlus control
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" viewastext tabIndex="-1"></object>
<%/* SCRIPTLETS */%>

<object data=scTable.htm height=1 id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet viewastext></object>
<span style="LEFT: 306px; POSITION: absolute; TOP: 435px">
	<object data=scPageScroll.htm id=scScrollPlus style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet viewastext></object>
</span>
<form id="frmToGN200" method="post" action="gn200.asp" style="DISPLAY: none"></form>
<form id="frmToGN210" method="post" action="gn210.asp" style="DISPLAY: none"></form>
<form id="frmWelcomeMenu" method="post" action="mn010.asp" style="DISPLAY: none"></form>
<form id="frmToMN060" method= "post" action = "mn060.asp" style="DISPLAY: none"></form>

<form id="frmScreen" mark  validate ="onchange">
<div id="divQuickSearch" style="HEIGHT: 115px; LEFT: 10px; POSITION: absolute; TOP: 60px; VISIBILITY: hidden; WIDTH: 604px" class="msgGroup" >
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Application Number
		<span style="LEFT: 102px; POSITION: absolute; TOP: -3px">
			<input id="txtApplicationNumber" maxlength="12" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 130px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 365px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Account Number
		<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountNumber" maxlength="20" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 136px" class="msgTxt">
		</span>
	</span>

	<% /* BM0035 */ %>
	<span style="TOP: 37px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Surname
		<span style="TOP: -3px; LEFT: 102px; POSITION: ABSOLUTE">
			<input id="txtSurname" maxlength="40" style="HEIGHT: 20px; WIDTH: 220px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 37px; LEFT: 365px; POSITION: ABSOLUTE" class="msgLabel">
		Date of Birth
		<span style="TOP: -3px; LEFT: 90px; POSITION: ABSOLUTE">
			<input type="text" id="txtDateOfBirth" maxlength="10" style="HEIGHT: 20px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 64px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Forename(s)
		<span style="TOP: -3px; LEFT: 102px; POSITION: ABSOLUTE">
			<input id="txtForenames" maxlength="40" style="HEIGHT: 20px; WIDTH: 220px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 64px; LEFT: 365px; POSITION: ABSOLUTE" class="msgLabel">
		PostCode
		<span style="TOP: -3px; LEFT: 90px; POSITION: ABSOLUTE">
			<input id="txtPostcode" maxlength="8" style="HEIGHT: 20px; WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxtUpper">
		</span>
	</span>	
	<% /* BM0035 End */ %>
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 91px" class="msgLabel">
		Unit Name 
		<span style="LEFT: 102px; POSITION: absolute; TOP: -3px">
			<input id="txtQuickUnitName"  maxlength="10" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 220px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 365px; POSITION: absolute; TOP: 91px" class="msgLabel">
		User Name
		<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
			<input id="txtQuickUserName"  maxlength="40" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 136px" class="msgTxt">
		</span>
	</span>	
</div>

<div id="divExtendedSearch" style="HEIGHT: 115px; LEFT: 10px; POSITION: absolute; TOP: 60px; VISIBILITY: hidden; WIDTH: 604px" class="msgGroup">
	<span id = "spnExtendedSearch">
		<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Unit 
			<span style="LEFT: 64px; POSITION: absolute; TOP: -3px">
				<input id="txtExtendedUnitName"  maxlength="10" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 110px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 195px; POSITION: absolute; TOP: 10px" class="msgLabel" id=SPAN1>
			User
			<span style="LEFT: 68px; POSITION: absolute; TOP: -3px">
				<input id="txtExtendedUserName"  maxlength="40" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 120px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 395px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Channel
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtChannelId"  maxlength="40" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 110px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 36px" class="msgLabel">
			Department
			<span style="LEFT: 64px; POSITION: absolute; TOP: -3px">
				<input id="txtDepartmentName"  maxlength="40" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 110px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 195px; POSITION: absolute; TOP: 36px" class="msgLabel">
			Lender
			<span style="LEFT: 68px; POSITION: absolute; TOP: -3px">
				<input id="txtLenderName"  maxlength="10" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 120px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 395px; POSITION: absolute; TOP: 36px" class="msgLabel">
			Mortgage Product
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtMortgageProductName"  maxlength="40" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 110px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 56px" class="msgLabel">
			Type Of <br>Application
			<span style="LEFT: 64px; POSITION: absolute; TOP: 3px">
				<input id="txtTypeOfApplication"  maxlength="40" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 110px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 195px; POSITION: absolute; TOP: 56px" class="msgLabel">
			Approval <br>Status
			<span style="LEFT: 68px; POSITION: absolute; TOP: 3px">
				<input id="txtApprovalStatus"  maxlength="10" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 120px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 395px; POSITION: absolute; TOP: 62px" class="msgLabel">
			From Stage
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtFromStage"  maxlength="40" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 110px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 89px" class="msgLabel">
			To Stage
			<span style="LEFT: 64px; POSITION: absolute; TOP: -3px">
				<input id="txtToStage"  maxlength="40" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 110px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 195px; POSITION: absolute; TOP: 89px" class="msgLabel">
			Date Range
			<span style="LEFT: 68px; POSITION: absolute; TOP: -3px">
				<input id="txtDateRange"  maxlength="10" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 120px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 395px; POSITION: absolute; TOP: 89px" class="msgLabel">
			Introducer
			<span style="LEFT: 95px; POSITION: absolute; TOP: -3px">
				<input id="txtIntroducerName"  maxlength="40" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 110px" class="msgTxt">
			</span>
		</span>
	</span>
</div>

<div id="divDispTable" style = "HEIGHT: 314px; LEFT: 0px; POSITION: relative; TOP: 170px; WIDTH: 605px" class="msgGroup">
	<div id="spnTable" style="LEFT: 5px; POSITION: absolute; TOP: 5px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="20%" class="TableHead">Customer Name</td>
				<td width="15%" class="TableHead">Application  Number</td>
				<td width="10%" class="TableHead">CurrentStage/ LockStatus</td>
				<td width="13%" class="TableHead">Date Created</td>
				<td width="15%" class="TableHead">Application Type</td>
				<td width="10%" class="TableHead">Loan Amount</td>
				<td width="17%" class="TableHead">Product Name</td>
			</tr>
			<tr id="row01">		
				<td  class="TableTopLeft">&nbsp;</td>		
				<td  class="TableTopCenter">&nbsp;</td>		
				<td  class="TableTopCenter">&nbsp;</td>		
				<td  class="TableTopCenter">&nbsp;</td>		
				<td  class="TableTopCenter">&nbsp;</td>		
				<td  class="TableTopCenter">&nbsp;</td>		
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row05">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row06">		
			<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row07">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row08">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row09">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row10">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row11">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row12">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row13">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row14">		
				<td  class="TableLeft">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td  class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row15">		
				<td  class="TableBottomLeft">&nbsp;</td>		
				<td  class="TableBottomCenter">&nbsp;</td>		
				<td  class="TableBottomCenter">&nbsp;</td>		
				<td  class="TableBottomCenter">&nbsp;</td>		
				<td  class="TableBottomCenter">&nbsp;</td>		
				<td  class="TableBottomCenter">&nbsp;</td>		
				<td class="TableBottomRight">&nbsp;</td>
			</tr>
		</table>
	</div>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 259px" class="msgLabel">
		* Locked Record		
	</span>
	<% /* BM0523 removed for performance reasons
	<span style="LEFT: 4px; POSITION: absolute; TOP: 288px" class="msgLabel" id=SPAN1>
		Total Loan Amount
		<span style="HEIGHT: 14px; LEFT: 17px; POSITION: absolute; TOP: -3px; WIDTH: 44px">
			<label style="TOP: 1px; LEFT: 87px; POSITION: relative" class="msgCurrency"></label>
			<input id="txtTotalLoanAmount" maxlength = "10" style="HEIGHT: 20px; LEFT: 97px; POSITION: absolute; TOP: -1px; WIDTH: 133px" class ="msgTxt" size=26>
		</span>
	</span>
	*/ %>
	<span style="LEFT: 325px; POSITION: absolute; TOP: 284px">
		<input id="btnPrint" value="Print" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
	<span style="LEFT: 425px; POSITION: absolute; TOP: 284px">
		<input id="btnSelect" value="Select" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
	<span style="LEFT: 525px; POSITION: absolute; TOP: 284px">
		<input id="btnReinstate" value="Reinstate" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
</div>	
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 4px; POSITION: absolute; TOP: 505px; WIDTH: 377px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen  */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->	

<% /*Include the file to lock the application */ %>
<!-- #include FILE="Includes/LockApp.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/gn215Attribs.asp" -->
	
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = "";
var DOWNLOADXTABLE = 15;
var xmlStoredData = null;
var xmlRequestTag = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sApplicationType = "";
var m_sApplicationTypeNo = "";
var m_sApplicationStageName = "";
//BMIDS00122 DRC module level switch
var m_bIsSelectOk = true;
var xmlSearchData = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();

	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
	var sButtonList = new Array("Back","Cancel");
	ShowMainButtons(sButtonList);
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	FW030SetTitles("Application Enquiry - Search Results ","GN215",scScreenFunctions); 
 	scScreenFunctions.SetCurrency(window, frmScreen);

	RetrieveContextData();
	scTable.initialise(tblTable, 0, "");
	window.status = "Searching...";
	frmScreen.style.cursor = "wait";
	window.setTimeout (PopulateScreen, 0);	
	frmScreen.btnPrint.disabled = true;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveContextData()
{
	// This is a Function to Retrieve Context Data and Fetch all the Records from the DB
	// Load into the Screen Controls and table .Search Criteria is from idXML String
	// passed from gn200.asp( QucikSearch)  or gn210.asp (ExtendedSearch)
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	var sStoredDataXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	xmlStoredData = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	xmlStoredData.LoadXML(sStoredDataXML);
	xmlRequestTag = xmlStoredData.SelectTag(null,"REQUEST");
	if (xmlRequestTag != null) 
	{
		if (m_sMetaAction == "QUICKSEARCH" ) 
		{
			//Make divQuickSearch Division visible
			scScreenFunctions.ShowCollection(divQuickSearch);
			// Load Screen Controls with the tag Text
			frmScreen.txtApplicationNumber.value = xmlStoredData.GetTagText("APPLICATIONNUMBER");
			frmScreen.txtAccountNumber.value = xmlStoredData.GetTagText("ACCOUNTNUMBER");
			frmScreen.txtQuickUnitName.value = xmlStoredData.GetTagText("UNITNAME");
			frmScreen.txtQuickUserName.value = xmlStoredData.GetTagText("USERNAME");				
			// Make them Read Only
			scScreenFunctions.SetFieldState(frmScreen, "txtApplicationNumber","R");
			scScreenFunctions.SetFieldState(frmScreen, "txtAccountNumber","R");
			scScreenFunctions.SetFieldState(frmScreen, "txtQuickUnitName","R");
			scScreenFunctions.SetFieldState(frmScreen, "txtQuickUserName","R");
			<% /* BM0035 */ %>
			frmScreen.txtSurname.value = xmlStoredData.GetTagText("SURNAME");
			frmScreen.txtForenames.value = xmlStoredData.GetTagText("FIRSTFORENAME") + ' ' + xmlStoredData.GetTagText("SECONDFORENAME") + ' ' + xmlStoredData.GetTagText("OTHERFORENAMES");
			frmScreen.txtDateOfBirth.value = xmlStoredData.GetTagText("DATEOFBIRTH");
			frmScreen.txtPostcode.value = xmlStoredData.GetTagText("POSTCODE");
			scScreenFunctions.SetFieldState(frmScreen, "txtSurname","R");
			scScreenFunctions.SetFieldState(frmScreen, "txtForenames","R");
			scScreenFunctions.SetFieldState(frmScreen, "txtDateOfBirth","R");
			scScreenFunctions.SetFieldState(frmScreen, "txtPostcode","R");
			<% /* BM0035 End */ %>
		}
		else if (m_sMetaAction == "EXTENDEDSEARCH" ) 
		{
			//Make divExtendedSearch Division visible
			scScreenFunctions.ShowCollection(divExtendedSearch);
			// Load Screen Controls with the tag Text
			// Make them Read Only
			frmScreen.txtExtendedUnitName.value = xmlStoredData.GetTagText("UNITNAME"); 
			scScreenFunctions.SetFieldState(frmScreen, "txtExtendedUnitName","R");
			
			frmScreen.txtExtendedUserName.value = xmlStoredData.GetTagText("USERNAME"); 
			scScreenFunctions.SetFieldState(frmScreen, "txtExtendedUserName","R");
			
			frmScreen.txtChannelId.value = xmlStoredData.GetTagText("CHANNELNAME"); 
			scScreenFunctions.SetFieldState(frmScreen, "txtChannelId","R");
			
			frmScreen.txtDepartmentName.value = xmlStoredData.GetTagText("DEPARTMENTNAME"); 
			scScreenFunctions.SetFieldState(frmScreen, "txtDepartmentName","R");
			
			frmScreen.txtLenderName.value = xmlStoredData.GetTagText("LENDERNAME"); 
			scScreenFunctions.SetFieldState(frmScreen, "txtLenderName","R");
			
			frmScreen.txtMortgageProductName.value = xmlStoredData.GetTagText("MORTGAGEPRODUCTNAME"); 
			scScreenFunctions.SetFieldState(frmScreen, "txtMortgageProductName","R");
			
			frmScreen.txtTypeOfApplication.value = xmlStoredData.GetTagText("APPLICATIONTYPENAME"); 
			scScreenFunctions.SetFieldState(frmScreen, "txtTypeOfApplication","R");
			
			// APS 09/01/2001 : Wrong tag name
			if ( xmlStoredData.GetTagText("APPLICATIONAPPROVEDMONTH") != "") 
				frmScreen.txtApprovalStatus.value = xmlStoredData.GetTagText("APPLICATIONAPPROVEDMONTHTEXT")  + " " +   xmlStoredData.GetTagText("APPLICATIONAPPROVEDYEAR");
			
			scScreenFunctions.SetFieldState(frmScreen, "txtApprovalStatus","R");
				
			// APS 09/01/2001 : Wrong tag name
			frmScreen.txtFromStage.value = xmlStoredData.GetTagText("APPLICATIONSTAGEFROMNAME"); 
			scScreenFunctions.SetFieldState(frmScreen, "txtFromStage","R");
			
			frmScreen.txtToStage.value = xmlStoredData.GetTagText("APPLICATIONSTAGETONAME"); 
			scScreenFunctions.SetFieldState(frmScreen, "txtToStage","R");
			
			// APS Wrong tag names
			if (xmlStoredData.GetTagText("APPLICATIONDATEFROM") != "") {
				frmScreen.txtDateRange.value = xmlStoredData.GetTagText("APPLICATIONDATEFROM")+ " - " +  xmlStoredData.GetTagText("APPLICATIONDATETO")  ;
				//scScreenFunctions.SetFieldState(frmScreen, "txtDateRange","R");
				}
			scScreenFunctions.SetFieldState(frmScreen, "txtDateRange","R");
			
			<% /* MO 02/10/2002 BMIDS00502 */ %>
			if (xmlStoredData.GetTagText("INTRODUCERNAME") != "") {
				frmScreen.txtIntroducerName.value = xmlStoredData.GetTagText("INTRODUCERNAME");
			} else {
				frmScreen.txtIntroducerName.value = xmlStoredData.GetTagText("INTRODUCERID");
			}
			scScreenFunctions.SetFieldState(frmScreen, "txtIntroducerName","R");
		}
	}
	
	//SYS2030
	if (xmlStoredData.GetTagText("INCCANCELLEDORDECLINEDAPPSCHECKED") == "") frmScreen.btnReinstate.disabled = true;
	
	<% /* BM0523 No longer required
	scScreenFunctions.SetFieldState(frmScreen, "txtTotalLoanAmount","R"); */ %>
}

function btnBack.onclick() 
{	
	// APS 08/01/2001 : SYS1782
	//scScreenFunctions.SetContextParameter(window,"idXML","");
	if (m_sMetaAction == "QUICKSEARCH")
		frmToGN200.submit();
	else if (m_sMetaAction == "EXTENDEDSEARCH")
		frmToGN210.submit();
}

function btnCancel.onclick() 
{	
	// APS 08/01/2001 : SYS1782
	scScreenFunctions.SetContextParameter(window,"idXML","");
	frmWelcomeMenu.submit();
}

function PopulateScreen()
{
	scTable.clear();
//	scScrollPlus.Clear();
	scScrollPlus.Initialise(FindApplications,ShowList,15,DOWNLOADXTABLE);
	if(scScrollPlus.GetTotalRows() > 0)	 scTable.setRowSelected(1);
	frmScreen.style.cursor = "default";
	window.status = "";
}

function FindApplications(nPageNo)
{
	scScreenFunctions.progressOn("Searching ...", 400);	<% /* MAR1453 GHun */ %>
	// This is a function to retrieve data from database for the criteria
	var sStoredDataXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	xmlStoredData = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	xmlStoredData.LoadXML(sStoredDataXML);
	xmlRequestTag = xmlStoredData.SelectTag(null,"REQUEST");
	if (xmlRequestTag != null) 
	{
		// using previously declared xmlRequestTag add in attributes
		xmlStoredData.SetAttribute("PAGESIZE",DOWNLOADXTABLE * 15);
		xmlStoredData.SetAttribute("PAGENUMBER",nPageNo);
	}	
	xmlStoredData.RunASP(document,"FindApplicationList.asp");
	scScreenFunctions.progressOff();	<% /* MAR1453 GHun */ %>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = xmlStoredData.CheckResponse(ErrorTypes);
	if(ErrorReturn[0])
	{
		if(xmlStoredData.SelectTag(null,"APPLICATIONLIST") != null)
			return xmlStoredData.GetAttribute("TOTALRECORDS");
	}
	else if(ErrorReturn[1] == ErrorTypes[0])
	{
		alert("Unable to find any Applications for your search criteria.");
		frmScreen.btnPrint.disabled = true;
  //BMIDS00122 - DRC 		
		frmScreen.btnSelect.disabled = true;
	}	
	return 0;
}

function ShowRow(nIndex,sCustomerName,sApplicationNumber,sApplicationFactFindNumber,
				sApplicationType,iApplicationTypeNo,sDateCreated,sCurrentStage,
				sApplicationLocked,sLoanAmount,sProductName)
{
	// This is a function to fill the tablerow with the corresponding passing parameters
	tblTable.rows(nIndex).setAttribute("ApplicationNumber",sApplicationNumber);			 					
	tblTable.rows(nIndex).setAttribute("ApplicationFactFindNumber", sApplicationFactFindNumber);
	tblTable.rows(nIndex).setAttribute("ApplicationType", sApplicationType);
	
	if (sApplicationLocked == "1") sCurrentStage = "*" + sCurrentStage;
	
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(0),sCustomerName);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(1),sApplicationNumber);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(2),sCurrentStage);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(3),sDateCreated);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(4),sApplicationType);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(5),sLoanAmount);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(6),sProductName);
}		

function ShowList(nStart)
{
	var i0=0, i1=0, iNumberOfApps=0, iApplicationTypeNo=0;
	var sApplicationNumber = "", sApplicationFactFindNumber = "";
	var sApplicationType = "", sDateCreated = "";
	var sCurrentStage = "", sApplicationLocked = "";
	var sLoanAmount = "", sProductName = "", sFirstForename = "";
	var sCustomerName = "";
	
	scTable.clear();
	
	var xmlTempRequestTag = xmlStoredData.SelectTag(null,"APPLICATIONLIST");			
	if (xmlTempRequestTag != null) 
	{
		iNumberOfApps = xmlStoredData.GetAttribute("TOTALRECORDS")
		<% /* BM0523 No longer required
		frmScreen.txtTotalLoanAmount.value = xmlStoredData.GetAttribute("TOTALLOANAMOUNT"); */ %>
	 		 	
		xmlStoredData.CreateTagList("APPLICATION");  
	 				
		for (i0 = 0; nStart + i0 < iNumberOfApps && i0 < DOWNLOADXTABLE ; i0++)
		{				
			if (xmlStoredData.SelectTagListItem(nStart+i0) == true) {
				
				sApplicationNumber = xmlStoredData.GetTagText("APPLICATIONNUMBER");
				sApplicationFactFindNumber = xmlStoredData.GetTagText("APPLICATIONFACTFINDNUMBER");
				sApplicationType = xmlStoredData.GetTagAttribute("TYPEOFAPPLICATION","TEXT");
				iApplicationTypeNo = xmlStoredData.GetTagText("TYPEOFAPPLICATION");
				sDateCreated = xmlStoredData.GetTagText("APPLICATIONDATE");
				sCurrentStage = xmlStoredData.GetTagText("APPLICATIONSTAGENAME");
				sApplicationLocked = xmlStoredData.GetTagText("APPLICATIONLOCKED");
				sLoanAmount = xmlStoredData.GetTagText("TOTALLOANAMOUNT");
				sProductName = xmlStoredData.GetTagText("PRODUCTNAME");
				sCustomerName = xmlStoredData.GetTagText("CORRESPONDENCESALUTATION");
				
				ShowRow (i0+1,sCustomerName,sApplicationNumber,sApplicationFactFindNumber,sApplicationType,iApplicationTypeNo,sDateCreated,
						 sCurrentStage,sApplicationLocked,sLoanAmount,sProductName);
				<% // PJO 13/12/2005 MAR 815 - permanently disbale button	
				//frmScreen.btnPrint.disabled = false; %>
			}
		}
	}	
}

function frmScreen.btnSelect.onclick()
{	
  //BMIDS00122 DRC Module level switch to prevent lock on double click
   if (m_bIsSelectOk)
   { 
		var iSelectedRow = scTable.getRowSelectedId();
	
		if (iSelectedRow == null){
			alert ("You must select an Application.");
			return;
		}

		<%/* BMIDS624 */ %>
		var sApplicationNumber = tblTable.rows(iSelectedRow).getAttribute("ApplicationNumber");
		var sApplicationFactFindNumber = tblTable.rows(iSelectedRow).getAttribute("ApplicationFactFindNumber");
		<%/* BMIDS624 End */ %>

		// APS 08/01/2001 : SYS1803
	
		//if(LockApplication("0", tblTable.rows(scTable.getRowSelectedId()).cells(1).innerText, 
		<% /* PSC 25/10/2005 MAR300 */ %>
		if (LockApplication("0", sApplicationNumber, sApplicationFactFindNumber, null, null, true))
		{
		      //BMIDS00122 revisited DRC - switch moved 
		    m_bIsSelectOk = false;
		    <%/* BMIDS624 */ %>
		    // JD BMIDS749 removed    CheckForChangedRates(sApplicationNumber, sApplicationFactFindNumber);
		    <%/* BMIDS624 End */ %>
			frmToMN060.submit();
		}
	}	
}	

function spnTable.ondblclick()
{
	if (scTable.getRowSelectedId() != null)
		frmScreen.btnSelect.onclick();
}


function frmScreen.btnReinstate.onclick()
{
	// APS 11/03/01 SYS2030
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
	var sApplicationNumber = tblTable.rows(scTable.getRowSelectedId()).getAttribute("ApplicationNumber");
	var sApplicationFactFindNumber = tblTable.rows(scTable.getRowSelectedId()).getAttribute("ApplicationFactFindNumber");	
	var sApplicationType = tblTable.rows(scTable.getRowSelectedId()).getAttribute("ApplicationType");
	
	GetApplicationPriority(sApplicationNumber);
	<% /* BM0523
	var sApplicationPriorityValue = tblTable.rows(scTable.getRowSelectedId()).getAttribute("ApplicationPriorityValue"); */ %>
	
	var tagRequest = XML.CreateRequestTag(window, "ReinstatePreviousStage");
	<% /* BM0523
	XML.SetAttribute("APPLICATIONPRIORITY", sApplicationPriorityValue);	*/ %>
	XML.CreateActiveTag("CASEACTIVITY");
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("CASEID", sApplicationNumber);
	XML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window, "idActivityId", "10"));
	XML.ActiveTag = tagRequest;
	XML.CreateActiveTag("APPLICATION");
	XML.SetAttribute("APPLICATIONNUMBER", sApplicationNumber);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	XML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window,"idApplicationPriority",""));
	// 	XML.RunASP(document, "OmigaTMBO.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "OmigaTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if (XML.IsResponseOK() == true)
	{
		scScreenFunctions.SetContextParameter(window,"idCancelDeclineFreezeDataIndicator", "0");
		if (LockApplication("0", sApplicationNumber, sApplicationFactFindNumber) != 0) 
		{
			<%/* BMIDS624 */ %>
			// JD BMIDS749 removed   CheckForChangedRates(sApplicationNumber, sApplicationFactFindNumber);
			<%/* BMIDS624 End */ %>
			frmToMN060.submit();
		}
	}			
	XML = null;
}

<% /* MO 02/10/2002 BMIDS00502 */ %>
function frmScreen.btnPrint.onclick()
{	
	xmlStoredData.SelectTag(null,"RESPONSE");	
	var iNumRecords = xmlStoredData.GetTagAttribute("APPLICATIONLIST", "TOTALRECORDS");
	if(confirm("You are about to print " + iNumRecords + " records.  Do you wish to continue?"))
	{	
		var GlobalXML = null;
		GlobalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var iDocumentID = parseInt(GlobalXML.GetGlobalParameterAmount(document,"PrintAppEnqHostTemplateId"));

		AttribsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		AttribsXML.CreateRequestTag(window, "GetPrintAttributes");
		AttribsXML.CreateActiveTag("FINDATTRIBUTES");
		AttribsXML.SetAttribute("HOSTTEMPLATEID", iDocumentID);
					
		// 		AttribsXML.RunASP(document, "PrintManager.asp");
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
						PrintXML.SetAttribute("PAGESIZE", "9999");
						PrintXML.SetAttribute("PAGENUMBER", "");
						PrintXML.SetAttribute("PAGESIZE", "9999");
						
						PrintXML.SelectTag(null, "REQUEST");			
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
						PrintXML.SetAttribute("DESTINATIONTYPE", sPrintType);
						
						PrintXML.SelectTag(null, "REQUEST");						
						PrintXML.CreateActiveTag("TEMPLATEDATA");
						
						var sStoredDataXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
						xmlSearchData = new top.frames[1].document.all.scXMLFunctions.XMLObject();
						xmlSearchData.LoadXML(sStoredDataXML);
						xmlSearchData.SelectTag(null,"REQUEST");
										
						PrintXML.SetAttribute("APPLICATIONNUMBER", xmlSearchData.GetTagText("APPLICATIONNUMBER"));
						PrintXML.SetAttribute("ACCOUNTNUMBER", xmlSearchData.GetTagText("ACCOUNTNUMBER"));
						PrintXML.SetAttribute("SEARCHUNITID", xmlSearchData.GetTagText("UNITID"));
						PrintXML.SetAttribute("SEARCHUNITNAME", xmlSearchData.GetTagText("UNITNAME"));
						PrintXML.SetAttribute("SEARCHUSERID", xmlSearchData.GetTagText("USERID"));
						PrintXML.SetAttribute("SEARCHUSERNAME", xmlSearchData.GetTagText("USERNAME"));
						PrintXML.SetAttribute("INCLUDEFLAG", xmlSearchData.GetTagText("INCCANCELLEDORDECLINEDAPPSCHECKED"));
						PrintXML.SetAttribute("DISTRIBUTIONCHANNELID", xmlSearchData.GetTagText("CHANNELID"));
						PrintXML.SetAttribute("DISTRIBUTIONCHANNELNAME", xmlSearchData.GetTagText("CHANNELNAME"));
						PrintXML.SetAttribute("DEPARTMENTID", xmlSearchData.GetTagText("DEPARTMENTID"));
						PrintXML.SetAttribute("DEPARTMENTNAME", xmlSearchData.GetTagText("DEPARTMENTNAME"));
						PrintXML.SetAttribute("LENDERNAME", xmlSearchData.GetTagText("LENDERNAME"));
						PrintXML.SetAttribute("PRODUCTNAME", xmlSearchData.GetTagText("MORTGAGEPRODUCTNAME"));
						PrintXML.SetAttribute("MORTGAGEPRODUCTCODE", xmlSearchData.GetTagText("MORTGAGEPRODUCTCODE"));
						PrintXML.SetAttribute("MORTGAGEPRODUCTSTARTDATE", xmlSearchData.GetTagText("MORTGAGEPRODUCTSTARTDATE"));
						PrintXML.SetAttribute("APPLICATIONTYPE", xmlSearchData.GetTagText("APPLICATIONTYPE"));
						PrintXML.SetAttribute("APPLICATIONSTAGEFROM", xmlSearchData.GetTagText("APPLICATIONSTAGEFROM"));
						PrintXML.SetAttribute("APPLICATIONSTAGEFROMNAME", xmlSearchData.GetTagText("APPLICATIONSTAGEFROMNAME"));
						PrintXML.SetAttribute("APPLICATIONSTAGETO", xmlSearchData.GetTagText("APPLICATIONSTAGETO"));
						PrintXML.SetAttribute("APPLICATIONSTAGETONAME", xmlSearchData.GetTagText("APPLICATIONSTAGETONAME"));
						PrintXML.SetAttribute("APPLICATIONDATEFROM", xmlSearchData.GetTagText("APPLICATIONDATEFROM"));
						PrintXML.SetAttribute("APPLICATIONDATETO", xmlSearchData.GetTagText("APPLICATIONDATETO"));
						PrintXML.SetAttribute("APPAPPROVEDMONTH", xmlSearchData.GetTagText("APPLICATIONAPPROVEDMONTH"));
						PrintXML.SetAttribute("APPAPPROVEDMONTHTEXT", xmlSearchData.GetTagText("APPLICATIONAPPROVEDMONTHTEXT"));
						PrintXML.SetAttribute("APPAPPROVEDYEAR", xmlSearchData.GetTagText("APPLICATIONAPPROVEDYEAR"));
						PrintXML.SetAttribute("SOBDIRECTORINDIRECT", xmlSearchData.GetTagText("INTERMEDIARYTYPE"));
						<% /* MO 02/10/2002 BMIDS00502 - Start
						PrintXML.SetAttribute("SOBCHANNELID", xmlSearchData.GetTagText("INTERMEDIARYPANELID"));
						PrintXML.SetAttribute("SOBNAME", xmlSearchData.GetTagText("INTERMEDIARYNAME"));
						      MO 02/10/2002 BMIDS00502 - End */ %>
						PrintXML.SetAttribute("SOBCHANNELID", xmlSearchData.GetTagText("INTRODUCERID"));
						PrintXML.SetAttribute("SOBNAME", xmlSearchData.GetTagText("INTRODUCERNAME"));
						
						PrintXML.SelectTag(null, "TEMPLATEDATA");
						PrintXML.CreateActiveTag("PRINTDETAIL");
						PrintXML.SetAttribute("DATE", Date());
						PrintXML.SetAttribute("PRINTUSERNAME", scScreenFunctions.GetContextParameter(window,"idUserName",""));
								
						PrintXML.SelectTag(null, "REQUEST");			
						PrintXML.CreateActiveTag("PRINTDATA");
						PrintXML.SetAttribute("METHODNAME", AttribsXML.GetTagAttribute("ATTRIBUTES", "PDMMETHOD"));	
						
						// 						PrintXML.RunASP(document, "PrintManager.asp");
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
}

<% /* Function converted from VBScript to JScript to improve performance */ %>
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
			alert("Rates are inconsistent on your current accepted or active quotation. Please remodel.");	
	}
}
<% /* BMIDS624 End */ %>

-->
</script>
</body>
</html>
