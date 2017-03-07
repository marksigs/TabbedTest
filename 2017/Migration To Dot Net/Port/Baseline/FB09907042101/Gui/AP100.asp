<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP100.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Employers reference
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		16/01/01	Created
JLD		18/01/01	Added some screen processing.
JLD		22/01/01	Altered creation of XML for create/update
JLD		23/01/01	Route to TM030, Don't set radio buttons if attribute not set, deal with Under Review.
JLD		25/01/01	change EMPLOYMENTSEQUENCENO  to EMPLOYMENTSEQUENCENUMBER in line with db
					change EMPLOYERSPRPSCHEME to EMPLOYEEPRPSCHEMEIND in line with db
JLD		14/02/01	Modifications to get it running with Task Management.
CL		05/03/01	SYS1920 Read only functionality added
GD		09/03/2001	sys2029 default combos to <SELECT>
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
SA		27/04/01	SYS2275 Removed Complete Task button and added Confirm Button
					Existing click code moved from Complete Task button to Confirm button
SA		03/12/01	SYS3285/3280 If Application set to "under Review" - don't make it read only. 
SG		10/12/01	SYS3364 Corrected xml attribute typo in PopulateFields()
PSC		12/12/01	SYS3388 Prompt before running confirm process
SG		01/02/02	SYS3837 Re-designed based on DC091
SR		16/02/02	SYS4101	Changes to logic for updating employers reference table
SR		16/02/02	SYS4101	Additional change to logic for updating employers reference table
SR		16/02/02	SYS4101	Remove error message for raised on initialise of employer reference.
AT		01/05/02	SYS4359 - Enable client customisation
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		15/05/2002	BMIDS00008	Amended IndustryType Combo to be Verification/Validation Type
								Removed References to PRP ,Modified UpdateEmpRef()
GD		24/06/02	BMIDS0077	Core Upgrade, plus fixes to unstable version from 7.0.2
MV		02/07/2002  BMIDS000109 Core Code Error - Modified the Tag name in PopulateFields()
MV		08/07/2002  BMIDS00109  Code Error 
MV		11/07/2002  BMIDS00109  Code Error - Modified PopulateFields();
MV		17/09/2002  BMIDS00468	Amended  UpdateEmpRef()
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MO		23/10/2002	BMIDS00649  Amended so that if you tab through the table you cant go off the end of the table
AW		24/10/2002	BMIDS00653	BM029 Call to Allowable Income
MDC		04/11/2002	BMIDS00654	BM088 Calculate Maximum Borrowing
MV		07/11/2002	BMIDS00833	AMended InitialiseTableXML()
MDC		16/11/2002	BMIDS00916  Amend CalcAllowableIncome to RunIncomeCalculation
MDC		21/11/2002	BMIDS01034  Return true from RunIncomeCalculation allowing user to progress
MO		21/11/2002	BMIDS01037 - Creating and Updating of references and casetasks result in 
					duplicate key violations
GHun	13/03/2003	BM0457		Include guarantors in allowable income calculation
HMA     16/09/2003  BM0063      Amend HTML text for radio buttons
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History
Prog	Date		AQR		Description 
MF		14/09/05	MAR30	Added new fields for Other income
MF		15/09/2005	MAR30	IncomeCalcs modified to send ActivityID into calculation
HMA     08/10/2005  MAR135  Do not update Customer Version when saving Tax Details.
PJO		23/11/2005  MAR477  Flag added to check for double click of SUBMIT button
PSC		25/01/2006	MAR1123	Use scroll table to allow extra items                           
PE		29/03/2006	MAR1525	Add Spinning Lifebelt
JD		25/05/2006  MAR1821 Add check for Underwriter, make sure IncomeSummary is loaded, exit back to TM030 on ok if no changes are made to screen.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History
Prog	Date		AQR		Description 
IK		10/08/2006	EP1039	fix confirm / ok conflict
PB		17/08/2006	EP10589 Merge MAR1908 - ensure task is set to pending on return to TM030
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* SG SYS3837 */ %>
<span style="TOP: 373px; LEFT: 305px; POSITION: absolute">
	<OBJECT data=scTableListScroll.asp id=scScrollTable name=scTable style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></OBJECT>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divDetails" style="HEIGHT: 513px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Employment Type
	<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
		<input id="txtCustEmployType" maxlength="50" style="WIDTH: 150px" class="msgReadOnly" readonly>
	</span>
</span>
<span style="LEFT: 310px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Employment Type
	<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
		<select id="cboEmployType" style="WIDTH: 150px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
	Verification /Validation
	<span style="LEFT: 125px; POSITION: absolute; TOP: -1px">
		<input id="txtCustIndustryType" maxlength="50" style="WIDTH: 150px" class="msgReadOnly" readonly>
	</span>
</span>
<span style="LEFT: 310px; POSITION: absolute; TOP: 30px" class="msgLabel">
	Verification /Validation
	<span style="LEFT: 125px; POSITION: absolute; TOP: -1px">
		<select id="cboIndustryType" style="WIDTH: 150px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 54px" class="msgLabel">
	Occupation Type
	<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
		<input id="txtCustOccupationType" maxlength="50" style="WIDTH: 150px" class="msgReadOnly" readonly>
	</span>
</span>
<span style="LEFT: 310px; POSITION: absolute; TOP: 54px" class="msgLabel">
	Occupation Type
	<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
		<select id="cboOccupationType" style="WIDTH: 150px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 76px" class="msgLabel">
	Job Title
	<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
		<input id="txtCustJobTitle" maxlength="50" style="WIDTH: 150px" class="msgReadOnly" readonly>
	</span>
</span>
<span style="LEFT: 310px; POSITION: absolute; TOP: 76px" class="msgLabel">
	Job Title
	<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
		<input id="txtJobTitle" maxlength="50" style="WIDTH: 150px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 98px" class="msgLabel">
	Date Started
	<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
		<input id="txtCustDateStarted" maxlength="10" style="WIDTH: 100px" class="msgReadOnly" readonly>
	</span>
</span>
<span style="LEFT: 310px; POSITION: absolute; TOP: 98px" class="msgLabel">
	Date Started
	<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
		<input id="txtDateStarted" maxlength="10" style="WIDTH: 100px" class="msgTxt">
	</span>
</span>
	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 137px" onclick="spnTable.onclick()">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="26%" class="TableHead">Income Type&nbsp;</td>
								<td width="16%" class="TableHead">Amount&nbsp;</td>
								<td width="16%" class="TableHead">Frequency&nbsp;</td>
								<td width="16%" class="TableHead">Confirmed Amount&nbsp;</td>								
								<td width="16%" class="TableHead">Confirmed Frequency&nbsp;</td></tr>
			<tr id="row01">		<td width="26%" class="TableTopLeft">&nbsp;</td>
								<td width="16%" class="TableTopCenter">	&nbsp;</td>
								<td width="16%" class="TableTopCenter">	&nbsp;</td>
								<td width="16%" class="TableTopCenter">	&nbsp;</td>
								<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="26%" class="TableLeft">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="26%" class="TableLeft">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="26%" class="TableLeft">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="26%" class="TableLeft">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="26%" class="TableLeft">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="26%" class="TableLeft">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="26%" class="TableLeft">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="26%" class="TableLeft">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td width="16%" class="TableCenter">&nbsp;</td>
								<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="26%" class="TableBottomLeft">&nbsp;</td>
								<td width="16%" class="TableBottomCenter">&nbsp;</td>
								<td width="16%" class="TableBottomCenter">&nbsp;</td>
								<td width="16%" class="TableBottomCenter">&nbsp;</td>
								<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>

	<div id="divBackground2" style="HEIGHT: 42px; LEFT: 4px; POSITION: absolute; TOP: 332px; WIDTH: 596px" class="msgGroupLight">
		<span id="spnToFrequency" tabindex="0"></span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 16px; WIDTH: 200px" class="msgLabel">
			Confirmed Amount
			<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtAmount" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
			</span> 
		</span>

		<span style="LEFT: 306px; POSITION: absolute; TOP: 16px; WIDTH: 200px" class="msgLabel">
			Confirmed Frequency
			<span style="LEFT: 125px; POSITION: absolute; TOP: -3px; WIDTH: 100px">
				<select id="cboFrequency" style="WIDTH: 100px" class="msgCombo">
				</select>
			</span> 
		</span>
		<span id="spnToAmount" tabindex="0"></span>
	</div>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 380px" class="msgLabel">
		Total Gross Income
		<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
			<input id="txtGrossIncome" name="GrossIncome" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgReadOnly" readonly tabindex=-1>
			</input>
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 407px" class="msgLabel">
		Include Other Income
		<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">			
			<span style="LEFT: 0px; POSITION: relative; TOP: -3px">
				<input id="optIncOtherIncomeYes" disabled name="OtherIncomeGroup" type="radio" value="1"><label for="optIncOtherIncomeYes" class="msgLabel">Yes</label> 
			</span>
			<span style="LEFT: 60px; POSITION: absolute;  TOP: -3px">
				<input id="optIncOtherIncomeNo" disabled name="OtherIncomeGroup" type="radio" value="0"><label for="optIncOtherIncomeNo" class="msgLabel">No</label> 
			</span> 	
		</span> 
	</span>
	
	<span style="LEFT: 310px; POSITION: absolute; TOP: 407px" class="msgLabel">
		Other Income Percentage
		<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">			
			<input id="txtOtherIncomePercentage" disabled maxlength="3" style="TOP: -3px; WIDTH: 100px; POSITION: absolute" class="msgTxt" NAME="txtOtherIncomePercentage">
			<label style="LEFT: 110px; POSITION: absolute; TOP: 0px" class="msgLabel">%</label>
		</span>
	</span>



<span style="LEFT: 4px; POSITION: absolute; TOP: 428px" class="msgLabel">
	Shares Held
	<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">
		<input id="txtCustSharesHeld" maxlength="3" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgReadOnly" readonly>
		<label style="LEFT: 110px; POSITION: absolute; TOP: 0px" class="msgLabel">%</label>
	</span>
</span>
<span style="LEFT: 310px; POSITION: absolute; TOP: 428px" class="msgLabel">
	Shares Held
	<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">
		<input id="txtSharesHeld" maxlength="3" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
		<label style="LEFT: 110px; POSITION: absolute; TOP: 0px" class="msgLabel">%</label>
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 451px" class="msgLabel">
	<LABEL id="idNINumber1"></LABEL>
	<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">
		<input id="txtCustNINumber" maxlength="9" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgReadOnly" readonly>
	</span>
</span>
<span style="LEFT: 310px; POSITION: absolute; TOP: 451px" class="msgLabel">
	<LABEL id="idNINumber2"></LABEL>
	<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">
		<input id="txtNINumber" maxlength="9" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 476px" class="msgLabel">
	Tax District
	<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">
		<input id="txtCustTaxDistrict" maxlength="60" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgReadOnly" readonly>
	</span>
</span>
<span style="LEFT: 310px; POSITION: absolute; TOP: 476px" class="msgLabel">
	Tax District
	<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">
		<input id="txtTaxDistrict" maxlength="60" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 497px" class="msgLabel">
	Tax Reference
	<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">
		<input id="txtCustTaxRef" maxlength="50" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgReadOnly" readonly>
	</span>
</span>
<span style="LEFT: 310px; POSITION: absolute; TOP: 497px" class="msgLabel">
	Tax Reference
	<span style="LEFT: 125px; POSITION: absolute; TOP: 0px">
		<input id="txtTaxRef" maxlength="50" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
	</span>
</span>
</div>
<div id="divQuestions" style="HEIGHT: 100px; LEFT: 10px; POSITION: absolute; TOP: 576px; WIDTH: 604px" class="msgGroup">
<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
	Are non-guaranteed payments<BR>likely to continue?
	<span style="LEFT: 150px; POSITION: absolute; TOP: 6px">
		<input id="optNonGuarYes" name="RadioGroup1" type="radio" value="1"><label for="optNonGuarYes" class="msgLabel">Yes</label>
	</span>
	<span style="LEFT: 200px; POSITION: absolute; TOP: 6px">
		<input id="optNonGuarNo" name="RadioGroup1" type="radio" value="0"><label for="optNonGuarNo" class="msgLabel">No</label>
	</span>
</span>
<span style="LEFT: 310px; POSITION: absolute; TOP: 19px" class="msgLabel">
	Signed and Stamped?
	<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
		<input id="optSignedYes" name="RadioGroup2" type="radio" value="1"><label for="optSignedYes" class="msgLabel">Yes</label>
	</span>
	<span style="LEFT: 200px; POSITION: absolute; TOP: -3px">
		<input id="optSignedNo" name="RadioGroup2" type="radio" value="0"><label for="optSignedNo" class="msgLabel">No</label>
	</span>
</span>
<% /*  BMIDS Specific - Removed PRP References
<span id="idPRPScheme" style="TOP:50px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Is the Employee part<BR>of a PRP scheme?
	<span style="TOP:6px; LEFT:150px; POSITION:ABSOLUTE">
		<input id="optPRPYes" name="RadioGroup3" type="radio" value="1">
		<label for="optPRPYes" class="msgLabel">Yes</label>
	</span>
	<span style="TOP:6px; LEFT:200px; POSITION:ABSOLUTE">
		<input id="optPRPNo" name="RadioGroup3" type="radio" value="0">
		<label for="optPRPNo" class="msgLabel">No</label>
	</span>
</span>
<span style="TOP:53px; LEFT:310px; POSITION:ABSOLUTE">
	<input id="btnPRP" value="PRP" type="button" style="WIDTH:60px" class="msgButton">
</span>*/ %>
<% /* SYS2275 SA Complete Task button replaced with confirm */ %>
<% /* <span style="TOP:53px; LEFT:420px; POSITION:ABSOLUTE"> */ %>
	<% /* <input id="btnComplete" value="Complete Task" type="button" style="WIDTH:100px" class="msgButton"> */ %>
<% /* </span> */ %>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 640px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP100attribs.asp" --><!-- #include FILE="Customise/AP100Customise.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_sReadOnly = null;
var m_sPRPXML = "";
var m_sTaskXML = "";
var m_TaskXML = null;
var m_CurrEmpXML = null;
var scScreenFunctions;

var m_bCreateFirstEmployersRef = true;
var m_blnReadOnly = false;
var m_sAppFactFindNo = "";
var m_sApplicationNumber = "";
var XML = null;
var TableXML = null;
var m_nIncomeTypeCount = 0;
<% /* MO - 21/11/2002 - BMIDS01037 */ %>
var m_bSetTaskToPending = null;
var m_bSubmitClicked = false ;	<% /* Added by PJO 23/11/2005 MAR477 */ %>
var m_nTableLength = 10;				<% /* PSC 25/01/2006 MAR1123 */ %>


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Cancel", "Confirm");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Employer's Reference","AP100",scScreenFunctions);

	scScreenFunctions.SetCurrency(window,frmScreen);

	m_bSubmitClicked = false ;	<% /* Added by PJO 23/11/2005 MAR477 */ %>
	RetrieveContextData();
	Customise();
	SetMasks();
	
	PopulateScreen();

	scScreenFunctions.SetFocusToFirstField(frmScreen);	

	//SG SYS3837
	spnTable.onclick()			

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP100");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
	
	<% /* MF MAR30 move to end to take snapshot of fields after populating them */ %>
	Validation_Init();	
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
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	//SYS2275 SA 1/5/01
	m_sAppFactFindNo = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
// DEBUG
//scScreenFunctions.SetContextParameter(window, "idCustomerNumber1", "00068012");
//scScreenFunctions.SetContextParameter(window, "idCustomerVersionNumber1", "1");
//m_sTaskXML = "<CASETASK SOURCEAPPLICATION=\"Omiga\" CASEID=\"C00071021\" ACTIVITYID=\"10\" TASKID=\"Second_Charge\" CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"1\" TASKSTATUS=\"20\" TASKINSTANCE=\"1\" STAGEID=\"60\" CASESTAGESEQUENCENO=\"6\"/>";
}
function SetFieldsReadOnly()
{
	scScreenFunctions.SetFieldState(frmScreen, "txtCustEmployType",		"R");
	scScreenFunctions.SetFieldState(frmScreen, "txtCustIndustryType",	"R");
	scScreenFunctions.SetFieldState(frmScreen, "txtCustOccupationType", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtCustJobTitle",		"R");
	scScreenFunctions.SetFieldState(frmScreen, "txtCustDateStarted",	"R");
	scScreenFunctions.SetFieldState(frmScreen, "txtCustSharesHeld",		"R");
	scScreenFunctions.SetFieldState(frmScreen, "txtCustNINumber",		"R");
	scScreenFunctions.SetFieldState(frmScreen, "txtCustTaxDistrict",	"R");
	scScreenFunctions.SetFieldState(frmScreen, "txtCustTaxRef",			"R");
	//frmScreen.btnPRP.disabled = true;
	
	<% /* JD MAR1821 if user is underwriter enable fields */ %>
	var bExcludeOthIncUnlessUnderwriter=false;
	var GlobalParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	bExcludeOthIncUnlessUnderwriter = GlobalParamXML.GetGlobalParameterBoolean(document,"ExcludeOthIncUnlessUnderwriter");
	if(bExcludeOthIncUnlessUnderwriter)
	{
		var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole");
		var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if (TempXML.IsInComboValidationList(document,"UserRole",sUserRole,["UW"]))
		{
			//enable fields
			frmScreen.txtOtherIncomePercentage.disabled = false;
			frmScreen.optIncOtherIncomeNo.disabled = false;
			frmScreen.optIncOtherIncomeYes.disabled = false;
		}
	}	
}

function LoadIncomeSummary()
{

	var EmploymentXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	EmploymentXML.CreateRequestTag(window,null);
	var tagCustomerList = EmploymentXML.CreateActiveTag("CUSTOMERLIST");
	EmploymentXML.CreateActiveTag("CUSTOMER");
	EmploymentXML.CreateTag("CUSTOMERNUMBER", m_TaskXML.GetAttribute("CUSTOMERIDENTIFIER"));
	EmploymentXML.CreateTag("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_TaskXML.GetAttribute("CUSTOMERIDENTIFIER")));
	
	EmploymentXML.RunASP(document,"FindEmploymentList.asp");	

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = EmploymentXML.CheckResponse(ErrorTypes);
	var nListLength = 0;	
	if ((ErrorReturn[1] == ErrorTypes[0]) | (EmploymentXML.XMLDocument.text == ""))
	{
		<% /* Error: record not found */ %>
		m_bIncomeSummaryExists=false;	
	} 
	else if (ErrorReturn[0] == true)
	{
		<% /* No error */ %>
		m_bIncomeSummaryExists=true;		
		if (EmploymentXML.SelectTag(null,"INCOMESUMMARYDETAILS") != null){		
			if (EmploymentXML.ActiveTag.text != "")
			{
				scScreenFunctions.SetRadioGroupValue(frmScreen,"OtherIncomeGroup",EmploymentXML.GetTagText("UNDERWRITEROVERRIDEINCLUDEOTHERINC"));
				if(frmScreen.optIncOtherIncomeYes.checked){
					frmScreen.optIncOtherIncomeYes.click();
				}
				frmScreen.txtOtherIncomePercentage.value=EmploymentXML.GetTagText("UNDERWRITEROTHERINCOMEPERCENTAGE");	
			}
		}
	}
}


function PopulateFields()
{
	if(m_TaskXML != null)
	{
		m_TaskXML.SelectTag(null, "CASETASK");

		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
		XML.CreateRequestTag(window , "GetEmpDetailsForRef");
		<% /* MV - 02/07/2002 - BMIDS000109 - Core Code Error */ %>
		XML.CreateActiveTag("EMPDETAILSFORREF");
		XML.SetAttribute("CUSTOMERNUMBER", m_TaskXML.GetAttribute("CUSTOMERIDENTIFIER"));
		XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_TaskXML.GetAttribute("CUSTOMERIDENTIFIER")));
		XML.SetAttribute("EMPLOYMENTSEQUENCENUMBER",m_TaskXML.GetAttribute("CONTEXT"));
		XML.SetAttribute("_COMBOLOOKUP_","y");

		//Populate the applicant-entered employment fields
		XML.RunASP(document, "omAppProc.asp");
		if(XML.IsResponseOK())
		{
			XML.SelectTag(null, "EMPLOYMENTDETAIL");
			frmScreen.txtCustEmployType.value = XML.GetAttribute("EMPLOYMENTTYPE_TEXT");
			frmScreen.txtCustIndustryType.value = XML.GetAttribute("INDUSTRYTYPE_TEXT");
			frmScreen.txtCustOccupationType.value = XML.GetAttribute("OCCUPATIONTYPE_TEXT");
			frmScreen.txtCustJobTitle.value = XML.GetAttribute("JOBTITLE");
			frmScreen.txtCustDateStarted.value = XML.GetAttribute("DATESTARTEDORESTABLISHED");
			frmScreen.txtCustSharesHeld.value = XML.GetAttribute("PERCENTSHARESHELD");
			frmScreen.txtCustNINumber.value = XML.GetAttribute("NATIONALINSURANCENUMBER");
			frmScreen.txtCustTaxDistrict.value = XML.GetAttribute("TAXOFFICE");
			frmScreen.txtCustTaxRef.value = XML.GetAttribute("TAXREFERENCENUMBER");

			var nRecordCount = 0;					
			nRecordCount = PopulateTableXML();
		}
<%		//Populate the Employers reference info if some already stored. The task status 
		//may be 'Pending' but still have no reference details input yet if the task
		//has an OutputDocument associated with it. We need to cater for this.
%>
		m_CurrEmpXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_CurrEmpXML.CreateRequestTag(window , "GetCurrEmployersRef");
		m_CurrEmpXML.CreateActiveTag("CURRENTEMPLOYERSREF");
		m_CurrEmpXML.SetAttribute("CUSTOMERNUMBER", m_TaskXML.GetAttribute("CUSTOMERIDENTIFIER"));
		m_CurrEmpXML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(m_TaskXML.GetAttribute("CUSTOMERIDENTIFIER")));
		m_CurrEmpXML.SetAttribute("EMPLOYMENTSEQUENCENUMBER",m_TaskXML.GetAttribute("CONTEXT"));

		m_CurrEmpXML.RunASP(document, "omAppProc.asp");  //ReferencesBO
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = m_CurrEmpXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{
			<% /* MV - 11/07/2002 - BMIDS00109 - Code Error */ %>
			<% /* MO - 21/11/2002 - BMIDS01037 - REMOVED */ %>
			<% /* if(m_TaskXML.GetAttribute("OUTPUTDOCUMENT") == "")
			{
				alert("Employers reference cannot be found");
				m_bCreateFirstEmployersRef = false;
			}
			else
			{
				m_bCreateFirstEmployersRef = true;
			} */ %>
			m_bCreateFirstEmployersRef = true;
		}
		else if(ErrorReturn[0] == true)
		{
			m_bCreateFirstEmployersRef = false;
			m_CurrEmpXML.SelectTag(null, "CURRENTEMPLOYERSREF")
			frmScreen.cboEmployType.value = m_CurrEmpXML.GetAttribute("EMPLOYMENTTYPE");
			frmScreen.cboIndustryType.value = m_CurrEmpXML.GetAttribute("INDUSTRYTYPE");
			frmScreen.cboOccupationType.value = m_CurrEmpXML.GetAttribute("OCCUPATIONTYPE");
			frmScreen.txtJobTitle.value = m_CurrEmpXML.GetAttribute("JOBTITLE");
			frmScreen.txtDateStarted.value = m_CurrEmpXML.GetAttribute("DATESTARTED");

			//At this point we need to update the table and tableXML with ConfirmedEarnedIncome information 			
			UpdateTableXMLConfirmed();
			m_CurrEmpXML.SelectTag(null, "CURRENTEMPLOYERSREF")
			CalculateTotals();
			
			frmScreen.txtSharesHeld.value = m_CurrEmpXML.GetAttribute("PERCENTSHARESHELD");
			<% /* JD MAR1821 load income summary outside of loop */ %>
			
			frmScreen.txtNINumber.value = m_CurrEmpXML.GetAttribute("NATIONALINSURANCENUMBER");
			frmScreen.txtTaxDistrict.value = m_CurrEmpXML.GetAttribute("TAXOFFICE");
			frmScreen.txtTaxRef.value = m_CurrEmpXML.GetAttribute("TAXREFERENCENUMBER");
			if(m_CurrEmpXML.GetAttribute("NONGUARANTEEDPAYMENTSIND") == "1")
				frmScreen.optNonGuarYes.checked = true;
			else if(m_CurrEmpXML.GetAttribute("NONGUARANTEEDPAYMENTSIND") == "0")
				frmScreen.optNonGuarNo.checked = true;
			if(m_CurrEmpXML.GetAttribute("SIGNEDSTAMPIND") == "1")
				frmScreen.optSignedYes.checked = true;
			else if(m_CurrEmpXML.GetAttribute("SIGNEDSTAMPIND") == "0")
				frmScreen.optSignedNo.checked = true;
			if(m_CurrEmpXML.GetAttribute("EMPLOYEEPRPSCHEMEIND") == "1")
			{
				//frmScreen.optPRPYes.checked = true;
				//frmScreen.btnPRP.disabled = false;
			}
			else if(m_CurrEmpXML.GetAttribute("EMPLOYEEPRPSCHEMEIND") == "0")
				//frmScreen.optPRPNo.checked = true;
			//set up the PRP info for AP101
			m_sPRPXML = "<PRP IRAPPROVAL=\"" + m_CurrEmpXML.GetAttribute("IRAPPROVALIND") + "\"";
			m_sPRPXML += " PAYMENTSRECLAIM=\"" + m_CurrEmpXML.GetAttribute("INTERIMPAYMENTSRECLAIMIND") + "\"";
			m_sPRPXML += " SALARYRESTORATION=\"" + m_CurrEmpXML.GetAttribute("SALARYRESTORATIONIND") + "\"";
			m_sPRPXML += " TAXLIABILITY=\"" + m_CurrEmpXML.GetAttribute("EMPLOYERTAXLIABILITYIND") + "\"";
			m_sPRPXML += " BASICPRP=\"" + m_CurrEmpXML.GetAttribute("BASICPAYINCPRP") + "\"";
			m_sPRPXML += " PRPFREQUENCY=\"" + m_CurrEmpXML.GetAttribute("FREQUENCYPRPRECEIVED") + "\"";
			m_sPRPXML += " ADDRESSAGREE=\"" + m_CurrEmpXML.GetAttribute("ADDRESSAGREEMENTIND") + "\"/>";
		}
		<% /* MF MAR30 populate new fields */ %>			
		LoadIncomeSummary();
			
		<% /* PSC 25/01/2006 MAR1123 - Start */ %>
		scTable.initialiseTable(tblTable,0,"",PopulateTable,m_nTableLength,m_nIncomeTypeCount);
		PopulateTable(0);

		if (m_nIncomeTypeCount > nRecordCount)
			<%/* Not all income types had records on the database for this employment, so force a 
			save of the default values for the missing income types */%>
			FlagChange(true);

		if (m_nIncomeTypeCount > 0)
		{
			scTable.setRowSelected(1);
			spnTable.onclick();
		}
		<% /* PSC 25/01/2006 MAR1123 - End */ %>
		
		if (m_TaskXML.GetAttribute("TASKSTATUS") != "10") {
			m_bSetTaskToPending = false;
		} else {
			m_bSetTaskToPending = true;
		}
		
		if(m_TaskXML.GetAttribute("TASKSTATUS") == "40")
		{
			//Task is completed so set screen to readonly
			m_sReadOnly = "1";
		}
	}
}
function PopulateScreen()
{
	if(m_sTaskXML != "")
	{
		m_TaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_TaskXML.LoadXML(m_sTaskXML);
	}
	PopulateCombos();

	//SG SYS3837
	InitialiseTableXML();
	DefaultFields();
	
	PopulateFields();
	SetFieldsReadOnly();
	if(m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);

		DisableMainButton("Confirm");
		DisableMainButton("Submit");
	}
}
function PopulateCombos()
{
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	var blnSuccess = true;
	
	var sGroupList = new Array("EmploymentType","IndustryType","OccupationType", "EmploymentPaymentFreq");

	if(XML.GetComboLists(document,sGroupList))
	{
		XML.PopulateCombo(document,frmScreen.cboEmployType,		"EmploymentType",		true);
		XML.PopulateCombo(document,frmScreen.cboIndustryType,	"IndustryType",			true);
		XML.PopulateCombo(document,frmScreen.cboOccupationType,	"OccupationType",		true);
		XML.PopulateCombo(document,frmScreen.cboFrequency,		"EmploymentPaymentFreq",true);
		
	}

	TableXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	sGroupList = new Array("IncomeType");
	if(TableXML.GetComboLists(document,sGroupList))
		blnSuccess = blnSuccess & (TableXML != null)

	if(!blnSuccess) scScreenFunctions.SetScreenToReadOnly(frmScreen);
	XML = null;
}
function GetCustomerVersion(sCustomerNumber)
{
	// find the customerversion number in context for this customernumber
	for(iCount = 1; iCount < 6; iCount++)
	{
		if(scScreenFunctions.GetContextParameter(window, "idCustomerNumber" + iCount.toString(), "") == sCustomerNumber)
			return(scScreenFunctions.GetContextParameter(window, "idCustomerVersionNumber" + iCount.toString(), ""));
	}
	alert("CustomerVersionNumber not found for customer " + sCustomerNumber);
}
function CalculatePayFields(XML)
{
	var nGrossBasic = 0;
	var nGtdOvertime = 0;
	var nRegOvertime = 0;
	XML.CreateTagList("EARNEDINCOME");
	for(var iCount = 0; iCount < XML.ActiveTagList.length; iCount++)
	{
		XML.SelectTagListItem(iCount);
		switch(XML.GetAttribute("EARNEDINCOMETYPE"))
		{
			case "1":
			case "6":
				nGrossBasic += parseFloat(XML.GetAttribute("EARNEDINCOMEAMOUNT")) * parseFloat(XML.GetAttribute("PAYMENTFREQUENCYTYPE"));
			break;
			case "2":
			case "4":
				nGtdOvertime += parseFloat(XML.GetAttribute("EARNEDINCOMEAMOUNT")) * parseFloat(XML.GetAttribute("PAYMENTFREQUENCYTYPE"));
			break;
			case "3":
			case "10":
			case "9":
			case "8":
				nRegOvertime += parseFloat(XML.GetAttribute("EARNEDINCOMEAMOUNT")) * parseFloat(XML.GetAttribute("PAYMENTFREQUENCYTYPE"));
			break;
		}
	}
}
<%

/*
function UpdateEmpRef()
{
	var XML = null;
	var sASPFile = "";
	
	if(m_bCreateFirstEmployersRef == true)
	{
		//CREATE employers ref
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var reqTag = XML.CreateRequestTag(window , "CreateCurrEmployersRef");
		XML.CreateActiveTag("CASETASK");
		XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		XML.SetAttribute("ACTIVITYID", m_TaskXML.GetAttribute("ACTIVITYID"));
		XML.SetAttribute("STAGEID", m_TaskXML.GetAttribute("STAGEID"));
		XML.SetAttribute("CASESTAGESEQUENCENO", m_TaskXML.GetAttribute("CASESTAGESEQUENCENO"));
		XML.SetAttribute("CASEID", m_TaskXML.GetAttribute("CASEID"));
		XML.SetAttribute("TASKID", m_TaskXML.GetAttribute("TASKID"));
		XML.SetAttribute("TASKINSTANCE", m_TaskXML.GetAttribute("TASKINSTANCE"));
		XML.ActiveTag = reqTag;
		XML.CreateActiveTag("CURRENTEMPLOYERSREF");
		sASPFile = "omAppProc.asp";
	
	}
	else
	{
		//UPDATE the employers ref
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window , "UpdateCurrEmployersRef");
		m_CurrEmpXML.SelectTag(null, "RESPONSE");
		XML.AddXMLBlock(m_CurrEmpXML.CreateFragment());
		XML.SelectTag(null, "CURRENTEMPLOYERSREF");		
		sASPFile = "omAppProc.asp";
	}
	var sCustomerNumber = m_TaskXML.GetAttribute("CUSTOMERIDENTIFIER");
	XML.SetAttribute("CUSTOMERNUMBER", sCustomerNumber);
	XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(sCustomerNumber));
	XML.SetAttribute("EMPLOYMENTSEQUENCENUMBER",m_TaskXML.GetAttribute("CONTEXT"));
	XML.SetAttribute("EMPLOYMENTTYPE",frmScreen.cboEmployType.value);
	XML.SetAttribute("INDUSTRYTYPE",frmScreen.cboIndustryType.value);
	XML.SetAttribute("OCCUPATIONTYPE",frmScreen.cboOccupationType.value);
	XML.SetAttribute("JOBTITLE",frmScreen.txtJobTitle.value);
	XML.SetAttribute("DATESTARTED",frmScreen.txtDateStarted.value);	
	XML.SetAttribute("PERCENTSHARESHELD",frmScreen.txtSharesHeld.value);
	XML.SetAttribute("NATIONALINSURANCENUMBER",frmScreen.txtNINumber.value);
	XML.SetAttribute("TAXOFFICE",frmScreen.txtTaxDistrict.value);
	XML.SetAttribute("TAXREFERENCENUMBER",frmScreen.txtTaxRef.value);
	if(frmScreen.optNonGuarYes.checked == true)
		XML.SetAttribute("NONGUARANTEEDPAYMENTSIND","1");
	else if(frmScreen.optNonGuarNo.checked == true)
		XML.SetAttribute("NONGUARANTEEDPAYMENTSIND","0");
	if(frmScreen.optSignedYes.checked == true)
		XML.SetAttribute("SIGNEDSTAMPIND","1");
	else if(frmScreen.optSignedNo.checked == true)
		XML.SetAttribute("SIGNEDSTAMPIND","0");
	if(frmScreen.optPRPYes.checked == true)
		XML.SetAttribute("EMPLOYEEPRPSCHEMEIND","1");
	else if(frmScreen.optPRPNo.checked == true)
		XML.SetAttribute("EMPLOYEEPRPSCHEMEIND","0");
	//Get the values from the PRP screen also
	if(m_sPRPXML.length > 0)
	{
		var prpXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		prpXML.LoadXML(m_sPRPXML);
		prpXML.SelectTag(null,"PRP");
		XML.SetAttribute("IRAPPROVALIND", prpXML.GetAttribute("IRAPPROVAL"));
		XML.SetAttribute("INTERIMPAYMENTSRECLAIMIND", prpXML.GetAttribute("PAYMENTSRECLAIM"));
		XML.SetAttribute("SALARYRESTORATIONIND", prpXML.GetAttribute("SALARYRESTORATION"));
		XML.SetAttribute("EMPLOYERTAXLIABILITYIND", prpXML.GetAttribute("TAXLIABILITY"));
		XML.SetAttribute("BASICPAYINCPRP", prpXML.GetAttribute("BASICPRP"));
		XML.SetAttribute("FREQUENCYPRPRECEIVED", prpXML.GetAttribute("PRPFREQUENCY"));
		XML.SetAttribute("ADDRESSAGREEMENTIND", prpXML.GetAttribute("ADDRESSAGREE"));
		prpXML = null;
	}
	else
	{
		XML.SetAttribute("IRAPPROVALIND", "");
		XML.SetAttribute("INTERIMPAYMENTSRECLAIMIND", "");
		XML.SetAttribute("SALARYRESTORATIONIND", "");
		XML.SetAttribute("EMPLOYERTAXLIABILITYIND", "");
		XML.SetAttribute("BASICPAYINCPRP", "");
		XML.SetAttribute("FREQUENCYPRPRECEIVED", "");
		XML.SetAttribute("ADDRESSAGREEMENTIND", "");
	}
	
	//Populate XML with parts of TableXML
	TableXML.ActiveTag = null;
	TableXML.CreateTagList("LISTENTRY");
	var iCount;
	for (iCount = 0; iCount < m_nIncomeTypeCount; iCount++)
	{
		TableXML.SelectTagListItem(iCount);
		XML.SelectTag(null,"CURRENTEMPLOYERSREF");
		
		XML.CreateActiveTag("CONFIRMEDEARNEDINCOME");
		XML.SetAttribute("EARNEDINCOMESEQUENCENUMBER",iCount+1);
		XML.SetAttribute("EARNEDINCOMEAMOUNT",TableXML.GetTagText("CONFIRMEDINCOMEAMOUNT"));
		XML.SetAttribute("EARNEDINCOMETYPE",TableXML.GetTagText("VALUEID"));
		XML.SetAttribute("PAYMENTFREQUENCYTYPE",TableXML.GetTagText("CONFIRMEDPAYMENTFREQUENCYTYPE"));
	}
		
	XML.RunASP(document, sASPFile);
	if(XML.IsResponseOK())
	{
		return true;
	}
	return false;
}
*/ %>

var m_bIncomeSummaryExists=false;
<% /* MAR30 Save the other income details using the new Other Income Percentage field value */ %>
function SaveOtherIncomeDetails()
{
	var bSuccess = true;
	var sAction = "";
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	<% /* XML Format
		<REQUEST ... ACTION=...>
			<INCOMESUMMARY>
				Income Summary fields in here
			</INCOMESUMMARY>
			<CUSTOMERVERSION>
				<NATIONALINSURANCENUMBER>.....</NATIONALINSURANCENUMBER>
			</CUSTOMERVERSION>
		</REQUEST>
	*/ %>	
	if(m_bIncomeSummaryExists)
		sAction = "UPDATE";
	else
		sAction = "CREATE";
		
	XML.CreateRequestTag(window,sAction);
	XML.SetAttribute("UPDATECUSTOMERVERSION", "False");     // MAR135
	XML.CreateActiveTag("INCOMESUMMARY");
	var sCustomerNumber = m_TaskXML.GetAttribute("CUSTOMERIDENTIFIER");
	XML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);	
	XML.CreateTag("CUSTOMERVERSIONNUMBER", GetCustomerVersion(sCustomerNumber));	
	XML.CreateTag("UNDERWRITEROVERRIDEINCLUDEOTHERINC", scScreenFunctions.GetRadioGroupValue(frmScreen,"OtherIncomeGroup"));
	XML.CreateTag("UNDERWRITEROTHERINCOMEPERCENTAGE", frmScreen.txtOtherIncomePercentage.value);	
	<% /* Save the details */ %>
	XML.RunASP(document,"SaveTaxDetails.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK		
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	bSuccess = XML.IsResponseOK();
	XML = null;
	return(bSuccess);
}



//GD BMIDS0077 Core Upgrade 24/06/02
function UpdateEmpRef()
{
	var XML = null;
	var sASPFile = "";
	if(m_bCreateFirstEmployersRef == true)
	{
		<% /* Create employers ref */ %>
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var reqTag = XML.CreateRequestTag(window , "CreateCurrEmployersRef");
		XML.CreateActiveTag("CASETASK");
		XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		XML.SetAttribute("ACTIVITYID", m_TaskXML.GetAttribute("ACTIVITYID"));
		XML.SetAttribute("STAGEID", m_TaskXML.GetAttribute("STAGEID"));
		XML.SetAttribute("CASESTAGESEQUENCENO", m_TaskXML.GetAttribute("CASESTAGESEQUENCENO"));
		XML.SetAttribute("CASEID", m_TaskXML.GetAttribute("CASEID"));
		XML.SetAttribute("TASKID", m_TaskXML.GetAttribute("TASKID"));
		XML.SetAttribute("TASKINSTANCE", m_TaskXML.GetAttribute("TASKINSTANCE"));
		XML.ActiveTag = reqTag;
		XML.CreateActiveTag("CURRENTEMPLOYERSREF");
		sASPFile = "omAppProc.asp";
	}
	else
	{
		<% /* UPDATE the employers ref */ %>
		XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window , "UpdateCurrEmployersRef");
		m_CurrEmpXML.SelectTag(null, "RESPONSE");
		XML.AddXMLBlock(m_CurrEmpXML.CreateFragment());
		XML.SelectTag(null, "CURRENTEMPLOYERSREF");
		sASPFile = "omAppProc.asp";
	}
	var sCustomerNumber = m_TaskXML.GetAttribute("CUSTOMERIDENTIFIER");
	XML.SetAttribute("CUSTOMERNUMBER", sCustomerNumber);
	XML.SetAttribute("CUSTOMERVERSIONNUMBER", GetCustomerVersion(sCustomerNumber));
	XML.SetAttribute("EMPLOYMENTSEQUENCENUMBER",m_TaskXML.GetAttribute("CONTEXT"));
	XML.SetAttribute("EMPLOYMENTTYPE",frmScreen.cboEmployType.value);
	XML.SetAttribute("INDUSTRYTYPE",frmScreen.cboIndustryType.value);
	XML.SetAttribute("OCCUPATIONTYPE",frmScreen.cboOccupationType.value);
	XML.SetAttribute("JOBTITLE",frmScreen.txtJobTitle.value);
	XML.SetAttribute("DATESTARTED",frmScreen.txtDateStarted.value);
	//XML.SetAttribute("GROSSBASICINCOME",frmScreen.txtGrossBasic.value);
	//GD BMIDS0077 XML.SetAttribute("GUARANTEEDOVERTIMEBONUS",frmScreen.txtGtd.value);
	//XML.SetAttribute("REGOVERTIMEPRPBONUS",frmScreen.txtReg.value);
	XML.SetAttribute("PERCENTSHARESHELD",frmScreen.txtSharesHeld.value);
	XML.SetAttribute("NATIONALINSURANCENUMBER",frmScreen.txtNINumber.value);
	XML.SetAttribute("TAXOFFICE",frmScreen.txtTaxDistrict.value);
	XML.SetAttribute("TAXREFERENCENUMBER",frmScreen.txtTaxRef.value);
	if(frmScreen.optNonGuarYes.checked == true)
		XML.SetAttribute("NONGUARANTEEDPAYMENTSIND","1");
	else if(frmScreen.optNonGuarNo.checked == true)
		XML.SetAttribute("NONGUARANTEEDPAYMENTSIND","0");
	if(frmScreen.optSignedYes.checked == true)
		XML.SetAttribute("SIGNEDSTAMPIND","1");
	else if(frmScreen.optSignedNo.checked == true)
		XML.SetAttribute("SIGNEDSTAMPIND","0");
	
	<% /* Since PRP functionality is removed update with nothing */ %>
	XML.SetAttribute("EMPLOYEEPRPSCHEMEIND","");
	XML.SetAttribute("IRAPPROVALIND", "");
	XML.SetAttribute("INTERIMPAYMENTSRECLAIMIND", "");
	XML.SetAttribute("SALARYRESTORATIONIND", "");
	XML.SetAttribute("EMPLOYERTAXLIABILITYIND", "");
	XML.SetAttribute("BASICPAYINCPRP", "");
	XML.SetAttribute("FREQUENCYPRPRECEIVED", "");
	XML.SetAttribute("ADDRESSAGREEMENTIND", "");
	
	//Populate XML with parts of TableXML
	TableXML.ActiveTag = null;
	TableXML.CreateTagList("LISTENTRY");
	var iCount;
	for (iCount = 0; iCount < m_nIncomeTypeCount; iCount++)
	{
		TableXML.SelectTagListItem(iCount);
		XML.SelectTag(null,"CURRENTEMPLOYERSREF");
		
		XML.CreateActiveTag("CONFIRMEDEARNEDINCOME");
		XML.SetAttribute("EARNEDINCOMESEQUENCENUMBER",iCount+1);
		XML.SetAttribute("EARNEDINCOMEAMOUNT",TableXML.GetTagText("CONFIRMEDINCOMEAMOUNT"));
		XML.SetAttribute("EARNEDINCOMETYPE",TableXML.GetTagText("VALUEID"));
		XML.SetAttribute("PAYMENTFREQUENCYTYPE",TableXML.GetTagText("CONFIRMEDPAYMENTFREQUENCYTYPE"));
	}
	
	// 	XML.RunASP(document, sASPFile);
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, sASPFile);
			break;
		default: // Error
			XML.SetErrorResponse();
		}
	
	<% /* MO - 21/11/2002 - BMIDS01037 - If everything worked ok and we have so far updated the employers ref but
	still need to set the task to pending, do it now */ %>
	if(XML.IsResponseOK()) {
		<% /* MF MAR30 Save other income radio and percentage */ %>
		if(SaveOtherIncomeDetails()){
			//PB 17/08/2006 EP1089 Begin
			//if ((m_bCreateFirstEmployersRef == false) && (m_bSetTaskToPending == true)) {
			if (m_bSetTaskToPending == true)
			{				
			//EP1089 End
				SetToPendingXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				SetToPendingXML.CreateRequestTag(window , "UpdateCaseTask");
				SetToPendingXML.CreateActiveTag("CASETASK");
				SetToPendingXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
				SetToPendingXML.SetAttribute("ACTIVITYID", m_TaskXML.GetAttribute("ACTIVITYID"));
				SetToPendingXML.SetAttribute("STAGEID", m_TaskXML.GetAttribute("STAGEID"));
				SetToPendingXML.SetAttribute("CASESTAGESEQUENCENO", m_TaskXML.GetAttribute("CASESTAGESEQUENCENO"));
				SetToPendingXML.SetAttribute("CASEID", m_TaskXML.GetAttribute("CASEID"));
				SetToPendingXML.SetAttribute("TASKID", m_TaskXML.GetAttribute("TASKID"));
				SetToPendingXML.SetAttribute("TASKINSTANCE", m_TaskXML.GetAttribute("TASKINSTANCE"));
				SetToPendingXML.SetAttribute("TASKSTATUS", "20"); <% /* Pending */ %>
				
				SetToPendingXML.RunASP(document, "msgTMBO.asp") ;
				
				if (SetToPendingXML.IsResponseOK()) {
					return true;
				} else {
					return false;
				}
				
			} else {
				return true;
			}
			return true;
		} else {
			return false;
		}
	} else {
		return false;
	}
}
function ValidateEmpRef()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = XML.CreateRequestTag(window , "ValidateCurrEmployersRef");
	XML.CreateActiveTag("CASETASK");
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("ACTIVITYID",					m_TaskXML.GetAttribute("ACTIVITYID"));
	XML.SetAttribute("CASEID",						m_TaskXML.GetAttribute("CASEID"));
	XML.SetAttribute("TASKID",						m_TaskXML.GetAttribute("TASKID"));
	XML.SetAttribute("STAGEID",						m_TaskXML.GetAttribute("STAGEID"));
	XML.SetAttribute("CASESTAGESEQUENCENO",			m_TaskXML.GetAttribute("CASESTAGESEQUENCENO"));
	XML.SetAttribute("TASKINSTANCE",				m_TaskXML.GetAttribute("TASKINSTANCE"));
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER",	m_sAppFactFindNo);
	XML.SetAttribute("CASETASKNAME",				m_TaskXML.GetAttribute("CASETASKNAME"));
	
	XML.ActiveTag = reqTag;
	XML.CreateActiveTag("CURRENTEMPLOYERSREF");
	XML.SetAttribute("CUSTOMERNUMBER",				m_TaskXML.GetAttribute("CUSTOMERIDENTIFIER"));
	XML.SetAttribute("CUSTOMERVERSIONNUMBER",		GetCustomerVersion(m_TaskXML.GetAttribute("CUSTOMERIDENTIFIER")));
	XML.SetAttribute("EMPLOYMENTSEQUENCENUMBER",	m_TaskXML.GetAttribute("CONTEXT"));
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

	if(XML.IsResponseOK())
	{
		XML.SelectTag(null, "APPSTATUS");
		if(XML.GetAttribute("UNDERREVIEWIND") == "1")
		{
			alert("The Application has been placed under review");
			scScreenFunctions.SetContextParameter(window,"idAppUnderReview","1");
		}	
		return true;
	}
	return false;
}
function btnConfirm.onclick()
{
	// EP1039
	if (m_bSubmitClicked == true) return ;
	m_bSubmitClicked = true ; 

	if (confirm("Please ensure all data has been entered correctly before continuing"))
	{
		//run rules and set task status.
		if(UpdateEmpRef())
		{
			if(ValidateEmpRef())
			{
				<% /* MF MAR30 Only call calcs if data has changed */ %>
				var bContinue = true;
				if(IsChanged()){
					bContinue = RunIncomeCalculations();
				}				
				if (bContinue)	
				{
					scScreenFunctions.SetContextParameter(window,"idTaskXML","");
					frmToTM030.submit();
				}
				// EP1039
				else m_bSubmitClicked = false ;
			}
			// EP1039
			else m_bSubmitClicked = false ;
		}
		// EP1039
		else m_bSubmitClicked = false ;
	}
	// EP1039
	else m_bSubmitClicked = false ;
}

function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idTaskXML","");
	frmToTM030.submit();
}
function btnSubmit.onclick()
{
	// PJO 23/11/2005 MARS477 - Check for doubleclick
	if (m_bSubmitClicked == true)
	{
		return ;
	}
	m_bSubmitClicked = true ;
	// PJO MARS477 End
		
	scScreenFunctions.progressOn("Submitting...", 400); <% //MAR1525 - Peter Edney - 29/03/2006 %>
	
	if(frmScreen.onsubmit())
	{
		if(IsChanged())
		{	
			if(UpdateEmpRef())
			{
				<% /* BMIDS00916 MDC 16/11/2002 */ %>
				// if (CalculateAllowableIncome())	//AW	24/10/2002	BMIDS00653				
				if (RunIncomeCalculations())
				{
					scScreenFunctions.SetContextParameter(window,"idTaskXML","");
					frmToTM030.submit();
				}
			}
		}
	}
	// EP1039
	//	m_bSubmitClicked == false ;
	scScreenFunctions.progressOff(); <% //MAR1525 - Peter Edney - 29/03/2006 %>
	frmToTM030.submit(); //MAR1821 exit cleanly
}
<% /* PSC 25/01/2006 MAR1123 */ %>
function PopulateTable(nStart)
{
	TableXML.ActiveTag = null;
	TableXML.CreateTagList("LISTENTRY");
	var iCount;

	<% /* PSC 25/01/2006 MAR1123 - Start */ %>
	for (iCount = 0; iCount < TableXML.ActiveTagList.length && iCount <  m_nTableLength; iCount++)
	{
		<% /* PSC 25/01/2006 MAR1123 */ %>
		TableXML.SelectTagListItem(iCount + nStart);

		var sIncomeType			= TableXML.GetTagText("VALUENAME");
		var sAmount				= TableXML.GetTagText("EARNEDINCOMEAMOUNT");
		var sFrequency			= TableXML.GetTagAttribute("PAYMENTFREQUENCYTYPE","TEXT");
		var sConfirmedAmount	= TableXML.GetTagText("CONFIRMEDINCOMEAMOUNT");
		var sConfirmedFrequency = TableXML.GetTagAttribute("CONFIRMEDPAYMENTFREQUENCYTYPE","TEXT");
		var sFrequency			= TableXML.GetTagText("PAYMENTFREQUENCYTYPE");
		var sFrequencyText		= "";
		
		//
		//
		//
		if(sFrequency != "")
		{
			for (iCount2 = 0; iCount2 < frmScreen.cboFrequency.options.length; iCount2++)
			{
				if (frmScreen.cboFrequency.options(iCount2).value == sFrequency)
				{
					sFrequencyText = frmScreen.cboFrequency.options(iCount2).text;
				}
			}
		}					

		// Add to the search table
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sIncomeType);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),sAmount);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),sFrequencyText);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),sConfirmedAmount);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),sConfirmedFrequency);
		<% /* PSC 25/01/2006 MAR1123 */ %>
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}
}
function PopulateTableXML()
{
	<%/* Populate the TableXML with the values retrieved in XML */%>
	XML.ActiveTag = null;
	XML.CreateTagList("EARNEDINCOME");
	var iCount;
	var iListLength = XML.ActiveTagList.length;
	var bPopulate = false;

	//Check whether there are elements
	if(iListLength==1)
	{
		XML.SelectTagListItem(0);
		if(XML.GetAttribute("EARNEDINCOMEAMOUNT")=="")
			bPopulate = false;	//Element is just <EARNEDINCOME/>
		else
			bPopulate = true;
	}
	else	
		bPopulate = true;

	if(bPopulate==true)
	{
		for (iCount = 0; iCount < iListLength; iCount++)
		{
			XML.SelectTagListItem(iCount);

			GetTableItemByIncomeType(XML.GetAttribute("EARNEDINCOMETYPE"));
						
			TableXML.SetTagText("EARNEDINCOMESEQUENCENUMBER",iCount+1);
			TableXML.SetTagText("EARNEDINCOMEAMOUNT",XML.GetAttribute("EARNEDINCOMEAMOUNT"));
			TableXML.SetTagText("PAYMENTFREQUENCYTYPE",XML.GetAttribute("PAYMENTFREQUENCYTYPE"));
			TableXML.SetAttributeOnTag("PAYMENTFREQUENCYTYPE","TEXT",XML.GetAttribute("PAYMENTFREQUENCYTYPETEXT"));
		}
	}
	return(iListLength);
}
function CalculateTotals()
{
	TableXML.ActiveTag = null;
	TableXML.CreateTagList("LISTENTRY");
	var iCount = 0;
	var nSingleGrossIncome = 0;
	var nTotalGrossIncome = 0;

	for (iCount = 0; iCount < m_nIncomeTypeCount; iCount++)
	{
		TableXML.SelectTagListItem(iCount);

		nSingleGrossIncome = parseFloat(TableXML.GetTagText("CONFIRMEDINCOMEAMOUNT")) *
							 parseInt(TableXML.GetTagText("CONFIRMEDPAYMENTFREQUENCYTYPE"));
		
		nTotalGrossIncome = nTotalGrossIncome + nSingleGrossIncome;
	}

	frmScreen.txtGrossIncome.value = nTotalGrossIncome;
}

function GetTableItemByIncomeType(sIncomeType)
{
	TableXML.ActiveTag = null;
	TableXML.CreateTagList("LISTENTRY");
	var iCount = 0;
	var iListLength = TableXML.ActiveTagList.length;
	var bFoundItem = false;

	for (iCount = 0; (iCount < iListLength) & !bFoundItem; iCount++)
	{
		TableXML.SelectTagListItem(iCount);
		bFoundItem = (TableXML.GetTagText("VALUEID") == sIncomeType);
	}
}
function GetTableItemBySelection()
{
	TableXML.ActiveTag = null;
	TableXML.CreateTagList("LISTENTRY");

	<% /* Get the index of the selected row */ %>
	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();

	TableXML.SelectTagListItem(nRowSelected-1);
}
function InitialiseTableXML()
{
	<%/* The TableXML will have been retrieved using GetComboLists for 'IncomeType'. That XML now
		 needs the extra fields adding to it that are used within the table onscreen. */%>
	TableXML.ActiveTag = null;
	TableXML.CreateTagList("LISTENTRY");
	var iCount = 0;
	var iListLength = TableXML.ActiveTagList.length;
	var sMonthly = "";

	<%/* Ascertain the string representing the 'Monthly' frequency */%>
	for (iCount = 0; iCount < frmScreen.cboFrequency.options.length; iCount++)
		if (frmScreen.cboFrequency.options(iCount).value == "1")
			sMonthly = frmScreen.cboFrequency.options(iCount).text;

	for (iCount = 0; iCount < iListLength ; iCount++)
	{
		TableXML.SelectTagListItem(iCount);
		TableXML.CreateTag("EARNEDINCOMEAMOUNT",			"0");
		TableXML.CreateTag("PAYMENTFREQUENCYTYPE",			"1");
		TableXML.SetAttributeOnTag("PAYMENTFREQUENCYTYPE","TEXT",sMonthly);
		TableXML.CreateTag("EARNEDINCOMESEQUENCENUMBER",iCount+1);
		TableXML.CreateTag("CONFIRMEDINCOMEAMOUNT",			"0");
		TableXML.CreateTag("CONFIRMEDPAYMENTFREQUENCYTYPE",	"1");
		TableXML.SetAttributeOnTag("CONFIRMEDPAYMENTFREQUENCYTYPE","TEXT",sMonthly);

	}
	m_nIncomeTypeCount = iListLength;
}

function frmScreen.cboFrequency.onchange()
{
	var nRowSelected = scTable.getRowSelected();
	UpdateTableXML(nRowSelected);
	CalculateTotals();
}

function frmScreen.txtAmount.onfocus()
{
	frmScreen.txtAmount.select();
}

function frmScreen.txtAmount.onchange()
{
	var nRowSelected = scTable.getRowSelected();
	UpdateTableXML(nRowSelected);
	
	if(frmScreen.txtAmount.value == "" ||frmScreen.txtAmount.value == "NaN")
	{
		frmScreen.txtAmount.value = "0"
	}
	CalculateTotals();
}

function spnToAmount.onfocus()
{
	var nRowSelected = scTable.getRowSelected();

	if ((nRowSelected + scTable.getOffset() <= m_nIncomeTypeCount) && (nRowSelected != 10))
	{
		<%/* Move onto the next earned income in the list */%>
		scTable.setRowSelected(nRowSelected + 1);
		spnTable.onclick();
		frmScreen.txtAmount.focus();
	}	
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToFrequency.onfocus()
{
	var nRowSelected = scTable.getRowSelected();
	if (nRowSelected == 1)
		btnCancel.focus();
	else
	{
		<%/* Move onto the previous earned income in the list */%>
		scTable.setRowSelected(nRowSelected - 1);
		spnTable.onclick();
		frmScreen.cboFrequency.focus();
	}
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function spnTable.onclick()
{
	<% /*MO - 23/10/2002 - BMIDS00649 Put in the check so that you cant go off the end of the table 
		if the selected row is greater than the number of rows in the table, set it to the last number */ %>
	if ((scTable.getOffset() + scTable.getRowSelected()) > m_nIncomeTypeCount) {
		scTable.setRowSelected(m_nIncomeTypeCount - scTable.getOffset());
	}
	
	GetTableItemBySelection();
	
	var dConfirmedAmount = parseFloat(TableXML.GetTagText("CONFIRMEDINCOMEAMOUNT"));
	var nConfirmedFrequency = TableXML.GetTagText("CONFIRMEDPAYMENTFREQUENCYTYPE");

	if (nConfirmedFrequency == "0") nConfirmedFrequency = "";

	frmScreen.txtAmount.value = dConfirmedAmount;
	frmScreen.cboFrequency.value = nConfirmedFrequency;
}

function frmScreen.optIncOtherIncomeNo.onclick()
{
	frmScreen.txtOtherIncomePercentage.value="";
	frmScreen.txtOtherIncomePercentage.disabled=true;
	frmScreen.txtOtherIncomePercentage.readonly=true;
	frmScreen.txtOtherIncomePercentage.setAttribute("required","false");
	<% /* Set the colour of the labels back to the normal colour */ %>
	frmScreen.txtOtherIncomePercentage.parentElement.parentElement.style.color
		= frmScreen.txtOtherIncomePercentage.style.color;	
	frmScreen.txtOtherIncomePercentage.parentElement.childNodes.item(2).style.color
		= frmScreen.txtOtherIncomePercentage.style.color;
}

function frmScreen.optIncOtherIncomeYes.onclick()
{
	frmScreen.txtOtherIncomePercentage.disabled=false;
	frmScreen.txtOtherIncomePercentage.readonly=false;
	frmScreen.txtOtherIncomePercentage.setAttribute("required","true");	
	<% /* Set the mandatory red colour on the label and the % sign */ %>
	frmScreen.txtOtherIncomePercentage.parentElement.parentElement.style.color = "red";	
	frmScreen.txtOtherIncomePercentage.parentElement.childNodes.item(2).style.color = "red";
}

function UpdateTableXML(nRowSelected)
{
	GetTableItemBySelection();
	
	if(frmScreen.txtAmount.value == "" ||frmScreen.txtAmount.value == "NaN")
	{
		frmScreen.txtAmount.value = "0"
	}
	
	var sAmount = frmScreen.txtAmount.value;
	var sFrequency = frmScreen.cboFrequency.options(frmScreen.cboFrequency.selectedIndex).text;
	var nFrequency = frmScreen.cboFrequency.value;

	TableXML.SetTagText("CONFIRMEDINCOMEAMOUNT",sAmount);
	TableXML.SetTagText("CONFIRMEDPAYMENTFREQUENCYTYPE",nFrequency);
	TableXML.SetAttributeOnTag("CONFIRMEDPAYMENTFREQUENCYTYPE","TEXT",sFrequency);

	scScreenFunctions.SizeTextToField(tblTable.rows(nRowSelected).cells(3),sAmount);
	scScreenFunctions.SizeTextToField(tblTable.rows(nRowSelected).cells(4),sFrequency);

}

function DefaultFields()
{
	frmScreen.cboFrequency.selectedIndex = 0;
	frmScreen.txtAmount.value = "0";
	<% /* MF 08/09/2005 MAR30 */ %>
	if(EnableOtherIncomeRadio()){
		frmScreen.optIncOtherIncomeNo.disabled=false;
		frmScreen.optIncOtherIncomeNo.checked=true;
		frmScreen.optIncOtherIncomeYes.disabled=false;		
	} 
}

<% /* MF 08/09/2005 MAR30 
	Read User Role, if this is Underwriter then return true
	Validation type for this is 'UW'
*/ %>
function EnableOtherIncomeRadio()
{	
	var bExcludeOthIncUnlessUnderwriter=false;
	var bUnderwriter=false;
	var GlobalParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	bExcludeOthIncUnlessUnderwriter = GlobalParamXML.GetGlobalParameterBoolean(document,"ExcludeOthIncUnlessUnderwriter");
	if(bExcludeOthIncUnlessUnderwriter){
		var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole");
		var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if (TempXML.IsInComboValidationList(document,"UserRole",sUserRole,["UW"]))
			bUnderwriter=true;
	}
	return bExcludeOthIncUnlessUnderwriter && bUnderwriter;
}

function UpdateTableXMLConfirmed()
{
	<%/* Populate the TableXML with the values retrieved in XML */%>
	m_CurrEmpXML.ActiveTag = null;
	m_CurrEmpXML.CreateTagList("CONFIRMEDEARNEDINCOME");
	var iCount;
	var iListLength = m_CurrEmpXML.ActiveTagList.length;

	for (iCount = 0; iCount < iListLength; iCount++)
	{
		m_CurrEmpXML.SelectTagListItem(iCount);

		GetTableItemByIncomeType(m_CurrEmpXML.GetAttribute("EARNEDINCOMETYPE"));
				
		TableXML.SetTagText("CONFIRMEDINCOMEAMOUNT",m_CurrEmpXML.GetAttribute("EARNEDINCOMEAMOUNT"));
		TableXML.SetTagText("CONFIRMEDPAYMENTFREQUENCYTYPE",m_CurrEmpXML.GetAttribute("PAYMENTFREQUENCYTYPE"));		
		
		<% /* PSC 25/01/2006 MAR1123 - Start */ %>
		var sConfirmedFrequency = m_CurrEmpXML.GetAttribute("PAYMENTFREQUENCYTYPE");
		var sConfirmedFrequencyText = "";
		//
		//
		//
		if(sConfirmedFrequency != "")
		{
			for (iCount2 = 0; iCount2 < frmScreen.cboFrequency.options.length; iCount2++)
			{
				if (frmScreen.cboFrequency.options(iCount2).value == sConfirmedFrequency)
				{
					sConfirmedFrequencyText = frmScreen.cboFrequency.options(iCount2).text;
				}
			}
		}			

		<% /* PSC 25/01/2006 MAR1123 - Start */ %>
		TableXML.SetAttributeOnTag("CONFIRMEDPAYMENTFREQUENCYTYPE","TEXT",sConfirmedFrequencyText);
		
		TableXML.SetTagText("EARNEDINCOMESEQUENCENUMBER",iCount+1);		
	}

	return(iListLength);
}

<% /* BMIDS00654 MDC 04/11/2002 */ %>
<% /* AW	24/10/2002	BMIDS00653 */ %>
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
	AllowableIncXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
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
	
	<% /* BMIDS01034 MDC 21/11/2002 - Always return true to allow user to progress
	if(AllowableIncXML.IsResponseOK())
	{
		return(true);
	}

	return(false); */ %>
	AllowableIncXML.IsResponseOK()
	return(true);
	<% /* BMIDS01034 MDC 21/11/2002 - End */ %>
}
<%	/*BMIDS00653  - End */ %>
<% /* BMIDS00654 MDC 04/11/2002 - End */ %>

-->
</script>
</body>
</html>
<% /* OMIGA BUILD VERSION 028.02.02.15.00 */ %>



