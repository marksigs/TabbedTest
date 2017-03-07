<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      DC160.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Employment Summary screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		06/12/1999	Created
AD		04/02/2000	Rework
AD		02/03/2000	Fixed SYS0336
AD		13/03/2000	Fixed SYS0444
AD		13/03/2000  Fixed SYS0434
AY		30/03/00	New top menu/scScreenFunctions change
IVW		06/04/2000	Fixed SYS0526 - Focus on first field only.
IW		14/04/00	Cancel routes back to DC155 (Regular Outgoings)
BG		SYS0746		Double clicking item in the table selects it in edit mode.
MC		23/05/00	SYS0756 - Main Employment count correct if > 10 employers
							  If read-only mode disable controls.
MC		01/06/00	Add Tax Details popup
BG		25/07/00	SYS0971 - Made customer name field longer to handle max length of customer name
BG		SYS1224		clear context parameter on btnAdd.onclick event.
CL		05/03/01	SYS1920 Read only functionality added
SA		16/05/01	SYS1947 Moved line of code to set context parameter idMainEmploymentCount
SA		24/05/01	SYS0910 Radio button UK Tax Payer fix.
JLD		10/12/01	SYS2806 use scScreenFunctions completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

prog	Date		AQR			Description
MV		22/08/2002	BMIDs00355	IE 5.5 upgrade - Modified the Msgbuttons Position , and 	width and height of the Popupwindow DC165.asp
TW      09/10/2002    Modified to incorporate client validation - SYS5115							
AW		17/10/2002	BMIDS00653 - BM089 Added Gross and Net Allowable Income fields
AW		18/10/2002	BMIDS00653 - BM089 Member of staff	
AW		24/10/2002	BMIDS00653	BM029 Call to Allowable Income							
MDC		01/11/2002	BMIDS00654  ICWP BM088 Income Calculations
GHun	08/11/2002	BMIDS00882	Non numeric customer number fails
AW		17/11/2002	BMIDS00932	Only disply Allowable income of first two applicants
								when more than 2 applicants exist and have been re-ordered.
								Call RunIncomeCalculations() when screen re-initialised
SA		17/11/2002	BMIDS00934	Disable Delete button when application is frozen
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
DPF		22/11/2002  BMIDS01032	Have disabled Edit / Delete buttons if there are no employment records
GHun	02/12/2002	BMIDS01093	Fixed word spacing and alignment
MV		20/02/2003	BM0342		Amended btnDelete.onclick()
GHun	13/03/2003	BM0457		Include guarantors in allowable income calculation
GHun	04/04/2003	BM0514		RunIncomeCalcs doesn't need to be called on entry when coming from DC155
BS		11/06/2003	BM0521		Pass in different ReadOnly parameter to DC165
HMA     17/09/2003  BM0063      Amend HTML text for radio buttons
PJO     24/09/2003  BMIDS621    If screen is read only don't run Income Calcs on OK
MC		19/04/2004	BMIDS517	Popup Dialog height incr. by 5px (DC165 dialog)
SR		19/06/2004  BMIDS772	route to DC110 instead of DC155 (on submit)
JD		29/09/2004	BMIDS895	No need to RunIncomeCalcs on Delete
JD		04/10/2004	BMIDS895	Catch exit from screen not via ok or cancel
JD		11/10/2004	BMIDS895	save customerXML on entry to screen as exit from app wipes the context before onBeforeUnload() is called.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog    Date			AQR		Description
MF		22/07/2005		MAR19	IA_WP01 process flow changes
MF		15/09/2005		MAR30	IncomeCalcs only called where data has changed
								IncomeCalcs modified to send ActivityID into calculation
PJO     17/11/2005      MAR440  Conditions for doing Max borrowing
PJO     20/12/2005      MAR825  Take DC195 out of routing on Global Parameter setting 
DRC     31/01/2006      MAR1170 Put in Critical Data check under RunIncomeCalculations
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom Specific History

Prog    Date			AQR		Description
MAH		20/11/2006		E2CR35	Added SelfAssess field to XML to pass to Income Tax popup
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<% // PJO 20/12/2005 MAR825 %>
<object data="scXMLFunctions.asp" height="1" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex="-1" type="text/x-scriptlet" width="1" viewastext></object>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>
<span id="spnListScroll">
	<span style="LEFT: 306px; POSITION: absolute; TOP: 274px">
<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable 
style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet 
VIEWASTEXT></OBJECT>
	</span> 
</span>


<%/* FORMS */%>
<form id="frmToDC170" method="post" action="DC170.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC195" method="post" action="DC195.asp" STYLE="DISPLAY: none"></form>
<% // PJO 20/12/2005 MAR825 %>
<form id="frmToDC085" method="post" action="DC085.asp" STYLE="DISPLAY: none"></form>
<% /* SR 19/06/2004 : BMIDS772 - route to DC110 instead of DC155 (on cancel) */ %>
<form id="frmToDC110" method="post" action="DC110.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC060" method="post" action="dc060.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark      validate ="onchange">
<div id="divBackground" style="HEIGHT: 380px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Customer Name
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtCustomerName" style="WIDTH: 430px" class="msgReadOnly" readonly tabindex=-1>
			</input>
		</span> 
	</span>

	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 36px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="25%" class="TableHead">Status&nbsp;</td>
								<td width="20%" class="TableHead">Employer Name&nbsp;</td>
								<td width="20%" class="TableHead">Occupation&nbsp;</td>
								<td width="15%" class="TableHead">Gross Salary pa&nbsp;</td>
								<td width="10%" class="TableHead">Date Started/Est&nbsp;</td>
								<td width="10%" class="TableHead">Main Employment?&nbsp;</td></tr>
			<tr id="row01">		<td width="25%" class="TableTopLeft">&nbsp;</td>			<td width="20%" class="TableTopCenter">	&nbsp;</td>			<td width="20%" class="TableTopCenter">&nbsp;</td>	<td width="15%" class="TableTopCenter">&nbsp;</td>	<td width="10%" class="TableTopCenter">&nbsp;</td>	<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="20%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	    <td width="15%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="20%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	    <td width="15%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="20%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	    <td width="15%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="20%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	    <td width="15%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="20%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	    <td width="15%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="20%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	    <td width="15%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="20%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	    <td width="15%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="20%" class="TableCenter">&nbsp;</td>			<td width="20%" class="TableCenter">&nbsp;</td>	    <td width="15%" class="TableCenter">&nbsp;</td>		<td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="25%" class="TableBottomLeft">&nbsp;</td>		<td width="20%" class="TableBottomCenter">&nbsp;</td>		<td width="20%" class="TableBottomCenter">&nbsp;</td> <td width="15%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td><td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 214px">
		<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnAdd.onClick()"> 
	</span>

	<span style="LEFT: 68px; POSITION: absolute; TOP: 214px">
		<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnEdit.onClick()"> 
	</span>

	<span style="LEFT: 132px; POSITION: absolute; TOP: 214px">
		<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnDelete.onClick()"> 
	</span> 

	<span style="LEFT: 10px; POSITION: absolute; TOP: 250px" class="msgLabel">
		Current Highest Tax Rate
		<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
			<select id="cboHighestTaxRate" name="HighestTaxRate" style="WIDTH: 80px" class="msgCombo">
			</select>
		</span> 
	</span>

	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 274px" class="msgLabel">
		UK Tax Payer?
		<span style="LEFT: 186px; POSITION: absolute; TOP: -3px">
			<input id="optUKTaxPayerYes" name="UKTaxPayerGroup" type="radio" value="1"><label for="optUKTaxPayerYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 231px; POSITION: absolute; TOP: -3px">
			<input id="optUKTaxPayerNo" name="UKTaxPayerGroup" type="radio" value="0"><label for="optUKTaxPayerNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 298px" class="msgLabel">
		Total Gross Earned Annual Income
		<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 3px" class="msgCurrency"></label>
			<input id="txtGrossIncome" name="GrossIncome" style="WIDTH: 80px" class="msgReadOnly" readonly tabindex=-1>
			</input>
		</span> 
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 322px" class="msgLabel">
		Total Net Monthly Income
		<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 3px" class="msgCurrency"></label>
			<input id="txtNetIncome" name="NetIncome" style="WIDTH: 80px" class="msgReadOnly" readonly tabindex=-1>
			</input>
		</span> 
	</span>
	
	<span id="spnGrossAllowable" style="LEFT: 305px; POSITION: absolute; TOP: 298px" class="msgLabel">
		Total Gross Allowable Annual Income
		<span style="LEFT: 207px; POSITION: absolute; TOP: -3px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 3px" class="msgCurrency"></label>
			<input id="txtGrossAllowableIncome" name="GrossAllowableIncome" style="WIDTH: 80px" class="msgReadOnly" readonly tabindex=-1>
			</input>
		</span> 
	</span>

	<span id="spnNetAllowable" style="LEFT: 305px; POSITION: absolute; TOP: 322px" class="msgLabel">
		Total Net Allowable Annual Income
		<span style="LEFT: 207px; POSITION: absolute; TOP: -3px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 3px" class="msgCurrency"></label>
			<input id="txtNetAllowableIncome" name="NetAllowableIncome" style="WIDTH: 80px" class="msgReadOnly" readonly tabindex=-1>
			</input>
		</span> 
	</span>
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 350px" class="msgLabel">
		Tax Details
		<span style="LEFT: 190px; POSITION: absolute; TOP: -8px">
			<input id="btnTaxDetails" type="button" style="HEIGHT: 26px; WIDTH: 26px" class ="msgDDButton">
		</span>
	</span>

</div> 
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 450px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc160Attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sCustomerName = "";
var m_nCustomerIndex = 0;
var EmploymentXML = null;
var scScreenFunctions;
var nMainEmploymentCount = 0;

var m_sTaxOffice = "";
var m_sTaxReference = "";
var m_sNINumber = "";
var m_sNonUKTaxArrangements = "";
var m_sTaxedOutsideUKInd = "";
var m_sSelfAssess = ""; <% /* MAH 20/11/2006 E2CR35 */ %>
var m_blnReadOnly = false;
//AW		24/10/2002	BMIDS00653
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
var bDelete = false; //JD BMIDS895
var bExitOk = false; //JD BMIDS895
var m_CustomerXML = null; //BMIDS895

/* EVENTS */

function frmScreen.btnAdd.onclick()
{
	if (CommitData())
	{
		scScreenFunctions.SetContextParameter(window,"idCustomerNumber", m_sCustomerNumber);
		scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber", m_sCustomerVersionNumber);
		scScreenFunctions.SetContextParameter(window,"idCustomerIndex", m_nCustomerIndex);
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "Add");
		//BG SYS1224 clear context parameter.
		scScreenFunctions.SetContextParameter(window,"idEmploymentSequenceNumber", "");
		frmToDC170.submit();
	}
}

function btnCancel.onclick()
{
	bExitOk = true;  //JD BMIDS895
	if(scScreenFunctions.CompletenessCheckRouting(window))
	{
		frmToGN300.submit();
		return;
	}
	m_nCustomerIndex--;
	scScreenFunctions.SetContextParameter(window,"idCustomerIndex", m_nCustomerIndex);

	if (RetrieveCustomerData())
	{
		scTable.clear();
		<% /* BM0514 call Initialise with bRunIncomeCalcs set to true when cancelling */ %>
		Initialise(true);
	}
	else
	{
		scScreenFunctions.SetContextParameter(window,"idCustomerIndex", "");	
		<% /* SR 19/06/2004 : BMIDS772 - route to DC110 instead of DC155 */ %>	
		<% /* MF 22/07/2005 MARS IA_WP01 route back to Address summary
		frmToDC110.submit(); */ %>
		frmToDC060.submit();
	}
}

function frmScreen.btnDelete.onclick()
{
	if (!confirm("Are you sure?")) return;

	//Get the XML that just contains the GroupConnection chosen in the listbox
	var XML = GetXMLBlock(true);
	var bMain = false;
	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();
	if(EmploymentXML.GetTagText("MAINSTATUS") == "1")
		bMain = true;
		

	// 	XML.RunASP(document,"DeleteEmployment.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"DeleteEmployment.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if (XML.IsResponseOK())
	{
		<% /*JD BMIDS895 RunIncomeCalculations(); no need to call this here. 
		     If we deleted the main employment we should set allowable income to zero*/ %>
		bDelete = true;
		scTable.clear();
		PopulateScreen();
		if(bMain)
		{
			frmScreen.txtGrossAllowableIncome.value = "0.00";
			frmScreen.txtNetAllowableIncome.value = "0.00";
		}
	}
}
function window.onbeforeunload()
{
<% /* JD BMIDS895 if we have done any deleting but are not exiting 
      via ok or cancel call RunIncomeCalcs */ %>
	if(bExitOk == false && bDelete == true)
	{
		RunIncomeCalculations();
	}
}
function frmScreen.btnEdit.onclick()
{
	if (CommitData())
	{
		var XML = GetXMLBlock(true);
		XML.SelectTag(null,"EMPLOYMENT");
		scScreenFunctions.SetContextParameter(window,"idCustomerNumber", m_sCustomerNumber);
		scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber", m_sCustomerVersionNumber);
		scScreenFunctions.SetContextParameter(window,"idCustomerIndex", m_nCustomerIndex);
		scScreenFunctions.SetContextParameter(window,"idEmploymentSequenceNumber", XML.GetTagText("EMPLOYMENTSEQUENCENUMBER"));
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "Edit");
		frmToDC170.submit();
	}
}

function frmScreen.btnTaxDetails.onclick()
{
	var sReturn = "";
	var ArrayArguments = new Array(4);
	var IncomeSummaryXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	IncomeSummaryXML.CreateActiveTag("INCOMESUMMARY");
	IncomeSummaryXML.CreateTag("TAXOFFICE", m_sTaxOffice);
	IncomeSummaryXML.CreateTag("TAXREFERENCENUMBER", m_sTaxReference);
	IncomeSummaryXML.CreateTag("NATIONALINSURANCENUMBER", m_sNINumber);
	if(m_sTaxedOutsideUKInd == "")
	{
		if(scScreenFunctions.GetRadioGroupValue(frmScreen,"UKTaxPayerGroup") == "1")
			IncomeSummaryXML.CreateTag("TAXEDOUTSIDEUKINDICATOR", "0");
		else
			IncomeSummaryXML.CreateTag("TAXEDOUTSIDEUKINDICATOR", "1");
	}
	else
		IncomeSummaryXML.CreateTag("TAXEDOUTSIDEUKINDICATOR", m_sTaxedOutsideUKInd);
		
	IncomeSummaryXML.CreateTag("SELFASSESS", m_sSelfAssess);<% /* MAH 20/11/2006 E2CR35 */ %>
	IncomeSummaryXML.CreateTag("NONUKTAXDETAILS", m_sNonUKTaxArrangements);
	
	ArrayArguments[0] = m_sCustomerName;
	ArrayArguments[1] = m_sCustomerNumber;
	ArrayArguments[2] = m_sCustomerVersionNumber;
	ArrayArguments[3] = IncomeSummaryXML.XMLDocument.xml;
	<% /* BS BM0521 11/06/03
	ArrayArguments[4] = m_sReadOnly;*/ %>
	if (m_blnReadOnly) ArrayArguments[4] = "1";
	else ArrayArguments[4] = "0";
	<% /* BS BM0521 End 11/06/03 */ %>

	sReturn = scScreenFunctions.DisplayPopup(window, document, "dc165.asp", ArrayArguments, 610, 390);<% /*MAH 20/11/2006 E2CR35*/ %>
	if(sReturn != "")
	{
		IncomeSummaryXML.ResetXMLDocument();
		IncomeSummaryXML.LoadXML(sReturn);
		IncomeSummaryXML.SelectTag(null, "INCOMESUMMARY");
		if(IncomeSummaryXML.GetAttribute("UPDATED") == "YES")
		{
			FlagChange(true);
			m_sTaxOffice = IncomeSummaryXML.GetTagText("TAXOFFICE");
			m_sTaxReference = IncomeSummaryXML.GetTagText("TAXREFERENCENUMBER");
			m_sNINumber = IncomeSummaryXML.GetTagText("NATIONALINSURANCENUMBER");
			m_sNonUKTaxArrangements = IncomeSummaryXML.GetTagText("NONUKTAXDETAILS");
			m_sTaxedOutsideUKInd = IncomeSummaryXML.GetTagText("TAXEDOUTSIDEUKINDICATOR");
			m_sSelfAssess = IncomeSummaryXML.GetTagText("SELFASSESS");<% /* MAH 20/11/2006 E2CR35 */ %>
			
		}
	}

}

function btnSubmit.onclick()
{
	if (ValidateScreen())
	{
		if(CommitData())
		{
			<% /* BMIDS00654 MDC 01/11/2002 */ %>
			<% /* MF MAR30 Call Income Calcs if data has changed and continue whether 
					return flag is true OR false */ %>
			if(IsChanged()){
				RunIncomeCalculations();		//AW	24/10/2002	BMIDS00653
			}
			
			bExitOk = true;  //JD BMIDS895
			if(scScreenFunctions.CompletenessCheckRouting(window))
			{
				frmToGN300.submit();
				return;
			}
			m_nCustomerIndex++;
			scScreenFunctions.SetContextParameter(window,"idCustomerIndex", m_nCustomerIndex);
			if (RetrieveCustomerData())
			{
				scTable.clear();
				<% /* BM0514 call Initialise with bRunIncomeCalcs set to true when submitting */ %>
				Initialise(true);
			}
			else
			{
				scScreenFunctions.SetContextParameter(window,"idCustomerIndex", "");
				<% // PJO 20/12/2005 MAR825 - Show / hide other income on global parameter %>
				var GlobalParamXML = new scXMLFunctions.XMLObject()
				if (GlobalParamXML.GetGlobalParameterBoolean(document,"OtherIncomeSummary"))
				{
					frmToDC195.submit();
				}
				else
				{	
					frmToDC085.submit();
				}
			}
		}
	}
}		

function spnTable.ondblclick()
{
	if (scTable.getRowSelectedIndex() != null) frmScreen.btnEdit.onclick();
}

function ValidateScreen()
{
	if (nMainEmploymentCount == 0)
	{
		alert("Main employment type must be specified");
		return false
	}	
	else
		return true
}
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Employment Summary","DC160",scScreenFunctions);
	//	AW	BMIDS00653
	scScreenFunctions.SetCurrency(window,frmScreen);
	RetrieveContextData();
	if (m_nCustomerIndex == "") m_nCustomerIndex = 1;
	RetrieveCustomerData();
	
	<% /* BMIDS621 We need to do this before initilise */ %>
    m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC160");
	<% /* BM0514 RunIncomeCalcs does not need to be rerun if coming from dc155 or read-only */ %>
	<% /* MAR440 PJO 17/11/2005 We no longer route to this screen from DC155
	if (document.referrer.toUpperCase().indexOf("DC155.ASP") > -1)
	{
		Initialise(false);
    }
	else   // PJO BMIDS621
	{
	MAR440 End */ %>

	    if (m_blnReadOnly)
	    {
	        Initialise(false);
	    }
	    else
		{
		    Initialise(true);
		}
    //}
	<% /* BM0514 End */ %>
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	if(m_sReadOnly == "1")
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	if (m_blnReadOnly == true) {
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnDelete.disabled = true;
		m_sReadOnly = "1" ; // PJO BMIDS621
	}
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

function CommitData()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				bSuccess = SaveTaxDetails();
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function DefaultFields()
{
	scScreenFunctions.SetRadioGroupValue(frmScreen,"UKTaxPayerGroup","1");
	frmScreen.cboHighestTaxRate.selectedIndex = 0;
}

function GetXMLBlock(bForEdit)
{
	EmploymentXML.ActiveTag = null;
	EmploymentXML.CreateTagList("EMPLOYMENT");

	<% /* Get the index of the selected row */ %>
	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	if(EmploymentXML.SelectTagListItem(nRowSelected-1) == true)
	{
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("EMPLOYMENT");

		XML.CreateTag("CUSTOMERNUMBER", EmploymentXML.GetTagText("CUSTOMERNUMBER"));
		XML.CreateTag("CUSTOMERVERSIONNUMBER", EmploymentXML.GetTagText("CUSTOMERVERSIONNUMBER"));
		if(bForEdit)
			XML.CreateTag("EMPLOYMENTSEQUENCENUMBER", EmploymentXML.GetTagText("EMPLOYMENTSEQUENCENUMBER"));
	}

	return(XML);
}
<% /* JD BMIDS895 SetUpCustomerXML added. saves XML block with customer information from the context for use in RunIncomeCalc */ %>
function SetUpCustomerXML()
{
	m_CustomerXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	TagCUSTOMERLIST = m_CustomerXML.CreateActiveTag("CUSTOMERLIST");
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
			m_CustomerXML.CreateActiveTag("CUSTOMER");
			m_CustomerXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			m_CustomerXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
			m_CustomerXML.CreateTag("CUSTOMERROLETYPE", sCustomerRoleType);
			m_CustomerXML.CreateTag("CUSTOMERORDER", sCustomerOrder);
			m_CustomerXML.ActiveTag = TagCUSTOMERLIST;
		}
	}	
}

<% /* BM0514 Added bRunIncomeCalcs parameter */ %>
function Initialise(bRunIncomeCalcs)
{
	scScreenFunctions.SetContextParameter(window,"idEmployerName", "");
	PopulateCombos();
	DefaultFields();
	//AW		17/11/2002	BMIDS00932
	
	<% /* BM0514 Only call RunIncomeCalcs if required */ %>
	if (bRunIncomeCalcs)
	{
		SetUpCustomerXML();  //JD BMIDS895
		RunIncomeCalculations();
	}
	
	PopulateScreen();
	if (EmploymentXML.SelectTag(null,"INCOMESUMMARYDETAILS") != null)
	{
		m_sMetaAction = "Edit";
		m_sTaxOffice = EmploymentXML.GetTagText("TAXOFFICE");
		m_sTaxReference = EmploymentXML.GetTagText("TAXREFERENCENUMBER");
		m_sNINumber = EmploymentXML.GetTagText("NATIONALINSURANCENUMBER");
		m_sTaxedOutsideUKInd = EmploymentXML.GetTagText("TAXEDOUTSIDEUKINDICATOR");
		m_sSelfAssess = EmploymentXML.GetTagText("SELFASSESS");<% /* MAH 20/11/2006 E2CR35 */ %>
		if(m_sTaxedOutsideUKInd == "1")
			m_sNonUKTaxArrangements = EmploymentXML.GetTagText("NONUKTAXDETAILS");
		else
			m_sNonUKTaxArrangements = "";
		
	}
	else
	{
		m_sMetaAction = "Add";
		m_sTaxOffice = "";
		m_sTaxReference = "";
		m_sNINumber = "";
		m_sNonUKTaxArrangements = "";
		m_sTaxedOutsideUKInd = "";
		m_sSelfAssess = ""; <% /*MAH 20/11/2006 E2CR35*/ %>
	}
	
}

function PopulateCombos()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var blnSuccess = true;

	var sGroupList = new Array("HighestTaxRate");
	if(XML.GetComboLists(document,sGroupList))
	{
		blnSuccess = blnSuccess & XML.PopulateCombo(document,frmScreen.cboHighestTaxRate,"HighestTaxRate",true);

		if(!blnSuccess)
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}

	XML = null;
}

function PopulateScreen()
{
	frmScreen.txtCustomerName.value = m_sCustomerName;
	EmploymentXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	EmploymentXML.CreateRequestTag(window,null);
	var tagCustomerList = EmploymentXML.CreateActiveTag("CUSTOMERLIST");
	EmploymentXML.CreateActiveTag("CUSTOMER");
	EmploymentXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	EmploymentXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);

	// 	EmploymentXML.RunASP(document,"FindEmploymentList.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			EmploymentXML.RunASP(document,"FindEmploymentList.asp");
			break;
		default: // Error
			EmploymentXML.SetErrorResponse();
		}


	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = EmploymentXML.CheckResponse(ErrorTypes);
	var nListLength = 0;
	
	if ((ErrorReturn[1] == ErrorTypes[0]) | (EmploymentXML.XMLDocument.text == ""))
	{
		//Error: record not found
		PopulateTotals();

		if(m_sReadOnly == "1")
		{
			frmScreen.btnAdd.disabled = true;
			//BMIDS00934 Disable Delete button too
			frmScreen.btnDelete.disabled = true;
		}
		else
			frmScreen.btnAdd.disabled = false;
		scScreenFunctions.SetContextParameter(window,"idMainEmploymentCount", "0");
		scScreenFunctions.SetContextParameter(window,"idEmploymentCount", "0");
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
		
		//frmScreen.cboHighestTaxRate.focus();
	}
	else if (ErrorReturn[0] == true)
	{
		// No error
		PopulateTable(0);
		<% /* SYS0756: Count main employments */ %>
		nMainEmploymentCount = CountMainEmployments();
		
		EmploymentXML.ActiveTag = null;
		EmploymentXML.CreateTagList("EMPLOYMENT");
		nListLength = EmploymentXML.ActiveTagList.length;
		scScreenFunctions.SetContextParameter(window,"idEmploymentCount", nListLength);
		// SYS1947 SA 16/5/01 Set Parameter after count of main employments has been done.
		scScreenFunctions.SetContextParameter(window,"idMainEmploymentCount", nMainEmploymentCount);

		PopulateTotals();
		if(m_sReadOnly == "1")
		{
			frmScreen.btnAdd.disabled = true;
			frmScreen.btnDelete.disabled = true;
			//IVW		06/04/2000	Fixed SYS0526 - Focus on first field only.
			//	frmScreen.btnEdit.focus();
		}
		else
		{
			//DPF 22/11/2002 - BMIDS01056 - If we have no employment records then disable Edit/Delete buttons
			if (nListLength == 0)
			{
				frmScreen.btnEdit.disabled = true;
				frmScreen.btnDelete.disabled = true;
				frmScreen.btnAdd.disabled = false;
			}
			else
			{
				frmScreen.btnEdit.disabled = false;
				frmScreen.btnDelete.disabled = false;
				frmScreen.btnAdd.disabled = false;
			}
			//END OF BMIDS01056
			
			//IVW		06/04/2000	Fixed SYS0526 - Focus on first field only.
			//	frmScreen.btnAdd.focus();
		}
	}
	
	ErrorTypes = null;
	ErrorReturn = null;

	scTable.initialiseTable(tblTable,0,"",PopulateTable,10,nListLength);
	if (nListLength > 0) scTable.setRowSelected(1);
}

function CountMainEmployments()
{
	var nEmpCount = 0;
	var iCount;
	var iNumberOfEmployments = EmploymentXML.ActiveTagList.length;

	for (iCount = 0; (iCount < iNumberOfEmployments); iCount++)
	{
		EmploymentXML.SelectTagListItem(iCount);
		if (EmploymentXML.GetTagText("MAINSTATUS") == "1")
			nEmpCount++;
	}
	
	return nEmpCount;
}

function PopulateTable(nStart)
{
	EmploymentXML.ActiveTag = null;
	EmploymentXML.CreateTagList("EMPLOYMENT");
	var iCount;
	var iNumberOfEmployments = EmploymentXML.ActiveTagList.length;
	
	for (iCount = 0; (iCount < iNumberOfEmployments) && (iCount < 10); iCount++)
	{
		EmploymentXML.SelectTagListItem(iCount + nStart);

		var sEmployerName = EmploymentXML.GetTagText("COMPANYNAME");
		var sEmploymentStatus = EmploymentXML.GetTagAttribute("EMPLOYMENTSTATUS", "TEXT");
		var sOccupation = EmploymentXML.GetTagText("JOBTITLE");
		var sSalary = EmploymentXML.GetTagText("YEARLYTOTALAMOUNT");
		if (sSalary == "")
			sSalary = EmploymentXML.GetTagText("AVERAGENETPROFIT");
		var sDateStarted = EmploymentXML.GetTagText("DATESTARTEDORESTABLISHED");

		var sMainEmployment = ""
		if (EmploymentXML.GetTagText("MAINSTATUS") == "1")
			sMainEmployment = "Yes";
		else
			sMainEmployment = "No";

		// Add to the search table
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sEmploymentStatus);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),sEmployerName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),sOccupation);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),sSalary);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),sDateStarted);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(5),sMainEmployment);
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount);
	}
	// SYS1947 SA 16/5/01 This is always being set to zero as the count is done after call 
	//to this function! Moved line below to after count has been done.
	//scScreenFunctions.SetContextParameter(window,"idMainEmploymentCount", nMainEmploymentCount);
}

function PopulateTotals()
{
	var sGrossIncome = "0.00";
	var sNetIncome = "0.00";
	var sGrossAllowableIncome = "0.00";
	var sNetAllowableIncome = "0.00";
	
	//	AW	18/10/2002 BMIDS00653 
	var CustomerXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var IsMemberOfStaff = 0;
	
	with (CustomerXML)
	{
		CreateRequestTag(window, null)
		CreateActiveTag("CUSTOMERVERSION");
		CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
		CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
		RunASP(document, "GetPersonalDetails.asp");
		IsResponseOK();
		
		CustomerXML.SelectTag(null,"CUSTOMERVERSION");
		if (!isNaN(parseInt(CustomerXML.GetTagText("MEMBEROFSTAFF"))))	{
			IsMemberOfStaff = parseInt(CustomerXML.GetTagText("MEMBEROFSTAFF"));
		}
	}
	
	if (EmploymentXML.SelectTag(null,"EMPLOYMENTANDINCOME") != null)
	{
		var sGrossIncome = EmploymentXML.GetTagText("TOTALGROSSEARNEDINCOME");
		var sNetIncome = EmploymentXML.GetTagText("TOTALNETMONTHLYINCOME");
	}

	if (EmploymentXML.SelectTag(null,"INCOMESUMMARYDETAILS") != null)
		if (EmploymentXML.ActiveTag.text != "")
		{
			frmScreen.cboHighestTaxRate.value = EmploymentXML.GetTagText("HIGHESTTAXRATE");
			scScreenFunctions.SetRadioGroupValue(frmScreen,"UKTaxPayerGroup",EmploymentXML.GetTagText("UKTAXPAYERINDICATOR"))
			//AW	BMIDS00653
			sGrossAllowableIncome = EmploymentXML.GetTagText("ALLOWABLEANNUALINCOME");
			sNetAllowableIncome = EmploymentXML.GetTagText("NETALLOWABLEANNUALINCOME");
			
		}

	if (sGrossIncome == "") sGrossIncome = "0.00";
	if (sNetIncome == "") sNetIncome = "0.00";
	frmScreen.txtGrossIncome.value = sGrossIncome;
	frmScreen.txtNetIncome.value = sNetIncome;
	//AW	BMIDS00653
	if (IsMemberOfStaff > 0)
	{
		scScreenFunctions.HideCollection(spnGrossAllowable);
		scScreenFunctions.HideCollection(spnNetAllowable);
	}
	else
	{
		//AW	17/11/2002	BMIDS00932
		if (sGrossAllowableIncome == "" || m_nCustomerIndex > 2 ) sGrossAllowableIncome = "0.00";
		if (sNetAllowableIncome == "" || m_nCustomerIndex > 2 ) sNetAllowableIncome = "0.00";
		//AW	17/11/2002	BMIDS00932 - End
		frmScreen.txtGrossAllowableIncome.value = sGrossAllowableIncome;
		frmScreen.txtNetAllowableIncome.value = sNetAllowableIncome;	
	}
}

function SaveTaxDetails()
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
			
	if(m_sMetaAction == "Add")
		sAction = "CREATE";
	else
		sAction = "UPDATE";
		
	XML.CreateRequestTag(window,sAction);
	XML.CreateActiveTag("INCOMESUMMARY");
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	XML.CreateTag("HIGHESTTAXRATE",frmScreen.cboHighestTaxRate.value);
	XML.CreateTag("UKTAXPAYERINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"UKTaxPayerGroup"));
	XML.CreateTag("TAXOFFICE", m_sTaxOffice);
	XML.CreateTag("TAXREFERENCENUMBER", m_sTaxReference);
	XML.CreateTag("TAXEDOUTSIDEUKINDICATOR", m_sTaxedOutsideUKInd);
	XML.CreateTag("SELFASSESS", m_sSelfAssess); <% /* MAH 20/11/2006 E2CR35 */ %>
	XML.CreateTag("NONUKTAXDETAILS", m_sNonUKTaxArrangements);
	
	XML.SelectTag(null, "REQUEST");
	XML.CreateActiveTag("CUSTOMERVERSION");
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	XML.CreateTag("NATIONALINSURANCENUMBER", m_sNINumber);

	// Save the details
	// 	XML.RunASP(document,"SaveTaxDetails.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"SaveTaxDetails.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	bSuccess = XML.IsResponseOK();
	XML = null;
	return(bSuccess);
}
		
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	//BMIDS00934 Need to check datafreeze indicator too.
	if (m_sReadOnly == "0")
		m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idFreezeDataIndicator","0");
	m_nCustomerIndex = scScreenFunctions.GetContextParameter(window,"idCustomerIndex", "1");
	//AW	24/10/2002	BMIDS00653
	m_sApplicationFactFindNumber	= scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
}

function RetrieveCustomerData()
{
	m_sCustomerNumber = "";
	m_sCustomerVersionNumber = "";
	m_sCustomerName = "";

	if ((m_nCustomerIndex > 0) && (m_nCustomerIndex < 6))
	{
		m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + m_nCustomerIndex,"1325");
		m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + m_nCustomerIndex,"1");
		m_sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + m_nCustomerIndex,"Cust");
	}
	
	<% /* BMIDS00882 Customer numbers can contain alphanumeric characters
	return(!isNaN(parseInt(m_sCustomerNumber)) & !isNaN(parseInt(m_sCustomerVersionNumber)));
	*/ %>
	if ((m_sCustomerNumber.length > 0) && (!isNaN(parseInt(m_sCustomerVersionNumber))))
		return true
	else
		return false;	
	<% /* BMIDS00882 End */ %>
}

<% /* BMIDS00654 MDC 01/11/2002 */ %>
<% /* AW	24/10/2002	BMIDS00653 */ %>
function RunIncomeCalculations()
{
    // PJO BMIDS621 Don't run income calcs in read only
  	if (m_sReadOnly =="1")
  		return(true) ;

		
	var AllowableIncXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AllowableIncXML.CreateRequestTag(window,null);
	AllowableIncXML.SetAttribute("OPERATION","CriticalDataCheck");
	<% /* MAR30 */ %> 
	
		
	// Set up request for Income Calculation
	var TagRequest = AllowableIncXML.CreateActiveTag("INCOMECALCULATION");
	AllowableIncXML.SetAttribute("CALCULATEMAXBORROWING","1");
	AllowableIncXML.SetAttribute("QUICKQUOTEMAXBORROWING","0");
	AllowableIncXML.CreateActiveTag("APPLICATION");
	AllowableIncXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AllowableIncXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	AllowableIncXML.ActiveTag = TagRequest;

	<% /* BMIDS895 add customer info */ %>
	AllowableIncXML.AddXMLBlock(m_CustomerXML.XMLDocument);
	<% /* MAR1170 Run this under Critical Data Check */ %>
	AllowableIncXML.SelectTag(null,"REQUEST");
	AllowableIncXML.CreateActiveTag("CRITICALDATACONTEXT");
	AllowableIncXML.SetAttribute("APPLICATIONNUMBER",scScreenFunctions.GetContextParameter(window, "idApplicationNumber",null));
	AllowableIncXML.SetAttribute("APPLICATIONFACTFINDNUMBER",scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber",null));
	AllowableIncXML.SetAttribute("SOURCEAPPLICATION","Omiga");
	AllowableIncXML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
	AllowableIncXML.SetAttribute("ACTIVITYINSTANCE","1");
	AllowableIncXML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
	AllowableIncXML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
	AllowableIncXML.SetAttribute("COMPONENT","omIC.IncomeCalcsBO");
	AllowableIncXML.SetAttribute("METHOD","RunIncomeCalculation");	
							
		window.status = "Critical Data Check - please wait";
		
	
	
	//switch (ScreenRules())
	//{
	//	case 1: // Warning
	//	case 0: // OK
	//				AllowableIncXML.RunASP(document,"RunIncomeCalculations.asp");
	//		break;
	//	default: // Error
	//		AllowableIncXML.SetErrorResponse();
	//}
	
	switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK					
					AllowableIncXML.RunASP(document,"OmigaTMBO.asp");
				break;
			default: // Error
				AllowableIncXML.SetErrorResponse();
			}
		
		window.status = "";	
    <% /* MAR1170  Critical Data Check Ends */ %>	
	
	if(AllowableIncXML.IsResponseOK())
	{
		return(true);
	}

	return(false);
}
<%	/*BMIDS00653  - End */ %>
<% /* BMIDS00654 MDC 01/11/2002 - End */ %>

-->
</script>
</body>
</html>



