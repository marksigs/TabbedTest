<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/*
Workfile:      DC190.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Income Breakdown
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		15/02/2000	Created
IW		22/03/2000  SYS0527
AY		31/03/00	New top menu/scScreenFunctions change
MH      15/05/00    SYS0741 field validation
BG		17/05/00	SYS0752 Removed Tooltips
MC		23/05/00	SYS0756 If Read-only mode, disabled all fields
MC		09/06/00	SYS0866 Standardise format of balance and payment fields
CL		05/03/01	SYS1920 Read only functionality added
JLD		10/12/01	SYS2806 completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

BMIDS Specific History:

Prog	Date		Description
DPF		02/07/02	BMIDS00149 added IsNumeric function to allow a check to be made
					to ensure values entered are numeric (not blank) and checks within
					CalculateTotals, PopulateTable, SpnTable_OnClick
MV		22/08/2002	BMIDS00355	IE 5.5 Upgrade - Modified the Msgbuttons Position 
ASu		05/09/02	BMIDS00147 - Increase income amount fields to 7 Char
TW      09/10/2002  Modified to incorporate client validation - SYS5115
MV		07/11/2002	BMIDS00833	Amended InitialiseTableXML
AW		04/12/02	BM00152		Do not route to GN300	
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History: 

Prog	Date		AQR			Description
MF		03/08/2005	MAR20		WP06 clicking Cancel now returns to DC170 Employment Details
MF		14/09/05	MAR30		Added new fields for Other income
HMA     08/10/2005  MAR135		Do not update Customer Version when saving Tax Details.
                                Improve checking for Income Summary details.
MV		18/10/2005	MAR174		Amended CommitChanges()    
SD		14/11/2005	MAR258		Critical Data Check changes 
PSC		25/01/2006	MAR1123		Use scroll table to allow extra items
DRC     31/01/2006  MAR1170     Move Critical Data Check to DC160
JD		25/05/2006	MAR1821		Check correct tag when looking for incomesummary details, Make sure data is saved correctly.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<HEAD> 
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>
<span style="TOP: 254px; LEFT: 305px; POSITION: absolute">
	<OBJECT data=scTableListScroll.asp id=scScrollTable name=scTable style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></OBJECT>
</span>
<%/* FORMS */%>
<form id="frmToDC160" method="post" action="DC160.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC181" method="post" action="DC181.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC170" method="post" action="DC170.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" validate ="onchange" mark>
<div id="divBackground" style="HEIGHT: 308px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Employer Name
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtEmployerName" maxlength="50" style="POSITION: absolute; WIDTH: 200px" class="msgReadOnly" readonly tabindex=-1>
		</span> 
	</span>

	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 34px" onclick="spnTable.onclick()">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="33%" class="TableHead">Income Type&nbsp;</td>
								<td width="33%" class="TableHead">Amount&nbsp;</td>
								<td width="33%" class="TableHead">Frequency&nbsp;</td></tr>
			<tr id="row01">		<td width="33%" class="TableTopLeft">&nbsp;</td>	<td width="33%" class="TableTopCenter">	&nbsp;</td>		<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="33%" class="TableBottomLeft">&nbsp;</td>	<td width="33%" class="TableBottomCenter">&nbsp;</td>	<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>

	<div id="divBackground2" style="HEIGHT: 42px; LEFT: 4px; POSITION: absolute; TOP: 214px; WIDTH: 596px" class="msgGroupLight">
		<span id="spnToFrequency" tabindex="0"></span>
		<span style="LEFT: 8px; POSITION: absolute; TOP: 16px; WIDTH: 200px" class="msgLabel">
			Amount
			<span style="LEFT: 60px; POSITION: absolute; TOP: 0px">
				<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtAmount" maxlength="7" style="TOP: -3px; WIDTH: 100px; POSITION: absolute" class="msgTxt">
			</span> 
		</span>

		<span style="LEFT: 190px; POSITION: absolute; TOP: 16px; WIDTH: 200px" class="msgLabel">
			Frequency
			<span style="LEFT: 60px; WIDTH: 100px; POSITION: absolute; TOP: -3px">
				<select id="cboFrequency" style="WIDTH: 100px" class="msgCombo">
				</select>
			</span> 
		</span>
		<span id="spnToAmount" tabindex="0"></span>
	</div>

	<span style="LEFT: 12px; POSITION: absolute; TOP: 274px" class="msgLabel">
		Total Gross Income
		<span style="LEFT: 150px; POSITION: absolute; TOP: 0px">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtGrossIncome" name="GrossIncome" style="TOP: -3px; WIDTH: 100px; POSITION: absolute" class="msgReadOnly" readonly tabindex=-1>
			</input>
		</span> 
	</span>

	<span style="LEFT: 12px; POSITION: absolute; TOP: 298px" class="msgLabel">
		Total Net Monthly Income
		<span style="LEFT: 150px; POSITION: absolute; TOP: 0px">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtNetIncome" maxlength="7" name="NetIncome" style="TOP: -3px; WIDTH: 100px; POSITION: absolute" class="msgTxt">
			</input>
		</span> 
	</span>
	
	<span style="LEFT: 330px; POSITION: absolute; TOP: 274px" class="msgLabel">
		Include Other Income
		<span style="LEFT: 150px; POSITION: absolute; TOP: 0px">			
			<span style="LEFT: 0px; POSITION: relative; TOP: -5px">
				<input id="optIncOtherIncomeYes" disabled name="OtherIncomeGroup" type="radio" value="1"><label for="optIncOtherIncomeYes" class="msgLabel">Yes</label> 
			</span>
			<span style="LEFT: 60px; POSITION: absolute;  TOP: -5px">
				<input id="optIncOtherIncomeNo" disabled name="OtherIncomeGroup" type="radio" value="0"><label for="optIncOtherIncomeNo" class="msgLabel">No</label> 
			</span> 	
		</span> 
	</span>

	<span style="LEFT: 330px; POSITION: absolute; TOP: 298px" class="msgLabel">
		Other Income Percentage
		<span style="LEFT: 150px; POSITION: absolute; TOP: 0px">			
			<input id="txtOtherIncomePercentage" disabled maxlength="3" style="TOP: -3px; WIDTH: 100px; POSITION: absolute" class="msgTxt">
			<label style="LEFT: 110px; POSITION: absolute; TOP: 0px" class="msgLabel">%</label>
		</span> 
	</span>	
	
</div> 
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/DC190attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sEmploymentSequenceNumber = "";
var m_sEmployerName = "";
var m_nIncomeTypeCount = 0;
var IncomeXML = null;
var TableXML = null;
var scScreenFunctions;
var m_blnReadOnly = false;
var m_bIncomeSummaryExists = false;
var m_nTableLength = 10;				<% /* PSC 25/01/2006 MAR1123 */ %>

/* EVENTS */

function btnCancel.onclick()
{
	//clear the contexts
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "Edit");
		<% /* MF MARS20 //frmToDC181.submit(); */ %>
		frmToDC170.submit();
	}
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		// AW	04/12/02	BM00152 -	Start
		//clear the contexts
		//if(scScreenFunctions.CompletenessCheckRouting(window))
		//	frmToGN300.submit();
		//else
		//{
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		frmToDC160.submit();
		//}
		// AW	04/12/02	BM00152 -	End
	}
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
	CalculateTotals();
}

function spnToAmount.onfocus()
{
	var nRowSelected = scTable.getRowSelected();
	if ((nRowSelected + scTable.getOffset() >= m_nIncomeTypeCount) | (nRowSelected == 10))
		frmScreen.txtNetIncome.focus();
	else
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

function spnTable.onclick()
{
	GetTableItemBySelection();
	var dAmount = parseFloat(TableXML.GetTagText("EARNEDINCOMEAMOUNT"));
	//BMIDS00149 - DPF 2/7/02 - check if the value is numeric, if not assign it to '0'
	if (!isNumeric(dAmount)) 
	{
		dAmount = 0;
	}
	
	var nFrequency = TableXML.GetTagText("PAYMENTFREQUENCYTYPE");

	if (nFrequency == "0") nFrequency = "";

	frmScreen.txtAmount.value = dAmount;
	frmScreen.cboFrequency.value = nFrequency;
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
	FW030SetTitles("Income Summary","DC190",scScreenFunctions);

	scScreenFunctions.SetCurrency(window,frmScreen);
	RetrieveContextData();
	Initialise();
	SetMasks();
	Validation_Init();

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC190");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	if(m_sReadOnly == "1")
	{
		scScreenFunctions.SetCollectionToReadOnly(divBackground2);
		frmScreen.txtNetIncome.disabled = true;
	}
	else
		scScreenFunctions.SetCollectionToWritable(divBackground2);

	scScreenFunctions.SetFocusToFirstField(frmScreen);	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

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
		nSingleGrossIncome = parseFloat(TableXML.GetTagText("EARNEDINCOMEAMOUNT")) *
							 parseInt(TableXML.GetTagText("PAYMENTFREQUENCYTYPE"));
		//BMIDS00149 - DPF 2/7/02 - check is numeric value before appending to total
		if (isNumeric(nSingleGrossIncome)) 
		{
			nTotalGrossIncome = nTotalGrossIncome + nSingleGrossIncome;
		}
	}
	
	frmScreen.txtGrossIncome.value = nTotalGrossIncome;
}

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
	{
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
			{
				<% /* MAR30 Save Tax details */ %>
				bSuccess = SaveOtherIncomeDetails();
				if(bSuccess){
					bSuccess = SaveEarnedIncome();
				}
			}
		}
		else
		{
			bSuccess = false
		}
	}
	return(bSuccess);
}

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
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
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

function Initialise()
{
	PopulateCombos();
	InitialiseTableXML();
	DefaultFields();
	frmScreen.txtEmployerName.value = m_sEmployerName;
	PopulateScreen();
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
		TableXML.CreateTag("EARNEDINCOMEAMOUNT","0");
		TableXML.CreateTag("PAYMENTFREQUENCYTYPE","1");
		TableXML.SetAttributeOnTag("PAYMENTFREQUENCYTYPE","TEXT",sMonthly);
		TableXML.CreateTag("EARNEDINCOMESEQUENCENUMBER","");
	}

	m_nIncomeTypeCount = iListLength;
}

function PopulateCombos()
{
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	<%/* Payment frequency */%>
	var sGroupList = new Array("EmploymentPaymentFreq");
	if(XML.GetComboLists(document,sGroupList))
		blnSuccess = blnSuccess & XML.PopulateCombo(document,frmScreen.cboFrequency,"EmploymentPaymentFreq",false);

	<%/* Income type */%>
	TableXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	sGroupList = new Array("IncomeType");
	if(TableXML.GetComboLists(document,sGroupList))
		blnSuccess = blnSuccess & (TableXML != null)

	if(!blnSuccess) scScreenFunctions.SetScreenToReadOnly(frmScreen);
	XML = null;
}

function LoadIncomeSummary()
{

	var EmploymentXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	EmploymentXML.CreateRequestTag(window,null);
	var tagCustomerList = EmploymentXML.CreateActiveTag("CUSTOMERLIST");
	EmploymentXML.CreateActiveTag("CUSTOMER");
	EmploymentXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	EmploymentXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	
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
		<% /* Error: record not found */ %>
		m_bIncomeSummaryExists=false;	
	} 
	else if (ErrorReturn[0] == true)
	{
		<% /* No error */ %>
		<% /*  MAR1821 look in IncomeSummaryDetails
		if (EmploymentXML.SelectTag(null,"INCOMESUMMARY") != null)
		{
			if (EmploymentXML.ActiveTag.text != "")
			{
				m_bIncomeSummaryExists=true;	
			}
		}	*/ %>

		if (EmploymentXML.SelectTag(null,"INCOMESUMMARYDETAILS") != null)
		{
			if (EmploymentXML.ActiveTag.text != "")
			{
				m_bIncomeSummaryExists=true; //MAR1821
				scScreenFunctions.SetRadioGroupValue(frmScreen,"OtherIncomeGroup",EmploymentXML.GetTagText("UNDERWRITEROVERRIDEINCLUDEOTHERINC"));
				if(frmScreen.optIncOtherIncomeYes.checked)
				{
					frmScreen.optIncOtherIncomeYes.click();
				}
				frmScreen.txtOtherIncomePercentage.value=EmploymentXML.GetTagText("UNDERWRITEROTHERINCOMEPERCENTAGE");	
			}
		}
	}
}


function PopulateScreen()
{
	IncomeXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	IncomeXML.CreateRequestTag(window,null);
	IncomeXML.CreateActiveTag("EARNEDINCOME");
	IncomeXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	IncomeXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	IncomeXML.CreateTag("EMPLOYMENTSEQUENCENUMBER", m_sEmploymentSequenceNumber);
	IncomeXML.RunASP(document,"FindEarnedIncomeList.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = IncomeXML.CheckResponse(ErrorTypes);
	var nRecordCount = 0;
	if (ErrorReturn[0] == true)
	{
		// No error
		nRecordCount = PopulateTableXML();
		IncomeXML.SelectTag(null,"EARNEDINCOMELIST");
		frmScreen.txtNetIncome.value = IncomeXML.GetTagText("NETMONTHLYINCOME");
	}

	<% /* MAR30 load income summary details */ %>
	LoadIncomeSummary();

	CalculateTotals();

	ErrorTypes = null;
	ErrorReturn = null;

	<% /* PSC 25/01/2006 MAR1123 - Start */ %>
	scTable.initialiseTable(tblTable,0,"",PopulateTable,m_nTableLength,m_nIncomeTypeCount);
	PopulateTable(0)
	<% /* PSC 25/01/2006 MAR1123 - End */ %>
	
	if (m_nIncomeTypeCount > nRecordCount)
		<%/* Not all income types had records on the database for this employment, so force a 
		save of the default values for the missing income types */%>
		FlagChange(true);

	if (m_nIncomeTypeCount > 0)
	{
		scTable.setRowSelected(1);
		spnTable.onclick();
	}
}

<% /* PSC 25/01/2006 MAR1123 */ %>
function PopulateTable(nStart)
{
	TableXML.ActiveTag = null;
	TableXML.CreateTagList("LISTENTRY");
	var iCount;
	varsAmount = 0;

	<% /* PSC 25/01/2006 MAR1123 */ %>
	for (iCount = 0; iCount < TableXML.ActiveTagList.length && iCount <  m_nTableLength; iCount++)
	{
		<% /* PSC 25/01/2006 MAR1123 */ %>
		TableXML.SelectTagListItem(iCount + nStart);

		var sIncomeType = TableXML.GetTagText("VALUENAME");
		//BMIDS00149 - DPF 2/7/02 - check is numeric value before writing to screen
		if (isNumeric(TableXML.GetTagText("EARNEDINCOMEAMOUNT")))
		{
			var sAmount = TableXML.GetTagText("EARNEDINCOMEAMOUNT");
		}
		var sFrequency = TableXML.GetTagAttribute("PAYMENTFREQUENCYTYPE","TEXT");

		// Add to the search table
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sIncomeType);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),sAmount);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),sFrequency);
		<% /* PSC 25/01/2006 MAR1123 */ %>
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}
}

function PopulateTableXML()
{
	<%/* Populate the TableXML with the values retrieved in IncomeXML */%>
	IncomeXML.ActiveTag = null;
	IncomeXML.CreateTagList("EARNEDINCOME");
	var iCount;
	var iListLength = IncomeXML.ActiveTagList.length;

	for (iCount = 0; iCount < iListLength; iCount++)
	{
		IncomeXML.SelectTagListItem(iCount);

		GetTableItemByIncomeType(IncomeXML.GetTagText("EARNEDINCOMETYPE"));
		TableXML.SetTagText("EARNEDINCOMEAMOUNT",IncomeXML.GetTagText("EARNEDINCOMEAMOUNT"));
		TableXML.SetTagText("PAYMENTFREQUENCYTYPE",IncomeXML.GetTagText("PAYMENTFREQUENCYTYPE"));
		TableXML.SetAttributeOnTag("PAYMENTFREQUENCYTYPE","TEXT",IncomeXML.GetTagAttribute("PAYMENTFREQUENCYTYPE","TEXT"));
		TableXML.SetTagText("EARNEDINCOMESEQUENCENUMBER",IncomeXML.GetTagText("EARNEDINCOMESEQUENCENUMBER"));
	}

	return(iListLength);
}
		
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","B00003743");
	m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","1");  
	m_sEmploymentSequenceNumber = scScreenFunctions.GetContextParameter(window,"idEmploymentSequenceNumber", "1");
	m_sEmployerName = scScreenFunctions.GetContextParameter(window,"idEmployerName", "");
}

function SaveEarnedIncome()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var TagRequestType = XML.CreateRequestTag(window, null);


	XML.CreateActiveTag("EARNEDINCOMELIST");
	XML.CreateTag("NETMONTHLYINCOME",frmScreen.txtNetIncome.value);

	<%/* Iterate through each item in the Table XML and add it to the XML to be saved */%>
	TableXML.ActiveTag = null;
	TableXML.CreateTagList("LISTENTRY");
	for (iCount = 0; (iCount < TableXML.ActiveTagList.length); iCount++)
	{	
		TableXML.SelectTagListItem(iCount);

		XML.CreateActiveTag("EARNEDINCOME");
		XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
		XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
		XML.CreateTag("EMPLOYMENTSEQUENCENUMBER",m_sEmploymentSequenceNumber);
		XML.CreateTag("EARNEDINCOMESEQUENCENUMBER",TableXML.GetTagText("EARNEDINCOMESEQUENCENUMBER"));
		XML.CreateTag("EARNEDINCOMEAMOUNT",TableXML.GetTagText("EARNEDINCOMEAMOUNT"));
		XML.CreateTag("EARNEDINCOMETYPE",TableXML.GetTagText("VALUEID"));
		XML.CreateTag("PAYMENTFREQUENCYTYPE",TableXML.GetTagText("PAYMENTFREQUENCYTYPE"));
	}

	if (iCount > 0)
	{
		
							
	<% /* MAR1170  Critical Data Check Moved to DC160 */ %>	
		
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK					
					XML.RunASP(document,"SaveEarnedIncome.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}
		
			
			
		bSuccess = XML.IsResponseOK();
	}

	XML = null;
	return(bSuccess);
}

function UpdateTableXML(nRowSelected)
{
	GetTableItemBySelection();

	var sAmount = frmScreen.txtAmount.value;
	var sFrequency = frmScreen.cboFrequency.options(frmScreen.cboFrequency.selectedIndex).text;
	var nFrequency = frmScreen.cboFrequency.value;

	TableXML.SetTagText("EARNEDINCOMEAMOUNT",sAmount);
	TableXML.SetTagText("PAYMENTFREQUENCYTYPE",nFrequency);
	TableXML.SetAttributeOnTag("PAYMENTFREQUENCYTYPE","TEXT",sFrequency);

	// Update the table itself
	scScreenFunctions.SizeTextToField(tblTable.rows(nRowSelected).cells(1),sAmount);
	scScreenFunctions.SizeTextToField(tblTable.rows(nRowSelected).cells(2),sFrequency);
}

//BMIDS00149 - DPF 2/7/02 - function added
function isNumeric(n) 
{ 
	return !isNaN(parseInt(n)); 
}
-->
</script>
</body>
</html>


