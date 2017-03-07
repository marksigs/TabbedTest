<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      DC192.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Net Profit Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		18/02/2000	Created
AD		22/03/2000	SYS0458 - Extra validation added
AY		31/03/00	New top menu/scScreenFunctions change
IVW		11/04/2000	Changed Prev/Next/Cancel to Ok/Cancel
AY		12/04/00	SYS0328 - Dynamic currency display
MH      03/05/00    SYS0458 - Improved validation on years
BG		17/05/00	SYS0752 Removed Tooltips
MC		30/05/00	SYS0756 If Read-only mode, disabled fields
BG		30/05/00	SYS0765	Made Value fields max characters 8
MC		09/06/00	SYS0866 Standardise format of balance and payment fields
CL		05/03/01	SYS1920 Read only functionality added
DJP		06/08/01	SYS2564(parent) SYS3808(child) Client cosmetic customisation
SA		23/04/02	SYS4437 Net Monthly Income length increased from 5 to 6
SA		25/04/02	SYS4437 Net Monthly Income length increased from 6-10 to be inline with Long datatype.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
ASu		05/09/2002	BMIDS00147 - Increase net profit amount fields to 7 characters
MV		11/10/2002	BMIDS00449 - Decreased the max length to 8 
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM History:

Prog	Date		Description
TW	02/02/2007  	EP2_1046 Cancel button does not work (Syntax error in btnCancel.onclick()).
LDM		07/02/2007	EP2_1252 Make sure the on cancell btn works. Related to EP2_1046.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

/* %>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
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
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<% /* FORMS */ %>
<form id="frmToDC160" method="post" action="DC160.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC182" method="post" action="DC182.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC183" method="post" action="DC183.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
<div style="HEIGHT: 160px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Net Profit Year Ending
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtYear1" maxlength="4" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span> 
	</span>

	<span id="spnYear1Amount" style="LEFT: 200px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Amount
		<span style="LEFT: 60px; POSITION: absolute; TOP: 0px">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtYear1Amount" maxlength="7" style="POSITION: absolute; WIDTH: 80px; TOP: -3px" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Net Profit Year Ending
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtYear2" maxlength="4" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span> 
	</span>

	<span id="spnYear2Amount" style="LEFT: 200px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Amount
		<span style="LEFT: 60px; POSITION: absolute; TOP: 0px">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtYear2Amount" maxlength="7" style="POSITION: absolute; WIDTH: 80px; TOP: -3px" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 58px" class="msgLabel">
		Net Profit Year Ending
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtYear3" maxlength="4" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span> 
	</span>

	<span id="spnYear3Amount" style="LEFT: 200px; POSITION: absolute; TOP: 58px" class="msgLabel">
		Amount
		<span style="LEFT: 60px; POSITION: absolute; TOP: 0px">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtYear3Amount" maxlength="7" style="POSITION: absolute; WIDTH: 80px; TOP: -3px" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 84px" class="msgLabel">
		<LABEL id="idAccountsSeen"></LABEL>
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="optEvidenceOfAccountsYes" name="EvidenceOfAccountsGroup" type="radio" value="1">
			<label for="optEvidenceOfAccountsYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px">
			<input id="optEvidenceOfAccountsNo" name="EvidenceOfAccountsGroup" type="radio" value="0">
			<label for="optEvidenceOfAccountsNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 110px" class="msgLabel">
		Average Net Profit
		<span style="LEFT: 130px; POSITION: absolute; TOP: 0px">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtAverageNetProfit" maxlength="7" style="POSITION: absolute; WIDTH: 80px; TOP: -3px" class="msgReadOnly" readonly tabindex=-1>
		</span> 
	</span>

	<span style="LEFT: 220px; POSITION: absolute; TOP: 104px">
		<input id="btnCalculate" value="Calculate" type="button" style="WIDTH: 60px" class="msgButton"> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 134px" class="msgLabel">
		Net Monthly Income
		<span style="LEFT: 130px; POSITION: absolute; TOP: 0px">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtNetMonthlyIncome" maxlength="8" style="POSITION: absolute; WIDTH: 80px; TOP: -3px" class="msgTxt">
		</span> 
	</span>
</div>

</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 440px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/DC192attribs.asp" -->
<!-- #include FILE="Customise/DC192Customise.asp" -->
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;
	
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sEmploymentSequenceNuymber = "";
var m_blnReadOnly = false;


/* EVENTS */

function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Edit");
	var sEmploymentStatusType = scScreenFunctions.GetContextParameter(window,"idEmploymentStatusType","");
	if (sEmploymentStatusType == "C")
		frmToDC183.submit();
	else if (sEmploymentStatusType == "S")
		frmToDC182.submit();
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		frmToDC160.submit();
	}
}

function frmScreen.btnCalculate.onclick()
{
	var dAverageNetProfit = 0;
	var dTotalNetProfit = 0;
	var nNetProfitCount = 0;
	var dAmount = 0;

	if (ValidateAmounts())
	{
		<%/* Year 1 */%>
		dAmount = parseFloat(frmScreen.txtYear1Amount.value);
		if (dAmount > 0)
		{
			nNetProfitCount++;
			dTotalNetProfit = dTotalNetProfit + dAmount;
		}

		<%/* Year 2 */%>
		dAmount = parseFloat(frmScreen.txtYear2Amount.value);
		if (dAmount > 0)
		{
			nNetProfitCount++;
			dTotalNetProfit = dTotalNetProfit + dAmount;
		}

		<%/* Year 3 */%>
		dAmount = parseFloat(frmScreen.txtYear3Amount.value);
		if (dAmount > 0)
		{
			nNetProfitCount++;
			dTotalNetProfit = dTotalNetProfit + dAmount;
		}

		if (nNetProfitCount > 0)
			dAverageNetProfit = dTotalNetProfit / nNetProfitCount;
		else
			dAverageNetProfit = 0;

		<% /* frmScreen.txtAverageNetProfit.value = top.frames[1].document.all.scMathFunctions.RoundValue(dAverageNetProfit,2); */ %>
		frmScreen.txtAverageNetProfit.value = top.frames[1].document.all.scMathFunctions.Truncate(dAverageNetProfit);
	}
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
	FW030SetTitles("Net Profit Details","DC192",scScreenFunctions);

	scScreenFunctions.SetCurrency(window,frmScreen);
	RetrieveContextData();
	SetMasks();

	Validation_Init();	
	Initialise();

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC192");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	if(m_sReadOnly == "1")
	{		
		frmScreen.btnCalculate.disabled = true;
	}
	scScreenFunctions.SetFocusToFirstField(frmScreen);	

	// DJP SYS2564/SYS3808 for client customisation
	Customise();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

function CommitChanges()
{
	var bSuccess = true;
	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
			if (IsChanged())
			{
				if (ValidateAmounts())
					bSuccess = SaveNetProfitDetails();
				else
					bSuccess=false;
			}

	return(bSuccess);
}

function DefaultFields()
{
	with (frmScreen)
	{
		txtAverageNetProfit.value = "0.00";
		txtNetMonthlyIncome.value = "0";
		txtYear1.value = "";
		txtYear1Amount.value = "";
		txtYear2.value = "";
		txtYear2Amount.value = "";
		txtYear3.value = "";
		txtYear3Amount.value = "";
	}
	scScreenFunctions.SetRadioGroupValue(frmScreen,"EvidenceOfAccountsGroup","0");
}

function Initialise()
// Initialises the screen
{
	PopulateScreen();

	if(m_sMetaAction == "Add")
		DefaultFields();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateScreen()
// Populates the screen with details of the item selected in 
{
	NetProfitXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	NetProfitXML.CreateRequestTag(window,null);
	NetProfitXML.CreateActiveTag("NETPROFIT");
	NetProfitXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	NetProfitXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	NetProfitXML.CreateTag("EMPLOYMENTSEQUENCENUMBER", m_sEmploymentSequenceNumber);

	NetProfitXML.RunASP(document,"GetNetProfitDetails.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = NetProfitXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[1] == ErrorTypes[0]))
	{
		//Error: record not found
		m_sMetaAction = "Add";
	}
	else if (ErrorReturn[0] == true)
	{
		// No error
		m_sMetaAction = "Edit";

		if(NetProfitXML.SelectTag(null, "NETPROFIT") != null)
			with (frmScreen)
			{
				txtNetMonthlyIncome.value = NetProfitXML.GetTagText("NETMONTHLYINCOME");
				txtYear1.value = NetProfitXML.GetTagText("YEAR1");
				txtYear1Amount.value = NetProfitXML.GetTagText("YEAR1AMOUNT");
				txtYear2.value = NetProfitXML.GetTagText("YEAR2");
				txtYear2Amount.value = NetProfitXML.GetTagText("YEAR2AMOUNT");
				txtYear3.value = NetProfitXML.GetTagText("YEAR3");
				txtYear3Amount.value = NetProfitXML.GetTagText("YEAR3AMOUNT");
				scScreenFunctions.SetRadioGroupValue(frmScreen,"EvidenceOfAccountsGroup",NetProfitXML.GetTagText("EVIDENCEOFACCOUNTS"));
			}
	}

	frmScreen.btnCalculate.onclick();

	ErrorTypes = null;
	ErrorReturn = null;
}

function RetrieveContextData()
{
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","1325");
	m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","1");
	m_sEmploymentSequenceNumber = scScreenFunctions.GetContextParameter(window,"idEmploymentSequenceNumber","1");    
}

function SaveNetProfitDetails()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateActiveTag("NETPROFIT");
	XML.CreateTag("NETMONTHLYINCOME",frmScreen.txtNetMonthlyIncome.value);
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	XML.CreateTag("EMPLOYMENTSEQUENCENUMBER",m_sEmploymentSequenceNumber);
	XML.CreateTag("YEAR1",frmScreen.txtYear1.value);
	XML.CreateTag("YEAR1AMOUNT",frmScreen.txtYear1Amount.value);
	XML.CreateTag("YEAR2",frmScreen.txtYear2.value);
	XML.CreateTag("YEAR2AMOUNT",frmScreen.txtYear2Amount.value);
	XML.CreateTag("YEAR3",frmScreen.txtYear3.value);
	XML.CreateTag("YEAR3AMOUNT",frmScreen.txtYear3Amount.value);
	XML.CreateTag("EVIDENCEOFACCOUNTS",scScreenFunctions.GetRadioGroupValue(frmScreen,"EvidenceOfAccountsGroup"));

	<%/* Save the details */%>
	// 	XML.RunASP(document,"SaveNetProfitDetails.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"SaveNetProfitDetails.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}

function ValidateAmount(ctrYear, ctrAmount)
{
	var sAmountInvalid = "A valid amount must be entered if a year has been specified.";
	var sYearInvalid = "A valid year must be entered if an amount has been specified.";
	var iYear;
	var thisYear;
	
	if ((ctrYear.value == "") && (ctrAmount.value == "")) return true;

	if ((ctrYear.value != "") && (ctrAmount.value == ""))
	{
		alert(sAmountInvalid);
		ctrAmount.focus();
		return(false);
	}

	if ((ctrYear.value == "") && (ctrAmount.value != ""))
	{
		alert(sYearInvalid);
		ctrYear.focus();
		return(false);
	}

	<% /* MO - BMIDS00807 */ %>
	var dtTempCurrentDate = scScreenFunctions.GetAppServerDate(true);
	<% /* var dtTempCurrentDate	= new Date(); */ %>

	iYear=new Number(ctrYear.value);
	thisYear=dtTempCurrentDate.getYear();
	
	if (thisYear < iYear)
	{
		alert ("Net Profit Year Ending cannot be in the future");
		ctrYear.focus();
		return(false);
	}
	
	if ( iYear < thisYear - 10)
	{
		alert ("Net Profit Year Ending cannot be in the distant past");
		ctrYear.focus();
		return(false);
	}

	return(true); 
}

function CompareYears(refYear1,refYear2)
{
	if ((refYear1.value=="") || (refYear2.value=="")) return true;
	
	if (refYear1.value==refYear2.value)
	{
		alert("Years must all be different");
		refYear2.focus();
		return false;
	}
	else 
		return true;
	
}
function ValidateAmounts()
{
	if (ValidateAmount(frmScreen.txtYear1, frmScreen.txtYear1Amount))
		if (ValidateAmount(frmScreen.txtYear2, frmScreen.txtYear2Amount))
			if (ValidateAmount(frmScreen.txtYear3, frmScreen.txtYear3Amount))
				if (CompareYears(frmScreen.txtYear1,frmScreen.txtYear2))
					if (CompareYears(frmScreen.txtYear2,frmScreen.txtYear3))
						if (CompareYears(frmScreen.txtYear1,frmScreen.txtYear3))
							return true;
	return false;
}

-->
</script>
</body>
</html>



