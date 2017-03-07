<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      cm040.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description: 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		18/11/99	Initial creation
JLD		03/02/2000	Rework for performance
AY		14/02/00	Change to msgButtons button types
APS		08/03/00	Retrieve product in SubQuote via ProductCode
AY		17/03/00	MetaAction changed to ApplicationMode
JLD		28/03/00	Add view subquote fnc
AY		29/03/00	New top menu/scScreenFunctions change
AY		12/04/00	SYS0328 - Dynamic currency display
BG		17/05/00	SYS0752 Removed Tooltips
APS		30/05/00	SYS0650 - Call Calculate on OK if Sub Quote not calculated
APS		30/05/00	SYS0647 - Moved Calulate button and added error msg when user exceeds Cover Amount
MS		08/07/00	SYS0824 - Modified Tag from APPLICATION to PAYMENTPROTECTION
BG		19/02/01	SYS1945 - Fixed problem with PPSubQuoteNumber
CL		05/03/01	SYS1920 Read only functionality added
IK		22/03/01	SYS2145 Add idNoCompleteCheck
JLD		4/12/01		SYS2806 add completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		06/06/02	SYS4800 - Error following SYS4727
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
GD		15/07/02	BMIDS00045 - Capture B&C and PP Insurance Premium - Complete re-work of screen
GD		15/08/02	BMIDS00312 - Add View and New buttons to allow PP to be re-instated correctly
ASu		14/10/02	BMIDS00597 - Premium and Reference fields on PP should not be mandatory.
HMA     04/10/04    BMIDS903   - Set CallingScreenID when LTV has changed.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/
%>
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
<link href="stylesheet.css" rel="STYLESHEET" type="text/css"><title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<script src="validation.js" language="JScript"></script>
<!-- Specify Forms Here -->
<form id="frmCostModellingMenu" method="post" action="cm010.asp" STYLE="DISPLAY: none">	</form>
<form id="frmToViewSubquotes" method="post" action="cm075.asp" STYLE="DISPLAY: none">	</form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none">	</form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 103px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span id = spnProduct>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 15px<% /*ASu BMIDS00597 - ; COLOR: 'red'*/ %>" class="msgLabel">
		Monthly Premium
		<span style="LEFT: 175px; POSITION: absolute; TOP: 0px">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
			<input id="txtTotalCost" name="TotalCost" maxlength="9" style="POSITION: absolute; TOP: -3px; WIDTH: 100px" class="msgTxt">
		</span>
		
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 45px<% /*ASu BMIDS00597 - ; COLOR: 'red'*/ %>" class="msgLabel">
		Reference Number
		<span style="LEFT: 175px; POSITION: absolute; TOP: 0px">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgTxt"></label>
			<input id="txtReferenceNumber" name="ReferenceNumber" maxlength="40" style="POSITION: absolute; TOP: -3px; WIDTH: 250px" class="msgTxt">
		</span>
		
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 75px<% /*ASu BMIDS00597 - ; COLOR: 'red'*/ %>" class="msgLabel">
		Cover Type
		<span style="LEFT: 174px; POSITION: absolute; TOP: -3px">
			<select id="cboCoverType" name="CoverType" style="WIDTH: 200px" class="msgCombo"> </select>
		</span>
	</span>	
</span>
<span style="LEFT: 500px; POSITION: absolute; TOP: 40px">
	<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
		<input id="btnViewSubQuotes" value="View Sub-Quotes" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	<span style="LEFT: 0px; POSITION: absolute; TOP: 30px">
		<input id="btnCreateNewSubQuote" value="New Sub-Quote" type="button" style="WIDTH: 100px" class="msgButton">
	</span>	
</span>
</div>
</form>
<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 180px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<!-- #include FILE="attribs/cm040attribs.asp" -->
<script language="JScript">
<!--
var m_sApplicationMode = null;
var m_sReadOnly	= null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sTypeOfApplication = null;
var m_sApplicant1Eligible = null;
var m_sApplicant2Eligible = null;
var m_sPPSubQuoteNumber	= "";
var m_sMortgageSubQuoteNumber = null;
var m_sQuotationNumber = null;
var m_sDistributionChannel = null;
var scScreenFunctions;
var m_blnReadOnly = false;


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
	FW030SetTitles("Payment Protection Details","CM040",scScreenFunctions);
	scScreenFunctions.SetCurrency(window,frmScreen);
	RetrieveContextData();	
	SetMasks();
	PopulateScreen();
	SetFieldsReadOnly();
	Validation_Init();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CM040");
	
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
function RetrieveContextData()
{
	m_sApplicationMode = scScreenFunctions.GetContextParameter(window,"idApplicationMode","Quick Quote");
	m_sReadOnly	= scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber", "APP0001");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber", "1");
	m_sTypeOfApplication = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue", "10");
	m_sApplicant1Eligible = scScreenFunctions.GetContextParameter(window,"idApplicant1PPEligible", "1");
	m_sApplicant2Eligible = scScreenFunctions.GetContextParameter(window,"idApplicant2PPEligible", "1");
	m_sPPSubQuoteNumber	= scScreenFunctions.GetContextParameter(window,"idPPSubQuoteNumber", "1");
	m_sMortgageSubQuoteNumber = scScreenFunctions.GetContextParameter(window,"idMortgageSubQuoteNumber", "1");
	m_sQuotationNumber = scScreenFunctions.GetContextParameter(window,"idQuotationNumber", "1");
	m_sDistributionChannel = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId", "1");
}

function GetCoverTypeList()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("PPCoverType");
	var bSuccess = false;
	if (XML.GetComboLists(document, sGroups) == true)
	{
		bSuccess = true;
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboCoverType,"PPCoverType",false);
	}
	if (bSuccess == false)scScreenFunctions.SetScreenToReadOnly(frmScreen);
}
function SetFieldsReadOnly()
{
	var sMonthlyCost = frmScreen.txtTotalCost.value;	
	if (m_sReadOnly == "1")
	{	
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		DisableMainButton("Submit");
	}	else
	{
		scScreenFunctions.SetFocusToFirstField(frmScreen);
	}
}
function SetFieldsOnExistingQuote()
{
	//scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.cboCoverType.id);
	//scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.txtReferenceNumber.id);
	//scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.txtTotalCost.id);
}
function PopulateScreen()
{
	GetCoverTypeList();
	GetPaymentProtectionSubQuote();
	CreateFirstPaymentProtectionSubQuote();
}

function GetPaymentProtectionSubQuote()
{
	if (m_sPPSubQuoteNumber != "")
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, "SEARCH");
		XML.CreateActiveTag("PAYMENTPROTECTIONSUBQUOTE");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);	
		XML.CreateTag("PPSUBQUOTENUMBER", m_sPPSubQuoteNumber);
		XML.RunASP(document, "GetPaymentProtectionSubQuote.asp");
		if (XML.IsResponseOK() == true)
		{			
			frmScreen.cboCoverType.value = XML.GetTagText("COVERTYPE");
			frmScreen.txtTotalCost.value = XML.GetTagText("TOTALPPMONTHLYCOST");
			frmScreen.txtReferenceNumber.value = XML.GetTagText("PPREFERENCENUMBER");
			scScreenFunctions.SetScreenToReadOnly(frmScreen);			
		}
		XML = null;
	}	
}
function CreateFirstPaymentProtectionSubQuote()
{
<%	// create first payment protection sub-quote if there is not an active one
%>	
	if (m_sReadOnly != "1" && m_sPPSubQuoteNumber == "")
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, "CREATE");
		XML.CreateActiveTag("QUOTATION");
		XML.CreateTag("CONTEXT", m_sApplicationMode);
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber); 
		XML.CreateTag("QUOTATIONNUMBER", m_sQuotationNumber);
		XML.RunASP(document, "CreateFirstPPSubQuote.asp");
		if (XML.IsResponseOK() == true)
		{
			m_sPPSubQuoteNumber = XML.GetTagText("PPSUBQUOTENUMBER");
			scScreenFunctions.SetContextParameter(window, "idPPSubQuoteNumber", m_sPPSubQuoteNumber);
			frmScreen.btnCreateNewSubQuote.disabled = true
		}
		else 
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			btnSubmit.focus();
		}
		XML = null;
	}
}

function SaveScreenData()
{
	var blnReturn = true;
	if (frmScreen.onsubmit() == true)
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window, null);
		XML.CreateActiveTag("CALCULATE");
		if (m_sApplicationMode == "Cost Modelling")
		{
			//XML.CreateActiveTag("PAYMENTPROTECTION");
			XML.CreateActiveTag("PAYMENTPROTECTIONSUBQUOTE");
			//sASPFile = "CalculateAndSaveAQPPCosts.asp";
			sASPFile = "SavePPSubquote.asp"
		}
		else
		{
			XML.CreateActiveTag("QUICKQUOTE");
			sASPFile = "CalculateAndSaveQQPPCosts.asp";

		}	
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("PPSUBQUOTENUMBER", m_sPPSubQuoteNumber); 
		XML.CreateTag("CUSTOMERNUMBER1", scScreenFunctions.GetContextParameter(window, "idCustomerNumber1", "AS1111"));
		XML.CreateTag("CUSTOMERVERSIONNUMBER1",	scScreenFunctions.GetContextParameter(window, "idCustomerVersionNumber1", "1"));
		XML.CreateTag("CUSTOMERNUMBER2", scScreenFunctions.GetContextParameter(window, "idCustomerNumber2", "AS1112"));
		XML.CreateTag("CUSTOMERVERSIONNUMBER2",	scScreenFunctions.GetContextParameter(window, "idCustomerVersionNumber2", "1"));
		XML.CreateTag("DISTRIBUTIONCHANNEL", m_sDistributionChannel);
		
		XML.CreateTag("COVERTYPE", frmScreen.cboCoverType.value); 
		XML.CreateTag("TOTALPPMONTHLYCOST", frmScreen.txtTotalCost.value); 
		XML.CreateTag("PPREFERENCENUMBER", frmScreen.txtReferenceNumber.value); 


		// 		XML.RunASP(document,sASPFile);				
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,sASPFile);				
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		if (XML.IsResponseOK() != true)
		{
			blnReturn = false
		}
		else blnReturn = true;
		XML = null;
	}
	else blnReturn = false;
	
	return blnReturn;
}
function btnSubmit.onclick()
{
	<% /* ASu BMIDS00597 - Start 
	if (DoMandatoryProcessing()) */ %>
	if ((frmScreen.onsubmit() == true))
	<% /* ASu - End */ %>
	{
		if (SaveScreenData())
		{
			scScreenFunctions.SetContextParameter(window,"idNoCompleteCheck","1");
			if(scScreenFunctions.CompletenessCheckRouting(window))
				frmToGN300.submit();
			else
				frmCostModellingMenu.submit();
				
			<% /* BMIDS903 Set the CallingScreenID context parameter for use by CM100 */ %>
			if (scScreenFunctions.GetContextParameter(window,"idLTVChanged") == '1')
			{
				scScreenFunctions.SetContextParameter(window, "idCallingScreenID", "CM040");
			}					
		}
	}
}
function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmCostModellingMenu.submit();
}
<% /* ASu BMIDS00597 - Start 
function DoMandatoryProcessing()
{
	var iIndex;
	var iFirstEmpty;
	var iLength = frmScreen.length;

	var blnReturn = true;
	for (iIndex = 0;((iIndex < iLength) && blnReturn); iIndex++)
	{
		if (frmScreen.item(iIndex).value == "")
		{
			blnReturn = false;
			iFirstEmpty = iIndex;
		}
	}
	if (blnReturn == false)
	{
		alert("Please enter a value.");
		frmScreen.item(iFirstEmpty).focus();
	}
	return(blnReturn);
}
<% /* ASu - End */ %>

function frmScreen.btnViewSubQuotes.onclick()
{
	<%  // Before we leave this screen we need to save the screen data	%>	
	//if (m_sPPSubQuoteNumber=="")	
	//{
		if(confirm("Viewing the sub-quotes will cause this screen data to be saved. Do you wish to continue?"))
		{
		<% /* ASu BMIDS00597 - Start 
		if (DoMandatoryProcessing()) */ %>
		if ((frmScreen.onsubmit() == true))
		<% /* ASu - End */ %>
			{
				if (SaveScreenData())
				{
						scScreenFunctions.SetContextParameter(window, "idBusinessType", "PP");
						scScreenFunctions.SetContextParameter(window, "idMetaAction", "CM040");
						frmToViewSubquotes.submit();
				}
			}	
		}	
	//} else
	//{
		//scScreenFunctions.SetContextParameter(window,"idBusinessType","BC");
		//scScreenFunctions.SetContextParameter(window,"idMetaAction","CM020");
		//frmViewSubQuotes.submit();
	//}
}
function frmScreen.btnCreateNewSubQuote.onclick()
{
<%	/* Creates a new payment protection sub-quote based on an active sub quote */ %>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "CREATE");
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("CONTEXT", m_sApplicationMode);
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber); 
	XML.CreateTag("QUOTATIONNUMBER", m_sQuotationNumber); 
	XML.CreateTag("PPSUBQUOTENUMBER", m_sPPSubQuoteNumber); 
	// 	XML.RunASP(document, "CreateNewPPSubQuote.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "CreateNewPPSubQuote.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if (XML.IsResponseOK() == true)
	{
		m_sPPSubQuoteNumber = XML.GetTagText("PPSUBQUOTENUMBER");
		scScreenFunctions.SetContextParameter(window, "idPPSubQuoteNumber", m_sPPSubQuoteNumber);
		//frmScreen.txtMonthlyCost.value = "";
		
		//if (IsApplicant1Eligible()==true && IsApplicant2Eligible()==true)
		//{
			//scScreenFunctions.SetRadioGroupToWritable(frmScreen, "SplitInd");
			//if (scScreenFunctions.GetRadioGroupValue(frmScreen, "SplitInd") == frmScreen.optSplitCoverYes.value)
			//{
			//	scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.txtApplicantOneMortgageCover.id);				
			//}
		//}		
		//scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.cboProduct.id);
		//scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.cboCoverType.id);
		//scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.txtMortgageCover.id);		
		//frmScreen.btnCalculate.disabled	= false;
		frmScreen.btnCreateNewSubQuote.disabled = true;
				//EMPTY OUT ALL THE FIELDS
		//frmScreen.txtTotalCost.value = "";
		//frmScreen.txtReferenceNumber.value = "";
		frmScreen.cboCoverType.selectedIndex = 0; //Default to First in list
		//frmScreen.txtReason.value = "";
		//scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTotalCost");
		//scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtStatus");
		scScreenFunctions.SetCollectionToWritable(spnProduct);
		//scScreenFunctions.SetCollectionToWritable(spnPayment);
		frmScreen.btnCreateNewSubQuote.disabled = true;	
		//frmScreen.btnCalculate.disabled = false;
		m_sQuoteFrozen = "0";
		frmScreen.txtTotalCost.focus();
	}
	XML = null;
}

-->
</script>
</body>
</html>
<% /* OMIGA BUILD VERSION 046.02.08.08.00 */ %>


