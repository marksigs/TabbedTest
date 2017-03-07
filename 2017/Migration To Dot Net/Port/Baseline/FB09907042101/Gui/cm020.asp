<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:		cm020.asp
Copyright:		Copyright © 2000 Marlborough Stirling

Description:	This screen diaplys the details of Buildings & Contents 
				sub-Quotes and allows the User to amend the details and 
				create new sub-Quotes.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		25/11/99	Initial Creation
AY		14/02/00	Change to msgButtons button types
APS		14/03/00	SYS0493
AP		30/03/00	Review of HTML layout
APS		04/05/00	Changed HTML layout in line with SYS0651
BG		17/05/00	SYS0752 Removed Tooltips
APS		31/05/00	SYS0650 OK Processing changed to remain on screen on calculate
MS		07/06/00	spelling correction indivdually to individually
CL		05/03/01	SYS1920 Read only functionality added
IK		22/03/01	SYS2145 Add idNoCompleteCheck
JLD		4/12/01		SYS2806 add completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		06/06/02	SYS4800 - Error following SYS4727

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
GD		15/07/02	BMIDS00045 - Capture B&C and PP Insurance Premium - Complete re-work of screen
GD		15/08/02	BMIDS00312 - Add View and New buttons to allow BC to be re-instated correctly
ASu		14/10/02	BMIDS00597 - Premium and Reference fields on B&C should not be mandatory.
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
KRW     16/06/2004  BMIDS776  -  Added Frequency Combo
MC		17/06/2004	BMIDS763	- Insurance Admin Fee validation and Fee Input box added.
MC		30/06/2004  BMIDS763    INSURANCEADMINFEE column added to "BuildingsAndContentsSubQuote" DB Table
								Save and Retrieve data operations changed.
MC		06/07/2004	BMIDS763	INSURANCEADMINFEE Field by default set to disabled.	
MC		13/07/2004	BMIDS763	SAME FIX AS ABOVE txtInsuranceAdminFee DISABLED.
SR		06/08/2004  BMIDS813	- When InsAdmin fee is 0 or empty do not populate it.
HMA     04/10/2004  BMISD903    Set CallingScreen context parameter when LTV has changed.
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
<body><!-- Form Manager Object - Controls Soft Coded Field Attributes -->

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp 
VIEWASTEXT></OBJECT>
<script src="validation.js" language="JScript"></script>
<!-- Specify Forms Here -->
<form id="frmCostModellingMenu" method="post" action="cm010.asp" style="DISPLAY: none">	</form>
<form id="frmValuables" method="post" action="cm035.asp" style="DISPLAY: none">	</form>
<form id="frmViewSubQuotes" method="post" action="cm075.asp" style="DISPLAY: none">	</form>
<form id="frmToGN300" method="post" action="GN300.asp" style="DISPLAY: none">	</form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" year4 validate="onchange" mark>
<div id="divBackground"	 style        ="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 555px">
<span id="spnProduct" style="LEFT: 0px; WIDTH: 605px; POSITION: absolute; TOP: 0px; HEIGHT: 152px" class="msgGroup">

	<span style="LEFT: 10px; POSITION: absolute; TOP: 15px" class ="msgLabel"      >
		Premium
		<span style="LEFT: 175px; POSITION: absolute; TOP: 0px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
			<input id="txtTotalCost" name="TotalCost" maxlength="9" style="LEFT: -1px; WIDTH: 100px; POSITION: absolute; TOP: -3px" class="msgTxt">
		</span>
		
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 68px" class ="msgLabel"      >
		Reference Number
		<span style="LEFT: 175px; POSITION: absolute; TOP: 0px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgTxt"></label>
			<input id="txtReferenceNumber" name="ReferenceNumber" maxlength="40" style="LEFT: -1px; WIDTH: 250px; POSITION: absolute; TOP: -3px" class="msgTxt">
		</span>
		
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 96px" class ="msgLabel"      >
		Cover Type
		<span style="LEFT: 174px; POSITION: absolute; TOP: -3px">
			<select id="cboCoverType" name="CoverType" style="WIDTH: 150px" class="msgCombo"> </select>
		</span>
	</span>	
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 41px" class ="msgLabel"      >
		Frequency
		<span style="LEFT: 174px; POSITION: absolute; TOP: -3px">
			<select id="cboFrequencyType" name="Frequency" style="LEFT: 1px; WIDTH: 150px; TOP: 0px" class="msgCombo"> </select>
		</span>
	</span>	
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 125px" class ="msgLabel">
		Insurance Admin Fee
		<span style="LEFT: 174px; POSITION: absolute; TOP: 0px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
			<input id="txtInsuranceAdminFee" name="InsuranceAdminFee" maxlength="9" style="WIDTH: 149px; POSITION: absolute; TOP: -3px; HEIGHT: 20px" class="msgTxt" size=29 readonly disabled=true>
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
</form><!-- Main Buttons -->
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 220px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div> 
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- #include FILE="attribs/cm020attribs.asp" -->

<script language="JScript">
<!--
var m_sApplicationMode = "";
var m_sReadOnly = "";
var m_sQuotationNumber = "";
var m_sBCSubQuoteNumber = "";
var m_sApplicationNo = "";
var m_sApplicationFFNo = "";
var m_sQuoteFrozen = "0";
var m_sCurrency = "";
var m_InsurancePaymentFrequency = "";
var subQuoteXML = null;
var CoverTypeXML = null;
var FrequencyTypeXML = null;
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
	FW030SetTitles("Buildings & Contents Details", "CM020", scScreenFunctions);
	m_sCurrency = scScreenFunctions.SetCurrency(window,frmScreen);
	RetrieveContextData();
	GetComboLists();
	SetMasks();	
	//DISABLE INSURANCE ADMIN FEE
	scScreenFunctions.SetFieldToDisabled(frmScreen,"txtInsuranceAdminFee");	
	PopulateScreen();
	SetFieldsToReadOnly();	
	Validation_Init();	
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CM020");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
	
}
function PopulateScreen()
{
	subQuoteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if(m_sBCSubQuoteNumber == "")
	{
		//If no sub quote number, then create a new sub quote
		subQuoteXML.CreateRequestTag(window,"CREATE");
		subQuoteXML.CreateActiveTag("QUOTATION");
		subQuoteXML.CreateTag("CONTEXT", m_sApplicationMode);
		subQuoteXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
		subQuoteXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
		subQuoteXML.CreateTag("QUOTATIONNUMBER", m_sQuotationNumber);
		subQuoteXML.RunASP(document,"CreateFirstBCSubQuote.asp");
		if(subQuoteXML.IsResponseOK())
		{
			m_sBCSubQuoteNumber = subQuoteXML.GetTagText("BCSUBQUOTENUMBER");
			scScreenFunctions.SetContextParameter(window,"idBCSubQuoteNumber", m_sBCSubQuoteNumber);
			frmScreen.cboCoverType.selectedIndex = 0;
			frmScreen.btnCreateNewSubQuote.disabled = true
		}
		else
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
		}

		m_sQuoteFrozen = "0";
	}
	else
	{
		//If there is a sub quote number, then go and get the details from the DB
		subQuoteXML.CreateRequestTag(window,null);
		subQuoteXML.CreateActiveTag("BCSUBQUOTEDETAILSNOTES");
		subQuoteXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
		subQuoteXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
		subQuoteXML.CreateTag("BCSUBQUOTENUMBER", m_sBCSubQuoteNumber);
		subQuoteXML.RunASP(document,"GetBCSubQuoteDetails.asp");
		if(subQuoteXML.IsResponseOK())
		{
			PopulateScreenFields();	
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			m_sQuoteFrozen = "1";						
		}	
	}
}
function PopulateScreenFields()
{
	if (subQuoteXML.SelectTag(null, "BCSUBQUOTEDETAILSNOTES") != null)
	{
		frmScreen.cboCoverType.value = subQuoteXML.GetTagText("COVERTYPE");
		frmScreen.cboFrequencyType.value = subQuoteXML.GetTagText("INSURANCEPAYFREQUENCY"); // KRW BMIDS0776 16/06/2004
		frmScreen.txtTotalCost.value = top.frames[1].document.all.scMathFunctions.RoundValue(subQuoteXML.GetTagText("TOTALBCMONTHLYCOST"),2);	
		frmScreen.txtReferenceNumber.value = subQuoteXML.GetTagText("BCREFERENCENUMBER");
		<%/*BMIDS763 - Insurance Admin Fee*/%>
		frmScreen.txtInsuranceAdminFee.value = subQuoteXML.GetTagText("INSURANCEADMINFEE");
		//SECTION END
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
function RetrieveContextData()
{
	m_sApplicationMode = scScreenFunctions.GetContextParameter(window,"idApplicationMode","Quick Quote");
	m_sReadOnly	= scScreenFunctions.GetContextParameter(window,"idReadOnly","0");			
	m_sBCSubQuoteNumber	= scScreenFunctions.GetContextParameter(window,"idBCSubQuoteNumber","1");
	m_sQuotationNumber = scScreenFunctions.GetContextParameter(window,"idQuotationNumber","1");
	m_sApplicationNo = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","APP0001");
	m_sApplicationFFNo = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
}
function SetFieldsToReadOnly()
{
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		DisableMainButton("Submit");
	}
	else
	{		
		scScreenFunctions.SetFocusToFirstField(frmScreen);
	}
}

function GetComboLists()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("BCPaymentFreq","BCCoverType");
	
	var bSuccess = false;
	if (XML.GetComboLists(document, sGroups) == true)
	{
		bSuccess = true;
		CoverTypeXML = XML.GetComboListXML("BCCoverType");
		bSuccess = XML.PopulateComboFromXML(document,frmScreen.cboCoverType,CoverTypeXML,false);
		
    	FrequencyTypeXML = XML.GetComboListXML("BCPaymentFreq");
		bSuccess = XML.PopulateComboFromXML(document,frmScreen.cboFrequencyType,FrequencyTypeXML,false); // KRW BMIDS0776 16/06/2004
		
		
		
	}

	else
		bSuccess = false;
			
	if (bSuccess == false)
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
}

function CreateSubQuoteXML()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequestTag = XML.CreateRequestTag(window, null);
	var xmlBCSubQuoteDetailsTag = XML.CreateActiveTag("BCSUBQUOTEDETAILSNOTES");
	XML.CreateActiveTag("CUSTOMERLIST");
	for (var iCustomerIndex=1; iCustomerIndex<=5; iCustomerIndex++)
	{
		XML.CreateActiveTag("CUSTOMER");
		XML.CreateTag("CUSTOMERNUMBER", scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCustomerIndex, ""))
		XML.CreateTag("CUSTOMERVERSIONNUMBER", scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCustomerIndex, ""));
	}
	XML.ActiveTag = xmlBCSubQuoteDetailsTag;
	XML.CreateActiveTag("BUILDINGSANDCONTENTSSUBQUOTE");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
	XML.CreateTag("BCSUBQUOTENUMBER", m_sBCSubQuoteNumber);
	<% /* MO - BMIDS00807 */ %>
	var today = scScreenFunctions.GetAppServerDate(true);
	<% /* var today = new Date(); */ %>
	XML.CreateTag("DATEANDTIMEGENERATED", today.toLocaleString());
	
	if (m_sApplicationMode == "Quick Quote") XML.CreateTag("QUOTATIONTYPE", "1");
	else XML.CreateTag("QUOTATIONTYPE", "2");
	<% /* ASu BMIDS00597 - Start. If a null value is saved the calculate method on CM010 will fail */ %>
	if (frmScreen.txtTotalCost.value == "")
		frmScreen.txtTotalCost.value = 0;
	<% /* ASu BMIDS00597 - End */ %>
	XML.CreateTag("TOTALBCMONTHLYCOST", frmScreen.txtTotalCost.value);
	XML.CreateTag("INSURANCEPAYFREQUENCY", frmScreen.cboFrequencyType.value); // KRW BMIDS0776 16/06/2004
	XML.CreateTag("BCREFERENCENUMBER", frmScreen.txtReferenceNumber.value);
	
	<%/*BMIDS763 - InsuranceAdminFee*/%>
	XML.CreateTag("INSURANCEADMINFEE",frmScreen.txtInsuranceAdminFee.value);
	<%/*SECTION END - BMIDS763*/%>
	
	XML.ActiveTag = xmlBCSubQuoteDetailsTag;	
	XML.CreateActiveTag("BUILDINGSANDCONTENTSDETAILS");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
	XML.CreateTag("BCSUBQUOTENUMBER", m_sBCSubQuoteNumber);	
	XML.CreateTag("ANNUALBUILDINGSPREMIUM", "");
	XML.CreateTag("ANNUALCONTENTSPREMIUM", "");
	XML.CreateTag("COVERTYPE", frmScreen.cboCoverType.value);
	XML.CreateTag("IPT", "");
	XML.CreateTag("MONTHLYBUILDINGSPREMIUM", "");
	XML.CreateTag("MONTHLYCONTENTSPREMIUM", "");
	XML.CreateTag("ANNUALBCPREMIUM", "");
	XML.CreateTag("STATUS", "");
	XML.CreateTag("BCRETURNTEXT", "");
//	
	if(subQuoteXML.SelectTag(null, "BUILDINGSANDCONTENTSDETAILS") != null)
	{ 
		XML.CreateTag("TOTALVALUABLESEXCEEDINGLIMIT", subQuoteXML.GetTagText("TOTALVALUABLESEXCEEDINGLIMIT"));
	}
	else XML.CreateTag("TOTALVALUABLESEXCEEDINGLIMIT", "");
	XML.ActiveTag = xmlBCSubQuoteDetailsTag;
	XML.CreateActiveTag("BUILDINGSANDCONTENTSNOTES");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
	XML.CreateTag("BCSUBQUOTENUMBER", m_sBCSubQuoteNumber);
	XML.ActiveTag = xmlRequestTag;	
	XML.CreateActiveTag("VALUABLESOVERLIMITLIST");
	if(subQuoteXML.SelectTag(null, "VALUABLESOVERLIMITLIST") != null) 
		XML.AddXMLBlock(subQuoteXML.CreateFragment());	
	return(XML)
}

function SaveScreenData()
{
	var blnSuccess = false;
	if (frmScreen.onsubmit() == true)
	{
		var saveXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		saveXML.CreateRequestTag(window, null);
		var xmlBCsubQuoteTag = saveXML.CreateActiveTag("BCSUBQUOTE");
		saveXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
		saveXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
		saveXML.CreateTag("CUSTOMERNUMBER1", scScreenFunctions.GetContextParameter(window,"idCustomerNumber1",null));
		saveXML.CreateTag("CUSTOMERVERSIONNUMBER1", scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1",null));
		saveXML.CreateTag("CUSTOMERNUMBER2", scScreenFunctions.GetContextParameter(window,"idCustomerNumber2",null));
		saveXML.CreateTag("CUSTOMERVERSIONNUMBER2", scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber2",null));		
		<%/*[MC]BMIDS763 - FEECHANGES*/%>
		//saveXML.CreateTag("INSURANCEADMINFEE", frmScreen.txtInsuranceAdminFee.value);
		//SECTION END
		var XML = CreateSubQuoteXML();
		if (XML.SelectTag(null, "BCSUBQUOTEDETAILSNOTES") != null)
		{
			var BCNotesXML = XML.CreateFragment()
			if (BCNotesXML != null)
			{
				if (BCNotesXML.childNodes.length > 0)
				{
					saveXML.CreateActiveTag("BCSUBQUOTEDETAILSNOTES");				
					saveXML.AddXMLBlock(BCNotesXML);
				}
			}		
			saveXML.ActiveTag = xmlBCsubQuoteTag;
		}
		if (XML.SelectTag(null, "VALUABLESOVERLIMITLIST") != null)
		{	
			var ValuablesXML = XML.CreateFragment();
			if (ValuablesXML != null)
			{
				if (ValuablesXML.childNodes.length > 0)
				{
					saveXML.CreateActiveTag("VALUABLESOVERLIMITLIST");				
					saveXML.AddXMLBlock(ValuablesXML);
				}
			}
		}
		if(m_sApplicationMode == "Quick Quote")
		{
			switch (ScreenRules())
			{
				case 1: // Warning
				case 0: // OK
					saveXML.RunASP(document,"QQCalculateBC.asp");
					break;
				default: // Error
					saveXML.SetErrorResponse();
			}
		}
		else
		{
			switch (ScreenRules())
			{
				case 1: // Warning
				case 0: // OK
					saveXML.RunASP(document, "AQCalculateBC.asp");
					break;
				default: // Error
					saveXML.SetErrorResponse();
			}
		}
		if(saveXML.IsResponseOK())
		{
			blnSuccess = true;
		}
	}
	return(blnSuccess);
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
				scScreenFunctions.SetContextParameter(window, "idCallingScreenID", "CM020");
			}	
		}
	}
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
	//if (m_sBCSubQuoteNumber=="")
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
					scScreenFunctions.SetContextParameter(window,"idBusinessType","BC");
					scScreenFunctions.SetContextParameter(window,"idMetaAction","CM020");
					frmViewSubQuotes.submit();
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
	var newXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	newXML.CreateRequestTag(window, "CREATE");
	newXML.CreateActiveTag("QUOTATION");
	newXML.CreateTag("CONTEXT", m_sApplicationMode);
	newXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
	newXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
	newXML.CreateTag("BCSUBQUOTENUMBER", m_sBCSubQuoteNumber);
	newXML.CreateTag("QUOTATIONNUMBER", m_sQuotationNumber);

	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			newXML.RunASP(document,"CreateNewBCSubQuote.asp");
			break;
		default: // Error
			newXML.SetErrorResponse();
		}

	if(newXML.IsResponseOK())
	{
		m_sBCSubQuoteNumber = newXML.GetTagText("BCSUBQUOTENUMBER");
		scScreenFunctions.SetContextParameter(window,"idBCSubQuoteNumber",m_sBCSubQuoteNumber);	
		<% /*
		if (subQuoteXML.SelectTag(null, "BCSUBQUOTEDETAILSNOTES") != null)
		{
			subQuoteXML.SetTagText("TOTALBCMONTHLYCOST", "");
			if (subQuoteXML.SelectTag(null,"BUILDINGSANDCONTENTSDETAILS") != null)
				subQuoteXML.SetTagText("STATUS", "");
		}
		//change all BCSUBQUOTENUMBER entries in subQuoteXML to the new number
		subQuoteXML.ActiveTag = null;
		var xmlBCSubQuoteNoTagList = subQuoteXML.CreateTagList("BCSUBQUOTENUMBER");
		for(var nLoop = 0;nLoop < xmlBCSubQuoteNoTagList.length;nLoop++)
		{
			var xmlNode = subQuoteXML.GetTagListItem(nLoop);
			xmlNode.text = m_sBCSubQuoteNumber;
		}	
		*/ %>
		//??GDfrmScreen.cboCoverType.onchange();
		//frmScreen.txtStatus.value = "";
		//EMPTY OUT ALL THE FIELDS
		//frmScreen.txtTotalCost.value = "";
		//frmScreen.txtReferenceNumber.value = "";
		frmScreen.cboFrequencyType.selectedIndex = 0; //Default to First in list KRW BMIDS0776 16/06/2004
		frmScreen.cboCoverType.selectedIndex = 0; //Default to First in list
		//frmScreen.txtReason.value = "";
		//scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTotalCost");
		//scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtStatus");
		scScreenFunctions.SetCollectionToWritable(spnProduct);
		
		//BMIDS763 - Insurance Adminfee alway readonly
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtInsuranceAdminFee");
		//section end
		
		//scScreenFunctions.SetCollectionToWritable(spnPayment);
		frmScreen.btnCreateNewSubQuote.disabled = true;	
		//frmScreen.btnCalculate.disabled = false;
		m_sQuoteFrozen = "0";
		frmScreen.txtTotalCost.focus();
	}
}
function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmCostModellingMenu.submit();
}

function frmScreen.cboCoverType.onchange()
{
	if(frmScreen.cboCoverType.selectedIndex<=0)
	{
		return;
	}
	if(IsIAFTypeSelected(frmScreen.cboCoverType.item(frmScreen.cboCoverType.selectedIndex).value,CoverTypeXML))
	{
		<%/*GET PRODUCT SWITCHFEE AND SET IN THE RELEVENT TEXT ELEMENT*/%>
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtInsuranceAdminFee");
		<% /* SR 06/08/2004 : BMIDS813 - When InsAdmin fee is 0 or empty do not populate it. */ %>
		var sTotalAmountDue = getInsuranceAdminFee();
		if (sTotalAmountDue !='0' && sTotalAmountDue !='') frmScreen.txtInsuranceAdminFee.value = sTotalAmountDue;		
		else frmScreen.txtInsuranceAdminFee.value = '';
		<% /* SR 06/08/2004 : BMIDS813 - End */ %>
	}
	else
	{
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtInsuranceAdminFee");
		//scScreenFunctions.SetFieldToWritable(frmScreen,"txtInsuranceAdminFee");
		frmScreen.txtInsuranceAdminFee.value="";
	}
}

function IsIAFTypeSelected(selectedValueId,BCCoverTypeComboXML)
{
	var xmlNodeList, xmlNode, xmlValidationNode ;
	var sPattern ;
	var xmlPSFNodeList;
	var bReturn = new Boolean();
	
	bReturn=false;
	//alert(BCCoverTypeComboXML.xml);
	xmlNodeList = BCCoverTypeComboXML.firstChild.childNodes ;
	
	for(var iCount = xmlNodeList.length -1 ; iCount>=0; --iCount)
	{
		xmlNode = xmlNodeList(iCount) ;
		
		if(xmlNode.selectNodes(".//VALIDATIONTYPELIST[ $any$ VALIDATIONTYPE='IAF']").length > 0)		
		{
			var nodValueID=xmlNode.selectSingleNode("./VALUEID");
			if(nodValueID!=null)
			{
				if(nodValueID.text==selectedValueId)
				{
					bReturn=true;
					break;
				}
			}
		}
		
		if(bReturn){
			break;
		}
	}
	
	return bReturn;
}

function getInsuranceAdminFee()
{
	var lAmount = 0;
	var ListXML=null;

	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(window,"GETINSURANCEADMINFEE")
	ListXML.CreateActiveTag("GETINSURANCEADMINFEE");
		
	ListXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNo);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNo);
			
	ListXML.RunASP(document, "GetInsuranceAdminFee.asp");
		
	<% /* Check Response */ %>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(ErrorTypes);

	if (sResponseArray[0] == true)
	{
		try
		{
			ListXML.ActiveTag = null;
			ListXML.SelectTag(null,"INSURANCEADMINFEE")
			lAmount = ListXML.GetTagText("AMOUNT");
		}
		catch(e)
		{
			//alert(e);
		}
	}
	
	return lAmount;
}
				
-->
</script>
</body>
</html>
<% /* OMIGA BUILD VERSION 046.02.08.08.00 */ %>


