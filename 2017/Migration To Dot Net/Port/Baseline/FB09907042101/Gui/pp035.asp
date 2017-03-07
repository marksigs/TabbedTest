<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<%
/*
Workfile:      pp035.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
IVW		21/12/2000	New Screen - just screen design
SR		15/03/2001	SYS2063
SA		17/05/2001	SYS2223 Changed "Ok" to "OK" on btnOK
SR		08/06/2001	SYS2298 Do not display ComboValues with validation-id 'XDIS'for 'FeeType'
SR		13/12/2001  SYS1845 Do not display combo values with validation XDIS
DRC     16/01/02    SYS3528 Make Other Fee type field mandatory when Other Fee type selected in combo
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog    Date           Description
MC		06/05/2004		White space added to page title
DRC     17/05/2004     CC071 - Regulation change to allow inclusion of Fee in APR calc and/or Add to Loan
MC		15/06/2004	   BMIDS763	ProductSwitch Functionality added
DRC     24/06/2004     BMIDS767 - Need to control add to loan for all fee types as per Lender
MC		30/06/2004		BMIDS763	- Validation rule removed. see sandra's note in clearquest.
SR		14/06/2004		BMIDS767 - Prevent clicking of OK button more than once 
SR		06/08/2004		BMIDS813 - When ProductSwitch fee is 0 or empty do not populate it.
GHun	23/09/2004		BMIDS893 Changed btnOk.onclick to handle multiple validation types.
GHun	28/10/2004		BMIDS936 Changed AllowAddToLoan to exclude END validation type
HMA     21/03/2005      BMIDS977 (CC089)  Call CreateAdHocFeeAndRemodel instead of UpdateQuotationForAddedCosts
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

HMA		23/08/2005		MAR28	Populate valuation fee amount. Tidied comments.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>

<title>PP035 - Add New Fee Type <!-- #include file="includes/TitleWhiteSpace.asp" --></title>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
</HEAD>

<BODY>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Scriptlets */ %>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* DRC 17/05/2004 */ %>
<% /* Specify forms here */ %>
<form id="frmToCM060" method="post" action="CM060.asp" none></form>
<form id="frmToPP010" method="post" action="PP010.asp" none></form>

<% /* DRC 17/05/04 */ %>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<!--<div style="HEIGHT: 190px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 300px" class="msgGroup">-->
<div style="LEFT: 10px; WIDTH: 341px; POSITION: absolute; TOP: 10px; HEIGHT: 172px" class="msgGroup">
	
	<span style="TOP: 13px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Fee Type
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<select id="cboFeeType" maxlength="30" style="POSITION: absolute; WIDTH: 150px" align="right" class="msgCombo"></select>
		</span>
	</span>
	
	<span style="TOP: 46px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Other Fee
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<input id="txtOtherFee" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	
	<span style="TOP: 79px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Total Amount Due
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<input id="txtTotalAmountDue" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	
		<% /* DRC 17/05/2004 Add "Add fee to loan?" Radios */ %>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 112px" class="msgLabel" id=SPAN2>
		Is the fee to be added to Loan
		<span style="LEFT: 165px; POSITION: absolute; TOP: -3px" id=SPAN3>
			<input id="optAddToLoanYes" name="RadioGroup1" type="radio" value="1"><label for="optAddToLoanYes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optAddToLoanNo" name="RadioGroup1" type="radio" value="0"><label for="optAddToLoanNo" class="msgLabel">No</label>
		</span>
	</span>
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 142px">
		<% /*SYS2223 Changed "Ok" to "OK"*/%>
		<input id="btnOk" value="OK" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
	<span style="LEFT: 100px; POSITION: absolute; TOP: 142px">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
</div>
</form>

<!-- #include FILE="attribs/PP035attribs.asp" -->

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;
var m_sApplicationNumber;
var m_sFeeType ;
var m_sFeeTypeDesc;
var m_sInitialAmount ;
var m_sNotes ;
var m_sFeeTypeSeq ;
var m_aArgArray = null;
var m_BaseNonPopupWindow = null;
<% /* DRC 17/05/2004 start */ %>
var m_sApplicationFactFindNumber;
var m_sMortgageSubQuoteNumber;
var m_sAcceptedQuoteNumber;
<% /* DRC 17/05/2004 */ %>
<% /* DRC 23/06/2004 Start*/ %>
var m_lenderXML = null;
<% /* DRC 23/06/2004 */ %>

var m_bIsSubmit = false; <% /* SR 14/06/2004 : BMIDS767  */ %>
var m_sValuationRefundAmount = "";   //MAR28

<% /**** Events *****/ %>

function frmScreen.btnCancel.onclick()
{
	window.close();
}

function frmScreen.btnOk.onclick()
{
	<% /* SR 14/06/2004 : BMIDS767 - Prevent clicking of OK button more than once  */ %>  
	if (m_bIsSubmit) return ;
	m_bIsSubmit = true ;
	frmScreen.btnOk.disabled = true ;
	<% /* SR 14/06/2004 : BMIDS767 - End */ %>  
	if (frmScreen.onsubmit())
		{
				switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
					<% /* DRC  17/05/2004 Add to Loan button */ %>
					if (frmScreen.optAddToLoanYes.checked == true )
					{    
						UpdateQuotation();
						<% /* TK 10/05/2004 BBG234 frmToCM060.submit(); */ %>
						var sReturn = new Array();
						sReturn[0] = "CM060";
						window.returnValue = sReturn;
						window.close();
					}
					else
					{	
						<% /* BMIDS893 GHun Check all validation types, not just the 1st one
						if (scScreenFunctions.GetComboValidationType(frmScreen, "cboFeeType") == "APR")
						*/ %>
						if (scScreenFunctions.IsOptionValidationType(frmScreen.cboFeeType, frmScreen.cboFeeType.selectedIndex, "APR"))
						{
							UpdateQuotation();
						}
						CreateApplicationFeeType();
		<% /* DRC 17/05 submit to PP010 directly  */ %>			
		
						var sReturn = new Array();
						sReturn[0] = null;
						window.returnValue = sReturn;
						window.close();
					}
					break ;
					
				default: // Error
					bContinue = false ;
				}
		}
	<% /* SR 14/06/2004 : BMIDS767 - Prevent clicking of OK button more than once  */ %>  
	m_bIsSubmit = false ;
	frmScreen.btnOk.disabled = false ;
	<% /* SR 14/06/2004 : BMIDS767 - End */ %>  
}

function frmScreen.cboFeeType.onchange()
{
	var sTotalAmountDue;
	
	<% /* SR 06/08/2004 : BMIDS813 - If button OK was disabled, enable it now */ %>
	if(frmScreen.btnOk.disabled) frmScreen.btnOk.disabled = false ; 
	
	<% /* BMIDS893 GHun Use validation type OTH rather than hard coded value 11
	if (frmScreen.cboFeeType.value=="11")
	*/ %>
	if (scScreenFunctions.IsOptionValidationType(frmScreen.cboFeeType, frmScreen.cboFeeType.selectedIndex, "OTH"))
		{
			frmScreen.txtOtherFee.value="";
	<% /* AQR SYS3528 DRC Make Field mandatory */ %>
			frmScreen.txtOtherFee.setAttribute("required", "true");
			frmScreen.txtOtherFee.parentElement.parentElement.style.color="red";
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtOtherFee");
		}
	else
		{
			
			frmScreen.txtOtherFee.setAttribute("required", "false");
			frmScreen.txtOtherFee.parentElement.parentElement.style.color="black";			
			scScreenFunctions.SetFieldToDisabled(frmScreen,"txtOtherFee");
		}
		
	    if (AllowAddToLoan(frmScreen.cboFeeType.item(frmScreen.cboFeeType.selectedIndex).value, m_OneOffCostXML) == true)
		{
			scScreenFunctions.SetRadioGroupState(frmScreen, "RadioGroup1", "W");
			frmScreen.optAddToLoanNo.checked=false;
			frmScreen.optAddToLoanYes.checked=true;
			
		}
		else
		{
			frmScreen.optAddToLoanNo.checked=true;
			frmScreen.optAddToLoanYes.checked=false;
			scScreenFunctions.SetRadioGroupState(frmScreen, "RadioGroup1", "R");
		}

	<%/*IsProductSwitchFeeSelected() validation has been dropped. see clearquest notes for more information.*/%>
	if ( IsPSFTypeSelected(frmScreen.cboFeeType.item(frmScreen.cboFeeType.selectedIndex).value,m_OneOffCostXML))
	{		
		<%/*GET PRODUCT SWITCHFEE AND SET IN THE RELEVENT TEXT ELEMENT*/%>
		<% /* SR 06/08/2004 : BMIDS813 - When switch fee is 0 or empty do not populate it. Disable the button OK */ %>
		sTotalAmountDue = getProductSwitchFee();
		if (sTotalAmountDue !='0' && sTotalAmountDue !='') frmScreen.txtTotalAmountDue.value = sTotalAmountDue ;
		else
		{
			frmScreen.txtTotalAmountDue.value = '';
			frmScreen.btnOk.disabled = true ;
		}
		
		<% /* SR 06/08/2004 : BMIDS813 - End */ %>	
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTotalAmountDue");
	}
	else
	{
		<% /* MAR28 If the type is Valuation Fee, look up the amount for the current scheme */ %>
		var ValidationList = new Array(1);
		ValidationList[0] = "VAL";
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		if( XML.IsInComboValidationList(document,"OneOffCost", frmScreen.cboFeeType.value, ValidationList))
		{
			<% /* Get Valuation Fee Set and Amount */ %>
			sTotalAmountDue = getValuationFee();
			
			if (sTotalAmountDue !="0" && sTotalAmountDue !="") 
			{
				frmScreen.txtTotalAmountDue.value = sTotalAmountDue ;
			}
			else
			{
				frmScreen.txtTotalAmountDue.value = "";
			}		
		}
	
		else
		{
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtTotalAmountDue");
			frmScreen.txtTotalAmountDue.value="";
		}
	}
}

function getProductSwitchFee()
{
	var lAmount = 0;
	var ListXML=null;

	ListXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(m_BaseNonPopupWindow,"GETPRODUCTSWITCHFEE")
	ListXML.CreateActiveTag("GETPRODUCTSWITCHFEE");
		
	ListXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
			
	ListXML.RunASP(m_BaseNonPopupWindow.top.frames[1].document, "GetProductSwitchFee.asp");
			
	<% /* Check Response */ %>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(ErrorTypes);

	if (sResponseArray[0] == true) 
	{
		try
		{
			ListXML.ActiveTag = null;
			ListXML.SelectTag(null,"PRODUCTSWITCHFEE")
			lAmount = ListXML.GetTagText("AMOUNT");
			
			if(lAmount==0 || lAmount == '') alert('A Product Switching Fee does not apply to this application');
		}
		catch(e)
		{
			alert(e);
		}
	}
	else 
		if(sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
		{
			alert('A Product Switching Fee does not apply to this application');
		}
	
	return lAmount;
}

<% /* MAR28 Add new function */ %>
function getValuationFee()
{
	var sAmount = "";
	var ListXML=null;

	ListXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	ListXML.CreateRequestTag(m_BaseNonPopupWindow,"GETVALUATIONFEE")
	ListXML.CreateActiveTag("GETVALUATIONFEE");
		
	ListXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	ListXML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber);
	
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			ListXML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			ListXML.SetErrorResponse();
		}

	<% /* Check Response - a record not found error is valid */%>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(ErrorTypes);

	if (sResponseArray[0] == true) 
	{
		sAmount = ListXML.GetTagText("AMOUNT");
		m_sValuationRefundAmount = ListXML.GetTagText("REFUNDAMOUNT");
	}
		
	return sAmount;
}

function getApplicationMortgageLender()
{
		var ListXML=null;

		ListXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		ListXML.CreateRequestTag(m_BaseNonPopupWindow,null)
		ListXML.CreateActiveTag("MORTGAGELENDER");
		
		ListXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		ListXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
			
		/**/	
		ListXML.RunASP(m_BaseNonPopupWindow.top.frames[1].document, "GetApplicationMortgageLender.asp");
		
		/*Check Response*/
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var sResponseArray = ListXML.CheckResponse(ErrorTypes);

		if (sResponseArray[0] == true || 
			sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
		{
			try
			{
				return ListXML;
			}
			catch(e)
			{
			}
		}
		
		return null;
}

<% /* MAR28 Remove unused function IsProductSwitchFeeSelected */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sArguments = window.dialogArguments ;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];	
	m_aArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sApplicationNumber = m_aArgArray[0];
	
	<% /* DRC 17/05/2004 start */ %>
	
	m_sApplicationFactFindNumber = m_aArgArray[2];
	m_sMortgageSubQuoteNumber = m_aArgArray[3];
	m_sAcceptedQuoteNumber = m_aArgArray[4];
	
	<% /* DRC 17/05/2004 end */ %>

    <% /* DRC 23/06/2004 Start		
    Move call to get the MortgageLender Data to here so it is only done once */ %>
	m_lenderXML = getApplicationMortgageLender();
	if(m_lenderXML!=null)
	{
		m_lenderXML.ActiveTag = null;
		m_lenderXML.SelectTag(null,"MORTGAGELENDER")
	}	
	<% /* DRC 23/06/2004 */ %>
	SetMasks();  
	Validation_Init();
	
	frmScreen.txtOtherFee.value="";
	scScreenFunctions.SetFieldToDisabled(frmScreen,"txtOtherFee");
	<% /* DRC 17/05/2004 */ %>
	frmScreen.optAddToLoanNo.checked = false;
	InitialiseScreen();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	ClientPopulateScreen();
}

<% /**** FUNCTIONS *******/ %>
function InitialiseScreen()
{
	PopulateCombos();
}

function PopulateCombos()
{
	var XMLCombos = null;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	var sGroupList = new Array("OneOffCost");
	if(XML.GetComboLists(document,sGroupList))
	{
		XMLCombos = XML.GetComboListXML("OneOffCost");
		<% /*[MC]BMIDS763 - CC075 */ %>
		m_OneOffCostXML = XMLCombos;
		<% /*SECTION END */ %>
		RemoveNonDisplayValues(XMLCombos) ;
		
		var blnSuccess = true;
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboFeeType,XMLCombos,true);
			
		if(!blnSuccess)
		{
			alert('Failed to populate combos');
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			frmScreen.btnOk.disabled = true ;
		}
	}
}

function RemoveNonDisplayValues(XMLCombos)
{
	var xmlNodeList, xmlNode, xmlValidationNode ;
	var sPattern ;
	
	xmlNodeList = XMLCombos.firstChild.childNodes ;
	
	for(var iCount = xmlNodeList.length -1 ; iCount>=0; --iCount)
	{
		xmlNode = xmlNodeList(iCount) ;
		<% /* SR 13-12-2001 : SYS1845 Do not display combo values with validation XDIS */ %>
		if(xmlNode.selectNodes(".//VALIDATIONTYPELIST[ $any$ VALIDATIONTYPE='XDIS']").length > 0)
			XMLCombos.firstChild.removeChild(xmlNode) ;
	}

}

function IsPSFTypeSelected(selectedValueId,OneOffCostComboXML)
{
	var xmlNodeList, xmlNode, xmlValidationNode ;
	var sPattern ;
	var xmlPSFNodeList;
	var bReturn = new Boolean();
	bReturn=false;
	
	xmlNodeList = OneOffCostComboXML.firstChild.childNodes ;
	
	for(var iCount = xmlNodeList.length -1 ; iCount>=0; --iCount)
	{
		xmlNode = xmlNodeList(iCount) ;
		
		if(xmlNode.selectNodes(".//VALIDATIONTYPELIST[ $any$ VALIDATIONTYPE='PSF']").length > 0)
		{
			try
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
			catch(e)
			{
				alert(e);
			}
		}
		
		if(bReturn)
		{
			break;
		}
	}
	
	return bReturn;
}
function AllowAddToLoan(selectedValueId,OneOffCostComboXML)
{
	var xmlNodeList, xmlNode, xmlValidationNode ;
	var sPattern ;
	var xmlPSFNodeList;
	var bReturn = new Boolean();
	bReturn=false;
	bFound=false;
	
	xmlNodeList = OneOffCostComboXML.firstChild.childNodes ;
	
	for(var iCount = xmlNodeList.length -1 ; iCount>=0; --iCount)
	{
		xmlNode = xmlNodeList(iCount) ;
		var nodValueID=xmlNode.selectSingleNode("./VALUEID");
		if(nodValueID!=null)
		{
			if(nodValueID.text==selectedValueId)
			{
			   <% /* Look for the Identfier */ %>
			   var xmlValTypeNode 
			   var xmlValTypeList = xmlNode.selectNodes(".//VALIDATIONTYPELIST/VALIDATIONTYPE")
			   for(var vCount = xmlValTypeList.length - 1; vCount>=0; --vCount)
			   {
					xmlValTypeNode = xmlValTypeList(vCount);
					var sValidationType = xmlValTypeNode.text;
					<% /* BMIDS936 GHun also exclude the END validation type */ %>
					if ((sValidationType != "APR")&&(sValidationType != "XAC")&&(sValidationType != "XDIS")&&(sValidationType != "END"))
					{
			     
						if( (sValidationType == "ADM" && m_lenderXML.GetTagText("ALLOWADMINFEEADDED") == "1") ||
							(sValidationType == "ARR" && m_lenderXML.GetTagText("ALLOWARRGEMTFEEADDED") == "1") ||
							(sValidationType == "DEE" && m_lenderXML.GetTagText("ALLOWDEEDSRELFEEADD") == "1") ||
							(sValidationType == "MIG" && m_lenderXML.GetTagText("ALLOWMIGFEEADDED") == "1") ||
							(sValidationType == "POR" && m_lenderXML.GetTagText("ALLOWPORTINGFEEADDED") == "1") ||
							(sValidationType == "REI" && m_lenderXML.GetTagText("ALLOWREINSPTFEEADDED") == "1") ||
							(sValidationType == "REV" && m_lenderXML.GetTagText("ALLOWREVALUATIONFEEADDED") == "1") ||
							(sValidationType == "SEA" && m_lenderXML.GetTagText("ALLOWSEALINGFEEADDED") == "1") ||
							(sValidationType == "TTF" && m_lenderXML.GetTagText("ALLOWTTFEEADDED") == "1") ||
							(sValidationType == "VAL" && m_lenderXML.GetTagText("ALLOWVALNFEEADDED") == "1") ||
							(sValidationType == "LEG" && m_lenderXML.GetTagText("ALLOWLEGALFEEADDED") == "1") ||
							(sValidationType == "OTH" && m_lenderXML.GetTagText("ALLOWOTHERFEEADDED") == "1") ||   
							(sValidationType == "IAF" && m_lenderXML.GetTagText("ALLOWNONLENDINSADMINFEEADDED") == "1") ||
							(sValidationType == "PSF" && m_lenderXML.GetTagText("ALLOWPRODUCTSWITCHFEEADDED") == "1"))

							{
									bReturn=true;
									break;
							}
				
				  } // End if
			   } // Inner For loop	
			   bFound = true;
			} // End if
		} // End If
		
		if(bFound)
		{
			break;
		}
	} // Outer For Loop
	
	return bReturn;
}



function CreateApplicationFeeType()
{
	var bContinue ;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	XML.CreateRequestTagFromArray(m_aArgArray[1], "CREATEAPPLICATIONFEETYPE");
	
	XML.CreateActiveTag("APPLICATIONFEETYPE");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("FEETYPE", frmScreen.cboFeeType.value);
	XML.SetAttribute("DESCRIPTION", frmScreen.txtOtherFee.value);
	XML.SetAttribute("AMOUNT", frmScreen.txtTotalAmountDue.value);
	<% /* DRC 17/05/2004 add ADHOC indicator */ %>
	XML.SetAttribute("ADHOCIND","1");
	XML.SetAttribute("REBATEORADDITION","0.0");
	XML.SetAttribute("REFUNDAMOUNT", m_sValuationRefundAmount);
	
	<% /* 	XML.RunASP(document,"PaymentProcessingRequest.asp");
	 Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	bContinue = XML.IsResponseOK();
	
	return bContinue
}
<% /* New functions added DRC 17/05/2004 (ex PJO for BBG) */ %>

function UpdateQuotation()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sAddToLoan ;
	
	if (frmScreen.optAddToLoanYes.checked == true)
	{
		sAddToLoan = "1" ;
	}
	else
	{
		sAddToLoan = "0" ;
	}
	
	XML.CreateRequestTagFromArray(m_aArgArray[1], "CREATEADHOCFEEANDREMODEL");   // BMIDS977
	
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber);
	XML.CreateTag("ACCEPTEDQUOTENUMBER", m_sAcceptedQuoteNumber);
	XML.CreateActiveTag("ADDEDONEOFFCOST")
	XML.CreateTag("FEETYPE", frmScreen.cboFeeType.value);
	XML.CreateTag("AMOUNT", frmScreen.txtTotalAmountDue.value);
	XML.CreateTag("ADDTOLOAN",sAddToLoan);
	XML.CreateTag("ADHOCIND","1");
	XML.CreateTag("CALLEDFROM","PP035");                                         // BMIDS977
	XML.CreateTag("REFUNDAMOUNT", m_sValuationRefundAmount);                     // MAR28
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"CreateAdHocFeeAndRemodel.asp");                 // BMIDS977
			break;
		default: // Error
			XML.SetErrorResponse();
		}
	
	bContinue = XML.IsResponseOK();
	
	return bContinue
}

-->
</SCRIPT>
</BODY>
</HTML>
