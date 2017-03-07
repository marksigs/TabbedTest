<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      cr044.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Mortgage Summary screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		15/11/2000	Amended screen for demo purposes.
RF		16/03/01	Amended to match new version of PSG adapter.
RF		10/04/01	Added timer functionality to window.onload
JLD		26/03/02	SYS4320 various errors
GD		24/04/02	SYS4468 - Optimus Data not being picked up correctly. 
							  Considerable rework of table display, and functionality.
DS		02/05/02	SYS4528 Rounding, percentage and repayment type fixed
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MV		13/08/2002  BMIDS00329  Use Cached Version
GHun	30/10/2002	BMIDS00769	Change OPERATION value to be mixed case
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
PSC		18/10/2005	MAR57		Change Routing
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
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>

<% /* Scriptlets - remove any which are not required */ %>
<span style="LEFT: 300px; POSITION: absolute; TOP: 500px">
	<OBJECT data=scTableListScroll.asp id=scScrollTable style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToCR040" method="post" action="CR040.asp" STYLE="DISPLAY: none"></form>
<% /* PSC 18/10/2005 MAR57 */ %>
<form id="frmToCR041" method="post" action="cr041.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div style="HEIGHT: 200px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 302px" class="msgGroup">
	<span id="spnAddressDetails" style="LEFT: 4px; POSITION: absolute; TOP: 4px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Home Telephone
			<span style="LEFT: 120px; POSITION: absolute; TOP: -1px">
				<input id="txtHomePhone" maxlength="20" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 25px" class="msgLabel">
			Post Code
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<input id="txtPostCode" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 50px" class="msgLabel">
			Property Address
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<input id="txtAddress1" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 70px" class="msgLabel">
			<br>
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<input id="txtAddress2" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 90px" class="msgLabel">
			<br>
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<input id="txtAddress3" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 110px" class="msgLabel">
			<br>
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<input id="txtAddress4" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 130px" class="msgLabel">
			<br>
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<input id="txtAddress5" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 150px" class="msgLabel">
			Applicants			
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<select id="cboApplicants" style="WIDTH: 150px" class="msgCombo"></select>
			</span>
		</span>
	</span>
</div>

<div style="HEIGHT: 200px; LEFT: 302px; POSITION: absolute; TOP: 60px; WIDTH: 315px" class="msgGroup">
	<span id="spnPaymentDetails" style="LEFT: 4px; POSITION: absolute; TOP: 4px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Arrears Status
			<span style="LEFT: 100px; POSITION: absolute; TOP: -1px">
				<input id="txtArrearsStatus" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 25px" class="msgLabel">
			Next payment
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtNextPaymentAmount" maxlength="20" style="POSITION: absolute; WIDTH: 70px" class="msgCurrency">
				<input id="txtNextPaymentDate" maxlength="20" style="LEFT: 80px; POSITION: absolute; WIDTH: 70px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 50px" class="msgLabel">
			Payment Method
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtPaymentMethod" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 75px" class="msgLabel">
			First advance date
			<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
				<input id="txtAdvDate" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgDate">
			</span>
		</span>
	</span>
</div>

<div style="HEIGHT: 258px; LEFT: 10px; POSITION: absolute; TOP: 290px; WIDTH: 315px" class="msgGroup">
	<span id="spnValuation" style="LEFT: 4px; POSITION: absolute; TOP: 4px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px" class="msgLabel">
			Valuation
			<span style="LEFT: 120px; POSITION: absolute; TOP: -1px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtValuation" maxlength="20" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 25px" class="msgLabel">
			Loan To Value Ratio
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<input id="txtLTV" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 50px" class="msgLabel">
			Suspense
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtSuspense" maxlength="20" style="POSITION: absolute; WIDTH: 140px" class="msgCurrency">
			</span>
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 75px" class="msgLabel">
			Paid in Advance
			<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
				<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
				<input id="txtPaidInAdvance" maxlength="20" style="POSITION: absolute; WIDTH: 140px" class="msgCurrency">
			</span>
		</span>
	</span>
</div>
<div  style="HEIGHT: 260px; LEFT: 302px; POSITION: absolute; TOP: 290px; WIDTH: 315px" class="msgGroup">
	<span id="spnPropertyAddress" style="LEFT: 4px; POSITION: absolute; TOP: 4px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 4px" class="msgLabel">
			<b>Insurances</b>
			<span style="LEFT: 0px; POSITION: absolute; TOP: 25px" class="msgLabel">
				Buildings and Contents
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtBuildings" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgCurrency">
				</span>
			</span>
			<span style="LEFT: 0px; POSITION: absolute; TOP: 50px" class="msgLabel">
				Payment Protection
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtPayment" maxlength="40" style="POSITION: absolute; WIDTH: 140px" class="msgCurrency">
				</span>
			</span>
		</span>
	</span>
</div>
<div style="HEIGHT: 100px; LEFT: 10px; POSITION: absolute; TOP: 386px; WIDTH: 596px" class="msgGroup">
	<span id="spnMortSummary" style="LEFT: 4px; POSITION: absolute; TOP: 0px">
	<table id="tblMortSummary" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
		<tr id="rowTitles">
			<td width="4%" class="TableHead">Product</td>	
			<td width="18%" class="TableHead">Type</td>	
			<td width="10%" class="TableHead">Interest Rate</td>		
			<td width="14%" class="TableHead">Product End</td>
			<td width="14%" class="TableHead">Loan End</td>
			<td width="8%" class="TableHead">Instalment</td><!--			<td width="8%" class="TableHead">REP</td>-->
			<td width="8%" class="TableHead">Arrears</td>
			<td width="16%" class="TableHead">Loan Balance</td>
		</tr>
		<tr id="row01">		
			<td class="TableTopLeft">&nbsp;</td>		
			<td class="TableTopCenter">&nbsp;</td>		
			<td class="TableTopCenter">&nbsp;</td>		
			<td class="TableTopCenter">&nbsp;</td>
			<td class="TableTopCenter">&nbsp;</td>		
			<td class="TableTopCenter">&nbsp;</td>		
			<td class="TableTopCenter">&nbsp;</td><!--			<td class="TableTopCenter">&nbsp;</td>-->
			<td class="TableTopRight">&nbsp;</td>
		</tr>
		<tr id="row02">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td><!--			<td class="TableCenter">&nbsp;</td>-->
			<td class="TableRight">&nbsp;</td>
		</tr>
		<tr id="row03">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td><!--			<td class="TableCenter">&nbsp;</td>-->
			<td class="TableRight">&nbsp;</td>
		</tr>
		<tr id="row04">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td><!--			<td class="TableCenter">&nbsp;</td>-->
			<td class="TableRight">&nbsp;</td>
		</tr>
		<tr id="row05">		
			<td class="TableLeft">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td>			
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td>	
			<td class="TableCenter">&nbsp;</td>		
			<td class="TableCenter">&nbsp;</td><!--			<td class="TableCenter">&nbsp;</td>-->
			<td class="TableRight">&nbsp;</td>
		</tr>
		<tr id="row06">		
			<td class="TableBottomLeft">&nbsp;</td>		
			<td class="TableBottomCenter">&nbsp;</td>		
			<td class="TableBottomCenter">&nbsp;</td>		
			<td class="TableBottomCenter">&nbsp;</td>
			<td class="TableBottomCenter">&nbsp;</td>		
			<td class="TableBottomCenter">&nbsp;</td>		
			<td class="TableBottomCenter">&nbsp;</td><!--			<td class="TableBottomCenter">&nbsp;</td>-->
			<td class="TableBottomRight">&nbsp;</td>
		</tr>
	</table>
</span><!--SYS4468 - Put Summary outside table-->
<span style="LEFT: 195px; POSITION: absolute; TOP: 167px" class="msgLabel">
	<B>BALANCES</B>
</span>
<span style="LEFT: 347px; POSITION: absolute; TOP: 167px" class="msgLabel">
	<label style="LEFT: -8px; POSITION: absolute; TOP: 3px" class="msgCurrency"></label>
	<input id="txtInstallmentSum" style="POSITION: absolute; WIDTH: 65px" class="msgReadOnly" READONLY tabindex=-1>
</span>
<span style="LEFT: 450px; POSITION: absolute; TOP: 167px" class="msgLabel">
	<label style="LEFT: -8px; POSITION: absolute; TOP: 3px" class="msgCurrency"></label>
	<input id="txtArearsSum" style="POSITION: absolute; WIDTH: 65px" class="msgReadOnly" READONLY tabindex=-1>
</span>
<span style="LEFT: 535px; POSITION: absolute; TOP: 167px" class="msgLabel">
	<label style="LEFT: -8px; POSITION: absolute; TOP: 3px" class="msgCurrency"></label>
	<input id="txtLoanBalanceSum" style="POSITION: absolute; WIDTH: 65px" class="msgReadOnly" READONLY tabindex=-1>
</span>

</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 590px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" -->
</div>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cr044Attribs.asp" -->
<% /* File containing field attributes (remove if not required) 
*/ %>
<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sOtherSystemCustomerNumber = null;
var m_sOtherSystemAccountNumber = "";
var m_sCurrentCustAddressSeqNo = null;
var m_sOptimusSessionId = null;
var m_optimusSessionEffectiveDate = null;

var m_nInstallmentTotal = 0;
var m_nArrearsTotal = 0;
var m_nBalanceTotal = 0;

var m_bIsSubmit = false;
var scScreenFunctions;
var XML;
var arrList, arrRow;
var iTableSize = 5;
var nNumPIComponents;
var xmlRepaymentTypeCombo;
var m_sXMLRepaymentTypes = "";
<% /* PSC 18/10/2005 MAR57 */%>
var	m_sCallingScreen = null;



<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	<% /*MV - 13/08/2002 - BMIDS00329 - Use Cached Version */%>
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	<%	// Make the required buttons available on the bottom of the screen %>	
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

		
	FW030SetTitles("Mortgage Summary","CR044",scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required 
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
<%	//This function is contained in the field attributes file (remove if not required)
%>	//SetMasks();
	//Validation_Init();
	PopulateScreen();
	
	scScreenFunctions.SetScreenToReadOnly(frmScreen);

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	frmScreen.cboApplicants.disabled = false;

	
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

function PopulateScreen()
{
	<% /*MV - 13/08/2002 - BMIDS00329 - Use Cached Version */%>
	xmlRepaymentTypeCombo = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var sGroups = new Array("RepaymentType");
	if (xmlRepaymentTypeCombo.GetComboLists(document, sGroups) != true)
	{
		alert("Error loading repayment type combo values");
	}
	
	m_sXMLRepaymentTypes = xmlRepaymentTypeCombo.XMLDocument.xml;
	
	<%/* MV - 13/08/2002 - BMIDS00329 - Use Cached Version */%>
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	<% /* BMIDS00769  XML.CreateRequestTag(window,"GETACCOUNTSUMMARY"); */ %>
	XML.CreateRequestTag(window,"GetAccountSummary");
	
	XML.CreateActiveTag("MORTGAGEACCOUNT");   
	var sCollateralNumber = m_sOtherSystemAccountNumber;
	XML.SetAttribute("ACCOUNTNUMBER", sCollateralNumber);
	
	XML.RunASP(document, "omAdmin.asp");

	if(XML.IsResponseOK())
	{
		XML.SelectTag(null,"CUSTOMERLIST");
		var xmlCustomerList = XML.CreateTagList("CUSTOMER");
		var xmlTelephoneList, xmlTelNum;
		
		if (xmlCustomerList.length != 0) 
		{
			xmlTelephoneList = xmlCustomerList.item(0).selectSingleNode("TELEPHONENUMBERLIST");
			xmlTelephoneDetails = xmlTelephoneList.getElementsByTagName("TELEPHONENUMBERDETAILS");
		
			for (var i = 0; i < xmlTelephoneDetails.length; i++)
			{
				if ((xmlTelephoneDetails.item(i).getAttribute("USAGE") != null) &&
					 (xmlTelephoneDetails.item(i).getAttribute("USAGE") == "H" || xmlTelephoneDetails.item(i).getAttribute("USAGE") == "1"))
				{
					frmScreen.txtHomePhone.value = xmlTelephoneDetails.item(i).getAttribute("COUNTRYCODE") + 
												   xmlTelephoneDetails.item(i).getAttribute("AREACODE") + " " + 
												   xmlTelephoneDetails.item(i).getAttribute("TELEPHONENUMBER");
				}
			}

			
			for (var i = 0; i < xmlCustomerList.length; i++)
			{
				var TagSELECT = document.createElement("OPTION");	  
				var sComboText = (xmlCustomerList.item(i).getAttribute("FIRSTFORENAME") +
				" " + xmlCustomerList.item(i).getAttribute("SECONDFORENAME") +
				" " + xmlCustomerList.item(i).getAttribute("SURNAME") + 
				" " + xmlCustomerList.item(i).getAttribute("OTHERFORENAMES"));
				TagSELECT.text = sComboText.trim();
				TagSELECT.value = ""; 
				frmScreen.cboApplicants.options.add(TagSELECT);
			}
		}
		
		XML.SelectTag(null,"PROPERTY");
		
		<%/* MV - 13/08/2002 - BMIDS00329 - Use Cached Version */%>
		frmScreen.txtValuation.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetAttribute("VALUATIONAMOUNT"),2);
		
		XML.SelectTag(null,"PROPERTYADDRESS");
		frmScreen.txtPostCode.value = XML.GetAttribute("POSTCODE");
		
		if (XML.GetAttribute("FLATNUMBER") != "")
		{
			frmScreen.txtAddress1.value = XML.GetAttribute("FLATNUMBER") + " ";
		}
		if (XML.GetAttribute("BUILDINGORHOUSENAME") != "")
		{
			frmScreen.txtAddress1.value += XML.GetAttribute("BUILDINGORHOUSENAME");
		}
		if (XML.GetAttribute("BUILDINGORHOUSENUMBER") != "")
		{
			frmScreen.txtAddress2.value = XML.GetAttribute("BUILDINGORHOUSENUMBER") + " ";
		}
		if (XML.GetAttribute("STREET") != "")
		{
			frmScreen.txtAddress2.value += XML.GetAttribute("STREET");
		}
		if (XML.GetAttribute("DISTRICT") != "")
		{
			frmScreen.txtAddress3.value = XML.GetAttribute("DISTRICT");
		}
		//GD SYS4468
		if (XML.GetAttribute("TOWN") != "")
		{
			frmScreen.txtAddress4.value = XML.GetAttribute("TOWN");
		}
		
		if (XML.GetAttribute("COUNTY") != "")
		{
			frmScreen.txtAddress5.value = XML.GetAttribute("COUNTY");
		}
		
		XML.SelectTag(null,"ARREARSDETAILS");
		frmScreen.txtArrearsStatus.value = XML.GetAttribute("ARREARSSTATUS");

		XML.SelectTag(null,"PAYMENTDETAILS");
		
		<%/*MV - 13/08/2002 - BMIDS00329 - Use Cached Versions */%>
		frmScreen.txtNextPaymentAmount.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetAttribute("PAYMENTAMOUNT"),2);
		
		frmScreen.txtNextPaymentDate.value = XML.GetAttribute("PAYMENTDATE");
		frmScreen.txtPaymentMethod.value = XML.GetAttribute("PAYMENTMETHOD");

		XML.SelectTag(null,"ADVANCE");
		frmScreen.txtAdvDate.value = XML.GetAttribute("FIRSTDATE");

		XML.SelectTag(null,"LTV");
		<%/* MV - 13/08/2002 - BMIDS00329 - Use Cached Versions */%>
		frmScreen.txtLTV.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetAttribute("RATIOVALUE"),2);

		XML.SelectTag(null,"SUSPENSEACCOUNT");
		<%/* MV - 13/08/2002 - BMIDS00329 - Use Cached Versions */%>
		frmScreen.txtSuspense.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetAttribute("BALANCE"),2);
		
		XML.SelectTag(null,"ADVANCEACCOUNT");
		<%/* MV - 13/08/2002 - BMIDS00329 - Use Cached Versions */%>
		frmScreen.txtPaidInAdvance.value = top.frames[1].document.all.scMathFunctions.RoundValue(XML.GetAttribute("BALANCE"),2);
		
		var curTotal = 0.00;
		var xmlNodeList = XML.XMLDocument.getElementsByTagName("BUILDINGSCONTENTSINSURANCES");
		for (var i=0; i < xmlNodeList.length ; i++ ) 
		{
			curTotal += parseFloat(xmlNodeList.item(i).getAttribute("PREMIUM"));
		}
		
		<%/* MV - 13/08/2002 - BMIDS00329 - Use Cached Versions */%>
		frmScreen.txtBuildings.value = top.frames[1].document.all.scMathFunctions.RoundValue(curTotal,2);
			
		curTotal = 0.00;
		xmlNodeList = XML.XMLDocument.getElementsByTagName("PAYMENTPROTECTION");
		for (var i=0; i < xmlNodeList.length ; i++ ) 
		{
			curTotal += parseFloat(xmlNodeList.item(i).getAttribute("PREMIUM"));
		}
		
		<%/* MV - 13/08/2002 - BMIDS00329 - Use Cached Versions */%>
		frmScreen.txtPayment.value = top.frames[1].document.all.scMathFunctions.RoundValue(curTotal,2);
		
		nNumPIComponents = XML.XMLDocument.getElementsByTagName("PRODUCTANDLOANDETAILS").length;
		setupListArray();
	//}
		scScrollTable.initialiseTable(tblMortSummary, 1, "", ShowList, iTableSize, nNumPIComponents);
		ShowList(0);
	}

}

function setupListArray()
{
	<%/* This function takes the XMl for PRODUCTSANDLOANS and creates an array with this information.
		The last entry in the array is the summary line for the table. */ %>

		arrList = new Array(nNumPIComponents);
		var nInstallmentTotal = 0, nArrearsTotal = 0, nBalanceTotal = 0;
		if(nNumPIComponents == 0)
		{
			arrRow[0] = "";
			arrRow[1] = "";
			arrRow[2] = "";
			arrRow[3] = "";
			arrRow[4] = "";
			arrRow[5] = "";
			arrRow[6] = "";
			arrRow[7] = "";
		}
		else
		{	
			var iPIComponent, strTemp;
			var txtProduct, txtType, txtIntRate, txtProdEnd, txtLoanEnd, txtInstallment, txtRep, txtArrears, txtLoan, txtBalance;
			for(iPIComponent=0;iPIComponent < nNumPIComponents;iPIComponent++)
			{
				thisPIComponent = XML.XMLDocument.getElementsByTagName("PRODUCTANDLOANDETAILS").item(iPIComponent);

				nInstallmentTotal += parseFloat(thisPIComponent.getAttribute("REPAYMENTINSTALLMENT"));
				nArrearsTotal += parseFloat(thisPIComponent.getAttribute("AMOUNTOFARREARS"));
				nBalanceTotal += parseFloat(thisPIComponent.getAttribute("OUTSTANDINGBALANCE"));
				arrRow = new Array(7);		
				arrRow[0] = SpaceFill(thisPIComponent.getAttribute("PRODUCTDESCRIPTION"))
				arrRow[1] = SpaceFill(thisPIComponent.getAttribute("PRODUCTTYPE"))
				xmlRepaymentTypeCombo.LoadXML(m_sXMLRepaymentTypes);
				xmlRepaymentTypeCombo.SelectNodes(".//LISTENTRY")
				var intRPIndex =0;
				while ( intRPIndex <  xmlRepaymentTypeCombo.ActiveTagList.length) 
				{

					xmlRepaymentTypeCombo.SelectTagListItem(intRPIndex);
					if (xmlRepaymentTypeCombo.ActiveTag.selectSingleNode("VALUEID").text == arrRow[1])
					{					
						arrRow[1] = xmlRepaymentTypeCombo.ActiveTag.selectSingleNode("VALUENAME").text;
					}
					intRPIndex++;
				}
				
				<%/* MV - 13/08/2002 - BMIDS00329 - Use Cached Versions */%>
				strTemp	  = SpaceFill(top.frames[1].document.all.scMathFunctions.RoundValue(thisPIComponent.getAttribute("INTERESTRATE"),2));
				
				if (strTemp != " ")
				{
					strTemp = strTemp + "%";
				}
				arrRow[2] = strTemp;
				arrRow[3] = SpaceFill(thisPIComponent.getAttribute("PRODUCTENDDATE"));
				arrRow[4] = SpaceFill(thisPIComponent.getAttribute("LOANENDDATE"));
				arrRow[5] = SpaceFill(scScreenFunctions.FormatAsCurrency(top.frames[1].document.all.scMathFunctions.RoundValue(parseFloat(thisPIComponent.getAttribute("REPAYMENTINSTALLMENT")),2)));
				arrRow[6] = SpaceFill(scScreenFunctions.FormatAsCurrency(top.frames[1].document.all.scMathFunctions.RoundValue(parseFloat(thisPIComponent.getAttribute("AMOUNTOFARREARS")),2)));
				arrRow[7] = SpaceFill(scScreenFunctions.FormatAsCurrency(top.frames[1].document.all.scMathFunctions.RoundValue(parseFloat(thisPIComponent.getAttribute("OUTSTANDINGBALANCE")),2)));
				arrList[iPIComponent] = arrRow;
			}			
		}
		<%
		/*
		arrRow = new Array(8);
		arrRow[0] = "&nbsp;";
		arrRow[1] = "";
		arrRow[2] = "";
		arrRow[3] = "<B>BALANCES</B>";
		arrRow[4] = "";
		arrRow[5] = "<B>" + scScreenFunctions.FormatAsCurrency(top.frames[1].document.all.scMathFunctions.RoundValue(nInstallmentTotal,2)) + "</B>";
		arrRow[6] = "";
		arrRow[7] = "<B>" + scScreenFunctions.FormatAsCurrency(top.frames[1].document.all.scMathFunctions.RoundValue(parseFloat(nArrearsTotal),2)) + "</B>";
		arrRow[8] = "<B>" + scScreenFunctions.FormatAsCurrency(top.frames[1].document.all.scMathFunctions.RoundValue(parseFloat(nBalanceTotal),2)) + "</B>";
		arrList[iPIComponent] = arrRow;
		*/
		%>
		m_nInstallmentTotal = nInstallmentTotal;
		m_nArrearsTotal = nArrearsTotal;
		m_nBalanceTotal = nBalanceTotal;
		
}
function SpaceFill(strValue)
{
	if (strValue == "") 
	{
		return(" ");
	} else
	{
		return(strValue);
	}
}



function ShowList(nScrollIndex)
{
		scScrollTable.clear();
		if(nNumPIComponents == 0)
		{
			tblMortSummary.rows(1).cells(1).innerText = "";
			tblMortSummary.rows(1).cells(2).innerText = "";
			tblMortSummary.rows(1).cells(3).innerText = "";
			tblMortSummary.rows(1).cells(4).innerText = "";
			tblMortSummary.rows(1).cells(5).innerText = "";
			tblMortSummary.rows(1).cells(6).innerText = "";
		}
		else
		{	
			var iPIComponent;
			for(iPIComponent= 0 ; iPIComponent+nScrollIndex < nNumPIComponents && iPIComponent < iTableSize ;iPIComponent++)
			{
				arrRow = arrList[iPIComponent + nScrollIndex];
				scScreenFunctions.SizeTextToField(tblMortSummary.rows(iPIComponent+1).cells(0), arrRow[0]);					
				scScreenFunctions.SizeTextToField(tblMortSummary.rows(iPIComponent+1).cells(1), arrRow[1]);
				scScreenFunctions.SizeTextToField(tblMortSummary.rows(iPIComponent+1).cells(2), arrRow[2]);
				scScreenFunctions.SizeTextToField(tblMortSummary.rows(iPIComponent+1).cells(3), arrRow[3]);
				scScreenFunctions.SizeTextToField(tblMortSummary.rows(iPIComponent+1).cells(4), arrRow[4]);
				scScreenFunctions.SizeTextToField(tblMortSummary.rows(iPIComponent+1).cells(5), arrRow[5]);
				scScreenFunctions.SizeTextToField(tblMortSummary.rows(iPIComponent+1).cells(6), arrRow[6]);
				scScreenFunctions.SizeTextToField(tblMortSummary.rows(iPIComponent+1).cells(7), arrRow[7]);
			}	

		}
		<% /* SUMMARY ROW - SYS4468 */ %>
		<%/* MV - 13/08/2002 - BMIDS00329 - USe Cached Versions */ %>
		frmScreen.txtInstallmentSum.value	= scScreenFunctions.FormatAsCurrency(top.frames[1].document.all.scMathFunctions.RoundValue(m_nInstallmentTotal,2));
		frmScreen.txtArearsSum.value		= scScreenFunctions.FormatAsCurrency(top.frames[1].document.all.scMathFunctions.RoundValue(parseFloat(m_nArrearsTotal),2));
		frmScreen.txtLoanBalanceSum.value	= scScreenFunctions.FormatAsCurrency(top.frames[1].document.all.scMathFunctions.RoundValue(parseFloat(m_nBalanceTotal),2));



}

function RetrieveContextData()
{
	<% /* m_sOptimusSessionId = scScreenFunctions.GetContextParameter(window,"idOptimusSessionId");
	m_optimusSessionEffectiveDate = scScreenFunctions.GetContextParameter(window,"optimusSessionEffectiveDate"); */ %>

	m_sOtherSystemCustomerNumber = scScreenFunctions.GetContextParameter(window,"idOtherSystemCustomerNumber");
	m_sOtherSystemAccountNumber = scScreenFunctions.GetContextParameter(window,"idOtherSystemAccountNumber");
	<% /* PSC 18/10/2005 MAR57 */%>
	m_sCallingScreen = scScreenFunctions.GetContextParameter(window,"idCallingScreenID",null);

}

function btnSubmit.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idOtherSystemAccountNumber", "");
	<% /* PSC 18/10/2005 MAR57 - Start */ %>
    if (m_sCallingScreen == "CR041")
		frmToCR041.submit();
	else
		frmToCR040.submit();
	<% /* PSC 18/10/2005 MAR57 - End */ %>
}

function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idOtherSystemAccountNumber", "");
	<% /* PSC 18/10/2005 MAR57 - Start */ %>
    if (m_sCallingScreen == "CR041")
		frmToCR041.submit();
	else
		frmToCR040.submit();
	<% /* PSC 18/10/2005 MAR57 - End */ %>
}
-->
</script>
</body>
</html>



