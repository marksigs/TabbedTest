<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      DC226.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Buy To Let Pop-up
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SG		04/03/02	Created
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
GD		05/07/2002	BMIDS00165 'Buy To Let' functionality added, plus applied performance enhancement AQR
GD		22/07/2002	BMIDS00230 Force FlagChange(true) in btnCalculate
TW      09/10/2002  SYS5115 Modified to incorporate client validation  
GHun	11/11/2002	BMIDS00903 Disable OK until calculate has been pressed
MV		06/12/2002	BMIDS0145 Amended PopulateScreen()
MC		21/04/2004	BMIDS517	Background div size changed to adjust title text (with IE 6.0 Display Bigfont)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
EPSOM 2 History

Prog	Date		Description
PE		13/01/2007	EP2_825  Error leaving screen if you do not complete question if buy to let.. at bottom of screen
INR		22/01/2007	EP2_677 Removed relative to occupy question.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
*/
%>
<HEAD>
<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Let Properties - Tenant & Rent Details <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange">
<div id="divBackground"  style="TOP: 10px; LEFT: 12px; HEIGHT: 260px; WIDTH: 400px; POSITION: ABSOLUTE" class="msgGroup">

	<span style="LEFT: 10px; POSITION: absolute; TOP: 15px" class="msgLabel" >	
		Number of Occupants
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtNumberOfOccupants"  style="WIDTH: 100px"  class="msgTxt" maxlength="2">
		</span>
	</span>
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 45px" class="msgLabel">
		Tenancy Type
		<span style="LEFT: 160px; POSITION: absolute; TOP: 0px">
			<select id="cboTenancy" style="WIDTH: 130px" class="msgCombo"></select>
		</span>
	</span>

	<span style="LEFT: 10px; POSITION: absolute; TOP: 75px" class="msgLabel" >	
		Monthly Rental Income
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 5px" class="msgCurrency"></label>
			<input id="txtMonthlyRentalIncome"  style="WIDTH: 100px"  class="msgTxt" maxlength="9">
		</span>
	</span>
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 105px" class="msgLabel">
		Rental Income Status
		<span style="LEFT: 160px; POSITION: absolute; TOP: 0px">
			<select id="cboRentalIncomeStatus" style="WIDTH: 130px" class="msgCombo"></select>
		</span>
	</span>
<!--GD BMIDS00165 start-->	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 135px" class="msgLabel" >	
		Requested Loan Amount
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 5px" class="msgCurrency"></label>
			<input id="txtRequestedLoanAmount"  style="WIDTH: 100px"  class="msgReadOnly" maxlength="9" readonly tabindex = -1>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 165px" class="msgLabel" >	
		Excess Monthly Rental Income
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 5px" class="msgCurrency"></label>
			<input id="txtExcessMonthlyRentalIncome"  style="WIDTH: 100px"  class="msgReadOnly" maxlength="9" readonly tabindex = -1>
		</span>
	</span>
	
	<span style="TOP: 162px; LEFT: 283px; POSITION: ABSOLUTE">
		<input id="btnCalculate" value="Calculate" type="button" style="WIDTH: 60px" class="msgButton">
	</span>	
<!--GD BMIDS00165 end-->		
</div>
	<span style="TOP: 220px; LEFT: 16px; POSITION: ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
		<span style="TOP: 0px; LEFT: 76px; POSITION: ABSOLUTE">
			<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>
</form>


<!-- #include FILE="attribs/DC226attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var XML = null;
var m_sXML = "";
var m_sMetaAction = "";
var m_BaseNonPopupWindow = null;
var m_CalcXML = null;
var m_sCalcXML = "";
var m_sRequestAttributes;
var m_sParameters;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();

	RetrieveData();	
	SetMasks();
	Validation_Init();	
	Initialise();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}
function RetrieveData()
{
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();

	//XML = new scXMLFunctions.XMLObject();

	var sArguments		= window.dialogArguments;
	window.dialogTop	= sArguments[0];	
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	
	m_sParameters		= sArguments[4];
	m_sRequestAttributes = m_sParameters[5];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();	
	m_CalcXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_sCalcXML = m_sParameters[4];
	m_CalcXML.LoadXML(m_sCalcXML);
	m_sReadOnly			= m_sParameters[0];
	XML					= m_sParameters[1];
	scScreenFunctions.SetThisCurrency(frmScreen,m_sParameters[2]);

}
function frmScreen.btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}
function frmScreen.btnOK.onclick()
{
	var sReturn = new Array();
	sReturn[0] = IsChanged();
	
	if(sReturn[0]) SaveBuyToLetDetails();	
	
	window.returnValue	= sReturn;
	window.close();
}
function frmScreen.btnCalculate.onclick()
{
	//var sRequestedLoanAmount = frmScreen.txtRequestedLoanAmount.value;
	var sMonthlyRentalIncome = frmScreen.txtMonthlyRentalIncome.value;
	if(sMonthlyRentalIncome == "")
	{
		alert("Monthly Rental Income must be present in order to perform the calculation.");
		return;
	}
	FlagChange(true);
	//Re-initialise CalcXML each time Calculate button is pressed
	m_CalcXML.LoadXML(m_sCalcXML);
	//Add REQUESTEDLOANAMOUNT and MONTHLYRENTALINCOME to RENTALDETAILS tag
	m_CalcXML.SelectTag(null,"RENTALDETAILS");
	m_CalcXML.CreateTag("REQUESTEDLOANAMOUNT", frmScreen.txtRequestedLoanAmount.value);
	m_CalcXML.CreateTag("MONTHLYRENTALINCOME", sMonthlyRentalIncome);
	// 	m_CalcXML.RunASP(document,"CalculateExcessRentalIncome.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			m_CalcXML.RunASP(document,"CalculateExcessRentalIncome.asp");
			break;
		default: // Error
			m_CalcXML.SetErrorResponse();
		}

	if (m_CalcXML.IsResponseOK())
	{
		m_CalcXML.SelectTag(null,"RENTALDETAILS");
		frmScreen.txtExcessMonthlyRentalIncome.value = m_CalcXML.GetTagText("EXCESSMONTHLYRENTALINCOME");
		<% /* BMIDS00903 Re-enable the OK button */ %>
		frmScreen.btnOK.disabled = false;
	}
}



function Initialise()
{
	PopulateCombos();
	PopulateScreen();
	
	if(m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnCalculate.disabled = true;
	}
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	<% /* BMIDS00903 Disable OK until calculate has been pressed */ %>
	frmScreen.btnOK.disabled = true;
}
function PopulateCombos()
{
	var bSuccess = true;
	//var ComboXML = new scXMLFunctions.XMLObject();
	var ComboXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("RentalIncomeStatus","BuyToLetTenancyType");

	if(ComboXML.GetComboLists(document,sGroupList))
	{
		bSuccess = true;
		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboTenancy,"BuyToLetTenancyType",true);
		bSuccess = bSuccess & ComboXML.PopulateCombo(document,frmScreen.cboRentalIncomeStatus,"RentalIncomeStatus",true);

		if(!bSuccess) scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
}
function PopulateScreen()
{
	var sApplicationNumber, sApplicationFactFindNumber;
	XML.SelectTag(null,"NEWPROPERTY")
	frmScreen.txtNumberOfOccupants.value = XML.GetTagText("NUMBEROFOCCUPANTS");
	frmScreen.cboTenancy.value = XML.GetTagText("TENANCYTYPE");
	frmScreen.txtMonthlyRentalIncome.value = XML.GetTagText("MONTHLYRENTALINCOME");		
	frmScreen.cboRentalIncomeStatus.value = XML.GetTagText("RENTALINCOMESTATUS");

	
	m_CalcXML.LoadXML(m_sCalcXML);
	m_CalcXML.SelectTag(null,"REQUEST/NEWPROPERTY/RENTALDETAILS")
	sApplicationNumber = m_CalcXML.GetTagText("APPLICATIONNUMBER");
	sApplicationFactFindNumber = m_CalcXML.GetTagText("APPLICATIONFACTFINDNUMBER");
	var xmlQuoteData = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	xmlQuoteData.CreateRequestTagFromArray(m_sRequestAttributes, null);
	xmlQuoteData.CreateActiveTag("APPLICATION");
	xmlQuoteData.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	xmlQuoteData.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	xmlQuoteData.RunASP(document,"GetAcceptedOrActiveQuoteData.asp");	
	
	if(xmlQuoteData.IsResponseOK())
	{
		var xmlTempTag = xmlQuoteData.SelectTag(null,"MORTGAGESUBQUOTE")
		if (xmlTempTag != null )
		{
			frmScreen.txtRequestedLoanAmount.value = xmlQuoteData.GetTagText("AMOUNTREQUESTED");
			frmScreen.btnCalculate.disabled = false;
		}
		else
			frmScreen.btnCalculate.disabled = true;
	} 
	frmScreen.txtExcessMonthlyRentalIncome.value = XML.GetTagText("EXCESSMONTHLYRENTALINCOME");
}
function SaveBuyToLetDetails()
{
	XML.SelectTag(null,"NEWPROPERTY")
	XML.SetTagText("NUMBEROFOCCUPANTS", frmScreen.txtNumberOfOccupants.value);
	XML.SetTagText("TENANCYTYPE", frmScreen.cboTenancy.value);
	XML.SetTagText("MONTHLYRENTALINCOME", frmScreen.txtMonthlyRentalIncome.value);	
	XML.SetTagText("RENTALINCOMESTATUS", frmScreen.cboRentalIncomeStatus.value);
	//GD BMIDS00165 start
	XML.SetTagText("REQUESTEDLOANAMOUNT", frmScreen.txtRequestedLoanAmount.value);	
	XML.SetTagText("EXCESSMONTHLYRENTALINCOME", frmScreen.txtExcessMonthlyRentalIncome.value);
	//GD BMIDS00165 end	

	
}
-->
</script>
</body>
</html>
