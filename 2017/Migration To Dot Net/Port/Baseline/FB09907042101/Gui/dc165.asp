<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc165.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Tax Details Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MDC		31/05/2000	Original
MDC		13/06/2000	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
MDC		13/06/2000	SYS0926 Do Cancel processing if clicking on window close button
BG		25/08/00	SYS1264 Was not populating data on entry.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MV		22/08/2002	BMIDS00355	IE 5.5 upgrade Errors - Modified the HTML layout
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Tax Details  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<% /* Scriptlets - remove any which are not required */ %>
<script src="validation.js" language="JScript"></script>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" style="VISIBILITY: hidden">

<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 314px; WIDTH: 580px; POSITION: ABSOLUTE" class="msgGroup">

	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Tax Details</strong>
	</span>

	<span style="TOP: 40px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Customer Name
		<span style="TOP: -3px; LEFT: 180px; POSITION: ABSOLUTE">
			<input id="txtCustomerName" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly">
		</span>
	</span>

	<span style="TOP: 65px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Name of Tax Office
		<span style="TOP: -3px; LEFT: 180px; POSITION: ABSOLUTE">
			<input id="txtTaxOffice" maxlength="60" style="WIDTH: 320px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 90px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Tax Reference Number
		<span style="TOP: -3px; LEFT: 180px; POSITION: ABSOLUTE">
			<input id="txtTaxReferenceNumber" maxlength="50" style="WIDTH: 320px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 115px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		National Insurance Number
		<span style="TOP: -3px; LEFT: 180px; POSITION: ABSOLUTE">
			<input id="txtNationalInsuranceNumber" maxlength="9" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	 <span style="LEFT: 10px; POSITION: absolute; TOP: 140px" class="msgLabel">
		Do you self assess to IR?
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="optSelfAssessYes" name="IRSelfAssesGroup" type="radio" value="1">
		</span> 
		<span style="LEFT: 200px; POSITION: absolute">
			<label for='optTaxedOutsideUKYes"' class="msgLabel">Yes</label> 
		<span>
		<span style="LEFT: 30px; POSITION: absolute; TOP: -3px">
			<input id="optSelfAssessNo" name="IRSelfAssesGroup" type="radio" value="0">
		</span> 
		<span style="LEFT: 50px; POSITION: absolute">
			<label for="optSelfAssessNo" class="msgLabel">No</label> 
		</span> 
	</span>	
	<span style="LEFT: -200px; POSITION: absolute; TOP: 24px" class="msgLabel">
		Taxed outside the UK?
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="optTaxedOutsideUKYes" name="TaxedOutsideUKIndicator" type="radio" value="1">
		</span> 
		<span style="LEFT: 200px; POSITION: absolute">
			<label for='optTaxedOutsideUKYes"' class="msgLabel">Yes</label> 
		<span>
		<span style="LEFT: 30px; POSITION: absolute; TOP: -3px">
			<input id="optTaxedOutsideUKNo" name="TaxedOutsideUKIndicator" type="radio" value="0">
		</span> 
		<span style="LEFT: 50px; POSITION: absolute">
			<label for="optTaxedOutsideUKNo" class="msgLabel">No</label> 
		</span> 
	</span>
	<span style="LEFT=-199px;Position:absolute;TOP=25px">
		Non-UK Tax Arrangements 
	</span>
	<span style="LEFT=-20px;Position:absolute;TOP=25px">
		<TEXTAREA class=msgTxt id=txtNonUKTaxArrangements style="WIDTH: 320px" name=NonUKTaxArrangements rows=5></TEXTAREA>
	</span>
	<span style="LEFT: -199px; POSITION: absolute; TOP: 120px">
		<input id="btnSubmit" value="OK" type="button" class="msgButton" style="WIDTH: 60px">
	</span>
	<span style="LEFT: -130px; POSITION: absolute; TOP: 120px">
		<input id="btnCancel" value="Cancel" type="button" class="msgButton" style="WIDTH: 60px">
	</span>
	<%/* <span style="LEFT: 10px; POSITION: absolute; TOP: 140px" class="msgLabel">
		Taxed outside the UK?
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="optTaxedOutsideUKYes" name="TaxedOutsideUKIndicator" type="radio" value="1">
			<label for='optTaxedOutsideUKYes"' class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 234px; POSITION: absolute; TOP: -3px">
			<input id="optTaxedOutsideUKNo" name="TaxedOutsideUKIndicator" type="radio" value="0">
			<label for="optTaxedOutsideUKNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="TOP: 165px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Non-UK Tax Arrangements
		<span style="TOP: -3px; LEFT: 180px; POSITION: ABSOLUTE">
			<textarea id="txtNonUKTaxArrangements" name="NonUKTaxArrangements" rows="5"  style="WIDTH: 320px; POSITION: ABSOLUTE" class="msgTxt"></textarea>
		</span>
	</span>

	<span style="TOP: 250px; LEFT: 10px; POSITION: ABSOLUTE">
		<input id="btnSubmit" value="OK" type="button" class="msgButton" style="WIDTH: 60px">
	</span>
	<span style="TOP: 250px; LEFT: 80px; POSITION: ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" class="msgButton" style="WIDTH: 60px">
	</span> */%>

</div>


</form>
<!-- #include FILE="attribs/DC165Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var IncomeSummaryXML = null;
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sCustomerName = "";
var m_sXML = "";
var m_sReadOnly = "";
var scScreenFunctions;
var bOKClicked = false;
var m_BaseNonPopupWindow = null;


function frmScreen.txtNonUKTaxArrangements.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtNonUKTaxArrangements", 255, true);
}

function window.onload()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];	
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	RetrieveContextData(sArgArray);
	
	Initialise();
	SetMasks();
	Validation_Init();
	if(m_sReadOnly == '1')
		SetAllFieldsToReadOnly();
	else
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCustomerName");
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
		
	window.returnValue = null;

}

function window.onbeforeunload()
{
	if(bOKClicked == false)
	{
	frmScreen.btnCancel.onclick();	
	}
	else
	window.close();
}

function SetAllFieldsToReadOnly()
{
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCustomerName");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTaxOffice");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTaxReferenceNumber");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtNationalInsuranceNumber");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtNonUKTaxArrangements");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"optTaxedOutsideUKYes");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"optTaxedOutsideUKNo");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"optSelfAssessYes"); <% /* MAH 20/11/2006 E2CR35 */ %>
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"optSelfAssessNo"); <% /* MAH 20/11/2006 E2CR35 */ %>
}

function RetrieveContextData(sArgArray)
{
	m_sCustomerName = sArgArray[0];
	m_sCustomerNumber = sArgArray[1];
	m_sCustomerVersionNumber = sArgArray[2];
	m_sXML = sArgArray[3];
	m_sReadOnly = sArgArray[4];
}

function PopulateScreen()
{
var sTaxedOutsideUK = "";

	IncomeSummaryXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	IncomeSummaryXML.LoadXML(m_sXML);
	IncomeSummaryXML.SelectTag(null,"INCOMESUMMARY");
	with (frmScreen)
	{
		<% /* Customer Name */ %>
		txtCustomerName.value = m_sCustomerName;
		txtTaxOffice.value = IncomeSummaryXML.GetTagText("TAXOFFICE");
		txtTaxReferenceNumber.value = IncomeSummaryXML.GetTagText("TAXREFERENCENUMBER");
		txtNationalInsuranceNumber.value = IncomeSummaryXML.GetTagText("NATIONALINSURANCENUMBER");
		sTaxedOutsideUK = IncomeSummaryXML.GetTagText("TAXEDOUTSIDEUKINDICATOR");
		scScreenFunctions.SetRadioGroupValue(frmScreen, "TaxedOutsideUKIndicator", sTaxedOutsideUK);
		scScreenFunctions.SetRadioGroupValue(frmScreen, "IRSelfAssesGroup",IncomeSummaryXML.GetTagText("SELFASSESS"));<% /* MAH 20/11/2006 E2CR35 */ %>
		if(sTaxedOutsideUK == "1")
		{
			txtNonUKTaxArrangements.disabled = false;
			txtNonUKTaxArrangements.value = IncomeSummaryXML.GetTagText("NONUKTAXDETAILS");
		}
		else
		{
			txtNonUKTaxArrangements.value = "";
			txtNonUKTaxArrangements.disabled = true;
		}
		
	}

}

function Initialise()
{
	scScreenFunctions.ShowCollection(frmScreen);
	PopulateScreen();
}

function frmScreen.optTaxedOutsideUKNo.onclick()
{
	frmScreen.txtNonUKTaxArrangements.value = "";
	frmScreen.txtNonUKTaxArrangements.disabled = true;		
}

function frmScreen.optTaxedOutsideUKYes.onclick()
{
	frmScreen.txtNonUKTaxArrangements.disabled = false;		
}

function frmScreen.btnSubmit.onclick()
{
var sReturn = "";
var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	if(scScreenFunctions.RestrictLength(frmScreen, "txtNonUKTaxArrangements", 255, true))
		return;

	XML.CreateActiveTag("INCOMESUMMARY");
	if(IsChanged())
	{
		XML.SetAttribute("UPDATED", "YES");
		XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
		XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
		XML.CreateTag("TAXOFFICE",frmScreen.txtTaxOffice.value);
		XML.CreateTag("TAXREFERENCENUMBER",frmScreen.txtTaxReferenceNumber.value);
		XML.CreateTag("NATIONALINSURANCENUMBER",frmScreen.txtNationalInsuranceNumber.value);
		XML.CreateTag("NONUKTAXDETAILS",frmScreen.txtNonUKTaxArrangements.value);
		XML.CreateTag("TAXEDOUTSIDEUKINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"TaxedOutsideUKIndicator"));
		XML.CreateTag("SELFASSESS",scScreenFunctions.GetRadioGroupValue(frmScreen,"IRSelfAssesGroup"));<% /* MAH 20/11/2006 E2CR35 */ %>
	}
	else
	{
		XML.SetAttribute("UPDATED", "NO");
	}
	
	sReturn = XML.XMLDocument.xml
	
	bOKClicked = true; //BG 25/08/00 SYS1264 Used to stop the onbeforeunload event clearing 
					   //variables if window is closed with the window close button.
	
	window.returnValue = sReturn;
	window.close();
}

function frmScreen.btnCancel.onclick()
{
	IncomeSummaryXML = null;
	window.returnValue = "";
	window.close();
}

-->
</script>
</body>
</html>
