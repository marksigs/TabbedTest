<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      CM101.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Interest Rate Manual Adjustment Percent
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		25/03/02	CREATED
JLD		13/04/02	MSMS0056 - call SetMasks to check the attribs.
JD		20/07/04	BMIDS763 use applicationdate
PSC		19/02/2007	EP2_1488 Use ORIGINALLTV
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MC		06/05/2005	BMIDS571 White Space added to page title
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Manual Adjustment Percent <!-- #include file="includes/TitleWhiteSpace.asp" -->       </title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scMathFunctions.asp" height="1px" id="scMathFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange" >
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 150px; WIDTH: 300px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:40px; LEFT:14px; POSITION:ABSOLUTE" class="msgLabel">
	Initial Interest Rate
	<span style="TOP:-3px; LEFT:150px; POSITION:ABSOLUTE">
		<input id="txtRate" maxlength="5" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:70px; LEFT:14px; POSITION:ABSOLUTE" class="msgLabel">
	Interest Rate Adjustment
	<span style="TOP:-3px; LEFT:150px; POSITION:ABSOLUTE">
		<input id="txtAdjustment" maxlength="6" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:120px; LEFT:14px; POSITION:ABSOLUTE">
	<input id="btnOK" value="OK" type="button" style="WIDTH:60px" class="msgButton">
</span>
<span style="TOP:120px; LEFT:80px; POSITION:ABSOLUTE">
	<input id="btnCancel" value="Cancel" type="button" style="WIDTH:60px" class="msgButton">
</span>
</div>
</form>

<!--  #include FILE="attribs/cm101attribs.asp" -->
<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_sApplicationMode = "";
var m_sInterestRate = "";
var m_sAdjustment = "";
var m_sLoanComponent = "";
var m_sApplicationInfo = "";
var m_sRequestAttributes = "";
var m_xmlLoanComponent = null;
var m_xmlAppInfo = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	RetrieveData();
	SetMasks();
	Validation_Init();
	SetScreenOnReadOnly();
	window.returnValue = null;
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}
function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
}
function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters		= sArguments[4];
	m_sReadOnly			= sParameters[0];
	m_sInterestRate		= sParameters[1];
	m_sAdjustment		= sParameters[2];
	m_sLoanComponent	= sParameters[3];
	m_sApplicationInfo  = sParameters[4];
	m_sApplicationMode	= sParameters[5];
	m_sRequestAttributes = sParameters[6];
}
function PopulateScreen()
{
	m_xmlLoanComponent = new scXMLFunctions.XMLObject();
	m_xmlLoanComponent.LoadXML(m_sLoanComponent);
	m_xmlAppInfo = new scXMLFunctions.XMLObject();
	m_xmlAppInfo.LoadXML(m_sApplicationInfo);
	scScreenFunctions.SetFieldState(frmScreen, "txtRate", "R");
	frmScreen.txtAdjustment.value = m_sAdjustment;
	frmScreen.txtRate.value = m_sInterestRate;
	frmScreen.txtAdjustment.focus();
}
function ValidateAdjustment()
{
	if(frmScreen.txtAdjustment.value.length == 0)
	{
		alert("Please enter a rate adjustment");
		return false;
	}
	if(isNaN(parseFloat(frmScreen.txtAdjustment.value)))
	{
		alert("Please enter a numeric value.");
		frmScreen.txtAdjustment.focus();
		return false;
	}
	//check for maximum 2dp
	var sAdjustment = frmScreen.txtAdjustment.value;
	var stringArray = sAdjustment.split(".");
	if(stringArray.length > 1)
	{
		var sPostDP = stringArray[1];
		if(sPostDP.length > 2)
		{
			alert("Rate adjustment allows a maximum of 2 decimal places.");
			frmScreen.txtAdjustment.focus();
			return false;
		}
	}
	
	var adjustment = parseFloat(frmScreen.txtAdjustment.value);
	m_xmlLoanComponent.SelectTag(null, "LOANCOMPONENT");
	m_xmlAppInfo.SelectTag(null, "RESPONSE");
	var validateXML = new scXMLFunctions.XMLObject();
	validateXML.CreateRequestTagFromArray(m_sRequestAttributes, "");
	validateXML.CreateActiveTag("VALIDATEDATA");
	validateXML.CreateTag("APPLICATIONNUMBER", m_xmlAppInfo.GetTagText("APPLICATIONNUMBER"));
	validateXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_xmlAppInfo.GetTagText("APPLICATIONFACTFINDNUMBER"));
	validateXML.CreateTag("APPLICATIONDATE", m_xmlAppInfo.GetTagText("APPLICATIONDATE")); // JD BMIDS763
	validateXML.CreateTag("MORTGAGESUBQUOTENUMBER", m_xmlLoanComponent.GetTagText("MORTGAGESUBQUOTENUMBER"));
	validateXML.CreateTag("MANUALADJUSTMENTPERCENT", adjustment);
	validateXML.CreateTag("MORTGAGEPRODUCTCODE", m_xmlLoanComponent.GetTagText("MORTGAGEPRODUCTCODE"));
	validateXML.CreateTag("STARTDATE", m_xmlLoanComponent.GetTagText("STARTDATE"));
	validateXML.CreateTag("LOANAMOUNT", m_xmlLoanComponent.GetTagText("LOANAMOUNT"));
	<% /* PSC 19/02/2007 EP2_1488 */ %>
	validateXML.CreateTag("ORIGINALLTV", m_xmlLoanComponent.GetTagText("ORIGINALLTV"));
	validateXML.CreateTag("LTV", m_xmlAppInfo.GetTagText("LTV"));
	validateXML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE", m_xmlAppInfo.GetTagText("PURCHASEPRICEORESTIMATEDVALUE"));
	validateXML.CreateTag("TYPEOFAPPLICATION", m_xmlAppInfo.GetTagText("TYPEOFAPPLICATION"));
	validateXML.CreateTag("LOCATION", m_xmlAppInfo.GetTagText("LOCATION"));
	validateXML.CreateTag("VALUATIONTYPE", m_xmlAppInfo.GetTagText("VALUATIONTYPE"));
	validateXML.CreateTag("LEGALFEETYPE", m_xmlAppInfo.GetTagText("LEGALFEETYPE"));
	var dblResolvedRate = parseFloat(m_sInterestRate) + parseFloat(frmScreen.txtAdjustment.value);
	validateXML.CreateTag("RESOLVEDRATE", scMathFunctions.RoundValue(dblResolvedRate,2));

	if(m_sApplicationMode == "Quick Quote")
		// 		validateXML.RunASP(document, "QQValidateManualAdjustment.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					validateXML.RunASP(document, "QQValidateManualAdjustment.asp");
				break;
			default: // Error
				validateXML.SetErrorResponse();
			}

	else
		validateXML.RunASP(document, "AQValidateManualAdjustment.asp");
		
	return(validateXML.IsResponseOK())
}
function frmScreen.btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}
function frmScreen.btnOK.onclick()
{
	var sReturn = new Array();
	if(ValidateAdjustment())
	{
		sReturn[0] = IsChanged();
		sReturn[1] = parseFloat(frmScreen.txtAdjustment.value);
		window.returnValue	= sReturn;
		window.close();
	}
	else frmScreen.txtAdjustment.focus();
}
-->
</script>
</body>
</html>


