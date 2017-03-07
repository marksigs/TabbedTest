<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      cr025.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Credit Check - CAIS Summary Popup Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MDC		17/03/2000	Original
AY		03/04/00	scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MC		20/04/2004	BMIDS517	White space padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date			Description
TW      09 Oct 2002		Modified to incorporate client validation - SYS5115
MDC		28/10/2002		BMIDS00733 - Remove line not required that was added by Screen Rules process
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Credit Check CAIS Summary  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" style="VISIBILITY: hidden">
<div id="divBackground" style="TOP: 0px; LEFT: 0px; HEIGHT: 275px; WIDTH: 370px; POSITION: ABSOLUTE" class="msgGroup">
	
	<span style="TOP: 20px; LEFT: 20px; POSITION: ABSOLUTE" class="msgLabel">
		Customer Name
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtName" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly">
		</span>
	</span>
	
	<div style="TOP: 50px; LEFT: 10px; HEIGHT: 170px; WIDTH: 350px; POSITION: ABSOLUTE" class="msgGroupLight">
	
		<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>CAIS Summary</strong>
		</span>
	
		<span style="TOP: 15px; LEFT: 153px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Current</strong>
		</span>
	
		<span style="TOP: 15px; LEFT: 235px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Total Number</strong>
		</span>
	
		<span style="TOP: 40px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			3 Months CAIS
			<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
				<input id="txtCurrentCAIS3M" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
			<span style="TOP: -3px; LEFT: 230px; POSITION: ABSOLUTE">
				<input id="txtTotalCAIS3M" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>
	
		<span style="TOP: 80px; LEFT: 153px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Current</strong>
		</span>
	
		<span style="TOP: 80px; LEFT: 253px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>&nbspWorst</strong>
		</span>
	
		<span style="TOP: 110px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			4 Months CAIS
			<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
				<input id="txtCurrentCAIS4M" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
			<span style="TOP: -3px; LEFT: 230px; POSITION: ABSOLUTE">
				<input id="txtWorstCAIS4M" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>
	
		<span style="TOP: 135px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			12 Months CAIS
			<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
				<input id="txtCurrentCAIS12M" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
			<span style="TOP: -3px; LEFT: 230px; POSITION: ABSOLUTE">
				<input id="txtWorstCAIS12M" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly">
			</span>
		</span>
	</div>
	
	<span style="TOP: 235px; LEFT: 150px; POSITION: ABSOLUTE">
		<input id="btnOK" value="OK" type="button" class="msgButton" style="WIDTH: 60px">
	</span>

</div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/ra025Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_sCustomerNumber = null;
var m_sCustomerVersion = null;
var m_nPersonType = 0;
var m_sCreditCheckGUID = null;
var m_CreditCheckXML = null;
var m_sCustomerName = null;
var scScreenFunctions;
var m_BaseNonPopupWindow = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	<% /* BMIDS00733 MDC 28/10/2002 - Remove line not required that was added by Screen Rules process 
	Validation_Init();
	*/ %>
	
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();	
	m_sCustomerNumber = sArgArray[0];
	m_sCustomerVersion = sArgArray[1];
	m_nPersonType = sArgArray[2];
	m_sCreditCheckGUID = sArgArray[3];
	m_sCustomerName = sArgArray[4];
	
	if (GetCAISSummary())
	{
		scScreenFunctions.ShowCollection(frmScreen);
		PopulateScreen();
	}
	else
		scScreenFunctions.ShowCollection(frmScreen);
	
	SetAllFieldsToReadOnly();
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function SetAllFieldsToReadOnly()
{
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCurrentCAIS3M");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTotalCAIS3M");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCurrentCAIS4M");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtWorstCAIS4M");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCurrentCAIS12M");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtWorstCAIS12M");
}

function GetCAISSummary()
{
var bSuccess = false;

	<% /* Retrieve CAIS Summary data */ %>
	
	m_CreditCheckXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	with (m_CreditCheckXML)
	{
		<% /* Create request block */ %>
		<% /* CreateRequestTag(window, "SEARCH"); */ %>
		CreateActiveTag("REQUEST");
		CreateActiveTag("SEARCH");
		CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
		CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersion);
		CreateTag("PERSONTYPE", m_nPersonType);
		CreateTag("CREDITCHECKGUID", m_sCreditCheckGUID);
		RunASP(document, "GetCreditCheckCAIS.asp");
		if (IsResponseOK())
			bSuccess = true;
	}
	
	return bSuccess;
}

function PopulateScreen()
{
	with (frmScreen)
	{
		txtName.value = m_sCustomerName
		txtCurrentCAIS3M.value = m_CreditCheckXML.GetTagText("TOTALVALUEACTIVECAIS3M");
		txtTotalCAIS3M.value =  m_CreditCheckXML.GetTagText("TOTALACTIVECAIS3M");
		txtCurrentCAIS4M.value =  m_CreditCheckXML.GetTagText("TOTALVALUEACTIVECAIS4M");
		txtWorstCAIS4M.value =  m_CreditCheckXML.GetTagText("WORSTSTATUSACTIVECAIS4M");
		txtCurrentCAIS12M.value =  m_CreditCheckXML.GetTagText("TOTALVALUEACTIVECAIS12M");
		txtWorstCAIS12M.value =  m_CreditCheckXML.GetTagText("WORSTSTATUSACTIVECAIS12M");
	}

}

function frmScreen.btnOK.onclick()
{
	window.close();
}

-->
</script>
</body>
</html>





