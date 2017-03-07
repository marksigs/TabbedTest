<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      mn010.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		14/02/00	msgButtons removed
AY		17/03/00	MetaAction changed to ApplicationMode
AY		03/04/00	New top menu/scScreenFunctions change
BG		10/05/2000	SYS0711 - Changed label on Customer Registration button and added new Pipeline Enquiry button
GD		05/12/2000	Task Management button added
CL		05/03/01	SYS1920 Read only functionality added

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:

Prog	Date		Description
PSC		11/10/2005	MAR57	Set customer search button based on global parameter
HMA		07/12/2005  MAR797  Change global parameter used for Customer Search button.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
<!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<script src="validation.js" language="JScript"></script>
<!-- Table Control Object -->
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex="-1" type="text/x-scriptlet" VIEWASTEXT></object>
<!-- Specify Forms Here -->
<form id="frmMortgageCalculator" method="post" action="mc010.asp" STYLE="DISPLAY: none"></form>
<form id="frmCustomerSearch" method="post" action="cr010.asp" STYLE="DISPLAY: none"></form>
<form id="frmApplicationQuickSearch" method="post" action="gn200.asp" STYLE="DISPLAY: none"></form>
<form id="frmTaskManagement" method="post" action="mn015.asp" STYLE="DISPLAY: none"></form>
<form id="frmPaymentSanction" method="post" action="SP090.asp" STYLE="DISPLAY: none"></form>

<% /* <form id="frmPipeLineEnquiry" method="post" action="  .asp" STYLE="DISPLAY: none"></form> */%>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange" >
<div id="divBackground" style="HEIGHT: 475px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 0px; POSITION: absolute; TOP: 40px" class="msgLabel">				
		
		<span id="spnCustomerEngagement" style="LEFT: 10px; POSITION: absolute; TOP: 0px" class="msgLabel">					
			<img alt border="0" id="imgCustomerEngagement1" src="images/cust_search_1.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 2" WIDTH="200" HEIGHT="100">				
			<img alt border="0" id="imgCustomerEngagement2" src="images/cust_search_2.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 1" WIDTH="200" HEIGHT="100"> 
			<img alt border="0" id="imgCustomerEngagement3" src="images/cust_search_3.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 0" WIDTH="200" HEIGHT="100"> 
		</span>
		
		<span id="spnPipeLineEnquiry" style="LEFT: 220px; POSITION: absolute; TOP: 0px" class="msgLabel">					
			<img alt border="0" id="imgPipeLineEnquiry1" src="images/app_enquiry_1.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 2" WIDTH="200" HEIGHT="100">					
			<img alt border="0" id="imgPipeLineEnquiry2" src="images/app_enquiry_2.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 1" WIDTH="200" HEIGHT="100"> 						
		</span>
				
		<span id="spnPaymentSanction" style="LEFT: 220px; POSITION: absolute; TOP: 130px" class="msgLabel">					
			<img alt border="0" id="imgPaymentSanction1" src="images/payment_sanction_1.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 2" WIDTH="200" HEIGHT="100">				
			<img alt border="0" id="imgPaymentSanction2" src="images/payment_sanction_2.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 1" WIDTH="200" HEIGHT="100"> 
		</span>
				
		<span id="spnTaskManagement" style="LEFT: 10px; POSITION: absolute; TOP: 130px" class="msgLabel">					
			<img alt border="0" id="imgTaskManagement1" src="images/task_man_1.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 2" WIDTH="200" HEIGHT="100">					
			<img alt border="0" id="imgTaskManagement2" src="images/task_man_2.jpg" style="LEFT: 100px; POSITION: absolute; TOP: 10px; Z-INDEX: 1" WIDTH="200" HEIGHT="100"> 						
		</span>
	
	</span> 
</div> 
</form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<!-- Specify Code Here -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/mn010Attribs.asp" -->
<script language="JScript">
<!--
var m_sMetaAction		= "";
var scScreenFunctions;
var m_blnReadOnly = false;



<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Welcome Menu","MN010",scScreenFunctions);
	RetrieveContextData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
<%	/* Initialise of the MultipleLender flag in context
	   this code should really be place in a one time initialisation
	   routine like just after logon */
%>	SetMultipleLenderInContext();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "MN010")
	
	<% /* PSC 11/10/2005 MAR57 - Start */ %>
	if (!m_blnReadOnly)
		SetCustomerSearchState();
	<% /* PSC 11/10/2005 MAR57 - End */ %>
	
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
	m_sMetaAction		= scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
}
function SetMultipleLenderInContext()
{
<% /* if the multiple lender flag is blank then we need to retrieve the
      parameter from the database*/
%>	var sMultipleLender	= scScreenFunctions.GetContextParameter(window,"idMultipleLender",null);
	if (sMultipleLender == "")
	{			
		<% /* MV - 08/07/2002 - BMIDS00188 - Core Upgrade Code Error */ %>
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.RunASP(document, "GetMultipleLender.asp");
		if (XML.IsResponseOK() == true)
		{
			sMultipleLender = XML.GetTagText("MULTIPLELENDER");
			scScreenFunctions.SetContextParameter(window, "idMultipleLender", sMultipleLender);
		}
		XML = null;
	}
}
function SwapImage(imgImage)
{
	imgImage.style.cursor="hand";
	imgImage.style.zIndex="3";
}
function RestoreImage(imgImage)
{
	imgImage.style.zIndex="1";
}
function GoCustomerSearch()
{
	frmCustomerSearch.submit();
}
function GoTaskManagement()
{
	frmTaskManagement.submit();
}
function GoMortgageCalculator()
{
	scScreenFunctions.SetContextParameter(window,"idApplicationMode","Mortgage Calculator");
	frmMortgageCalculator.submit();
}

function GoPaymentSanction()
{
	frmPaymentSanction.submit();
}

//*********************************

function frmScreen.imgCustomerEngagement1.onmouseover()  // customer registration
{
	SwapImage(frmScreen.imgCustomerEngagement2);
}
function frmScreen.imgCustomerEngagement2.onmouseout()
{
	RestoreImage(frmScreen.imgCustomerEngagement2);
}
function frmScreen.imgCustomerEngagement2.onclick()
{
	GoCustomerSearch();
}

//**********************************

function frmScreen.imgTaskManagement1.onmouseover() // task manager
{
	SwapImage(frmScreen.imgTaskManagement2);
}
function frmScreen.imgTaskManagement2.onmouseout()
{
	RestoreImage(frmScreen.imgTaskManagement2);
}
function frmScreen.imgTaskManagement2.onclick()
{
	GoTaskManagement();
}

//***************************************

function frmScreen.imgPipeLineEnquiry1.onmouseover() //application enquiry
{
	SwapImage(frmScreen.imgPipeLineEnquiry2);
}
function frmScreen.imgPipeLineEnquiry2.onmouseout()
{
	RestoreImage(frmScreen.imgPipeLineEnquiry2);
}
function frmScreen.imgPipeLineEnquiry2.onclick()
{
	//window.alert("Pipeline Enquiry to be developed")
	frmApplicationQuickSearch.submit();
}

//*********************************************

function frmScreen.imgPaymentSanction1.onmouseover() // Payment sanction
{
	SwapImage(frmScreen.imgPaymentSanction2);
}
function frmScreen.imgPaymentSanction2.onmouseout()
{
	RestoreImage(frmScreen.imgPaymentSanction2);
}
function frmScreen.imgPaymentSanction2.onclick()
{
	GoPaymentSanction();
}
<% /* PSC 11/10/2005 MAR57 - Start */ %>
function SetCustomerSearchState()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bDisableCustomerSearch = XML.GetGlobalParameterBoolean(document,"DisableCustSearchButton");   // MAR797
	if 	(bDisableCustomerSearch)
		frmScreen.imgCustomerEngagement3.style.zIndex="3";
}
<% /* PSC 11/10/2005 MAR57 - End */ %>

-->
</script>
</body>
</html>




