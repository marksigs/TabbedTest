<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      mn070.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
IVW		28/11/2000	Created
ADP		13/02/2001	Changed page title to "Application Processing menu" SYS1944
SR		13/02/2001  Call PP070 on clicking 'Payee' button
CL		13/02/2001	SYS1944 New Buttons and functionality added
AP		15/02/2001	SYS1944 New Buttons altered to correct size and position 
MC		05/03/2001	SYS1987 Enable Payee and Disbursement buttons
CL		05/03/01	SYS1920 Read only functionality added
JLD		06/03/01	SYS1879 Added routing to Application Summary and removed readonly fnc (unnecessary)
MV		19/03/01	SYS2049	setting the ReturnScreenId Context Parameter
ADP		20/03/01	SYS2078 Added new button for valuer instructions
SG		05/12/01	SYS3357 Removed Valuer Instructions button
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog    Date		Description
TW      09/10/2002	Modified to incorporate client validation - SYS5115
MV		19/03/2003	BM0371	- Amended frmScreen.imgPayees2.onclick();frmScreen.imgAppnCosts2.onclick();
					frmScreen.imgDisbursements2.onclick();frmScreen.imgAppnTransfer2.onclick()
					frmScreen.imgAppnSummary2.onclick();frmScreen.imgAppnStage2.onclick()
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>
<body><!-- Form Manager Object - Controls Soft Coded Field Attributes -->

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>
<!-- XML Functions Object - see comments within scXMLFunctions.htm for details -->
<!-- Table Control Object -->
<OBJECT data=scTable.htm height=1 id=scTable 
style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet 
VIEWASTEXT></OBJECT><!-- Specify Forms Here -->

<form id="frmApplicationStage" Method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>
<form id="frmApplicationSummary" Method="post" action="AP010.asp" STYLE="DISPLAY: none"></form>
<form id="frmApplicationTransfer" Method="post" action="AT010.asp" STYLE="DISPLAY: none"></form>
<form id="frmApplicationCosts" Method="post" action="pp010.asp" STYLE="DISPLAY: none"></form>
<form id="frmPayees" method="post" action="pp070.asp" STYLE="DISPLAY: none"></form>
<form id="frmDisbursements" method="post" action="pp050.asp" STYLE="DISPLAY: none"></form>
<!-- SG 5/12/01 SYS3357 <form id="frmValInstructions" method="post" action="AP250.asp" STYLE="DISPLAY: none"></form>-->

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate ="onchange">
<div id="divBackground" style="HEIGHT: 480px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 85px; POSITION: absolute; TOP: 0px" class="msgLabel">				
				
		<span id="spnApplicationSummary" style="LEFT: 0px; POSITION: absolute; TOP: 20px" class="msgLabel">					
			<IMG alt ="" border=0 height=100 id=imgAppnSummary1 src     ="images/aproc_summary1.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 2" width=200 >				
			<IMG alt ="" border=0 height=100 id=imgAppnSummary2 src     ="images/aproc_summary2.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 1" width=200 >
		</span>
		
		<span id="spnApplicationStage" style="LEFT: 0px; POSITION: absolute; TOP: 130px" class="msgLabel">					
			<IMG alt ="" border=0 height=100 id=imgAppnStage1 src     ="images/aproc_stage1.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 2" width=200 >				
			<IMG alt ="" border=0 height=100 id=imgAppnStage2 src      ="images/aproc_stage2.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 1" width=200 >
		</span>
		
		<span id="spnApplicationTransfer" style="LEFT: 0px; POSITION: absolute; TOP: 240px" class="msgLabel">					
			<IMG alt ="" border=0 height=100 id=imgAppnTransfer1 src     ="images/aproc_transfer1.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 2" width=200 >				
			<IMG alt ="" border=0 height=100 id=imgAppnTransfer2 src     ="images/aproc_transfer2.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 1" width=200 >
		</span>
							
		<span id="spnMortgageCalculator" style="LEFT: 230px; POSITION: absolute; TOP: 20px" class="msgLabel">					
			<IMG alt ="" border=0 height=100 id=imgAppnCosts1 src     ="images/aproc_costs1.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 2" width=200 >				
			<IMG alt ="" border=0 height=100 id=imgAppnCosts2 src     ="images/aproc_costs2.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 1" width=200 >
		</span>
		
		<span id="spnCustomerEngagement" style="LEFT: 230px; POSITION: absolute; TOP: 130px" class="msgLabel">					
			<IMG alt ="" border=0 height=100 id=imgPayees1 src     ="images/aproc_payees1.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 2" width=200 >					
			<IMG alt ="" border=0 height=100 id=imgPayees2 src     ="images/aproc_payees2.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 1" width=200 > 						
		</span>
		
		<span id="spnPipeLineEnquiry" style="LEFT: 230px; POSITION: absolute; TOP: 240px" class="msgLabel">					
			<IMG alt ="" border=0 height=100 id=imgDisbursements1 src     ="images/aproc_disbursements1.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 2" width=200 >					
			<IMG alt ="" border=0 height=100 id=imgDisbursements2 src     ="images/aproc_disbursements2.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 1" width=200 > 						
		</span>
		
		<!-- SG 5/12/01 SYS3357 <span id="spnValuerInstructions" style="LEFT: 115px; POSITION: absolute; TOP: 350px" class="msgLabel">					
			<IMG alt ="" border=0 height=100 id=imgValInstruct1 src     ="images/val_instruct_1.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 2" width=200 >					
			<IMG alt ="" border=0 height=100 id=imgValInstruct2 src     ="images/val_instruct_2.jpg" style="LEFT: 0px; POSITION: absolute; TOP: 0px; Z-INDEX: 1" width=200 > 						
		</span>-->

	</span> 
</div> 
</form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- Specify Code Here -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/mn070Attribs.asp" -->
<script language="JScript">
<!--
var m_sMetaAction		= "";
var scScreenFunctions;
var m_sAppCostsUnavailable = false;
var m_sApplicationNumber = "";

function RetrieveContextData()
{
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Application Processing Menu","MN070",scScreenFunctions);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	
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
	m_sApplicationNumber = scScrFunctions.GetContextParameter(window,"idApplicationNumber","");
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

function frmScreen.imgAppnCosts1.onmouseover()
{
	SwapImage(frmScreen.imgAppnCosts2);
}

function frmScreen.imgAppnStage1.onmouseover()
{
	SwapImage(frmScreen.imgAppnStage2);
}
function frmScreen.imgAppnSummary1.onmouseover()
{
	SwapImage(frmScreen.imgAppnSummary2);
}

function frmScreen.imgAppnTransfer1.onmouseover()
{
	SwapImage(frmScreen.imgAppnTransfer2);
}

function frmScreen.imgAppnCosts1.onmouseover()
{
	SwapImage(frmScreen.imgAppnCosts2);
}

function frmScreen.imgPayees1.onmouseover()
{
	SwapImage(frmScreen.imgPayees2);
}

function frmScreen.imgDisbursements1.onmouseover()
{
	SwapImage(frmScreen.imgDisbursements2);
}

function frmScreen.imgAppnStage2.onmouseout()
{
	RestoreImage(frmScreen.imgAppnStage2);
}

function frmScreen.imgAppnSummary2.onmouseout()
{
	RestoreImage(frmScreen.imgAppnSummary2);
}

function frmScreen.imgAppnCosts2.onmouseout()
{
	RestoreImage(frmScreen.imgAppnCosts2);
}

function frmScreen.imgAppnTransfer2.onmouseout()
{
	RestoreImage(frmScreen.imgAppnTransfer2);
}

function frmScreen.imgAppnCosts2.onmouseout()
{
	RestoreImage(frmScreen.imgAppnCosts2);
}

function frmScreen.imgPayees2.onmouseout()
{
	RestoreImage(frmScreen.imgPayees2);
}

function frmScreen.imgDisbursements2.onmouseout()
{
	RestoreImage(frmScreen.imgDisbursements2);
}

function frmScreen.imgAppnStage2.onclick()
{
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			frmApplicationStage.submit();
			break;
		default: // Error
	}
	
}
function frmScreen.imgAppnSummary2.onclick()
{
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
				frmApplicationSummary.submit();	
				break;
		default: // Error
	}
}

function frmScreen.imgAppnTransfer2.onclick()
{
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
				frmApplicationTransfer.submit();
				break;
		default: // Error
	}
}

function frmScreen.imgDisbursements2.onclick()
{
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
				frmDisbursements.submit();
				break;
		default: // Error
	}
}

function frmScreen.imgAppnCosts2.onclick()
{
	if(m_sAppCostsUnavailable==false)
	{
		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
					frmApplicationCosts.submit();
					break;
			default: // Error
		}
	}
		
}

function frmScreen.imgPayees2.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idReturnScreenId", "MN070.asp");
	
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
				frmPayees.submit();	
				break;
		default: // Error
	}
	
}

-->
</script>
</body>
</html>




