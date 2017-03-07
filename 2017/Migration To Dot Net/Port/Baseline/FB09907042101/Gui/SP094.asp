<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      SP093.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Sanction/Unsanction Authorisation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
CL		14/02/2001	SYS1982 Initial creation 
CL		05/03/01	SYS1920 Read only functionality added
CL		07/03/01	SYS2002 Removed FW030 detail
CL		07/03/01	SYS2002 Adjustment for Popup functionality 
APS		22/03/01	SYS2146
SR		06/09/01	SYS2412
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS SPECIFIC
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
GD		14/10/02	BMIDS00280 - Make sure XML object has been defined, and change maxlength on TTref1 and TTref 2 to match DB
MV		31/01/2003	BM0254 ApplicationField is not long enough - made length = 12
MC		05/05/2004	BMIDS751	- White space added to the title
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Manual Payments<!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange" year4>
	<div id="divSanctionDetails" style="TOP: 10px; LEFT: 10px; HEIGHT: 170px; WIDTH: 350px; POSITION: ABSOLUTE" class="msgGroup">
	
		<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
			Application number
			<span style="TOP:-3px; LEFT:130px; POSITION:ABSOLUTE">
				<input id="txtApplicationNumber" maxlength="12" style="WIDTH:100px" class="msgTxt">
			</span>
		</span>
		<span style="TOP:35px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
			Payment amount
			<span style="TOP:-3px; LEFT:130px; POSITION:ABSOLUTE">
				<input id="txtValueOfPayment" maxlength="10" style="WIDTH:200px" class="msgTxt">
			</span>
		</span>
		<span style="TOP:60px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
			Manual cheque number
			<span style="TOP:-3px; LEFT:130px; POSITION:ABSOLUTE">
				<input id="txtManualChequeNumber" maxlength="10" style="WIDTH:200px" class="msgTxt">
			</span>
		</span>
		<span style="TOP:85px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
			TT reference 1
			<span style="TOP:-3px; LEFT:130px; POSITION:ABSOLUTE">
				<input id="txtTTReference1" maxlength="7" style="WIDTH:200px" class="msgTxt">
			</span>
		</span>
		<span style="TOP:110px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
			TT reference 2
			<span style="TOP:-3px; LEFT:130px; POSITION:ABSOLUTE">
				<input id="txtTTReference2" maxlength="9" style="WIDTH:200px" class="msgTxt">
			</span>
		</span>
		<!-- Buttons -->
		<span id="spnButtons" style="LEFT: 4px; POSITION: absolute; TOP: 145px">
			<span style="LEFT: 1px; POSITION: absolute; TOP: 0px">
				<input id="btnOK" value="OK" type="button" 
					style="WIDTH: 80px" class="msgButton">
				</span>
			<span style="LEFT: 104px; POSITION: absolute; TOP: 0px">
				<input id="btnCancel" value="Cancel" type="button" 
					style="WIDTH: 80px" class="msgButton">
			</span>
		</span>
	</div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/SP094Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_ArraySelectedRows = null; 
var m_blnReadOnly = false;
var XML = null;
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
	Validation_Init();
	<% //GD BMIDS00280 START
	//RetrieveData();
	//XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	%>

	RetrieveData();
	
	//FW030SetTitles("Manual Payment Details","SP094",scScreenFunctions);
	//m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "SP094")
	
	scScreenFunctions.ShowCollection(frmScreen);	
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}


function RetrieveData()
{
	var ArraySelectedRows = new Array();
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters	= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	frmScreen.txtValueOfPayment.value = sParameters[0];
	frmScreen.txtApplicationNumber.value = sParameters[1];
	m_ArraySelectedRows = sParameters[2];
	XML.LoadXML(m_ArraySelectedRows);
}


function frmScreen.btnOK.onclick()
{
	//Load the form data into the XML
	XML.SelectTag(XML.ActiveTag,"PAYMENTRECORD");
	XML.SetAttribute("CHEQUENUMBER", frmScreen.txtManualChequeNumber.value);
	XML.SelectTag(XML.ActiveTag,"DISBURSEMENTPAYMENT");
	XML.SetAttribute("TTREFNUMBER1", frmScreen.txtTTReference1.value);
	XML.SetAttribute("TTREFNUMBER2", frmScreen.txtTTReference2.value);
	//Run the update method
	// 	XML.RunASP(document,"UpdateDisbursement.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"UpdateDisbursement.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	if (XML.IsResponseOK()==true){
		var sReturn = new Array();

		sReturn[0] = true;

		window.returnValue = sReturn;
		window.close();
	}
}

function frmScreen.btnCancel.onclick()
{
	var sReturn = new Array();
	sReturn[0] = false;
	window.returnValue = sReturn;		
	window.close();
}

-->
</script>
</body>
</html>




