<%@ LANGUAGE="JSCRIPT" %>
<% /*
Workfile:		cm115.asp
Copyright:		Copyright © 2000 Marlborough Stirling

Description:	Product Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		29/03/00	scScreenFunctions change
AY		29/03/00	Incorrect comment
BG		17/05/00	SYS0752 Removed Tooltips
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MC		19/04/2004	BMIDS517	space padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<html>
<head>
	<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Mortgage Product Details <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>
<!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<script src="validation.js" language="JScript"></script>
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 220px; WIDTH: 320px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Product Details...</strong>
	</span>
	<span style="TOP: 30px; LEFT: 10px; POSITION: ABSOLUTE">
		<textarea id="txtProductDetails" name="ProductDetails" 
			rows="10" tabindex="-1" 
			style="WIDTH: 300px; POSITION: absolute" class="msgReadOnly">
		</textarea>
	</span>
	<span style="TOP: 190px; LEFT: 4px; POSITION: ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
</div>
</form>
<script language="JScript">
<!--
var m_sProductDetails	= null;
var scScreenFunctions;
var m_BaseNonPopupWindow = null;

function window.onload()
{
	RetrieveData();
	Validation_Init();
	PopulateScreen();
	scScreenFunctions.SetScreenToReadOnly(frmScreen);
	frmScreen.btnOK.focus();
	window.returnValue = null;
}
function PopulateScreen()
{
	frmScreen.txtProductDetails.value = m_sProductDetails;
}
function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters		= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sProductDetails	= sParameters[0];
}
function frmScreen.btnOK.onclick()
{
	window.close();
}
-->
</script>
</body>
</html>
