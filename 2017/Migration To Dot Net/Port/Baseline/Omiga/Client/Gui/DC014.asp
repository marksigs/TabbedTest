<%@ Language=JScript %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<% /*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      dc014.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Third Party Data Declaration Screen (a popup).
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
INR		15/06/04	Created.
INR		29/06/2004	BMIDS744 Changed presentation of Radio Buttons.
GHun	16/07/2004	BMIDS744 If no radio button has been selected then don't save a value
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		Description
GHun	22/07/2005	MAR10 Removed the need for scrollbars when using XP
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>DC014 - Third Party Data Declaration <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" mark validate="onchange">

<div style="LEFT: 10px; WIDTH: 520px; POSITION: absolute; TOP: 15px; HEIGHT: 360px" class="msgGroup">
	
	<div id="divDeclaration" class="msgInput" style="LEFT: 10px; TOP: 10px; OVERFLOW: auto; HEIGHT: 260px; WIDTH:500px; POSITION: ABSOLUTE">
		<label id="TPDDeclarationText" class="msgTxt"></label>
	</div>
	
	<div style="TOP: 280px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<label id="TPDDeclarationLabel" class="msgLabel"></label>
		<span style="TOP: -3px; LEFT: 410px; POSITION: ABSOLUTE">
			<input id="CustomerSeenTPDDeclarationYes" name="CustomerSeenTPDDeclaration" type="radio" value="1"><label for="CustomerSeenTPDDeclarationYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: -3px; LEFT: 460px; POSITION: ABSOLUTE">
			<input id="CustomerSeenTPDDeclarationNo" name="CustomerSeenTPDDeclaration" type="radio" value="0"><label for="CustomerSeenTPDDeclarationNo" class="msgLabel">No</label>
		</span>  
	</div> 
	
	<div id="dc014Buttons" style="TOP: 320px; LEFT: 390px; POSITION: ABSOLUTE" class="msgLabel">
		<input type="button" value="OK" class="msgButton" id="btnOK" style="width:60px"> 
		<input type="button" value="Cancel" class="msgButton" id="btnCancel" style="width:60px">
	</div>
	
</div>

</form>


<% /* File containing field attributes */ %>
<!--  #include FILE="attribs/dc014attribs.asp" -->
<!--  #include FILE="Customise/DC014Customise.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = null;
var m_sTPData = null;
var m_sReadOnly = null;
var scScreenFunctions;
var m_BaseNonPopupWindow = null;

function window.onload()
{
	RetrieveData();
	SetMasks();
	// MC SYS2564/SYS2757 for client customisation
	Customise();

	Validation_Init();
	PopulateScreen();
	Initialise();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	window.returnValue = null;
}

function Initialise()
{
	if(m_sReadOnly == "1")
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
}

function PopulateScreen()
{
	if(m_sTPData == "1")
		frmScreen.CustomerSeenTPDDeclarationYes.checked = true;
	else if(m_sTPData == "0")
		frmScreen.CustomerSeenTPDDeclarationNo.checked = true;
}

function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters	= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sTPData	= sParameters[0];
	m_sReadOnly	= sParameters[1];
}

function frmScreen.btnOK.onclick()
{
	var sReturn = new Array();

	if(frmScreen.CustomerSeenTPDDeclarationYes.checked)
		m_sTPData = "1";
	<% /* BMIDS744 GHun If no radio button has been selected then don't save a value */ %>
	else if (frmScreen.CustomerSeenTPDDeclarationNo.checked)
		m_sTPData = "0";
	else 
		m_sTPData = "";
	<% /* BMIDS744 End */ %>
		
	sReturn[0] = IsChanged();
	sReturn[1] = m_sTPData;
	window.returnValue	= sReturn;
	window.close();
}

function frmScreen.btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}
-->
</script>
</body>
</html>
