<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      RA033.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Address Match
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		05/12/00	Screen Design
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<head>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Address Match <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>

<script src="validation.js" language="JScript"></script>

<% //Span to keep tabbing within this screen %> 
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" validate ="onchange" mark>
<div id="divCCData" style="HEIGHT: 56px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		CCN Ref. No.
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtCreditCheckReferenceNumber" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Name
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtFullBureauName" maxlength="50" style="POSITION: absolute; WIDTH: 250px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
</div>

<div id="divAddress" style="HEIGHT: 200px; LEFT: 10px; POSITION: absolute; TOP: 70px; WIDTH: 604px" class="msgGroup">
	<span id="spnCurrentAddress" style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
			<strong>Current Address</strong>
		</span>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel">
			Flat No. 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtFlatNo" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>
		
		<span style="LEFT: 4px; POSITION: absolute; TOP: 42px" class="msgLabel">
			House <br>Name &amp; No. 
			<span style="LEFT: 80px; POSITION: absolute; TOP: 2px">
				<input id="txtHouseName" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
			<span style="LEFT: 231px; POSITION: absolute; TOP: 2px">
				<input id="txtHouseNo" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 72px" class="msgLabel">
			Street 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtStreet" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 96px" class="msgLabel">
			District 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtDistrict" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 120px" class="msgLabel">
			Town
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtTown" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 144px" class="msgLabel">
			County
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtCounty" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>	
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 168px" class="msgLabel">
			Post Code
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtPostCode" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>	
	</span>
	
	<span id="spnFBAddress" style="TOP: 4px; LEFT: 310px; POSITION: ABSOLUTE">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
			<strong>Bureau Address</strong>
		</span>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel">
			Flat No. 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtFBFlatNo" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 42px" class="msgLabel">
			House<br>Name &amp; No. 
			<span style="LEFT: 80px; POSITION: absolute; TOP: 2px">
				<input id="txtFBHouseName" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
			<span style="LEFT: 231px; POSITION: absolute; TOP: 2px">
				<input id="txtFBHouseNo" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 72px" class="msgLabel">
			Street 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtFBStreet" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 96px" class="msgLabel">
			District 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtFBDistrict" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 120px" class="msgLabel">
			Town
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtFBTown" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 144px" class="msgLabel">
			County
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtFBCounty" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>	
	
		<span style="LEFT: 4px; POSITION: absolute; TOP: 168px" class="msgLabel">
			Post Code
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtFBPostCode" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" READONLY tabindex=-1>
			</span>
		</span>	
	</span>
	
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 300px; WIDTH: 450px">
<!-- #include FILE="msgButtons.asp" -->
</div>

<script language="JScript">
<!--

var m_aArgArray = null;
var m_aFBAddress = null;
var m_aAddress = null;
var m_BaseNonPopupWindow = null;

function window.onload()
{
	var sArguments		= window.dialogArguments ;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];	
	m_aArgArray			= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
	
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	
	PopulateScreen();
}

function btnSubmit.onclick()
{
	window.close();	
}

function PopulateScreen()
{
	frmScreen.txtCreditCheckReferenceNumber.value = m_aArgArray[0];
	frmScreen.txtFullBureauName.value = m_aArgArray[1];
	
	m_aFBAddress = m_aArgArray[2];
	m_aAddress   = m_aArgArray[3];
	
	frmScreen.txtFBFlatNo.value		= m_aFBAddress[0] ;
	frmScreen.txtFBHouseName.value	= m_aFBAddress[1] ;
	frmScreen.txtFBHouseNo.value	= m_aFBAddress[2] ;
	frmScreen.txtFBStreet.value		= m_aFBAddress[3] ;
	frmScreen.txtFBTown.value		= m_aFBAddress[4] ;
	frmScreen.txtFBDistrict.value	= m_aFBAddress[5] ;
	frmScreen.txtFBCounty.value		= m_aFBAddress[6] ;
	frmScreen.txtFBPostCode.value	= m_aFBAddress[7] ;

	frmScreen.txtFlatNo.value		= m_aAddress[0] ;
	frmScreen.txtHouseName.value	= m_aAddress[1] ;
	frmScreen.txtHouseNo.value		= m_aAddress[2] ;
	frmScreen.txtStreet.value		= m_aAddress[3] ;
	frmScreen.txtTown.value			= m_aAddress[4] ;
	frmScreen.txtDistrict.value		= m_aAddress[5] ;
	frmScreen.txtCounty.value		= m_aAddress[6] ;
	frmScreen.txtPostCode.value		= m_aAddress[7] ;
}

-->
</script>
</body>
</html>