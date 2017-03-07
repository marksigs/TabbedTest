<%@ LANGUAGE="JSCRIPT" %>
 
<% /*
Workfile:      cm160.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Part and Part screen.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		24/11/99	AQR 9: Ensure split amounts initialised correctly on screen entry.
RF		08/02/00	Reworked for IE5 and performance.
JLD		29/02/00	SYS0311	Altered tabbing sequence
MCS		01/03/00	SYS0370	moved button and altered tabbing sequnce
JLD		02/03/00	removed unnecessary spnToFirstField etc. 
AY		29/03/00	scScreenFunctions change
AY		13/04/00	SYS0328 - Dynamic currency display
JLD		03/12/01	SYS3253 - Output error message correctly.
LD		23/05/02	SYS4727 Use cached versions of frame functions
MC		06/05/2005	BMIDS571 White Space added to page title and page alignment fix
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Part and Part <!-- #include file="includes/TitleWhiteSpace.asp" -->       </title>
</head>

<body>
<script src="validation.js" language="JScript"></script>

<OBJECT id=scTable style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 data=scTable.htm VIEWASTEXT></OBJECT>

<% // Specify Forms Here %>

<% // Specify Screen Layout Here %>
<form id="frmScreen" year4 validate="onchange" mark>
	<div id="divBackground" style="LEFT: 10px; WIDTH: 472px; POSITION: absolute; 
		TOP: 10px; HEIGHT: 130px" class="msgGroup">

		<div id="divInterestOnlyAmount" 
			style="LEFT: 0px; WIDTH: 466px; POSITION: absolute; TOP: 0px; HEIGHT: 110px"
			 class="msgGroup">
			<span style="LEFT: 31px; POSITION: absolute; TOP: 10px" class="msgLabel">
				Interest Only Amount
				<span style="LEFT: 135px; POSITION: absolute; TOP: 0px">
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtInterestOnlyAmount" name="InterestOnlyAmount" maxlength="7" 
						style="WIDTH: 65px; POSITION: absolute; TOP: -3px" class="msgTxt"> 
				</span> 
			</span>					
		</div>
			

		<span id=spnCalculateSplit style="LEFT: 0px; POSITION: absolute; TOP: 5px">
			<span style="LEFT: 250px; POSITION: absolute; TOP: 0px">
				<input id="btnCalculateSplit" value="CalculateSplit" 
					type="button" style="WIDTH: 80px" class="msgButton">
			</span>
		</span>

		<div id="divAmount" 
			style="LEFT: 0px; WIDTH: 466px; POSITION: absolute; TOP: 40px; HEIGHT: 40px"
			 class="msgGroup">
			<span style="LEFT: 31px; POSITION: absolute; TOP: 10px" class="msgLabel">
				Capital &amp; Interest Amount
				<span style="LEFT: 135px; POSITION: absolute; TOP: 0px">						
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtCapitalAndInterestAmount" name="CapitalAndInterestAmount" maxlength="7" 
						style="WIDTH: 65px; POSITION: absolute; TOP: -3px" class="msgReadOnly"> 
				</span> 
			</span>
			<span style="LEFT: 250px; POSITION: absolute; TOP: 10px" class="msgLabel">
				Total Loan Amount
				<span style="LEFT: 110px; POSITION: absolute; TOP: 0px">						
					<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
					<input id="txtTotalLoanAmount" name="TotalLoanAmount" maxlength="7" 
						style="WIDTH: 65px; POSITION: absolute; TOP: -3px" class="msgReadOnly"> 
				</span> 
			</span>
		</div>

		<span id=spnOkCancelButtons style="LEFT: 0px; POSITION: absolute; TOP: 90px">
			<span style="LEFT: 175px; POSITION: absolute; TOP: 0px">
				<input id="btnOK" value="OK" type="button" 
					style="WIDTH: 60px" disabled class="msgButton">
			</span>
			<span style="LEFT: 245px; POSITION: absolute; TOP: 0px">
				<input id="btnCancel" value="Cancel" type="button" 
					style="WIDTH: 60px" class="msgButton">
			</span>
		</span>
	
	</div>
</form>
	
<% // File containing field attributes (remove if not required) %>
<!--  #include FILE="attribs/cm160attribs.asp" -->

<% // Specify Code Here %>
<script language="JScript">
<!--	// JScript Code
var m_sMetaAction = null;
var scScreenFunctions;
var m_BaseNonPopupWindow = null;

function window.onload()
{
	RetrieveData();			
	SetMasks();
	Validation_Init();
	SetScreenOnReadOnly();
	scScreenFunctions.SetFocusToFirstField(frmScreen);			
	window.returnValue = null;
}

function RetrieveData()
{
	<% /*
		Description:	Retrieve information from mc010
		RF	24/11/99 AQR 9: Ensure split amounts initialised correctly on screen entry
	*/ %>

	var sArguments = window.dialogArguments;
			
	window.dialogTop		= sArguments[0];
	window.dialogLeft		= sArguments[1];
	window.dialogWidth		= sArguments[2];
	window.dialogHeight		= sArguments[3];			
	var sParameters	= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
																	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	frmScreen.txtTotalLoanAmount.value = sParameters[0];
	frmScreen.txtInterestOnlyAmount.value = sParameters[1];
	frmScreen.txtCapitalAndInterestAmount.value = sParameters[2];

	m_sReadOnly	= sParameters[3];
	scScreenFunctions.SetThisCurrency(frmScreen,sParameters[4]);
}

function SetScreenOnReadOnly()
{
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnCalculateSplit.disabled = true;
		frmScreen.btnOK.disabled = true;
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtInterestOnlyAmount");
		frmScreen.btnCancel.focus();
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCapitalAndInterestAmount");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTotalLoanAmount");
		frmScreen.txtInterestOnlyAmount.focus();
	}		
}		

function frmScreen.btnCalculateSplit.onclick()
{
	var iInterestOnlyAmount = parseInt(frmScreen.txtInterestOnlyAmount.value);
	var iTotalLoanAmount	= parseInt(frmScreen.txtTotalLoanAmount.value);
			
	if ((isNaN(iInterestOnlyAmount)) ||
		(isNaN(iTotalLoanAmount)) ||
		(iInterestOnlyAmount == 0) ||
		(iInterestOnlyAmount >= iTotalLoanAmount))
	{
		alert("Invalid interest-only amount");
	}
	else
	{
		var TotalLoanAmount = frmScreen.txtTotalLoanAmount.value
		var InterestOnlyAmount = frmScreen.txtInterestOnlyAmount.value
		frmScreen.txtCapitalAndInterestAmount.value = TotalLoanAmount-InterestOnlyAmount;
		frmScreen.btnOK.disabled = false;
	}				
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
	sReturn[1] = frmScreen.txtInterestOnlyAmount.value;	
	sReturn[2] = frmScreen.txtCapitalAndInterestAmount.value;			
	window.returnValue	= sReturn;
	window.close();
}
-->
</script>
</body>
</html>
