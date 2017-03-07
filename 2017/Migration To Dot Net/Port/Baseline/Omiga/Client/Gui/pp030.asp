<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<%
/*
Workfile:      pp030.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Adjust Fee Amount
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		08/02/01	SYS1864	- Adjustment amount cannot be larger than the outstanding amount
SR		22/02/01	SYS1958 - Do all the validation on clicking OK button
SR		22/03/01	SYS2153	- Change in the layout and Amout validation
SA		17/05/01	SYS2223 - "Ok" should be "OK"
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>

<title>PP030 - Adjust Fee Amount</title>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
</HEAD>

<BODY>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Scriptlets */ %>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 150px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 325px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Fee Type
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtFeeType" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
		
	<span style="TOP: 44px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Total Amount Due
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtTotalAmountDue" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
	
	<span style="TOP: 78px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Adjustment
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<select id="cboAdjustment" maxlength="30" style="POSITION: absolute; WIDTH: 150px" align="right" class="msgCombo"></select>
		</span>
	</span>
	
	<span style="TOP: 112px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Amount 
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtRebateAmount" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
</div>

<div style="HEIGHT: 150px; LEFT: 340px; POSITION: absolute; TOP: 10px; WIDTH: 276px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Amount Outstanding</strong>
	</span>
	
	<span style="TOP: 30px; LEFT: 30px; POSITION: ABSOLUTE" class="msgLabel">
		Actual
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtAmountOutStanding" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>

	<span style="TOP: 64px; LEFT: 30px; POSITION: ABSOLUTE" class="msgLabel">
		After Rebate/Addition
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtActualAmountOutStanding" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>
</div>

<div style="HEIGHT: 130px; LEFT: 10px; POSITION: absolute; TOP: 165px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Notes
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<TEXTAREA class=msgTxt id="txtNotes" name=Notes rows=5 style="POSITION: absolute; WIDTH: 450px"></TEXTAREA>  
		</span>	
	</span>
	
	<span style="TOP: 96px; LEFT: 4px; POSITION: ABSOLUTE">
		<% /*SYS2223 Changed "Ok" to "OK"*/%>
		<input id="btnOk" value="OK" type="button" style="WIDTH: 75px" class="msgButton">
	</span>	
	
	<span style="TOP: 96px; LEFT: 100px; POSITION: ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
</div>

</form>


<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="attribs/PP030attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;
var m_sApplicationNumber;
var m_sFeeType ;
var m_sFeeTypeDesc;
var m_sInitialAmount ;
var m_sNotes ;
var m_sFeeTypeComboSeq ;
var m_aArgArray = null;
var m_sAmountDue;
var m_sAmountOutStanding;
var m_RebateAmount;
var m_AdditionalAmount;
var m_nActualAmountOutstanding ;
var m_BaseNonPopupWindow = null;

<% /**** Events *****/  %>

function frmScreen.cboAdjustment.onchange()
{
	frmScreen.txtRebateAmount.value = "";
}

function frmScreen.btnCancel.onclick()
{
	window.close();
}

function frmScreen.btnOk.onclick()
{
	if (UpdateApplicationFeeType())
	{
		window.returnValue	= "Ok"; 
		window.close();
	}
}

function frmScreen.txtRebateAmount.onchange()
{	
	<% /* SR 07-02-01 : SYS1864 - rebate amount cannot be larger than the outstanding amount */ %>
	if(frmScreen.cboAdjustment.value=="1")
	{
		if(parseFloat(frmScreen.txtRebateAmount.value) > m_nActualAmountOutstanding)
		{
			alert('Rebate cannot be greater than the amount outstanding after rebate/addition.');
			return false
		}
	}
}
<%
/* //SYS1958 - validations on the field RebateAmount are to be done when OK is clicked
function frmScreen.txtRebateAmount.onblur()
{
	if(frmScreen.cboAdjustment.value=="1" && frmScreen.txtRebateAmount.value <= 0)
		window.alert("Rebate Amount must be greater than zero") ;	
	else if(frmScreen.cboAdjustment.value=="2" && frmScreen.txtRebateAmount.value <= 0)
		window.alert("Additional Amount must be greater than zero") ;		
}
*/
%>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sArguments = window.dialogArguments ;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	m_aArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	m_sApplicationNumber = m_aArgArray[0];
	m_sFeeTypeDesc = m_aArgArray[1];
	m_sAmountDue = m_aArgArray[2];
	m_sAmountOutStanding = m_aArgArray[3];
	m_sNotes = m_aArgArray[4];
	m_sFeeTypeComboSeq = m_aArgArray[5];
	m_sInitialAmount = parseFloat(m_aArgArray[6]);
	m_sFeeTypeRecordSeq = m_aArgArray[7];

	SetMasks();  

	AddComboOption( frmScreen.cboAdjustment, "<SELECT>","0" );
	AddComboOption( frmScreen.cboAdjustment, "Rebate","1" );
	AddComboOption( frmScreen.cboAdjustment, "Addition","2" );
	
	Validation_Init();
	InitialiseScreen();
	
	// fixed combo	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function AddComboOption(cboToAdd, strName, strValue)
{
	var TagOPTION = document.createElement("OPTION");
	TagOPTION = document.createElement("OPTION");
	TagOPTION.value	= strValue;
	TagOPTION.text = strName;
	cboToAdd.add(TagOPTION);
}


<% /**** FUNCTIONS *******/ %>

function InitialiseScreen()
{
	frmScreen.txtFeeType.value = m_sFeeTypeDesc ;
	frmScreen.txtNotes.value = m_sNotes ;
	frmScreen.txtTotalAmountDue.value = m_sAmountDue;
	frmScreen.txtAmountOutStanding.value = m_sAmountOutStanding;
	
	<% /* SYS2153  */ %>
	m_nActualAmountOutstanding =  parseFloat(frmScreen.txtAmountOutStanding.value) + ((-1) * parseFloat(m_sInitialAmount));
	
	if(parseInt(m_sInitialAmount) > 0) 
	{
		frmScreen.cboAdjustment.selectedIndex = 2;
		frmScreen.txtRebateAmount.value = m_sInitialAmount ;
	}
	else
	{ 
		frmScreen.cboAdjustment.selectedIndex = 1;
		frmScreen.txtRebateAmount.value = (-1) * m_sInitialAmount ;
	}
	
	frmScreen.txtActualAmountOutStanding.value = m_nActualAmountOutstanding ;
	
}

function UpdateApplicationFeeType()
{
	if(!ValidateBeforeSave()) return ;
	
	var bContinue ;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
<%  /*
		Do field level validations here before saving
		2)Rebate and AdditionalAmount are mutually exclusive; should be greater than 0
	*/
 %>
		
	var sRebatOrAdditionalAmount = "0";
	
	if(frmScreen.cboAdjustment.value=="1") // rebate means negative # 
		sRebatOrAdditionalAmount = 	(-1) * frmScreen.txtRebateAmount.value ;
	else if(frmScreen.cboAdjustment.value=="2")
		sRebatOrAdditionalAmount = 	frmScreen.txtRebateAmount.value ;
	
	XML.CreateRequestTagFromArray(m_aArgArray[8], "UPDATEAPPLICATIONFEETYPE");
	XML.CreateActiveTag("APPLICATIONFEETYPE");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("FEETYPE", m_sFeeTypeComboSeq);
	XML.SetAttribute("REBATEORADDITION", sRebatOrAdditionalAmount);
	XML.SetAttribute("NOTES", frmScreen.txtNotes.value);
	//XML.SetAttribute("AMOUNT", frmScreen.txtRebateAmount.value);
	XML.SetAttribute("FEETYPESEQUENCENUMBER", m_sFeeTypeRecordSeq);
			
	// 	XML.RunASP(document,"PaymentProcessingRequest.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"PaymentProcessingRequest.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	bContinue = XML.IsResponseOK();
	
	return bContinue
}

function ValidateBeforeSave()
{
	if(frmScreen.txtRebateAmount.value == '')
	{
		alert('Adjustment amount cannot be empty');
		frmScreen.txtRebateAmount.focus();
		return false;
	}
	
	if(parseFloat(frmScreen.txtRebateAmount.value) <= 0)
	{
		alert('Adjustment amount should be greater than zero');
		frmScreen.txtRebateAmount.focus();
		return false;
	}
	
	<% /* SR 07-02-01 : SYS1864 - rebate amount cannot be larger than the outstanding amount */ %>
	if(frmScreen.cboAdjustment.value=="1")
	{
		if(parseFloat(frmScreen.txtRebateAmount.value) > m_nActualAmountOutstanding)
		{
			alert('Rebate cannot be greater than the amount outstanding after rebate/addition.');
			frmScreen.txtRebateAmount.focus();
			return false;
		}
	}
	
	<% /* Note can be maximum upto 255 chars */ %>
	var sNotes = frmScreen.txtNotes.value ;
	if(sNotes.length > 255)
	{
		alert('Note can be upto a maximum of 255 characters only');
		return false;
	}
	
	return true;
}

-->
</SCRIPT>
</BODY>
</HTML>


