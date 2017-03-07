<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<% /*
Workfile:      pp051.asp
Copyright:     Copyright © 2006 Vertex

Description:   Completion Date   THIS IS A POP-UP SCREEN
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DRC 	21/03/06	Created (MAR 1408)

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Delayed Completion Date   <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark    validate ="onchange"><!--style="VISIBILITY: hidden"> -->


<div id = "divDetails" style="LEFT: 10px; POSITION: absolute; TOP: 0px; WIDTH: 300px" class="msgGroup">
	<span style="LEFT: 40px; POSITION: absolute; TOP: 25px" class="msgLabel">
		Current Completion Date
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtCurrCompDate" maxlength="10" style="POSITION: absolute; WIDTH: 60px" class="msgReadOnly" readonly tabindex="-1">
		</span>
	</span>

	<span style="LEFT: 40px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Revised Completion Date
		<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
			<input id="txtRevCompDate" maxlength="10" style="POSITION: absolute; WIDTH: 60px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 0px; POSITION: absolute; TOP: 120px">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
	</span>	
	<span style="LEFT: 65px; POSITION: absolute; TOP: 120px">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 70px" class="msgButton">
	</span>
		
</div>

</form>

<!-- File containing field attributes -->
<!--  #include FILE="attribs/pp051attribs.asp" -->

<!-- Specify Code Here -->
<script language="JScript">
<!--
var m_sCurrentCompletionDate = null;
var m_sRevisedCompletionDate = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sRequestAttributes = null;
var m_bCompletionDateValid = false;
var m_sUserID = null;

var scClientScreenFunctions;
function window.onload()
{
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	scScreenFunctions.ShowCollection(frmScreen);
	m_sRequestAttributes = sArgArray[0];
	m_sApplicationNumber = sArgArray[1];
	m_sApplicationFactFindNumber = sArgArray[2];
	m_sCurrentCompletionDate = sArgArray[3];
	m_sRevisedCompletionDate = m_sCurrentCompletionDate;
	m_sUserID = sArgArray[4];
	
	SetMasks();
	Validation_Init();

	PopulateScreen();
		
	window.returnValue = m_sCurrentCompletionDate;   
	
	ClientPopulateScreen();
}


function PopulateScreen()
{
	<% /* Set the Current Completion Date to the value passed in from PP050 */ %>
	frmScreen.txtCurrCompDate.value =  m_sCurrentCompletionDate;
	frmScreen.txtCurrCompDate.disabled = true;
	frmScreen.txtRevCompDate.value =  m_sRevisedCompletionDate;

}

function frmScreen.btnCancel.onclick()
{	
    window.returnValue = m_sCurrentCompletionDate;
   	window.close();
}


function frmScreen.btnOK.onclick()
{
	//Validate completion date
	m_bCompletionDateValid = true;
	
	if (frmScreen.txtRevCompDate.valid() == true)
	{
		if ((frmScreen.txtRevCompDate.value != "") && (m_sCurrentCompletionDate != frmScreen.txtRevCompDate.value))
		{
			if (ValidateCompletionDate())
			{
				m_sRevisedCompletionDate = frmScreen.txtRevCompDate.value;
				
			}
			else
			{
				frmScreen.txtRevCompDate.focus();
			}
		}
		else if (frmScreen.txtRevCompDate.value == "")
		{
			m_bCompletionDateValid = false;
			m_sRevisedCompletionDate = "";
			frmScreen.txtRevCompDate.focus();
		
			alert("Please enter a valid Completion Date in the format DD/MM/YYYY");
		}
	}
	else
	{
		m_bCompletionDateValid = false;
		m_sRevisedCompletionDate = "";
		frmScreen.txtRevComp.focus();
		
		alert("The Completion Date is invalid - please enter a valid date in the format DD/MM/YYYY");

	}
	
	<% /* If the revised date is valid, exit the screen */ %>
    if (m_bCompletionDateValid)
    {
		window.returnValue = m_sRevisedCompletionDate;
		window.close();
	}
}


function ValidateCompletionDate()
{
	// if the date is not valid, set focus on date field
		m_bCompletionDateValid = false;
		XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTagFromArray(m_sRequestAttributes, "REQUEST");
		XML.CreateTag("ApplicationNumber", m_sApplicationNumber);
		//XML.CreateRequestTag(window, "REQUEST")	//XML.CreateRequestTag(window, null)
		XML.SetAttribute("OPERATION","ValidateCompletionDate");
		XML.CreateActiveTag("VALIDATECOMPLETIONDATE");
		XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.SetAttribute("COMPLETIONDATE", frmScreen.txtRevCompDate.value );


		XML.RunASP(document,"PaymentProcessingRequest.asp");
		
		var TagRESPONSE = XML.SelectTag(null,"RESPONSE");
		if (XML.SelectTag(TagRESPONSE,"ERROR") != null)
		{
			var sErrorMessage = XML.GetTagText("DESCRIPTION");
			alert(sErrorMessage);
		}
		else
		{
			m_bCompletionDateValid = true;
		}
		var XML=null;
		return m_bCompletionDateValid;
}

-->
</script>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/Jscript"></script>
</body>
</html>
