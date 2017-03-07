<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<HTML>
  <HEAD>
	<title>Non Resident Details<!-- #include file="includes/TitleWhiteSpace.asp" --></title>
<%/*
Workfile:      DC036.asp
Copyright:     Copyright © 2006 Vertex Financial Services
Author			Martin Heys

Description:   Residency details  POP-UP
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MAH		14/11/2006	Created
INR		15/01/2007	EP2_677 removed if not resident label
*/ %>
	<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
  </HEAD>

<body>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript" type="text/javascript"></script>

<%/* FORMS */%>
<form id="frmScreen" mark validate="onchange">
	<div id="divResidency" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 10px; HEIGHT: 288px" class="msgGroup">
		<span style="LEFT: 50px; POSITION: absolute; TOP: 60px" class="msglabel">Where do you reside?
			<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
				<input class=msgTxt id="txtCountry"  name="txtCountry" style="WIDTH: 168px; POSITION: absolute; HEIGHT: 20px" maxLength="30" size="22">
			</span>
		</span>
		<span style="LEFT: 50px; POSITION: absolute; TOP: 90px" class="msglabel">Why?
			<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
				<textarea class=msgTxt id="txtWhy" name="txtWhy" rows="4" style="LEFT: 0px; WIDTH: 320px; POSITION: absolute; TOP: 2px; HEIGHT: 62px" cols=37></textarea>
			</span>
		</span>
		<span style="LEFT: 50px; POSITION: absolute; TOP: 160px" class="msgLabel">Do you have the right to work and<br>reside in the UK permanently?
			<span style="LEFT: 220px; POSITION: absolute; TOP: 4px">
				<input id="optRightToWorkYes" name="RightToWorkGroup" type="radio" value="1"><label for="optRightToWorkYes" class="msgLabel">Yes</label>
			</span> 
			<span style="LEFT: 280px; POSITION: absolute; TOP: 4px">
				<input id="optRightToWorkNo" name="RightToWorkGroup" type="radio" value="0"><label for="optRightToWorkNo" class="msgLabel">No</label>
			</span> 
		</span>
		<span style="LEFT: 50px; POSITION: absolute; TOP: 200px" class="msgLabel">Do you have diplomatic immunity?
			<span style="LEFT: 220px; POSITION: absolute; TOP: -4px">
				<input id="optDiplomaticImmunityYes" name="DiplomatGroup" type="radio" value="1"><label for="optDiplomaticImmunityYes" class="msgLabel">Yes</label>
			</span> 
			<span style="LEFT: 280px; POSITION: absolute; TOP: -4px">
				<input id="optDiplomaticImmunityNo" name="DiplomatGroup" type="radio" value="0"><label for="optDiplomaticImmunityNo" class="msgLabel">No</label>
			</span> 
		</span>
		<span style="LEFT: 20px; POSITION: absolute; TOP: 250px">
			<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton" NAME="btnOK">
			<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">
				<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton" NAME="btnCancel">
			</span>
		</span>	
	</div>			
</form>

<!--  #include FILE="attribs/DC036attribs.asp" -->

<%/* CODE */%>
<script language="JScript" type="text/javascript">
<!--
var m_sXMLIn = "";
var m_sReadOnly = "";
var CustomerXML = null;
//var m_sCountryOfResidency = "";
//var m_sReasonForCountryOfResidency = "";
var m_sRightToWork = "";
var m_sDiplomaticImmunity = "";
var scScreenFunctions;
var m_BaseNonPopupWindow = null;

function window.onload()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions.ShowCollection(frmScreen);

	SetMasks();

	m_sReadOnly = sArgArray[4];
	m_sXMLIn = sArgArray[5]; <% //CustomerXML %>

	Initialise();
	Validation_Init();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	// Read Only processing
	if (m_sReadOnly == "1")
	window.returnValue = null;
}
function PopulateScreen()
{
	CustomerXML.SelectTag(null,"CUSTOMERVERSION");
	frmScreen.txtCountry.value = CustomerXML.GetTagText("COUNTRYOFRESIDENCY");
	frmScreen.txtWhy.value = CustomerXML.GetTagText("REASONFORCOUNTRYOFRESIDENCY");
	scScreenFunctions.SetRadioGroupValue(frmScreen, "RightToWorkGroup", CustomerXML.GetTagText("RIGHTTOWORK"));		
	scScreenFunctions.SetRadioGroupValue(frmScreen, "DiplomatGroup", CustomerXML.GetTagText("DIPLOMATICIMMUNITY"));		
}
function Initialise()
{
	CustomerXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	CustomerXML.LoadXML(m_sXMLIn);
	PopulateScreen();
	if(m_sReadOnly == "1")
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
}
function frmScreen.btnOK.onclick()
{
	var sValidateData = ValidateResidency();
	if (sValidateData)
	{
		var sReturn = new Array();

		sReturn[0] = IsChanged();		// Has there been a change made
		sReturn[1] = WriteResidencyDetails();		// The XML string

		window.returnValue = sReturn;
		window.close();
	}	
}
function frmScreen.btnCancel.onclick()
{
	window.close();
}
function ValidateResidency()
{
	var bValidOK = false;
	if(frmScreen.txtCountry.value == "")
		{
		alert("Please enter Country");
		frmScreen.txtCountry.focus();
		return(bValidOK);
		}
	if(frmScreen.txtWhy.value == "")
		{
		alert("Please enter \"Why\"");
		frmScreen.txtWhy.focus();
		return(bValidOK);
		}		
	if(!frmScreen.optRightToWorkYes.checked && !frmScreen.optRightToWorkNo.checked)
		{
		alert("Please indicate eligibility to work and reside in the UK");
		frmScreen.optRightToWorkNo.focus();
		return(bValidOK);
		}				
	if(!frmScreen.optDiplomaticImmunityYes.checked && !frmScreen.optDiplomaticImmunityNo.checked)
		{
		alert("Please indicate presence of Diplomatic Immunity");
		frmScreen.optDiplomaticImmunityNo.focus();
		return(bValidOK);
		}				
	bValidOK = true;
	return(bValidOK);
}
function frmScreen.txtWhy.onkeyup()
{
	scScreenFunctions.RestrictLength(frmScreen, "txtWhy", 255, true);
}
function WriteResidencyDetails()
{
	//save any changed data
	if ((m_sReadOnly != "1") && frmScreen.onsubmit())
	{
		if (IsChanged() == true )
		{
			//CustomerXML.ActiveTag = null;
			CustomerXML.SelectTag(null,"CUSTOMERVERSION");
			
			if(frmScreen.txtCountry.value != CustomerXML.GetTagText("COUNTRYOFRESIDENCY"))
				CustomerXML.SetTagText("COUNTRYOFRESIDENCY", frmScreen.txtCountry.value);
				
			if(frmScreen.txtWhy.value != CustomerXML.GetTagText("REASONFORCOUNTRYOFRESIDENCY"))
				CustomerXML.SetTagText("REASONFORCOUNTRYOFRESIDENCY", frmScreen.txtWhy.value);
				
			if(frmScreen.optRightToWorkYes.checked)
				m_sRightToWork = "1";
			else if(frmScreen.optRightToWorkNo.checked)
				m_sRightToWork = "0";
			else
				m_sRightToWork = "";	
						
			CustomerXML.SetTagText("RIGHTTOWORK", m_sRightToWork);
			
			if(frmScreen.optDiplomaticImmunityYes.checked)
				m_sDiplomaticImmunity = "1";
			else if(frmScreen.optDiplomaticImmunityNo.checked)
				m_sDiplomaticImmunity = "0";
			else
				m_sDiplomaticImmunity = "";
				
			CustomerXML.SetTagText("DIPLOMATICIMMUNITY", m_sDiplomaticImmunity);
		}
	}
	return(CustomerXML.XMLDocument.xml);
}
-->
</script>
</body>
</HTML>
