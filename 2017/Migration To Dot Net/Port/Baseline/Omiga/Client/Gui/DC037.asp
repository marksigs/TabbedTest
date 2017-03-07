<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<HTML>
  <HEAD>
	<title>Additional Personal Details<!-- #include file="includes/TitleWhiteSpace.asp" --></title>
<%/*
Workfile:      DC037.asp
Copyright:     Copyright © 2006 Vertex Financial Services
Author			Ian Ross

Description:   Additional Personal Details  POP-UP
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
INR		16/01/2007	Created
Ashaw	22/02/2007	EP2_1747 - Disable OK Btn if Read only.
*/ %>
	<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
  </HEAD>

<body>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript" type="text/javascript"></script>

<%/* FORMS */%>
<form id="frmScreen" mark validate="onchange">
	<div id="divResidency" style="LEFT: 10px; WIDTH: 500px; POSITION: absolute; TOP: 10px; HEIGHT: 60px" class="msgGroup">
		<span style="LEFT: 20px; POSITION: absolute; TOP: 20px" class="msgLabel">Have you been in receipt of housing benefit in the last 12 months?
			<span style="LEFT: 340px; POSITION: absolute; TOP: -4px">
				<input id="optHousingBenefitYes" name="HousingBenefitGroup" type="radio" value="1"><label for="optHousingBenefitYes" class="msgLabel">Yes</label>
			</span> 
			<span style="LEFT: 400px; POSITION: absolute; TOP: -4px">
				<input id="optHousingBenefitNo" name="HousingBenefitGroup" type="radio" value="0"><label for="optHousingBenefitNo" class="msgLabel">No</label>
			</span> 
		</span>

		<span style="LEFT: 20px; POSITION: absolute; TOP: 70px">
			<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton" NAME="btnOK">
			<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">
				<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton" NAME="btnCancel">
			</span>
		</span>	
	</div>			
</form>

<!--  #include FILE="attribs/DC037attribs.asp" -->

<%/* CODE */%>
<script language="JScript" type="text/javascript">
<!--
var m_sXMLIn = "";
var m_sReadOnly = "";
var CustomerXML = null;
var m_sReceiveHouseBenefit = "";
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
	{
		window.returnValue = null;
		// EP2_1747 - Disable OK Btn if Read only.
		frmScreen.btnOK.disabled = true;
	}
}
function PopulateScreen()
{
	CustomerXML.SelectTag(null,"CUSTOMERVERSION");
	scScreenFunctions.SetRadioGroupValue(frmScreen, "HousingBenefitGroup", CustomerXML.GetTagText("HOUSINGBENEFIT"));		
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
	var sReturn = new Array();
	
	// EP2_1007 - Must enter Benefit option
	if(!frmScreen.optHousingBenefitNo.checked && !frmScreen.optHousingBenefitYes.checked)
	{
		alert("Please enter Housing Benefit details");
		frmScreen.optHousingBenefitYes.focus();			
		return false;
	}


	sReturn[0] = IsChanged();		// Has there been a change made
	sReturn[1] = WriteHousingBenefitDetails();		// The XML string

	window.returnValue = sReturn;
	window.close();
}
function frmScreen.btnCancel.onclick()
{
	window.close();
}

function WriteHousingBenefitDetails()
{
	//save any changed data
	if ((m_sReadOnly != "1") && frmScreen.onsubmit())
	{
		if (IsChanged() == true )
		{
			//CustomerXML.ActiveTag = null;
			CustomerXML.SelectTag(null,"CUSTOMERVERSION");
			
			if(frmScreen.optHousingBenefitYes.checked)
				m_sHousingBenefit = "1";
			else if(frmScreen.optHousingBenefitNo.checked)
				m_sHousingBenefit = "0";
			else
				m_sHousingBenefit = "";	
						
			CustomerXML.SetTagText("HOUSINGBENEFIT", m_sHousingBenefit);
			
		}
	}
	return(CustomerXML.XMLDocument.xml);
}
-->
</script>
</body>
</HTML>
