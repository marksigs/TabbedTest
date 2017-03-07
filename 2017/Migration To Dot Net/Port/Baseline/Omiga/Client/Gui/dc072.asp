<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc072.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Additional Property Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		02/02/2000	ie5 changes 
AY		14/02/00	Change to msgButtons button types
AY		30/03/00	New top menu/scScreenFunctions change
BG		08/08/00	Added window functionality to allow it to save and 
					retrieve data with validation etc.
BS		01/03/00	SYS2564/SYS4211 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
PSC		09/07/2002	BMIDS00006  Change to use new account structure
GD		18/11/2002	BMIDS00376	Make readonly if appropriate
MC		19/04/2004	CC057		blank space padded to the title text to hide IE default message.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
	<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Additional Property Details  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>
<script src="validation.js" language="JScript"></script>

<% //Span to keep tabbing within this screen %>
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 225px; WIDTH: 460px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Property Details</strong>
	</span>
	<span style="TOP: 30px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<LABEL id="idValuerName"></LABEL> <% /* SYS2564/SYS4211 BS */ %>
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<input id="txtLastValuerName" name="LastValuerName" title="" maxlength="50" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 56px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<LABEL id="idValuationDate"></LABEL> <% /* SYS2564/SYS4211 BS */ %>
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<input id="txtLastValuationDate" name="LastValuationDate" title="" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 82px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		<LABEL id="idValuationAmount"></LABEL> <% /* SYS2564/SYS4211 BS */ %>
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<input id="txtLastValuationAmount" name="LastValuationAmount" title="" maxlength="7" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 108px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Reinstatement Amount
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<input id="txtReinstatementAmount" name="ReinstatementAmount" title="" maxlength="7" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 134px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Estimated Current Value
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<input id="txtEstimatedCurrentValue" name="EstimatedCurrentValue" title="" maxlength="7" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 160px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Type of Insurance
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<select id="cboHomeInsuranceType" name="HomeInsuranceType" title="" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	</span>
	<span style="TOP: 186px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Buildings Sum Insured
		<span style="TOP: -3px; LEFT: 170px; POSITION: ABSOLUTE">
			<input id="txtBuildingsSumInsured" name="BuildingsSumInsured" title="" maxlength="7" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
</div>
</form>

<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 245px; WIDTH: 460px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% //span to keep tabbing within this screen %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="attribs/DC072attribs.asp" -->
<!--  #include FILE="Customise/DC072Customise.asp" --> <% /* SYS2564/SYS4211 BS */ %>

<!-- Specify Code Here -->
<script language="JScript">
<!--
var m_sMetaAction = null;
var scScreenFunctions;
var PropertyXML = null;
var bUpdate = false;
var m_BaseNonPopupWindow = null;
var PropertySaveXML;
var m_sReadOnly = false;

function window.onload()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];	
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
	
	<%	/* Make the required buttons available on the bottom of the screen
		(see msgButtons.asp for details) */ %>
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	RetrieveContextData(sArgArray);
	
	SetMasks();
	// BS SYS2564/SYS4211 for client customisation
	Customise();
	
	PopulateCombo();
	PopulateScreen();
	
	Validation_Init();
					
	
	//GD BMIDS00376
	if (m_sReadOnly == true)
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	} else
	{
		scScreenFunctions.SetFocusToFirstField(frmScreen);
	}
}

<% /* keep the focus within	this screen when using the tab key */%>
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<%/* BG 08/08/00 Retrieves all the data required for this screen */%>
function RetrieveContextData(sArgArray)
{
	
	PropertyXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	PropertyXML.LoadXML(sArgArray[0]);
	//GD BMIDS00376
	m_sReadOnly = sArgArray[1];
}

function PopulateScreen()
{	
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	if(PropertyXML.SelectTag(null, "PROPERTYDETAILS") != null)
	{	
		with (frmScreen)
		{
			<% /* PSC 08/08/2002 BMIDS00006 - Start */ %>
			frmScreen.txtLastValuerName.value = PropertyXML.GetTagText("LASTVALUERNAME");
			frmScreen.txtLastValuationDate.value = PropertyXML.GetTagText("LASTVALUATIONDATE");
			frmScreen.txtLastValuationAmount.value = PropertyXML.GetTagText("LASTVALUATIONAMOUNT");
			frmScreen.txtReinstatementAmount.value = PropertyXML.GetTagText("REINSTATEMENTAMOUNT");
			frmScreen.txtEstimatedCurrentValue.value = PropertyXML.GetTagText("ESTIMATEDCURRENTVALUE");
			frmScreen.txtBuildingsSumInsured.value = PropertyXML.GetTagText("BUILDINGSSUMINSURED");
			frmScreen.cboHomeInsuranceType.value = PropertyXML.GetTagText("HOMEINSURANCETYPE");
			<% /* PSC 08/08/2002 BMIDS00006 - End */ %>
		}
	}
}
function btnSubmit.onclick()
{	
	<%/* BG 08/08/00 Run Validation checks */%>
	if(frmScreen.onsubmit() && CheckValuationDate())
	{
		<% /* PSC 08/08/2002 BMIDS00006 - Start */ %>
		PropertyXML.SelectTag(null, "PROPERTYDETAILS")
		
		if (PropertyXML.ActiveTag != null)
		{
			PropertyXML.SetTagText("LASTVALUERNAME",frmScreen.txtLastValuerName.value);
			PropertyXML.SetTagText("LASTVALUATIONDATE",frmScreen.txtLastValuationDate.value);
			PropertyXML.SetTagText("LASTVALUATIONAMOUNT",frmScreen.txtLastValuationAmount.value);
			PropertyXML.SetTagText("REINSTATEMENTAMOUNT",frmScreen.txtReinstatementAmount.value);
			PropertyXML.SetTagText("ESTIMATEDCURRENTVALUE",frmScreen.txtEstimatedCurrentValue.value);
			PropertyXML.SetTagText("BUILDINGSSUMINSURED",frmScreen.txtBuildingsSumInsured.value);
			PropertyXML.SetTagText("HOMEINSURANCETYPE",frmScreen.cboHomeInsuranceType.value);
		}
		else
		{
			PropertyXML.CreateActiveTag("PROPERTYDETAILS");
			PropertyXML.CreateTag("LASTVALUERNAME",frmScreen.txtLastValuerName.value);
			PropertyXML.CreateTag("LASTVALUATIONDATE",frmScreen.txtLastValuationDate.value);
			PropertyXML.CreateTag("LASTVALUATIONAMOUNT",frmScreen.txtLastValuationAmount.value);
			PropertyXML.CreateTag("REINSTATEMENTAMOUNT",frmScreen.txtReinstatementAmount.value);
			PropertyXML.CreateTag("ESTIMATEDCURRENTVALUE",frmScreen.txtEstimatedCurrentValue.value);
			PropertyXML.CreateTag("BUILDINGSSUMINSURED",frmScreen.txtBuildingsSumInsured.value);
			PropertyXML.CreateTag("HOMEINSURANCETYPE",frmScreen.cboHomeInsuranceType.value);
		}
		var sReturn = new Array();
		sReturn[0] = IsChanged();
		sReturn[1] = PropertyXML.XMLDocument.xml;
		window.returnValue = sReturn
		<% /* PSC 08/08/2002 BMIDS00006 - End */ %>
		
		window.close();
	}
}
<%/* BG 08/08/00 On cancel ignore all changes and close the window */%>
function btnCancel.onclick()
{		
	<% /* PSC 08/08/2002 BMIDS00006 - Start */ %>	
	var sReturn = new Array();
	sReturn[0] = IsChanged();
	sReturn[1] = PropertyXML.XMLDocument.xml;
	window.returnValue = sReturn
	<% /* PSC 08/08/2002 BMIDS00006 - End */ %>
	
	window.close();
}

function PopulateCombo()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("InsuranceType");

	if(XML.GetComboLists(document,sGroupList))
		// Populate Country combo
		blnSuccess = XML.PopulateCombo(document,frmScreen.cboHomeInsuranceType,"InsuranceType",true);

	XML = null;
}

<%/* BG 08/08/00 Check the date entered is before the current system date */%>
function CheckValuationDate()
{
	if(scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtLastValuationDate,">") == true)
	{
		alert("Last Valuation Date must be earlier than current Date");
		frmScreen.txtLastValuationDate.focus();
		return false;
	}
	return true;
}
-->
</script>
</body>

</html>
