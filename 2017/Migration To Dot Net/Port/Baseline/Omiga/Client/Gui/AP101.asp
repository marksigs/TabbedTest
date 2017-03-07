<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP101.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Employers reference PRP details   POP-UP
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		16/01/01	Created
JLD		23/01/01	Don't save radio as NO if not set.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Employer's Reference PRP Details</title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* style="VISIBILITY: hidden" Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen"  mark validate="onchange" >
<div id="div1" style="TOP: 10px; LEFT: 10px; HEIGHT: 200px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Has the scheme been given Inland Revenue approval?
	<span style="TOP:-3px; LEFT:400px; POSITION:ABSOLUTE">
		<input id="optIRApprovalYes" name="RadioGroup1" type="radio" value="1">
		<label for="optIRApprovalYes" class="msgLabel">Yes</label>
	</span>
	<span style="TOP:-3px; LEFT:450px; POSITION:ABSOLUTE">
		<input id="optIRApprovalNo" name="RadioGroup1" type="radio" value="0">
		<label for="optIRApprovalNo" class="msgLabel">No</label>
	</span>
</span>
<span style="TOP:40px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Can the interim payments made under the scheme be reclaimed from the<BR>employee once paid?
	<span style="TOP:-3px; LEFT:400px; POSITION:ABSOLUTE">
		<input id="optInterimPaymentsReclaimYes" name="RadioGroup2" type="radio" value="1">
		<label for="optInterimPaymentsReclaimYes" class="msgLabel">Yes</label>
	</span>
	<span style="TOP:-3px; LEFT:450px; POSITION:ABSOLUTE">
		<input id="optInterimPaymentsReclaimNo" name="RadioGroup2" type="radio" value="0">
		<label for="optInterimPaymentsReclaimNo" class="msgLabel">No</label>
	</span>
</span>
<span style="TOP:90px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	If the profits for the period of pay are insufficient to make appropriate<BR>
	profit-related payments, will the employee's salary be restored to the level<BR>
	it was before the scheme began?
	<span style="TOP:-3px; LEFT:400px; POSITION:ABSOLUTE">
		<input id="optSalaryRestorationYes" name="RadioGroup3" type="radio" value="1">
		<label for="optSalaryRestorationYes" class="msgLabel">Yes</label>
	</span>
	<span style="TOP:-3px; LEFT:450px; POSITION:ABSOLUTE">
		<input id="optSalaryRestorationNo" name="RadioGroup3" type="radio" value="0">
		<label for="optSalaryRestorationNo" class="msgLabel">No</label>
	</span>
</span>
<span style="TOP:160px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Do you meet any tax liability if any part of the tax free payments become<BR>subject to taxation?
	<span style="TOP:-3px; LEFT:400px; POSITION:ABSOLUTE">
		<input id="optEmpTaxLiabilityYes" name="RadioGroup4" type="radio" value="1">
		<label for="optEmpTaxLiabilityYes" class="msgLabel">Yes</label>
	</span>
	<span style="TOP:-3px; LEFT:450px; POSITION:ABSOLUTE">
		<input id="optEmpTaxLiabilityNo" name="RadioGroup4" type="radio" value="0">
		<label for="optEmpTaxLiabilityNo" class="msgLabel">No</label>
	</span>
</span>
</div>
<div id="div2" style="TOP: 220px; LEFT: 10px; HEIGHT: 65px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Basic pay including Profit Related Pay
	<span style="TOP:0px; LEFT:350px; POSITION:ABSOLUTE">
		<label style="TOP:0px; LEFT:-10px; POSITION:absolute" class="msgCurrency">£</label>
		<input id="txtBasicPayPRP" maxlength="10" style="TOP:-3px; WIDTH:110px; POSITION:ABSOLUTE" class="msgTxt">
	</span>
</span>
<span style="TOP:35px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Frequency Profit Related Pay received
	<span style="TOP:-3px; LEFT:350px; POSITION:ABSOLUTE">
		<select id="cboPRPFrequency" style="WIDTH:110px" class="msgCombo"></select>
	</span>
</span>
</div>
<div id="div3" style="TOP: 295px; LEFT: 10px; HEIGHT: 30px; WIDTH: 510px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Does the address agree with your records?
	<span style="TOP:-3px; LEFT:400px; POSITION:ABSOLUTE">
		<input id="optAddressAgreeYes" name="RadioGroup5" type="radio" value="1">
		<label for="optAddressAgreeYes" class="msgLabel">Yes</label>
	</span>
	<span style="TOP:-3px; LEFT:450px; POSITION:ABSOLUTE">
		<input id="optAddressAgreeNo" name="RadioGroup5" type="radio" value="0">
		<label for="optAddressAgreeNo" class="msgLabel">No</label>
	</span>
</span>
<span style="TOP:40px; LEFT:4px; POSITION:ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH:60px" class="msgButton">
</span>
<span style="TOP:40px; LEFT:70px; POSITION:ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH:60px" class="msgButton">
</span>
</div>
</form>

<!-- #include FILE="attribs/AP101attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_sPRPXML = "";
var m_PRPXML = null;
var m_BaseNonPopupWindow = null;

function window.onload()
{
	RetrieveData();
	SetMasks();
	Validation_Init();
	PopulateScreen();
	SetScreenOnReadOnly();
	window.returnValue = null;
}
function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
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
	m_sReadOnly			= sParameters[0];
	m_sPRPXML			= sParameters[1];
}
function PopulateCombo()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("EmploymentPaymentFreq");
	if(XML.GetComboLists(document,sGroupList))
	{
		XML.PopulateCombo(document,frmScreen.cboPRPFrequency,"EmploymentPaymentFreq",false);
	}
	XML = null;
}
function PopulateScreen()
{
<%	// m_sPRPXML passed in looks like this :
	// <PRP IRAPPROVAL="1" PAYMENTSRECLAIM="1" SALARYRESTORATION="1" TAXLIABILITY="1"
	//      BASICPRP="20000" PRPFREQUENCY="1" ADDRESSAGREE="1" ></PRP>
%>	PopulateCombo();
	if(m_sPRPXML.length > 0)
	{
		m_PRPXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_PRPXML.LoadXML(m_sPRPXML);
		m_PRPXML.SelectTag(null, "PRP");
		if(m_PRPXML.GetAttribute("IRAPPROVAL") == "1")
			frmScreen.optIRApprovalYes.checked = true;
		else if(m_PRPXML.GetAttribute("IRAPPROVAL") == "0")
			frmScreen.optIRApprovalNo.checked = true;
		if(m_PRPXML.GetAttribute("PAYMENTSRECLAIM") == "1")
			frmScreen.optInterimPaymentsReclaimYes.checked = true;
		else if(m_PRPXML.GetAttribute("PAYMENTSRECLAIM") == "0")
			frmScreen.optInterimPaymentsReclaimNo.checked = true;
		if(m_PRPXML.GetAttribute("SALARYRESTORATION") == "1")
			frmScreen.optSalaryRestorationYes.checked = true;
		else if(m_PRPXML.GetAttribute("SALARYRESTORATION") == "0")
			frmScreen.optSalaryRestorationNo.checked = true;
		if(m_PRPXML.GetAttribute("TAXLIABILITY") == "1")
			frmScreen.optEmpTaxLiabilityYes.checked = true;
		else if(m_PRPXML.GetAttribute("TAXLIABILITY") == "0")
			frmScreen.optEmpTaxLiabilityNo.checked = true;
		frmScreen.txtBasicPayPRP.value = m_PRPXML.GetAttribute("BASICPRP");
		frmScreen.cboPRPFrequency.value = m_PRPXML.GetAttribute("PRPFREQUENCY");
		if(m_PRPXML.GetAttribute("ADDRESSAGREE") == "1")
			frmScreen.optAddressAgreeYes.checked = true;
		else if(m_PRPXML.GetAttribute("ADDRESSAGREE") == "0")
			frmScreen.optAddressAgreeNo.checked = true;
	}
}
function SaveXML()
{
	if(m_PRPXML == null)
	{
		m_PRPXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_PRPXML.CreateActiveTag("PRP");
	}
	else m_PRPXML.SelectTag(null, "PRP");
	if(frmScreen.optIRApprovalYes.checked == true)
		m_PRPXML.SetAttribute("IRAPPROVAL","1");
	else if(frmScreen.optIRApprovalNo.checked == true)
		m_PRPXML.SetAttribute("IRAPPROVAL","0");
	else m_PRPXML.SetAttribute("IRAPPROVAL","");
	if(frmScreen.optInterimPaymentsReclaimYes.checked == true)
		m_PRPXML.SetAttribute("PAYMENTSRECLAIM","1");
	else if(frmScreen.optInterimPaymentsReclaimNo.checked == true)
		m_PRPXML.SetAttribute("PAYMENTSRECLAIM","0");
	else m_PRPXML.SetAttribute("PAYMENTSRECLAIM","");
	if(frmScreen.optSalaryRestorationYes.checked == true)
		m_PRPXML.SetAttribute("SALARYRESTORATION","1");
	else if(frmScreen.optSalaryRestorationNo.checked == true)
		m_PRPXML.SetAttribute("SALARYRESTORATION","0");
	else m_PRPXML.SetAttribute("SALARYRESTORATION","");
	if(frmScreen.optEmpTaxLiabilityYes.checked == true)
		m_PRPXML.SetAttribute("TAXLIABILITY","1");
	else if(frmScreen.optEmpTaxLiabilityNo.checked == true)
		m_PRPXML.SetAttribute("TAXLIABILITY","0");
	else m_PRPXML.SetAttribute("TAXLIABILITY","");
	m_PRPXML.SetAttribute("BASICPRP",frmScreen.txtBasicPayPRP.value);
	m_PRPXML.SetAttribute("PRPFREQUENCY",frmScreen.cboPRPFrequency.value);
	if(frmScreen.optAddressAgreeYes.checked == true)
		m_PRPXML.SetAttribute("ADDRESSAGREE","1");
	else if(frmScreen.optAddressAgreeNo.checked == true)
		m_PRPXML.SetAttribute("ADDRESSAGREE","0");	
	else m_PRPXML.SetAttribute("ADDRESSAGREE","");
}
function frmScreen.btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}
function frmScreen.btnOK.onclick()
{
	var sReturn = new Array();
	SaveXML();
	sReturn[0] = IsChanged();
	sReturn[1] = m_PRPXML.XMLDocument.xml;
	window.returnValue	= sReturn;
	window.close();
}
-->
</script>
</body>
</html>
