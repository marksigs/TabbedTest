<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      DC077.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Add/Edit Mortgage Account Owner
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
PSC		31/07/2002	BMIDS00006 Created
PSC		23/08/2002	BMIDS00356 Amend screen title
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		Description
DS		26/03/2007	EP2_1794 - In Edit mode, Changed Applicant name to be non editable

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<head>
<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">

<title>Add/Edit Existing Account Owner <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 170px; WIDTH: 350px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP:40px; LEFT:20px; POSITION:ABSOLUTE" class="msgLabel">
		Applicant
		<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
			<select id="cboApplicant" style="WIDTH:200px" class="msgCombo"></select>
		</span>
	</span>
	<span style="TOP:70px; LEFT:20px; POSITION:ABSOLUTE" class="msgLabel">
		Role
		<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
			<select id="cboRole" style="WIDTH:200px" class="msgCombo"></select>
		</span>
	</span>
	<span style="TOP:110px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH:60px" class="msgButton">
	</span>
	<span style="TOP:110px; LEFT:180px; POSITION:ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH:60px" class="msgButton">
	</span>
</div>
</form>

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_BaseNonPopupWindow = null;
var m_XMLApplicants = null;
var m_XMLOwners = null;
var m_sMetaAction = null;
var m_sAccountGuid = null;
var m_blnReadOnly = false;

function window.onload()
{
	RetrieveData();
	Validation_Init();
	SetScreenOnReadOnly();
	PopulateApplicantCombo();
	SetUpComboList();
	
	window.returnValue = null;
}
function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	if (m_blnReadOnly == true)
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
	m_XMLApplicants = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_XMLOwners = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	m_sMetaAction = sParameters[0];
	m_XMLOwners.LoadXML(sParameters[1]);
	m_XMLApplicants.LoadXML(sParameters[2]);
	m_blnReadOnly = sParameters[4];
	
	var nIndex = null;
	
	<% /* If in edit mode set the active tag to the position in the XML structure where the edit
	      is to take place */ %>
	if (m_sMetaAction == "Edit")
	{
		<%/*  DS - EP2_1794 */ %>
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboApplicant");
		nIndex = sParameters[3];
		m_XMLOwners.CreateTagList("ACCOUNTRELATIONSHIP");
		m_XMLOwners.SelectTagListItem(nIndex);
	}
	else
		m_sAccountGuid = sParameters[3];
}
function frmScreen.btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}

function frmScreen.btnOK.onclick()
{


	if (frmScreen.cboRole.value != "")
	{
		var nSelectedRole = frmScreen.cboRole.selectedIndex;
		var RoleOption = frmScreen.cboRole.options.item(nSelectedRole);

		if (m_sMetaAction == "Add")
		{
			m_XMLOwners.SelectTag(null,"ACCOUNTRELATIONSHIPLIST")
			m_XMLOwners.CreateActiveTag("ACCOUNTRELATIONSHIP");
			
			if (m_sAccountGuid != null)
				m_XMLOwners.CreateTag("ACCOUNTGUID", m_sAccountGuid);
				
			m_XMLOwners.CreateTag("CUSTOMERNUMBER", frmScreen.cboApplicant.value);
			
			var nSelectedCustomer = frmScreen.cboApplicant.selectedIndex;
			var CustOption = frmScreen.cboApplicant.options.item(nSelectedCustomer);
			
			m_XMLOwners.CreateTag("CUSTOMERVERSIONNUMBER",CustOption.getAttribute("CustomerVersionNumber"));
			m_XMLOwners.CreateTag("CUSTOMERROLETYPE", frmScreen.cboRole.value);
			m_XMLOwners.SetAttributeOnTag("CUSTOMERROLETYPE", "TEXT", RoleOption.text);
			m_XMLOwners.CreateTag("CUSTOMERNUMBER", frmScreen.cboApplicant.value);
			m_XMLOwners.CreateTag("CUSTOMERVERSIONNUMBER",CustOption.getAttribute("CustomerVersionNumber"));
			
			var nSpace = CustOption.text.indexOf(" ");
			
			m_XMLOwners.CreateTag("FIRSTFORENAME",CustOption.text.substring(0,nSpace));
			m_XMLOwners.CreateTag("SURNAME",CustOption.text.substring(nSpace + 1));	
		}
		else
		{
			m_XMLOwners.SetTagText("CUSTOMERROLETYPE", frmScreen.cboRole.value);	
			m_XMLOwners.SetAttributeOnTag("CUSTOMERROLETYPE", "TEXT", RoleOption.text);
		}
		
		var sReturn = new Array();
		sReturn[0] = IsChanged();
		sReturn[1] = m_XMLOwners.XMLDocument.xml;
		window.returnValue = sReturn
		
		window.close();
	}
	else
	{
		alert ("A valid role must be selected");
	}
}

function PopulateApplicantCombo()
{

	m_XMLApplicants.CreateTagList("CUSTOMER");

	var nCustomerCount = m_XMLApplicants.ActiveTagList.length;
			
	<% /* Loop through all customer context entries */ %>
	for(var nLoop = 0; nLoop < nCustomerCount; nLoop++)
	{
		m_XMLApplicants.SelectTagListItem(nLoop);
		
		TagOPTION = document.createElement("OPTION");
		TagOPTION.value	= m_XMLApplicants.GetTagText("CUSTOMERNUMBER");
		TagOPTION.text = m_XMLApplicants.GetTagText("CUSTOMERNAME");
		TagOPTION.setAttribute("CustomerVersionNumber", m_XMLApplicants.GetTagText("CUSTOMERVERSIONNUMBER"));
		frmScreen.cboApplicant.add(TagOPTION);
	}
			
	frmScreen.cboApplicant.selectedIndex = 0;
}

function SetUpComboList()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups;
			
	sGroups = new Array("CustomerRoleType");
			
	if(XML.GetComboLists(document, sGroups) == true)
		XML.PopulateCombo(document, frmScreen.cboRole,"CustomerRoleType",true);
		
	<% /* If in edit mode set the combo to the currently selected value */ %>	
	if (m_sMetaAction == "Edit")		
		frmScreen.cboRole.value = m_XMLOwners.GetTagText("CUSTOMERROLETYPE");

	XML = null;
}

-->
</script>
</body>
</html>
