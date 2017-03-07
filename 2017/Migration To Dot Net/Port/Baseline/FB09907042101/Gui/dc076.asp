<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      DC077.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Add/Edit Mortgage Account Features
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
PSC		05/08/2002	BMIDS00006 Created
GD		18/11/2002	BMIDS00376 Ensure readonly gets set
MC		20/04/2004	BMIDS517	White space padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">

<title>Add/Edit Special Features  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 170px; WIDTH: 350px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP:40px; LEFT:20px; POSITION:ABSOLUTE" class="msgLabel">
		Description
		<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
			<input id="txtDescription" maxlength="50" style="WIDTH:200px" class="msgTxt">
		</span>
	</span>
	<span style="TOP:70px; LEFT:20px; POSITION:ABSOLUTE" class="msgLabel">
		Indicator
		<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
			<input id="txtIndicator" maxlength="5" style="WIDTH:200px" class="msgTxt">
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

<!-- #include FILE="attribs/DC076attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_BaseNonPopupWindow = null;
var m_XMLFeatures;
var m_sMetaAction = null;
var m_sAccountGuid = null;
var m_RequestParams = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	RetrieveData();
	SetMasks();
	Validation_Init();
	SetScreenOnReadOnly();
	scScreenFunctions.SetFocusToFirstField(frmScreen);	
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
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
	m_XMLFeatures = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	m_RequestParams = sParameters[0];
	m_sMetaAction = sParameters[1];
	m_XMLFeatures.LoadXML(sParameters[2]);
	//GD BMIDS00376
	m_sReadOnly = sParameters[4];
	
	
	var nIndex = null;
	
	<% /* If in edit mode set the active tag to the position in the XML structure where the edit
	      is to take place */ %>
	if (m_sMetaAction == "Edit")
	{
		nIndex = sParameters[3];
		m_XMLFeatures.CreateTagList("MORTGAGEACCOUNTSPECIALFEATURE");
		m_XMLFeatures.SelectTagListItem(nIndex);
		frmScreen.txtDescription.value = m_XMLFeatures.GetTagText("MORTGAGEACCOUNTSPECIALFEATUREDESC");
		frmScreen.txtIndicator.value = m_XMLFeatures.GetTagText("MORTGAGEACCOUNTSPECIALFEATUREIND");
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
	if (frmScreen.txtDescription.value != "" && frmScreen.txtIndicator.value != "")
	{
		var XMLRequest = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject(); 
		XMLRequest.CreateRequestTagFromArray(m_RequestParams, null);

		if (m_sMetaAction == "Add")
		{
			m_XMLFeatures.SelectTag(null,"MORTGAGEACCOUNTSPECIALFEATURELIST")
			m_XMLFeatures.CreateActiveTag("MORTGAGEACCOUNTSPECIALFEATURE");
			m_XMLFeatures.CreateTag("ACCOUNTGUID", m_sAccountGuid);
			m_XMLFeatures.CreateTag("MORTGAGEACCOUNTSPECIALFEATUREIND", frmScreen.txtIndicator.value);
			m_XMLFeatures.CreateTag("MORTGAGEACCOUNTSPECIALFEATUREDESC", frmScreen.txtDescription.value);
		
			XMLRequest.ActiveTag.appendChild(m_XMLFeatures.ActiveTag.cloneNode(true));
			// 			XMLRequest.RunASP(document,"CreateSpecialFeature.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XMLRequest.RunASP(document,"CreateSpecialFeature.asp");
					break;
				default: // Error
					XMLRequest.SetErrorResponse();
				}

			
			if(XMLRequest.IsResponseOK())
			{
				XMLRequest.SelectTag(null,"GENERATEDKEYS");
				m_XMLFeatures.CreateTag("MORTGAGEACCOUNTSPECIALFEATURESEQNO", XMLRequest.GetTagText("MORTGAGEACCOUNTSPECIALFEATURESEQNO"));
			}
			else
				return;
		}
		else
		{
			m_XMLFeatures.SetTagText("MORTGAGEACCOUNTSPECIALFEATUREIND", frmScreen.txtIndicator.value);	
			m_XMLFeatures.SetTagText("MORTGAGEACCOUNTSPECIALFEATUREDESC", frmScreen.txtDescription.value);	

			XMLRequest.ActiveTag.appendChild(m_XMLFeatures.ActiveTag.cloneNode(true));
			// 			XMLRequest.RunASP(document,"UpdateSpecialFeature.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XMLRequest.RunASP(document,"UpdateSpecialFeature.asp");
					break;
				default: // Error
					XMLRequest.SetErrorResponse();
				}

			
			if (!XMLRequest.IsResponseOK())
				return;
		}
		
		var sReturn = new Array();
		sReturn[0] = IsChanged();
		sReturn[1] = m_XMLFeatures.XMLDocument.xml;
		window.returnValue = sReturn
		
		window.close();
	}
	else
	{
		alert ("A valid Description and Indicator must be entered");
	}
}

-->
</script>
</body>
</html>


