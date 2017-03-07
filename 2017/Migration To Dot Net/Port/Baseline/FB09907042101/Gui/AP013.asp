<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP013.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Application review Add/Overide
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		06/03/01	Added method calls.
JLD		08/03/01	Fix for setting under review indicator.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

BMIDS History:

Prog	Date		Description
ASu		18/10/2002	Where Application Mode = 'Override' set mandatory fields
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client

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
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Forms Here */ %>
<form id="frmToAP011" method="post" action="AP011.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange">
<div id="divReviewBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 160px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span id="spnReview" style="TOP: 10px; LEFT: 40px; POSITION: ABSOLUTE">
	<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
		Review Reason
		<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
			<select id="cboReviewReason" style="WIDTH:300px" class="msgCombo"></select>
		</span>
	</span>
	<span style="TOP:60px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
		Comments
		<span style="TOP:0px; LEFT:100px; POSITION:ABSOLUTE">
			<textarea id="txtReviewComments" rows="5" style="WIDTH:300px" class="msgTxt"></textarea>
		</span>
	</span>
</span>
</div>
<div id="divOverrideBackground" style="TOP: 230px; LEFT: 10px; HEIGHT: 160px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span id="spnReview" style="TOP: 10px; LEFT: 40px; POSITION: ABSOLUTE">
	<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
		Override Reason
		<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
			<select id="cboOverrideReason" style="WIDTH:300px" class="msgCombo"></select>
		</span>
	</span>
	<span style="TOP:60px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
		Comments
		<span style="TOP:0px; LEFT:100px; POSITION:ABSOLUTE">
			<textarea id="txtOverrideComments" rows="5" style="WIDTH:300px" class="msgTxt"></textarea>
		</span>
	</span>
</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/AP013Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_sReasonXML = "";
var m_ReasonXML = null;
var m_sAuthUserId = "";
var m_sContext = "";
var m_sApplicationNumber = "";
var m_sApplicationFFNumber = "";
var m_sUnitId = "";
var scScreenFunctions;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Application Review Add/Override","AP013",scScreenFunctions);

	RetrieveContextData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	<% /* ASu BMIDS00253 - Start. Please note that the SetMasks() is only called whilst in 'Override' Mode to set mandatory 
	fields and should the 'Set Masks' function be required in any other mode the calling .asp will need additional intelligence */ %>
	if (m_sContext == "Override")
		SetMasks();
	<% /* ASu - End */ %>	
	Validation_Init();
	PopulateScreen();
// AQR SYS 3134 - Screen should never be readonly
//	scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP013");
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function RetrieveContextData()
{
<%  //m_sMetaAction will be one of 'Add_userid' 'View_' 'Override_userid'. We need
	//to pick off the userId and context from this.
%>	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReasonXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFFNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sUnitId = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	
	var sArray = m_sMetaAction.split("_");
	m_sContext = sArray[0];
	if(sArray.length > 1) m_sAuthUserId = sArray[1];
}
function PopulateScreen()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();			
	var sGroups = new Array("ReviewReason", "OverrideReason");
	if (XML.GetComboLists(document, sGroups) == true)
	{
		XML.PopulateCombo(document, frmScreen.cboReviewReason, "ReviewReason", false);
		XML.PopulateCombo(document, frmScreen.cboOverrideReason, "OverrideReason", false);
	}
	XML = null;
	scScreenFunctions.SetCollectionState(divReviewBackground, "R");
	scScreenFunctions.SetCollectionState(divOverrideBackground, "R");
	
	if(m_sReasonXML != "")
	{
		m_ReasonXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_ReasonXML.LoadXML(m_sReasonXML);
		m_ReasonXML.SelectTag(null, "APPLICATIONREVIEWHISTORY");
	}
	if(m_sContext == "Override")
	{
		frmScreen.cboReviewReason.value = m_ReasonXML.GetAttribute("REVIEWREASON");
		frmScreen.txtReviewComments.value = m_ReasonXML.GetAttribute("REVIEWCOMMENTS");
		scScreenFunctions.SetCollectionState(divOverrideBackground, "W");
	}
	else if(m_sContext == "View")
	{
		frmScreen.cboReviewReason.value = m_ReasonXML.GetAttribute("REVIEWREASON");
		frmScreen.txtReviewComments.value = m_ReasonXML.GetAttribute("REVIEWCOMMENTS");
		frmScreen.cboOverrideReason.value = m_ReasonXML.GetAttribute("OVERRIDEREASON");
		frmScreen.txtOverrideComments.value = m_ReasonXML.GetAttribute("OVERRIDECOMMENTS");
	}
	else scScreenFunctions.SetCollectionState(divReviewBackground, "W");	
}
function btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
	{
		if(IsChanged())
		{
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			if(m_sContext == "Add")
				XML.CreateRequestTag(window, "CreateApplicationReviewHistory");
			else if(m_sContext == "Override")
				XML.CreateRequestTag(window, "UpdateApplicationReviewHistory");
			XML.CreateActiveTag("APPLICATIONREVIEWHISTORY");
			XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber)
			XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFFNumber)
			if(m_sContext == "Add")
			{
				XML.SetAttribute("REVIEWREASON", frmScreen.cboReviewReason.value);
				<% /* MO - BMIDS00807 */ %>
				<% /* XML.SetAttribute("REVIEWDATETIME", scScreenFunctions.DateTimeToString(scScreenFunctions.GetSystemDateTime())); */ %>
				XML.SetAttribute("REVIEWDATETIME", scScreenFunctions.DateTimeToString(scScreenFunctions.GetAppServerDate(true)));
				XML.SetAttribute("REVIEWUSERID", m_sAuthUserId);
				XML.SetAttribute("REVIEWUNITID", m_sUnitId);
				XML.SetAttribute("REVIEWCOMMENTS", frmScreen.txtReviewComments.value);
				// 				XML.RunASP(document, "omAppProc.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									XML.RunASP(document, "omAppProc.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
					}

				if(XML.IsResponseOK())
				{
					scScreenFunctions.SetContextParameter(window,"idAppUnderReview","1");
				}
			}
			else if(m_sContext == "Override")
			{
				XML.SetAttribute("REVIEWREASON", m_ReasonXML.GetAttribute("REVIEWREASON"));
				XML.SetAttribute("REVIEWDATETIME", m_ReasonXML.GetAttribute("REVIEWDATETIME"));
				XML.SetAttribute("OVERRIDENBYUSERID", m_sAuthUserId);
				XML.SetAttribute("OVERRIDENBYUNITID", m_sUnitId);
				XML.SetAttribute("OVERRIDEREASON", frmScreen.cboOverrideReason.value);
				XML.SetAttribute("OVERRIDECOMMENTS", frmScreen.txtOverrideComments.value);
				
				<% /* MO - BMIDS00807 */ %>
				<% /* XML.SetAttribute("OVERRIDEDATETIME", scScreenFunctions.DateTimeToString(scScreenFunctions.GetSystemDateTime())); */ %>
				XML.SetAttribute("OVERRIDEDATETIME", scScreenFunctions.DateTimeToString(scScreenFunctions.GetAppServerDate(true)));
				// 				XML.RunASP(document, "omAppProc.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									XML.RunASP(document, "omAppProc.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
					}

				if(XML.IsResponseOK())
				{
					if(parseInt(m_ReasonXML.GetAttribute("OUTSTANDINGREASONS")) == 1)
						scScreenFunctions.SetContextParameter(window,"idAppUnderReview","0");
				}
			}
		}
		frmToAP011.submit();
	}
}
function btnCancel.onclick()
{
	frmToAP011.submit();
}
-->
</script>
</body>
</html>




