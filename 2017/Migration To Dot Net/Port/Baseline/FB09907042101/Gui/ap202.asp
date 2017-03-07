<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP202.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:	Valuation Report - General Observations
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JD		19/09/2005	MAR40 Created screen
JD		08/11/2005	MAR434 Changed title to make it shorter
JD		18/11/2005  MAR647 Check that the task has an instructionSeqNo. If not, get it from the context.
AW		02/05/2006	EP451  Limit text entry to database column size.
PE		01/02/2007	EP2_1029 - Map txtGeneralObservations to GeneralRemarks.
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css"><title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
data=scClientFunctions.asp width=1 height=1 type=text/x-scriptlet 
VIEWASTEXT></OBJECT><!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<script src="validation.js" language="JScript"></script>
<!-- Screen Functions Object - see comments within scScreenFunctions.htm for details -->

<% /* Specify Forms Here */ %>
<form id="frmToAP200" method="post" action="AP200.asp" STYLE="DISPLAY: none"></form>


<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here
	 Amended 28/08/2002 - for APWP3 by DPF
 -->
<form id="frmScreen" mark validate ="onchange">
	<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 200px" class="msgGroup">
		<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		General Observations
		<span style="LEFT: 0px; POSITION: absolute; TOP: 16px">
			<TEXTAREA id=txtGeneralObservations rows=10 style="WIDTH: 470px" maxlength="255" class=msgTxt ></TEXTAREA>
		</span>
	</span>		  
	</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP:  330px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP202attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sXML = "";
var m_XML = null;
var m_sReadOnly = "";
var m_sInsSeqNo = "";
var m_sAppNo = "";
var m_sAppFactFindNo = "";

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
	FW030SetTitles("General Observations","AP202",scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
	
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
	PopulateScreen();
	
 	scScreenFunctions.SetFocusToFirstField(frmScreen);

	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function PopulateScreen()
{
	// Populate the screen field from valnRepPropertyDetails.GeneralObservations
	if(m_XML != null)
		frmScreen.txtGeneralObservations.value = m_XML.GetAttribute("GENERALREMARKS");
	//scScreenFunctions.SetScreenToReadOnly(frmScreen);
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
	var sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	if(sTaskXML != "")
	{
		var TaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		TaskXML.LoadXML(sTaskXML);
		TaskXML.SelectTag(null, "CASETASK");
		m_sInsSeqNo = TaskXML.GetAttribute("CONTEXT");
	}
	//JD MAR647
	if (m_sInsSeqNo == "")
		m_sInsSeqNo = scScreenFunctions.GetContextParameter(window,"idInstructionSequenceNo",null);
		
	m_sAppNo = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sAppFactFindNo = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);

	if(m_sReadOnly != "1")
	{
		var sRet;
		sRet = scScreenFunctions.GetContextParameter(window,"idMetaAction", null);	
		if(sRet == "1")
		{
			m_sReadOnly = "1";
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			DisableMainButton("Submit");
		}
	}
	m_sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	if(m_sXML != "")
	{
		m_XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_XML.LoadXML(m_sXML);
		m_XML.SelectTag(null, "GETVALUATIONREPORT");
	}
}
function btnSubmit.onclick()
{
	//update the valuation report for the new text
	var bSuccess = true;
	if(IsChanged())
	{
		var ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var reqTag = ValuationXML.CreateRequestTag(window , "UpdateValuationReport");
		ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");

		ValuationXML.CreateActiveTag("VALUATION");
		ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
		ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
		ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInsSeqNo);
		ValuationXML.SetAttribute("GENERALREMARKS", frmScreen.txtGeneralObservations.value);
		ValuationXML.RunASP(document,"omAppProc.asp");
		bSuccess = ValuationXML.IsResponseOK()
	}
	if (bSuccess)
		frmToAP200.submit();
}
function btnCancel.onclick()
{
	frmToAP200.submit();
}

<% /* AW 02/05/2006 EP451 Start  */ %>
function frmScreen.txtGeneralObservations.onkeyup()
    {	
       	scScreenFunctions.RestrictLength(frmScreen, "txtGeneralObservations", 450, true);
    }
<% /* AW 02/05/2006 EP451 End  */ %>

-->
</script>
</STRONG>
</body>
</html>


