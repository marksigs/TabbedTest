<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP203.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:	Valuation Report - Property Essential Repairs
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JD		19/09/2005	MAR40 Created screen
JD		08/11/2005	MAR434 Changed title to make it shorter
JD		18/11/2005  MAR647 Check that the task has an instructionSeqNo. If not, get it from the context.
PE		01/02/2007	EP2_1029 - Modified txtEssentialMatters to use RepairsNotes.
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
	<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 320px" class="msgGroup">
		<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Essential Matters
		<span style="LEFT: 0px; POSITION: absolute; TOP: 16px">
			<TEXTAREA id=txtEssentialMatters rows=6 style="WIDTH: 470px" maxlength="600" class=msgTxt ></TEXTAREA>
		</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 120px" class="msgLabel">	
			Structural Movement?
			<span style="LEFT: 250px; POSITION: absolute; TOP: -3px">
				<input id="StructuralMovement_Yes" name="StructuralMovement" type="radio" value="1"><label for="StructuralMovement_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
				<input id="StructuralMovement_No" name="StructuralMovement" type="radio" value="0"><label for="StructuralMovement_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 140px" class="msgLabel">	
			Significant and Progressive Structural Movement?
			<span style="LEFT: 250px; POSITION: absolute; TOP: -3px">
				<input id="ProgStructuralMovement_Yes" name="ProgStructuralMovement" type="radio" value="1"><label for="ProgStructuralMovement_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
				<input id="ProgStructuralMovement_No" name="ProgStructuralMovement" type="radio" value="0"><label for="ProgStructuralMovement_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 160px" class="msgLabel">	
			Serious Rot?
			<span style="LEFT: 250px; POSITION: absolute; TOP: -3px">
				<input id="SeriousRot_Yes" name="SeriousRot" type="radio" value="1"><label for="SeriousRot_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
				<input id="SeriousRot_No" name="SeriousRot" type="radio" value="0"><label for="SeriousRot_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 180px" class="msgLabel">	
			Asbestos Poor Condition?
			<span style="LEFT: 250px; POSITION: absolute; TOP: -3px">
				<input id="AsbestosPoorCondition_Yes" name="AsbestosPoorCondition" type="radio" value="1"><label for="AsbestosPoorCondition_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
				<input id="AsbestosPoorCondition_No" name="AsbestosPoorCondition" type="radio" value="0"><label for="AsbestosPoorCondition_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 200px" class="msgLabel">	
			Cavity Wall Tie Failure?
			<span style="LEFT: 250px; POSITION: absolute; TOP: -3px">
				<input id="CavityWallFailure_Yes" name="CavityWallFailure" type="radio" value="1"><label for="CavityWallFailure_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
				<input id="CavityWallFailure_No" name="CavityWallFailure" type="radio" value="0"><label for="CavityWallFailure_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 220px" class="msgLabel">	
			Large Panel System Appraised?
			<span style="LEFT: 250px; POSITION: absolute; TOP: -3px">
				<input id="LargePanelSystem_Yes" name="LargePanelSystem" type="radio" value="1"><label for="LargePanelSystem_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
				<input id="LargePanelSystem_No" name="LargePanelSystem" type="radio" value="0"><label for="LargePanelSystem_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 240px" class="msgLabel">	
			Historic Building Repairs Required?
			<span style="LEFT: 250px; POSITION: absolute; TOP: -3px">
				<input id="HistoricBuilding_Yes" name="HistoricBuilding" type="radio" value="1"><label for="HistoricBuilding_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
				<input id="HistoricBuilding_No" name="HistoricBuilding" type="radio" value="0"><label for="HistoricBuilding_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 260px" class="msgLabel">	
			ReType Indicator?
			<span style="LEFT: 250px; POSITION: absolute; TOP: -3px">
				<input id="ReType_Yes" name="ReType" type="radio" value="1"><label for="ReType_Yes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
				<input id="ReType_No" name="ReType" type="radio" value="0"><label for="ReType_No" class="msgLabel">No</label>
			</span>
		</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 280px" class="msgLabel">	
			Original Lender
			<span style="LEFT: 250px; POSITION: absolute; TOP: -3px">
				<input id="txtOriginalLender"  maxlength=40 style="WIDTH: 180px"  class="msgTxt">
			</span>
		</span>
	</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP:  430px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP203attribs.asp" -->

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
%>	var sButtonList = new Array("Submit", "Cancel");

	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Property Essential Repairs","AP203",scScreenFunctions);

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
function SetOptionValue( objOptionYes, objOptionNo , sVal )
{
	if( sVal == "1" )
	{
		objOptionYes.checked = true;
	}
	else 
		if( sVal == "0" )
		{
			objOptionNo.checked = true;
		}
}
function GetOptionValue( objOption )
{
	var sVal = "0";
	
	if(objOption.checked == true)
	{
		sVal = "1";
	}
	return(sVal)
}
function PopulateScreen()
{
	if(m_XML != null)
	{
		frmScreen.txtEssentialMatters.value = m_XML.GetAttribute("REPAIRSNOTES"); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.StructuralMovement_Yes, frmScreen.StructuralMovement_No, m_XML.GetAttribute("PROPERTYSUBSIDENCE")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.ProgStructuralMovement_Yes, frmScreen.ProgStructuralMovement_No, m_XML.GetAttribute("LONGSTANDINGSUBIDENCE")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.SeriousRot_Yes, frmScreen.SeriousRot_No, m_XML.GetAttribute("TIMBERDAMPREPORTREQ")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.AsbestosPoorCondition_Yes, frmScreen.AsbestosPoorCondition_No, m_XML.GetAttribute("ASBESTOSPOORCONDITION")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.CavityWallFailure_Yes, frmScreen.CavityWallFailure_No, m_XML.GetAttribute("CAVITYWALLTIEFAILURE")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.LargePanelSystem_Yes, frmScreen.LargePanelSystem_No, m_XML.GetAttribute("LARGEPANELSYSTEMAPPRAISED")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.HistoricBuilding_Yes, frmScreen.HistoricBuilding_No, m_XML.GetAttribute("HISTORICBUILDINGREPAIRSREQ")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.ReType_Yes, frmScreen.ReType_No, m_XML.GetAttribute("RETYPEIND")); //ValnRepPropertyRisks
		frmScreen.txtOriginalLender.value = m_XML.GetAttribute("RETYPEORIGINALLENDERNAME"); //ValnRepPropertyRisks
		
	}
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

	m_sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	if(m_sXML != "")
	{
		m_XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_XML.LoadXML(m_sXML);
		m_XML.SelectTag(null, "GETVALUATIONREPORT");
	}
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
}
function btnSubmit.onclick()
{
	//update the valuation report for the new fields
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
		
		ValuationXML.SetAttribute("REPAIRSNOTES", frmScreen.txtEssentialMatters.value);
		ValuationXML.SetAttribute("PROPERTYSUBSIDENCE", GetOptionValue(frmScreen.StructuralMovement_Yes));
		ValuationXML.SetAttribute("LONGSTANDINGSUBIDENCE", GetOptionValue(frmScreen.ProgStructuralMovement_Yes));
		ValuationXML.SetAttribute("TIMBERDAMPREPORTREQ", GetOptionValue(frmScreen.SeriousRot_Yes));
		ValuationXML.SetAttribute("ASBESTOSPOORCONDITION", GetOptionValue(frmScreen.AsbestosPoorCondition_Yes));
		ValuationXML.SetAttribute("CAVITYWALLTIEFAILURE", GetOptionValue(frmScreen.CavityWallFailure_Yes) ); 
		ValuationXML.SetAttribute("LARGEPANELSYSTEMAPPRAISED", GetOptionValue(frmScreen.LargePanelSystem_Yes));
		ValuationXML.SetAttribute("HISTORICBUILDINGREPAIRSREQ", GetOptionValue(frmScreen.HistoricBuilding_Yes)); 
		ValuationXML.SetAttribute("RETYPEIND", GetOptionValue(frmScreen.ReType_Yes)); 
		ValuationXML.SetAttribute("RETYPEORIGINALLENDERNAME", frmScreen.txtOriginalLender.value); 

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

-->
</script>
</STRONG>
</body>
</html>


