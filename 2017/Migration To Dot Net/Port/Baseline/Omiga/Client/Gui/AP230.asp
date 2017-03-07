<%@ LANGUAGE="JSCRIPT" %>
<%
/*
Workfile:      AP230.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		01/02/01	SYS1839 Created
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
DRC     21/11/01    SYS3048 Get Instruction Sequence No if not present in the Task Context from
                    the omiga menu context
BG		03/11/2001  SYS3048	moved "var m_ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();"
					from declarations at top of JScript to GetValuationReport.	
DB		18/12/2001	SYS3401 If optpropertysubsidence = true then enabled optlongsubs & optaffectproperty.					
AT		30/04/02	SYS4359 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
IK		12/05/2006	EP540 - revised layout
EPSOM 2
-------
AShaw	22/11/2006	EP2_2 - Three new 'Construction' option fields.
PEdney	02/02/2007	EP2_1136 - WP8 - Xit2 Interface - missing UI fields
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<html>
	<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
	<head>
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
		<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
		<title></title>
		<style>
			.radioLabel1{width:25%;padding-top:4px;vertical-align:top}
			.radioLabel2{width:28%;padding-top:4px;padding-left:2px;vertical-align:top}
		</style>
	</head>
	<body style="width:610px">
		<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
		<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			type="text/x-scriptlet" width="1" viewastext tabindex="-1">
		</object>
		<!-- Form Manager Object - Controls Soft Coded Field Attributes -->
		<script src="validation.js" language="JScript"></script> <!-- Screen Functions Object - see comments within scScreenFunctions.htm for details -->
		<% /* Specify Forms Here */ %>
		<form id="frmToAP205" method="post" action="AP205.asp" style="DISPLAY: none">
		</form>
		<% /* Span to keep tabbing within this screen */ %>
		<span id="spnToLastField" tabindex="0"></span>
		<form id="frmScreen" validate="onchange" mark>
			<!--<div id="divBackground" style="HEIGHT: 450px; LEFT: 10px; POSITION: absolute; TOP: 40px; WIDTH: 604px" class="msgGroup">-->
			<!-- Screen section one -->
			<div style="position:relative;top:0;margin-top:65;left:0;margin-left:8" class="msgLabel">
				<span style="font-weight:bold">Subsidence</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span id="spnLocalSubsidence">
					<span class="radioLabel1">Local Subsidence?</span> 
					<span style="width:10%;vertical-align:top">
						<input id="optLocalSubsidenceYes" name="LocalSubsidenceRadioGroup" type="radio" value="1">
						<label for="optLocalSubsidenceYes" class="msgLabel">Yes</label>
					</span>
					<span style="width:12%;vertical-align:top">
						<input id="optLocalSubsidenceNo" name="LocalSubsidenceRadioGroup" type="radio" value="0">
						<label for="optLocalSubsidenceNo" class="msgLabel">No</label>
					</span>
				</span>
				<span id="spnLongStandingSubsidence">
					<span class="radioLabel2">Longstanding Subsidence?</span> 
					<span style="width:10%;vertical-align:top">
						<input id="optLongSubsYes" name="LongSubsRadioGroup" type="radio" value="1">
						<label for="optLongSubsYes" class="msgLabel">Yes</label>
					</span>
					<span style="vertical-align:top">
						<input id="optLongSubsNo" name="LongSubsRadioGroup" type="radio" value="0">
						<label for="optLongSubsNo" class="msgLabel">No</label>
					</span>
				</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span id="spnAffectProperty">
					<span class="radioLabel1">Likely to Affect Property?</span> 
					<span style="width:10%;vertical-align:top">
						<input id="optAffectPropertyYes" name="AffectPropertyRadioGroup" type="radio" value="1">
						<label for="optAffectPropertyYes" class="msgLabel">Yes</label>
					</span>
					<span style="width:12%;vertical-align:top">
						<input id="optAffectPropertyNo" name="AffectPropertyRadioGroup" type="radio" value="0">
						<label for="optAffectPropertyNo" class="msgLabel">No</label>
					</span>
				</span>
				<span class="radioLabel2">Structural Engineer's Report Required?</span> 
				<span style="width:10%;vertical-align:top">
					<input id="optStructuralReportYes" name="StructuralReportRadioGroup" type="radio" value="1">
					<label for="optStructuralReportYes" class="msgLabel">Yes</label>
				</span>
				<span style="vertical-align:top">
					<input id="optStructuralReportNo" name="StructuralReportRadioGroup" type="radio" value="0">
					<label for="optStructuralReportNo" class="msgLabel">No</label>
				</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span id="spnPropertySubsidence">
					<span class="radioLabel1">Property Subsidence?</span> 
					<span style="width:10%;vertical-align:top">
						<input id="optPropertySubsidenceYes" name="PropertySubsidenceRadioGroup" type="radio" value="1">
						<label for="optPropertySubsidenceYes" class="msgLabel">Yes</label>
					</span>
					<span style="width:12%;vertical-align:top">
						<input id="optPropertySubsidenceNo" name="PropertySubsidenceRadioGroup" type="radio" value="0">
						<label for="optPropertySubsidenceNo" class="msgLabel">No</label>
					</span>
				</span>
				<span class="radioLabel2">Affect Sale?</span> 
				<span style="width:10%;vertical-align:top">
					<input id="optAffectSaleYes" name="AffectSaleRadioGroup" type="radio" value="1">
					<label for="optAffectSaleYes" class="msgLabel">Yes</label>
				</span>
				<span style="vertical-align:top">
					<input id="optAffectSaleNo" name="AffectSaleRadioGroup" type="radio" value="0">
					<label for="optAffectSaleNo" class="msgLabel">No</label>
				</span>
			</div>
			<div style="position:relative;top:0;margin-top:12;left:0;margin-left:8" class="msgLabel">
				<span style="font-weight:bold">Construction</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span id="Span1">
					<span class="radioLabel1">Risk of Mundic problems?</span> 
					<span style="width:10%;vertical-align:top">
						<input id="optMundicYes" name="MundicRadioGroup" type="radio" value="1">
						<label for="optMundicYes" class="msgLabel">Yes</label>
					</span>
					<span style="width:12%;vertical-align:top">
						<input id="optMundicNo" name="MundicRadioGroup" type="radio" value="0">
						<label for="optMundicNo" class="msgLabel">No</label>
					</span>
				</span>
				<span id="Span2">
					<span class="radioLabel2">Longstanding and Non-Progressive<br>Structural Issues?</span> 
					<span style="width:10%;vertical-align:top">
						<input id="optLongStructYes" name="LongStructRadioGroup" type="radio" value="1">
						<label for="optLongStructYes" class="msgLabel">Yes</label>
					</span>
					<span style="vertical-align:top">
						<input id="optLongStructNo" name="LongStructRadioGroup" type="radio" value="0">
						<label for="optLongStructNo" class="msgLabel">No</label>
					</span>
				</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span id="Span3">
					<span class="radioLabel1">Structural Issues?</span> 
					<span style="width:10%;vertical-align:top">
						<input id="optStructuralIssuesYes" name="StructuralRadioGroup" type="radio" value="1">
						<label for="optStructuralIssuesYes" class="msgLabel">Yes</label>
					</span>
					<span style="width:12%;vertical-align:top">
						<input id="optStructuralIssuesNo" name="StructuralRadioGroup" type="radio" value="0">
						<label for="optStructuralIssuesNo" class="msgLabel">No</label>
					</span>
				</span>
			</div>
			<div style="position:relative;top:0;margin-top:12;left:0;margin-left:8" class="msgLabel">
				<span style="font-weight:bold">Risk from Trees</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span id="spnTreeReportRequired">
					<span class="radioLabel1">Trees Report Required?</span> 
					<span style="width:10%;vertical-align:top">
						<input id="optTreeReportYes" name="TreeReportRadioGroup" type="radio" value="1">
						<label for="optTreeReportYes" class="msgLabel">Yes</label>
					</span>
					<span style="width:12%;vertical-align:top">
							<input id="optTreeReportNo" name="TreeReportRadioGroup" type="radio" value="0">
							<label for="optTreeReportNo" class="msgLabel">No</label>
					</span>
				</span>
				<span class="radioLabel2">Tree Height - ft</span> 
				<span style="width:18%;vertical-align:top">
					<input id="txtTreeHeight" maxlength="10" style="width: 100px" class="msgTxt" name="txtTreeHeight">				
				</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span class="radioLabel1">Nature of Risks</span> 
				<span style="width:22%;vertical-align:top">
					<select id="cboNatureOfRisks" style="width:90%" class="msgCombo" name="cboNatureOfRisks"></select>				</span>
				<span class="radioLabel2">Tree Distance - ft</span> 
				<span style="width:18%;vertical-align:top">
					<input id="txtTreeDistance" maxlength="10" style="width: 100px" class="msgTxt" name="txtTreeDistance">				
				</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span class="radioLabel1">Tree Type</span> 
				<span style="width:22%;vertical-align:top">
					<select id="cboTreeType" style="width:90%" class="msgCombo" name="cboTreeType"></select>				
				</span>
				<span class="radioLabel2">Within Curtiledge?</span> 
				<span style="width:10%;vertical-align:top">
					<input id="optWithinCurtiledgeYes" name="WithinCurtiledgeRadioGroup" type="radio" value="1">
					<label for="optWithinCurtiledgeYes" class="msgLabel">Yes</label>
				</span>
				<span style="vertical-align:top">
					<input id="optWithinCurtiledgeNo" name="WithinCurtiledgeRadioGroup" type="radio" value="0">
					<label for="optWithinCurtiledgeNo" class="msgLabel">No</label>
				</span>
			</div>
			<div style="position:relative;top:0;margin-top:12;left:0;margin-left:8" class="msgLabel">
				<span style="font-weight:bold">Reports</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span class="radioLabel1">Mining Reports Required?</span> 
				<span style="width:10%;vertical-align:top">
					<input id="optMiningReportYes" name="MiningReportRadioGroup" type="radio" value="1">
					<label for="optMiningReportYes" class="msgLabel">Yes</label>
				</span>
				<span style="width:12%;vertical-align:top">
					<input id="optMiningReportNo" name="MiningReportRadioGroup" type="radio" value="0">
					<label for="optMiningReportNo" class="msgLabel">No</label>
				</span>
				<span class="radioLabel2">Timber and Damp Report Required?</span> 
				<span style="width:10%;vertical-align:top">
					<input id="optTimberYes" name="TimberReportRadioGroup" type="radio" value="1">
					<label for="optTimberYes" class="msgLabel">Yes</label>
				</span>
				<span style="vertical-align:top">
					<input id="optTimberNo" name="TimberReportRadioGroup" type="radio" value="0">
					<label for="optTimberNo" class="msgLabel">No</label>
				</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span id="spnSpecialistReport">
					<span class="radioLabel1">Specialist Report Required?</span> 
					<span style="width:10%;vertical-align:top">
						<input id="optSpecialistReportYes" name="SpecialistReportRadioGroup" type="radio" value="1">
						<label for="optSpecialistReportYes" class="msgLabel">Yes</label>
					</span>
					<span style="width:12%;vertical-align:top">
						<input id="optSpecialistReportNo" name="SpecialistReportRadioGroup" type="radio" value="0">
						<label for="optSpecialistReportNo" class="msgLabel">No</label>
					</span>
				</span>
				<span class="radioLabel2">Electrical Report Required?</span> 
				<span style="width:10%;vertical-align:top">
					<input id="optElectriclReportYes" name="ElectriclReportRadioGroup" type="radio" value="1">
					<label for="optElectriclReportYes" class="msgLabel">Yes</label>
				</span>
				<span style="vertical-align:top">
					<input id="optElectriclReportNo" name="ElectriclReportRadioGroup" type="radio" value="0">
					<label for="optElectriclReportNo" class="msgLabel">No</label>
				</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span class="radioLabel1">Specialist Report</span> 
				<span style="width:22%;vertical-align:top">
					<select id="cboSpecialistReport" style="width:90%" class="msgCombo" name="cboSpecialistReport"></select>				
				</span>
				<span id="spnSolicitorRef">
					<span class="radioLabel2">Solicitor Reference Required?</span> 
					<span style="width:10%;vertical-align:top">
						<input id="optSolicitorRefYes" name="SolicotorReferenceReq" type="radio" value="1">
						<label for="optSolicitorRefYes" class="msgLabel">Yes</label>
					</span>
					<span style="vertical-align:top">
						<input id="optSolicitorRefNo" name="SolicotorReferenceReq" type="radio" value="0">
						<label for="optSolicitorRefNo" class="msgLabel">No</label>
					</span>
				</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span class="radioLabel1">Prone to Flooding?</span> 
				<span style="width:10%;vertical-align:top">
					<input id="optFloodingProneYes" name="FloodingProneRadioGroup" type="radio" value="1">
					<label for="optFloodingProneYes" class="msgLabel">Yes</label>
				</span>
				<span style="width:12%;vertical-align:top">
					<input id="optFloodingProneNo" name="FloodingProneRadioGroup" type="radio" value="0">
					<label for="optFloodingProneNo" class="msgLabel">No</label>
				</span>
				<span style="width:46%" class="msgLabel">
					<textarea class="msgTxt" id="txtSolicitorNotes" rows="3" style="width:100%" name="txtSolicitorNotes"></textarea>
				</span>
			</div>
			<div style="position:relative;top:0;margin-top:12;left:0;margin-left:8" class="msgLabel">
				<span style="font-weight:bold">Locality</span>
			</div>			
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span class="radioLabel1">Radon Gas Area</span> 
				<span style="width:10%;vertical-align:top">
					<input id="optRadonAreaYes" name="RadonAreaRadioGroup" type="radio" value="1">
					<label for="optRadonAreaYes" class="msgLabel">Yes</label>
				</span>
				<span style="width:12%;vertical-align:top">
					<input id="optRadonAreaNo" name="RadonAreaRadioGroup" type="radio" value="0">
					<label for="optRadonAreaNo" class="msgLabel">No</label>
				</span>
				<span class="radioLabel1">Allowed for in Valuation</span> 
				<span style="width:10%;vertical-align:top">
					<input id="optAllowedForYes" name="AllowedForRadioGroup" type="radio" value="1">
					<label for="optAllowedForYes" class="msgLabel">Yes</label>
				</span>
				<span style="width:12%;vertical-align:top">
					<input id="optAllowedForNo" name="AllowedForRadioGroup" type="radio" value="0">
					<label for="optAllowedForNo" class="msgLabel">No</label>
				</span>
			</div>
			<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
				<span class="radioLabel1">Boarded up or Damaged Property in area</span> 
				<span style="width:10%;vertical-align:top">
					<input id="optBoardedUpYes" name="BoardedUpRadioGroup" type="radio" value="1">
					<label for="optMiningReportYes" class="msgLabel">Yes</label>
				</span>
				<span style="width:12%;vertical-align:top">
					<input id="optBoardedUpNo" name="BoardedUpRadioGroup" type="radio" value="0">
					<label for="optBoardedUpNo" class="msgLabel">No</label>
				</span>
			</div>
		</form>
		<%/* Main Buttons */ %>
		<div style="position:relative;top:0;left:0;margin-top:8;margin-left:10;width:100%" class="msgLabel">
			<!-- #include FILE="msgButtons.asp" -->
		</div>		
		</div>
		<% /* Span to keep tabbing within this screen */ %>
		<span id="spnToFirstField" tabindex="0"></span>
	</body>
	<!-- #include FILE="fw030.asp" -->
		<% /* File containing field attributes (remove if not required) */ %>
		<!-- #include FILE="attribs/AP230attribs.asp" -->
		<!-- #include FILE="Customise/AP230Customise.asp" -->
		<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
//BG 03/11/2001 SYS3048 can't var as new here
//var m_ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
var m_ValuationXML = null;
var m_TaskXML = null;
var m_sTaskXML = "";
var m_sAppNo = "";
var m_sAppFactFindNo = "";
var m_sTaskID = "";
var m_sInsSeqNo = "";
var csValuerTypeStaff = "10";
var csTaskComplete = "40";
var csTaskPending = "20";
var csTaskNotStarted = "10";
var m_sCallingScreen = "";

var m_sReadOnly;
var m_sReportNotFound = "0";

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
	FW030SetTitles("Valuation Report - Risks & Reports","AP230",scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
	Customise();
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
	PopulateScreen();
	
 	scScreenFunctions.SetFocusToFirstField(frmScreen);
	if(m_sReadOnly == "1")
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
		}
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function PopulateScreen()
{
	var bSuccess;

	PopulateCombos();
	bSuccess = GetValuationReport();

	if(bSuccess)
	{
		// Populate the screen fields
		frmScreen.cboNatureOfRisks.value = m_ValuationXML.GetAttribute("NATUREOFRISKS");
		frmScreen.cboTreeType.value = m_ValuationXML.GetAttribute("TREETYPE");
		frmScreen.cboSpecialistReport.value = m_ValuationXML.GetAttribute("SPECIALISTREPORT");

		SetOptionValue(frmScreen.optLocalSubsidenceYes, frmScreen.optLocalSubsidenceNo, m_ValuationXML.GetAttribute("LOCALSUBSIDENCE"))
		SetOptionValue(frmScreen.optAffectPropertyYes, frmScreen.optAffectPropertyNo, m_ValuationXML.GetAttribute("LIKELYTOAFFECTPROPERTY"))
		SetOptionValue(frmScreen.optPropertySubsidenceYes, frmScreen.optPropertySubsidenceNo, m_ValuationXML.GetAttribute("PROPERTYSUBSIDENCE"))
		SetOptionValue(frmScreen.optLongSubsYes, frmScreen.optLongSubsNo, m_ValuationXML.GetAttribute("LONGSTANDINGSUBIDENCE"))
		SetOptionValue(frmScreen.optStructuralReportYes, frmScreen.optStructuralReportNo, m_ValuationXML.GetAttribute("STRUCENGREPORTREQ"))
		SetOptionValue(frmScreen.optAffectSaleYes, frmScreen.optAffectSaleNo, m_ValuationXML.GetAttribute("AFFECTSALE"))
		SetOptionValue(frmScreen.optWithinCurtiledgeYes, frmScreen.optWithinCurtiledgeNo, m_ValuationXML.GetAttribute("WITHINCURTILEDGE"))
		SetOptionValue(frmScreen.optMiningReportYes, frmScreen.optMiningReportNo, m_ValuationXML.GetAttribute("MININGREPORTSREQ"))
		SetOptionValue(frmScreen.optSpecialistReportYes, frmScreen.optSpecialistReportNo, m_ValuationXML.GetAttribute("SPECIALISTREPORTREQ"))
		SetOptionValue(frmScreen.optFloodingProneYes, frmScreen.optFloodingProneNo, m_ValuationXML.GetAttribute("PRONETOFLOODING"))
		SetOptionValue(frmScreen.optTimberYes, frmScreen.optTimberNo, m_ValuationXML.GetAttribute("TIMBERDAMPREPORTREQ"))
		SetOptionValue(frmScreen.optElectriclReportYes, frmScreen.optElectriclReportNo, m_ValuationXML.GetAttribute("ELECTICALREPORTREQ"))
		SetOptionValue(frmScreen.optSolicitorRefYes, frmScreen.optSolicitorRefNo, m_ValuationXML.GetAttribute("SOLICITORREFERENCEREQ"))

		SetOptionValue(frmScreen.optTreeReportYes, frmScreen.optTreeReportNo, m_ValuationXML.GetAttribute("TREEREPORTREQ"))
		SetOptionValue(frmScreen.optLongSubsYes, frmScreen.optLongSubsNo, m_ValuationXML.GetAttribute("LONGSTANDINGSUBIDENCE"))
		SetOptionValue(frmScreen.optStructuralReportYes, frmScreen.optStructuralReportNo, m_ValuationXML.GetAttribute("STRUCENGREPORTREQ"))

		frmScreen.txtTreeHeight.value = m_ValuationXML.GetAttribute("TREEHEIGHT");
		frmScreen.txtTreeDistance.value = m_ValuationXML.GetAttribute("TREEDISTANCE");
		frmScreen.txtSolicitorNotes.value = m_ValuationXML.GetAttribute("SOLICITORNOTES");
		//EP2_2 New fields
		SetOptionValue(frmScreen.optMundicYes, frmScreen.optMundicNo, m_ValuationXML.GetAttribute("MUNDICRISK"));
		SetOptionValue(frmScreen.optLongStructYes, frmScreen.optLongStructNo, m_ValuationXML.GetAttribute("STRUCTURALISSUES"));
		SetOptionValue(frmScreen.optStructuralIssuesYes, frmScreen.optStructuralIssuesNo, m_ValuationXML.GetAttribute("NONPROGRESSIVESTRUCTURALISSUES"));
		
		// EP2_1136 - WP8 - Xit2 Interface - missing UI fields
		SetOptionValue(frmScreen.optRadonAreaYes, frmScreen.optRadonAreaNo, m_ValuationXML.GetAttribute("RADONGASAREA"));
		SetOptionValue(frmScreen.optAllowedForYes, frmScreen.optAllowedForNo, m_ValuationXML.GetAttribute("RADONGASALLOWEDFOR"));
		SetOptionValue(frmScreen.optBoardedUpYes, frmScreen.optBoardedUpNo, m_ValuationXML.GetAttribute("BOARDEDUPORDAMAGEDPROPERTY"));
		
	}	
	else
	{
		frmScreen.optLocalSubsidenceNo.checked ="1";
		frmScreen.optPropertySubsidenceNo.checked ="1";
		frmScreen.optLongSubsNo.checked ="1";
		frmScreen.optStructuralReportNo.checked ="1";
		frmScreen.optAffectSaleNo.checked ="1";
		frmScreen.optWithinCurtiledgeNo.checked ="1";
		frmScreen.optMiningReportNo.checked ="1";
		frmScreen.optSpecialistReportNo.checked ="1";
		frmScreen.optFloodingProneNo.checkedvalue ="1";
		frmScreen.optTimberNo.checked ="1";
		frmScreen.optElectriclReportNo.checked ="1";
		frmScreen.optSolicitorRefNo.checked ="1";																				
		frmScreen.optTreeReportNo.checked ="1";																				
		frmScreen.optLongSubsNo.checked ="1";																				
		frmScreen.optStructuralReportNo.checked ="1";																				
		frmScreen.optAffectPropertyNo.checked ="1";																				
		frmScreen.optFloodingProneNo.checked ="1";	
		m_sReportNotFound = "1";	
		//EP2_2 New fields
		frmScreen.optMundicNo.checked ="1";	
		frmScreen.optLongStructNo.checked ="1";	
		frmScreen.optStructuralIssuesNo.checked ="1";	
	}
	if(frmScreen.txtSolicitorNotes.value.length > 0 && frmScreen.optSolicitorRefYes.checked == true )

	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtSolicitorNotes")			
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtSolicitorNotes")			
	}

	SetRadioState();
	return(bSuccess);
}

function SetRadioState()
{
	SetLocalSubsidenceState();
	SetTreeReportState();
	SetSpecialisRepState();
}

function SetPropertySubsidenceState()
{
	// If property subsidence is 1 enable likely to affect property
	// DB SYS3401 18/12/01 - Enable & disable optlongsubs, optaffectproperty dependent on value of optpropertysubsidence

	if (frmScreen.optPropertySubsidenceYes.checked == true) 
	//	frmScreen.optLongSubsYes.checked == true && 
	//	frmScreen.optAffectPropertyYes.checked == true ) // &&

	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"optAffectSaleNo")
		scScreenFunctions.SetFieldToWritable(frmScreen,"optAffectSaleYes")
		scScreenFunctions.SetFieldToWritable(frmScreen,"optStructuralReportNo")
		scScreenFunctions.SetFieldToWritable(frmScreen,"optStructuralReportYes")
		// DB 18/12/01 - Enable long standing subsidence & affect property
		scScreenFunctions.SetFieldToWritable(frmScreen,"optLongSubsNo")
		scScreenFunctions.SetFieldToWritable(frmScreen,"optLongSubsYes")
		scScreenFunctions.SetFieldToWritable(frmScreen,"optAffectPropertyNo")
		scScreenFunctions.SetFieldToWritable(frmScreen,"optAffectPropertyYes")
	}
	else
	{
		frmScreen.optStructuralReportNo.checked = true;
		frmScreen.optAffectSaleNo.checked = true;
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optStructuralReportNo")
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optStructuralReportYes")
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optAffectSaleNo")
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optAffectSaleYes")
		// DB 18/12/01 - Else set to read only
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optLongSubsNo")
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optLongSubsYes")
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optAffectPropertyNo")
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optAffectPropertyYes")
	}
}

function spnTreeReportRequired.onclick()
{
	SetTreeReportState();
}

function SetTreeReportState()
{
	if(frmScreen.optTreeReportYes.checked == true)

	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboNatureOfRisks")
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboTreeType")
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtTreeDistance")
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtTreeHeight")			
		scScreenFunctions.SetFieldToWritable(frmScreen,"optWithinCurtiledgeNo")			
		scScreenFunctions.SetFieldToWritable(frmScreen,"optWithinCurtiledgeYes")			
	}
	else
	{
		frmScreen.optWithinCurtiledgeNo.checked = true;
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboNatureOfRisks")
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboTreeType")
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTreeDistance")
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTreeHeight")			
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optWithinCurtiledgeNo")			
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optWithinCurtiledgeYes")			
		frmScreen.cboTreeType.value = "";
		frmScreen.cboNatureOfRisks.value = ""
		frmScreen.txtTreeDistance.value = "";
		frmScreen.txtTreeHeight.value = "";
	}
}

function SetLocalSubsidenceState()
{
	// If local subsidence is 1 enable likely to affect property
	if(frmScreen.optLocalSubsidenceYes.checked == true )

	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"optAffectPropertyNo")
		scScreenFunctions.SetFieldToWritable(frmScreen,"optAffectPropertyYes")
	}
	else
	{
		frmScreen.optAffectPropertyNo.checked = true;
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optAffectPropertyNo")
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optAffectPropertyYes")
	}
	SetPropertySubsidenceState();
}

function spnSpecialistReport.onclick()
{
	SetSpecialisRepState();
}

function SetSpecialisRepState()
{
	if(frmScreen.optSpecialistReportYes.checked == true)

	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboSpecialistReport")
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboSpecialistReport")
	}
}

function spnPropertySubsidence.onclick()
{
	SetPropertySubsidenceState();
}
function spnLongStandingSubsidence.onclick()
{
	SetPropertySubsidenceState();
}
function spnAffectProperty.onclick()
{
	SetPropertySubsidenceState();
}

function spnLocalSubsidence.onclick()
{
	SetLocalSubsidenceState();
}

function GetValuationReport()
{
	//BG 03/11/2001 SYS3048 
	m_ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = m_ValuationXML.CreateRequestTag(window , "GetValuationReport");
	var bSuccess = false;
	
	m_ValuationXML.CreateActiveTag("VALUATION");
	m_ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	m_ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
	m_ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
	// AQR SYS3048 - if Instruction sequence number is not in CASETASK context then
	//               get it from the main context (set in AP205)
	if (m_sInsSeqNo == "" )
	{
		m_sInsSeqNo = scScreenFunctions.GetContextParameter(window,"idInstructionSequenceNo",null)
	}
	m_ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO",m_sInsSeqNo);

	m_ValuationXML.RunASP(document, "omAppProc.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = m_ValuationXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)	
	{
		if(m_ValuationXML.SelectTag(null, "GETVALUATIONREPORT") != null)
		{
			bSuccess = true;
		}
	}
	return(bSuccess);
}

function PopulateCombos()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("NatureOfRisk","TreeType", "SpecialistReport");

	if(XML.GetComboLists(document,sGroupList))
	{
		XML.PopulateCombo(document,frmScreen.cboNatureOfRisks,"NatureOfRisk",true);
		XML.PopulateCombo(document,frmScreen.cboTreeType,"TreeType",true);
		XML.PopulateCombo(document,frmScreen.cboSpecialistReport,"SpecialistReport",true);
	}

	XML = null;
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
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	if(m_sTaskXML != "")
	{
		m_TaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_TaskXML.LoadXML(m_sTaskXML);
		m_TaskXML.SelectTag(null, "CASETASK");
		m_sInsSeqNo = m_TaskXML.GetAttribute("CONTEXT");
		m_sTaskID = m_TaskXML.GetAttribute("TASKID")	
	}

	m_sAppNo = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sAppFactFindNo = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sCallingScreen = scScreenFunctions.GetContextParameter(window,"idREturnScreenId",null);
   	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idMetaAction", null);	
}

function CommitChanges()
{
	var bSuccess = true;
	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				if (ValidateScreen())
					bSuccess = SaveValuation();
				else
					bSuccess = false;
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function ValidateScreen()
{
	var bSuccess = true;
	return(bSuccess);
}

function ValidateInvoiceAmount()
{
	var bSuccess = true;
	return( bSuccess );
}

function SaveScreenValues( XML )
{
	var bSuccess;
	var sVal;
	
	bSuccess = false
	if(XML != null)
	{
		XML.SetAttribute("NATUREOFRISKS", frmScreen.cboNatureOfRisks.value);
		XML.SetAttribute("TREETYPE", frmScreen.cboTreeType.value);
		XML.SetAttribute("SPECIALISTREPORT", frmScreen.cboSpecialistReport.value);

		XML.SetAttribute("LOCALSUBSIDENCE", GetOptionValue(frmScreen.optLocalSubsidenceYes));
		XML.SetAttribute("LIKELYTOAFFECTPROPERTY", GetOptionValue(frmScreen.optAffectPropertyYes));
		XML.SetAttribute("PROPERTYSUBSIDENCE", GetOptionValue(frmScreen.optPropertySubsidenceYes));
		XML.SetAttribute("LONGSTANDINGSUBIDENCE", GetOptionValue(frmScreen.optLongSubsYes));
		XML.SetAttribute("STRUCENGREPORTREQ", GetOptionValue(frmScreen.optStructuralReportYes));
		XML.SetAttribute("AFFECTSALE", GetOptionValue(frmScreen.optAffectSaleYes));
		XML.SetAttribute("WITHINCURTILEDGE", GetOptionValue(frmScreen.optWithinCurtiledgeYes));
		XML.SetAttribute("MININGREPORTSREQ", GetOptionValue(frmScreen.optMiningReportYes));
		XML.SetAttribute("SPECIALISTREPORTREQ", GetOptionValue(frmScreen.optSpecialistReportYes));				
		XML.SetAttribute("PRONETOFLOODING", GetOptionValue(frmScreen.optFloodingProneYes));				
		XML.SetAttribute("ELECTICALREPORTREQ", GetOptionValue(frmScreen.optElectriclReportYes));				
		XML.SetAttribute("SOLICITORREFERENCEREQ", GetOptionValue(frmScreen.optSolicitorRefYes));				
		XML.SetAttribute("TREEREPORTREQ", GetOptionValue(frmScreen.optTreeReportYes));				
		XML.SetAttribute("WITHINCURTILEDGE", GetOptionValue(frmScreen.optWithinCurtiledgeYes));				
		XML.SetAttribute("TIMBERDAMPREPORTREQ", GetOptionValue(frmScreen.optTimberYes));				
		XML.SetAttribute("TREEHEIGHT", frmScreen.txtTreeHeight.value);
		XML.SetAttribute("TREEDISTANCE", frmScreen.txtTreeDistance.value);
		XML.SetAttribute("SOLICITORNOTES", frmScreen.txtSolicitorNotes.value);
		// EP2_2 New fields
		XML.SetAttribute("MUNDICRISK", GetOptionValue(frmScreen.optMundicYes));				
		XML.SetAttribute("STRUCTURALISSUES", GetOptionValue(frmScreen.optLongStructYes));				
		XML.SetAttribute("NONPROGRESSIVESTRUCTURALISSUES", GetOptionValue(frmScreen.optStructuralIssuesYes));				

		// EP2_1136 - WP8 - Xit2 Interface - missing UI fields
		XML.SetAttribute("RADONGASAREA", GetOptionValue(frmScreen.optRadonAreaYes));				
		XML.SetAttribute("RADONGASALLOWEDFOR", GetOptionValue(frmScreen.optAllowedForYes));				
		XML.SetAttribute("BOARDEDUPORDAMAGEDPROPERTY", GetOptionValue(frmScreen.optBoardedUpYes));				

		bSuccess = true;
	}

	return(bSuccess);
}

function SaveValuation()
{
	var bSuccess = false;
	var ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTaskStatus;
	var sASPFile;
		
	sTaskStatus = m_TaskXML.GetAttribute("TASKSTATUS");
	
	if(m_sReportNotFound == "0") 
	{
		var reqTag = ValuationXML.CreateRequestTag(window , "UpdateValuationReport");
		ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");

		sASPFile = "omAppProc.asp"
		ValuationXML.CreateActiveTag("VALUATION");
		ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
		ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
		ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInsSeqNo);

		if(SaveScreenValues(ValuationXML))
		{
			bSuccess = true;
		}			
	}
	else if(m_sReportNotFound == "1")
	{
		var reqTag = ValuationXML.CreateRequestTag(window , "CreateValuationReport");
		ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");

		ValuationXML.CreateActiveTag("VALUATION");
		if(SaveScreenValues(ValuationXML))
		{
			sASPFile = "OmigaTmBO.asp"
			ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
			ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
			ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInsSeqNo);

			ValuationXML.ActiveTag = reqTag;
			ValuationXML.CreateActiveTag("CASETASK");
			ValuationXML.SetAttribute("ACTIVITYID", m_TaskXML.GetAttribute("ACTIVITYID"));
			ValuationXML.SetAttribute("STAGEID", m_TaskXML.GetAttribute("STAGEID"));
			ValuationXML.SetAttribute("CASESTAGESEQUENCENO", m_TaskXML.GetAttribute("CASESTAGESEQUENCENO"));
			ValuationXML.SetAttribute("CASEID", m_TaskXML.GetAttribute("CASEID"));
			ValuationXML.SetAttribute("TASKID", m_TaskXML.GetAttribute("TASKID"));
			ValuationXML.SetAttribute("TASKINSTANCE", m_TaskXML.GetAttribute("TASKINSTANCE"));
			
			bSuccess = true;
		}	
	}

	if(bSuccess)
	{
		// 		ValuationXML.RunASP(document, sASPFile);
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					ValuationXML.RunASP(document, sASPFile);
				break;
			default: // Error
				ValuationXML.SetErrorResponse();
			}

		if(ValuationXML.IsResponseOK())
		{
			// Update the context
			m_TaskXML.SetAttribute("TASKSTATUS",csTaskPending);
			scScreenFunctions.SetContextParameter(window,"idTaskXML", m_TaskXML.XMLDocument.xml);
			bSuccess = true;
		}
		else
		{
			bSuccess = false
		}
	}	

	return(bSuccess);
}

function GetOptionValue( objOption )
{
	var sVal = "0";
	
	if(objOption.checked == true)
	{
		sVal = "1"
	}
	return(sVal)
}

function SetOptionValue( objOptionYes, objOptionNo , sVal )
{
	if( sVal == "1" )
	{
		objOptionYes.checked = true;
	}
	else
	{
		objOptionNo.checked = true;
	}
}

function spnSolicitorRef.onclick()
{
	var sMessage = "Setting Solicitor Reference to No will remove all Solicitor Notes. Do you want to continue?";
	HandleOption(frmScreen.optSolicitorRefYes,frmScreen.txtSolicitorNotes,sMessage); 
}

function HandleOption( optYes, txtField, sMessage )
{
	var bRemarks;
	var bEnable = false;
	
	bRemarks = optYes.checked 

	if(bRemarks == true)
	{
		bEnable = true;
	}

	if(!bEnable && txtField.disabled == false && txtField.value.length > 0)
	{
		if (confirm(sMessage))
		{		
			txtField.value = "";
		}
		else
		{
			optYes.checked = true;
			bEnable = true;
		}
	}

	if( bEnable && m_sReadOnly != "1")
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,txtField.id)			
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,txtField.id)			
	}
}

function btnCancel.onclick()
{
	frmToAP205.submit();
}

function btnSubmit.onclick()
{
	if(CommitChanges())
	{
		frmToAP205.submit();
	}
}


-->
</script>
</html>
