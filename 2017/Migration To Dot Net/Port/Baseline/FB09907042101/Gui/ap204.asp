<%@ LANGUAGE="JSCRIPT" %>
<HTML>
	<HEAD>
		<title></title>
		<%
/*
Workfile:      AP204.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:	Valuation Report - Comments for Solicitor
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JD		19/09/2005	MAR40 Created screen
JD		08/11/05	MAR434 Changed title to make it shorter
JD		18/11/2005  MAR647 Check that the task has an instructionSeqNo. If not, get it from the context.

Epsom2
------
AShaw	22/11/2006	EP2_2 - 11 New fields.
PEdney	01/02/2007	EP2_1029 - Changed MiningArea to use MININGAREA.
PEdney	19/03/2008	EP2_1548 - Default to unchecked.
*/

%>
		<META content="Microsoft Visual Studio 6.0" name="GENERATOR">
		<LINK href="stylesheet.css" type="text/css" rel="STYLESHEET">
	</HEAD>
	<BODY>
		<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
		<OBJECT id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px"
			tabIndex="-1" type="text/x-scriptlet" height="1" width="1" data="scClientFunctions.asp"
			VIEWASTEXT>
		</OBJECT>
		<!-- Form Manager Object - Controls Soft Coded Field Attributes -->
		<SCRIPT language="JScript" src="validation.js"></SCRIPT>
		<!-- Screen Functions Object - see comments within scScreenFunctions.htm for details -->
		<% /* Specify Forms Here */ %>
		<FORM id="frmToAP200" style="DISPLAY: none" action="AP200.asp" method="post">
		</FORM>
		<% /* Span to keep tabbing within this screen */ %>
		<SPAN id="spnToLastField" tabIndex="0"></SPAN><!-- Specify Screen Layout Here
	 Amended 28/08/2002 - for APWP3 by DPF
 -->
		<FORM id="frmScreen" validate="onchange" mark>
			<DIV class="msgGroup" id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 470px"><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 10px">Comments For 
Solicitor <SPAN style="LEFT: 0px; POSITION: absolute; TOP: 16px"><TEXTAREA class="msgTxt" id="txtComments" style="WIDTH: 470px" rows="6" maxlength="255"></TEXTAREA>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 120px">Rights of Way? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="RightsOfWay_Yes" type="radio" value="1" name="RightsOfWay"><LABEL class="msgLabel" for="RightsOfWay_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="RightsOfWay_No" type="radio" value="0" name="RightsOfWay"><LABEL class="msgLabel" for="RightsOfWay_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 140px">Shared Drive / Access? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="SharedDrive_Yes" type="radio" value="1" name="SharedDrive"><LABEL class="msgLabel" for="SharedDrive_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="SharedDrive_No" type="radio" value="0" name="SharedDrive"><LABEL class="msgLabel" for="SharedDrive_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 160px">Garage on Separate Site? 
<SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="GarageSeparate_Yes" type="radio" value="1" name="GarageSeparate"><LABEL class="msgLabel" for="GarageSeparate_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="GarageSeparate_No" type="radio" value="0" name="GarageSeparate"><LABEL class="msgLabel" for="GarageSeparate_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 180px">Flying Freehold? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="FlyingFreehold_Yes" type="radio" value="1" name="FlyingFreehold"><LABEL class="msgLabel" for="FlyingFreehold_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="FlyingFreehold_No" type="radio" value="0" name="FlyingFreehold"><LABEL class="msgLabel" for="FlyingFreehold_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 200px">Flying Freehold Percentage 
<SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT class="msgTxt" id="txtFFPercentage" style="WIDTH: 140px" maxLength="3" name="txtFFPercentage">
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 220px">Shared Services? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="SharedServices_Yes" type="radio" value="1" name="SharedServices"><LABEL class="msgLabel" for="SharedServices_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="SharedServices_No" type="radio" value="0" name="SharedServices"><LABEL class="msgLabel" for="SharedServices_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 240px">Ill Defined Boundaries? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="IllDefinedBoundaries_Yes" type="radio" value="1" name="IllDefinedBoundaries"><LABEL class="msgLabel" for="IllDefinedBoundaries_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="IllDefinedBoundaries_No" type="radio" value="0" name="IllDefinedBoundaries"><LABEL class="msgLabel" for="IllDefinedBoundaries_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 260px">Other Legal 
Issues? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="OtherLegalIssues_Yes" type="radio" value="1" name="OtherLegalIssues"><LABEL class="msgLabel" for="OtherLegalIssues_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="OtherLegalIssues_No" type="radio" value="0" name="OtherLegalIssues"><LABEL class="msgLabel" for="OtherLegalIssues_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 280px">Surveyor to Inspect Title 
Plans? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="SurveyorInspection_Yes" type="radio" value="1" name="SurveyorInspection"><LABEL class="msgLabel" for="SurveyorInspection_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="SurveyorInspection_No" type="radio" value="0" name="SurveyorInspection"><LABEL class="msgLabel" for="SurveyorInspection_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 300px">Occupancy 
Restrictions? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="OccupancyRestrictions_Yes" type="radio" value="1" name="OccupancyRestrictions"><LABEL class="msgLabel" for="OccupancyRestrictions_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="OccupancyRestrictions_No" type="radio" value="0" name="OccupancyRestrictions"><LABEL class="msgLabel" for="OccupancyRestrictions_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 320px">Plot More Than 
10 Acres? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="MoreThan10Acres_Yes" type="radio" value="1" name="MoreThan10Acres"><LABEL class="msgLabel" for="MoreThan10Acres_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="MoreThan10Acres_No" type="radio" value="0" name="MoreThan10Acres"><LABEL class="msgLabel" for="MoreThan10Acres_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 340px">Has Extensions/Alterations? 
<SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="Extensions_Yes" type="radio" value="1" name="Extensions"><LABEL class="msgLabel" for="Extensions_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="Extensions_No" type="radio" value="0" name="Extensions"><LABEL class="msgLabel" for="Extensions_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 360px">Land not Residential? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="LandNotResidential_Yes" type="radio" value="1" name="LandNotResidential"><LABEL class="msgLabel" for="LandNotResidential_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="LandNotResidential_No" type="radio" value="0" name="LandNotResidential"><LABEL class="msgLabel" for="LandNotResidential_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 380px">Unadopted Road 
Issues? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="UnadoptedRoad_Yes" type="radio" value="1" name="UnadoptedRoad"><LABEL class="msgLabel" for="UnadoptedRoad_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="UnadoptedRoad_No" type="radio" value="0" name="UnadoptedRoad"><LABEL class="msgLabel" for="UnadoptedRoad_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 400px">Tenanted Property? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="TenantedProperty_Yes" type="radio" value="1" name="TenantedProperty"><LABEL class="msgLabel" for="TenantedProperty_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="TenantedProperty_No" type="radio" value="0" name="TenantedProperty"><LABEL class="msgLabel" for="TenantedProperty_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 420px">Development Proposals? <SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="DevelopmentProposals_Yes" type="radio" value="1" name="DevelopmentProposals"><LABEL class="msgLabel" for="DevelopmentProposals_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="DevelopmentProposals_No" type="radio" value="0" name="DevelopmentProposals"><LABEL class="msgLabel" for="DevelopmentProposals_No">No</LABEL>
					</SPAN></SPAN><SPAN class="msgLabel" style="LEFT: 10px; POSITION: absolute; TOP: 440px">Mining Area? 
<SPAN style="LEFT: 250px; POSITION: absolute; TOP: -3px"><INPUT id="MiningArea_Yes" type="radio" value="0" name="MiningArea"><LABEL class="msgLabel" for="MiningArea_Yes">Yes</LABEL>
					</SPAN><SPAN style="LEFT: 300px; POSITION: absolute; TOP: -3px">
						<INPUT id="MiningArea_No" type="radio" value="0" name="MiningArea"><LABEL class="msgLabel" for="MiningArea_No">No</LABEL>
					</SPAN></SPAN></DIV>
		</FORM>
		<%/* Main Buttons */ %>
		<DIV id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 540px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --></DIV>
		<% /* Span to keep tabbing within this screen */ %>
		<SPAN id="spnToFirstField" tabIndex="0"></SPAN><!-- #include FILE="fw030.asp" -->
		<% /* File containing field attributes (remove if not required) */ %>
		<!-- #include FILE="attribs/AP204attribs.asp" -->
		<% /* Specify Code Here */ %>
		<SCRIPT language="JScript">
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
	FW030SetTitles("Comments for Solicitor","AP204",scScreenFunctions);

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
		frmScreen.txtComments.value = m_XML.GetAttribute("SOLICITORNOTES"); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.Extensions_Yes, frmScreen.Extensions_No, m_XML.GetAttribute("EXTENSIONSORALTERATIONS")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.LandNotResidential_Yes, frmScreen.LandNotResidential_No, m_XML.GetAttribute("NONRESIDENTIALLANDIND")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.UnadoptedRoad_Yes, frmScreen.UnadoptedRoad_No, m_XML.GetAttribute("UNADOPTEDSHAREDACCESSISSUES")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.TenantedProperty_Yes, frmScreen.TenantedProperty_No, m_XML.GetAttribute("TENANTEDPROPERTYIND")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.DevelopmentProposals_Yes, frmScreen.DevelopmentProposals_No, m_XML.GetAttribute("DEVELOPMENTPROPOSALS")); //ValnRepPropertyRisks
		SetOptionValue(frmScreen.MiningArea_Yes, frmScreen.MiningArea_No, m_XML.GetAttribute("MININGAREA")); //ValnRepPropertyRisks
	
		//EP2_2 Add new fields.
		SetOptionValue(frmScreen.RightsOfWay_Yes, frmScreen.RightsOfWay_No, m_XML.GetAttribute("RIGHTSOFWAY")); 
		SetOptionValue(frmScreen.SharedDrive_Yes, frmScreen.SharedDrive_No, m_XML.GetAttribute("SHAREDDRIVEORACCESS")); 
		SetOptionValue(frmScreen.GarageSeparate_Yes, frmScreen.GarageSeparate_No, m_XML.GetAttribute("GARAGEONSEPARATESITE")); 
		SetOptionValue(frmScreen.FlyingFreehold_Yes, frmScreen.FlyingFreehold_No, m_XML.GetAttribute("FLYINGFREEHOLD")); 
		if (frmScreen.FlyingFreehold_Yes.checked == true)
			frmScreen.txtFFPercentage.disabled = false;	
		SetOptionValue(frmScreen.SharedServices_Yes, frmScreen.SharedServices_No, m_XML.GetAttribute("SHAREDSERVICES")); 
		SetOptionValue(frmScreen.IllDefinedBoundaries_Yes, frmScreen.IllDefinedBoundaries_No, m_XML.GetAttribute("BOUNDARYILLDEFINED")); 
		SetOptionValue(frmScreen.OtherLegalIssues_Yes, frmScreen.OtherLegalIssues_No, m_XML.GetAttribute("OTHERLEGALISSUES")); 
		SetOptionValue(frmScreen.SurveyorInspection_Yes, frmScreen.SurveyorInspection_No, m_XML.GetAttribute("SURVEYORTOINSPECTTITLE")); 
		SetOptionValue(frmScreen.OccupancyRestrictions_Yes, frmScreen.OccupancyRestrictions_No, m_XML.GetAttribute("OCCUPANCYRESTRICTIONS")); 
		SetOptionValue(frmScreen.MoreThan10Acres_Yes, frmScreen.MoreThan10Acres_No, m_XML.GetAttribute("LARGEPLOT")); 
		frmScreen.txtFFPercentage.value = m_XML.GetAttribute("FLYINGFREEHOLDPERCENT"); 
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
	//Update the valuation report for the new fields
	var bSuccess = true;
	if(!frmScreen.onsubmit())
		bSuccess = false;
		
	else if (!ValidateScreen())
		{
			alert("Flying Freehold percentage must be between 0 and 100.");
			frmScreen.txtFFPercentage.focus();
			bSuccess = false;
		}
	
	else
	{
		if(IsChanged())
		{
			var ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var reqTag = ValuationXML.CreateRequestTag(window , "UpdateValuationReport");
			ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");

			ValuationXML.CreateActiveTag("VALUATION");
			ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
			ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
			ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInsSeqNo);
			
			ValuationXML.SetAttribute("SOLICITORNOTES", frmScreen.txtComments.value);
			ValuationXML.SetAttribute("EXTENSIONSORALTERATIONS", GetOptionValue(frmScreen.Extensions_Yes));
			ValuationXML.SetAttribute("NONRESIDENTIALLANDIND", GetOptionValue(frmScreen.LandNotResidential_Yes));
			ValuationXML.SetAttribute("UNADOPTEDSHAREDACCESSISSUES", GetOptionValue(frmScreen.UnadoptedRoad_Yes));
			ValuationXML.SetAttribute("TENANTEDPROPERTYIND", GetOptionValue(frmScreen.TenantedProperty_Yes));
			ValuationXML.SetAttribute("DEVELOPMENTPROPOSALS", GetOptionValue(frmScreen.DevelopmentProposals_Yes ));
			ValuationXML.SetAttribute("MININGAREA", GetOptionValue(frmScreen.MiningArea_Yes));

			//EP2_2 Add new fields.
			ValuationXML.SetAttribute("RIGHTSOFWAY", GetOptionValue(frmScreen.RightsOfWay_Yes));
			ValuationXML.SetAttribute("SHAREDDRIVEORACCESS", GetOptionValue(frmScreen.SharedDrive_Yes));
			ValuationXML.SetAttribute("GARAGEONSEPARATESITE", GetOptionValue(frmScreen.GarageSeparate_Yes));
			ValuationXML.SetAttribute("FLYINGFREEHOLD", GetOptionValue(frmScreen.FlyingFreehold_Yes));
			ValuationXML.SetAttribute("SHAREDSERVICES", GetOptionValue(frmScreen.SharedServices_Yes));
			ValuationXML.SetAttribute("BOUNDARYILLDEFINED", GetOptionValue(frmScreen.IllDefinedBoundaries_Yes));
			ValuationXML.SetAttribute("OTHERLEGALISSUES", GetOptionValue(frmScreen.OtherLegalIssues_Yes));
			ValuationXML.SetAttribute("SURVEYORTOINSPECTTITLE", GetOptionValue(frmScreen.SurveyorInspection_Yes));
			ValuationXML.SetAttribute("OCCUPANCYRESTRICTIONS", GetOptionValue(frmScreen.OccupancyRestrictions_Yes));
			ValuationXML.SetAttribute("LARGEPLOT", GetOptionValue(frmScreen.MoreThan10Acres_Yes));
			ValuationXML.SetAttribute("FLYINGFREEHOLDPERCENT", frmScreen.txtFFPercentage.value);

			ValuationXML.RunASP(document,"omAppProc.asp");
			bSuccess = ValuationXML.IsResponseOK()
		}
	}
	if (bSuccess)
		frmToAP200.submit();
}
function btnCancel.onclick()
{
	frmToAP200.submit();
}

<% /* EP2_2 - New Methods = AShaw - 22Nov06 */ %>
function frmScreen.FlyingFreehold_Yes.onclick()
{
	frmScreen.txtFFPercentage.disabled = false;
	frmScreen.txtFFPercentage.setAttribute("required", "true");
}

function frmScreen.FlyingFreehold_No.onclick()
{
	frmScreen.txtFFPercentage.value = "";	
	frmScreen.txtFFPercentage.disabled = true;	
	frmScreen.txtFFPercentage.removeAttribute("required");
}

function ValidateScreen()
{
	<% /* Check that the Flying freehold percentage is between 0 and 100 if enabled. */ %>
	if(frmScreen.FlyingFreehold_Yes.checked == true && (frmScreen.txtFFPercentage.value == "" || parseInt(frmScreen.txtFFPercentage.value) > 100 ))
		return(false);
	else
		return(true);
	
}
<% /* EP2_2 - END New Methods = AShaw - 22Nov06 */ %>

-->
		</SCRIPT>
		</STRONG>
	</BODY>
</HTML>
