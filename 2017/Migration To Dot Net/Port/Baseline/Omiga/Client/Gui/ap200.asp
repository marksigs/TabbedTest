<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP200.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		01/02/01	SYS1839 Created
DJP		01/02/01	SYS1839 Don't move to next screen in validation fails.
DJP		16/03/01	SYS2040 Added Directory Search
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
DRC     21/11/01    SYS3048 Get Instruction Sequence No if not present in the Task Context from
                    the omiga menu context
BG		03/11/2001  SYS3048	moved "var m_ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();"
					from declarations at top of JScript to GetValuationReport.			
JLD		17/12/01	SY3392 check retention amount against amount requested				
DRC     20/12/01	SYS3523 Trap to stop finding the Repayment period where there's no accepted quote
MEVA	19/04/02	SYS3336 Add hidden Fee Authorisation field and exit processing
AT		25/04/02	SYS4359 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
DPF		23/08/2002	Re-structuring of screen, including new fields & functions, removal of fields &
					functions and amendment of existing functions.
DPF		12/11/2002	Have made 'Comment box#' permanently enabled.  Rules for check boxes will still apply.
DPF		13/11/2002	Have added Yes / No question for Re-inspection.
MDC		18/11/2002	CC014
GHun	16/01/2003	BM0243 No data displayed when InstructionSequenceNo is blank
HMA     16/09/2003  BM0063 Amend HTML text for radio buttons
KRW     25/09/2003    BM0063 Changed from ! = 1 for when case is (Read Only) 
KRW     26/09/2003  BM0063 Changed back to != -1 (not saving values !!) 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
INGDUK History:

Prog	Date		Description
JD		19/09/2005	MAR40 Updates for Esurv/Hometrack
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<form id="frmToAP205" method="post" action="AP205.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP210" method="post" action="AP210.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP201" method="post" action="AP201.asp" STYLE="DISPLAY: none"></form>
<% /* JD MAR40 add new screens */ %>
<form id="frmToAP202" method="post" action="AP202.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP203" method="post" action="AP203.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP204" method="post" action="AP204.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here
	 Amended 28/08/2002 - for APWP3 by DPF
 -->
<form id="frmScreen" mark validate ="onchange">
	<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 175px" class="msgGroup">
		  <span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel" >	
				<strong>Valuation</strong>:
			</span>
			<span style="LEFT: 4px; POSITION: absolute; TOP: 25px" class="msgLabel">	
				Date Received
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtDateReceived" maxlength=10 style="WIDTH: 100px" class ="msgTxt">
				</span>
			</span>
			<span style="LEFT: 230px; POSITION: absolute; TOP: 20px" class="msgLabel">						
				<input id="btnValuerDetails" value="Valuer Details" type="button" style="WIDTH: 100px; HEIGHT: 22px" class="msgButton">
			</span>
			<span style="LEFT: 4px; POSITION: absolute; TOP: 50px" class="msgLabel">	
				Present Valuation
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtPresentValuation"  maxlength=10 style="WIDTH: 100px"  class="msgTxt">
				</span>
			</span>
			<span style="LEFT: 4px; POSITION: absolute; TOP: 75px" class="msgLabel">	
				Reinstatement Value
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtReinstatementValue" maxlength=10 style="WIDTH: 100px"  class="msgTxt">
				</span>
			</span>
			<span style="LEFT: 4px; POSITION: absolute; TOP: 100px" class="msgLabel">	
				Post Works Valuation
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtPostWorksValuation"  maxlength=10 style="WIDTH: 100px"  class="msgTxt">
				</span>
			</span>
			<div style="LEFT: 350px; POSITION: absolute; TOP: 25px" class="msgLabel">	
				Retention Amount
				<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
					<input id="txtRetentionAmount"  maxlength=10 style="WIDTH: 100px"  class="msgTxt">
				</span>
			</div>
			<div style="TOP:125px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
				Overall Condition
				<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
					<select id="cboOverallCondition" style="WIDTH:100px" class="msgCombo"></select>
				</span>
			</div>
			<div style="TOP:150px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
				Saleability
				<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
					<select id="cboSaleability" style="WIDTH:100px" class="msgCombo"></select>
				</span>
			</div>			
			<div style="LEFT: 350px; POSITION: absolute; TOP: 50px" class="msgLabel">	
				Multiple Occupancy?
				<span style="LEFT: 116px; POSITION: absolute; TOP: -3px">
					<input id="MultipleOccupancy_Yes" name="MultipleOccupancy" type="radio" value="1"><label for="MultipleOccupancy_Yes" class="msgLabel">Yes</label>
				</span>
				<span style="LEFT: 184px; POSITION: absolute; TOP: -3px">
					<input id="MultipleOccupancy_No" name="MultipleOccupancy" type="radio" value="0"><label for="MultipleOccupancy_No" class="msgLabel">No</label>
				</span>
			</div>
			<div style="LEFT: 350px; POSITION: absolute; TOP: 75px" class ="msgLabel">	
				Refer to Insurance?
				<span style="LEFT: 116px; POSITION: absolute; TOP: -3px">
					<input id="ReferToInsurance_Yes" name="ReferToInsurance" type="radio" value="1"><label for="ReferToInsurance_Yes" class="msgLabel">Yes</label>
				</span>
				<span style="LEFT: 184px; POSITION: absolute; TOP: -3px">
					<input id="ReferToInsurance_No" name="ReferToInsurance" type="radio" value="0"><label for="ReferToInsurance_No" class="msgLabel">No</label>
				</span>
			</div><!-- field added for re-inspection as part of AQR BMIDS00902 -->
			<div style="LEFT: 350px; POSITION: absolute; TOP: 100px" class ="msgLabel">	
				Re-inspection Required?
				<span style="LEFT: 116px; POSITION: absolute; TOP: -3px">
					<input id="ReInspection_Yes" name="ReInspection" type="radio" value="1"><label for="ReInspection_Yes" class="msgLabel">Yes</label>
				</span>
				<span style="LEFT: 184px; POSITION: absolute; TOP: -3px">
					<input id="ReInspection_No" name="ReInspection" type="radio" value="0"><label for="ReInspection_No" class="msgLabel">No</label>
				</span>
			</div>
			<div style="LEFT: 350px; POSITION: absolute; TOP: 125px" class ="msgLabel">	
				Signature Returned?
				<span style="LEFT: 116px; POSITION: absolute; TOP: -3px">
					<input id="SignatureReturned_Yes" name="SignatureReturned" type="radio" value="1"><label for="SignatureReturned_Yes" class="msgLabel">Yes</label>
				</span>
				<span style="LEFT: 184px; POSITION: absolute; TOP: -3px">
					<input id="SignatureReturned_No" name="SignatureReturned" type="radio" value="0"><label for="SignatureReturned_No" class="msgLabel">No</label>
				</span>
			</div>
		</div><!-- Second part of screen		 Amended 28/08/2002 - for APWP3 by DPF -->
		
		<!-- JD MAR40 added 3 buttons in new div-->
		<div style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 251px; HEIGHT: 50px" class ="msgGroup" >
			<span style="LEFT: 100px; POSITION: absolute; TOP: 10px">
				<input id="btnGeneralObservations" value="General Observations" type="button" style="HEIGHT: 20px;WIDTH:120px" class="msgButton">
			</span>
			<span style="LEFT: 230px; POSITION: absolute; TOP: 10px">
				<input id="btnPropertyEssentialRepairs" value="Property Essential Repairs" type="button" style="HEIGHT: 20px;WIDTH:135px" class="msgButton">
			</span>
			<span style="LEFT: 380px; POSITION: absolute; TOP: 10px">
				<input id="btnCommentsForSolicitor" value="Comments For Solicitor" type="button" style="HEIGHT: 20px;WIDTH:135px" class="msgButton">
			</span>
		</div>
		
<!-- JD MAR40 removed Hazards div
	<div style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 201px; HEIGHT: 121px" class ="msgGroup" >	
	<span id="spnHazards" style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<span  style="LEFT: 12px; WIDTH: 63px; POSITION: absolute; TOP: -3px; HEIGHT: 18px">
			<label for="Trees_Yes" class="msgLabel">Trees</label>
			<input id="Trees_Yes" name="RemarksRadioGroup" type="checkbox" value="0">
		</span>
		<span  style="LEFT: 90px; WIDTH: 60px; POSITION: absolute; TOP: -3px; HEIGHT: 18px" id=SPAN2>
			<label for="DryRot_Yes" class="msgLabel">Dry Rot</label>
			<input id="DryRot_yes" name="DryRot_yes" type="checkbox" value="0" style="LEFT: 38px; TOP: 1px">
		</span>
		<span  style="LEFT: 177px; WIDTH: 122px; POSITION: absolute; TOP: -3px; HEIGHT: 18px" id=SPAN1>
			<label for="OngoingMove_Yes" class="msgLabel">Ongoing Movement</label>
			<input id="OngoingMove_Yes" name="OngoingMove_yes" type="checkbox" value="0">
		</span>
		<span  style="LEFT: 327px; WIDTH: 53px; POSITION: absolute; TOP: -3px; HEIGHT: 21px">
			<label for="Other_Yes" class="msgLabel">Other</label>
			<input id="Other_Yes" name="Other_Yes" type="checkbox" value="0">
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 25px"><TEXTAREA class=msgTxt id=txtNotes style="WIDTH: 550px" rows=5></TEXTAREA>
		</span>
	</span>
	</div>
-->
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP:  330px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP200attribs.asp" -->

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
var m_sUnitId = "";
var m_sUnitName = "";
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
%>	var sButtonList = new Array("Submit","Cancel","Next");

	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Valuation Report - Valuation","AP200",scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	RetrieveContextData();
	
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
	PopulateScreen();
	
 	scScreenFunctions.SetFocusToFirstField(frmScreen);

	if(m_sReadOnly != "1") // Changed back to != -1 (not saving values !!) KRW 26/09/03
	{
		var bReadOnly = false;
		bReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP200");	

		if(bReadOnly)
		{
			m_sReadOnly = "1";		
		}
	}
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}
function PopulateCombos()
{
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var sGroupList = new Array("ValuationOverAllCondition","ValuationSaleability");

	if(XML.GetComboLists(document,sGroupList))
	{
		XML.PopulateCombo(document,frmScreen.cboOverallCondition,"ValuationOverAllCondition",true);
		XML.PopulateCombo(document,frmScreen.cboSaleability,"ValuationSaleability",true);
	}
}
function PopulateScreen()
{
	var bSuccess;

	PopulateCombos(); //JD MAR40 added
	
	bSuccess = GetValuationReport();

	if(bSuccess)
	{
		// Populate the screen fields
		frmScreen.txtDateReceived.value = m_ValuationXML.GetAttribute("DATERECEIVED");
		frmScreen.txtPresentValuation.value = m_ValuationXML.GetAttribute("PRESENTVALUATION");
		frmScreen.txtReinstatementValue.value = m_ValuationXML.GetAttribute("REINSTATEMENTVALUE");
		frmScreen.txtPostWorksValuation.value = m_ValuationXML.GetAttribute("POSTWORKSVALUATION");
		frmScreen.txtRetentionAmount.value = m_ValuationXML.GetAttribute("RETENTIONWORKS");
		
		//New radio values - APWP3 - DPF 28/08/2002
		SetOptionValue(frmScreen.MultipleOccupancy_Yes, frmScreen.MultipleOccupancy_No, m_ValuationXML.GetAttribute("MULTIPLEOCCUPANCY"));
		SetOptionValue(frmScreen.ReferToInsurance_Yes,frmScreen.ReferToInsurance_No,m_ValuationXML.GetAttribute("REFERTOINSURANCE"));
		SetOptionValue(frmScreen.ReInspection_Yes,frmScreen.ReInspection_No,m_ValuationXML.GetAttribute("FUTUREINSPECTION"));
		
		//JD MARS40 add Overall condition, Saleability, signature returned
		SetOptionValue(frmScreen.SignatureReturned_Yes,frmScreen.SignatureReturned_No,m_ValuationXML.GetAttribute("SIGNATURERETURNED"));
		frmScreen.cboOverallCondition.value = m_ValuationXML.GetAttribute("OVERALLCONDITION");
		frmScreen.cboSaleability.value = m_ValuationXML.GetAttribute("SALEABILITY");
		
		/*JD MAR40 removed hazards
		//Hazards - APWP3 - DPF 28/08/2002
		frmScreen.txtNotes.value = m_ValuationXML.GetAttribute("GENERALREMARKNOTES");
		SetCheckBoxValue(frmScreen.Trees_Yes, m_ValuationXML.GetAttribute("TREEREPORTREQ"));
		SetCheckBoxValue(frmScreen.DryRot_yes, m_ValuationXML.GetAttribute("TIMBERDAMPREPORTREQ"));
		SetCheckBoxValue(frmScreen.OngoingMove_Yes, m_ValuationXML.GetAttribute("LOCALSUBSIDENCE"));
		SetCheckBoxValue(frmScreen.Other_Yes, m_ValuationXML.GetAttribute("OTHERHAZARDS"));
		JD END   */
	}	
	else
	{
		//populate new fields - APWP3 - DPF 28/08/2002
		frmScreen.MultipleOccupancy_No.checked ="1";	
		frmScreen.ReferToInsurance_No.checked ="1";
		frmScreen.ReInspection_No.checked ="1";
		/*JD MAR40 removed hazards	
		frmScreen.Trees_Yes.checked = false;
		frmScreen.DryRot_yes.checked = false;
		frmScreen.OngoingMove_Yes.checked = false;
		frmScreen.Other_Yes.checked = false;
		JD END */
		m_sReportNotFound = "1";
	}
	
	/* JD MAR40 removed	
	if(m_sReadOnly != "1")
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtNotes")			
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtNotes")			
	}
	JD END */

	return(bSuccess);
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
		m_sInsSeqNo = scScreenFunctions.GetContextParameter(window,"idInstructionSequenceNo",null)
			
	<% /* BM0243 Only set InstructionSequenceNo if it is not blank */ %>
	if (m_sInsSeqNo.length > 0)
		m_ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO",m_sInsSeqNo);
	<% /* BM0243 End */ %>
	
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
	m_sUnitId = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	m_sUnitName = scScreenFunctions.GetContextParameter(window,"idUnitName",null);
	m_sAppFactFindNo = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sCallingScreen = scScreenFunctions.GetContextParameter(window,"idREturnScreenId",null);

	if(m_sReadOnly != "1")
	{
		var sRet;
		sRet = scScreenFunctions.GetContextParameter(window,"idMetaAction", null);	
		if(sRet == "1")
		{
			m_sReadOnly = "1";
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
		}
	}
}

function CommitChanges()
{
	var bSuccess = true;
	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			//if (IsChanged())
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
	var sNotes;
	
	/*JD MAR40 removed hazards
	// Validate hazards - APWP3 - DPF 28/08/2002
	sNotes = frmScreen.txtNotes.value
	if ((frmScreen.Trees_Yes.checked == true || frmScreen.DryRot_yes.checked == true || frmScreen.OngoingMove_Yes.checked == true || frmScreen.Other_Yes.checked == true))
	{
		alert("Please ensure you have entered notes for hazard(s) recorded");
		frmScreen.txtNotes.focus();
		bSuccess = true;
	}
	JD END */

	<% /* BMIDS00938 MDC 18/11/2002 */ %>
	var dblPostWorksValuation = parseFloat(frmScreen.txtPostWorksValuation.value);
	if (isNaN(dblPostWorksValuation))
		dblPostWorksValuation = 0

	var dblRetentionAmount = parseFloat(frmScreen.txtRetentionAmount.value);
	if (isNaN(dblRetentionAmount))
		dblRetentionAmount = 0;
		
	if (dblPostWorksValuation > 0 && dblRetentionAmount == 0)
	{
		alert("If a Post Works Valuation is entered, you must specify the Retention Amount.");
		frmScreen.txtRetentionAmount.focus();
		return false;
	}
	
	if (dblPostWorksValuation == 0 && dblRetentionAmount > 0)
	{
		alert("If a Retention Amount is entered, you must specify the Post Works Valuation.");
		frmScreen.txtPostWorksValuation.focus();
		return false;
	}
	<% /* BMIDS00938 MDC 18/11/2002 - End */ %>
	
	return(bSuccess);
}

//2 new functions to Set / Retrieve values on the screen for Hazard Check Boxes - APWP3 - DPF 28/08/2002
function GetCheckBoxValue( objOption )
{
	var sVal = "0";
	
	if(objOption.checked == true)
	{
		sVal = "1";
	}
	return(sVal)
}

function SetCheckBoxValue( objOption, sVal )
{
	if( sVal == "1" )
	{
		objOption.checked = true;
	}
	else
	{
		objOption.checked = false;
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

function SaveScreenValues( XML )
{
	var bSuccess;
	var sVal;
	
	bSuccess = false;
	if(XML != null)
	{
		XML.SetAttribute("DATERECEIVED", frmScreen.txtDateReceived.value);
		XML.SetAttribute("PRESENTVALUATION", frmScreen.txtPresentValuation.value);
		XML.SetAttribute("REINSTATEMENTVALUE", frmScreen.txtReinstatementValue.value);
		XML.SetAttribute("POSTWORKSVALUATION", frmScreen.txtPostWorksValuation.value);
		XML.SetAttribute("RETENTIONWORKS",frmScreen.txtRetentionAmount.value);
		
		//New radio values - APWP3 - DPF 28/08/2002
		XML.SetAttribute("MULTIPLEOCCUPANCY", GetOptionValue(frmScreen.MultipleOccupancy_Yes));
		XML.SetAttribute("REFERTOINSURANCE", GetOptionValue(frmScreen.ReferToInsurance_Yes));
		XML.SetAttribute("FUTUREINSPECTION", GetOptionValue(frmScreen.ReInspection_Yes));
		
		//JD MAR40 add new fields
		XML.SetAttribute("SIGNATURERETURNED", GetOptionValue(frmScreen.SignatureReturned_Yes));
		XML.SetAttribute("OVERALLCONDITION", frmScreen.cboOverallCondition.value);
		XML.SetAttribute("SALEABILITY", frmScreen.cboSaleability.value);
		
		/*JD MAR40 removed hazards		
		XML.SetAttribute("NOTES",frmScreen.txtNotes.value);
		
		
		//Hazards - APWP3 - DPF 28/08/2002
		XML.SetAttribute("GENERALREMARKNOTES",frmScreen.txtNotes.value);
		XML.SetAttribute("TREEREPORTREQ", GetCheckBoxValue(frmScreen.Trees_Yes));
		XML.SetAttribute("TIMBERDAMPREPORTREQ", GetCheckBoxValue(frmScreen.DryRot_yes));
		XML.SetAttribute("LOCALSUBSIDENCE", GetCheckBoxValue(frmScreen.OngoingMove_Yes));
		XML.SetAttribute("OTHERHAZARDS", GetCheckBoxValue(frmScreen.Other_Yes));
		JD END */
		
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

function btnCancel.onclick()
{
	frmToAP205.submit();
}

function btnNext.onclick()
{
	if(CommitChanges())
	{
		frmToAP210.submit();
	}
}
function frmScreen.btnGeneralObservations.onclick()
{
	//JD MAR40 Added button. Route to AP202 passing valuationXML
	//Save any changes first
	if (IsChanged())
		CommitChanges();
	scScreenFunctions.SetContextParameter(window,"idXML", m_ValuationXML.XMLDocument.xml);
	frmToAP202.submit();
}
function frmScreen.btnPropertyEssentialRepairs.onclick()
{
	//JD MAR40 Added button. Route to AP203 passing valuationXML
	//Save any changes first
	if (IsChanged())
		CommitChanges();
	scScreenFunctions.SetContextParameter(window,"idXML", m_ValuationXML.XMLDocument.xml);
	frmToAP203.submit();
}
function frmScreen.btnCommentsForSolicitor.onclick()
{
	//JD MAR40 Added button. Route to AP204 passing valuationXML
	//Save any changes first
	if (IsChanged())
		CommitChanges();
	scScreenFunctions.SetContextParameter(window,"idXML", m_ValuationXML.XMLDocument.xml);
	frmToAP204.submit();
}
function frmScreen.btnValuerDetails.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTmp;
	
	ArrayArguments[0]= XML.CreateRequestAttributeArray(window);	
	ArrayArguments[1]= m_sAppNo;
	ArrayArguments[2]= m_sAppFactFindNo;	
	ArrayArguments[3]= m_sInsSeqNo;	
	ArrayArguments[4]= m_sTaskXML;	
	ArrayArguments[5]= m_sReadOnly;	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "AP201.asp", ArrayArguments, 630, 528);
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
</STRONG>
</body>
</html>


