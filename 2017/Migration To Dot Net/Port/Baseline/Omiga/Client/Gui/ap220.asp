<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP220.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		01/02/01	SYS1839 Created
DJP		03/03/01	SYS1839 Added Main Services processing
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
DRC     21/11/01    SYS3048 Get Instruction Sequence No if not present in the Task Context from
                    the omiga menu context	
BG		03/11/2001  SYS3048	moved "var m_ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();"
					from declarations at top of JScript to GetValuationReport.
AT		30/04/02	SYS4359 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
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

EPSOM-Specific History

Description: 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SAB		13/07/2006	EP976 Improved radio button presentation when disabled.  SetServices()
						now checks to see if the radio buttons are disabled before updating
						fields within the screen.
AW		05/10/2006	EP1196 - Amended radio reference
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css"><title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<script src="validation.js" language="JScript"></script><!-- Screen Functions Object - see comments within scScreenFunctions.htm for details -->
<!-- Specify Forms Here -->

<% /* Specify Forms Here */ %>
<form id="frmToAP205" method="post" action="AP205.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP230" method="post" action="AP230.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange">
	<!--
	<div id="divBackground" style="TOP: 40px; LEFT: 10px; HEIGHT: 350px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
 Titles -->
	
	
	<!-- Screen section one -->
	<div style="LEFT: 10px; POSITION: absolute; TOP: 60px; HEIGHT: 240px; WIDTH: 604px;" class="msgGroup" >		
		<span style="POSITION: absolute; TOP: 4px;LEFT: 4px;" class="msgLabel" >		
			<strong>Services
		</span>
		<div style="POSITION: absolute; TOP: 25px;LEFT: 4px;" class="msgLabel" >	
			Number of Bedrooms
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtBedrooms"  style="WIDTH: 100px ">
			</span>
		</div>
		<span id="spnMainServices" style="TOP:50px;LEFT: 4px; POSITION:ABSOLUTE" class="msgLabel">
			All Main Services?
			<span style="TOP:-3px; LEFT:116px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optMainServicesYes" name="MainServicesRadioGroup" type="radio" value="1">
				<label for="optMainServicesYes" class="msgLabel">Yes</label>
			</span>
			<span style="TOP:-3px; LEFT:184px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optMainServicesNo" name="MainServicesRadioGroup" type="radio" value="0">
				<label for="optMainServicesNo" class="msgLabel">No</label>
			</span>
		</span>
		<span style="TOP:40px;LEFT: 324px; POSITION:ABSOLUTE" class="msgLabel">
			Hot Water/<BR>Central Heating?
			<span style="TOP:6px; LEFT:116px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optHotWaterYes" name="HotWaterRadioGroup" type="radio" value="1">
				<label for="optHotWaterYes" class="msgLabel">Yes</label>
			</span>
			<span style="TOP:6px; LEFT:184px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optHotWaterNo" name="HotWaterRadioGroup" type="radio" value="0">
				<label for="optHotWaterNo" class="msgLabel">No</label>
			</span>
		</span>
		<span style="TOP:75px;LEFT: 4px; POSITION:ABSOLUTE" class="msgLabel">
		Gas
			<span style="TOP:-3px; LEFT:125px; POSITION:ABSOLUTE">
				<select id="cboGas" style="WIDTH:150px" class="msgCombo"></select>
			</span>
		</span>
		<span style="TOP:75px;LEFT: 324px; POSITION:ABSOLUTE" class="msgLabel">
		Electricity
			<span style="TOP:-3px; LEFT:125px; POSITION:ABSOLUTE">
				<select id="cboElectricity" style="WIDTH:150px" class="msgCombo"></select>
			</span>
		</span>		
		<span style="TOP:100px;LEFT: 4px;POSITION:ABSOLUTE" class="msgLabel">
		Water
			<span style="TOP:-3px; LEFT:125px; POSITION:ABSOLUTE">
				<select id="cboWater" style="WIDTH:150px" class="msgCombo"></select>
			</span>
		</span>
		<span style="TOP:100px;LEFT: 324px;POSITION:ABSOLUTE" class="msgLabel">
		Drainage
			<span style="TOP:-3px; LEFT:125px; POSITION:ABSOLUTE">
				<select id="cboDrainage" style="WIDTH:150px" class="msgCombo"></select>
			</span>
		</span>
		<span style="TOP:125px;LEFT: 4px; POSITION:ABSOLUTE" class="msgLabel">
		Demand in Area
			<span style="TOP:-3px; LEFT:125px; POSITION:ABSOLUTE">
				<select id="cboDemandInArea" style="WIDTH:150px" class="msgCombo"></select>
			</span>
		</span>
		<span id="spnOtherFactors" style="TOP:150px;LEFT: 4px; POSITION:ABSOLUTE" class="msgLabel">
		Other Factors?
			<span style="TOP:-3px; LEFT:116px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optOtherFactorsYes" name="OtherFactorsRadioGroup" type="radio" value="1">
				<label for="optOtherFactorsYes" class="msgLabel">Yes</label>
			</span>
			<span style="TOP:-3px; LEFT:184px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optOtherFactorsNo" name="OtherFactorsRadioGroup" type="radio" value="0">
				<label for="optOtherFactorsNo" class="msgLabel">No</label>
			</span>
		</span>
		<span style="TOP:175px; LEFT:4px; POSITION:ABSOLUTE">
			<textarea id="txtOtherFactors" rows="3" style="WIDTH:225px" class="msgTxt"></textarea>
		</span>
		
		<span style="TOP:135px;LEFT: 324px; POSITION:ABSOLUTE" class="msgLabel">
			Any items of disrepair<BR> which may threaten <BR>future saleability?
			<span style="TOP:8px; LEFT:120px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optDisrepairYes" name="DisrepairRadioGroup" type="radio" value="1">
				<label for="optDisrepairYes" class="msgLabel">Yes</label>
			</span>
			<span style="TOP:8px;  LEFT:188px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optDisrepairNo" name="DisrepairRadioGroup" type="radio" value="0">
				<label for="optDisrepairNo" class="msgLabel">No</label>
			</span>
		</span>

	</div>
	<!-- Screen section two -->
	<div style="LEFT: 10px; POSITION: absolute; TOP: 306px; HEIGHT: 50px; WIDTH: 604px;" class="msgGroup" >			
		<span style="POSITION: absolute; TOP: 4px;LEFT: 4px;" class="msgLabel" >		
			<strong>Main Roads
		</span>
		<span style="TOP:25px;LEFT: 4px; POSITION:ABSOLUTE" class="msgLabel">
		Made Up/Adopted?
			<span style="TOP:-3px; LEFT:116px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optAdoptedYes" name="AdoptedRadioGroup" type="radio" value="1">
				<label for="optAdoptedYes" class="msgLabel">Yes</label>
			</span>
			<span style="TOP:-3px; LEFT:184px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optAdoptedNo" name="AdoptedRadioGroup" type="radio" value="0">
				<label for="optAdoptedNo" class="msgLabel">No</label>
			</span>
		</span>		
		<span style="TOP:25px;LEFT: 324px; POSITION:ABSOLUTE" class="msgLabel">
		To be Adopted?
			<span style="TOP:-3px; LEFT:116px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optToBeAdoptedYes" name="ToBeAdoptedRadioGroup" type="radio" value="1">
				<label for="optToBeAdoptedYes" class="msgLabel">Yes</label>
			</span>
			<span style="TOP:-3px; LEFT:184px; WIDTH: 50px; POSITION:ABSOLUTE">
				<input id="optToBeAdoptedNo" name="ToBeAdoptedRadioGroup" type="radio" value="0">
				<label for="optToBeAdoptedNo" class="msgLabel">No</label>
			</span>
		</span>
	</div>
</div>				
</form>
<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute;TOP:420;  WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP220attribs.asp" -->
<!-- #include FILE="Customise/AP220Customise.asp" -->

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
var csTenureLeashold = "2";
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
	FW030SetTitles("Valuation Report - Services","AP220",scScreenFunctions);

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
		frmScreen.cboWater.value = m_ValuationXML.GetAttribute("WATER");
		frmScreen.cboElectricity.value = m_ValuationXML.GetAttribute("ELECTRICITY");
		frmScreen.cboDrainage.value = m_ValuationXML.GetAttribute("DRAINAGE");
		frmScreen.cboDemandInArea.value = m_ValuationXML.GetAttribute("DEMANDINAREA");
		frmScreen.cboGas.value = m_ValuationXML.GetAttribute("GAS");
		frmScreen.txtOtherFactors.value  = m_ValuationXML.GetAttribute("OTHERFACTORSNOTES");
		frmScreen.txtBedrooms.value = m_ValuationXML.GetAttribute("NUMBEROFBEDROOMS");

		SetOptionValue(frmScreen.optMainServicesYes , frmScreen.optMainServicesNo, m_ValuationXML.GetAttribute("MAINSERVICES"))
		SetOptionValue(frmScreen.optHotWaterYes  , frmScreen.optHotWaterNo, m_ValuationXML.GetAttribute("HOTWATERCENTRALHEATING"))
		SetOptionValue(frmScreen.optOtherFactorsYes   , frmScreen.optOtherFactorsNo, m_ValuationXML.GetAttribute("OTHERFACTORS"))
		SetOptionValue(frmScreen.optDisrepairYes, frmScreen.optDisrepairNo, m_ValuationXML.GetAttribute("FUTURESALEABILITY"))
		SetOptionValue(frmScreen.optAdoptedYes , frmScreen.optAdoptedNo, m_ValuationXML.GetAttribute("ROADSMADEUPADOPTED"))
		SetOptionValue(frmScreen.optToBeAdoptedYes  , frmScreen.optToBeAdoptedNo, m_ValuationXML.GetAttribute("ROADSTOBEADOPTED"))
	}	
	else
	{
		frmScreen.optMainServicesNo.checked = "1";
		frmScreen.optHotWaterNo.checked = "1";
		frmScreen.optOtherFactorsNo.checked = "1";
		frmScreen.optDisrepairNo.checked = "1";
		frmScreen.optAdoptedNo.checked = "1";
		frmScreen.optToBeAdoptedNo.checked = "1";
		m_sReportNotFound = "1";
	}
	if(frmScreen.txtOtherFactors.value.length > 0 && frmScreen.optOtherFactorsYes.checked == true)
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtOtherFactors")			
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtOtherFactors")			
	}

	SetServices();
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
	var sGroupList = new Array("GasSupply","WaterSupply", "ElectricitySupply", "PropertyDrainage", "AreaDemand");

	if(XML.GetComboLists(document,sGroupList))
	{
		XML.PopulateCombo(document,frmScreen.cboWater,"WaterSupply",true);
		XML.PopulateCombo(document,frmScreen.cboGas,"GasSupply",true);
		XML.PopulateCombo(document,frmScreen.cboElectricity,"ElectricitySupply",true);
		XML.PopulateCombo(document,frmScreen.cboDrainage,"PropertyDrainage",true);
		XML.PopulateCombo(document,frmScreen.cboDemandInArea,"AreaDemand",true);
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

	<%/* EP976 - Check for the application has been accessed in Read-Only Mode */%>
	if (m_sReadOnly != "1")
		m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly", null);
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
		XML.SetAttribute("MAINSERVICES", GetOptionValue(frmScreen.optMainServicesYes));
		XML.SetAttribute("OTHERFACTORS", GetOptionValue(frmScreen.optOtherFactorsYes));
		XML.SetAttribute("FUTURESALEABILITY", GetOptionValue(frmScreen.optDisrepairYes));
		XML.SetAttribute("ROADSMADEUPADOPTED", GetOptionValue(frmScreen.optAdoptedYes));
		<%/* EP1196 - Amended radio reference */%>
		XML.SetAttribute("ROADSTOBEADOPTED", GetOptionValue(frmScreen.optToBeAdoptedYes));
		XML.SetAttribute("HOTWATERCENTRALHEATING", GetOptionValue(frmScreen.optHotWaterYes));
		XML.SetAttribute("WATER", frmScreen.cboWater.value);
		XML.SetAttribute("ELECTRICITY", frmScreen.cboElectricity.value);
		XML.SetAttribute("DRAINAGE", frmScreen.cboDrainage.value);
		XML.SetAttribute("DEMANDINAREA", frmScreen.cboDemandInArea.value);
		XML.SetAttribute("GAS", frmScreen.cboGas.value);
		XML.SetAttribute("WATER", frmScreen.cboWater.value);						
		XML.SetAttribute("OTHERFACTORSNOTES", frmScreen.txtOtherFactors.value);						
		XML.SetAttribute("NUMBEROFBEDROOMS", frmScreen.txtBedrooms.value);						

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

function spnMainServices.onclick()
{
	SetServices();
}

function SetServices()
{
	if (!frmScreen.optMainServicesYes.disabled)
	{
		if(frmScreen.optMainServicesYes.checked == false)

		{
			scScreenFunctions.SetFieldToWritable(frmScreen,"cboGas")
			scScreenFunctions.SetFieldToWritable(frmScreen,"cboWater")
			scScreenFunctions.SetFieldToWritable(frmScreen,"cboElectricity")
			scScreenFunctions.SetFieldToWritable(frmScreen,"cboDrainage")
		}
		else
		{
			frmScreen.cboGas.value = "";
			frmScreen.cboWater.value = ""
			frmScreen.cboElectricity.value = "";
			frmScreen.cboDrainage.value = "";
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboGas")
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboWater")
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboElectricity")
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboDrainage")
		}
	}
}


function spnOtherFactors.onclick()
{
	var sMessage = "Setting Other Factors to No will remove all related notes. Do you want to continue?";
	HandleOption(frmScreen.optOtherFactorsYes,frmScreen.txtOtherFactors,sMessage);
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

function btnNext.onclick()
{
	if(CommitChanges())
	{
		frmToAP230.submit();
	}
}

-->
</script>
</body>
</html>


