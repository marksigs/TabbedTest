<%@ LANGUAGE="JSCRIPT" %>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<form id="frmTogn215" method="post" action="gn215.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC015" method="post" action="DC015.asp" STYLE="DISPLAY: none"></form>
<form id="frmWelcomeMenu" method="post" action="mn010.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span> 

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 416px" class="msgGroup">
<span style="LEFT: 6px; POSITION: absolute; TOP: 22px" class="msgLabel">
	Distribuition Channel
	<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
		<select id="cboDistributionChannel" style="WIDTH: 130px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 333px; POSITION: absolute; TOP: 22px" class="msgLabel">
	Department
	<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
		<select id="cboDepartment" style="WIDTH: 130px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 248px; POSITION: absolute; TOP: -3px">&nbsp;</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 52px" class="msgLabel">
	Unit Name
	<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
		<select id="cboUnitName" style="WIDTH: 130px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 333px; POSITION: absolute; TOP: 52px" class="msgLabel">
	User Name
	<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
		<select id="cboUserName" style="WIDTH: 130px" class="msgCombo"></select>
	</span>
</span>
<span id="spnMortgageLenderDetails" style="LEFT: 4px; POSITION: absolute; TOP: 73px" class="msgLabelHead">
	<span>
		<input id="Mortgage" style="WIDTH: 15px; HEIGHT: 20px" type="checkbox" value="on" For="Mortgage" size="15" <LABEL>Mortgage</LABEL>
	</span>		
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 100px" class="msgLabel">
	Lender
	<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
		<select id="cboLenderName" style="WIDTH: 130px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 332px; WIDTH: 70px; POSITION: absolute; TOP: 100px; HEIGHT: 28px" class="msgLabel">
	Product Name
	<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
		<select id="cboProductName" style="WIDTH: 130px" class="msgCombo"></select>
	</span>
</span>	
<span id="spnApplication" style="LEFT: 4px; POSITION: absolute; TOP: 135px" class="msgLabelHead">
	<input id="Application" style="WIDTH: 17px; TOP: 5px; HEIGHT: 20px" type="checkbox" value="on" For="Application" size="17" <LABEL>Application</LABEL>
</span>	
<span style="LEFT: 330px; POSITION: absolute; TOP: 141px" class="msgLabel">
	Type
	<span style="LEFT: 80px; POSITION: absolute; TOP: 2px">
		<select id = "cboApplicationType" class="msgCombo"  style="WIDTH: 130px"></select> 
	</span>
</span>
<span style="LEFT: 2px; POSITION: absolute; TOP: 180px" class="msgLabel">
	From Stage
	<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
		<select id="cboFromStage" style="WIDTH: 130px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 331px; POSITION: absolute; TOP: 176px" class="msgLabel" id="SPAN1">
	To Stage
	<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
		<select id="cboToStage" style="WIDTH: 130px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 212px" class="msgLabel" id="SPAN1">
	Date From 
	<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
		<input id="txtDateFrom" maxlength="30" style="LEFT: -4px; WIDTH: 131px; POSITION: absolute; TOP: 4px; HEIGHT: 20px" class="msgTxt" size="18">
	</span>
</span>
<span style="LEFT: 329px; POSITION: absolute; TOP: 213px" class="msgLabel">
	Date To 
	<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
		<input id="txtDateTo" maxlength="30" style="LEFT: 4px; WIDTH: 129px; POSITION: absolute; TOP: 0px; HEIGHT: 20px" class="msgTxt" size="18">
	</span>
</span>
<span id="spnApplicationApproved" style="LEFT: 4px; POSITION: absolute; TOP: 260px" class="msgLabelHead">
	<input id="ApplicationApproved" style="WIDTH: 17px; TOP: 5px; HEIGHT: 20px" type="checkbox" value="on" For="ApplicationApproved" size="17" <LABEL>Application 
Approved</LABEL>
</span>
<span style="LEFT: 334px; POSITION: absolute; TOP: 260px" class="msgLabel">
	Month
	<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
		<input id="txtAppApprovedMonth" maxlength="30" style="LEFT: -85px; WIDTH: 57px; POSITION: absolute; TOP: 1px; HEIGHT: 20px" class="msgTxt" size="8">
	</span>
</span>
<span style="LEFT: 453px; POSITION: absolute; TOP: 260px" class="msgLabel">
	Year
	<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
		<input id="txtAppApprovedYear" maxlength="30" style="LEFT: -100px; WIDTH: 57px; POSITION: absolute; TOP: 1px; HEIGHT: 20px" class="msgTxt" size="8">
	</span>
</span>
<span id="spnSrcOfBusiness" style="LEFT: 4px; POSITION: absolute; TOP: 300px" class="msgLabelHead">
	<input id="SrcOfBusiness" style="WIDTH: 17px; TOP: 5px; HEIGHT: 20px" type="checkbox" value="on" For="SrcOfBusiness" size="17" <LABEL>Source Of Business</LABEL>
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 330px" class="msgLabel">
	Direct / Indirect
	<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
		<select id="cboDirectIndirect" style="WIDTH: 130px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 410px; POSITION: absolute; TOP: 325px">
	<input id="btnIntermediarySearch" style="WIDTH: 125px" type="button" value="Intermediary Search" class="msgButton" disabled>
</span>
<span style="LEFT: 6px; POSITION: absolute; TOP: 360px" class="msgLabel">
	Intermediary Panel ID 
	<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
		<input id="txtIntermediaryPanelID" maxlength="30" style="LEFT: -4px; WIDTH: 135px; POSITION: absolute; TOP: 3px; HEIGHT: 20px" class="msgTxt" size="19">
	</span>
</span>
<span style="LEFT: 296px; POSITION: absolute; TOP: 360px" class="msgLabel">
	Intermediary Name
	<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
		<input id="txtIntermediaryName" maxlength="30" style="LEFT: -25px; WIDTH: 134px; POSITION: absolute; TOP: 1px; HEIGHT: 20px" class="msgTxt" size="19">
	</span>	
</span>
<span style="LEFT: 4px; POSITION: absolute; TOP: 385px" class="msgLabelHead">
	<input id="ChkCancelledorDeclinedApplications" type="checkbox" value="on" For="ChkCancelledorDeclinedApplications" size="17" <LABEL>Include 
Cancelled/Declined Applications</LABEL>
</span>	
<span style="LEFT: 253px; POSITION: absolute; TOP: 385px" class="msgLabelHead">
	<input id="ProvideTotalLoanAmount" type="checkbox" value="on" For="ProvideTotalLoanAmount" size="17" <LABEL>Provide Total Loan Amount</LABEL>
</span>
<span style="LEFT: 433px; POSITION: absolute; TOP: 385px">
	<input id="btnSearch" style="WIDTH: 100px" type="button" value="Search" class="msgButton">
</span>	
</div>		
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 480px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen  */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->	

<% /* File containing field attributes (remove if not required) */ %>
	<!-- #include FILE="attribs/pp210attribs.asp"  -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var scScreenFunctions;
var m_blnReadOnly = false;


function window.onload()
{
	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
	var sButtonList = new Array("Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Payment Protection - Extended Search ","PP210",scScreenFunctions); 
	RetrieveContextData();
	SetMasks();
	Validation_Init();
	// Initialise the table
	PopulateCombos()
	//frmScreen.txtPanelId.focus();
	//DisableMainButton("Submit");
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP210");
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
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
}

function frmScreen.btnSearch.onclick()
{
	alert ("under construction");
/*	
	//save the combo settings in idXML context to use when we return
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	//Preparing XML Request string 
	XML.CreateActiveTag("Request");
	XML.SetAttribute("Operation", "ExtendedSearch");
	XML.SetAttribute("UnitId",scScreenFunctions.GetContextParameter(window,"idUnitId",null));
	XML.SetAttribute("UserId",scScreenFunctions.GetContextParameter(window,"idUserId",null));
	XML.SetAttribute("ChannelId",scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null));
	XML.SetAttribute("MachineId",scScreenFunctions.GetContextParameter(window,"idMachineId",null));
	XML.CreateTag("Applicationnumber",null);
	XML.CreateTag("AccountNumber",null);
	if (frmScreen.ChkCancelledorDeclinedApplications.checked == true) 
	{
		XML.CreateTag("IncludeCancelledApps","1");
		XML.CreateTag("IncludeDeclinedApps","1");
	}
	else
	{
		XML.CreateTag("IncludeCancelledApps","0");
		XML.CreateTag("IncludeDeclinedApps","0");
	}
	XML.CreateTag("ChannelId",frmScreen.cboDistributionChannel.value );
	XML.CreateTag("DepartmentId",frmScreen.cboDepartment);
	XML.CreateTag("UnitId",null);
	XML.CreateTag("UnitName",frmScreen.cboUnitName.value);
	XML.CreateTag("UserId",null);
	XML.CreateTag("UserName",frmScreen.cboUsesrName .value);
	XML.CreateTag("LenderName",frmScreen.cboLenderName.value);
	XML.CreateTag("ProductName",frmScreen.cboProductName.value );
	XML.CreateTag("ApplicationType",frmScreen.cboApplicationType.value );
	XML.CreateTag("FromStage",frmScreen.cboFromStage.value );
	XML.CreateTag("ToStage",frmScreen.cboToStage.value);
	XML.CreateTag("DateFrom",frmScreen.txtDateFrom.value);
	XML.CreateTag("DateTo",frmScreen.txtDateTo.value);
	XML.CreateTag("AppApprovedMonth",frmScreen.txtAppApprovedMonth.value) ;
	XML.CreateTag("AppApprovedYear",frmScreen.txtAppApprovedYear.value) ;
	XML.CreateTag("SOBDirectorIndirect",frmScreen.cboDirectIndirect.value);
	XML.CreateTag("SOBPannelId",frmScreen.txtIntermediaryPanelID.value  );
	XML.CreateTag("SOBName",frmScreen.txtIntermediaryName.value);
	scScreenFunctions.SetContextParameter(window,"idXML",XML.ActiveTag.xml);
	scScreenFunctions.SetContextParameter(window,"idMetaAction","ExtendedSearch");
	frmTogn215.submit();
*/
}
function frmScreen.btnIntermediarySearch.onclick()
{
	//frmToDC015.submit();
	alert("under construction");
}

function PopulateCombos()
{	
	var XMLDirectIndirect = null;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("Direct/Indirect");
	
	if (XML.GetComboLists(document,sGroupList))
	{
		XMLDirectIndirect = XML.GetComboListXML("Direct/Indirect");
		var blnSuccess = true;
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboDirectIndirect,XMLDirectIndirect,false);
		if(blnSuccess == false)
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			DisableMainButton("Submit");
		}
		else
		{
			// default combo DirectIndirect to 'Direct'
			frmScreen.cboDirectIndirect.value = 1;
		}
	}	
	XML = null;
}

function frmScreen.cboDirectIndirect.onchange()
{
	var selIndex = frmScreen.cboDirectIndirect.selectedIndex;

	if(selIndex != -1 && !scScreenFunctions.IsOptionValidationType(frmScreen.cboDirectIndirect,selIndex,"I"))
		frmScreen.btnIntermediarySearch.disabled = true;			
	else
		frmScreen.btnIntermediarySearch.disabled = false;			
}

function btnCancel.onclick()
{
		frmWelcomeMenu.submit();
}

-->
</script>
</body>
</html>
