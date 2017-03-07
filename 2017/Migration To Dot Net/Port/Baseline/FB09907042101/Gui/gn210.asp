<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<% /*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      gn210.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Application Enquiry Quick Search screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MV		30/10/2000	Created
APS		08/01/2001	SYS1801
CL		05/03/01	SYS1920 Read only functionality added
MEVA	24/04/02	SYS2242 Resetting of screen values
LD		23/05/02	SYS4727 Use cached versions of frame functions
TW      09/10/2002  SYS5115 Modified to incorporate client validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

DPF		11/09/2002	APWP1/BM010 - Amended Product Name combo to display product	code & product name for clarity
MO		01/10/2002	BM061/INWP1 - Amended Intermediary Search to Introducer Search
DPF		03/10/2002	BMIDS00555 - Fixed bug with Product Name combo (wasn't emptying list)
MO		01/10/2002	BMIDS00663 - Fix bug pass the direct/indirect validation type not the value
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
GHun	18/12/2002	BM0034 - Show UserName instead of UserID in user name combo, improve alignment and spacing
BS		17/04/2003	BM0499 - Add <ALL> option to Application Type combo
HMA     08/09/2003  BMIDS607   - Correct combo boxes in extended search
HMA     24/09/2003  BMIDS607   - Further corrections.
HMA     30/09/2003  BMIDS607 Check for sufficient data to perform search.
MC		20/04/2004	BMIDS517	Title text added (DC015p - Find Intermediary)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

GHun	26/07/2005	MAR10 Change Introducer ID label to be more generic
HMA     21/09/2005  MAR46 Add search for Product Switch and Transfer of Equity completed applications
MahaT	10/10/2005	MAR94 Disable Introducer Search button.
PE		21/02/2006	MAR1152 Remove authority level checks for extended search GN200/GN210
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<form id="frmTogn215" method="post" action="gn215.asp" style="DISPLAY: none"></form>
<form id="frmToDC015" method="post" action="DC015.asp" style="DISPLAY: none"></form>
<form id="frmWelcomeMenu" method="post" action="mn010.asp" style="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span> 

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4 >
<div id="divBackground1" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 470px" class="msgGroup">
<span id="spnDistributionChannel" style="LEFT: 10px; POSITION: absolute; TOP: 14px" class="msgLabel">
	Distribution Channel
	<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
		<select id="cboDistributionChannel" style="WIDTH: 180px" class="msgCombo"></select>
	</span>
</span>
<span id="spnDepartmentNames" style="LEFT: 335px; POSITION: absolute; TOP: 14px" class="msgLabel">
	Department
	<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
		<select id="cboDepartment" style="WIDTH: 180px" class="msgCombo"></select>
	</span>
</span>
<span id="spnUnitNames" style="LEFT: 10px; POSITION: absolute; TOP: 40px" class="msgLabel">
	Unit Name
	<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
		<select id="cboUnitName" style="WIDTH: 180px" class="msgCombo"></select>
	</span>
</span>
<span id="spnUserNames" style="LEFT: 335px; POSITION: absolute; TOP: 40px" class="msgLabel">
	User Name
	<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
		<select id="cboUserName" style="WIDTH: 180px" class="msgCombo"></select>
	</span>
</span>

<hr style="LEFT: 10px; WIDTH: 584px; TOP: 66; POSITION: absolute">

<span id="spnMortgageLenderDetails" style="LEFT: 6px; POSITION: absolute; TOP: 75px" >
	<span id="spnchkMortgageLenderDetails">
		<input id="chkMortgageLenderDetails" type="checkbox" value="1" >
		<label for="chkMortgageLenderDetails" class="msgLabel">Mortgage</label>
	</span>		
	<span id="spnLenderProductDetails">
		<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
			Lender
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<select id="cboLenderName" style="WIDTH: 295px" class="msgCombo"></select>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 56px" class="msgLabel">
			Product
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<select id="cboProductName" style="WIDTH: 470px" class="msgCombo"></select>
			</span>
		</span>
	</span>
</span>

<hr style="LEFT: 10px; WIDTH: 584px; TOP: 157; POSITION: absolute">

<span id="spnApplication" style="LEFT: 6px; POSITION: absolute; TOP: 167px" > 
	<span id="spnchkApplication">
		<input id="chkApplication" type="checkbox" value="1">
		<label for="chkApplication" class="msgLabel">Application</label>
	</span>
	<span id="spnApplicationDetails">
		<span style="LEFT: 329px; POSITION: absolute; TOP: 5px" class="msgLabel">
			Type
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<select id="cboApplicationType" style="WIDTH: 180px" class="msgCombo" ></select> 
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 31px" class="msgLabel">
			From Stage
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<select id="cboFromStage" style="WIDTH: 180px" class="msgCombo"></select>
			</span>
		</span>
		<span style="LEFT: 329px; POSITION: absolute; TOP: 31px; WIDTH: 50px" class="msgLabel">
			To Stage
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<select id="cboToStage" style="WIDTH: 180px" class="msgCombo"></select>
			</span>
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 57px" class="msgLabel" >
			Date From 
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<input id="txtDateFrom" maxlength="10" style="WIDTH: 70px; POSITION: absolute" class="msgTxt">
			</span>
		</span>
		<span style="LEFT: 329px; POSITION: absolute; TOP: 57px; WIDTH: 50px" class="msgLabel">
			Date To 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<input id="txtDateTo" maxlength="10" style="WIDTH: 70px; POSITION: absolute" class="msgTxt">
			</span>
		</span>
	</span>
</span>

<hr style="LEFT: 10px; WIDTH: 584px; TOP: 251; POSITION: absolute">

<span id="spnApplicationApproved" style="LEFT: 6px; POSITION: absolute; TOP: 261px" >
	<span id="spnchkApplicationApproved">
		<input id="chkApplicationApproved" type="checkbox" value="on">
		<label for="chkApplicationApproved" class="msgLabel">Application Approved</label>
	</span>
	<span id="spnApplicationApprovedDetails">
		<span style="LEFT: 329px; POSITION: absolute; TOP: 5px" class="msgLabel">
			Month&nbsp;&amp;&nbsp;Year
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
				<select id="cboApplicationApprovedMonth" style="WIDTH: 88px" class="msgCombo"></select>
			</span>
			<span style="LEFT: 175px; POSITION: absolute; TOP: -3px">
				<select id="cboApplicationApprovedYear" style="WIDTH: 57px" class="msgCombo"></select>
			</span>
		</span>		
	</span>
</span>

<hr style="LEFT: 10px; WIDTH: 584px; TOP: 293; POSITION: absolute">

<span id="spnSrcOfBusiness" style="LEFT: 6px; POSITION: absolute; TOP: 302px">
	<span id="spnchkSrcOfBusiness">
		<input id="chkSrcOfBusiness" type="checkbox" value="1">
		<label for="chkSrcOfBusiness" class="msgLabel">Source Of Business</label>
	</span>
	<span id="spnSrcOfBusinessDetails">
		<span id = "spnIntermediaryType" style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
			Direct / Indirect
			<span style="LEFT: 115px; POSITION: absolute; TOP: -3px">
				<select id="cboDirectIndirect" style="WIDTH: 180px" class="msgCombo"></select>
			</span>
		</span>
		<span style="LEFT: 329px; POSITION: absolute; TOP: 25px">
			<input id="btnIntermediarySearch" style="WIDTH: 100px" type="button" value="Introducer Search" class="msgButton" disabled>
		</span>
		<span id="spnIntermediaryDetails">
			<span style="LEFT: 4px; POSITION: absolute; TOP: 55px" class="msgLabel">
				FSA / Introducer<br>System ID<br>
				<span style="LEFT: 115px; POSITION: absolute; TOP: 0">
					<input id="txtIntroducerId" maxlength="30" style="WIDTH: 180px; POSITION: absolute" class="msgTxt" disabled>
				</span>
			</span>
			<span id="spnIntroducerName" style="LEFT: 329px; POSITION: absolute; TOP: 55px" class="msgLabel">
				Introducer<br>Name
				<span style="LEFT: 80px; POSITION: absolute; TOP: 0">
					<input id="txtIntroducerName" maxlength="30" style="WIDTH: 180px; POSITION: absolute" class="msgTxt" disabled>
				</span>	
			</span>
		</span>
	</span>
</span>

<hr style="LEFT: 10px; WIDTH: 584px; TOP: 395; POSITION: absolute">

<span style="LEFT: 6px; POSITION: absolute; TOP: 404px">
	<input id="ChkCancelledorDeclinedApplications" type="checkbox" value="1">
	<label for="ChkCancelledorDeclinedApplications" class="msgLabelHead">Include Cancelled/Declined Applications</label>
</span>	
<span style="LEFT: 6px; POSITION: absolute; TOP: 424px">
	<input id="ChkProductSwitchApplications" type="checkbox" value="1">
	<label for="ChkProductSwitchApplications" class="msgLabelHead">Include Product Switch Completed Applications</label>
</span>	
<span style="LEFT: 6px; POSITION: absolute; TOP: 444px">
	<input id="ChkTOEApplications" type="checkbox" value="1">
	<label for="ChkTOEApplications" class="msgLabelHead">Include Transfer of Equity Completed Applications</label>
</span>	
<span style="LEFT: 493px; POSITION: absolute; TOP: 404px">
	<input id="btnSearch" style="WIDTH: 100px" type="button" value="Search" class="msgButton">
</span>	
</div>		
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 538px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen  */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->	

<% /* File containing field attributes (remove if not required) */ %>
	<!-- #include FILE="attribs/gn210attribs.asp"  -->

<% /* Specify Code Here */ %>
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = null;
<% /*MO, INWP1 var m_sIntermediaryGUID ;*/ %>
var m_sIntroducerId;
var XMLProductStartDate = null;
var XMLTemp = null;
var strFromStageId;
var strToStageId;

// These Varibles are used to store values retrieved from GlobalParameter DB Table
// to Check the Authority level
// var m_iAppEnqChannelRole; MAR1152
// var m_iAppEnqDepartmentRole; MAR1152
// var m_iAppEnqUnitRole; MAR1152
// var m_iAppEnqUserRole; MAR1152
// var m_iAppEnqMortgageRole; MAR1152
// var m_iAppEnqApplicationRole; MAR1152
// var m_iAppEnqApprovedRole; MAR1152
// var m_iAppEnqBusinessSourceRole; MAR1152
var m_iCancelledStageValue;
var m_iDeclinedStageValue;
var m_iApplicationApprovalYears;

// These Varibles are used to Validate the Authority level
var m_blnAppEnqChannelRole  = true ;
var m_blnAppEnqDepartmentRole  = true ;
var m_blnAppEnqUnitRole  = true ;
var m_blnAppEnqUserRole  = true ;
var m_blnAppEnqMortgageRole = true ;
var m_blnAppEnqApplicationRole  = true ;
var m_blnAppEnqApprovedRole  = true ;
var m_blnAppEnqBusinessSourceRole  = true ;

// These Varibles are used to check the UserAccess With Screen Controls
// Whether selected or not
var m_chkblnAppEnqChannelRole  = false ;
var m_chkblnAppEnqDepartmentRole  = false ;
var m_chkblnAppEnqUnitRole  =false ;
var m_chkblnAppEnqUserRole  = false ;
var m_chkblnAppEnqMortgageRole = false ;
var m_chkblnAppEnqApplicationRole  = false ;
var m_chkblnAppEnqApprovedRole  = false ;
var m_chkblnAppEnqBusinessSourceRole  = false;

//Local Variables to Store Context Parameters
var m_DistributionChannelId ;
var m_DistributionChannelName;
var m_DepartmentId ;
var m_DepartmentName;
var m_UnitId;
var m_UnitName;
var m_UserId;
var m_UserName;
var m_PreviousDistributionChannel = 0;
var m_PreviousDepartment= 0;
var m_PreviousUnit = 0;
var m_PreviousUser = 0;
var m_PreviousLenderName = 0;
var blnFormLoad = false;
var scScreenFunctions;
var XML = null;
var sXML = null;
var sProductCode, sProductName, sStartDate;

<%/* MO - 01/10/2002 - BM061/INWP1 */%>
var m_MachineId ;

<% /* BS BM0499 17/04/03 */ %>
var m_AppType = "";

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Cancel");
	ShowMainButtons(sButtonList);
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions.HideCollection(spnIntroducerName);
	FW030SetTitles("Application Enquiry - Search ","GN210",scScreenFunctions); 
	SetMasks();
	Validation_Init();
	if ( GetParameterValues ()== true ) 
	{
		RetrieveContextData();
		sXML = scScreenFunctions.GetContextParameter(window,"idXML","");
		if ( sXML != "") PopulateScreenValues(sXML);
		ValidateUserAuthorityLevel();
	}
	else
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	
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
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_iContextRole = scScreenFunctions.GetContextParameter(window,"idRole",null);
	m_DistributionChannelId = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
	m_DistributionChannelName = scScreenFunctions.GetContextParameter(window,"idDistributionChannelName",null);
	m_DepartmentId = scScreenFunctions.GetContextParameter(window,"idDepartmentId",null);
	m_DepartmentName = scScreenFunctions.GetContextParameter(window,"idDepartmentName",null);
	m_UnitId = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	m_UnitName = scScreenFunctions.GetContextParameter(window,"idUnitName",null);
	m_UserId = scScreenFunctions.GetContextParameter(window,"idUserId",null);
	m_UserName = scScreenFunctions.GetContextParameter(window,"idUserName",null);
	
	<%/* MO - 01/10/2002 - BM061/INWP1 */%>
	m_MachineId = scScreenFunctions.GetContextParameter(window,"idMachineId",null);
	
}

function GetParameterValues ()
{
	//MAR1152: REMOVE AUTHORITY LEVEL CHECKS FOR EXTEND SEARCH GN200/GN210

	// This is a Function to retrieve the Global Parameter Values from database 	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_iAppApprovalYears = XML.GetGlobalParameterAmount(document,"APPLICATIONAPPROVALYEARS");
	XML = null; 
	return true;
}

function ValidateUserAuthorityLevel()
{

	// MAR1152: REMOVE AUTHORITY LEVEL CHECKS FOR EXTEND SEARCH GN200/GN210

	//DistributionChannelList
	PopulateDistributionChannelCombo();
	SetComboDefaultValue(frmScreen.cboDistributionChannel,m_DistributionChannelId,m_DistributionChannelName);

	//Departmentlist
	PopulateDepartmentCombo();
	SetComboDefaultValue(frmScreen.cboDepartment,m_DepartmentId,m_DepartmentName);
	
	//UnitList
	PopulateUnitListCombo();	
	SetComboDefaultValue(frmScreen.cboUnitName,m_UnitId,m_UnitName); 
	
	//UserList
	PopulateUserListCombo();
	SetComboDefaultValue(frmScreen.cboUserName,m_UserId,m_UserName); 
	
	if (frmScreen.chkMortgageLenderDetails.checked == false){
		//MortgageLender
		scScreenFunctions.SetCollectionState(spnMortgageLenderDetails,"R");
		scScreenFunctions.SetCollectionState(spnchkMortgageLenderDetails,"W");
	}
		
	if (frmScreen.chkApplication.checked == false) {
		//Application		
		scScreenFunctions.SetCollectionState(spnApplication,"R");
		scScreenFunctions.SetCollectionState(spnchkApplication,"W");
	}
		
	if (frmScreen.chkApplicationApproved.checked == false){
		// ApplicationApproved			
		scScreenFunctions.SetCollectionState(spnApplicationApproved,"R"); 
		scScreenFunctions.SetCollectionState(spnchkApplicationApproved,"W");
	}
	
	if (frmScreen.chkSrcOfBusiness.checked == false){			
		// Source of Business	
		scScreenFunctions.SetCollectionState(spnSrcOfBusiness,"R"); 
		scScreenFunctions.SetCollectionState(spnchkSrcOfBusiness,"W"); 
	}

}

function PopulateDistributionChannelCombo()
{	
	// This is a function to load DistributionChannelID's and ChannelNames into the Combo
	var blnSuccess = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.RunASP(document,"FindDistributionChannelList.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0])
	{
		//If it is Success then Load into Combo box
		blnSuccess = XML.PopulateComboFromXML(document,frmScreen.cboDistributionChannel,XML.XMLDocument,false,false,"DISTRIBUTIONCHANNEL","CHANNELID","CHANNELNAME");
		if(blnSuccess == false)
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			return;
		}
	}
	else 
	if(ErrorReturn[1] == ErrorTypes[0])
	{
		alert("No DistributionChannels are available");
		
		return;
	}
	XML= null;
}

function frmScreen.cboDistributionChannel.onclick ()
{
	// Store Selected list item index 
	m_PreviousDistributionChannel =  frmScreen.cboDistributionChannel.selectedIndex; 
}

<% /* MO 02/10/2002 BMIDS00502 - If the Introducer Id changes, remove the Introducer name */ %>
function frmScreen.txtIntroducerId.onchange() {
	frmScreen.txtIntroducerName.value = "";
	scScreenFunctions.HideCollection(spnIntroducerName);
}

function frmScreen.cboDistributionChannel.onchange()
{
	
	if ( m_blnAppEnqDepartmentRole == true ) 
	{	
		scScreenFunctions.SetFieldState(frmScreen, "cboDepartment",  "W");
		PopulateDepartmentCombo();
		
		<% /* BMIDS607 Clear Unit and User combo entries */ %>
		ResetCombo(frmScreen.cboUnitName);
		ResetCombo(frmScreen.cboUserName);
	}
} 

function PopulateDepartmentCombo()
{
	//This is  a Function to populate DepartmentCombo with DepartmentID's and DepartmentName
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window);
	XML.CreateActiveTag("DEPARTMENT");
	XML.CreateTag("CHANNELID",frmScreen.cboDistributionChannel.value );
	XML.RunASP(document,"FindDepartmentList.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0])
	{
		blnSuccess = XML.PopulateComboFromXML(document,frmScreen.cboDepartment,XML.XMLDocument,false,false,"DEPARTMENT","DEPARTMENTID","DEPARTMENTNAME");
		if(blnSuccess == false)
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			return;
		}
		else
			<% /*BMIDS607   Clear the current Department value.
			SetComboDefaultValue(frmScreen.cboDepartment  ,m_DepartmentId,m_DepartmentName);*/ %>
			frmScreen.cboDepartment.selectedIndex = -1;
	}
	else
	if(ErrorReturn[1] == ErrorTypes[0])
	{	
		alert("No Departments are available");
		<% /*BMIDS607   Clear the current Department value.
		frmScreen.cboDistributionChannel.selectedIndex = m_PreviousDistributionChannel; */ %>
		frmScreen.cboDepartment.selectedIndex = -1;
		ResetCombo(frmScreen.cboDepartment);
	}
	XML= null;
}

function frmScreen.cboDepartment.onclick ()
{
	m_PreviousDepartment =  frmScreen.cboDepartment.selectedIndex; 
}


function frmScreen.cboDepartment.onchange()
{
	if (m_blnAppEnqUnitRole == true ) 
	{
		scScreenFunctions.SetFieldState(frmScreen, "cboUnitName",  "W");
		PopulateUnitListCombo();
		
		<% /* BMIDS607 Clear User combo entries */ %>
		ResetCombo(frmScreen.cboUserName);
		
	}
}

function PopulateUnitListCombo()
{
	//This is a Function to Populate UnitList COmbo with UnitId's and UnitNames
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	//Preparing XML Request string 
	XML.CreateRequestTag(window);
	XML.CreateActiveTag("UNIT");
	XML.CreateTag("CHANNELID",frmScreen.cboDistributionChannel.value);
	XML.CreateTag("DEPARTMENTID",frmScreen.cboDepartment.value  );		
	XML.RunASP(document,"FindUnitList.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0])
	{
		blnSuccess = XML.PopulateComboFromXML(document,frmScreen.cboUnitName,XML.XMLDocument,false,false, "UNIT","UNITID","UNITNAME");	
		if(blnSuccess == false)
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			return;
		}
		else
		<% /* BM0034 <All> is added in PopulateUserListCombo */ %>
		<% /* {
			var TagOPTION = document.createElement("OPTION");
			TagOPTION.value = "<ALL>"
			TagOPTION.text  = "<ALL>";
			frmScreen.cboUserName.add(TagOPTION);
			TagOPTION = null;
		*/ %>
			<% /* BMIDS607  Clear the current Unit name */  %>
			frmScreen.cboUnitName.selectedIndex = -1;
	
			return;
		<% /* } */ %>
	}
	else
	if(ErrorReturn[1] == ErrorTypes[0])
	{	
		alert("No Units are available");
		<% /* BMIDS607  Clear the current Unit name 
		frmScreen.cboDepartment.selectedIndex = m_PreviousDepartment; */ %>
		frmScreen.cboUnitName.selectedIndex = -1;
		ResetCombo(frmScreen.cboUnitName)	
		return;
	}
	XML= null;
}

function frmScreen.cboUnitName.onclick ()
{
	m_PreviousUnit =  frmScreen.cboUnitName.selectedIndex; 
}

function frmScreen.cboUnitName.onchange()
{
	if (m_blnAppEnqUserRole == true ) 
	{
		scScreenFunctions.SetFieldState(frmScreen, "cboUserName",  "W");
		PopulateUserListCombo();
	}
}

function PopulateUserListCombo()
{
	// This is a Function to Populate userList Combo with userId and Usernames
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	//Preparing XML Request string 
	XML.CreateRequestTag(window);
	XML.CreateActiveTag("USERLIST");
	XML.CreateTag("UNITID",frmScreen.cboUnitName.value );
	XML.RunASP(document,"FindUserList.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0])
	{	
		<% /* BM0034 Clear existing combo entries before populating */ %>
		ResetCombo(frmScreen.cboUserName);
	
		<% /* BM0034 <ALL> should be displayed at the top of the list */ %>
		var TagOPTION = document.createElement("OPTION");
		TagOPTION.value = "<ALL>"
		TagOPTION.text  = "<ALL>";
		frmScreen.cboUserName.add(TagOPTION);
		TagOPTION = null;
		
		<% /* BM0034 */ %>
		<%
		// PopulateComboFromXML can no longer be used as 2 values need to concatenated to display the name
		// blnSuccess = XML.PopulateComboFromXML(document,frmScreen.cboUserName,XML.XMLDocument,false,false,"OMIGAUSER","USERID","USERID");
		%>
		var TagListLISTENTRY = XML.CreateTagList("OMIGAUSER");
		XML.ActiveTagList = TagListLISTENTRY;
		
		for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
		{	
			XML.SelectTagListItem(nLoop);
			TagOPTION = document.createElement("OPTION");
				
			TagOPTION.value = XML.GetTagText("USERID");
			TagOPTION.text = XML.GetTagText("USERFORENAME") + " " + XML.GetTagText("USERSURNAME");;
			frmScreen.cboUserName.options.add(TagOPTION);				
		}

		frmScreen.cboUserName.selectedIndex = -1;
		<% /* BM0034 End */ %>
	}	
	else
	if(ErrorReturn[1] == ErrorTypes[0])
	{	
		alert("No Users are available");
		<% /* Clear the current User Name
		frmScreen.cboUnitName.selectedIndex = m_PreviousUnit; */ %>
		frmScreen.cboUserName.selectedIndex = -1;
		ResetCombo(frmScreen.cboUserName);
	}
	XML= null;
}	

function frmScreen.chkMortgageLenderDetails.onclick ()
{
	if (m_blnAppEnqMortgageRole == true) 
	{
		if (frmScreen.chkMortgageLenderDetails.checked == true ) 
		{
			scScreenFunctions.SetCollectionState(spnLenderProductDetails,"W");
			if (m_chkblnAppEnqMortgageRole == false )	
			{
				PopulateLenderNamesCombo();
				PopulateProductNamesCombo();
				m_chkblnAppEnqMortgageRole = true;
			}
		}
		else
		{
			if (frmScreen.chkMortgageLenderDetails.checked == false ) 
			{
				//
				// Reset Screen values
				//
				frmScreen.cboLenderName.selectedIndex=-1;
				frmScreen.cboProductName.selectedIndex=-1;
				
				scScreenFunctions.SetCollectionState(spnLenderProductDetails,"R");
			}
		}
	}
}	

function PopulateLenderNamesCombo()
{
	// This is a Function to Populate Lendernames Combo with OrganiationID's and LenderNames
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateActiveTag("MORTGAGELENDER");
	// 	XML.RunASP(document,"FindLenderNamesList.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"FindLenderNamesList.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0])
	{
		blnSuccess = XML.PopulateComboFromXML(document,frmScreen.cboLenderName,XML.XMLDocument,false,false,"MORTGAGELENDER","ORGANISATIONID","LENDERNAME")	;
		if(blnSuccess == false)
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			return;
		}
	}
	else
	if(ErrorReturn[1] == ErrorTypes[0])
		alert("No Lenders are available");
		
	XML= null;
}

function frmScreen.cboLenderName.onclick ()
{
	m_PreviousLenderName =  frmScreen.cboLenderName.selectedIndex; 
}

function frmScreen.cboLenderName.onchange ()
{
	if (frmScreen.chkMortgageLenderDetails.checked == true ) 
		PopulateProductNamesCombo();
}		
	
function PopulateProductNamesCombo()
{
	// This is a Function To populate ProductCombo with Productcode and ProductNames
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLProductStartDate = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"PRODUCTNAMES");
	XML.CreateActiveTag("MORTGAGEPRODUCT");
	XML.CreateTag("ORGANISATIONID",frmScreen.cboLenderName.value );
	// 	XML.RunASP(document,"FindProductNamesList.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"FindProductNamesList.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0])
	{
		//blnSuccess = XML.PopulateComboFromXML(document,frmScreen.cboProductName ,XML.XMLDocument,false,false,"MORTGAGEPRODUCT","MORTGAGEPRODUCTCODE","MORTGAGEPRODUCTCODE");
		
		//DPF 11/09/02 APWP3/BM008 - manually build up Combo list so we can concatenate the code and product name
		XML.ActiveTag = null;
		XML.CreateTagList("MORTGAGEPRODUCT"); 			
		var nNumProducts = XML.ActiveTagList.length;
		var TagOPTION;				
					
		if (nNumProducts > 0) 
		{
			while(frmScreen.cboProductName.options.length > 0) frmScreen.cboProductName.options.remove(0);
			var nIndex = null;
			for(nIndex = 1; nIndex <= nNumProducts; nIndex++) 
			{
				XML.SelectTagListItem(nIndex - 1)
				sProductCode = XML.GetTagText("MORTGAGEPRODUCTCODE");
				sProductName = XML.GetTagText("PRODUCTNAME");
				sStartDate = XML.GetTagText("STARTDATE");
						
				TagOPTION	= document.createElement("OPTION");
				TagOPTION.value	= sProductCode
				TagOPTION.text	= sProductCode + " - " + sProductName;
				<% /* DPF 01/10/2002 - AQR00555. */ %>
				TagOPTION.setAttribute("ProductName", sProductName);
				TagOPTION.setAttribute("StartDate", sStartDate);
				//TagOPTION.StartDate = sStartDate
				frmScreen.cboProductName.add(TagOPTION);
							
			}

			// Default to the first option
			frmScreen.cboProductName.selectedIndex = 0;
		}				

		<% /* END OF GREGS COD! */ %>
		
		if(blnSuccess == false)
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			return;
		}
	}
	else
	if(ErrorReturn[1] == ErrorTypes[0])
	{	
		alert("No Products are available");
		frmScreen.cboLenderName.selectedIndex = m_PreviousLenderName;
		return;
	}
	XML = null;
}	

function frmScreen.chkApplication.onclick ()
{
	if (m_blnAppEnqApplicationRole == true )
	{
		if (frmScreen.chkApplication.checked == true ) 
		{
			scScreenFunctions.SetCollectionState(spnApplicationDetails,"W");
			if (m_chkblnAppEnqApplicationRole == false )
			{
				PopulateComboLists("TypeOfMortgage",frmScreen.cboApplicationType );
				<% /* BS BM0499 17/04/03 */ %>
				var TagOPTION = document.createElement("OPTION");
				TagOPTION.value = "<ALL>"
				TagOPTION.text  = "<ALL>";
				frmScreen.cboApplicationType.add(TagOPTION);
				TagOPTION = null;
				<% /* BS BM0499 End 17/04/03 */ %>				
				PopulateApplicationStages();
				m_chkblnAppApplicationRole = true;
			}
		}
		else
		{
			if (frmScreen.chkApplication.checked == false ) 
			{
				//
				// Reset Screen values
				//
				frmScreen.cboApplicationType.selectedIndex=-1;
				frmScreen.cboFromStage.selectedIndex=-1;
				frmScreen.cboToStage.selectedIndex=-1;
				frmScreen.txtDateFrom.value="";
				frmScreen.txtDateTo.value="";

				scScreenFunctions.SetCollectionState(spnApplicationDetails,"R");
			}
		}
	}
}

function PopulateApplicationStages()
{
	// This is a Function to Populate Application Stages 
	var blnSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"FINDSTAGELIST");
	XML.CreateActiveTag("STAGE");
	XML.SetAttribute("ACTIVITYID","10");
	XML.SetAttribute("EXCEPTIONSTAGEINDICATOR","0");
	// 	XML.RunASP(document,"FindStageList.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"FindStageList.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0])
	{
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboFromStage,XML.XMLDocument,true,true,"STAGE","STAGESEQUENCENO","STAGENAME");
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboToStage,XML.XMLDocument,true,true,"STAGE","STAGESEQUENCENO","STAGENAME");

		if(blnSuccess == false)
		{
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			return;
		}
	}
	else
	if(ErrorReturn[1] == ErrorTypes[0])
		alert("No Stages are available");
		
	XML= null;
}

<% /*function frmScreen.cboToStage.onblur()
{
	ValidateStages();
}

function frmScreen.txtDateFrom.onblur()
{
	ValidateDates();
}

function frmScreen.txtDateTo.onblur()
{
	ValidateDates();		
}*/ %>

function frmScreen.chkApplicationApproved.onclick ()
{
	if (m_blnAppEnqApprovedRole == true)
	{
		scScreenFunctions.SetCollectionState(spnApplicationApprovedDetails,"R");
		if (frmScreen.chkApplicationApproved.checked == true ) 
		{
			scScreenFunctions.SetCollectionState(spnApplicationApprovedDetails,"W");
			if (m_chkblnAppEnqApprovedRole == false)
			{
				PopulateComboLists("MonthsInYear",frmScreen.cboApplicationApprovedMonth);
				PopulateApplicationApporvalYears();
				m_chkblnAppEnqApprovedRole =true;
			}
		}
		else
		{
			//
			// Reset Screen values
			//
			frmScreen.cboApplicationApprovedMonth.selectedIndex=-1;
			frmScreen.cboApplicationApprovedYear.selectedIndex=-1;
		}
	}
}	

 function PopulateApplicationApporvalYears() 
{
	// Function to load up last APPAPPROVALYEARS i.e. [ year(sysdate) - m_iAppApprovalYears ] to [ Year(sysDate) ]
	var nYear ;
	<% /* MO - BMIDS00807 */ %>
	var dtTempCurrentDate = scScreenFunctions.GetAppServerDate(true);
	<% /* var dtTempCurrentDate = new Date(); */ %>

	for ( nYear = (dtTempCurrentDate.getFullYear() - m_iAppApprovalYears)  ; nYear <= dtTempCurrentDate.getFullYear()  ; nYear++)
	{
		var TagOPTION = document.createElement("OPTION");
		TagOPTION.value	= nYear;
		TagOPTION.text = nYear;
		frmScreen.cboApplicationApprovedYear.add(TagOPTION);
		TagOPTION = null;
	}
}

function frmScreen.chkSrcOfBusiness.onclick()
{
	if (m_blnAppEnqBusinessSourceRole == true) 
	{
		if (frmScreen.chkSrcOfBusiness.checked == true) 
		{
			scScreenFunctions.SetCollectionState(spnSrcOfBusinessDetails,"W");
			if (m_chkblnAppEnqBusinessSourceRole == false) 
			{
				<% /* MAR94  Maha T - Disable Introducer Search */ %>
				frmScreen.btnIntermediarySearch.disabled = true; 
				scScreenFunctions.SetCollectionState(spnIntermediaryDetails,"R");
				scScreenFunctions.SetFieldState(frmScreen, frmScreen.txtIntroducerId.id, "W");
				PopulateComboLists("Direct/Indirect",frmScreen.cboDirectIndirect);
				m_chkblnAppEnqBusinessSourceRole = true ;
			}
			else
			{
				<% /* MAR94  Maha T - Disable Introducer Search */ %>
				frmScreen.btnIntermediarySearch.disabled = true; 
				if ( frmScreen.cboDirectIndirect.value  == "1" )
				{
					<% /* MAR94  Maha T - Disable Introducer Search */ %>
					frmScreen.btnIntermediarySearch.disabled = true; 
					scScreenFunctions.SetCollectionState(spnIntermediaryDetails,"R");
					scScreenFunctions.SetFieldState(frmScreen, frmScreen.txtIntroducerId.id, "W");
				}
			}
		}
		else
		if (frmScreen.chkSrcOfBusiness.checked == false)
		{
			//
			// Reset Screen values
			//
			frmScreen.cboDirectIndirect.selectedIndex=-1;
			frmScreen.txtIntroducerId.value="";
			frmScreen.txtIntroducerName.value="";

	 		scScreenFunctions.SetCollectionState(spnSrcOfBusinessDetails,"R");
	 		scScreenFunctions.HideCollection(spnIntroducerName);
	 		frmScreen.btnIntermediarySearch.disabled = true; 
		}
	}
}

function frmScreen.btnSearch.onclick()
{
	// Validate input values ,If success then 
	// On click Button Search build an XML String to query 
	// Extra XML Nodes are built in to Request Tag to populate GN215.asp form Controls like ApplicationNumber , 
	// AccountNumber,UnitName ,UserName,channelName ,DepartmentName,LenderName,ProductName, ApplicationTypeName,
	// FromStageName,ToStageName,DateRange,IntermediaryName
	
	<% /* BMIDS607 Check that the Distribution Channel, Department, User and Unit are present */ %>
	if (MandatoryFields() == false) return 
	
	if (frmScreen.onsubmit() == false) return
	
	if (frmScreen.chkApplication.checked == true )  
	{
		if (ValidateStages()== false)
			return;
		if (ValidateDates() == false) 
			return;		
	}
	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"EXTENDEDSEARCH");
	
	XML.CreateTag("CHANNELID",frmScreen.cboDistributionChannel.value );
	XML.CreateTag("CHANNELNAME",frmScreen.cboDistributionChannel.options(frmScreen.cboDistributionChannel.selectedIndex).text);
	
	XML.CreateTag("DEPARTMENTID",frmScreen.cboDepartment.value);
	XML.CreateTag("DEPARTMENTNAME",frmScreen.cboDepartment.options(frmScreen.cboDepartment.selectedIndex).text);
	
	if (frmScreen.cboUnitName.value  != "<ALL>" ) 
	{ 	
		XML.CreateTag("UNITID",frmScreen.cboUnitName.value);
		XML.CreateTag("UNITNAME",frmScreen.cboUnitName.options(frmScreen.cboUnitName.selectedIndex).text);
	}
	
	if (frmScreen.cboUserName.value != "<ALL>" ) 
	{
		XML.CreateTag("USERID",frmScreen.cboUserName .value);
		XML.CreateTag("USERNAME",frmScreen.cboUserName.options(frmScreen.cboUserName.selectedIndex).text);
	}
	
	if (frmScreen.chkMortgageLenderDetails.checked == true )  
	{	
		XML.CreateTag("MORTGAGELENDERDETAILSCHECKED","1");	
		XML.CreateTag("LENDERNAME",frmScreen.cboLenderName.options(frmScreen.cboLenderName.selectedIndex).text);
		XML.CreateTag("ORGANISATIONID",frmScreen.cboLenderName.value);
		XML.CreateTag("MORTGAGEPRODUCTCODE",frmScreen.cboProductName.value);
		//DPF 02/10/2002 - AQR00555.
		XML.CreateTag("MORTGAGEPRODUCTNAME",frmScreen.cboProductName.options(frmScreen.cboProductName.selectedIndex).getAttribute("ProductName"));
		XML.CreateTag("MORTGAGEPRODUCTSTARTDATE",frmScreen.cboProductName.options(frmScreen.cboProductName.selectedIndex).getAttribute("StartDate"));
	}
	
	if (frmScreen.chkApplication.checked == true ) 
	{
		XML.CreateTag("APPLICATIONCHECKED","1");
		<% /* BS BM0499 17/04/03 */ %>
		if (frmScreen.cboApplicationType.value != "<ALL>" ) 
		{
			XML.CreateTag("APPLICATIONTYPE",frmScreen.cboApplicationType.value );
			XML.CreateTag("APPLICATIONTYPENAME",frmScreen.cboApplicationType.options(frmScreen.cboApplicationType.selectedIndex).text );
		}
		
		if (frmScreen.cboFromStage.value != "" && frmScreen.cboToStage.value != "" )
		{
			XML.CreateTag("APPLICATIONSTAGEFROM",frmScreen.cboFromStage.value );
			XML.CreateTag("APPLICATIONSTAGEFROMNAME",frmScreen.cboFromStage.options(frmScreen.cboFromStage.selectedIndex).text);
			XML.CreateTag("APPLICATIONSTAGETO",frmScreen.cboToStage.value);
			XML.CreateTag("APPLICATIONSTAGETONAME",frmScreen.cboToStage.options(frmScreen.cboToStage.selectedIndex).text);
		}
		
		if (frmScreen.txtDateFrom.value != "" && frmScreen.txtDateTo.value != "" ) 
		{
			XML.CreateTag("APPLICATIONDATEFROM",frmScreen.txtDateFrom.value);
			XML.CreateTag("APPLICATIONDATETO",frmScreen.txtDateTo.value);
		}
	}
	
	if (frmScreen.chkApplicationApproved.checked == true)  	
	{
		XML.CreateTag("APPLICATIONAPPROVEDCHECKED","1");
		XML.CreateTag("APPLICATIONAPPROVEDMONTH",frmScreen.cboApplicationApprovedMonth.value);
		XML.CreateTag("APPLICATIONAPPROVEDMONTHTEXT",frmScreen.cboApplicationApprovedMonth.options(frmScreen.cboApplicationApprovedMonth.selectedIndex).innerText);
		XML.CreateTag("APPLICATIONAPPROVEDYEAR",frmScreen.cboApplicationApprovedYear.value);
	}
		
	if (frmScreen.chkSrcOfBusiness.checked == true) 
	{	
		XML.CreateTag("SRCOFBUSINESSCHECKED","1");
		XML.CreateTag("INTERMEDIARYTYPE",frmScreen.cboDirectIndirect.value); 
		<% /* MO - 01/10/2002 - BM061/INWP1 - Need to search by direct indirect business now*/ %>
		XML.CreateTag("DIRECTINDIRECTBUSINESS",frmScreen.cboDirectIndirect.value); 
		
		if (frmScreen.txtIntroducerId.value != "") 
		{
			XML.CreateTag("INTRODUCERID",frmScreen.txtIntroducerId.value ); 
			XML.CreateTag("INTRODUCERNAME",frmScreen.txtIntroducerName.value  );
		}
		<% /* MO - 01/10/2002 - BM061/INWP1 - You no longer need to select an intermediary 
		else
		{
			alert("Select an Introducer");
			frmScreen.btnIntermediarySearch.focus(); 
			return ;
		} */ %>
	}
	
	if (frmScreen.ChkCancelledorDeclinedApplications.checked == true) 
	{
		XML.CreateTag("INCCANCELLEDORDECLINEDAPPSCHECKED","1")
		XML.CreateTag("INCLUDECANCELLEDAPPS","1");
		XML.CreateTag("INCLUDEDECLINEDAPPS","1");
	}	
	
	<% /* MAR46 Add search for Product Switch and Transfer of Equity completed applications*/ %>
	if (frmScreen.ChkProductSwitchApplications.checked == true) 
	{
		XML.CreateTag("INCPRODUCTSWITCHAPPSCHECKED","1")
	}	
	if (frmScreen.ChkTOEApplications.checked == true)   
	{
		XML.CreateTag("INCTOEAPPSCHECKED","1")
	}	
	scScreenFunctions.SetContextParameter(window,"idXML",XML.XMLDocument.xml);	
	scScreenFunctions.SetContextParameter(window,"idMetaAction","EXTENDEDSEARCH");
	frmTogn215.submit();
}

function frmScreen.btnIntermediarySearch.onclick()
{
	var ArrayArguments = new Array();
	
	<% /* MO - 01/10/2002 - BM061/INWP1 - Fixed these arguments cause they are wrong */ %>
	ArrayArguments[0] = m_UnitId;
	ArrayArguments[1] = m_UnitId;
	ArrayArguments[2] = m_MachineId;
	ArrayArguments[3] = m_DistributionChannelId;
	<% /* MO - 01/10/2002 - BMIDS00663 - Pass the direct/indirect validation type */ %>
	<% /* MO - 01/10/2002 - BM061/INWP1 - Pass the direct/indirect indicator
	ArrayArguments[4] = frmScreen.cboDirectIndirect.value;  */ %>
	ArrayArguments[4] = scScreenFunctions.GetComboValidationType(frmScreen, frmScreen.cboDirectIndirect.id);
	
	var sReturn = scScreenFunctions.DisplayPopup(window,document,"DC015Popup.asp",ArrayArguments,630,420);
	if (sReturn != null)
	{
		frmScreen.txtIntroducerId.value = sReturn[0];
		scScreenFunctions.SetFieldState(frmScreen, "txtIntroducerId", "W");
		frmScreen.txtIntroducerName.value = sReturn[1];
		scScreenFunctions.SetFieldState(frmScreen, "txtIntroducerName", "R");
		scScreenFunctions.ShowCollection(spnIntroducerName);
		m_sIntroducerId = sReturn[2];
		return;	
	}
}

function SetComboDefaultValue(ComboId,DefaultContextParamValueId,DefaultContextParamValuetext)
{
	// Function to set the Default Value .If there are elements more than 1 than loop thru 
	// else initiase with the contextparameters 
	var NumberOfElements = ComboId.length 
	if ( NumberOfElements > 0 ) 
	{
		ComboId.value = DefaultContextParamValueId; 
		//for (var i0 = 0 ; i0 < NumberOfElements ; i0++)
		//{
		//	if (ComboId.options(i0).text == DefaultContextParamValueId)
		//	{
		//		 ComboId.selectedIndex  =  i0;
				 return;
		//	}
		//}
	}
	else
	{
		var TagOPTION = document.createElement("OPTION");
		TagOPTION.value	= DefaultContextParamValueId;
		TagOPTION.text = DefaultContextParamValuetext;
		ComboId.add(TagOPTION);
		TagOPTION = null;	
	}
}

function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idXML","");
	frmWelcomeMenu.submit();
}

function ResetCombo(ComboId)
{
	// Function to Clear Combo Contents
	var nLength = ComboId.options.length;
	if ( nLength > 0 ) 
	{
		for(var iCount = nLength -1; iCount >= 0 ; iCount--)
		{
			ComboId.options.remove(iCount);
		}		
	}
}

function PopulateComboLists(ComboValueName,FormFieldId)
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array(ComboValueName);
	var bSuccess = false;

	if(XML.GetComboLists(document,sGroups))
	{
		bSuccess = true;
		bSuccess = bSuccess & XML.PopulateCombo(document,FormFieldId,ComboValueName,false);
	}
	XML= null;
}

function ValidateStages()
{
	if ( frmScreen.cboFromStage.value != "" ) 
	{		
		if (frmScreen.cboToStage.value == "" )
		{
			alert("Missing To Stage");
			frmScreen.cboToStage.focus();
			return false;
		}
	}
	
	if ( frmScreen.cboToStage.value != "") 
	{
		if (frmScreen.cboFromStage.value == "")
		{
			alert("Missing From Stage");
			frmScreen.cboFromStage.focus();
			return false;
		}
	}
	
	if (( frmScreen.cboToStage.value != "") && (frmScreen.cboFromStage.value != "") ) 
	{
		if ( parseInt(frmScreen.cboFromStage.value) > parseInt(frmScreen.cboToStage.value))
		{ 
			alert("The From Stage cannot be greater than the To Stage");
			frmScreen.cboFromStage.focus();  
			return false;
		}
	}
}
function ValidateDates()
{
	if (frmScreen.txtDateFrom.value != "" )  
	{
		// DateFrom cannot be Greater than sysdate (today)
		if ( frmScreen.txtDateFrom.value != "")
		{
			if (scScreenFunctions.CompareDateStringToSystemDate(frmScreen.txtDateFrom.value,">"))
			{
				alert("Date From cannot be greater than today ");
				frmScreen.txtDateFrom.focus();
				return false;
			}
		}

		if (frmScreen.txtDateTo.value == "" )
		{	
			alert("Missing To Date");
			frmScreen.txtDateTo.focus (); 			
			return false;
		}
	}
	
	if ( frmScreen.txtDateTo.value  != "" )
	{
		if (scScreenFunctions.CompareDateStringToSystemDate(frmScreen.txtDateTo.value ,">")) 
		{
			alert("The Date To cannot be greater than today's date");
			frmScreen.txtDateTo.focus();
			return false;
		}
	
		if (frmScreen.txtDateFrom.value == "")
		{
			alert("Missing From Date");
			frmScreen.txtDateFrom.focus (); 			
			return false;
		}
	}
	
	if ((frmScreen.txtDateFrom.value != "") && (frmScreen.txtDateTo.value != "" ))
	{
		if (scScreenFunctions.CompareDateFields(frmScreen.txtDateFrom,">",frmScreen.txtDateTo))
		{
			alert("The Date From cannot be greater than Date To");
			frmScreen.txtDateFrom.focus();
			return false;
		}
	}
}

<% /*  BMIDS607 Add function to check that Distribution Channel, Department, User and Unit are present */ %>
function MandatoryFields()
{
	if ( frmScreen.cboDistributionChannel.value == "" ) 
	{		
		alert("Missing Distribution Channel");
		frmScreen.cboDistributionChannel.focus();
		return false;
	}
	
	if ( frmScreen.cboDepartment.value == "" ) 
	{		
		alert("Missing Department");
		frmScreen.cboDepartment.focus();
		return false;
	}
	
	if ( frmScreen.cboUnitName.value == "" ) 
	{		
		alert("Missing Unit Name");
		frmScreen.cboUnitName.focus();
		return false;
	}
	
	if ( frmScreen.cboUserName.value == "" ) 
	{		
		alert("Missing User Name");
		frmScreen.cboUserName.focus();
		return false;
	}
}

function PopulateScreenValues(sXML)
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(sXML);
	if (XML.SelectTag(null,"REQUEST") != null)
	{
		// APS 09/01/2001 : Added in functionality to reselect combo box values
		// for distribution channel, department, unit and user
		m_DistributionChannelId = XML.GetTagText("CHANNELID");
		m_DepartmentId = XML.GetTagText("DEPARTMENTID");
		m_UnitId = XML.GetTagText("UNITID");
		if (m_UnitId == "") m_UnitId = "<ALL>";
		m_UserId = XML.GetTagText("USERID");		
		if (m_UserId == "") m_UserId = "<ALL>";
		if (XML.GetTagText("MORTGAGELENDERDETAILSCHECKED") == "1" ) 
		{
			frmScreen.chkMortgageLenderDetails.checked = true ; 
			scScreenFunctions.SetCollectionState(spnLenderProductDetails,"W");
			PopulateLenderNamesCombo();
			frmScreen.cboLenderName.value = XML.GetTagText("ORGANISATIONID");
			PopulateProductNamesCombo();
			frmScreen.cboProductName.value = XML.GetTagText("MORTGAGEPRODUCTCODE");
		}
		
		if (XML.GetTagText("APPLICATIONCHECKED") == "1")
		{
			frmScreen.chkApplication.checked = true;
			scScreenFunctions.SetCollectionState(spnApplicationDetails,"W");
			PopulateComboLists("TypeOfMortgage",frmScreen.cboApplicationType );
			<% /* BS BM0499 17/04/03 */ %>
			var TagOPTION = document.createElement("OPTION");
			TagOPTION.value = "<ALL>"
			TagOPTION.text  = "<ALL>";
			frmScreen.cboApplicationType.add(TagOPTION);
			TagOPTION = null;
			//frmScreen.cboApplicationType.value = XML.GetTagText("APPLICATIONTYPE"); 
			m_AppType = XML.GetTagText("APPLICATIONTYPE"); 
			if (m_AppType == "") m_AppType = "<ALL>";
			frmScreen.cboApplicationType.value = m_AppType; 
			<% /* BS BM0499 End 17/04/03 */ %>
			PopulateApplicationStages();
			frmScreen.cboFromStage.value = XML.GetTagText("APPLICATIONSTAGEFROM");
			frmScreen.cboToStage.value = XML.GetTagText("APPLICATIONSTAGETO");
			frmScreen.txtDateFrom.value = XML.GetTagText("APPLICATIONDATEFROM");
			frmScreen.txtDateTo.value = XML.GetTagText("APPLICATIONDATETO"); 
		}
		
		if (XML.GetTagText("APPLICATIONAPPROVEDCHECKED") == "1" ) 
		{
			frmScreen.chkApplicationApproved.checked = true;
			scScreenFunctions.SetCollectionState(spnApplicationApprovedDetails,"W");
			PopulateComboLists("MonthsInYear",frmScreen.cboApplicationApprovedMonth);
			frmScreen.cboApplicationApprovedMonth.value = XML.GetTagText("APPLICATIONAPPROVEDMONTH");
			PopulateApplicationApporvalYears();
			frmScreen.cboApplicationApprovedYear.value = XML.GetTagText("APPLICATIONAPPROVEDYEAR"); 
		}
		
		if ( XML.GetTagText("SRCOFBUSINESSCHECKED") == "1")
		{
			scScreenFunctions.SetCollectionState(spnIntermediaryType,"W");
			frmScreen.chkSrcOfBusiness.checked = true ;
			PopulateComboLists("Direct/Indirect",frmScreen.cboDirectIndirect);	
			frmScreen.cboDirectIndirect.value = XML.GetTagText("INTERMEDIARYTYPE");
			scScreenFunctions.SetCollectionState(spnIntermediaryDetails,"R");
			scScreenFunctions.SetFieldState(frmScreen, frmScreen.txtIntroducerId.id, "W");
			frmScreen.txtIntroducerId.value = XML.GetTagText("INTRODUCERID"); 
			frmScreen.txtIntroducerName.value = XML.GetTagText("INTRODUCERNAME");
			if (frmScreen.txtIntroducerName.value.length > 0) {
				scScreenFunctions.ShowCollection(spnIntroducerName);
			} else {
				scScreenFunctions.HideCollection(spnIntroducerName);
			}
			<% /* MAR94  Maha T - Disable Introducer Search */ %>
			frmScreen.btnIntermediarySearch.disabled = true;
		}
		
		if (XML.GetTagText("INCCANCELLEDORDECLINEDAPPSCHECKED") == "1" )
		{
			frmScreen.ChkCancelledorDeclinedApplications.checked = true ; 
		}	
	}
}

-->
</script>
</body>
</html>
