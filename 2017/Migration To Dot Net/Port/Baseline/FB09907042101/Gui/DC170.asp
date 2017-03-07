<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*  
Workfile:      DC170.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Employment Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		07/02/2000	Created
AD		15/03/2000	Incorporated third party include files.
AD		21/03/2000	SYS0423 - Customer Name added
AY		30/03/00	New top menu/scScreenFunctions change
AY		05/04/00	SYS0582 - XML Response not being checked and
					GetTagText calls outside appropriate if statement
IVW		11/04/2000	Changed Prev/Next/Cancel to Ok/Cancel
BG		12/04/2000	SYS0521 - Text box now appears when user selects "Other" from cboEmploymentStatus
BG		12/04/2000	SYS0616 - Changed text from Main Employment to read Current Main Employment.
IVW		20/04/2000	SYS0612	- Switches Mandatory on and off for Trade/Company Name
IW		27/04/00	SYS0504 - Only allow 1 main current employment type to exist.
MH      17/05/00    SYS0750 - Date validation return value was being ignored +
					 Postcode validation on save + Current Main Employment enablement
BG		17/05/00	SYS0752 Removed Tooltips
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
MC		23/05/00	SYS0756 If Read-only mode, disabled fields
MC		31/05/00	SYS0754	Ensure Submit not actioned more than once
BG		30/08/00	SYS1224	stop error message appearing about main employment if info is changed.
BG		08/09/00	SYS1063	disable add to Directory button if Retired or Not employed.
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
CL		05/03/01	SYS1920 Read only functionality added
TJ		30/03/01	SYS2050 Critical Data functionality added
SA		23/04/01	SYS1947 Main Employment Indicator Validation altered.
GD		11/05/01	SYS2050 Critical Data functionality ROLLED BACK
SA		16/05/01	SYS1947	Main Employment Indicator Validation altered. Date Left field disabled in Edit mode
DC      20/07/01    SYS2038 Critical Data functionality ROLLED FORWARD AGAIN
JR		14/09/01	Omiplus24 - Telephone Number change in populate screen
MDC		01/10/01	SYS2785 Enable client versions to override labels
JLD		10/12/01	SYS2806 Completeness check routing

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		15/05/2002	BMIDS00008	Removed cboIndustryType functionality
MV		17/06/2002	BMIDS00008	Removed cboIndustryType functionality in window.OnLoad() and PopulateCombos();
DPF		20/06/02	BMIDS00077	Changes made to file to bring in line with Core V7.0.2, changes are...
					SYS4727 Use cached versions of frame functions
GD		25/06/02	BMIDS00077 - Problems fixed with mismatching { }
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MV		16/08/2002	BMIDS00163  Core SYS5025 Previous Employment should only require DC170 to be completed 
								Amended btnSubmit.Onclick()
MV		22/08/2002	BMIDS00355	IE 5.5 upgrade Errors - Modified the HTML layout								
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		11/10/2002	BMIDS00272	Amended frmScreen.cboEmploymentStatus.onclick()
MV		17/10/2002	BMIDS00272	Amended frmScreen.cboEmploymentStatus.onclick()
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
SA		17/11/2002	BMIDS00934 - Disable screen if data freeze indicator set.
MO		20/11/2002	BMIDS01007 - Disable Date left if employment status is not employed
MDC		29/11/2002	BMIDS01110 - Add ScreenRules wrapper around SaveEmployment.
AW		04/12/02	BM00152		 Do not route to GN300	
JR		23/01/2003	BM0271		Check for ProcessingInd.	
BS		12/06/2003	BM0521		Disable Contact Details on ThirdPartyDetails.asp
KRW     25/09/03    BM00063     Corrections to screen alignment
HMA     26/09/03    BMIDS639    Allow for changing Employment Status using keyboard
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
MF		03/08/2005	MAR20		WP06 clicking Ok now routes to DC190 Income summary
JJ		10/10/2005	MAR119		Validating Flat no., House no. and House name.
								PostCode,Town,Street made mandatory fields.
PJO     10/22/2005  MAR483      Date Started in madatory - set red									
PE		23/02/2006	MAR1313		Unable to enter income details when employment status "CONTRACT" selected
PE		28/03/2006	MAR1395		Make 'Employer' and 'Date Started' not required for employment status types (Not Earning,Homemaker,Student,Retired,Other)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
pct		07/03/2006	EP8			MAR20 Reversed - now routs to DC181
PB		06/07/2006	EP543		Populate 'Title-Other' field from XML
MAH		16/11/2006  E2CR35		Add NatureOfBusiness
GHun	19/01/2007	EP2_974		Merged EP1230 - need to populate m_sEmploymentStatusType for use in Context (used in DC192)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


*/ %>
<HEAD>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
data=scClientFunctions.asp width=1 height=1 type=text/x-scriptlet 
VIEWASTEXT></OBJECT>
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<% /* removed as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scMathFunctions.asp height=1 id=scMathFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>

FORMS */ %>
<form id="frmToDC160" method="post" action="DC160.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC181" method="post" action="DC181.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC182" method="post" action="DC182.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC183" method="post" action="DC183.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC190" method="post" action="DC190.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange" year4>
<div style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 40px; HEIGHT: 300px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Employment Details</strong> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel"> 
		Customer Name
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="txtCustomerName" style="WIDTH: 200px" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Employment Status
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<select id="cboEmploymentStatus" style="WIDTH: 200px" class="msgCombo"></select>
			<span id="spnOtherEmploymentStatus" style="LEFT: 205px; POSITION: absolute; TOP: 0px">
				<input id="txtOtherEmploymentStatus" maxlength="30" style="VISIBILITY: hidden; WIDTH: 150px; POSITION: absolute" class="msgTxt">
			</span>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 84px" class="msgLabel">
		Employment Type
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<select id="cboEmploymentType" style="WIDTH: 200px" class="msgCombo"></select>
		</span>
	</span>
	
	<span style="LEFT: 4px; COLOR: red; POSITION: absolute; TOP: 108px" class="msgLabel">
		Employer/Trading Name
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyName" maxlength="45" style="WIDTH: 200px; POSITION: absolute" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 350px; POSITION: absolute; TOP: 103px" class="msgLabel">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton" onclick="btnDirectorySearch.onClick()">
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 132px" class="msgLabel">
		Is in the directory?
		<span style="LEFT: 140px; WIDTH: 60px; POSITION: absolute; TOP: -3px">
			<input id="optInDirectoryYes" name="InDirectoryGroup" type="radio" value="1" onclick="InDirectoryChanged()">
			<label for="optInDirectoryYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 194px; WIDTH: 60px; POSITION: absolute; TOP: -3px">
			<input id="optInDirectoryNo" name="InDirectoryGroup" type="radio" value="0" onclick="InDirectoryChanged()">
			<label for="optInDirectoryNo" class="msgLabel">No</label> 
		</span> 
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 156px" class="msgLabel">
		Industry Type
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<select id="cboIndustryType" style="WIDTH: 200px" class="msgCombo" NAME="cboIndustryType"></select>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 180px" class="msgLabel">
		Occupation Type
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<select id="cboOccupationType" style="WIDTH: 200px" class="msgCombo"></select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 204px" class="msgLabel">
		Job Title
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="txtJobTitle" maxlength="50" style="WIDTH: 200px; POSITION: absolute" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 224px" class="msgLabel">
		Current Main Employment?
		<span style="LEFT: 10px; POSITION: relative; TOP: 1px">
			<input id="optMainEmploymentYes" name="MainEmploymentGroup" type="radio" value="1" onclick="MainEmploymentChanged()">
			<label for="optMainEmploymentYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 17px; POSITION: relative; TOP: 1px">
			<input id="optMainEmploymentNo" name="MainEmploymentGroup" type="radio" value="0" onclick="MainEmploymentChanged()">
			<label for="optMainEmploymentNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; COLOR: red; POSITION: absolute; TOP: 252px" class="msgLabel">
		Date Started/Established
		<span style="LEFT: 140px; COLOR: red; POSITION: absolute; TOP: -3px">
			<input id="txtDateStartedOrEstablished" maxlength="10" style="WIDTH: 100px; POSITION: absolute" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 276px" class="msgLabel">
		Date Left/Ceased Trading
		<span style="LEFT: 140px; POSITION: absolute; TOP: -3px">
			<input id="txtDateLeftOrCeasedTrading" maxlength="10" style="WIDTH: 100px; POSITION: absolute" class="msgTxt">
		</span> 
	</span>
	
</div>

<div id="divThirdPartyDetails" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 346px; HEIGHT: 232px" class="msgGroup"><!-- #include FILE="includes/thirdpartydetails.htm" -->
</div> 
</form>

<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 580px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- #include FILE="attribs/DC170attribs.asp" --><!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sCustomerName = "";
var m_sEmploymentSequenceNumber = "";
var EmploymentXML = null;
var m_bAllowMainEmployment = true;
var ComboXML = null;
var m_sEmploymentStatusType = "";
var scScreenFunctions;
var m_bIsSubmit = false;
var m_bEmploymentRecordsExist = false;
var m_blnReadOnly = false;
var m_sProcessInd = ""; //JR BM0271

/* EVENTS */

function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
	scScreenFunctions.SetContextParameter(window,"idEmploymentStatusType", "");
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmToDC160.submit();
}

function btnSubmit.onclick()
{
	if(m_bIsSubmit)
		return;
		
	m_bIsSubmit = true;
	
	//START: (MAR119) - New code added by Joyce Joseph on 10-Oct-2005
	//Validating Flat no., House no. and House name.
	if (ThirdPartyDetailsScreenRules() == 2)
	{
		m_bIsSubmit = false;
		return;
	}
	//END: (MAR119)
	if (CommitChanges())
	{
		<% /* clear the contexts */ %>
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		scScreenFunctions.SetContextParameter(window,"idEmployerName", frmScreen.txtCompanyName.value);
		scScreenFunctions.SetContextParameter(window,"idEmploymentStatusType", m_sEmploymentStatusType);
		
		// AW	04/12/02	BM00152 -	Start
		//if(scScreenFunctions.CompletenessCheckRouting(window))
		//	frmToGN300.submit();
		
        <%/* MV - 16/08/2002 - BMIDS00163 - Core SYS5025 Previous Employment shuold only require DC170 to be completed */%>
        //else
        // AW	04/12/02	BM00152 -	End 
        if (frmScreen.txtDateLeftOrCeasedTrading.value.length > 0)
            frmToDC160.submit();
			
		else
		
			<% //MAR1313 %>
			<% //Peter Edney - 23/02/2006 %>
			var bEmploymentStatusTypeE = scScreenFunctions.IsValidationType(frmScreen.cboEmploymentStatus,"E");
			var bEmploymentStatusTypeS = scScreenFunctions.IsValidationType(frmScreen.cboEmploymentStatus,"S");
			var bEmploymentStatusTypeC = scScreenFunctions.IsValidationType(frmScreen.cboEmploymentStatus,"C");						
			switch (true)
			{			
				case bEmploymentStatusTypeE:
					<% /* MF MARS20 //frmToDC181.submit(); */ %>
					<% /* pct EP8 //frmToDC190.submit(); */ %>
					frmToDC181.submit();

					break;
				case bEmploymentStatusTypeS:
					frmToDC182.submit();
					break;
				case bEmploymentStatusTypeC:
					frmToDC183.submit();
					break;
				default:
					frmToDC160.submit();
					break;					
			}
					
	}
	else
		m_bIsSubmit = false;
}

function frmScreen.optMainEmploymentYes.onclick()
{
	<% /* MO 20/11/2002 BMIDS01007 */ %>
	<% /* if ((m_sEmploymentStatusType != "N") && (m_sEmploymentStatusType != "R")) { (MAR1395) */ %> 
	if (!GetEmpValidation("N") && !GetEmpValidation("R")) {
		if (frmScreen.optMainEmploymentYes.checked)
		{
			frmScreen.txtDateLeftOrCeasedTrading.value = "";
			scScreenFunctions.SetFieldState(frmScreen, "txtDateLeftOrCeasedTrading", "D");
		}
	}
}

function frmScreen.optMainEmploymentNo.onclick()
{
	<% /* MO 20/11/2002 BMIDS01007 */ %>
	<% /* if ((m_sEmploymentStatusType != "N") && (m_sEmploymentStatusType != "R")) { (MAR1395) */ %>
	if (!GetEmpValidation("N") && !GetEmpValidation("R")) {
		if (frmScreen.optMainEmploymentNo.checked) {
			scScreenFunctions.SetFieldState(frmScreen, "txtDateLeftOrCeasedTrading", "W");
		}
	}
}

<% /* BMIDS639 Use onchange event so that processing is done whenever Employment Status is changed */ %>
function frmScreen.cboEmploymentStatus.onchange()
{
	<% //m_sEmploymentStatusType = scScreenFunctions.GetComboValidationType(frmScreen,"cboEmploymentStatus"); (MAR1395) %>
	<%// EP1230 - need to populate m_sEmploymentStatusType for use in Context %>
	if (scScreenFunctions.IsValidationType(frmScreen.cboEmploymentStatus, "CON"))
		m_sEmploymentStatusType = "C";
	else
	  if (scScreenFunctions.IsValidationType(frmScreen.cboEmploymentStatus, "SELF"))
	     m_sEmploymentStatusType = "S";
	
	<%// EP1230 - End %>

	PopulateEmploymentTypes();
	
	//if ((m_sEmploymentStatusType == "N") | (m_sEmploymentStatusType == "R"))
	if(GetEmpValidation("N") || GetEmpValidation("R"))
	{
		<%/* Disable almost all fields related fields */%>
		frmScreen.txtCompanyName.value = "";
		frmScreen.cboOccupationType.selectedIndex = 0;
		frmScreen.cboIndustryType.selectedIndex = "";
		frmScreen.txtJobTitle.value = "";
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCompanyName");
		frmScreen.txtCompanyName.disabled = true;
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboOccupationType");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboIndustryType");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtJobTitle");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboEmploymentType");
		scScreenFunctions.ClearCollection(divThirdPartyDetails);
		scScreenFunctions.SetCollectionToReadOnly(divThirdPartyDetails);
		frmScreen.txtCompanyName.removeAttribute("required");
		frmScreen.txtDateStartedOrEstablished.removeAttribute("required");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtDateStartedOrEstablished");
		<% /* MO 20/11/2002 BMIDS01007 */ %>
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtDateLeftOrCeasedTrading");
	}
	else
	{
		frmScreen.txtCompanyName.setAttribute("required", "true");
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtCompanyName");
		scScreenFunctions.SetCollectionToWritable(divThirdPartyDetails);
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtCompanyName");
		frmScreen.txtCompanyName.disabled = false;
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboOccupationType");
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboIndustryType");
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtJobTitle");
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboEmploymentType");
		scScreenFunctions.SetRadioGroupToWritable(frmScreen,"MainEmploymentGroup");
		frmScreen.txtDateStartedOrEstablished.setAttribute("required", "true");
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtDateStartedOrEstablished");
		<% /* MO 20/11/2002 BMIDS01007 */ %>
		if (frmScreen.optMainEmploymentNo.checked) {
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtDateLeftOrCeasedTrading");
		}
		
	}
	
	<% /* if (m_sEmploymentStatusType == "O") (MAR1395) */ %>
	if(GetEmpValidation("O"))
		scScreenFunctions.ShowCollection(spnOtherEmploymentStatus);
	else
		scScreenFunctions.HideCollection(spnOtherEmploymentStatus);	
			
	<% // MAR1395 %>
	if(GetEmpValidation("NOEMPLOYER"))
	{
		frmScreen.txtCompanyName.removeAttribute("required");	    	    
		frmScreen.txtCompanyName.parentElement.parentElement.style.color = "";		
		frmScreen.txtDateStartedOrEstablished.removeAttribute("required");	    	    
		frmScreen.txtDateStartedOrEstablished.parentElement.parentElement.style.color = "";		
	}
	else
	{
		frmScreen.txtCompanyName.setAttribute("required","true");	    
		frmScreen.txtCompanyName.parentElement.parentElement.style.color = "red";					
		frmScreen.txtDateStartedOrEstablished.setAttribute("required","true");	    
		frmScreen.txtDateStartedOrEstablished.parentElement.parentElement.style.color = "red";			
	}
			
	SetAvailableFunctionality();
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	<%/* Parameters for thirdpartydetails.asp */%>
	m_fSetAvailableFunctionalityOverride = SetAvailableFunctionalityOverride;
	m_sThirdPartyType = "5";

	var sGroups = new Array("Country","EmploymentStatus","EmploymentType","OccupationType","NatureOfBusiness");
	objDerivedOperations = new DerivedScreen(sGroups);

	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	FW030SetTitles("Add/Edit Employment Details","DC170",scScreenFunctions);

	RetrieveContextData();
	SetThirdPartyDetailsMasks();

	frmScreen.txtCompanyName.removeAttribute("required");
	SetMasks();
	
	ThirdPartyCustomise();

	Validation_Init();	
	Initialise(true);
	
	if(m_sMetaAction == "Add") 
	{
		if(m_bAllowMainEmployment == false)
			frmScreen.optMainEmploymentNo.checked=true;
		else
		{
			frmScreen.optMainEmploymentYes.checked=true;
			scScreenFunctions.SetFieldState(frmScreen, "txtDateLeftOrCeasedTrading", "D");
		}
	}

	<% /* BS BM0521 12/06/03 
	if(m_sReadOnly == "1" || m_sProcessInd == "0") //JR BM0271 - add ProcessInd check.
		SetScreenToReadOnly();
	else
		scScreenFunctions.SetFocusToFirstField(frmScreen);*/ %>
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC170");
	
	<% /* BS BM0521 12/06/03 */ %>
	if (m_blnReadOnly)
	{
		m_sReadOnly = "1";
		SetScreenToReadOnly();
	}
	else
	{
		m_sReadOnly = "0";
		scScreenFunctions.SetFocusToFirstField(frmScreen);
	}	
	<% /* BS BM0521 End 12/06/03 */ %>
	
	if(m_sMetaAction == "Edit")
	{
		if(frmScreen.optMainEmploymentYes.checked==true)
			scScreenFunctions.SetFieldState(frmScreen, "txtDateLeftOrCeasedTrading", "D");
	}

	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

function SetScreenToReadOnly()
{
	<% /* BS BM0521 12/06/03 Commented out to avoid duplication
	scScreenFunctions.SetScreenToReadOnly(frmScreen);*/ %>

	frmScreen.btnDirectorySearch.disabled = true;
	frmScreen.btnClear.disabled = true;
	frmScreen.btnPAFSearch.disabled = true;
	frmScreen.btnAddToDirectory.disabled = true;
}


function CommitChanges()
{
	function ValidateDates()
	{
		<% /* MO - BMIDS00807 */ %>
		var dteToday = scScreenFunctions.GetAppServerDate(true);
		<% /* var dteToday = new Date(); */ %>
		var dtActiveFrom = scScreenFunctions.GetDateObject(frmScreen.txtDateStartedOrEstablished);
		var dtActiveTo = scScreenFunctions.GetDateObject(frmScreen.txtDateLeftOrCeasedTrading);
		<%/* Date started must not be in the future or more than a 100 years in the past */%>
		if (dtActiveFrom != null)
			if (scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtDateStartedOrEstablished,">"))
			{
				alert("Date started cannot be in the future.");
				frmScreen.txtDateStartedOrEstablished.focus();
				return(false);
			}
			//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
			//else if (scMathFunctions.GetYearsBetweenDates(dtActiveFrom,dteToday) >= 100)
			else if (top.frames[1].document.all.scMathFunctions.GetYearsBetweenDates(dtActiveFrom,dteToday) >= 100)
			{
				alert("Date started cannot more than 100 years in the past.");
				frmScreen.txtDateStartedOrEstablished.focus();
				return(false);
			}

		<%/* Date left must not be in the future or more than a 100 years in the past */%>
		if (dtActiveTo != null)
			if (scScreenFunctions.CompareDateFieldToSystemDate(frmScreen.txtDateLeftOrCeasedTrading,">"))
			{
				alert("Date left cannot be in the future.");
				frmScreen.txtDateLeftOrCeasedTrading.focus();
				return(false);
			}
			//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
			//else if (scMathFunctions.GetYearsBetweenDates(dtActiveTo,dteToday) >= 100)
			else if (top.frames[1].document.all.scMathFunctions.GetYearsBetweenDates(dtActiveTo,dteToday) >= 100)
			
			{
				alert("Date left cannot more than 100 years in the past.");
				frmScreen.txtDateLeftOrCeasedTrading.focus();
				return(false);
			}

		<%/* The moved in date cannot be after the moved out date */%>
		if ((dtActiveFrom != null) & (dtActiveTo != null))
			if (scScreenFunctions.CompareDateFields(frmScreen.txtDateStartedOrEstablished,">",frmScreen.txtDateLeftOrCeasedTrading))
			{
				alert("Date started cannot be after the date left.");
				frmScreen.txtDateStartedOrEstablished.focus();
				return(false);
			}
		return(true);
	}

	var bSuccess = true;
	var bSaveEmploymentDetails = true;

	if (m_sReadOnly != "1")
	{
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
			{
				<% /*  SYS0612 Must do mandatory processing on Employer/Trading name
				 if it is required. This is a code around because for certain 
				 employment types the field is none mandatory and made read only,
				 however the validation still requires read only mandatory fields
				 to be filled. */ %>
				 		
				bSuccess = false;
	
				if (MainEmploymentCheck())
					if (ValidateDates())
						if (scScreenFunctions.ValidatePostcode(frmScreen.txtPostcode))
							bSuccess = SaveEmploymentDetails()
			}
		}
		else
			bSuccess = false;
	}		
	return(bSuccess);
}

<% /*  Inserts default values into all fields */ %>
function DefaultFields()
{
	ClearFields(true, true);
	with (frmScreen)
	{
		txtFlatNumber.value = "";
		txtHouseName.value = "";
		txtHouseNumber.value = "";
		txtPostcode.value = "";
		txtStreet.value = "";

		objDerivedOperations.ClearCounty();
		
		txtDistrict.value = "";
		txtTown.value = "";
		cboCountry.selectedIndex = 0;
		frmScreen.optMainEmploymentNo.checked=true;
	}
}

<% /*  Initialises the screen */ %>
function Initialise(bOnLoad)
{
	if(bOnLoad == true)
		PopulateCombos();

	frmScreen.txtCustomerName.value = m_sCustomerName;

	if(m_sMetaAction == "Edit")
	{
		PopulateScreen();
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboEmploymentStatus");
	}
	else
	{
		DefaultFields();
		frmScreen.cboEmploymentStatus.onchange();   //BMIDS639
	}

	SetAvailableFunctionality();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function MainEmploymentChanged()
{
	if ((scScreenFunctions.GetRadioGroupValue(frmScreen,"MainEmploymentGroup") == "1") & !m_bAllowMainEmployment)
	{
		alert("A main employment already exists for this customer.");
		scScreenFunctions.SetRadioGroupValue(frmScreen,"MainEmploymentGroup","0");
	}
}

<% /*  Populates all combos on the screen */ %>
function PopulateCombos()
{
	PopulateTPTitleCombo();

	var sControlList = new Array(frmScreen.cboCountry, frmScreen.cboEmploymentStatus, frmScreen.cboEmploymentType,frmScreen.cboOccupationType,frmScreen.cboIndustryType);
	objDerivedOperations.GetComboLists(sControlList);

	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//ComboXML = new scXMLFunctions.XMLObject();
	ComboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("Country","EmploymentStatus","EmploymentType","OccupationType","NatureOfBusiness");
	ComboXML.GetComboLists(document,sGroupList);
	<%/*MAH 30/11/2006 E2CR35*/%>
	ComboXML.PopulateCombo(document,frmScreen.cboIndustryType,"NatureOfBusiness",true);
}

function PopulateEmploymentTypes()
{
	if (ComboXML.PopulateCombo(document,frmScreen.cboEmploymentType,"EmploymentType",true))
	{
		optOption = null;
		for (var iCount = frmScreen.cboEmploymentType.length - 1; iCount > -1; iCount--)
		{
			optOption = frmScreen.cboEmploymentType.item(iCount);

			<% /* if (((m_sEmploymentStatusType == "N") || (m_sEmploymentStatusType == "R")) ^ (MAR1395) */ %>
			if ((GetEmpValidation("N") || GetEmpValidation("R")) ^
				scScreenFunctions.IsOptionValidationType(frmScreen.cboEmploymentType, iCount, "NA"))
				frmScreen.cboEmploymentType.remove(iCount);
		}
	}
}

function PopulateScreen()
{
	<%/* Populates the screen with details of the item selected in dc160 
	next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077*/ %>
	//EmploymentXML = new scXMLFunctions.XMLObject();
	EmploymentXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	EmploymentXML.CreateRequestTag(window,null);
	EmploymentXML.CreateActiveTag("EMPLOYMENT");
	EmploymentXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	EmploymentXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	EmploymentXML.CreateTag("EMPLOYMENTSEQUENCENUMBER", m_sEmploymentSequenceNumber);

	EmploymentXML.RunASP(document,"GetEmploymentDetails.asp");

	if(!EmploymentXML.IsResponseOK()) return;

	if(EmploymentXML.SelectTag(null, "EMPLOYMENT") != null)
	{
		with (frmScreen)
		{
			txtCompanyName.value = EmploymentXML.GetTagText("COMPANYNAME");
			cboEmploymentStatus.value = EmploymentXML.GetTagText("EMPLOYMENTSTATUS");
			txtOtherEmploymentStatus.value = EmploymentXML.GetTagText("OTHEREMPLOYMENTSTATUS");
			frmScreen.cboEmploymentStatus.onchange();  //BMIDS639
			cboEmploymentType.value = EmploymentXML.GetTagText("EMPLOYMENTTYPE");
			cboOccupationType.value = EmploymentXML.GetTagText("OCCUPATIONTYPE");
			cboIndustryType.value = EmploymentXML.GetTagText("INDUSTRYTYPE");
			txtJobTitle.value = EmploymentXML.GetTagText("JOBTITLE");
			scScreenFunctions.SetRadioGroupValue(frmScreen,"MainEmploymentGroup",EmploymentXML.GetTagText("MAINSTATUS"));
			txtDateStartedOrEstablished.value = EmploymentXML.GetTagText("DATESTARTEDORESTABLISHED");
			txtDateLeftOrCeasedTrading.value = EmploymentXML.GetTagText("DATELEFTORCEASEDTRADING");
			txtContactForename.value = EmploymentXML.GetTagText("CONTACTFORENAME");
			txtContactSurname.value = EmploymentXML.GetTagText("CONTACTSURNAME");
			cboTitle.value = EmploymentXML.GetTagText("CONTACTTITLE");
			<% /* PB 07/06/2007 EP543 Begin */ %>
			checkOtherTitleField();
			txtTitleOther.value = EmploymentXML.GetTagText("CONTACTTITLEOTHER");
			<% /* EP543 End */ %>

			objDerivedOperations.LoadCounty(EmploymentXML);
			
			txtDistrict.value = EmploymentXML.GetTagText("DISTRICT");
			txtFlatNumber.value = EmploymentXML.GetTagText("FLATNUMBER");
			txtHouseName.value = EmploymentXML.GetTagText("BUILDINGORHOUSENAME");
			txtHouseNumber.value = EmploymentXML.GetTagText("BUILDINGORHOUSENUMBER");
			txtPostcode.value = EmploymentXML.GetTagText("POSTCODE");
			txtStreet.value = EmploymentXML.GetTagText("STREET");
			txtTown.value = EmploymentXML.GetTagText("TOWN");
			cboCountry.value = EmploymentXML.GetTagText("COUNTRY");
		}

		if (EmploymentXML.GetTagText("MAINSTATUS") == "1") 
			m_bAllowMainEmployment = true;

		m_sDirectoryGUID = EmploymentXML.GetTagText("DIRECTORYGUID");
		m_sThirdPartyGUID = EmploymentXML.GetTagText("THIRDPARTYGUID");
		m_bDirectoryAddress = (m_sDirectoryGUID != "");
		
		var TempXML = EmploymentXML.ActiveTag;
		var ContactXML = EmploymentXML.SelectTag(null, "CONTACTDETAILS");
		if(ContactXML != null)
			m_sXMLContact = ContactXML.xml;
		EmploymentXML.ActiveTag = TempXML; 
	}
}

function RetrieveContextData()
{
	/*Debug
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber","02013118");
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber","1");
	scScreenFunctions.SetContextParameter(window,"idEmploymentSequenceNumber","1");
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","00017884");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
	scScreenFunctions.SetContextParameter(window,"idApplicationPriority","10");
	scScreenFunctions.SetContextParameter(window,"idStageId","40");
	End */

	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	<% /* BS BM0521 12/06/03  
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	//BMIDS00934 Need to check datafreeze indicator too.
	if (m_sReadOnly == "0")
		m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idFreezeDataIndicator","0");*/ %>
	
	
	m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","1325");
	m_sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + scScreenFunctions.GetContextParameter(window,"idCustomerIndex","1"),"");
	m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","1");
	m_bAllowMainEmployment = (parseInt(scScreenFunctions.GetContextParameter(window,"idMainEmploymentCount","1")) == 0);
	m_bEmploymentRecordsExist = (parseInt(scScreenFunctions.GetContextParameter(window,"idEmploymentCount","1")) == 1);	
	m_sProcessInd = scScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null); //JR BM0271
	
	if (m_sMetaAction == "Edit")
		m_sEmploymentSequenceNumber = scScreenFunctions.GetContextParameter(window,"idEmploymentSequenceNumber","1");    

	var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bCanAddToDirectory = (sUserRole >= XML.GetGlobalParameterAmount(document,"AddToDirectoryRole"));

	XML = null;
}


function SaveEmploymentDetails()
{
	var bSuccess = true;
	//next line replaced by line below as per Core V7.0.2
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,null)
	XML.SetAttribute("OPERATION","CriticalDataCheck");
	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");

	var tagEmploymentDetails = XML.CreateActiveTag("EMPLOYMENT");
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	XML.CreateTag("DATELEFTORCEASEDTRADING",frmScreen.txtDateLeftOrCeasedTrading.value);
	XML.CreateTag("DATESTARTEDORESTABLISHED",frmScreen.txtDateStartedOrEstablished.value);
	XML.CreateTag("EMPLOYMENTSTATUS",frmScreen.cboEmploymentStatus.value);
	XML.CreateTag("INDUSTRYTYPE",frmScreen.cboIndustryType.value);
	XML.CreateTag("OTHEREMPLOYMENTSTATUS",frmScreen.txtOtherEmploymentStatus.value);
	XML.CreateTag("EMPLOYMENTTYPE",frmScreen.cboEmploymentType.value);
	XML.CreateTag("OCCUPATIONTYPE",frmScreen.cboOccupationType.value);
	XML.CreateTag("JOBTITLE",frmScreen.txtJobTitle.value);
	XML.CreateTag("MAINSTATUS",scScreenFunctions.GetRadioGroupValue(frmScreen,"MainEmploymentGroup"));

	if (m_sMetaAction == "Edit")
	{
		XML.CreateTag("EMPLOYMENTSEQUENCENUMBER", m_sEmploymentSequenceNumber);

		<% /* Retrieve the original third party/directory GUIDs */%>
		var sOriginalThirdPartyGUID = EmploymentXML.GetTagText("THIRDPARTYGUID");
		var sOriginalDirectoryGUID = EmploymentXML.GetTagText("DIRECTORYGUID");
		<% /* Only retrieve the address/contact details GUID if we are updating an existing third party record */%>
		var sAddressGUID = (sOriginalThirdPartyGUID != "") ? EmploymentXML.GetTagText("ADDRESSGUID") : "";
		var sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? EmploymentXML.GetTagText("CONTACTDETAILSGUID") : "";
	}

	<% /* If changing from a directory to a third party, or vice-versa, the old directory/thirdparty GUID
	 should still be specified to alert the middler tier to the fact that the old link needs deleting */%>
	XML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID != "" ? m_sDirectoryGUID : sOriginalDirectoryGUID);
	XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);

	<% /* Note the check for company name in the next if - this is to cater for the situation when NO third party
	data exists for the employment */%>
	if (!m_bDirectoryAddress && (frmScreen.txtCompanyName.value != ""))
	{
		<% /*  Store the third party details */%>
		XML.CreateActiveTag("THIRDPARTY");
		XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
		
		XML.CreateTag("THIRDPARTYTYPE", 5); <% /* Employer */%>
		XML.CreateTag("COMPANYNAME", frmScreen.txtCompanyName.value);

		<% /*  Address */%>
		XML.CreateTag("ADDRESSGUID", sAddressGUID);
		SaveAddress(XML, sAddressGUID);

		<% /*  Contact Details */%>
		XML.SelectTag(null, "THIRDPARTY");
		XML.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
		SaveContactDetails(XML, sContactDetailsGUID);
	}

	<% /*  Save the details */%>
	XML.SelectTag(null,"REQUEST");
	XML.CreateActiveTag("CRITICALDATACONTEXT");
	XML.SetAttribute("APPLICATIONNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null));
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null));
	XML.SetAttribute("SOURCEAPPLICATION","Omiga");
	XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
	XML.SetAttribute("ACTIVITYINSTANCE","1");
	XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
	XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
	XML.SetAttribute("COMPONENT","omCE.CustomerEmploymentBO");
	XML.SetAttribute("METHOD","SaveEmploymentDetails");

	<% /* BMIDS01110 MDC 29/11/2002 */ %>	
//	window.status = "Critical Data Check - please wait";
//	XML.RunASP(document,"OmigaTMBO.asp");
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			window.status = "Critical Data Check - please wait";
			XML.RunASP(document,"OmigaTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
	}
	<% /* BMIDS01110 MDC 29/11/2002 - End */ %>	

	window.status = "";
	
	bSuccess = XML.IsResponseOK();
	if (m_sMetaAction != "Edit")
		scScreenFunctions.SetContextParameter(window,"idEmploymentSequenceNumber", XML.GetTagText("EMPLOYMENTSEQUENCENUMBER"));

	XML = null;
	return(bSuccess);
}

function MainEmploymentCheck()
{
	var bRetVal = true;
	if (scScreenFunctions.GetRadioGroupValue(frmScreen,"MainEmploymentGroup") == "1")
	{
		var EmploymentXML = null;
		//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
		//EmploymentXML = new scXMLFunctions.XMLObject();
		EmploymentXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		EmploymentXML.CreateRequestTag(window,null);
		var tagCustomerList = EmploymentXML.CreateActiveTag("CUSTOMERLIST");
		EmploymentXML.CreateActiveTag("CUSTOMER");
		EmploymentXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
		EmploymentXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
		EmploymentXML.RunASP(document,"FindEmploymentList.asp");
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = EmploymentXML.CheckResponse(ErrorTypes);
		if ((ErrorReturn[1] == ErrorTypes[0]) | (EmploymentXML.XMLDocument.text == ""))
		{
			 <% /* Error: record not found */ %>
			bRetVal =  true
		}
		else if (ErrorReturn[0] == true)
		{
			<% /*  No error */ %>
			EmploymentXML.ActiveTag = null;
			nMainEmploymentCount = 0
			EmploymentXML.CreateTagList("EMPLOYMENT");
			var iCount;
			var iNumberOfEmployments = EmploymentXML.ActiveTagList.length;
			var iCurrentEmpSeqNumb = scScreenFunctions.GetContextParameter(window,"idEmploymentSequenceNumber",null)
			
			for (iCount = 0; (iCount < iNumberOfEmployments) && (iCount < 10); iCount++)
			{
				EmploymentXML.SelectTagListItem(iCount);
				if((EmploymentXML.GetTagText("EMPLOYMENTSEQUENCENUMBER")!=(iCurrentEmpSeqNumb)) && (EmploymentXML.GetTagText("MAINSTATUS") == "1"))
				{
					nMainEmploymentCount++;
				}
			}
			bRetVal = (nMainEmploymentCount == 0);
		}
		if (bRetVal == false) 
		{
			alert("Cannot have more than one Main Employment");
			frmScreen.optMainEmploymentYes.focus();
		}	
	}
	return bRetVal;
}

function SetAvailableFunctionalityOverride()
{
	<% /* if ((m_sEmploymentStatusType == "N") || (m_sEmploymentStatusType == "R")) (MAR1395) */ %>
	if (GetEmpValidation("N") || GetEmpValidation("R"))
	{
		scScreenFunctions.SetCollectionToReadOnly(spnContactDetails);
		scScreenFunctions.SetCollectionToReadOnly(spnAddress);
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCompanyName");
		frmScreen.btnPAFSearch.disabled = true;
		frmScreen.btnClear.disabled = true;
		frmScreen.btnDirectorySearch.disabled = true;
		//BG SYS1063 disable button.
		frmScreen.btnAddToDirectory.disabled = true;
	}
}

<% /* Peter Edney - 28/03/2006 */ %>
<% /* MAR1395  */ %>
function GetEmpValidation(sType)
{
	return scScreenFunctions.IsValidationType(frmScreen.cboEmploymentStatus, sType);
}

-->
</script>
<!-- #include FILE="includes/thirdpartydetails.asp" -->
</body>
</html>


