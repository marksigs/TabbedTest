<%@ LANGUAGE="JSCRIPT" %>
<html> 
<% /*
Workfile:      DC280.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Legal Representative Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		22/02/2000	Created
AD		15/03/2000	Incorporated third party include files.
AY		03/04/00	New top menu/scScreenFunctions change
MH		03/05/00	SYS0618 Postcode validation
MH      17/05/00     SYS0696 Validation and stuff
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
CL		05/03/01	SYS1920 Read only functionality added
JR		17/09/01	Omiplus24 - Removed FaxNumber, TelephoneNumber & EmailAddress. Now included in TPD.asp
MDC		01/10/01	SYS2785 Enable client versions to override labels
DPF		20/06/02	BMIDS00077 Changes made to bring file in line with Core V7.0.2
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		13/05/2002	BMIDS00004	Modified PopulateScreen() - Added DXID,DXLocation 
MV		12/07/2002  BMIDS00198  Display DxID and DxLocation
								Amended frmScreen.btnDirectorySearch.onclick() and PopulateScreen()
MV		31/07/2002	BMIDS00075	Modified SaveLegalRep(); Created CreatePayeeHistoryDetailsXML();						
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
DRC     01/10/2002  BMIDS00463  Warning message if Legal Rep is non-panel
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client
PSC		27/11/2002	BMIDS00998 - Amend to cater for Panel Bank Account Details not being present
MDC		12/12/2002	BM0094 Legal Rep Contact Details
MDC		10/01/2003	BM0244 Enable DxId and DxLocation saving to ApplicationLegalRep
MDC		13/01/2003	BM0250 
HMA     17/09/2003  BM0063     - Amend HTML text for radio buttons
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History

Prog	Date		AQR		Description
MF		04/08/2005	MAR20	Routes back to DC201
MF		08/08/2005	MAR20	Route depending on Global Parameters
MF		07/09/2005	MAR20	Add new fields, Load default legal rep, create task
							where selected legal rep is inactive
MF		12/09/2005	MAR20	Modified text for status label to reflect status of
							date and status db column.
HMA     01/11/2005  MAR144	Use TMLegalRepIncativeTaskID Global is used
PJO     15/12/2005  MAR827  Disable Directory Search on remortgage
HMA     07/03/2006  MAR1171 Disable radio buttons for Remortgage applications.
							Ensure correct default Legal Rep is selected.
HMA     08/03/2006  MAR1336 Save Payee History correctly.
						
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History

Prog	Date		AQR		Description
PB		01/06/2006	EP646	Removed MARS feature of disabling directory search for
							re-mortgage applications.
PB		07/07/2006	EP543	Populate 'Title-Other' field from XML
*/

%>
<HEAD>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>
<% /* removed as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077 
<OBJECT data="scScreenFunctions.asp" height=1 id=scScrFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data="scXMLFunctions.asp" height=1 id=scXMLFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
*/ %>

<% /* FORMS */ %>
<form id="frmToDC240" method="post" action="DC240.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC201" method="post" action="DC201.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC270" method="post" action="DC270.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC230" method="post" action="DC230.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
<div style="HEIGHT: 160px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Legal Representative Details</strong> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Company Name
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyName" maxlength="45" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" onchange="ContactDetailsChanged()">
		</span> 
	</span>

	<span style="LEFT: 370px; POSITION: absolute; TOP: 31px" class="msgLabel">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton">
	</span>

	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span style="TOP: 60px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Is in the directory?
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryYes" name="InDirectoryGroup" type="radio" value="1" onclick="InDirectoryChanged()"><label for="optInDirectoryYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryNo" name="InDirectoryGroup" type="radio" value="0"><label for="optInDirectoryNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 84px" class="msgLabel">
		Different Legal Rep for Lender?
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="optSeparateLegalRepYes" name="SeparateLegalRepGroup" type="radio" value="1"><label for="optSeparateLegalRepYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px">
			<input id="optSeparateLegalRepNo" name="SeparateLegalRepGroup" type="radio" value="0"><label for="optSeparateLegalRepNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 108px" class="msgLabel">
		Solicitor Status
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtSolicitorStatus" maxlength="45" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" readonly tabindex=-1>
		</span> 
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 135px" class="msgLabel">
		DX Number
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtDxNumber" maxlength="9" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
		</span> 
	</span>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 135px" class="msgLabel">
		DX Location
		<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
			<input id="txtDXLocation" maxlength="20" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
		</span> 
	</span>
	
	<span style="LEFT: 315px; POSITION: absolute; TOP: 10px; display:none;" class="msgLabel">		
		Status
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<select id="cboStatus" name="cboStatus" style="WIDTH: 150px; " class="msgCombo">
			</select>
		</span>
	</span>	
</div>

<div id="divThirdPartyDetails" style="HEIGHT: 232px; LEFT: 10px; POSITION: absolute; TOP: 225px; WIDTH: 604px" class="msgGroup">
<!-- #include FILE="includes/thirdpartydetails.htm" -->
</div> 
</form>

<% /* have amended top measurement for buttons from 440 -> 460 px - DPF 25/06/02 - BMIDS00077 */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 470px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->
<!-- #include FILE="includes/thirdpartydetails.asp" -->
<!-- #include FILE="includes/directorysearch.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/DC280Attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sXML = "";
var scScreenFunctions;

var LegalRepXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber  = "";
var m_blnReadOnly = false;
var m_sUserId = "";
var xmlActiveTag = "";
var XML = null;
var sOriginalThirdPartyGUID = "" ;
var sOriginalDirectoryGUID = "";
var sAddressGUID = "";
var sContactDetailsGUID = "";
var m_bUsingDirectory = false;
var m_bRemortgage = false;          <% /* MAR1171 */ %>
		
function frmScreen.btnDirectorySearch.onclick()
{
	<% /* Call the routine in TPD and then update the screen when it returns */ %>
	var nCurrentValue=scScreenFunctions.GetRadioGroupValue(frmScreen,"InDirectoryGroup");	
	
	ThirdPartyDetailsDirectorySearch();	
	ReadLegalRepStatusFromXML();
	
	<% /* MV - 12/07/2002 - BMIDS00198 - Display DxID and DxLocation*/%>		
	frmScreen.txtDxNumber.value = sDxID;
	frmScreen.txtDXLocation.value = sDxLocation;

	<% /* BM0094 MDC 11/12/2002 */ %>
	scScreenFunctions.SetCollectionToWritable(spnContactDetails);
	<% /* BM0244 MDC 10/01/2003 
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDxNumber");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDXLocation"); */ %>
	<% /* BM0094 MDC 11/12/2002 - End */ %>
	 
	if (nCurrentValue!=scScreenFunctions.GetRadioGroupValue(frmScreen,"InDirectoryGroup"))
		scScreenFunctions.SetRadioGroupValue(frmScreen,"SeparateLegalRepGroup","0");	
}

<% /* MF 15/08/2005 compare status with validation type in hidden combo */ %>
function ReadLegalRepStatusFromXML()
{		
	if(ThirdPartyDetailsXML.ActiveTag!=null){
		var sStatus = ThirdPartyDetailsXML.GetTagText("STATUS");
		var sActiveFrom = ThirdPartyDetailsXML.GetTagText("NAMEANDADDRESSACTIVEFROM");
		var sActiveTo = ThirdPartyDetailsXML.GetTagText("NAMEANDADDRESSACTIVETO");
		SetStatusText(sActiveFrom,sActiveTo,sStatus);
	}
}

function frmScreen.optInDirectoryYes.onclick ()
{
	InDirectoryChanged();
	scScreenFunctions.SetRadioGroupValue(frmScreen,"SeparateLegalRepGroup","0")
	
	<% /* BM0244 MDC 10/01/2003 
	// BM0094 MDC 13/12/2002
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDxNumber");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDXLocation");
	// BM0094 MDC 13/12/2002 
	*/ %>
	
}

function frmScreen.optInDirectoryNo.onclick ()
{
	frmScreen.txtSolicitorStatus.value="";
	InDirectoryChanged();
	scScreenFunctions.SetRadioGroupValue(frmScreen,"SeparateLegalRepGroup","0")
	<% /* BM0244 MDC 10/01/2003 
	// BM0094 MDC 13/12/2002 
	scScreenFunctions.SetFieldToWritable(frmScreen, "txtDxNumber");
	scScreenFunctions.SetFieldToWritable(frmScreen, "txtDXLocation");
	// BM0094 MDC 13/12/2002 
	*/ %>
}

function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
	scScreenFunctions.SetContextParameter(window,"idXML", "");	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
	var bThirdPartySummary = XML.GetGlobalParameterBoolean(document,"ThirdPartySummary");			
	if(bThirdPartySummary)
		frmToDC240.submit();
	else{
		var bNewPropertySummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");			
		if(bNewPropertySummary)	
			frmToDC201.submit();
		else
			frmToDC230.submit();		
	}
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		scScreenFunctions.SetContextParameter(window,"idXML", "");		
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
		var bThirdPartySummary = XML.GetGlobalParameterBoolean(document,"ThirdPartySummary");			
		if(bThirdPartySummary)
			frmToDC240.submit();
		else
			frmToDC270.submit();
	}
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
	m_sThirdPartyType = "10";
	m_bDirectoryOnly = false;

	frmScreen.btnAddToDirectory.style.visibility = "hidden";

	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	<% /* MAR1171  Check for Remortgage application */ %>
	m_bRemortgage = IsRemortgageApplication();
			
	FW030SetTitles("Add/Edit Legal Representative","DC280",scScreenFunctions);

	var sGroups = new Array("Country");
	objDerivedOperations = new DerivedScreen(sGroups);

	RetrieveContextData();
	SetThirdPartyDetailsMasks();

	ThirdPartyCustomise();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();	
	Initialise(true);
	
	if (m_bRemortgage == true)   <% /* MAR1171 */ %>
	{
		// frmScreen.btnDirectorySearch.disabled = true ; PB 01/06/2006 EP646
		
		<% /* MAR1171  Disable the "Is In Directory" radio buttons. */ %>
		scScreenFunctions.SetRadioGroupState(frmScreen, "InDirectoryGroup", "R");
		
		<% /* MAR1171  Set the value and disable the "Different Legal Rep For Lender" radio buttons. */ %>
		scScreenFunctions.SetRadioGroupValue(frmScreen, "SeparateLegalRepGroup","0");
		scScreenFunctions.SetRadioGroupState(frmScreen, "SeparateLegalRepGroup", "R");
	}
		
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC280");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	SetScreenToReadOnly();
	scScreenFunctions.SetFocusToFirstField(frmScreen);	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function IsFormChanged()
{		
	return (m_bDirectoryAddress != m_bUsingDirectory) || IsChanged();	
}

function CommitChanges()
{
	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsFormChanged() || m_sMetaAction == "Add")
			{
				if (scScreenFunctions.ValidatePostcode(frmScreen.txtPostcode))
					bSuccess = SaveLegalRep();
				else
					bSuccess=false;
			}
			ProcessInactiveLegalRepTask();
		}
		else
			bSuccess = false;
	return(bSuccess);
}

<% /* MAR20 Add task where selected Legal Rep is not Active.
			Need to read DB to check status in case it has changed 
			since this screen loaded. */ %>
function ProcessInactiveLegalRepTask()
{	
	if(m_sDirectoryGUID!=""){
		<% /* MAR20 Don't create task if Customer lock indicator is true
				and readonly flag is true (from context) */ %>
		var bCustLock = false;
		bCustLock = (scScreenFunctions.GetContextParameter(window,"CustomerLockIndicator","0") != "0");		
		if(!(bCustLock && m_blnReadOnly)){	
			var st = GetLegalRepStatus();
			if(!IsActiveFromStatusCombo(st)){						
				<% /* Create Task for Inactive Legal rep*/ %>				
				var bRet=LaunchAdHocTask();	
			}
		}
	}	
}

<% /* MF MAR20 Launch task.
	Get Global Parameter TMLegalRepInactiveTaskID
	Call Create Adhoc Task with task ID returned from 
	TMLegalRepInactiveTaskID global Parameter.  */ %>
function LaunchAdHocTask()
{
	var bSuccess = false;
	var sUserRole;
	var CaseTaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = CaseTaskXML.CreateRequestTag(window, "CreateAdhocCaseTask");	
	var sTaskID = ParamXML.GetGlobalParameterString(document,"TMLegalRepInactiveTaskID");   // MAR144				
															
	stageXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	stageXML.CreateRequestTag(window , "GetCurrentStage");
	stageXML.CreateActiveTag("CASEACTIVITY");
	stageXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	stageXML.SetAttribute("CASEID", scScreenFunctions.GetContextParameter(window,"idApplicationNumber",""));
	stageXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));
	stageXML.SetAttribute("ACTIVITYINSTANCE", "1");
	stageXML.RunASP(document, "MsgTMBO.asp");

	if(stageXML.IsResponseOK())
	{			
		<% /* Check if this task already exists as incomplete, only add if it does 
				not exist or if has been set to complete */ %>
		var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();						
		var tagList = stageXML.CreateTagList("CASETASK");			
		for (var nItem = 0; nItem < tagList.length && bSuccess == false; nItem++)
		{
			stageXML.SelectTagListItem(nItem);			
			if(stageXML.GetAttribute("TASKID")==sTaskID){				
				var sTaskStatus=stageXML.GetAttribute("TASKSTATUS");
				if (TempXML.IsInComboValidationList(document,"TaskStatus",sTaskStatus,["I"])){
					bSuccess = true;
				}		
			}				
		}	
		if(bSuccess == false){
			var sSeqNo="";					
			stageXML.SelectTag(null, "CASESTAGE");
			sSeqNo  = stageXML.GetAttribute("CASESTAGESEQUENCENO");				
			CaseTaskXML.SetAttribute("USERID", scScreenFunctions.GetContextParameter(window,"idUserID",null));
			CaseTaskXML.SetAttribute("UNITID", scScreenFunctions.GetContextParameter(window,"idUnitID",null));		
			sUserRole = scScreenFunctions.GetContextParameter(window,"idRole",null);
			CaseTaskXML.SetAttribute("USERAUTHORITYLEVEL", sUserRole);				
			CaseTaskXML.CreateActiveTag("APPLICATION");		
			CaseTaskXML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
			CaseTaskXML.ActiveTag = reqTag;	
			CaseTaskXML.CreateActiveTag("CASETASK");
			CaseTaskXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
			CaseTaskXML.SetAttribute("CASEID", scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null));	
			CaseTaskXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",null));
			CaseTaskXML.SetAttribute("ACTIVITYINSTANCE", "1");		
			CaseTaskXML.SetAttribute("CASESTAGESEQUENCENO", sSeqNo);		
			CaseTaskXML.SetAttribute("STAGEID", scScreenFunctions.GetContextParameter(window,"idStageID",null));		
			CaseTaskXML.SetAttribute("TASKID", sTaskID);		
			CaseTaskXML.RunASP(document, "OmigaTmBO.asp");		
			if(CaseTaskXML.IsResponseOK())			
				bSuccess = true;					
		}
	}
	return(bSuccess);
}

<% /* Inserts default values into all fields */ %>
function DefaultFields()
{
	ClearFields(true,true);
	scScreenFunctions.SetRadioGroupValue(frmScreen,"SeparateLegalRepGroup","0");
	m_sDirectoryGUID = "";
	m_bDirectoryAddress = false;
}

function Initialise()
{
	PopulateCombos();
	if(m_sMetaAction == "Edit")		
		PopulateScreen();
	else {
		DefaultFields();
	}
	
	<% /* MAR1171 Set Default Rep for Remortgage applications. */ %>
	if (m_bRemortgage == true)
	{
		LoadDefaultRep();
	}
	
	SetAvailableFunctionality();
	<% /* BM0094 MDC 11/12/2002 */ %>
	scScreenFunctions.SetCollectionToWritable(spnContactDetails);
	<% /* BM0244 MDC 10/01/2003 
	if (scScreenFunctions.GetRadioGroupValue(frmScreen,"InDirectoryGroup") == 1)
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDxNumber");
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDXLocation");
	} */ %>
	bLegalRep = true;
	<% /* BM0094 MDC 11/12/2002 - End */ %>
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}


<% /* MF MAR20 Read the directory details where the legal rep is on the panel 
		and determine the status of the rep
*/ %>
function GetLegalRepStatus()
{
	var bActive=false;
	var DetailsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	DetailsXML.CreateActiveTag("NAMEANDADDRESSDIRECTORY");
	DetailsXML.CreateTag("DIRECTORYGUID",m_sDirectoryGUID); 
	DetailsXML.RunASP(document,"GetDirectoryDetails.asp");
	DetailsXML.SelectTag(null,"NAMEANDADDRESSDIRECTORY");		
	var tag = DetailsXML.SelectTag(null, "STATUS");		
	return tag==null ? "" : tag.text;
}

function IsActiveFromStatusCombo(v_sOptionVal)
{
	var bActive=false;
	for(var iOption=0;iOption<frmScreen.cboStatus.children.length;iOption++){								
		if(v_sOptionVal==frmScreen.cboStatus.children[iOption].value){
			<% /* MF If Status = "A" then set solicitor status active */ %>
			if(frmScreen.cboStatus.children[iOption].getAttribute("ValidationType0")=="A")
				bActive=true;	
		}
	}
	return bActive;
}

<% /* MF Sets the text for the status textbox: 
	"Active" - If date from and to are in range,
	"Inactive - Date" - If status is active but date is out of range 
	"Inactive - Status" - If Status is Inactive */ %>
function SetStatusText(v_sActiveFromDate, v_sActiveToDate, v_sStatus)
{
	var sStatusText = "";	
	<% /* MF Now check status from PanelLegalRep table */ %>		
		if(IsActiveFromStatusCombo(v_sStatus)){			
			if (scScreenFunctions.CompareDateStringToSystemDate(v_sActiveFromDate,"<") &
				(scScreenFunctions.CompareDateStringToSystemDate(v_sActiveToDate,">=") | (v_sActiveToDate == "")))
			{
				sStatusText = "Active";	
			}else{
				sStatusText = "Inactive Date";						
			} 	
		}else
			sStatusText = "Inactive Status";							
	
	frmScreen.txtSolicitorStatus.value = sStatusText;
}

<% /* Populates the screen with details of the item selected in */ %>
function PopulateScreen()
{
	//next line replaced with line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//LegalRepXML = new scXMLFunctions.XMLObject();
	LegalRepXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	LegalRepXML.LoadXML(m_sXML);	
	LegalRepXML.SelectTag(null,"APPLICATIONTHIRDPARTY");

	<% /* BM0244 MDC 10/01/2003 */ %>
	frmScreen.txtDxNumber.value =   LegalRepXML.GetTagText("DXID");
	frmScreen.txtDXLocation.value = LegalRepXML.GetTagText("DXLOCATION");
	<% /* BM0244 MDC 10/01/2003 - End */ %>

	m_sDirectoryGUID = LegalRepXML.GetTagText("DIRECTORYGUID");
	m_sThirdPartyGUID = LegalRepXML.GetTagText("THIRDPARTYGUID");
	m_bDirectoryAddress = (m_sDirectoryGUID != "");
	m_bUsingDirectory = m_bDirectoryAddress; 
	with (frmScreen)
	{
		txtCompanyName.value = LegalRepXML.GetTagText("COMPANYNAME");
		scScreenFunctions.SetRadioGroupValue(frmScreen,"SeparateLegalRepGroup",LegalRepXML.GetTagText("SEPARATELEGALREPRESENTATIVE"));

		var sActiveFromDate = LegalRepXML.GetTagText("NAMEANDADDRESSACTIVEFROM");
		var sActiveToDate = LegalRepXML.GetTagText("NAMEANDADDRESSACTIVETO");
		
		if (m_bDirectoryAddress) 
		{
			var sStatus = GetLegalRepStatus();
			SetStatusText(sActiveFromDate,sActiveToDate,sStatus);		
		}
		else txtSolicitorStatus.value =""
		
		txtContactForename.value = LegalRepXML.GetTagText("CONTACTFORENAME");
		txtContactSurname.value = LegalRepXML.GetTagText("CONTACTSURNAME");
		cboTitle.value = LegalRepXML.GetTagText("CONTACTTITLE");
		<% /* PB 07/07/2006 EP543 Begin */ %>
		checkOtherTitleField();
		txtTitleOther.value = LegalRepXML.GetTagText("CONTACTTITLEOTHER");
		<% /* EP543 End */ %>
		
		objDerivedOperations.LoadCounty(LegalRepXML);
		
		txtDistrict.value = LegalRepXML.GetTagText("DISTRICT");
		txtFlatNumber.value = LegalRepXML.GetTagText("FLATNUMBER");
		txtHouseName.value = LegalRepXML.GetTagText("BUILDINGORHOUSENAME");
		txtHouseNumber.value = LegalRepXML.GetTagText("BUILDINGORHOUSENUMBER");
		txtPostcode.value = LegalRepXML.GetTagText("POSTCODE");
		txtStreet.value = LegalRepXML.GetTagText("STREET");
		txtTown.value = LegalRepXML.GetTagText("TOWN");
		//cboCountry.value = LegalRepXML.GetTagAttribute("COUNTRY","TEXT");
		cboCountry.value = LegalRepXML.GetTagText("COUNTRY");
	}
	
	var TempXML = LegalRepXML.ActiveTag;
	var ContactXML = LegalRepXML.SelectTag(null, "CONTACTDETAILS");
	if(ContactXML != null)
		m_sXMLContact = ContactXML.xml;
	LegalRepXML.ActiveTag = TempXML;
	
	<% /* BM0094 MDC 13/12/2002 */ %>
	<% /* MV - 12/07/2002 - BMIDS00198 - Display DxID and DxLocation*/%>		
	// if (LegalRepXML.SelectTag(null, "NAMEANDADDRESSDIRECTORY") != null )
	// {
	<% /* BM0244 MDC 10/01/2003
		frmScreen.txtDxNumber.value =   LegalRepXML.GetTagText("DXID");
		frmScreen.txtDXLocation.value = LegalRepXML.GetTagText("DXLOCATION"); */ %>
	// 	LegalRepXML.ActiveTag = TempXML;
	// }
	<% /* BM0094 MDC 13/12/2002 - End */ %>
}


<% /* MF MAR20 09/08/2005 Created to load XML for third party data
		where screen DC240 is not present in screen flow. DC240 normally 
		loads third party data and stores in the Context object.
*/ %>
function LoadThirdPartyData(v_sTPType)
{
	var sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B00003743");
	var sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");  

	// Retrieve main data
	ThirdPartyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ThirdPartyXML.CreateRequestTag(window,null);
	var tagApplication = ThirdPartyXML.CreateActiveTag("APPLICATION");
	ThirdPartyXML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	ThirdPartyXML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	ThirdPartyXML.RunASP(document,"FindApplicationThirdPartyList.asp");
	ThirdPartyXML.ActiveTag = null;
	ThirdPartyXML.CreateTagList("APPLICATIONTHIRDPARTY");
	<% /* Loop through third parties and select v_sTPType */ %>
	var iCount=0;
	var bFound=false;	
	while (iCount < ThirdPartyXML.ActiveTagList.length && !bFound)
	{
		ThirdPartyXML.SelectTagListItem(iCount);
		var sThirdPartyType = ThirdPartyXML.GetTagText("THIRDPARTYTYPE");
		if(sThirdPartyType=="")	sThirdPartyType = ThirdPartyXML.GetTagText("NAMEANDADDRESSTYPE");		
		if(sThirdPartyType == v_sTPType){
			bFound=true;
		}
		iCount++;
	}
	return bFound==true ? ThirdPartyXML.ActiveTag.xml : "";	
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	<% /* MF If no third party summary screen (DC240) present then read in the Third Party 
		data normally loaded inside DC240 */ %> 
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
	var bThirdPartySummary = XML.GetGlobalParameterBoolean(document,"ThirdPartySummary");			
	if(!bThirdPartySummary){
		m_sXML = LoadThirdPartyData("10");
		if(m_sXML==""){
			m_sMetaAction = "Add";			
		} else	 
			m_sMetaAction = "Edit";	
	}else if(m_sMetaAction == "Edit")
		m_sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
		
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B00003743");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");  
	m_sUserId = scScreenFunctions.GetContextParameter(window,"idUserId","1");  
	var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	//next line removed as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	//var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bCanAddToDirectory = false;
	XML = null;
}

<% /* MAR20 Is this application type remortgage */ %>
function IsRemortgageApplication()
{
	var sTypeOfApplication = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","");
	var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
	if (TempXML.IsInComboValidationList(document,"TypeOfMortgage",sTypeOfApplication,["R"]))
		return true;
	else
		return false;
}

<% /* MAR20 Get Default Legal Rep for either 'England & Wales' or 'Scotland'
	as set in Global Parameters */ %> 
function LoadDefaultRep()
{
	<% /* MAR1171 Theis function is only called for Remortgage applications */ %>

	var sPanelID="";
	var SearchXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	<% /* Load PanelID from Global Parameter */ %>
	var LoanPropertyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	LoanPropertyXML.CreateRequestTag(window,null);
	LoanPropertyXML.CreateActiveTag("LOANPROPERTYDETAILS");
	LoanPropertyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	LoanPropertyXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	LoanPropertyXML.RunASP(document,"GetLoanPropertyDetails.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = LoanPropertyXML.CheckResponse(ErrorTypes);
	if(ErrorReturn[1] == ErrorTypes[0])	{}
		<%/* Record not found so no panel Id to use */%>					
	else if (ErrorReturn[0] == true)
		<%/* Record found */%>
	{
		<%/* Get the property location held in the XML */%>
		LoanPropertyXML.SelectTag(null, "LOANPROPERTYDETAILS")	
		var sProp=LoanPropertyXML.GetTagText("PROPERTYLOCATION");		
		var ComboValidationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if(ComboValidationXML.IsInComboValidationList(document,"PropertyLocation",sProp,["FTE"]))
		{
			<% /* England default rep */ %>			
			sPanelID = SearchXML.GetGlobalParameterString(document,"RemortDefaultLegalRepEngland");						
		}
		else if(ComboValidationXML.IsInComboValidationList(document,"PropertyLocation",sProp,["FTS"]))
		{
			<% /* Scotland default rep */ %>
			sPanelID = SearchXML.GetGlobalParameterString(document,"RemortDefaultLegalRepScotland");
		}
	}		
	
	if(sPanelID!="")
	{
		var ArrayArguments = SearchXML.CreateRequestAttributeArray(window);
		DirectorySearch(SearchXML, ArrayArguments, sPanelID, "", "", "", "", "", "", "", "", "", "")	
		<% /* Now get the ThirdParty data from the GUID found here & 
			Populate the screen from this XML */ %>	
		SearchXML.SelectTag(null, "NAMEANDADDRESSDIRECTORY");
		m_sDirectoryGUID = SearchXML.GetTagText("DIRECTORYGUID");
		<% /* Initailise and use the ThirdPartyDetailsXML object from the include file */ %>
		ThirdPartyDetailsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		ThirdPartyDetailsXML.CreateRequestTag(window, "");		
		ThirdPartyDetailsXML.CreateActiveTag("NAMEANDADDRESSDIRECTORY");
		ThirdPartyDetailsXML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID);		
		ThirdPartyDetailsXML.RunASP(document, "GetThirdParty.asp");
		if(ThirdPartyDetailsXML.IsResponseOK())
		{
			ThirdPartyDetailsXML.SelectTag(null,"NAMEANDADDRESSDIRECTORY");	
			m_bDirectoryAddress = (m_sDirectoryGUID != "");
			m_bUsingDirectory = m_bDirectoryAddress; 
			with (frmScreen)
			{
				txtDxNumber.value =   ThirdPartyDetailsXML.GetTagText("DXID");
				txtDXLocation.value = ThirdPartyDetailsXML.GetTagText("DXLOCATION");
				txtCompanyName.value = ThirdPartyDetailsXML.GetTagText("COMPANYNAME");		
				var sActiveFromDate = ThirdPartyDetailsXML.GetTagText("NAMEANDADDRESSACTIVEFROM");
				var sActiveToDate = ThirdPartyDetailsXML.GetTagText("NAMEANDADDRESSACTIVETO");		
				if (m_bDirectoryAddress) 
				{
					var sStatus = GetLegalRepStatus();
					SetStatusText(sActiveFromDate,sActiveToDate,sStatus);		
				}
				else txtSolicitorStatus.value =""
				
				txtContactForename.value = ThirdPartyDetailsXML.GetTagText("CONTACTFORENAME");
				txtContactSurname.value = ThirdPartyDetailsXML.GetTagText("CONTACTSURNAME");
				cboTitle.value = ThirdPartyDetailsXML.GetTagText("CONTACTTITLE");		
				objDerivedOperations.LoadCounty(ThirdPartyDetailsXML);		
				txtDistrict.value = ThirdPartyDetailsXML.GetTagText("DISTRICT");
				txtFlatNumber.value = ThirdPartyDetailsXML.GetTagText("FLATNUMBER");
				txtHouseName.value = ThirdPartyDetailsXML.GetTagText("BUILDINGORHOUSENAME");
				txtHouseNumber.value = ThirdPartyDetailsXML.GetTagText("BUILDINGORHOUSENUMBER");
				txtPostcode.value = ThirdPartyDetailsXML.GetTagText("POSTCODE");
				txtStreet.value = ThirdPartyDetailsXML.GetTagText("STREET");
				txtTown.value = ThirdPartyDetailsXML.GetTagText("TOWN");		
				cboCountry.value = ThirdPartyDetailsXML.GetTagText("COUNTRY");			
			}
		}		
	}
}


function SaveLegalRep()
{
	var bSuccess = true;
	//next line replaced with line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,null)	
	<% /* MF 16/08/2005 set Directory address flag according to option 
			for "Is in directory" */ %>
	if(scScreenFunctions.GetRadioGroupValue(frmScreen,"InDirectoryGroup")==1)
		m_bDirectoryAddress=true;
	else
		m_bDirectoryAddress=false;

	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");
	xmlActiveTag = XML.CreateActiveTag("APPLICATIONLEGALREP");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("SEPARATELEGALREPRESENTATIVE", scScreenFunctions.GetRadioGroupValue(frmScreen,"SeparateLegalRepGroup"));

	<% /* BM0244 MDC 10/01/2003 */ %>
	XML.CreateTag("DXID", frmScreen.txtDxNumber.value);
	XML.CreateTag("DXLOCATION", frmScreen.txtDXLocation.value);
	<% /* BM0244 MDC 10/01/2003 - End */ %>	
	if (m_sMetaAction == "Edit")
	{
		LegalRepXML.SelectTag(null, "APPLICATIONTHIRDPARTY")
		sOriginalThirdPartyGUID = LegalRepXML.GetTagText("THIRDPARTYGUID");
		sOriginalDirectoryGUID = LegalRepXML.GetTagText("DIRECTORYGUID");
		sAddressGUID = (sOriginalThirdPartyGUID != "") ? LegalRepXML.GetTagText("ADDRESSGUID") : "";
		<% /* BM0094 MDC 12/12/2002
		sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? LegalRepXML.GetTagText("CONTACTDETAILSGUID") : ""; */ %>
		sContactDetailsGUID = LegalRepXML.GetTagText("CONTACTDETAILSGUID");
		<% /* BM0094 MDC 12/12/2002 - End */ %>
	}

	<% /* If changing from a directory to a third party, or vice-versa, the old directory/thirdparty GUID
	 should still be specified to alert the middler tier to the fact that the old link needs deleting */ %>
	<% /* MF MAR20 16/08/2005 The middle tier stores the directory guid in a column
			within the applicationlegalrep table, no link record is created.
	//XML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID != "" ? m_sDirectoryGUID : sOriginalDirectoryGUID);
	*/ %>
	XML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID);
	XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);

	if (!m_bDirectoryAddress)
	{
		<% /* Store the third party details */ %>
		XML.CreateActiveTag("THIRDPARTY");
		XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
		XML.CreateTag("THIRDPARTYTYPE", 10); 
		XML.CreateTag("COMPANYNAME", frmScreen.txtCompanyName.value);

		<% /* BM0094 MDC 13/12/2002 */ %>
		XML.CreateTag("DXID", frmScreen.txtDxNumber.value);
		XML.CreateTag("DXLOCATION", frmScreen.txtDXLocation.value);
		<% /* BM0094 MDC 13/12/2002 - End */ %>
		
		<% /* Address */%>
		XML.CreateTag("ADDRESSGUID", sAddressGUID);
		SaveAddress(XML, sAddressGUID);

		<% /* Contact Details */ %>
		<% /* BM0094 MDC 11/12/2002
		XML.SelectTag(null, "THIRDPARTY"); */ %>
		XML.SelectTag(null, "APPLICATIONLEGALREP");
		<% /* BM0094 MDC 11/12/2002 - End */ %>
		XML.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
		SaveContactDetails(XML, sContactDetailsGUID);
        <% /* DRC - BMIDS00463  */ %>
        alert("Warning: Legal Representative is not on Panel");
	}
	<% /* BM0094 MDC 11/12/2002 */ %>
	else
	{	
		XML.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
		SaveContactDetails(XML, sContactDetailsGUID);
	}
	<% /* BM0094 MDC 11/12/2002  - End */ %>

	<% /* MV - 31/07/2002 - BMIDS00179 */ %>					
	CreatePayeeHistoryDetailsXML();
	// 	XML.RunASP(document,"SaveLegalRep.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"SaveLegalRep.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}

<%/* MV - 31/07/2002 - BMIDS00179 */%>
function CreatePayeeHistoryDetailsXML()
{
	if (((m_sMetaAction == "Edit") && (sOriginalThirdPartyGUID != m_sThirdPartyGUID)) ||
		((m_sMetaAction == "Edit") && (sOriginalDirectoryGUID != m_sDirectoryGUID))   ||
		 (m_sMetaAction == "Add"))
	{
		XML.ActiveTag = xmlActiveTag;
			
		xmlActiveTag = XML.CreateActiveTag("PAYEEHISTORY");
		XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
		
		if (m_bDirectoryAddress)
		{
			<% /* PSC 27/11/2002 BMIDS00998 - Start */ %>
			var sSortCode = "";
			var sAccountNo = "";
			var sBankName = "";
			
			ThirdPartyDetailsXML.SelectTag(null,"PANELBANKACCOUNTLIST/PANELBANKACCOUNT");
			
			if (ThirdPartyDetailsXML.ActiveTag != null)
			{
				sSortCode = ThirdPartyDetailsXML.GetTagText("BANKSORTCODE");
				sAccountNo = ThirdPartyDetailsXML.GetTagText("ACCOUNTNUMBER");
				sBankName = ThirdPartyDetailsXML.GetTagText("BANKNAME");
			}		
			XML.CreateTag("BANKSORTCODE",sSortCode);
			XML.CreateTag("ACCOUNTNUMBER",sAccountNo);
			XML.CreateTag("BANKNAME",sBankName);
			<% /* PSC 27/11/2002 BMIDS00998 - End */ %>
			
			XML.CreateTag("USERID",m_sUserId);

			<% /* MO - BMIDS00807 */ %>
			var dteToday = scScreenFunctions.GetAppServerDate(true);
			<% /* var dteToday = new Date(); */ %>
			XML.CreateTag("CREATEDATETIME",dteToday);
		
			XML.CreateTag("DXID",frmScreen.txtDxNumber.value);
			XML.CreateTag("DXLOCATION",frmScreen.txtDXLocation.value);
		}
			
		XML.CreateActiveTag("THIRDPARTY");
		XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
		XML.CreateTag("THIRDPARTYTYPE",10);
		XML.CreateTag("COMPANYNAME",frmScreen.txtCompanyName.value);
			
		if (m_bDirectoryAddress)
		{
			XML.CreateActiveTag("ADDRESS");
			XML.AddXMLBlock(ThirdPartyDetailsXML.SelectTag(null,"ADDRESS"));
		}
		else
		{
			XML.CreateTag("ADDRESSGUID", sAddressGUID);
			SaveAddress(XML, sAddressGUID);
		}
				
	}

}

<% /* Populates all combos on the screen */ %>
function PopulateCombos()
{	
	PopulateTPTitleCombo();

	var ComboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sLegalRepStatus = new Array("LegalRepStatus");
	if(ComboXML.GetComboLists(document,sLegalRepStatus))	
		ComboXML.PopulateCombo(document,frmScreen.cboStatus,"LegalRepStatus",true);
	ComboXML = null;
	
	var sControlList = new Array(frmScreen.cboCountry);
	objDerivedOperations.GetComboLists(sControlList);
}

function SetScreenToReadOnly()
{
	if (m_sReadOnly=="1")
	{		
		frmScreen.btnDirectorySearch.disabled =true;
		frmScreen.btnPAFSearch.disabled = true;
		frmScreen.btnClear.disabled = true;
	}
}

-->

</script>
</body>
</html>




