<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP430.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Details of Property and Insurance
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Descrip
JR		31/1/01		Screen Design
JR		8/2/01		Added functionality
CL		12/03/01	SYS1920 Read only functionality added
JR		15/03/01	SYS1878 Change create report on title to call omigatmbo.asp
JR		16/03/01	SYS2089 Use scScreenFunctions instead of formscreen to enable/diable text and radio groups
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
JR		22/03/01	SYS2048 Correct setting radio btn attributes
JR		10/04/01	SYS2251 Fixed runtime error which was a result of variable not declared.
JR		19/04/01	SYS2048 Added AddressType attribute when setting address attributes for LeaseHold and Freudal values
JR		18/05/01	SYS2048	Added STAGEID and CASESTAGESEQUENCENO attributes to CreateReportonTitle request. 
BG		26/11/01	SYS3107 need to change the TaskXML held in context to reflect the task 
					has been moved to pending from not started in ReportOnTitleAction function.
JLD		04/02/02	SYS3109 changed tab order for divAddress
JLD		04/02/02	SYS2670 added attribs file for date field.
DRC     11/02/02    SYS4056 changed wording to opposite for Property Match question
MEVA	16/04/02	SYS3091 Enable / Disable Free Holder name depending on Tenure type
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog    Date           Description
DB      11/11/2002	   Removed any reference/route to ap440 as this is not required.
HMA     18/09/2003     Amend HTML text for radio buttons
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History : 

Prog	Date		AQR			Description
LDM		27/04/2006	MAR1624		Only create 1 report on title record
*/
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
db History:
Prog    Date           Description
DRC     27/04/2006     EP440 - remove routing to AP450 (further questions)
PB		06/06/2006	   EP696 MAR1698 only update the address if it is a Freeholder address
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/


/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"> 
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>

<% /* Specify Forms Here */%>
<!-- <form id="frmToAP440" method="post" action="AP440.asp" STYLE="DISPLAY: none"></form> -->
<form id="frmToAP450" method="post" action="AP450.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP400" method="post" action="AP400.asp" STYLE="DISPLAY: none"></form>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<form id="frmScreen" validate  ="onchange" mark year4>

<div id="divProperty" style="HEIGHT: 60px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Tenure of Property
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboPropertyTenure" style="WIDTH: 200px" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Name of Freeholder
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtFreeholderName" maxlength="26" style="POSITION: absolute; WIDTH: 200px" class="msgTxt">
		</span>
	</span>
	
</div>
<div id="divAddress" style="HEIGHT: 110px; LEFT: 10px; POSITION: absolute; TOP: 125px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Postcode
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtFreeholderPostcode" maxlength="8" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 35px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Flat No.
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtFreeholderFlatNumber" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 55px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		House Name<br> &amp; No.
		<span style="TOP: 1px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtFreeholderHouseName" maxlength="26" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
		<span style="TOP: 1px; LEFT: 270px; POSITION: ABSOLUTE">
			<input id="txtFreeholderHouseNumber" maxlength="10" style="WIDTH: 50px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 87px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Street
		<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
			<input id="txtFreeholderStreet" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 10px; LEFT: 350px; POSITION: ABSOLUTE" class="msgLabel">
		District
		<span style="TOP: -3px; LEFT: 40px; POSITION: ABSOLUTE">
			<input id="txtFreeholderDistrict" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 35px; LEFT: 350px; POSITION: ABSOLUTE" class="msgLabel">
		Town
		<span style="TOP: -3px; LEFT: 40px; POSITION: ABSOLUTE">
			<input id="txtFreeholderTown" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 60px; LEFT: 350px; POSITION: ABSOLUTE" class="msgLabel">
		County
		<span style="TOP: -3px; LEFT: 40px; POSITION: ABSOLUTE">
			<input id="txtFreeholderCounty" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>	
	<span style="TOP: 87px; LEFT: 350px; POSITION: ABSOLUTE" class="msgLabel">
		Country
		<span style="TOP: -3px; LEFT: 40px; POSITION: ABSOLUTE">
			<select id="cboFreeholderCountry" style="WIDTH: 200px" class="msgCombo" readonly></select>
		</span>
	</span>
</div>	
<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
<div id="divLeaseTie" style="HEIGHT: 80px; LEFT: 10px; POSITION: absolute; TOP: 240px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Does the lease tie the buildings insurance to
		<span style="LEFT: 0px; POSITION: absolute; TOP: 15px" class="msgLabel">
			a particular insurance company or agency?
		</span>	
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optLeaseTie_Yes" name="LeaseTie" type="radio" value="1" disabled><label for="optLeaseTie_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 270px; POSITION: absolute; TOP: -3px">
			<input id="optLeaseTie_No" name="LeaseTie" type="radio" value="0" disabled><label for="optLeaseTie_No" class="msgLabel">No</label>
		</span>	
	</span>
	<span style="LEFT: 350px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Details:
		<span style="LEFT: 40px; POSITION: absolute; TOP: 1px">
			<TEXTAREA id=txtLeaseTieDetails rows=4 style="WIDTH: 205px" maxlength="255" class=msgTxt></TEXTAREA>
		</span>
	</span>
</div>
<div id="divInsurance" style="HEIGHT: 80px; LEFT: 10px; POSITION: absolute; TOP: 325px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Are the insurance arrangements stipulated
		<span style="LEFT: 0px; POSITION: absolute; TOP: 15px" class="msgLabel">
			in the rent or lease charge deed?
		</span>	
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optInsuranceArrangements_Yes" name="InsuranceArrangements" type="radio" value="1"><label for="optInsuranceArrangements_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 270px; POSITION: absolute; TOP: -3px">
			<input id="optInsuranceArrangements_No" name="InsuranceArrangements" type="radio" value="0"><label for="optInsuranceArrangements_No" class="msgLabel">No</label>
		</span>	
	</span>
	<span style="LEFT: 350px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Details:
		<span style="LEFT: 40px; POSITION: absolute; TOP: 1px">
			<TEXTAREA id=txtInsuranceArrangementDetails rows=4 style="WIDTH: 205px" maxlength="255" class=msgTxt></TEXTAREA>
		</span>
	</span>
</div>
<div id="divFireCover" style="HEIGHT: 30px; LEFT: 10px; POSITION: absolute; TOP: 410px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Does this call for fire cover?
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optFireCover_Yes" name="FireCover" type="radio" value="1"><label for="optFireCover_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 270px; POSITION: absolute; TOP: -3px">
			<input id="optFireCover_No" name="FireCover" type="radio" value="0"><label for="optFireCover_No" class="msgLabel">No</label>
		</span>	
	</span>
</div>
<div id="divPropertyMatch" style="HEIGHT: 80px; LEFT: 10px; POSITION: absolute; TOP: 445px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
	 
		Do the details of the property and of the
		<span style="LEFT: 0px; POSITION: absolute; TOP: 15px" class="msgLabel">
<% /* AQR SYS4056  DRC - changed wording to opposite */%>		
			transaction vary from the particulars
		</span>
		<span style="LEFT: 0px; POSITION: absolute; TOP: 30px" class="msgLabel">
			in the offer of advance?
		</span>		
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optPropertyMatch_Yes" name="PropertyMatch" type="radio" value="1"><label for="optPropertyMatch_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 270px; POSITION: absolute; TOP: -3px">
			<input id="optPropertyMatch_No" name="PropertyMatch" type="radio" value="0"><label for="optPropertyMatch_No" class="msgLabel">No</label>
		</span>	
	</span>
	<span style="LEFT: 350px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Details:
		<span style="LEFT: 40px; POSITION: absolute; TOP: 1px">
			<TEXTAREA id=txtPropertyMatchDetails rows=4 style="WIDTH: 205px" maxlength="255" class=msgTxt></TEXTAREA>
		</span>
	</span>
</div>
<div id="divFurtherInspection" style="HEIGHT: 45px; LEFT: 10px; POSITION: absolute; TOP: 530px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Further inspection, Society's Branch 
		<span style="LEFT: 0px; POSITION: absolute; TOP: 15px" class="msgLabel">
			Office to arrange this?
		</span>
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optFurtherInspection_Yes" name="FurtherInspection" type="radio" value="1"><label for="optFurtherInspection_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 270px; POSITION: absolute; TOP: -3px">
			<input id="optFurtherInspection_No" name="FurtherInspection" type="radio" value="0"><label for="optFurtherInspection_No" class="msgLabel">No</label>
		</span>	
	</span>
</div>
<div id="divValuersReport" style="HEIGHT: 45px; LEFT: 10px; POSITION: absolute; TOP: 580px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Is the release of the funds subject 
		<span style="LEFT: 0px; POSITION: absolute; TOP: 15px" class="msgLabel">
			to a Valuer's report?
		</span>
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optValuersReport_Yes" name="ValuersReport" type="radio" value="1"><label for="optValuersReport_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 270px; POSITION: absolute; TOP: -3px">
			<input id="optValuersReport_No" name="ValuersReport" type="radio" value="0"><label for="optValuersReport_No" class="msgLabel">No</label>
		</span>	
	</span>
</div>
<div id="divInspectionDate" style="HEIGHT: 35px; LEFT: 10px; POSITION: absolute; TOP: 630px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Ready for inspection on 
		<span style="LEFT: 225px; POSITION: absolute; TOP: -3px">
			<input id="txtInspectionDate" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
</div>
</form>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 690px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/AP430attribs.asp" -->
<%/* CODE */%>
<script language="JScript">
<!--

var scScreenFunctions;
var xmlCombos = null ;
var RotXML = null ; //XML results returned from ROT
var XML = null ;
var taskXML = null ;
var strAddressGuid ;

var m_sApplicationNumber, m_sApplicationFactFindNumber ;
var m_sUnitId = "" ;
var m_sUserId = "" ;
var m_sUserRole = "" ;
var m_sMetaAction = "" ;
var m_sReadOnly = "";
var m_sTaskXML = "";
var m_sDistributionChannelId = "" ;
var m_blnReadOnly = false;
var m_EntryStatus = null;
//PB 06/06/06 EP696/MAR1698
var m_bHaveFreeholderAddress = false;
//PB EP696/MAR1698 End

/** EVENTS **/


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<% /* DRC EP440 - remove routing to AP450 (further questions) */ %>
	//var sButtonList = new Array("Submit","Cancel","Next");
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Details of Property & Insurance","AP430",scScreenFunctions);
		
	GetCountryList() ;
	GetPropertyTenureList() ;
	RetrieveContextData();
	SetMasks();	
	Initialise();

	Validation_Init();
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP430");
	
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
	/* TEST
	
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","C00059072");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
	//scScreenFunctions.SetContextParameter(window,"idCustomerRoleType","1");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber","00000094");
	scScreenFunctions.SetContextParameter(window,"idMortgageApplicationValue","10");
	
	END TEST */
			
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	
	m_sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0")
	m_sUserId = scScreenFunctions.GetContextParameter(window,"idUserId","");
	m_sUnitId =	scScreenFunctions.GetContextParameter(window,"idUnitId","");
	
	m_sUserId = scScreenFunctions.GetContextParameter(window,"idUserId","");
	m_sUnitId =	scScreenFunctions.GetContextParameter(window,"idUnitId","");
	m_sDistributionChannelId =	scScreenFunctions.GetContextParameter(window,"idDistributionChannelId","");
		
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = 	scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","");
			
//DEBUG
//m_sMetaAction = "0" ;
//m_sTaskXML = "<CASETASK TASKID=\"22\" CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"2\" TASKSTATUS=\"40\"/>";	
}

function Initialise()
{
	if(m_sTaskXML.length > 0) 
	{
		xmlCombos = new top.frames[1].document.all.scXMLFunctions.XMLObject() ;
		taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject() ;
		taskXML.LoadXML(m_sTaskXML) ;
		taskXML.ActiveTag = null;
		taskXML.SelectTag(null, "CASETASK");
		GetPropertyTenureList() ;
		GetCountryList() ;
		m_EntryStatus = taskXML.GetAttribute ("TASKSTATUS");<%/* LDM 27/04/2006 MAR1624 need to know what the task status on entry to the screen */%> 
		PopulateScreen() ;
	}	
}

function PopulateScreen()
{	
	<%/* LDM 27/04/2006 MAR1624 */%> 
	RotXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	RotXML.CreateRequestTag(window , "GetReportOnTitleData");
	RotXML.CreateActiveTag("REPORTONTITLE");
	RotXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	RotXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);

	//Pass XML to omROTBO
	RotXML.RunASP(document,"ReportOnTitle.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = RotXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{
		taskXML.SetAttribute("TASKSTATUS",20);	<%/* LDM 27/04/2006 MAR1624 force to be a status of update if ROT rec is there already */%> 
		scScreenFunctions.SetContextParameter(window,"idTaskXML",taskXML.XMLDocument.xml);
	}
	<%/* LDM 27/04/2006 MAR1624 */%> 

	DisablingTextFields("D") ; //Disabling Address and lease tie details
	if (taskXML.GetAttribute ("TASKSTATUS") != "10")
	{
		//Verify XML response
		if(RotXML.IsResponseOK())
		{
			PopulateScreenFields() ;			
		}	
	}	
	if (m_sReadOnly == "1") scScreenFunctions.SetScreenToReadOnly(frmScreen) ;
}

function PopulateScreenFields()
{
	// Process the returning XML
	RotXML.ActiveTag = null;
	RotXML.SelectTag(null, "REPORTONTITLE");
		
	frmScreen.cboPropertyTenure.value  = RotXML.GetAttribute("PROPERTYTENURE") ;
	frmScreen.txtFreeholderName.value = (RotXML.GetAttribute("NAMEOFFREEHOLDER") != null) ? RotXML.GetAttribute("NAMEOFFREEHOLDER"): "" ;
	scScreenFunctions.SetRadioGroupValue(frmScreen,"LeaseTie", RotXML.GetAttribute("BUILDINGINSURANCETIE")) ;
	frmScreen.txtLeaseTieDetails.value = (RotXML.GetAttribute("BUILDINGINSURANCETIEDETAILS") != null) ? RotXML.GetAttribute("BUILDINGINSURANCETIEDETAILS") : "" ;
	scScreenFunctions.SetRadioGroupValue(frmScreen,"InsuranceArrangements", RotXML.GetAttribute("INSURANCEARRANGEMENTS")) ;
	frmScreen.txtInsuranceArrangementDetails.value = (RotXML.GetAttribute("INSURANCEARRANGEMENTSDETAILS") != null) ? RotXML.GetAttribute("INSURANCEARRANGEMENTSDETAILS") : "" ;
	scScreenFunctions.SetRadioGroupValue(frmScreen,"FireCover", RotXML.GetAttribute("FIRECOVERREQUIRED")) ;
	scScreenFunctions.SetRadioGroupValue(frmScreen,"PropertyMatch", RotXML.GetAttribute("DETAILSMATCHOFFER")) ;
	frmScreen.txtPropertyMatchDetails.value = (RotXML.GetAttribute("DETAILSMATCHOFFERDETAILS") != null) ? RotXML.GetAttribute("DETAILSMATCHOFFERDETAILS") : "" ;
	scScreenFunctions.SetRadioGroupValue(frmScreen,"FurtherInspection", RotXML.GetAttribute("FURTHERINSPBRANCHARRANGEMENT")) ;
	scScreenFunctions.SetRadioGroupValue(frmScreen,"ValuersReport", RotXML.GetAttribute("RELEASEOFFUNDSVALUERSREPORT")) ;
	frmScreen.txtInspectionDate.value = (RotXML.GetAttribute("READYFORINSPECTIONDATE") != null) ? RotXML.GetAttribute("READYFORINSPECTIONDATE") : "" ;
	
	//Ensure relevant screen fields is disabled
	//if ( (RotXML.GetAttribute("PROPERTYTENURE") == 2) )
	
	//Force the combo to display correct option	
	frmScreen.cboPropertyTenure.onchange() ;
	
	if (RotXML.GetAttribute("BUILDINGINSURANCETIE") == 0)
		scScreenFunctions.SetFieldState(frmScreen, "txtLeaseTieDetails", "D");
	if (RotXML.GetAttribute("INSURANCEARRANGEMENTS") == 0)
		scScreenFunctions.SetFieldState(frmScreen, "txtInsuranceArrangementDetails", "D")
	if (RotXML.GetAttribute("DETAILSMATCHOFFER") == 0)
		scScreenFunctions.SetFieldState(frmScreen, "txtPropertyMatchDetails", "D")				
	
	//populate address details
	<% /* PB 06/06/06 EP696/MAR1698
	RotXML.SelectTag(null, "ADDRESS") ;
	strAddressGuid = RotXML.GetAttribute("ADDRESSGUID") ; */ %>
	var xmlAddrList = RotXML.CreateTagList("ADDRESS")
	for (var iCount=0; iCount < RotXML.ActiveTagList.length ; iCount++)
	{
		<% /* PB EP696/MAR1698 End */ %>
		RotXML.SelectTagListItem(iCount) ;
	
		//If AddressType is not Completion Correspondence Address, populate address fields.
		if (RotXML.GetAttribute("ADDRESSTYPE") != "10")
		{
			frmScreen.txtFreeholderHouseName.value = (RotXML.GetAttribute("BUILDINGORHOUSENAME") != null) ? RotXML.GetAttribute("BUILDINGORHOUSENAME") : "" ;
			frmScreen.txtFreeholderHouseNumber.value = (RotXML.GetAttribute("BUILDINGORHOUSENUMBER") != null) ? RotXML.GetAttribute("BUILDINGORHOUSENUMBER") : "" ; 						
			frmScreen.txtFreeholderFlatNumber.value = (RotXML.GetAttribute("FLATNUMBER") != null) ? RotXML.GetAttribute("FLATNUMBER") : "" ;
			frmScreen.txtFreeholderStreet.value = (RotXML.GetAttribute("STREET") != null) ? RotXML.GetAttribute("STREET") : "" ;
			frmScreen.txtFreeholderDistrict.value = (RotXML.GetAttribute("DISTRICT") != null) ? RotXML.GetAttribute("DISTRICT") : "" ;
			frmScreen.txtFreeholderTown.value = (RotXML.GetAttribute("TOWN") != null) ? RotXML.GetAttribute("TOWN") : "" ;
			frmScreen.txtFreeholderCounty.value = (RotXML.GetAttribute("COUNTY") != null) ? RotXML.GetAttribute("COUNTY") : "" ;
			frmScreen.cboFreeholderCountry.value = RotXML.GetAttribute("COUNTRY") ;
			frmScreen.txtFreeholderPostcode.value = (RotXML.GetAttribute("POSTCODE") != null) ? RotXML.GetAttribute("POSTCODE") : "" ;
		}
	}
} 

function ReportOnTitleAction()
{
	var reqTag ;
	XML = null ;
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject()
	
	if ( (taskXML.GetAttribute ("TASKID") == null) || (m_sReadOnly == "1") ) return true ;

	if (taskXML.GetAttribute ("TASKSTATUS") == "20")
	{
		reqTag = XML.CreateRequestTag(window , "UpdateReportOnTitle");
							
		//Complete the XML with ROT Attributes 
		SetROTAttributes(reqTag, "Update") ; 
	
		//Pass XML to OmRotBO
		// 		XML.RunASP(document,"ReportOnTitle.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"ReportOnTitle.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	
		//Verify XML response
		<% /* LDM 27/04/2006 MAR1624 */ %>
		if(XML.IsResponseOK())
		{
			if (m_EntryStatus == "20")
			{
				return true;
			}
			else (m_EntryStatus == "10")
			{
				<% /* LDM 27/04/2006 MAR1624 - If everything worked ok, need to set the task to pending.
					If this is the first time into the cot screens by "any" action; then the status will be automatically moved
					to pending, when the createreportontitle is run.
					If we got here by the details btn in the application stage. Then the status is already pending, so do not reset to pending
					If we got here by creating another cot action and this is the first time in need to reset the status to pending  */ %>
				SetToPendingXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				SetToPendingXML.CreateRequestTag(window , "UpdateCaseTask");
				SetToPendingXML.CreateActiveTag("CASETASK");
				SetToPendingXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
				SetToPendingXML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
				SetToPendingXML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
				SetToPendingXML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
				SetToPendingXML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
				SetToPendingXML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
				SetToPendingXML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
				SetToPendingXML.SetAttribute("TASKSTATUS", "20"); <% /* Pending */ %>
				
				SetToPendingXML.RunASP(document, "msgTMBO.asp") ;
				
				if (SetToPendingXML.IsResponseOK())
				{
					return true;
				}
				else
				{
					return false;
				}
			}
		}
		else
		{
			return false;
		}
		<% /* LDM 27/04/2006 MAR1624 */ %>
	}
	else
	{
		if (taskXML.GetAttribute ("TASKSTATUS") == "10")
		{					
			reqTag = XML.CreateRequestTag(window , "CreateReportOnTitle");
			XML.SetAttribute("USERID", m_sUserId) ;
			XML.SetAttribute("UNITID", m_sUnitId) ;
			XML.SetAttribute("USERAUTHORITYLEVEL", m_sUserRole) ;
						
			XML.ActiveTag = reqTag ;
			XML.CreateActiveTag("CASETASK");
			XML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
			XML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
			XML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
			XML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
			XML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
			// JR - SYS2048, add the following attributes to Request
			XML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
			XML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
					
			//Complete the XML with ROT Attributes
			SetROTAttributes(reqTag, "Create") ;
						
			//Pass XML to omTmBO
			// 			XML.RunASP(document,"OmigaTMBO.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document,"OmigaTMBO.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

	
			//Verify XML response
			//BG 26/11/01 SYS3107 need to change the TaskXML held in context to reflect
			//the task has been moved to pending from not started.
			//if(!XML.IsResponseOK())	return false ;
		
			if(!XML.IsResponseOK())	
			{
				return false ;
			}
			else
			{
				taskXML.SetAttribute("TASKSTATUS",20);
				scScreenFunctions.SetContextParameter(window,"idTaskXML",taskXML.XMLDocument.xml);
			}				
		}
	}
	return true ;
}

function SetROTAttributes(reqTag, sOperation) 
{
	var strRadioBtnChk ;
	
	XML.ActiveTag = reqTag ;
	XML.CreateActiveTag("REPORTONTITLE");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber) ;
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	if (sOperation == "Update") XML.SetAttribute("ROTGUID", RotXML.GetAttribute("ROTGUID")) ;
	XML.SetAttribute("PROPERTYTENURE", frmScreen.cboPropertyTenure.value) ;
	XML.SetAttribute("NAMEOFFREEHOLDER", frmScreen.txtFreeholderName.value) ;
	
	//SYS2048
	strRadioBtnChk = (scScreenFunctions.GetRadioGroupValue(frmScreen,"LeaseTie") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"LeaseTie") : "" ;
	XML.SetAttribute("BUILDINGINSURANCETIE", strRadioBtnChk) ;
	XML.SetAttribute("BUILDINGINSURANCETIEDETAILS", frmScreen.txtLeaseTieDetails.value) ;
	
	strRadioBtnChk = (scScreenFunctions.GetRadioGroupValue(frmScreen,"InsuranceArrangements") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"InsuranceArrangements") : "" ;
	XML.SetAttribute("INSURANCEARRANGEMENTS", strRadioBtnChk) ;
	XML.SetAttribute("INSURANCEARRANGEMENTSDETAILS", frmScreen.txtInsuranceArrangementDetails.value) ;

	strRadioBtnChk = (scScreenFunctions.GetRadioGroupValue(frmScreen,"PropertyMatch") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"PropertyMatch") : "" ;
	XML.SetAttribute("DETAILSMATCHOFFER", strRadioBtnChk) ;
	XML.SetAttribute("DETAILSMATCHOFFERDETAILS", frmScreen.txtPropertyMatchDetails.value) ;

	strRadioBtnChk = (scScreenFunctions.GetRadioGroupValue(frmScreen,"FireCover") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"FireCover") : "" ;
	XML.SetAttribute("FIRECOVERREQUIRED", strRadioBtnChk) ;
	
	strRadioBtnChk = (scScreenFunctions.GetRadioGroupValue(frmScreen,"FurtherInspection") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"FurtherInspection") : "" ;
	XML.SetAttribute("FURTHERINSPBRANCHARRANGEMENT", strRadioBtnChk) ;

	strRadioBtnChk = (scScreenFunctions.GetRadioGroupValue(frmScreen,"ValuersReport") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"ValuersReport") : "" ;
	XML.SetAttribute("RELEASEOFFUNDSVALUERSREPORT", strRadioBtnChk) ;
		
	XML.SetAttribute("READYFORINSPECTIONDATE", frmScreen.txtInspectionDate.value) ;
	
	<% /* PB 06/06/06 EP696 MAR1698
	XML.ActiveTag = reqTag ;
	XML.CreateActiveTag("ADDRESS") ;
	if (sOperation == "Update") XML.SetAttribute("ADDRESSGUID", strAddressGuid);
	XML.SetAttribute("BUILDINGORHOUSENAME", frmScreen.txtFreeholderHouseName.value);
	XML.SetAttribute("BUILDINGORHOUSENUMBER", frmScreen.txtFreeholderHouseNumber.value);
	XML.SetAttribute("FLATNUMBER", frmScreen.txtFreeholderFlatNumber.value);
	XML.SetAttribute("STREET", frmScreen.txtFreeholderStreet.value);
	XML.SetAttribute("DISTRICT", frmScreen.txtFreeholderDistrict.value);
	XML.SetAttribute("TOWN", frmScreen.txtFreeholderTown.value);
	XML.SetAttribute("COUNTY", frmScreen.txtFreeholderCounty.value);
	XML.SetAttribute("COUNTRY", frmScreen.cboFreeholderCountry.value);
	XML.SetAttribute("POSTCODE", frmScreen.txtFreeholderPostcode.value);
	*/ %>
	
	//Add AddressType attribute only if selected combo property tenure = 'LeaseHold' or 'Feudal'
	var sCurrentComboValue = frmScreen.cboPropertyTenure.value ; 
	var sROTAddressTypeValue = "" ;
	var bReturn = xmlCombos.IsInComboValidationList(document,"PropertyTenure", sCurrentComboValue, ["L"]);
	
	if (bReturn == true)
	{
		sROTAddressTypeValue = GetRotAddressTypeValue() ;	
		<% /* PB EP696/MAR1698 only update the address if it is a Freeholder address.*/ %>
		XML.ActiveTag = reqTag ;
		XML.CreateActiveTag("ADDRESS") ;
		if (sOperation == "Update" && m_bHaveFreeholderAddress) XML.SetAttribute("ADDRESSGUID", strAddressGuid);
		XML.SetAttribute("BUILDINGORHOUSENAME", frmScreen.txtFreeholderHouseName.value);
		XML.SetAttribute("BUILDINGORHOUSENUMBER", frmScreen.txtFreeholderHouseNumber.value);
		XML.SetAttribute("FLATNUMBER", frmScreen.txtFreeholderFlatNumber.value);
		XML.SetAttribute("STREET", frmScreen.txtFreeholderStreet.value);
		XML.SetAttribute("DISTRICT", frmScreen.txtFreeholderDistrict.value);
		XML.SetAttribute("TOWN", frmScreen.txtFreeholderTown.value);
		XML.SetAttribute("COUNTY", frmScreen.txtFreeholderCounty.value);
		XML.SetAttribute("COUNTRY", frmScreen.cboFreeholderCountry.value);
		XML.SetAttribute("POSTCODE", frmScreen.txtFreeholderPostcode.value);
		<% /* PB EP696/MAR1698 End */ %>
		XML.SetAttribute("ADDRESSTYPE", sROTAddressTypeValue) ;		
	}
}

function GetRotAddressTypeValue()
{
	var addressXML = null ;
	addressXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("ROTAddressType");
	var sRecipientList = addressXML.GetComboLists(document, sGroups);
	var sValidationType = "F"; //this is the validation type for Freeholder Address 
	var sROTAddressType = addressXML.GetTagText("LIST/LISTNAME/LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='" + sValidationType + "']/VALUEID");		

	return sROTAddressType ;	
}

function GetPropertyTenureList()
{
	XML = null ;
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("PropertyTenure");
	var bSuccess = false;	
	if (XML.GetComboLists(document, sGroups) == true)
	{
		bSuccess = true;
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboPropertyTenure,"PropertyTenure",true);
	}
}

function GetCountryList()
{
	XML = null ;
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("Country");
	var bSuccess = false;
	if (XML.GetComboLists(document, sGroups) == true)
	{
		bSuccess = true;
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboFreeholderCountry,"Country",true);
	}
}

function CommitChanges()
{
	var bSuccess = true;
	//if (m_sReadOnly != "1")
	if(frmScreen.onsubmit())
	{
		if (IsChanged()) bSuccess = ReportOnTitleAction();
	}
	else
		bSuccess = false;
	return(bSuccess);
}

function frmScreen.cboPropertyTenure.onchange()
{
	DisplayFreeHolder();
		
	// If Selected combo is either 'LeaseHold' or 'Fedual' then
	// enable the appropriate text fields, otherwise disable.		
	var sComboValue = frmScreen.cboPropertyTenure.value ;
	if (sComboValue != "")
	{
		var bReturn = xmlCombos.IsInComboValidationList(document,"PropertyTenure", sComboValue, ["L"]);
		if (bReturn == true)
		{
			DisablingTextFields("W");
			if (scScreenFunctions.GetRadioGroupValue(frmScreen, "LeaseTie") == 0) 
				scScreenFunctions.SetFieldState(frmScreen, "txtLeaseTieDetails", "D") ;						
			return ;
		}
	}

	DisablingTextFields("D");	
}

function DisablingTextFields(sAction)
{
	scScreenFunctions.SetCollectionState(divAddress, sAction);
	scScreenFunctions.SetCollectionState(divLeaseTie, sAction);
}

function frmScreen.optLeaseTie_Yes.onclick()
{
	
	scScreenFunctions.SetFieldState(frmScreen, "txtLeaseTieDetails", "W");
	frmScreen.txtLeaseTieDetails.select() ;	
}

function frmScreen.optLeaseTie_No.onclick()
{
	if (frmScreen.txtLeaseTieDetails.value != "")
	{
		if (confirm("Setting Lease Tie to no will remove all Lease Tie details. Do you wish to continue?"))		
			scScreenFunctions.SetFieldState(frmScreen, "txtLeaseTieDetails", "D");	
		else
			frmScreen.optLeaseTie_Yes.checked = true ;
	}
	else
		scScreenFunctions.SetFieldState(frmScreen, "txtLeaseTieDetails", "D");
}

function frmScreen.optInsuranceArrangements_Yes.onclick()
{
	scScreenFunctions.SetFieldState(frmScreen, "txtInsuranceArrangementDetails", "W") ;
	frmScreen.txtInsuranceArrangementDetails.select() ;
}

function frmScreen.optInsuranceArrangements_No.onclick ()
{
	if (frmScreen.txtInsuranceArrangementDetails.value != "")
	{
		if (confirm("Setting Insurance Arrangements to No will remove all Insurance Arrangements details\nDo you wish to continue?"))
			scScreenFunctions.SetFieldState(frmScreen, "txtInsuranceArrangementDetails", "D") ;
		else
			frmScreen.optInsuranceArrangements_Yes.checked = true ;
	}
	else
		scScreenFunctions.SetFieldState(frmScreen, "txtInsuranceArrangementDetails", "D") ;
}

function frmScreen.optPropertyMatch_Yes.onclick()
{
	scScreenFunctions.SetFieldState(frmScreen, "txtPropertyMatchDetails", "W") ;
	frmScreen.txtPropertyMatchDetails.select() ;
}

function frmScreen.optPropertyMatch_No.onclick()
{
	if (frmScreen.txtPropertyMatchDetails.value != "")
	{
		if (confirm("Setting Property Match to No will remove all Property Match details\nDo you wish to continue?"))
			scScreenFunctions.SetFieldState(frmScreen, "txtPropertyMatchDetails", "D") ;
		else
			frmScreen.optPropertyMatch_Yes.checked = true ;				
	}
	else
		scScreenFunctions.SetFieldState(frmScreen, "txtPropertyMatchDetails", "D") ;
}

function btnSubmit.onclick()
{	
	if (CommitChanges()) frmToAP400.submit() ;
}

function btnCancel.onclick()
{
	taskXML.SetAttribute("TASKSTATUS",m_EntryStatus); <%/* LDM 02/04/2006 MAR1624 reset to initial status */%> 
	scScreenFunctions.SetContextParameter(window,"idTaskXML",taskXML.XMLDocument.xml); <%/* LDM 02/05/2006 MAR1624 */%>
	frmToAP400.submit() ;
}

function btnNext.onclick()
{
	//DB BMIDS00862 - Move to ap450
	//if (CommitChanges()) frmToAP440.submit() ;
	if (CommitChanges()) frmToAP450.submit() ;
}

function DisplayFreeHolder()
{
	var sCurrentComboValue = frmScreen.cboPropertyTenure.value ;
	
	if (sCurrentComboValue == "" || sCurrentComboValue == "1" || sCurrentComboValue == "4")
	{
		//
		// Lease Holder - Collect Free Holder name
		//	
		scScreenFunctions.SetFieldState(frmScreen, "txtFreeholderName", "D") ;	
	}
	else
	{	
		//
		// Free Holder - Do not collect Free Holder name
		//
		scScreenFunctions.SetFieldState(frmScreen, "txtFreeholderName", "W") ;	
	}
}
-->
</script>
</body>
</HTML>

