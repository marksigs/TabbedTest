<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/* 
Workfile:      dc295.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Other Insurance Company Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MC		24/03/2000	Original
MH		15/05/2000  SYS0706 currency symbol and date validation
MC		21/06/2000	SYS0866 Standardise format of balance and payment fields
CL		05/03/01	SYS1920 Read only functionality added
JLD		10/12/01	SYS2806 completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MO		14/11/2002	BMIDS00807 - Made change to take the date from the app server not the client

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History

Prog	Date		AQR		Description
MF		04/08/2005	MARS20	Routes back to DC201
MF		08/08/2005	MAR20	Route depending on Global Parameters
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Forms Here */ %>
<form id="frmToDC300" method="post" action="dc300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC240" method="post" action="dc240.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC201" method="post" action="DC201.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC270" method="post" action="DC270.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 175px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Other Insurance Company Details</strong>
	</span>
	<div style="TOP: 34px; LEFT: 38px; HEIGHT: 120px; WIDTH: 528px; POSITION: ABSOLUTE" class="msgGroupLight">
		<span style="TOP: 15px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Insurance Company
			<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
				<input id="txtInsuranceCompany" maxlength="45" style="WIDTH: 300px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<span style="TOP: 40px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Buildings Sum Insured Amount
			<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<label style="TOP: 0px; LEFT: -10px; POSITION: absolute" class="msgCurrency"></label>
				<input id="txtBuildingsSIAmt" maxlength="6" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<span style="TOP: 65px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Policy Number
			<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
				<input id="txtPolicyNumber" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<span style="TOP: 90px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Next Renewal Date
			<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
				<input id="txtRenewalDate" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
	</div>


</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 250px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!--  #include FILE="attribs/dc295attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var InsCoXML = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var scScreenFunctions;
var m_blnReadOnly = false;



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
	FW030SetTitles("Other Insurance Company","DC295",scScreenFunctions);

	RetrieveContextData();
	Initialise();
	
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC295");
	scScreenFunctions.SetFocusToFirstField(frmScreen);	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function PopulateScreen()
{
	<% /* populate fields in form */ %>
	frmScreen.txtInsuranceCompany.value = InsCoXML.GetTagText("INSURANCECOMPANYNAME");
	frmScreen.txtBuildingsSIAmt.value = InsCoXML.GetTagText("BUILDINGSSUMINSURED");
	frmScreen.txtPolicyNumber.value = InsCoXML.GetTagText("POLICYNUMBER");
	frmScreen.txtRenewalDate.value = InsCoXML.GetTagText("RENEWALDATE");
}

function GetOtherInsuranceCompanyData()
{
var strNewInsuranceCompany = "";

	<% /* Retrieve Credit Check data */ %>
	InsCoXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	with (InsCoXML)
	{
		<% /* Create request block */ %>
		CreateRequestTag(window, "SEARCH");
		CreateActiveTag("SEARCH");
		<% /* CreateActiveTag("OTHERINSURANCECOMPANY"); */ %>
		CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		
		<% /* Run the corresponding ASP page script to retrieve the data */ %>
		RunASP(document, "GetOtherInsuranceCompany.asp");
		<% /* Check response */ %>
		if (IsResponseOK())
		{
			<% /* Check if creating a new record */ %>
			strNewInsuranceCompany = GetTagText("NEWINSURANCECOMPANY")
			if (strNewInsuranceCompany != "")
				m_sMetaAction = "CreateInsuranceCompany";
			else
				m_sMetaAction = "UpdateInsuranceCompany";
				
			PopulateScreen();
			frmScreen.txtInsuranceCompany.focus();
		}
		else
			DisableMainButton("Submit");
	}
}

function CheckForBuildingsAndContentsQuote()
{
	// see dc300.asp for example
	return false;
}

function Initialise()
{

	scScreenFunctions.SetCurrency(window,frmScreen);
	
	if (CheckForBuildingsAndContentsQuote())
	{
		alert("Buildings and Contents Quote exists");
		frmToDC300.submit();
	}
	else
		GetOtherInsuranceCompanyData();
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
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","C00010014");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
}

function btnSubmit.onclick()
{
var RetVal = true;
	
	if(frmScreen.onsubmit())
	{
		if(IsChanged())
			if (IsRenewalDateValid())
				RetVal = CommitScreen();
			else RetVal=false;
	
		if(RetVal)
		{
			if(scScreenFunctions.CompletenessCheckRouting(window))
				frmToGN300.submit();
			else{
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
				var bNewPropertySummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");			
				if(bNewPropertySummary)	
					frmToDC201.submit();
				else
					frmToDC300.submit();
			}
		}
	}
}

function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
		var bNewPropertySummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");			
		if(bNewPropertySummary)
			frmToDC201.submit();
		else{
			var bThirdPartySummary = XML.GetGlobalParameterBoolean(document,"ThirdPartySummary");			
			if(bThirdPartySummary)			
				frmToDC240.submit();
			else
				frmToDC270.submit();
		}
	}
}

function IsRenewalDateValid()
{
	var dteRenewal;
	var dteToday;
	var dteNextYear;
	
	if (frmScreen.txtRenewalDate=="") return true;
	
	dteRenewal=scScreenFunctions.GetDateObject(frmScreen.txtRenewalDate);
	<% /* MO - BMIDS00807 */ %>
	dteToday = scScreenFunctions.GetAppServerDate(true);
	<% /* dteToday=new Date(); */ %>
	
	if (dteRenewal < dteToday) 
	{
		alert ("Date must be within the next 12 months");
		frmScreen.txtRenewalDate.focus();
		return false;
	}
	dteNextYear=new Date(dteToday.getYear()+1,dteToday.getMonth(),dteToday.getDate())	
	
	if (dteRenewal > dteNextYear) 
	{
		alert ("Date must be within the next 12 months");
		frmScreen.txtRenewalDate.focus();
		return false;
	}
	return true;
}

function CommitScreen()
{
<%	/*	XML structure:
		<REQUEST USERID=?,USERTYPE=?,UNIT=?,ACTION="CREATE" or "UPDATE">
			Other Insurance Company Details
		<REQUEST>
	*/
%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	var TagREQUEST = XML.CreateRequestTag(window,null);
	var TagRequestType = null;
	var TagREQUEST = null;
	var sASPFile;

	if(m_sMetaAction == "CreateInsuranceCompany")
	{
		TagRequestType = XML.CreateActiveTag("CREATE");
		// scScreenFunctions.SetContextParameter(window,"idCustomerLockIndicator","1");
		sASPFile = "CreateOtherInsuranceCompany.asp";
	}

	if (m_sMetaAction == "UpdateInsuranceCompany")
	{
		TagRequestType = XML.CreateActiveTag("UPDATE");
		sASPFile = "UpdateOtherInsuranceCompany.asp";
	}

	if(TagRequestType != null)
	{

	<%	/*	XML structure:
			<OTHERINSURANCECOMPANY>
				fields
			</OTHERINSURANCECOMPANY>
		*/
	%>	XML.ActiveTag = TagRequestType;
		XML.CreateActiveTag("OTHERINSURANCECOMPANY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("INSURANCECOMPANYNAME", frmScreen.txtInsuranceCompany.value);
		XML.CreateTag("BUILDINGSSUMINSURED", frmScreen.txtBuildingsSIAmt.value);
		XML.CreateTag("POLICYNUMBER", frmScreen.txtPolicyNumber.value);
		XML.CreateTag("RENEWALDATE", frmScreen.txtRenewalDate.value);
	
		// 		XML.RunASP(document,sASPFile);
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,sASPFile);
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		if(XML.IsResponseOK())
			return true;
	}
	return false;
}

-->
</script>
</body>
</html>


