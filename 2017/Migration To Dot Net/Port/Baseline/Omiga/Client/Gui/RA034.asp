<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      RA034.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Full Bureau CAIS Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		05/12/00	Screen Design
JR		27/12/00	Added Functionality
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MDC		08/10/2002	BMIDS00561	Add Date of Birth
TW      09/10/2002  Modified to incorporate client validation - SYS5115
MDC		19/11/2002	BMIDS01000	Display Balance and Monthly Repayment correctly
MDC		11/12/2002	BM0185		Only display repossess date if CML not blank
MV		11/02/2003	BM0294		Amended FillScreen() to display CreditCheckAccType Combo Value
BS		12/03/2003	BM0441/442	Amended FillScreen() to display correct details
GD		09/07/2003	BM0529		Check for account status of 8 AND 9.
MC		20/04/2004	BMIDS517	White space padded to the title text
KRW     13/10/2004  BM0568      Description not value to be displayed for Company type 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Full Bureau CAIS Summary  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>

<% //Span to keep tabbing within this screen %> 
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" validate ="onchange" mark>
<div id="divBackground" style="HEIGHT: 130px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 480px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		CCN Ref. No.
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtCreditCheckReferenceNumber" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>		
	<span style="TOP: 35px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Sequence Number
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtSequenceNumber" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Name
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtFullBureauName" maxlength="100" style="POSITION: absolute; WIDTH: 250px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 85px" class="msgLabel">
		Alias/Association
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtAliasAssociation" maxlength="50" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
	<% /* BMIDS00561 MDC 08/10/2002 */ %>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 110px" class="msgLabel">
		Date of Birth
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtDateOfBirth" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
	<% /* BMIDS00561 MDC 08/10/2002 - End */ %>
</div>
<div style="HEIGHT: 200px; LEFT: 10px; POSITION: absolute; TOP: 145px; WIDTH: 480px" class="msgGroup">
	<span style="TOP: 5px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Current Address</strong>
	</span>	
	<span style="TOP: 30px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Flat No.
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressFlatNumber" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 55px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		House Name & No.
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressHouseName" maxlength="26" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
		<span style="TOP: -3px; LEFT: 302px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressHouseNumber" maxlength="10" style="WIDTH: 50px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 80px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Street
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressStreet" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 105px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		District
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressDistrict" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 130px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Town
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressTown" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 155px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		County
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressCounty" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 180px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Postcode
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressPostcode" maxlength="8" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
</div>
<!--
	<span style="LEFT: 4px; POSITION: absolute; TOP: 82px" class="msgLabel">
		Current Address
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtCurrentAddress" maxlength="100" style="POSITION: absolute; WIDTH: 250px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
-->

<div style="HEIGHT: 130px; LEFT: 10px; POSITION: absolute; TOP: 350px; WIDTH: 480px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Company Type
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyType" maxlength="20" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Account Type
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountType" maxlength="20" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 250px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Own Group Account
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtOwnGroupAccount" maxlength="20" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Account Start Date
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISStartDate" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>	
	<span style="LEFT: 250px; POSITION: absolute; TOP: 85px" class="msgLabel">
		Account Updated Date
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISLastUpdatedDate" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 85px" class="msgLabel">
		Current Balance
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISCurrentBalance" maxlength="20" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 110px" class="msgLabel">
		Current Limit
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISCreditLimit" maxlength="20" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
</div>	
<div style="HEIGHT: 130px; LEFT: 10px; POSITION: absolute; TOP: 485px; WIDTH: 480px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Payment/Term
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISMonthlyPayment" maxlength="20" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
		<span style="LEFT: 202px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISRepayPeriod" maxlength="20" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Active/Settled
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtActiveSettled" maxlength="20" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>	
	<span style="LEFT: 250px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Settlement Date
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISSettlementDate" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Default Date
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISDefaultDate" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>	
	<span style="LEFT: 250px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Default Amount
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISCurrentDefaultBal" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 85px" class="msgLabel">
		Status C/W
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISStatus" maxlength="20" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>	
	<span style="LEFT: 250px; POSITION: absolute; TOP: 85px" class="msgLabel">
		Special Inst
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISSplInstrFlag" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 110px" class="msgLabel">
		Repossess Date
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISSettledDate" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>	
</div>
<div style="HEIGHT: 100px; LEFT: 10px; POSITION: absolute; TOP: 620px; WIDTH: 480px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Account History 1-2
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISAccountInstance1or2" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Account History in <br>the last * Months
		<span style="LEFT: 100px; POSITION: absolute; TOP: 5px">
			<input id="txtFBCAISStatusCountMonths" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Account History Status 3
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAISAccountInStatus3" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" tabindex=-1>
		</span>
	</span>
</div>
<div style="HEIGHT: 60px; LEFT: 10px; POSITION: absolute; TOP: 725px; WIDTH: 480px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 300px" class="msgLabel">
		Notices of Correction?
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="optNoticesOfCorrectionYes" name="NoticesOfCorrection" type="radio" value="1" tabindex=-1>
			<label for="optNoticesOfCorrectionYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
			<input id="optNoticesOfCorrectionNo" name="NoticesOfCorrection" type="radio" value="0" tabindex=-1>
			<label for="optNoticesOfCorrectionNo" class="msgLabel">No</label>
		</span> 
	</span>	
	<span style="LEFT: 320px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<input id="ddNoticesOfCorrection" type="button" style="HEIGHT: 32px; WIDTH: 32px" class ="msgDDButton">
	</span>		
	<span style="LEFT: 4px; POSITION: absolute; TOP: 35px">
		<input id="btnAddressMatch" value="Address Match" type="button" style="WIDTH: 100px" class="msgButton">
	</span>	
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 805px; WIDTH: 420px">
<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/RA034Attribs.asp" -->

<script language="JScript">
<!--
var scScreenFunctions;

var m_sCustomerNumber, m_sCustomerVersionNumber, m_sCustomerOrder ;
var m_sCreditCheckGuid, m_sFBBlockId, m_sFBHeaderSeqNo, m_sFBNameIndicator;
var m_sCreditCheckRefNo ;
var m_sXMLDataHeaderKeys, XMLDataHeaderKeys ;
var m_sFBResultsXML, XMLFBResults ; 

var m_iaFBCAISSequence = new Array() ; 
var m_iaFBResultsSequence = new Array() ; 
var m_iaCustomerAddressSequence = new Array() ;

var XML ;
var iCurrentBlock ;

var m_aArgArray = null ;
var m_aRequestAttribs = null ;

var sFBAddressIndicator, sFBNameIndicator, sFBBlockId, sFBHeaderSequence ;
var sFBTitle, sFBForename, sFBSurname, sFBInitials;
var sFBFlat, sFBHouseName, sFBHouseNumber, sFBStreet, sFBDistrict, sFBTown, sFBCounty, sFBPostCode ;
var sFBAssociationInfoType ;
var m_BaseNonPopupWindow = null;

var xmlCombos = null;
var sGroupList;
var sCompanyGroupList ;



<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	var sArguments		= window.dialogArguments ;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];	
	m_aArgArray			= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
	
	var sFBResults ;
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	m_sCreditCheckGuid		= m_aArgArray[0];
	m_sCustomerNumber		= m_aArgArray[1];
	m_sCustomerVersionNumber= m_aArgArray[2];
	m_sCustomerOrder		= m_aArgArray[3];
	m_sFBNameIndicator		= m_sCustomerOrder ;
	
	m_sFBBlockId			= m_aArgArray[4];
	m_sFBHeaderSeqNo		= m_aArgArray[5];
	m_sCreditCheckRefNo		= m_aArgArray[6];
	
	m_sFBResultsXML			= m_aArgArray[7]; 
	m_sXMLDataHeaderKeys	= m_aArgArray[8]; // Keys of table FBDataHeader
	
	m_aRequestAttribs		= m_aArgArray[9]; // Request attribs
	
	var sButtonList = new Array("Submit", "Previous", "Next");
	ShowMainButtons(sButtonList);

	scScreenFunctions.SetRadioGroupState(frmScreen, "NoticesOfCorrection", "D");
	
	XMLFBResults = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLFBResults.LoadXML(m_sFBResultsXML) ;
	
	XMLDataHeaderKeys = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLDataHeaderKeys.LoadXML();
	
	xmlCombos = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	sGroupList = new Array("CreditCheckAccType"); 
	sCompanyGroupList = new Array("CreditCheckCompanyType");// KRW 13/10/04 BM0568   

	
			
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function btnSubmit.onclick()
{
	window.close();
}

function btnNext.onclick()
{
	ClearScreen();
	++iCurrentBlock ;
	FillScreen();
}

function btnPrevious.onclick()
{
	ClearScreen();
	--iCurrentBlock ;
	FillScreen();
}

function frmScreen.btnAddressMatch.onclick()
{
	var m_saFBAddress = new Array();
	var m_saAddress = new Array();

	// Fill the array of Bureau Address
	m_saFBAddress[0] = frmScreen.txtCurrentAddressFlatNumber.value  ;
	m_saFBAddress[1] = frmScreen.txtCurrentAddressHouseName.value ;
	m_saFBAddress[2] = frmScreen.txtCurrentAddressHouseNumber.value ;
	m_saFBAddress[3] = frmScreen.txtCurrentAddressStreet.value ;
	m_saFBAddress[4] = frmScreen.txtCurrentAddressTown.value ;
	m_saFBAddress[5] = frmScreen.txtCurrentAddressDistrict.value ;
	m_saFBAddress[6] = frmScreen.txtCurrentAddressCounty.value ;
	m_saFBAddress[7] = frmScreen.txtCurrentAddressPostcode.value ;
	
	// Fill the array with Address from the table 'CustomerAddress'
	if(m_iaCustomerAddressSequence[iCurrentBlock] != null)
	{
		xmlAddress.SelectTagListItem(m_iaCustomerAddressSequence[iCurrentBlock]);
		m_saAddress[0] = xmlAddress.GetTagText("FLATNUMBER");
		m_saAddress[1] = xmlAddress.GetTagText("BUILDINGORHOUSENAME");
		m_saAddress[2] = xmlAddress.GetTagText("BUILDINGORHOUSENUMBER");
		m_saAddress[3] = xmlAddress.GetTagText("STREET");
		m_saAddress[4] = xmlAddress.GetTagText("TOWN");
		m_saAddress[5] = xmlAddress.GetTagText("DISTRICT");
		m_saAddress[6] = xmlAddress.GetTagText("COUNTY");
		m_saAddress[7] = xmlAddress.GetTagText("POSTCODE");
	}
	else
	{
		m_saAddress[0] = '';
		m_saAddress[1] = '';
		m_saAddress[2] = '';
		m_saAddress[3] = '';
		m_saAddress[4] = '';
		m_saAddress[5] = '';
		m_saAddress[6] = '';
		m_saAddress[7] = '';
	}
		
	// Call the screen 'AddressMatch' RA038
	var ArrayArguments = new Array();
	ArrayArguments[0] = frmScreen.txtCreditCheckReferenceNumber.value ;
	ArrayArguments[1] = frmScreen.txtFullBureauName.value   ;
	ArrayArguments[2] = m_saFBAddress ;
	ArrayArguments[3] = m_saAddress ;

	scScreenFunctions.DisplayPopup(window, document, "RA038.asp", ArrayArguments, 630, 390);
}

function GetCustomerAddressData()
{
	xmlAddress = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	var xmlRequest, xmlCustomerAddressList ;
	
	// Get current address data by default; 
	xmlRequest = xmlAddress.CreateRequestTagFromArray(m_aRequestAttribs, "SEARCH");
	xmlCustomerAddressList = xmlAddress.CreateActiveTag("CUSTOMERADDRESSLIST");
	xmlAddress.CreateActiveTag("CUSTOMERADDRESS");
	xmlAddress.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber) ;
	xmlAddress.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	xmlAddress.CreateTag("ADDRESSTYPE", 1);
	
	// fetch previous adresses, if required, i.e., any of FBADDRESSINDICATOR <> 'C' in dataheader
	var sCondition = ".//FULLBUREAUDATAHEADER[CREDITCHECKGUID='" + m_sCreditCheckGuid + "'"
					 + " && FBBLOCKID = '" + m_sFBBlockId  + "'" 
                     + " && FBHEADERSEQUENCE ='" + m_sFBHeaderSeqNo + "'" 
                     + " && FBADDRESSINDICATOR != 'C']" ;
                     
    var xmlNodeList = XMLDataHeaderKeys.XMLDocument.selectNodes(sCondition);
    if(xmlNodeList.length > 0)
    {
		xml.ActiveTag = xmlCustomerAddressList ;
		xmlAddress.CreateActiveTag("CUSTOMERADDRESS");
		xmlAddress.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber) ;
		xmlAddress.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
		xmlAddress.CreateTag("ADDRESSTYPE", 3);	
    }   
    xmlAddress.RunASP(document,"FindCustomerAddressList.asp");
	
	if(!xmlAddress.IsResponseOK())
	{
		alert('Error retreiving Customer Address data');
	}
	
	xmlAddress.CreateTagList("CUSTOMERADDRESS");
}

function frmScreen.ddNoticesOfCorrection.onclick()
{
	var m_saFBAddress = new Array();

	// Fill the array of Bureau Address
	m_saFBAddress[0] = frmScreen.txtCurrentAddressFlatNumber.value ;
	m_saFBAddress[1] = frmScreen.txtCurrentAddressHouseName.value ;
	m_saFBAddress[2] = frmScreen.txtCurrentAddressHouseNumber.value ;
	m_saFBAddress[3] = frmScreen.txtCurrentAddressStreet.value ;
	m_saFBAddress[4] = frmScreen.txtCurrentAddressTown.value ;
	m_saFBAddress[5] = frmScreen.txtCurrentAddressDistrict.value ;
	m_saFBAddress[6] = frmScreen.txtCurrentAddressCounty.value ;
	m_saFBAddress[7] = frmScreen.txtCurrentAddressPostcode.value ;
	
	// Call the screen RA033
	var ArrayArguments = new Array();
	ArrayArguments[0] = frmScreen.txtCreditCheckReferenceNumber.value ;
	ArrayArguments[1] = frmScreen.txtFullBureauName.value  ;
	ArrayArguments[2] = m_saFBAddress ;
	ArrayArguments[3] = XML.ActiveTag.xml ;
	
	scScreenFunctions.DisplayPopup(window, document, "RA033.asp", ArrayArguments, 427, 590);
}


/** FUNCTIONS **/
function PopulateScreen()
{
	XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var XMLTemp = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequest, xmlCurrentBlock ;

	// Build the request and call the method to fetch data
	xmlRequest = XML.CreateRequestTagFromArray(m_aRequestAttribs, "SEARCH");
	
	XML.CreateActiveTag("FULLBUREAUCAIS");
	XML.CreateTag("CREDITCHECKGUID", m_sCreditCheckGuid);
	XML.CreateTag("FBBLOCKID", m_sFBBlockId);
	
	XML.ActiveTag =	xmlRequest ;
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	
	// Add FullBureau Data Header keys (xml) to the request ; this is optional tag in the request
	XMLTemp.LoadXML(m_sXMLDataHeaderKeys);
	xmlRequest.appendChild(XMLTemp.XMLDocument.documentElement); //Append the xmlDataHeaderKeys to the Request

	XML.RunASP(document,"GetCurrentFullBureauCAISSummary.asp");
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	
	if(ErrorReturn[0] == false) 
	{
		alert('Error retreiving CAIS Summary');
		return ;
	}
	
	GetCustomerAddressData() ;
	
	// The method returns data for all the FBHeaderSequenceNos with this FBBlockId and 
	// CreditCheckGuid. Go to the correct row in the returned XML, and populate the screen
	XML.CreateTagList("FULLBUREAUCAIS");
	XMLFBResults.CreateTagList("FBRESULTS");
	
	var iCount, iFBCAISSCount ;
	var sFBBLockId, sFBHeaderSeqNo ;
	
	// Build arrays with indexes of the node corresponding to the current customer and BlockId ,
	// for both FullBureauResults and CAIS
	for (iCount=0; iCount < XMLFBResults.ActiveTagList.length; ++iCount)
	{
		XMLFBResults.SelectTagListItem(iCount);
		sFBBLockId		= XMLFBResults.GetTagText("FBBLOCKID") ;
		sFBHeaderSeqNo	= XMLFBResults.GetTagText("FBHEADERSEQUENCE") ;
		
		// From  FullBureau results, show only the CAIS data
		if(sFBBLockId == m_sFBBlockId)
		{
			m_iaFBResultsSequence[m_iaFBResultsSequence.length] = iCount ;
			for(iFBCAISSCount = 0; iFBCAISSCount < XML.ActiveTagList.length; ++iFBCAISSCount)
			{
				XML.SelectTagListItem(iFBCAISSCount);
				
				// if this is corresponding CAIS node
				if(sFBBLockId == XML.GetTagText("FBBLOCKID") && sFBHeaderSeqNo == XML.GetTagText("FBHEADERSEQUENCE"))
				{
					m_iaFBCAISSequence[m_iaFBCAISSequence.length] = iFBCAISSCount ;
					GetRelatedAddress(XMLFBResults.GetTagText("FBADDRESSINDICATOR")) ;
					break;
				}
			}
			// Find the current CAIS node to be displayed
			if(sFBBLockId == m_sFBBlockId && sFBHeaderSeqNo == m_sFBHeaderSeqNo)
			{
				 iCurrentBlock = m_iaFBCAISSequence.length - 1 ;
			}
		}
	}
		
	FillScreen() ;
}

function FillScreen()
{
	// Select the nodes (in FullBureau results, CAIS) from which the data is to be displayed in the screen
	XMLFBResults.SelectTagListItem(m_iaFBResultsSequence[iCurrentBlock]);
	XML.SelectTagListItem(m_iaFBCAISSequence[iCurrentBlock]);
		
	frmScreen.txtCreditCheckReferenceNumber.value	= m_sCreditCheckRefNo ;
	frmScreen.txtFullBureauName.value 				= XMLFBResults.GetTagText("FBTITLE") + " " +
													  XMLFBResults.GetTagText("FBFORENAME") + " " +
													  XMLFBResults.GetTagText("FBSURNAME");
													  							  
	frmScreen.txtSequenceNumber.value				= XMLFBResults.GetTagText("FBHEADERSEQUENCE");	
	
	m_sFBBlockId		= XML.GetTagText("FBBLOCKID");
	m_sFBHeaderSeqNo	= XML.GetTagText("FBHEADERSEQUENCE");
	
	/******Other details******/
	
	// Fill Current Address
	FillCurrentAddress() ;
	
	//Get selected row for retrieving FBAssociationInfoType
	GetSelectedRowData();
	
	switch (sFBAssociationInfoType)
	{
		case "S":
			frmScreen.txtAliasAssociation.value = "Alias";
			break;
		case "A":
			frmScreen.txtAliasAssociation.value = "Association";
			break;
		default:
			frmScreen.txtAliasAssociation.value = "";	
	}
	<% /* BMIDS00561 MDC 08/10/2002 */ %>
	frmScreen.txtDateOfBirth.value				= XML.GetTagText("FBDATEOFBIRTH");
	<% /* BMIDS00561 MDC 08/10/2002 - End */ %>
	
	<% /* BM0568 KRW 13/10/04 */ %>
	xmlCombos.GetComboLists(document, sCompanyGroupList) ;
	
	var sComboCompanyType                       = XML.GetTagText("FBCAISCOMPANYTYPE");
	frmScreen.txtCompanyType.value				=  xmlCombos.GetComboDescriptionForValidation("CreditCheckCompanyType",sComboCompanyType);
	<% /* BM0568 KRW 13/10/04 - End */ %>
	
	xmlCombos.GetComboLists(document, sGroupList) ;
	
	var sComboValidation = XML.GetTagText("FBCAISACCOUNTTYPE")
	frmScreen.txtAccountType.value				=  xmlCombos.GetComboDescriptionForValidation("CreditCheckAccType",sComboValidation);
	
	frmScreen.txtOwnGroupAccount.value			= XML.GetTagText("FBCAISOWNDATAFLAG");
	frmScreen.txtFBCAISStartDate.value			= XML.GetTagText("FBCAISSTARTDATE");
	frmScreen.txtFBCAISLastUpdatedDate.value	= XML.GetTagText("FBCAISLASTUPDATEDDATE");
	<% /* BMIDS01000 MDC 19/11/2002
	frmScreen.txtFBCAISCurrentBalance.value		= XML.GetTagText("FBCAISCURRENTDEFAULTBAL");  */ %>
	<% /* BS BM0441/442 12/03/2003
	frmScreen.txtFBCAISCurrentBalance.value		= XML.GetTagText("FBCAISBALANCE"); */ %>
	frmScreen.txtFBCAISMonthlyPayment.value		= XML.GetTagText("FBCAISMONTHLYPAYMENT");
	<% /* BMIDS01000 MDC 19/11/2002 - End */ %>
	frmScreen.txtFBCAISCreditLimit.value		= XML.GetTagText("FBCAISCREDITLIMIT");
	frmScreen.txtFBCAISRepayPeriod.value		= XML.GetTagText("FBCAISREPAYPERIOD");
	<% /* BS BM0441/442 12/03/2003
	frmScreen.txtActiveSettled.value			= XML.GetTagText("FBCAISSETTLEDDATE");
	frmScreen.txtFBCAISSettlementDate.value		= XML.GetTagText("FBCAISSETTLEDDATE");
	frmScreen.txtFBCAISDefaultDate.value		= XML.GetTagText("FBCAISDEFAULTDATE");
	frmScreen.txtFBCAISCurrentDefaultBal.value	= XML.GetTagText("FBCAISCURRENTDEFAULTBAL"); */ %>
	frmScreen.txtFBCAISStatus.value				= XML.GetTagText("FBCAISACCOUNTSTATUS");
	
	<% /* BS BM0441/442 12/03/2003 */ %>
	var AccStatus
	AccStatus = XML.GetTagText("FBCAISACCOUNTSTATUS"); 
	<% // GD BM0529 09/07/2003 %>
	if ((AccStatus.substr(0,1) == "8") || (AccStatus.substr(0,1) == "9"))
	{
		if (XML.GetTagText("FBCAISCURRENTDEFAULTBAL") > 0)
		{
			//Default
			frmScreen.txtFBCAISCurrentBalance.value		= XML.GetTagText("FBCAISCURRENTDEFAULTBAL");
			frmScreen.txtActiveSettled.value			= "Default";
			frmScreen.txtFBCAISSettlementDate.value		= "";
			frmScreen.txtFBCAISDefaultDate.value		= XML.GetTagText("FBCAISSETTLEDDATE");
			frmScreen.txtFBCAISCurrentDefaultBal.value	= XML.GetTagText("FBCAISBALANCE");
		}
		else
		{
			//Satisfied default
			frmScreen.txtFBCAISCurrentBalance.value		= "Satisfied";
			frmScreen.txtActiveSettled.value			= "Default";
			frmScreen.txtFBCAISSettlementDate.value		= "";
			frmScreen.txtFBCAISDefaultDate.value		= XML.GetTagText("FBCAISSETTLEDDATE");
			frmScreen.txtFBCAISCurrentDefaultBal.value	= XML.GetTagText("FBCAISBALANCE");		}
	}
	else
	{
		if (XML.GetTagText("FBCAISSETTLEDDATE") == "")
		{
			//Active CAIS
			frmScreen.txtFBCAISCurrentBalance.value		= XML.GetTagText("FBCAISBALANCE");
			frmScreen.txtActiveSettled.value			= "Active";
			frmScreen.txtFBCAISSettlementDate.value		= "";
			frmScreen.txtFBCAISDefaultDate.value		= "";
			frmScreen.txtFBCAISCurrentDefaultBal.value	= "";
		}
		else
		{
			//Settled CAIS
			frmScreen.txtFBCAISCurrentBalance.value		= XML.GetTagText("FBCAISBALANCE");
			frmScreen.txtActiveSettled.value			= "Settled";
			frmScreen.txtFBCAISSettlementDate.value		= XML.GetTagText("FBCAISSETTLEDDATE");
			frmScreen.txtFBCAISDefaultDate.value		= "";
			frmScreen.txtFBCAISCurrentDefaultBal.value	= "";		
		}
	}
	<% /* BS BM0441/442 End 12/03/2003 */ %>
	
	frmScreen.txtFBCAISSplInstrFlag.value		= XML.GetTagText("FBCAISSPECIALINSTRFLAG");
	<% /* BM0185 MDC 11/12/2002
	frmScreen.txtFBCAISSettledDate.value		= XML.GetTagText("FBCAISSETTLEDDATE");  */ %>
	if(XML.GetTagText("FBCAISCMLADDRESSTYPE") == "")
		frmScreen.txtFBCAISSettledDate.value	= "";
	else
		frmScreen.txtFBCAISSettledDate.value	= XML.GetTagText("FBCAISSETTLEDDATE");
	<% /* BM0185 MDC 11/12/2002 - End */ %>
		
	frmScreen.txtFBCAISAccountInstance1or2.value= XML.GetTagText("FBCAISACCOUNTINSTATUS1OR3");
	frmScreen.txtFBCAISStatusCountMonths.value	= XML.GetTagText("FBCAISSTATUSCOUNTINONTHS");
	frmScreen.txtFBCAISAccountInStatus3.value	= XML.GetTagText("FBCAISACCOUNTINSTATUS3");
	
	// enable or disable the button (Notices Of Correction)
	if (XML.GetTagText("FBNOTICEOFCORRECTIONREF") == "") 
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen,"NoticesOfCorrection","0");
		scScreenFunctions.DisableDrillDown(frmScreen.ddNoticesOfCorrection);
	}
	else
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen,"NoticesOfCorrection", "1");
		scScreenFunctions.EnableDrillDown(frmScreen.ddNoticesOfCorrection);
	}
	
	// Disable navigation buttons appropriately
	DisableNavigationButtons() ;
}

function GetRelatedAddress(sFBAddressIndicator)
{
	switch (sFBAddressIndicator)
	{
		case 'C':			
			m_iaCustomerAddressSequence[m_iaCustomerAddressSequence.length] = 
									 (xmlAddress.ActiveTagList.length == 1)? 0 : null ;
			break ;
		case 'P':
			m_iaCustomerAddressSequence[m_iaCustomerAddressSequence.length] =
									 (xmlAddress.ActiveTagList.length == 2)? 1 : null ;
			break ;
		case '3':
			m_iaCustomerAddressSequence[m_iaCustomerAddressSequence.length] =
									 (xmlAddress.ActiveTagList.length == 3)? 2 : null ;
			break ;
		default:
			m_iaCustomerAddressSequence[m_iaCustomerAddressSequence.length] =
									 (xmlAddress.ActiveTagList.length == 4)? 3 : null ;
	}
}

function FillCurrentAddress()
{	

	frmScreen.txtCurrentAddressFlatNumber.value = XMLFBResults.GetTagText("FBFLAT");
	frmScreen.txtCurrentAddressHouseName.value 	= XMLFBResults.GetTagText("FBHOUSENAME");
	frmScreen.txtCurrentAddressHouseName.value 	= XMLFBResults.GetTagText("FBHOUSENUMBER"); 
	frmScreen.txtCurrentAddressStreet.value 	= XMLFBResults.GetTagText("FBSTREET");
	frmScreen.txtCurrentAddressTown.value 		= XMLFBResults.GetTagText("FBTOWN");
	frmScreen.txtCurrentAddressDistrict.value 	= XMLFBResults.GetTagText("FBDISTRICT");
	frmScreen.txtCurrentAddressCounty.value 	= XMLFBResults.GetTagText("FBCOUNTY");
	frmScreen.txtCurrentAddressPostcode.value 	= XMLFBResults.GetTagText("FBPOSTCODE");
}

function ClearScreen()
{
	frmScreen.txtCreditCheckReferenceNumber.value	= "" ;
	frmScreen.txtFullBureauName.value				= "" ;
	frmScreen.txtSequenceNumber.value				= "" ;	
	
	// Current Address
	frmScreen.txtCurrentAddressFlatNumber.value		= "";
	frmScreen.txtCurrentAddressHouseName.value 		= "";
	frmScreen.txtCurrentAddressHouseName.value 		= "";
	frmScreen.txtCurrentAddressStreet.value 		= "";
	frmScreen.txtCurrentAddressTown.value 			= "";
	frmScreen.txtCurrentAddressDistrict.value 		= "";
	frmScreen.txtCurrentAddressCounty.value 		= "";
	frmScreen.txtCurrentAddressPostcode.value 		= "";

	// Other Details
	frmScreen.txtAliasAssociation.value			= "";
	frmScreen.txtCompanyType.value				= "";
	frmScreen.txtAccountType.value				= "";
	frmScreen.txtOwnGroupAccount.value			= "";
	frmScreen.txtFBCAISStartDate.value			= "";
	frmScreen.txtFBCAISLastUpdatedDate.value	= "";
	frmScreen.txtFBCAISCurrentBalance.value		= "";
	frmScreen.txtFBCAISCreditLimit.value		= "";
	frmScreen.txtFBCAISRepayPeriod.value		= "";
	frmScreen.txtActiveSettled.value			= "";
	frmScreen.txtFBCAISSettlementDate.value		= "";
	frmScreen.txtFBCAISDefaultDate.value		= "";
	frmScreen.txtFBCAISCurrentDefaultBal.value	= "";
	frmScreen.txtFBCAISStatus.value				= "";
	frmScreen.txtFBCAISSplInstrFlag.value		= "";
	frmScreen.txtFBCAISSettledDate.value		= "";
	frmScreen.txtFBCAISAccountInstance1or2.value = "";
	frmScreen.txtFBCAISStatusCountMonths.value	= "";
	frmScreen.txtFBCAISAccountInStatus3.value	= "";
	
}

function GetSelectedRowData()
{
	//Retrieve data from selected row on RA030
	for (var iSelectedCount=0; iSelectedCount < XMLFBResults.ActiveTagList.length; iSelectedCount++)
	{
		XMLFBResults.SelectTagListItem(iSelectedCount) ;
		sFBBlockId = XMLFBResults.GetTagText ("FBBLOCKID") ;
		sFBHeaderSequence = XMLFBResults.GetTagText ("FBHEADERSEQUENCE") ;
		sFBNameIndicator = XMLFBResults.GetTagText ("FBNAMEINDICATOR") ;
		
		//if this is the selected row		
		if ((sFBBlockId == m_sFBBlockId) && (sFBHeaderSequence == m_sFBHeaderSeqNo) && (sFBNameIndicator == m_sFBNameIndicator))
		{
			sFBAddressIndicator = XMLFBResults.GetTagText ("FBADDRESSINDICATOR") ;
			sFBTitle			= XMLFBResults.GetTagText ("FBTITLE") ;			
			sFBForename			= XMLFBResults.GetTagText ("FBFORENAME") ;		
			sFBSurname			= XMLFBResults.GetTagText ("FBSURNAME") ;		
			sFBFlat				= XMLFBResults.GetTagText ("FBFLAT") ;			
			sFBHouseName		= XMLFBResults.GetTagText ("FBHOUSENAME") ;		
			sFBHouseNumber		= XMLFBResults.GetTagText ("FBHOUSENUMBER") ;	
			sFBStreet			= XMLFBResults.GetTagText ("FBSTREET") ;			
			sFBDistrict			= XMLFBResults.GetTagText ("FBDISTRICT") ;		
			sFBTown				= XMLFBResults.GetTagText ("FBTOWN") ;			
			sFBCounty			= XMLFBResults.GetTagText ("FBCOUNTY") ;			
			sFBPostCode			= XMLFBResults.GetTagText ("FBPOSTCODE") ;
			
			//Now find a match within the remaining XML
			GetFBAssociationInfoType() ;
			break ;
		}
	}
}

function GetFBAssociationInfoType()
{
	var iLoop ;

	//Exit loop when found a match
	for (iLoop=0; iLoop < XMLFBResults.ActiveTagList.length; iLoop++)
	{
		XMLFBResults.SelectTagListItem(iLoop);
		
		if ( (sFBAddressIndicator== XMLFBResults.GetTagText ("FBADDRESSINDICATOR"))&& 
			(sFBNameIndicator	 == XMLFBResults.GetTagText ("FBNAMEINDICATOR"))	&&
			(sFBTitle			 == XMLFBResults.GetTagText ("FBTITLE"))			&&
			(sFBForename		 == XMLFBResults.GetTagText ("FBFORENAME"))			&&
			(sFBSurname			 == XMLFBResults.GetTagText ("FBSURNAME"))			&&
			(sFBFlat			 == XMLFBResults.GetTagText ("FBFLAT"))				&&
			(sFBHouseName		 == XMLFBResults.GetTagText ("FBHOUSENAME"))		&&
			(sFBHouseNumber		 == XMLFBResults.GetTagText ("FBHOUSENUMBER"))		&&
			(sFBStreet			 == XMLFBResults.GetTagText ("FBSTREET"))			&&
			(sFBDistrict		 == XMLFBResults.GetTagText ("FBDISTRICT"))			&&
			(sFBTown			 == XMLFBResults.GetTagText ("FBTOWN"))				&&
			(sFBCounty			 == XMLFBResults.GetTagText ("FBCOUNTY"))			&&
			(sFBPostCode		 == XMLFBResults.GetTagText ("FBPOSTCODE")) )	
		{
			sFBAssociationInfoType = XMLFBResults.GetTagText ("FBASSOCIATIONINFOTYPE");
			break;
		}
	}
}

function DisableNavigationButtons()
{
	if(m_iaFBCAISSequence.length == 1 || m_iaFBCAISSequence.length == 0)
	{
		DisableMainButton("Previous");
		DisableMainButton("Next");
	}
	else 
	{
		if(iCurrentBlock == 0)
		{
			DisableMainButton("Previous");
			EnableMainButton("Next");
		}
		else 
		{
			if(parseInt(iCurrentBlock) + 1 == m_iaFBCAISSequence.length)
			{
				EnableMainButton("Previous");
				DisableMainButton("Next");
			}
			else
			{
				EnableMainButton("Previous");
				EnableMainButton("Next");
			}
		}
	}
}

-->
</script>

</body>
</html>




