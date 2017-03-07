<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      ra036.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Credit Check Summary Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JR		5/12/00		Screen Design
SR		15/01/01	included functionality
LD		23/05/02	SYS4727 Use cached versions of frame functions
HMA     10/12/03    BMIDS675  Enable ddNoticesOfCorrection button and add OnClick function.
						      Add title. Remove extra PostCode field.
MC		19/04/2004	BMIDS517	White space padded to the title text	
HMA     01/12/2004  BMIDS954    Correct display of Address Type and Own Group ID.			      
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
*/

%>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"> 
	<title>Previous Applications CAPS Summary  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>
<% //Span to keep tabbing within this screen %> 
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate ="onchange">

<div id="divBackground" style="HEIGHT: 55px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 450px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		CCN Ref. No.
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtCreditCheckReferenceNumber" name="CreditCheckReferenceNumber" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 280px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Sequence
		<span style="LEFT: 60px; POSITION: absolute; TOP: -3px">
			<input id="txtSequenceNumber" name="Sequence" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Name
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtName" name="Name" maxlength="55" style="POSITION: absolute; WIDTH: 250px" class="msgReadOnly" readonly>
		</span>
	</span>
</div>
<div style="HEIGHT: 251px; LEFT: 10px; POSITION: absolute; TOP: 70px; WIDTH: 450px" class="msgGroup">
	<span style="TOP: 5px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Current Address</strong>
	</span>	
	<span style="TOP: 34px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Flat No.
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtFlatNo" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 58px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		House Name & No.
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtHouseName" maxlength="26" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
		<span style="TOP: -3px; LEFT: 331px; POSITION: ABSOLUTE">
			<input id="txtHouseNo" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 82px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Street
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtStreet" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 106px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		District
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtDistrict" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 130px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Town
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtTown" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 154px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		County
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtCounty" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	
	<span style="TOP: 178px; LEFT: 10px; POSITION: absolute" class="msgLabel">
		Post Code
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtPostCode" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>	
	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 202px" class="msgLabel">
		Date of Birth
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAPSDateofBirth" name="DateofBirth" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 226px" class="msgLabel">
		Time At Address
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAPSTimeAtAddress" name="TimeAtAddress" maxlength="4" style="POSITION: absolute; WIDTH: 60px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 208px; POSITION: absolute; TOP: 226px" class="msgLabel">
		Address Type
		<span style="LEFT: 72px; POSITION: absolute; TOP: -3px">
			<input id="txtAddressType" name="AddressType" maxlength="40" style="POSITION: absolute; WIDTH: 163px" class="msgReadOnly" readonly>
		</span>
	</span>
</div>
<div style="HEIGHT: 127px; LEFT: 10px; POSITION: absolute; TOP: 326px; WIDTH: 450px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Search Date
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAPSApplicationDate" name="Search Date" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 10; POSITION: absolute; TOP: 34" class="msgLabel">
		Amount
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAPSAmount" name="Amount" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 308px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Term
		<span style="LEFT: 50px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAPSTermInMonths" name="Term" maxlength="5" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 58px" class="msgLabel">
		Own Group Id
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAPSOwnApplicationInd" name="OwnGroupId" maxlength="9" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 82px" class="msgLabel">
		Company Type
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtFBCAPSCompanyType" name="CompanyType" maxlength="40" style="POSITION: absolute; WIDTH: 130px" class="msgReadOnly" readonly>
		</span>
	</span>	
	<span style="LEFT: 10px; POSITION: absolute; TOP: 106px" class="msgLabel">
		Account Type
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountType" name="CompanyType" maxlength="40" style="POSITION: absolute; WIDTH: 130px" class="msgReadOnly" readonly>
		</span>
	</span>
</div>
<div style="HEIGHT: 70px; LEFT: 10px; POSITION: absolute; TOP: 458px; WIDTH: 450px" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Notices of Correction
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="optNoticesofCorrectionYes" name="NoticesOfCorrection" type="radio" value="1" tabindex=-1>
			<label for="optNoticesofCorrectionYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="optNoticesofCorrectionNo" name="NoticesOfCorrection" type="radio" value="0" tabindex=-1>
			<label for="optNoticesofCorrectionNo" class="msgLabel">No</label>
		</span> 
		<span style="TOP: -6px; LEFT: 260px; POSITION: ABSOLUTE">
			<input id="ddNoticesOfCorrection" type="button" style="HEIGHT: 32px; WIDTH: 32px" class="msgDDButton">
		</span>		
	</span>
	
	<span style="TOP: 40px; LEFT: 10px; POSITION: ABSOLUTE">
		<input id="btnAddressMatch" value="Address Match" type="button" style="WIDTH: 90px" class="msgButton">
	</span>

</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 540px; WIDTH: 450px">
<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/ra036Attribs.asp" -->

<script language="JScript">
<!--

var scScreenFunctions;

var m_sCustomerNumber, m_sCustomerVersionNumber, m_sCustomerOrder ;
var m_sCreditCheckGuid, m_sFBBlockId, m_sFBHeaderSeqNo;
var m_sCreditCheckRefNo ;
var m_sXMLDataHeaderKeys, XMLDataHeaderKeys ;
var m_sFBResultsXML, XMLFBResults ; ;

var m_iaFBCAPSSequence = new Array(); // Array Containing the header sequence numbers of CAPS for current customer
var m_iaFBResultsSequence = new Array(); // Array Containing the header sequence numbers of FB Results for current customer
var m_iaCustomerAddressSequence = new Array(); // Array containing the customer address sequence for current customer

var XML, xmlAddress ;
var iCurrentBlock ;

var m_aArgArray = null ;
var m_aRequestAttribs = null ;

var xmlCombos ;
var m_BaseNonPopupWindow = null;


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
	
	m_sFBBlockId			= m_aArgArray[4];
	m_sFBHeaderSeqNo		= m_aArgArray[5];
	m_sCreditCheckRefNo		= m_aArgArray[6];
	
	m_sFBResultsXML			= m_aArgArray[7];
	m_sXMLDataHeaderKeys	= m_aArgArray[8]; // Keys of table FBDataHeader
	
	m_aRequestAttribs		= m_aArgArray[9]; // Request attribs
		
	var sButtonList = new Array("Submit", "Previous", "Next");
	ShowMainButtons(sButtonList);

	scScreenFunctions.SetRadioGroupToDisabled(frmScreen, "NoticesOfCorrection") ;
	
	XMLFBResults = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLFBResults.LoadXML(m_sFBResultsXML) ;
	
	XMLDataHeaderKeys = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLDataHeaderKeys.LoadXML();

	// Fetch the combo values and the respective validations for the follwing combos ;
	// later used to populate value name in the form
	xmlCombos = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("JointAppInd", "ApplicationType", "CAPSCompanyType");  // BMIDS954
	xmlCombos.GetComboLists(document, sGroupList) ;
		
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
	m_saFBAddress[0] = frmScreen.txtFlatNo.value  ;
	m_saFBAddress[1] = frmScreen.txtHouseName.value  ;
	m_saFBAddress[2] = frmScreen.txtHouseNo.value  ;
	m_saFBAddress[3] = frmScreen.txtStreet.value  ;
	m_saFBAddress[4] = frmScreen.txtTown.value  ;
	m_saFBAddress[5] = frmScreen.txtDistrict.value  ;
	m_saFBAddress[6] = frmScreen.txtCounty.value  ;
	m_saFBAddress[7] = frmScreen.txtPostCode.value  ;
	
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
	ArrayArguments[1] = frmScreen.txtName.value  ;
	ArrayArguments[2] = m_saFBAddress ;
	ArrayArguments[3] = m_saAddress ;

	scScreenFunctions.DisplayPopup(window, document, "RA038.asp", ArrayArguments, 630, 390);
}
//BMIDS675
function frmScreen.ddNoticesOfCorrection.onclick()
{
	var m_saFBAddress = new Array();

	// Fill the array of Bureau Address
	m_saFBAddress[0] = frmScreen.txtFlatNo.value ;
	m_saFBAddress[1] = frmScreen.txtHouseName.value ;
	m_saFBAddress[2] = frmScreen.txtHouseNo.value ;
	m_saFBAddress[3] = frmScreen.txtStreet.value ;
	m_saFBAddress[4] = frmScreen.txtTown.value ;
	m_saFBAddress[5] = frmScreen.txtDistrict.value ;
	m_saFBAddress[6] = frmScreen.txtCounty.value ;
	m_saFBAddress[7] = frmScreen.txtPostCode.value ;
	
	// Call the screen RA033
	var ArrayArguments = new Array();
	ArrayArguments[0] = frmScreen.txtCreditCheckReferenceNumber.value ;
	ArrayArguments[1] = frmScreen.txtName.value  ;
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
	XML.CreateActiveTag("FULLBUREAUCAPS");
	XML.CreateTag("CREDITCHECKGUID", m_sCreditCheckGuid);
	XML.CreateTag("FBBLOCKID", m_sFBBlockId);
	
	XML.ActiveTag =	xmlRequest ;
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	
	// Add FullBureau Data Header keys (xml) to the request ; this is optional tag in the request
	XMLTemp.LoadXML(m_sXMLDataHeaderKeys);
	xmlRequest.appendChild(XMLTemp.XMLDocument.documentElement); //Append the xmlDataHeaderKeys to the Request

	XML.RunASP(document,"GetCurrentCAPSSummary.asp");
	
	if(!XML.IsResponseOK())
	{
		alert('Error retreiving public info summary');
		return ;
	}	
	
	GetCustomerAddressData();
	
	// The method returns data for all the FBHeaderSequenceNos with this FBBlockId and 
	// CreditCheckGuid. Go to the correct row in the returned XML, and populate the screen
	XML.CreateTagList("FULLBUREAUCAPS");
	XMLFBResults.CreateTagList("FBRESULTS");
	
	var iCount, iFBCAPSCount ;
	var sFBBLockId, sFBHeaderSeqNo ;
	
	// Build arrays with indexes of the node corresponding to the current customer and BlockId ,
	// for both FullBureauResults and CAPS
	for (iCount=0; iCount < XMLFBResults.ActiveTagList.length; ++iCount)
	{
		XMLFBResults.SelectTagListItem(iCount);
		sFBBLockId		= XMLFBResults.GetTagText("FBBLOCKID") ;
		sFBHeaderSeqNo	= XMLFBResults.GetTagText("FBHEADERSEQUENCE") ;
		
		// From  FullBureau results, show only the CAPS data
		if(sFBBLockId == m_sFBBlockId)
		{
			m_iaFBResultsSequence[m_iaFBResultsSequence.length] = iCount ;			
			for(iFBCAPSCount = 0; iFBCAPSCount < XML.ActiveTagList.length; ++iFBCAPSCount)
			{
				XML.SelectTagListItem(iFBCAPSCount);
				// if this is corresponding CAPS node
				if(sFBBLockId == XML.GetTagText("FBBLOCKID") && sFBHeaderSeqNo == XML.GetTagText("FBHEADERSEQUENCE"))
				{
					m_iaFBCAPSSequence[m_iaFBCAPSSequence.length] = iFBCAPSCount ;
					GetRelatedAddress(XMLFBResults.GetTagText("FBADDRESSINDICATOR")) ;
					break ;
				}
			}		
			// Find the current CAPS node to be displayed
			if(sFBBLockId == m_sFBBlockId && sFBHeaderSeqNo == m_sFBHeaderSeqNo) iCurrentBlock = m_iaFBCAPSSequence.length - 1 ;
		}
	}
	FillScreen() ;
}

function FillScreen()
{
	// Select the nodes (in FullBureau results, CAPS) from which the data is to be displayed in the screen
	XMLFBResults.SelectTagListItem(m_iaFBResultsSequence[iCurrentBlock]);
	XML.SelectTagListItem(m_iaFBCAPSSequence[iCurrentBlock]);

	frmScreen.txtCreditCheckReferenceNumber.value	= m_sCreditCheckRefNo ;
	frmScreen.txtName.value							= XMLFBResults.GetTagText("FBTITLE") + " " +
													  XMLFBResults.GetTagText("FBFORENAME") + " " +
													  XMLFBResults.GetTagText("FBSURNAME");
													  
	frmScreen.txtSequenceNumber.value				= XMLFBResults.GetTagText("FBHEADERSEQUENCE");	
	
	m_sFBBlockId		= XML.GetTagText("FBBLOCKID");
	m_sFBHeaderSeqNo	= XML.GetTagText("FBHEADERSEQUENCE");
		
	// Fill Current Address
	FillCurrentAddress() ;
	
	// Other Details
	frmScreen.txtFBCAPSDateofBirth.value			= XML.GetTagText("FBCAPSDATEOFBIRTH") ;
	frmScreen.txtFBCAPSTimeAtAddress.value			= XML.GetTagText("FBCAPSTIMEATADDRESS") ;
	frmScreen.txtFBCAPSApplicationDate.value		= XML.GetTagText("FBCAPSAPPLICATIONDATE") ;
	frmScreen.txtFBCAPSAmount.value					= XML.GetTagText("FBCAPSAMOUNT") ;
	frmScreen.txtFBCAPSTermInMonths.value			= XML.GetTagText("FBCAPSTERMINMONTHS") ;
	frmScreen.txtFBCAPSOwnApplicationInd.value		= XML.GetTagText("FBCAPSOWNAPPLICATIONIND") ;  // BMIDS954
	
	// Populate derived values
	var sComboValidation ;
	sComboValidation = XML.GetTagText("FBCAPSJOINTAPPLICATIONIND");
	frmScreen.txtAddressType.value = (sComboValidation == "") ? sComboValidation : 
								xmlCombos.GetComboDescriptionForValidation("JointAppInd", sComboValidation);   // BMIDS954

	sComboValidation = XML.GetTagText("FBCAPSCOMPANYTYPE");
	frmScreen.txtFBCAPSCompanyType.value = (sComboValidation == "") ? sComboValidation : 
								xmlCombos.GetComboDescriptionForValidation("CAPSCompanyType", sComboValidation);

	sComboValidation = XML.GetTagText("FBCAPSAPPLICATIONTYPE");
	frmScreen.txtAccountType.value = (sComboValidation == "") ? sComboValidation : 
								xmlCombos.GetComboDescriptionForValidation("ApplicationType", sComboValidation);
	
	// enable or disable the button (Notices Of Correction)
	if (XML.GetTagText("FBNOTICEOFCORRECTIONREF") == "") 
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen,"NoticesOfCorrection","0");
		scScreenFunctions.DisableDrillDown(frmScreen.ddNoticesOfCorrection)
	}
	else
	{
		scScreenFunctions.SetRadioGroupValue(frmScreen,"NoticesOfCorrection", "1");
		scScreenFunctions.EnableDrillDown(frmScreen.ddNoticesOfCorrection);
	}
	
	// Disable navigation buttons appropriately
	DisableNavigationButtons() ;
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
	frmScreen.txtFlatNo.value		= XMLFBResults.GetTagText("FBFLAT");
	frmScreen.txtHouseName.value	= XMLFBResults.GetTagText("FBHOUSENAME");
	frmScreen.txtHouseNo.value		= XMLFBResults.GetTagText("FBHOUSENUMBER"); 
	frmScreen.txtStreet.value		= XMLFBResults.GetTagText("FBSTREET");
	frmScreen.txtTown.value			= XMLFBResults.GetTagText("FBTOWN");
	frmScreen.txtDistrict.value		= XMLFBResults.GetTagText("FBDISTRICT");
	frmScreen.txtCounty.value		= XMLFBResults.GetTagText("FBCOUNTY");
	frmScreen.txtPostCode.value		= XMLFBResults.GetTagText("FBPOSTCODE");
}

function ClearScreen()
{
	frmScreen.txtCreditCheckReferenceNumber.value	= "" ;
	frmScreen.txtName.value							= "" ;
	frmScreen.txtSequenceNumber.value				= "" ;	
	
	// Current Address
	frmScreen.txtFlatNo.value						= "" ;
	frmScreen.txtHouseName.value					= "" ;
	frmScreen.txtHouseNo.value						= "" ;
	frmScreen.txtStreet.value						= "" ;
	frmScreen.txtTown.value							= "" ;
	frmScreen.txtDistrict.value						= "" ;
	frmScreen.txtCounty.value						= "" ;
	frmScreen.txtPostCode.value						= "" ;
	
	// Other Details
	frmScreen.txtFBCAPSDateofBirth.value			= "" ;
	frmScreen.txtFBCAPSTimeAtAddress.value			= "" ;
	frmScreen.txtAddressType.value					= "" ;
	frmScreen.txtFBCAPSApplicationDate.value		= "" ;
	frmScreen.txtFBCAPSAmount.value					= "" ;
	frmScreen.txtFBCAPSTermInMonths.value			= "" ;
	frmScreen.txtFBCAPSOwnApplicationInd.value		= "" ;
	frmScreen.txtFBCAPSCompanyType.value			= "" ;
	frmScreen.txtAccountType.value					= "" ;	
}

function DisableNavigationButtons()
{
	if(m_iaFBCAPSSequence.length == 1 || m_iaFBCAPSSequence.length == 0)
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
			if(parseInt(iCurrentBlock) + 1 == m_iaFBCAPSSequence.length)
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

//-->
</script>

</body>
</html>




