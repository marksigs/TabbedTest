<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      ra032.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Public Info Summary Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:
Prog	Date		Descrip
JR		5/12/00		Screen Design
JR		20/12/00	Added functionlity
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MC		19/04/2004	BMIDS517	White space padded to the title text
KRW     13/10/04    BM0568      Description not value to be displayed for debt type 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
	<title>Public Information Summary  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate ="onchange">
<div id="divBackground" style="HEIGHT: 80px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 430px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		CCN Ref. No.
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtCreditCheckReferenceNumber" name="CreditCheckReferenceNumber" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Sequence Number
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtSequenceNumber" name="Sequence" maxlength="30" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Name
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtFullBureauName" name="Name" maxlength="55" style="POSITION: absolute; WIDTH: 250px" class="msgReadOnly" readonly>
		</span>
	</span>
</div>
<div style="HEIGHT: 200px; LEFT: 10px; POSITION: absolute; TOP: 95px; WIDTH: 430px" class="msgGroup">
	<span style="TOP: 5px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Current Address</strong>
	</span>	
	<span style="TOP: 30px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Flat No.
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressFlatNumber" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 55px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		House Name & No.
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressHouseName" maxlength="26" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
		<span style="TOP: -3px; LEFT: 329px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressHouseNumber" maxlength="10" style="WIDTH: 50px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 80px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Street
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressStreet" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 105px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		District
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressDistrict" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 130px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Town
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressTown" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 155px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		County
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressCounty" maxlength="20" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="TOP: 180px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Postcode
		<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="txtCurrentAddressPostcode" maxlength="8" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgReadOnly" readonly>
		</span>
	</span>
</div>
<div style="HEIGHT: 105px; LEFT: 10px; POSITION: absolute; TOP: 300px; WIDTH: 430px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Debt Type
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtDebtType" name="DebtType" maxlength="40" style="POSITION: absolute; WIDTH: 250px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Debt Amount
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtDebtAmount" name="DebtType" maxlength="15" style="POSITION: absolute; WIDTH: 200px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Date Registered
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtFBPublicInfoRegistrationDate" name="DebtType" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" readonly>
		</span>
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 85px" class="msgLabel">
		Date Satisfied
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtFBPublicInfoSatisfiedDate" name="DebtType" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" readonly>
		</span>
	</span>
</div>
<div style="HEIGHT: 70px; LEFT: 10px; POSITION: absolute; TOP: 410px; WIDTH: 430px" class="msgGroup">
	<span id=optNoticesOfCorrection name=NoticesOfCorrection style="TOP: 20px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Notices of Correction
		<span style="TOP: 3px; LEFT: 130px; POSITION: ABSOLUTE">
			<input id="optNoticesOfCorrectionYes" name="NoticesOfCorrection" type="radio" value="1" tabindex=-1>
			<label for="optNoticesOfCorrectionYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: 3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="optNoticesOfCorrectionNo" name="NoticesOfCorrection" type="radio" value="0" tabindex=-1>
			<label for="optNoticesOfCorrectionNo" class="msgLabel">No</label>
		</span> 
	</span>	
	
	<span style="TOP: 15px; LEFT: 270px; POSITION: ABSOLUTE">
			<input id="ddNoticesOfCorrection" type="button" style="HEIGHT: 32px; WIDTH: 32px" class="msgDDButton">
	</span>
	<span style="TOP: 45px; LEFT: 10px; POSITION: ABSOLUTE">
			<input id="btnAddressMatch" value="AddressMatch" type="button" style="WIDTH: 90px" class="msgButton">
	</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 500px; WIDTH: 420px">
<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/ra032Attribs.asp" -->


<script language="JScript">
<!--
var scScreenFunctions;

var m_sCustomerNumber, m_sCustomerVersionNumber, m_sCustomerOrder ;
var m_sCreditCheckGuid, m_sFBBlockId, m_sFBHeaderSeqNo ;
var m_sCreditCheckRefNo ;
var m_sXMLDataHeaderKeys, XMLDataHeaderKeys ;
var m_sFBResultsXML, XMLFBResults ; ;

var m_iaFBPublicInfoSequence = new Array(); //Array containing the header sequence numbers of public info for current customer
var m_iaFBResultsSequence = new Array(); // Array Containing the header sequence numbers of FB Results for current customer
var m_iaCustomerAddressSequence = new Array(); // Array containing the customer address sequence for current customer

var XML, xmlAddress ;
var iCurrentBlock ;

var m_aArgArray = null ;
var m_aRequestAttribs = null ;
var m_BaseNonPopupWindow = null;

var xmlCombos = null;
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

	scScreenFunctions.SetRadioGroupState(frmScreen, "NoticesOfCorrection", "D");
	
	XMLFBResults = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLFBResults.LoadXML(m_sFBResultsXML) ;
	
	XMLDataHeaderKeys = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLDataHeaderKeys.LoadXML();
	
	xmlCombos = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("CreditCheckDebtType");
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
	m_saFBAddress[0] = frmScreen.txtCurrentAddressFlatNumber.value  ;
	m_saFBAddress[1] = frmScreen.txtCurrentAddressHouseName.value  ;
	m_saFBAddress[2] = frmScreen.txtCurrentAddressHouseNumber.value  ;
	m_saFBAddress[3] = frmScreen.txtCurrentAddressStreet.value  ;
	m_saFBAddress[4] = frmScreen.txtCurrentAddressTown.value  ;
	m_saFBAddress[5] = frmScreen.txtCurrentAddressDistrict.value  ;
	m_saFBAddress[6] = frmScreen.txtCurrentAddressCounty.value  ;
	m_saFBAddress[7] = frmScreen.txtCurrentAddressPostcode.value  ;
	
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
	ArrayArguments[1] = frmScreen.txtFullBureauName.value  ;
	ArrayArguments[2] = m_saFBAddress ;
	ArrayArguments[3] = m_saAddress ;

	scScreenFunctions.DisplayPopup(window, document, "RA038.asp", ArrayArguments, 630, 390);
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
	XML.CreateActiveTag("FULLBUREAUPUBLICINFO");
	XML.CreateTag("CREDITCHECKGUID", m_sCreditCheckGuid);
	XML.CreateTag("FBBLOCKID", m_sFBBlockId);
	
	XML.ActiveTag =	xmlRequest ;
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	
	// Add FullBuyreau Data Header keys (xml) to the request ; this is optional tag in the request
	XMLTemp.LoadXML(m_sXMLDataHeaderKeys);
	xmlRequest.appendChild(XMLTemp.XMLDocument.documentElement); //Append the xmlDataHeaderKeys to the Request

	XML.RunASP(document,"GetCurrentPublicInfoSummary.asp");
	
	if(!XML.IsResponseOK())
	{
		alert('Error retreiving public info summary');
		return ;
	}
	
	GetCustomerAddressData();
	
	// The method returns data for all the FBHeaderSequenceNos with this FBBlockId and 
	// CreditCheckGuid. Go to the correct row in the returned XML, and populate the screen
	XML.CreateTagList("FULLBUREAUPUBLICINFO");
	XMLFBResults.CreateTagList("FBRESULTS");
	
	var iCount, iFBPublicInfoCount ;
	var sFBBLockId, sFBHeaderSeqNo ;
	
	// Build arrays with indexes of the node corresponding to the current customer and BlockId ,
	// for both FullBureauResults and VotersRoll
	for (iCount=0; iCount < XMLFBResults.ActiveTagList.length; ++iCount)
	{
		XMLFBResults.SelectTagListItem(iCount);
		sFBBLockId		= XMLFBResults.GetTagText("FBBLOCKID") ;
		sFBHeaderSeqNo	= XMLFBResults.GetTagText("FBHEADERSEQUENCE") ;
		
		// From  FullBureau results, show only the Public Info Data data
		if(sFBBLockId == m_sFBBlockId)
		{
			m_iaFBResultsSequence[m_iaFBResultsSequence.length] = iCount ;			
			for(iFBPublicInfoCount = 0; iFBPublicInfoCount < XML.ActiveTagList.length; ++iFBPublicInfoCount)
			{
				XML.SelectTagListItem(iFBPublicInfoCount);
				// if this is corresponding PublicRoll node
				if(sFBBLockId == XML.GetTagText("FBBLOCKID") && sFBHeaderSeqNo == XML.GetTagText("FBHEADERSEQUENCE"))
				{
					m_iaFBPublicInfoSequence[m_iaFBPublicInfoSequence.length] = iFBPublicInfoCount ;
					GetRelatedAddress(XMLFBResults.GetTagText("FBADDRESSINDICATOR")) ;
					break ;
				}
			}		
			// Find the current public info node to be displayed
			if(sFBBLockId == m_sFBBlockId && sFBHeaderSeqNo == m_sFBHeaderSeqNo) iCurrentBlock = m_iaFBPublicInfoSequence.length - 1 ;
		}
	}
	FillScreen() ;
}

function FillScreen()
{
	// Select the nodes (in FullBureau results, PublicInfo) from which the data is to be displayed in the screen
	XMLFBResults.SelectTagListItem(m_iaFBResultsSequence[iCurrentBlock]);
	XML.SelectTagListItem(m_iaFBPublicInfoSequence[iCurrentBlock]);
		
	frmScreen.txtCreditCheckReferenceNumber.value	= m_sCreditCheckRefNo ;
	frmScreen.txtFullBureauName.value 				= XMLFBResults.GetTagText("FBTITLE") + " " +
													  XMLFBResults.GetTagText("FBFORENAME") + " " +
													  XMLFBResults.GetTagText("FBSURNAME");
													  
	frmScreen.txtSequenceNumber.value				= XMLFBResults.GetTagText("FBHEADERSEQUENCE");	
	
	m_sFBBlockId		= XML.GetTagText("FBBLOCKID");
	m_sFBHeaderSeqNo	= XML.GetTagText("FBHEADERSEQUENCE");
	
	// Other Details
	//frmScreen.txtDebtType.value = XML.GetTagText("FBPUBLICINFODATATYPE");
	
	var sComboDebtType = XML.GetTagText("FBPUBLICINFODATATYPE"); // KRW BM0568 13/10/04
	frmScreen.txtDebtType.value	 	        =  xmlCombos.GetComboDescriptionForValidation("CreditCheckDebtType",sComboDebtType);
	
	
	frmScreen.txtDebtAmount.value = XML.GetTagText("FBPUBLICINFOAMOUNTINPOUNDS") + "." +
									XML.GetTagText("FBPUBLICINFOAMOUNTINPENCE");
	frmScreen.txtFBPublicInfoRegistrationDate.value = XML.GetTagText("FBPUBLICINFOREGISTRATIONDATE");
	frmScreen.txtFBPublicInfoSatisfiedDate.value = XML.GetTagText("FBPUBLICINFOSATISFIEDDATE");
			
	// Fill Current Address
	FillCurrentAddress() ;
	
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
	frmScreen.txtDebtAmount.value					= "";
	frmScreen.txtDebtType.value						= "";
	frmScreen.txtFBPublicInfoRegistrationDate.value = "";
	frmScreen.txtFBPublicInfoSatisfiedDate.value	= "";
	
}

function DisableNavigationButtons()
{
	if(m_iaFBPublicInfoSequence.length == 1 || m_iaFBPublicInfoSequence.length == 0)
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
			if(parseInt(iCurrentBlock) + 1 == m_iaFBPublicInfoSequence.length)
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




