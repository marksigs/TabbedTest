<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      RA037.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Aliases and Associations Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		07/12/00	Screen Design
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MC		19/04/2004	BMIDS517	White space padded to the title text
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
	<title>Aliases and Associations Summary  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>

<% //Span to keep tabbing within this screen %> 
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" validate ="onchange" mark>
<div id="divBackground" style="HEIGHT: 80px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 450px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		CCN Ref. No.
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtCreditCheckReferenceNumber" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Sequence Number
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtSequenceNumber" maxlength="5" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 58px" class="msgLabel">
		Name
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtName" maxlength="50" style="POSITION: absolute; WIDTH: 250px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
</div>	

<div id="divAddress" style="HEIGHT: 190px; LEFT: 10px; POSITION: absolute; TOP: 95px; WIDTH: 450px" class="msgGroup">
	<span id="spnCurrentAddress" style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
		<strong>Current Address</strong>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel">
		Flat No. 
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtFlatNo" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
		
	<span style="LEFT: 4px; POSITION: absolute; TOP: 48px" class="msgLabel">
		House Name &amp; No. 
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtHouseName" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
		<span style="LEFT: 301px; POSITION: absolute; TOP: -3px">
			<input id="txtHouseNo" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 72px" class="msgLabel">
		Street 
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtStreet" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 96px" class="msgLabel">
		District 
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtDistrict" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 120px" class="msgLabel">
		Town
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtTown" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 144px" class="msgLabel">
		County
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtCounty" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>	
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 168px" class="msgLabel">
		Post Code
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtPostCode" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>	
</div>

<div id="divAssociationData" style="HEIGHT: 180px; LEFT: 10px; POSITION: absolute; TOP: 290px; WIDTH: 450px" class="msgGroup">		
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Alias Or Association
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtAliasOrAssociation" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>	
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 34px" class="msgLabel">
		Alias/Association Name
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtAliasOrAssociationName" maxlength="40" style="POSITION: absolute; WIDTH: 200px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 58px" class="msgLabel">
		Date Application Made
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtFBAssociationInfoDate" maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>	
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 82px" class="msgLabel">
		Search Type 
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtSearchType" maxlength="40" style="POSITION: absolute; WIDTH: 200px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>	
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 106px" class="msgLabel">
		Application Type
		<span style="LEFT: 150px; POSITION: absolute; TOP: -3px">
			<input id="txtApplicationType" maxlength="40" style="POSITION: absolute; WIDTH: 200px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>	
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 130px; WIDTH: 300px" class="msgLabel">
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
	
	<span style="LEFT: 320px; POSITION: absolute; TOP: 124px" class="msgLabel">
		<input id="ddNoticesOfCorrection" type="button" style="HEIGHT: 32px; WIDTH: 32px" class ="msgDDButton">
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 150px">
		<input id="btnAddressMatch" value="Address Match" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
</div>
	
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 490px; WIDTH: 450px">
<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/RA037Attribs.asp" -->

<script language="JScript">
<!--
var scScreenFunctions;

var m_aArgArray = null ;
var m_aRequestAttribs = null ;

var XML, xmlAddress ;
var iCurrentBlock ;

var m_sCustomerNumber, m_sCustomerVersionNumber, m_sCustomerOrder ;
var m_sCreditCheckGuid, m_sFBBlockId, m_sFBHeaderSeqNo;
var m_sCreditCheckRefNo ;

var m_sXMLDataHeaderKeys, XMLDataHeaderKeys ;
var m_sFBResultsXML, XMLFBResults ; 

var m_iaFBAliasSequence = new Array(); // Array Containing the header sequence numbers of voters roll for current customer
var m_iaFBResultsSequence = new Array(); // Array Containing the header sequence numbers of FB Results for current customer
var m_iaCustomerAddressSequence = new Array(); // Array containing the customer address sequence for current customer

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
	
	scScreenFunctions.SetRadioGroupState(frmScreen, "NoticesOfCorrection", "D");
	
	XMLFBResults = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLFBResults.LoadXML(m_sFBResultsXML) ;
	
	XMLDataHeaderKeys = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLDataHeaderKeys.LoadXML();
	
	// Fetch the combo values and the respective validations for the follwing combos ;
	// later used to populate value name in the form
	xmlCombos = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("AssocInfoType", "CAPSCompanyType", "InfoSource");
	xmlCombos.GetComboLists(document,sGroupList)
		
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
	m_saFBAddress[0] = frmScreen.txtFlatNo.value ;
	m_saFBAddress[1] = frmScreen.txtHouseName.value ;
	m_saFBAddress[2] = frmScreen.txtHouseNo.value ;
	m_saFBAddress[3] = frmScreen.txtStreet.value ;
	m_saFBAddress[4] = frmScreen.txtTown.value ;
	m_saFBAddress[5] = frmScreen.txtDistrict.value ;
	m_saFBAddress[6] = frmScreen.txtCounty.value ;
	m_saFBAddress[7] = frmScreen.txtPostCode.value ;
	
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
	ArrayArguments[1] = frmScreen.txtName.value ;
	ArrayArguments[2] = m_saFBAddress ;
	ArrayArguments[3] = m_saAddress ;

	scScreenFunctions.DisplayPopup(window, document, "RA038.asp", ArrayArguments, 630, 390);
}

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
	ArrayArguments[1] = frmScreen.txtName.value ;
	ArrayArguments[2] = m_saFBAddress ;
	ArrayArguments[3] = XML.ActiveTag.xml ;
	
	scScreenFunctions.DisplayPopup(window, document, "RA033.asp", ArrayArguments, 427, 590);
}

/* FUNCTIONS */
function PopulateScreen()
{
	XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var XMLTemp = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequest, xmlCurrentBlock ;

	// Build the request and call the method to fetch data
	xmlRequest = XML.CreateRequestTagFromArray(m_aRequestAttribs, "SEARCH");
	XML.CreateActiveTag("FULLBUREAUASSOCIATIONS");
	XML.CreateTag("CREDITCHECKGUID", m_sCreditCheckGuid);
	XML.CreateTag("FBBLOCKID", m_sFBBlockId);
	
	XML.ActiveTag =	xmlRequest ;
	XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);
	
	// Add FullBuyreau Data Header keys (xml) to the request ; this is optional tag in the request
	XMLTemp.LoadXML(m_sXMLDataHeaderKeys);
	xmlRequest.appendChild(XMLTemp.XMLDocument.documentElement); //Append the xmlDataHeaderKeys to the Request

	XML.RunASP(document,"GetCurrentAliasSummary.asp");
	
	if(!XML.IsResponseOK())
	{
		alert('Error retreiving public info summary');
		return ;
	}	
	
	GetCustomerAddressData();

	// The method returns data for all the FBHeaderSequenceNos with this FBBlockId and 
	// CreditCheckGuid. Go to the correct row in the returned XML, and populate the screen
	XML.CreateTagList("FULLBUREAUASSOCIATIONS");
	XMLFBResults.CreateTagList("FBRESULTS");
	
	var iCount, iFBVotersRollCount ;
	var sFBBLockId, sFBHeaderSeqNo ;
	
	// Build arrays with indexes of the node corresponding to the current customer and BlockId ,
	// for both FullBureauResults and VotersRoll
	for (iCount=0; iCount < XMLFBResults.ActiveTagList.length; ++iCount)
	{
		XMLFBResults.SelectTagListItem(iCount);
		sFBBLockId		= XMLFBResults.GetTagText("FBBLOCKID") ;
		sFBHeaderSeqNo	= XMLFBResults.GetTagText("FBHEADERSEQUENCE") ;
		
		// From  FullBureau results, show only the Voters Roll data
		if(sFBBLockId == m_sFBBlockId)
		{
			m_iaFBResultsSequence[m_iaFBResultsSequence.length] = iCount ;			
			for(iFBAliasCount = 0; iFBAliasCount < XML.ActiveTagList.length; ++iFBAliasCount)
			{
				XML.SelectTagListItem(iFBAliasCount);
				// if this is corresponding Assocations node
				if(sFBBLockId == XML.GetTagText("FBBLOCKID") && sFBHeaderSeqNo == XML.GetTagText("FBHEADERSEQUENCE"))
				{
					m_iaFBAliasSequence[m_iaFBAliasSequence.length] = iFBAliasCount ;
					GetRelatedAddress(XMLFBResults.GetTagText("FBADDRESSINDICATOR")) ;
					break ;
				}
			}		
			// Find the current voters roll node to be displayed
			if(sFBBLockId == m_sFBBlockId && sFBHeaderSeqNo == m_sFBHeaderSeqNo) iCurrentBlock = m_iaFBAliasSequence.length - 1 ;
		}
	}
	FillScreen() ;
}

function ClearScreen()
{
	frmScreen.txtCreditCheckReferenceNumber.value = "" ;
	frmScreen.txtName.value = "" ;
	frmScreen.txtSequenceNumber.value = "" ;
	
	frmScreen.txtFlatNo.value		= "";
	frmScreen.txtHouseName.value	= "";
	frmScreen.txtHouseNo.value		= ""; 
	frmScreen.txtStreet.value		= "";
	frmScreen.txtTown.value			= "";
	frmScreen.txtDistrict.value		= "";
	frmScreen.txtCounty.value		= "";
	frmScreen.txtPostCode.value		= "";
	
	frmScreen.txtAliasOrAssociation.value		= "";
	frmScreen.txtAliasOrAssociationName.value	= "";
	frmScreen.txtFBAssociationInfoDate.value	= "";
	frmScreen.txtSearchType.value				= "";
	frmScreen.txtApplicationType.value			= "";
}

function FillScreen()
{
	// Select the nodes (in FullBureau results, Alias) from which the data is to be displayed in the screen
	XMLFBResults.SelectTagListItem(m_iaFBResultsSequence[iCurrentBlock]);
	XML.SelectTagListItem(m_iaFBAliasSequence[iCurrentBlock]);
		
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
	frmScreen.txtAliasOrAssociationName.value = XML.GetTagText("FBASSOCIATIONTITLE") + " " +
												XML.GetTagText("FBASSOCIATIONFORENAME") + " " +
												XML.GetTagText("FBASSOCIATIONSURNAME") ;
	
	frmScreen.txtFBAssociationInfoDate.value  = XML.GetTagText("FBASSOCIATIONINFODATE");
	
	var sComboValidation ;
	sComboValidation = XML.GetTagText("FBASSOCIATIONINFOTYPE");
	frmScreen.txtAliasOrAssociation.value = (sComboValidation == "") ? sComboValidation : 
								xmlCombos.GetComboDescriptionForValidation("AssocInfoType", sComboValidation);

	sComboValidation = XML.GetTagText("FBASSOCIATIONCOTYPE");
	frmScreen.txtSearchType.value = (sComboValidation == "") ? sComboValidation : 
								xmlCombos.GetComboDescriptionForValidation("CAPSCompanyType", sComboValidation);

	sComboValidation = XML.GetTagText("FBASSOCIAITIONINFOSOURCE");
	frmScreen.txtApplicationType.value = (sComboValidation == "") ? sComboValidation : 
								xmlCombos.GetComboDescriptionForValidation("InfoSource", sComboValidation);
			
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

function DisableNavigationButtons()
{
	if(m_iaFBAliasSequence.length == 1 || m_iaFBAliasSequence.length == 0)
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
			if(parseInt(iCurrentBlock) + 1 == m_iaFBAliasSequence.length)
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




