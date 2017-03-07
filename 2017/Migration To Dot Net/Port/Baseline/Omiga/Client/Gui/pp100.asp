<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<%
/*
Workfile:      pp100.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Maintain Payee Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		06/02/01	New Screen
CL		05/03/01	SYS1920 Read only functionality added
MC		21/03/01	SYS2131 - ThirdParty Search
CL		22/03/01	SYS2134 - Town and District changed around. so that District comed before Town.
CL		23/03/01	SYS2134 - Revamped incorrect array to feed District and Town correctly.
SA		17/07/01	SYS2224 - Set focus to payee name on entry. Prevent Payee name from disappearing!
SA		18/7/01		SYS2178 - #Include PP100attribs.asp and call setmasks()
SA		19/07/01	SYS2159 - Display bank account details correctly.
SA		19/07/01	SYS2182	- Enable Payee Search button in Edit mode.
SR		06/09/01	SYS2412  
BG		22/11/01	SYS3061  - enabled payee type combo after a search. Onchange of the combo after a search
								I have enabled the search button.  Also made the screen writable after doing
								a search.
LD		23/05/02	SYS4727	  Use cached versions of frame functions
DB		29/05/02	SYS4767 - MSMS to core Integration
STB		31/05/02	SYS4536 - If payment details exist for a payee, set PP100 to read-only.
SG		05/06/02	SYS4818 - Fix error in SYS4767
GD		26/06/02	BMIDS0077 applied SG 25/06/02 SYS4930
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Prog	Date		AQR			Description 
MV		06/08/2002	BMIDS00294  Core Upgrade inLine with the functionality - Ref AQR's SYS4728,SYS4847,SYS4930,SYS4205
								Modified - HTML, PopulateScreen();frmScreen.btnBankWizard.onclick(); PopulateThirdPartyData()
TW      09/10/2002	SYS5115		Modified to incorporate client validation 
MO		01/11/2002	BMIDS00725  Modified to save whether the bank wizard search was successful - 
MV		11/11/2002  BMIDS00819	Bank Details are not mandatory
PSC		16/11/2002	BMIDS00461	Amended to clear down details on pressing 'Another' but keep them
								on change of payee type
MV		06/12/2002	BMIDS0020	Modified PAF Search Tab Order
MV		17/03/2003	BM0410		AMended SavePayeeHistory() 
KRW     26/10/2004  BMIDS931    Disable OnOK button on selection (due to duplicate records being created)
TLiu	02/09/2005	MAR38		Changed layout for Flat No., House Name & No.
SAB		24/03/2006	EP288		Removed hard-coded population of the country code

HMA     21/09/2006  EP2_3       Add roll number. Correct display of Bank Sort Code
PSC		23/03/2007	EP2_2087	Change length of roll number to 20 characters
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<HEAD>
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
</HEAD>

<BODY>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets */ %>
<% //DB SYS4767 - MSMS Integration %>
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<% //DB End %>

<% /* FORMS */ %>
<form id="frmToPP070" method="post" action="PP070.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% // Specify Screen Layout Here %>
<form id="frmScreen" mark validate="onchange" year4>
	<div id="divBackground" style="HEIGHT: 449px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">		
		<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
			Payee Type
			<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
				<select id="cboPayeeType" style="WIDTH: 300px" class="msgCombo"></select>
			</span>	
			
			<span style="LEFT: 450px; POSITION: absolute; TOP: -5px">
				<input id="btnSearch" value="Payee Search" type="button" style="WIDTH: 90px" class="msgButton">
			</span>	
		</span>

		<span style="LEFT: 10px; POSITION: absolute; TOP: 34px" class="msgLabel">
			Payee Name
			<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
				<input id="txtPayeeName" maxlength="45" style="POSITION: absolute; WIDTH: 300px" class="msgTxt">
			</span>	
		</span>
		
		<span style="LEFT: 10px; POSITION: absolute; TOP: 58px" class="msgLabel">
			<strong>Address</strong>
			<% /* MV - 05/08/2002 - BMIDS0294 
				<font face="MS Sans Serif" size="1"> 
				<strong>Address</strong>
			</font> */ %>
		</span>				

		<span id="spnAddress" style="LEFT: 10px; POSITION: absolute; TOP: 82px">
			<span style="LEFT: 0px; POSITION: absolute; TOP: 0px" class="msgLabel">
				Postcode
				<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
					<input id="txtPostcode" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgTxtUpper">
				</span>
				<span style="LEFT: 220px; POSITION: absolute; TOP: -3px">
					<input id="btnClear" value="Clear" type="button" style="WIDTH: 90px" class="msgButton">
				</span>
			</span>
		
			<span style="LEFT: 0px; POSITION: absolute; TOP: 24px" class="msgLabel">
				Flat No./ Name
				<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
					<input id="txtFlatNo"  maxlength="10" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
				</span>
			
				
			</span>
		
			<span style="LEFT: 0px; POSITION: absolute; TOP: 48px" class="msgLabel">
				Building No. &amp; Name
				<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
					<input id="txtHouseNumber" name="HouseNumber" maxlength="10" style="POSITION: absolute; WIDTH: 45px" class="msgTxt">
				</span>
				
				<span style="LEFT: 170px; POSITION: absolute; TOP: -3px">
					<input id="txtHouseName" name="HouseName" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgTxt">
				</span>
			
			</span>
			
			<span style="LEFT: 220px; POSITION: absolute; TOP: 20px">
					<input id="btnPAF" value="PAF Search" type="button" style="WIDTH: 90px" class="msgButton">
			</span>
						
			<span style="LEFT: 0px; POSITION: absolute; TOP: 72px" class="msgLabel">
				Street
				<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
					<input id="txtStreet" name="Street" maxlength="40" style="POSITION: absolute; WIDTH: 250px" class="msgTxt">
				</span>
			</span>

			<span style="LEFT: 0px; POSITION: absolute; TOP: 96px" class="msgLabel">
				District
				<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
					<input id="txtDistrict" name="District" maxlength="40" style="POSITION: absolute; WIDTH: 250px" class="msgTxt">
				</span>
			</span>
						
			<span style="LEFT: 0px; POSITION: absolute; TOP: 120px" class="msgLabel">
				Town
				<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
					<input id="txtTown" name="Town" maxlength="40" style="POSITION: absolute; WIDTH: 250px" class="msgTxt">
				</span>
			</span>
				
			<span style="LEFT: 0px; POSITION: absolute; TOP: 144px" class="msgLabel">
				County
				<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
					<input id="txtCounty" name="County" maxlength="40" style="POSITION: absolute; WIDTH: 250px" class="msgTxt">
				</span>
			</span> 
		
			<span style="LEFT: 0px; POSITION: absolute; TOP: 168px" class="msgLabel">
				Country
				<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
					<select id="cboCountry" style="WIDTH: 250px" class="msgCombo"></select>
				</span>
			</span> 
		</span>
		
		<span style="LEFT: 10px; POSITION: absolute; TOP: 274px" class="msgLabel">
			<strong>Bank account details</strong>
			
			<% /* MV - BMIDS0294 - 05/08/2002 
			<font face="MS Sans Serif" size="1">
				<strong>Bank account details</strong>
			</font> */ %>
		</span>				
				
		<span style="LEFT: 10px; POSITION: absolute; TOP: 298px" class="msgLabel">
			Bank Sort Code
			<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
				<input id="txtBankSortCode" name="BankSortCode" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgTxt" onkeyup="BankWizardDetailsChanged()">
			</span>

			<span style="LEFT: 450px; POSITION: absolute; TOP: -3px">
				<input id="btnBankWizard" value="Bank Wizard" type="button" style="WIDTH: 90px" class="msgButton">
			</span>			
		</span> 
		
		<span style="LEFT: 10px; POSITION: absolute; TOP: 322px" class="msgLabel">
			Account Number
			<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
				<input id="txtAccountNumber" name="Accounnumber" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgTxt" onkeyup="BankWizardDetailsChanged()">
			</span>
		</span> 		
		
		<span style="LEFT: 222px; POSITION: absolute; TOP: 322px" class="msgLabel">
			Roll Number
			<span style="LEFT: 63px; POSITION: absolute; TOP: -3px">
				<input id="txtRollNumber" name="RollNumber" maxlength="20" style="POSITION: absolute; WIDTH: 130px" class="msgTxt" onkeyup="BankWizardDetailsChanged()">
			</span>
		</span> 		

		<span style="LEFT: 10px; POSITION: absolute; TOP: 346px" class="msgLabel">
			Bank Name
			<span style="LEFT: 125px; POSITION: absolute; TOP: -3px">
				<input id="txtBankName" name="BankName" maxlength="50" style="POSITION: absolute; WIDTH: 280px" class="msgTxt">
			</span>
		</span> 		
		
		<span style="LEFT: 10px; POSITION: absolute; TOP: 370px" class="msgLabel">
			Notes
			<span style="LEFT: 125px; POSITION: absolute; TOP: -3px"><TEXTAREA class=msgTxt id=txtNotes name=Notes rows=5 style="POSITION: absolute; WIDTH: 450px"></TEXTAREA>  
			</span>	
		</span>		
	</div>
</form>

<% /* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 520px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<% /* DB SYS4767 - MSMS Integration */ %>
<% /*<span id="spnToFirstField" tabindex="0"></span>
	<!-- #include FILE="fw030.asp" 
	<!-- #include FILE="includes/pafsearch.asp" -->
	<!-- #include FILE="attribs/PP100attribs.asp" --> */ %>
<span id="spnToFirstField" tabindex="0"></span>
<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="includes/pafsearch.asp" -->
<!-- #include FILE="attribs/PP100attribs.asp" -->
<!-- #include FILE="includes/BankWizard.asp" -->
<% /* DB End */ %>

<%/* CODE */ %>
<script language="JScript">
<!--
var m_sApplicationNumber, m_sAFFNumber, m_sMetaAction, m_sUserId ;
var m_sThirdPartyGuid = '';
var m_sAddressGuid = '';

var m_sContext = '';
var m_sPayeeHistorySeqNo ;
var m_sReadOnly ;
var sOtherTPType ;

var XMLContext, ThirdPartyXML;
var m_bBankAccountValid ;
var m_blnReadOnly = false;
var m_sScreenId = "PP100";
var m_bIsPopup = false;

function RetrieveContextData()
{
<%	
	// TEST	
	
	//scScreenFunctions.SetContextParameter(window, "idApplicationNumber", "C00078387");
	//scScreenFunctions.SetContextParameter(window, "idApplicationFactFindNumber", "1");
	//scScreenFunctions.SetContextParameter(window, "idReadOnly", "0");
	//scScreenFunctions.SetContextParameter(window, "idUserId", "SR");
	// END TEST
%>	
	m_sApplicationNumber= scScreenFunctions.GetContextParameter(window,"idApplicationNumber", "");
	m_sAFFNumber		= scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber", "");
	m_sMetaAction		= scScreenFunctions.GetContextParameter(window,"idMetaAction", "");
	m_sReadOnly			= scScreenFunctions.GetContextParameter(window,"idReadOnly", "0");
	m_sUserId			= scScreenFunctions.GetContextParameter(window,"idUserId", "");
	
	if(m_sMetaAction == "Edit")
	{
		m_sContext = scScreenFunctions.GetContextParameter(window,"idXML", "");

		XMLContext = new top.frames[1].document.all.scXMLFunctions.XMLObject();

		XMLContext.LoadXML(m_sContext);	
		XMLContext.CreateTagList('PAYEEHISTORYDATA');
		XMLContext.SelectTagListItem(0);
		m_sPayeeHistorySeqNo = XMLContext.GetTagText("PAYEEHISTORYSEQNO");
	}
}

<% /* Events */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Maintain Payee Details","PP100",scScreenFunctions)
	
	var saButtonList = new Array("Submit", "Cancel", "Another");
	ShowMainButtons(saButtonList);
	
	SetMasks();
	Validation_Init();
	RetrieveContextData();
	if(!PopulateCombos()) 
	{
		alert('Error loading combos');
		return ;
	}
	
	PopulateScreen();
	
	<% /* SYS4536 - If payment details exist for a payee, set PP100 to read-only. */ %>
	if (scScreenFunctions.GetContextParameter(window, "idPayeeReadOnly", "0") == "1")
	{
		m_sReadOnly = "1";
	}
	else
	{
		m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "PP100");
		if (m_blnReadOnly == true) m_sReadOnly = "1";		
	}
	
	SetScreenToReadOnly();
	
	<% /* MV - 05/08/2002 - BMIDS00294 */ %>
	<% /* SR 22/08/2001 : SYS2412 - In edit mode, for type 'Other' edits are not allowed */ %>
	//var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	//if (TempXML.IsInComboValidationList(document,"PaymentStatus",m_sPayeementStatus,["C"]))
	//	DoEditProcessing();
		
	
	//SA SYS2224 Set focus on screen
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function btnSubmit.onclick()
{	
		
	if(frmScreen.onsubmit())
	{
		if(!ValidateBeforeSave()) return;

		if(SavePayeeHistory()) frmToPP070.submit();
		
		DisableMainButton("Submit"); // BMIDS931 KRW 26/10/2004
	}
}

function btnCancel.onclick()
{
	frmToPP070.submit();
}

function btnAnother.onclick()
{
	if(!ValidateBeforeSave()) return;
	
	if(SavePayeeHistory())
	{
		ClearScreen(true);
		frmScreen.txtPayeeName.focus();
	}
	else return ;
}

function frmScreen.btnClear.onclick()
{
	scScreenFunctions.ClearCollection(spnAddress);
	frmScreen.txtPostcode.focus();
}

function frmScreen.btnPAF.onclick()
{
	with (frmScreen)
		m_bAddressPAFIndicator = PAFSearch(txtPostcode, txtHouseName, txtHouseNumber, txtFlatNo,
				txtStreet,txtDistrict, txtTown, txtCounty, cboCountry);

}

function frmScreen.btnBankWizard.onclick()
{
	<% /* INR 21/06/02 SYS2412 Call to INCLUDE/BANKWIZARD.ASP */ %>
	<% /* EP2_3 Add roll number */ %>
	with (frmScreen)
		m_bBankAccountValid = BankWizard(txtBankSortCode,txtAccountNumber,txtRollNumber);


	<%/* var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(window,null)

	XML.CreateActiveTag("BANKDETAILS");
	// Create tags for the key
	XML.CreateTag("SORTCODE", frmScreen.txtBankSortCode.value);
	XML.CreateTag("ACCOUNTNUMBER", frmScreen.txtAccountNumber.value);
	XML.RunASP(document,"GetBankDetails.asp");
	
	ErrorArray = XML.CheckResponse();
	if (ErrorArray[0] == true)
	{
		if (ErrorArray[3] == false)
		{
			m_bBankAccountValid = ErrorArray[0];
			alert("The Bank Sort Code and Account Number have been successfully validated.");
		}
		else m_bBankAccountValid = false;
	}	
	else m_bBankAccountValid = ErrorArray[0];		
	
	XML = null; */ %>
}

function frmScreen.btnSearch.onclick()
{		
	if(frmScreen.cboPayeeType.value == "") 
	{
		alert("Choose a particular Payee Type for searching") ;
		frmScreen.cboPayeeType.focus();
		return;
	}
	
	FindThirdPartyList();
}

function frmScreen.cboPayeeType.onchange()
{	
	<% /* PSC 16/11/2002 BMIDS00461 */%>
	<% /* ClearScreen(false);  */%>
	<% /*  BG 22/11/01 SYS3061 */%>
	frmScreen.btnSearch.disabled = false;
}

<% /* Functions */  %>
function PopulateCombos()
{
	// Populate combos from table 'ComboValue'
	var XMLCombos = null;

	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var XML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>

	var blnSuccess = true;
	var sGroupList = new Array("ThirdPartyType", "Country");
	if(XML.GetComboLists(document,sGroupList))
	{
		XMLCombos = XML.GetComboListXML("Country");
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document, 
								frmScreen.cboCountry,XMLCombos,true);
	
		XMLCombos = XML.GetComboListXML("ThirdPartyType");
		sOtherTPType = XML.GetComboIdForValidation("ThirdPartyType", "O", XMLCombos) ; ;
		
		RemoveNonDisplayValues(XMLCombos);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,
								frmScreen.cboPayeeType, XMLCombos,true);
	}
	
	return blnSuccess ;
}

function RemoveNonDisplayValues(XMLCombos)
{
	var xmlNodeList, xmlNode ;
	var xmlValidationList, xmlValidation, iValidationCount ;
	var sPattern ;
	
	xmlNodeList = XMLCombos.firstChild.childNodes ;
	
	for(var iCount = xmlNodeList.length -1 ; iCount>=0; --iCount)
	{
		xmlNode = xmlNodeList(iCount) ;
		
		xmlValidationList = xmlNode.selectNodes(".//VALIDATIONTYPE")
		bFound = false ;
		for(iValidationCount = xmlValidationList.length - 1; iValidationCount >= 0; --iValidationCount)
		{
			xmlValidation = xmlValidationList(iValidationCount);
			if(xmlValidation.text == 'XP')
			{
				bFound = true ;
				break ;
			} 
		}
		
		if(!bFound) XMLCombos.firstChild.removeChild(xmlNode) ;	
	}
}

function DoEditProcessing()
{
	if(m_sMetaAction == 'Edit')
	{
		m_sReadOnly = 1;
		SetScreenToReadOnly() ;
		SetAllFieldsToReadOnly();
		
		<% /* MV - 05/08/2002 - BMIDS00294  
		if(frmScreen.cboPayeeType.value != sOtherTPType)
		{
			m_sReadOnly = 1;
			SetScreenToReadOnly() ;
			
			SetAllFieldsToReadOnly();
		}	 */ %>
	}
}

function PopulateScreen()
{
	if(m_sMetaAction == "Edit")
	{	
		<%//SYS2182 Shouldn't be disabled in Edit mode because there is no Delete option.
		//frmScreen.btnSearch.disabled = true ;  // disable Search button
		%>
		DisableMainButton('Another') ; // Another button is required only for 'Create'
		
		XMLContext.ActiveTag = null ;
		XMLContext.CreateTagList("PAYEEHISTORY");
		
		if(XMLContext.ActiveTagList.length > 0) XMLContext.SelectTagListItem(0);
		else
		{
			XMLContext.ResetXMLDocument();
			XMLContext.CreateRequestTag(window,"FINDPAYEEHISTORYLIST");
			XMLContext.CreateActiveTag("PAYEEHISTORY");
			XMLContext.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
			XMLContext.SetAttribute("PAYEEHISTORYSEQNO", m_sPayeeHistorySeqNo);
			XMLContext.SetAttribute("_COMBOLOOKUP_", "1");
			XMLContext.RunASP(document, "PaymentProcessingRequest.asp");
			
			var ErrorTypes = new Array("RECORDNOTFOUND");
			var ErrorReturn = XMLContext.CheckResponse(ErrorTypes);
	
			if(ErrorReturn[0] == false)
			{
				alert('Error retrieving payee details');
				SetScreenToReadOnly();
				return ;
			}
			else
			{
				XMLContext.CreateTagList("PAYEEHISTORY");
				XMLContext.SelectTagListItem(0);
			}
		}
		
		m_sThirdPartyGuid				= XMLContext.GetAttribute("THIRDPARTYGUID");
		frmScreen.txtPayeeName.value	= XMLContext.GetTagAttribute("THIRDPARTY", "COMPANYNAME");
		frmScreen.cboPayeeType.value	= XMLContext.GetTagAttribute("THIRDPARTY", "THIRDPARTYTYPE");
		m_sAddressGuid					= XMLContext.GetTagAttribute("ADDRESS", "ADDRESSGUID");
		frmScreen.txtPostcode.value		= XMLContext.GetTagAttribute("ADDRESS", "POSTCODE");
		frmScreen.txtFlatNo.value		= XMLContext.GetTagAttribute("ADDRESS", "FLATNUMBER");
		frmScreen.txtHouseName.value	= XMLContext.GetTagAttribute("ADDRESS", "BUILDINGORHOUSENAME");
		frmScreen.txtHouseNumber.value	= XMLContext.GetTagAttribute("ADDRESS", "BUILDINGORHOUSENUMBER");
		frmScreen.txtStreet.value		= XMLContext.GetTagAttribute("ADDRESS", "STREET");
		frmScreen.txtDistrict.value		= XMLContext.GetTagAttribute("ADDRESS", "DISTRICT");
		frmScreen.txtTown.value			= XMLContext.GetTagAttribute("ADDRESS", "TOWN");
		frmScreen.txtCounty.value		= XMLContext.GetTagAttribute("ADDRESS", "COUNTY");
		frmScreen.cboCountry.value		= XMLContext.GetTagAttribute("ADDRESS", "COUNTRY");
		
		frmScreen.txtBankSortCode.value = XMLContext.GetAttribute("BANKSORTCODE");
		frmScreen.txtAccountNumber.value= XMLContext.GetAttribute("ACCOUNTNUMBER");
		frmScreen.txtRollNumber.value   = XMLContext.GetAttribute("ROLLNUMBER");    <%/* EP2_3 */%>
		frmScreen.txtBankName.value		= XMLContext.GetAttribute("BANKNAME");
		frmScreen.txtNotes.value		= XMLContext.GetAttribute("NOTES");
		
		<% /*MO - 01/11/2002 - BMIDS00725*/ %>
		if (XMLContext.GetAttribute("BANKWIZARDINDICATOR") == "1") {
			m_bBankAccountValid = true ;
		} else {
			m_bBankAccountValid = false ;
		}
	}
	else
	{
		<% /* SAB - EP288 */ %>
		//frmScreen.cboCountry.value = "1" ;
		m_bBankAccountValid = false ;
	}
	
	
}

function SetScreenToReadOnly()
{
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);

		frmScreen.btnSearch.disabled = true;
		frmScreen.btnClear.disabled = true;
		frmScreen.btnPAF.disabled = true;
		frmScreen.btnBankWizard.disabled = true;
		
		DisableMainButton("Submit");
		DisableMainButton("Another");
		
	}
}

function SetAllFieldsToReadOnly()
{	
	<% /*  BG 22/11/01 SYS3061 */%>
	frmScreen.cboPayeeType.disabled = false;
	<%/*frmScreen.cboPayeeType.disabled = true ;*/%>
	frmScreen.txtPayeeName.disabled = true ;
	frmScreen.txtPostcode.disabled = true ;
	frmScreen.txtFlatNo.disabled = true ;
	frmScreen.txtHouseName.disabled = true ;
	frmScreen.txtHouseNumber.disabled = true ;
	frmScreen.txtStreet.disabled = true ;
	frmScreen.txtDistrict.disabled = true ;
	frmScreen.txtTown.disabled = true ;
	frmScreen.txtCounty.disabled = true ;
	frmScreen.txtDistrict.disabled = true ;
	frmScreen.cboCountry.disabled = true ;
	
	frmScreen.txtBankSortCode.disabled = true ;
	frmScreen.txtBankName.disabled = true ;
	frmScreen.txtAccountNumber.disabled = true ;
	frmScreen.txtRollNumber.disabled = true ;      <%/* EP2_3 */%>

	frmScreen.txtNotes.disabled = true ;

}

function ValidateBeforeSave()
{
	<% // Postcode cannot be empty
	%>
	if(frmScreen.txtPostcode.value == '')
	{
		alert('Postcode cannot be left empty');
		frmScreen.txtPostcode.focus();
		return false ;
	}
	
	<% // validate the postcode
	%>
	if(!scScreenFunctions.ValidatePostcode(frmScreen.txtPostcode)) return false
		
	<% /* MV - 11/11/2002 - BMIDS00819 
	Bank sort code & account number cannot be empty
	if(frmScreen.txtBankSortCode.value == '')
	{
		alert('Bank Sort Code cannot be left empty');
		frmScreen.txtBankSortCode.focus();
		return false ;
	}		
	
	if(frmScreen.txtAccountNumber.value == '')
	{
		alert('Account Number cannot be left empty');
		frmScreen.txtAccountNumber.focus();
		return false ;
	}*/%>
	
	if (!scScreenFunctions.ValidatePostcode(frmScreen.txtPostcode)) return false ;
		
	<% // check whether bank sort code and account number have been validates thru bank wizard
	%>
	if(!m_bBankAccountValid)	
		if (!confirm("The bank account details have not been validated using the Bank Wizard. Continue anyway?"))
				return false ; 
		
	return true ;
}

function SavePayeeHistory()
{	
	var XMLPayeeHistory = null;

	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	var XMLPayeeHistory = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* var XMLPayeeHistory = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>

	if(m_sMetaAction == 'Edit') XMLPayeeHistory.CreateRequestTag(window,"UPDATEPAYEEHISTORYDETAILS");
	else XMLPayeeHistory.CreateRequestTag(window,"CREATEPAYEEHISTORYDETAILS");
	
	XMLPayeeHistory.CreateActiveTag("PAYEEHISTORY");
	XMLPayeeHistory.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	if(m_sMetaAction == 'Edit') XMLPayeeHistory.SetAttribute("PAYEEHISTORYSEQNO", m_sPayeeHistorySeqNo);	
	XMLPayeeHistory.SetAttribute("BANKSORTCODE", frmScreen.txtBankSortCode.value);
	XMLPayeeHistory.SetAttribute("ACCOUNTNUMBER", frmScreen.txtAccountNumber.value);
	XMLPayeeHistory.SetAttribute("BANKNAME", frmScreen.txtBankName.value);
	XMLPayeeHistory.SetAttribute("NOTES", frmScreen.txtNotes.value);
	XMLPayeeHistory.SetAttribute("USERID", m_sUserId);
	<% /*MO - 01/11/2002 - BMIDS00725*/ %>
	if (m_bBankAccountValid == true) {
		XMLPayeeHistory.SetAttribute("BANKWIZARDINDICATOR", "1");
	} else {
		XMLPayeeHistory.SetAttribute("BANKWIZARDINDICATOR", "0");
	}
	<%/* EP2_3 */%>
	XMLPayeeHistory.SetAttribute("ROLLNUMBER", frmScreen.txtRollNumber.value);
	
	XMLPayeeHistory.CreateActiveTag("THIRDPARTY") ;
	if(m_sMetaAction == 'Edit') XMLPayeeHistory.SetAttribute("THIRDPARTYGUID", m_sThirdPartyGuid) ;
	XMLPayeeHistory.SetAttribute("THIRDPARTYTYPE", frmScreen.cboPayeeType.value) ;
	XMLPayeeHistory.SetAttribute("COMPANYNAME", frmScreen.txtPayeeName.value) ;
<%
// TEST - AccessAuditGuid is a not null column. Remove it after tesing
// XMLPayeeHistory.SetAttribute("ACCESSAUDITGUID", '1234') ;
// END TEST
%>
	XMLPayeeHistory.CreateActiveTag("ADDRESS") ;
	if(m_sMetaAction == 'Edit') XMLPayeeHistory.SetAttribute("ADDRESSGUID", m_sAddressGuid) ;
	XMLPayeeHistory.SetAttribute("BUILDINGORHOUSENAME", frmScreen.txtHouseName.value);
	XMLPayeeHistory.SetAttribute("BUILDINGORHOUSENUMBER", frmScreen.txtHouseNumber.value);
	XMLPayeeHistory.SetAttribute("FLATNUMBER", frmScreen.txtFlatNo.value);
	XMLPayeeHistory.SetAttribute("STREET", frmScreen.txtStreet.value);
	XMLPayeeHistory.SetAttribute("DISTRICT", frmScreen.txtDistrict.value);
	XMLPayeeHistory.SetAttribute("TOWN", frmScreen.txtTown.value);
	XMLPayeeHistory.SetAttribute("COUNTY", frmScreen.txtCounty.value);
	XMLPayeeHistory.SetAttribute("COUNTRY", frmScreen.cboCountry.value);
	XMLPayeeHistory.SetAttribute("POSTCODE", frmScreen.txtPostcode.value);
	
	// 	XMLPayeeHistory.RunASP(document, "PaymentProcessingRequest.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XMLPayeeHistory.RunASP(document, "PaymentProcessingRequest.asp");
			break;
		default: // Error
			XMLPayeeHistory.SetErrorResponse();
		}

	
	if(!XMLPayeeHistory.IsResponseOK()) return false
	
	return true ;
}

function ClearScreen()
{<%	// Clears all the values entered by the user. Set default value to Country
	
	//SA SYS2224 Don't clear the name as this is the first field on the screen
	// so user will have input value and the will select payee type.
 %>
	<% /* PSC 16/11/2002 BMIDS00461 - Start */ %>
	frmScreen.txtPayeeName.value	= '';
	frmScreen.cboPayeeType.value	= '';
	<% /* PSC 16/11/2002 BMIDS00461 - End */ %>
		
	frmScreen.txtPostcode.value		= '';
	frmScreen.txtFlatNo.value		= '';
	frmScreen.txtHouseName.value	= '';
	frmScreen.txtHouseNumber.value	= '';
	frmScreen.txtStreet.value		= '';
	frmScreen.txtDistrict.value		= '';
	frmScreen.txtTown.value			= '';
	frmScreen.txtCounty.value		= '';
	<% /* SAB EP288 - Set to <SELECT> */ %>
	frmScreen.cboCountry.selectedIndex = 0;
		
	frmScreen.txtBankSortCode.value = '';
	frmScreen.txtAccountNumber.value= '';
	frmScreen.txtBankName.value		= '';
	frmScreen.txtNotes.value		= '';
	frmScreen.txtRollNumber.value   = '';  <%/* EP2_3 */%>
}

function BankWizardDetailsChanged()
{
	frmScreen.btnBankWizard.disabled = ((frmScreen.txtAccountNumber.value == "") ||
			(frmScreen.txtBankSortCode.value == ""));
	m_bBankAccountValid = false;
}

function FindThirdPartyList()
{
	function GetThirdPartyDataFromArray(saTPData)
	{
		frmScreen.txtPayeeName.value	= saTPData[0] ;
		//frmScreen.cboPayeeType.value	= saTPData[1] ;
		frmScreen.txtPostcode.value		= saTPData[2] ;
		frmScreen.txtFlatNo.value		= saTPData[3] ;
		frmScreen.txtHouseName.value	= saTPData[4] ;
		frmScreen.txtHouseNumber.value	= saTPData[5] ;
		frmScreen.txtStreet.value		= saTPData[6] ;
		frmScreen.txtDistrict.value		= saTPData[8] ; <%/*was saTPData[7]*/%>
		frmScreen.txtTown.value			= saTPData[7] ; <%/*was saTPData[8]*/%>
		frmScreen.txtCounty.value		= saTPData[9] ;
		frmScreen.cboCountry.value		= saTPData[10] ;
		frmScreen.txtBankName.value		= saTPData[11] ;
		frmScreen.txtBankSortCode.value	= saTPData[12] ;
		//SYS2159 Display Bank Account Number too
		frmScreen.txtAccountNumber.value= saTPData[13];
		frmScreen.txtRollNumber.value   = saTPData[14]; <%/* EP2_3 */%>
	}
	
	function PopulateThirdPartyData()
	{
		frmScreen.txtPayeeName.value	= ThirdPartyXML.GetTagText("COMPANYNAME") ;
		
		<% /*  BG 22/11/01 SYS3061 */%>
		<% /*  DB SYS4767 - MSMS Integration
		  frmScreen.cboPayeeType.value	= ThirdPartyXML.GetTagText("NAMEANDADDRESSTYPE") ;
		  DB End
		  frmScreen.cboPayeeType.value	= ThirdPartyXML.GetTagText("THIRDPARTYTYPE") ; */%>
		frmScreen.txtPostcode.value		= ThirdPartyXML.GetTagText("POSTCODE") ;
		frmScreen.txtFlatNo.value		= ThirdPartyXML.GetTagText("FLATNUMBER") ;
		frmScreen.txtHouseName.value	= ThirdPartyXML.GetTagText("BUILDINGORHOUSENAME") ;
		frmScreen.txtHouseNumber.value	= ThirdPartyXML.GetTagText("BUILDINGORHOUSENUMBER") ;
		frmScreen.txtStreet.value		= ThirdPartyXML.GetTagText("STREET") ;
		frmScreen.txtDistrict.value		= ThirdPartyXML.GetTagText("DISTRICT") ;
		frmScreen.txtTown.value			= ThirdPartyXML.GetTagText("TOWN") ;
		frmScreen.txtCounty.value		= ThirdPartyXML.GetTagText("COUNTY") ;
		frmScreen.cboCountry.value		= ThirdPartyXML.GetTagText("COUNTRY") ;
		
		<% /* MV - 05/08/2002 - BMIDS0294 */ %>
		//SYS4205
		frmScreen.txtBankName.value     = ThirdPartyXML.GetTagText("BRANCHNAME") ;
		
		<%/* EP2_3 Display the Sort Code correctly */%>
		if (ThirdPartyXML.GetTagText("NAMEANDADDRESSTYPE") != "")
		{
			frmScreen.txtBankSortCode.value = ThirdPartyXML.GetTagText("NAMEANDADDRESSBANKSORTCODE") ;
		}
		else
		{
			frmScreen.txtBankSortCode.value = ThirdPartyXML.GetTagText("THIRDPARTYBANKSORTCODE") ;
		}
		
		frmScreen.txtAccountNumber.value = ThirdPartyXML.GetTagText("ACCOUNTNUMBER") ;
		frmScreen.txtRollNumber.value = ThirdPartyXML.GetTagText("ROLLNUMBER") ;   <%/* EP2_3 */%>

	}

<%	// Find ThirdPartyrecords that correspond to the ThirdPartyType selected. If no matching
	// records exist, prompt the user and return. 		
%>
	<%/* SG 05/06/02 SYS4818 */%>	
	<%/* DB SYS4767 - MSMS Integration */%>
	ThirdPartyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/* ThirdPartyXML = new scXMLFunctions.XMLObject(); */%>
	<%/* DB End */%>	
	<%/* SG 05/06/02 SYS4818 */%>
	ThirdPartyXML.CreateRequestTag(window,null);

	var tagApplication = ThirdPartyXML.CreateActiveTag("APPLICATION");
	ThirdPartyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	ThirdPartyXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAFFNumber);
	ThirdPartyXML.RunASP(document,"FindApplicationThirdPartyList.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = ThirdPartyXML.CheckResponse(ErrorTypes);

	if(ErrorReturn[0] == false)
	{
		if(ErrorReturn[1] == ErrorTypes[0]) 
		{
			alert("No Third Parties exist matching the selection criteria");
			frmScreen.cboPayeeType.focus();
			return false;
		}	
		else
		{	
			alert("Error in searching Third Party data");
			frmScreen.cboPayeeType.focus();
			return false;
		}
	}
<%	
	// if one record is found, populate the ThirdParty data in the screen. else, go to PP110.
	// Attempt to find thirdparty of the required type
%>
	var sCondition = "APPLICATIONTHIRDPARTY[THIRDPARTY/THIRDPARTYTYPE='" + frmScreen.cboPayeeType.value + "' || NAMEANDADDRESSDIRECTORY/NAMEANDADDRESSTYPE='" + frmScreen.cboPayeeType.value + "']";
	ThirdPartyXML.CreateTagList(sCondition);
	if(ThirdPartyXML.ActiveTagList.length ==0)
	{
		alert("No Third Parties exist matching the selection criteria");
		frmScreen.cboPayeeType.focus();
		return false;
	}	
	else if(ThirdPartyXML.ActiveTagList.length ==1)
	{
		ThirdPartyXML.SelectTagListItem(0);
		PopulateThirdPartyData();
		
		<% /*  BG 22/11/01 SYS3061 */%>
		//SetAllFieldsToReadOnly();

		frmScreen.btnSearch.disabled = true;
		frmScreen.btnClear.disabled = true;
		frmScreen.btnPAF.disabled = true;
		frmScreen.btnBankWizard.disabled = true;
	}
	else
	{	// call PP110.asp - Payee Search Results
		var ArrayArguments = new Array();
		ArrayArguments[0] = frmScreen.cboPayeeType.value ;
		ArrayArguments[1] = ThirdPartyXML.XMLDocument.xml ;
		ArrayArguments[2] = m_sApplicationNumber ;
		ArrayArguments[3] = m_sAFFNumber ;
		
		var saReturnArray = scScreenFunctions.DisplayPopup(window, document, "PP110.asp", ArrayArguments, 630, 340);
		if(saReturnArray == null) frmScreen.cboPayeeType.focus()
		else GetThirdPartyDataFromArray(saReturnArray);
	}
	
	return true;
}

-->
</script>
</BODY>
</HTML>

