<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      dc101.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		14/02/00	Change to msgButtons button types
IW		13/03/00    SYS0262 (Refer to Policy rather than Product etc)
AY		30/03/00	New top menu/scScreenFunctions change
IW		02/05/00	SYS0601 - (Lose PolicyRelationship table)
MH      02/05/00    SYS0618 - Postcode validation
MC		05/05/00	SYS0677 - Validation of Maturity Year
IW		05/05/00	SYS0676 - Cosmetics
BG		17/05/00	SYS0752 Removed Tooltips
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
MC		09/06/00	SYS0866 Standardise format of balance and payment fields
MH      23/06/00    SYS0933 Readonly stuff
MC		06/07/00	SYS0846 Various errors. See Deuce for more detail.
BG		29/08/00	SYS0862 Changed heading from Lender Details to Life Office Details,
					removed Organisation type combo from screen.
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
BG		04/11/00	SYS1037 When populating country combo, set final parameter to false so there is
							no "SELECT" option and remove validation for "0" in ValidateThirdPartyDetails.
CL		05/03/01	SYS1920 Read only functionality added
SA		06/06/01	SYS1672 Guarantors added to applicants combo
JR		14/09/01	Omiplus24 - Removed FaxNumber, TelephoneNumber & EmailAddress. Now included in TPD.asp
MDC		01/10/01	SYS2785 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
TW      09/10/2002  SYS5115 Modified to incorporate client validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
PB		21/07/2006	EP543		Added 'TITLEOTHER' field for third parties
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="Validation.js" language=""></script>

<form id="frmToDC100" method="post" action="dc100.asp" STYLE="DISPLAY: none"></form>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 358px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Life Policy Owner
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<select id="cboApplicant" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	</span>
	<span style="TOP: 36px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Policy Number
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtAccountNumber" type="text" maxlength="20" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 60px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Policy Type
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<select id="cboPolicyType" style="WIDTH: 200px" class="msgCombo">
			</select>
		</span>
	</span>
	<span style="TOP: 84px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Monthly Premium
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtMonthlyPremium" type="text" maxlength="9" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 108px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Year of Maturity
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtYearOfMaturity" type="text" maxlength="4" msg="Year of Maturity must be before 2100 and cannot be in the past." style="WIDTH: 40px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 132px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Minimum Death Benefit
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtMinimumDeathBenefit" type="text" maxlength="6" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 156px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Maturity Value
		<span style="TOP: -3px; LEFT: 200px; POSITION: ABSOLUTE">
			<input id="txtMaturityValue" type="text" maxlength="6" style="WIDTH: 65px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 174px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Is the Policy to be assigned<br>to the New Loan?
		<span style="TOP: 3px; LEFT: 196px; POSITION: ABSOLUTE">
			<input id="optASSIGNEDTOLOANINDICATORYes" name="AssignToNewLoanInd" type="radio" value="1">
			<label for="optASSIGNEDTOLOANINDICATORYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: 3px; LEFT: 250px; POSITION: ABSOLUTE">
			<input id="optAssignToNewLoanNo" name="AssignToNewLoanInd" type="radio" value="0" checked>
			<label for="optAssignToNewLoanNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span style="TOP: 204px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Is this a New Policy?
		<span style="TOP: -3px; LEFT: 196px; POSITION: ABSOLUTE">
			<input id="optNewPolicyYes" name="NewPolicyInd" type="radio" value="1">
			<label for="optNewPolicyYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: -3px; LEFT: 250px; POSITION: ABSOLUTE">
			<input id="optNewPolicyNo" name="NewPolicyInd" type="radio" value="0" checked>
			<label for="optNewPolicyNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span style="TOP: 228px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Are you the Life Assured?
		<span style="TOP: -3px; LEFT: 196px; POSITION: ABSOLUTE">
			<input id="optLIFEASSUREDINDICATORYes" name="grpLIFEASSUREDINDICATOR" type="radio" value="1">
			<label for="optLIFEASSUREDINDICATORYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: -3px; LEFT: 250px; POSITION: ABSOLUTE">
			<input id="optLIFEASSUREDINDICATORNo" name="grpLIFEASSUREDINDICATOR" type="radio" value="0" checked>
			<label for="optLIFEASSUREDINDICATORNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span style="TOP: 252px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Are there any Additional Policy Holders?
		<span style="TOP: -3px; LEFT: 196px; POSITION: ABSOLUTE">
			<input id="optAdditionalIndYes" name="AdditionalInd" type="radio" value="1">
			<label for="optAdditionalIndYes" class="msgLabel">Yes</label>
		</span> 
		<span style="TOP: -3px; LEFT: 250px; POSITION: ABSOLUTE">
			<input id="optAdditionalIndNo" name="AdditionalInd" type="radio" value="0" checked>
			<label for="optAdditionalIndNo" class="msgLabel">No</label>
		</span> 
	</span>		
	<span id=lblAdditionalDetails style="TOP: 276px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
		Additional Policy Holder(s) Details
		<span style="TOP: 0px; LEFT: 200px; POSITION: ABSOLUTE">
			<textarea id="txtAdditionalDetails" rows="5" style="WIDTH: 300px; POSITION: absolute" class="msgTxt"></textarea>
		</span>
	</span>
</div>

<div style="HEIGHT: 304px; LEFT: 10px; POSITION: absolute; TOP: 424px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Life Office Details</strong> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Life Office
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyName" type="text" maxlength="45" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" onchange="ContactDetailsChanged()">
		</span> 
	</span>

	<span style="LEFT: 330px; POSITION: absolute; TOP: 31px" class="msgLabel">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton" onclick="btnDirectorySearch.onClick()">
	</span>

	<span style="TOP: 60px; LEFT: 4px; POSITION: absolute" class="msgLabel">
		Is in the directory?
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryYes" name="InDirectoryGroup" type="radio" value="1" onclick="InDirectoryChanged()">
			<label for="optInDirectoryYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 174px; POSITION: absolute; TOP: -3px; WIDTH: 60">
			<input id="optInDirectoryNo" name="InDirectoryGroup" type="radio" value="0" onclick="InDirectoryChanged()">
			<label for="optInDirectoryNo" class="msgLabel">No</label> 
		</span> 
	</span>
	
</div> 

<div style="HEIGHT: 232px; LEFT: 10px; POSITION: absolute; TOP: 500px; WIDTH: 604px" class="msgGroup">
<!-- #include FILE="includes/thirdpartydetails.htm" -->
</div> 
</form>

<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 738px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/dc101attribs.asp" -->
<!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = null;
var m_sReadOnly = "";
var XMLOnEntry = null;
var m_sSequenceNumber = null;
var m_sAccountGUID = null;
var scScreenFunctions;
var m_blnReadOnly = false;

function btnAnother.onclick()
{
	if (CommitChanges())
	{
		PopulateApplicantCombo();
		scScreenFunctions.SetFocusToFirstField(frmScreen);
		DefaultFields();

		m_sDirectoryGUID = "";
		m_sThirdPartyGUID = "";
		m_bDirectoryAddress = false;
		m_bPAFIndicator = false;
		SetAvailableFunctionality();
	}
}

function btnCancel.onclick()
{
	frmToDC100.submit();
}

function btnSubmit.onclick()
{
	if (CommitChanges())
		frmToDC100.submit();
}

<% /* Removes the <SELECT> option */ %>
function frmScreen.cboApplicant.onchange()
{
	if(frmScreen.cboApplicant.value != "")
		if(frmScreen.cboApplicant.item(0).text == "<SELECT>")
			frmScreen.cboApplicant.remove(0);
}

function frmScreen.optAdditionalIndYes.onclick()
{
	if(frmScreen.optAdditionalIndYes.checked)
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtAdditionalDetails");
}

function frmScreen.optAdditionalIndNo.onclick()
{
	if(frmScreen.optAdditionalIndNo.checked)
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtAdditionalDetails");
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
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	RetrieveContextData();

	<% /* Make the required buttons available on the bottom of the screen
		(see msgButtons.asp for details) */ %>
	var sButtonList = new Array("Submit","Cancel","Another");
	if(m_sMetaAction != "Add")
		sButtonList = new Array("Submit","Cancel");
	else
		FlagChange(true);
		
	ShowMainButtons(sButtonList);
	FW030SetTitles("Add/Edit Mortgage Life Policies","DC101",scScreenFunctions);

	// MDC SYS2564 / SYS2785 Client Customisation
	var sGroups = new Array("PolicyType","Country","OrganisationType");
	objDerivedOperations = new DerivedScreen(sGroups);

	<%/* Parameters for thirdpartydetails.asp */%>
	m_sThirdPartyType = "9";
	m_fValidateScreen = ValidateThirdPartyDetails;

	// MC SYS2564/SYS2785 for client customisation
	ThirdPartyCustomise();

	PopulateCombos();
	PopulateApplicantCombo();			
	if (m_sMetaAction == "Edit")
		PopulateScreen();
			
	frmScreen.cboApplicant.onchange();
	frmScreen.optAdditionalIndNo.onclick();
			
	SetThirdPartyDetailsMasks();
	SetMasks();
	Validation_Init();
	SetAvailableFunctionality();
			
	<% /* Reset label colour to 616161 as only mandatory if radio button set to yes */ %>
	lblAdditionalDetails.style.color = "#616161";

	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	
	if (m_sReadOnly=="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		ThirdPartyDetailsDisableButtons();
		DisableMainButton("Another");
	}
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC101");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function CommitChanges()
{
	var bSuccess = true;
	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				if (ValidateScreen())
					bSuccess = SaveLifePolicy();
				else
					bSuccess = false;
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function DefaultFields()
{
	scScreenFunctions.ClearCollection(frmScreen);
	frmScreen.optASSIGNEDTOLOANINDICATORYes.checked = false;
	frmScreen.optNewPolicyNo.checked = true;
	frmScreen.optAdditionalIndNo.checked = true;
	frmScreen.optAdditionalIndNo.onclick();
	ClearFields(true,true);
}

<% /* Populates the Applicant combo with all applicants	currently held in context */ %>
function PopulateApplicantCombo()
{
	<% /* Clear any <OPTION> elements from the combo */ %>
	while(frmScreen.cboApplicant.options.length > 0)
	{
		frmScreen.cboApplicant.options.remove(0);
	}

	<% /* Add a <SELECT> option */ %>
	var TagOPTION	= document.createElement("OPTION");
	TagOPTION.value	= "";
	TagOPTION.text	= "<SELECT>";
	frmScreen.cboApplicant.add(TagOPTION);

	var nCustomerCount = 0;

	<% /* Loop through all customer context entries */ %>
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + nLoop);
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		<% /* If the customer is an applicant, add him/her as an option */ %>
		<% /* SYS1672 - or guarantor */ %>
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			TagOPTION		= document.createElement("OPTION");
			TagOPTION.value	= sCustomerNumber;
			TagOPTION.text	= sCustomerName;
			TagOPTION.setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
			frmScreen.cboApplicant.add(TagOPTION);
			nCustomerCount++;
		}
	}
			
	if(nCustomerCount == 1)
	{
		frmScreen.cboApplicant.remove(0);
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboApplicant");
	}
	<% /* Default to the first option */ %>
	frmScreen.cboApplicant.selectedIndex = 0;
}

function PopulateCombos()
{
	PopulateTPTitleCombo();
	// MDC SYS2564 / SYS2785 Client Customisation
	var sControlList = new Array(frmScreen.cboPolicyType, frmScreen.cboCountry)
	objDerivedOperations.GetComboLists(sControlList);

/*	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();			
	var sGroups = new Array("PolicyType","Country","OrganisationType");
			
	if(XML.GetComboLists(document, sGroups) == true)
	{
		XML.PopulateCombo(document, frmScreen.cboPolicyType,"PolicyType",true);
		XML.PopulateCombo(document, frmScreen.cboCountry,"Country",false);
	}
	PopulateApplicantCombo();
	XML = null;			
*/
}

function PopulateScreen()
{
	if(m_sMetaAction == "Edit")
	{
		var sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
		XMLOnEntry = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XMLOnEntry.LoadXML(sXML);
				
		if(XMLOnEntry.SelectTag(null,"LIFEPRODUCT") != null)
		{
			var TagLIFEPRODUCT = XMLOnEntry.ActiveTag;
					
			frmScreen.cboApplicant.value = XMLOnEntry.GetTagText("CUSTOMERNUMBER");
			if(XMLOnEntry.SelectTag(TagLIFEPRODUCT,"ACCOUNT") != null)
			{
				frmScreen.txtAccountNumber.value = XMLOnEntry.GetTagText("ACCOUNTNUMBER");
				m_sAccountGUID = XMLOnEntry.GetTagText("ACCOUNTGUID");

				with (frmScreen)
				{
					txtCompanyName.value = XMLOnEntry.GetTagText("COMPANYNAME");
					txtContactForename.value = XMLOnEntry.GetTagText("CONTACTFORENAME");
					txtContactSurname.value = XMLOnEntry.GetTagText("CONTACTSURNAME");
					cboTitle.value = XMLOnEntry.GetTagText("CONTACTTITLE");
					<% /* PB 07/07/2006 EP543 Begin */ %>
					checkOtherTitleField();
					txtTitleOther.value = LegalRepXML.GetTagText("CONTACTTITLEOTHER");
					<% /* EP543 End */ %>

					// MDC SYS2564/SYS2785 
					objDerivedOperations.LoadCounty(XMLOnEntry);
					// txtCounty.value = XMLOnEntry.GetTagText("COUNTY");
					
					txtDistrict.value = XMLOnEntry.GetTagText("DISTRICT");
					//txtEMailAddress.value = XMLOnEntry.GetTagText("EMAILADDRESS");
					//txtFaxNo.value = XMLOnEntry.GetTagText("FAXNUMBER");
					txtFlatNumber.value = XMLOnEntry.GetTagText("FLATNUMBER");
					txtHouseName.value = XMLOnEntry.GetTagText("BUILDINGORHOUSENAME");
					txtHouseNumber.value = XMLOnEntry.GetTagText("BUILDINGORHOUSENUMBER");
					txtPostcode.value = XMLOnEntry.GetTagText("POSTCODE");
					txtStreet.value = XMLOnEntry.GetTagText("STREET");
					//txtTelephoneNo.value = XMLOnEntry.GetTagText("TELEPHONENUMBER");
					txtTown.value = XMLOnEntry.GetTagText("TOWN");
					cboCountry.value = XMLOnEntry.GetTagText("COUNTRY");
					
					m_sDirectoryGUID = XMLOnEntry.GetTagText("DIRECTORYGUID");
					m_bDirectoryAddress = (m_sDirectoryGUID != "");
					m_sThirdPartyGUID = XMLOnEntry.GetTagText("THIRDPARTYGUID");
				}
			}

			if(XMLOnEntry.SelectTag(TagLIFEPRODUCT,"LIFEPOLICY") != null)
			{
				frmScreen.cboPolicyType.value = XMLOnEntry.GetTagText("POLICYTYPE");
				frmScreen.txtMonthlyPremium.value = XMLOnEntry.GetTagText("MONTHLYPREMIUM");
				frmScreen.txtYearOfMaturity.value = XMLOnEntry.GetTagText("YEAROFMATURITY");
				frmScreen.txtMinimumDeathBenefit.value = XMLOnEntry.GetTagText("MINIMUMDEATHBENEFIT");
				frmScreen.txtMaturityValue.value = XMLOnEntry.GetTagText("MATURITYVALUE");
			}

			if(XMLOnEntry.SelectTag(TagLIFEPRODUCT,"MORTGAGERELATEDCONTRACTS") != null)
			{
				scScreenFunctions.SetRadioGroupValue(frmScreen, "NewPolicyInd", XMLOnEntry.GetTagText("NEWPOLICYINDICATOR"));
				scScreenFunctions.SetRadioGroupValue(frmScreen, "grpLIFEASSUREDINDICATOR", XMLOnEntry.GetTagText("LIFEASSUREDINDICATOR"));
				scScreenFunctions.SetRadioGroupValue(frmScreen, "AdditionalInd", XMLOnEntry.GetTagText("ADDITIONALINDICATOR"));
				frmScreen.txtAdditionalDetails.value = XMLOnEntry.GetTagText("ADDITIONALDETAILS");
			}
			if(XMLOnEntry.SelectTag(TagLIFEPRODUCT,"APPLICATIONCONTRACT") != null)
			{
				scScreenFunctions.SetRadioGroupValue(frmScreen, "AssignToNewLoanInd", XMLOnEntry.GetTagText("ASSIGNEDTOLOANINDICATOR"));
			}
			
			// 14/09/2001 JR OmiPlus 24 
			var TempXML = XMLOnEntry.ActiveTag;
			var ContactXML = XMLOnEntry.SelectTag(null, "CONTACTDETAILS");
			if(ContactXML != null)
				m_sXMLContact = ContactXML.xml;
			XMLOnEntry.ActiveTag = TempXML;
		}
	}
}

function RetrieveContextData()
{
	/*JR Test
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
	scScreenFunctions.GetContextParameter(window, "idApplicationNumber", "C00078174");
	scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", "1");
	End*/
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);

	var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bCanAddToDirectory = (sUserRole >= XML.GetGlobalParameterAmount(document,"AddToDirectoryRole"));

	XML = null;
}

function SaveLifePolicy()
{
	var bOK = false;
	var TagRequestType = null;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	if(m_sMetaAction == "Add")
		TagRequestType = XML.CreateRequestTag(window, "CREATE");
	else if(m_sMetaAction == "Edit")
		TagRequestType = XML.CreateRequestTag(window, "UPDATE");
	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");

	if(TagRequestType != null)
	{
		var TagLIFEPRODUCT = XML.CreateActiveTag("LIFEPRODUCT");
		XML.CreateActiveTag("POLICYRELATIONSHIP");

		<% /* For an update we need to specify the Account GUID */ %>
		if(m_sMetaAction == "Edit")
			XML.CreateTag("ACCOUNTGUID",m_sAccountGUID);
		var strCustomerNumber = frmScreen.cboApplicant.value;
		XML.CreateTag("CUSTOMERNUMBER", strCustomerNumber);

		var nSelectedCustomer = frmScreen.cboApplicant.selectedIndex;
		var TagOption = frmScreen.cboApplicant.options.item(nSelectedCustomer);
								
		XML.CreateTag("CUSTOMERVERSIONNUMBER",TagOption.getAttribute("CustomerVersionNumber"));
		XML.CreateTag("RELATIONSHIP",null);

		XML.ActiveTag = TagLIFEPRODUCT;
		XML.CreateActiveTag("LIFEPOLICY");

		<% /* For an update we need to specify the Account GUID */ %>
		if(m_sMetaAction == "Edit")
			XML.CreateTag("ACCOUNTGUID",m_sAccountGUID);
		XML.CreateTag("MATURITYVALUE",frmScreen.txtMaturityValue.value);
		XML.CreateTag("MINIMUMDEATHBENEFIT",frmScreen.txtMinimumDeathBenefit.value);
		XML.CreateTag("MONTHLYPREMIUM",frmScreen.txtMonthlyPremium.value);
		XML.CreateTag("POLICYTYPE",frmScreen.cboPolicyType.value);
		XML.CreateTag("YEAROFMATURITY",frmScreen.txtYearOfMaturity.value);
				
		XML.ActiveTag = TagLIFEPRODUCT;
		XML.CreateActiveTag("MORTGAGERELATEDCONTRACTS");			

		<% /* For an update we need to specify the Account GUID */ %>
		if(m_sMetaAction == "Edit")
			XML.CreateTag("ACCOUNTGUID",m_sAccountGUID);
		XML.CreateTag("CUSTOMERNUMBER", strCustomerNumber);
		XML.CreateTag("CUSTOMERVERSIONNUMBER",TagOption.getAttribute("CustomerVersionNumber"));
		XML.CreateTag("NEWPOLICYINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"NewPolicyInd"));
		XML.CreateTag("LIFEASSUREDINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"grpLIFEASSUREDINDICATOR"));
		XML.CreateTag("ADDITIONALINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"AdditionalInd"));
		XML.CreateTag("ADDITIONALDETAILS",frmScreen.txtAdditionalDetails.value);
				
		XML.ActiveTag = TagLIFEPRODUCT;
		XML.CreateActiveTag("APPLICATIONCONTRACT");			

		<% /* For an update we need to specify the Account GUID */ %>
		if(m_sMetaAction == "Edit")
			XML.CreateTag("ACCOUNTGUID",m_sAccountGUID);
		XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.CreateTag("ASSIGNEDTOLOANINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"AssignToNewLoanInd"));

		<%/* ACCOUNT */%>
		XML.ActiveTag = TagLIFEPRODUCT;
		XML.CreateActiveTag("ACCOUNT");
		if(m_sMetaAction == "Edit")
			XML.CreateTag("ACCOUNTGUID",m_sAccountGUID);
		XML.CreateTag("ACCOUNTNUMBER",frmScreen.txtAccountNumber.value);

		if (m_sMetaAction == "Edit")
		{
			// Retrieve the original third party/directory GUIDs
			var sOriginalThirdPartyGUID = XMLOnEntry.GetTagText("THIRDPARTYGUID");
			var sOriginalDirectoryGUID = XMLOnEntry.GetTagText("DIRECTORYGUID");
			// Only retrieve the address/contact details GUID if we are updating an existing third party record
			var sAddressGUID = (sOriginalThirdPartyGUID != "") ? XMLOnEntry.GetTagText("ADDRESSGUID") : "";
			var sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? XMLOnEntry.GetTagText("CONTACTDETAILSGUID") : "";
		}

		// If changing from a directory to a third party, or vice-versa, the old directory/thirdparty GUID
		// should still be specified to alert the middler tier to the fact that the old link needs deleting
		XML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID != "" ? m_sDirectoryGUID : sOriginalDirectoryGUID);
		XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);

		if (!m_bDirectoryAddress && !AllFieldsEmpty()) // Note that AllFieldsEmpty is in the ThirdPartyDetails.asp
		{
			// Store the third party details
			XML.CreateActiveTag("THIRDPARTY");
			XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
			XML.CreateTag("THIRDPARTYTYPE", 9); // Life policy
			XML.CreateTag("COMPANYNAME", frmScreen.txtCompanyName.value);
			
			// Address
			XML.CreateTag("ADDRESSGUID", sAddressGUID);
			SaveAddress(XML, sAddressGUID);

			// Contact Details
			XML.SelectTag(null, "THIRDPARTY");
			XML.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
			SaveContactDetails(XML, sContactDetailsGUID);
		}
				
		if(m_sMetaAction == "Add")
			// 			XML.RunASP(document, "CreateLifeProduct.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "CreateLifeProduct.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

		else
		{			
			<% /* A PREVIOUS KEY section may need to be created because we may have 
				changed the customer to who the loan or liability applies*/ %>					
			var OldXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
			OldXML.LoadXML(XMLOnEntry);			
			if (OldXML.SelectTag(null,"LIFEPRODUCT") != null)
			{
				var strOrigCustomerNumber = OldXML.GetTagText("CUSTOMERNUMBER");					
				if (strOrigCustomerNumber != strCustomerNumber)
				{						
					XML.CreateActiveTag("PREVIOUSKEY");
					XML.CreateActiveTag("LIFEPRODUCT");
					XML.CreateTag("CUSTOMERNUMBER", strOrigCustomerNumber);
					XML.CreateTag("CUSTOMERVERSIONNUMBER", OldXML.GetTagText("CUSTOMERVERSIONNUMBER"));
					XML.CreateTag("SEQUENCENUMBER", OldXML.GetTagText("SEQUENCENUMBER"));
				}
			}
			OldXML = null;	
			// 			XML.RunASP(document, "UpdateLifeProduct.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "UpdateLifeProduct.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

		}
		bOK = XML.IsResponseOK();
	}

	XML = null;
	return bOK;
}

function ValidateScreen()
{
	if(scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true))
		return false;

	if (scScreenFunctions.ValidatePostcode(frmScreen.txtPostcode))
		return(ValidateThirdPartyDetails());
	else
		return false;
}

function ValidateThirdPartyDetails()
{
	<%/* If third party details have been specified then so must a lender name (note that the
	     call to AllFieldsEmpty only checks the THIRDPARTY related fields) */%>
	if((frmScreen.txtCompanyName.value == "") && !AllFieldsEmpty())
	{
		alert("Lender details have been specified, therefore the Lender Name cannot be blank.");
		scScreenFunctions.DoFocusProcessing(frmScreen,"txtCompanyName");
		return(false);
	}

	//if((frmScreen.txtCompanyName.value != "") && (frmScreen.cboCountry.selectedIndex == 0))
	//{
	//	alert("The Country must be specified for the lender.");
	//	scScreenFunctions.DoFocusProcessing(frmScreen,"txtCompanyName");
	//	return(false);
	//}

	return(true);
}

function frmScreen.txtAdditionalDetails.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtAdditionalDetails", 255, true);
}

-->
</script>
<!-- #include FILE="includes/thirdpartydetails.asp" -->
</body>
</html>


