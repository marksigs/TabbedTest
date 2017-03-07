<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc095.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Lenders popup
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		03/02/2000	ie5 changes 
AY		30/03/00	scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
BG		18/05/00	SYS0752 Changed Search Button label to Directory Search
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog    Date        Description
TW      09/10/2002	Modified to incorporate client validation - SYS5115
MV		06/12/2002	Modified PAF Search Tab Order
TLiu	02/09/2005	MAR38 Changed layout for Flat No., House Name & No.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Lender Details</title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="Validation.js" language="JScript"></script>

<form id="frmScreen" mark validate="onchange" style="VISIBILITY: hidden">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 368px; WIDTH: 454px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Name
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<input id="txtCompanyName" name="CompanyName" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	<span style="TOP: 5px; LEFT: 350px; POSITION: ABSOLUTE">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 94px" class="msgButton">
	</span>
	<span style="TOP: 36px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Lender Type
		<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
			<select id="cboLenderType" name="LenderType" style="WIDTH: 150px" class="msgCombo">
			</select>
		</span>
	</span>
	<span id="spnAddress" style="TOP: 56px; LEFT: 0px; POSITION: ABSOLUTE" class="msgLabel">
		<span style="TOP: 6px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			Postcode
			<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
				<input id="txtPostcode" name="Postcode" maxlength="8" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxtUpper">
			</span>
		</span>
		<span style="TOP: 32px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			House No. &amp; Name
			<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
				<input id="txtBuildingNumber" name="BuildingNumber" maxlength="10" style="WIDTH: 45px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 150px; POSITION: ABSOLUTE">
				<input id="txtBuildingName" name="BuildingNameName" maxlength="40" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<span style="TOP: 1px; LEFT: 350px; POSITION: ABSOLUTE">
			<input id="btnPAFSearch" value="PAF Search" type="button" style="WIDTH: 100px" class="msgButton">
		</span>
		
		<span style="TOP: 58px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			Street
			<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
				<input id="txtStreet" name="Street" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<span style="TOP: 84px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			District
			<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
				<input id="txtDistrict" name="District" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<span style="TOP: 110px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			Town
			<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
				<input id="txtTown" name="Town" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<span style="TOP: 136px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			County
			<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
				<input id="txtCounty" name="County" maxlength="40" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<span style="TOP: 162px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			Country
			<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
				<select id="cboCountry" name="Country" style="WIDTH: 200px" class="msgCombo">
				</select>
			</span>
		</span>
		<span style="TOP: 188px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			Telephone Number
			<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
				<input id="txtTelephone" name="Telephone" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<span style="TOP: 214px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			Fax Number
			<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
				<input id="txtFax" name="Fax" maxlength="30" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
		<span style="TOP: 240px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			E-mail Address
			<span style="TOP: -3px; LEFT: 100px; POSITION: ABSOLUTE">
				<input id="txtContactEmailAddress" name="ContactEmailAddress" maxlength="100" style="WIDTH: 346px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
	</span>
	<span style="TOP: 83px; LEFT: 350px; POSITION: ABSOLUTE">
		<input id="btnClear" value="Clear Address" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
	<span style="TOP: 340px; LEFT: 4px; POSITION: ABSOLUTE">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
		<span style="TOP: 0px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="TOP: 0px; LEFT: 140px; POSITION: ABSOLUTE">
			<input id="btnAddToDirectory" value="Add to Directory" type="button" style="WIDTH: 100px" class="msgButton">
		</span>
	</span>
</div>
</form>
	
<!--  #include FILE="attribs/dc095attribs.asp" -->
<!--  #include FILE="includes/pafsearch.asp" -->

<script language="JScript">
<!--		
var m_sRequestAttributes = null		
<% /* Type of company */ %>
var m_sNameAndAddressType = null;	
<% /* GUID parameters */ %>
var m_sDirectoryGUID = "";
var m_sThirdPartyGUID = "";
var m_sAddressGUID = "";
var m_sContactDetailsGUID = "";		
<% /* Flags */ %>
var m_bPAFIndicator	= false;
var m_bDirectoryAddressIndicator = false;
var scScreenFunctions;
var m_BaseNonPopupWindow = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sDirectoryGUID = sArgArray[0];
	m_sThirdPartyGUID = sArgArray[1];
	m_sRequestAttributes = sArgArray[2];
	m_sNameAndAddressType = sArgArray[3];

	GetComboLists();
	scScreenFunctions.ShowCollection(frmScreen);
			
	GetRecordsOnEntry();
	SetLenderTypeCombo();						
	SetMasks();
	Validation_Init();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
			
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

<% /* Populates all combos with their options */ %>
function GetComboLists()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups;
			
	if(m_sNameAndAddressType == "3")
		sGroups = new Array("OrganisationType","Country");
	else
		sGroups = new Array("Country");
			
	if(XML.GetComboLists(document, sGroups) == true)
	{
		if(m_sNameAndAddressType == "3") XML.PopulateCombo(document, frmScreen.cboLenderType,"OrganisationType",true);
		XML.PopulateCombo(document, frmScreen.cboCountry,"Country",false);
	}
	XML = null;
}

<% /* Retrieves the data and sets the screen accordingly */ %>
function GetRecordsOnEntry()
{
	var SearchXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bSearch	= false;

	if(m_sDirectoryGUID != "" || m_sThirdPartyGUID != "")
	{
		<% /* <REQUEST USERID=? USERTYPE=? UNIT=?> */ %>
		SearchXML.CreateRequestTagFromArray(m_sRequestAttributes);
		SearchXML.CreateActiveTag("SEARCH");

		if(m_sDirectoryGUID != "")
		{
			SearchXML.CreateActiveTag("NAMEANDADDRESSDIRECTORY");
			SearchXML.CreateTag("DIRECTORYGUID",m_sDirectoryGUID);
			m_bDirectoryAddressIndicator = true;
		}
		else
		{
			if(m_sThirdPartyGUID != "")
			{
				SearchXML.CreateActiveTag("THIRDPARTY");
				SearchXML.CreateTag("THIRDPARTYGUID",m_sThirdPartyGUID);
			}
		}					
		SearchXML.RunASP(document, "GetThirdParty.asp");
				
		if(SearchXML.IsResponseOK())
		{
			PopulateScreen(SearchXML);
		}
	}
}
		
function PopulateScreen(XML)
{
	<% /* If this is a directory record */ %>
	if(m_bDirectoryAddressIndicator)
	{
		XML.SelectTag(null,"NAMEANDADDRESSDIRECTORY");
		m_sDirectoryGUID = XML.GetTagText("DIRECTORYGUID");
	}
	else
	{
		XML.SelectTag(null,"THIRDPARTY");
		if(m_sNameAndAddressType == "3") frmScreen.cboLenderType.value = XML.GetTagText("THIRDPARTYTYPE");

		m_sThirdPartyGUID = XML.GetTagText("THIRDPARTYGUID");
		m_sAddressGUID = XML.GetTagText("ADDRESSGUID");
		m_sContactDetailsGUID = XML.GetTagText("CONTACTDETAILSGUID");
	}
	frmScreen.txtCompanyName.value = XML.GetTagText("COMPANYNAME");
	frmScreen.txtPostcode.value	= XML.GetTagText("POSTCODE");
	frmScreen.txtBuildingName.value = XML.GetTagText("BUILDINGORHOUSENAME");
	frmScreen.txtBuildingNumber.value = XML.GetTagText("BUILDINGORHOUSENUMBER");
	frmScreen.txtStreet.value = XML.GetTagText("STREET");
	frmScreen.txtDistrict.value	= XML.GetTagText("DISTRICT");
	frmScreen.txtTown.value	= XML.GetTagText("TOWN");
	frmScreen.txtCounty.value = XML.GetTagText("COUNTY");
	frmScreen.cboCountry.value = XML.GetTagText("COUNTRY");
	frmScreen.txtTelephone.value = XML.GetTagText("TELEPHONENUMBER");
	frmScreen.txtFax.value = XML.GetTagText("FAXNUMBER");
	frmScreen.txtContactEmailAddress.value = XML.GetTagText("EMAILADDRESS");
			
	m_bPAFIndicator = (XML.GetTagText("PAFINDICATOR") == "1");
}
		
function ValidateIndicators(bPAFIndicator)
{
	if(bPAFIndicator)
	{
		if(m_bPAFIndicator) m_bPAFIndicator = false;
	}
	if(m_bDirectoryAddressIndicator) m_bDirectoryAddressIndicator = false;
	SetLenderTypeCombo();
}

function SetLenderTypeCombo()
{
	<% /* The lender type combo is only applicable to thirdparty records */ %>
	if(m_bDirectoryAddressIndicator || m_sNameAndAddressType != "3")
		scScreenFunctions.SetFieldToDisabled(frmScreen,"cboLenderType");
	else
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboLenderType");
}
		
function frmScreen.txtCompanyName.onchange()
{
	ValidateIndicators(false);
}

function frmScreen.btnClear.onclick()
{
	scScreenFunctions.ClearCollection(spnAddress);
	ValidateIndicators(true);
}

function frmScreen.txtPostcode.onchange()
{
	ValidateIndicators(true);
}

function frmScreen.txtBuildingName.onchange()
{
	ValidateIndicators(true);
}

function frmScreen.txtBuildingNumber.onchange()
{
	ValidateIndicators(true);
}

function frmScreen.txtStreet.onchange()
{
	ValidateIndicators(true);
}

function frmScreen.txtDistrict.onchange()
{
	ValidateIndicators(true);
}

function frmScreen.txtTown.onchange()
{
	ValidateIndicators(true);
}

function frmScreen.txtCounty.onchange()
{
	ValidateIndicators(true);
}

function frmScreen.cboCountry.onchange()
{
	ValidateIndicators(true);
}

function frmScreen.txtTelephone.onchange()
{
	ValidateIndicators(false);
}

function frmScreen.txtFax.onchange()
{
	ValidateIndicators(false);
}

function frmScreen.txtContactEmailAddress.onchange()
{
	ValidateIndicators(false);
}

function frmScreen.btnDirectorySearch.onclick()
{
	if (m_sReadOnly == "1") return;

	//build up XML for directory search
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var ArrayArguments = XML.CreateRequestAttributeArray(window)
	var sCompanyName = frmScreen.txtCompanyName.value;
	
	/* Route to ZA020 */
	ArrayArguments[4] = m_sNameAndAddressType;
	ArrayArguments[5] = sCompanyName;

	var sReturn = scScreenFunctions.DisplayPopup(window, document, "ZA020.asp", ArrayArguments, 630, 440);
	if(sReturn != null)
	{
		FlagChange(sReturn[0]);
		XML.LoadXML(sReturn[1]);
		PopulateDirectoryAddress(XML);

		m_bDirectoryAddressIndicator = true;
		var ReturnXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		ReturnXML.LoadXML(sReturn[1]);
		PopulateScreen(ReturnXML);
		ReturnXML = null;
	}

	frmScreen.txtCompanyName.focus();
	XML = null;
	ArrayArguments = null;
}

function frmScreen.btnPAFSearch.onclick()
{
	with (frmScreen)
		m_bPAFIndicator = PAFSearch(txtPostcode,txtBuildingName,txtBuildingNumber,null,txtStreet,txtDistrict,txtTown,txtCounty,cboCountry);
}

function frmScreen.btnOK.onclick()
{
	var bExit = false;
	var ResponseXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	if (frmScreen.onsubmit())
	{
		<% /* Start generating the XML to pass back to the calling screen */ %>
		ResponseXML.CreateActiveTag("RESPONSE");
		ResponseXML.CreateTag("COMPANYNAME",frmScreen.txtCompanyName.value);
			
		<% /* If this is a directory record */ %>
		if(m_bDirectoryAddressIndicator)
		{
			ResponseXML.CreateTag("DIRECTORYGUID",m_sDirectoryGUID);
			ResponseXML.CreateTag("THIRDPARTYGUID",m_sThirdPartyGUID);
			bExit = true;
		}
		else
		{
			<% /* We are creating or updating a ThirdParty record */ %>
			var XML = GenerateCommitXML(false);
				
			<% /* SaveThirdParty can perform a delete. 
				Ensure the XML is never set up to do this */ %>
			// 			XML.RunASP(document, "SaveThirdParty.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "SaveThirdParty.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

				
			<% /* If there is a successful response, check which type it is:
				CREATE	- get the new guid value
				UPDATE	- pass back the guid as is */ %>						
			if(XML.IsResponseOK())
			{
				XML.SelectTag(null,"RESPONSE");
				var sType = XML.GetAttribute("OPERATION");
					
				if(sType == "CREATE")
				{
					m_sThirdPartyGUID = XML.GetTagText("THIRDPARTYGUID");
				}
				ResponseXML.CreateTag("THIRDPARTYGUID",m_sThirdPartyGUID);
				bExit = true;
			}
		}
	}
	if(bExit)
	{
		window.returnValue = ResponseXML.XMLDocument.xml;
		window.close();
	}
}

function frmScreen.btnCancel.onclick()
{
	window.close();
}
		
function GenerateCommitXML(bCreateDirectory)
{
	<% /* bCreateDirectory is set to true when we are 
		generating the XML to pass to ZA030 */ %>
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	<% /* <REQUEST USERID=? USERTYPE=? UNIT=?> */ %>
	XML.CreateRequestTagFromArray(m_sRequestAttributes);
				
	var TagParent;			
	if(bCreateDirectory)
	{
		XML.CreateActiveTag("CREATE");
		TagParent = XML.CreateActiveTag("NAMEANDADDRESSDIRECTORY");
		XML.CreateTag("NAMEANDADDRESSTYPE",m_sNameAndAddressType);
	}
	else
	{
		XML.CreateActiveTag("SAVE");
		TagParent = XML.CreateActiveTag("THIRDPARTY");
				
		if(m_sThirdPartyGUID != "") 
			XML.CreateTag("THIRDPARTYGUID",m_sThirdPartyGUID);
		if(m_sNameAndAddressType == "3")
			XML.CreateTag("THIRDPARTYTYPE",frmScreen.cboLenderType.value);
		else
			XML.CreateTag("THIRDPARTYTYPE","9");
	}
	XML.CreateTag("COMPANYNAME",frmScreen.txtCompanyName.value);
				
	if(m_sAddressGUID != "" && !bCreateDirectory)
		XML.CreateTag("ADDRESSGUID",m_sAddressGUID);
				
	<% /* In the case of no existing address record, check 
		whether any address fields other than the country field are set */ %>
	if(m_sAddressGUID != "" || (m_sAddressGUID == "" && !AreAddressFieldsBlank()))
	{				
		XML.CreateActiveTag("ADDRESS");
				
		if(m_sAddressGUID != "" && !bCreateDirectory) XML.CreateTag("ADDRESSGUID",m_sAddressGUID);

		var sPAFIndicator = "0";
		if(m_bPAFIndicator) sPAFIndicator = "1";

		XML.CreateTag("BUILDINGORHOUSENAME",frmScreen.txtBuildingName.value);
		XML.CreateTag("BUILDINGORHOUSENUMBER",frmScreen.txtBuildingNumber.value);
		XML.CreateTag("STREET",frmScreen.txtStreet.value);
		XML.CreateTag("DISTRICT",frmScreen.txtDistrict.value);
		XML.CreateTag("TOWN",frmScreen.txtTown.value);
		XML.CreateTag("COUNTY",frmScreen.txtCounty.value);
		XML.CreateTag("COUNTRY",frmScreen.cboCountry.value);
		XML.CreateTag("POSTCODE",frmScreen.txtPostcode.value);
		XML.CreateTag("PAFINDICATOR",sPAFIndicator);
	}
	XML.ActiveTag = TagParent;
	XML.CreateActiveTag("CONTACTDETAILS");
			
	if(m_sContactDetailsGUID != "" && !bCreateDirectory) XML.CreateTag("CONTACTDETAILSGUID",m_sContactDetailsGUID);

	XML.CreateTag("EMAILADDRESS",frmScreen.txtContactEmailAddress.value);
	XML.CreateTag("FAXNUMBER",frmScreen.txtFax.value);
	XML.CreateTag("TELEPHONENUMBER",frmScreen.txtTelephone.value);
	return XML;	
}
		
function AreAddressAndContactFieldsBlank()
{
	var bAreBlank = false;
			
	if(AreAddressFieldsBlank() && AreContactFieldsBlank())
	{
		bAreBlank = true;
	}			
	return bAreBlank;
}
		
function AreAddressFieldsBlank()
{
	var bAreBlank = false;
			
	if(frmScreen.txtPostcode.value == "" && frmScreen.txtBuildingName.value == ""
	   && frmScreen.txtBuildingNumber.value == "" && frmScreen.txtStreet.value == ""
	   && frmScreen.txtDistrict.value == "" && frmScreen.txtTown.value == "" && frmScreen.txtCounty.value == "")
	{
		bAreBlank = true;
	}
	return bAreBlank;
}

function AreContactFieldsBlank()
{
	var AreBlank = false;
	if(frmScreen.txtTelephone.value == "" && frmScreen.txtFax.value == "" && frmScreen.txtContactEmailAddress.value == "")
	{
		bAreBlank = true;
	}
	return bAreBlank;
}

function frmScreen.btnAddToDirectory.onclick()
{
	if (frmScreen.onsubmit() == true)
	{
		var XML = GenerateCommitXML(true);
		
		var ArrayArguments = m_sRequestAttributes;
		ArrayArguments[4] = XML.XMLDocument.xml;
		var sReturn = scScreenFunctions.DisplayPopup(window,document,"za030.asp",ArrayArguments,630,410);
		if(sReturn != null)
		{
			FlagChange(sReturn[0]);
			m_sDirectoryGUID = sReturn[1];
			m_bDirectoryAddressIndicator = true;
		}
	}
}
-->
</script>
</body>
</html>


