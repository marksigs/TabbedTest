<% /*
	SR		22/05/00	Modified frmScreen.btnAddToDirectory.onclick() - changed name of tag from 
						BankSortCode to NameAndAdressBankSortCode						
	SR		31/05/00	SYS0683 - Clear button should not clear Company Name
	BG		15/08/00	Added function IsScreenClear() to check to see if anything has been entered
						on the ThirdPartyDetails.htm screen - for SYS1428 screen DC071.
	BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
						NB This change was made to be consistent with the way Supervisor Stores combo 
						values for titles.  Due to this change in this file we are currently storing
						the combo value ID (numeric) in a column of type String, as this used to hold 
						the input title in the text box Contact Title that was on this screen before 
						this change.
	SA		05/06/01	SYS0923 In function DirectoryEntryAlreadyExists - need to cope with window being a popup.
						Call to DirectorysEarch() - add in Sortcode parameter.
	MDC		15/06/01	OmiPlus24 Telephone Numbers						
	MDC		08/10/01	SYS2785 Client specific cosmetic customisation
	MDC		06/12/01	SYS3408 Fix IsCountyEmpty function
	STB		21/02/02	SYS4091 Fix IsContactEmpty so that telephone numbers and e-mail adresses are also checked.
	SA		22/04/02	SYS4439 County set to read only - new function	
	MEVA	25/04/02	SYS1154 Provide a default Contact Type
	LD		23/05/02	SYS4727 Use cached versions of frame functions
	---------BMIDS----------------------------------------------------------------
	GD		24/06/02	BMIDS Fixed SYS4727 for core.
	GD		26/06/02	BMIDS0077 Applied SYS4930
	MV		12/07/02	BMIDS00198 - Display DxID and DxLocation in ThirdPartyDetailsDirectorySearch() 
	MV		31/07/02	BMIDS0075 - Modified ThirdPartyDetailsDirectorySearch()
	MV		23/08/02	BMIDS00355	- Modified Width and Height of DC241.asp
	PSC		26/11/02	BMIDS00998 - Modify PopulateDirectoryAddress to cater for Address and ContactDetails
	                                 not being present
	PSC		28/11/02	BMIDS00900 - Modify to take into account blank details
	MDC		12/12/2002	BM0094 - Legal Rep Contact Details
	DRC     09/02/2004  BMIDS690 - Stop duplicating of contactelephone details when changing from Directory to thirdparty
	DRC     11/03/2002  BMIDS690 Revisited - Bugfix
	
	
	MARS History:

	Prog	Date		Description
	KRW     09/11/2005  MAR180 AddressChanged() // MAR180 KRW 	
	KRW     09/11/2005  Attach code to clearquest no change // MAR180 KRW
	PJO     21/11/2005  MAR 633 Ensure Country combo has <SELECT>
	
	EPSOM History
	
	Prog	Date		Description
	PE		22/06/2006	EP828 Ensure AccountantQualifications has select
	PB		06/07/2006	EP543 Allow free-text field for title-other
	
	SR		20/04/2007	EP2_2413 - Show <Select> for combo LandlordType
	*/			
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
%>

<!-- #include FILE="directorysearch.asp" -->
<!-- #include FILE="pafsearch.asp" -->
<!-- #include FILE="../Customise/ThirdPartyCustomise.asp" -->

<script language="JScript">
<!--
var m_fSetAvailableFunctionalityOverride = null;
var m_fValidateScreen = null;

var m_bIsPopup = false;
var m_sThirdPartyType = "0";
var m_bCanAddToDirectory = true;
var m_bDirectoryOnly = false;

var m_bDirectoryAddress = false;
var m_sDirectoryGUID = "";
var m_sThirdPartyGUID = "";
var m_bPAFIndicator = false;

var m_ctrSortCode = null;
var m_ctrBranchName = null;
var m_ctrOrganisationType = null;
var m_sContactType = "99";

var m_sXMLContact = "";
var objDerivedOperations;

var sNameAndAddressActiveFROMDate = "";
var sNameAndAddressActiveTODate = "";

var sDxID = "";
var sDxLocation = "";

var bLegalRep = false; <% /* BM0094 MDC 12/12/2002 */ %>

var ThirdPartyDetailsXML = null;

function BaseScreen(sBaseGroupList)
{
	<% // Base methods %>
	this.SaveCounty = SaveCounty;
	this.LoadCounty = LoadCounty;
	this.ClearCounty = ClearCounty;
	this.IsCountyEmpty = IsCountyEmpty;
	//SYS4439
	this.SetCountyToReadOnly = SetCountyToReadOnly;
	this.GetComboLists = GetComboLists;
	this.PopulateDerivedCombos = PopulateDerivedCombos;
	<% // Allow clients to override the number and type of combos to be obtained from the middle tier %>
	this.sGroupList = sBaseGroupList; //new Array("Country","LandlordType");
	<% // Allow clients to populate their own combos over and above the ones used already in this screen %>
	function PopulateDerivedCombos(XML){ return(true);}
	function SaveCounty(XML)
	{
		XML.CreateTag("COUNTY", frmScreen.txtCounty.value);	
	}
	function LoadCounty(XML)
	{
		frmScreen.txtCounty.value = XML.GetTagText("COUNTY");
	}
	function IsCountyEmpty()
	{
		if(frmScreen.txtCounty.value == "")
			return true;
		else
			return false;
	}
	function ClearCounty()
	{
		frmScreen.txtCounty.value = "";	
	}
	//SYS4439 County is a client configurable field.
	function SetCountyToReadOnly()
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtCounty");
	}
	
	function GetComboLists(sControlList)
	{
		var blnSuccess = true;
		//GD BMIDS0077
		<%/*SG 25/06/02 SYS4930 START */%>
		if (m_bIsPopup)
			var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		else
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		<%/*SG 25/06/02 SYS4930 END */%>

		if(XML.GetComboLists(document,this.sGroupList))
		{
			// Populate each combo
			for (var intLoop = 0; intLoop < sControlList.length; intLoop++)
			{
				// PJO 21/11/2005 MAR633 - Ensure Country has select
				// PE  22/06/2006 EP828  - Ensure AccountantQualifications has select
				// SR  20/04/2007 EP2_2413 - Add <Select> for combo LandlordType
				blnSuccess = blnSuccess && XML.PopulateCombo(document,
															 sControlList[intLoop],
															 this.sGroupList[intLoop],
															 this.sGroupList[intLoop] == "Country" || this.sGroupList[intLoop] == "AccountantQualifications" || this.sGroupList[intLoop]=="LandlordType");
			}
			this.PopulateDerivedCombos(XML);
			
			if(blnSuccess == false)
				scScreenFunctions.SetScreenToReadOnly(frmScreen);

		}
	}
}

function IsScreenClear()
{	
	var blnReturn = false;
	
	if(frmScreen.txtPostcode.value == "" && frmScreen.txtFlatNumber.value == "" &&
		frmScreen.txtHouseName.value == "" && frmScreen.txtHouseNumber.value == "" &&
		frmScreen.txtStreet.value == "" && frmScreen.txtDistrict.value == "" &&
		frmScreen.txtTown.value == "" && objDerivedOperations.IsCountyEmpty() &&
		frmScreen.cboTitle.value == "" && frmScreen.txtContactForename.value == "" && <% /*
		PB 06/07/06 EP543 Begin
		frmScreen.txtContactSurname.value == "") */ %>
		frmScreen.txtTitleOther.value == "" && frmScreen.txtContactSurname.value == "") <% /*
		PB EP543 End */ %>
	
		blnReturn = true;
		return blnReturn;
}

function frmScreen.btnContact.onclick()
{
// Display popup DC241 - Contact Details
// =====================================
var sReturn = null;
var ArrayArguments = new Array(2);

	//ArrayArguments[0] = m_sReadOnly;
	ArrayArguments[0] = IsReadOnly();
	ArrayArguments[1] = m_sXMLContact;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "dc241.asp", ArrayArguments, 450, 290);
	if(sReturn != null)
	{
		// Save returned XML
		FlagChange(sReturn[0]);
		m_sXMLContact = sReturn[1];
	}

}

function frmScreen.btnAddToDirectory.onclick()
// Calls pop-up ZA030
{
	if (!frmScreen.onsubmit()) return;
	if (m_fValidateScreen != null)
		if (!m_fValidateScreen()) return;
	if (DirectoryEntryAlreadyExists())
	{
		alert("A directory entry for this company address already exists. Please select that entry from " +
              "the Directory Search screen.");
		return;
	}
	//GD BMIDS0077
	<%/*SG 25/06/02 SYS4930 START */%>
	if (m_bIsPopup)
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	else
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/*SG 25/06/02 SYS4930 END */%>

	if (m_bIsPopup)
		XML.CreateRequestTagFromArray(m_aArgArray,null);
	else
		XML.CreateRequestTag(window,null);

	var tagNAMEANDADDRESS = XML.CreateActiveTag("NAMEANDADDRESSDIRECTORY");
	XML.CreateTag("DIRECTORYGUID", "");
	XML.CreateTag("NAMEANDADDRESSTYPE", m_sThirdPartyType);
	XML.CreateTag("COMPANYNAME", frmScreen.txtCompanyName.value);
	if (m_ctrBranchName != null) XML.CreateTag("BRANCHNAME",m_ctrBranchName.value);
	if (m_ctrOrganisationType != null) XML.CreateTag("ORGANISATIONTYPE",m_ctrOrganisationType.value);
	<%/* SR 22-05-00 / SYS0686: Change the name of the tag from BankSortCode to NameAndAddressBankSortCode */%>
	if (m_ctrSortCode != null) XML.CreateTag("NAMEANDADDRESSBANKSORTCODE",m_ctrSortCode.value);
	SaveAddress(XML, "");
	XML.SelectTag(null, "NAMEANDADDRESSDIRECTORY");
	
	<%/* SR-22/05/00 SYS0698 - Do not add ContactDetails to Request tag, if no data was entered for it */%>
	if (!IsContactEmpty()) SaveContactDetails(XML, "");
	XML.SelectTag(null, "REQUEST");

	if (m_bIsPopup)
		var ArrayArguments = m_aArgArray;
	else
		var ArrayArguments = XML.CreateRequestAttributeArray(window);

	<%/* Display the ZA030 popup */%>
	ArrayArguments[4] = XML.XMLDocument.xml;
	var sReturn = scScreenFunctions.DisplayPopup(window, document, "ZA030.asp", ArrayArguments, 520, 319);

	if(sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sDirectoryGUID = sReturn[1];
		m_bDirectoryAddress = true;
		if (m_ctrOrganisationType != null) m_ctrOrganisationType.value = sReturn[2];
	}

	SetAvailableFunctionality();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function frmScreen.btnClear.onclick()
{
<% /* MH 17/05/00 SYS0698  */ %>
	ClearFields(true, true);
	frmScreen.txtPostcode.focus();
}

function frmScreen.btnDirectorySearch.onclick()
{
	ThirdPartyDetailsDirectorySearch();
}

<% /* This function is here so that DC290 specifically can override the button
  behaviour to do its own thing as well */ %>
function ThirdPartyDetailsDirectorySearch()
{

	if (m_sReadOnly == "1") return;

	//build up XML for directory search
	//GD BMIDS0077
	<%/*SG 25/06/02 SYS4930 START */%>
	if (m_bIsPopup)
		<% /* MV - 31/07/2002 - BMIDS00179 */ %>
		ThirdPartyDetailsXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	else
		<% /* MV - 31/07/2002 - BMIDS00179 */ %>
		ThirdPartyDetailsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/*SG 25/06/02 SYS4930 END */%>
	var sCompanyName = frmScreen.txtCompanyName.value;

	if (m_bIsPopup)
		var ArrayArguments = m_aArgArray;
	else
		<% /* MV - 31/07/2002 - BMIDS00179 */ %>
		var ArrayArguments = ThirdPartyDetailsXML.CreateRequestAttributeArray(window);

	<%/* Display the ZA020 popup */%>
	ArrayArguments[4] = m_sThirdPartyType;
	ArrayArguments[5] = sCompanyName;
	var sReturn = scScreenFunctions.DisplayPopup(window, document, "ZA020.asp", ArrayArguments, 630, 450);

	if(sReturn != null)
	{
		FlagChange(sReturn[0]);
		<% /* MV - 31/07/2002 - BMIDS00179 */ %>
		ThirdPartyDetailsXML.LoadXML(sReturn[1]);
		PopulateDirectoryAddress(ThirdPartyDetailsXML);
		
		var xmlActiveTag = ThirdPartyDetailsXML.SelectTag(null, "NAMEANDADDRESSDIRECTORY");
		sNameAndAddressActiveFROMDate = ThirdPartyDetailsXML.GetTagText("NAMEANDADDRESSACTIVEFROM");
		sNameAndAddressActiveTODate = ThirdPartyDetailsXML.GetTagText("NAMEANDADDRESSACTIVETO");
		
		<% /* MV - 12/07/2002 - BMIDS00198 - Display DxID and DxLocation*/%>		
		sDxID =  ThirdPartyDetailsXML.GetTagText("DXID");
		sDxLocation =  ThirdPartyDetailsXML.GetTagText("DXLOCATION");
		
		<% /* MV - 31/07/2002 - BMIDS00179 */ %>					
		ThirdPartyDetailsXML.ActiveTag = xmlActiveTag;
	}

	XML = null;
	ArrayArguments = null;
	SetAvailableFunctionality();
	<%/* BG 16/05/2000 SYS0682 sets focus to Employer/Trading Name Text Box on return from search*/%>
	
	if (frmScreen.txtCompanyName.readOnly == false) 
	{
		frmScreen.txtCompanyName.focus();
	}
	else
		frmScreen.optInDirectoryYes.focus();
	
}

function frmScreen.btnPAFSearch.onclick()
{
	with (frmScreen)
		m_bPAFIndicator = PAFSearch(txtPostcode,txtHouseName,txtHouseNumber,txtFlatNumber,txtStreet,txtDistrict,txtTown,txtCounty,cboCountry);

	if (m_bPAFIndicator) m_bDirectoryAddress = false;
}

function AddressChanged()
{
	m_bPAFIndicator = false;
}

function AllFieldsEmpty()
{
	var bReturn = (frmScreen.txtCompanyName.value == "") &&
		(frmScreen.txtPostcode.value == "") &&
		(frmScreen.txtFlatNumber.value == "") &&
		(frmScreen.txtHouseName.value == "") &&
		(frmScreen.txtHouseNumber.value == "") &&
		(frmScreen.txtStreet.value == "") &&
		(frmScreen.txtDistrict.value == "") &&
		(objDerivedOperations.IsCountyEmpty()) &&
		(frmScreen.txtContactForename.value == "") &&
		(frmScreen.txtContactSurname.value == "")

	if (m_ctrBranchName != null) bReturn = bReturn && (m_ctrBranchName.value == "");
	if (m_ctrOrganisationType != null) bReturn = bReturn && (m_ctrOrganisationType.value == "");
	if (m_ctrSortCode != null) bReturn = bReturn && (m_ctrSortCode.value == "");

	return(bReturn);
}

function ClearFields(bAddress, bContactDetails)
// Clears all fields of data
{
	with (frmScreen)
	{
		if (bContactDetails)
		{
			<% /* SR 31/05/00 - SYS0683 : Do not clear the Company Name on clicking Clear button 
			txtCompanyName.value = "";  */ %>
			txtContactForename.value = "";
			txtContactSurname.value = "";
			cboTitle.value = ""; <% /*
			PB 06/07/2006 EP543 Begin */ %>
			txtTitleOther.value = ""; <% /*
			PB EP543 End */ %>
		}

		if (bAddress)
		{
			txtFlatNumber.value = "";
			txtHouseName.value = "";
			txtHouseNumber.value = "";
			txtPostcode.value = "";
			txtStreet.value = "";
			// MDC SYS2564 / SYS2785
			objDerivedOperations.ClearCounty();
			//txtCounty.value = "";
			
			txtDistrict.value = "";
			txtTown.value = "";
			cboCountry.selectedIndex = 0;
			m_bIsChanged = true ; // MAR180 KRW 09/11/2005
		}
	}

	SetAvailableFunctionality();
}

function ContactDetailsChanged()
{
	// Do nothing - can be added to later
}

<% /* PB 06/07/2006 EP543 Begin */ %>
function frmScreen.cboTitle.onchange()
{
	checkOtherTitleField();
}

function checkOtherTitleField()
{
	// Need to show/hide 'Other' field depending whether 'Other' has been selected or deselected.
	//if(scScreenFunctions.IsValidationType(frmScreen.cboTitle,"O");
	if(frmScreen.cboTitle.value=='99')
	{
		// 'Other' selected - show field
		document.all.txtTitleOther.style.visibility='visible';
		document.all.idTitleOther.style.visibility='visible';
	}
	else
	{
		// 'Other' not selected - hide field
		document.all.txtTitleOther.style.visibility='hidden';
		document.all.idTitleOther.style.visibility='hidden';
	}
}
<% /* PB EP543 End */ %>

function DirectoryEntryAlreadyExists()
{
	<%/* Returns whether data matching that displayed in the address details onscreen
		 already exists on the directory */%>
	var bReturn = false;
	//GD BMIDS0077
	<%/*SG 25/06/02 SYS4930 START */%>
	if (m_bIsPopup)
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	else
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/*SG 25/06/02 SYS4930 END */%>
	// SA SYS0923 If the window is a popup - use the array values already setup.
	if (m_bIsPopup)
		var AttributeArray = m_aArgArray;
	else
		var AttributeArray = XML.CreateRequestAttributeArray(window);

	//SA SYS0923 Not passing through expected argument SortCode - pass through null after "type".
	bReturn = (DirectorySearch(XML, AttributeArray,
				null,
				frmScreen.txtCompanyName.value, 
				frmScreen.txtTown.value,
				null,
				m_sThirdPartyType,
				null,
				frmScreen.txtPostcode.value,
				frmScreen.txtHouseName.value,
				frmScreen.txtHouseNumber.value,
				frmScreen.txtFlatNumber.value,
				frmScreen.txtStreet.value));

	XML = null;
	AttributeArray = null;
	return(bReturn);
}

function InDirectoryChanged()
{
	if (scScreenFunctions.GetRadioGroupValue(frmScreen,"InDirectoryGroup") == "0")
	{
		m_bDirectoryAddress = false;
		m_sDirectoryGUID = "";
		SetAvailableFunctionality();
	}
}

function PopulateDirectoryAddress(XML)
// Populates the address fields with the details retrieved from a directory search
{	
	
	with (frmScreen)
	{
		XML.SelectTag(null, "NAMEANDADDRESSDIRECTORY");
		txtCompanyName.value = XML.GetTagText("COMPANYNAME");
		m_bDirectoryAddress = true;
		m_sDirectoryGUID = XML.GetTagText("DIRECTORYGUID");

		if (m_ctrBranchName != null) m_ctrBranchName.value = XML.GetTagText("BRANCHNAME");
		if (m_ctrOrganisationType != null) m_ctrOrganisationType.value = XML.GetTagText("ORGANISATIONTYPE");
		if (m_ctrSortCode != null) m_ctrSortCode.value = XML.GetTagText("NAMEANDADDRESSBANKSORTCODE");

		XML.SelectTag(null, "ADDRESS");
		
		<% /* PSC 26/11/2002 BMIDS00998 - Start */ %>
		if(XML.ActiveTag != null)
		{
			txtPostcode.value = XML.GetTagText("POSTCODE");
			txtFlatNumber.value = XML.GetTagText("FLATNUMBER");
			txtHouseName.value = XML.GetTagText("BUILDINGORHOUSENAME");
			txtHouseNumber.value = XML.GetTagText("BUILDINGORHOUSENUMBER");
			txtStreet.value = XML.GetTagText("STREET");
			txtDistrict.value = XML.GetTagText("DISTRICT");
			txtTown.value = XML.GetTagText("TOWN");
		
			// MDC SYS2564 / SYS2785 
			objDerivedOperations.LoadCounty(XML);
		
			cboCountry.value = XML.GetTagText("COUNTRY");
		}
		<% /* PSC 26/11/2002 BMIDS00998 - End */ %>
		
		XML.SelectTag(null, "CONTACTDETAILS");
		
		<% /* PSC 26/11/2002 BMIDS00998 - Start */ %>
		if (XML.ActiveTag != null)
		{
			cboTitle.value = XML.GetTagText("CONTACTTITLE"); <% /*
			PB 06/07/2006 EP543 Begin */ %>
			checkOtherTitleField();
			txtTitleOther.value = XML.GetTagText("CONTACTTITLEOTHER"); <% /*
			PB EP543 End */ %>
			checkOtherTitleField();
			txtTitleOther.value = XML.GetTagText("CONTACTTITLEOTHER");
			txtContactForename.value = XML.GetTagText("CONTACTFORENAME");
			txtContactSurname.value = XML.GetTagText("CONTACTSURNAME");
			m_sContactType = XML.GetTagText("CONTACTTYPE");
		}
		<% /* PSC 26/11/2002 BMIDS00998 - End */ %>
		
		// 05/10/01 JR OmiPlus 24 
		var ContactXML = XML.SelectTag(null, "CONTACTDETAILS");
		if(ContactXML != null)
			m_sXMLContact = ContactXML.xml;
	}
}

function SaveAddress(XML, sAddressGUID)
{
	XML.CreateActiveTag("ADDRESS");
	XML.CreateTag("ADDRESSGUID", sAddressGUID);
	XML.CreateTag("BUILDINGORHOUSENAME", frmScreen.txtHouseName.value);
	XML.CreateTag("BUILDINGORHOUSENUMBER", frmScreen.txtHouseNumber.value);
	XML.CreateTag("FLATNUMBER", frmScreen.txtFlatNumber.value);
	XML.CreateTag("STREET", frmScreen.txtStreet.value);
	XML.CreateTag("DISTRICT", frmScreen.txtDistrict.value);
	XML.CreateTag("TOWN", frmScreen.txtTown.value);

	// MDC SYS2564 / SYS2785
	objDerivedOperations.SaveCounty(XML);

	XML.CreateTag("COUNTRY", frmScreen.cboCountry.value);
	XML.CreateTag("POSTCODE", frmScreen.txtPostcode.value);
	XML.CreateTag("PAFINDICATOR", m_bPAFIndicator ? "1" : "0");

}

function IsAddressEmpty()
{
	var bReturn = (frmScreen.txtPostcode.value == "") &&
		(frmScreen.txtFlatNumber.value == "") &&
		(frmScreen.txtHouseName.value == "") &&
		(frmScreen.txtHouseNumber.value == "") &&
		(frmScreen.txtStreet.value == "") &&
		(frmScreen.txtDistrict.value == "") &&
		(frmScreen.txtTown.value == "") &&
		(objDerivedOperations.IsCountyEmpty())	// MDC SYS2564 / SYS2785
		//(frmScreen.txtCounty.value == "")
	
	<% /* Country is not added in the above condition because, it has a default value  */  %>
	return (bReturn) ;
}

function SaveContactDetails(XML, sContactDetailsGUID)
{
	if(m_sContactType == "")
	{
		m_sContactType = "99";
	}
	//BMIDS690 check whether a string value was passed in 
	if (typeof(sContactDetailsGUID) != "string")
	   sContactDetailsGUID = "";
	  
	XML.CreateActiveTag("CONTACTDETAILS");
	XML.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
	XML.CreateTag("CONTACTFORENAME", frmScreen.txtContactForename.value);
	XML.CreateTag("CONTACTSURNAME", frmScreen.txtContactSurname.value);
	XML.CreateTag("CONTACTTITLE", frmScreen.cboTitle.value); <% /*
	PB 06/07/2006 EP543 Begin */ %>
	XML.CreateTag("CONTACTTITLEOTHER", frmScreen.txtTitleOther.value);
	XML.CreateTag("CONTACTTYPE", m_sContactType);

	if(m_sXMLContact != "")
	{
		// Add Contact Telephone Details from popup DC241
		<%/*SG 25/06/02 SYS4930 START */%>
		if (m_bIsPopup)
			var XMLContact = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		else
			var XMLContact = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		<%/*SG 25/06/02 SYS4930 END */%>
		XMLContact.LoadXML(m_sXMLContact);
		var xmlRootNode = XMLContact.ActiveTag;	// Bookmark the root
		
		XMLContact.SelectTag(null, "CONTACTDETAILS");
		// Email Address
		<% /* PSC 28/11/2002 BMIDS00900 */ %>
		XML.CreateTag("EMAILADDRESS", XMLContact.GetTagText("EMAILADDRESS"));	

		XMLContact.CreateTagList("CONTACTTELEPHONEDETAILS");
		for (var iLoop=0;iLoop < XMLContact.ActiveTagList.length; iLoop++)
		{
			<%/*DRC BMIDS690 - stop duplication of original records START */%>
			if (m_bIsPopup)
				var XMLContactTeleDetails = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			else
				var XMLContactTeleDetails = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XMLContactTeleDetails.ActiveTag = XMLContact.GetTagListItem(iLoop);
			if (XMLContactTeleDetails.GetTagText("CONTACTDETAILSGUID") != "")
			  {
			   if (sContactDetailsGUID != "")
			    {
				XMLContactTeleDetails.SetTagText("CONTACTDETAILSGUID", sContactDetailsGUID);
				}
			   else
			    {
			    XMLContactTeleDetails.ActiveTag = XMLContactTeleDetails.SelectTag(XMLContactTeleDetails.ActiveTag,"CONTACTDETAILSGUID");
			   	XMLContactTeleDetails.RemoveActiveTag();
			   	}
			  }
			else
			  {
			  if (sContactDetailsGUID != "")
			    {
				XMLContactTeleDetails.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
				}
			  }
			
			XML.ActiveTag.appendChild(XMLContactTeleDetails.ActiveTag);
			<%/*DRC BMIDS690- stop duplication of original records END */%>			
			
		}
		XMLContact.ActiveTag = xmlRootNode;	// Reset Active Tag to the root
	}
}

// Returns true if all of the contact fields are blank.
function IsContactEmpty()
{	
	// Assume the contact information is missing until we find something.
	var bEmpty = true;
	
	// If additional contact information has been entered.
	if (m_sXMLContact != "")
	{
		// Create an XML object to load the extra contact details into.
		//GD BMIDS0077
		<%/*SG 25/06/02 SYS4930 START */%>
		if (m_bIsPopup)
			var XMLContact = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		else
			var XMLContact = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		<%/*SG 25/06/02 SYS4930 END */%>
	
		// Load any e-mail addresses and telephone numbers into the XML.
		XMLContact.LoadXML(m_sXMLContact);
	
		// Check e-mail address.	
		var clsEmailNode = XMLContact.SelectTag(null, "EMAILADDRESS");
			
		// If there is an e-mail node and its not blank, then we're not empty.
		if (clsEmailNode != null) 
		{	
			if(clsEmailNode.text != "")
			{
				bEmpty = false;
			}
		}
	
		// Check phone numbers.
		if (bEmpty == true)
		{
			// Populate a tag list of telephone numbers.
			XMLContact.SelectTag(null, "CONTACTDETAILS");
			XMLContact.CreateTagList("CONTACTTELEPHONEDETAILS");
							
			// Loop through any phone numbers looking for values.
			for (var iLoop=0; iLoop < XMLContact.ActiveTagList.length; iLoop++)
			{
				// The text property of the CONTACTTELEPHONEDETAILS node is a
				// composite of all the child nodes.
				var oNode = XMLContact.GetTagListItem(iLoop);
				
				if (oNode.text != "")
				{
					bEmpty = false;
					break;
				}				
			}						
		} 
	}
	
	// Check forename, surname and title.
	if (bEmpty == true)
	{
		bEmpty = ((frmScreen.txtContactForename.value == "") &&
					(frmScreen.txtContactSurname.value == "") && 
					(frmScreen.cboTitle.value == ""))
	}
			
	return (bEmpty)
}


function SetAvailableFunctionality()
{

	<%/* 'Directory Only' functionality */%>
	if (m_bDirectoryOnly)
	{
		scScreenFunctions.SetCollectionToReadOnly(spnContactDetails);
		scScreenFunctions.SetCollectionToReadOnly(spnAddress);
		scScreenFunctions.SetRadioGroupToDisabled(frmScreen,"InDirectoryGroup");
		scScreenFunctions.SetRadioGroupValue(frmScreen,"InDirectoryGroup","1");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCompanyName");
		if (m_ctrBranchName != null) scScreenFunctions.SetFieldToReadOnly(frmScreen,m_ctrBranchName.id);
		if (m_ctrOrganisationType != null) scScreenFunctions.SetFieldToReadOnly(frmScreen,m_ctrOrganisationType.id);
		if (m_ctrSortCode != null) scScreenFunctions.SetFieldToReadOnly(frmScreen,m_ctrSortCode.id);

		frmScreen.btnPAFSearch.style.visibility = "hidden";
		frmScreen.btnClear.style.visibility = "hidden";
		frmScreen.btnAddToDirectory.style.visibility = "hidden";
		return;
	}

	<%/* 'In Directory' option buttons and address/contact fields */%>
	if (m_bDirectoryAddress)
	{
		scScreenFunctions.SetRadioGroupToWritable(frmScreen,"InDirectoryGroup");
		scScreenFunctions.SetRadioGroupValue(frmScreen,"InDirectoryGroup","1");
		scScreenFunctions.SetCollectionToReadOnly(spnContactDetails);
		scScreenFunctions.SetCollectionToReadOnly(spnAddress);
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCompanyName");
		if (m_ctrBranchName != null) scScreenFunctions.SetFieldToReadOnly(frmScreen,m_ctrBranchName.id);
		if (m_ctrOrganisationType != null) scScreenFunctions.SetFieldToReadOnly(frmScreen,m_ctrOrganisationType.id);
		if (m_ctrSortCode != null) scScreenFunctions.SetFieldToReadOnly(frmScreen,m_ctrSortCode.id);
	}
	else
	{
		scScreenFunctions.SetRadioGroupToDisabled(frmScreen,"InDirectoryGroup");
		scScreenFunctions.SetRadioGroupValue(frmScreen,"InDirectoryGroup","0");
		scScreenFunctions.SetCollectionToWritable(spnContactDetails);
		scScreenFunctions.SetCollectionToWritable(spnAddress);
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtCompanyName");
		if (m_ctrBranchName != null) scScreenFunctions.SetFieldToWritable(frmScreen,m_ctrBranchName.id);
		if (m_ctrOrganisationType != null) scScreenFunctions.SetFieldToWritable(frmScreen,m_ctrOrganisationType.id);
		if (m_ctrSortCode != null) scScreenFunctions.SetFieldToWritable(frmScreen,m_ctrSortCode.id);
	}

	<%/* Buttons */%>
	frmScreen.btnPAFSearch.disabled = m_bDirectoryAddress
	frmScreen.btnClear.disabled = m_bDirectoryAddress
	frmScreen.btnDirectorySearch.disabled = false;

	<%/* Add to Directory */%>
	if (!m_bCanAddToDirectory)
		frmScreen.btnAddToDirectory.style.visibility = "hidden";
	else
	{
		frmScreen.btnAddToDirectory.style.visibility = "visible";
		frmScreen.btnAddToDirectory.disabled = m_bDirectoryAddress;
	}

	if (m_fSetAvailableFunctionalityOverride != null) m_fSetAvailableFunctionalityOverride();
}

function ThirdPartyDetailsDisableButtons()
{
	frmScreen.btnAddToDirectory.disabled=true;
	frmScreen.btnClear.disabled=true;
	frmScreen.btnPAFSearch.disabled=true;
	frmScreen.btnDirectorySearch.disabled = true;
	//JR - Omiplus24
	frmScreen.btnContact.disabled = true; 
}

function PopulateTPTitleCombo()
{
	var blnSuccess = true;
	//GD BMIDS0077
	<%/*SG 25/06/02 SYS4930 START */%>
	if (m_bIsPopup)
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	else
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<%/*SG 25/06/02 SYS4930 END */%>
	var sTPGroupList = new Array("Title");

	if(XML.GetComboLists(document,sTPGroupList)) <% /*
	PB 06/07/2006 EP543 Begin */ %>
	{
		XML.PopulateCombo(document,frmScreen.cboTitle,"Title",true);
		checkOtherTitleField();
	} <% /*
	PB EP543 End
	*/ %>
}

function IsReadOnly()
{
	//JR - Omiplus24, Used for DC241 pop-up to determined ReadOnly action
	var strReadOnly = "0";
	<% /* BM0094 MDC 12/12/2002 */ %>
	// if (m_sReadOnly=="1" || m_bDirectoryOnly || m_bDirectoryAddress)
	if (bLegalRep && m_sReadOnly=="0")
		strReadOnly = "0";
	else if (m_sReadOnly=="1" || m_bDirectoryOnly || m_bDirectoryAddress)
		strReadOnly = "1";
	<% /* BM0094 MDC 12/12/2002 - End */ %>
		
	return strReadOnly;	
}
-->
</script>
<% /* OMIGA BUILD VERSION 045.04.03.29.03 */ %>
