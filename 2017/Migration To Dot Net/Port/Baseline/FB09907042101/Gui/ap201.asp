<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP201.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Directory Search screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		15/03/01	SYS2040 Created
JR		24/09/01	Omiplus24, TelephoneNumber change 
PSC		10/12/01	SYS3454 Add Closing } and correct variables
PSC		10/12/01	SYS3367 Display and save panelid on directory search
PSC		11/12/01	SYS3367 Add Search Criteria clear button
AT		29/04/02	SYS4359 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

DPF		12/09/02	APWP1/BM007 - re-arranged screen layout and added town to search criteria
SA		25/10/02	BMIDS00659 - Changed valuer details not being saved.
SA		30/10/02	BMIDS00659 - Clear button should clear ALL details.
SA		31/10/02	BMIDS00658/660 - ValuerType should be sent to ZA020 as 10th param - not 8th.
SA		07/11/02	BMIDS00659 - SaveValuerInstructions - changed to store new Valuer type combo value.
PSC		03/12/02    BM0105 - Amend search for performance
BS		17/02/2003	BM0187 - Remove valuation types with validation type 'N' from combo
MC		05/05/2004	BMIDS751	- White space added to the title
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

PB		29/06/2006	EP435	Added reminder dialog to enter valuer details before submitting..
AShaw	21Nov06		EP2_2	Add Qual and RICS No.
AShaw	07/12/2006	EP2_344 Alter Qualification code (different source table).
AShaw	08/12/2006	EP2_364 Add VALUERNAME field.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
INGDUK Specific History:
JD		21/09/2005	MAR40 - added DateOfInspection
PSC		06/12/2005	MAR816 - Enabled DateOfInspection
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
<%/* FILL IN THE TITLE OF THE SCREEN HERE */%>
<title>Valuer's Details <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<script src="validation.js" language="JSCRIPT"></script>

<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 400px; WIDTH: 590px; POSITION: ABSOLUTE" class="msgGroup">
	<!-- Search Crteria -->
	<div style="LEFT: 4px; POSITION: absolute; TOP: 5px" class="msgLabel">
		<div style= "LEFT: 20px; POSITION: absolute; TOP: 5px" class="msgLabel" >	
			<LABEL id="idValuationType"></LABEL>
			<span style="LEFT: 90; POSITION:ABSOLUTE; TOP: -3">
				<select id="cboValuationType" style="WIDTH: 130px" class="msgCombo"></select>
			</span>
		</div>
		<div style="LEFT: 20px; POSITION: absolute; TOP: 30px" class="msgLabel" >	
			Company
			<span style="LEFT: 90; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtCompany"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT:320px; POSITION: absolute; TOP: 5px" class="msgLabel" >	
			Date Of Inspection
			<span style="LEFT: 95; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtDateOfInspection"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT:320px; POSITION: absolute; TOP: 30px" class="msgLabel" >	
			Panel No.
			<span style="LEFT: 95; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtSearchPanelNo"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT: 20px; POSITION: absolute; TOP: 55px" class="msgLabel" >	
			Town
			<span style="LEFT: 90; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtSearchTown"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT: 320px; POSITION: absolute; TOP: 55px" class="msgLabel" >	
			<LABEL id="idValuerType"></LABEL>
			<span style="LEFT: 95; POSITION:ABSOLUTE; TOP: -3">
				<select id="cboValuerType" style="WIDTH: 130px" class="msgCombo"></select>
			</span>
		</div>
		<div style="LEFT: 320px; POSITION: ABSOLUTE; TOP: 85px" class="msgLabel">						
			<input id="btnClearSearch" value="Clear" type="button" style="WIDTH: 100px; HEIGHT:22" class="msgButton">
		</div>
		
		<div style="LEFT:420px; POSITION: ABSOLUTE; TOP: 85px" class="msgLabel">						
			<input id="btnSearch" value="Directory Search" type="button" style="WIDTH: 100px; HEIGHT:22" class="msgButton">
		</div>
	</div>
	<!--Individual  valuer details -->
	<div style="LEFT: 4px; POSITION: absolute; TOP: 120px" class="msgLabel">
		<div style="LEFT: 5px; POSITION: absolute; TOP: 5px" class="msgLabel" >	
			<LABEL id="idForename"></LABEL>
			<span style="LEFT:120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtForename"  style="WIDTH: 140px ">
			</span>
		</div>
		<div style="LEFT: 5px; POSITION: absolute; TOP: 30px" class="msgLabel" >	
			<LABEL id="idSurname"></LABEL>
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtSurname"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT: 5px; POSITION: absolute; TOP: 55px" class="msgLabel" >	
			E-Mail
			<span style="LEFT: 40; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtEmail"  style="WIDTH: 220px"  class="msgTxt" NAME="txtEmail">
			</span>
		</div>
		<!-- EP2_364 - Add new field -->
		<div style="LEFT: 5px; POSITION: absolute; TOP: 80px" class="msgLabel" >	
			Vex Valuer Name
			<span style="LEFT: 100; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtVexValuerName"  style="WIDTH: 160px"  class="msgTxt" NAME="txtVexValuerName">
			</span>
		</div>
		<!-- EP2_344 (EP2_2) - Add new field -->
		<div style="LEFT: 5px; POSITION: absolute; TOP: 105px" class="msgLabel" >	
			Qualifications
			<span style="LEFT: 160; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtQualAndRICSNo"  style="WIDTH: 100px"  class="msgTxt" NAME="txtQualAndRICSNo">
			</span>
		</div>
		<!-- Company Details -->
		<div style="LEFT: 320px; POSITION: absolute; TOP: 5px" class="msgLabel" >	
			Company
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtCompanyName"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
	
		<div style="LEFT: 320px; POSITION: absolute; TOP: 30px" class="msgLabel" >	
			Panel No.
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtPanelNo"  style="WIDTH: 70px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT: 320px; POSITION: absolute; TOP: 55px" class="msgLabel" >	
			<LABEL id="idPostcode"></LABEL>
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtPostcode"  style="WIDTH: 70px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT: 320px; POSITION: absolute; TOP: 80px" class="msgLabel" >	
			<LABEL  id="idBuildingNameNo"></LABEL>
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtBuildingName"  style="WIDTH: 100px"  class="msgTxt">
			</span>
			<span style="LEFT: 235; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtBuildingNo"  style="WIDTH: 25px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT: 320px; POSITION: absolute; TOP: 105px" class="msgLabel" >	
			Street
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtStreet"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT: 320px; POSITION: absolute; TOP: 130px" class="msgLabel" >	
			<LABEL id="idDistrict"></LABEL>
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtDistrict"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT: 320px; POSITION: absolute; TOP: 155px" class="msgLabel" >	
			<LABEL id="idTown"></LABEL>
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtTown"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT: 320px; POSITION: absolute; TOP: 180px" class="msgLabel" >	
			<LABEL id="idCounty"></LABEL>
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtCounty"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT: 320px; POSITION: absolute; TOP: 205px" class="msgLabel" >	
			Country
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<select id="cboCountry" style="WIDTH: 130px" class="msgCombo"></select>
			</span>
		</div>
		<div style="LEFT: 320px; POSITION: absolute; TOP: 230px" class="msgLabel" >	
			Telephone No.
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtTelephone"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
		<div style="LEFT: 320px; POSITION: absolute; TOP: 255px" class="msgLabel" >	
			Fax No.
			<span style="LEFT: 120; POSITION:ABSOLUTE; TOP: -3">
				<input id="txtFax"  style="WIDTH: 140px"  class="msgTxt">
			</span>
		</div>
	</div>
</div>				
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 420px; WIDTH: 590px">
	<!-- #include FILE="msgButtons.asp" -->
</div>
<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP201attribs.asp" -->
<!-- #include FILE="Customise/AP201Customise.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sReadOnly = "";
var m_ValuationXML; 
var m_sAppNo = "";
var m_sAppFactFindNo = "";
var m_sInsSeqNo = "";
var m_sContextParams = "";
var m_sDirectoryGUID = "";
var cstrValuer = 11;
//JR - Omiplus24
var m_sWorkComboId = "";
var m_sFaxComboId = "";
var m_BaseNonPopupWindow = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	RetrieveContextData();

<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Cancel");
	m_ValuationXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	ShowMainButtons(sButtonList);
	Customise();

<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
	SetScreenOnReadOnly();
	SetReadOnlyFields();
	PopulateScreen();
	
 	//scScreenFunctions.SetFocusToFirstField(frmScreen);
 	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function SetReadOnlyFields()
{

	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtPostcode");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtBuildingName");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtBuildingNo");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtStreet");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDistrict");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTown");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtCounty");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "cboCountry");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtForename");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtSurname");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtTelephone");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtFax");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtEmail");
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtPanelNo");	
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtCompanyName");
	
	<% /* PSC 06/12/2005 MAR816
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtDateOfInspection"); //JD MAR40
	 */%>
}
function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnSearch.disabled =true;
	}
}

function PopulateDirectoryAddress(xmlDirectory)
{
	if(m_sDirectoryGUID.length > 0)
	{
		frmScreen.cboCountry.value = xmlDirectory.GetAttribute("COUNTRY");
		frmScreen.txtForename.value = xmlDirectory.GetAttribute("CONTACTFORENAME");
		frmScreen.txtSurname.value = xmlDirectory.GetAttribute("CONTACTSURNAME");
		frmScreen.txtPostcode.value = xmlDirectory.GetAttribute("POSTCODE");
		frmScreen.txtBuildingNo.value = xmlDirectory.GetAttribute("BUILDINGORHOUSENUMBER");
		frmScreen.txtBuildingName.value = xmlDirectory.GetAttribute("BUILDINGORHOUSENAME");
		frmScreen.txtStreet.value = xmlDirectory.GetAttribute("STREET");
		frmScreen.txtDistrict.value = xmlDirectory.GetAttribute("DISTRICT");
		frmScreen.txtTown.value = xmlDirectory.GetAttribute("TOWN");		
		frmScreen.txtCounty.value = xmlDirectory.GetAttribute("COUNTY");		
		//frmScreen.txtTelephone.value = xmlDirectory.GetAttribute("TELEPHONENUMBER");		
		//frmScreen.txtFax.value = xmlDirectory.GetAttribute("FAXNUMBER");		
		frmScreen.txtEmail.value = xmlDirectory.GetAttribute("EMAILADDRESS");		
		frmScreen.txtCompany.value = xmlDirectory.GetAttribute("COMPANYNAME");
		
		// PSC 10/12/01 SYS3367 - Start
		frmScreen.txtPanelNo.value = xmlDirectory.GetAttribute("VALUERPANELNO");
		frmScreen.txtCompanyName.value = xmlDirectory.GetAttribute("COMPANYNAME");
		// PSC 10/12/01 SYS3367 - End
		//BMIDS00658/660 Populate search town.
		frmScreen.txtSearchTown.value = xmlDirectory.GetAttribute("TOWN");		
				
		//JR - Omiplus24
		m_ValuationXML.ActiveTag = null;
		m_ValuationXML.SelectTag(null, "VALUERINSTRUCTIONS[@USAGE='" + m_sWorkComboId + "']");
		if (m_ValuationXML.ActiveTag !=null)
		{
			var sPhoneNo = m_ValuationXML.GetAttribute("COUNTRYCODE");
			sPhoneNo = sPhoneNo + " " + m_ValuationXML.GetAttribute("AREACODE");
			sPhoneNo = sPhoneNo + " " + m_ValuationXML.GetAttribute("TELENUMBER");
			sPhoneNo = sPhoneNo + " " + m_ValuationXML.GetAttribute("EXTENSIONNUMBER");
			frmScreen.txtTelephone.value = sPhoneNo;			
					
			m_ValuationXML.ActiveTag = null;
		}
		m_ValuationXML.SelectTag(null, "VALUERINSTRUCTIONS[@USAGE='" + m_sFaxComboId + "']");
		if (m_ValuationXML.ActiveTag !=null)
		{
			var sFaxNo = m_ValuationXML.GetAttribute("COUNTRYCODE");
			sFaxNo = sFaxNo + " " + m_ValuationXML.GetAttribute("AREACODE");
			sFaxNo = sFaxNo + " " + m_ValuationXML.GetAttribute("TELENUMBER");
			frmScreen.txtFax.value = sFaxNo;
			
			m_ValuationXML.ActiveTag = null;
		}
		//End
	}
}
function PopulateDirectoryAddressElem(xmlDirectory)
{
	<% /* PSC 03/12/2002 BM0105 - Start */ %>
	xmlDirectory.SelectTag(null,"NAMEANDADDRESSDIRECTORY");
	
	if(xmlDirectory.ActiveTag != null)
	{
	
		m_sDirectoryGUID = xmlDirectory.GetTagText("DIRECTORYGUID");
		frmScreen.txtCompany.value = xmlDirectory.GetTagText("COMPANYNAME");
		frmScreen.txtCompanyName.value = xmlDirectory.GetTagText("COMPANYNAME");
		
		xmlDirectory.SelectTag(null,"ADDRESS");
		
		if (xmlDirectory.ActiveTag != null)
		{
			frmScreen.cboCountry.value = xmlDirectory.GetTagText("COUNTRY");
			frmScreen.txtPostcode.value = xmlDirectory.GetTagText("POSTCODE");
			frmScreen.txtBuildingNo.value = xmlDirectory.GetTagText("BUILDINGORHOUSENUMBER");
			frmScreen.txtBuildingName.value = xmlDirectory.GetTagText("BUILDINGORHOUSENAME");
			frmScreen.txtStreet.value = xmlDirectory.GetTagText("STREET");
			frmScreen.txtDistrict.value = xmlDirectory.GetTagText("DISTRICT");
			frmScreen.txtTown.value = xmlDirectory.GetTagText("TOWN");
			frmScreen.txtSearchTown.value		= xmlDirectory.GetTagText("TOWN");
			frmScreen.txtCounty.value = xmlDirectory.GetTagText("COUNTY");
		}
		
		xmlDirectory.SelectTag(null,"CONTACTDETAILS");
		
		if (xmlDirectory.ActiveTag != null)
		{
			frmScreen.txtForename.value = xmlDirectory.GetTagText("CONTACTFORENAME");
			frmScreen.txtSurname.value = xmlDirectory.GetTagText("CONTACTSURNAME");
			frmScreen.txtEmail.value = xmlDirectory.GetTagText("EMAILADDRESS");		
	}
		
		xmlDirectory.SelectTag(null,"PANEL");
		
		if (xmlDirectory.ActiveTag != null)
		{
			frmScreen.txtSearchPanelNo.value = xmlDirectory.GetTagText("PANELID");
			frmScreen.txtPanelNo.value = xmlDirectory.GetTagText("PANELID");
		}
		<% /* PSC 03/12/2002 BM0105 - End */ %>
		
		//JR - Omiplus24, Get Work(W) ContactTelephoneDetails	
		var TempXML = xmlDirectory.ActiveTag;
		var sTelephone = "";
		var sFaxNumber = "";
		
		xmlDirectory.SelectTag(null,"CONTACTTELEPHONEDETAILS[USAGE='" + m_sWorkComboId + "']");
		if (xmlDirectory.ActiveTag != null)
		{
			sTelephone = xmlDirectory.GetTagText("COUNTRYCODE");
			sTelephone = sTelephone + " " + xmlDirectory.GetTagText("AREACODE");
			sTelephone = sTelephone + " " + xmlDirectory.GetTagText("TELENUMBER");
			sTelephone = sTelephone + " " + xmlDirectory.GetTagText("EXTENSIONNUMBER");
			xmlDirectory.ActiveTag = null; // PSC 10/12/01 SYS3454
		}
		//JR - Omiplus24, Get Fax(F) ContactTelephoneDetails
		// PSC 10/12/01 SYS3454 - Start	
		xmlDirectory.SelectTag(null,"CONTACTTELEPHONEDETAILS[USAGE='" + m_sFaxComboId + "']");
		if (xmlDirectory.ActiveTag != null)
		{	
			sFaxNumber = xmlDirectory.GetTagText("COUNTRYCODE");
			sFaxNumber = sFaxNumber + " " + xmlDirectory.GetTagText("AREACODE");
			sFaxNumber = sFaxNumber + " " + xmlDirectory.GetTagText("TELENUMBER");
			xmlDirectory.ActiveTag = null;
		
		}
		<% /* PSC 03/12/2002 BM0105 - Start */ %>
		frmScreen.txtTelephone.value = sTelephone;
		frmScreen.txtFax.value = sFaxNumber;
		<% /* PSC 03/12/2002 BM0105 - End */ %>

		// PSC 10/12/01 SYS3454 - End
		xmlDirectory.ActiveTag = TempXML;
	} // PSC 10/12/01 - SYS
}

function PopulateScreen()
{
	var bSuccess;

	PopulateCombos();
	//JR - Omiplus24
	m_sFaxComboId = m_ValuationXML.GetComboIdForValidation("ContactTelephoneUsage", "F", null, document);
	m_sWorkComboId = m_ValuationXML.GetComboIdForValidation("ContactTelephoneUsage", "W", null, document);	
	//End
	bSuccess = GetValuerInstructions();

	if(bSuccess)
	{
		// Populate the screen fields
		frmScreen.cboValuerType.value = m_ValuationXML.GetAttribute("VALUERTYPE");		
		frmScreen.cboValuationType.value = m_ValuationXML.GetAttribute("VALUATIONTYPE");		
		frmScreen.txtSearchPanelNo.value = m_ValuationXML.GetAttribute("VALUERPANELNO");
		frmScreen.txtDateOfInspection.value = m_ValuationXML.GetAttribute("APPOINTMENTDATE");	//JD MAR40
		//EP2_344 (EP2_2)
		frmScreen.txtQualAndRICSNo.value = m_ValuationXML.GetAttribute("QUALIFICATION"); 
		//EP2_364 
		frmScreen.txtVexValuerName.value = m_ValuationXML.GetAttribute("VALUERNAME"); 

		PopulateDirectoryAddress(m_ValuationXML);
	}	

	return(bSuccess);
}

function GetValuerInstructions()
{
	var bSuccess = false;
	var reqTag = m_ValuationXML.CreateRequestTagFromArray(m_sContextParams, "GetValuerInstructions");
	m_ValuationXML.CreateActiveTag("VALUERINSTRUCTIONS");
	m_ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	m_ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
	m_ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
	m_ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO",m_sInsSeqNo);

	m_ValuationXML.RunASP(document, "omAppProc.asp");

	if(m_ValuationXML.IsResponseOK())
	{
		bSuccess = true;
		m_ValuationXML.SelectTag(null, "VALUERINSTRUCTIONS");
		
		m_sDirectoryGUID= m_ValuationXML.GetAttribute("DIRECTORYGUID");
		
		if(m_sDirectoryGUID.length == 0 && m_sReadOnly != "1")
		{
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtCompany");
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtSearchPanelNo");
			scScreenFunctions.SetFieldToWritable(frmScreen,"cboValuerType");
			scScreenFunctions.SetFieldToWritable(frmScreen,"cboValuationType");
			//BMIDS00658 - correct spelling mistake found!
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtSearchTown");
			//EP2_344 (EP2_2)
			scScreenFunctions.SetFieldToWritable (frmScreen,"txtQualAndRICSNo"); 
			//EP2_364
			scScreenFunctions.SetFieldToWritable (frmScreen,"txtVexValuerName"); 

			frmScreen.btnSearch.disabled = false;
		}
	}
	return(bSuccess);
}

function PopulateCombos()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("Country","ValuerType","ValuationType");

	if(XML.GetComboLists(document,sGroupList))
	{
		XML.PopulateCombo(document,frmScreen.cboCountry ,"Country",true);
		XML.PopulateCombo(document,frmScreen.cboValuerType ,"ValuerType",true);
		XML.PopulateCombo(document,frmScreen.cboValuationType,"ValuationType",true);
		
		<% /* BS BM0187 17/02/03 
		Remove ValuationTypes with a ValidationType = 'N' from the combo*/ %>
		var iCount = 0;
		for (iCount = frmScreen.cboValuationType.length - 1; iCount > 0; iCount--)
		{
			if (scScreenFunctions.IsOptionValidationType(frmScreen.cboValuationType, iCount, "N") )
				frmScreen.cboValuationType.remove(iCount);
		}		
		<% /* BS BM0187 End 17/02/03 */ %>
	}

	XML = null;
}

function RetrieveContextData()
{

	var sParameters;
	var sArguments = window.dialogArguments;
		
	//Grab the window sizing properties...
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	sParameters	= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
		
	//Grab the array with my parameters from calling screen, plus the readonly flag
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sContextParams	= sParameters[0];
	m_sAppNo			= sParameters[1];
	m_sAppFactFindNo	= sParameters[2];	
	m_sInsSeqNo			= sParameters[3];	
	m_sTaskXML			= sParameters[4];	
	m_sReadOnly			= sParameters[5];	

	
}

function CommitChanges()
{
	var bSuccess = true;
	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				if (ValidateScreen())
					bSuccess = SaveValuerInstructions();
				else
					bSuccess = false;
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function ValidateScreen()
{
	var bSuccess = true;
	return(bSuccess);
}

function ValidateInvoiceAmount()
{
	var bSuccess = true;
	return( bSuccess );
}

function SaveScreenValues( XML )
{
	var bSuccess;
	var sVal;
	
	bSuccess = false
	if(XML != null)
	{
		if(m_sDirectoryGUID.length > 0)
		{
			XML.SetAttribute("DIRECTORYGUID", m_sDirectoryGUID);
		}
		bSuccess = true;
	}

	return(bSuccess);
}

function SaveValuerInstructions()
{
	var bSuccess = false;
	var ValuationXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTaskStatus;
	var sASPFile;
			
	if(m_sDirectoryGUID.length > 0)
	{
		var reqTag = ValuationXML.CreateRequestTagFromArray(m_sContextParams, "UpdateValuerInstructions");

		ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");

		sASPFile = "omAppProc.asp"
		ValuationXML.CreateActiveTag("VALUERINSTRUCTION");
		ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
		ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
		ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInsSeqNo);
		ValuationXML.SetAttribute("DIRECTORYGUID", m_sDirectoryGUID);
		
		// PSC 10/12/01 SYS3367
		ValuationXML.SetAttribute("VALUERPANELNO", frmScreen.txtPanelNo.value);
		
		//SA 08/11/02 BMIDS00659
		ValuationXML.SetAttribute("VALUERTYPE", frmScreen.cboValuerType.value);
		ValuationXML.SetAttribute("VALUATIONTYPE", frmScreen.cboValuationType.value);
		ValuationXML.SetAttribute("APPOINTMENTDATE", frmScreen.txtDateOfInspection.value); // JD MAR40
		bSuccess = true;

		//EP2_344 (EP2_2)
		ValuationXML.SetAttribute("QUALIFICATION", frmScreen.txtQualAndRICSNo.value);
		//EP2_364 
		ValuationXML.SetAttribute("VALUERNAME", frmScreen.txtVexValuerName.value);

	}
	<% /* PB 29/06/2006 EP435 Begin */ %>
	else
	{
		alert('Enter search criteria and click "Directory search" to find a valuer\ror choose "Cancel" to exit');
	}
	<% /* PB EP435 End */ %>

	if(bSuccess)
	{
		// 		ValuationXML.RunASP(document, sASPFile);
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					ValuationXML.RunASP(document, sASPFile);
				break;
			default: // Error
				ValuationXML.SetErrorResponse();
			}

		if(ValuationXML.IsResponseOK())
		{
			// Update the context
			bSuccess = true;
		}
		else
		{
			bSuccess = false
		}
	}	

	return(bSuccess);
}

function DoSearch()
{
	var bSuccess = false;
	var sPanelNo;
	var sCompany;
	var sValuerType;
	var sValuationType;
	var sSearchTown;
	sValuationType	= frmScreen.cboValuationType.value;

	<% /* PSC 03/12/2002 BM0105 - Start */ %>
	if (sValuationType.length <= 0)
	{
		alert("Valuation Type is required for a directory search.");
		return;
	}
	
	if (frmScreen.txtSearchPanelNo.value != "")
		if (!CheckField(frmScreen.txtSearchPanelNo, "Search Panel No")) return;

	if (frmScreen.txtCompany.value != "")
	{
		if (!CheckField(frmScreen.txtCompany, "Company Name")) return;
	}
	else
	{
		frmScreen.txtCompany.focus; 
		alert("Company Name must be entered");
		return;
	}

	if (frmScreen.txtSearchTown.value != "")
		if (!CheckField(frmScreen.txtSearchTown, "Search Town")) return;

	sPanelNo		= frmScreen.txtSearchPanelNo.value; 
	sCompany		= frmScreen.txtCompany.value; 
	sValuerType		= frmScreen.cboValuerType.value; 
	sSearchTown		= frmScreen.txtSearchTown.value

	//PASS RESPONSE TO za020
	var paramsArray = new Array();
			
	/* Route to ZA020 */
	paramsArray[0] = m_sContextParams[0];
	paramsArray[1] = m_sContextParams[1];
	paramsArray[2] = m_sContextParams[2];			
	paramsArray[3] = m_sContextParams[3];
	paramsArray[4] = cstrValuer;
	paramsArray[5] = sCompany;
	paramsArray[6] = true;
	paramsArray[7] = sValuationType;
	paramsArray[8] = sSearchTown;
	paramsArray[9] = sPanelNo;
	//BMIDS00658 Pass in Valuer Type
	paramsArray[10] = sValuerType;
				
	var sReturn = scScreenFunctions.DisplayPopup(window, document, "ZA020.asp", paramsArray, 630, 440);
	if (sReturn != null)
	{
		var XMLPanelValuerList 	= new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
		XMLPanelValuerList.LoadXML(sReturn[1]);
		PopulateDirectoryAddressElem(XMLPanelValuerList);
		bSuccess = true;
	}
	<% /* PSC 03/12/2002 BM0105 - End */ %>
	
	return(bSuccess);
}

function frmScreen.btnSearch.onclick()
{
	DoSearch();
	//BMIDS00659 SA The valuer has changed - so need to flag it.
	m_bIsChanged = true;
	//m_sValuationType
}

function btnCancel.onclick()
{
	//nothing is returned from here....
	window.close();
}

function btnSubmit.onclick()
{
	if(CommitChanges())
	{
		//nothing is returned from here....
		window.close();
	}
}

function frmScreen.btnClearSearch.onclick ()
{
	frmScreen.txtSearchPanelNo.value = "";
	frmScreen.txtCompany.value = "";
	frmScreen.txtSearchTown.value = "";	//BMIDS00658 clear town as well.
	//BMIDS00659 SA Clear down all fields
	frmScreen.cboCountry.value = "";
	frmScreen.txtForename.value = "";
	frmScreen.txtSurname.value = "";
	frmScreen.txtPostcode.value = "";
	frmScreen.txtBuildingNo.value = "";
	frmScreen.txtBuildingName.value = "";
	frmScreen.txtStreet.value = "";
	frmScreen.txtDistrict.value = "";
	frmScreen.txtTown.value = "";		
	frmScreen.txtCounty.value = "";		
	frmScreen.txtEmail.value = "";		
	frmScreen.txtPanelNo.value = "";
	frmScreen.txtCompanyName.value = "";
	frmScreen.txtTelephone.value = "";			
	frmScreen.txtFax.value = "";
	frmScreen.txtDateOfInspection.value = "";  //JD MAR40
	frmScreen.txtQualAndRICSNo.value = "";  //EP2_344 (EP2_2)
	frmScreen.txtVexValuerName.value = "";  //EP2_344 (EP2_2)
	//BMIDS00659 end
		
}
<% /* PSC 03/12/2002 BM0105 - Start */ %>
function CheckField(refField,sName)
{
	var sField=refField.value ;
	if(sField.indexOf("*") == 0)
	{
		alert("You must enter at least 1 preceding character to wildcard the " + sName);
		refField.focus();
		return false;
	}
	else 
		return true;

}
<% /* PSC 03/12/2002 BM0105 - End */ %>
-->
</script>

</body>
</html>


