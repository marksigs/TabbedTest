<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      DC011.asp
Copyright:     Copyright © 2006 Vertex Financial Services

Description:   Popup screen to show introducer details gained from BM's introducer system
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MO		30/09/2002	Created, INWP1, BM061, Introducer details
MO		18/10/2002  BMIDS00663, Fixed bug with MCCB number		
MC		06/05/2005	BMIDS571 White Space added to page title
KRW     24/05/2004  BMIDS762   - FSA Validation
IK		05/04/2006	EP15 - EPSOM changes
PB		21/09/2006	EP1117 - CC59 - Additional packager contact details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History :

Prog	Date		AQR			Description
INR		31/10/2006	EP2_12		 Changes for New Intermediary structure
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Introducer Details <!-- #include file="includes/TitleWhiteSpace.asp" -->       </title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<form id="frmScreen" >
<div id="divBackground" style="HEIGHT: 480px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 381px" class="msgGroup">

	<span style="LEFT: 4px; POSITION: absolute; TOP: 28px" class="msgLabel">
		Introducer Id
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtIntroducerId" style="WIDTH: 100px" class ="msgTxt" >
		</span>
	</span>
	
	<span id="lblCompany" style="LEFT: 4px; POSITION: absolute; TOP: 107px" class="msgLabel">
		Company Name
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtCompany" style="WIDTH: 250px" class ="msgTxt" >
		</span>
	</span>	
		
	<span style="LEFT: 4px; POSITION: absolute; TOP: 54px" class="msgLabel" id=SPAN1>
		FSA Ref. No.
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtFSANumberId" style="WIDTH: 100px" class ="msgTxt" >
		</span>
	</span>
	
	<span id="lblName" style="LEFT: 4px; POSITION: absolute; TOP: 136px" class="msgLabel">
		Name
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtName" style="WIDTH: 250px" class ="msgTxt" >
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 164px" class="msgLabel">
		Type Of Introducer
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtTypeOfIntroducer" style="WIDTH: 250px" class ="msgTxt" >
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 192px" class="msgLabel">
		Address
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtBuildingNo" style="WIDTH: 75px" class ="msgTxt" >
		</span>
		<span style="LEFT: 190px; POSITION: absolute; TOP: 0px">
			<input id="txtBuildingName" style="WIDTH: 160px" class ="msgTxt" >
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 220px" class="msgLabel">
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtStreet" style="WIDTH: 250px" class ="msgTxt" >
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 250px" class="msgLabel">
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtDistrict" style="WIDTH: 250px" class ="msgTxt" >
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 280px" class="msgLabel">
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtTown" style="WIDTH: 250px" class ="msgTxt" >
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 310px" class="msgLabel">
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtCounty" style="WIDTH: 250px" class ="msgTxt" >
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 340px" class="msgLabel">
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtPostcode" style="LEFT: 0px; TOP: 2px; WIDTH: 100px" class="msgTxt">
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 372px" class="msgLabel">
		Telephone No. <% /* PB 21/09/2006 EP1117 */ %>
		<span style="padding-left:1" id="lblPhone1"></span>
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtPhone1" style="WIDTH: 200px" class ="msgTxt" >
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 400px" class="msgLabel">
		Fax No. 
		<span style="padding-left:1" id="lblPhone2"></span>
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtPhone2" style="WIDTH: 200px" class ="msgTxt" >
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 427px" class="msgLabel">
		Email address
		<span style="LEFT: 100px; POSITION: absolute; TOP: 0px">
			<input id="txtEmailAddress" style="WIDTH: 200px" class ="msgTxt" >
		</span>
	</span>

</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 483px; WIDTH: 312px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sIntroducerXML = "";
var IntroducerXML = null;
var m_BaseNonPopupWindow = null;
var XMLIntroducerSearchType = null;
var XMLListingStatus = null;

function window.onload()
{
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	RetrieveData();
	IntroducerXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	IntroducerXML.LoadXML(m_sIntroducerXML);
	
	<% /* EP2_12 */ %>
	PopulateCombos();
	PopulateScreen();
	
	SetScreenOnReadOnly();
	
	window.returnValue = null;
}

function SetScreenOnReadOnly()
{
	scScreenFunctions.SetScreenToReadOnly(frmScreen);
}

function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters		= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sIntroducerXML = sParameters[0];
	
}

<% /*EP2_12 Just Need combo info */ %>
function PopulateCombos()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("IntroducerSearchType","ListingStatus");
	
	if(XML.GetComboLists(document,sGroupList))
	{
		XMLIntroducerSearchType = XML.GetComboListXML("IntroducerSearchType");
		XMLListingStatus = XML.GetComboListXML("ListingStatus");
	}

	XML = null;
}

function PopulateScreen() 
{
	var xn = IntroducerXML.XMLDocument.documentElement.selectSingleNode("INTERMEDIARY");
	if(!xn) return;
	
	var sType;
	var listingStatus;
	var introducerId;
	var nodValueID;
	var sPrincipalFirmID = xn.getAttribute("PRINCIPALFIRMID");
	var sARFirmID = xn.getAttribute("ARFIRMID");
	var sClubAssocID = xn.getAttribute("CLUBNETWORKASSOCIATIONID");
	var sIntroducerID = xn.getAttribute("INTRODUCERID");
	if(sPrincipalFirmID&&(sPrincipalFirmID.length > 0))
	{
		introducerId = sPrincipalFirmID;
	}
	if(sARFirmID&&(sARFirmID.length > 0))
	{
		introducerId = sARFirmID;
	}
	if(sClubAssocID&&(sClubAssocID.length > 0))
	{
		introducerId = sClubAssocID;
	}
	if(sIntroducerID&&(sIntroducerID.length > 0))
	{	
		introducerId = sIntroducerID;
		frmScreen.txtName.value =  xn.getAttribute("NAME");
		lblCompany.style.display = "none";
	}
	else
	{
		frmScreen.txtCompany.value = xn.getAttribute("NAME");
		lblName.style.display = "none";
	}
	
	
	frmScreen.txtIntroducerId.value = introducerId;
	
	if(xn.getAttribute("LISTSTATUSVALIDATION"))
	{
		sValidateType =  xn.getAttribute("LISTSTATUSVALIDATION");
		nodValueID=XMLListingStatus.selectSingleNode(".//LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='"+ sValidateType +"']");
		if (nodValueID)
		{
			listingStatus =  nodValueID.selectSingleNode(".//VALUENAME").text;
		}
	}

	if(xn.getAttribute("TYPEVALIDATION"))
	{
		sValidateType = xn.getAttribute("TYPEVALIDATION");
		nodValueID=XMLIntroducerSearchType.selectSingleNode(".//LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='"+ sValidateType +"']");
		if (nodValueID)
		{
			sType = nodValueID.selectSingleNode(".//VALUENAME").text;
			frmScreen.txtTypeOfIntroducer.value = sType;
		}
	}
	var sStatus =  xn.getAttribute("STATUS");
	//ToDo if(xn.getAttribute("STATUS"))frmScreen.txtStatus.value = xn.GetAttribute("STATUS");

	if(xn.getAttribute("EMAILADDRESS"))frmScreen.txtEmailAddress.value = xn.getAttribute("EMAILADDRESS");
	if(xn.getAttribute("TELEPHONE"))frmScreen.txtPhone1.value = xn.getAttribute("TELEPHONE");
	if(xn.getAttribute("FAX"))frmScreen.txtPhone2.value = xn.getAttribute("FAX");
	if(xn.getAttribute("FSAREFNUMBER"))frmScreen.txtFSANumberId.value = xn.getAttribute("FSAREFNUMBER");
	
	if(xn.getAttribute("ADDRESSLINE1")) frmScreen.txtBuildingNo.value = xn.getAttribute("ADDRESSLINE1");
	if(xn.getAttribute("ADDRESSLINE2")) frmScreen.txtBuildingName.value = xn.getAttribute("ADDRESSLINE2");
	if(xn.getAttribute("ADDRESSLINE3")) frmScreen.txtStreet.value = xn.getAttribute("ADDRESSLINE3");
	if(xn.getAttribute("ADDRESSLINE4")) frmScreen.txtDistrict.value = xn.getAttribute("ADDRESSLINE4");
	if(xn.getAttribute("ADDRESSLINE5")) frmScreen.txtTown.value = xn.getAttribute("ADDRESSLINE5");
	if(xn.getAttribute("ADDRESSLINE6")) frmScreen.txtCounty.value = xn.getAttribute("ADDRESSLINE6");
	if(xn.getAttribute("POSTCODE")) frmScreen.txtPostcode.value = xn.getAttribute("POSTCODE");


/*	switch (IntroducerXML.GetAttribute("PREFERREDCONTACTMETHOD")) {
		case "T":
			frmScreen.optPrefContactTel.checked = true;
			break;
		case "F":
			frmScreen.optPrefContactFax.checked = true;
			break;
		case "M":
			frmScreen.optPrefContactMobile.checked = true;
			break;
		case "E":
			frmScreen.optPrefContactEmail.checked = true;
			break;
	}
*/
	
}

function btnSubmit.onclick()
{
	window.close();	
}

-->
</script>
</body>
</html>
