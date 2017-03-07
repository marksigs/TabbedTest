<%@ LANGUAGE="JSCRIPT" %>
<html>
	<%
/*
Workfile:      RA041.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Ambiguous Address Resolution
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
INR		26/02/2004	Screen Design
INR		19/03/2004	BMIDS730 Address Targeting Processing
SAB		28/10/2005	MAR245 Disabled Accept Declared Address checkbox

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom Specific History :

Prog	Date		AQR			Description
SR		05/03/2007	EP2_1644	modified LEFT attribute for scScrollPlus
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/
%>
	<head>
		<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
			<title>Ambiguous Address Resolution</title>
	</head>
	<body>
		<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
		<OBJECT data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden"
			tabIndex="-1" type="text/x-scriptlet" width="1" VIEWASTEXT>
		</OBJECT>
		<% /* Validation script - Controls Soft Coded Field Attributes */ %>
		<script src="validation.js" language="JScript"></script>
		<% /* Scriptlets - remove any which are not required */ %>
		<!-- List Scroll object -->
		<span style="LEFT: 302px; POSITION: absolute; TOP: 315px">
			<OBJECT data="scPageScroll.htm" id="scScrollPlus" style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px"
				tabIndex="-1" type="text/x-scriptlet" VIEWASTEXT>
			</OBJECT>
		</span>
		<% //Span to keep tabbing within this screen %>
		<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
		<form id="frmScreen" validate="onchange" mark>
			<div id="divCCData" style="HEIGHT: 56px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px"
				class="msgGroup">
				<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Customer Name
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
						<input id="txtCustomerName" maxlength="70" style="POSITION: absolute; WIDTH: 250px" class="msgReadOnly"
							READONLY tabindex="-1">
					</span>
	</span>
			</div>
			<div id="divAddress" style="HEIGHT: 241px; LEFT: 10px; POSITION: absolute; TOP: 70px; WIDTH: 604px"
				class="msgGroup">
				<span id="spnCurrentAddress" style="LEFT: 4px; POSITION: absolute; TOP: 4px">
					<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
						<strong>Declared Address</strong>
					</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel">
			Post Code
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtPostCode" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 48px" class="msgLabel">
			Flat No. 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtFlatNo" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 66px" class="msgLabel">
			House <br>Name &amp; No. 
			<span style="LEFT: 80px; POSITION: absolute; TOP: 2px">
							<input id="txtHouseName" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
			<span style="LEFT: 231px; POSITION: absolute; TOP: 2px">
							<input id="txtHouseNo" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 96px" class="msgLabel">
			Street 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtStreet" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 120px" class="msgLabel">
			District 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtDistrict" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 144px" class="msgLabel">
			Town
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtTown" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 168px" class="msgLabel">
			County
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtCounty" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span id="spnAccDecAddr" style="LEFT: 4px; POSITION: absolute; TOP: 194px" class="msgLabel">
						<STRONG>Accept Declared Address</STRONG>
						<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
							<input id="chkDECLAREACCEPT" type="checkbox" name="DECLAREACCEPT" style="POSITION: absolute; WIDTH: 70px"
								onclick="EnableDisableDECLAREACCEPT()">
						</span>
					</span>
				</span>
				<span id="spnTargetAddress" style="LEFT: 310px; POSITION: absolute; TOP: 4px">
					<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
						<strong>Bureau Address</strong>
					</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel">
			Post Code
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtTargetPostCode" maxlength="8" style="POSITION: absolute; WIDTH: 70px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 48px" class="msgLabel">
			Flat No. 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtTargetFlatNo" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 66px" class="msgLabel">
			House<br>Name &amp; No. 
			<span style="LEFT: 80px; POSITION: absolute; TOP: 2px">
							<input id="txtTargetHouseName" maxlength="40" style="POSITION: absolute; WIDTH: 150px"
								class="msgReadOnly" READONLY tabindex="-1">
						</span>
			<span style="LEFT: 231px; POSITION: absolute; TOP: 2px">
							<input id="txtTargetHouseNo" maxlength="5" style="POSITION: absolute; WIDTH: 50px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 96px" class="msgLabel">
			Street 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtTargetStreet" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 120px" class="msgLabel">
			District 
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtTargetDistrict" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 144px" class="msgLabel">
			Town
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtTargetTown" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 168px" class="msgLabel">
			County
			<span style="LEFT: 80px; POSITION: absolute; TOP: -3px">
							<input id="txtTargetCounty" maxlength="40" style="POSITION: absolute; WIDTH: 150px" class="msgReadOnly"
								READONLY tabindex="-1">
						</span>
		</span>
					<span style="LEFT: 4px; POSITION: absolute; TOP: 194px" class="msgLabel">
						<STRONG>Accept Bureau Address</STRONG>
						<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
							<input id="chkACCEPT" type="checkbox" name="ACCEPT" style="POSITION: absolute; WIDTH: 70px"
								onclick="EnableDisableACCEPT()">
						</span>
					</span>
				</span>
			</div>
		</form>
		<%/* Main Buttons */ %>
		<div id="msgButtons" style="HEIGHT: 19px; LEFT: 18px; POSITION: absolute; TOP: 315px; WIDTH: 450px"><!-- #include FILE="msgButtons.asp" -->
		</div> <!-- #include FILE="attribs/ra041Attribs.asp" -->
		<script language="JScript">
<!--

var m_BaseNonPopupWindow = null;
var AddressTargetXML = null;
var DeclareAddressTargetXML = null;
var m_sCustomerName		= null;
var m_sBlockSeqNo		= null;
var m_sAddressTarget	= null;
var m_sReadOnly	= null;
//BMIDS730
var m_sDeclareAddress = null;
var m_blnEnableDeclared = null;

function window.onload()
{

	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	RetrieveData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	SetScreenOnReadOnly();

	// MAR245
	if (m_blnEnableDeclared == false)
	{
		scScreenFunctions.HideCollection(spnAccDecAddr);
	}
	// MAR245
	PopulateScreen(0);
	window.returnValue = null;

	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();

}

function btnSubmit.onclick()
{	
	var sReturn = new Array();
	
	// return the BLOCKCOUNT to indicate which AUK1 block has been accepted
	if(frmScreen.chkACCEPT.checked)
	{	
		sReturn[0] = true ;
		//Return the selected AUK1 block
		sReturn[1] = AddressTargetXML.ActiveTag.xml;

		window.returnValue	= sReturn;
		window.close();	
	}
	else
	{
		//BMIDS730 May want to accept the declared address
		if(frmScreen.chkDECLAREACCEPT.checked)
		{
			sReturn[0] = true ;
			//Return the currently selected AUK1 block so we have all the correct
			//Experian references, but update it with the declared address details
			GetDeclaredAUK1() ;
			sReturn[1] = AddressTargetXML.ActiveTag.xml;

			window.returnValue	= sReturn;
			window.close();	
		}
		else
		{
			alert('Please "Accept" an Address before selecting the "Ok" button');
		}
	}
}

function btnCancel.onclick()
{
	var sReturn = new Array();
	sReturn[0] = false ;
	
	//BMIDS730 return the value
	window.returnValue	= sReturn;
	window.close();
}

function RetrieveData()
{
	var sArguments = window.dialogArguments;
	window.dialogTop		= sArguments[0];
	window.dialogLeft		= sArguments[1];
	window.dialogWidth		= sArguments[2];
	window.dialogHeight		= sArguments[3];
	var sParameters			= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();

	m_sCustomerName		= sParameters[0];
	m_sBlockSeqNo		= sParameters[1];
	m_sAddressTarget	= sParameters[2];
	//BMIDS730
	m_sDeclareAddress	= sParameters[3];
	//MAR245
	m_blnEnableDeclared = sParameters[4];

	AddressTargetXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	DeclareAddressTargetXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	// Possibility of multiple ambiguous addresses or a single invalid address
	AddressTargetXML.LoadXML(m_sAddressTarget);
	AddressTargetXML.ActiveTag = null;
	strPattern = "ADDRESSTARGET[@BLOCKSEQNUMBER='" + m_sBlockSeqNo + "']";
	AddressTargetXML.CreateTagList(strPattern);

	// There will only be a single declared address
	DeclareAddressTargetXML.LoadXML(m_sDeclareAddress);
	DeclareAddressTargetXML.ActiveTag = null;
	strPattern = "DECLAREADDRESS[@BLOCKSEQNUMBER='" + m_sBlockSeqNo + "']";
	DeclareAddressTargetXML.SelectTag(null, strPattern);

	scScrollPlus.Initialise(GetAddresses,PopulateScreen,1,1);

}

function GetAddresses(nPageNo)
{
	AddressTargetXML.SelectTagListItem(nPageNo - 1);
	return AddressTargetXML.ActiveTagList.length;
}


function PopulateScreen(nStart)
{
	
	frmScreen.txtCustomerName.value = m_sCustomerName;
	
	frmScreen.txtTargetFlatNo.value		= AddressTargetXML.GetAttribute("FLAT");
	frmScreen.txtTargetHouseName.value	= AddressTargetXML.GetAttribute("HOUSENAME");
	frmScreen.txtTargetHouseNo.value	= AddressTargetXML.GetAttribute("HOUSENUMBER");
	frmScreen.txtTargetStreet.value		= AddressTargetXML.GetAttribute("STREET");
	frmScreen.txtTargetTown.value		= AddressTargetXML.GetAttribute("TOWN");
	frmScreen.txtTargetDistrict.value	= AddressTargetXML.GetAttribute("DISTRICT");
	frmScreen.txtTargetCounty.value		= AddressTargetXML.GetAttribute("COUNTY");
	frmScreen.txtTargetPostCode.value	= AddressTargetXML.GetAttribute("POSTCODE");

	frmScreen.txtFlatNo.value		= DeclareAddressTargetXML.GetAttribute("FLAT");
	frmScreen.txtHouseName.value	= DeclareAddressTargetXML.GetAttribute("HOUSENAME");
	frmScreen.txtHouseNo.value		= DeclareAddressTargetXML.GetAttribute("HOUSENUMBER");
	frmScreen.txtStreet.value		= DeclareAddressTargetXML.GetAttribute("STREET");
	frmScreen.txtTown.value			= DeclareAddressTargetXML.GetAttribute("TOWN");
	frmScreen.txtDistrict.value		= DeclareAddressTargetXML.GetAttribute("DISTRICT");
	frmScreen.txtCounty.value		= DeclareAddressTargetXML.GetAttribute("COUNTY");
	frmScreen.txtPostCode.value		= DeclareAddressTargetXML.GetAttribute("POSTCODE");
}

function EnableDisableACCEPT()
{
	//BMIDS730 Disable/Enable Scroll if accept is checked/unchecked
	if(frmScreen.chkACCEPT.checked)
	{	
		//Can't have both checkboxes checked
		frmScreen.chkDECLAREACCEPT.checked = false;
		scScrollPlus.disableScroll()
	}
	else
	{
		scScrollPlus.enableScroll()
	}
}

//BMIDS730 Disable/Enable Target text fields if accept is checked/unchecked
function EnableDisableDECLAREACCEPT()
{
	if(frmScreen.chkDECLAREACCEPT.checked)
	{
		//Can't have both checkboxes checked
		frmScreen.chkACCEPT.checked = false;

		scScrollPlus.disableScroll()
	}
	else
	{
		//May want to change something, so enable scroll
		scScrollPlus.enableScroll()
	}

}

function GetDeclaredAUK1()
{
	
	AddressTargetXML.SetAttribute("FLAT", frmScreen.txtFlatNo.value);
	AddressTargetXML.SetAttribute("HOUSENAME", frmScreen.txtHouseName.value);
	AddressTargetXML.SetAttribute("HOUSENUMBER", frmScreen.txtHouseNo.value);
	AddressTargetXML.SetAttribute("STREET", frmScreen.txtStreet.value);
	AddressTargetXML.SetAttribute("TOWN", frmScreen.txtTown.value);
	AddressTargetXML.SetAttribute("DISTRICT", frmScreen.txtDistrict.value);
	AddressTargetXML.SetAttribute("COUNTY", frmScreen.txtCounty.value);
	AddressTargetXML.SetAttribute("POSTCODE", frmScreen.txtPostCode.value);
	
}

function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnOK.disabled = true;
		frmScreen.btnCancel.focus();
	}
}
-->
		</script>
	</body>
</html>
<% /* OMIGA BUILD VERSION 045.04.02.12.00 */ %>
