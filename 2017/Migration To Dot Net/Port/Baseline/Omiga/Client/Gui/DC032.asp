<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%/*
Workfile:      DC032.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Contact details  POP-UP
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		25/11/1999	Created
JLD		14/12/1999	DC/033 - made customer name bigger
					DC/046 - changed the look of it
					DC/047 - Expanded field lengths
					SYS0067 - all of the above.
AD		30/01/2000	Rework
IW		08/03/00	sys0096 - Enabled OK button when screen is readonly for routing purposes
AY		30/03/00	scScreenFunctions change
IW		19/04/00	Added Work Extension Number Field
IW		03/05/00	Added preffered Contact Radio Buttons
IW		09/05/00	SYS0085 Changes/consolidation in Contact Validation
IW		10/05/00	SYS0730 Changed order of validation routines.
BG		17/05/00	SYS0752 Removed Tooltips
IW		25/05/00	SYS0776 Added EmailPreffered option
MC		14/06/00	SYS0935 Amended to use PREFERREDMETHODOFCONTACT
BG		25/07/00	SYS0971 - Made customer name field longer to handle max length of customer name
JR		13/09/01	Omiplus24 - Included Country/Area Code for telephone number change
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
DB		25/02/2003	BM0042 Make correspondence salutation read-only.
MC		19/04/04	CC057	white spaces padded to the right of title text to hide standard IE message.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History

GHun	31/01/2006	MAR1158	Removed redundant OtherSystemReadOnlyCustomer logic
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<head>
	<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Contact Details  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript" type="text/javascript"></script>

<%/* FORMS */%>
<form id="frmScreen" mark validate="onchange">
	<div id="divTelephone" style="TOP: 10px; LEFT: 10px; HEIGHT: 190px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
		<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			Customer Name
			<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
				<input id="txtCustomerName" name="CustomerName" maxlength="70" style="WIDTH: 430px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 36px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			<strong>Telephone Details</strong>
		</span>

		<span style="TOP: 62px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			Type
		</span>			
		<span style="TOP: 53px; LEFT: 96px; POSITION: ABSOLUTE" class="msgLabel">
			Country<BR>Code
		</span>	
		<span style="TOP: 62px; LEFT: 142px; POSITION: ABSOLUTE" class="msgLabel">
			Area Code
		</span>
		<span style="TOP: 62px; LEFT: 202px; POSITION: ABSOLUTE" class="msgLabel">
			Telephone Number
		</span>	
		<span style="TOP: 53px; LEFT: 320px; POSITION: ABSOLUTE" class="msgLabel">
			Work Extension<BR> Number
		</span>	
		<span style="TOP: 62px; LEFT: 404px; POSITION: ABSOLUTE" class="msgLabel">
			Contact Time
		</span>				

		<span style="TOP: 53px; LEFT: 545px; POSITION: ABSOLUTE" class="msgLabel">
			Preferred<BR>Contact?
		</span>				

		<span id="spnTelephone1" style="TOP: 88px; LEFT: 4px; POSITION: ABSOLUTE">
			<span style="TOP: -3px; LEFT: 0px; POSITION: ABSOLUTE">
				<select id="cboType1" name="Type1" style="WIDTH: 80px" class="msgCombo" onchange="return cboType1_onchange()">
				</select>
			</span>
			<span style="TOP: -3px; LEFT: 90px; POSITION: ABSOLUTE">
				<input id="txtCountryCode1" name="CountryCode1" maxlength="3" style="WIDTH: 40px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 135px; POSITION: ABSOLUTE">
				<input id="txtAreaCode1" name="AreaCode1" maxlength="6" style="WIDTH: 55px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 195px; POSITION: ABSOLUTE">
				<input id="txtTelNumber1" name="TelNumber1" maxlength="30" style="WIDTH: 115px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 315px; POSITION: ABSOLUTE">
				<input id="txtExtensionNumber1" name="ExtNumber1" maxlength="8" style="WIDTH: 80px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 400px; POSITION: ABSOLUTE">
				<input id="txtTime1" name="Time1" maxlength="30" style="WIDTH: 140px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 558px; POSITION: ABSOLUTE">
				<input id="optPREFERREDCONTACT1" name="RadioGroup" type="radio" value="1">
			</span>
		</span>

		<span id="spnTelephone2" style="TOP: 114px; LEFT: 4px; POSITION: ABSOLUTE">
			<span style="TOP: -3px; LEFT: 0px; POSITION: ABSOLUTE">
				<select id="cboType2" name="Type2" style="WIDTH: 80px" class="msgCombo">
				</select>
			</span>
			<span style="TOP: -3px; LEFT: 90px; POSITION: ABSOLUTE">
				<input id="txtCountryCode2" name="CountryCode2" maxlength="3" style="WIDTH: 40px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 135px; POSITION: ABSOLUTE">
				<input id="txtAreaCode2" name="AreaCode2" maxlength="6" style="WIDTH: 55px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 195px; POSITION: ABSOLUTE">
				<input id="txtTelNumber2" name="TelNumber2" maxlength="30" style="WIDTH: 115px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 315px; POSITION: ABSOLUTE">
				<input id="txtExtensionNumber2" name="ExtNumber2" maxlength="8" style="WIDTH: 80px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 400px; POSITION: ABSOLUTE">
				<input id="txtTime2" name="Time2" maxlength="30" style="WIDTH: 140px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 558px; POSITION: ABSOLUTE">
				<input id="optPREFERREDCONTACT2" name="RadioGroup" type="radio" value="1">
			</span>
		</span>

		<span id="spnTelephone3" style="TOP: 140px; LEFT: 4px; POSITION: ABSOLUTE">
			<span style="TOP: -3px; LEFT: 0px; POSITION: ABSOLUTE">
				<select id="cboType3" name="Type3" style="WIDTH: 80px" class="msgCombo">
				</select>
			</span>
			<span style="TOP: -3px; LEFT: 90px; POSITION: ABSOLUTE">
				<input id="txtCountryCode3" name="CountryCode3" maxlength="3" style="WIDTH: 40px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 135px; POSITION: ABSOLUTE">
				<input id="txtAreaCode3" name="AreaCode3" maxlength="6" style="WIDTH: 55px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 195px; POSITION: ABSOLUTE">
				<input id="txtTelNumber3" name="TelNumber3" maxlength="30" style="WIDTH: 115px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 315px; POSITION: ABSOLUTE">
				<input id="txtExtensionNumber3" name="ExtNumber3" maxlength="8" style="WIDTH: 80px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 400px; POSITION: ABSOLUTE">
				<input id="txtTime3" name="Time3" maxlength="30" style="WIDTH: 140px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 558px; POSITION: ABSOLUTE">
				<input id="optPREFERREDCONTACT3" name="RadioGroup" type="radio" value="1">
			</span>
		</span>

		<span id="spnTelephone4" style="TOP: 166px; LEFT: 4px; POSITION: ABSOLUTE">
			<span style="TOP: -3px; LEFT: 0px; POSITION: ABSOLUTE">
				<select id="cboType4" name="Type4" style="WIDTH: 80px" class="msgCombo">
				</select>
			</span>
			<span style="TOP: -3px; LEFT: 90px; POSITION: ABSOLUTE">
				<input id="txtCountryCode4" name="CountryCode4" maxlength="3" style="WIDTH: 40px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 135px; POSITION: ABSOLUTE">
				<input id="txtAreaCode4" name="AreaCode4" maxlength="6" style="WIDTH: 55px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 195px; POSITION: ABSOLUTE">
				<input id="txtTelNumber4" name="TelNumber4" maxlength="30" style="WIDTH: 115px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 315px; POSITION: ABSOLUTE">
				<input id="txtExtensionNumber4" name="ExtNumber4" maxlength="8" style="WIDTH: 80px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 400px; POSITION: ABSOLUTE">
				<input id="txtTime4" name="Time4" maxlength="30" style="WIDTH: 140px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 558px; POSITION: ABSOLUTE">
				<input id="optPREFERREDCONTACT4" name="RadioGroup" type="radio" value="1">
			</span>
		</span>
	</div>

	<div id="divOtherContact" style="TOP: 210px; LEFT: 10px; HEIGHT: 120px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">										
		<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			E-Mail Address
			<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
				<input id="txtEMailAddress" name="EMailAddress" maxlength="100" style="WIDTH: 430px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
			<span style="TOP: -3px; LEFT: 558px; POSITION: ABSOLUTE">
				<input id="optPREFERREDCONTACT5" name="RadioGroup" type="radio" value="1">
			</span>
		</span>

		<span style="TOP: 30px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			Customer<BR>Correspondence<BR>Salutation
			<span style="TOP: 10px; LEFT: 110px; POSITION: ABSOLUTE">
				<input id="txtSalutation" name="Salutation" maxlength="50" style="WIDTH: 480px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 85px; LEFT: 10px; POSITION: ABSOLUTE">
			<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
			<span style="TOP: 0px; LEFT: 64px; POSITION: ABSOLUTE">
				<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton">
			</span>
		</span>		
	</div>
</form>

<!--  #include FILE="attribs/DC032attribs.asp" -->

<%/* CODE */%>
<script language="JScript" type="text/javascript">
<!--
var m_sXMLIn = "";
var m_sReadOnly = "";
var CustomerXML = null;
var m_sCustomerName = "";
var scScreenFunctions;
var m_BaseNonPopupWindow = null;

function window.onload()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions.ShowCollection(frmScreen);

	SetMasks();

	m_sReadOnly = sArgArray[4];
	m_sXMLIn = sArgArray[5]; //CustomerXML
	m_sCustomerName = sArgArray[6];

	Initialise();
	Validation_Init();

	// Read Only processing
	if (m_sReadOnly == "1")
	window.returnValue = null;
}

function PopulateCombos()
{
	var XMLTelephoneUsage	= null;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("TelephoneUsage");

	if(XML.GetComboLists(document,sGroupList))
	{
		XMLTelephoneUsage = XML.GetComboListXML("TelephoneUsage");

		var blnSuccess = true;
		blnSuccess = XML.PopulateComboFromXML(document,frmScreen.cboType1,XMLTelephoneUsage,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboType2,XMLTelephoneUsage,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboType3,XMLTelephoneUsage,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboType4,XMLTelephoneUsage,true);

		if(!blnSuccess)
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
	XML = null;
}

function PopulateTelephone(nSpanNum)
{
	switch (nSpanNum)
	{
		case 1: frmScreen.cboType1.value = CustomerXML.GetTagText("USAGE");
				frmScreen.txtTelNumber1.value = CustomerXML.GetTagText("TELEPHONENUMBER");
				frmScreen.txtExtensionNumber1.value = CustomerXML.GetTagText("EXTENSIONNUMBER");
				frmScreen.txtTime1.value = CustomerXML.GetTagText("CONTACTTIME");
				frmScreen.optPREFERREDCONTACT1.checked = (CustomerXML.GetTagText("PREFERREDMETHODOFCONTACT") == "1")? true:false;
				//JR - Omiplus24
				frmScreen.txtCountryCode1.value = CustomerXML.GetTagText("COUNTRYCODE");
				frmScreen.txtAreaCode1.value = CustomerXML.GetTagText("AREACODE");
				break;
		case 2: frmScreen.cboType2.value = CustomerXML.GetTagText("USAGE");
				frmScreen.txtTelNumber2.value = CustomerXML.GetTagText("TELEPHONENUMBER");
				frmScreen.txtExtensionNumber2.value = CustomerXML.GetTagText("EXTENSIONNUMBER");
				frmScreen.txtTime2.value = CustomerXML.GetTagText("CONTACTTIME");
				frmScreen.optPREFERREDCONTACT2.checked = (CustomerXML.GetTagText("PREFERREDMETHODOFCONTACT") == "1")? true:false;
				//JR - Omiplus24
				frmScreen.txtCountryCode2.value = CustomerXML.GetTagText("COUNTRYCODE");
				frmScreen.txtAreaCode2.value = CustomerXML.GetTagText("AREACODE");
				break;
		case 3: frmScreen.cboType3.value = CustomerXML.GetTagText("USAGE");
				frmScreen.txtTelNumber3.value = CustomerXML.GetTagText("TELEPHONENUMBER");
				frmScreen.txtExtensionNumber3.value = CustomerXML.GetTagText("EXTENSIONNUMBER");
				frmScreen.txtTime3.value = CustomerXML.GetTagText("CONTACTTIME");
				frmScreen.optPREFERREDCONTACT3.checked = (CustomerXML.GetTagText("PREFERREDMETHODOFCONTACT") == "1")? true:false;
				//JR - Omiplus24
				frmScreen.txtCountryCode3.value = CustomerXML.GetTagText("COUNTRYCODE");
				frmScreen.txtAreaCode3.value = CustomerXML.GetTagText("AREACODE");
				break;
		case 4: frmScreen.cboType4.value = CustomerXML.GetTagText("USAGE");
				frmScreen.txtTelNumber4.value = CustomerXML.GetTagText("TELEPHONENUMBER");
				frmScreen.txtExtensionNumber4.value = CustomerXML.GetTagText("EXTENSIONNUMBER");
				frmScreen.txtTime4.value = CustomerXML.GetTagText("CONTACTTIME");
				frmScreen.optPREFERREDCONTACT4.checked = (CustomerXML.GetTagText("PREFERREDMETHODOFCONTACT") == "1")? true:false;
				//JR - Omiplus24
				frmScreen.txtCountryCode4.value = CustomerXML.GetTagText("COUNTRYCODE");
				frmScreen.txtAreaCode4.value = CustomerXML.GetTagText("AREACODE");
				break;
		default: break;
	}
}

function SaveTelephone(nSpanNum)
{
	switch (nSpanNum)
	{
		case 1: if(frmScreen.cboType1.value != CustomerXML.GetTagText("USAGE"))
					CustomerXML.SetTagText("USAGE", frmScreen.cboType1.value);
				if(frmScreen.txtTelNumber1.value != CustomerXML.GetTagText("TELEPHONENUMBER"))
					CustomerXML.SetTagText("TELEPHONENUMBER", frmScreen.txtTelNumber1.value);
				if(frmScreen.txtExtensionNumber1.value != CustomerXML.GetTagText("EXTENSIONNUMBER"))
					CustomerXML.SetTagText("EXTENSIONNUMBER", frmScreen.txtExtensionNumber1.value);
				if(frmScreen.txtTime1.value != CustomerXML.GetTagText("CONTACTTIME"))
					CustomerXML.SetTagText("CONTACTTIME", frmScreen.txtTime1.value);
					CustomerXML.SetTagText("PREFERREDMETHODOFCONTACT", (frmScreen.optPREFERREDCONTACT1.checked)? "1":"0");
				//JR - Omiplus24
				if(frmScreen.txtCountryCode1.value != CustomerXML.GetTagText("COUNTRYCODE"))
					CustomerXML.SetTagText("COUNTRYCODE", frmScreen.txtCountryCode1.value );
				if(frmScreen.txtAreaCode1.value != CustomerXML.GetTagText("AREACODE"))
					CustomerXML.SetTagText("AREACODE", frmScreen.txtAreaCode1.value);					
				break;
		case 2: if(frmScreen.cboType2.value != CustomerXML.GetTagText("USAGE"))
					CustomerXML.SetTagText("USAGE", frmScreen.cboType2.value);
				if(frmScreen.txtTelNumber2.value != CustomerXML.GetTagText("TELEPHONENUMBER"))
					CustomerXML.SetTagText("TELEPHONENUMBER", frmScreen.txtTelNumber2.value);
				if(frmScreen.txtExtensionNumber2.value != CustomerXML.GetTagText("EXTENSIONNUMBER"))
					CustomerXML.SetTagText("EXTENSIONNUMBER", frmScreen.txtExtensionNumber2.value);
				if(frmScreen.txtTime2.value != CustomerXML.GetTagText("CONTACTTIME"))
					CustomerXML.SetTagText("CONTACTTIME", frmScreen.txtTime2.value);
					CustomerXML.SetTagText("PREFERREDMETHODOFCONTACT", (frmScreen.optPREFERREDCONTACT2.checked)? "1":"0");
				//JR - Omiplus24
				if(frmScreen.txtCountryCode2.value != CustomerXML.GetTagText("COUNTRYCODE"))
					CustomerXML.SetTagText("COUNTRYCODE", frmScreen.txtCountryCode2.value );
				if(frmScreen.txtAreaCode2.value != CustomerXML.GetTagText("AREACODE"))
					CustomerXML.SetTagText("AREACODE", frmScreen.txtAreaCode2.value);
				break;
		case 3: if(frmScreen.cboType3.value != CustomerXML.GetTagText("USAGE"))
					CustomerXML.SetTagText("USAGE", frmScreen.cboType3.value);
				if(frmScreen.txtTelNumber3.value != CustomerXML.GetTagText("TELEPHONENUMBER"))
					CustomerXML.SetTagText("TELEPHONENUMBER", frmScreen.txtTelNumber3.value);
				if(frmScreen.txtExtensionNumber3.value != CustomerXML.GetTagText("EXTENSIONNUMBER"))
					CustomerXML.SetTagText("EXTENSIONNUMBER", frmScreen.txtExtensionNumber3.value);
				if(frmScreen.txtTime3.value != CustomerXML.GetTagText("CONTACTTIME"))
					CustomerXML.SetTagText("CONTACTTIME", frmScreen.txtTime3.value);
					CustomerXML.SetTagText("PREFERREDMETHODOFCONTACT", (frmScreen.optPREFERREDCONTACT3.checked)? "1":"0");
				//JR - Omiplus24
				if(frmScreen.txtCountryCode3.value != CustomerXML.GetTagText("COUNTRYCODE"))
					CustomerXML.SetTagText("COUNTRYCODE", frmScreen.txtCountryCode3.value );
				if(frmScreen.txtAreaCode3.value != CustomerXML.GetTagText("AREACODE"))
					CustomerXML.SetTagText("AREACODE", frmScreen.txtAreaCode3.value);
				break;
		case 4: if(frmScreen.cboType4.value != CustomerXML.GetTagText("USAGE"))
					CustomerXML.SetTagText("USAGE", frmScreen.cboType4.value);
				if(frmScreen.txtTelNumber4.value != CustomerXML.GetTagText("TELEPHONENUMBER"))
					CustomerXML.SetTagText("TELEPHONENUMBER", frmScreen.txtTelNumber4.value);
				if(frmScreen.txtExtensionNumber4.value != CustomerXML.GetTagText("EXTENSIONNUMBER"))
					CustomerXML.SetTagText("EXTENSIONNUMBER", frmScreen.txtExtensionNumber4.value);
				if(frmScreen.txtTime4.value != CustomerXML.GetTagText("CONTACTTIME"))
					CustomerXML.SetTagText("CONTACTTIME", frmScreen.txtTime4.value);
					CustomerXML.SetTagText("PREFERREDMETHODOFCONTACT", (frmScreen.optPREFERREDCONTACT4.checked)? "1":"0");
				//JR - Omiplus24
				if(frmScreen.txtCountryCode4.value != CustomerXML.GetTagText("COUNTRYCODE"))
					CustomerXML.SetTagText("COUNTRYCODE", frmScreen.txtCountryCode4.value );
				if(frmScreen.txtAreaCode4.value != CustomerXML.GetTagText("AREACODE"))
					CustomerXML.SetTagText("AREACODE", frmScreen.txtAreaCode4.value);
				break;
		default: break;
	}
}

function SaveNewTelephone(nSpanNum)
{
	var newXML = null;
	CustomerXML.SelectTag(null, "CUSTOMERVERSION");
	var sCustomerNumber = CustomerXML.GetTagText("CUSTOMERNUMBER");
	var sCustomerVersion = CustomerXML.GetTagText("CUSTOMERVERSIONNUMBER");
	switch (nSpanNum)
	{
		case 1: if(frmScreen.cboType1.value != "" || 
				   frmScreen.txtTelNumber1.value != "" ||
				   frmScreen.txtExtensionNumber1.value != "" ||
				   frmScreen.txtTime1.value != "")
				{
					//we have a new entry
					newXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
					newXML.CreateActiveTag("START");
					newXML.CreateActiveTag("CUSTOMERTELEPHONENUMBER");
					newXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
					newXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersion);
					newXML.CreateTag("TELEPHONESEQUENCENUMBER","");
					newXML.CreateTag("USAGE", frmScreen.cboType1.value);
					newXML.CreateTag("TELEPHONENUMBER", frmScreen.txtTelNumber1.value);
					newXML.CreateTag("EXTENSIONNUMBER", frmScreen.txtExtensionNumber1.value);
					newXML.CreateTag("CONTACTTIME", frmScreen.txtTime1.value);
					newXML.CreateTag("PREFERREDMETHODOFCONTACT", (frmScreen.optPREFERREDCONTACT1.checked)? "1":"0");
					//JR - Omiplus24
					newXML.CreateTag("COUNTRYCODE", frmScreen.txtCountryCode1.value);
					newXML.CreateTag("AREACODE", frmScreen.txtAreaCode1.value);
				}
				break;
		case 2: if(frmScreen.cboType2.value != "" || 
				   frmScreen.txtTelNumber2.value != "" ||
				   frmScreen.txtExtensionNumber2.value != "" ||
				   frmScreen.txtTime2.value != "")
				{
					//we have a new entry
					newXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
					newXML.CreateActiveTag("START");
					newXML.CreateActiveTag("CUSTOMERTELEPHONENUMBER");
					newXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
					newXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersion);
					newXML.CreateTag("TELEPHONESEQUENCENUMBER","");
					newXML.CreateTag("USAGE", frmScreen.cboType2.value);
					newXML.CreateTag("TELEPHONENUMBER", frmScreen.txtTelNumber2.value);
					newXML.CreateTag("EXTENSIONNUMBER", frmScreen.txtExtensionNumber2.value);
					newXML.CreateTag("CONTACTTIME", frmScreen.txtTime2.value);
					newXML.CreateTag("PREFERREDMETHODOFCONTACT", (frmScreen.optPREFERREDCONTACT2.checked)? "1":"0");
					//JR - Omiplus24
					newXML.CreateTag("COUNTRYCODE", frmScreen.txtCountryCode2.value);
					newXML.CreateTag("AREACODE", frmScreen.txtAreaCode2.value);
				}
				break;
		case 3: if(frmScreen.cboType3.value != "" || 
				   frmScreen.txtTelNumber3.value != "" ||
				   frmScreen.txtExtensionNumber3.value != "" ||
				   frmScreen.txtTime3.value != "")
				{
					//we have a new entry
					newXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
					newXML.CreateActiveTag("START");
					newXML.CreateActiveTag("CUSTOMERTELEPHONENUMBER");
					newXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
					newXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersion);
					newXML.CreateTag("TELEPHONESEQUENCENUMBER","");
					newXML.CreateTag("USAGE", frmScreen.cboType3.value);
					newXML.CreateTag("TELEPHONENUMBER", frmScreen.txtTelNumber3.value);
					newXML.CreateTag("EXTENSIONNUMBER", frmScreen.txtExtensionNumber3.value);
					newXML.CreateTag("CONTACTTIME", frmScreen.txtTime3.value);
					newXML.CreateTag("PREFERREDMETHODOFCONTACT", (frmScreen.optPREFERREDCONTACT3.checked)? "1":"0");
					//JR - Omiplus24
					newXML.CreateTag("COUNTRYCODE", frmScreen.txtCountryCode3.value);
					newXML.CreateTag("AREACODE", frmScreen.txtAreaCode3.value);
				}
				break;
		case 4: if(frmScreen.cboType4.value != "" || 
				   frmScreen.txtTelNumber4.value != "" ||
				   frmScreen.txtExtensionNumber4.value != "" ||
				   frmScreen.txtTime4.value != "")
				{
					//we have a new entry
					newXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
					newXML.CreateActiveTag("START");
					newXML.CreateActiveTag("CUSTOMERTELEPHONENUMBER");
					newXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
					newXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersion);
					newXML.CreateTag("TELEPHONESEQUENCENUMBER","");
					newXML.CreateTag("USAGE", frmScreen.cboType4.value);
					newXML.CreateTag("TELEPHONENUMBER", frmScreen.txtTelNumber4.value);
					newXML.CreateTag("EXTENSIONNUMBER", frmScreen.txtExtensionNumber4.value);
					newXML.CreateTag("CONTACTTIME", frmScreen.txtTime4.value);
					newXML.CreateTag("PREFERREDMETHODOFCONTACT", (frmScreen.optPREFERREDCONTACT4.checked)? "1":"0");
					//JR - Omiplus24
					newXML.CreateTag("COUNTRYCODE", frmScreen.txtCountryCode4.value);
					newXML.CreateTag("AREACODE", frmScreen.txtAreaCode4.value);
				}
				break;
		default: break;
	}
	return(newXML);
}

function PopulateScreen()
{
	//Set Customer Name
	frmScreen.txtCustomerName.value = m_sCustomerName;
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtCustomerName");

	CustomerXML.ActiveTag = null;
	CustomerXML.CreateTagList("CUSTOMERTELEPHONENUMBER");
	var iCount;
	var iNumberOfEntries = CustomerXML.ActiveTagList.length;

	for (iCount = 0; iCount < iNumberOfEntries ; iCount++)
	{
		CustomerXML.ActiveTag = null;
		CustomerXML.CreateTagList("CUSTOMERTELEPHONENUMBER");  
		CustomerXML.SelectTagListItem(iCount);

		//For each telephone area populate
		PopulateTelephone(iCount + 1);
	}

	CustomerXML.SelectTag(null,"CUSTOMERVERSION");
	frmScreen.txtEMailAddress.value = CustomerXML.GetTagText("CONTACTEMAILADDRESS");
	frmScreen.optPREFERREDCONTACT5.checked = (CustomerXML.GetTagText("EMAILPREFERRED") == "1")? true:false;
	frmScreen.txtSalutation.value = CustomerXML.GetTagText("CORRESPONDENCESALUTATION");
	//DB BM0042 25/02/03 - Set Salutation to read-only.
	scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtSalutation");
}

function Initialise()
{
	CustomerXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	CustomerXML.LoadXML(m_sXMLIn);

	PopulateCombos();
	PopulateScreen();

	if(m_sReadOnly == "1")
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	else
	{
		frmScreen.cboType1.onchange();
		frmScreen.cboType2.onchange();
		frmScreen.cboType3.onchange();
		frmScreen.cboType4.onchange();
		frmScreen.cboType1.focus();	
	}
}

function WriteContactDetails()
{
	//save any changed data
	if ((m_sReadOnly != "1") && frmScreen.onsubmit())
	{
		if (IsChanged() == true )
		{
			//Change the relevant parts of the CustomerXML
			CustomerXML.ActiveTag = null;
			CustomerXML.CreateTagList("CUSTOMERTELEPHONENUMBER");
			var iCount;
			var iNumberOfEntries = CustomerXML.ActiveTagList.length;

			for (iCount = 0; iCount < iNumberOfEntries ; iCount++)
			{
				CustomerXML.ActiveTag = null;
				CustomerXML.CreateTagList("CUSTOMERTELEPHONENUMBER");  
				CustomerXML.SelectTagListItem(iCount);
				SaveTelephone(iCount+1);
			}

			CustomerXML.ActiveTag = null;
			CustomerXML.SelectTag(null,"CUSTOMERVERSION");
			if(frmScreen.txtEMailAddress.value != CustomerXML.GetTagText("CONTACTEMAILADDRESS"))
				CustomerXML.SetTagText("CONTACTEMAILADDRESS", frmScreen.txtEMailAddress.value);
			if(frmScreen.txtSalutation.value != CustomerXML.GetTagText("CORRESPONDENCESALUTATION"))
				CustomerXML.SetTagText("CORRESPONDENCESALUTATION", frmScreen.txtSalutation.value);
			// No Check for existing value as negligable benefit in terms of processing:
			CustomerXML.SetTagText("EMAILPREFERRED", frmScreen.optPREFERREDCONTACT5.checked ? "1":"0");

			//see if we have any new ones to add to the XML
			if(iNumberOfEntries < 4)
			{
				var newXML;
				if( iNumberOfEntries == 0 )
				{
					//We need to create the CUSTOMERTELEPHONENUMBERLIST tag
					CustomerXML.SelectTag(null, "CUSTOMER");
					CustomerXML.CreateActiveTag("CUSTOMERTELEPHONENUMBERLIST");
				}
				var iStartBlock = iNumberOfEntries + 1;
				var iNewBlocksAdded = 0;
				for(var iNew = 0; iNew < (4 - iNumberOfEntries); iNew++)
				{
					CustomerXML.ActiveTag = null;

					newXML = SaveNewTelephone(iStartBlock);
					if(newXML != null)
					{
						var tag = newXML.SelectTag(null,"START");
						CustomerXML.SelectTag(null,"CUSTOMERTELEPHONENUMBERLIST");
						CustomerXML.AddXMLBlock(tag);
						iNewBlocksAdded++;
					}
					iStartBlock++;
				}
			}
		}
	}

	return(CustomerXML.XMLDocument.xml);
}

function frmScreen.btnOK.onclick()
{
	var sValidatePhones = ValidateTelephoneDetails() ;
	if (sValidatePhones)
	{
		var sReturn = new Array();

		sReturn[0] = IsChanged();		// Has there been a change made
		sReturn[1] = WriteContactDetails();		// The XML string

		window.returnValue = sReturn;
		window.close();
	}	
}

function frmScreen.cboType1.onchange()
{
scScreenFunctions.SetFieldState(frmScreen, "txtExtensionNumber1", ((frmScreen.cboType1.value) == "2") ? "W":"D");
scScreenFunctions.SetFieldState(frmScreen, "optPREFERREDCONTACT1", ((frmScreen.cboType1.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTelNumber1", ((frmScreen.cboType1.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTime1", ((frmScreen.cboType1.value) == "") ? "D":"W");
//JR - Omiplus24
scScreenFunctions.SetFieldState(frmScreen, "txtCountryCode1", ((frmScreen.cboType1.value) =="") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtAreaCode1", ((frmScreen.cboType1.value) =="") ? "D":"W");
}
function frmScreen.cboType2.onchange()
{
scScreenFunctions.SetFieldState(frmScreen, "txtExtensionNumber2", ((frmScreen.cboType2.value) == "2") ? "W":"D");
scScreenFunctions.SetFieldState(frmScreen, "optPREFERREDCONTACT2", ((frmScreen.cboType2.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTelNumber2", ((frmScreen.cboType2.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTime2", ((frmScreen.cboType2.value) == "") ? "D":"W");
//JR - Omiplus24
scScreenFunctions.SetFieldState(frmScreen, "txtCountryCode2", ((frmScreen.cboType2.value) =="") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtAreaCode2", ((frmScreen.cboType2.value) =="") ? "D":"W");
}
function frmScreen.cboType3.onchange()
{
scScreenFunctions.SetFieldState(frmScreen, "txtExtensionNumber3", ((frmScreen.cboType3.value) == "2") ? "W":"D");
scScreenFunctions.SetFieldState(frmScreen, "optPREFERREDCONTACT3", ((frmScreen.cboType3.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTelNumber3", ((frmScreen.cboType3.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTime3", ((frmScreen.cboType3.value) == "") ? "D":"W");
//JR - Omiplus24
scScreenFunctions.SetFieldState(frmScreen, "txtCountryCode3", ((frmScreen.cboType3.value) =="") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtAreaCode3", ((frmScreen.cboType3.value) =="") ? "D":"W");
}
function frmScreen.cboType4.onchange()
{
scScreenFunctions.SetFieldState(frmScreen, "txtExtensionNumber4", ((frmScreen.cboType4.value) == "2") ? "W":"D");
scScreenFunctions.SetFieldState(frmScreen, "optPREFERREDCONTACT4", ((frmScreen.cboType4.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTelNumber4", ((frmScreen.cboType4.value) == "") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtTime4", ((frmScreen.cboType4.value) == "") ? "D":"W");
//JR - Omiplus24
scScreenFunctions.SetFieldState(frmScreen, "txtCountryCode4", ((frmScreen.cboType4.value) =="") ? "D":"W");
scScreenFunctions.SetFieldState(frmScreen, "txtAreaCode4", ((frmScreen.cboType4.value) =="") ? "D":"W");
}

function frmScreen.btnCancel.onclick()
{
	window.close();
}
function ValidateTelephoneDetails()
{
	var sType	= new Array("cboType1", "cboType2", "cboType3" , "cboType4");
	var	sNumber	= new Array("txtTelNumber1", "txtTelNumber2", "txtTelNumber3", "txtTelNumber4");
	var sTime	= new Array("txtTime1", "txtTime2", "txtTime3", "txtTime4");
	
	for(var nLoop = 0;nLoop < 4;nLoop++)
	{
		if (frmScreen(sType[nLoop]).value !== '')
		{
			if (frmScreen(sNumber[nLoop]).value === '')
			{
				alert('Please enter a Telephone Number for each Contact Type')
				frmScreen.all(sType[nLoop]).focus()
				return false ;	
			}
		}
		else
		{
			if (frmScreen(sNumber[nLoop]).value !== '' || frmScreen(sTime[nLoop]).value !== '')
			{
				alert('For each telephone number please enter the Usage and Telephone Number')
				frmScreen.all(sType[nLoop]).focus()
				return false ;	
			}
		}
	}
	return true;
}
-->
</script>
</body>
</html>
