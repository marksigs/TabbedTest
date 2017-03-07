<%@ LANGUAGE="JSCRIPT" %>
<html>
<comment>
Workfile:      DC222.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Leasehold Details (Popup)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		07/03/2000	Created
AY		31/03/00	scScreenFunctions change
AY		12/04/00	SYS0328 - Dynamic currency display
IW		03/05/00	SYS0624 - MISC (Validation changes)
MH      03/05/00    SYS0450 - Validation changes
IW		10/05/00	SYS0722 Resolved ReadOnly Screen issues
MC		01/10/01	SYS2759 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
HMA     17/09/03    BM0063  Amend HTML text for radio buttons
MC		19/04/2004	BMIDS517	white space padded to the title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
</comment>
<head>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>DC222 - Leasehold Details <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 264px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 308px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 20px" class="msgLabel">
		<LABEL id="idOriginalTerm"></LABEL>
		<span style="LEFT: 145px; POSITION: absolute; TOP: -3px">
			<input id="txtOriginalTermOfLeaseYears" maxlength="3" style="POSITION: absolute; WIDTH: 30px" class="msgTxt">
			<span style="LEFT: 34px; POSITION: absolute; TOP: 3px" class="msgLabel">
				Years
			</span>
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 44px" class="msgLabel">
		<LABEL id="idUnexpiredTerm"></LABEL>
		<span style="LEFT: 145px; POSITION: absolute; TOP: -3px">
			<input id="txtUnexpiredTermOfLeaseYears"  maxlength="3" style="POSITION: absolute; WIDTH: 30px" class="msgTxt">
			<span style="LEFT: 34px; POSITION: absolute; TOP: 3px" class="msgLabel">
				Years
			</span>
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 68px" class="msgLabel">
		Annual Ground Rent
		<span style="LEFT: 145px; POSITION: absolute; TOP: 0px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
			<input id="txtGroundRent" maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 80px" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 92px" class="msgLabel">
		Date Lease Started
		<span style="LEFT: 145px; POSITION: absolute; TOP: -3px">
			<input id="txtDateLeaseStarted"  maxlength="10" style="POSITION: absolute; WIDTH: 80px" class="msgTxt">
		</span> 
	</span>
	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 109px" class="msgLabel">
		Can Ground Rent<br>be increased?
		<span style="LEFT: 145px; POSITION: absolute; TOP: 4px">
			<input id="optGroundRentYes" name="GroundRentGroup" type="radio" value="1" onclick="CanGroundRentIncreaseChanged()"><label for="optGroundRentYes" class="msgLabel">Yes</label> 
		</span> 
		<span style="LEFT: 195px; POSITION: absolute; TOP: 4px">
			<input id="optGroundRentNo" name="GroundRentGroup" type="radio" value="0" onclick="CanGroundRentIncreaseChanged()"><label for="optGroundRentNo" class="msgLabel">No</label> 
		</span> 
	</span>		

	<span style="LEFT: 4px; POSITION: absolute; TOP: 140px" class="msgLabel">
		Increased Rent Details
		<span style="LEFT: 0px; POSITION: absolute; TOP: 16px"><TEXTAREA class=msgTxt id=txtIncreasedRentDetails rows=5 style="POSITION: absolute; WIDTH: 300px">			</TEXTAREA> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 240px" class="msgLabel">
		<LABEL id="idServiceCharges"></LABEL>
		<span style="LEFT: 145px; POSITION: absolute; TOP: 0px">
			<label style="LEFT: -10px; POSITION: absolute; TOP: 0px" class="msgCurrency"></label>
			<input id="txtServiceCharges"  maxlength="10" style="POSITION: absolute; TOP: -3px; WIDTH: 80px" class="msgTxt">
		</span> 
	</span>
</div>

<span style="LEFT: 10px; POSITION: absolute; TOP: 284px">
	<input id="btnOK" value="OK"  type="button" style="WIDTH: 60px" class="msgButton">
	<span style="LEFT: 70px; POSITION: absolute; TOP: 0px">
		<input id="btnCancel" value="Cancel"  type="button" style="WIDTH: 60px" class="msgButton">
	</span>
</span>
</form>

<!-- #include FILE="attribs/DC222attribs.asp" -->
<!--  #include FILE="Customise/DC222Customise.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;
	
var LeaseholdXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_BaseNonPopupWindow = null;

/* EVENTS */

function frmScreen.btnCancel.onclick()
{
	LeaseholdXML = null;
	window.close();
}

function frmScreen.btnOK.onclick()
{
	if (CommitChanges())
	{
		LeaseholdXML = null;
		window.close();
	}
}

function window.onload()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	var sParameters	= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	LeaseholdXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	LeaseholdXML.LoadXML(sParameters[0]);
	m_sReadOnly	= sParameters[1];
	m_sApplicationNumber = sParameters[2];
	m_sApplicationFactFindNumber = sParameters[3];
	scScreenFunctions.SetThisCurrency(frmScreen,sParameters[4]);

	SetMasks();
	// MC SYS2564/SYS2759 for client customisation
	Customise();
	
	Validation_Init();	
	Initialise();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

/* FUNCTIONS */

function frmScreen.txtIncreasedRentDetails.onkeydown()
{	
if (this.value.length>=255)
	{
		alert("Cannot enter more than 255 characters");
		this.value = this.value.substr(0,this.value.length-1)
		this.focus();
	}	
}

function CanGroundRentIncreaseChanged()
{
	if (scScreenFunctions.GetRadioGroupValue(frmScreen,"GroundRentGroup") == "0")
	{
		frmScreen.txtIncreasedRentDetails.value = "";
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtIncreasedRentDetails");
	}
	else
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtIncreasedRentDetails");
}

function CommitChanges()
{
	function ValidateScreen()
	{
		if ((frmScreen.txtOriginalTermOfLeaseYears.value!="") && (frmScreen.txtOriginalTermOfLeaseYears.value - frmScreen.txtUnexpiredTermOfLeaseYears.value < 0))
		{
			alert("Unexpired Term of Lease cannot be greater than the Original Term of Lease.");
			frmScreen.txtUnexpiredTermOfLeaseYears.focus();
			return(false);
		}	
		return(true);
	}

	var bSuccess = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				if (ValidateScreen())
					bSuccess = SaveLeasehold();
				else
					bSuccess = false;
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function Initialise()
// Initialises the screen
{
	PopulateScreen();
	CanGroundRentIncreaseChanged();
	if(m_sReadOnly == "1")
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateScreen()
// Populates the screen with details of the item selected in 
{
	LeaseholdXML.SelectTag(null,"NEWPROPERTYLEASEHOLD");
	if (LeaseholdXML.ActiveTag != null)
		with (frmScreen)
		{
			txtDateLeaseStarted.value = LeaseholdXML.GetTagText("DATELEASESTARTED");
			txtGroundRent.value = LeaseholdXML.GetTagText("GROUNDRENT");
			txtOriginalTermOfLeaseYears.value = LeaseholdXML.GetTagText("ORIGINALTERMOFLEASEYEARS");
			txtServiceCharges.value = LeaseholdXML.GetTagText("SERVICECHARGE");
			txtUnexpiredTermOfLeaseYears.value = LeaseholdXML.GetTagText("UNEXPIREDTERMOFLEASEYEARS");
			txtIncreasedRentDetails.value = LeaseholdXML.GetTagText("INCREASEDRENTDETAILS");
			scScreenFunctions.SetRadioGroupValue(frmScreen,"GroundRentGroup",(txtIncreasedRentDetails.value == "") ? "0" : "1");
		}
	else
		scScreenFunctions.SetRadioGroupValue(frmScreen,"GroundRentGroup","0");
}

function SaveLeasehold()
{
	var sReturn = new Array();
	if (IsChanged())
	{
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateActiveTag("NEWPROPERTYLEASEHOLD")
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("DATELEASESTARTED", frmScreen.txtDateLeaseStarted.value);
		XML.CreateTag("GROUNDRENT", frmScreen.txtGroundRent.value);
		XML.CreateTag("INCREASEDRENTDETAILS", frmScreen.txtIncreasedRentDetails.value);
		XML.CreateTag("ORIGINALTERMOFLEASEYEARS", frmScreen.txtOriginalTermOfLeaseYears.value);
		XML.CreateTag("SERVICECHARGE", frmScreen.txtServiceCharges.value);
		XML.CreateTag("UNEXPIREDTERMOFLEASEYEARS", frmScreen.txtUnexpiredTermOfLeaseYears.value);
		LeaseholdXML = XML.XMLDocument.xml;
	}
	
	sReturn[0] = IsChanged();
	sReturn[1] = LeaseholdXML;
	window.returnValue	= sReturn;
	window.close();
}
-->
</script>
</body>
</html>


