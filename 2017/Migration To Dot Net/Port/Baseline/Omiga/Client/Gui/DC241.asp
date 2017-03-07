<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      DC241.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Popup screen to display/capture Contact 
			   Telephone Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MDC		18/06/01	Created
JR		09/10/01	Omiplus24, amended functionality of screen
DRC     24/01/02    AQR 3487 Change width of area code fields to 5
GD		06/02/02    SYS4005 Ensure phone number details aren't nested in XML on a create (for first time)
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
ASU		18/09/2002	BMIDS00106	Telephone number validation for alpha characters added  
MDC		12/12/2002	BM0094 Legal Rep Contact Details
KRW     18/10/2004  BM0472 Problems with Type Combo resolved 
KRW     19/10/2004  BM0472 Blank out corresponding fields when SELECT (erm... selected) 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<head>
	<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Contact Details <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 180px; WIDTH: 420px; POSITION: ABSOLUTE" class="msgGroup">

	<span style="LEFT: 4px; POSITION: absolute; TOP: 30px" class="msgLabel">
		<strong>Telephone Details</strong>
	</span>

	<% /* Telephone Numbers  */  %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 56px" class="msgLabel">
		Type
		<span style="LEFT: 0px; POSITION: absolute; TOP: 20px">
			<select id="cboType1" style="WIDTH: 100px" class="msgCombo" onchange="SetWorkExtState(1, null)"></select>
		</span>
	</span>
	<span style="LEFT: 110px; POSITION: absolute; TOP: 46px" class="msgLabel">
		&nbsp;Country
		<span style="LEFT: 0px; POSITION: absolute; TOP: 10px" class="msgLabel">
			&nbsp;Code
			<span style="LEFT: 0px; POSITION: absolute; TOP: 20px">
				<input id="txtCountryCode1" name="CountryCode1" maxlength="3" style="WIDTH: 50px; POSITION: absolute" class="msgTxt">
			</span>
		</span>	
	</span>
	<span style="LEFT: 166px; POSITION: absolute; TOP: 46px" class="msgLabel">
		&nbsp;Area
		<span style="LEFT: 0px; POSITION: absolute; TOP: 10px" class="msgLabel">
			&nbsp;Code			
			<span style="LEFT: 0px; POSITION: absolute; TOP: 20px">
				<input id="txtAreaCode1" name="AreaCode1" maxlength="6" style="WIDTH: 50px; POSITION: absolute" class="msgTxt">
			</span>
		</span>			
	</span>
	<span style="LEFT: 222px; POSITION: absolute; TOP: 46px" class="msgLabel">
		&nbsp;Telephone
		<span style="LEFT: 0px; POSITION: absolute; TOP: 10px" class="msgLabel">
		&nbsp;Number
			<span style="LEFT: 0px; POSITION: absolute; TOP: 20px">
				<input id="txtTelephoneNumber1" name="TelephoneNumber1" maxlength="10" style="WIDTH: 100px; POSITION: absolute" class="msgTxt">
			</span>
		</span>		
	</span>
	<span style="LEFT: 327px; POSITION: absolute; TOP: 46px" class="msgLabel">
		&nbsp;Work Extension
		<span style="LEFT: 0px; POSITION: absolute; TOP: 10px" class="msgLabel">
		&nbsp;Number
			<span style="LEFT: 0px; POSITION: absolute; TOP: 20px">
				<input id="txtWorkExtNo1" name="WorkExtensionNumber1" maxlength="6" style="WIDTH: 60px; POSITION: absolute" class="msgTxt">
			</span>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 96px" class="msgLabel">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 10px">
			<select id="cboType2" style="WIDTH: 100px; POSITION: absolute" class="msgCombo" onchange="SetWorkExtState(2, null)"></select>
		</span>		
		<span style="LEFT: 106px; POSITION: absolute; TOP: 10px" class="msgLabel">
			<input id="txtCountryCode2" name="CountryCode2" maxlength="3" style="WIDTH: 50px; POSITION: absolute" class="msgTxt">
		</span>
		<span style="LEFT: 162px; POSITION: absolute; TOP: 10px">
			<input id="txtAreaCode2" name="AreaCode2" maxlength="6" style="WIDTH: 50px; POSITION: absolute" class="msgTxt">
		</span>
		<span style="LEFT: 218px; POSITION: absolute; TOP: 10px">
			<input id="txtTelephoneNumber2" name="TelephoneNumber2" maxlength="10" style="WIDTH: 100px; POSITION: absolute" class="msgTxt">
		</span>
		<span style="LEFT: 323px; POSITION: absolute; TOP: 10px">
			<input id="txtWorkExtNo2" name="WorkExtensionNumber2" maxlength="6" style="WIDTH: 60px; POSITION: absolute" class="msgTxt">
		</span>
	</span>
			
	<% /* E-mail Address */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 150px" class="msgLabel">
		E-mail Address
		<span style="LEFT: 80px; POSITION: absolute; TOP: 0px">
			<input id="txtEmail" name="Email" maxlength="255" style="WIDTH: 302px; POSITION: absolute" class="msgTxt">
		</span>
	</span>

</div>
</form>

<div id="msgButtons" style="LEFT: 8px; WIDTH: 312px; POSITION: absolute; TOP: 200px; HEIGHT: 19px">
	<!-- #include FILE="msgButtons.asp" --> 
</div> 

<!--  #include FILE="attribs/DC241attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var scScreenFunctions;
var m_sReadOnly = "";
var m_sContactXML = "";
var XMLContact = null;
var m_XMLTelList = null;
var m_ContactDetailsTag = null;
var m_strComboId = "";
var m_strFaxComboId = ""; //JR - Omiplus24
var m_BaseNonPopupWindow = null;

function window.onload()
{
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	RetrieveData();
	XMLContact = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	SetMasks();

	GetComboLists();
	m_strComboId = XMLContact.GetComboIdForValidation("ContactTelephoneUsage", "W", null, document);	RetrieveData();
	m_strFaxComboId = XMLContact.GetComboIdForValidation("ContactTelephoneUsage", "F", null, document);	RetrieveData();
	Validation_Init();
	SetScreenOnReadOnly();
	PopulateContactTelephoneDetails();
	window.returnValue = null;
}

function GetComboLists()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("ContactTelephoneUsage");
	var bSuccess = false;

	if(XML.GetComboLists(document,sGroups))
	{
		bSuccess = true;
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboType1,"ContactTelephoneUsage",true);
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboType2,"ContactTelephoneUsage",true);
	}

	if(!bSuccess)
	{
		SetScreenToReadOnly();
		DisableMainButton("Submit");
	}
}

function SetScreenOnReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
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
	m_sReadOnly			= sParameters[0];
	m_sContactXML		= sParameters[1];
	
}

function btnCancel.onclick()
{
	window.returnValue	= null;
	window.close();
}

function btnSubmit.onclick()
{
	var sReturn = new Array();
	
	sReturn[0] = IsChanged();

	if(IsChanged())
	{
		//ASu BMIDS00106 - Start
		if(frmScreen.onsubmit())
		{
			if(ValidateContactDetails())
				SaveContactTelephoneDetails();
			else
			{
				alert("Type, Area Code and Telephone Number must be entered for each required item.");
				return;
			}
			sReturn[1] = m_sContactXML;
			window.returnValue	= sReturn;
			window.close();
		} 
		else return;
		//ASu - End
	}
	sReturn[1] = m_sContactXML;
	window.returnValue	= sReturn;
	window.close();	

}

function SetWorkExtState(intIndex, strValue)
{

	var intSelectOne = 0;
	var intSelectTwo = 0;


	if(strValue == null)
		strValue = "";
	
	
	if(frmScreen.cboType1.value == "" ) // KRW BM0472 18/10/1004
    {
		intSelectOne = 1;
		frmScreen.txtCountryCode1.value = "";
		frmScreen.txtAreaCode1.value = "";
		frmScreen.txtTelephoneNumber1.value = "";
		frmScreen.txtWorkExtNo1.value = "";
	}
	
	if(frmScreen.cboType2.value == "") // KRW BM0472 18/10/1004
	{
	   intSelectTwo = 1;
		frmScreen.txtCountryCode2.value = "";
		frmScreen.txtAreaCode2.value = "";
		frmScreen.txtTelephoneNumber2.value = "";
		frmScreen.txtWorkExtNo2.value = "";
	}
	

	
	if(intIndex == 1)
	{
		//JR - if value = fax, force it to Work
		/* NOTE *****************
		This will later on need to be more efficiently coded to re-populate
		the other combo, removing this selected combogroup
		********************************************************************/ 
		if(frmScreen.cboType1.value == frmScreen.cboType2.value && (intSelectOne + intSelectTwo < 1))// KRW BM0472 18/10/1004
		{
			if(frmScreen.cboType2.value == m_strComboId)
				frmScreen.cboType1.value = m_strFaxComboId;
			else
				frmScreen.cboType1.value = m_strComboId;
		}
		//End
		if(frmScreen.cboType1.value == m_strComboId)
		{
			// Usage is 'Work' so enable work telephone number
			frmScreen.txtWorkExtNo1.value = strValue;
			frmScreen.txtWorkExtNo1.disabled = false;
			if (m_sReadOnly=="0")
				scScreenFunctions.SetFieldToWritable(frmScreen, "txtWorkExtNo1");			
		}
		else
		{
			// Usage not 'work' do reset and disable telephone number
			frmScreen.txtWorkExtNo1.value = "";
			frmScreen.txtWorkExtNo1.disabled = true;
			scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtWorkExtNo1");
		}
	}
	else if(intIndex == 2)
	{
		//JR - if value = work, force it to Fax
		/* NOTE *****************
		This will later on need to be more efficiently coded to re-populate
		the other combo, removing this selected combogroup
		********************************************************************/ 
		if(frmScreen.cboType2.value == frmScreen.cboType1.value && (intSelectOne + intSelectTwo < 1)) // KRW BM0472 18/10/1004
		{
			if(frmScreen.cboType1.value == m_strComboId)
				frmScreen.cboType2.value = m_strFaxComboId;
			else
				frmScreen.cboType2.value = m_strComboId;
		}		
		//End
		if(frmScreen.cboType2.value == m_strComboId)
		{
			// Usage is 'Work' so enable work telephone number
			frmScreen.txtWorkExtNo2.value = strValue;
			frmScreen.txtWorkExtNo2.disabled = false;
			if (m_sReadOnly=="0")
				scScreenFunctions.SetFieldToWritable(frmScreen, "txtWorkExtNo2");
		}
		else
		{
			// Usage not 'work' do reset and disable telephone number
			frmScreen.txtWorkExtNo2.value = "";
			frmScreen.txtWorkExtNo2.disabled = true;			
			scScreenFunctions.SetFieldToReadOnly(frmScreen, "txtWorkExtNo2");
		}
	}
}

function PopulateContactTelephoneDetails()
{
	if(m_sContactXML != "")
	{
		XMLContact.LoadXML(m_sContactXML);
		m_ContactDetailsTag = XMLContact.SelectTag(null, "CONTACTDETAILS");
		
		// Email
		frmScreen.txtEmail.value = XMLContact.GetTagText("EMAILADDRESS");
		
		// Telephone Numbers
		m_XMLTelList = XMLContact.CreateTagList("CONTACTTELEPHONEDETAILS");
		if(m_XMLTelList.length > 0)
		{
			// Telephone Number 1
			if(XMLContact.SelectTagListItem(0))
			{
				frmScreen.cboType1.value = XMLContact.GetTagText("USAGE");
				frmScreen.txtCountryCode1.value = XMLContact.GetTagText("COUNTRYCODE");
				frmScreen.txtAreaCode1.value = XMLContact.GetTagText("AREACODE");
				frmScreen.txtTelephoneNumber1.value = XMLContact.GetTagText("TELENUMBER");
				SetWorkExtState(1, XMLContact.GetTagText("EXTENSIONNUMBER"));
			}		
			if(m_XMLTelList.length > 1)
			{
				// Telephone Number 2
				if(XMLContact.SelectTagListItem(1))
				{
					frmScreen.cboType2.value = XMLContact.GetTagText("USAGE");
					frmScreen.txtCountryCode2.value = XMLContact.GetTagText("COUNTRYCODE");
					frmScreen.txtAreaCode2.value = XMLContact.GetTagText("AREACODE");
					frmScreen.txtTelephoneNumber2.value = XMLContact.GetTagText("TELENUMBER");
					SetWorkExtState(2, XMLContact.GetTagText("EXTENSIONNUMBER"))
				}
			}
		}
	}
}

function SaveContactTelephoneDetails()
{
var xmlListTag = null;
<% /* BM0094 MDC 12/12/2002 */ %>
var TempTag = null;
var xmlActiveTag = null;
<% /* BM0094 MDC 12/12/2002 - End */ %>

	if(m_ContactDetailsTag == null)
	{
		//Create Contact Details
		m_ContactDetailsTag = XMLContact.CreateActiveTag("CONTACTDETAILS");

		// Email Address
		XMLContact.CreateTag("EMAILADDRESS", frmScreen.txtEmail.value); 
		CreateTelephoneDetails();
	}
	else
	{
		// Email Address
		XMLContact.ActiveTag = m_ContactDetailsTag;
		<% /* BM0094 MDC 12/12/2002 - Handle the case where EMAILADDRESS is not present */ %>
		// XMLContact.SetTagText("EMAILADDRESS", frmScreen.txtEmail.value);
		TempTag = XMLContact.SelectTag(m_ContactDetailsTag,"EMAILADDRESS")
		XMLContact.ActiveTag = m_ContactDetailsTag;
		if(TempTag == null)
			XMLContact.CreateTag("EMAILADDRESS", frmScreen.txtEmail.value); 
		else
			XMLContact.SetTagText("EMAILADDRESS", frmScreen.txtEmail.value);
		<% /* BM0094 MDC 12/12/2002 - End */ %>
		
		if(m_XMLTelList != null)
		{			
			// Telephone Number 1
			if(m_XMLTelList.length > 0)
			{
				// Update existing first telephone number
				XMLContact.SelectTagListItem(0);
				XMLContact.SetTagText("USAGE",frmScreen.cboType1.value);		
				<% /* BM0094 MDC 12/12/2002 - Handle the case where EMAILADDRESS is not present */ %>
				xmlActiveTag = XMLContact.ActiveTag;
				// XMLContact.SetTagText("COUNTRYCODE",frmScreen.txtCountryCode1.value);		
				TempTag = XMLContact.SelectTag(xmlActiveTag,"COUNTRYCODE")
				XMLContact.ActiveTag = xmlActiveTag;
				if(TempTag == null)
					XMLContact.CreateTag("COUNTRYCODE", frmScreen.txtCountryCode1.value); 
				else
					XMLContact.SetTagText("COUNTRYCODE", frmScreen.txtCountryCode1.value);
				// XMLContact.SetTagText("AREACODE",frmScreen.txtAreaCode1.value);		
				TempTag = XMLContact.SelectTag(xmlActiveTag,"AREACODE")
				XMLContact.ActiveTag = xmlActiveTag;
				if(TempTag == null)
					XMLContact.CreateTag("AREACODE", frmScreen.txtAreaCode1.value); 
				else
					XMLContact.SetTagText("AREACODE", frmScreen.txtAreaCode1.value);
				// XMLContact.SetTagText("TELENUMBER",frmScreen.txtTelephoneNumber1.value);		
				TempTag = XMLContact.SelectTag(xmlActiveTag,"TELENUMBER")
				XMLContact.ActiveTag = xmlActiveTag;
				if(TempTag == null)
					XMLContact.CreateTag("TELENUMBER", frmScreen.txtTelephoneNumber1.value); 
				else
					XMLContact.SetTagText("TELENUMBER", frmScreen.txtTelephoneNumber1.value);
				// XMLContact.SetTagText("EXTENSIONNUMBER",frmScreen.txtWorkExtNo1.value);						
				TempTag = XMLContact.SelectTag(xmlActiveTag,"EXTENSIONNUMBER")
				XMLContact.ActiveTag = xmlActiveTag;
				if(TempTag == null)
					XMLContact.CreateTag("EXTENSIONNUMBER", frmScreen.txtWorkExtNo1.value); 
				else
					XMLContact.SetTagText("EXTENSIONNUMBER", frmScreen.txtWorkExtNo1.value);
				<% /* BM0094 MDC 12/12/2002 - End */ %>
			}
			else if(IsTelephoneEntered(1))
			{
				//XMLContact.CreateActiveTag("CONTACTTELEPHONEDETAILSLIST");
				XMLContact.ActiveTag = m_ContactDetailsTag; //JR - omiplus24
				CreateTelephoneDetailsListItem(1);
			}

			// Telephone Number 2
			if(m_XMLTelList.length > 1)
			{
				// Update existing second telephone number
				XMLContact.SelectTagListItem(1);
				XMLContact.SetTagText("USAGE",frmScreen.cboType2.value);		
				<% /* BM0094 MDC 12/12/2002 - Handle the case where tags are not present */ %>
				// XMLContact.SetTagText("COUNTRYCODE",frmScreen.txtCountryCode2.value);		
				// XMLContact.SetTagText("AREACODE",frmScreen.txtAreaCode2.value);		
				// XMLContact.SetTagText("TELENUMBER",frmScreen.txtTelephoneNumber2.value);		
				// XMLContact.SetTagText("EXTENSIONNUMBER",frmScreen.txtWorkExtNo2.value);		
				xmlActiveTag = XMLContact.ActiveTag;
				TempTag = XMLContact.SelectTag(xmlActiveTag,"COUNTRYCODE")
				XMLContact.ActiveTag = xmlActiveTag;
				if(TempTag == null)
					XMLContact.CreateTag("COUNTRYCODE", frmScreen.txtCountryCode2.value); 
				else
					XMLContact.SetTagText("COUNTRYCODE", frmScreen.txtCountryCode2.value);

				TempTag = XMLContact.SelectTag(xmlActiveTag,"AREACODE")
				XMLContact.ActiveTag = xmlActiveTag;
				if(TempTag == null)
					XMLContact.CreateTag("AREACODE", frmScreen.txtAreaCode2.value); 
				else
					XMLContact.SetTagText("AREACODE", frmScreen.txtAreaCode2.value);

				TempTag = XMLContact.SelectTag(xmlActiveTag,"TELENUMBER")
				XMLContact.ActiveTag = xmlActiveTag;
				if(TempTag == null)
					XMLContact.CreateTag("TELENUMBER", frmScreen.txtTelephoneNumber2.value); 
				else
					XMLContact.SetTagText("TELENUMBER", frmScreen.txtTelephoneNumber2.value);

				TempTag = XMLContact.SelectTag(xmlActiveTag,"EXTENSIONNUMBER")
				XMLContact.ActiveTag = xmlActiveTag;
				if(TempTag == null)
					XMLContact.CreateTag("EXTENSIONNUMBER", frmScreen.txtWorkExtNo2.value); 
				else
					XMLContact.SetTagText("EXTENSIONNUMBER", frmScreen.txtWorkExtNo2.value);
				<% /* BM0094 MDC 12/12/2002 - End */ %>
			}
			else if(IsTelephoneEntered(2))
			{
				// Second Telephone Number is new
				XMLContact.ActiveTag = m_ContactDetailsTag;
				CreateTelephoneDetailsListItem(2);
			}
		}
		else
		{
			CreateTelephoneDetails();
		}
			
	}

	m_sContactXML = XMLContact.XMLDocument.xml;

}

function IsTelephoneEntered(intIndex)
{
	if(intIndex == 1)
	{
		// Telephone Number 1
		if(frmScreen.cboType1.value != "")
			return true;
	}
	else if(intIndex == 2)
	{
		// Telephone Number 2
		if(frmScreen.cboType2.value != "")
			return true;
	
	}
	
	return false;
}

function ValidateContactDetails()
{
	if(frmScreen.cboType1.value != "" || frmScreen.txtCountryCode1.value != "" || frmScreen.txtAreaCode1.value != "" 
									|| frmScreen.txtTelephoneNumber1.value != "" || frmScreen.txtWorkExtNo1.value != "")
		if(frmScreen.cboType1.value == "" || frmScreen.txtAreaCode1.value == "" || frmScreen.txtTelephoneNumber1.value == "")
			return false;
	
	if(frmScreen.cboType2.value != "" || frmScreen.txtCountryCode2.value != "" || frmScreen.txtAreaCode2.value != "" 
									|| frmScreen.txtTelephoneNumber2.value != "" || frmScreen.txtWorkExtNo2.value != "")
		if(frmScreen.cboType2.value == "" || frmScreen.txtAreaCode2.value == "" || frmScreen.txtTelephoneNumber2.value == "")
			return false;
		
	return true;	
}

function CreateTelephoneDetails()
{
	//JR - omiplus24, comment out
	//var TelListNode = XMLContact.CreateActiveTag("CONTACTTELEPHONEDETAILSLIST");

	// Telephone Number 1
	if(IsTelephoneEntered(1))
	{
		CreateTelephoneDetailsListItem(1);
	}
		
	// Telephone Number 2
	if(IsTelephoneEntered(2))
	{
		//XMLContact.ActiveTag = TelListNode;
		//GD	06/02/02    SYS4005 Ensure phone number details aren't nested in XML on a create (for first time)
		XMLContact.ActiveTag = m_ContactDetailsTag;
		CreateTelephoneDetailsListItem(2);
	}
}


function CreateTelephoneDetailsListItem(intIndex)
{
	// NB ActiveTag in XMLContact must have been set correctly
	// to CONTACTTELEPHONEDETAILSLIST before calling this function.
	
	if(intIndex == 1)
	{
		XMLContact.CreateActiveTag("CONTACTTELEPHONEDETAILS");
		XMLContact.CreateTag("USAGE",frmScreen.cboType1.value);		
		XMLContact.CreateTag("COUNTRYCODE",frmScreen.txtCountryCode1.value);		
		XMLContact.CreateTag("AREACODE",frmScreen.txtAreaCode1.value);		
		XMLContact.CreateTag("TELENUMBER",frmScreen.txtTelephoneNumber1.value);		
		XMLContact.CreateTag("EXTENSIONNUMBER",frmScreen.txtWorkExtNo1.value);		
	}
	else if(intIndex == 2)
	{
		XMLContact.CreateActiveTag("CONTACTTELEPHONEDETAILS");
		XMLContact.CreateTag("USAGE",frmScreen.cboType2.value);		
		XMLContact.CreateTag("COUNTRYCODE",frmScreen.txtCountryCode2.value);		
		XMLContact.CreateTag("AREACODE",frmScreen.txtAreaCode2.value);		
		XMLContact.CreateTag("TELENUMBER",frmScreen.txtTelephoneNumber2.value);		
		XMLContact.CreateTag("EXTENSIONNUMBER",frmScreen.txtWorkExtNo2.value);		
	}	
	
}


-->
</script>
</body>
</html>