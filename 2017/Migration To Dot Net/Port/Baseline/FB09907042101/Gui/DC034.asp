<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/*
Workfile:      DC034.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Add/Edit Alias/Association
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MDC		11/11/2002	BMIDS00911	Created
DRC     11/02/2004  BMIDS693    Added Credit Search
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

Prog	Date		AQR			Description
AW		08/05/2006	EP390		Make Other title mandatory
*/ %>

<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>
<body>

<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>

<% /* FORMS */%>
<form id="frmToDC033" method="post" action="DC033.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 30px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<% /*<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<font face="MS Sans Serif" size="1">
			<strong>Alias Association Details</strong>
		</font>
		</span>
	*/ %>

	<span style="TOP: 8px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Customer Name
		<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
			<input id="txtCustomerName" name="CustomerName" maxlength="70" style="WIDTH: 190px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
</div>

<div id="divAlias1" style="TOP: 100px; LEFT: 10px; HEIGHT: 200px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Title
		<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
			<select id="cboTitle1" name="Title1" style="WIDTH: 100px" class="msgCombo">
			</select>
		</span>
	</span>

	<span id="spnOther1" style="TOP: 10px; LEFT: 280px; POSITION: ABSOLUTE" class="msgLabel">
		Other Title
		<span style="TOP: -3px; LEFT: 70px; POSITION: ABSOLUTE">
			<input id="txtOtherTitle1" name="OtherTitle1" maxlength="20" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 36px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Forenames
		<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
			<input id="txtFirstForename1" name="FirstForename1" maxlength="30" style="WIDTH: 150px; POSITION: ABSOLUTE"class="msgTxt">
		</span>

		<span style="TOP: -3px; LEFT: 270px; POSITION: ABSOLUTE">
			<input id="txtSecondForename1" name="SecondForename1" maxlength="30" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>

		<span style="TOP: -3px; LEFT: 430px; POSITION: ABSOLUTE">
			<input id="txtOtherForenames1" name="OtherForenames1" maxlength="30" style="WIDTH: 150px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 62px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Surname
		<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
			<input id="txtSurname1" name="Surname1" maxlength="40" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 88px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Alias Type
		<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
			<select id="cboAliasType1" name="AliasType1" style="WIDTH: 100px" class="msgCombo">
			</select>
		</span>
	</span>

	<span style="TOP: 114px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Date Of Change
		<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
			<input id="txtDateOfChange1" name="DateOfChange1" maxlength="10" style="WIDTH: 100px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>

	<span style="TOP: 140px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Method Of Change
		<span style="TOP: -3px; LEFT: 110px; POSITION: ABSOLUTE">
			<select id="cboMethodOfChange1" name="MethodOfChange1" style="WIDTH: 100px" class="msgCombo">
			</select>
		</span>
	</span>
	<% /*DRC BMIDS693  11/02/2004 */ %>
	<span style="LEFT:4px; POSITION: absolute; TOP: 165px" class="msgLabel">
			Credit Search
			<span style="LEFT: 33px; POSITION: relative; TOP: 2px">
				<input id="optCreditCheckYes" name="CreditCheckIndicator" type="radio" value="1"><label for="optCreditCheckYes" class="msgLabel">Yes</label>
			</span>
			<span style="LEFT: 150px; POSITION: absolute; TOP: 2px" id=SPAN1>
				<input id="optCreditCheckNo" name="CreditCheckIndicator" type="radio" value="0" checked><label for="optCreditCheckNo" class="msgLabel">No</label>
			</span>
	</span>	
	
</div>

</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 400px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" -->
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<!--  #include FILE="attribs/DC034Attribs.asp" -->
<%/* CODE */%>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly	= "";
var m_sCurrentCustomerNumber = "";
var m_sCurrentCustomerVersion = "";
var m_sAlias1SeqNum = "";
var m_sAlias2SeqNum = "";
var m_sAlias3SeqNum = "";
var m_sAliasPersonGuid = "";
var scScreenFunctions;
var m_sReadOnly = null;
var m_blnReadOnly = false;
var m_sXML = null;
var m_sAliasSequenceNumber = "";

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;

function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	//next line replaced by line below as per V7.0.2 of Core - DPF 20/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Add/Edit Alias/Association","DC034",scScreenFunctions);

	RetrieveContextData();

	//This function is contained in the field attributes file (remove if not required)
	SetMasks();

	Validation_Init();
	Initialise();
	scScreenFunctions.SetRadioGroupToReadOnly(frmScreen, "CreditCheckIndicator");
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC033");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly	= scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sCurrentCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","1007");
	m_sCurrentCustomerVersion = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","1");
	m_sXML = scScreenFunctions.GetContextParameter(window,"idXML","");
	//EP578: critical data LH_23/05/2001
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","APP0001");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","2");
}

function onComboChange(ComboId, SpanId, TxtOtherId)
{
	var selIndex = ComboId.selectedIndex;
	if(selIndex != -1 &&
	   scScreenFunctions.IsOptionValidationType(ComboId,selIndex,"O") == true)
	{
		scScreenFunctions.ShowCollection(SpanId);
	}
	else
	{
		scScreenFunctions.HideCollection(SpanId);
		TxtOtherId.value = "";
	}
}

function PopulateCombos()
{
	var XMLTitle = null;
	var XMLMethodOfChange = null;
	var XMLAliasType = null;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("Title", "AliasMethodOfChange", "AliasType");

	if(XML.GetComboLists(document,sGroupList))
	{
		XMLTitle = XML.GetComboListXML("Title");
		XMLMethodOfChange = XML.GetComboListXML("AliasMethodOfChange");
		XMLAliasType = XML.GetComboListXML("AliasType");

		var blnSuccess = true;
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboTitle1,XMLTitle,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboAliasType1,XMLAliasType,true);
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboMethodOfChange1,XMLMethodOfChange,true);

		if(!blnSuccess)
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
	XML = null;
}

function PopulateScreen()
{
	//set the Customer Name
	var sCustomerName = scScreenFunctions.GetContextCustomerName(window, m_sCurrentCustomerNumber);
	frmScreen.txtCustomerName.value = sCustomerName;
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCustomerName");

	if(m_sMetaAction == "Edit")
	{
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.LoadXML(m_sXML);
	
		XML.SelectTag(null, "ALIAS");
		m_sAliasPersonGuid = XML.GetTagText("PERSONGUID");
		m_sAliasSequenceNumber = XML.GetTagText("ALIASSEQUENCENUMBER");

		// Populate Alias fields
		frmScreen.txtFirstForename1.value = XML.GetTagText("FIRSTFORENAME");
		frmScreen.txtSecondForename1.value = XML.GetTagText("SECONDFORENAME");
		frmScreen.txtOtherForenames1.value = XML.GetTagText("OTHERFORENAMES");
		frmScreen.txtSurname1.value = XML.GetTagText("SURNAME");
		frmScreen.txtDateOfChange1.value = XML.GetTagText("DATEOFCHANGE");
		frmScreen.cboTitle1.value = XML.GetTagText("TITLE");
		frmScreen.txtOtherTitle1.value = XML.GetTagText("TITLEOTHER");
		frmScreen.cboMethodOfChange1.value = XML.GetTagText("METHODOFCHANGE");
		frmScreen.cboAliasType1.value = XML.GetTagText("ALIASTYPE");

		//Set the field visibility etc
		onComboChange(frmScreen.cboTitle1,spnOther1,frmScreen.txtOtherTitle1);
	    <% /*DRC BMIDS693  11/02/2004 START*/ %>

		scScreenFunctions.SetRadioGroupValue(frmScreen, "CreditCheckIndicator", XML.GetTagText("CREDITSEARCH"));

		
	}
	else
	scScreenFunctions.SetRadioGroupValue(frmScreen, "CreditCheckIndicator", "0");
	<% /*DRC BMIDS693  11/02/2004 END */ %>
}

function Initialise()
{
	PopulateCombos();
	PopulateScreen();

	if(m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
	else
	{
		frmScreen.cboTitle1.focus();
		// If Association selected - disable fields
		if (frmScreen.cboAliasType1.value =="20")
		{
			frmScreen.txtDateOfChange1.disabled = true;
			frmScreen.cboMethodOfChange1.disabled = true;
		}
	}	
}

function frmScreen.cboTitle1.onchange()
{
	onComboChange(frmScreen.cboTitle1,spnOther1,frmScreen.txtOtherTitle1);
}

function frmScreen.cboAliasType1.onchange() 
{
	if (frmScreen.cboAliasType1.value == "20")
	{
		frmScreen.txtDateOfChange1.disabled = true;
		frmScreen.cboMethodOfChange1.disabled = true;
		frmScreen.txtDateOfChange1.value = "";
		frmScreen.cboMethodOfChange1.value = 0;
	}
	else
	{
		frmScreen.txtDateOfChange1.disabled = false;
		frmScreen.cboMethodOfChange1.disabled = false;
	}
}

function btnSubmit.onclick()
{
	if (m_sReadOnly == true) 
	{
		btnCancel.onclick()
	}	
	else
	{
		<% /* AW EP390 Other title mandatory */ %>
		if(scScreenFunctions.IsValidationType(frmScreen.cboTitle1,"O") && frmScreen.txtOtherTitle1.value == "") 
		{
			alert("Please enter the 'Other' Title or Change Title selection ");
			frmScreen.cboTitle1.focus();
			return false;
		}
		
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
			{
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
				XML.CreateRequestTag(window,null)
				
				//EP578: OPERATION LH_23/05/2001
				XML.SetAttribute("OPERATION","CriticalDataCheck");
		
				XML.CreateActiveTag("ALIASPERSONLIST");

				XML.CreateActiveTag("ALIAS");
				XML.CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersion);
				if(m_sMetaAction == "Edit")
					XML.CreateTag("ALIASSEQUENCENUMBER", m_sAliasSequenceNumber);
				<% /*DRC BMIDS693  11/02/2004 */ %>
				XML.CreateTag("CREDITSEARCH", scScreenFunctions.GetRadioGroupValue(frmScreen,"CreditCheckIndicator"));
				XML.CreateTag("ALIASTYPE", frmScreen.cboAliasType1.value);
				XML.CreateTag("DATEOFCHANGE", frmScreen.txtDateOfChange1.value);
				XML.CreateTag("METHODOFCHANGE", frmScreen.cboMethodOfChange1.value);
				XML.CreateTag("PERSONGUID", m_sAliasPersonGuid);

				XML.CreateActiveTag("PERSON");
				XML.CreateTag("PERSONGUID", m_sAliasPersonGuid);

				XML.CreateTag("FIRSTFORENAME", frmScreen.txtFirstForename1.value);
				XML.CreateTag("OTHERFORENAMES", frmScreen.txtOtherForenames1.value);
				XML.CreateTag("SECONDFORENAME", frmScreen.txtSecondForename1.value);
				XML.CreateTag("SURNAME", frmScreen.txtSurname1.value);
				XML.CreateTag("TITLE", frmScreen.cboTitle1.value);
				XML.CreateTag("TITLEOTHER", frmScreen.txtOtherTitle1.value);

				// EP578: add CRITICALDATACONTEXT LH 23/05/2006
				XML.SelectTag(null,"REQUEST");
				XML.CreateActiveTag("CRITICALDATACONTEXT");
				XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
				XML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
				XML.SetAttribute("SOURCEAPPLICATION","Omiga");
				XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
				XML.SetAttribute("ACTIVITYINSTANCE","1");
				XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
				XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
				XML.SetAttribute("COMPONENT","omCust.CustomerBO");
				XML.SetAttribute("METHOD","SaveAlias");
						
				window.status = "Critical Data Check - please wait";
		
				switch (ScreenRules())
				{
					case 1: // Warning
					case 0: // OK
						XML.RunASP(document,"OmigaTMBO.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
				}
				
				window.status = "";
				
				if(XML.IsResponseOK())
				{
					scScreenFunctions.SetContextParameter(window,"idXML","");
					scScreenFunctions.SetContextParameter(window,"idMetaAction","");
					frmToDC033.submit();
				}
			}
			else
				btnCancel.onclick();
		}
	}
}

function btnCancel.onclick()
{
	//set metaAction
	scScreenFunctions.SetContextParameter(window,"idXML","");
	scScreenFunctions.SetContextParameter(window,"idMetaAction","");
	frmToDC033.submit();
}

-->
</script>
</body>
</html>


