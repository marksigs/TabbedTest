<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/*
Workfile:      DC225.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   New Property Description
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
IW		26/04/2000	Created
BG		17/05/00	SYS0752 Removed Tooltips
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
CL		05/03/01	SYS1920 Read only functionality added
SA		30/05/01	SYS1026 optGRANTAPPLIEDFORINDYes label altered.
					Top position of GrantAmount textbox altered.
SA		30/05/01	SYS1059
JLD		10/12/01	SYS2806 completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		29/05/02	SYS4767 MSMS to Core integration
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
GD		05/07/2002	BMIDS00165 'Buy To Let' functionality added
MV		22/08/2002	BMIDS0355 - amended frmScreen.btnBuyToLet.onclick()- Modified the width and height of the popup window
HMA     17/09/2003  BM0063    - Amend HTML text for radion buttons
MC		21/04/2004	BMIDS517	LetProperty (DC226) dialog size changed.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom 2
Prog    Date           Description
MAH		23/11/2006	   E2CR35 Implemented New NewProperty fields
AShaw	05/02/2007	   EP2_814 - New logic in IsOK method to enforce entry of BTL details.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

%>
<HEAD>
	<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<% /* FORMS */ %>
<form id="frmToDC220" method="post" action="DC220.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC230" method="post" action="DC230.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">

<% /* SG 29/05/02 SYS4767 START */ %>
<div id="divBackground1" style="TOP: 60px; LEFT: 4px; HEIGHT: 130px; WIDTH: 617px; POSITION: ABSOLUTE" class="msgGroup">
<% /* SG 29/05/02 SYS4767 END */ %>
<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>

	<span style="LEFT: 4px; POSITION: absolute; WIDTH: 349px; TOP: 10px" class="msgLabel">
		Self build? 
		<span style="LEFT: 180px; POSITION: absolute">
			<input id="optSELFBUILDINDICATORYes" name="GroupSELFBUILDINDICATOR" type="radio" value="1"><label for="optSELFBUILDINDICATORYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 240px; POSITION: absolute; WIDTH: 50px">
			<input id="optSELFBUILDINDICATORNo" name="GroupSELFBUILDINDICATOR" type="radio" value="0"><label for="optSELFBUILDINDICATORNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
		Currently occupied? 
		<span style="LEFT: 180px; POSITION: absolute">
			<input id="optCURRENTLYOCCUPIEDINDICATORYes" name="GroupCURRENTLYOCCUPIEDINDICATOR" type="radio" value="1"><label for="optCURRENTLYOCCUPIEDINDICATORYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 240px; POSITION: absolute">
			<input id="optCURRENTLYOCCUPIEDINDICATORNo" name="GroupCURRENTLYOCCUPIEDINDICATOR" type="radio" value="0"><label for="optCURRENTLYOCCUPIEDINDICATORNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Applicant to occupy? 
		<span style="LEFT: 180px; POSITION: absolute">
			<input id="optAPPLICANTTOOCCUPYINDICATORYes" name="GroupAPPLICANTTOOCCUPYINDICATOR" type="radio" value="1" ><label for="optAPPLICANTTOOCCUPYINDICATORYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 240px; POSITION: absolute">
			<input id="optAPPLICANTTOOCCUPYINDICATORNo" name="GroupAPPLICANTTOOCCUPYINDICATOR" type="radio" value="0"><label for="optAPPLICANTTOOCCUPYINDICATORNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 324px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Let or part let?
		<span style="LEFT: 173px; POSITION: absolute; WIDTH: 50px">
			<input id="optLETORPARTLETINDICATORYes" name="GroupLETORPARTLETINDICATOR" type="radio" value="1"><label for="optLETORPARTLETINDICATORYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 233px; POSITION: absolute; WIDTH: 50px">
			<input id="optLETORPARTLETINDICATORNo" name="GroupLETORPARTLETINDICATOR" type="radio" value="0"><label for="optLETORPARTLETINDICATORNo" class="msgLabel">No</label> 
		</span> 
	</span>
	<% /* SG 29/05/02 SYS4767 START */ %>
	<!-- SG 04/03/02 SYS4187 -->
	<span style="LEFT: 500px; POSITION: absolute; TOP: 40px">
		<input id="btnBuyToLet" value="Buy To Let" type="button" style="HEIGHT: 20px; WIDTH: 83px" class="msgButton">
	</span>
	<!-- SG 04/03/02 SYS4187 -->

	<span style="LEFT: 324px; POSITION: absolute; TOP: 70px" class="msgLabel">
	<% /* SG 29/05/02 SYS4767 END */ %>
		First occupier?
		<span style="LEFT: 173px; POSITION: absolute; WIDTH: 50px">
			<input id="optFIRSTOCCUPIERINDICATOR Yes" name="GroupFIRSTOCCUPIERINDICATOR" type="radio" value="1"><label for="optFIRSTOCCUPIERINDICATOR Yes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 233px; POSITION: absolute; WIDTH: 50px">
			<input id="optFIRSTOCCUPIERINDICATORNO" name="GroupFIRSTOCCUPIERINDICATOR" type="radio" value="0">
			<label for="optFIRSTOCCUPIERINDICATORNO" class="msgLabel">No</label> 
		</span> 
	</span>

	<% /* SG 29/05/02 SYS4767 START */ %>
	<span style="LEFT: 324px; POSITION: absolute; TOP: 100px" class="msgLabel">
	<% /* SG 29/05/02 SYS4767 END */ %>
		Ex. Local Authority? 
		<span style="LEFT: 173px; POSITION: absolute; WIDTH: 50px">
			<input id="optEXLOCALAUTHORITYINDICATORYes" name="GroupEXLOCALAUTHORITYINDICATOR" type="radio" value="1"><label for="optEXLOCALAUTHORITYINDICATORYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 233px; POSITION: absolute; WIDTH: 50px">
			<input id="optEXLOCALAUTHORITYINDICATORNo" name="GroupEXLOCALAUTHORITYINDICATOR" type="radio" value="0"><label for="optEXLOCALAUTHORITYINDICATORNo" class="msgLabel">No</label> 
		</span> 
	</span>

</div>

<% /* SG 29/05/02 SYS4767 START */ %>
<div id="divBackground2" style="TOP: 200px; LEFT: 4px; HEIGHT: 120Px; WIDTH: 303px; POSITION: ABSOLUTE" class="msgGroup">
<% /* SG 29/05/02 SYS4767 END */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Has the structure or use changed<BR>since any previous advance? 
		<span style="LEFT: 180px; POSITION: absolute; WIDTH: 50px">
			<input id="optSTRUCTUREUSAGECHANGEINDYes" name="GroupSTRUCTUREUSAGECHANGEIND" type="radio" value="1"><label for="optSTRUCTUREUSAGECHANGEINDYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 240px; POSITION: absolute; WIDTH: 50px">
			<input id="optSTRUCTUREUSAGECHANGEINDNo" name="GroupSTRUCTUREUSAGECHANGEIND" type="radio" value="0"><label for="optSTRUCTUREUSAGECHANGEINDNo" class="msgLabel">No</label> 
		</span> 

		<span style="LEFT: 0px; POSITION: absolute; TOP: 35px" class="msgLabel">
		 Details
			<span style="LEFT: 60px; POSITION: absolute; TOP: 5px"><TEXTAREA class=msgTxt id=txtSTRUCTUREUSAGECHANGE rows=4 style="POSITION: absolute; WIDTH: 225px"></TEXTAREA>
			</span>
		</span>

	</span>
</div>

<% /* SG 29/05/02 SYS4767 START */ %>
<div id="divBackground3" style="TOP: 200px; LEFT: 318px; HEIGHT: 120px; WIDTH: 303px; POSITION: ABSOLUTE" class="msgGroup">
<% /* SG 29/05/02 SYS4767 END */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Grant applied for? 
		<span style="LEFT: 180px; POSITION: absolute; WIDTH: 50px">
			<input id="optGRANTAPPLIEDFORINDYes" name="GroupGRANTAPPLIEDFORIND" type="radio" value="1"><label for="optGRANTAPPLIEDFORINDYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 240px; POSITION: absolute; WIDTH: 50px">
			<input id="optGRANTAPPLIEDFORINDNo" name="GroupGRANTAPPLIEDFORIND" type="radio" value="0"><label for="optGRANTAPPLIEDFORINDNo" class="msgLabel">No</label> 
		</span> 

		<span style="LEFT: 0px; POSITION: absolute; TOP: 35px" class="msgLabel">
		 Grant amount £
			<span style="TOP: 0px; LEFT: 90px; POSITION: ABSOLUTE">
				<% /* SG 29/05/02 SYS4767 START */ %>
				<input id="txtGRANTAMOUNT" maxlength="10" style="WIDTH: 80px; POSITION: ABSOLUTE" class="msgTxt">
				<% /* SG 29/05/02 SYS4767 END */ %>
			</span>
		</span>

	</span>
</div>
<% /* SG 29/05/02 SYS4767 START */ %>
<div id="divBackground4" style="TOP: 330px; LEFT: 4px; HEIGHT: 120Px; WIDTH: 303px; POSITION: ABSOLUTE" class="msgGroup">
<% /* SG 29/05/02 SYS4767 END */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Resale restrictions? 
		<span style="LEFT: 180px; POSITION: absolute; WIDTH: 50px">
			<input id="optRESALERESTRICTIONSINDYes" name="GroupRESALERESTRICTIONSIND" type="radio" value="1"><label for="optRESALERESTRICTIONSINDYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 240px; POSITION: absolute; WIDTH: 50px">
			<input id="optRESALERESTRICTIONSINDNo" name="GroupRESALERESTRICTIONSIND" type="radio" value="0"><label for="optRESALERESTRICTIONSINDNo" class="msgLabel">No</label> 
		</span> 

		<span style="LEFT: 0px; POSITION: absolute; TOP: 35px" class="msgLabel">
		 Details
			<span style="LEFT: 60px; POSITION: absolute; TOP: 5px"><TEXTAREA class=msgTxt id=txtRESALERESTRICTIONSDETAILS rows=4 style="POSITION: absolute; WIDTH: 225px"></TEXTAREA>
			</span>
		</span>

	</span>
</div>
<% /* SG 29/05/02 SYS4767 START */ %>
<div id="divBackground5" style="TOP: 330px; LEFT: 318px; HEIGHT: 120px; WIDTH: 303px; POSITION: ABSOLUTE" class="msgGroup">
<% /* SG 29/05/02 SYS4767 END */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Extras included in price? 
		<span style="LEFT: 180px; POSITION: absolute; WIDTH: 50px">
			<input id="optEXTRASINCLUDEDINDYes" name="GroupEXTRASINCLUDEDIND" type="radio" value="1"><label for="optEXTRASINCLUDEDINDYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 240px; POSITION: absolute; WIDTH: 50px">
			<input id="optEXTRASINCLUDEDINDNo" name="GroupEXTRASINCLUDEDIND" type="radio" value="0"><label for="optEXTRASINCLUDEDINDNo" class="msgLabel">No</label> 
		</span> 

		<span style="LEFT: 0px; POSITION: absolute; TOP: 35px" class="msgLabel">
		 Details
			<span style="LEFT: 60px; POSITION: absolute; TOP: 5px"><TEXTAREA class=msgTxt id=txtEXTRASINCLUDEDDETAILS rows=4 style="POSITION: absolute; WIDTH: 225px"></TEXTAREA>
			</span>
		</span>
	</span>	
</div>

</form>

<%/* Main Buttons */ %>
<% /* SG 29/05/02 SYS4767 START */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 490px; WIDTH: 612px">
<% /* SG 29/05/02 SYS4767 END */ %>
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/DC225attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var NewPropertyGeneralXML = null;
var ComboXML = null;
var scScreenFunctions;
var m_blnReadOnly = false;

//SG 29/05/02 SYS4767 START
//SG 06/03/02 SYS4187
var BuyToLetXML = null;
//SG 29/05/02 SYS4767 END

/* EVENTS */

function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		frmToDC220.submit();
	}
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else
			frmToDC230.submit();
	}
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

function frmScreen.txtEXTRASINCLUDEDDETAILS.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtEXTRASINCLUDEDDETAILS", 255, true);
}

function frmScreen.txtRESALERESTRICTIONSDETAILS.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtRESALERESTRICTIONSDETAILS", 255, true);
}

function frmScreen.txtSTRUCTUREUSAGECHANGE.onkeyup()
{	
	scScreenFunctions.RestrictLength(frmScreen, "txtSTRUCTUREUSAGECHANGE", 255, true);
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("New Property Details","DC225",scScreenFunctions);
	RetrieveContextData();
	SetMasks();
	Validation_Init();	
	Initialise();	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC225");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	scScreenFunctions.SetFocusToFirstField(frmScreen);	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}


function ClearFields()
{
	scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupSELFBUILDINDICATOR","0");
	scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupCURRENTLYOCCUPIEDINDICATOR","0");
	scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupAPPLICANTTOOCCUPYINDICATOR","0");
	scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupLETORPARTLETINDICATOR","0");
	scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupFIRSTOCCUPIERINDICATOR","0");
	scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupEXLOCALAUTHORITYINDICATOR","0");
	scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupSTRUCTUREUSAGECHANGEIND","0");
	frmScreen.txtSTRUCTUREUSAGECHANGE.value = "";
	scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupGRANTAPPLIEDFORIND","0");
	frmScreen.txtGRANTAMOUNT.value = "";
	scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupRESALERESTRICTIONSIND","0");
	frmScreen.txtRESALERESTRICTIONSDETAILS.value = "";
	scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupEXTRASINCLUDEDIND","0");
	frmScreen.txtEXTRASINCLUDEDDETAILS.value = "";
}

function CommitChanges()
{
	var bSuccess = true;
	
	if(scScreenFunctions.RestrictLength(frmScreen, "txtEXTRASINCLUDEDDETAILS", 255, true))
		return false;
	if(scScreenFunctions.RestrictLength(frmScreen, "txtRESALERESTRICTIONSDETAILS", 255, true))
		return false;
	if(scScreenFunctions.RestrictLength(frmScreen, "txtSTRUCTUREUSAGECHANGE", 255, true))
		return false;
	
	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
			{
				if (IsOK())
				{
					bSuccess = SaveNewPropertyGeneralDetails();
				}
				else
				{
					bSuccess = false;
				}
			}
		}
		else
			bSuccess = false;
	return(bSuccess);
}


function IsOK()
{
	if ((frmScreen.optGRANTAPPLIEDFORINDYes.checked) && (frmScreen.txtGRANTAMOUNT.value.trim() == ""))
	{
		frmScreen.txtGRANTAMOUNT.focus();
		alert("Please enter value for the Grant Amount!");
		return false;
	}
	
	if ((frmScreen.txtEXTRASINCLUDEDDETAILS.value.trim() == "") && (frmScreen.optEXTRASINCLUDEDINDYes.checked))
	{
		frmScreen.txtEXTRASINCLUDEDDETAILS.focus();
		alert("Please enter details for extras included!");
		return false;
	}
	
	if ((frmScreen.txtRESALERESTRICTIONSDETAILS.value.trim() == "") && (frmScreen.optRESALERESTRICTIONSINDYes.checked))
	{
		frmScreen.txtRESALERESTRICTIONSDETAILS.focus();
		alert("Please enter details for resale restrictions!");
		return false;
	}
	
	if ((frmScreen.txtSTRUCTUREUSAGECHANGE.value.trim() == "") && (frmScreen.optSTRUCTUREUSAGECHANGEINDYes.checked))
	{
		frmScreen.txtSTRUCTUREUSAGECHANGE.focus();
		alert("Please enter details for structure usage!");
		return false;
	}
	
	// EP2_814 - New logic to enforce entry of Rental details.
	if (frmScreen.optLETORPARTLETINDICATORYes.checked) 
	{
		var varTenancy = BuyToLetXML.GetTagText("TENANCYTYPE");
		var varMonthlyRental = BuyToLetXML.GetTagText("MONTHLYRENTALINCOME");
		var varRentalStatus = BuyToLetXML.GetTagText("RENTALINCOMESTATUS");

		if (varTenancy == "" || varMonthlyRental == "" || varMonthlyRental == 0 || varRentalStatus == "")
		{
			frmScreen.btnBuyToLet.focus();
			alert("Please complete buy-to-let details.");
			return false;
		}

	}
	
	return true;
}


function DefaultFields()
{
	ClearFields();
}

function Initialise()
{
	//SG 29/05/02 SYS4767 START
	//SG 04/03/02 SYS4187
	frmScreen.btnBuyToLet.disabled = true;
	//SG 29/05/02 SYS4767 END

	PopulateScreen();

	if (m_sMetaAction == "Add")	DefaultFields();

	frmScreen.optSTRUCTUREUSAGECHANGEINDNo.onclick();
	frmScreen.optGRANTAPPLIEDFORINDNo.onclick();
	frmScreen.optRESALERESTRICTIONSINDNo.onclick();
	frmScreen.optEXTRASINCLUDEDINDNo.onclick();
	
	//SG 29/05/02 SYS4767 START	
	//SG 96/03/02 SYS4187
	frmScreen.optLETORPARTLETINDICATORYes.onclick();
	//SG 29/05/02 SYS4767 END
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}


function PopulateScreen()
{
	NewPropertyGeneralXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	NewPropertyGeneralXML.CreateRequestTag(window,null);
	NewPropertyGeneralXML.CreateActiveTag("NEWPROPERTY");
	NewPropertyGeneralXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	NewPropertyGeneralXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	NewPropertyGeneralXML.RunASP(document,"GetNewPropertyGeneral.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = NewPropertyGeneralXML.CheckResponse(ErrorTypes);

	if(ErrorReturn[1] == ErrorTypes[0])
		<%/* Record not found */%>
		m_sMetaAction = "Add";
	else if (ErrorReturn[0] == true)
		<%/* Record found */%>
	{
		m_sMetaAction = "Edit";

		<%/* Populate the screen with the details held in the XML */%>
		NewPropertyGeneralXML.SelectTag(null, "NEWPROPERTY")
		with (frmScreen)
		{
			scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupSELFBUILDINDICATOR",NewPropertyGeneralXML.GetTagText("SELFBUILDINDICATOR"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupCURRENTLYOCCUPIEDINDICATOR",NewPropertyGeneralXML.GetTagText("CURRENTLYOCCUPIEDINDICATOR"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupAPPLICANTTOOCCUPYINDICATOR",NewPropertyGeneralXML.GetTagText("APPLICANTTOOCCUPYINDICATOR"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupLETORPARTLETINDICATOR",NewPropertyGeneralXML.GetTagText("LETORPARTLETINDICATOR"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupFIRSTOCCUPIERINDICATOR",NewPropertyGeneralXML.GetTagText("FIRSTOCCUPIERINDICATOR"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupSTRUCTUREUSAGECHANGEIND",NewPropertyGeneralXML.GetTagText("STRUCTUREUSAGECHANGEIND"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupEXLOCALAUTHORITYINDICATOR",NewPropertyGeneralXML.GetTagText("EXLOCALAUTHORITYINDICATOR"));
			frmScreen.txtSTRUCTUREUSAGECHANGE.value = NewPropertyGeneralXML.GetTagText("STRUCTUREUSAGECHANGE");
			scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupGRANTAPPLIEDFORIND",NewPropertyGeneralXML.GetTagText("GRANTAPPLIEDFORIND"));
			frmScreen.txtGRANTAMOUNT.value = NewPropertyGeneralXML.GetTagText("GRANTAMOUNT");
			scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupRESALERESTRICTIONSIND",NewPropertyGeneralXML.GetTagText("RESALERESTRICTIONSIND"));
			frmScreen.txtRESALERESTRICTIONSDETAILS.value = NewPropertyGeneralXML.GetTagText("RESALERESTRICTIONSDETAILS");
			scScreenFunctions.SetRadioGroupValue(frmScreen,"GroupEXTRASINCLUDEDIND",NewPropertyGeneralXML.GetTagText("EXTRASINCLUDEDIND"));
			frmScreen.txtEXTRASINCLUDEDDETAILS.value = NewPropertyGeneralXML.GetTagText("EXTRASINCLUDEDDETAILS");
		}
	}

	ErrorTypes = null;
	ErrorReturn = null;
	
	//SG 29/05/02 SYS4767 START
	//SG 06/03/02 SYS4187
	ConstructBuyToLetXML();
	//SG 29/05/02 SYS4767 END
	
}

function RetrieveContextData()
{
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","1325");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");

	//SG 29/05/02 SYS4767 START
	//SG 04/03/02 SYS4187
	m_sCurrency = scScreenFunctions.GetContextParameter(window,"idCurrency","");
	//SG 29/05/02 SYS4767 END

}
function SaveNewPropertyGeneralDetails()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null)
	XML.CreateActiveTag("NEWPROPERTY");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("SELFBUILDINDICATOR",(scScreenFunctions.GetRadioGroupValue(frmScreen,"GroupSELFBUILDINDICATOR")=="1")? "1":"0");
	XML.CreateTag("CURRENTLYOCCUPIEDINDICATOR",(scScreenFunctions.GetRadioGroupValue(frmScreen,"GroupCURRENTLYOCCUPIEDINDICATOR")=="1")? "1":"0");
	XML.CreateTag("APPLICANTTOOCCUPYINDICATOR",(scScreenFunctions.GetRadioGroupValue(frmScreen,"GroupAPPLICANTTOOCCUPYINDICATOR")=="1")? "1":"0");
	XML.CreateTag("LETORPARTLETINDICATOR",(scScreenFunctions.GetRadioGroupValue(frmScreen,"GroupLETORPARTLETINDICATOR")=="1")? "1":"0");
	XML.CreateTag("FIRSTOCCUPIERINDICATOR",(scScreenFunctions.GetRadioGroupValue(frmScreen,"GroupFIRSTOCCUPIERINDICATOR")=="1")? "1":"0");
	XML.CreateTag("EXLOCALAUTHORITYINDICATOR",(scScreenFunctions.GetRadioGroupValue(frmScreen,"GroupEXLOCALAUTHORITYINDICATOR")=="1")? "1":"0");
	XML.CreateTag("STRUCTUREUSAGECHANGEIND",(scScreenFunctions.GetRadioGroupValue(frmScreen,"GroupSTRUCTUREUSAGECHANGEIND")=="1")? "1":"0");
	XML.CreateTag("STRUCTUREUSAGECHANGE", frmScreen.txtSTRUCTUREUSAGECHANGE.value);
	XML.CreateTag("GRANTAPPLIEDFORIND",(scScreenFunctions.GetRadioGroupValue(frmScreen,"GroupGRANTAPPLIEDFORIND")=="1")? "1":"0");
	XML.CreateTag("GRANTAMOUNT", frmScreen.txtGRANTAMOUNT.value);
	XML.CreateTag("RESALERESTRICTIONSIND",(scScreenFunctions.GetRadioGroupValue(frmScreen,"GroupRESALERESTRICTIONSIND")=="1")? "1":"0");
	XML.CreateTag("RESALERESTRICTIONSDETAILS", frmScreen.txtRESALERESTRICTIONSDETAILS.value);
	XML.CreateTag("EXTRASINCLUDEDIND",(scScreenFunctions.GetRadioGroupValue(frmScreen,"GroupEXTRASINCLUDEDIND")=="1")? "1":"0");
	XML.CreateTag("EXTRASINCLUDEDDETAILS", frmScreen.txtEXTRASINCLUDEDDETAILS.value);

	//SG 29/05/02 SYS4767 START
	//SG 06/03/02 SYS4187
	if (frmScreen.optLETORPARTLETINDICATORNo.checked)
	{
		XML.CreateTag("NUMBEROFOCCUPANTS", "");
		XML.CreateTag("TENANCYTYPE", "");
		XML.CreateTag("MONTHLYRENTALINCOME", "");
		XML.CreateTag("RENTALINCOMESTATUS", "");
		<%//GD BMIDS00165 start%>
		XML.CreateTag("EXCESSMONTHLYRENTALINCOME", "");
		<%//GD BMIDS00165 end%>
		XML.CreateTag("OCCUPYPROPERTY", "");<%/*MAH 23/11/2006*/%>
	}
	else
	{
		BuyToLetXML.SelectTag(null,"NEWPROPERTY");
		XML.CreateTag("NUMBEROFOCCUPANTS", BuyToLetXML.GetTagText("NUMBEROFOCCUPANTS"));
		XML.CreateTag("TENANCYTYPE", BuyToLetXML.GetTagText("TENANCYTYPE"));
		XML.CreateTag("MONTHLYRENTALINCOME", BuyToLetXML.GetTagText("MONTHLYRENTALINCOME"));
		XML.CreateTag("RENTALINCOMESTATUS", BuyToLetXML.GetTagText("RENTALINCOMESTATUS"));
		<%//GD BMIDS00165 start%>
		XML.CreateTag("EXCESSMONTHLYRENTALINCOME", BuyToLetXML.GetTagText("EXCESSMONTHLYRENTALINCOME"));
		<%//GD BMIDS00165 end%>
		XML.CreateTag("OCCUPYPROPERTY", BuyToLetXML.GetTagText("OCCUPYPROPERTY"));<%/*MAH 23/11/2006*/%>
		
	}
	//SG 06/03/02 SYS4187
	//SG 29/05/02 SYS4767 END

	if (m_sMetaAction == "Add")
		// 		XML.RunASP(document,"CreateNewPropertyGeneral.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateNewPropertyGeneral.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	else
		
		// 		XML.RunASP(document,"UpdateNewPropertyGeneral.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateNewPropertyGeneral.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}


	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}

function frmScreen.optSTRUCTUREUSAGECHANGEINDNo.onclick()
{
	if(frmScreen.optSTRUCTUREUSAGECHANGEINDNo.checked)
		scScreenFunctions.SetFieldState(frmScreen, "txtSTRUCTUREUSAGECHANGE", "D")
}

function frmScreen.optSTRUCTUREUSAGECHANGEINDYes.onclick()
{
	if(frmScreen.optSTRUCTUREUSAGECHANGEINDYes.checked)
		scScreenFunctions.SetFieldState(frmScreen, "txtSTRUCTUREUSAGECHANGE", "W")
}

function frmScreen.optGRANTAPPLIEDFORINDNo.onclick()
{
	if(frmScreen.optGRANTAPPLIEDFORINDNo.checked)
		scScreenFunctions.SetFieldState(frmScreen, "txtGRANTAMOUNT", "D")
}

function frmScreen.optGRANTAPPLIEDFORINDYes.onclick()
{
	if(frmScreen.optGRANTAPPLIEDFORINDYes.checked)
		scScreenFunctions.SetFieldState(frmScreen, "txtGRANTAMOUNT", "W")
}

function frmScreen.optRESALERESTRICTIONSINDNo.onclick()
{
	if(frmScreen.optRESALERESTRICTIONSINDNo.checked)
		scScreenFunctions.SetFieldState(frmScreen, "txtRESALERESTRICTIONSDETAILS", "D")
}

function frmScreen.optRESALERESTRICTIONSINDYes.onclick()
{
	if(frmScreen.optRESALERESTRICTIONSINDYes.checked)
		scScreenFunctions.SetFieldState(frmScreen, "txtRESALERESTRICTIONSDETAILS", "W")
}

function frmScreen.optEXTRASINCLUDEDINDNo.onclick()
{
	if(frmScreen.optEXTRASINCLUDEDINDNo.checked)
		scScreenFunctions.SetFieldState(frmScreen, "txtEXTRASINCLUDEDDETAILS", "D")
}

function frmScreen.optEXTRASINCLUDEDINDYes.onclick()
{
	if(frmScreen.optEXTRASINCLUDEDINDYes.checked)
		scScreenFunctions.SetFieldState(frmScreen, "txtEXTRASINCLUDEDDETAILS", "W")
}

function frmScreen.optLETORPARTLETINDICATORYes.onclick()	//SG 04/03/02 SYS4187
{	
	if(frmScreen.optLETORPARTLETINDICATORYes.checked)
		frmScreen.btnBuyToLet.disabled = false;
}
//SG 29/05/02 SYS4767 Function added
function frmScreen.optLETORPARTLETINDICATORNo.onclick()		//SG 04/03/02 SYS4187
{
	if(frmScreen.optLETORPARTLETINDICATORNo.checked)
		frmScreen.btnBuyToLet.disabled = true;
}
//SG 29/05/02 SYS4767 Function added
function frmScreen.btnBuyToLet.onclick()					//SG 04/03/02 SYS4187
{

	var sReturn = null;
	var ArrayArguments = new Array();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ArrayArguments[0] = m_sReadOnly;
	ArrayArguments[1] = BuyToLetXML;
	ArrayArguments[2] = m_sCurrency;
	ArrayArguments[3] = m_sMetaAction;
	//GD BMIDS00165 start
	ArrayArguments[4] = BuildXMLForCalcDC226Button()
	ArrayArguments[5] = XML.CreateRequestAttributeArray(window);
	//GD BMIDS00165 end
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "DC226.asp", ArrayArguments, 435, 328);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);	
	}
}
//SG 29/05/02 SYS4767 Function added
function ConstructBuyToLetXML()					//SG 04/03/02 SYS4187
{	
	//Contruct XML to pass to popup.
	//Do this because NewPropertyGeneralXML may contain Record Not Found error

	BuyToLetXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	BuyToLetXML.CreateActiveTag("NEWPROPERTY");		

	if (m_sMetaAction == "Add")
	{	
		//Populate with blank data
		BuyToLetXML.CreateTag("NUMBEROFOCCUPANTS", "");
		BuyToLetXML.CreateTag("TENANCYTYPE", "");
		BuyToLetXML.CreateTag("MONTHLYRENTALINCOME", "");
		BuyToLetXML.CreateTag("RENTALINCOMESTATUS", "");
		<%//GD BMIDS00165 start%>
		BuyToLetXML.CreateTag("EXCESSMONTHLYRENTALINCOME", "");
		<%//GD BMIDS00165 end%>
		BuyToLetXML.CreateTag("OCCUPYPROPERTY",""); <%/*MAH 23/11/2006 E2CR35*/%>
	}
	else
	{
		//Populate from NewPropertyGeneralXML
		NewPropertyGeneralXML.SelectTag(null,"NEWPROPERTY");
		BuyToLetXML.CreateTag("NUMBEROFOCCUPANTS", NewPropertyGeneralXML.GetTagText("NUMBEROFOCCUPANTS"));
		BuyToLetXML.CreateTag("TENANCYTYPE", NewPropertyGeneralXML.GetTagText("TENANCYTYPE"));
		BuyToLetXML.CreateTag("MONTHLYRENTALINCOME", NewPropertyGeneralXML.GetTagText("MONTHLYRENTALINCOME"));
		BuyToLetXML.CreateTag("RENTALINCOMESTATUS", NewPropertyGeneralXML.GetTagText("RENTALINCOMESTATUS"));
		<%//GD BMIDS00165 start%>
		BuyToLetXML.CreateTag("EXCESSMONTHLYRENTALINCOME", NewPropertyGeneralXML.GetTagText("EXCESSMONTHLYRENTALINCOME"));
		<%//GD BMIDS00165 end%>	
		BuyToLetXML.CreateTag("OCCUPYPROPERTY",NewPropertyGeneralXML.GetTagText("OCCUPYPROPERTY")); <%/*MAH 23/11/2006 E2CR35*/%>
		}
}
//GD BMIDS00165 New Function BuildXMLForCalcDC226Button
function BuildXMLForCalcDC226Button()
{
	//Build the XML here for dc226, rather than passing, potentially, many arguments to the popup.
	var CalcXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var CustNum, CustVerNum, iCount;
	CalcXML.CreateRequestTag(window,null);
	CalcXML.CreateActiveTag("NEWPROPERTY");
	CalcXML.CreateActiveTag("RENTALDETAILS");
	CalcXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	CalcXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	//Build up customer list
	CalcXML.SelectTag(null, "RENTALDETAILS");
	CalcXML.CreateActiveTag("CUSTOMERLIST");
	for(iCount = 1;iCount <= 5;iCount++)
	{
		CustNum = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
		CustVerNum = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);
		if (CustNum != "")
		{
			CalcXML.CreateActiveTag("CUSTOMER");
			CalcXML.CreateTag("CUSTOMERNUMBER", CustNum);
			CalcXML.CreateTag("CUSTOMERVERSIONNUMBER", CustVerNum);
			CalcXML.SelectTag(null, "CUSTOMERLIST");
		}
	}
	//Return as string, so object stays intact
	return(CalcXML.XMLDocument.xml)
}

-->
</script>
</body>
</html>



