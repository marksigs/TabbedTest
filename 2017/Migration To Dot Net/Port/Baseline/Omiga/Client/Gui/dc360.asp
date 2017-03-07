<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<%/*
Workfile:      dc360.asp
Copyright:     Copyright © 2005 Marlborough Stirling

Description:   Wrap Up Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
HM		05/09/2005	MAR29	Original
MahaT	12/10/2005	MAR57	WP03 Changes
MahaT	17/10/2005	MAR203	Call Date & Time validation
PSC		28/10/2005	MAR300  Comment out call to UpdateCRSCustomer
MV		31/10/2005	MAR257	Amended GetTelephoneDetails() and ValidateAddHocTaskDateTime()
PJO     09/11/2005  MAR321  Make validation of passwords non case-sensitive
Maha T	18/11/2005	MAR323	1) Change mask setings for Country, Area & Telephone codes to accept only numbers. 
							2) Ammened GetTelephoneDetails
HMA     28/11/2005  MAR324  Use AllowExitFromOmiga context variable.
JD		17/02/2006	MAR1287 Password fields are not mandatory. Changed ValidatePassword.
IK		02/03/2006	MAR1349 UpdateCRSustomer on marketing button change
PE		07/02/2006	MAR1366	Warns that data needs to be completed but then closes on Ok
PE		07/02/2006	MAR1374	Wrap up call customer doesn't say who to call
MHeys	18/04/2006	MAR1587	Fixed Password validation error
MHeys	18/04/2006	MAR1587	Call WrapUp when ExitOmigaFromWrapUp set and when window closed by browser "X" push
TK		06/05/2006	MAR1384 Added PreferredContact validation
							Associate in ClearQuest.
PSC		08/05/2006	MAR1384 Correct Validation
PSC		08/05/2006	MAR1384 Correct Validation to check for email
JD		14/07/2006	MAR1691 set boolean flags correctly.
JD		25/07/2006	MAR1922 change to always do an UpdateCRS call on OK
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
	<HEAD>
		<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
		<title>Wrap Up <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
	</HEAD>
	<body>
<script src="validation.js" language="JScript"></script>
<% /* XML Functions Object - see comments within scXMLFunctions.asp for details */ %>
<OBJECT id=scClientFunctions style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp VIEWASTEXT></OBJECT>
<OBJECT id=scTable style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet height=1 data=scTable.htm VIEWASTEXT></OBJECT>
		<span style="DISPLAY: none; LEFT: 0px; TOP: 0px">
<OBJECT id=scScrollTable style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet data=scTableListScroll.asp VIEWASTEXT></OBJECT>
		</span>
		<% /* Scriptlets - remove any which are not required */ %>
		<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
		<FORM id="frmScreen" style="VISIBILITY: hidden" mark validate="onchange" year4>
			<TABLE cellSpacing="5" cellPadding="0" width="607" border="0">
				<tr>
					<td class="msgLabelHead" colspan="3">Wrap Up</td>
				</tr>
				<TR>
					<TD colSpan="3">
						<TEXTAREA class="msgReadOnly" id="txtAgreeNextSteps" rows=3 style="WIDTH: 605px;"></TEXTAREA>
					</TD>
				</TR>
				<TR>
					<TD width="155"></TD>
					<TD></TD>
					<TD width="275"></TD>
				</TR>
				<TR>
					<TD class="msgLabel">Sales / Lead Status
					</TD>
					<TD><SELECT class="msgCombo" id="cboSalesLeadStatus" style="WIDTH: 180px" name="SalesLeadStatus"></SELECT></TD>
					<TD class="msgLabelHead">
						Marketing Information Programme Opt In
					</TD>
				</TR>
				<TR>
					<TD><label id="lblPassword" class="msgLabel">Enter Password</label></TD>
					<TD><INPUT class="msgTxt" id="txtPassword" style="WIDTH: 100px" type="PASSWORD" name="Password" maxlength=8>
					</TD>
					<TD rowspan="3">
						<TABLE class="msgTable" id="tblMarketingInfo" cellSpacing="0" cellPadding="0" width="100%"
							 border="0">
							<TR id="row01">
								<TD class="TableTopLeft" width="60%">&nbsp;</TD>
								<TD class="TableTopRight">
									<% /* Start: MAR57  - Maha T */ %>
									<span id="MailShotRequiredGroup1">
										<INPUT id="optMailShotRequired1Yes" style="LEFT: 1px; TOP: 0px" type="radio" value="1" name="MailShotRequired1" readonly>
										<LABEL class="msgLabel" for="optMailShotRequired1Yes">Yes </LABEL>
										<INPUT id="optMailShotRequired1No" style="LEFT: 1px; TOP: 0px" type="radio" value="0" name="MailShotRequired1" readonly>
										<LABEL class="msgLabel" for="optMailShotRequired1No">No </LABEL>
									</span>
									<% /* End: MAR57 */ %>
								</TD>
							</TR>
							<TR id="row02">
								<TD class="TableLeft">&nbsp;</TD>
								<TD class="TableRight">
									<% /* Start: MAR57  - Maha T */ %>
									<span id="MailShotRequiredGroup2">
										<INPUT id="optMailShotRequired2Yes" style="LEFT: 1px; TOP: 0px" type="radio" value="1" name="MailShotRequired2" readonly>
										<LABEL class="msgLabel" for="optMailShotRequired2Yes">Yes </LABEL>
										<INPUT id="optMailShotRequired2No" style="LEFT: 1px; TOP: 0px" type="radio" value="0" name="MailShotRequired2" readonly>
										<LABEL class="msgLabel" for="optMailShotRequired2No">No </LABEL>
									</span>
									<% /* End: MAR57 */ %>
								</TD>
							</TR>
							<TR id="row03">
								<TD class="TableLeft">&nbsp;</TD>
								<TD class="TableRight">
									<% /* Start: MAR57  - Maha T */ %>
									<span id="MailShotRequiredGroup3">
										<INPUT id="optMailShotRequired3Yes" style="LEFT: 1px; TOP: 0px" type="radio" value="1" name="MailShotRequired3" readonly>
										<LABEL class="msgLabel" for="optMailShotRequired3Yes">Yes </LABEL>
										<INPUT id="optMailShotRequired3No" style="LEFT: 1px; TOP: 0px" type="radio" value="0" name="MailShotRequired3" readonly>
										<LABEL class="msgLabel" for="optMailShotRequired3No">No </LABEL>
									</span>
									<% /* End: MAR57 */ %>
								</TD>
							</TR>
							<TR id="row04">
								<TD class="TableBottomLeft">&nbsp;</TD>
								<TD class="TableBottomRight">
									<% /* Start: MAR57  - Maha T */ %>
									<span id="MailShotRequiredGroup4">
										<INPUT id="optMailShotRequired4Yes" style="LEFT: 1px; TOP: 0px" type="radio" value="1" name="MailShotRequired4" readonly>
										<LABEL class="msgLabel" for="optMailShotRequired4Yes">Yes </LABEL>
										<INPUT id="optMailShotRequired4No" style="LEFT: 1px; TOP: 0px" type="radio" value="0" name="MailShotRequired4" readonly>
										<LABEL class="msgLabel" for="optMailShotRequired4No">No </LABEL>
									</span>
									<% /* End: MAR57 */ %>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="msgLabel"><label id="lblPasswordConfirm">Confirm Password</label></TD>
					<TD><INPUT class="msgTxt" id="txtPasswordConfirm" style="WIDTH: 100px" type="PASSWORD" name="PasswordConfirm" maxlength=8>
					</TD>
				</TR>
				<TR>
					<TD class="msgLabel">Schedule a call</TD>
					<TD class="msgLabel">
						<INPUT id="optScheduleACallYes" style="LEFT: 1px; TOP: 0px" type="radio" value="1" name="ScheduleACall">
						<LABEL class="msgLabel" for="optScheduleACallYes">Yes </LABEL>
						<INPUT id="optScheduleACallNo" style="LEFT: 1px; TOP: 0px" type="radio" value="0" name="ScheduleACall">
						<LABEL class="msgLabel" for="optScheduleACallNo">No </LABEL>
					</TD>
				</TR>
				<TR>
					<TD class="msgLabel">Call Date</TD>
					<TD colspan="2"><INPUT class="msgTxt" id="txtCallDate" style="WIDTH: 100px" maxLength="10" name="CallDate"></TD> 
					
				</TR>
				<TR>
					<TD class="msgLabel">Call Time</TD>
					<TD colspan="2"><INPUT class="msgTxt" id="txtCallTime" style="WIDTH: 100px" maxLength="10" name="CallTime"></TD>
				</TR>
				<TR>
					<TD class="msgLabel">Or</TD>
					<TD colspan="2"><SELECT class="msgCombo" id="cboCallTime" style="WIDTH: 100px" name="OrField"></SELECT>
					</TD>
				</TR>
				<TR>
					<TD class="msgLabel">Who to contact</TD>
					<TD colspan="2"><SELECT class="msgCombo" id="cboWhoToContact" style="WIDTH: 180px" name="WhoToContact"></SELECT>
					</TD>
				</TR>
				<TR class="msgGroup">
					<TD class="msgLabelWait" colSpan="3">
						<DIV class="msgGroup" id="DIV1" style="LEFT: 4px; WIDTH: 604px; POSITION: relative; HEIGHT: 170px">
							<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabelHead">	Contact Details	</span>
							<span style="LEFT: 4px; POSITION: absolute; TOP: 24px" class="msgLabel"> Type 
								<span style="LEFT: 90px; POSITION: absolute; TOP: 0px" class="msgLabel"> Country 
								</span> 
								<span style="LEFT: 90px; POSITION: absolute; TOP: 10px" class="msgLabel"> Code 
								</span> 
								<span style="LEFT: 150px; POSITION: absolute; TOP: 0px" class="msgLabel"> Area Code 
								</span> 
								<span style="LEFT: 210px; POSITION: absolute; TOP: 0px" class="msgLabel"> Telephone Number 
								</span> 
								<span style="LEFT: 305px; POSITION: absolute; TOP: 0px" class="msgLabel"> &nbsp;Work Extension 
								</span> 
								<span style="LEFT: 395px; POSITION: absolute; TOP: 0px" class="msgLabel"> Contact Time 
								</span> 
								<span style="LEFT: 545px; POSITION: absolute; TOP: -14px" class="msgLabel"> Preferred<BR>
								Contact ? 
								</span> 
							</span>
							<span style="LEFT: 4px; POSITION: absolute; TOP: 48px">
								<select name="select" class="msgCombo" id="cboType1" style="WIDTH: 80px">
								</select>
								<span style="LEFT: 90px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtCountryCode1" style="WIDTH: 50px; POSITION: absolute"
										 maxlength="3">
								</span>
								<span style="LEFT: 148px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtAreaCode1" style="WIDTH: 50px; POSITION: absolute"
										 maxlength="6">
								</span>
								<span style="LEFT: 208px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtTelNumber1" style="WIDTH: 90px; POSITION: absolute"
										 maxlength="30">
								</span>
								<span style="LEFT: 305px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtExtensionNumber1" style="WIDTH: 80px; POSITION: absolute"
										 maxlength="8">
								</span>
								<span style="LEFT: 395px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtTime1" style="WIDTH: 140px; POSITION: absolute"
										 maxlength="30">
								</span>
								<span style="LEFT: 558px; POSITION: absolute; TOP: 0px">
									<input id="optPREFERREDCONTACT1" name="RadioGroup" type="radio" value="1">
								</span>
							</span>
							<span style="LEFT: 4px; POSITION: absolute; TOP: 72px">
								<select name="select" class="msgCombo" id="cboType2" style="WIDTH: 80px">
								</select>
								<span style="LEFT: 90px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtCountryCode2" style="WIDTH: 50px; POSITION: absolute"
										 maxlength="3">
								</span>
								<span style="LEFT: 148px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtAreaCode2" style="WIDTH: 50px; POSITION: absolute"
										 maxlength="6">
								</span>
								<span style="LEFT: 208px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtTelNumber2" style="WIDTH: 90px; POSITION: absolute"
										 maxlength="30">
								</span>
								<span style="LEFT: 305px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtExtensionNumber2" style="WIDTH: 80px; POSITION: absolute"
										 maxlength="8">
								</span>
								<span style="LEFT: 395px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtTime2" style="WIDTH: 140px; POSITION: absolute"
										 maxlength="30">
								</span>
								<span style="LEFT: 558px; POSITION: absolute; TOP: 0px">
									<input id="optPREFERREDCONTACT2" name="RadioGroup" type="radio" value="1">
								</span>
							</span>
							<span style="LEFT: 4px; POSITION: absolute; TOP: 96px">
								<select name="select" class="msgCombo" id="cboType3" style="WIDTH: 80px">
								</select>
								<span style="LEFT: 90px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtCountryCode3" style="WIDTH: 50px; POSITION: absolute"
										 maxlength="3">
								</span>
								<span style="LEFT: 148px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtAreaCode3" style="WIDTH: 50px; POSITION: absolute"
										 maxlength="6">
								</span>
								<span style="LEFT: 208px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtTelNumber3" style="WIDTH: 90px; POSITION: absolute"
										 maxlength="30">
								</span>
								<span style="LEFT: 305px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtExtensionNumber3" style="WIDTH: 80px; POSITION: absolute"
										 maxlength="8">
								</span>
								<span style="LEFT: 395px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtTime3" style="WIDTH: 140px; POSITION: absolute"
										 maxlength="30">
								</span>
								<span style="LEFT: 558px; POSITION: absolute; TOP: 0px">
									<input id="optPREFERREDCONTACT3" name="RadioGroup" type="radio" value="1">
								</span>
							</span>
							<span style="LEFT: 4px; POSITION: absolute; TOP: 120px">
								<select name="select" class="msgCombo" id="cboType4" style="WIDTH: 80px">
								</select>
								<span style="LEFT: 90px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtCountryCode4" style="WIDTH: 50px; POSITION: absolute"
										 maxlength="3">
								</span>
								<span style="LEFT: 148px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtAreaCode4" style="WIDTH: 50px; POSITION: absolute"
										 maxlength="6">
								</span>
								<span style="LEFT: 208px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtTelNumber4" style="WIDTH: 90px; POSITION: absolute"
										 maxlength="30">
								</span>
								<span style="LEFT: 305px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtExtensionNumber4" style="WIDTH: 80px; POSITION: absolute"
										 maxlength="8">
								</span>
								<span style="LEFT: 395px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtTime4" style="WIDTH: 140px; POSITION: absolute"
										 maxlength="30">
								</span>
								<span style="LEFT: 558px; POSITION: absolute; TOP: 0px">
									<input id="optPREFERREDCONTACT4" name="RadioGroup" type="radio" value="1">
								</span>
							</span>
							<span style="LEFT: 4px; POSITION: absolute; TOP: 148px" class="msgLabel"> E-Mail Address 
								<span style="LEFT: 90px; POSITION: absolute; TOP: 0px">
									<input name="Input" class="msgTxt" id="txtContactEMailAddress" style="WIDTH: 445px; POSITION: absolute"
										 maxlength="100">
								</span> 
								<span style="LEFT: 558px; POSITION: absolute; TOP: 0px">
									<input id="optPREFERREDCONTACT5" name="RadioGroup" type="radio" value="1">
								</span> 
							</span>
						</DIV>
					</TD>
				</TR>
			</TABLE>

		</FORM>
<%/* Main Buttons */ %>
<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 500px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" -->
</div>				

<% /* File containing field attributes */ %>
<!-- #include FILE="attribs/dc360attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">

var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_blnReadOnly = false;
var m_sReadOnly = "";
var m_sUnitID = null;
var m_sUserID = null;
var m_XMLCust = null;
var m_sAllowExitFromWrap = null;   // MAR324

var m_sCustomerName = new Array();
var m_sCustomerNumber = new Array();
var m_sCustomerVersion = new Array();
var m_nMaxCustomers = 0;
var m_nMinPwdLength = 6;
<% /* Maha T  MAR57 */ %>
var m_nMaxPwdLength = 8;

//Marketing Info table total rows 
var m_iTableLength = 5;
var m_sMarketingInfo = new Array();
var m_bMarketingEnabled = new Array();
var m_nSeqNo = new Array();
<% /* Start: MAR57 - Maha T*/ %>
var m_sProspectPassword = new Array();
var m_bProspectPasswordTaken = false;
<% /* End: MAR57 */ %>

var scScreenFunctions;
var m_sRequestAttribs = "";

var m_BaseNonPopupWindow = null;
var m_sArgArray = null;
var m_bWrapUpPasswordTaken = false;
var m_sSalesLeadStatus = "";

var m_bLaunchAdHocTask = false;
var m_sAddHocTaskNote = "";
var m_sAddHocTaskDateTime = "";

var scClientScreenFunctions;

// used to manipulate Contact details
var sType =			new Array("cboType1",				"cboType2",				"cboType3",				"cboType4");
var	sCountryCode =	new Array("txtCountryCode1",		"txtCountryCode2",		"txtCountryCode3",		"txtCountryCode4");
var	sAreaCode =		new Array("txtAreaCode1",			"txtAreaCode2",			"txtAreaCode3",			"txtAreaCode4");
var	sNumber	=		new Array("txtTelNumber1",			"txtTelNumber2",		"txtTelNumber3",		"txtTelNumber4");
var	sExtension =	new Array("txtExtensionNumber1",	"txtExtensionNumber2",	"txtExtensionNumber3",	"txtExtensionNumber4");
var sTime =			new Array("txtTime1",				"txtTime2",				"txtTime3",				"txtTime4");
var sOption	=		new Array("optPREFERREDCONTACT1",	"optPREFERREDCONTACT2",	"optPREFERREDCONTACT3", "optPREFERREDCONTACT4");

var m_blnMailShotUpdated  = false;		<% /* PSC 27/10/2005 MAR300 */%>

var m_blnSubmitProcessed = false;		<% /* MAR1587 M Heys 18/04/2006 */%>

function window.onload()
{
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	SetMasks();
	Validation_Init();
	
	<%	// Make the required buttons available on the bottom of the screen
		// (see msgButtons.asp for details)
	%>	
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];	
	m_sArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
							
	RetrieveContextData();
	FindApplicationData();
	GetComboLists();
	PopulateCustomers();
	PopulateMarketingInfo();
	//fill up customer
	frmScreen.cboWhoToContact.onchange();
	SetAllFieldsAttributes();

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(m_BaseNonPopupWindow, frmScreen, "DC360");
	if (m_blnReadOnly == true) m_sReadOnly = "1";

	// set focus on SalesLeadStatus combo 
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	window.returnValue = null;
	ClientPopulateScreen();		//attribs func
}

function RetrieveContextData()
{
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idApplicationNumber",null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idApplicationFactFindNumber",null);  
	m_sReadOnly = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idReadOnly", "0");
	m_sUnitID = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idUnitID",null)
	m_sUserID = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idUserID",null)	
	m_sRequestAttribs = m_sArgArray[2];
	m_sAllowExitFromWrap = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idAllowExitFromWrap",null);   // MAR324
}

function FindApplicationData()
{
	var blnSuccess = false;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	XML.CreateRequestTag(m_BaseNonPopupWindow,"SEARCH");
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.RunASP(document,"GetApplicationData.asp");

	if (XML.IsResponseOK()) 
	{
		m_sSalesLeadStatus = XML.GetTagText("SALESLEADSTATUS");
		m_bWrapUpPasswordTaken = XML.GetTagInt("WRAPUPPASSWORDTAKEN"); //MAR1691
		blnSuccess = true;
	}
	return blnSuccess;
}

function GetComboLists()
{			  
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var XMLTelephoneUsage	= null;

	var sGroups = new Array("SalesLeadStatus","CallTime","TelephoneUsage"); 
	if(XML.GetComboLists(document,sGroups))
	{
		XML.PopulateCombo(document,frmScreen.cboSalesLeadStatus,"SalesLeadStatus",true);
		XML.PopulateCombo(document,frmScreen.cboCallTime,"CallTime",true);
		// telephone usage
		XMLTelephoneUsage = XML.GetComboListXML("TelephoneUsage");
		XML.PopulateComboFromXML(document,frmScreen.cboType1,XMLTelephoneUsage,true);
		XML.PopulateComboFromXML(document,frmScreen.cboType2,XMLTelephoneUsage,true);
		XML.PopulateComboFromXML(document,frmScreen.cboType3,XMLTelephoneUsage,true);
		XML.PopulateComboFromXML(document,frmScreen.cboType4,XMLTelephoneUsage,true);
	}
	<% /* set Sales Lead status to be from Application Fact Find*/ %>
	if (m_sSalesLeadStatus != "")
			frmScreen.cboSalesLeadStatus.value = m_sSalesLeadStatus;
	
	<% /* populate Who To Contact from the customers in context. Add a <SELECT> option */ %>	
	var TagOPTION	= document.createElement("OPTION");
	TagOPTION.value	= "";
	TagOPTION.text	= "<SELECT>";
	
	frmScreen.cboWhoToContact.add(TagOPTION);
	m_nMaxCustomers = 0;
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerName" + nLoop);
		m_sCustomerName[nLoop] = sCustomerName
		
		var sCustomerNumber = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerNumber" + nLoop);
		m_sCustomerNumber[nLoop] = sCustomerNumber
		
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerVersionNumber" + nLoop);
		m_sCustomerVersion[nLoop] = sCustomerVersionNumber
		
		
		if(sCustomerName != "" && sCustomerNumber != "")
		{
			m_nMaxCustomers++;
			TagOPTION		= document.createElement("OPTION");
			TagOPTION.value	= sCustomerNumber;
			TagOPTION.text	= sCustomerName;
			TagOPTION.setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
			frmScreen.cboWhoToContact.add(TagOPTION);
		}
	}
	<%	/* Default to SELECT or the only option if there is only one */ %>	
	if(m_nMaxCustomers == 1) frmScreen.cboWhoToContact.selectedIndex = 1;
	else frmScreen.cboWhoToContact.selectedIndex = 0;
	
	XML = null;
	XMLTelephoneUsage = null;
}

function SetAllFieldsAttributes()
{
	scScreenFunctions.ShowCollection(frmScreen);
	
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtAgreeNextSteps");
	
	//set checkboxes to readonly in MarketingInfo table
	for (var iLoop = 1; iLoop < m_nMaxCustomers+1 ; iLoop++)
	{
		if (m_bMarketingEnabled[iLoop] == false)
		{
			<% /* Start: MAR57  - Maha T */ %>
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"MailShotRequiredGroup" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"optMailShotRequired" + iLoop + "Yes");
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"optMailShotRequired" + iLoop + "No");
			<% /* End: MAR57 */ %>
		}
		
		<% /* Start: MAR57  - Maha T */ %>
		// check for ProspectPasswordTaken to true for any of the customers
		if (m_sProspectPassword[iLoop] == true)
		{
			m_bProspectPasswordTaken = true;
		}
	}
	//hide extra check boxes in MarketingInfo table
	for (iLoop = m_nMaxCustomers +1 ; iLoop < m_iTableLength; iLoop++)
	{
		<% /* Start: MAR57  - Maha T */ %>
		var chk = document.getElementById("MailShotRequiredGroup" + iLoop);
		chk.style.visibility = "hidden";
		
		var chk = document.getElementById("optMailShotRequired" + iLoop + "Yes");
		chk.style.visibility = "hidden";
		
		var chk = document.getElementById("optMailShotRequired" + iLoop + "No");
		chk.style.visibility = "hidden";
		<% /* End: MAR57 */ %>
	}
	
	<% /* Start: MAR57  - Maha T */ %>
	<% /* Disable password fields if prospectpasswordtaken for any customer */ %>
	if (m_bProspectPasswordTaken  == true)
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPassword");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPasswordConfirm");
	}
	else
	{
		if (m_bWrapUpPasswordTaken == true)
		{		 
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPassword");
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPasswordConfirm");
		}
		else
		{
			<% /* Reset label colour to red as only mandatory */ %>
			if(frmScreen.txtPassword.getAttribute("required") == "true")
				lblPassword.style.color = "red";		
			if(frmScreen.txtPasswordConfirm.getAttribute("required") == "true")
				lblPasswordConfirm.style.color = "red";	
		}
	}
	<% /* End: MAR57 */ %>
	
	//set Call Dtl
	frmScreen.optScheduleACallNo.checked = true
	frmScreen.optScheduleACallNo.onclick()
}

function PopulateCustomers()
{
	// read in loop all customers info into one xml object
	for(var nLoop = 1;nLoop <= m_nMaxCustomers;nLoop++)	
	{
		//Get the customer details
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

		XML.CreateRequestTag(m_BaseNonPopupWindow,"SEARCH");
		XML.CreateActiveTag("SEARCH");
		XML.CreateActiveTag("CUSTOMER");
		XML.CreateTag("CUSTOMERNUMBER",m_sCustomerNumber[nLoop]);
		XML.CreateTag("CUSTOMERVERSIONNUMBER",m_sCustomerVersion[nLoop]);
		
		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				XML.RunASP(document,"GetCustomerDetails.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
		}
		if(XML.IsResponseOK())
		{
			//create customerS xml (root)
			if (m_XMLCust == null) 
			{
				m_XMLCust =	new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
				m_XMLCust.CreateActiveTag("CUSTOMERS");
				// init tblMarketingInfo
				scScrollTable.initialiseTable(tblMarketingInfo, 0, "", ShowTable, m_iTableLength, m_nMaxCustomers);
			}
			// add XML.Customer node as a child
			m_XMLCust.AddXMLBlock(XML.SelectTag(null,"RESPONSE"));
			// update MarketingInfo array values
			m_sMarketingInfo[nLoop]		= XML.GetTagText("MAILSHOTREQUIRED")
			<% /* MAR57  - Maha T */ %>
			// Get ProspectPasswordTaken details 
			m_sProspectPassword[nLoop]	= XML.GetTagInt("PROSPECTPASSWORDTAKEN") //MAR1691
			
			if (m_sMarketingInfo[nLoop] == "1" || m_sMarketingInfo[nLoop] == "0")
				m_bMarketingEnabled[nLoop] = false;
			else
				m_bMarketingEnabled[nLoop] = true;
			
		}
		XML = null;	
	}
}

function GetTelephoneDetails(sTag)
{
<%	
	//fill up contacts from "wrap up" node (1 priority) or from "customer telephone number" (2-nd)
	//active node is sent as a parameter
	//tag CUSTOMERTELEPHONENUMBER and/or CUSTOMERWRAPUPDETAILS ovewrites previous version
%>
	m_XMLCust.SelectSingleNode(sTag);
	if ( m_XMLCust.ActiveTag != null ) 
	{
		//read wrapup or customer email address
		frmScreen.txtContactEMailAddress.value		 = m_XMLCust.GetTagText("WRAPUPEMAILADDRESS");
		frmScreen.optPREFERREDCONTACT5.checked		 = (m_XMLCust.GetTagText("WRAPUPEMAILPREFERRED") == "1")? true:false; 
		
		if (frmScreen.txtContactEMailAddress.value   == "")
		{
			frmScreen.txtContactEMailAddress.value	 = m_XMLCust.GetTagText("CONTACTEMAILADDRESS");
			frmScreen.optPREFERREDCONTACT5.checked	 = (m_XMLCust.GetTagText("EMAILPREFERRED") == "1")? true:false; 
		}
		
		//blank the seq nos array
		for (var nLoop = 0; nLoop < 4; nLoop++)	
					{m_nSeqNo[nLoop] = "";}
		//telephone
		m_XMLCust.CreateTagList("CUSTOMERWRAPUPDETAILS");
		if (m_XMLCust.ActiveTagList.length == 0)
		{
			var strPattern = sTag + "/CUSTOMERVERSION/CUSTOMERTELEPHONENUMBERLIST/CUSTOMERTELEPHONENUMBER"
			var xmlTagList = m_XMLCust.XMLDocument.selectNodes(strPattern);
		}
		<% /* START: MAR323 - (Maha T) If customerwrapupdetails found than ... */ %>
		else
		{
			var xmlTagList = m_XMLCust.XMLDocument.selectNodes(sTag + "/CUSTOMERVERSION/WRAPUP/CUSTOMERWRAPUPDETAILS");
		}
		<% /* END: MAR323  */ %>
		
		if (xmlTagList != null)
		{
			for(nLoop = 0;nLoop < xmlTagList.length && nLoop < 4;nLoop++)
			{
				var xmlTagXML  = xmlTagList.item(nLoop);
				
				frmScreen(sType[nLoop]).value		 = xmlTagXML.selectSingleNode("USAGE").text;
				frmScreen(sCountryCode[nLoop]).value = xmlTagXML.selectSingleNode("COUNTRYCODE").text;
				frmScreen(sAreaCode[nLoop]).value	 = xmlTagXML.selectSingleNode("AREACODE").text;
				frmScreen(sNumber[nLoop]).value		 = xmlTagXML.selectSingleNode("TELEPHONENUMBER").text;
				frmScreen(sExtension[nLoop]).value   = xmlTagXML.selectSingleNode("EXTENSIONNUMBER").text;
				frmScreen(sTime[nLoop]).value		 = xmlTagXML.selectSingleNode("CONTACTTIME").text;
				frmScreen(sOption[nLoop]).checked	 = (xmlTagXML.selectSingleNode("PREFERREDMETHODOFCONTACT").text == "1")? true:false; 
				if ( xmlTagXML.selectSingleNode("WRAPUPSEQUENCENUMBER") != null)
					m_nSeqNo[nLoop]	= xmlTagXML.selectSingleNode("WRAPUPSEQUENCENUMBER").text;
				
			}
		}
	}
}

function frmScreen.optScheduleACallYes.onclick()
{
	if ( GetOptionValue (frmScreen.optScheduleACallYes)=="1")
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtCallDate");
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtCallTime");
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboCallTime");
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboWhoToContact");
		//set contact list fields attribute
		frmScreen.cboWhoToContact.onchange();
	}
}

function frmScreen.optScheduleACallNo.onclick()
{
	if ( GetOptionValue (frmScreen.optScheduleACallNo)=="1")
	{
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtCallDate");
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtCallTime");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboCallTime");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboWhoToContact");
		for (var iLoop = 1; iLoop < 5 ; iLoop++)
		{
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboType" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCountryCode" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtAreaCode" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTelNumber" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtExtensionNumber" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTime" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"optPREFERREDCONTACT" + iLoop);
		}
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtContactEMailAddress");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optPREFERREDCONTACT5");
	}
}

function frmScreen.txtCallTime.onchange()
{
	if (frmScreen.txtCallTime.value != "")
	{
		frmScreen.cboCallTime.selectedIndex=0;
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboCallTime");
	}
	else
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboCallTime");
	}
}

function frmScreen.cboCallTime.onchange()
{
	if (frmScreen.cboCallTime.selectedIndex > 0)
	{
		frmScreen.txtCallTime.value="";
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCallTime");
	}
	else
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtCallTime");
	}
}

function frmScreen.cboWhoToContact.onchange()
{
	var value = frmScreen.cboWhoToContact.value;
	
	ClearContactCombos();
	
	if (frmScreen.cboWhoToContact.selectedIndex > 0) 
	{
		// refresh telephone contact list 
		m_XMLCust.ActiveTag = null;
		GetTelephoneDetails("CUSTOMERS/CUSTOMER[CUSTOMERNUMBER='" + value + "']");
		// set fields attributes
		for (var iLoop = 1; iLoop < 5 ; iLoop++)
		{
			scScreenFunctions.SetFieldToWritable(frmScreen,"cboType" + iLoop);
			SetContactRowAttributes(iLoop)
		}
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtContactEMailAddress");
		
		<% // MAR1374 - Peter Edney %>		
		if (frmScreen.txtContactEMailAddress.value.length > 0)
			scScreenFunctions.SetFieldToWritable(frmScreen,"optPREFERREDCONTACT5");
	}
}

<% /* Start: MAR57  - Maha T */ %>
function ClearContactCombos()
{
	var value = frmScreen.cboWhoToContact.value;
	for (var nLoop = 0; nLoop < 4 ; nLoop++)
	{
		// clear Contact List
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboType" + (nLoop+1));
		frmScreen(sType[nLoop]).selectedIndex = 0;
		frmScreen(sCountryCode[nLoop]).value = "";
		frmScreen(sAreaCode[nLoop]).value = "";
		frmScreen(sNumber[nLoop]).value = "";
		frmScreen(sExtension[nLoop]).value = "";
		frmScreen(sTime[nLoop]).value = "";
		frmScreen(sOption[nLoop]).checked = false; 
	}
	//set all rows to read only
	SetContactRowAttributes(0)
}
<% /* End: MAR57 */ %>

function frmScreen.cboType1.onchange()
{
	SetContactRowAttributes(1)
}

function frmScreen.cboType2.onchange()
{
	SetContactRowAttributes(2)
}

function frmScreen.cboType3.onchange()
{
	SetContactRowAttributes(3)
}

function frmScreen.cboType4.onchange()
{
	SetContactRowAttributes(4)
}

function frmScreen.txtContactEMailAddress.onchange()
{
	scScreenFunctions.SetFieldState(frmScreen, "optPREFERREDCONTACT5", ((frmScreen.txtContactEMailAddress.value) == "") ? "D":"W");
}

function SetContactRowAttributes(iRow)
{
	if (iRow == 0 )
	{
		//set all rows to read only
		for (var iLoop = 1; iLoop < 5 ; iLoop++)
		{
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboType" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCountryCode" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtAreaCode" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTelNumber" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtExtensionNumber" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTime" + iLoop);
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"optPREFERREDCONTACT" + iLoop);
		}
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtContactEMailAddress");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"optPREFERREDCONTACT5");
	}
	else
	{
		iRow-=1;
		scScreenFunctions.SetFieldState(frmScreen, sCountryCode[iRow], (frmScreen(sType[iRow]).value == "") ? "D":"W");
		scScreenFunctions.SetFieldState(frmScreen, sAreaCode[iRow], (frmScreen(sType[iRow]).value == "") ? "D":"W");
		scScreenFunctions.SetFieldState(frmScreen, sNumber[iRow], (frmScreen(sType[iRow]).value == "") ? "D":"W");
		scScreenFunctions.SetFieldState(frmScreen, sExtension[iRow], (frmScreen(sType[iRow]).value == "2") ? "W":"D");
		scScreenFunctions.SetFieldState(frmScreen, sTime[iRow], (frmScreen(sType[iRow]).value == "") ? "D":"W");
		scScreenFunctions.SetFieldState(frmScreen, sOption[iRow], (frmScreen(sType[iRow]).value == "") ? "D":"W");
	}
}

function PopulateMarketingInfo()
{
	if (m_nMaxCustomers > 0)
	{
		scScrollTable.initialiseTable(tblMarketingInfo, 0, "", ShowTable, m_iTableLength, m_nMaxCustomers);
		ShowTable(0);
		scScrollTable.setRowSelected(0);
	}
}

function ShowTable(iStart)
{			
	for (var iLoop = 0; iLoop < m_nMaxCustomers && iLoop < m_iTableLength; iLoop++)
	{
		scScreenFunctions.SizeTextToField(tblMarketingInfo.rows(iLoop).cells(0),m_sCustomerName[iLoop+1]);
		<% /* Start: MAR57  - Maha T */ %>
		SetCheckBoxValue(document.getElementById("optMailShotRequired" + (iLoop+1) + "Yes"), 
						 document.getElementById("optMailShotRequired" + (iLoop+1) + "No"),
						 m_sMarketingInfo[iLoop+1]);
		<% /* End: MAR57 */ %>
			
	}											
}


function SetCheckBoxValue( objOptionYes, objOptionNo, sVal )
{
	<% /* Start: MAR57  - Maha T */ %>
	if( sVal == "1" )
	{
		objOptionYes.checked = true;
	}
	else if( sVal == "0" )
	{
		objOptionNo.checked = true;
	}
	else
	{
		objOptionYes.checked = false;
		objOptionNo.checked = false;
	}
	<% /* End: MAR57 */ %>
}

function GetOptionValue( objOption )
{
	var sVal = "0";
	
	if(objOption.checked == true)
	{
		sVal = "1";
	}
	return(sVal)
}

function ValidatePassword(sPwdValue,sPwdValue2, bIsConfirmedPwd)
{	
	var bSuccess = false;
	var sPwdName = ((bIsConfirmedPwd == true) ? "The Confirmed password":"The Password");
	
	// JD MAR1287 only validate password if they have been entered 
	if(sPwdValue == "" && sPwdValue2 == "")
	{
		return true;
	}
	
	
	//compare confirmed password and password values	
	if (bIsConfirmedPwd == true)
	{
		// PJO 09/11/2005 MAR321 Make validation of passwords non case-sensitive
		bSuccess =((sPwdValue.toUpperCase() == sPwdValue2.toUpperCase()) ? true:false);
		if (bSuccess == false)
		{
			alert("The Confirmed password does not match the entered password");
			return bSuccess;
		}
	}
	
	<% /* MAR57  - Maha T */ %>
	//check password length
	if ( (sPwdValue.length < m_nMinPwdLength) || (sPwdValue.length > m_nMaxPwdLength) )
	{
		alert( sPwdName + " must be at least " + m_nMinPwdLength + " characters and no more that " + m_nMaxPwdLength + "characters");
		return bSuccess;
	}
	
	<% /* Start: MAR57  - Maha T */ %>
	//check presents of 1 letter and 1 numeric
	//var srchChr = /[a-zA-z]+/;
	var srchChr = /[a-zA-Z]+/; //MAR1587 M Heys 18/04/2006
	var srchInt = /[0-9]+/;
	if ( (sPwdValue.search(srchChr) != -1) && (sPwdValue.search(srchInt) != -1) )
	{
		bSuccess = true;
	}
	<% /* End: MAR57  */ %>
	
	if (bSuccess == false)
	{
		alert(sPwdName + " must contain both numbers and characters");
		return bSuccess;
	}
	
	return bSuccess;
}

function ValidateAddHocTaskDateTime()
{
	var bSuccess = false;
	m_sAddHocTaskDateTime = "";
	m_bLaunchAdHocTask = false;
	var gParamXML =		new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	//check if there is a data to initiate the task
	if ( frmScreen.txtCallDate.value == "" || (frmScreen.txtCallTime.value + frmScreen.cboCallTime.value) == "")
		return (bSuccess);

	var ValType= scScreenFunctions.GetComboValidationType(frmScreen, "cboCallTime");
	var sTaskTime = "";
	switch (ValType) {
		case "AM": 
			sTaskTime = gParamXML.GetGlobalParameterString(document,"TMWrapUpDefaultAM");
			break;
		case "PM":
			sTaskTime = gParamXML.GetGlobalParameterString(document,"TMWrapUpDefaultPM");
			break;
		default: 
			var sTaskTime = frmScreen.txtCallTime.value;
	}
	
	<% /* Start: MAR203  - Maha T */ %>
	// Check if its valid time format (00:00 to 23:59)
	if (/^([01]?[0-9]|[2][0-3])(:[0-5][0-9])?$/.test(sTaskTime) == false)
		return false;
	
	// Now check for date
	var sTaskDateTime = frmScreen.txtCallDate.value + " " + sTaskTime + ":00";
	bSuccess = scScreenFunctions.CompareDateStringToSystemDateTime (sTaskDateTime ,">=");
	<% /* End: MAR203 */ %>
		
	if (bSuccess == true) {
		m_sAddHocTaskDateTime = sTaskDateTime;
		m_bLaunchAdHocTask = true;
	}
	gParamXML = null;
	return bSuccess;
}

function ValidateScreen()
{
	// validate passwords
	var bSuccess = false;
	var bOptionChecked = false;

	if 	(m_bWrapUpPasswordTaken == false)
	{
		bSuccess = ValidatePassword (frmScreen.txtPassword.value,"",false)
		if (bSuccess == false)
		{	
			frmScreen.txtPassword.focus();
			return bSuccess;
		}
		else
		{
			bSuccess = ValidatePassword (frmScreen.txtPasswordConfirm.value,frmScreen.txtPassword.value,true)
			if (bSuccess == false)
			{
				frmScreen.txtPasswordConfirm.focus();
				return bSuccess;
			}
		}
	}
	
	if (GetOptionValue (frmScreen.optScheduleACallYes)=="1")
	{	
		// validate adhoc task date time
		bSuccess = ValidateAddHocTaskDateTime();
		if (bSuccess == false)
		{
			alert("Please enter valid call date, call time, contact details to schedule a call");
			frmScreen.txtCallDate.focus();
			return bSuccess;
		}
	
		// validate cboWhoToContact
		if (frmScreen.cboWhoToContact.value == "")
		{
			alert("Please fill up contact details");
			frmScreen.cboWhoToContact.focus();
			return false;
		}
		
		// validate contact details
		if ( (frmScreen(sType[0]).value + frmScreen(sType[1]).value + frmScreen(sType[2]).value + frmScreen(sType[3]).value) == "" )
		{
			alert("Please provide contact details for this Contact")
			frmScreen.all(sType[0]).focus();
			return false;
		}
	
		// validate telephone list
		for(nLoop = 0; nLoop < 4; nLoop++)
		{
			if (frmScreen(sType[nLoop]).value != "")
			{
				if (frmScreen(sNumber[nLoop]).value == "")
				{
					alert("Please enter a Telephone Number for each Contact Type");
					frmScreen.all(sType[nLoop]).focus();
					return false ;	
				}
				else
					bSuccess = true;
				<% /* TK 06/05/2006 MAR1384 */ %>
				if (frmScreen(sOption[nLoop]).checked == true)
					bOptionChecked = true;
			}
			else
			{
				if (frmScreen(sNumber[nLoop]).value != "" || frmScreen(sTime[nLoop]).value != "")
				{
					alert("For each telephone number please enter the Usage and Telephone Number");
					frmScreen.all(sType[nLoop]).focus();
					return false ;	
				}
				else
					bSuccess = true;
			}
		}	
		<% /* TK 06/05/2006 MAR1384 */ %>
		<% /* PSC 08/05/2006 MAR1384 */ %>
		if (bOptionChecked == false && frmScreen.optPREFERREDCONTACT5.checked == false)
		{
			alert("You must set the type of contact preferred by the customer");
			frmScreen.all(sOption[0]).focus();
			return false;
		}
		<% /* TK 06/05/2006 MAR1384 End */ %>
	}
	else
		bSuccess = true;

	return bSuccess;
}

function UpdateCRSustomer()
{	
	
	bSuccess = false;
	
	// check for password has been taken for the first time or any marketing programme radio button changes
	<% /* PSC 27/10/2005 MAR300 - Start */ %>
	<% /* MAR1349 update required on marketing button change when no password */ %>
	<% /* JD MAR1922 change to always do an UpdateCRS call*/ %>
	<% /*if 
		(
			(
				(!m_bWrapUpPasswordTaken && !m_bProspectPasswordTaken) 
				&& (frmScreen.txtPassword.value.length != 0)
			)
			|| m_blnMailShotUpdated
		)
	{*/ %>
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(m_BaseNonPopupWindow,"UPDATE");
		XML.SetAttribute("MessageType", "WrapUp");

		var xmlCustomers = XML.CreateActiveTag("CUSTOMERS");
				
		XML.CreateTag("PASSWORD", frmScreen.txtPassword.value);
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);

		XML.RunASP(document,"UpdateCRSCustomer.asp");
		if(XML.IsResponseOK()) bSuccess = true;					
		XML = null;
	<% /*}
	else
		bSuccess = true; */ %>
	<% /* PSC 27/10/2005 MAR300 - End */ %>
				
			
	return (bSuccess);
}

function SaveApplicationFact()
{
	var bSuccess = true;
	<% /* MAR1287 password optional */ %>
	if 	(m_bWrapUpPasswordTaken == false && frmScreen.txtPassword.value.length != 0)
	{	// password has been taken for the first time
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(m_BaseNonPopupWindow,"UPDATE");
		XML.CreateActiveTag("APPLICATIONFACTFIND");
		XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML.CreateTag("WRAPUPPASSWORDTAKEN","1");
		XML.RunASP(document,"UpdateApplicationFactFind.asp");
		if(XML.IsResponseOK()) bSuccess = true;
		else bSuccess = false;					
		XML = null;
	}
	if (frmScreen.cboSalesLeadStatus.value != "" && frmScreen.cboSalesLeadStatus.value != m_sSalesLeadStatus)
	{
		var XML2 = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML2.CreateRequestTag(m_BaseNonPopupWindow,"UPDATE");
		XML2.CreateActiveTag("APPLICATIONFACTFIND");
		XML2.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		XML2.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
		XML2.CreateTag("SALESLEADSTATUS",frmScreen.cboSalesLeadStatus.value);
		XML2.RunASP(document,"UpdateApplicationFactFind.asp");
		if(XML2.IsResponseOK()) bSuccess = true;
		else bSuccess = false;
		XML2 = null;
	}
	return bSuccess;
}

function SaveWrapUpContacts()
{
	var nSeq = frmScreen.cboWhoToContact.selectedIndex
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var requestTag = XML.CreateRequestTag(m_BaseNonPopupWindow,"UPDATE");
	var tagWrapUp = XML.CreateActiveTag("WRAPUP");
	var bSuccess = false;
	
	if (frmScreen.optScheduleACallYes.checked == true)
	{
		for (var nLoop = 0; nLoop < 4; nLoop++)
		{
			XML.ActiveTag =  tagWrapUp;
			XML.CreateActiveTag("CUSTOMERWRAPUPDETAILS");
			XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber[nSeq]);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersion[nSeq]);
			if (m_nSeqNo[nLoop] != "") 
				XML.CreateTag("WRAPUPSEQUENCENUMBER",m_nSeqNo[nLoop]);
		
			XML.CreateTag("USAGE", frmScreen(sType[nLoop]).value);
			XML.CreateTag("TELEPHONENUMBER", frmScreen(sNumber[nLoop]).value);
			XML.CreateTag("EXTENSIONNUMBER", frmScreen(sExtension[nLoop]).value);
			XML.CreateTag("CONTACTTIME", frmScreen(sTime[nLoop]).value);
			XML.CreateTag("PREFERREDMETHODOFCONTACT", (frmScreen(sOption[nLoop]).checked)? "1":"0");
			XML.CreateTag("COUNTRYCODE", frmScreen(sCountryCode[nLoop]).value);
			XML.CreateTag("AREACODE", frmScreen(sAreaCode[nLoop]).value);
			
			//update value m_sAddHocTaskNote for initiating Adhoc Task in the next step
			if (frmScreen(sOption[nLoop]).checked == true)
			{ 
				m_sAddHocTaskNote = "Contact Name: " +	frmScreen.cboWhoToContact.options(nSeq).text + ", "; 
				m_sAddHocTaskNote = m_sAddHocTaskNote + frmScreen(sType[nLoop]).options(frmScreen(sType[nLoop]).selectedIndex).text + " telephone: ";
				m_sAddHocTaskNote = m_sAddHocTaskNote + frmScreen(sCountryCode[nLoop]).value + " ";  
				m_sAddHocTaskNote = m_sAddHocTaskNote + frmScreen(sAreaCode[nLoop]).value + " " ;
				m_sAddHocTaskNote = m_sAddHocTaskNote + frmScreen(sNumber[nLoop]).value + " "; 
				m_sAddHocTaskNote = m_sAddHocTaskNote + ((frmScreen(sExtension[nLoop]).value=="")? "" : "ext." + frmScreen(sExtension[nLoop]).value) ;
				m_sAddHocTaskNote = m_sAddHocTaskNote + ", date: "  + frmScreen.txtCallDate.value + ", time: "; 
				m_sAddHocTaskNote = m_sAddHocTaskNote + ((frmScreen.cboCallTime.value == "") ? frmScreen.txtCallTime.value : frmScreen.cboCallTime.options(frmScreen.cboCallTime.selectedIndex).text);
			}
		}
	}
	
	//save only changed MarketingInfo opt in and save WrapUp email only for selected customer
	for (nLoop = 1; nLoop < 5; nLoop++)
	{
		//update mailshot required
		if (m_bMarketingEnabled[nLoop] == true)
		{ 
			<% /* Start: MAR57  - Maha T */ %>
			//if (frmScreen(chkMailShotRequired + nLoop).checked)
			XML.ActiveTag = requestTag;
			XML.CreateActiveTag("CUSTOMERVERSION");
			XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber[nLoop]);
			XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersion[nLoop]);
			if (frmScreen("optMailShotRequired" + nLoop + "Yes").checked)
			{
				XML.CreateTag("MAILSHOTREQUIRED","1");
				<% /* PSC 27/10/2005 MAR300 */%>
				m_blnMailShotUpdated  = true;

			}
			else if (frmScreen("optMailShotRequired" + nLoop + "No").checked)
			{
				XML.CreateTag("MAILSHOTREQUIRED","0");
				<% /* PSC 27/10/2005 MAR300 */%>
				m_blnMailShotUpdated  = true;
			}
			else
			{
				XML.CreateTag("MAILSHOTREQUIRED","NULL");
			}
			<% /* End: MAR57 */ %>
		}
		//update wrapup email
		if (nSeq == nLoop)
		{
			if (XML.ActiveTag.baseName != "CUSTOMERVERSION")
			{
				XML.ActiveTag = requestTag;
				XML.CreateActiveTag("CUSTOMERVERSION");
				XML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber[nLoop]);
				XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersion[nLoop]);
			}
			XML.CreateTag("WRAPUPEMAILADDRESS", frmScreen.txtContactEMailAddress.value);
			XML.CreateTag("WRAPUPEMAILPREFERRED", (frmScreen.optPREFERREDCONTACT5.checked)? "1":"0");
			
			//update value m_sAddHocTaskNote for initiating Adhoc Task in the next step
			if (frmScreen.optPREFERREDCONTACT5.checked == true)
			{ 
				m_sAddHocTaskNote = "Contact Name " + frmScreen.cboWhoToContact.options(nSeq).text + ", ";
				m_sAddHocTaskNote = m_sAddHocTaskNote + " email address: " + frmScreen.txtContactEMailAddress.value ;
				m_sAddHocTaskNote = m_sAddHocTaskNote + ", date: "  + frmScreen.txtCallDate.value + ", time: " ;
				m_sAddHocTaskNote = m_sAddHocTaskNote + ((frmScreen.cboCallTime.value == "")? frmScreen.txtCallTime.value : frmScreen.cboCallTime.options(frmScreen.cboCallTime.selectedIndex).text);
			}
		}
	}
	XML.RunASP(document, "SaveWrapUpDetails.asp");
	if(XML.IsResponseOK()) bSuccess = true;					
	return (bSuccess);
}

function LaunchAdHocTask()
{
	var bSuccess = false;
	var bTaskUpdate	= false;
	var sTaskSeqNo = "";		
	
	//get params..
	var stageXML =		new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var taskXML =		new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var gParamXML =		new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTaskID =		gParamXML.GetGlobalParameterString(document,"TMWrapUpScheduleCallTaskID");				
	var tagSourceAppl	= "Omiga"
	var sUserRole =		scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idRole",null);
	var sUserID =		scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idUserID",null)	
	var sUnitID =		scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idUnitID",null)
	var sApplPriority = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idApplicationPriority",null)
	var sActivityID =	scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idActivityId",null)
	//var	sTaskXML =		scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idTaskXML",null);
	var sStageID =		scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idStageID",null)
	
	//check if exists
	stageXML.CreateRequestTag(m_BaseNonPopupWindow , "GetCurrentStage");
	stageXML.CreateActiveTag("CASEACTIVITY");
	stageXML.SetAttribute("SOURCEAPPLICATION", tagSourceAppl);
	stageXML.SetAttribute("CASEID", m_sApplicationNumber);
	stageXML.SetAttribute("ACTIVITYID", sActivityID);
	stageXML.SetAttribute("ACTIVITYINSTANCE", "1");
	stageXML.RunASP(document, "MsgTMBO.asp");
	
	if(stageXML.IsResponseOK() == false)
		return (bSuccess);
		
	sTaskSeqNo = stageXML.GetTagAttribute("CASESTAGE","CASESTAGESEQUENCENO");
	
	<% /* Check if this task already exists as incomplete, only add if it does 
		not exist or if has been set to complete */ %>
	var TempXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();						
	var tagList = stageXML.CreateTagList("CASETASK");			
	
	for (var nItem = 0; nItem < tagList.length && bSuccess == false; nItem++) {
		stageXML.SelectTagListItem(nItem);			
		if(stageXML.GetAttribute("TASKID")==sTaskID) {				
			var sTaskStatus = stageXML.GetAttribute("TASKSTATUS");
			if (TempXML.IsInComboValidationList(document,"TaskStatus",sTaskStatus,["I"])) 
			{
				//incomplete then check due date, time
				//var sOldTaskDateTime = stageXML.GetAttribute("TASKDUEDATEANDTIME");
				//if (scScreenFunctions.CompareDateStringToSystemDateTime (sOldTaskDateTime ,">=")) {
					
				//}
				bTaskUpdate = true;
				bSuccess = true; 
				break;
			}		
		}				
	}
	bSuccess = false;
	if (bTaskUpdate == false) 
	{
		// create AdHocTask
		var reqTag = taskXML.CreateRequestTag(m_BaseNonPopupWindow, "CreateAdhocCaseTask");	
		taskXML.SetAttribute("USERID", sUserID);
		taskXML.SetAttribute("UNITID", sUnitID);		
		taskXML.SetAttribute("USERAUTHORITYLEVEL", sUserRole);				
		taskXML.CreateActiveTag("APPLICATION");		
		taskXML.SetAttribute("APPLICATIONPRIORITY", sApplPriority);
		taskXML.ActiveTag = reqTag;	
		taskXML.CreateActiveTag("CASETASK");
		taskXML.SetAttribute("SOURCEAPPLICATION", tagSourceAppl);
		taskXML.SetAttribute("CASEID", m_sApplicationNumber);	
		taskXML.SetAttribute("ACTIVITYID", sActivityID);
		taskXML.SetAttribute("ACTIVITYINSTANCE", "1");		
		taskXML.SetAttribute("CASESTAGESEQUENCENO", sTaskSeqNo);
		taskXML.SetAttribute("STAGEID",sStageID);		
		taskXML.SetAttribute("TASKID", sTaskID);
		taskXML.SetAttribute("TASKDUEDATEANDTIME", m_sAddHocTaskDateTime);
		
		taskXML.RunASP(document, "OmigaTmBO.asp");
	}
	else
	{
		// Update Case Task
		reqTag = taskXML.CreateRequestTag(m_BaseNonPopupWindow, "UpdateCaseTask");
			
		taskXML.CreateActiveTag("CASETASK");
		taskXML.SetAttribute("SOURCEAPPLICATION", tagSourceAppl);
		taskXML.SetAttribute("CASEID",m_sApplicationNumber);
		taskXML.SetAttribute("ACTIVITYID", sActivityID);
		taskXML.SetAttribute("STAGEID",sStageID);
		taskXML.SetAttribute("CASESTAGESEQUENCENO", sTaskSeqNo);
		taskXML.SetAttribute("TASKID", sTaskID);
		taskXML.SetAttribute("TASKINSTANCE", stageXML.GetAttribute("TASKINSTANCE"));
		taskXML.SetAttribute("TASKDUEDATEANDTIME", m_sAddHocTaskDateTime);
			
		taskXML.RunASP(document, "msgTMBO.asp") ;
	}				
	if(taskXML.IsResponseOK())
	{	
		bSuccess = true;
		<% /*	// to be revised due to change in omTM 
				// to do : add a new method to create adhoc task and note under one transaction
		bSuccess = false;
		// create task note
		var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(m_BaseNonPopupWindow, "CreateTaskNote");
		XML.CreateActiveTag("TASKNOTE");
		XML.SetAttribute("SOURCEAPPLICATION", tagSourceAppl);
		XML.SetAttribute("CASEID", m_sApplicationNumber);
		XML.SetAttribute("ACTIVITYID", sActivityID);
		XML.SetAttribute("ACTIVITYINSTANCE", "1");
		XML.SetAttribute("STAGEID", sStageID);
		XML.SetAttribute("CASESTAGESEQUENCENO", stageXML.GetAttribute("CASESTAGESEQUENCENO"));
		XML.SetAttribute("TASKID", sTaskID);
		XML.SetAttribute("TASKINSTANCE", stageXML.GetAttribute("TASKINSTANCE"));

		XML.SetAttribute("NOTEENTRY", m_sAddHocTaskNote);
				//XML.SetAttribute("NOTEORIGINATINGUSERID", scScreenFunctions.GetContextParameter(window, "idUserId", ""));
				//XML.SetAttribute("NOTETYPE", frmScreen.cboNoteType.value);
		XML.SetAttribute("CONTACTNOTEREASON", "20");//Contact
		
		XML.RunASP(document, "MsgTMBO.asp");
		if(XML.IsResponseOK()) bSuccess = true; 
		*/ %>		

		<% // MAR1374 - Peter Edney %>
		if(bTaskUpdate == false){
			var stageXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			stageXML.CreateRequestTag(m_BaseNonPopupWindow , "GetCurrentStage");
			stageXML.CreateActiveTag("CASEACTIVITY");
			stageXML.SetAttribute("SOURCEAPPLICATION", tagSourceAppl);
			stageXML.SetAttribute("CASEID", m_sApplicationNumber);
			stageXML.SetAttribute("ACTIVITYID", sActivityID);
			stageXML.SetAttribute("ACTIVITYINSTANCE", "1");
			stageXML.RunASP(document, "MsgTMBO.asp");					
			bSuccess = stageXML.IsResponseOK();
			if(bSuccess){
				stageXML.XMLDocument.setProperty("SelectionLanguage","XPath");
				stageXML.SelectSingleNode("/RESPONSE/CASESTAGE/CASETASK[last()]");			
			}
		}
		
		if(bSuccess){		
			var tasknoteXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();		

			tasknoteXML.CreateRequestTag(m_BaseNonPopupWindow, "CreateTaskNote");
			tasknoteXML.CreateActiveTag("TASKNOTE");
			tasknoteXML.SetAttribute("SOURCEAPPLICATION", tagSourceAppl);
			tasknoteXML.SetAttribute("CASEID", m_sApplicationNumber);
			tasknoteXML.SetAttribute("ACTIVITYID", sActivityID);
			tasknoteXML.SetAttribute("ACTIVITYINSTANCE","1");
			tasknoteXML.SetAttribute("STAGEID", sStageID);
			tasknoteXML.SetAttribute("CASESTAGESEQUENCENO", sTaskSeqNo);
			tasknoteXML.SetAttribute("TASKID", sTaskID);
			tasknoteXML.SetAttribute("TASKINSTANCE", stageXML.GetAttribute("TASKINSTANCE"));
			tasknoteXML.SetAttribute("NOTEENTRY", m_sAddHocTaskNote);
			tasknoteXML.SetAttribute("NOTEORIGINATINGUSERID", m_sUserID);
			tasknoteXML.SetAttribute("NOTETYPE", 10);
			
			tasknoteXML.RunASP(document, "msgTMBO.asp") ;
			bSuccess = tasknoteXML.IsResponseOK();
		}

	}
	stageXML = null;
	taskXML = null;
	gParamXML = null;
	TempXML = null;
	tagList = null;
	//XML=null;
	return bSuccess;
}

<% /* MAR324  Context variable AllowExitFromWrap is now used instead of this function */ %>
<% /*
function CheckCloseApplication()
{
	bSuccess = false;
	
	//get exitfromwrapup for this user & unit
	var UnitXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	UnitXML.CreateRequestTag(m_BaseNonPopupWindow , "FindCurrentUnitList");
	UnitXML.CreateActiveTag("UNIT");
	UnitXML.CreateTag("USERID",m_sUserID);
	UnitXML.CreateTag("UNITID",m_sUnitID);
	UnitXML.RunASP(document, "FindUnitsForUser.asp");
	if (UnitXML.IsResponseOK())
	{
		var sAllowOmigaExitFromWrap = UnitXML.GetTagText("ALLOWOMIGAEXITFROMWRAP")
		if (sAllowOmigaExitFromWrap == 1)
			bSuccess = true;
		else
			bSuccess = false;
	}
	
	return bSuccess;
}   */ %>

function UnLockApplication()
{
	// call ValidateApplicationOrCustomer from modal_omigamenu.asp to unlock currect customer & unit
	m_BaseNonPopupWindow.top.frames["omigamenu"].ValidateApplicationOrCustomer(false);
}
	
function LogOffAdminSystem()
{
	// call omAdmin.AdministrationInterfaceBO.LogOffUser 
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateActiveTag("REQUEST");
	XML.SetAttribute("OPERATION", "LogOffUser");
	XML.SetAttribute("ADMINSYSTEMSTATE", objHiddenFrameForm.txtAdminSystemState.value);
	// alert(XML.XMLDocument.xml)

	// Run server-side asp script
	XML.RunASP(document, "omAdmin.asp");
	// alert(XML.XMLDocument.xml); 

	if (XML.IsResponseOK() == true)
	{
		// delete ADMINSYSTEMSTATE
		objHiddenFrameForm.txtAdminSystemState.value = null;
	}
	XML = null;
}


function btnSubmit.onclick()
{
	var bSuccess = false;
	
	if (m_sReadOnly != "1")
	{
		<% // MAR1366 - Peter Edney - 07/03/2006
		/* MAR1287 check mandatory settings if they are set */
		//if(frmScreen.txtPassword.getAttribute("required") == "true" ||
		//   frmScreen.txtPasswordConfirm.getAttribute("required") == "true")
		//	bSuccess = frmScreen.onsubmit();
		//else bSuccess = true; %>				
		bSuccess = frmScreen.onsubmit();
		
		if(bSuccess)
		{
			// MAR1366 - Peter Edney - 07/03/2006
			bSuccess = ValidateScreen();		
			if(bSuccess)
			{
				//save ApplicationFact: WrapUpPasswordTaken and SalesLeadTime
				bSuccess = SaveApplicationFact();
				//save WrapUpDeatils
				if (bSuccess)
					bSuccess = SaveWrapUpContacts();
				//update CRS Customer with Password
				<% /* MAR1287 password optional */ %>
				<% /* MAR1349 update required on marketing button change */ %>
				if (bSuccess)
					bSuccess = UpdateCRSustomer();

				//create adhoc Task
				if (bSuccess && m_bLaunchAdHocTask)
				{
					bSuccess = LaunchAdHocTask();
				}
			}
		}
	}
	else
		bSuccess=true;
	
	if (bSuccess==true)
	{
		<% /* MAR324  Use context variable to test ALLOWOMIGAEXITFROMWRAP */ %>
		if (m_sAllowExitFromWrap == "1")
		// close application
		{
			var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			var bLegacyCustomer = XML.GetGlobalParameterBoolean(document,"FindLegacyCustomer");
			XML = null;

			if (bLegacyCustomer)
			{
				UnLockApplication();
				//LogOffAdminSystem();
				window.returnValue = "exitOmiga";
				window.close();
			}
		}
		else
			window.close();
	}	
	m_blnSubmitProcessed = true;      <% /* MAR1587 M Heys 18/04/2006  this is a status flag for the OnUnload 
																and does not imply error free completion*/ %>
}
<% /* MAR1587 M Heys 18/04/2006 Start */ %>
function window.onunload()
{	
	if (!m_blnSubmitProcessed)
	{
		btnSubmit.onclick();
	}
}
<% /* MAR1587 M Heys 18/04/2006 end */ %>
</script>
</body>
</HTML>
