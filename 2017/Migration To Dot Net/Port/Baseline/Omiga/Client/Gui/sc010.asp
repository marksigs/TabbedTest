<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<% /*
Workfile:      sc010.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Logon screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		03/11/99	Created.
RF		18/11/99	Changed from div-based to table-based form.
RF		25/11/99	AQR SC003: Change to 2 labels.
					AQR SC005: Ensure combo shows all units found.
					AQR SC004: Fix combo width.
RF		01/12/99	AQR SC008: Change to when FindUnitsForUser is called;
					function frmScreen.txtPassword.onblur() no longer required 
					AQR SC002: Add screen id.
RF		28/01/00	Reworked for IE5 and performance.
AY		03/04/00	scScreenFunctions change
MC		11/05/00	SYS0178 Do not FindUnitsForUser if userid not entered
MC		11/05/00	SYS0481. Use UnitName in SC030
MH      17/05/00    SYS0481  Use UnitName in SC010
AP		04/07/00	Improved html creen layout and used latest graphics
GD		20/12/00	SYS1742 changes made to allow use of hidden frame
					and eliminate session variables
GD		10/01/01	SYS0203: In window.onload added call to LogonLegallyOpened()
MV		05/03/01	SYS2001: commented AccessAuditGUID
JLD		09/10/01	SYS2737: Removed hard coding of machineName. Now set up in SC030 
RF		08/11/01	SYS2927: Omiplus change control 34 - Security: Log on to Optimus.
RF		22/01/02	SYS3866: Omiplus change control 42 - Security: 
					Screen SC010 to allow separate passwords for Omiga and Optimus.
					Allow access to Omiga irrespective of whether Opimus logon succeeds.
MDC		16/04/02	SYS4396: Enable client customisation of system name
TW      09/10/02	SYS5115   Modified to incorporate client validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
AW		12/07/02	BMIDS00178	Removed Admin System Login
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
JR		23/01/2003	BM0271		Select correct ProcessingInd from returned XML.
BS		23/03/2003	BM0286		Fix ProcessingIndicator change 
HMA     29/01/2004  BMIDS678    Set up CreditCheckAccess context parameter
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :

Prog	Date		AQR			Description
GHun	21/07/2005	MAR14		Apply ING Style Sheet and GUI Images
LD		19/10/2005  MAR123		NT Authentication
PSC		15/11/2005	MAR310		Check for ALLOWOMIGALOGON
PJO     21/11/2005  MAR417      Increase User ID screen size to 30 chars
HMA     25/11/2005  MAR324      Add AllowExitFromOmiga context parameter.
PJO     30/11/2005	MAR417      The USer ID should allow input of 64 chars
PE		15/12/2005	MAR798		Using Windows authentication should never  route to change password screen.
AW		12/12/2006	EP1240		Amendments for Quality Checker, Approved user (DMS availability)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<html>

<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="LogonStylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title>Logon</title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets - remove any which are not required */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>

<% /* Specify Forms Here */ %>

<!-- GD	
RF 08/11/01	SYS2927 Added AdminSystemState
<form id="frmGoNextScreen" method="post" STYLE="DISPLAY: none">	
	<div id="ELEMENTS" style="display:none"> 	
		<input id="txtDebug" name="Debug">	
		<input id="txtUserId" name="UserId">	
		<input id="txtUserName" name="UserName">	
		<input id="txtAccessType" name="AccessType">	
		<input id="txtUserCompetency" name="UserCompetency">	
		<input id="txtCreditCheckAccess" name="CreditCheckAccess">
		<input id="txtQualificationsListXml" name="QualificationsListXml">	
		<input id="txtRole" name="Role">	
		<input id="txtUnitId" name="UnitId">	
		<input id="txtDepartmentId" name="DepartmentId">	
		<input id="txtDistributionChannelId" name="DistributionChannelId">	
		<input id="txtProcessingIndicator" name="ProcessingIndicator">	
		<input id="txtMachineId" name="MachineId">	
		<input id="txtUnitName" name="UnitName">
		<input id="txtAdminSystemState name="AdminSystemState">
		<input id="txtAllowExitFromWrap" name="AllowExitFromWrap">
		<input id="txtIsUserQualityChecker name="IsUserQualityChecker">
		<input id="txtIsUserFulfillApproved" name="IsUserFulfillApproved">
	</div>
</form>
GD -->
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% // Specify Screen Layout Here %>
<form id="frmScreen" mark validate="onchange">

<% // ADP: Outer table ensures everything inside is centrally aligned both horizontally and vertically %>
<table width="100%" height="90%" border="0" cellspacing="0" cellpadding="0" onkeydown="checkKeyPress()">
<tr>
	<td align="center" valign="middle">	
	
		<% // ADP: This table defines the overall size of the logon screen area %>	
		<table id="tblMain" width="640" height="385" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
		<tr>
			<td align="center">
		
				<%	// ADP: Another table divides the screen into two columns. The first holds the 
					// MainBack.gif graphic and the second contains the title bar and logon fields %>
				<table width="100%" height="100%" border="2" bordercolor="#616161" cellspacing="0" cellpadding="0">
				<tr>
					<td width="152" valign="top"><img src="images/mainback.gif" width="152" height="265" alt="" border="0"></td>	
					<td bgcolor="#FFFFFF" valign="top">
						
						<% // ADP: This table contains three rows, the first to represent the menubar, the second to hold the page title and screen Id in two columns, and the third to hold another table where the logon fields are located %>
						<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
							<tr bgcolor="#FFFFFF">
								<td height="20" colspan="2"></td>
							</tr>
							<tr >
							<td height="30" class="msgVerdanaLeft">
								<% // Screen Title and Id %>
								System Logon
							</td>
							<td height="30" align="right" class="msgVerdanaRight">
								SC010
							</td>
						</tr>
						<tr>
							<td align="center" colspan="2" valign="middle">
							
								<% // ADP: Here's the table containing the logon fields %>
								<table width="350" height="150" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">	
								<tr>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td width="30" height="40">&nbsp;</td>
											<td class="msgLabel"><strong>Please enter your logon details ...</strong></td>
										</tr>
										</table>
										<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td width="30" height="30">&nbsp;</td>
											<td width="70" height="30" class="msgLabel">User Id.</td>
											<td width="340" height="30"><input id="txtUserId" name="UserId" maxlength="64" style="WIDTH: 240px;" class="msgTxt">&nbsp;</td>
										</tr>
										<tr>
											<td width="30" height="30">&nbsp;</td>
											<td width="70" height="30" class="msgLabel"><label id="idPassword">Password</label></td>
											<td width="170" height="30"><input id="txtOmigaPassword" type="PASSWORD" name="Password" maxlength="15" style="WIDTH: 100px;" class="msgTxt"></td>
										</tr>
										<tr>
											<td width="30" height="30">&nbsp;</td>
											<td width="70" height="30" class="msgLabel">Unit Name</td>
											<td width="170" height="30"><select id="cboUnitName" name="UnitName" style="WIDTH: 200px" class="msgCombo"></select></td>
										</tr>
										</table>
									</td>
								</tr>	
								</table>
			
								<% // ADP: OK and Change Password buttons outside of the logon fields group %>
								<table border="0" width="350" cellpadding="0" cellspacing="0">
								<tr><td height="5"></td></tr>
								<tr>
									<td width="65"><input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">&nbsp;</td>
									<td><input id="btnChangePassword" value="Change Password" type="button" style="WIDTH: 120px" class="msgButton">&nbsp;</td>
								</tr>
								</table>
							</td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

</form>
	
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* SYS4396 MDC 16/04/2002 - Customisation of System Name */ %>
<!-- #include FILE="customise/sc010customise.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/sc010Attribs.asp" -->
<% /* SYS4396 MDC 16/04/2002 - End */ %>

<% /* Specify Code Here */ %>
<script language="JScript" type="text/javascript">
<!--
<% // JScript Code %>
var m_sMetaAction = null;		
var m_bDebug = false;
		
<%// xml returned from FindUnitsForUser %>
var m_sFindUnitsForUserXml;		
var m_sUserId = null;
var scScreenfunctions;
<%
//GD New - global object points to form on hidden frame
//NB MAINTENANCE - If for any reason the location of the hidden frame/form changes
//					then objHiddenFrameForm will need to change. 
//					It is defined in window on load.
%>
var objHiddenFrameForm;
		
//var m_bLegacyCustomer;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	<%	//Check user has reached this page by the correct path	%>
	
	if (scScreenFunctions.LogonLegallyOpened()==true)
	{
		objHiddenFrameForm=top.frames("fraDetails").document.forms("frmDetails");

		<% /* SYS4396 MDC 16/04/2002 - Customisation of System Name */ %>
		Customise();
		<% /* SYS4396 MDC 16/04/2002 - End */ %>
		
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
		Validation_Init();				
		Initialise();
		scScreenFunctions.SetFocusToFirstField(frmScreen);
	}
	else
	{
		<%	//The user has accessed this page via a dubious route - send back to logon	%>
		top.location.href="LogonFrameset.asp";
	}	
	
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

function Initialise()
{
	scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.cboUnitName.id)
	<%	/* 
		set context variable id fields to null - 
		this prevents potential security problems if  browser
		"back" and "forward" buttons are used
	*/	%>
	<% //GD New  %>
	objHiddenFrameForm.txtDebug.value = m_bDebug;
	objHiddenFrameForm.txtUserId.value = null;
	objHiddenFrameForm.txtUserName.value = null;
	objHiddenFrameForm.txtAccessType.value = null;
	objHiddenFrameForm.txtUserCompetency.value = null;
	objHiddenFrameForm.txtCreditCheckAccess.value = null;  
	objHiddenFrameForm.txtQualificationsListXml.value = null;
	objHiddenFrameForm.txtRole.value = null;
	objHiddenFrameForm.txtUnitId.value = null;
	objHiddenFrameForm.txtDepartmentId.value = null;
	objHiddenFrameForm.txtDepartmentName.value = null;
	objHiddenFrameForm.txtDistributionChannelId.value = null;
	objHiddenFrameForm.txtDistributionChannelName.value = null;
	objHiddenFrameForm.txtProcessingIndicator.value = null;
	objHiddenFrameForm.txtMachineId.value = null;
	objHiddenFrameForm.txtUnitName.value = null;
	// objHiddenFrameForm.txtAccessAuditGUID.value = null;
	<% //GD New %>
	<% // RF 08/11/01 SYS2927 Added AdminSystemState %>
	objHiddenFrameForm.txtAdminSystemState.value = null;
	
	<% // MAR324 Added AllowExitFromWrap %>
	objHiddenFrameForm.txtAllowExitFromWrap.value = null;                
	
	<% // AW 30/10/06 EP1240  %>
	objHiddenFrameForm.txtIsUserQualityChecker.value = null;
	objHiddenFrameForm.txtIsUserFulfillApproved.value = null;
		
	//InitialiseFindLegacyCustomer();	
	
	m_sUserId = null;
	frmScreen.txtUserId.value = "";
/*	AW	12/07/02	BMIDS00178	
	<% // populate Environment combo %>
	if (IsLegacyCustomer())
	{
		var XML = new scXMLFunctions.XMLObject();			
		var sGroups = new Array("AdminSystemEnvironment");
		if (XML.GetComboLists(document, sGroups) == true)
		{
			XML.PopulateCombo(document, frmScreen.cboEnvironment, "AdminSystemEnvironment", true);
		}
		XML = null;
	}
	else
	{
		<% //frmScreen.cboEnvironment.disabled = true; %>
		<% //frmScreen.cboEnvironment.style.visibility= "hidden"; %>
		scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.cboEnvironment.id);
		scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.txtAdminSystemPassword.id);
		
	}
*/
}

function frmScreen.btnOK.onclick()
{
	if(frmScreen.onsubmit() == true)
		Logon();
}
		
function frmScreen.btnChangePassword.onclick()
{
	if(frmScreen.onsubmit() == true)
		ChangePassword();
}		
		
function ChangePassword()
{
	<% // Validate logon and route to SC020 %>
	
	var bReturnArray = new Array(2);
	bReturnArray = ValidateUser();
				
	if (bReturnArray[0] == true)	// logon ok
	{
		var sUnitName = frmScreen.cboUnitName.options(frmScreen.cboUnitName.selectedIndex).text;

		objHiddenFrameForm.txtUnitName.value = sUnitName;

		<% // route to SC020 ChangePassword - sc020.asp %>
		top.document.frames("fraMain").location.href="sc020.asp";
	}
}		

function Logon()
{
	<% // Validate logon and route to SC020 or SC030 %>

	var bReturnArray = new Array(2);
	bReturnArray = ValidateUser();
				
	if (bReturnArray[0] == true)	<% // logon ok %>
	{
		var TagOption = frmScreen.cboUnitName.options(frmScreen.cboUnitName.selectedIndex);   // MAR324
		var sUnitName = TagOption.text;														  // MAR324
		var sAllowExitFromWrap = TagOption.getAttribute("AllowExitFromWrap");			      // MAR324

		objHiddenFrameForm.txtUnitName.value = sUnitName;
		objHiddenFrameForm.txtAllowExitFromWrap.value = sAllowExitFromWrap;                   // MAR324
	
		<% // check where to route to %>
		if (bReturnArray[1] == true)	<% // change password %>
			<% // route to SC020 ChangePassword %>
			top.document.frames("fraMain").location.href="sc020.asp";
		else
			<% // route to SC030 LaunchOmiga4 %>
			top.document.frames("fraMain").location.href="sc030.asp";					
	}
}
		
function ValidateUser()
{
	<% /* 
		Returns:		
			bReturnArray, a one dimensional array:
			bReturnArray[0] - A boolean indicating if logon was successful
			bReturnArray[1] - A boolean indicating ChangePasswordIndicator value 
	*/ %>

	var bContinue = true;
	bLogonOk = false;
	bChangePassword = false;
				
	m_sUserId = frmScreen.txtUserId.value;			
	if (m_sUserId == "")
	{
		alert("Please enter your User Id");
		frmScreen.txtUserId.focus();
		bContinue = false;
	}
				
	if (bContinue)
	{
		if (frmScreen.txtOmigaPassword.value.length == 0)
		{
			alert("Please enter your password");
			frmScreen.txtOmigaPassword.focus();
			bContinue = false;
		}
	}
/*	AW	12/07/02	BMIDS00178
	if (bContinue && IsLegacyCustomer())
	{
		if (frmScreen.txtAdminSystemPassword.value.length == 0)
		{
			alert("Please enter your Admin System Password");
			frmScreen.txtAdminSystemPassword.focus();
			bContinue = false;
		}
	}
*/
	if (bContinue)
	{
		if (frmScreen.cboUnitName.value == "")
		{
			if (frmScreen.cboUnitName.options.length == 0)
			{
				FindUnitsForUser();
				if (frmScreen.cboUnitName.options.length != 1)
					bContinue = false;
			}
			else
			{
				alert("Please select a Unit");
				frmScreen.cboUnitName.focus();
				bContinue = false;
			}
		}
	}
/*	AW	12/07/02	BMIDS00178				
	if (bContinue && IsLegacyCustomer())
	{
		<% // validate the Optimus logon %>
		
		<%	// RF 08/11/01 SYS2927 Attempt logon to Optimus - 
			// call omAdmin.AdministrationInterfaceBO.ValidateUserLogon %>  
			
		if (frmScreen.cboEnvironment.options.length > 1 && 
			frmScreen.cboEnvironment.selectedIndex == 0)
		{
			frmScreen.cboEnvironment.focus();
			alert("Please select an Environment");
			bContinue = false;
		}
	}
			
	if (bContinue && IsLegacyCustomer())
	{
		var XML = new scXMLFunctions.XMLObject();
		var TagRequest = XML.CreateActiveTag("REQUEST");
		XML.SetAttribute("OPERATION", "ValidateUserLogon");

		XML.CreateActiveTag("USER");
		XML.SetAttribute("USERNAME", m_sUserId.toUpperCase());
		XML.SetAttribute("PASSWORDVALUE", frmScreen.txtAdminSystemPassword.value.toUpperCase());

		XML.ActiveTag = TagRequest;		
		XML.CreateActiveTag("ODIINITIALISATION");
		XML.SetAttribute("ODIENVIRONMENT", 
			frmScreen.cboEnvironment.options(frmScreen.cboEnvironment.selectedIndex).text);
		
		<% // alert(XML.XMLDocument.xml); %>

		<% // Run server-side asp script %>
		XML.RunASP(document, "omAdmin.asp");
		
		if (m_bDebug)
			XML.WriteXMLToFile("C:\\temp\\sc010omAdmin.xml");
		
		<% // alert(XML.XMLDocument.xml); %>

		if (XML.IsResponseOK() == true)
		{
			<% // save ADMINSYSTEMSTATE %>
			var tagState = XML.SelectTag(null, "ADMINSYSTEMSTATE");						
			if (tagState != null)
				objHiddenFrameForm.txtAdminSystemState.value = tagState.xml;
			else
			{
				alert("Admin System State not found");
				bContinue = false;
			}
		}
		else
		{
			<% // RF 24/01/02 Failure to log on to Admin System is not a show stopper %>
			<% // bContinue = false; %>
			alert("Log on to Admin System failed");
		}
		XML = null;
	}
*/	
	if (bContinue)
	{		
		<% // validate the Omiga logon %>

		var XML = new scXMLFunctions.XMLObject();
		
		XML.CreateActiveTag("REQUEST");
		XML.SetAttribute("USERID", m_sUserId);
		
		XML.CreateActiveTag("OMIGAUSER");
		XML.CreateTag("USERID", m_sUserId);
		XML.CreateTag("PASSWORDVALUE", frmScreen.txtOmigaPassword.value);
		XML.CreateTag("UNITID", frmScreen.cboUnitName.options(frmScreen.cboUnitName.selectedIndex).value);
		XML.CreateTag("AUDITRECORDTYPE", "1");
					
		<% // Run server-side asp script %>
		// 		XML.RunASP(document, "ValidateUserLogon.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "ValidateUserLogon.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

					
		if (m_bDebug)
			XML.WriteXMLToFile("C:\\temp\\sc010ValidateUserLogon.xml");
						
		if (XML.IsResponseOK() == true)
		{
			var bAccessAuditAttempt = XML.GetTagBoolean("ACCESSAUDITATTEMPT");
			if (bAccessAuditAttempt == true)
			{
				bLogonOk = true;

				// MAR798 - Windows authentication should never route to password screen.
				// Peter Edney - 15/12/2005
				var xmlFunc = new scXMLFunctions.XMLObject();
				var sSecurityType = xmlFunc.GetGlobalParameterString(document,"SecurityCredentialsType");
				if (sSecurityType.toUpperCase() == "WINDOWSAUTHENTICATION")
					bChangePassword = 0;
				else
					bChangePassword = XML.GetTagBoolean("CHANGEPASSWORDINDICATOR");					
								
				<% /*			
				set context variable id fields 
				(n.b. these are required whether routing is to either SC020 or SC030)
							
				txtDebug.value is set up on screen initialisation
				*/ %>

				objHiddenFrameForm.txtUserId.value = XML.GetTagText("USERID");
				objHiddenFrameForm.txtUserName.value = XML.GetTagText("USERNAME");
				objHiddenFrameForm.txtAccessType.value = XML.GetTagText("ACCESSTYPE");
				objHiddenFrameForm.txtUserCompetency.value = XML.GetTagText("COMPETENCYTYPE");
				<% /* BMIDS678 */ %>
				objHiddenFrameForm.txtCreditCheckAccess.value = XML.GetTagText("CREDITCHECKACCESS");
				//objHiddenFrameForm.txtAccessAuditGUID.value = XML.GetTagText("ACCESSAUDITGUID");
				//AW	EP1240 12/12/06
				var XMLCombo = new scXMLFunctions.XMLObject();
				var sGroupList = new Array("UserRole", "UserQualification");
				
				objHiddenFrameForm.txtIsUserFulfillApproved.value = "0";
				
				var tagQList = XML.SelectTag(null, "QUALIFICATIONLIST");						
				if (tagQList != null)
				{
					<% /*
					RF 09/11/99 Not implemented yet as causes problems due
					to quote marks in xml string being interpreted in html as the 
					end of the variable
					txtQualificationsListXml.value = tagQList.xml;
					*/ %>					
					//AW	EP1240 12/12/06
					var iQualification = "";
					var xmlQualificationNode;
					var xmlQualificationList = XML.XMLDocument.documentElement.selectNodes("//QUALIFICATIONTYPE");
					var iCount;
					
					if(XMLCombo.GetComboLists(document,sGroupList) == true)
					{	
						for(iCount = 0 ; iCount <= xmlQualificationList.length - 1; ++iCount)
						{
							xmlQualificationNode = xmlQualificationList.item(iCount);
							iQualification = xmlQualificationNode.text
							
							if (XMLCombo.IsInComboValidationList(document,"UserQualification",iQualification,["AF"]))
							{
								objHiddenFrameForm.txtIsUserFulfillApproved.value = "1";			
							}
						}
					}
					
					objHiddenFrameForm.txtQualificationsListXml.value = "";
					
				}
				else
				{
					objHiddenFrameForm.txtQualificationsListXml.value = "";
					objHiddenFrameForm.txtIsUserFulfillApproved.value = "";
				}
				
				<% // reset active tag %>
				XML.SelectTag(null,"RESPONSE");

				var iUserRole = XML.GetTagText("ROLE");
				
				if(XMLCombo.GetComboLists(document,sGroupList) == true)
				{	
					if (XMLCombo.IsInComboValidationList(document,"UserRole",iUserRole,["QC"]))
					{
						objHiddenFrameForm.txtIsUserQualityChecker.value = "1";			
					}
					else
					{
						objHiddenFrameForm.txtIsUserQualityChecker.value = "0";
					}
				}
				
				objHiddenFrameForm.txtRole.value = iUserRole;
				//AW	EP1240 12/12/06 - End
				var sUnitId = XML.GetTagText("UNITID");
	
				objHiddenFrameForm.txtUnitId.value = sUnitId;
							
				<% // set up entries from FindUnitsForUser result %>
				XML.LoadXML(m_sFindUnitsForUserXml);
							
				<% // set active tag %>
				XML.SelectTag(null,"RESPONSE");

				//objHiddenFrameForm.txtProcessingIndicator.value = XML.GetTagText("PROCESSINGUNITINDICATOR"); JR BM0271 - moved below
				
				<% // get the department info for the unit %>
				var sSearchPattern = "DEPARTMENTLIST/DEPARTMENT[UNITLIST/UNIT/UNITID='" +  
									 sUnitId + "']";
				XML.ActiveTag = XML.XMLDocument.documentElement.selectSingleNode(sSearchPattern);														
							
				objHiddenFrameForm.txtDepartmentId.value = XML.GetTagText("DEPARTMENTID");
				
				objHiddenFrameForm.txtDistributionChannelId.value = XML.GetTagText("CHANNELID");
				
				//GD New (originally MV 20/12/00)
				objHiddenFrameForm.txtDepartmentName.value = XML.GetTagText("DEPARTMENTNAME");
				//GD New (originally MV 20/12/00)
				objHiddenFrameForm.txtDistributionChannelName.value = XML.GetTagText("CHANNELNAME");
				
				<% /* BS BM0286 20/03/03 */ %>
				var sSearchPattern = "DEPARTMENTLIST/DEPARTMENT/UNITLIST/UNIT[UNITID='" +  
									 sUnitId + "']";
				XML.ActiveTag = XML.XMLDocument.documentElement.selectSingleNode(sSearchPattern);
				<% /* BS BM0286 End 20/03/03 */ %>
				objHiddenFrameForm.txtProcessingIndicator.value = XML.GetTagText("PROCESSINGUNITINDICATOR"); //JR BM0271
			}
		}				
	}
				
	var bReturnArray = new Array(2);
	bReturnArray[0] = bLogonOk;
	bReturnArray[1] = bChangePassword;			
	return bReturnArray;
}

function frmScreen.txtUserId.onblur() 
{
	<% /* SYS0178 Do not FindUnitsForUser if userid not entered */ %>
	if (frmScreen.txtUserId.value != m_sUserId && frmScreen.txtUserId.value != "")
	{
		m_sUserId = frmScreen.txtUserId.value;					
		FindUnitsForUser();					
		frmScreen.txtOmigaPassword.focus();
	}			
}

function FindUnitsForUser()
{
	<% /* Get a list of active units for the user id and populate the combo with it. */ %>

	<% // clear combo details %>
	while (frmScreen.cboUnitName.options.length > 0)
		frmScreen.cboUnitName.options.remove(0);
				
	<%// get unit list%>
				
	var XML = new scXMLFunctions.XMLObject();

	XML.CreateActiveTag("REQUEST");
	XML.SetAttribute("USERID", m_sUserId);
	XML.CreateActiveTag("SEARCH");
	XML.CreateTag("USERID", m_sUserId);
	<% /* PSC 15/11/2005 MAR310 */ %>
	XML.CreateTag("ALLOWOMIGALOGON", "1");
				
	// 	XML.RunASP(document, "FindUnitsForUser.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "FindUnitsForUser.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

			
	if (m_bDebug)
		XML.WriteXMLToFile("C:\\temp\\sc010FindUnitsForUser.xml");
					
	if (XML.IsResponseOK() == true)
	{
		<% // store the xml for use in ValidateUser %>
		m_sFindUnitsForUserXml = XML.XMLDocument.xml;
					
		<% // populate the UnitName combo %>
						
		XML.ActiveTag = null;
		XML.CreateTagList("UNIT");			
		var nNumUnits = XML.ActiveTagList.length;
					
		var sUnitName, sUnitId;
		var TagOPTION;				
					
		if (nNumUnits > 1) 
		{
			<% // First add a <SELECT> option %>
			TagOPTION	= document.createElement("OPTION");
			TagOPTION.value	= "";
			TagOPTION.text	= "<SELECT>";
			frmScreen.cboUnitName.add(TagOPTION);
		}

		if (nNumUnits > 0) 
		{
			<% // RF 25/11/99 AQR SC005: Ensure combo shows all units found %>
			var nUnit = null;
			var sAllowExitFromWrap = null;
			
			for(nUnit = 1; nUnit <= nNumUnits; nUnit++) 
			{
				XML.SelectTagListItem(nUnit - 1)
				sUnitId = XML.GetTagText("UNITID")
				sUnitName = XML.GetTagText("UNITNAME")
				sAllowExitFromWrap = XML.GetTagText("ALLOWOMIGAEXITFROMWRAP")      // MAR324
						
				TagOPTION		= document.createElement("OPTION");
				TagOPTION.value	= sUnitId;
				TagOPTION.text	= sUnitName;
				TagOPTION.setAttribute("AllowExitFromWrap", sAllowExitFromWrap);   // MAR324
				
				frmScreen.cboUnitName.add(TagOPTION);
			}

			<% // Default to the first option %>
			frmScreen.cboUnitName.selectedIndex = 0;
		}				
	}
				
	if (frmScreen.cboUnitName.options.length > 1) 
	{
		<% // enable UnitName combo and give it the focus %>
		scScreenFunctions.SetFieldToWritable(frmScreen, frmScreen.cboUnitName.id)
		frmScreen.cboUnitName.focus();
	}
	else
	{
		<% // disable UnitName combo and give OK button the focus %>
		scScreenFunctions.SetFieldToReadOnly(frmScreen, frmScreen.cboUnitName.id)
		frmScreen.btnOK.focus();
	}
				
	XML = null;
}
/*	AW	12/07/02	BMIDS00178
function IsLegacyCustomer()
{
	return (m_bLegacyCustomer == true);
}

function InitialiseFindLegacyCustomer()
{
	<% // Check the Legacy Customer global parameter %>	
	
	var XML = new scXMLFunctions.XMLObject();
	m_bLegacyCustomer = XML.GetGlobalParameterBoolean(document,"FindLegacyCustomer");
	XML = null;
}
*/
function checkKeyPress()
{
	if( window.event.keyCode == 13)
		Logon();
}
-->
</script>
</body>
</html>
