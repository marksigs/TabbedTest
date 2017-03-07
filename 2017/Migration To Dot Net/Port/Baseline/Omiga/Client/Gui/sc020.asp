<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<% /*
Workfile:      sc020.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Change password screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		03/11/99	Created.
RF		18/11/99	Changed from div-based to table-based form.
RF		01/12/99	AQR SC002: Add screen id.
RF		13/12/99	AQR SC025: Ensure context is set for SC030.
RF		28/01/00	Reworked for IE5 and performance.
AY		03/04/00	scScreenFunctions change
MC		11/05/00	SYS0481. Use UnitName in SC030
AP		04/07/00	Improved html creen layout and used latest graphics
GD		21/12/00	SYS1742 Changed to allow elimination of session variables
GD		10/01/01	SYS0203:In window.onload added call to LogonLegallyOpened()
PSC		21/02/01	SYS1970 Amend change password to set unit id correctly
GD		23/02/01	SYS1970 Removed code from onload (originally for SYS0203)
					to allow stable version to go into build. Code in
					onload should not have been put back into dev because
					SYS0203 is not completed.
TW      09/10/02	SYS5115   Modified to incorporate client validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :

Prog	Date		AQR			Description
GHun	21/07/2005	MAR14		Apply ING Style Sheet and GUI Images
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<html>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="LogonStylesheet.css" rel="STYLESHEET" type="text/css">
	<title>Change password</title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets - remove any which are not required */ %>
<script src="validation.js" language="JScript"></script>
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>

<% /* Specify Forms Here */ %>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% // Specify Screen Layout Here %>
<form id="frmScreen" mark validate="onchange">
	
<% // ADP: Outer table ensures everything inside is centrally aligned both horizontally and vertically %>
<table width="100%" height="90%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td align="center" valign="middle">	
	
		<% // ADP: This table defines the overall size of the screen area %>	
		<table id="tblMain" width="640" height="385" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
		<tr>
			<td align="center">
		
				<% // ADP: Another table divides the screen into two columns. The first holds the MainBack.gif graphic and the second contains the title bar and fields %>
				<table width="100%" height="100%" border="2" bordercolor="#616161" cellspacing="0" cellpadding="0">
				<tr>
					<td width="152" valign="top"><img src="images/mainback.gif" width="152" height="265" alt="" border="0"></td>	
					<td bgcolor="#FFFFFF" valign="top">
						
						<% // ADP: This table contains three rows, the first to represent the menubar, the second to hold the page title and screen Id in two columns, and the third to hold another table where the fields are located %>
						<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
						<tr bgcolor="#FFFFFF">
							<td height="20" colspan="2"></td>
						</tr>
						<tr>
							<td height="30" class="msgVerdanaLeft">
								<% // Screen Title and Id %>
								<b>Change Password</b>
							</td>
							<td height="30" align="right" class="msgVerdanaRight">
								<b>SC020</b>
							</td>
						</tr>
						<tr>
							<td align="center" colspan="2" valign="middle">
							
								<% // ADP: Here's the table containing the fields %>
								<table width="350" height="150" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">	
								<tr>
									<td valign="middle">
										<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td width="30" height="30">&nbsp;</td>
											<td width="140" class="msgLabel">Current Password</td>
											<td width="150"><input id="txtCurPwd" type=PASSWORD name="CurPwd" maxlength="15" style="WIDTH:100px" class="msgTxt">&nbsp;</td>
										</tr>
										<tr>
											<td width="30" height="30">&nbsp;</td>
											<td width="140" class="msgLabel">New Password</td>
											<td width="150"><input id="txtNewPwd" type=PASSWORD name="NewPwd" maxlength="15" style="WIDTH:100px" class="msgTxt">&nbsp;</td>
										</tr>
										<tr>
											<td width="30" height="30">&nbsp;</td>
											<td width="140" class="msgLabel">Confirm New Password</td>
											<td width="150"><input id="txtConfirmedNewPwd" type=PASSWORD name="ConfirmedNewPwd" maxlength="15" style="WIDTH:100px" class="msgTxt">&nbsp;</td>
										</tr>
										</table>
									</td>
								</tr>	
								</table>
			
								<% // ADP: OK and Change Password buttons outside of the logon fields group %>
								<table border="0" width="350" cellpadding="0" cellspacing="0">
								<tr><td height="5"></td></tr>
								<tr>
									<td width="65"><input id="btnSubmit" value="OK" type="button" style="WIDTH: 60px" class="msgButton"></td>
									<td><input id="btnCancel" value="Cancel" type="button" style="WIDTH: 60px" class="msgButton"></td>
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

<% // File containing field attributes %>
<!--  #include FILE="attribs/sc020attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!-- // JScript Code

var m_sMetaAction = null;
var scScreenFunctions;
<%
//GD New - global object points to form on hidden frame
//NB MAINTENANCE - If for any reason the location of the hidden frame/form change
//					then objHiddenFrameForm will need to change.
%>
var objHiddenFrameForm;
		

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	objHiddenFrameForm = top.frames("fraDetails").document.forms("frmDetails");
	SetMasks();
	Validation_Init();			
	Initialise();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
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
	<%  
	//set context variable id fields to null - 
	//this prevents potential security problems if  browser
	//"back" and "forward" buttons are used
	%>
			
}

function frmScreen.btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
		ChangePassword();
}
		
function frmScreen.btnCancel.onclick()
{
<%	
	// Route to SC010 
	// GD frmGoPreviousScreen.submit();
	//GD New
	//Leave this objHiddenFrameForm.txtDebug.value = m_bDebug;
%>
	objHiddenFrameForm.txtUserId.value = null;
	objHiddenFrameForm.txtUserName.value = null;
	objHiddenFrameForm.txtAccessType.value = null;
	objHiddenFrameForm.txtUserCompetency.value = null;
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

	top.document.frames("fraMain").location.href="sc010.asp";
	
}

function ChangePassword()
{
	<% // Routes to SC030 
	%>
	var bContinue = true;
			
	if (frmScreen.txtNewPwd.value != frmScreen.txtConfirmedNewPwd.value)
	{
		alert("Confirmed new password does not match new password entered");
		frmScreen.txtConfirmedNewPwd.focus();
		bContinue = false;
	}
			
	if (bContinue)
	{
		var XML = new scXMLFunctions.XMLObject();

		XML.CreateActiveTag("REQUEST");
		XML.SetAttribute("USERID", objHiddenFrameForm.txtUserId.value);
		
		
		XML.CreateActiveTag("OMIGAUSER");
		XML.CreateTag("USERID", objHiddenFrameForm.txtUserId.value);
		
		XML.CreateTag("CURRENTPASSWORD", frmScreen.txtCurPwd.value);
		XML.CreateTag("NEWPASSWORD", frmScreen.txtNewPwd.value);
		XML.CreateTag("CONFIRMEDNEWPASSWORD", frmScreen.txtConfirmedNewPwd.value);
		XML.CreateTag("UNITID", objHiddenFrameForm.txtUnitId.value);
		
			
		// 		XML.RunASP(document, "ChangePassword.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "ChangePassword.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

			
		if (objHiddenFrameForm.txtDebug.value == true)
			XML.WriteXMLToFile("C:\\temp\\sc020ChangePassword.xml");
				
		if (XML.IsResponseOK() == true)
			GoNextScreen();
	}
}			

	//Description:	
function GoNextScreen()
{
	<% /* Routes to SC030 LaunchOmiga4 */ %>
	
	<% /* 
		set context variable id fields; 						
		n.b.txtDebug.value is set up on screen initialisation
	*/ %>

	top.document.frames("fraMain").location.href="sc030.asp";
}

-->
</script>
</body>
</html>


