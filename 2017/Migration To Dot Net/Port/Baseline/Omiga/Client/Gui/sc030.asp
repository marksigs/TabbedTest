<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<% /*
Workfile:      sc030.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Launch Omiga 4 screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date     Description
RF		03/11/99 Created.
RF		18/11/99 Changed from div-based to table-based form.
RF		01/12/99 AQR SC002: Enhancement - Add screen id.
RF		06/12/99 AQR SC011: Enhancement - Add user name and unit id and version info.
RF		28/01/00 Reworked for IE5 and performance.
AY		14/04/00 scScreenFunctions change
MC		11/05/00 SYS0481. Use UnitName in SC030
MC		13/06/00 SYS0481. Increase width of User & Unit Name to allow 45 chars
AP		04/07/00 Improved html screen layout and used latest graphics
MV		14/11/00 Added m_sUnitname Variable and added UnitName to URL in btnSearchClick
MV		05/12/00 CORE00015 - Added m_sDepartmentname ,m_sDistributionChannelName Variable and added to URL in btnSearchClick
GD		20/12/00 SYS1742 changes made to eliminate session variables
				 and allow use of hidden frame.
GD		10/01/01 SYS0203:In window.onload added call to LogonLegallyOpened()
CL		30/03/01 SYS2042: Resize for SC035
PF		06/06/01 SYS2346: Resize Tweaking to avoid scroll bars
JLD		08/10/01 SYS2736: check omPC before continuing
JLD		09/10/01 SYS2737  Find machine name and update context
JLD		19/11/01   TEMPORARILY REMOVE omPC UNTIL IT WORKS
RF		13/11/01 SYS2927: Log out of Admin System on close of screen.
AD		22/11/01 Re-enable the omPC dowload. Now a dll instead of an ocx.
MDC		17/04/02 SYS4396 Customisation of System Name
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
AW		12/07/02	BMIDS00178	Removed Admin System Login
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MV		22/08/2002	BMIDS00355	IE 5.5 upgrade Errors - Modified the HTML layout
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
PSC		05/11/2002	BMIDS00793	Amend to use omPC.cab rather than omPC.dll
MV		27/02/2003	BM0393	slow logons - Amended PCAttributesBO object declaration
AD		05/09/2003	BMIDS634	Re-signed ompc so incremented version number
HMA     08/12/2004  BMIDS957    Remove VBScript
JD		23/03/2005  BMIDS988	New omPC version
JD		30/03/2005  BMIDS988	New omPC version
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :

Prog	Date		AQR			Description
GHun	15/07/2005	MAR7		Integrate local printing. Added axword
GHun	21/07/2005  MAR14		Apply ING Style Sheet and GUI Images
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History :
AS		09/05/2006	EP523 Axword 5.1.0.16 Viewing AFPs and fix to wRTF2PDF02
AS		18/05/2006	EP557 Axword 5.1.0.17 First address line missing on some of the print templates
SAB		07/02/2006	EP2_848 Axword 5.1.0.20
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<html>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="LogonStylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title>Launch Omiga 4</title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets - remove any which are not required */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<% /* MAR7 Not used
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
*/ %>
<object id="PCAttributesBO" style="DISPLAY: none; VISIBILITY: hidden"
classid="CLSID:E31AAD92-37B6-45E1-A84A-481094EBAECE"
	codebase="omPC.CAB#version=1,0,0,4">
<param name="vstrXMLRequest" value=strXML>
</object>
<% /* MAR7 GHun */ %>
<object id="axwordclass" 
	style="DISPLAY: none; VISIBILITY: hidden;" 
	tabindex="-1" 
	classid="CLSID:68575265-E41C-4C2E-808A-CBA63B53D0EF" 
	codebase="axword.CAB#version=5,1,0,20" 
	viewastext>
</object>
<% /* MAR7 End */ %>

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
		<table id="tblMain" width="640" height="385" border="0" cellspacing="0" cellpadding="0">
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
								<b><LABEL id="idTitle">Launch</LABEL></b>
							</td>
							<td height="30" align="right" class="msgVerdanaRight">
								<b>SC030</b>
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
											<td width="10" height="50">&nbsp;</td>
											<td width="70" class="msgLabel">User Name</td>
											<td><input id="txtUserName" name="UserName" maxlength="45" style="WIDTH:200px" class="msgTxt">&nbsp;</td>
										</tr>
										<tr>
											<td width="10" height="30">&nbsp;</td>
											<td width="70" class="msgLabel">Unit Name</td>
											<td><input id="txtUnitName" name="UnitName" maxlength="45" style="WIDTH:200px" class="msgTxt">&nbsp;</td>
										</tr>
										</table>
									</td>
								</tr>	
								</table>
			
								<% // ADP: Navigation buttons outside of the logon fields group %>
								<table border="0" width="350" cellpadding="0" cellspacing="0">
								<tr><td height="5"></td></tr>
								<tr>
									<td width="105px"><input id="btnLaunch" value="Launch Omiga 4" type="button" style="WIDTH:100px" class="msgButton"></td>
									<td><input id="btnAbout" value="About Omiga 4" type="button" style="WIDTH:100px" class="msgButton"></td>
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
<!-- #include FILE="customise/sc030customise.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/sc030Attribs.asp" -->
<% /* SYS4396 MDC 16/04/2002 - End */ %>

<% // Specify Code Here %>
<script language="JScript" type="text/javascript">
<!--	// JScript Code
<%
//GD New - global object points to form on hidden frame
//NB MAINTENANCE - If for any reason the location of the hidden frame/form change
//					then objHiddenFrameForm will need to change.
%>
var objHiddenFrameForm;

var m_sMetaAction = null;
var m_bObjectError = false;
var scScreenFunctions;
		

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
		objHiddenFrameForm = top.frames("fraDetails").document.forms("frmDetails");

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
/*	AW		12/07/02	BMIDS00178		
function window.onbeforeunload()
{
	<% // RF 13/11/01 SYS2927 %>

	ExitOmiga();
}
	
function ExitOmiga()
{
	<% // RF 13/11/01 SYS2927 %>

	var XML = new scXMLFunctions.XMLObject();
	var bLegacyCustomer = XML.GetGlobalParameterBoolean(document,"FindLegacyCustomer");
	XML = null;
		
	if (bLegacyCustomer)
		LogOffAdminSystem();
}
*/
function frmScreen.btnAbout.onclick()
{
	scScreenFunctions.DisplayPopup(window, document, "sc035.asp", "", 470, 420);
}				
function PCAttributesBO.onerror()
{
	m_bObjectError = true;
}
function Initialise()
{
	scScreenFunctions.SetScreenToReadOnly(frmScreen);
	//JLD set launch button based on printer object availability
	if(m_bObjectError == false)
	{
		objHiddenFrameForm.txtMachineId.value =GetComputerName();
		frmScreen.btnLaunch.disabled = false;
	}
	else 
	{
		frmScreen.btnLaunch.disabled = true;
		alert("object omPC has not been installed. Omiga cannot be launched");
	}
	//GD New
	frmScreen.txtUnitName.value =objHiddenFrameForm.txtUnitName.value
		
	//GD New
	frmScreen.txtUserName.value =objHiddenFrameForm.txtUserName.value
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}
/*	AW		12/07/02	BMIDS00178
function LogOffAdminSystem()
{
	<% // RF 13/11/01 Function added for SYS2927 -
	   // call omAdmin.AdministrationInterfaceBO.LogOffUser %>  
	   
	var XML = new scXMLFunctions.XMLObject();
	XML.CreateActiveTag("REQUEST");
	XML.SetAttribute("OPERATION", "LogOffUser");
	XML.SetAttribute("ADMINSYSTEMSTATE", objHiddenFrameForm.txtAdminSystemState.value);
	
	<% // alert(XML.XMLDocument.xml); %>

	<% // Run server-side asp script %>
	XML.RunASP(document, "omAdmin.asp");
		
	<% // alert(XML.XMLDocument.xml); %>

	if (XML.IsResponseOK() == true)
	{
		// delete ADMINSYSTEMSTATE
		objHiddenFrameForm.txtAdminSystemState.value = null
	}
	XML = null;
}
*/
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Description:	
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
function frmScreen.btnLaunch.onclick()
{
	<% // call Welcome Menu %>
	
	<% // set up query string and send it into goOmiga4.asp GD no longer reqd %>
<%
	//Iterate through hidden frame/form looking for nulls - replicates original code
%>
	for(intIndex=0;intIndex < objHiddenFrameForm.length; intIndex++)
	{
		if (objHiddenFrameForm.item(intIndex).value==null)
		{
			alert("Missing " + objHiddenFrameForm.item(intIndex).name + " context value.");
		}
	}	

	var sURL = "modal_om4frameset.asp"
	
	<% // minimum screen resolution 800x600 %>
	<% //SYS2346 Previous iWindowWidth value 792 %>
	var iWindowWidth = 800;
	var iWindowHeight = 552;		
	var iLeft = window.screen.availWidth - iWindowWidth - 10;
	var iTop = window.screen.availHeight - iWindowHeight - 20;
			
	if (iLeft < 0) iLeft = 0;
	else iLeft = iLeft/2;
		
	if (iTop < 0) iTop = 0;
	else iTop = iTop/2;

	var sWidth = ",Width=" + iWindowWidth.toString();
	var sHeight = ",Height=" + iWindowHeight.toString();
	var sTop = ",Top=" + iTop.toString();
	var sLeft = ",Left=" + iLeft.toString();
	<%
	// original attribute were:
	// "toolbar=No,location=No,directories=No,menubar=No,scrollbars=No,
	// resizable=No,status=Yes,Width=792,height=560,top=80,left=80"
	%>			
	<% /* MV-21/08/2002-BMIDS0355*/%>
	var sAttributes = "toolbar=No,location=No,directories=No,menubar=No," +
		"scrollbars=No,resizable=Yes,status=Yes" + sWidth + sHeight + sTop +	sLeft;
				
	window.open(sURL,"", sAttributes);
}		

<% /* BMIDS957  Function converted from VBScript to JScript */ %>
function GetComputerName()
{
	var strOut = "";
	var objOmPC = new ActiveXObject("omPC.PCAttributesBO");
	if (objOmPC != null)
	{
		strOut = objOmPC.NameOfPC();
		objOmPC = null;
	}
	return strOut;

}
-->
</script>
</body>
</html>
