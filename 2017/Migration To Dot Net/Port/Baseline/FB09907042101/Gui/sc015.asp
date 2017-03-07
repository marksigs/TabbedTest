<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<% /*
Workfile:      sc015.asp
Copyright:     Copyright © 2005 Marlborough Stirling

Description:   Launch Main Menu Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
PSC		13/10/2005	Created.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History :
AS		09/05/2006	EP523 Axword 5.1.0.16 Viewing AFPs and fix to wRTF2PDF02
AS		18/05/2006	EP557 Axword 5.1.0.17 First address line missing on some of the print templates
SAB		07/02/2006	EP2_848 Axword 5.1.0.20
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<html>
<head>
	<link href="LogonStylesheet.css" rel="STYLESHEET" type="text/css" />
	<title>Launch Omiga</title>
</head>
<body>

<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<% /* Scriptlets - remove any which are not required */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>

<object id="PCAttributesBO" style="DISPLAY: none; VISIBILITY: hidden"
classid="CLSID:E31AAD92-37B6-45E1-A84A-481094EBAECE"
	 codebase="../../omPC.CAB#version=1,0,0,4">
<param name="vstrXMLRequest" value=strXML>
</object>
<object id="axwordclass" 
	style="DISPLAY: none; VISIBILITY: hidden;" 
	tabindex="-1" 
	classid="CLSID:68575265-E41C-4C2E-808A-CBA63B53D0EF" 
	 codebase="../../axword.CAB#version=5,1,0,20" 
	viewastext>
</object>

<% /* Specify Forms Here */ %>
	
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
					<td width="152" valign="top"><img src="images/mainback.gif" width="152" height="265" alt border="0"></td>	
					<td bgcolor="#FFFFFF" valign="top">
						
						<% // ADP: This table contains three rows, the first to represent the menubar, the second to hold the page title and screen Id in two columns, and the third to hold another table where the logon fields are located %>
						<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
							<tr bgcolor="#FFFFFF">
								<td height="20" colspan="2"></td>
							</tr>
							<tr>
							<td height="30" class="msgVerdanaLeft">
								<% // Screen Title and Id %>
								System Logon
							</td>
							<td height="30" align="right" class="msgVerdanaRight">
								SC015
							</td>
						</tr>
						<tr>
							<td align="center" colspan="2" valign="middle">						
								<table width="350" height="150" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">	
									<tr>
										<td valign="middle" align="center">
											<table cellpadding="0" cellspacing="0">
											<tr>
												<td align="center" id="lblStatus" class="msgLabelBold">Logging onto Omiga. Please wait...</td>
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
	</td>
</tr>
</table>

</form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="customise/sc015customise.asp" -->
<!-- #include FILE="attribs/sc015Attribs.asp" -->

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
var m_sObjectName = "";
var scScreenFunctions;
var m_sFindUnitsForUserXml = "";
var m_sUserId = "";
var m_sPassword = "";
var m_sScreen = "";
		
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
		Customise();
		SetMasks();
		Validation_Init();
		GetLaunchData();
				
		if(!m_bObjectError)
		{
			window.setTimeout(LaunchOmiga, 0);
		}
		else
		{
			lblStatus.innerHTML = "Omiga cannot be launched.<br>The " + m_sObjectName + " object has not been installed.<br>Please contact the system administrator.";
		}
	}
	else
	{
		<%	//The user has accessed this page via a dubious route - send back to logon	%>
		top.location.href="LaunchOmiga.asp";
	}
	
	ClientPopulateScreen();
}
function PCAttributesBO.onerror()
{
	m_bObjectError = true;
	m_sObjectName = "omPC";
}

function axwordclass.onerror()
{
	m_bObjectError = true;
	m_sObjectName = "axWord";
}


function LaunchOmiga()
{
	bReturnArray = ValidateUser();

	if (!bReturnArray[0] || bReturnArray[1])
	{
		lblStatus.innerHTML = "Omiga cannot be launched.<br>The supplied user details cannot be validated.";
		return;
	}
	
	objHiddenFrameForm.txtMachineId.value =GetComputerName();
	
	for(intIndex=0;intIndex < objHiddenFrameForm.length; intIndex++)
	{
		if (objHiddenFrameForm.item(intIndex).value==null)
		{
			alert("Missing " + objHiddenFrameForm.item(intIndex).name + " context value.");
		}
	}	

	var sURL = "";
	
	sURL = "modal_om4frameset.asp?Screen=" + m_sScreen;
	
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
	var sAttributes = "toolbar=No,location=No,directories=No,menubar=No," +
		"scrollbars=No,resizable=Yes,status=Yes" + sWidth + sHeight + sTop +	sLeft;
				
	window.open(sURL,"", sAttributes);
	lblStatus.innerHTML = "Logged into Omiga successfully";
	
}
function GetUnitId(sUserId)
{
	var XML = new scXMLFunctions.XMLObject();
	var sUnitId = "";

	XML.CreateActiveTag("REQUEST");
	XML.SetAttribute("USERID", sUserId);
	XML.CreateActiveTag("SEARCH");
	XML.CreateTag("USERID", sUserId);
				
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "FindUnitsForUser.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}
	
	XML.SelectTag(null, "RESPONSE");
		
	if (XML.GetAttribute("TYPE") == "SUCCESS")
	{
		<% // store the xml for use in ValidateUser %>
		m_sFindUnitsForUserXml = XML.XMLDocument.xml;

		XML.SelectSingleNode("DEPARTMENTLIST/DEPARTMENT/UNITLIST/UNIT[ALLOWOMIGALOGON != '']");
		
		if (XML.ActiveTag != null)
			sUnitId = XML.GetTagText("UNITID");
	}			
				
	return sUnitId;
}


function ValidateUser()
{
	var bLogonOk = false;
	var bChangePassword = false;
	
	var sUnitId = GetUnitId(m_sUserId);
	
	if (sUnitId.length == 0)
		return false;
		
	var XML = new scXMLFunctions.XMLObject();
		
	XML.CreateActiveTag("REQUEST");
	XML.SetAttribute("USERID", m_sUserId);
		
	XML.CreateActiveTag("OMIGAUSER");
	XML.CreateTag("USERID", m_sUserId);
	XML.CreateTag("PASSWORDVALUE", m_sPassword);
	XML.CreateTag("UNITID", sUnitId);
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
	
	XML.SelectTag(null, "RESPONSE");
									
	if (XML.GetAttribute("TYPE") == "SUCCESS")
	{
		var bAccessAuditAttempt = XML.GetTagBoolean("ACCESSAUDITATTEMPT");
		if (bAccessAuditAttempt == true)
		{
			bLogonOk = true;
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
			objHiddenFrameForm.txtCreditCheckAccess.value = XML.GetTagText("CREDITCHECKACCESS");
			
			var tagQList = XML.SelectTag(null, "QUALIFICATIONLIST");						
			if (tagQList != null)
			{
				<% /*
				RF 09/11/99 Not implemented yet as causes problems due
				to quote marks in xml string being interpreted in html as the 
				end of the variable
				txtQualificationsListXml.value = tagQList.xml;
				*/ %>					

				objHiddenFrameForm.txtQualificationsListXml.value = "";
			}
			else
				objHiddenFrameForm.txtQualificationsListXml.value = "";

			<% // reset active tag %>
			XML.SelectTag(null,"RESPONSE");

			objHiddenFrameForm.txtRole.value = XML.GetTagText("ROLE");
					
			objHiddenFrameForm.txtUnitId.value = sUnitId;
							
			<% // set up entries from FindUnitsForUser result %>
			XML.LoadXML(m_sFindUnitsForUserXml);
							
			<% // set active tag %>
			XML.SelectTag(null,"RESPONSE");

			<% // get the department info for the unit %>
			var sSearchPattern = "DEPARTMENTLIST/DEPARTMENT[UNITLIST/UNIT/UNITID='" +  
								 sUnitId + "']";
			XML.ActiveTag = XML.XMLDocument.documentElement.selectSingleNode(sSearchPattern);														
							
			objHiddenFrameForm.txtDepartmentId.value = XML.GetTagText("DEPARTMENTID");
			objHiddenFrameForm.txtDistributionChannelId.value = XML.GetTagText("CHANNELID");
			objHiddenFrameForm.txtDepartmentName.value = XML.GetTagText("DEPARTMENTNAME");
			objHiddenFrameForm.txtDistributionChannelName.value = XML.GetTagText("CHANNELNAME");
			objHiddenFrameForm.txtProcessingIndicator.value = XML.GetTagText("PROCESSINGUNITINDICATOR"); //JR BM0271
		}
	}
			
	var bReturnArray = new Array(2);
	bReturnArray[0] = bLogonOk;
	bReturnArray[1] = bChangePassword;			
	return bReturnArray;		
}
function GetLaunchData()
{
	var XML = new scXMLFunctions.XMLObject();
	XML.LoadXML(window.top.xmlLaunchData.xml);
	XML.SelectTag(null, "REQUEST")
	m_sUserId = XML.GetAttribute("USERID");
	m_sPassword = XML.GetAttribute("PASSWORD");
	m_sScreen = XML.GetAttribute("SCREEN");
	
	if (m_sScreen.toUpperCase() != "CR041.ASP")
		m_sScreen = "MN010.asp";
	
	var xmlCustomerData = XML.SelectTag(XML.ActiveTag, "CUSTOMER")
	
	if (xmlCustomerData != null)
		objHiddenFrameForm.txtLaunchXML.value = xmlCustomerData.xml;	
}
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
