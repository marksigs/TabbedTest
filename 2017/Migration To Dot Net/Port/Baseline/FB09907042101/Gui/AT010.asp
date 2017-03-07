<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AT010.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Application Transfer
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		07/12/00	Created
DJP		04/01/01	SYS1739 Added user authority validation
DJP		05/01/01	SYS1739 Disable Transfer button on entry, and
					after transfer, populate listbox before displaying
					message (Application Transferred)
DJP		15/01/01	SYS1739 Remove disable code from main applicant and app no,
					and don't include current owner in new owner combo.
ADP		15/02/01	SYS1952 Amended submit routing to go to MN070, not MN060
CL		05/03/01	SYS1920 Read only functionality added
LD		23/05/02	SYS4727 Use cached versions of frame functions

BMIDS History:

Prog	Date		Description
MO		15/11/2002	BMIDS00814, Modified to call new method omTM.TransferApplicationOwnership 

MARS History:

Prog	Date		Description
PJO		10/11/2005	MAR479 Disable application number & applicant name fields
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ 
%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<OBJECT data=scTable.htm height=1 id=scTable style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
<span style="LEFT: 310px; POSITION: absolute; TOP: 290px">
	<OBJECT data=scTableListScroll.asp id=scScrollTable style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToMN070" method="post" action="MN070.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" year4 validate="onchange" mark>

<div id="divOwnershipHistory" style="HEIGHT: 260px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="TOP:3px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
		Application Transfer
	</span>
	<span style="LEFT: 24px; POSITION: absolute; TOP: 25px" class="msgLabel">
		Application Number
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtApplicationNumber" maxlength="12" style="WIDTH: 120px" class="msgReadOnly">
		</span>
	</span>
	<span style="LEFT: 270px; POSITION: absolute; TOP: 25px" class="msgLabel">
		Main Applicant
		<span style="LEFT: 78px; POSITION: absolute; TOP: -3px">
			<input id="txtMainApplicant" maxlength="12" style="WIDTH: 200px" class="msgReadOnly">
		</span>
	</span>
	<span style="TOP:50px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
		Ownership History
	</span>

<span id="spnTable" style="TOP: 70px; LEFT: 4px; POSITION: ABSOLUTE">
	<table id="tblTable" width="596px" border="0" cellspacing="0" cellpadding="0" class="msgTable">
	<tr id="rowTitles">
		<td width="30%" class="TableHead">Start Date</td>	
		<td width="30%" class="TableHead">User Name</td>
		<td width="15%" class="TableHead">Unit ID</td>		
		<td width="20%" class="TableHead">Unit Name</td>
	</tr>
	<tr id="row01">		
		<td class="TableTopLeft">&nbsp;</td>		
		<td class="TableTopCenter">&nbsp;</td>		
		<td class="TableTopCenter">&nbsp;</td>
		<td class="TableTopRight">&nbsp;</td>
	</tr>
	<tr id="row02">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row03">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row04">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row05">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row06">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row07">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row08">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	 <tr id="row09">		
		<td class="TableLeft">&nbsp;</td>		
		<td class="TableCenter">&nbsp;</td>			
		<td class="TableCenter">&nbsp;</td>
		<td class="TableRight">&nbsp;</td>
	</tr>
	<tr id="row10">		
		<td class="TableBottomLeft">&nbsp;</td>		
		<td class="TableBottomCenter">&nbsp;</td>			
		<td class="TableBottomCenter">&nbsp;</td>
		<td class="TableBottomRight">&nbsp;</td>
	</tr>
	</table>
</span>

</div>
<div id="divTransferDetails" style="HEIGHT: 120px; LEFT: 10px; POSITION: absolute; TOP: 330px; WIDTH: 604px" class="msgGroup">
	<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
		Transfer Details to...
	</span>
<span style="LEFT: 24px; POSITION: absolute; TOP: 30px" class="msgLabel">
	New Unit No.
	<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
		<select id="cboNewUnitNo" style="WIDTH: 190px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 24px; POSITION: absolute; TOP: 90px" class="msgLabel">
	New Owner
	<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
		<select id="cboNewOwner" style="WIDTH: 190px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 24px; POSITION: absolute; TOP: 60px" class="msgLabel">
	New Unit Name
	<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
		<input id="txtNewUnitName" maxlength="12" style="WIDTH: 190px" class="msgReadOnly">
	</span>
</span>

<span style="TOP:60px; LEFT:450px; POSITION:ABSOLUTE">
	<input id="btnTransferOwnership" value="Transfer Ownership" type="button" style="WIDTH:110px" class="msgButton">
</span>
</div>

</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 460px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AT010attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var scScreenFunctions;
var m_sDistributionChannel;
var m_iTableLength = 10;
var userHistoryXML;
var m_sApplicationNumber;
var m_sApplicantName;
var m_sApplicationFactFindNumber;
var m_TM020XML;
var m_nReadOnly;
var m_nUserRole;
var m_sCurrentOwner;
var m_sCurrentUnit;
var m_blnReadOnly = false;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit");

	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Application Transfer","AT010",scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

/*rmScreen.UserName.disabled = true; 
	frmScreen.btnSelect.disabled =true;
	frmScreen.btnEdit.disabled =true;
*/
	RetrieveContextData();
	
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();

	PopulateScreen();
	
	frmScreen.txtNewUnitName.disabled = true;	
	frmScreen.cboNewOwner.disabled = true;
	frmScreen.btnTransferOwnership.disabled = true;
	EnableTransfer();

 	scScreenFunctions.SetFocusToFirstField(frmScreen);
	ValidateUserAuthority();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AT010");
	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function ValidateUserAuthority()
{
	var sValid;

	if( m_nReadOnly != "1" )
	{
		var authXML = null;

		authXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

		authXML.CreateRequestTag(window , null);
		authXML.CreateActiveTag("VALIDATION");
		authXML.CreateTag("USERROLE",m_nUserRole);
		authXML.CreateTag("AUTHORITYREQUIREMENT","AppTfOrgUserRole");

		authXML.RunASP(document, "ValidateUserAuthority.asp")

		if(authXML.IsResponseOK())
		{
			//userXML.SelectTag(null, "VALIDITY");
			sValid = authXML.GetTagText("VALIDITY");
			if( sValid != "1" )
			{
				m_nReadOnly = "1"
			}
		}
	}
	
	if( m_nReadOnly == "1" )
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);	
	}
	
}

function EnableTransfer()
{
	var strUserID;
	var strUnitID;
	
	strUserID = frmScreen.cboNewOwner.value;
	strUnitID = frmScreen.cboNewUnitNo.value;
	
	if( strUserID.length > 0 && strUnitID .length > 0 )
	{
		frmScreen.btnTransferOwnership.disabled =false;
	}
	else
	{
		frmScreen.btnTransferOwnership.disabled =true;	
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

function RetrieveContextData()
{
	m_sApplicantName = scScreenFunctions.GetContextParameter(window,"idCustomerName1","");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sDistributionChannel = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId","");
	m_nReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","");
	m_nUserRole = scScreenFunctions.GetContextParameter(window,"idRole","");
}

function SetUserCombo( strUserID )
{
	if(strUserID.length > 0 )
	{
		frmScreen.cboNewOwner.disabled = false;
		frmScreen.cboNewOwner.value = strUserID;
	}
}
function SetUnitCombo( strUnitID )
{
	frmScreen.cboUnitName.value = strUnitID;
}

function PopulateScreen()
{
<%	//set the due date to today
%>	
	var bSuccess = false;
	frmScreen.txtApplicationNumber.value  = m_sApplicationNumber;
	frmScreen.txtMainApplicant.value  = m_sApplicantName;
	frmScreen.txtApplicationNumber.disabled =true;
	frmScreen.txtMainApplicant.disabled =true;	
	
	bSuccess = PopulateUserHistory()

	if(bSuccess)
	{
		PopulateUnit();
	}
}

function PopulateUser()
{
	var userXML = null;
	var sUnitID;
	var sFile;
	var bDisableUserCombo;

	bDisableUserCombo = true;

	userXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	while(frmScreen.cboNewOwner.options.length > 0) frmScreen.cboNewOwner.remove(0);

	// Get the UnitID from the unit combo
	sUnitID = frmScreen.cboNewUnitNo.options(frmScreen.cboNewUnitNo.selectedIndex).text;
	frmScreen.txtNewUnitName.value = frmScreen.cboNewUnitNo.value  

	if( sUnitID.length > 0 )
	{
		userXML.CreateRequestTag(window , null);
		userXML.CreateActiveTag("USERLIST");
		userXML.CreateTag("UNITID",sUnitID);
		// 		userXML.RunASP(document, "FindUserList.asp")
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					userXML.RunASP(document, "FindUserList.asp")
				break;
			default: // Error
				userXML.SetErrorResponse();
			}


		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = userXML.CheckResponse(ErrorTypes);
			
		if ( ErrorReturn[0] == true )
		{
			var strForeName;
			var strSurname;
			var strUserID;
			var intCounter;
			var intCount;

			userXML.CreateTagList("OMIGAUSER");
			intCount = userXML.ActiveTagList.length;

			if( intCount > 0 )
			{
				bDisableUserCombo = false;
				AddComboOption( frmScreen.cboNewOwner, "<SELECT>","" );
				
				for( intCounter = 0; intCounter < intCount; intCounter++ )
				{
					userXML.SelectTagListItem( intCounter );
					strForeName = userXML.GetTagText("USERFORENAME");
					strSurname = userXML.GetTagText("USERSURNAME");
					strUserID = userXML.GetTagText("USERID");
					
					if( m_sCurrentOwner != strUserID )
					{
						AddComboOption( frmScreen.cboNewOwner , strForeName + " " + strSurname, strUserID );
					}
				}
			}
		}
	}
	
	// Enable user combo
	frmScreen.cboNewOwner.disabled = bDisableUserCombo;		
}

function PopulateUserHistory()
{
	var bSuccess;
	bSuccess = false;
	userHistoryXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	userHistoryXML.CreateRequestTag(window , null);
	userHistoryXML.CreateActiveTag("APPLICATIONOWNERSHIP");
	userHistoryXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);

	// 	userHistoryXML.RunASP(document, "FindApplicationOwnership.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			userHistoryXML.RunASP(document, "FindApplicationOwnership.asp");
			break;
		default: // Error
			userHistoryXML.SetErrorResponse();
		}


	if(userHistoryXML.IsResponseOK())
	{
		PopulateListBox();
		bSuccess = true;
	}

	EnableTransfer();
	return(bSuccess);
}

function PopulateListBox()
{
	var iCount;

	userHistoryXML.CreateTagList("USERHISTORY");
	iCount = userHistoryXML.ActiveTagList.length;
	
	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iCount);
	ShowList(0);	
}

function ShowList(nStart)
{
	var strListItem;
	var strForeName;
	var strSurname;
	var nCount;
	
	scScrollTable.clear();
	nCount = userHistoryXML.ActiveTagList.length;

	if( nCount > 0 )
	{
		userHistoryXML.SelectTagListItem(0);
		m_sCurrentOwner = userHistoryXML.GetTagText("USERID");
		m_sCurrentUnit = userHistoryXML.GetTagText("UNITID");
	}

	for (var iCount = 0; iCount < nCount && iCount < m_iTableLength; iCount++)
	{
		userHistoryXML.SelectTagListItem(iCount + nStart);
		strListItem = userHistoryXML.GetTagText("USERHISTORYDATE");
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), strListItem);

		strForeName = userHistoryXML.GetTagText("USERFORENAME");
		strSurname = userHistoryXML.GetTagText("USERSURNAME");
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), strForeName + " " + strSurname);

		strListItem = userHistoryXML.GetTagText("UNITID");
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), strListItem);

		strListItem = userHistoryXML.GetTagText("UNITNAME");
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), strListItem);

		tblTable.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));
	}
}


function AddComboOption(cboToAdd, strName, strValue)
{
	var TagOPTION = document.createElement("OPTION");
	TagOPTION = document.createElement("OPTION");
	TagOPTION.value	= strValue;
	TagOPTION.text = strName;
	cboToAdd.add(TagOPTION);
}

function window.onunload ()
{
}

function btnSubmit.onclick()
{
<% // An example submit function showing the use of the validation functions 
%>	
	frmToMN070.submit();
}

function PopulateUnit()
{
	//
	var unitXML = null;
	unitXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	unitXML.CreateRequestTag(window , null);
	unitXML.CreateActiveTag("UNIT");
	unitXML.CreateTag("CHANNELID",m_sDistributionChannel);

	unitXML.RunASP(document, "FindUnitList.asp");

	if(unitXML.IsResponseOK())
	{
		var strUnitName;
		var strUnitID;
		var intCounter;
		var intCount;

		unitXML.CreateTagList("UNIT");
		intCount = unitXML.ActiveTagList.length;
		
		if( intCount > 0 )
		{
			AddComboOption( frmScreen.cboNewUnitNo, "<SELECT>","" );

			for( intCounter = 0; intCounter < intCount; intCounter++ )
			{
				unitXML.SelectTagListItem( intCounter );
				strUnitID = unitXML.GetTagText("UNITID");
				strUnitName = unitXML.GetTagText("UNITNAME");

				AddComboOption( frmScreen.cboNewUnitNo, strUnitID, strUnitName  );
			}

			//Set the selection to the first one
			frmScreen.cboNewUnitNo.selectedIndex = 0;
		}
	}
}

function frmScreen.btnTransferOwnership.onclick()
{
	TransferApplication();
}

<% /* MO 15/11/2002 BMIDS00814 - Modified greatly!*/ %>
function TransferApplication()
{
	var transferAppXML = null;
	var strVal;
	var strActivityId = scScreenFunctions.GetContextParameter(window,"idActivityId","1");
	var strOldUserId = "";
	var strOldUnitId = "";
	
	transferAppXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	transferAppXML.CreateRequestTag(window , "TRANSFERAPPLICATIONOWNERSHIP");
	<% /*
	
	MO 15/11/2002 BMIDS00814 - Removed
	
	transferAppXML.CreateActiveTag("USERHISTORY");

	transferAppXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);

	strVal = frmScreen.cboNewUnitNo.options(frmScreen.cboNewUnitNo.selectedIndex).text;
	transferAppXML.CreateTag("UNITID",strVal);

	strVal = frmScreen.cboNewOwner.value
	transferAppXML.CreateTag("USERID",strVal);
	
	*/ %>
	transferAppXML.CreateActiveTag("CASEACTIVITY");
	transferAppXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	transferAppXML.SetAttribute("ACTIVITYID", strActivityId);
	transferAppXML.SetAttribute("CASEID", m_sApplicationNumber);
	
	transferAppXML.SelectTag(null, "REQUEST");
	
	transferAppXML.CreateActiveTag("USERHISTORY");
	transferAppXML.SetAttribute("OLDUSERID", m_sCurrentOwner);
	transferAppXML.SetAttribute("OLDUNITID", m_sCurrentUnit);
	transferAppXML.SetAttribute("NEWUSERID", frmScreen.cboNewOwner.value);
	transferAppXML.SetAttribute("NEWUNITID",  frmScreen.cboNewUnitNo.options(frmScreen.cboNewUnitNo.selectedIndex).text);
	
	// 	transferAppXML.RunASP(document, "CreateUserHistory.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			//transferAppXML.RunASP(document, "CreateUserHistory.asp");
			transferAppXML.RunASP(document, "OmigaTMBO.asp");
			break;
		default: // Error
			transferAppXML.SetErrorResponse();
		}


	if(transferAppXML.IsResponseOK())
	{
		PopulateUserHistory();
		alert("Application transferred successfully");
	}
	PopulateUser();
	frmScreen.btnTransferOwnership.disabled = true;
}

function frmScreen.cboNewOwner.onchange()
{
	EnableTransfer();
}

function frmScreen.cboNewUnitNo.onchange()
{
	PopulateUser();
	EnableTransfer();
}


function FindCustomersForApplication()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	m_sApplicationNumber = userHistoryXML.GetAttribute("CASEID");	
	m_sApplicationFactFindNumber = 1;
	
	XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.RunASP(document,"FindCustomersForApplication.asp");

	if (XML.IsResponseOK()) {

		XML.CreateTagList("CUSTOMERROLE");
		
		var sCustomerNumber = "";
		var sCustomerVersionNumber = "";
		var sCustomerRoleType = "";
		var sCustomerOrder = "";
		var sSurname = "";
		var sForename = "";
		var iCustomerIndex = 0;
		
		for (var i0=0; i0<5; i0++)
		{
			sCustomerNumber = "";
			sCustomerVersionNumber = "";
			sCustomerRoleType = "";
			sCustomerOrder = "";
			sSurname = "";
			sForename = "";
			
			if (i0 < XML.ActiveTagList.length){
				if (XML.SelectTagListItem(i0) == true){
					sCustomerNumber = XML.GetTagText("CUSTOMERNUMBER");
					sCustomerVersionNumber = XML.GetTagText("CUSTOMERVERSIONNUMBER");
					sCustomerRoleType = XML.GetTagText("CUSTOMERROLETYPE");
					sCustomerOrder = XML.GetTagText("CUSTOMERORDER");				
					sSurname = XML.GetTagText("SURNAME");
					sForename = XML.GetTagText("FIRSTFORENAME");		
				}
			}
			iCustomerIndex = i0 + 1;						
			scScreenFunctions.SetContextParameter(window,"idCustomerName" + iCustomerIndex, sForename + " " + sSurname);
			scScreenFunctions.SetContextParameter(window,"idCustomerNumber" + iCustomerIndex, sCustomerNumber);
			scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber" + iCustomerIndex, sCustomerVersionNumber);
			scScreenFunctions.SetContextParameter(window,"idCustomerRoleType" + iCustomerIndex, sCustomerRoleType);
			scScreenFunctions.SetContextParameter(window,"idCustomerOrder" + iCustomerIndex, sCustomerOrder);
		}
	}
}
-->
</script>
</body>
</html>


