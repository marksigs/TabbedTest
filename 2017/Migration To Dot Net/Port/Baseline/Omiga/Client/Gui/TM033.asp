<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      TM033.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Edit Action
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		30/10/00	Created (screen paint)
JLD		16/11/00	Implementation added (not complete)
JLD		22/11/00	Added BO calls
JLD		03/01/01	SYS1778 check the time along with dates too.
CL		09/01/01	SYS1910 Revision for Date change
CL		05/03/01	SYS1920 Read only functionality added
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
DRC     17/07/01    SYS2208 Fixed logic bug in code 
DRC     17/04/02    SYS1853 Added the TASKSTATUSSETBY paramters to the call to update casetask					
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:
******** NOTE THAT TM033P.asp WILL ALSO NEED THE SAME CHANGES APPLIED TO IT AS THIS SCREEN, AS IT'S ESSENTIALLY THE SAME*****
Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
GD		10/07/2003	   Removed checks for non-working day.
HMA     24/11/2004     BMIDS725  Do not allow new unit to be selected without new user and vice versa.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
PSC     2209/2005	MAR32		Add processing for Sales Agents
Maha T	15/11/2005	MAR319		Disable newunit combo And show workgroup users in newowner combo,
								if current loggged on user is SA and userole < global parm.
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

<% /* Specify Forms Here */ %>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM020" method="post" action="TM020.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divOwnerDetails" style="TOP: 60px; LEFT: 10px; HEIGHT: 170px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
	Owner Details...
</span>
<span style="TOP:35px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Existing Owner
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtExistingOwner" maxlength="10" style="WIDTH:250px" class="msgTxt">
	</span>
</span>
<span style="TOP:60px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	New Unit
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<select id="cboNewUnit" style="WIDTH:250px" class="msgCombo"></select>
	</span>
</span>
<span style="TOP:85px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	New Owner
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<select id="cboNewOwner" style="WIDTH:250px" class="msgCombo"></select>
	</span>
</span>
<span style="TOP:110px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	User Id
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtUserId" maxlength="10" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
<span style="TOP:135px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Unit Code
	<span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		<input id="txtUnitCode" maxlength="10" style="WIDTH:150px" class="msgTxt">
	</span>
</span>
</div>
<div id="divDueDate" style="TOP: 240px; LEFT: 10px; HEIGHT: 95px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
<span style="TOP:10px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabelHead">
	Due Date...
</span>
<span style="TOP:35px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	Existing Due Date/Time
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtExistingDueDate" maxlength="10" style="WIDTH:100px" class="msgTxt">
	</span>
	<span style="TOP:-3px; LEFT:223px; POSITION:ABSOLUTE">
		<input id="txtExistingDueTime" maxlength="10" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
<span style="TOP:60px; LEFT:4px; POSITION:ABSOLUTE" class="msgLabel">
	New Due Date/Time
	<span style="TOP:-3px; LEFT:120px; POSITION:ABSOLUTE">
		<input id="txtNewDueDate" maxlength="10" style="WIDTH:100px" class="msgTxt">
	</span>
	<span style="TOP:-3px; LEFT:223px; POSITION:ABSOLUTE">
		<input id="txtNewDueTime" maxlength="8" style="WIDTH:100px" class="msgTxt">
	</span>
</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 340px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/TM033attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_sTaskXML = "";
var taskXML = null;
var m_sExistingOwner = "";
var m_sExistingDueDate = "";
var m_sExistingOwnerType = "";
var m_sDistributionChannel = "";
var m_sUnitId = "";
var m_sUnitName = "";
var scScreenFunctions;

var bDifferentOwners = false;
var bDifferentTypes = false;
var m_blnReadOnly = false;

<% /* PSC 22/09/2005 MAR32 - Start */ %>
var m_nUserRole = "-1";
var m_nMinSALevel = 0;
var m_bIsSARole = false;
<% /* PSC 22/09/2005 MAR32 - End */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Cancel");

	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Edit Action","TM033",scScreenFunctions);

	RetrieveContextData();
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
	PopulateScreen();
	
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

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	m_sDistributionChannel = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId",null);
	m_sUnitId = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	<% /* PSC 22/09/2005 MAR32 */ %>
	m_nUserRole = parseInt(scScreenFunctions.GetContextParameter(window,"idRole","-1"));

}
function PopulateScreen()
{
	scScreenFunctions.SetFieldState(frmScreen, "txtExistingOwner", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtUserId", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtUnitCode", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtExistingDueDate", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtExistingDueTime", "R");
	
	<% /* PSC 22/09/2005 MAR32 - Start */ %>
	var GlobalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_nMinSALevel = parseInt(GlobalXML.GetGlobalParameterAmount(document,"TMMinSAAuthLevel"));

	var userRoleXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bIsSARole = userRoleXML.IsInComboValidationList(document,"UserRole", m_nUserRole, ["SA"]);
	<% /* PSC 22/09/2005 MAR32 - End */ %>
	
	if(m_sMetaAction == "TM030")
		scScreenFunctions.SetFieldState(frmScreen, "cboNewUnit", "R");
	
	taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	taskXML.LoadXML(m_sTaskXML);
	taskXML.SelectTag(null, "REQUEST");
	taskXML.CreateTagList("CASETASK");
	taskXML.SelectTagListItem(0);
	m_sExistingOwner = taskXML.GetAttribute("OWNINGUSERID");
	m_sExistingDueDate = taskXML.GetAttribute("TASKDUEDATEANDTIME");
	if( m_sExistingDueDate == null )
		m_sExistingDueDate = "";

	m_sExistingOwnerType = taskXML.GetAttribute("TASKOWNERTYPE");
	bDifferentOwners = false;
	bDifferentTypes = false;
	if(taskXML.ActiveTagList.length > 1)
	{
		for(var iCount = 0; iCount < taskXML.ActiveTagList.length && bDifferentOwners == false; iCount++)
		{
			taskXML.SelectTagListItem(iCount);
			if(taskXML.GetAttribute("OWNINGUSERID") != m_sExistingOwner ||
			   taskXML.GetAttribute("TASKDUEDATEANDTIME") != m_sExistingDueDate )
					bDifferentOwners = true;
			if(taskXML.GetAttribute("TASKOWNERTYPE") != m_sExistingOwnerType)
					bDifferentTypes = true;
		}
	}
	if(!bDifferentOwners)
	{
		frmScreen.txtExistingOwner.value = m_sExistingOwner;
		frmScreen.txtExistingDueDate.value = m_sExistingDueDate.substr(0,10);
		frmScreen.txtExistingDueTime.value = m_sExistingDueDate.substr(11,8);
	}
	
	PopulateNewUnitCombo();
	if(!bDifferentTypes)
		PopulateNewOwnerCombo();
	frmScreen.txtUnitCode.value = m_sUnitName;
}
function PopulateNewUnitCombo()
{
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window , "");
	XML.CreateActiveTag("UNIT");
	XML.CreateTag("CHANNELID", m_sDistributionChannel);
	XML.RunASP(document, "FindUnitList.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[1] == ErrorTypes[0])
			alert("No Units could be found for the logged on Distribution Channel.");
		else
		{
			var nSelect = -1;
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= "";
			TagOPTION.text = "<SELECT>";
			frmScreen.cboNewUnit.add(TagOPTION);
			
			XML.CreateTagList("UNIT");
			for(var iCount = 0; iCount < XML.ActiveTagList.length; iCount++)
			{
				XML.SelectTagListItem(iCount);
				TagOPTION = document.createElement("OPTION");
				TagOPTION.value	= XML.GetTagText("UNITID");
				TagOPTION.text =  XML.GetTagText("UNITNAME");
				frmScreen.cboNewUnit.add(TagOPTION);
				if(m_sUnitId == XML.GetTagText("UNITID")) 
				{
					nSelect = iCount+1;
					m_sUnitName = TagOPTION.text;
				}
			}
			//select the current users unit
			frmScreen.cboNewUnit.selectedIndex = nSelect;
		}
	}
	
	<% /* START: MAR319 Maha */ %>
	if (m_bIsSARole && m_nUserRole < m_nMinSALevel)
		frmScreen.cboNewUnit.disabled = true;
	else
		frmScreen.cboNewUnit.disabled = false;	
	<% /* END: MAR319 */ %>	
}
function frmScreen.cboNewUnit.onchange()
{
	if(!bDifferentTypes)
		PopulateNewOwnerCombo();

	frmScreen.txtUnitCode.value = frmScreen.cboNewUnit.options(frmScreen.cboNewUnit.selectedIndex).text;
	frmScreen.txtUserId.value = "";
}
function PopulateNewOwnerCombo()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bDisableOwnerCombo;
	bDisableOwnerCombo = true;
	
	while(frmScreen.cboNewOwner.options.length > 0) frmScreen.cboNewOwner.remove(0);
	
	XML.CreateRequestTag(window , "");
	XML.CreateActiveTag("USERLIST");
	if(frmScreen.cboNewUnit.value == "")
		XML.CreateTag("UNITID",m_sUnitId);
	else
		XML.CreateTag("UNITID",frmScreen.cboNewUnit.value);
	XML.CreateTag("USERROLE", m_sExistingOwnerType);
	XML.RunASP(document, "FindUserList.asp")

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
			
	if ( ErrorReturn[0] == true )
	{
		bDisableOwnerCombo = false;

		<% /* PSC 22/09/2005 MAR32 - Start */ %>		
		if (m_bIsSARole && m_nUserRole < m_nMinSALevel)
		{
			XML.SelectNodes("OMIGAUSERLIST/OMIGAUSER[WORKGROUPUSER='1']");
		}
		else
			XML.SelectNodes("OMIGAUSERLIST/OMIGAUSER");
		<% /* PSC 22/09/2005 MAR32 - End */ %>
		
		TagOPTION = document.createElement("OPTION");
		TagOPTION.value	= "";
		TagOPTION.text = "<SELECT>";
		frmScreen.cboNewOwner.add(TagOPTION);
		for(var iCount = 0; iCount < XML.ActiveTagList.length; iCount++)
		{
			XML.SelectTagListItem(iCount);
			TagOPTION = document.createElement("OPTION");
			TagOPTION.value	= XML.GetTagText("USERID");
			TagOPTION.text =  XML.GetTagText("USERID");
			frmScreen.cboNewOwner.add(TagOPTION);
		}
	}

	// Enable/disable New Owner combo
	frmScreen.cboNewOwner.disabled = bDisableOwnerCombo;		

}
function frmScreen.cboNewOwner.onchange()
{
	frmScreen.txtUserId.value = frmScreen.cboNewOwner.value;
}
function GetNewDueDateTime()
{
	if(frmScreen.txtNewDueDate.value.length == 0)
		return "";
		
	var sDateTime = frmScreen.txtNewDueDate.value + " " + frmScreen.txtNewDueTime.value;
	return sDateTime;
}
function CheckTimeFields()
{
	var bSuccess = true;
	var sTime = frmScreen.txtNewDueTime.value;
	var sHours = sTime.substr(0,2);
	var sMins = sTime.substr(3,2);
	var sSecs = sTime.substr(6,2);
	if(sTime.length == 0)
		frmScreen.txtNewDueTime.value = "00:00:00";
	else
	{
		if( sTime.length < 8 || 
		    isNaN(parseInt(sHours)) || isNaN(parseInt(sMins)) || isNaN(parseInt(sSecs)) )
			bSuccess = false;
		if( bSuccess )
		{
			if( (parseInt(sHours) < 0 || parseInt(sHours) > 23) ||
			    (parseInt(sMins) < 0 || parseInt(sMins) > 59) ||
			    (parseInt(sSecs) < 0 || parseInt(sSecs) > 59)      )
					bSuccess = false;
			if( bSuccess &&
				(sTime.charAt(2) != ":" || sTime.charAt(5) != ":") )
					bSuccess = false;
			
		}
	
		if(!bSuccess)
		{
			alert("The time you have entered is invalid");
			frmScreen.txtNewDueTime.focus();
		}
	}
	return(bSuccess);
}
function btnCancel.onclick()
{
	if(m_sMetaAction == "TM030")
		frmToTM030.submit();
	else frmToTM020.submit();
}
function btnSubmit.onclick()
{
	if(CheckTimeFields())
	{
		if(frmScreen.onsubmit())
		{
			if(m_sExistingOwner != frmScreen.cboNewOwner.value || m_sExistingDueDate != GetNewDueDateTime())
			{
				<% /* BMIDS725 If a unit has been selected, make sure that a user has also been selected and vice versa */ %>
				if ((frmScreen.cboNewUnit.value != "") && (frmScreen.cboNewOwner.value == ""))
				{
					alert("Please enter a new owner");
					return;
				}
				else if ((frmScreen.cboNewOwner.value != "") && (frmScreen.cboNewUnit.value == ""))
				{
					alert("Please enter a new unit");
					return;
				}
		
				if(scScreenFunctions.CompareDateStringToSystemDateTime(GetNewDueDateTime(),"<"))
				{
					alert("The new due date and time must not be in the past");
					return;
				}
				<%/* GD BM0340 - Remove working hours checks
				var  varNewDueDate  = frmScreen.txtNewDueDate.value
				var  varCurrentDueDate  = frmScreen.txtExistingDueDate.value
				var  varDistributionChannel =  scScreenFunctions.GetContextParameter(window,"idDistributionChannelID");  		
				
				if(IsChanged())
				{
					if(frmScreen.txtNewDueDate.value != "")
					{
						var DateXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
						var tagRequest = DateXML.CreateRequestTag(window, "CheckNonWorkingOccurence");
																						
						DateXML.CreateActiveTag("SYSTEMDATE");
						DateXML.CreateTag("DATE",varNewDueDate );
						DateXML.CreateTag("CHANNELID",varDistributionChannel );
						// 						DateXML.RunASP(document, "CheckNonWorkingOccurence.asp"); 
						// Added by automated update TW 09 Oct 2002 SYS5115
						switch (ScreenRules())
							{
							case 1: // Warning
							case 0: // OK
													DateXML.RunASP(document, "CheckNonWorkingOccurence.asp"); 
								break;
							default: // Error
								DateXML.SetErrorResponse();
							}

					}
					else
					{
						var DateXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
						var tagRequest = DateXML.CreateRequestTag(window, "CheckNonWorkingOccurence");
												
					
						DateXML.CreateActiveTag("SYSTEMDATE");
						DateXML.CreateTag("DATE",varCurrentDueDate );
						DateXML.CreateTag("CHANNELID",varDistributionChannel );
						// 						DateXML.RunASP(document, "CheckNonWorkingOccurence.asp"); 	
						// Added by automated update TW 09 Oct 2002 SYS5115
						switch (ScreenRules())
							{
							case 1: // Warning
							case 0: // OK
													DateXML.RunASP(document, "CheckNonWorkingOccurence.asp"); 	
								break;
							default: // Error
								DateXML.SetErrorResponse();
							}

					}
					
					if (DateXML.IsResponseOK()==true)
					{
						var varDateAccept = DateXML.GetTagText("NONWORKINGIND");
						if(varDateAccept == 1)
						{
							alert("This is not a valid working day"); 
							return;
						}
					}	
				}
				*/%>
				var strUserID;
				var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				var tagReq = XML.CreateRequestTag(window, "UpdateCaseTask");
				for(var iCount = 0; iCount < taskXML.ActiveTagList.length; iCount++)
				{
					taskXML.SelectTagListItem(iCount);
					XML.ActiveTag = tagReq;
					XML.CreateActiveTag("CASETASK");
					XML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
					XML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
					XML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
					XML.SetAttribute("ACTIVITYINSTANCE", taskXML.GetAttribute("ACTIVITYINSTANCE"));
					XML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
					XML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
					XML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
					XML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
					if(GetNewDueDateTime() != "")
						XML.SetAttribute("TASKDUEDATEANDTIME", GetNewDueDateTime());
						
					XML.SetAttribute("OWNINGUNITID", frmScreen.cboNewUnit.value);
					strUserID = frmScreen.txtUserId.value;
					
					if(strUserID.length > 0	)
						XML.SetAttribute("OWNINGUSERID", frmScreen.txtUserId.value);
						
					XML.SetAttribute("TASKSTATUSSETBYUSERID", scScreenFunctions.GetContextParameter(window, "idUserID", null));
					XML.SetAttribute("TASKSTATUSSETBYUNITID", scScreenFunctions.GetContextParameter(window, "idUnitID", null));
						
				}
				// 				XML.RunASP(document, "MsgTMBO.asp");
				// Added by automated update TW 09 Oct 2002 SYS5115
				switch (ScreenRules())
					{
					case 1: // Warning
					case 0: // OK
									XML.RunASP(document, "MsgTMBO.asp");
						break;
					default: // Error
						XML.SetErrorResponse();
					}

				if(XML.IsResponseOK())
				{
					if(m_sMetaAction == "TM030")
					{
						frmToTM030.submit();
					}	
					else frmToTM020.submit();
				
				}// IF XML.IsResponseOK(1)
			}//Compare Dates
		}//frmScreen.onsubmit
	}//IF CheckTimeFields
}//btnSubmit.onclick
-->
</script>
</body>
</html>


