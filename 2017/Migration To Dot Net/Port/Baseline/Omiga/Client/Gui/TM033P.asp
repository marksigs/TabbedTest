<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      TM033P.asp
Copyright:     Copyright © 2003 Marlborough Stirling

Description:   Edit Action Popup
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:
******** NOTE THAT TM033.asp WILL ALSO NEED THE SAME CHANGES APPLIED TO IT AS THIS SCREEN, AS IT'S ESSENTIALLY THE SAME*****
Prog	Date		Description
MDC		04/04/2003	Created 
GD		10/07/2003	Removed checks for non-working day.
HMA     05/09/2003  Allow editing of New Unit  
HMA     30/11/2004  BMIDS725  Do not allow new unit to be selected without new user and vice versa.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
PSC     2209/2005	MAR32		Add processing for Sales Agents
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Edit Action  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</HEAD>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>

<div id="divOwnerDetails" style="TOP: 4px; LEFT: 10px; HEIGHT: 170px; WIDTH: 370px; POSITION: ABSOLUTE" class="msgGroup">
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

<div id="divDueDate" style="TOP: 184px; LEFT: 10px; HEIGHT: 95px; WIDTH: 370px; POSITION: ABSOLUTE" class="msgGroup">
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
	
	<span style="TOP:100px; LEFT:4px; POSITION:ABSOLUTE">
		<input id="btnSubmit" value="OK" type="button" style="WIDTH:60px" class="msgButton">
	</span>
	<span style="TOP:100px; LEFT:69px; POSITION:ABSOLUTE">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH:60px" class="msgButton">
	</span>

</div>

</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/TM033Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
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

var m_sRequestAttributes;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;

<% /* PSC 22/09/2005 MAR32 - Start */ %>
var m_nUserRole = "-1";
var m_nMinSALevel = 0;
var m_bIsSARole = false;
<% /* PSC 22/09/2005 MAR32 - End */ %>


function window.onload()
{
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	RetrieveData();
	
	SetMasks();
	Validation_Init();
//	SetScreenOnReadOnly();
	PopulateScreen();
	window.returnValue = null;
	
	ClientPopulateScreen();
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
	m_sReadOnly				= sParameters[0];
	m_sTaskXML				= sParameters[1];
	m_sRequestAttributes	= sParameters[2];
	m_sDistributionChannel	= m_sRequestAttributes[3];
	m_sUnitId				= m_sRequestAttributes[1];
	<% /* PSC 22/09/2005 MAR32 */ %>
	m_nUserRole             = sParameters[3];
}

function PopulateScreen()
{
	scScreenFunctions.SetFieldState(frmScreen, "txtExistingOwner", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtUserId", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtUnitCode", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtExistingDueDate", "R");
	scScreenFunctions.SetFieldState(frmScreen, "txtExistingDueTime", "R");
	<% /* Allow editing of New Unit combo */ %>
	<% /* scScreenFunctions.SetFieldState(frmScreen, "cboNewUnit", "R");*/ %>
	
	<% /* PSC 22/09/2005 MAR32 - Start */ %>
	var GlobalXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_nMinSALevel = parseInt(GlobalXML.GetGlobalParameterAmount(document,"TMMinSAAuthLevel"));

	var userRoleXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bIsSARole = userRoleXML.IsInComboValidationList(document,"UserRole", m_nUserRole, ["SA"]);
	<% /* PSC 22/09/2005 MAR32 - End */ %>
	
	taskXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
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
	XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_sRequestAttributes, "");
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
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bDisableOwnerCombo;
	bDisableOwnerCombo = true;
	
	while(frmScreen.cboNewOwner.options.length > 0) frmScreen.cboNewOwner.remove(0);
	
	XML.CreateRequestTagFromArray(m_sRequestAttributes, "");
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
			XML.SelectNodes("OMIGAUSERLIST/OMIGAUSER[WORKGROUPUSER='1']");
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

function frmScreen.btnCancel.onclick()
{
	window.close();
}

function frmScreen.btnSubmit.onclick()
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
				var  varDistributionChannel = m_sDistributionChannel // scScreenFunctions.GetContextParameter(window,"idDistributionChannelID");  		
				
				if(IsChanged())
				{
					if(frmScreen.txtNewDueDate.value != "")
					{
						var DateXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
						var tagRequest = DateXML.CreateRequestTagFromArray(m_sRequestAttributes, "CheckNonWorkingOccurence");
																						
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
						var DateXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
						var tagRequest = DateXML.CreateRequestTagFromArray(m_sRequestAttributes, "CheckNonWorkingOccurence");
					
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
				} */%>
				var strUserID;
				var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
				var tagReq = XML.CreateRequestTagFromArray(m_sRequestAttributes, "UpdateCaseTask");
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
						
					XML.SetAttribute("TASKSTATUSSETBYUSERID", m_sRequestAttributes[0], null);
					XML.SetAttribute("TASKSTATUSSETBYUNITID", m_sUnitId, null);

				}

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
					var sReturn = new Array();
					sReturn[0] = IsChanged();
					window.returnValue	= sReturn;
					window.close();
				}// IF XML.IsResponseOK(1)
			}//Compare Dates
		}//frmScreen.onsubmit
	}//IF CheckTimeFields
}//btnSubmit.onclick

-->
</script>
</body>
</html>





