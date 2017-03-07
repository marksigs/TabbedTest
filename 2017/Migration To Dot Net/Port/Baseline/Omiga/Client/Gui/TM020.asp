<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      TM020.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Task Search
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		24/10/00	Created (screen paint)
JLD		04/01/01	SYS1767 lock application on download.
JLD		05/01/01	SYS1769 make list box multi-select for editing tasks.
JLD		08/01/01	SYS1767 move LockApp.asp file to the includes directory
JLD		08/01/01	SYS1800 deal with scrolling table.
JLD		15/01/01	SYS1808 set up the stage context params.
APS		03/03/01	SYS1993	LockApp processing changes
ADP		12/03/01	SYS1985 Change routing on select from MN060 to TM030
ADP		29/03/01	SYS2191	Task context changed from using idXML to idTaskXML
BG		10/04/01	SYS2096 Added functionality to print Button.
JLD		08/10/01	SYS2736 Don't download omPC again
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		29/05/02	SYS4767 MSMS to Core integration
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS History:

Prog	Date		Description
PSC  	29/06/2005	MAR5 Add Delivery type for printing templates.
GHun	19/07/2005	MAR7 Integerate local printing
PSC		22/08/2005	MAR32 WP08 changes
MahaT	12/10/2005	MAR170 Change label "Task Date Due" to "Task Due Date"
PSC		25/10/2005	MAR300 Amend call to LockApplication to synchronise customer details
PSC		28/11/2005	MAR307 Amend Populate Task Description to handle RECORDNOTFOUND
GHun	05/12/2005	MAR796 Turn off document compression
PJO     13/12/2005  MAR676 Disable the print button
Shiva.S 30/12/2005  MAR950 Label Correction:Label Task due date should read Tasks due date.
PJO     08/03/2006  MAR1378 Don't use hardcoded values for cancelled / declined stage numbers
PJO     09/03/2006  MAR1359 Add Progress Message during search
PJO     09/03/2006  MAR1359 Parameterise Progress Message width
GHun	20/03/2006	MAR1453	Hide progress window before popping up an error
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM History:

Prog	Date		Description
AW		24/05/2006	EP602	Hide Print button 
HMA     08/08/2006  EP602   Enable Print List button when appropriate.
PB		13/09/2006	EP602	CC47 - New Task List print with no document archiving
PB		14/09/2006	EP602	Allow printing if records exceed 500 but show a warning
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ 
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex="-1" type="text/x-scriptlet" width="1" viewastext></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex="-1" type="text/x-scriptlet" viewastext></object>
<span style="LEFT: 310px; POSITION: absolute; TOP: 470px">
	<object data="scTableListScroll.asp" id="scScrollTable" style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex="-1" type="text/x-scriptlet" viewastext></object>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToTM033" method="post" action="TM033.asp" style="DISPLAY: none"></form>
<!-- <form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form> -->
<form id="frmToMN015" method="post" action="MN015.asp" style="DISPLAY: none"></form>

<% /* SYS1985 */ %>
<form id="frmToTM030" method="post" action="TM030.asp" style="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange" year4>
<div id="divSearch" style="HEIGHT: 210px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabelHead">
	Search Criteria
</span>
<span style="LEFT: 15px; POSITION: absolute; TOP: 30px" class="msgLabel">
	Unit Name
	<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
		<select id="cboUnitName" style="WIDTH: 190px" class="msgCombo" menusafe="true"></select>
	</span>
</span>
<span style="LEFT: 306px; POSITION: absolute; TOP: 30px" class="msgLabel">
	User Name
	<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
		<select id="cboUserName" style="WIDTH: 190px" class="msgCombo"></select>
	</span>
</span>
<span style="LEFT: 15px; POSITION: absolute; TOP: 60px" class="msgLabel">
	<% /* MahaT  MAR170: Change "Task Date Due" to "Task Due Date" */ %>
	<% /* Shiva S:30/12/2005:Defect:MAR950:Label Task due date should read Tasks due date%*/%>
	Tasks Due Date:
	<span style="LEFT: 90px; POSITION: absolute; TOP: 0px" class="msgLabel">
		Start
		<span style="LEFT: 40px; POSITION: absolute; TOP: -3px">
			<input id="txtDueDateStart" maxlength="10" style="WIDTH: 100px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 90px; POSITION: absolute; TOP: 30px" class="msgLabel">
		End
		<span style="LEFT: 40px; POSITION: absolute; TOP: -3px">
			<input id="txtDueDateEnd" maxlength="10" style="WIDTH: 100px" class="msgTxt">
		</span>
	</span>
</span>	
<span style="LEFT: 306px; POSITION: absolute; TOP: 60px" class="msgLabel">
	Application Number
	<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
		<input id="txtApplicationNum" maxlength="12" style="WIDTH: 150px" class="msgTxt">
	</span>
</span>
<span style="LEFT: 306px; POSITION: absolute; TOP: 90px" class="msgLabel">
	Sales Lead Status
	<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
		<select id="cboSalesLeadStatus" style="WIDTH: 190px" class="msgCombo" menusafe="true"></select>
	</span>
</span>

<% //SG 29/05/02 SYS4767 START %>
<!-- SG 26/02/02 SYS4183 Added combo-->
<span style="LEFT: 15px; POSITION: absolute; TOP: 120px" class="msgLabel">
	Task Type
	<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
		<select id="cboTaskType" style="WIDTH: 190px" class="msgCombo" menusafe="true"></select>
	</span>
</span>
<span style="LEFT: 306px; POSITION: absolute; TOP: 120px" class="msgLabel">
	SLA Expiry Within
	<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
		<input id="txtSLAExpiryWithin" maxlength="4" style="WIDTH: 35px" class="msgTxt">
	</span>
	<span style="LEFT: 140px; POSITION: absolute; TOP: 0px" class="msgLabel">
		Days
	</span>
</span>

<% //SG 29/05/02 SYS4767 END %>
<% /* BMIDS702 START */ %>
<span style="LEFT: 15px; POSITION: absolute; TOP: 150px" class="msgLabel">
	Task Status
	<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
		<select id="cboTaskStatus" style="WIDTH: 190px" class="msgCombo" menusafe="true"></select>
	</span>
</span>
<span style="LEFT: 15px; POSITION: absolute; TOP: 180px" class="msgLabel">
	Task Description
	<span style="LEFT: 90px; POSITION: absolute; TOP: -3px">
		<select id="cboTaskDescription" style="WIDTH: 190px" class="msgCombo" menusafe="true"></select>
	</span>
</span>
<% /* BMIDS702 END */ %>
<span style="LEFT: 534px; POSITION: absolute; TOP: 180px">
	<input id="btnSearch" value="Search" type="button" style="WIDTH: 60px" class="msgButton">
</span>
</div>

<div id="divTaskList" style="HEIGHT: 220px; LEFT: 10px; POSITION: absolute; TOP: 280px; WIDTH: 604px" class="msgGroup">
<div id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 10px">
	<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
	<tr id="rowTitles">
		<td width="25%" class="TableHead">Application/<br>Account No</td>	
		<td width="35%" class="TableHead">Task</td>
		<td width="15%" class="TableHead">Status</td>		
		<td width="20%" class="TableHead">Due Date/Time</td>
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
</div>
<span style="LEFT: 4px; POSITION: absolute; TOP: 190px">
		<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
</span>
<span id="spnPrintButton" style="LEFT: 90px; POSITION: absolute; TOP: 190px">
	<input id="btnPrint" value="Print List" type="button" style="WIDTH: 60px" class="msgButton">
</span>
<span style="LEFT: 176px; POSITION: absolute; TOP: 190px">
	<input id="btnSelect" value="Select" type="button" style="WIDTH: 60px" class="msgButton">
</span>

</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 510px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /*Include the file to lock the application */ %>
<!-- #include FILE="Includes/LockApp.asp" --> 

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/TM020attribs.asp" -->

<% /* Specify Code Here */ %>
<script src="includes/Documents.js" language="jscript" type="text/javascript"></script>	<% /* MAR7 GHun */ %>
<script language="JScript" type="text/javascript">
<!--
<% /* PB 13/09/2006 EP602 */ %>
var HTMLToPrint = '';
var m_sMetaAction = null;
var m_sReadOnly = "";
var scScreenFunctions;
var m_sDistributionChannel;
var m_iTableLength = 10;
var searchXML = null;
var m_sApplicationNumber;
var m_sApplicationFactFindNumber;
var m_TM020XML;
<% /* PSC 22/08/2005 MAR32 - Start */ %>
var m_nUserRole = -1;
var m_nMinSALevel = 0;
var m_nMinNSALevel = 0;
var m_bIsNSARole = false;
var m_bIsSARole = false;
<% /* PSC 22/08/2005 MAR32 - End */ %>
var m_iTotalRecords; <% /* EP602 */ %>


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit");
	var bStateReceived;
	bStateReceived = false;
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Task Search","TM020",scScreenFunctions);

<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);

	frmScreen.cboUserName.disabled = true; 
	frmScreen.btnSelect.disabled =true;
	frmScreen.btnEdit.disabled =true;
	frmScreen.btnPrint.disabled = true;
	<% /* EP602 Do not hide print button */ %>
	
	RetrieveContextData();
	<% /* PSC 22/08/2005 MAR32 */ %>
	GetUserRoleDetails();
	PopulateUnit(0);
	bStateReceived = RetrieveStateContext();
		
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();

	if( !bStateReceived )
	{
		PopulateScreen();
	}
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	ClientPopulateScreen();
}

function spnTable.ondblclick()
{
	<% /* BMIDS541 double clicking should only work if btnSelect is enabled */ %>
	if (!frmScreen.btnSelect.disabled)
		frmScreen.btnSelect.onclick();
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
	m_sDistributionChannel = scScreenFunctions.GetContextParameter(window,"idDistributionChannelId","");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","");

	<% /* PSC 22/08/2005 MAR32 */ %>
	m_nUserRole = parseInt(scScreenFunctions.GetContextParameter(window,"idRole","-1"));
}

function SetUserCombo( strUserID )
{
	if(strUserID.length > 0 )
	{
		frmScreen.cboUserName.value = strUserID;
	}

	<% /* PSC 22/08/2005 MAR32 - Start */ %>
	if (frmScreen.cboUserName.selectedIndex < 0)
		frmScreen.cboUserName.selectedIndex = 0;
	<% /* PSC 22/08/2005 MAR32 - End */ %>
		
}
function SetUnitCombo( strUnitID )
{
	frmScreen.cboUnitName.value = strUnitID;
}

function RetrieveStateContext()
{
	<% /* Incase we've come back here after an Edit */ %>
	var strTM020XML = scScreenFunctions.GetContextParameter(window,"idXML2","");	
	var bStateReceived;
	bStateReceived = false;

	if( strTM020XML.length > 0 )
	{
		m_TM020XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
		<% /* PSC 22/08/2005 MAR32 - Start */ %>
		var strUnitID,
			strUserID,
			strDueDateStart,
			strDueDateEnd,
			strCaseNo,
			strCaseTask,
			strSalesLeadStatus,
			strSLAExpiryWithin;
		<% /* PSC 22/08/2005 MAR32 - End */ %>
					
		m_TM020XML.LoadXML(strTM020XML);
		m_TM020XML.CreateTagList("TM020CONTEXT");
		var iCount = m_TM020XML.ActiveTagList.length;

		if(iCount > 0)
		{
			bStateReceived = true;
			m_TM020XML.SelectTagListItem(0)

			strUnitID = m_TM020XML.GetAttribute("TM020UNITID");
			strUserID = m_TM020XML.GetAttribute("TM020USERID");
			
			<% /* PSC 22/08/2005 MAR32 - Start */ %>
			strDueDateStart = m_TM020XML.GetAttribute("TM020DUEDATESTART");
			strDueDateEnd = m_TM020XML.GetAttribute("TM020DUEDATEEND");
			strCaseNo = m_TM020XML.GetAttribute("TM020CASENUMBER");
			strTaskType = m_TM020XML.GetAttribute("TM020TASKTYPE"); 
			strTaskStatus = m_TM020XML.GetAttribute("TM020TASKSTATUS"); 
			strCaseTaskDesc = m_TM020XML.GetAttribute("TM020CASETASKDESC"); 
			strSalesLeadStatus = m_TM020XML.GetAttribute("TM020SALESLEADSTATUS");
			strSLAExpiryWithin = m_TM020XML.GetAttribute("TM020SLAEXPIRYWITHIN");
			<% /* PSC 22/08/2005 MAR32 - End */ %>

			SetUnitCombo(strUnitID);
			PopulateUser(0);
			
			<% /* PSC 22/08/2005 MAR32 */ %>
			PopulateCombos();			
			
			<% /* SG 29/05/02 SYS4767 START
			DRC 30/04/02 MSMS0051	*/%>
			if(strTaskType.length > 0)
				frmScreen.cboTaskType.value = strTaskType; 
			
			<% /*SG 29/05/02 SYS4767 END */ %>
			
			SetUserCombo( strUserID );
			if(strTaskStatus > 0)
				frmScreen.cboTaskStatus.value= strTaskStatus; 
			
			if(strTaskType.length > 0)
				PopulateTaskDescription(strTaskType);
								
			if(strCaseTaskDesc.length > 0)
			   frmScreen.cboTaskDescription.value= strCaseTaskDesc	;
			   
			<% /* PSC 22/08/2005 MAR32 - Start */ %>
			if (strSalesLeadStatus.length > 0)
			   frmScreen.cboSalesLeadStatus.value= strCaseTaskDesc;


			frmScreen.txtDueDateStart.value  = strDueDateStart;
			frmScreen.txtDueDateEnd.value  = strDueDateEnd;	
			frmScreen.txtSLAExpiryWithin.value  = strSLAExpiryWithin;			
			frmScreen.txtApplicationNum.value  = strCaseNo;
			<% /* PSC 22/08/2005 MAR32 - End */ %>
				
			m_TM020XML.RemoveActiveTag();

			var strActiveTag;
			
			if( m_TM020XML.ActiveTag != null )
			{
				strActiveTag = m_TM020XML.ActiveTag.xml
			}
			
			scScreenFunctions.SetContextParameter(window,"idXML2",strActiveTag );
			DoSearch();
		}
	}
	return(bStateReceived);
}

function PopulateScreen()
{
<%	//set the due date to today
%>	
	var strUnitID;
	var strUserID;
	
	<% /* MO - BMIDS00807 */ %>
	<% /* frmScreen.txtDueDate.value = scScreenFunctions.DateToString(scScreenFunctions.GetSystemDate()); */ %>
	<% /* PSC 22/08/2005 MAR32 - Start */ %>
	frmScreen.txtDueDateStart.value;
	frmScreen.txtDueDateEnd.value = scScreenFunctions.DateToString(scScreenFunctions.GetAppServerDate());
	<% /* PSC 22/08/2005 MAR32 - End */ %>

	strUnitID = scScreenFunctions.GetContextParameter(window,"idUnitId","");
	strUserID = scScreenFunctions.GetContextParameter(window,"idUserId","");

	if( strUnitID.length > 0 ) 
	{
		SetUnitCombo( strUnitID );
		
		PopulateUser(0);
		
		if (frmScreen.cboUserName.length > 0 )
		{
			SetUserCombo( strUserID );
		}
	}
	<% /* PSC 22/08/2005 MAR32 */ %>
	PopulateCombos();

	<% /* SG 29/05/02 SYS4767 END	*/ %>
}

function PopulateUser(nDEBUG_USE)
{
	var userXML = null;
	var sUnitID;
	var sFile;
	var bDisableUserCombo;
	
	bDisableUserCombo = true;

	userXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	while(frmScreen.cboUserName.options.length > 0) frmScreen.cboUserName.remove(0);

	<% /* Get the UnitID from the unit combo */ %>
	sUnitID = frmScreen.cboUnitName.value 

	if( sUnitID.length > 0 )
	{

		if(nDEBUG_USE == 0) 
		{
			userXML.CreateRequestTag(window , null);
			userXML.CreateActiveTag("USERLIST");
			userXML.CreateTag("UNITID",sUnitID);
			userXML.RunASP(document, "FindUserList.asp")
		}
		else
		{
			sFile = "C:\\projects\\dev\\XML\\FindUserList.xml";

			var sXMLString = userXML.ReadXMLFromFile(sFile);
			userXML.LoadXML(sXMLString);
		}	

		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = userXML.CheckResponse(ErrorTypes);
			
		if ( ErrorReturn[0] == true )
		{
			var strUserName;
			var intCounter;
			var intCount;
			var xmlActiveTag = userXML.ActiveTag;
			var xmlTagList = null;
			
			<% /* PSC 22/08/2005 MAR32 - Start */ %>
			if (m_bIsNSARole && m_nUserRole < m_nMinNSALevel)
				xmlTagList = xmlActiveTag.selectNodes("OMIGAUSERLIST/OMIGAUSER[WORKGROUPUSER='1']/USERID");
			else
				xmlTagList = xmlActiveTag.selectNodes("OMIGAUSERLIST/OMIGAUSER/USERID");
			<% /* PSC 22/08/2005 MAR32 - End */ %>


			intCount = xmlTagList.length;
			
			<% /* PSC 22/08/2005 MAR32 - Start */ %>
			if( intCount > 1 )
				AddComboOption( frmScreen.cboUserName, "<ALL>","" );
			<% /* PSC 22/08/2005 MAR32 - End */ %>
				
			if( intCount > 0 )
			{
				bDisableUserCombo = false;
				
				for( intCounter = 0; intCounter < intCount; intCounter++ )
				{
					strUserName = xmlTagList.item(intCounter).text
					AddComboOption( frmScreen.cboUserName , strUserName, strUserName );
				}
			}
		}
	}
	
	frmScreen.cboUserName.disabled = bDisableUserCombo;		
}

function PopulateUnit(nDEBUG_USE)
{
	
	var unitXML = null;
	unitXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	unitXML.CreateRequestTag(window , null);
	unitXML.CreateActiveTag("UNIT");
	unitXML.CreateTag("CHANNELID",m_sDistributionChannel);

	var sFile;

	if(nDEBUG_USE == 0) 
	{
		unitXML.RunASP(document, "FindUnitList.asp");
	}
	else
	{
		sFile = "C:\\projects\\dev\\XML\\FindUnitList.xml"
		var sXMLString = unitXML.ReadXMLFromFile(sFile);
		unitXML.LoadXML(sXMLString);
	}
	if(unitXML.IsResponseOK())
	{
		var strUnitName;
		var strUnitID;
		var intCounter;
		var intCount;
		var xmlActiveTag = unitXML.ActiveTag;
		var xmlUnitIDList;
		var xmlUnitNameList;
		
		<% /* PSC 22/08/2005 MAR32 - Start */ %>
		if (m_bIsSARole && m_nUserRole < m_nMinSALevel)
		{
			var strCurrentUnit = scScreenFunctions.GetContextParameter(window,"idUnitId","");
			xmlUnitIDList = unitXML.SelectNodes("UNITLIST/UNIT[UNITID='" + strCurrentUnit + "']/UNITID");
			xmlUnitNameList = unitXML.SelectNodes("UNITLIST/UNIT[UNITID='" + strCurrentUnit + "']/UNITNAME");
		}
		else
		{
			xmlUnitIDList = xmlActiveTag.getElementsByTagName("UNITID");
			xmlUnitNameList = xmlActiveTag.getElementsByTagName("UNITNAME");
		}
		<% /* PSC 22/08/2005 MAR32 - End */ %>
 
		intCount = xmlUnitIDList.length;
		
		<% /* PSC 22/08/2005 MAR32 */ %>
		if (intCount > 1)
			AddComboOption( frmScreen.cboUnitName, "<ALL>","" );

		if( intCount > 0 )
		{
			for( intCounter = 0; intCounter < intCount; intCounter++ )
			{
				strUnitName = xmlUnitNameList.item(intCounter).text;
				strUnitID = xmlUnitIDList.item(intCounter).text;

				AddComboOption( frmScreen.cboUnitName, strUnitName, strUnitID );
			}

			<% /*Set the selection to the first one */ %>
			frmScreen.cboUnitName.selectedIndex = 0;
		}
		
		<% /* PSC 22/08/2005 MAR32 - Start */ %>
		if (m_bIsSARole && m_nUserRole < m_nMinSALevel)
			frmScreen.cboUnitName.disabled = true;
		else
			frmScreen.cboUnitName.disabled = false;	
		<% /* PSC 22/08/2005 MAR32 - End */ %>		
	}
}

function spnTable.onclick()
{
	SetButtonStates();
}

function SetButtonStates()
{
	var saRowArray = null;
	var saRowArray = scScrollTable.getArrayofRowsSelected();
	
	if( saRowArray.length > 1 )
	{
		<% /* PSC 22/08/2005 MAR32 - Start */ %>
		if (frmScreen.cboTaskType.value.length != 0 || frmScreen.txtSLAExpiryWithin.value == 0) 
			frmScreen.btnEdit.disabled = false;
		else
			frmScreen.btnEdit.disabled = true;
		<% /* PSC 22/08/2005 MAR32 - End */ %>
			
		
		frmScreen.btnSelect.disabled = true;
	}
	else if(saRowArray.length > 0)
	{
		<% /* PSC 22/08/2005 MAR32 - Start */ %>
		if (frmScreen.cboTaskType.value.length != 0 || frmScreen.txtSLAExpiryWithin.value == 0) 
			frmScreen.btnEdit.disabled = false;
		else
			frmScreen.btnEdit.disabled = true;
		<% /* PSC 22/08/2005 MAR32 - End */ %>
			
		
		frmScreen.btnSelect.disabled = false;
	}
	else
	{
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnSelect.disabled = true;
	}
}

function GetUserID()
{
	return(frmScreen.cboUserName.value);
}

function GetUnitID()
{
	return(frmScreen.cboUnitName.value);
}

<% /* PSC 22/08/2005 MAR32 - Start */ %>
function GetDueDateStart()
{
	return(frmScreen.txtDueDateStart.value);
}

function GetDueDateEnd()
{
	return (frmScreen.txtDueDateEnd.value);
}

function GetSalesLeadStatus()
{
	return (frmScreen.cboSalesLeadStatus.value);
}

function GetSLAExpiryWithin()
{
	return (frmScreen.txtSLAExpiryWithin.value);
}
<% /* PSC 22/08/2005 MAR32 - End */ %>

function GetCaseNumber()
{
	return(frmScreen.txtApplicationNum.value);
}

<% /* SG 29/05/02 SYS4767 Function added
      SG 27/02/02 SYS4183  */ %>
function GetTaskType()	 
{
	return(frmScreen.cboTaskType.value);
}

<% /* BMIDS702 Start */ %>
function GetTaskID()	
{
	return(frmScreen.cboTaskDescription.value);
}

function GetTaskStatus()	
{
	return(frmScreen.cboTaskStatus.value);
}
<% /* BMIDS702 End */ %>

function DoSearch()
{
	<% /* BMIDS702  Add Task ID and Status to search */ %>

	var strAppNo;
	<% /* PSC 22/08/2005 MAR32 - Start */ %>
	var strDueDateStart;
	var strDueDateEnd;
	<% /* PSC 22/08/2005 MAR32 - End */ %>
	
	var strUserID;
	var strUnitID;
	var strTaskType;
	var strTaskStatus;
	var strTaskID;
	
	<% /* PSC 22/08/2005 MAR32 - Start */ %>
	var strSalesLeadStatus;
	var strSLAExpiryWithin;
	var intSLAExpiryWithin = 0 
		
	<% /* Validate SLA Expiry */ %>
	strSLAExpiryWithin = GetSLAExpiryWithin();
	
	if (strSLAExpiryWithin.length > 0)
	{
		var intFirstMinusPos = strSLAExpiryWithin.indexOf("-");
		var intSecondMinusPos = strSLAExpiryWithin.indexOf("-", intFirstMinusPos + 1);
		
		if ((intFirstMinusPos != -1 && intFirstMinusPos != 0) || 
			 intSecondMinusPos != -1 || parseInt(strSLAExpiryWithin) > 999)		
		{
			alert("SLA Expiry must be between -999 and 999 days");
			frmScreen.txtSLAExpiryWithin.focus();
			return;
		}
	}
	
	<% /* Check start date is on or before end dete */ %>
	strDueDateStart = GetDueDateStart();
	strDueDateEnd = GetDueDateEnd();
	
	if  (strDueDateStart.length > 0 && strDueDateEnd.length > 0)
	{
		var dteStartDate = scScreenFunctions.GetDateObject(frmScreen.txtDueDateStart);
		var dteEndDate = scScreenFunctions.GetDateObject(frmScreen.txtDueDateEnd);
		
		if (dteEndDate < dteStartDate)
		{
			alert ("End date cannot be before start date")
			frmScreen.txtDueDateEnd.focus();
			return;
		} 
	}
	<% /* PSC 22/08/2005 MAR32 - End */ %>

	<% /* PJO 09/03/2006 MAR1359 Add progress message */ %>
	scScreenFunctions.progressOn("Please wait ... Searching for Tasks", 400);
	
	scScrollTable.clear();
	searchXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	searchXML.CreateRequestTag(window , "FindCaseTaskListLite");
	searchXML.CreateActiveTag("CASETASK");

	strAppNo = frmScreen.txtApplicationNum.value 
	if( strAppNo.length > 0)
	{
		searchXML.SetAttribute("CASEID",strAppNo);
		<% /* BMIDS702 
		   Ensure that only Incomplete or Pending tasks are returned for this application */ %>
		searchXML.SetAttribute("TASKSTATUS","I");
	}
	else	
	{
		searchXML.SetAttribute("SOURCEAPPLICATION","Omiga");
		
		<% /* PSC 22/08/2005 MAR32 - Start */ %>
		if( strDueDateStart.length > 0)
			searchXML.SetAttribute("TASKDUESTARTDATEANDTIME",strDueDateStart);
			
		if( strDueDateEnd.length > 0)
			searchXML.SetAttribute("TASKDUEENDDATEANDTIME",strDueDateEnd);
		<% /* PSC 22/08/2005 MAR32 - End */ %>
			
		strUnitID = GetUnitID();
		if( strUnitID.length > 0)
			searchXML.SetAttribute("OWNINGUNITID",strUnitID);

		strUserID = GetUserID();
		if( strUserID.length > 0)
			searchXML.SetAttribute("OWNINGUSERID",strUserID);

		strTaskType = GetTaskType();
		if( strTaskType.length > 0)
			searchXML.SetAttribute("TASKTYPE",strTaskType);
		
		strTaskID = GetTaskID();
		if( strTaskID.length > 0)
			searchXML.SetAttribute("TASKID",strTaskID);	
			
		strTaskStatus = GetTaskStatus();
		if( strTaskStatus.length > 0)
			searchXML.SetAttribute("TASKSTATUS",strTaskStatus);	
		else  <% /* Default to <ALL> */ %>
			searchXML.SetAttribute("TASKSTATUS","I");
			
		<% /* PSC 22/08/2005 MAR32 - Start */ %>
		strSalesLeadStatus = GetSalesLeadStatus();
		if (strSalesLeadStatus.length > 0)
			searchXML.SetAttribute("SALESLEADSTATUS",strSalesLeadStatus);
			
		if (strSLAExpiryWithin.length > 0)
			searchXML.SetAttribute("SLAEXPIRYWITHIN",strSLAExpiryWithin);
		<% /* PSC 22/08/2005 MAR32 - End */ %>		
	}

	<% /* Ensure Task Status is displayed */ %>
	searchXML.SetAttribute("_COMBOLOOKUP_","y");
				
	switch (ScreenRules())
	{
		case 1: <% /* Warning */ %>
		case 0: <% /* OK  */ %>
			searchXML.RunASP(document, "OmigaTMBO.asp");
			break;
		default: <% /* Error */ %>
			searchXML.SetErrorResponse();
	}

	<% /* PJO 09/03/2006 MAR1359 Add progress message */ %>
	scScreenFunctions.progressOff();

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = searchXML.CheckResponse(ErrorTypes);
	
	if ( ErrorReturn[0] == true )
	{
		PopulateListBox();
		<% // PJO 13/12/2005 MAR676 - Leave Print List disabled %>
		<% /* EP602 Enable print button if less than 500 records have been returned */ %>
		<% /* EP602 Now always enable print button - but warn prior to printing when records exceed 500 */ %>
		//if (m_iTotalRecords <= 500)
			frmScreen.btnPrint.disabled = false;
		//else
		//	frmScreen.btnPrint.disabled = true;
	
	}
	else
	{	
		frmScreen.btnPrint.disabled = true;
		alert("Unable to find any outstanding tasks for your search criteria");
	}

	<% /* PSC 22/08/2005 MAR32 */ %>
	SetButtonStates();
}

function frmScreen.btnSearch.onclick()
{
	DoSearch();	
}

function PopulateListBox()
{
	var iCount;
	searchXML.CreateTagList("CASETASK");
	var iCount = searchXML.ActiveTagList.length;
	
	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iCount);
	scScrollTable.EnableMultiSelectTable();
	ShowList(0);	
}

function ShowList(nStart)
{
	scScrollTable.clear();
	for (var iCount = 0; iCount < searchXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		searchXML.SelectTagListItem(iCount + nStart);

		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), searchXML.GetAttribute("CASEID"));
		<% /* BMIDS702 Display Task name instead of ID */ %>
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), searchXML.GetAttribute("CASETASKNAME"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), searchXML.GetAttribute("TASKSTATUS_TEXT"));
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), searchXML.GetAttribute("TASKDUEDATEANDTIME"));
		tblTable.rows(iCount+1).setAttribute("TagListItem", (iCount + nStart));
	}
	
	<% /* EP602 */ %>
	m_iTotalRecords = searchXML.ActiveTagList.length;
}

function AddComboOption(cboToAdd, strName, strValue, intPosition)
{
	var TagOPTION = document.createElement("OPTION");
	TagOPTION = document.createElement("OPTION");
	TagOPTION.value	= strValue;
	TagOPTION.text = strName;
	
	<% /* PSC 22/08/2005 MAR32 - Start */ %>
	if (intPosition == undefined || intPosition == null)
		cboToAdd.add(TagOPTION);
	else
		cboToAdd.add(TagOPTION, intPosition);
	<% /* PSC 22/08/2005 MAR32 - End */ %>

}

function GetTasksXML()
{
<%	//For each row of the table selected get the associated TASK xml and return them all
%>	var saRowArray = scScrollTable.getArrayofRowsSelected();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var tagRequest = XML.CreateRequestTag(window , "");
	for(var iCount = 0; iCount < saRowArray.length; iCount++)
	{
		searchXML.SelectTagListItem(saRowArray[iCount] -1);
		var newNode = searchXML.GetTagListItem(saRowArray[iCount] -1);
		var node = newNode.cloneNode(true);
		XML.ActiveTag.appendChild(node);
	}
	return(XML.XMLDocument.xml);
}
function frmScreen.btnEdit.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idTaskXML",GetTasksXML() );
	scScreenFunctions.SetContextParameter(window,"idMetaAction","TM020");
	SaveContextXML();
	frmToTM033.submit();
}

function SaveContextXML()
{
	var strXML;
	strXML = scScreenFunctions.GetContextParameter(window,"idXML2","");	

	if( strXML.length == 0 )
		m_TM020XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();	
	else
		m_TM020XML.LoadXML(strXML);
	
	m_TM020XML.CreateActiveTag("TM020CONTEXT");
	m_TM020XML.SetAttribute("TM020UNITID",GetUnitID());
	m_TM020XML.SetAttribute("TM020USERID",GetUserID());
	
	<% /* PSC 22/08/2005 MAR32 - Start */ %>
	m_TM020XML.SetAttribute("TM020DUEDATESTART",GetDueDateStart());
	m_TM020XML.SetAttribute("TM020DUEDATEEND",GetDueDateEnd());
	m_TM020XML.SetAttribute("TM020CASENUMBER",GetCaseNumber());
	m_TM020XML.SetAttribute("TM020TASKTYPE",GetTaskType());
	m_TM020XML.SetAttribute("TM020TASKSTATUS",GetTaskStatus());
	m_TM020XML.SetAttribute("TM020CASETASKDESC",GetTaskID());
	m_TM020XML.SetAttribute("TM020SALESLEADSTATUS",GetSalesLeadStatus());
	m_TM020XML.SetAttribute("TM020SLAEXPIRYWITHIN",GetSLAExpiryWithin());
	<% /* PSC 22/08/2005 MAR32 - End */ %>
		 
	scScreenFunctions.SetContextParameter(window,"idXML2",m_TM020XML.ActiveTag.xml );
}

function btnSubmit.onclick()
{
	frmToMN015.submit();
}

function frmScreen.cboUnitName.onchange()
{
	PopulateUser(0);
}

function frmScreen.btnSelect.onclick()
{
	var sApplicationStage;

<%	//find the task in the XML list for the one selected.
	//Select button is only enabled when there is only one row selected
%>	var saRowArray = scScrollTable.getArrayofRowsSelected();
	searchXML.SelectTagListItem(saRowArray[0] -1);
	
	m_sApplicationNumber = searchXML.GetAttribute("CASEID");	
	m_sApplicationFactFindNumber = 1;

	<% /* Setup the Stage info */ %>
	sApplicationStage = searchXML.GetAttribute("STAGEID");
	
	<% /* PJO 08/03/2006 MAR1378 - Don't use hardcoded values
	if((sApplicationStage == 910)||(sApplicationStage == 920)) */ %>
	if((sApplicationStage == scScreenFunctions.getCancelledStageValue(window)) ||
	   (sApplicationStage == scScreenFunctions.getDeclinedStageValue(window)))
	{
		scScreenFunctions.SetContextParameter(window,"idCancelDeclineFreezeDataIndicator","1")
	}
	if( sApplicationStage.length > 0 )
	{
		<% /* can we route to destination? (see include\LockApp.asp for definition of return values) */ %>
		<% /* PSC 25/10/2005 MAR300 */ %>
		if(LockApplication(m_sReadOnly, m_sApplicationNumber, m_sApplicationFactFindNumber, null, null, true) != 0)
		{
			<% /* BMIDS624 */ %>
			<% /* JD BMIDS749 removed   CheckForChangedRates(m_sApplicationNumber, m_sApplicationFactFindNumber); */ %>
			<% /* BMIDS624 End */ %>
			frmToTM030.submit();
		}
		else
			m_sReadOnly = "";
	}
	else 
		alert("Unable to set ApplicationStage - it is empty");
}

function frmScreen.btnPrint.onclick()
{	<% /* PB 14/09/2006 EP602 Begin */ %>
	if(m_iTotalRecords>500)
	{
		var iPages=parseInt(((m_iTotalRecords+6)/28)+1);
		if(!confirm('Warning - this printout will be approximately ' + iPages + ' pages in size.\r Click "OK" to continue or "Cancel" to abort.'))
			return;
	}
	<% /* PB 14/09/2006 EP602 End */ %>
	var GlobalXML = null;
	GlobalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var iDocumentID = parseInt(GlobalXML.GetGlobalParameterAmount(document,"PrintOutTasksHostTemplateId"));

	AttribsXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AttribsXML.CreateRequestTag(window, "GetPrintAttributes");
	AttribsXML.CreateActiveTag("FINDATTRIBUTES");
	AttribsXML.SetAttribute("HOSTTEMPLATEID", iDocumentID);
				
	<% /*  	AttribsXML.RunASP(document, "PrintManager.asp"); */ %>
	<% /*  Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	switch (ScreenRules())
		{
		case 1: <% /* Warning */ %>
		case 0: <% /* OK */ %>
			AttribsXML.RunASP(document, "PrintManager.asp");
			break;
		default: <% /* Error */ %>
			AttribsXML.SetErrorResponse();
		}

	if(AttribsXML.IsResponseOK())
	{
		<% /* MAR7 GHun */ %>
		var sDeliveryType = AttribsXML.GetTagAttribute("ATTRIBUTES", "DELIVERYTYPE")
		var sCompressionMethod = "";	<% /* MAR796 GHun turn off compression */ %>
		<% /* MAR7 End */ %>
	
		if(AttribsXML.GetTagAttribute("ATTRIBUTES", "INACTIVEINDICATOR") == 1)
		{
			alert("This document template " + AttribsXML.GetTagAttribute("ATTRIBUTES", "HOSTTEMPLATENAME") + " is inactive.");		
		}
		else
		{	
			<% /* MAR7 GHun */ %>
			var printerTypeId = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
			var printerType = getPrinterType(window, printerTypeId);
			var defaultPrinter = getDefaultPrinterFromTypeId(window, printerTypeId);

			if (defaultPrinter != null)
			{
			
			<% /*
			var sPrinterType = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
			var ValidXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			if (ValidXML.IsInComboValidationList(document,"PrinterDestination",sPrinterType,["L"]) == false)
			{				
				alert("The document template printer destination has been defined incorrectly.  See your system administrator.");
			}
			else
			{
				//Call PrintBO.PrintDocument
		
				var sLocalPrinters = GetLocalPrinters();
				LocalPrintersXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				LocalPrintersXML.LoadXML(sLocalPrinters);
				LocalPrintersXML.SelectTag(null, "RESPONSE");	
		
				var sPrinter = LocalPrintersXML.GetTagText("PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
								
				if(sPrinter == "")
				{					
					alert("You do not have a default printer set on your PC.");			
				}		
				else
				{
					TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
					var sGroups = new Array("PrinterDestination");
					var sList = TempXML.GetComboLists(document,sGroups);
					var sValueID = AttribsXML.GetTagAttribute("ATTRIBUTES", "PRINTERDESTINATIONTYPE");
					var sPrintType = TempXML.GetTagText("LIST/LISTNAME/LISTENTRY[VALUEID='" + sValueID + "']/VALIDATIONTYPELIST/VALIDATIONTYPE");
				MAR7 End */ %>			
				PrintXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				PrintXML.CreateRequestTag(window, "PrintDocument");
				PrintXML.SetAttribute("PRINTINDICATOR", printerType == "W" ? "0" : "1");	<% /* MAR7 GHun */ %>
	
				PrintXML.CreateActiveTag("TEMPLATEDATA");
				PrintXML.CreateActiveTag("PRINTDOCSEARCH");
				
				<% /* MAR7 GHun */ %>
				var sPrintText = frmScreen.cboUnitName.options(frmScreen.cboUnitName.selectedIndex).text;
				if (sPrintText == "<ALL>")
				{
					sPrintText = "";
				}
				PrintXML.SetAttribute("SEARCHUNITNAME", sPrintText);
				
				sPrintText = frmScreen.cboUserName.value;
				if (sPrintText == "<ALL>")
				{
					sPrintText = "";
				}
				PrintXML.SetAttribute("SEARCHUSERNAME", sPrintText);
					
				sPrintText = frmScreen.cboTaskType.options(frmScreen.cboTaskType.selectedIndex).text;
				if (sPrintText == "<ALL>")
				{
					sPrintText = "";
				}	
				PrintXML.SetAttribute("SEARCHTASKTYPE", sPrintText); 
				<% /* MAR7 End */ %>
				
				PrintXML.SetAttribute("SEARCHAPPNUMBER", frmScreen.txtApplicationNum.value);
				PrintXML.SetAttribute("SEARCHDUEDATESTART", frmScreen.txtDueDateStart.value);
				PrintXML.SetAttribute("SEARCHDUEDATEEND", frmScreen.txtDueDateEnd.value);
				
				sPrintText = frmScreen.cboSalesLeadStatus.options(frmScreen.cboSalesLeadStatus.selectedIndex).text;
				if (sPrintText == "<ALL>")
				{
					sPrintText = "";
				}	
				PrintXML.SetAttribute("SEARCHSALESLEADSTATUS", sPrintText);
				PrintXML.SetAttribute("SEARCHSLAEXPIRYWITHIN", frmScreen.txtSLAExpiryWithin.value);

				var iTotalRecords = searchXML.ActiveTagList.length;		<% /* MAR7 GHun */ %>
								
				for (var iCount = 0; iCount < iTotalRecords; iCount++)	<% /* MAR7 GHun */ %>
				{		
					PrintXML.SelectTag(null, "TEMPLATEDATA");
					PrintXML.CreateActiveTag("PRINTDOCRESULT");
					searchXML.SelectTagListItem(iCount);			
					PrintXML.SetAttribute("RESULTAPPNUMBER", searchXML.GetAttribute("CASEID"));
					PrintXML.SetAttribute("RESULTTASK", searchXML.GetAttribute("CASETASKNAME")); // BMIDS739 22/09/2004 Was TASKNAME "wrong value"
					PrintXML.SetAttribute("RESULTSTATUS", searchXML.GetAttribute("TASKSTATUS_TEXT"));
					PrintXML.SetAttribute("RESULTDATETIME", searchXML.GetAttribute("TASKDUEDATEANDTIME"));
				}
				
				// PB 13/09/2006 EP602
				var xslTransform = new ActiveXObject("Microsoft.XMLDOM")
				xslTransform.async = false
				xslTransform.load("XML/TaskList.xslt")
				HTMLToPrint=PrintXML.XMLDocument.transformNode(xslTransform);
				window.open('TM020_PrintFrame.asp','winTaskList','width=800,height=600,resizable=no,scrollbars=yes,location=no,copyhistory=no');
			}
		}
	}		
}

<% /* PB 13/09/2006 EP602 */ %>
function getHTMLToPrint()
{
	return HTMLToPrint;
}

<% /* BMIDS624 GHun 22/10/2003 */ %>
function CheckForChangedRates(sApplicationNumber, sApplicationFactFindNumber)
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,"");
	XML.CreateActiveTag("QUOTATION");
	XML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	XML.RunASP(document, "HaveRatesChanged.asp");
	if (XML.IsResponseOK())
	{
		if (XML.GetTagAttribute("QUOTATION", "RATESINCONSISTENT") == "1")
			alert("Rates are inconsistent on your current accepted or active quotation. Please remodel.");	
	}

}
<% /* BMIDS624 End */ %>


function PopulateTaskDescription(strTaskType)
{
	<% /* Get all Task descriptions for this task type */ %>
	var TaskXML = null;

	TaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	TaskXML.CreateRequestTag(window , "FindTaskDescriptions");
	TaskXML.CreateActiveTag("TASKLIST");

	TaskXML.SetAttribute("TASKTYPE",strTaskType);

	TaskXML.RunASP(document, "MsgTMBO.asp");

	<% /* Populate the task description combo. */ %>
	<% /* PSC 28/11/2005 MAR307 - Start */ %>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = TaskXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		var strTaskName;
		var strTaskID;
		var intCounter;
		var intCount;

		TaskXML.CreateTagList("TASK");
		intCount = TaskXML.ActiveTagList.length;

		if( intCount > 0 )
		{
			AddComboOption( frmScreen.cboTaskDescription, "<ALL>","" );

			for( intCounter = 0; intCounter < intCount; intCounter++ )
			{			
				TaskXML.SelectTagListItem(intCounter);
				strTaskName = TaskXML.GetAttribute("TASKNAME")
				strTaskID = TaskXML.GetAttribute("TASKID")
				
				AddComboOption( frmScreen.cboTaskDescription, strTaskName, strTaskID);
			}

			<% /*Set the selection to the first one */ %>
			frmScreen.cboTaskDescription.selectedIndex = 0;
		}
	}
	<% /* PSC 28/11/2005 MAR307 - End */ %>
}

function frmScreen.cboTaskType.onchange()
{
	<% // clear Task description combo %>
	while (frmScreen.cboTaskDescription.options.length > 0)
		frmScreen.cboTaskDescription.options.remove(0);

	var strTaskType = GetTaskType();
	
	if( strTaskType.length > 0)
	{
		PopulateTaskDescription(strTaskType);
		frmScreen.cboTaskDescription.disabled = false;
	}
	else
		frmScreen.cboTaskDescription.disabled = true;
}
<% /* BMIDS702 End */ %>
<% /* PSC 22/08/2005 MAR32 - Start */ %>
function PopulateCombos()
{
	var comboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("TaskStatus", "TaskType", "SalesLeadStatus" );

	var blnSuccess = comboXML.GetComboLists(document, sGroups);
	if (blnSuccess)
	{	
		if(comboXML.PopulateCombo(document,frmScreen.cboTaskType ,"TaskType" , false, true))
		{
			if (frmScreen.cboTaskType.options.length > 1)
				AddComboOption(frmScreen.cboTaskType, "<ALL>","", 0);
			
			frmScreen.cboTaskType.selectedIndex = 0
			frmScreen.cboTaskType.onchange();
		}
		
		if (comboXML.PopulateCombo(document,frmScreen.cboSalesLeadStatus, "SalesLeadStatus", false, true))
		{
			if (frmScreen.cboSalesLeadStatus.options.length > 1)
				AddComboOption(frmScreen.cboSalesLeadStatus, "<ALL>","", 0);
			
			frmScreen.cboSalesLeadStatus.selectedIndex = 0;
		}
	
		var xmlTaskStatusList = comboXML.GetComboListXML("TaskStatus");
		var xmlNotIncompleteList = xmlTaskStatusList.selectNodes("LISTNAME/LISTENTRY[not(VALIDATIONTYPELIST[VALIDATIONTYPE='I'])]");
		
		for (var nIndex = 0; nIndex < xmlNotIncompleteList.length; nIndex++)
		{
			xmlNotIncompleteList.item(nIndex).parentNode.removeChild(xmlNotIncompleteList.item(nIndex));
		}
		
		
		var undefinedValue;
		if(comboXML.PopulateComboFromXML(document, frmScreen.cboTaskStatus, xmlTaskStatusList, false, undefinedValue, undefinedValue, undefinedValue, undefinedValue, true))
		{
			if (frmScreen.cboTaskStatus.options.length > 1)
			{
				AddComboOption(frmScreen.cboTaskStatus, "<ALL>","", 0);
				frmScreen.cboTaskStatus.selectedIndex = 1;
			}
			else if	(frmScreen.cboTaskStatus.options.length == 1)
				frmScreen.cboTaskStatus.selectedIndex = 0;	
		}
	}
	else
	{
		<% /* Display error message */ %>
		alert("Unable to retrieve combo values");
		<% /*Disable screen. */ %>				
		scScreenFunctions.SetScreenToReadOnly(frmScreen)
		frmScreen.btnSearch.disabled = true;
	}
}

function GetUserRoleDetails()
{
	var GlobalXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_nMinNSALevel = parseInt(GlobalXML.GetGlobalParameterAmount(document,"TMMinNSAAuthLevel"));
	m_nMinSALevel = parseInt(GlobalXML.GetGlobalParameterAmount(document,"TMMinSAAuthLevel"));
	
	var userRoleXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bIsNSARole = userRoleXML.IsInComboValidationList(document,"UserRole", m_nUserRole, ["NSA"]);
	m_bIsSARole = userRoleXML.IsInComboValidationList(document,"UserRole", m_nUserRole, ["SA"]);
}
<% /* PSC 22/08/2005 MAR32 - End */ %>


-->
</script>
</body>
</html>
