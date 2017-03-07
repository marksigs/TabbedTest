<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      CR025.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Contact History Summary 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
CL		16/01/2001	SYS1756 Screen first created
CL		05/03/01	SYS1920 Read only functionality added
CL		08/03/01	SYS1981 Adjustment to Edit facility 
CL		15/03/01	SYS2044 Re-Routing changes
CL		15/03/01	SYS2057 Sorting changes
LD		23/05/02	SYS4727 Use cached versions of frame functions


BMIDS Specific History:

Prog	Date		AQR			Description
MV		13/05/2002	BMIDS00004	Modified frmScreen.btnEdit.onclick(),RetrieveContextData()
MV		14/05/2002	BMIDS00004	Modified alert Message in frmScreen.btnEdit.onclick()
MV		14/05/2002	BMIDS00004	Modified frmScreen.btnEdit.onclick(),Added DisableButtons();
GHun	16/05/2002  BMIDS00005  Added View and Get Contact Data buttons, disabled Edit for validation type C
MV		16/08/2002  BMIDS00194	CRWP2 BM065 - Disable btnGetContactData
GHun	11/09/2002	BMIDS00442	Get CIF number from context
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		09/10/2002	BMIDS00556	Removed Validation_Init()
MV		07/11/2002	BMIDS00723	Altered the coordinates of GN450.asp
GHun	02/12/2002	BMIDS01121	Changed MetaAction on cancel to avoid locking problem
MC		19/04/2004	CC057		cr450 popup Dialog height increased to fit the text (+10)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<% /* Scriptlets - remove any which are not required */ %>
<% /* CORE UPGRADE 702 <object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object> */ %>
<object data="scTable.htm" height="1px" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 275px; LEFT: 310px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scTableList" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Specify Forms Here */ %>
<form id="frmToCR020" method="post" action="cr020.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCR040" method="post" action="cr040.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCR030" method="post" action="cr030.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Start of form */ %>

<form id="frmScreen" mark validate="onchange">
	<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 246px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
		<div style="TOP: 10px; LEFT: 4px; HEIGHT: 23px; WIDTH: 55px; POSITION: ABSOLUTE" >
			<span style="TOP:4px; LEFT:4px; WIDTH: 80px; POSITION:ABSOLUTE" class="msgLabel">
				Customer Name
				<span style="TOP:-3px; LEFT:85px; POSITION:ABSOLUTE">
					<input id="txtCustomer" style="WIDTH:315px" class="msgTxt">
				</span>
			</span>		
			<span style="TOP:4px; LEFT:410px; WIDTH: 100px; POSITION:ABSOLUTE" class="msgLabel">
				Sort List By...
				<span style="TOP:-3px; LEFT:75px; POSITION:ABSOLUTE">
					<select id="cboSortBy" style="WIDTH:111px" class="msgCombo"></select>
				</span>
			</span>	
		</div>
		<span id="spnContactHistoryList" style="TOP: 40px; LEFT: 4px; POSITION: ABSOLUTE">
			<table id="tblContactHistoryList" width="596px" border="0" cellspacing="0" cellpadding="0" class="msgTable">				
				<tr id="rowTitles"><td width="23%" class="TableHead">Date/Time</td><td width="20%" class="TableHead">Contact Reason</td><td width="20%" class="TableHead">User Name</td><td width="32%" class="TableHead">Details</td><td width="5%"class="TableHead">Deleted?</td></tr>				
				<tr id="row01"><td width="23%" class="TableTopLeft">&nbsp</td><td width="20%" class="TableTopCenter">&nbsp</td><td width="20%" class="TableTopCenter">&nbsp</td><td width="30%" class="TableTopCenter">&nbsp</td><td width="5%" class="TableTopRight">&nbsp</td></tr>
				<tr id="row02"><td width="23%" class="TableLeft">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="32%" class="TableCenter">&nbsp</td><td width="5%" class="TableRight">&nbsp</td></tr>
				<tr id="row03"><td width="23%" class="TableLeft">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="32%" class="TableCenter">&nbsp</td><td width="5%" class="TableRight">&nbsp</td></tr>
				<tr id="row04"><td width="23%" class="TableLeft">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="32%" class="TableCenter">&nbsp</td><td width="5%" class="TableRight">&nbsp</td></tr>
				<tr id="row05"><td width="23%" class="TableLeft">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="32%" class="TableCenter">&nbsp</td><td width="5%" class="TableRight">&nbsp</td></tr>
				<tr id="row06"><td width="23%" class="TableLeft">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="32%" class="TableCenter">&nbsp</td><td width="5%" class="TableRight">&nbsp</td></tr>
				<tr id="row07"><td width="23%" class="TableLeft">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="32%" class="TableCenter">&nbsp</td><td width="5%" class="TableRight">&nbsp</td></tr>
				<tr id="row08"><td width="23%" class="TableLeft">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="32%" class="TableCenter">&nbsp</td><td width="5%" class="TableRight">&nbsp</td></tr>
				<tr id="row09"><td width="23%" class="TableLeft">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="32%" class="TableCenter">&nbsp</td><td width="5%" class="TableRight">&nbsp</td></tr>
				<tr id="row10"><td width="23%" class="TableBottomLeft">&nbsp</td><td width="20%" class="TableBottomCenter">&nbsp</td><td width="20%" class="TableBottomCenter">&nbsp</td><td width="32%" class="TableBottomCenter">&nbsp</td><td width="5%" class="TableBottomRight">&nbsp</td></tr>
			</table>
		</span>
		<span id="spnButtons" style="TOP: 215px; LEFT: 4px; POSITION: ABSOLUTE">
			<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
				<input id="btnAdd" value="Add" type="button" 
					style="WIDTH: 60px" class="msgButton">
			</span>
			<% /* BMIDS00005 Added View button */ %>
			<span style="TOP: 0px; LEFT: 64px; POSITION: ABSOLUTE">
				<input id="btnView" value="View" type="button" 
					style="WIDTH: 60px" class="msgButton">
			</span>
			<span style="LEFT: 128px; POSITION: absolute; TOP: 0px">
				<input id="btnEdit" value="Edit" type="button" 
					style="WIDTH: 60px" class="msgButton">
			</span>
			<% /* BMIDS00005 Added Get Contact Data button */ %>
			<span style="LEFT: 192px; POSITION: absolute; TOP: 0px">
				<input id="btnGetContactData" value="Get Contact Data" type="button" 
					style="WIDTH: 96px" class="msgButton">
			</span>
		</span>
	</div>
</form>

<% /* End of form */ %>


<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 320px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cr025Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_iTableLength = 10;
var m_iTotalRecords = 0;
var m_strCustomerNumber;
var XML = null;
var scScreenFunctions;
var m_sReadOnly = null;
var m_sCustomerName;  
var m_sContext;
var m_blnReadOnly = false;
var m_sUserRole = "";
var m_sCIFNumber;
var sContactHistoryEditAuthority;
var ContactReasonXML;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	/* CORE UPGRADE 702 scScreenFunctions = new scScrFunctions.ScreenFunctionsObject(); */
	FW030SetTitles("Customer Contact History","CR025",scScreenFunctions);

	GetComboLists();
	RetrieveContextData();
	PopulateTable();
	<% /* BMIDS00005 DisableButtons functionality has been moved to SetEditButtonStatus
	DisableButtons();
	 */ %>
	scScreenFunctions.SetFieldState(frmScreen, "txtCustomer", "R");
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "CR025");
	frmScreen.btnEdit.disabled = true;
	<% /* BMIDS00005 View button should be disabled until a row is selected */ %>
	frmScreen.btnView.disabled = true;
	
	<% /* BMIDS00005 Get the Contact History Edit Authority global parameter */ %>
	/* CORE UPGRADE 702 var GlobalParamXML = new scXMLFunctions.XMLObject(); */
	var GlobalParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	sContactHistoryEditAuthority = GlobalParamXML.GetGlobalParameterAmount(document,"ContactHistoryEditAuthority");
	
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
	m_sContext = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName");  		
	m_sUserRole = scScreenFunctions.GetContextParameter(window,"idRole");
	
	if(m_sCustomerName == "")
	{
		m_sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName1");
	}	  			
	
	frmScreen.txtCustomer.value = m_sCustomerName;
	m_strCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber");
	
	<% /* BMIDS00442 */ %>
	m_sCIFNumber = scScreenFunctions.GetContextParameter(window,"idOtherSystemCustomerNumber");
	<% /* BMIDS00442 End */ %>
	<% /* Get Contact Data button is only enabled if there is a CIF Number */ %>
	if (m_sCIFNumber == "")
		frmScreen.btnGetContactData.disabled = true;
	
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	
	if (m_sReadOnly == "1")
	{
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnEdit.disabled = true;
		//frmScreen.btnView.disabled = true;
		frmScreen.btnGetContactData.disabled = true;
	}
	
	<%/* MV - 16/08/2002 - BMIDS00194 - CRWP2 BM065 */%>
	if ( m_sMetaAction == "CreateNewCustomerForNewApplication")
		frmScreen.btnGetContactData.disabled = true;
}

function SelectData() // this extracts the selected Date/Time of the history record 
{
	var iXMLKey = scTableList.getOffset() + scTableList.getRowSelected();
	if (iXMLKey != "-1"); //A row has been selected 
	{	
		XML.SelectTagListItem(iXMLKey-1);
		SendData(XML.ActiveTag.xml);
	}
}

function SendData(sContactHistoryRecordXML) // this selects the tag from XML for the record in SelectData
{
	// now set the extracted data to the array
	var sReturn = null;
	var strUserName;
	var strUnitName;
	var ArrayArguments = new Array();
	ArrayArguments[0] = XML.CreateRequestAttributeArray(window);	
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = sContactHistoryRecordXML;	
	ArrayArguments[3] = m_sMetaAction;
	ArrayArguments[4] = m_sCustomerName;
	
	if (m_sMetaAction == "Add")
	{
		ArrayArguments[5] = scScreenFunctions.GetContextParameter(window,"idUserName");
		ArrayArguments[6] = scScreenFunctions.GetContextParameter(window,"idUnitName");
		ArrayArguments[7] = scScreenFunctions.GetContextParameter(window,"idCustomerNumber");
	}
	
	ArrayArguments[8] = scScreenFunctions.GetContextParameter(window,"idUserID");
	ArrayArguments[9] = scScreenFunctions.GetContextParameter(window,"idUnitID");
	ArrayArguments[10] = scScreenFunctions.GetContextParameter(window,"idRole");
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "gn450.asp", ArrayArguments, 630, 375);
			
	if (sReturn != null)
	{
		if (sReturn[0] == true)
		{
			PopulateTable();
		}
	}
}

function PopulateTable()
{		
	/* CORE UPGRADE 702 XML = new scXMLFunctions.XMLObject(); */
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
	/*
	<REQUEST>
		<CONTACTHISTORY>
			<CUSTOMERNUMBER>
			<SORTORDER>	
	*/
		
	XML.CreateRequestTag(window,"REQUEST");
	XML.CreateActiveTag("CONTACTHISTORY");
	XML.CreateTag("CUSTOMERNUMBER",m_strCustomerNumber);
	XML.CreateTag("SORTORDER",frmScreen.cboSortBy.value);
	// 	XML.RunASP(document,"FindContactHistoryList.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"FindContactHistoryList.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
		
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[0] == true)
		{
			/* CORE UPGRADE 702 if (XML.CreateTagList("CONTACTHISTORY") != null) 
				m_iTotalRecords = XML.ActiveTagList.length;
			else 
				m_iTotalRecords = 0;
			scTableList.initialiseTable(tblContactHistoryList, 0, "", ShowList, m_iTableLength, m_iTotalRecords);*/
			if (XML.CreateTagList("CONTACTHISTORY") != null) m_iTotalRecords = XML.ActiveTagList.length;
			else m_iTotalRecords = 0;
			scTableList.initialiseTable(tblContactHistoryList, 0, "", ShowList, m_iTableLength, m_iTotalRecords)
			
			ShowList(0);
		}
	}			
	ErrorTypes = null;
	ErrorReturn = null;
			
}

function ShowList(iOffset)
{
	var strDateTime;
	var strContactReason;
	var strUserID;
	var strDetails;
	var strDeleted;
	
	for (var iCount=0; iCount < m_iTotalRecords && iCount < m_iTableLength; iCount++)
	{
		XML.SelectTagListItem(iCount+iOffset);
		strDateTime = XML.GetTagText("CONTACTHISTORYDATETIME");
		strContactReason = XML.GetTagAttribute("CONTACTREASONCODE","TEXT");
		strUserID = XML.GetTagText("USERNAME");
		strDetails = XML.GetTagText("CONTACTTEXT");
		strDeleted = XML.GetTagText("STATUSINDICATOR");
					
		scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(0), strDateTime);
		scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(1), strContactReason);
		scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(2), strUserID);
		scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(3), strDetails);
				
		if (strDeleted == "1")
		{
			scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(4), "Yes");
		}
		else
		{
			scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(4), "");
		}
				
		if (m_sReadOnly == "1")
		{
			frmScreen.btnEdit.disabled = true;
		}

		<% /* BMIDS00005 btnEdit should not be enabled just because the form is not read only	
			as other conditions may have disabled it 
		else
		{
			frmScreen.btnEdit.disabled = false;
		}
		*/ %>
	}	
}

function frmScreen.cboSortBy.onchange()
{
	scTableList.clear();		
	PopulateTable();
	frmScreen.btnEdit.disabled = true;
}

function GetComboLists()
{
	/* CORE UPGRADE 702 var XML = new scXMLFunctions.XMLObject(); */
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sSortBy = new Array("ContactReasonSortOrder");
	var sContactReason = new Array("ContactHistoryReason");
	var bSuccess = false;

	<% /* BMIDS00005 Get ContactReasonXML once now, for multiple use later */ %>
	/* CORE UPGRADE 702 ContactReasonXML = new scXMLFunctions.XMLObject(); */
	ContactReasonXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ContactReasonXML.GetComboLists(document,sContactReason);

	if(XML.GetComboLists(document,sSortBy))
	{
		bSuccess = XML.PopulateCombo(document,frmScreen.cboSortBy,"ContactReasonSortOrder",false);
		if (bSuccess)
			frmScreen.cboSortBy.value = "2";
	}

	if(!bSuccess)
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		DisableMainButton("Submit");
	}
}

function frmScreen.cboSortBy.onchange()
{
	scTableList.clear();		
	PopulateTable();
	frmScreen.btnEdit.disabled = true;
	frmScreen.btnView.disabled = true;
}

function frmScreen.btnAdd.onclick()
{
	<% //scScreenFunctions.SetContextParameter(window,"idMetaAction","Add"); %>
	m_sMetaAction = "Add";
	SendData("");
	<% //scScreenFunctions.SetContextParameter(window,"idMetaAction",m_sContext);  %>
}

function frmScreen.btnEdit.onclick()
{	
	<% //scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit"); %>
	m_sMetaAction = "Edit";
	SelectData();
	<% //scScreenFunctions.SetContextParameter(window,"idMetaAction",m_sContext); %>
}

<% /* BMIDS00005 View button click event */ %>
function frmScreen.btnView.onclick()
{	
	m_sMetaAction = "View";
	SelectData();
}

<% /* BMIDS00005 Get Contact Data button click event */ %>
function frmScreen.btnGetContactData.onclick()
{
	<% /* Get the last CRS date before clearing the data */ %>
	var sLastCRSDate = GetLastCRSDate();

	/* CORE UPGRADE 702 var NewXML = new scXMLFunctions.XMLObject(); */
	var NewXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	NewXML.CreateRequestTag(window,"REQUEST");
	NewXML.CreateActiveTag("CONTACTLOG");
	NewXML.SetAttribute("OMIGACUSTOMERNUMBER", m_strCustomerNumber);
	NewXML.SetAttribute("CIFNUMBER", m_sCIFNumber);
	NewXML.SetAttribute("LASTCRSDATE", sLastCRSDate);
	NewXML.SetAttribute("SORTORDER", frmScreen.cboSortBy.value);
	// 	NewXML.RunASP(document, "GetCRSContactData.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			NewXML.RunASP(document, "GetCRSContactData.asp");
			break;
		default: // Error
			NewXML.SetErrorResponse();
		}

	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = NewXML.CheckResponse(ErrorTypes);

	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[0] == true)
		{
			scTableList.clear();
			NewXML.ActiveTag = null;
		
			if (NewXML.CreateTagList("CONTACTHISTORY") != null)
				m_iTotalRecords = NewXML.ActiveTagList.length;
			else
				m_iTotalRecords = 0;
			
			XML = NewXML;
			scTableList.initialiseTable(tblContactHistoryList, 0, "", ShowList, m_iTableLength, m_iTotalRecords);
			ShowList(0);
		}
	}			
	ErrorTypes = null;
	ErrorReturn = null;
	
	frmScreen.btnEdit.disabled = true;
	frmScreen.btnView.disabled = true;
}

function btnSubmit.onclick()
{
	//Get context value of application number
	var sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber");  	
	
	if (sApplicationNumber != "")
	{
		frmToCR030.submit();
	}
	else
	{
		frmToCR040.submit();
	}
}

function btnCancel.onclick()
{
	<% /* BMIDS01121 When re-entering CR020 with a new customer, the MetaAction should be existing, not new */ %>
	if (m_sMetaAction == "CreateNewCustomerForNewApplication")
	{
		m_sMetaAction = "CreateExistingCustomerForNewApplication";
		scScreenFunctions.SetContextParameter(window, "idMetaAction", m_sMetaAction);
	}
	<% /* BMIDS01121 End */ %>
	frmToCR020.submit();
}

function CheckRowSelected()
{
	var iIndex = scTableList.getRowSelectedIndex();  // returns the table index 
	
	if (iIndex != null)
	{   
		frmScreen.btnView.disabled = false;
		SetEditButtonStatus();
	}
	else
	{
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnView.disabled = true;
	}
}

function spnContactHistoryList.onclick()
{
	CheckRowSelected();
}	

<% /* BMIDS00005 Get the last date a CRS record was inserted into the table */ %>
function GetLastCRSDate()
{
	var dtMaxCRSDate = new Date("1 Jan 1800");
	var sContactDate;
	var dtContactDate;
	var iContactReason;
	
	for (var iRow=0; iRow < m_iTotalRecords; iRow++)
	{
		XML.SelectTagListItem(iRow);
		iContactReasonCode = XML.GetTagText("CONTACTREASONCODE");
		ContactReasonXML.ActiveTag = null;
		ContactReasonXML.SelectSingleNode("RESPONSE/LIST/LISTNAME/LISTENTRY[VALUEID[.='" + iContactReasonCode + "']]/VALIDATIONTYPELIST/VALIDATIONTYPE[.='C']");
		if (ContactReasonXML.ActiveTag != null)
		{	
			sContactDate = XML.GetTagText("CONTACTHISTORYDATETIME");
			dtContactDate = scScreenFunctions.GetDateObjectFromString(sContactDate);
			if (dtContactDate > dtMaxCRSDate)
				dtMaxCRSDate = dtContactDate;
		}
	}

	if (dtMaxCRSDate.toString() != (new Date("1 Jan 1800").toString()))
		return scScreenFunctions.DateTimeToString(dtMaxCRSDate);
	else
		return "";
}	

<% /* BMIDS00005 Either enable or disable the Edit button */ %>
function SetEditButtonStatus()
{	
	<% /* Enable the Edit button by default */ %>
	var blnEditDisabled = false;			
	
	<% /* Disable it if the screen is read only */ %>
	if (m_sReadOnly == "1")
		blnEditDisabled = true;
	else
	{
		<% /* Disable it if there is insufficient authority */ %>
		if (parseInt(m_sUserRole) < parseInt(sContactHistoryEditAuthority))
			blnEditDisabled = true;
		else
		{
			<% /* Disable it if the contact reason is CRS (validation type = 'C') */ %>
			var iXMLKey = scTableList.getOffset() + scTableList.getRowSelected();
			XML.SelectTagListItem(iXMLKey-1);
			var iContactReasonCode = XML.GetTagText("CONTACTREASONCODE");
				
			ContactReasonXML.ActiveTag = null;	
			ContactReasonXML.SelectSingleNode("RESPONSE/LIST/LISTNAME/LISTENTRY[VALUEID[.='" + iContactReasonCode + "']]/VALIDATIONTYPELIST/VALIDATIONTYPE[.='C']");
			if (ContactReasonXML.ActiveTag != null)
				blnEditDisabled = true;
		}
	}
			
	frmScreen.btnEdit.disabled = blnEditDisabled;
}

-->
</script>
</body>
</html>




