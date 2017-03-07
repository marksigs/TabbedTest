<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      GN400.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Contact History Summary 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
CL		22/01/2001	SYS1756 Sceen first created
DPF		20/06/2002	BMIDS00077  Altered file to be in line with Core V7.0.2
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		10/05/2002	BMIDS00004	Modified frmScreen.btnEdit.onclick(),RetrieveData()
MV		15/05/2002	BMIDS00004	Modified frmScreen.btnEdit.onclick(),Added DisableButtons()
GHun	22/05/2002	BMIDS00005  Added View and Get Contact Data buttons, disabled Edit for validation type C
GHun	28/05/2002	BMIDS00108	Fix errors introduced by Core upgrade
GHun	13/09/2002	BMIDS00442	Get CIF number for passing to GetCRSContactData
GHun	23/09/2002	BMIDS00506	m_iTotalRecords not always initialised when changing between customers
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		09/10/2002	BMIDS00556	Removed Validation_Init()
MO		13/11/2002	BMIDS00723	Made changes after the font sizes were changed
MC		19/04/2004	CC057		Contact History Details screen dialog size changed
								White spaces padded to hide std. IE message on the title.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Contact History  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets - remove any which are not required
remove these two as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
 */ %>
<object data="scTable.htm" height="1px" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 220px; LEFT: 300px; POSITION: absolute">
	<object data="scTableListScroll.asp" id="scTableList" style="LEFT: 0px; TOP: 0px; HEIGHT: 24px; WIDTH: 304px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
</span>

<% /* Start of form */ %>
<form id="frmScreen" style="VISIBILITY: visible" mark validate="onchange" year4>
	<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 270px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
		<div style="TOP: 10px; LEFT: 4px; HEIGHT: 23px; WIDTH: 55px; POSITION: ABSOLUTE" >
			<span style="TOP:4px; LEFT:4px; WIDTH: 80px; POSITION:ABSOLUTE" class="msgLabel">
				Customer Name
				<span style="TOP:-3px; LEFT:85px; POSITION:ABSOLUTE">
					<Select id="cboCustomer" style="WIDTH:315px" class="msgCombo"></Select>
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
		<span id="spnButtons" style="TOP: 240px; LEFT: 4px; POSITION: ABSOLUTE">
			<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
				<input id="btnAdd" value="Add" type="button" 
					style="WIDTH: 60px" class="msgButton">
			</span>
			<% /* BMIDS00005 Added View button */ %>
			<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">
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
			<span style="TOP: 0px; LEFT: 535px; POSITION: ABSOLUTE">
				<input id="btnClose" value="Close" type="button" 
					style="WIDTH: 60px" class="msgButton">
			</span>
		</span>
	</div>
</form>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/gn400Attribs.asp" -->

<% /* End of form */ %>

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_iTableLength = 10;
var m_iTotalRecords = 0;
//var m_strCustomerNumber;
var XML = null;
var scScreenFunctions;
var m_sReadOnly = "";
//var m_sCustomerName = "";
//var m_sCustomerNumber = "";
var m_sCustomerNameArray = null;
var m_sCustomerNumberArray = null;
var m_sCustomerVersionNumberArray = null; <% /* BMIDS00442 */ %>
var m_BaseNonPopupWindow = null;
var m_sUserRole= "";
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
	
	RetrieveData();
	
	GetComboLists();
		
	SetScreenToReadOnly();
	
	PopulateTable();
	SetGetContactDataButtonStatus();	<% /* BMIDS00442 */ %>
	
	<% /* BMIDS00005 DisableButtons functionality has been moved to SetEditButtonStatus
	DisableButtons();
	 */ %>
	
	window.returnValue = null;
	
	frmScreen.btnEdit.disabled = true;
	<% /* BMIDS00005 View button should be disabled until a row is selected */ %>
	frmScreen.btnView.disabled = true;
	
	<% /* BMIDS00005 Get the Contact History Edit Authority global parameter */ %>
	<% //removed as per Core V7.0.2 - DPF 10/06/02 - BMIDS00077
	//var GlobalParamXML = new scXMLFunctions.XMLObject();
	//sContactHistoryEditAuthority = GlobalParamXML.GetGlobalParameterAmount(document,"ContactHistoryEditAuthority");
	//BMIDS000108 The above 2 lines should not have been commented out, but changed to the following %>
	var GlobalParamXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	sContactHistoryEditAuthority = GlobalParamXML.GetGlobalParameterAmount(document,"ContactHistoryEditAuthority");
	<% /* BMIDS00108 End */ %>
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveData()
{		
	var sArguments			= window.dialogArguments;
	
	window.dialogTop		= sArguments[0]; 
	window.dialogLeft		= sArguments[1];
	window.dialogWidth		= sArguments[2];
	window.dialogHeight		= sArguments[3];
		
	var RecordArguments		= sArguments[4];
	
	//next 2 lines added as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	m_BaseNonPopupWindow = sArguments[5];
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
				
	m_sCustomerNameArray	= RecordArguments[0]; // Customer Name Array;
	m_sCustomerNumberArray	= RecordArguments[1]; // Customer Number Array
	m_sReadOnly				= RecordArguments[2]; // ReadOnly
	m_sUserID				= RecordArguments[3]; // UserId
	m_sUnitID				= RecordArguments[4]; // UnitId
	m_sAttributeArray		= RecordArguments[5]; // FW030RequestArray
	m_sUserName				= RecordArguments[6]; // User Name
	m_sUnitName				= RecordArguments[7]; // Unit Name
	m_sUserRole				= RecordArguments[8]; // User Role
	m_sCustomerVersionNumberArray = RecordArguments[9]; <% /* BMIDS00442 Customer Version Number Array */ %>
}

function PopulateCustomerCombo()
{
	for(var nLoop = 0;nLoop < m_sCustomerNameArray.length;nLoop++)
	{
		var objOPTION = document.createElement("OPTION");
		objOPTION.text = m_sCustomerNameArray[nLoop]; //  Name
		objOPTION.value = m_sCustomerNumberArray[nLoop]; // Number
		 <% /* BMIDS00442 */ %>
		var sCIFNumber = GetCIFNumber(m_sCustomerNumberArray[nLoop], m_sCustomerVersionNumberArray[nLoop]);
		if (sCIFNumber.length > 0)
		{
			objOPTION.setAttribute("HasCIFNumber", "true");
			objOPTION.setAttribute("CIFNumber", sCIFNumber);
		}
		else
			objOPTION.setAttribute("HasCIFNumber", "false");
		<% /* BMIDS00442 End */ %>
		frmScreen.cboCustomer.add(objOPTION);
	}	
}

function SetScreenToReadOnly()
{
	scScreenFunctions.ShowCollection(frmScreen);
	if (m_sReadOnly == "1")
	{
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnGetContactData.disabled = true;
	}
}

<% /* this extracts the selected Date/Time of the history record */ %>
function SelectData() 
{
	var iXMLKey = scTableList.getOffset() + scTableList.getRowSelected();
	if (iXMLKey != "-1") 
	{
		XML.SelectTagListItem(iXMLKey-1);
		SendData(XML.ActiveTag.xml);
	}
}

<%/* this selects the tag from XML for the record in SelectData */%>
function SendData(sContactHistoryRecordXML) 
{
	var sReturn = null;
	var strUserName;
	var strUnitName;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sAttributeArray;
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = sContactHistoryRecordXML;	
	ArrayArguments[3] = m_sMetaAction;
	ArrayArguments[4] = frmScreen.cboCustomer.options(frmScreen.cboCustomer.selectedIndex).text;  //customer name
			
	if (m_sMetaAction == "Add")
	{
		ArrayArguments[5] = m_sUserName;
		ArrayArguments[6] = m_sUnitName;
		ArrayArguments[7] = frmScreen.cboCustomer.value;  // customer id
		ArrayArguments[8] = m_sUserID;
		ArrayArguments[9] = m_sUnitID;
	}
	sReturn = scScreenFunctions.DisplayPopup(window, document, "gn450.asp", ArrayArguments, 630, 370);
			
	if (sReturn != null)
	{
		if (sReturn[0] == true)
			PopulateTable();
	}
}

function PopulateTable()
{		
	//next line replaced by line below as per Core V7.0.2
	//XML = new scXMLFunctions.XMLObject();
	XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		
	scTableList.clear();
	m_iTotalRecords = 0; <% /* BMIDS00506 */ %>
	
	XML.CreateRequestTagFromArray(m_sAttributeArray, "SEARCH");
	XML.CreateActiveTag("CONTACTHISTORY");
	XML.CreateTag("CUSTOMERNUMBER",frmScreen.cboCustomer.value);
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
			<% /* BMIDS00506 Set m_iTotalRecords = 0 after clearing the table */ %>
			// if (XML.CreateTagList("CONTACTHISTORY") != null) m_iTotalRecords = XML.ActiveTagList.length;
			// else m_iTotalRecords = 0;
			if (XML.CreateTagList("CONTACTHISTORY") != null)
				m_iTotalRecords = XML.ActiveTagList.length;
			<% /* BMIDS00506 End */ %>
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
		strUserID = XML.GetTagText("USERID");
		strDetails = XML.GetTagText("CONTACTTEXT");
		strDeleted = XML.GetTagText("STATUSINDICATOR");
					
		scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(0), strDateTime);
		scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(1), strContactReason);
		scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(2), strUserID);
		scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(3), strDetails);
		
		if (strDeleted == "1")
			scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(4), "Yes");
		else
			scScreenFunctions.SizeTextToField(tblContactHistoryList.rows(iCount+1).cells(4), "");
	}	
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
	<% /* BMIDS00442 */ %>
	var sCIFNumber = frmScreen.cboCustomer.item(frmScreen.cboCustomer.selectedIndex).getAttribute("CIFNumber");
	if (sCIFNumber.length > 0)
	{
	<% /* BMIDS00442 End */ %>
	
		<% /* Get the last CRS date before clearing the data */ %>
		var sLastCRSDate = GetLastCRSDate();

		<% /* BMIDS00108 Core upgrade (BMIDS00077) should have changed the following */ %>
		<% /* var NewXML = new scXMLFunctions.XMLObject(); */ %>
		var NewXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		<% /* BMIDS00108 End */ %>
				
		//XML.CreateRequestTag(window,"REQUEST");
		NewXML.CreateRequestTagFromArray(m_sAttributeArray, "REQUEST");
		NewXML.CreateActiveTag("CONTACTLOG");
		NewXML.SetAttribute("OMIGACUSTOMERNUMBER", frmScreen.cboCustomer.value);
		NewXML.SetAttribute("CIFNUMBER", sCIFNumber);	<% /* BMIDS00442 */ %>
		NewXML.SetAttribute("LASTCRSDATE", sLastCRSDate);
		NewXML.SetAttribute("SORTORDER", frmScreen.cboSortBy.value);
		// 		NewXML.RunASP(document, "GetCRSContactData.asp");
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
	}	<% /* BMIDS00442 */ %>
}

<% /* BMIDS00005 This fucntion has been split between window.onload() and SetEditButtonStatus() */ %>
<% /* function DisableButtons()
{
	var XML = new scXMLFunctions.XMLObject();
	var sContactHistoryUserRole = XML.GetGlobalParameterAmount(document,"ContactHistoryUserRole");
	
	if (parseInt(m_sUserRole) <  parseInt(sContactHistoryUserRole))
		frmScreen.btnEdit.disabled = true ;
} */ %>

function spnContactHistoryList.onclick()
{ 
	CheckRowSelected();
}

function frmScreen.cboSortBy.onchange()
{
	PopulateTable();
	frmScreen.btnEdit.disabled = true;
}

function frmScreen.cboCustomer.onchange()
{
	PopulateTable();
	frmScreen.btnEdit.disabled = true;
	SetGetContactDataButtonStatus();	<% /* BMIDS00442 */ %>
}

function GetComboLists()
{
	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sSortBy = new Array("ContactReasonSortOrder");
	var sContactReason = new Array("ContactHistoryReason");
	var bSuccess = false;

	<% /* BMIDS000005 Get ContactReasonXML once in the beginning to use later */ %>
	<% //ContactReasonXML = new scXMLFunctions.XMLObject(); %>
	<% /* BMIDS00108 Core upgrade added a var in front of ContactReasonXML which should not have been there */ %>
	<% // var ContactReasonXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject(); %>
	ContactReasonXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	<% /* BMIDS00108 End */ %>
	ContactReasonXML.GetComboLists(document,sContactReason);

	if(XML.GetComboLists(document,sSortBy))
	{
		bSuccess = true;
		bSuccess = XML.PopulateCombo(document,frmScreen.cboSortBy,"ContactReasonSortOrder",false);	
		if (bSuccess)
			frmScreen.cboSortBy.value = "2" ;	
	}

	if(!bSuccess)
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
	
	PopulateCustomerCombo(); 
}

function frmScreen.btnAdd.onclick()
{
	m_sMetaAction = "Add"
	SendData("");
}

function frmScreen.btnEdit.onclick()
{	
	<% /* scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit"); */ %>
	m_sMetaAction = "Edit"
	SelectData();
	<% /* scScreenFunctions.SetContextParameter(window,"idMetaAction",m_sContext); */ %>
}

function frmScreen.btnClose.onclick()
{
	window.close();	
}

function CheckRowSelected()
{
	var RowSelected =(scTableList.getRowSelected())
	if (RowSelected == -1)
	{
		<% /* BMIDS00005 Also disable the View button if no row is selected */ %>
		frmScreen.btnView.disabled = true;
		frmScreen.btnEdit.disabled = true;
	}
	else
	{
		<% /* BMIDS00005 Enable the View button and enable Edit if permitted to */ %>
		frmScreen.btnView.disabled = false;
		SetEditButtonStatus();
	}
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

<% /* BMIDS00442 */ %>
function GetCIFNumber(sCustomerNumber, sCustomerVersionNumber)
{
	<% /* Query GetCustomerDetails to get OtherSystemCustomerNumber */ %>
	var CustXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	CustXML.CreateRequestTagFromArray(m_sAttributeArray, "REQUEST");
	CustXML.CreateActiveTag("SEARCH");
	CustXML.CreateActiveTag("CUSTOMER");
	CustXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
	CustXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
	CustXML.RunASP(document, "GetCustomerDetails.asp");
	if(CustXML.IsResponseOK())
	{
		CustXML.SelectTag(null, "CUSTOMER");
		var sCIFNumber = CustXML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER");
		return sCIFNumber;
	}
	else
		return "";
}

function SetGetContactDataButtonStatus()
{
	<% /* Enable the GetContactData button if there is an OtherSystemCustomerNumber */ %>
	if (frmScreen.cboCustomer.item(frmScreen.cboCustomer.selectedIndex).getAttribute("HasCIFNumber") == "true")
		frmScreen.btnGetContactData.disabled = false;
	else
		frmScreen.btnGetContactData.disabled = true;
}

<% /* BMIDS00442 End */ %>

-->
</script>
</body>
</html>





