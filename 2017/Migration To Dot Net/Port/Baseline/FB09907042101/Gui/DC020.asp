<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/*
Workfile:      DC020.asp
Copyright:     Copyright © 1999 Marlborough Stirling
Description:   Group Connection Maintenance screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:
Prog	Date		Description
JLD		15/11/1999	Created
JLD		10/12/1999	DC/013	Add scrolling to listbox
JLD		13/12/1999	DC/043	Deal with empty listbox on load.
JLD		14/12/1999	DC/044 - Listbox now refreshed after last line deleted
					DC/045 - Readonly processing corrected for empty listbox.
JLD		23/12/1999	SYS0116 - re-named screen 'Group Connections'
JLD		01/02/2000	Rework
AY		14/02/00	Change to msgButtons button types
IW		17/3/00		Route Cancel back to DC012
AY		30/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
MC		31/05/00	SYS0403 Removed unnecessary context fields from loop
CL		05/03/01	SYS1920 Read only functionality added
JLD		4/12/01		SYS2806 completeness checkk routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		22/08/2002	BMIDS00355	IE 5.5 Upgrade 
GD		17/11/2002		BMIDS00376 Disable screen if readonly
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/%>
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
<link href="stylesheet.css" rel="STYLESHEET" type="text/css"><title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<script src="validation.js" language="JScript"></script>
<!-- List Scroll object ( and table I believe) -->
<span id="spnListScroll">
	<span style="LEFT: 230px; POSITION: absolute; TOP: 248px">
		<object data="scTableListScroll.asp" id="scScrollTable" style="LEFT: 0px; TOP: 0px; height:24; width:304" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
	</span> 
</span>	
<!-- Specify Forms Here -->
<form id="frmToDC012" method="post" action="DC012.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC030" method="post" action="DC030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC025" method="post" action="DC025.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange">
	<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 218px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
		<span id="spnTable" style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE">
			<table id="tblTable" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
				<tr id="rowTitles">	<td width="30%" class="TableHead">Customer Name</td>	<td width="15%" class="TableHead">Account Type</td>	<td width="15%" class="TableHead">Account Number</td>	<td class="TableHead">Other Details</td></tr>
				<tr id="row01">		<td width="30%" class="TableTopLeft">&nbsp</td>		<td width="15%" class="TableTopCenter">&nbsp</td>		<td width="15%" class="TableTopCenter">&nbsp</td>		<td class="TableTopRight">&nbsp</td></tr>
				<tr id="row02">		<td width="30%" class="TableLeft">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
				<tr id="row03">		<td width="30%" class="TableLeft">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
				<tr id="row04">		<td width="30%" class="TableLeft">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
				<tr id="row05">		<td width="30%" class="TableLeft">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
				<tr id="row06">		<td width="30%" class="TableLeft">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
				<tr id="row07">		<td width="30%" class="TableLeft">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
				<tr id="row08">		<td width="30%" class="TableLeft">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
				<tr id="row09">		<td width="30%" class="TableLeft">&nbsp</td>				<td width="15%" class="TableCenter">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
				<tr id="row10">		<td width="30%" class="TableBottomLeft">&nbsp</td>		<td width="15%" class="TableBottomCenter">&nbsp</td>		<td width="15%" class="TableBottomCenter">&nbsp</td>	<td class="TableBottomRight">&nbsp</td></tr>
			</table>
		</span>
		<span style="TOP: 188px; LEFT: 4px; POSITION: ABSOLUTE">
			<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="TOP: 188px; LEFT: 68px; POSITION: ABSOLUTE">
			<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="TOP: 188px; LEFT: 132px; POSITION: ABSOLUTE">
			<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</div>
</form>
<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 290px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/DC020Attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction				= "";
var m_sReadOnly					= "";
var m_sCustomerNumber			= "";
var m_sCustomerVersionNumber	= "";
var m_sCustomerNumber1			= "";
var m_sCustomerVersionNumber1	= "";
var m_sCustomerNumber2			= "";
var m_sCustomerVersionNumber2	= "";
var m_sCustomerNumber3			= "";
var m_sCustomerVersionNumber3	= "";
var m_sCustomerNumber4			= "";
var m_sCustomerVersionNumber4	= "";
var m_sCustomerNumber5			= "";
var m_sCustomerVersionNumber5	= "";
var m_sCustomerName1			= "";
var m_sCustomerName2			= "";
var m_sCustomerName3			= "";
var m_sCustomerName4			= "";
var m_sCustomerName5			= "";
var m_iTableLength = 10;
var GroupConnXML				= null;
var scScreenFunctions;
var m_blnReadOnly = false;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Group Connections","DC020",scScreenFunctions);

	RetrieveContextData();
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC020");
	<% // GD BMIDS00376 START %>
	if (m_blnReadOnly == true)
	{
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnDelete.disabled = true;
	}
	<% // GD BMIDS00376 END %>	
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
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber",null);
	m_sCustomerNumber1 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber1","1112");
	m_sCustomerNumber2 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber2",null);
	m_sCustomerNumber3 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber3",null);
	m_sCustomerNumber4 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber4",null);
	m_sCustomerNumber5 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber5",null);
	m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber",null);
	m_sCustomerVersionNumber1 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1","1");
	m_sCustomerVersionNumber2 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber2",null);
	m_sCustomerVersionNumber3 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber3",null);
	m_sCustomerVersionNumber4 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber4",null);
	m_sCustomerVersionNumber5 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber5",null);
	m_sCustomerName1 = scScreenFunctions.GetContextParameter(window,"idCustomerName1","Jane");
	m_sCustomerName2 = scScreenFunctions.GetContextParameter(window,"idCustomerName2",null);
	m_sCustomerName3 = scScreenFunctions.GetContextParameter(window,"idCustomerName3",null);
	m_sCustomerName4 = scScreenFunctions.GetContextParameter(window,"idCustomerName4",null);
	m_sCustomerName5 = scScreenFunctions.GetContextParameter(window,"idCustomerName5",null);
}
function AddGroupConnectionToRequest(XML, tagActiveTag, sTagName, iCustomerNumber)
{
<%  /* Add each customer to the list. XML input is the XML to add to. */
%>	var sCustomerNumber, sCustomerVersionNumber;
	switch (iCustomerNumber)
	{
		case 1:		sCustomerNumber			= m_sCustomerNumber1;
					sCustomerVersionNumber	= m_sCustomerVersionNumber1;
					break;
		case 2:		sCustomerNumber			= m_sCustomerNumber2;
					sCustomerVersionNumber	= m_sCustomerVersionNumber2;
					break;
		case 3:		sCustomerNumber			= m_sCustomerNumber3;
					sCustomerVersionNumber	= m_sCustomerVersionNumber3;
					break;
		case 4:		sCustomerNumber			= m_sCustomerNumber4;
					sCustomerVersionNumber	= m_sCustomerVersionNumber4;
					break;
		case 5:		sCustomerNumber			= m_sCustomerNumber5;
					sCustomerVersionNumber	= m_sCustomerVersionNumber5;
					break;
		default:	sCustomerNumber			= "";
					sCustomerVersionNumber	= "";
					break;			
	}
	if(sCustomerNumber != "" && sCustomerVersionNumber != "")
	{						
		XML.ActiveTag = tagActiveTag;			
		XML.CreateActiveTag(sTagName);
		XML.CreateTag("CUSTOMERNUMBER",				sCustomerNumber);
		XML.CreateTag("CUSTOMERVERSIONNUMBER",		sCustomerVersionNumber);
	}
}
function ShowList(nStart)
{
<%  /* Populate the listbox with values from GroupConnXML Also used by the list scroll object */
%>  scScrollTable.clear();
	for (var iCount = 0; iCount < GroupConnXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		GroupConnXML.SelectTagListItem(iCount + nStart);
		var sAccountType = GroupConnXML.GetTagAttribute("ACCOUNTTYPE", "TEXT");
		var sAccountNumber = GroupConnXML.GetTagText("ACCOUNTNUMBER");
		var sAdditionalDetails = GroupConnXML.GetTagText("ADDITIONALDETAILS");
		var sCustomerName = scScreenFunctions.GetContextCustomerName(window, GroupConnXML.GetTagText("CUSTOMERNUMBER"));
		// Add to the search table
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sCustomerName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),sAccountType);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),sAccountNumber);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),sAdditionalDetails);
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}
}
function PopulateListBox(nStart)
{
<%  /* Populate the listbox with values from GroupConnXML */ 
%>	GroupConnXML.ActiveTag = null;
	GroupConnXML.CreateTagList("GROUPCONNECTION");
	var iNumberOfCustomers = GroupConnXML.ActiveTagList.length;
	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfCustomers);
	ShowList(nStart);
	if(iNumberOfCustomers == 0)
	{
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
	}
}
function PopulateScreen()
{
	GroupConnXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	GroupConnXML.CreateRequestTag(window, null);
	var tagGROUPCONNECTIONLIST = GroupConnXML.CreateActiveTag("GROUPCONNECTIONLIST");
	for(var nCount = 1; nCount <=5; nCount++)
	{
		AddGroupConnectionToRequest(GroupConnXML, tagGROUPCONNECTIONLIST, "GROUPCONNECTION", nCount);
	}
	GroupConnXML.RunASP(document, "FindGroupConnectionList.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = GroupConnXML.CheckResponse(ErrorTypes);
	if ( ErrorReturn[0] == true )
	{
		PopulateListBox(0);
		scScrollTable.setRowSelected(1);
		if(m_sReadOnly == "1")
		{
			frmScreen.btnAdd.disabled = true;
			frmScreen.btnDelete.disabled = true;
			frmScreen.btnEdit.focus();
		}
		else frmScreen.btnAdd.focus();
	}
	else
	{
		if(m_sReadOnly == "1")
		{
			frmScreen.btnDelete.disabled = true;
			frmScreen.btnEdit.disabled = true;
			frmScreen.btnAdd.disabled = true;
		}
		else
		{
			frmScreen.btnDelete.disabled = true;
			frmScreen.btnAdd.focus();
			frmScreen.btnEdit.disabled = true;
		}
	}
	ErrorTypes = null;
	ErrorReturn = null;
}
function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Add");
	frmToDC025.submit();
}
function frmScreen.btnDelete.onclick()
{
	var bAllowDelete = confirm("Are you sure?");
	if(bAllowDelete)
	{
<%		//Get the XML that just contains the GroupConnection chosen in the listbox
%>		var XML = GetXMLBlock(false);
		// 		XML.RunASP(document, "SaveGroupConnection.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "SaveGroupConnection.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		if (XML.IsResponseOK() == true)
		{
<%			//Rather than re-search the database, remove the relevant section from the GroupConnXML
%>			var nDeletedRow = DeleteGroupFromXML();
			PopulateListBox(0);
			if(nDeletedRow > GroupConnXML.ActiveTagList.length)	nDeletedRow = nDeletedRow - 1;
			if(nDeletedRow >=0)	scScrollTable.setRowSelected(nDeletedRow);
			else
			{
				frmScreen.btnDelete.disabled = true;
				frmScreen.btnEdit.disabled = true;
			}
		}
	}
	else frmScreen.btnDelete.focus();
}
function frmScreen.btnEdit.onclick()
{
<%	//Set idXML in context to the GROUPCONNECTION record for the selected list box line
%>	var XML = GetXMLBlock(true);
	scScreenFunctions.SetContextParameter(window,"idXML", XML.ActiveTag.xml);
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Edit");
	frmToDC025.submit();
}
function DeleteGroupFromXML()
{
<%  /* Deletes a GROUPCONNECTION block from the GroupConnXML
	   The GROUPCONNECTION block deleted is the one selected in the listbox 
	   Returns nRow - the number of the row deleted */
%>	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	GroupConnXML.ActiveTag = null;
	var tagNode = GroupConnXML.SelectTag(null, "GROUPCONNECTIONLIST");
	tagNode.removeChild(tagNode.childNodes.item(nRowSelected -1));
	return nRowSelected;
}
function GetXMLBlock(bForEdit)
{
<%  /* Description:	Returns an XML DOM for a single GROUPCONNECTION block
	   from GroupConnXML - the one relating to the selected	line in the listbox
	   Args Passed:	bForEdit - if true, we want all the data to pass to DC025.
					if false, we want only the primary key for deletion purposes */
	//From the row selected, find the index into the XML from which the table
	//was created. Create a new XML request block with just this block of XML
%>	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	GroupConnXML.ActiveTag = null;
	GroupConnXML.CreateTagList("GROUPCONNECTION");
	GroupConnXML.SelectTagListItem(nRowSelected -1);
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window, "SAVE");
	XML.CreateActiveTag("GROUPCONNECTION");
	XML.CreateTag("CUSTOMERNUMBER", GroupConnXML.GetTagText("CUSTOMERNUMBER"));
	XML.CreateTag("CUSTOMERVERSIONNUMBER", GroupConnXML.GetTagText("CUSTOMERVERSIONNUMBER"));
	XML.CreateTag("GROUPCONNECTIONSEQUENCENUMBER", GroupConnXML.GetTagText("GROUPCONNECTIONSEQUENCENUMBER"));
	if(bForEdit == true)
	{
		XML.CreateTag("ACCOUNTTYPE", GroupConnXML.GetTagText("ACCOUNTTYPE"));
		XML.CreateTag("ACCOUNTNUMBER", GroupConnXML.GetTagText("ACCOUNTNUMBER"));
		XML.CreateTag("ADDITIONALINDICATOR", GroupConnXML.GetTagText("ADDITIONALINDICATOR"));
		XML.CreateTag("ADDITIONALDETAILS", GroupConnXML.GetTagText("ADDITIONALDETAILS"));
	}
	return(XML);
}
function btnSubmit.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmToDC030.submit();
}
function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmToDC012.submit();
}
-->
</script>
</body>
</html>




