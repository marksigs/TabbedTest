<%@ LANGUAGE="JSCRIPT" %>
<html>
<comment>
Workfile:      DC040.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Group Connection Maintenance screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		23/11/1999	Created
JLD		10/12/1999	DC/013 - make table scrollable
					DC/022 - Correct routing from OK
JLD		14/12/1999	DC/044 - checked against dc040 and changes made to make sure deletion of last
							 record works ok.
JLD		01/02/2000	rework for performance
AY		14/02/00	Change to msgButtons button types
IW		14/03/00	SYS0070 Tidied up dependant name in list
AY		30/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
CL		05/03/01	SYS1920 Read only functionality added
JLD		4/12/01		SYS2806 completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
</comment>
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
<form id="frmToDC060" method="post" action="DC060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC030" method="post" action="DC030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC045" method="post" action="DC045.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="HEIGHT: 218px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 4px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="25%" class="TableHead">Customer Name</td><td width="25%" class="TableHead">Dependant Name</td><td width="13%" class="TableHead">Date of Birth</td><td width="10%" class="TableHead">Other Resident?</td><td width="27%" class="TableHead">Other Details</td></tr>
			<tr id="row01">		<td width="25%" class="TableTopLeft">&nbsp</td>		<td width="25%" class="TableTopCenter">&nbsp</td>		<td width="13%" class="TableTopCenter">&nbsp</td>	<td width="10%" class="TableTopCenter">&nbsp</td>	<td class="TableTopRight">&nbsp</td></tr>
			<tr id="row02">		<td width="25%" class="TableLeft">&nbsp</td>				<td width="25%" class="TableCenter">&nbsp</td>			<td width="13%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>	    <td class="TableRight">&nbsp</td></tr>
			<tr id="row03">		<td width="25%" class="TableLeft">&nbsp</td>				<td width="25%" class="TableCenter">&nbsp</td>			<td width="13%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>	    <td class="TableRight">&nbsp</td></tr>
			<tr id="row04">		<td width="25%" class="TableLeft">&nbsp</td>				<td width="25%" class="TableCenter">&nbsp</td>			<td width="13%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>	    <td class="TableRight">&nbsp</td></tr>
			<tr id="row05">		<td width="25%" class="TableLeft">&nbsp</td>				<td width="25%" class="TableCenter">&nbsp</td>			<td width="13%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>	    <td class="TableRight">&nbsp</td></tr>
			<tr id="row06">		<td width="25%" class="TableLeft">&nbsp</td>				<td width="25%" class="TableCenter">&nbsp</td>			<td width="13%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>	    <td class="TableRight">&nbsp</td></tr>
			<tr id="row07">		<td width="25%" class="TableLeft">&nbsp</td>				<td width="25%" class="TableCenter">&nbsp</td>			<td width="13%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>	    <td class="TableRight">&nbsp</td></tr>
			<tr id="row08">		<td width="25%" class="TableLeft">&nbsp</td>				<td width="25%" class="TableCenter">&nbsp</td>			<td width="13%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>	    <td class="TableRight">&nbsp</td></tr>
			<tr id="row09">		<td width="25%" class="TableLeft">&nbsp</td>				<td width="25%" class="TableCenter">&nbsp</td>			<td width="13%" class="TableCenter">&nbsp</td>		<td width="10%" class="TableCenter">&nbsp</td>	    <td class="TableRight">&nbsp</td></tr>
			<tr id="row10">		<td width="25%" class="TableBottomLeft">&nbsp</td>		<td width="25%" class="TableBottomCenter">&nbsp</td>		<td width="13%" class="TableBottomCenter">&nbsp</td>	<td width="10%" class="TableBottomCenter">&nbsp</td><td class="TableBottomRight">&nbsp</td></tr>
		</table>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 188px">
		<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnAdd.onClick()"> 
	</span>
	<span style="LEFT: 68px; POSITION: absolute; TOP: 188px">
		<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnEdit.onClick()"> 
	</span>
	<span style="LEFT: 132px; POSITION: absolute; TOP: 188px">
		<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnDelete.onClick()"> 
	</span> 
</div> 
</form>
<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" -->
</div> 
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc040Attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction				= "";
var m_sReadOnly					= "";
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
var m_sApplicationNumber		= "";
var m_sApplicationFactFindNumber= "";
var DependantXML				= null;
var m_iTableLength = 10;
var scScreenFunctions;
var m_blnReadOnly = false;


function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Add");
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber", m_sApplicationNumber);
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber", m_sApplicationFactFindNumber);
	frmToDC045.submit();
}
function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmToDC030.submit();
}
function frmScreen.btnDelete.onclick()
{
	var bAllowDelete = confirm("Are you sure?");
	if(bAllowDelete)
	{
<%		//Get the XML that just contains the dependant chosen
		//in the listbox
%>		var XML = GetXMLBlock(false);
		// 		XML.RunASP(document, "DeleteDependantForCustomer.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "DeleteDependantForCustomer.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		if (XML.IsResponseOK() == true)
		{
<%			//Rather than re-search the database, remove the relevant section from the DependantXML
%>			var nDeletedRow = DeleteDependantFromXML();
			PopulateListBox(0);
			if(nDeletedRow > DependantXML.ActiveTagList.length)nDeletedRow = nDeletedRow - 1;
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
<%	//Set idXML in context to the DEPENDANT record for the selected list box line
%>	var XML = GetXMLBlock(true);
	XML.SelectTag(null,"DEPENDANT");
	scScreenFunctions.SetContextParameter(window,"idXML", XML.ActiveTag.xml);
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Edit");
	frmToDC045.submit();
}
function btnSubmit.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmToDC060.submit();
}
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

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
	FW030SetTitles("Dependant Summary","DC040",scScreenFunctions);

	RetrieveContextData();
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC040");
	
	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}
function AddDependantToRequest(XML, tagActiveTag, sTagName, iCustomerNumber)
{
<%  /*Description:	Add each customer to the list
	  Args Passed:	XML to add to
					Active tag
					tagname
					customer number */
%>	var sCustomerNumber = "";
	var sCustomerVersionNumber = "";
	switch (iCustomerNumber)
	{
		case 1:	sCustomerNumber			= m_sCustomerNumber1;
				sCustomerVersionNumber	= m_sCustomerVersionNumber1;
				break;
		case 2:	sCustomerNumber			= m_sCustomerNumber2;
				sCustomerVersionNumber	= m_sCustomerVersionNumber2;
				break;
		case 3:	sCustomerNumber			= m_sCustomerNumber3;
				sCustomerVersionNumber	= m_sCustomerVersionNumber3;
				break;
		case 4:	sCustomerNumber			= m_sCustomerNumber4;
				sCustomerVersionNumber	= m_sCustomerVersionNumber4;
				break;
		case 5:	sCustomerNumber			= m_sCustomerNumber5;
				sCustomerVersionNumber	= m_sCustomerVersionNumber5;
				break;
		default: sCustomerNumber			= "";
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
function DeleteDependantFromXML()
{
<%  /* Description:	Deletes a GROUPCONNECTION block from the DependantXML
					The GROUPCONNECTION block deleted is the one selected
					in the listbox */
%>	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	DependantXML.ActiveTag = null;
	var tagNode = DependantXML.SelectTag(null, "DEPENDANTLIST");
	tagNode.removeChild(tagNode.childNodes.item(nRowSelected -1));
	return(nRowSelected);
}
function GetXMLBlock(bForEdit)
{
<% /* Description:	Returns an XML DOM for a single DEPENDANT block
					from DependantXML - the one relating to the selected
					line in the listbox
      Args Passed:	bForEdit - if true, we want all the data to pass to DC045.
					if false, we want only the primary key for deletion purposes
	From the row selected, find the index into the XML from which the table
	was created. Create a new XML request block with just this block of XML */
%>	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	DependantXML.ActiveTag = null;			
	DependantXML.CreateTagList("DEPENDANT");
	DependantXML.SelectTagListItem(nRowSelected -1);
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
<%  /* Set the action tag to 'delete' here, even though the function is called in 2 places. Only 
       the delete function is run with this xml, the 'save' function is run from DC045 which creates its own request block */
%>	XML.CreateRequestTag(window, "DELETE");
	XML.CreateActiveTag("DEPENDANT");
	XML.CreateTag("CUSTOMERNUMBER",				DependantXML.GetTagText("CUSTOMERNUMBER"));
	XML.CreateTag("CUSTOMERVERSIONNUMBER",		DependantXML.GetTagText("CUSTOMERVERSIONNUMBER"));
	XML.CreateTag("DEPENDANTSEQUENCENUMBER",	DependantXML.GetTagText("DEPENDANTSEQUENCENUMBER"));
	XML.CreateTag("APPLICATIONNUMBER",			m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",	m_sApplicationFactFindNumber);
	if(bForEdit == true)
	{
		XML.CreateTag("PERSONGUID",				DependantXML.GetTagText("PERSONGUID"));
		XML.CreateTag("DEPENDANTRELATIONSHIP",	DependantXML.GetTagText("DEPENDANTRELATIONSHIP"));
		XML.CreateTag("ADDITIONALINDICATOR",	DependantXML.GetTagText("ADDITIONALINDICATOR"));
		XML.CreateTag("ADDITIONALDETAILS",		DependantXML.GetTagText("ADDITIONALDETAILS"));
		XML.CreateTag("DATEOFBIRTH",			DependantXML.GetTagText("DATEOFBIRTH"));
		XML.CreateTag("GENDER",					DependantXML.GetTagText("GENDER"));
		XML.CreateTag("MARITALSTATUS",			DependantXML.GetTagText("MARITALSTATUS"));
		XML.CreateTag("NATIONALITY",			DependantXML.GetTagText("NATIONALITY"));
		XML.CreateTag("OTHERFORENAMES",			DependantXML.GetTagText("OTHERFORENAMES"));
		XML.CreateTag("SECONDFORENAME",			DependantXML.GetTagText("SECONDFORENAME"));
		XML.CreateTag("SURNAME",				DependantXML.GetTagText("SURNAME"));
		XML.CreateTag("FIRSTFORENAME",			DependantXML.GetTagText("FIRSTFORENAME"));
		XML.CreateTag("TITLE",					DependantXML.GetTagText("TITLE"));
		XML.CreateTag("TITLEOTHER",				DependantXML.GetTagText("TITLEOTHER"));
		XML.CreateActiveTag("OTHERRESIDENT");
		XML.CreateTag("OTHERRESIDENTSEQUENCENUMBER", DependantXML.GetTagText("OTHERRESIDENTSEQUENCENUMBER"));
	}
	return(XML);
}
function PopulateListBox(nStart)
{
<%  /* Populate the listbox with values from DependantXML */
%>	DependantXML.ActiveTag = null;			
	DependantXML.CreateTagList("DEPENDANT");
	var iNumberOfCustomers = DependantXML.ActiveTagList.length;
	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfCustomers);
	ShowList(nStart);
	if(iNumberOfCustomers > 0)
	{
		scScrollTable.setRowSelected(1);
		frmScreen.btnDelete.disabled = false;
		frmScreen.btnEdit.disabled = false;
	}
	else
	{
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
	}
}
function ShowList(nStart)
{
	scScrollTable.clear();
	for (var iCount = 0; iCount < DependantXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		DependantXML.SelectTagListItem(iCount + nStart);
		var sCustomerName = scScreenFunctions.GetContextCustomerName(window, DependantXML.GetTagText("CUSTOMERNUMBER"));
		var sSecondForeName	   = (DependantXML.GetTagText("SECONDFORENAME") == "") ? "" : DependantXML.GetTagText("SECONDFORENAME") + " ";
		var sOtherForeName	   = (DependantXML.GetTagText("OTHERFORENAMES") == "") ? "" : DependantXML.GetTagText("OTHERFORENAMES") + " ";
		var sDependantName     = DependantXML.GetTagText("FIRSTFORENAME")  + ' ' + 	sSecondForeName + sOtherForeName + DependantXML.GetTagText("SURNAME");
		var sDateOfBirth	   = DependantXML.GetTagText("DATEOFBIRTH");
		var sIsOtherResident   = (DependantXML.GetTagText("OTHERRESIDENTSEQUENCENUMBER") == "") ? "No" : "Yes";				
		var sAdditionalDetails = DependantXML.GetTagText("ADDITIONALDETAILS");
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sCustomerName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),sDependantName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),sDateOfBirth);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3),sIsOtherResident);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(4),sAdditionalDetails);
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}
}
function PopulateScreen()
{
	DependantXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	DependantXML.CreateRequestTag(window, null);
	var tagDependantList = DependantXML.CreateActiveTag("DEPENDANTLIST");
	for(var nCount = 1; nCount <=5; nCount++)
	{
		AddDependantToRequest(DependantXML, tagDependantList, "DEPENDANT", nCount);
	}
	DependantXML.ActiveTag = tagDependantList
	DependantXML.CreateTag("APPLICATIONNUMBER",		    m_sApplicationNumber);
	DependantXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	DependantXML.RunASP(document, "FindDependantList.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = DependantXML.CheckResponse(ErrorTypes);
	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[1] == ErrorTypes[0])
		{
<%			//record not found
%>			if(m_sReadOnly == "1")frmScreen.btnAdd.disabled = true;
			else frmScreen.btnAdd.disabled = false;
			frmScreen.btnDelete.disabled = true;
			frmScreen.btnEdit.disabled = true;
		}
		else
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
	}
	ErrorTypes = null;
	ErrorReturn = null;
}
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sCustomerNumber1 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber1","1074");
	m_sCustomerNumber2 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber2",null);
	m_sCustomerNumber3 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber3",null);
	m_sCustomerNumber4 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber4",null);
	m_sCustomerNumber5 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber5",null);
	m_sCustomerVersionNumber1 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber1","1");
	m_sCustomerVersionNumber2 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber2",null);
	m_sCustomerVersionNumber3 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber3",null);
	m_sCustomerVersionNumber4 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber4",null);
	m_sCustomerVersionNumber5 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber5",null);
	m_sCustomerName1 = scScreenFunctions.GetContextParameter(window,"idCustomerName1","Cust");
	m_sCustomerName2 = scScreenFunctions.GetContextParameter(window,"idCustomerName2",null);
	m_sCustomerName3 = scScreenFunctions.GetContextParameter(window,"idCustomerName3",null);
	m_sCustomerName4 = scScreenFunctions.GetContextParameter(window,"idCustomerName4",null);
	m_sCustomerName5 = scScreenFunctions.GetContextParameter(window,"idCustomerName5",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B00003743");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");  
}
-->
</script>
</body>
</html>




