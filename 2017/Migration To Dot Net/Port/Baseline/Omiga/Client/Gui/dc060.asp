<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc060.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   This screen is responsible for the maintenance of
				the customer's address history
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DPol	15/11/1999	Creation
DPol	16/11/1999	Still in the process of creating it, but as the
					Business Object isn't ready, this is as far as
					I can go at the moment, so checked in for safety.
JLD		09/12/99	Finished off.
JLD		13/12/1999	DC/021 - added country to address line
					DC/028 - fixed.
JLD		16/12/1999	SYS0077 - removed idHasDependents. 
AD		31/01/2000	Rework
AY		14/02/00	Change to msgButtons button types
AY		30/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
MH      22/06/00    SYS0933 Readonly processing
CL		05/03/01	SYS1920 Read only functionality added
DJP		27/09/01	SYS2564/SYS2751 (child) Make Cancel and Next routing overrideable by
					a client
JLD		4/12/01		SYS2806 use scScreenfunctions CompletenessCheckRouting
DPF		21/06/02	BMIDS00077 - changes to file to bring in line with Core V7.0.2
					SYS4727 Use cached versions of frame functions
					SYS4767 MSMS to Core integration
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

prog	Date		AQR			Description
MV		22/08/2002	BMIDS00355	IE 5.5 Upgrade - Modified the Msgbuttons Position 
GD		17/11/2002	BMIDS00376 Disable screen if readonly
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog    Date			AQR		Description
MF		22/07/2005		MAR19	Added MSG omiga schema for intellisense
HM		05/10/2005		MAR19	Modifications for MAR19. Displaying of various controls
								now dependant on global paramater RestrictUseOfCurrentOrCorAddress
RCl		02/06/2006		MAR1847	Modify how the Delete button is disabled according to the global
								parameter RestrictUseOfCurrentOrCorAddress
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific history

Prog    Date			AQR		Description
PE		26/06/2006		EP234	Taken a copy of dc060.asp to get the MAR1847 change.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>
<% /* removed - DPF 21/6/02 - BMIDS00077 
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
BMIDS00077 END */ %>

<span id="spnAddressSummaryListScroll">
	<% /* span left setting altered from 250 -> 300 px - DPF 21/06/02 - BMIDS00077 */ %>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 242px">
<OBJECT data=scTableListScroll.asp id=scAddressSummaryTable 
style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 
type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>	

<%/* FORMS */%>
<form id="frmToDC065" method="post" action="dc065.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" validate ="onchange" mark>
<div id="divBackground" style="HEIGHT: 246px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span id="spnAddressSummary" style="LEFT: 4px; POSITION: absolute; TOP: 4px">
		<table id="tblAddressSummary" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="25%" class="TableHead">Customer Name&nbsp;</td>	<td width="13%" class="TableHead">Address Type&nbsp;</td>	<td width="36%" class="TableHead">Address&nbsp;</td>	<td width="13%" class="TableHead">Resident From&nbsp;</td>	<td class="TableHead">To&nbsp;</td></tr>
			<tr id="row01">		<td width="25%" class="TableTopLeft">	&nbsp;</td>		<td width="13%" class="TableTopCenter">&nbsp;</td>		<td width="36%" class="TableTopCenter">&nbsp;</td>		<td width="13%" class="TableTopCenter">&nbsp;</td>			<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="13%" class="TableCenter">&nbsp;</td>			<td width="36%" class="TableCenter">&nbsp;</td>			<td width="13%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="13%" class="TableCenter">&nbsp;</td>			<td width="36%" class="TableCenter">&nbsp;</td>			<td width="13%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="13%" class="TableCenter">&nbsp;</td>			<td width="36%" class="TableCenter">&nbsp;</td>			<td width="13%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="13%" class="TableCenter">&nbsp;</td>			<td width="36%" class="TableCenter">&nbsp;</td>			<td width="13%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="13%" class="TableCenter">&nbsp;</td>			<td width="36%" class="TableCenter">&nbsp;</td>			<td width="13%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="13%" class="TableCenter">&nbsp;</td>			<td width="36%" class="TableCenter">&nbsp;</td>			<td width="13%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="13%" class="TableCenter">&nbsp;</td>			<td width="36%" class="TableCenter">&nbsp;</td>			<td width="13%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="13%" class="TableCenter">&nbsp;</td>			<td width="36%" class="TableCenter">&nbsp;</td>			<td width="13%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="25%" class="TableBottomLeft">&nbsp;</td>		<td width="13%" class="TableBottomCenter">&nbsp;</td>		<td width="36%" class="TableBottomCenter">&nbsp;</td>	<td width="13%" class="TableBottomCenter">&nbsp;</td>		<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>									

	<span style="LEFT: 4px; POSITION: absolute; TOP: 182px">
		<input id="btnAddButton" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
	</span>

	<span style="LEFT: 68px; POSITION: absolute; TOP: 182px">
		<input id="btnEditButton" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
	</span>

	<span style="LEFT: 132px; POSITION: absolute; TOP: 182px">
		<input id="btnDeleteButton" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
	
	<!-- button added - DPF 21/06/02 - BMIDS00077 -->
	<!-- SG 27/02/02 SYS4186 -->
	<span style="LEFT: 196px; POSITION: absolute; TOP: 182px">
		<input id="btnCopyButton" value="Copy" type="button" style="WIDTH: 60px" class="msgButton" LANGUAGE=javascript onclick="return btnCopyButton_onclick()">
	</span>
	<!-- BMIDS00077 END -->

	<span style="LEFT: 4px; POSITION: absolute; TOP: 218px" class="msgLabel">
		Application Correspondence Salutation
		<span style="LEFT: 200px; POSITION: absolute; TOP: -3px">
			<input id="txtCorrespondenceSalutation" name="CorrespondenceSalutation" maxlength="50" style="POSITION: absolute; WIDTH: 200px" class="msgTxt">
		</span>
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 320px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="Routing/DC060Routing.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc060Attribs.asp" -->
<%/* CODE */%>
<script language="JScript">
<!--
var m_XMLAddressSummary;
var m_sMetaAction = null;
var m_iTableLength = 10;
var m_iCurrentMaximumNoOfCustomers = 5;		
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sReadOnly = null;	
var m_blnReadOnly = false;
<% /* WP01 MAR19  */ %>
var m_bIsAddressRestriction = false;
var xmlCombos ;
var m_sAddressDescrHome = "";
var m_sAddressDescrCorr = "";

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	//next line replaced by line below - DPF 21/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Address Summary","DC060",scScreenFunctions);

	<%/* Fetch the combo values and the respective validations for the follwing combos ;
	 later used to populate value name in the form */
	%> 
	xmlCombos = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("CustomerAddressType");
	xmlCombos.GetComboLists(document,sGroupList);

	RetrieveContextData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();

	// Default Add to enabled, Edit/Delete to disabled
	frmScreen.btnAddButton.disabled = false;
	frmScreen.btnEditButton.disabled = true; 
	frmScreen.btnDeleteButton.disabled = true; 
	
	//line added - DPF 21/06/02 - BMIDS00077
	//SG 27/02/02 SYS4186
	frmScreen.btnCopyButton.disabled = true;
	//BMIDS00077 END

	frmScreen.btnAddButton.focus();

	// if readonly indicator = true, disable add and delete button
	PopulateScreen();
	

<% // GD BMIDS00376 START %>
	//m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC060");
<% // GD BMIDS00376 END %>
	
	if (m_sReadOnly == "1")
	{
		frmScreen.btnAddButton.disabled = true;
		frmScreen.btnDeleteButton.disabled = true;
		
		//line added - DPF 21/06/02 - BMIDS00077
		//SG 27/02/02 SYS4186
		frmScreen.btnCopyButton.disabled = true;
		//BMIDS00077 END
		
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	//m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC060");
	
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
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
		
	<% /* HM MAR19 WP01 Check state of Global parameter RestrictUseOfCurrentOrCorAddress */ %>
	var ParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bRestrict = ParamXML.GetGlobalParameterBoolean(document,"RestrictUseOfCurrOrCorrAddress");	
	m_bIsAddressRestriction = bRestrict;
	ParamXML = null;
	
	m_sAddressDescrHome = xmlCombos.GetComboDescriptionForValidation("CustomerAddressType", "H")
	m_sAddressDescrCorr = xmlCombos.GetComboDescriptionForValidation("CustomerAddressType", "C")
}

function frmScreen.btnAddButton.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToDC065.submit();
}

function frmScreen.btnDeleteButton.onclick()
{
	var bAllowDelete = confirm("Are you sure?");

	if (bAllowDelete)
	{
		// Get the index of the selected row
		var nRowSelected = scAddressSummaryTable.getOffset() + scAddressSummaryTable.getRowSelected();

		m_XMLAddressSummary.SelectTagListItem(nRowSelected-1);

		//create a new XML for deletion
		//next line replaced with line below
		//var XML = new scXMLFunctions.XMLObject();
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window,null)
		XML.CreateActiveTag("CUSTOMERADDRESS");
		XML.AddXMLBlock(m_XMLAddressSummary.ActiveTag);

		// 		XML.RunASP(document,"DeleteCustomerAddress.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"DeleteCustomerAddress.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

		if(XML.IsResponseOK())
		{
			// Remove the entry from the XML and redisplay the list
			m_XMLAddressSummary.RemoveActiveTag();
			scAddressSummaryTable.RowDeleted();
				
			// If the last entry was deleted make sure the selection returns to the new last entry 
			if(m_XMLAddressSummary.ActiveTagList.length < nRowSelected)
				nRowSelected = nRowSelected - 1;
							
			// If there is still a line to select, select it; otherwise, disable the Edit/Delete buttons					
			if(nRowSelected > 0)
				scAddressSummaryTable.setRowSelected(nRowSelected - scAddressSummaryTable.getOffset());
			else
			{
				frmScreen.btnEditButton.disabled = true;
				frmScreen.btnDeleteButton.disabled = true;
				
				//line added - DPF 21/06/02 - BMIDS00077
				//SG 27/02/02 SYS4186
				frmScreen.btnCopyButton.disabled = true;				
				//BMIDS00077 END
				
				scScreenFunctions.SetFocusToFirstField(frmScreen);
			}
		}
	}
	else
		frmScreen.btnDeleteButton.focus();
}

function frmScreen.btnEditButton.onclick()
{
	// Get the index of the selected row
	var nRowSelected = scAddressSummaryTable.getOffset() + scAddressSummaryTable.getRowSelected();
	m_XMLAddressSummary.ActiveTag = null;
	m_XMLAddressSummary.CreateTagList("CUSTOMERADDRESS");
	if(m_XMLAddressSummary.SelectTagListItem(nRowSelected-1) == true)
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		scScreenFunctions.SetContextParameter(window,"idXML",m_XMLAddressSummary.ActiveTag.xml);
		frmToDC065.submit();
	}
}

function PopulateScreen()
{
	//next line replaced by line below - DPF 21/06/02 - BMIDS00077
	//m_XMLAddressSummary = new scXMLFunctions.XMLObject();
	m_XMLAddressSummary = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_XMLAddressSummary.CreateRequestTag(window,null)
	m_XMLAddressSummary.CreateActiveTag("APPLICATION");
	m_XMLAddressSummary.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	m_XMLAddressSummary.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);

	var tagLIST = m_XMLAddressSummary.CreateActiveTag("CUSTOMERADDRESSLIST");

	// Need to pass over the customers that only have any value
	// m_iCurrentMaximumNoOfCustomers is set at the top......HARD CODED
	for (var iLoop=1; iLoop<=m_iCurrentMaximumNoOfCustomers; iLoop++)
	{
		var sLoop = iLoop; 
		// get data from context,
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + sLoop,null);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + sLoop,null);

		// now check that there is some data to be sent in the request.....
		if ((sCustomerNumber != "") && (sCustomerVersionNumber != ""))
		{
			m_XMLAddressSummary.ActiveTag = tagLIST;
			m_XMLAddressSummary.CreateActiveTag("CUSTOMERADDRESS");
			m_XMLAddressSummary.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			m_XMLAddressSummary.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
		}
	}

	// Run server-side asp script
	m_XMLAddressSummary.RunASP(document,"FindCustomerAddressListAndSalutation.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = m_XMLAddressSummary.CheckResponse(ErrorTypes);

	if ( ErrorReturn[0] == true )
	{
		//Record was found
		m_XMLAddressSummary.SelectTag(null,"APPLICATION");
		frmScreen.txtCorrespondenceSalutation.value = m_XMLAddressSummary.GetTagText("CORRESPONDENCESALUTATION");

		PopulateTable();
	}

	ErrorTypes = null;
	ErrorReturn = null;
}

function PopulateTable()
{
	m_XMLAddressSummary.ActiveTag = null;
	m_XMLAddressSummary.CreateTagList("CUSTOMERADDRESS");
	var iNoOfCustomerAddresses = m_XMLAddressSummary.ActiveTagList.length;

	if (iNoOfCustomerAddresses > 0)
	{
		scAddressSummaryTable.initialiseTable(tblAddressSummary, 0, "", ShowTable, m_iTableLength, iNoOfCustomerAddresses);
		ShowTable(0);
		scAddressSummaryTable.setRowSelected(1);
		<% /* HM MAR19 WP01 Disable Delete button when AddressType = Current or Correspondence*/ %>
		//frmScreen.btnDeleteButton.disabled = false;
		spnAddressSummary.onclick();
		frmScreen.btnEditButton.disabled = false;
		
		//line added - DPF 21/06/02 - BMIDS00077
		//SG 27/02/02 SYS4186
		frmScreen.btnCopyButton.disabled = false;
		//BMIDS00077 END
	}
}

function ShowTable(iStart)
{			
	var sCustomerNumber, sCustomerVersionNumber, sCustomerName, sAddressType, sAddress, sResidentFrom, sResidentTo;
	var ParentTag = null;

	m_XMLAddressSummary.ActiveTag = null;
	m_XMLAddressSummary.CreateTagList("CUSTOMERADDRESS");
	for (var iLoop = 0; iLoop < m_XMLAddressSummary.ActiveTagList.length && iLoop < m_iTableLength; iLoop++)
	{				
		m_XMLAddressSummary.SelectTagListItem(iLoop + iStart);

		ParentTag = m_XMLAddressSummary.ActiveTag;

		sCustomerNumber = m_XMLAddressSummary.GetTagText("CUSTOMERNUMBER");
		sCustomerVersionNumber = m_XMLAddressSummary.GetTagText("CUSTOMERVERSIONNUMBER");
		sCustomerName = scScreenFunctions.GetContextCustomerName(window,sCustomerNumber);
		sAddressType = m_XMLAddressSummary.GetTagAttribute("ADDRESSTYPE","TEXT");
		sResidentFrom = m_XMLAddressSummary.GetTagText("DATEMOVEDIN");
		sResidentTo = m_XMLAddressSummary.GetTagText("DATEMOVEDOUT");

		m_XMLAddressSummary.SelectTag(ParentTag, "ADDRESS");
		// build up the address string with the details required
		sAddress = GetAddressString();

		ShowRow(iLoop+1, sCustomerName, sAddressType, sAddress, sResidentFrom, sResidentTo);
	}											
}

function GetAddressString()
{		
	function AddComma(sAddress, sTagValue, bAddComma)
	{																						
		if ((bAddComma==true) && (sTagValue != "") && (sAddress != "")) sTagValue = ", " + sTagValue;
		sAddress += sTagValue;
		return sAddress;
	}	

	var sAddress = "";
						
	sAddress = AddComma(sAddress, m_XMLAddressSummary.GetTagText("POSTCODE"), false);
	sAddress = AddComma(sAddress, m_XMLAddressSummary.GetTagText("FLATNUMBER"), true);
	sAddress = AddComma(sAddress, m_XMLAddressSummary.GetTagText("BUILDINGORHOUSENAME"), true);
	sAddress = AddComma(sAddress, m_XMLAddressSummary.GetTagText("BUILDINGORHOUSENUMBER"), true); <% /* APS UNIT TEST REF 21 */ %>
	sAddress = AddComma(sAddress, m_XMLAddressSummary.GetTagText("STREET"), true);
	sAddress = AddComma(sAddress, m_XMLAddressSummary.GetTagText("DISTRICT"), true);
	sAddress = AddComma(sAddress, m_XMLAddressSummary.GetTagText("TOWN"), true);
	sAddress = AddComma(sAddress, m_XMLAddressSummary.GetTagText("COUNTY"), true);
	sAddress = AddComma(sAddress, m_XMLAddressSummary.GetTagAttribute("COUNTRY","TEXT"), true);
						
	return sAddress;
}

function ShowRow(nIndex, sCustomerName, sAddressType, sAddress, sResidentFrom, sResidentTo)
{			 					
	scScreenFunctions.SizeTextToField(tblAddressSummary.rows(nIndex).cells(0),sCustomerName);
	scScreenFunctions.SizeTextToField(tblAddressSummary.rows(nIndex).cells(1),sAddressType);
	scScreenFunctions.SizeTextToField(tblAddressSummary.rows(nIndex).cells(2),sAddress);
	scScreenFunctions.SizeTextToField(tblAddressSummary.rows(nIndex).cells(3),sResidentFrom);
	scScreenFunctions.SizeTextToField(tblAddressSummary.rows(nIndex).cells(4),sResidentTo);
}


function btnSubmit.onclick()
{
	var bSuccess = false;
	if(frmScreen.onsubmit())
	{
		bSuccess = true;
		// if any changes have been made
		if (m_sReadOnly != "1" && IsChanged())
			bSuccess = CommitChanges();
	}

	if(bSuccess)
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else		
		{
			<% // DJP SYS2564/SYS2751 %>
			RouteNext();
		}
	}
}

function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else		
	{
		<% // DJP SYS2564/SYS2751 %>
		RoutePrevious();
	}
}

function CommitChanges()
{
	var bSuccess = true;

	//next line replaced with line below - DPF 21/06/02 - BMIDS00077
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null)
	XML.CreateActiveTag("APPLICATION");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("CORRESPONDENCESALUTATION", frmScreen.txtCorrespondenceSalutation.value);
	// 	XML.RunASP(document,"UpdateApplication.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"UpdateApplication.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}

//SG 29/05/02 SYS4767 Function Added
function frmScreen.btnCopyButton.onclick()	//SG 27/02/02 SYS4186
{
	// Get the index of the selected row
	var nRowSelected = scAddressSummaryTable.getOffset() + scAddressSummaryTable.getRowSelected();
	m_XMLAddressSummary.ActiveTag = null;
	m_XMLAddressSummary.CreateTagList("CUSTOMERADDRESS");
	if(m_XMLAddressSummary.SelectTagListItem(nRowSelected-1) == true)
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Copy");
		scScreenFunctions.SetContextParameter(window,"idXML",m_XMLAddressSummary.ActiveTag.xml);
		frmToDC065.submit();
	}
}

<% /* HM MAR19 WP01 Disable Delete button when AddressType = Current or Correspondence*/ %>
//function tblAddressSummary.onclick()
<% /* RCl MAR1847 Disable Delete button only when AddressRestriction is True*/ %>
function spnAddressSummary.onclick()
{
	var nRowSelected = scAddressSummaryTable.getOffset() + scAddressSummaryTable.getRowSelected();
	if (nRowSelected > 0 )
	{
		if (m_bIsAddressRestriction == false)
		{
		<% /* RCl 02/06/2006 MAR1847 Only disable when bIsAddressRestriction is true*/ %>
			frmScreen.btnDeleteButton.disabled=false;
		}
		else
		{
			m_XMLAddressSummary.ActiveTag = null;
			m_XMLAddressSummary.CreateTagList("CUSTOMERADDRESS");
			if (m_XMLAddressSummary.SelectTagListItem(nRowSelected-1)==true)
			{
				var sAddressType = m_XMLAddressSummary.GetTagAttribute("ADDRESSTYPE","TEXT");
				if (sAddressType==m_sAddressDescrHome ||  sAddressType==m_sAddressDescrCorr)
					frmScreen.btnDeleteButton.disabled=true;
				else
					frmScreen.btnDeleteButton.disabled=false;
			}
		}
	}
	else
		frmScreen.btnDeleteButton.disabled=true;
}

-->
</script>
</body>
</html>




