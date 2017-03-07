<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/*
Workfile:      DC033.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Alias/Association Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		24/11/99	Created
JLD		23/12/1999	SYS0099 - Removed 4th alias section
AD		31/01/2000	Rework
AY		14/02/00	Change to msgButtons button types
IW		08/03/00	SYS0097 (OK now calls cancel if ReadOnly screen)
IW		10/03/00	SYS0109 Logic of screen navigation ammended
IW		23/03/00	SYS0109 Defaulted Focus on validation failure
AY		30/03/00	New top menu/scScreenFunctions change
IW		27/04/00	SYS0207 Clear buttons added
BG		17/05/00	SYS0752 Removed Tooltips
MH      22/06/00    SYS0933 Readonly processing
CL		05/03/01	SYS1920 Read only functionality added
SA		04/06/01	SYS1062 Disable date and method of change if alias = "association"
SA		08/08/01	SYS1062 Further fixes around disabling/enabling controls.
DJP		26/09/01	SYS2564/SYS2743 (child) Client cosmetic customisation
DJP		01/02/02	SYS2564/SYS2743 (child) Client cosmetic customisation, fixed id problem.
DPF		20/06/02	BMIDS00077 - Changes made to bring this copy of file in line with Core V7.0.2
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		16/05/2002	BMIDS00008	Modified AliasOK() Removed Validation for MethodOfChange and AliasType
MV		17/05/2002	BMIDS00008	Modified AliasOK() - Included AliasType Validation
MV		22/08/2002	BMIDS00355	IE 5.5 Upgrade 
MDC		11/11/2002	BMIDS00911	Redesign screen in standard Omiga form.
GD		17/11/2002	BMIDS00376 Disable screen if readonly
DRC     11/02/2004  BMIDS693    Added Credit Search
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
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
<% /* BMIDS00911 MDC 11/11/2002 */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>

<%/* SCRIPTLETS */%>
<!-- List Scroll object -->
<span style="TOP: 275px; LEFT: 310px; POSITION: absolute">
	<object data="scListScroll.htm" id="scScrollPlus" 
	style="LEFT: 0px; TOP: 0px; height:24; width:304" 
	type=text/x-scriptlet VIEWASTEXT tabindex="-1"></object>
</span> 
<% /* BMIDS00911 MDC 11/11/2002 - End */ %>

<script src="validation.js" language="JScript"></script>
<% /*the following has been removed as per V7.0.2 of core - DPF 20/06/02 - BMIDS00077
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>

FORMS */%>
<form id="frmToDC030" method="post" action="DC030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC034" method="post" action="DC034.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange" year4>
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 250px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<% /*<span style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<font face="MS Sans Serif" size="1">
			<strong>Alias Association Details</strong>
		</font>
		</span>
	*/ %>

	<span style="TOP: 8px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		Customer Name
		<span style="TOP: -3px; LEFT: 90px; POSITION: ABSOLUTE">
			<input id="txtCustomerName" name="CustomerName" maxlength="70" style="WIDTH: 190px; POSITION: ABSOLUTE" class="msgTxt">
		</span>
	</span>
	
	<span id="spnExistingAliases" style="TOP: 40px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblExistingAliases" width="596px" border="0" cellspacing="0" cellpadding="0" class="msgTable">				
			<tr id="rowTitles"><td width="40%" class="TableHead">Name</td><td width="15%" class="TableHead">Alias Type</td><td width="15%" class="TableHead">Date of Change</td><td width="15%" class="TableHead">Method of Change</td><td width="15%" class="TableHead">Credit Search</td></tr>				
			<tr id="row01"><td width="40%" class="TableTopLeft">&nbsp</td><td width="15%" class="TableTopCenter">&nbsp</td><td width="15%" class="TableTopCenter">&nbsp</td><td width="15%" class="TableTopCenter">&nbsp</td><td width="15%" class="TableTopRight">&nbsp</td></tr>
			<tr id="row02"><td width="40%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableRight">&nbsp</td></tr>
			<tr id="row03"><td width="40%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableRight">&nbsp</td></tr>
			<tr id="row04"><td width="40%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableRight">&nbsp</td></tr>
			<tr id="row05"><td width="40%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableRight">&nbsp</td></tr>
			<tr id="row06"><td width="40%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableRight">&nbsp</td></tr>
			<tr id="row07"><td width="40%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableRight">&nbsp</td></tr>
			<tr id="row08"><td width="40%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableRight">&nbsp</td></tr>
			<tr id="row09"><td width="40%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableRight">&nbsp</td></tr>
			<tr id="row10"><td width="40%" class="TableBottomLeft">&nbsp</td><td width="15%" class="TableBottomCenter">&nbsp</td><td width="15%" class="TableBottomCenter">&nbsp</td><td width="15%" class="TableBottomCenter">&nbsp</td><td width="15%" class="TableBottomRight">&nbsp</td></tr>
		</table>
	</span>
	<span id="spnButtons" style="TOP: 215px; LEFT: 4px; POSITION: ABSOLUTE">
		<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
			<input id="btnAdd" value="Add" type="button" 
				style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="TOP: 0px; LEFT: 64px; POSITION: ABSOLUTE">
			<input id="btnEdit" value="Edit" type="button" 
				style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="TOP: 0px; LEFT: 128px; POSITION: ABSOLUTE">
			<input id="btnDelete" value="Delete" type="button" 
				style="WIDTH: 60px" class="msgButton">
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

<!--  #include FILE="attribs/DC033attribs.asp" -->
<!--  #include FILE="Customise/DC033Customise.asp" -->
<%/* CODE */%>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly	= "";
var m_sCurrentCustomerNumber = "";
var m_sCurrentCustomerVersion = "";
var scScreenFunctions;
var m_sReadOnly = null;
var m_blnReadOnly = false;
var XML = null;
var m_iTableLength = 10;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();

	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	//next line replaced by line below as per V7.0.2 of Core - DPF 20/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Alias/Association Summary","DC033",scScreenFunctions);

	scTable.initialise(tblExistingAliases, 0, "");

	RetrieveContextData();

	//This function is contained in the field attributes file (remove if not required)
	SetMasks();

	Validation_Init();
	Initialise();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC033");
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
	m_sReadOnly	= scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sCurrentCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","1007");
	m_sCurrentCustomerVersion = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","1");
}

function spnExistingAliases.onclick()
{
	if (scTable.getRowSelectedId() != null) 
	{
		frmScreen.btnEdit.disabled = false;
		<% // GD BMIDS00376 START %>
		if (m_blnReadOnly == false)
		{
			frmScreen.btnDelete.disabled = false;			
		}		
		<% // GD BMIDS00376 END %>	
	}
}

function spnExistingAliases.ondblclick()
{
	if (!frmScreen.btnEdit.disabled )
		frmScreen.btnEdit.onclick();
}


function PopulateScreen()
{
	//set the Customer Name
	var sCustomerName = scScreenFunctions.GetContextCustomerName(window, m_sCurrentCustomerNumber);
	frmScreen.txtCustomerName.value = sCustomerName;
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtCustomerName");

	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null)
	XML.CreateActiveTag("ALIAS");
	XML.CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
	XML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersion);

	XML.RunASP(document,"FindAliasList.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);

	if(ErrorReturn[0] == true && (ErrorReturn[1] != ErrorTypes[0]))
	{
		// records found, populate the screen fields
		XML.ActiveTag = null;
		XML.CreateTagList("ALIASPERSON");

		PopulateTable(0);
					
//		XML.ActiveTag = null;
//		XML.CreateTagList("ALIASPERSON");
		scScrollPlus.Initialise(PopulateTable,10,XML.ActiveTagList.length);
	}
	else
	{
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
	}
	
}

<% /* BMIDS00911 MDC 11/11/2002 */ %>
function PopulateTable(nStart)
{
	scTable.clear();

	XML.ActiveTag = null;
	XML.CreateTagList("ALIASPERSON");
	var iNoOfAliases = XML.ActiveTagList.length;

	if (iNoOfAliases > 0)
	{
		ShowTable(nStart);
		scTable.setRowSelected(1);
		frmScreen.btnDelete.disabled = false;
		frmScreen.btnEdit.disabled = false;
	}
	else
	{
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
	}
}

function ShowTable(iStart)
{			
	var sAliasName, sAliasType, sDateOfChange, sMethodOfChange, sSequenceNumber, sCreditSearch;
	var ParentTag = null;

	XML.ActiveTag = null;
	XML.CreateTagList("ALIASPERSON");
	for (var iLoop = 0; iLoop < XML.ActiveTagList.length && iLoop < m_iTableLength; iLoop++)
	{				
		XML.SelectTagListItem(iLoop + iStart);

		ParentTag = XML.ActiveTag;

		sSequenceNumber = XML.GetTagText("ALIASSEQUENCENUMBER");
		sAliasType = XML.GetTagAttribute("ALIASTYPE","TEXT");
		sDateOfChange = XML.GetTagText("DATEOFCHANGE");
		sMethodOfChange = XML.GetTagAttribute("METHODOFCHANGE","TEXT");
	    <% /* DRC BMIDS693  11/02/2004 START */ %>
		sCreditSearch = XML.GetTagText("CREDITSEARCH");
		if (sCreditSearch == "1")
		  sCreditSearch = "Yes";
		else
		  sCreditSearch = "";  
	    <% /* DRC BMIDS693  11/02/2004 END */ %>
		var tagPerson = XML.SelectTag(ParentTag, "PERSON");
		if(tagPerson != null)
			sAliasName = GetAliasName();
		else
			sAliasName = "Unspecified";

		scScreenFunctions.SizeTextToField(tblExistingAliases.rows(iLoop+1).cells(0),sAliasName);
		scScreenFunctions.SizeTextToField(tblExistingAliases.rows(iLoop+1).cells(1),sAliasType);
		scScreenFunctions.SizeTextToField(tblExistingAliases.rows(iLoop+1).cells(2),sDateOfChange);
		scScreenFunctions.SizeTextToField(tblExistingAliases.rows(iLoop+1).cells(3),sMethodOfChange);
		scScreenFunctions.SizeTextToField(tblExistingAliases.rows(iLoop+1).cells(4),sCreditSearch);
		
		
		tblExistingAliases.rows(iLoop+1).setAttribute("SequenceNumber", sSequenceNumber);
	}						
}

function GetAliasName()
{
	var sAliasName = "";
	var	xmlCombos = new top.frames[1].document.all.scXMLFunctions.XMLObject() ;
	
	if(xmlCombos.IsInComboValidationList(document,"Title", XML.GetTagText("TITLE"), ["O"]))
		sAliasName = XML.GetTagText("TITLEOTHER");
	else
		sAliasName = XML.GetTagAttribute("TITLE","TEXT");
		
	sAliasName += " ";
	sAliasName += XML.GetTagText("FIRSTFORENAME");
	sAliasName += " ";
	sAliasName += XML.GetTagText("SURNAME");

	return sAliasName;
}
<% /* BMIDS00911 MDC 11/11/2002 - End */ %>

function Initialise()
{
	PopulateScreen();

	if(m_sReadOnly == "1")
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		
}


function btnSubmit.onclick()
{
	//set metaAction
	scScreenFunctions.SetContextParameter(window,"idMetaAction","FromDC033");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber",m_sCurrentCustomerNumber);
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber",m_sCurrentCustomerVersion);

	frmToDC030.submit();
}

function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idXML","");
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToDC034.submit();
}

function frmScreen.btnEdit.onclick()
{
	var iSelectedRowNo = scTable.getRowSelected();
	var sSequenceNumber = tblExistingAliases.rows(iSelectedRowNo).getAttribute("SequenceNumber");
	
	XML.SelectSingleNode("//ALIASPERSON/ALIAS[ALIASSEQUENCENUMBER='" + sSequenceNumber + "']");
	scScreenFunctions.SetContextParameter(window,"idXML",XML.ActiveTag.xml);
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
	frmToDC034.submit();
}

function frmScreen.btnDelete.onclick()
{
	var bAllowDelete = confirm("Are you sure?");
				
	if (bAllowDelete == true)
	{
		var iSelectedRowNo = scTable.getRowSelected();
		var sSequenceNumber = tblExistingAliases.rows(iSelectedRowNo).getAttribute("SequenceNumber");
		XML.SelectSingleNode("//ALIASPERSON[ALIAS/ALIASSEQUENCENUMBER='" + sSequenceNumber + "']");

		var DeleteXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		DeleteXML.CreateRequestTag(window,null)
		DeleteXML.CreateActiveTag("ALIAS");
		DeleteXML.CreateTag("CUSTOMERNUMBER", m_sCurrentCustomerNumber);
		DeleteXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCurrentCustomerVersion);
		DeleteXML.CreateTag("ALIASSEQUENCENUMBER", sSequenceNumber);

		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				DeleteXML.RunASP(document,"DeleteAlias.asp");
				break;
			default: // Error
				DeleteXML.SetErrorResponse();
		}

		<% /* If the deletion is successful remove the entry from the 
			list xml and the screen and redisplay the list */ %>
		if(DeleteXML.IsResponseOK() == true)
		{					
			XML.RemoveActiveTag();
			scScrollPlus.RowDeleted();
		}
	}
}

-->
</script>
</body>
</html>


