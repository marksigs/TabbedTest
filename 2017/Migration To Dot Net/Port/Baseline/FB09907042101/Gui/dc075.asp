<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      DC075.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Mortgage Account Special Features
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
PSC		05/08/2002  BMIDS00006 Created
PSC		23/08/2002	BMIDS00359 Amend screen title
GHun	05/09/2002	BMIDS00406 Remove the word "Mortgage" from the screen
GD		18/11/2002	BMIDS00376 Ensure readonly happens
MC		20/04/2004	BMIDS517	DC076 Popup dialog height change
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

<span id="spnListScroll">
	<span style="LEFT: 300px; POSITION: absolute; TOP: 320px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<% /* Specify Forms Here. GN300 is needed for MAIN screens to route back to completeness check screen.
      A MAIN screen is not a pop-up or a screen routed to via any button other than the main submit/cancel buttons */ %>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC071" method="post" action="DC071.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here - remove year4 attribute if no date fields are on this screen */ %>
<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 340px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP:40px; LEFT:10px; POSITION:ABSOLUTE" class="msgLabel">
		<% /* BMIDS00406 remove mortgage		
		// Mortgage Account
		// <span style="TOP:-3px; LEFT:100px; POSITION:ABSOLUTE">
		*/ %>
		Account
		<span style="TOP:-3px; LEFT:60px; POSITION:ABSOLUTE">
		<% /* BMIDS00406 End */ %>
			<input id="txtMortgageAccount" maxlength="10" style="WIDTH:120px" class="msgTxt">
		</span>
	</span>
	<span id=lblFeatures style="LEFT: 10px; POSITION: absolute; TOP: 80px" class="msgLabel">
		Special Features
		<span id="spnFeatures" style="TOP: 20px; LEFT: 0px; POSITION: ABSOLUTE">
			<table id="tblFeatures" width="585px" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">
				<td width="80%" class="TableHead">Description</td>	
				<td width="20%" class="TableHead">Indicator</td>
			</tr>
			<tr id="row01">		
				<td class="TableTopLeft">&nbsp;</td>		
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">		
				<td class="TableLeft">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">		
				<td class="TableLeft">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">		
				<td class="TableLeft">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row05">		
				<td class="TableLeft">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row06">		
				<td class="TableLeft">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row07">		
				<td class="TableLeft">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row08">		
				<td class="TableLeft">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row09">		
				<td class="TableLeft">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row10">		
				<td class="TableBottomLeft">&nbsp;</td>		
				<td class="TableBottomRight">&nbsp;</td>
			</tr>
			</table>
		</span>
	</span>
	<span style="TOP:260px; LEFT:10px; POSITION:ABSOLUTE">
		<input id="btnAdd" value="Add" type="button" style="WIDTH:60px" class="msgButton">
	</span>
	<span style="TOP:260px; LEFT:80px; POSITION:ABSOLUTE">
		<input id="btnEdit" value="Edit" type="button" style="WIDTH:60px" class="msgButton">
	</span>
	<span style="TOP:260px; LEFT:150px; POSITION:ABSOLUTE">
		<input id="btnDelete" value="Delete" type="button" style="WIDTH:60px" class="msgButton">
	</span>

</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/dc075attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript">
<!--
var m_sMetaAction = null;
var m_blnReadOnly = false;
var m_sReadOnly = "";
var scScreenFunctions;
var m_XMLFeatures = null;
var m_sAccountGuid = null;
var m_iTableLength = 10;
var m_bIsEdit = false;
var m_nNumberOfFeatures = 0;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sScreenId = "DC075";
	
<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
%>	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	<% /* PSC 23/08/2002 BMIDS00359 */ %>
	FW030SetTitles("Account Related Special Features",sScreenId,scScreenFunctions);

	m_bIsEdit = false;

	RetrieveContextData();
	SetMasks();
	Validation_Init();
		
	m_bReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, sScreenId);
	//GD BMIDS00376
	if (m_bReadOnly == true) m_sReadOnly = "1";
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
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
}


function PopulateScreen()
{
	var sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	var XML	= new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(sXML);
	
	XML.SelectTag(null, "MORTGAGEACCOUNT");
	m_sAccountGuid = XML.GetTagText("ACCOUNTGUID");
	frmScreen.txtMortgageAccount.value = XML.GetTagText("ACCOUNTNUMBER");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtMortgageAccount");
	
	GetSpecialFeatures();
	SetButtonStates();
}

function btnSubmit.onclick()
{
	if(frmScreen.onsubmit())
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else
			{
				scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
				frmToDC071.submit();
			}
	}
}

function frmScreen.btnAdd.onclick()
{	
	var sReturn = "";
	var ArrayArguments = new Array(4);
	
	m_bIsEdit = false;

	ArrayArguments[0] = m_XMLFeatures.CreateRequestAttributeArray(window);
	ArrayArguments[1] = "Add";
	ArrayArguments[2] = m_XMLFeatures.XMLDocument.xml;
	ArrayArguments[3] = m_sAccountGuid;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "dc076.asp", ArrayArguments, 375, 245);
	
	if (sReturn != null)
	{
		m_XMLFeatures.LoadXML(sReturn[1]);
		PopulateListBox(0);
		scScrollTable.setRowSelected(1);	
		SetButtonStates();
		FlagChange(sReturn[0]);
	}
}

function frmScreen.btnDelete.onclick()
{
	var nRowSelected = scScrollTable.getOffset() + scScrollTable.getRowSelected();
	
	m_bIsEdit = false;
	
	m_XMLFeatures.ActiveTag = null;
	m_XMLFeatures.CreateTagList("MORTGAGEACCOUNTSPECIALFEATURE");
	m_XMLFeatures.SelectTagListItem(nRowSelected -1);
	
	var XMLRequest = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLRequest.CreateRequestTag(window, null);
	XMLRequest.ActiveTag.appendChild(m_XMLFeatures.ActiveTag.cloneNode(true));
	
	// 	XMLRequest.RunASP(document,"DeleteSpecialFeature.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XMLRequest.RunASP(document,"DeleteSpecialFeature.asp");
			break;
		default: // Error
			XMLRequest.SetErrorResponse();
		}


	if (XMLRequest.IsResponseOK())
	{
		m_XMLFeatures.RemoveActiveTag();
		PopulateListBox(0);	
		if (m_nNumberOfFeatures > 0)
			scScrollTable.setRowSelected(1);	
		SetButtonStates();
		FlagChange(true);
	}
}


function frmScreen.btnEdit.onclick()
{
	var sReturn = "";
	var ArrayArguments = new Array(4);
	var nTableRow = scScrollTable.getRowSelected();
	var nRowSelected = scScrollTable.getOffset() + nTableRow;
	
	m_bIsEdit = true;
	
	ArrayArguments[0] = m_XMLFeatures.CreateRequestAttributeArray(window);
	ArrayArguments[1]= "Edit";
	ArrayArguments[2] = m_XMLFeatures.XMLDocument.xml;
	ArrayArguments[3] = nRowSelected -1;
	//GD BMIDS00376
	ArrayArguments[4] = m_bReadOnly;	
	
	//m_bReadOnly
	sReturn = scScreenFunctions.DisplayPopup(window, document, "dc076.asp", ArrayArguments, 375, 245);

	if (sReturn != null)
	{
		m_XMLFeatures.LoadXML(sReturn[1]);
		PopulateListBox(scScrollTable.getOffset());
		scScrollTable.setRowSelected(nTableRow);	
		SetButtonStates();
		FlagChange(sReturn[0]);
	}
}

function PopulateListBox(nStart)
{
<%  /* Populate the listbox with values from m_XMLFeatures */ 
%>	m_XMLFeatures.ActiveTag = null;
	m_XMLFeatures.CreateTagList("MORTGAGEACCOUNTSPECIALFEATURE");
	
	m_nNumberOfFeatures = m_XMLFeatures.ActiveTagList.length;
	
	if (!m_bIsEdit)
		scScrollTable.initialiseTable(tblFeatures, 0, "", ShowList, m_iTableLength, m_nNumberOfFeatures);
	
	ShowList(nStart);	
}

function ShowList(nStart)
{
<%  /* Populate the listbox with values from m_XMLFeatures Also used by the list scroll object */
%> 
	scScrollTable.clear();
	
	for (var iCount = 0; iCount < m_XMLFeatures.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		m_XMLFeatures.SelectTagListItem(iCount + nStart);
		
		var sDesc = m_XMLFeatures.GetTagText("MORTGAGEACCOUNTSPECIALFEATUREDESC");
		var sInd = m_XMLFeatures.GetTagText("MORTGAGEACCOUNTSPECIALFEATUREIND");
		
		scScreenFunctions.SizeTextToField(tblFeatures.rows(iCount+1).cells(0),sDesc);
		scScreenFunctions.SizeTextToField(tblFeatures.rows(iCount+1).cells(1),sInd);
		tblFeatures.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);
	}
}

function GetSpecialFeatures()
{
	m_XMLFeatures = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_XMLFeatures.CreateRequestTag(window, null)
	m_XMLFeatures.CreateActiveTag("MORTGAGEACCOUNTSPECIALFEATURE");
	m_XMLFeatures.CreateTag("ACCOUNTGUID", m_sAccountGuid);
	
	m_XMLFeatures.RunASP(document,"FindSpecialFeatureList.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = m_XMLFeatures.CheckResponse(ErrorTypes);
	if(ErrorReturn[0] == true)
	{
		PopulateListBox(0);
		SetButtonStates();
		scScrollTable.setRowSelected(1);	
	}
	else if (ErrorReturn[1] == ErrorTypes[0])
	{
		m_XMLFeatures = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_XMLFeatures.CreateActiveTag("MORTGAGEACCOUNTSPECIALFEATURELIST");
	}
}

function SetButtonStates()
{
	if (m_sReadOnly == "1")
	{
		frmScreen.btnDelete.disabled = true;
		//GD BMIDS00376 frmScreen.btnEdit.disabled = true;
		if(m_nNumberOfFeatures > 0)
		{
			frmScreen.btnEdit.disabled = false;
		} else
		{
			frmScreen.btnEdit.disabled = true;
		}
		frmScreen.btnAdd.disabled = true;
	}
	else
	{
		frmScreen.btnAdd.disabled = false;

		if(m_nNumberOfFeatures == 0)
		{
			frmScreen.btnDelete.disabled = true;
			frmScreen.btnEdit.disabled = true;
		}
		else
		{
			frmScreen.btnDelete.disabled = false;
			frmScreen.btnEdit.disabled = false;	
		}	
	}
}
-->
</script>
</body>
</html>


