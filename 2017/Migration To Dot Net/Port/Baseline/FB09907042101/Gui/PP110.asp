<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<%
/*
Workfile:      pp110.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Payee Search Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		09/02/01	New Screen
MC		21/03/01	SYS2131 - ThirdParty Search
SA		18/05/01	SYS2133 - Disabled OK button on entry. New Function added
					spnTable.OnClick to enable OK button again once entry selected.
SA		19/07/01	SYS2159 Bank Details not being displayed correctly.
SA		19/07/01	SYS2176	Bank name is shown as branch name Column header should be sort code not a/c number.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MC		05/05/2004	BMIDS751	- White space added to the title
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History :

Prog	Date		AQR			Description
HMA		22/09/2006	EP2_3		Add Roll Number.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
<meta NAME="GENERATOR" Content="MSHTML 5.00.2014.210"  >
<LINK href="stylesheet.css" rel=STYLESHEET type=text/css>
<TITLE>Payee Search Results<!-- #include file="includes/TitleWhiteSpace.asp" --></TITLE>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
</HEAD>

<BODY>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* Scriptlets */ %>
<span id="spnListScroll">
	<span style="LEFT: 310px; POSITION: absolute; TOP: 200px">
		<OBJECT data=scTableListScroll.asp id=scScrollTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 300px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>


<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 220px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px" class="msgGroup">
	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Third Party list</strong>
	</span>
	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 24px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="30%" class="TableHead">Third Party Name</td>	
				<td width="20%" class="TableHead">Third Party Type</td>	
				<td width="30%" class="TableHead">Bank Name</td>
				<td class="TableHead">Bank Sort Code</td></tr>
			<tr id="row01">		
				<td width="30%" class="TableTopLeft">&nbsp;</td>		
				<td width="20%" class="TableTopCenter">&nbsp;</td>		
				<td width="30%" class="TableTopCenter">&nbsp;</td>		
				<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		
				<td width="30%" class="TableLeft">&nbsp;</td>		
				<td width="20%" class="TableCenter">&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		
				<td width="30%" class="TableBottomLeft">&nbsp;</td>		
				<td width="20%" class="TableBottomCenter">&nbsp;</td>		
				<td width="30%" class="TableBottomCenter">&nbsp;</td>		
				<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>
</div>
</form>

<% /* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 250px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" -->
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/PP110Attribs.asp" -->

<%/* CODE */ %>
<script language="JScript">
<!--
var m_sPayeeType
var m_sXML ;
var m_sAppNumber = null;
var m_sAFFNumber = null;
//var m_sContext ;
var m_iTableLength = 10;

var XMLContext ;
var scScreenFunctions = null;
var m_BaseNonPopupWindow = null;

/* Events */

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sArguments = window.dialogArguments ;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	m_aArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	var sButtonList = new Array("Submit", "Cancel");
	ShowMainButtons(sButtonList);
	<% /* SYS2133 SA 18/5/01 Disable button until entry selected */ %>
	DisableMainButton("Submit");
	
	
	<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	SetMasks();
	Validation_Init();
	
	m_sPayeeType = m_aArgArray[0];
	m_sXML = m_aArgArray[1];
	m_sAppNumber = m_aArgArray[2];
	m_sAFFNumber = m_aArgArray[3];

	XMLContext = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	PopulateScreen();
	
	<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
	ClientPopulateScreen();
}

function btnCancel.onclick()
{
	window.close();
}

function btnSubmit.onclick()
{
	var saReturnArray = new Array();
	var iTagListItemCount ;
	
	var iCurrentRow = scScrollTable.getRowSelected();
	if(iCurrentRow == -1)
	{
		window.alert("Select a payee.") ;
		return ;
	}
	
	iTagListCount = tblTable.rows(iCurrentRow).getAttribute("TagListItemCount");
	XMLContext.SelectTagListItem(iTagListCount);
	
	saReturnArray[0] =  XMLContext.GetTagText('COMPANYNAME');
	saReturnArray[1] =  ""; //XMLContext.GetTagText('THIRDPARTYTYPE');
	saReturnArray[2] =  XMLContext.GetTagText('POSTCODE');
	saReturnArray[3] =  XMLContext.GetTagText('FLATNUMBER');
	saReturnArray[4] =  XMLContext.GetTagText('BUILDINGORHOUSENAME');
	saReturnArray[5] =  XMLContext.GetTagText('BUILDINGORHOUSENUMBER');
	saReturnArray[6] =  XMLContext.GetTagText('STREET');
	saReturnArray[7] =  XMLContext.GetTagText('TOWN');
	saReturnArray[8] =  XMLContext.GetTagText('DISTRICT');
	saReturnArray[9] =  XMLContext.GetTagText('COUNTY');
	saReturnArray[10] = XMLContext.GetTagText('COUNTRY'); 
	<% /* SYS2176 Shouldn't be branch name but actual name of bank */ %>
	<% /* saReturnArray[11] = XMLContext.GetTagText('BRANCHNAME'); */ %>
	saReturnArray[11] = XMLContext.GetTagText('COMPANYNAME');
	<% /* SYS2159 The sort code could be in the NAMEANDADDRESSBANKSORTCODE element. */ %>
	saReturnArray[12] = XMLContext.GetTagText('THIRDPARTYBANKSORTCODE'); 
	if (saReturnArray[12] == "")
	{
		saReturnArray[12] = XMLContext.GetTagText('NAMEANDADDRESSBANKSORTCODE'); 
	}
	<% /* SYS2159 Also should be returning the Account number */ %>
	saReturnArray[13] = XMLContext.GetTagText('ACCOUNTNUMBER'); 
	
	<%/* EP2_3 Return Roll Number */%>
	saReturnArray[14] = XMLContext.GetTagText('ROLLNUMBER'); 
	
	window.returnValue	= saReturnArray;
	window.close();
}

/* Functions */
function PopulateScreen()
{
	
	if(m_sXML != '')
	{	<% /* Third Party XML was passed from the calling screen */ %>
		XMLContext.LoadXML(m_sXML);		
	}
	else
	{	<% /* Fetch Third Party data from database */ %>
		XMLContext.CreateRequestTag(window,null);

		XMLContext.CreateActiveTag("APPLICATION");
		XMLContext.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XMLContext.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAFFNumber);
		XMLContext.RunASP(document,"FindApplicationThirdPartyList.asp");

		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = XMLContext.CheckResponse(ErrorTypes);

		if(ErrorReturn[0] == false)
		{
			if(ErrorReturn[1] == ErrorTypes[0]) 
				alert("No Third Parties exist for this Third Party Type.");
			else
				alert("Error in searching Third Party data");

			return;
		}
	}
	
	var sCondition = "APPLICATIONTHIRDPARTY[THIRDPARTY/THIRDPARTYTYPE='" + m_sPayeeType + "' || NAMEANDADDRESSDIRECTORY/NAMEANDADDRESSTYPE='" + m_sPayeeType + "']";
	XMLContext.CreateTagList(sCondition);

	PopulateListBox();
}

function PopulateListBox()
{
	var iNumberOfRows = XMLContext.ActiveTagList.length;

	scScrollTable.initialiseTable(tblTable, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(0);	
}

function ShowList(nStart)
{
	var sThirdPartyName, sThirdPartyType, sBankName, sBankSortCode;
	<% /* var sNAType, sNABankSortCode, sType="", sSortCode=""; */ %>
	
	for (var iCount = 0; iCount < XMLContext.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		XMLContext.SelectTagListItem(iCount + nStart);
		XMLContext.SelectSingleNode("THIRDPARTY");
		if(XMLContext.ActiveTag != null)
		{
			sThirdPartyName = XMLContext.GetTagText("COMPANYNAME");
			sThirdPartyType	= XMLContext.GetTagAttribute("THIRDPARTYTYPE", "TEXT")
			<% /* SYS2176 Should be displaying name of bank not branch. */ %>
			<% /* sBankName		= XMLContext.GetTagText("BRANCHNAME");  */ %>
			sBankName		= XMLContext.GetTagText("COMPANYNAME");
			sBankSortCode	= XMLContext.GetTagText("THIRDPARTYBANKSORTCODE");
		}
		else
		{
			XMLContext.SelectTagListItem(iCount + nStart);		
			XMLContext.SelectSingleNode("NAMEANDADDRESSDIRECTORY");
			if(XMLContext.ActiveTag != null)
			{
				sThirdPartyName = XMLContext.GetTagText("COMPANYNAME");
				sThirdPartyType	= XMLContext.GetTagAttribute("NAMEANDADDRESSTYPE", "TEXT")
				<% /* SYS2176 Should be displaying name of bank not branch. */ %>
				<% /* sBankName		= XMLContext.GetTagText("BRANCHNAME"); */ %>
				sBankName		= XMLContext.GetTagText("COMPANYNAME");
				sBankSortCode	= XMLContext.GetTagText("NAMEANDADDRESSBANKSORTCODE");
			}
		}
		
		<% /* Add to the table */ %>
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0), sThirdPartyName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1), sThirdPartyType);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2), sBankName);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(3), sBankSortCode);
		
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount + nStart);		
	}
}
<% /* SYS2133 SA Enable Submit button when entry selected */ %>
function spnTable.onclick()
{
	EnableMainButton("Submit");
}
-->
</script>
</BODY>
</HTML>



