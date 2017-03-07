<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc120.asp
Copyright:     Copyright © 1999 Marlborough Stirling

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		20/09/1999	Created
AD		31/01/2000	Rework
AY		14/02/00	Change to msgButtons button types
IVW		24/03/2000	SYS0435 - Default the selection to the first one in the list
AY		30/03/00	New top menu/scScreenFunctions change
SR		12/05/00    SYS0732 - Set the width of 'Span' covering the radio group to
					appropriate value.
BG		17/05/00	SYS0752 Removed Tooltips
CL		05/03/01	SYS1920 Read only functionality added
SA		31/05/01	SYS0933 Read only - disable radio/add/delete buttons
SA		06/06/01	SYS1672 Collect details for guarantors as well
DJP		06/08/01	SYS2564/SYS2786 (child) Client cosmetic customisation which in this case
					is routing.
JLD		4/12/01		SYS2806 completeness check routing
DPF		20/06/02	BMIDS00077 Changes made to bring file in line with Core V7.0.2. Changes are...
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		17/05/2002	BMIDS00008	Modified Routing Screen on btnSubmit.Onclick() 
ASu		11/10/2002	BMIDS00610	Change routing from DC100 to DC080 (DC100 removed) 
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
BS		10/06/2003	BM0521	Don't enable Delete when record selected if screen is in read only
INR		07/07/2003  BMIDS597 Need to default to an indicator of 'NO' if DECLINEDMORTGAGEINDICATOR has never been set
SR		15/06/2004	BMIDS772	Remove the question and cancel button. Change processing accordingly
								Using asp files instead of html for ScrollList functionality
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog	Date		AQR			Description
MF		26/07/2005	MAR019		Changed screen routing to financial lists DC085

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
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>
<% /* removed as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
*/ %>

<span id="spnListScroll">
	<span style="LEFT: 310px; POSITION: absolute; TOP: 245px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<%/* FORMS */%>
<form id="frmToDC121" method="post" action="dc121.asp" STYLE="DISPLAY: none"></form>
<% /* SR 14/06/2004 : BMIDS772 - Route to DC110 instead of DC080 on submit */ %>
<form id="frmToDC110" method="post" action="dc110.asp" STYLE="DISPLAY: none"></form> 
<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC085" method="post" action="DC085.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 215px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span id="spnDeclinedMortgages" style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblDeclinedMortgages" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="25%" class="TableHead">Name&nbsp</td>		<td width="15%" class="TableHead">Date<br>Declined&nbsp</td>	<td width="45%" class="TableHead">Details&nbsp</td>		<td class="TableHead">Additional<br>Other(s)&nbsp</td></tr>
			<tr id="row01">		<td width="25%" class="TableTopLeft">&nbsp</td>		<td width="15%" class="TableTopCenter">&nbsp</td>			<td width="45%" class="TableTopCenter">&nbsp</td>		<td class="TableTopRight">&nbsp</td></tr>
			<tr id="row02">		<td width="25%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>				<td width="45%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
			<tr id="row03">		<td width="25%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>				<td width="45%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
			<tr id="row04">		<td width="25%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>				<td width="45%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
			<tr id="row05">		<td width="25%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>				<td width="45%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
			<tr id="row06">		<td width="25%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>				<td width="45%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
			<tr id="row07">		<td width="25%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>				<td width="45%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
			<tr id="row08">		<td width="25%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>				<td width="45%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
			<tr id="row09">		<td width="25%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>				<td width="45%" class="TableCenter">&nbsp</td>			<td class="TableRight">&nbsp</td></tr>
			<tr id="row10">		<td width="25%" class="TableBottomLeft">&nbsp</td>	<td width="15%" class="TableBottomCenter">&nbsp</td>			<td width="45%" class="TableBottomCenter">&nbsp</td>		<td class="TableBottomRight">&nbsp</td></tr>
		</table>
	</span>

	<span id="spnButtons" style="TOP: 185px; LEFT: 4px; POSITION: ABSOLUTE">
		<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
			<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="TOP: 0px; LEFT: 64px; POSITION: ABSOLUTE">
			<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="TOP: 0px; LEFT: 128px; POSITION: ABSOLUTE">
			<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
	</span>
</div>
</form>

<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 320px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="routing/dc120routing.asp" -->	
<!-- Specify Code Here -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc120Attribs.asp" -->
<script language="JScript">
<!--
var ListXML;
var m_sMetaAction = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sQuestionOnEntry	= null;
var m_bIsFinancialSummary = false;
var scScreenFunctions;
var m_blnReadOnly = false;
var m_iTableLength = 10 ; <% /* SR 14/06/2004 : BMIDS772 */ %>


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	FW030SetTitles("Declined Mortgage Details","DC120",scScreenFunctions);
	
<%	// IVW - 24/03/2000 - SYS0435 - Default the first item in the list box.
	// IVW - Do before the populate functionality otherwise whatever defaults are set
	// are lost when we return here.
%>	
	RetrieveContextData();
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();

	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC120");
	if (m_blnReadOnly == true)
	{	
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnDelete.disabled = true;	
	}
	
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
	m_sMetaAction	= scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
}

function spnDeclinedMortgages.onclick()
{
	if (scTable.getRowSelected() != -1) 
	{
		frmScreen.btnEdit.disabled = false;
		<% /* BS BM0521 Only enable delete if screen is in edit mode */%>
		if (m_blnReadOnly != true)
		frmScreen.btnDelete.disabled = false;
	}
}

function PopulateScreen()
{
	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//ListXML = new scXMLFunctions.XMLObject();
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	ListXML.CreateRequestTag(window,null);
	var TagSEARCH = ListXML.CreateActiveTag("SEARCH");
	ListXML.CreateActiveTag("FINANCIALSUMMARY");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);

	ListXML.ActiveTag = TagSEARCH;
	var TagDECLINEDMORTGAGELIST = ListXML.CreateActiveTag("DECLINEDMORTGAGELIST");

	// Loop through all customer context entries
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		// If the customer is an applicant, add him/her to the search
		//SYS1672 Guarantor details too
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			ListXML.CreateActiveTag("DECLINEDMORTGAGE");
			ListXML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			ListXML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			ListXML.ActiveTag = TagDECLINEDMORTGAGELIST;
		}
	}

	ListXML.RunASP(document,"FindDeclinedMortgageSummary.asp");

	// A record not found error is valid
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);

	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		PopulateTable(0);
	}
}

function PopulateTable(nStart)
{
	var iNumberOfRows ;
	
	scTable.clear();

	ListXML.ActiveTag = null;
	ListXML.CreateTagList("DECLINEDMORTGAGE");
	iNumberOfRows = ListXML.ActiveTagList.length
	
	scTable.initialiseTable(tblDeclinedMortgages, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}

function ShowList(nStart)
{	
	var sCustomerNumber, sName, sDateDeclined, sDetails, sOthersResponsible, sOthersResponsibleField
	var nRow = 0;

	// Populate the table with a set of records, starting with record nStart
	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) != false && nLoop < 10; nLoop++)
	{
		sCustomerNumber = ListXML.GetTagText("CUSTOMERNUMBER");
		sName = scScreenFunctions.GetContextCustomerName(window,sCustomerNumber);
		sDateDeclined = ListXML.GetTagText("DATEDECLINED");
		sDetails = ListXML.GetTagText("DECLINEDDETAILS");
		sOthersResponsible = ListXML.GetTagText("ADDITIONALINDICATOR");

		<% /* Display Yes or No for if there are others responsible or not */ %>
		sOthersResponsibleField = (sOthersResponsible == "1") ? "Yes" : "No"

		<% /* Display the details in the appropriate table row */ %>
		nRow = nLoop + 1;
		scScreenFunctions.SizeTextToField(tblDeclinedMortgages.rows(nRow).cells(0), sName);
		scScreenFunctions.SizeTextToField(tblDeclinedMortgages.rows(nRow).cells(1), sDateDeclined);
		scScreenFunctions.SizeTextToField(tblDeclinedMortgages.rows(nRow).cells(2), sDetails);
		scScreenFunctions.SizeTextToField(tblDeclinedMortgages.rows(nRow).cells(3), sOthersResponsibleField);
	}
	
	// IVW - 24/03/2000 - SYS0435 - Default the first item in the list box.	
	if(nRow > 0)
	{
		scTable.setRowSelected(1);
		frmScreen.btnEdit.disabled = false;
		frmScreen.btnDelete.disabled = false;
	}
	else
	{
		frmScreen.btnEdit.disabled = true;
		frmScreen.btnDelete.disabled = true;	
	}		
}

function frmScreen.btnEdit.onclick()
{
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("DECLINEDMORTGAGE");

	// Get the index of the selected row
	var nRowSelected =  scTable.getRowSelectedIndex();

	if(ListXML.SelectTagListItem(nRowSelected-1) == true)
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		scScreenFunctions.SetContextParameter(window,"idXML",ListXML.ActiveTag.xml);
		frmToDC121.submit();
	}
}

function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToDC121.submit();
}

function frmScreen.btnDelete.onclick()
{	
	if (!confirm("Are you sure?")) return;

	ListXML.ActiveTag = null;
	ListXML.CreateTagList("DECLINEDMORTGAGE");

	// Get the index of the selected row
	var nRowSelected = scTable.getRowSelectedIndex();

	ListXML.SelectTagListItem(nRowSelected-1);

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	var xmlRequest = XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("DELETE");
	XML.CreateActiveTag("DECLINEDMORTGAGE");
	XML.CreateTag("CUSTOMERNUMBER", ListXML.GetTagText("CUSTOMERNUMBER"));
	XML.CreateTag("CUSTOMERVERSIONNUMBER", ListXML.GetTagText("CUSTOMERVERSIONNUMBER"));
	XML.CreateTag("SEQUENCENUMBER", ListXML.GetTagText("SEQUENCENUMBER"));
	
	<% /* SR 14/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
			node to CustomerFinancialBO. If the form accessed through completeness check, update FS always.
			Else, do it only when all the records are deleted from the list box
	*/ %>
	if( scScreenFunctions.CompletenessCheckRouting(window) || scTable.getTotalRecords() == 1)
	{
		XML.ActiveTag = xmlRequest ;
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		if(scTable.getTotalRecords() == 1) XML.CreateTag("DECLINEDMORTGAGEINDICATOR", 0)
		else XML.CreateTag("DECLINEDMORTGAGEINDICATOR", 1) 
	} 
	<% /* SR 14/06/2004 : BMIDS772 - End */ %>

	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"DeleteDeclinedMortgage.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}


	// If the deletion is successful remove the entry from the list xml and the screen
	if(XML.IsResponseOK())
	{
		PopulateScreen();
	}
	XML = null;
}

function btnSubmit.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window)) frmToGN300.submit();
	else frmToDC085.submit();  <% /* MF 26/07/2005 frmToDC110.submit();  SR 14/06/2004 : BMIDS772 - route to DC110 on submit*/ %>
}

-->
</script>
</body>
</html>




