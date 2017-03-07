<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      DC150.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   CCJ History Summary
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		28/09/1999	Created
AD		03/02/2000	Rework
AY		14/02/00	Change to msgButtons button types
IW		23/02/00	SYSS0149 - Cosmetic Changes
IVW		24/03/2000	SYS0439 - Default the selection to the first one in the list
AY		30/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
MH      23/06/00    SYS0933 Readonly status
BG		22/08/00	SYS0803 Changed routing from OK to go to DC155 instead of DC160
CL		05/03/01	SYS1920 Read only functionality added
SA		31/05/01	SYS0933 Enable Edit button and disable Delete button in read only
SA		06/06/01	SYS1672 Collect details for guarantors as well
JLD		4/12/01		SYS2806 use screen functions to check completeness check routing
DPF		20/06/02	BMIDS00077	Amended file to bring in line with Core V7.0.2, changes made...
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		15/05/2002	BMIDS00008	Modified btnSubmit.Onclick()
MV		17/05/2002	BMIDS00008	Modified Routing Screen on btnSubmit.Onclick()
GHun	31/07/2002	BMIDS00190	Support multiple customers per CCJ history record 
MDC		02/09/2002	BMIDS00336	Credit Check & Bureau Download
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
MDC		03/12/2002	BMIDS01124	Show 'Unassigned' as customer name if unassigned.
JR		23/01/2003	BM0271		Check for ReadOnly on clicking Table.
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
<script src="validation.js" language="JScript"></script>
<span id="spnListScroll">
	<span style="LEFT: 310px; POSITION: absolute; TOP: 245px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<%/* FORMS */%>
<form id="frmToDC151" method="post" action="dc151.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC140" method="post" action="dc140.asp" STYLE="DISPLAY: none"></form>
<% /* SR 15/06/2004 : BMIDS772 - Route to DC110 on submit (instead of DC120) */ %>
<form id="frmToDC110" method="post" action="DC110.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC085" method="post" action="DC085.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 280px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span id="spnCCJHistory" style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblCCJHistory" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="15%" class="TableHead">Name&nbsp</td>		<td width="15%" class="TableHead">Type&nbsp</td>	<td width="15%" class="TableHead">Value Of<br>Judgement&nbsp</td>	<td width="15%" class="TableHead">Plaintiff&nbsp</td>	<td width="10%" class="TableHead">Monthly<br>Repayment&nbsp</td>	<td width="10%" class="TableHead">Date<br>Cleared&nbsp</td>	<td width="10%" class="TableHead">Others<br>Resp.&nbsp</td>		<td class="TableHead">Unassigned</td></tr>
			<tr id="row01">		<td width="15%" class="TableTopLeft">&nbsp</td>			<td width="15%" class="TableTopCenter">&nbsp</td>	<td width="15%" class="TableTopCenter">&nbsp</td>					<td width="15%" class="TableTopCenter">&nbsp</td>		<td width="10%" class="TableTopCenter">&nbsp</td>					<td width="10%" class="TableTopCenter">&nbsp</td>			<td width="10%" class="TableTopCenter">&nbsp</td>				<td class="TableTopRight">&nbsp</td></tr>
			<tr id="row02">		<td width="15%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>		<td width="15%" class="TableCenter">&nbsp</td>						<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>						<td width="10%" class="TableCenter">&nbsp</td>				<td width="10%" class="TableCenter">&nbsp</td>					<td class="TableRight">&nbsp</td></tr>
			<tr id="row03">		<td width="15%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>		<td width="15%" class="TableCenter">&nbsp</td>						<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>						<td width="10%" class="TableCenter">&nbsp</td>				<td width="10%" class="TableCenter">&nbsp</td>					<td class="TableRight">&nbsp</td></tr>
			<tr id="row04">		<td width="15%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>		<td width="15%" class="TableCenter">&nbsp</td>						<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>						<td width="10%" class="TableCenter">&nbsp</td>				<td width="10%" class="TableCenter">&nbsp</td>					<td class="TableRight">&nbsp</td></tr>
			<tr id="row05">		<td width="15%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>		<td width="15%" class="TableCenter">&nbsp</td>						<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>						<td width="10%" class="TableCenter">&nbsp</td>				<td width="10%" class="TableCenter">&nbsp</td>					<td class="TableRight">&nbsp</td></tr>
			<tr id="row06">		<td width="15%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>		<td width="15%" class="TableCenter">&nbsp</td>						<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>						<td width="10%" class="TableCenter">&nbsp</td>				<td width="10%" class="TableCenter">&nbsp</td>					<td class="TableRight">&nbsp</td></tr>
			<tr id="row07">		<td width="15%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>		<td width="15%" class="TableCenter">&nbsp</td>						<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>						<td width="10%" class="TableCenter">&nbsp</td>				<td width="10%" class="TableCenter">&nbsp</td>					<td class="TableRight">&nbsp</td></tr>
			<tr id="row08">		<td width="15%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>		<td width="15%" class="TableCenter">&nbsp</td>						<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>						<td width="10%" class="TableCenter">&nbsp</td>				<td width="10%" class="TableCenter">&nbsp</td>					<td class="TableRight">&nbsp</td></tr>
			<tr id="row09">		<td width="15%" class="TableLeft">&nbsp</td>			<td width="15%" class="TableCenter">&nbsp</td>		<td width="15%" class="TableCenter">&nbsp</td>						<td width="15%" class="TableCenter">&nbsp</td>			<td width="10%" class="TableCenter">&nbsp</td>						<td width="10%" class="TableCenter">&nbsp</td>				<td width="10%" class="TableCenter">&nbsp</td>					<td class="TableRight">&nbsp</td></tr>
			<tr id="row10">		<td width="15%" class="TableBottomLeft">&nbsp</td>		<td width="15%" class="TableBottomCenter">&nbsp</td><td width="15%" class="TableBottomCenter">&nbsp</td>				<td width="15%" class="TableBottomCenter">&nbsp</td>	<td width="10%" class="TableBottomCenter">&nbsp</td>				<td width="10%" class="TableBottomCenter">&nbsp</td>		<td width="10%" class="TableBottomCenter">&nbsp</td>			<td class="TableBottomRight">&nbsp</td></tr>
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

	<div style="TOP: 222px; LEFT: 390px; HEIGHT: 52px; WIDTH: 208px; POSITION: ABSOLUTE" class="msgGroupLight">
		<span style="TOP: 7px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			Total CCJs
			<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
				<input id="txtTotalCCJs" name="TotalCCJs" maxlength="3" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 33px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
			Total Judgement Amount
			<span style="TOP: -3px; LEFT: 130px; POSITION: ABSOLUTE">
				<input id="txtTotalJudgementAmount" name="TotalJudgementAmount" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
	</div>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 380px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc150Attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--
var ListXML;
var m_sMetaAction = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sQuestionOnEntry	= null;
var m_sReadOnly ="";
var scScreenFunctions;
var m_blnReadOnly = false;
var m_iTableLength = 10; <% /* SR 14/06/2004 : BMIDS772 */ %>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<% /* SR 15/06/2004 : BMIDS772 - Remove Cancel button */ %>
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	FW030SetTitles("CCJ History","DC150",scScreenFunctions);
	
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTotalCCJs");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtTotalJudgementAmount");

	RetrieveContextData();
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	
	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC150");

	if (m_sReadOnly =="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnAdd.disabled =true;
		frmScreen.btnEdit.disabled =true;
		frmScreen.btnDelete.disabled=true;
	}
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	<% /* MO - 18/11/2002 - BMIDS00376m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC150"); */ %>
	
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
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);	
}

function spnCCJHistory.onclick()
{
	if (scTable.getRowSelected() != -1) 
	{
		frmScreen.btnEdit.disabled = false;
		if(m_sReadOnly == true) //JR BM0271
			frmScreen.btnDelete.disabled = true;
		else
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
	var TagCCJHISTORYLIST = ListXML.CreateActiveTag("CCJHISTORYLIST");

	// Loop through all customer context entries
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		// If the customer is an applicant, add him/her to the search
		<% /* SYS1672 Collect for Guarantors as well */ %>
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			ListXML.CreateActiveTag("CCJHISTORY");
			ListXML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			ListXML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			ListXML.ActiveTag = TagCCJHISTORYLIST;
		}
	}

	ListXML.RunASP(document,"FindCCJHistorySummary.asp");

	<% /*  A record not found error is valid */ %>
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);

	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		PopulateTable(0);

		if(ListXML.SelectTag(null,"CCJHISTORYLIST") != null)
		{
			frmScreen.txtTotalCCJs.value = ListXML.GetTagText("TOTALCCJS");
			frmScreen.txtTotalJudgementAmount.value	= ListXML.GetTagText("TOTALJUDGEMENTAMOUNT");
		}
		else
		{
			frmScreen.txtTotalCCJs.value = "" ;
			frmScreen.txtTotalJudgementAmount.value	= "";
		}
	}
}

function PopulateTable(nStart)
{
	var iNumberOfRows ;
	
	scTable.clear();

	ListXML.ActiveTag = null;
	ListXML.CreateTagList("CCJHISTORY");
	iNumberOfRows = ListXML.ActiveTagList.length ;
	
	scTable.initialiseTable(tblCCJHistory, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}

function ShowList(nStart)
{
	var nRow = 0;

	<% /* BMIDS00190 */ %>
	var sCustomers;
	var sName;
	var CustomersXML;
	<% /* BMIDS00190 End */ %>
	var sValueOfJudgement, sPlaintiff, sMonthlyRepayment, sDateCleared
	var sOthersResponsible, sOthersResponsibleField, sUnassigned, sCCJType

	<% /* Populate the table with a set of records, starting with record nStart */ %>
	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) != false && nLoop < 10; nLoop++)
	{				
		<% /* BMIDS00190 support display of multiple customers */ %>
		sCustomers = "";
		CustomersXML = ListXML.ActiveTag.cloneNode(true).selectNodes("CUSTOMERVERSIONCCJHISTORY");
		
		for(var nCust=0; nCust < CustomersXML.length; nCust++)
		{
			sCustomerNumber = CustomersXML.item(nCust).selectSingleNode("CUSTOMERNUMBER").text;
			sName = scScreenFunctions.GetContextCustomerName(window,sCustomerNumber);
			if (sCustomers.length > 0)
				sCustomers = sCustomers + " & " + sName;
			else
				sCustomers = sName;
		}
		<% /* BMIDS00190 End */ %>
		
		sValueOfJudgement = ListXML.GetTagText("VALUEOFJUDGEMENT");
		sPlaintiff = ListXML.GetTagText("PLAINTIFF");
		sMonthlyRepayment = ListXML.GetTagText("MONTHLYREPAYMENT");
		sDateCleared = ListXML.GetTagText("DATECLEARED");
		sOthersResponsible = ListXML.GetTagText("ADDITIONALINDICATOR");
		sOthersResponsibleField = (sOthersResponsible == "1") ? "Yes" : "No";

		<% /* BMIDS00336 MDC 02/09/2002 */ %>
		sUnassigned	= ListXML.GetTagText("UNASSIGNED");
		if(sUnassigned == "1")
		{
			sUnassigned = "Yes";
			sCustomers = "Unassigned";	<% /* BMIDS01124 MDC 03/12/2002 */ %>
		}
		else
			sUnassigned = "No";

		sCCJType = ListXML.GetTagAttribute("CCJTYPE","TEXT");

		<% /*  Display the details in the appropriate table row */ %>
		nRow = nLoop + 1;
		scScreenFunctions.SizeTextToField(tblCCJHistory.rows(nRow).cells(0), sCustomers);
		scScreenFunctions.SizeTextToField(tblCCJHistory.rows(nRow).cells(1), sCCJType);
		scScreenFunctions.SizeTextToField(tblCCJHistory.rows(nRow).cells(2), sValueOfJudgement);
		scScreenFunctions.SizeTextToField(tblCCJHistory.rows(nRow).cells(3), sPlaintiff);
		scScreenFunctions.SizeTextToField(tblCCJHistory.rows(nRow).cells(4), sMonthlyRepayment);
		scScreenFunctions.SizeTextToField(tblCCJHistory.rows(nRow).cells(5), sDateCleared);
		scScreenFunctions.SizeTextToField(tblCCJHistory.rows(nRow).cells(6), sOthersResponsibleField);
		scScreenFunctions.SizeTextToField(tblCCJHistory.rows(nRow).cells(7), sUnassigned);
		<% /* BMIDS00336 MDC 02/09/2002 - End */ %>
	}
	
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
	ListXML.CreateTagList("CCJHISTORY");

	// Get the index of the selected row
	var nRowSelected =  scTable.getRowSelectedIndex();

	if(ListXML.SelectTagListItem(nRowSelected-1))
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		scScreenFunctions.SetContextParameter(window,"idXML",ListXML.ActiveTag.xml);
		frmToDC151.submit();
	}
}

function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToDC151.submit();

}

function frmScreen.btnDelete.onclick()
{	
	if (!confirm("Are you sure?")) return;

	ListXML.ActiveTag = null;
	ListXML.CreateTagList("CCJHISTORY");

	var nRowSelected = scTable.getRowSelectedIndex();
	ListXML.SelectTagListItem(nRowSelected-1);

	// Set up the deletion XML
	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02
	//var XML = new scXMLFunctions.XMLObject();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var xmlRequest = XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("DELETE");
	XML.CreateActiveTag("CCJHISTORY");
	
	<% /* BMIDS00190 */ %>
	XML.CreateTag("CCJHISTORYGUID", ListXML.GetTagText("CCJHISTORYGUID"));
	
	var CustomersXML = ListXML.ActiveTag.selectNodes("CUSTOMERVERSIONCCJHISTORY");
	for(var nCust=0; nCust < CustomersXML.length; nCust++)
		XML.ActiveTag.appendChild(CustomersXML.item(nCust).cloneNode(true));
	<%
	//XML.CreateTag("CUSTOMERNUMBER", ListXML.GetTagText("CUSTOMERNUMBER"));
	//XML.CreateTag("CUSTOMERVERSIONNUMBER", ListXML.GetTagText("CUSTOMERVERSIONNUMBER"));
	//XML.CreateTag("SEQUENCENUMBER", ListXML.GetTagText("SEQUENCENUMBER"));
	/* BMIDS00190 End */ %>

	<% /* SR 16/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
			node to CustomerFinancialBO. If the form accessed through completeness check, update FS always.
			Else, do it only when all the records are deleted from the list box
	*/ %>
	if( scScreenFunctions.CompletenessCheckRouting(window) || scTable.getTotalRecords() == 1)
	{
		XML.ActiveTag = xmlRequest ;
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		if(scTable.getTotalRecords() == 1) XML.CreateTag("CCJHISTORYINDICATOR", 0)
		else XML.CreateTag("CCJHISTORYINDICATOR", 1) 
	} 
	<% /* SR 16/06/2004 : BMIDS772 - End */ %>
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"DeleteCCJHistory.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}


	// If the deletion is successful remove the entry from the list xml and the screen
	if(XML.IsResponseOK() == true)
	{
		PopulateScreen();
	}
	XML = null;
}

function btnSubmit.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window)) frmToGN300.submit();
	else frmToDC085.submit();<% /* MF 26/07/2005 changed routing frmToDC110.submit(); */ %>
}

-->
</script>
</body>
</html>




