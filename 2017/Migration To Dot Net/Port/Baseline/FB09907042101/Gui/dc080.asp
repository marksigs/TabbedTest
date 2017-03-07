<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc080.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Bank/Credit Card Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		02/02/2000	ie5 changes 
AY		14/02/00	Change to msgButtons button types
MC		21/03/2000	Fixed SYS0431
AY		30/03/00	New top menu/scScreenFunctions change
MH      03/05/00    SYS0492 Collect CC details for guarantors as well
BG		17/05/00	SYS0752 Removed Tooltips
BG		19/05/00	SYS0703	After delete processing highlight first entry in listbox.
MH      22/05/00    SYS0933 Readonly handling
CL		05/03/01	SYS1920 Read only functionality added
JLD		4/12/01		SYS2806 Completeness Check Routing
DPF		20/06/02	BMIDS00077 - amendments made to this file to bring it in line with
					V7.0.2 of Core.  The following actions have been taken...
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		15/05/2002	BMIDS00008	Modified btnCancel.Onclick()
MV		17/05/2002	BMIDS00008	Modified Routing Screen on btnSubmit.Onclick() , btnCancel.onclick()
ASu		11/10/2002	BMIDS00610	Change 'Cancel' routing from DC100 to DC120 (DC100 removed) 
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
BS		10/06/2003	BM0521	Don't enable Delete when record selected if screen is in read only
INR		07/07/2003  BMIDS597 Need to default to an indicator of 'NO' if BANKCARDINDICATOR has never been set
SR		01/06/2004	BMIDS772	Remove the question and cancel button. Change processing accordingly.
								Using asp files instead of html for ScrollList functionality
     								
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARSS Specific History:

Prog	Date		AQR			Description
DRC     03/01/2006  MAR1189     Enable/Disable Add Button based on global param FSDisableBankCredCardAddButton
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
<script src="validation.js" language="JScript"></script>
<% /* removed as per V7.0.2 of Core - DPF 20/06/02 - BMIDS00077 
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
*/ %>
<span id="spnListScroll">
	<span style="LEFT: 310px; POSITION: absolute; TOP: 245px">
		<OBJECT data=scTableListScroll.asp id=scScrollTable name=scTable  style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<form id="frmToDC085" method="post" action="DC085.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC081" method="post" action="dc081.asp" STYLE="DISPLAY: none"></form>
<% /* ASu - BMIDS00610 - Start. Removed DC100.asp
<form id="frmToDC100" method="post" action="dc100.asp" STYLE="DISPLAY: none"></form>  */%>
<form id="frmToDC120" method="post" action="dc120.asp" STYLE="DISPLAY: none"></form>
<% /* ASu - BMIDS00610 - Start */ %> 
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 230px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span id="spnBankCreditCardDetails" style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblBankCreditCardDetails" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
			<tr id="rowTitles"><td width="25%" class="TableHead">Name</td><td width="15%" class="TableHead">Card Type</td><td width="22%" class="TableHead">Card Provider</td><td width="12%" class="TableHead">Outstanding<br>Balance</td>	<td width="12%" class="TableHead">Average<br>Repayment</td><td class="TableHead">Additional<br>Holder(s)</td></tr>
			<tr id="row01"><td width="25%" class="TableTopLeft">&nbsp</td><td width="15%" class="TableTopCenter">&nbsp</td><td width="22%" class="TableTopCenter">&nbsp</td><td width="12%" class="TableTopCenter">&nbsp</td><td width="12%" class="TableTopCenter">&nbsp</td><td class="TableTopRight">&nbsp</td></tr>
			<tr id="row02"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="22%" class="TableCenter">&nbsp</td><td width="12%" class="TableCenter"></td><td width="12%" class="TableCenter"></td><td class="TableRight">&nbsp</td></tr>
			<tr id="row03"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="22%" class="TableCenter">&nbsp</td><td width="12%" class="TableCenter"></td><td width="12%" class="TableCenter"></td><td class="TableRight">&nbsp</td></tr>
			<tr id="row04"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="22%" class="TableCenter">&nbsp</td><td width="12%" class="TableCenter"></td><td width="12%" class="TableCenter"></td><td class="TableRight">&nbsp</td></tr>
			<tr id="row05"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="22%" class="TableCenter">&nbsp</td><td width="12%" class="TableCenter"></td><td width="12%" class="TableCenter"></td><td class="TableRight">&nbsp</td></tr>
			<tr id="row06"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="22%" class="TableCenter">&nbsp</td><td width="12%" class="TableCenter"></td><td width="12%" class="TableCenter"></td><td class="TableRight">&nbsp</td></tr>
			<tr id="row07"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="22%" class="TableCenter">&nbsp</td><td width="12%" class="TableCenter"></td><td width="12%" class="TableCenter"></td><td class="TableRight">&nbsp</td></tr>
			<tr id="row08"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="22%" class="TableCenter">&nbsp</td><td width="12%" class="TableCenter"></td><td width="12%" class="TableCenter"></td><td class="TableRight">&nbsp</td></tr>
			<tr id="row09"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="22%" class="TableCenter">&nbsp</td><td width="12%" class="TableCenter"></td><td width="12%" class="TableCenter"></td><td class="TableRight">&nbsp</td></tr>
			<tr id="row10"><td width="25%" class="TableBottomLeft">&nbsp</td><td width="15%" class="TableBottomCenter">&nbsp</td><td width="22%" class="TableBottomCenter">&nbsp</td><td width="12%" class="TableBottomCenter">&nbsp</td><td width="12%" class="TableBottomCenter">&nbsp</td><td class="TableBottomRight">&nbsp</td></tr>
		</table>
	</span>

	<span id="spnButtons" style="TOP: 200px; LEFT: 4px; POSITION: ABSOLUTE">
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
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 330px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc080Attribs.asp" -->

<script language="JScript">
<!--
var ListXML;
var m_sMetaAction					= null;
var m_sUserType						= null;
var m_sUnitId						= null;
var m_sUserId						= null;
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
<% /* Radio button value on entry */ %>
var m_sQuestionOnEntry	= null;		
<% /* Is there a FinancialSummary record */ %>
var scScreenFunctions;
var m_sReadOnly = null;
var m_blnReadOnly = false;
var m_iTableLength = 10;

		

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<% /* Make the required buttons available on the bottom of the screen
		(see msgButtons.asp for details) */ %>
	<% /*  SR 01/06/2004 : BMIDS772 - Remove Cancel button */ %>
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	FW030SetTitles("Bank/Credit Card Details","DC080",scScreenFunctions);

	RetrieveContextData();
			
	<% /* Default Edit/Delete to disabled */ %>
	frmScreen.btnEdit.disabled = true;
	frmScreen.btnDelete.disabled = true;
	
	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	
	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC080");

	if (m_sReadOnly=="1") 
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnAdd.disabled=true;
		frmScreen.btnDelete.disabled=true;
	}	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	<% /* MO - 18/11/2002 - BMIDS00376 m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC080"); */ %>
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

<% /* keep the focus within	this screen when using the tab key */ %>
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<% /* keep the focus within	this screen when using the tab key */ %>
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

<% /* Retrieves all context data required for use within this screen */ %>
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", "1627");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", "1");
}

<% /* Handles the onclick event from the span surrounding 
	the table. This is done here to handle the enabling 
	of buttons when a row is selected. Using the principle 
	of event bubbling we pick up the onclick event after 
	the table_onclick event in the scTable.htm scriptlet */ %>
function spnBankCreditCardDetails.onclick()
{
	if (scTable.getRowSelected() != -1) 
	{
		frmScreen.btnEdit.disabled = false;	
		<% /* BS BM0521 Only enable delete if screen is in edit mode */%>
		if (m_sReadOnly != "1")			
		frmScreen.btnDelete.disabled = false;					
	}
}

<% /* Retrieves the data and sets the screen accordingly */ %>
function PopulateScreen()
{
	//next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//ListXML = new scXMLFunctions.XMLObject();
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	<% /* <REQUEST USERID=? USERTYPE=? UNIT=?> */ %>
	var TagSEARCH = ListXML.CreateRequestTag(window, "SEARCH");
	ListXML.CreateActiveTag("FINANCIALSUMMARY");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);

	ListXML.ActiveTag = TagSEARCH;
	var TagBANKCREDITCARDLIST = ListXML.CreateActiveTag("BANKCREDITCARDLIST");
			
	<% /* Loop through all customer context entries */ %>
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		<% /* If the customer is an applicant, add him/her to the search */ %>
		<% /* SYS0492 Collect for Guarantors as well */ %>
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			ListXML.CreateActiveTag("BANKCREDITCARD");
			ListXML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			ListXML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			ListXML.ActiveTag = TagBANKCREDITCARDLIST;
		}
	}
	ListXML.RunASP(document, "FindBankCardSummary.asp");
			
	<% /* A record not found error is valid */ %>
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);
			
	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{		
		PopulateTable(0);
	}
}

<% /* Displays a set of records in the table. */ %>
function PopulateTable(nStart)
{	
	var iNumberOfRows ;
				
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("BANKCREDITCARD");
	iNumberOfRows = ListXML.ActiveTagList.length
	
	scTable.initialiseTable(tblBankCreditCardDetails, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}
			
function ShowList(nStart)
{
	var sCustomerNumber, sName, sCardType, sProvider, sBalance, sRepayment, sAdditionalCardHolder
	
	scTable.clear();
	
	<% /* Populate the table with a set of records, starting with record nStart */ %>
	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) != false && nLoop < 10; nLoop++)
	{				
		sCustomerNumber = ListXML.GetTagText("CUSTOMERNUMBER");
		sName = scScreenFunctions.GetContextCustomerName(window,sCustomerNumber);
		sCardType = ListXML.GetTagAttribute("CARDTYPE","TEXT");
		sProvider = ListXML.GetTagText("CARDPROVIDER");
		sBalance = ListXML.GetTagText("TOTALOUTSTANDINGBALANCE");
		sRepayment = ListXML.GetTagText("AVERAGEMONTHLYREPAYMENT");
		sAdditionalCardHolder = ListXML.GetTagText("ADDITIONALINDICATOR");
				
		<% /* Display Yes or No for if there is a cardholder or not */%>
		var sCardHolderField = "";
		if(sAdditionalCardHolder == "1")
		{
			sCardHolderField = "Yes";
		}
		if(sAdditionalCardHolder == "0")
		{
			sCardHolderField = "No";
		}

		<% /* Display the details in the appropriate table row */ %>
		var nRow = nLoop + 1;
		scScreenFunctions.SizeTextToField(tblBankCreditCardDetails.rows(nRow).cells(0), sName);
		scScreenFunctions.SizeTextToField(tblBankCreditCardDetails.rows(nRow).cells(1), sCardType);
		scScreenFunctions.SizeTextToField(tblBankCreditCardDetails.rows(nRow).cells(2), sProvider);
		scScreenFunctions.SizeTextToField(tblBankCreditCardDetails.rows(nRow).cells(3), sBalance);
		scScreenFunctions.SizeTextToField(tblBankCreditCardDetails.rows(nRow).cells(4), sRepayment);
		scScreenFunctions.SizeTextToField(tblBankCreditCardDetails.rows(nRow).cells(5), sCardHolderField);
	}
	if (nLoop + nStart > 0)
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
	//DRC MAR1189 - check if ADD button should be disabled
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bDisableAddButton = XML.GetGlobalParameterBoolean(document,"FSDisableBankCredCardAddButton");				
	if (bDisableAddButton)
	  frmScreen.btnAdd.disabled=true;
	else
	  frmScreen.btnAdd.disabled=false;  
	  
}

function frmScreen.btnEdit.onclick()
{
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("BANKCREDITCARD");

	<% /* Get the index of the selected row */ %>
	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();

	if(ListXML.SelectTagListItem(nRowSelected-1) == true)
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		scScreenFunctions.SetContextParameter(window,"idXML",ListXML.ActiveTag.xml);
		frmToDC081.submit();
	}
}

function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToDC081.submit();
}
		
function frmScreen.btnDelete.onclick()
{	
	var bAllowDelete = confirm("Are you sure?");
				
	if (bAllowDelete == true)
	{
		ListXML.ActiveTag = null;
		ListXML.CreateTagList("BANKCREDITCARD");

		<% /* Get the index of the selected row */ %>
		var nRowSelected = scTable.getRowSelectedIndex();
		ListXML.SelectTagListItem(nRowSelected-1);

		<% /* Set up the deletion XML 
		next line replaced by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077*/ %>
		//var XML = new scXMLFunctions.XMLObject();
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				
		var xmlRequest = XML.CreateRequestTag(window , "DELETE");
		XML.CreateActiveTag("BANKCREDITCARD");
		XML.CreateTag("CUSTOMERNUMBER", ListXML.GetTagText("CUSTOMERNUMBER"));
		XML.CreateTag("CUSTOMERVERSIONNUMBER", ListXML.GetTagText("CUSTOMERVERSIONNUMBER"));
		XML.CreateTag("SEQUENCENUMBER", ListXML.GetTagText("SEQUENCENUMBER"));				
		
		<% /* SR 09/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
			  node to CustomerFinancialBO. If the form accessed through completeness check, update FS always.
			  Else, do it only when all the records are deleted from the list box
		*/ %>
		if( scScreenFunctions.CompletenessCheckRouting(window) || scTable.getTotalRecords() == 1)
		{
			XML.ActiveTag = xmlRequest ;
			XML.CreateActiveTag("FINANCIALSUMMARY");
			XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
			if(scTable.getTotalRecords() == 1) XML.CreateTag("BANKCARDINDICATOR", 0)
			else XML.CreateTag("BANKCARDINDICATOR", 1) 
		}
		<% /* SR 09/06/2004 : BMIDS772 - End */ %>
		
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "DeleteBankCard.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}
				
		<% /* If the deletion is successful remove the entry from the list xml and the screen */ %>
		if(XML.IsResponseOK() == true)
		{
			PopulateScreen();
		}
		XML = null;					
	}				
}

function btnSubmit.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window)) frmToGN300.submit();
	else frmToDC085.submit();
}

-->
</script>
</body>
</html>



