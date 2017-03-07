<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc070.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Existing/Previous Mortgages
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		02/02/2000	ie5 changes 
AY		14/02/00	Change to msgButtons button types
AD		02/03/2000	Fixed SYS0115
MC		21/03/2000	Fixed SYS0429
MC		23/03/2000	Fixed SYS0542
AD		23/03/2000	Fixed SYS0237
AY		30/03/00	New top menu/scScreenFunctions change
MH		08/05/00	SYS0695 Doubleclick and cosmetic
IVW		09/06/00	SYS0707 Added Guarantors into the summary of mortgage accounts for this application.
BG		17/05/00	SYS0752 Removed Tooltips
MH      22/06/00    SYS0933 Readonly processing
CL		05/03/01	SYS1920 Read only functionality added
JLD		4/12/01		SYS2806 use scScreenFunctions CompletenessCheckRouting
JLD		22/01/02	SYS3851 don't output duplicate mortgage accounts.
DPF		20/06/02	BMIDS00077 - Changes made to bring file in line with V7.0.2 of Core, the change being
					implemented is...
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		15/05/2002	BMIDS00008	Modified btnsubmit.Onclick()
PSC		30/07/2002  BMIDS00006  Amend to use new Mortgage Account structure  
MV		22/08/2002	BMIDS00355	IE 5.5 Upgrade - Modified the Msgbuttons Position 
PSC		23/08/2002	BMIDS00354  Amend screen title
MDC		29/08/2002	BMIDS00336	Credit Check & Bureau Download
GHun	05/09/2002	BMIDS00406	Remove the word mortgage from the screen
GD		17/11/2002	BMIDS00376 Disable screen if readonly
MO		18/11/2002	BMIDS00723	Make changes for bigger fonts
LDM		12/05/2003	BM0492 Show total of loans not redeemed in the list box
BS		10/06/2003	BM0521 Do not enable Delete button when record selected if screen in in read mode
SR		01/06/2004	BMIDS772	Remove the question and cancel button. Change processing accordingly
SR		14/06/2004	BMIDS772	Using asp files instead of html for ScrollList functionality
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog    Date			AQR		Description
Maha T	17/11/2005		MAR174	Call IncomeCalcs when Submit button is clicked.
PE		25/01/2006		MAR1109	DC070 - Should be change to display the outstanding balance (Collateral Balance) from Mortgage Account
PSC		16/05/2006		MAR1798	Run Critical Data 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
*/%>

<span id="spnListScroll">
	<span style="LEFT: 310px; POSITION: absolute; TOP: 245px">
		<OBJECT data=scTableListScroll.asp id=scTable name=scScrollTable style="HEIGHT: 24px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<form id="frmToDC060" method="post" action="dc060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC071" method="post" action="dc071.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC073" method="post" action="dc073.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC085" method="post" action="dc085.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>
<% // span to keep tabbing within this screen %>
<span id="spnToLastField" tabindex="0"></span>
	
<form id="frmScreen" mark validate="onchange">
	<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 230px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
		<span id="spnExistingMortgages" style="TOP: 10px; LEFT: 4px; POSITION: ABSOLUTE">
			<table id="tblExistingMortgages" width="596px" border="0" cellspacing="0" cellpadding="0" class="msgTable">				
				<% /* BMIDS00336 MDC 29/08/2002 - Add Unassigned column 
					  BM0492 LDM 12/05/2003	Change Total Loans Not being Redeemed to Total Loans Not Redeemed*/ %>
				<tr id="rowTitles"><td width="25%" class="TableHead">Name</td><td width="15%" class="TableHead">Account Number</td><td width="10%" class="TableHead">Second<br>Charge?</td><td width="20%" class="TableHead">Lender</td><td width="20%" class="TableHead">Total Loans Not Redeemed</td><td width="10%" class="TableHead">Unassigned</td></tr>				
				<tr id="row01"><td width="25%" class="TableTopLeft">&nbsp</td><td width="15%" class="TableTopCenter">&nbsp</td><td width="10%" class="TableTopCenter">&nbsp</td><td width="20%" class="TableTopCenter">&nbsp</td><td width="20%" class="TableTopCenter">&nbsp</td><td width="10%" class="TableTopRight">&nbsp</td></tr>
				<tr id="row02"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
				<tr id="row03"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
				<tr id="row04"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
				<tr id="row05"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
				<tr id="row06"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
				<tr id="row07"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
				<tr id="row08"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
				<tr id="row09"><td width="25%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="20%" class="TableCenter">&nbsp</td><td width="10%" class="TableRight">&nbsp</td></tr>
				<tr id="row10"><td width="25%" class="TableBottomLeft">&nbsp</td><td width="15%" class="TableBottomCenter">&nbsp</td><td width="10%" class="TableBottomCenter">&nbsp</td><td width="20%" class="TableBottomCenter">&nbsp</td><td width="20%" class="TableBottomCenter">&nbsp</td><td width="10%" class="TableBottomRight">&nbsp</td></tr>
				<% /* BMIDS00336 MDC 29/08/2002 - End */ %>
			</table>
		</span>
		<span id="spnRemortgageIndicator" style="TOP: 182px; LEFT: 4px; POSITION: ABSOLUTE; VISIBILITY: hidden" class="msgLabel">
		* Account to be remortgaged
		</span>
		<span id="spnButtons" style="TOP: 200px; LEFT: 4px; POSITION: ABSOLUTE">
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
	
<!-- Main Buttons -->
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 320px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within the screen */%>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc070Attribs.asp" -->

<script language="JScript">
<!--
var ListXML;				
var m_sMetaAction					= null;
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
var m_sMortgageApplicationValue		= null;
var scScreenFunctions;
var m_sReadOnly = null;
var m_blnReadOnly = false;
var XMLArray = new Array();
var m_blnIsRemortgage = false;	<% /* BMIDS00444 */ %>
var m_iTableLength = 10;  <% /* SR 14/06/2004 : BMIDS772 */ %> 


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<% /* Make the required buttons available on the bottom of the screen
		(see msgButtons.asp for details) */ %>
	<%/* SR 01/06/2004 : BMIDS772 - remove Cancel button */ %>		
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);

<%	//next line replaced by line below as per V7.0.2 - DPF 20/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
%>	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	<% /* PSC 23/08/2002 BMIDS00354 */ %>
	FW030SetTitles("Existing Accounts","DC070",scScreenFunctions);

	RetrieveContextData();
			
	frmScreen.btnEdit.disabled = true;
	frmScreen.btnDelete.disabled = true;

	PopulateScreen();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
<%	//GD BMIDS00376 m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
%>	m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC070");
	if (m_sReadOnly=="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnAdd.disabled =true;
		frmScreen.btnDelete.disabled =true;
	}
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
<%	//GD BMIDS00376 m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC070");
%>	scScreenFunctions.SetContextParameter(window,"idAccountGuid",null);	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

<% /* keep the focus within	this screen when using the tab key */ %>
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<% /* keep the focus within	this screen when using the tab key */%>
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

<% /* Retrieves all context data required */ %>
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", null);			
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", "1627");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", "1");
	m_sMortgageApplicationValue = scScreenFunctions.GetContextParameter(window, "idMortgageApplicationValue", null);
}

<% /* Handles the onclick event from the span surrounding 
	the table. This is done here to handle the enabling 
	of buttons when a row is selected. Using the principle 
	of event bubbling we pick up the onclick event after 
	the table_onclick event in the scTable.asp file */%>
function spnExistingMortgages.onclick()
{
	if (scTable.getRowSelected() != -1) 
	{
		frmScreen.btnEdit.disabled = false;
		<% /* BS BM0521 10/06/03 Only enable Delete button if screen is in edit mode */ %>
		if (m_sReadOnly!="1")
		frmScreen.btnDelete.disabled = false;					
	}
}

function spnExistingMortgages.ondblclick()
{
	if (!frmScreen.btnEdit.disabled )
		frmScreen.btnEdit.onclick();
}

<% /* Retrieves the data and sets the screen accordingly */%>
function PopulateScreen()
{
	SetRemortgageIndicatorVisibility();
	//next line replace by line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//ListXML = new scXMLFunctions.XMLObject();
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	var TagSEARCH = ListXML.CreateRequestTag(window, "SEARCH")
									
	ListXML.CreateActiveTag("FINANCIALSUMMARY");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);

	ListXML.ActiveTag = TagSEARCH;
	var TagMORTGAGEACCOUNTLIST = ListXML.CreateActiveTag("MORTGAGEACCOUNTLIST");
			
	<% /* Loop through all customer context entries */ %>
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber			= scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber	= scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType		= scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		<% /* If the customer is an applicant, add him/her to the search */ %>
		if(sCustomerRoleType != "")
		{
			ListXML.CreateActiveTag("MORTGAGEACCOUNT");
			ListXML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			ListXML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			ListXML.ActiveTag = TagMORTGAGEACCOUNTLIST;
		}
	}
<% /*  BM0492 LDM 12/05/2003 Displaying mortgage account summary 
    Show outstanding balance where the redemption status is not redeemed (ie where != "already redeemed") */ %>
	ListXML.ActiveTag = TagSEARCH;
	ListXML.CreateTag("SHOWNOTREDEEMED","TRUE");
	
	ListXML.RunASP(document, "FindMortgageAccountSummary.asp");
			
	<%/* A record not found error is valid */%>
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);
			
	if (sResponseArray[0] == true || 
		sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		SetUpXMLArray();
		PopulateTable(0);				
	}
}
function SetUpXMLArray()
{
<% /* Only unique mortgage information to be displayed - based on accountguid - so
	  create an array linking the table row and the taglist index */
%>	ListXML.ActiveTag = null;
	ListXML.CreateTagList("MORTGAGEACCOUNT");
	var sGUID = "";
	for(var nIdx = 0; ListXML.SelectTagListItem(nIdx); nIdx++)
	{
		if(ListXML.GetTagText("ACCOUNTGUID") != sGUID)
		{
			sGUID = ListXML.GetTagText("ACCOUNTGUID");
			XMLArray[XMLArray.length] = nIdx;
		}
	}
}
<%	/* Displays a set of records in the table. This function is also used 
	by the scListScroll object. */%>
function PopulateTable(nStart)
{
	var iNumberOfRows ;		
	scTable.clear();	

	ListXML.ActiveTag = null;
	ListXML.CreateTagList("MORTGAGEACCOUNT");
	iNumberOfRows = ListXML.ActiveTagList.length ;
	
	scTable.initialiseTable(tblExistingMortgages, 0, "", ShowList, m_iTableLength, iNumberOfRows);
	ShowList(nStart);
}

function ShowList(nStart)
{			
	<% /* Populate the table with a set of records, starting with record nStart */ %>
	for (var nLoop=0; ListXML.SelectTagListItem(XMLArray[nLoop + nStart]) != false && nLoop < m_iTableLength; nLoop++)
	{								
		
		<% /* PSC 30/07/2002 BMIDS00006 - Start
		var sCustomerNumber = ListXML.GetTagText("CUSTOMERNUMBER");
		var sName			= scScreenFunctions.GetContextCustomerName(window,sCustomerNumber);
		*/ %>
		
		var xmlCustomerList = ListXML.ActiveTag.selectNodes("ACCOUNTRELATIONSHIPLIST/ACCOUNTRELATIONSHIP");
		var sName = "";
		var sForename = "";
		var sSurname = "";
	
		for (var nIndex=0; nIndex < xmlCustomerList.length; nIndex++)
		{
			if (nIndex > 0)
				if (nIndex == xmlCustomerList.length - 1)
					sName = sName + " & ";
				else
					sName = sName + ", ";
			
			sForename = xmlCustomerList.item(nIndex).selectSingleNode("FIRSTFORENAME").text;
			sSurname = xmlCustomerList.item(nIndex).selectSingleNode("SURNAME").text;
			
			sName = sName + sForename + " " + sSurname;			
		}
	
		var sAccountNumber	= ListXML.GetTagText("ACCOUNTNUMBER");
		var sSecondCharge	= ListXML.GetTagText("SECONDCHARGEINDICATOR");
		var sLenderName		= ListXML.GetTagText("COMPANYNAME");
		/* MAR1109 - Peter Edney - 25/01/2006 */		
		/*	var sTotalOS		= ListXML.GetTagText("OUTSTANDINGBALANCE"); */
		var sTotalOS		= ListXML.GetTagText("TOTALCOLLATERALBALANCE");
		var sAddress		= ListXML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER");
		var sUnassigned		= ListXML.GetTagText("UNASSIGNED");	<% /* MDC 29/08/2002 BMIDS00336 */ %>
		<% /* PSC 30/07/2002 BMIDS00006 - End */ %>
		
		<% /* BMIDS00444 Indicate which account is to be remortgaged */ %>
		if (m_blnIsRemortgage)
		{
			var sRemortgageIndicator = ListXML.GetTagText("REMORTGAGEINDICATOR");
			if (sRemortgageIndicator == "1")
				sAccountNumber = "*" + sAccountNumber;
		}
		<% /* BMIDS00444 End */ %>
		
		<% /* Display Yes or No for if there is a second charge or not */ %>
		var sSecondChargeField = "";
		if(sSecondCharge == "1")
		{
			sSecondChargeField = "Yes";
		}

		if(sSecondCharge == "0")
		{
			sSecondChargeField = "No";
		}

		<% /* MDC 29/08/2002 BMIDS00336 */ %>
		if(sUnassigned == "" || sUnassigned == "0")
			sUnassigned = "No";
		else
			sUnassigned = "Yes";
		<% /* MDC 29/08/2002 BMIDS00336 - End */ %>

		scScreenFunctions.SizeTextToField(tblExistingMortgages.rows(nLoop + 1).cells(0), sName);
		scScreenFunctions.SizeTextToField(tblExistingMortgages.rows(nLoop + 1).cells(1), sAccountNumber);
		scScreenFunctions.SizeTextToField(tblExistingMortgages.rows(nLoop + 1).cells(2), sSecondChargeField);
		scScreenFunctions.SizeTextToField(tblExistingMortgages.rows(nLoop + 1).cells(3), sLenderName);
		scScreenFunctions.SizeTextToField(tblExistingMortgages.rows(nLoop + 1).cells(4), sTotalOS);
		<% /* MDC 29/08/2002 BMIDS00336 */ %>
		scScreenFunctions.SizeTextToField(tblExistingMortgages.rows(nLoop + 1).cells(5), sUnassigned);
		<% /* MDC 29/08/2002 BMIDS00336 - End */ %>
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
}
	
function frmScreen.btnEdit.onclick()
{
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("MORTGAGEACCOUNT");
			
	<% /* Get the index of the selected row */ %>
	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();

	if(ListXML.SelectTagListItem(XMLArray[nRowSelected-1]) == true)
	{ 
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
		scScreenFunctions.SetContextParameter(window,"idAccountGuid",ListXML.GetTagText("ACCOUNTGUID"));
		frmToDC071.submit();  
	}
}

function frmScreen.btnAdd.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	frmToDC071.submit();
}
		
function frmScreen.btnDelete.onclick()
{	
	var bAllowDelete = confirm("Are you sure?");
				
	if (bAllowDelete == true)
	{
		ListXML.ActiveTag = null;
		ListXML.CreateTagList("MORTGAGEACCOUNT");

		<% /* Get the index of the selected row */ %>
		var nRowSelected = scTable.getRowSelectedIndex();

		ListXML.SelectTagListItem(XMLArray[nRowSelected-1]);

		<% /* Set up the deletion XML 
		next line replaced by line below as per Core 7.0.2 - DPF 20/06/02 - BMIDS00077 */ %>
		//var XML = new scXMLFunctions.XMLObject();
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				
		var xmlRequest = XML.CreateRequestTag(window, "DELETE");
		XML.CreateActiveTag("MORTGAGEACCOUNT");
		XML.CreateTag("CUSTOMERNUMBER",	ListXML.GetTagText("CUSTOMERNUMBER"));
		XML.CreateTag("CUSTOMERVERSIONNUMBER", ListXML.GetTagText("CUSTOMERVERSIONNUMBER"));
		XML.CreateTag("ACCOUNTGUID", ListXML.GetTagText("ACCOUNTGUID"));
		XML.CreateTag("THIRDPARTYGUID", ListXML.GetTagText("THIRDPARTYGUID"));
		XML.CreateTag("CUSTOMERADDRESSSEQUENCENUMBER", ListXML.GetTagText("CUSTOMERADDRESSSEQUENCENUMBER"));
		
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
			if(scTable.getTotalRecords() == 1) XML.CreateTag("EXISTINGMORTGAGEINDICATOR", 0)
			else XML.CreateTag("EXISTINGMORTGAGEINDICATOR", 1) 
		} 
		<% /* SR 09/06/2004 : BMIDS772 - End */ %>
		
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "DeleteMortgageAccount.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}
			
		if(XML.IsResponseOK())
		{					
			PopulateScreen();
		}
		XML = null;					
	}
}

function btnSubmit.onclick()
{
	ListXML = null;
	
	<% /* START: MAR174 - Maha T */ %>
	var bContinue = true;
	bContinue = RunIncomeCalculations();
	<% /* END: MAR174 */ %>
	
	<% /* MAR174 - Maha T */ %>
	if (bContinue)
	{
		<% /* PSC 09/08/2002 BMIDS00006 */ %>
		scScreenFunctions.SetContextParameter(window,"idAccountGuid",null);
			
		if(scScreenFunctions.CompletenessCheckRouting(window)) frmToGN300.submit();
		else frmToDC085.submit();
	}
}

<% /* START: MAR174 - Maha T */ %>
function RunIncomeCalculations()
{
    // PJO BMIDS621 Don't run income calcs in read only
  	if (m_sReadOnly =="1")
  		return(true) ;

	var AllowableIncXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AllowableIncXML.CreateRequestTag(window,null);
	<% /* MAR30 */ %> 
	AllowableIncXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));	
		
	// Set up request for Income Calculation
	var TagRequest = AllowableIncXML.CreateActiveTag("INCOMECALCULATION");
	AllowableIncXML.SetAttribute("CALCULATEMAXBORROWING","1");
	AllowableIncXML.SetAttribute("QUICKQUOTEMAXBORROWING","0");
	AllowableIncXML.CreateActiveTag("APPLICATION");
	AllowableIncXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AllowableIncXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	AllowableIncXML.ActiveTag = TagRequest;
	
	var TagCUSTOMERLIST = AllowableIncXML.CreateActiveTag("CUSTOMERLIST");

	for(var nLoop = 1; nLoop <= 2; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		<% /* BM0457 The customer must exist */ %>
		if (sCustomerNumber.trim().length > 0)
		{
		<% /* BM0457 End */ %>
			var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
			var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);
			var sCustomerOrder = scScreenFunctions.GetContextParameter(window,"idCustomerOrder" + nLoop);

			<% /* BM0457
			if(sCustomerRoleType == "1")
			{ */ %>
			AllowableIncXML.CreateActiveTag("CUSTOMER");
			AllowableIncXML.CreateTag("CUSTOMERNUMBER", sCustomerNumber);
			AllowableIncXML.CreateTag("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);
			AllowableIncXML.CreateTag("CUSTOMERROLETYPE", sCustomerRoleType);
			AllowableIncXML.CreateTag("CUSTOMERORDER", sCustomerOrder);
			AllowableIncXML.ActiveTag = TagCUSTOMERLIST;
		}
	}
	
		<% /* PSC 16/05/2006 MAR1798 - Start */ %>
	AllowableIncXML.SelectTag(null,"REQUEST");
	AllowableIncXML.SetAttribute("OPERATION","CriticalDataCheck");	
	AllowableIncXML.CreateActiveTag("CRITICALDATACONTEXT");
	AllowableIncXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	AllowableIncXML.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	AllowableIncXML.SetAttribute("SOURCEAPPLICATION","Omiga");
	AllowableIncXML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
	AllowableIncXML.SetAttribute("ACTIVITYINSTANCE","1");
	AllowableIncXML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
	AllowableIncXML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
	AllowableIncXML.SetAttribute("COMPONENT","omIC.IncomeCalcsBO");
	AllowableIncXML.SetAttribute("METHOD","RunIncomeCalculation");	

	window.status = "Critical Data Check - please wait";

	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
					AllowableIncXML.RunASP(document,"OmigaTMBO.asp");
			break;
		default: // Error
			AllowableIncXML.SetErrorResponse();
	}
	window.status = "";
	<% /* PSC 16/05/2006 MAR1798 - End */ %>

	AllowableIncXML.IsResponseOK()
	return(true);
	<% /* BMIDS01034 MDC 21/11/2002 - End */ %>
}
<% /* END: MAR174 */ %>

<% /* BMIDS00444 */ %>
function SetRemortgageIndicatorVisibility()
{
	<% /* Indicator should only be visible if TypeOfMortgage is "Remortgage" (validation type "R") */ %>
	var sMortgageType = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","");
	var ValidationList = new Array(1);
	ValidationList[0] = "R";
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_blnIsRemortgage = XML.IsInComboValidationList(document,"TypeOfMortgage",sMortgageType, ValidationList);
	
	if (m_blnIsRemortgage)
		spnRemortgageIndicator.style.visibility = "visible";
	else
		spnRemortgageIndicator.style.visibility = "hidden";
}
<% /* BMIDS00444 End */ %>
-->
</script>
</body>
</html>




<% /* OMIGA BUILD VERSION 045.03.04.28.00 */ %>
