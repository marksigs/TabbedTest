<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      ra030.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Credit Check Summary Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Descrip
JR		5/12/00		Screen Design
JR		14/12		included functionality to populate the list box
SR		18/12/00	included functionality to call popup screen for Voters Roll
SR		3/01/01		screen to be called as pop-up window
CL		05/03/01	SYS1920 Read only functionality added
SR		20/04/01	SYS2267 
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMids History:

Prog	Date		Descrip
MDC		29/08/2002	BMIDS00336 Credit Check and Full Bureau Download
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		09/10/2002	BMIDS00556	Removed Validate_Int()
MDC		22/11/2002	BMIDS00570  Pass all FB Records through to detail screen
GD		06/02/2003	BM0317		Display problems in ListBox - filter not being applied in ShowList()
GD		19/03/2003	BM0480		Display appropriate message if call to FBDownloadSummary.asp fails
MC		20/04/2004	BMIDS517	White space padded to the title text.
INR		10/06/2004	BMIDS744	ThirdPartyData changes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

*/
%>

<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"> 
	<title>Full Bureau Results Summary  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>

<span id="spnAddressSummaryListScroll">
	<% /* BMIDS00336 MDC 29/08/2002 
	<span style="LEFT: 300px; POSITION: absolute; TOP: 270px"> */ %>
	<span style="LEFT: 300px; POSITION: absolute; TOP: 300px">
	<% /* BMIDS00336 MDC 29/08/2002 - End */ %>
		<OBJECT data=scColorTableListScroll.asp id=scFBResultsSummaryTable style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
	</span> 
</span>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<!--
<form id="frmScreen" style="VISIBILITY: hidden" mark validate ="onchange">
-->
<form id="frmScreen" mark validate ="onchange">
<% /* BMIDS00336 MDC 29/08/2002 
<div id="divBackground" style="HEIGHT: 290px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px" class="msgGroup"> */ %>
<div id="divBackground" style="HEIGHT: 320px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px" class="msgGroup">
<% /* BMIDS00336 MDC 29/08/2002 - End */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		CCN Ref. No.
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtCreditCheckReferenceNumber" name="CreditCheckReferenceNumber" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
		Applicant Name
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<select id="cboApplicantName" style="WIDTH: 200px" class="msgCombo"></select>
		</span>
	</span>
	<% /* BMIDS00336 MDC 29/08/2002 */ %>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Voters Roll
		<span style="LEFT: 60px; POSITION: absolute; TOP: -3px">
			<input id="chkVotersRoll" type="checkbox" onclick="PopulateListBox()">
		</span>
	</span>
	<span style="LEFT: 104px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Public Info
		<span style="LEFT: 60px; POSITION: absolute; TOP: -3px">
			<input id="chkPublicInfo" type="checkbox" onclick="PopulateListBox()">
		</span>
	</span>
	<span style="LEFT: 204px; POSITION: absolute; TOP: 70px" class="msgLabel">
		CAIS
		<span style="LEFT: 30px; POSITION: absolute; TOP: -3px">
			<input id="chkCAIS" type="checkbox" value="1" onclick="PopulateListBox()">
		</span>
	</span>
	<span style="LEFT: 284px; POSITION: absolute; TOP: 70px" class="msgLabel">
		CIFAS
		<span style="LEFT: 30px; POSITION: absolute; TOP: -3px">
			<input id="chkCIFAS" type="checkbox" onclick="PopulateListBox()">
		</span>
	</span>
	<span style="LEFT: 354px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Prev Apps
		<span style="LEFT: 60px; POSITION: absolute; TOP: -3px">
			<input id="chkPrevApps" type="checkbox" onclick="PopulateListBox()">
		</span>
	</span>
	<span style="LEFT: 454px; POSITION: absolute; TOP: 70px" class="msgLabel">
		Alias/Assoc
		<span style="LEFT: 60px; POSITION: absolute; TOP: -3px">
			<input id="chkAlias" type="checkbox" onclick="PopulateListBox()">
		</span>
	</span>
	<% /* <span id="spnExperianBureauSummary" style="LEFT: 4px; POSITION: absolute; TOP: 70px"> */ %>
	<span id="spnExperianBureauSummary" style="LEFT: 4px; POSITION: absolute; TOP: 100px">
	<% /* BMIDS00336 MDC 29/08/2002 - End */ %>
	
	
		<table id="tblExperianBureauSummary" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	
				<td width="20%" class="TableHead">Name&nbsp;</td>	
				<td width="30%" class="TableHead">Address&nbsp;</td>	
				<td width="5%" class="TableHead">Sequence&nbsp;</td>	
				<td width="15%" class="TableHead">Type&nbsp;</td>
				<td width="15%" class="TableHead">Status&nbsp;</td>
				<td width="10%" class="TableHead">Current/ Worst&nbsp;</td>
				<td class="TableHead">Own Group Ind&nbsp;</td>
			</tr>
			<tr id="row01">		
				<td width="20%" class="TableTopLeft">	&nbsp;</td>		
				<td width="30%" class="TableTopCenter">&nbsp;</td>		
				<td width="5%" class="TableTopCenter">&nbsp;</td>		
				<td width="15%" class="TableTopCenter">&nbsp;</td>
				<td width="15%" class="TableTopCenter">&nbsp;</td>
				<td width="10%" class="TableTopCenter">&nbsp;</td>			
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">		
				<td width="20%" class="TableLeft">	&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td width="5%" class="TableCenter">&nbsp;</td>		
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">		
				<td width="20%" class="TableLeft">	&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td width="5%" class="TableCenter">&nbsp;</td>		
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">		
				<td width="20%" class="TableLeft">	&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td width="5%" class="TableCenter">&nbsp;</td>		
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row05">		
				<td width="20%" class="TableLeft">	&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td width="5%" class="TableCenter">&nbsp;</td>		
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row06">		
				<td width="20%" class="TableLeft">	&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td width="5%" class="TableCenter">&nbsp;</td>		
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row07">		
				<td width="20%" class="TableLeft">	&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td width="5%" class="TableCenter">&nbsp;</td>		
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row08">		
				<td width="20%" class="TableLeft">	&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td width="5%" class="TableCenter">&nbsp;</td>		
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row09">		
				<td width="20%" class="TableLeft">	&nbsp;</td>		
				<td width="30%" class="TableCenter">&nbsp;</td>		
				<td width="5%" class="TableCenter">&nbsp;</td>		
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="15%" class="TableCenter">&nbsp;</td>
				<td width="10%" class="TableCenter">&nbsp;</td>			
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row10">		
				<td width="20%" class="TableBottomLeft">	&nbsp;</td>		
				<td width="30%" class="TableBottomCenter">&nbsp;</td>		
				<td width="5%" class="TableBottomCenter">&nbsp;</td>		
				<td width="15%" class="TableBottomCenter">&nbsp;</td>
				<td width="15%" class="TableBottomCenter">&nbsp;</td>
				<td width="10%" class="TableBottomCenter">&nbsp;</td>		
				<td class="TableBottomRight">&nbsp;</td>
			</tr>
		</table>
	</span>	
</div>
</form>
<% /* BMIDS00336 MDC 29/08/2002 
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 340px; WIDTH: 612px"> */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 370px; WIDTH: 612px"> 
<% /* BMIDS00336 MDC 29/08/2002 - End */ %>
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/ra030Attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--

var scScreenFunctions;
var FBResultsXML = null;
var xmlresponseNode ;
var xmlDataHeaderKeys ;
var xmlCombos ;

var m_sCreditCheckGuid;
var m_sCustomerNumber1, m_sCustomerVersionNumber1 ;
var m_sCustomerNumber2, m_sCustomerVersionNumber2 ;
var m_sCustomerName1, m_sCustomerName2  ;
var m_sApplicationNumber, m_sApplicationFactFindNumber ;

var m_sCurrentCustomerFBResults ;

var iCurrentRow ;
var m_iTableLength = 10;

var m_aArgArray = null ;
var m_aRequestAttribs = null ;
var m_blnReadOnly = false;
var m_BaseNonPopupWindow = null;
<% /* BMIDS744 */ %>
var m_aBGColor = new Array(); 
var m_aFGColor = new Array(); 
<% /** EVENTS **/ %>


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{


	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();

	var sArguments		= window.dialogArguments ;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];	
	m_aArgArray			= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
		
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	RetrieveContextData();	
	Initialise();
	<%/* Fetch the combo values and the respective validations for the follwing combos ;
	 later used to populate value name in the form */
	%> 
	xmlCombos = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("SpecInst", 'CMLAddrFlag', 'StandardStatus');
	xmlCombos.GetComboLists(document,sGroupList);
	<% /* m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "RA030");  */ %>
	scScreenFunctions.ShowCollection(frmScreen);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function btnSubmit.onclick()
{
	iCurrentRow = scFBResultsSummaryTable.getRowSelected();
	
	if(iCurrentRow == -1)
	{
		window.alert("Select a row in the list") ;
		return ;
	}
	
	var sSelectedBlockId = tblExperianBureauSummary.rows(iCurrentRow).getAttribute("FBBlockId");
	var sSelectedHeaderSeq = tblExperianBureauSummary.rows(iCurrentRow).cells(2).innerText; 
	ShowBlockSummary(sSelectedBlockId, sSelectedHeaderSeq);
}

function btnCancel.onclick()
{
	window.close();
}

function frmScreen.cboApplicantName.onchange()
{
	scFBResultsSummaryTable.clear();
	PopulateListBox();
} 

function spnExperianBureauSummary.ondblclick()
{
	<%//BM0480 - Check for row selected being -1, too %>
	if ((scFBResultsSummaryTable.getRowSelected() != null) && (scFBResultsSummaryTable.getRowSelected() != -1)) btnSubmit.onclick();
}


<% /** FUNCTIONS **/  %>

function ShowBlockSummary(sSelectedBlockId, sSelectedHeaderSeq)
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	var ArrayArguments = new Array();
	<%/* All the elements added to ArrayArguments is common to all pop-ups */ %>
	ArrayArguments[0] = m_sCreditCheckGuid ;
	ArrayArguments[1] = (frmScreen.cboApplicantName.value == "1")? m_sCustomerNumber1 : m_sCustomerNumber2 ;
	ArrayArguments[2] = (frmScreen.cboApplicantName.value == "1")? m_sCustomerVersionNumber1 : m_sCustomerVersionNumber2 ;
	ArrayArguments[3] = frmScreen.cboApplicantName.value ;
	ArrayArguments[4] = sSelectedBlockId ;
	ArrayArguments[5] = sSelectedHeaderSeq ;
	ArrayArguments[6] =	frmScreen.txtCreditCheckReferenceNumber.value	// Credit Check reference number 
	
	ArrayArguments[7] = m_sCurrentCustomerFBResults ; // Pass the results corresponding to the Customer selected
	ArrayArguments[8] = xmlDataHeaderKeys.XMLDocument.xml ;
	
	<% /* ArrayArguments[9] = XML.CreateRequestAttributeArray(window); */ %>
	ArrayArguments[9] = m_aRequestAttribs;
	var saPopupData = new Array()
		
	switch (sSelectedBlockId.substr(0,2))
	{
		case "BE":
			<% /* Add popup-specific data to be passed as an array to ArrayArguments */ %>
			scScreenFunctions.DisplayPopup(window, document, "ra031.asp", ArrayArguments, 478, 540);
			break;
		case "BJ":
			scScreenFunctions.DisplayPopup(window, document, "ra032.asp", ArrayArguments, 460, 590);
			break;
		case "BF":
			scScreenFunctions.DisplayPopup(window, document, "ra035.asp", ArrayArguments, 460, 620);
			break;
		case "BC":
			scScreenFunctions.DisplayPopup(window, document, "RA034.asp", ArrayArguments, 525, 590);
			break;
		case "BA":
			scScreenFunctions.DisplayPopup(window, document, "RA036.asp", ArrayArguments, 480, 630);
			break;
		case "BL":
			scScreenFunctions.DisplayPopup(window, document, "RA037.asp", ArrayArguments, 477, 580);		
			break ;
		default:	
			N/A  <%/* error */%>
	}
}

function RetrieveContextData()
{
<%
	/* TEST	*/
	/*
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber","00073261");
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber","1");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber2","00073962");
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber2","1");
	
	scScreenFunctions.SetContextParameter(window,"idCustomerName1","Customer Name 1");
	scScreenFunctions.SetContextParameter(window,"idCustomerName2","Customer Name 2");
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","00018856");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
	
	scScreenFunctions.SetContextParameter(window,"idUserId","USER");
	scScreenFunctions.SetContextParameter(window,"idUnitId", "UNIT1");
	scScreenFunctions.SetContextParameter(window,"idMachineId", "MACH1");
	scScreenFunctions.SetContextParameter(window,"idDistributionChannelId", "CHAN1");
	*/
	
	/*
	m_sCustomerNumber1 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber","");
	m_sCustomerVersionNumber1 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber","");
	m_sCustomerNumber2 = scScreenFunctions.GetContextParameter(window,"idCustomerNumber2","");
	m_sCustomerVersionNumber1 = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber2","");
	m_sCustomerName1 = scScreenFunctions.GetContextParameter(window,"idCustomerName1","");
	m_sCustomerName2 = scScreenFunctions.GetContextParameter(window,"idCustomerName2","");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = 	scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	*/
	/* END TEST */
%>	
	m_sApplicationNumber			= m_aArgArray[0];
	m_sApplicationFactFindNumber	= m_aArgArray[1];
	m_sCustomerNumber1				= m_aArgArray[2];
	m_sCustomerVersionNumber1		= m_aArgArray[3];
	m_sCustomerNumber2				= m_aArgArray[4];
	m_sCustomerVersionNumber2		= m_aArgArray[5];
	m_sCustomerName1				= m_aArgArray[6];
	m_sCustomerName2				= m_aArgArray[7];
	m_aRequestAttribs				= m_aArgArray[10];
}

function Initialise()
{
	PopulateApplicantNames();
	frmScreen.cboApplicantName.value = 1 ;
	
	<% /* BMIDS00336 MDC 02/09/2002 */ %>
	frmScreen.chkAlias.checked = true;
	frmScreen.chkCAIS.checked = true;
	frmScreen.chkCIFAS.checked = true;
	frmScreen.chkPrevApps.checked = true;
	frmScreen.chkPublicInfo.checked = true;
	frmScreen.chkVotersRoll.checked = true;
	<% /* BMIDS00336 MDC 02/09/2002 - End */ %>
	
	<%//BM0480 don't access FBResultsXML is it doesn't contain useable data %>
	if (PopulateListBox() != false)
	{
		<% /* BMIDS00570 MDC 22/11/2002 - Pass all FB Records through to detail screen */ %>
		var TagFBResultsList = FBResultsXML.SelectTag(null,"FBRESULTSLIST");
		m_sCurrentCustomerFBResults = TagFBResultsList.xml;
		<% /* BMIDS00570 MDC 22/11/2002 - End */ %>
	
		FBResultsXML.ActiveTag = xmlresponseNode ;
		frmScreen.txtCreditCheckReferenceNumber.value = FBResultsXML.GetTagText("CREDITCHECKREFERENCENUMBER");
	
		GenerateXmlDataHeaderKeys();
	}
}

function PopulateApplicantNames()
{
	var TagOPTION;
	
	TagOPTION = document.createElement("OPTION");
	TagOPTION.value	= "0";
	TagOPTION.text	= "<SELECT>";
	frmScreen.cboApplicantName.add(TagOPTION);
	
	TagOPTION = document.createElement("OPTION");
	TagOPTION.value	= "1";
	TagOPTION.text	= m_sCustomerName1;
	frmScreen.cboApplicantName.add(TagOPTION);
	
	if(m_sCustomerName2 != '')
	{
		TagOPTION = document.createElement("OPTION");
		TagOPTION.value	= "2";
		TagOPTION.text	= m_sCustomerName2;
		frmScreen.cboApplicantName.add(TagOPTION);
	}
}

function PopulateListBox()
{
	<% /* BMIDS744 */ %>
	var iCount;
	var bOptOutIndicator;
	
	<% /* Do not call the method, if it was already called, i.e., FBResultsXML is null	 */ %>

	if(FBResultsXML == null)
	{
		<% /* Set up the passing XML to BO */ %>
		FBResultsXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
		FBResultsXML.CreateRequestTagFromArray(m_aRequestAttribs, "SEARCH");
		
		FBResultsXML.CreateActiveTag("APPLICATIONCREDITCHECK");
		FBResultsXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		FBResultsXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	
		<% /* Pass XML to BO  */ %>
		
		FBResultsXML.RunASP(document,"FBDownloadSummary.asp");
		<% //BM0480 Changes to trap RNF error %>
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = FBResultsXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[0] == false) <%//If there is some sort of error%>
		{
			<%//Disable OK button - Set Screen to readonly %>
			DisableMainButton("Submit");
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			if (ErrorReturn[1] == ErrorTypes[0]) <%//Record Not Found%>
			{
				alert('Full Bureau Download returned no results');
			}
			<%// BM0480 Return false from the method if FBResultsXML contains an usable XML %>
			return(false);
		}
	
		<% /* Process the returning XML  */ %>
		xmlresponseNode = FBResultsXML.ActiveTag ;
		FBResultsXML.ActiveTag = null;
		<% /* BMIDS744  */ %>
		FBResultsXML.SelectTag(null, "RESPONSE");
		bOptOutIndicator = FBResultsXML.GetTagText("OPTOUTINDICATOR");

		FBResultsXML.ActiveTag = null;
		FBResultsXML.CreateTagList("FBRESULTS");
		FBResultsXML.SelectTag(null, "FBRESULTS");
	
		<% /* Setup the scrolling for list box */ %>
		m_sCreditCheckGuid = FBResultsXML.GetTagText("CREDITCHECKGUID");
	}
	<%//GD BM0317 START%>
	if (ApplyFilters() == false)
	{
		return;
	}
	<%//GD BM0317 END%>
	
	<% /*BMIDS744 Display Colour is dependant on the OptOutIndicator
	and the bureau Ref Category */ %>
	
	var iNumberOfBlocks = FBResultsXML.ActiveTagList.length;
	for (iCount=0; iCount < iNumberOfBlocks; iCount++)
	{
		FBResultsXML.SelectTagListItem(iCount);
		sFBBureauCategory = FBResultsXML.GetTagText ("FBBUREAUREFCATEGORY");

		if (bOptOutIndicator == false)
		{
			switch (sFBBureauCategory)
			{
				case "1":
				case "2":
				case "5":
					m_aFGColor[iCount + 1] = 'black';
					break;
				case "3":
				case "4":
					m_aFGColor[iCount + 1] = 'green';
					break;
				default:
					<%/* Anything arriving which is unexpected, make it green */%>	
					m_aFGColor[iCount + 1] = 'green';				
			}
		}
		else
		{	
			switch (sFBBureauCategory)
			{
				case "1":
				case "5":
					m_aFGColor[iCount + 1] = 'black';
					break;
				case "2":
					m_aFGColor[iCount + 1] = 'blue';
					break;
				case "3":
				case "4":
					m_aFGColor[iCount + 1] = 'green';
					break;
				default:
					<%/* Anything arriving which is unexpected, make it green */%>	
					m_aFGColor[iCount + 1] = 'green';				
			}
		}
	}
	scFBResultsSummaryTable.initialiseTable(tblExperianBureauSummary, 0, "", ShowList, m_iTableLength, iNumberOfBlocks, m_aFGColor);
	scFBResultsSummaryTable.clear();
	<% /* BMIDS00336 MDC 05/09/2002 - End */ %>

	if(iNumberOfBlocks > 0)
		ShowList(0);
	
}

function ShowList(nStart)
{
	var iCount;
	var sTitle, sFirstForename, sSurname;
	var sName;
	var sAddress;
	var sFlatNumber, sHouseName, sHouseNumber, sStreet, sDistrict, sTown, sCounty, sPostCode;
	var sSequence;
	var sBlockId, sBlockType;
	var sStatus;
	var sAccountStatus;
	var sOwnGrpId;
	var sFBNameIndicator, sFBAssociationInfoType ;
		
	var iRowIndex = 0 ;
	<% /* BMIDS00570 MDC 22/11/2002
	m_sCurrentCustomerFBResults = "" ; */ %>
	
	var iCurrentCustomer = frmScreen.cboApplicantName.value ;
	<%//GD BM0317 START %>
	if (ApplyFilters() == false)
	{
		return;
	}
	<%//GD BM0317 END	%>
	for (iCount=0; iCount < FBResultsXML.ActiveTagList.length && iRowIndex < m_iTableLength; iCount++)
	{ 

		
		FBResultsXML.SelectTagListItem(iCount);
		sFBNameIndicator = FBResultsXML.GetTagText ("FBNAMEINDICATOR");
		
		FBResultsXML.SelectTagListItem(iCount + nStart);

			
		if(((sFBNameIndicator == "1" || sFBNameIndicator == "") && iCurrentCustomer ==1) 
		    || (sFBNameIndicator == "2" && iCurrentCustomer == 2))
		{	
			sAddress = "";
		
			<% /* Extract data from the XML */ %>
			sTitle = FBResultsXML.GetTagText ("FBTITLE");
			sFirstForename = FBResultsXML.GetTagText ("FBFORENAME");
			sSurname = FBResultsXML.GetTagText ("FBSURNAME");
			sFlatNumber = FBResultsXML.GetTagText ("FBFLAT");
			sHouseName = FBResultsXML.GetTagText ("FBHOUSENAME");
			sHouseNumber = FBResultsXML.GetTagText ("FBHOUSENUMBER");
			sStreet = FBResultsXML.GetTagText ("FBSTREET");
			sDistrict = FBResultsXML.GetTagText ("FBDISTRICT");
			sTown = FBResultsXML.GetTagText ("FBTOWN");
			sCounty = FBResultsXML.GetTagText ("FBCOUNTY");
			sPostCode = FBResultsXML.GetTagText ("FBPOSTCODE");
			sSequence = FBResultsXML.GetTagText ("FBHEADERSEQUENCE");
			sBlockId = FBResultsXML.GetTagText ("FBBLOCKID");
			
			sStatus = "" ;
			sAccountStatus = "" ;
			sOwnGrpId = "" ;
			
			if(sBlockId == "BC01") //CAIS block
			{
				if(FBResultsXML.GetTagText("FBCAISOWNDATAFLAG") == 'Y') sOwnGrpId = '*' ;
				
				sAccountStatus = GetCurrentWorstStatus(FBResultsXML.GetTagText("FBCAISACCOUNTSTATUS"))
			}
			else
			{
				if(sBlockId == "BA01") // CAPS block
					if(FBResultsXML.GetTagText("FBCAPSOWNAPPLICATIONIND") != '') sOwnGrpId = '*' ;
			}
						
			switch (sBlockId.substr(0,2))
			{
				case "BE":
					sBlockType = "Voters Roll";
					break;
				case "BJ":
					sBlockType = "Public Info";
					break;
				case "BF":
					sBlockType = "CIFAS";
					break;
				case "BC":
					sBlockType = "CAIS";
					break;
				case "BA":
					sBlockType = "Prev App";
					break;
				case "BL":
					sBlockType = "Alias/Assoc";
					break ;
				default:	
					sBlockType = sBlockId ;				
			}
			
			if(sFlatNumber != "") sAddress += sFlatNumber + ",";
			if(sHouseName != "") sAddress += sHouseName + ",";
			if(sHouseNumber != "") sAddress += sHouseNumber + ",";
			if(sStreet != "") sAddress += sStreet + ",";
			if(sDistrict != "") sAddress += sDistrict + ",";
			if(sTown != "") sAddress += sTown + ",";
			if(sCounty != "") sAddress += sCounty + ",";
			if(sPostCode != "") sAddress += sPostCode + ",";
				
			sName = sTitle + " " + sFirstForename + " " + sSurname;
				
			<% /* populate the list box */ %>
			scScreenFunctions.SizeTextToField(tblExperianBureauSummary.rows(iRowIndex+1).cells(0),sName);
			scScreenFunctions.SizeTextToField(tblExperianBureauSummary.rows(iRowIndex+1).cells(1),sAddress);
			scScreenFunctions.SizeTextToField(tblExperianBureauSummary.rows(iRowIndex+1).cells(2),sSequence);
			scScreenFunctions.SizeTextToField(tblExperianBureauSummary.rows(iRowIndex+1).cells(3),sBlockType);
			scScreenFunctions.SizeTextToField(tblExperianBureauSummary.rows(iRowIndex+1).cells(4),sStatus);
			scScreenFunctions.SizeTextToField(tblExperianBureauSummary.rows(iRowIndex+1).cells(5),sAccountStatus);
			scScreenFunctions.SizeTextToField(tblExperianBureauSummary.rows(iRowIndex+1).cells(6),sOwnGrpId);
			
			tblExperianBureauSummary.rows(iRowIndex+1).setAttribute("FBBlockId", sBlockId);
			tblExperianBureauSummary.rows(iRowIndex+1).setAttribute("FBNameIndicator", sFBNameIndicator);
			tblExperianBureauSummary.rows(iRowIndex+1).setAttribute("FBAssociationInfoType", sFBAssociationInfoType);


			<% /* BMIDS00570 MDC 22/11/2002 - Pass all FB Records through to detail screen
			m_sCurrentCustomerFBResults += FBResultsXML.ActiveTag.xml ; */ %>					
			++iRowIndex;			
		}
	}
	<% /* BMIDS00570 MDC 22/11/2002 - Pass all FB Records through to detail screen 
	m_sCurrentCustomerFBResults = "<FBRESULTSLIST>" + m_sCurrentCustomerFBResults + "</FBRESULTSLIST>" ;  */ %>
}

function GenerateXmlDataHeaderKeys()
{	<% /* Generate the XML with KEY values for all the rows (blocks) to be populated in the ListBox */ %>

	FBResultsXML.ActiveTag = null;
	FBResultsXML.CreateTagList("FBRESULTS");

	xmlDataHeaderKeys = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var xmlListNode = xmlDataHeaderKeys.CreateActiveTag("FULLBUREAUDATAHEADERLIST")
	
	for (var iCount=0; iCount < FBResultsXML.ActiveTagList.length; iCount++)	
	{
		xmlDataHeaderKeys.CreateActiveTag("FULLBUREAUDATAHEADER")
		FBResultsXML.SelectTagListItem(iCount);
		
		xmlDataHeaderKeys.CreateTag("CREDITCHECKGUID", FBResultsXML.GetTagText("CREDITCHECKGUID"));
		xmlDataHeaderKeys.CreateTag("FBBLOCKID", FBResultsXML.GetTagText("FBBLOCKID"));
		xmlDataHeaderKeys.CreateTag("FBHEADERSEQUENCE", FBResultsXML.GetTagText("FBHEADERSEQUENCE"));
		xmlDataHeaderKeys.CreateTag("FBADDRESSINDICATOR", FBResultsXML.GetTagText("FBADDRESSINDICATOR"));
		
		xmlDataHeaderKeys.ActiveTag = xmlListNode ;
	}
}

function GetCurrentWorstStatus(sAccountStatus)
{
	var sNums = "[0123456789]"
	var sCurrentStatus, sWorstStatus ;
	var sTemp ;
	
	sCurrentStatus = sAccountStatus.charAt(0);
	
	sWorstStatus = '-1' ;
	
	<% /* if numerics are found in the string from the character 2 to 12, assign the highest one to worst status */ %>
	for(var nLoop=1; nLoop<sAccountStatus.length && nLoop < 12; nLoop++)
	{
		sTemp = sAccountStatus.charAt(nLoop)
		if(sTemp.search(sNums) != -1)
			if(parseInt(sTemp) > parseInt(sWorstStatus)) sWorstStatus = sTemp ;			
	}
	
	<% /* if numericas are not found */ %>
	if(sWorstStatus == '-1')
	{
		sWorstStatus = '';
		for(var nLoop=1; nLoop<sAccountStatus.length && nLoop < 12; nLoop++)
		{
			sTemp = sAccountStatus.charAt(nLoop)
			
			if(sTemp == 'U')
			{
				sWorstStatus = 'U'
				break;
			}
			else if(sTemp == 'S')sWorstStatus = 'S'				
		}
	}
	
	return sCurrentStatus + '/' + sWorstStatus ;
}

function GetStatus()
{
	var sStatus ;
	var sCMLAddrFlag, sStandardStatus, sAccountStatus;
		
	var sSplInstFlag = FBResultsXML.GetTagText("FBCAISSPECIALINSTRFLAG") ;

	if(sSplInstFlag != '')
	{
		sStatus =  xmlCombos.GetComboDescriptionForValidation('SpecInst', sSplInstFlag);
		
		if(sStatus != '') return sStatus;	
	}
	
	sCMLAddrFlag = FBResultsXML.GetTagText("FBCAISCMLADDRESSTYPE") ;
	
	if(sCMLAddrFlag != '')
	{
		sStatus = xmlCombos.GetComboDescriptionForValidation('CMLAddrFlag', sCMLAddrFlag);
		if(sStatus != '') return sStatus ;
	}
	
	sStandardStatus = '';
	
	if(FBResultsXML.GetTagText("FBCAISSETTLEDDATE") != '')
		sStandardStatus = '1' ;
	else
	{
		sAccountStatus = FBResultsXML.GetTagText("FBCAISACCOUNTSTATUS") ;
		
		var sNums = "[89]"
		var sTemp ;
		
		for(var nLoop=1; nLoop<sAccountStatus.length; nLoop++)
		{
			sTemp = sAccountStatus.charAt(nLoop) ;
			if(sTemp == '8' || sTemp == '9')
			{
				sStandardStatus = '2' ;
				break ;
			}
		}
		
		if(sStandardStatus == '') sStandardStatus = '3';
	}
	
	sStatus =  xmlCombos.GetComboDescriptionForValidation('StandardStatus', sStandardStatus);
	return sStatus ;
}
 
<%// GD BM0317 START%>

function ApplyFilters()
{
	var sFbNameIndicator ;
	sFBNameIndicator = frmScreen.cboApplicantName.value  ;
	
	<% /* Find out the number of rows to be displayed in the list box (for this customer) */ %>
	var sCondition ;
	sCondition = ".//FBRESULTS[(FBNAMEINDICATOR='" + sFBNameIndicator + "'"
	
	<% /* if there is only one applicant, FBNameIndicator might not have been mentioned in the table */ %>
	if(sFBNameIndicator ==1) sCondition += " || FBNAMEINDICATOR=''" ;
	
	<% /* BMIDS00336 MDC 02/09/2002 - Additional filtering */ %>
	sCondition += ")"
	var sSubCondition = "";

	if(frmScreen.chkAlias.checked == false && frmScreen.chkCAIS.checked == false && frmScreen.chkCIFAS.checked == false &&
		frmScreen.chkPrevApps.checked == false && frmScreen.chkPublicInfo.checked == false && frmScreen.chkVotersRoll.checked == false)
	{
		alert("You must select one option.");
		return(false);
	}
	
	if(frmScreen.chkAlias.checked == false || frmScreen.chkCAIS.checked == false || frmScreen.chkCIFAS.checked == false ||
		frmScreen.chkPrevApps.checked == false || frmScreen.chkPublicInfo.checked == false || frmScreen.chkVotersRoll.checked == false)
	{
		if(frmScreen.chkAlias.checked)
			sSubCondition = " || FBBLOCKID = 'BL01'";

		if(frmScreen.chkCAIS.checked)
			sSubCondition += " || FBBLOCKID = 'BC01'";

		if(frmScreen.chkCIFAS.checked)
			sSubCondition += " || FBBLOCKID = 'BF01'";

		if(frmScreen.chkPrevApps.checked)
			sSubCondition += " || FBBLOCKID = 'BA01'";

		if(frmScreen.chkPublicInfo.checked)
			sSubCondition += " || FBBLOCKID = 'BJ01'";

		if(frmScreen.chkVotersRoll.checked)
			sSubCondition += " || FBBLOCKID = 'BE02'";
			
		sCondition += " && (" + sSubCondition.substr(4, sSubCondition.length - 4) + ")";
	}		
	<% /* BMIDS00336 MDC 02/09/2002 - End */ %>
	
	sCondition += "]" ;
	
	<% /* BMIDS00336 MDC 05/09/2002 */ %>
	// var iNumberOfBlocks = FBResultsXML.XMLDocument.selectNodes(sCondition).length;
	FBResultsXML.ActiveTag = null;
	FBResultsXML.SelectNodes(sCondition).length;
	return(true)
}

<%//GD BM0317 END%>

-->
</script>
</BODY>
</HTML>

