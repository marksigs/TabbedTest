<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      dc100.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Mortgage Related Life Products
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
APS		04/02/2000	ie5 changes 
AY		14/02/00	Change to msgButtons button types
AD		02/03/2000	Fixed SYS0263.
IW		02/05/00	SYS0137 (2 & 4) .. 
MC		21/03/2000	Fixed SYS0433.
AY		30/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
MH      23/06/00    SYS0933 Readonly stuff
CL		05/03/01	SYS1920 Read only functionality added
SA		06/06/01	SYS1672 Collect details for guarantors as well
JLD		4/12/01		SYS2806 Completeness check routing
DPF		20/06/02	BMIDS00077 Changes made to file to bring in line with
					Core V7.0.2.  Changes are...
					SYS1254 Add scroll bars
					SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		17/05/2002	BMIDS00008	Modified Routing Screen on btnSubmit.Onclick() , btnCancel.onclick()
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
<script src="Validation.js" language=""></script>
<% /* removed as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
<object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
*/ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
<span style="TOP: 275px; LEFT: 310px; POSITION: absolute">
	<object data="scListScroll.htm" id="scScrollPlus" style="LEFT: 0px; TOP: 0px; height:24; width:304" type=text/x-scriptlet VIEWASTEXT tabindex="-1"></object>
</span> 

<form id="frmToDC101" method="post" action="dc101.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC120" method="post" action="dc120.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC080" method="post" action="dc080.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 245px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<div style="TOP: 10px; LEFT: 4px; HEIGHT: 23px; WIDTH: 335px; POSITION: ABSOLUTE" class="msgGroupLight">
		<span style="TOP: 4px; LEFT: 4px; WIDTH: 400px; POSITION: ABSOLUTE" class="msgLabel">
		Do you have any mortgage related life policies?
			<span style="TOP: -3px; LEFT: 240px; POSITION: ABSOLUTE">
				<input id="optLifeProductYes" name="LifeProductInd" type="radio" value="1">
				<label for="optLifeProductYes" class="msgLabel">Yes</label>
			</span> 
			<span style="TOP: -3px; LEFT: 290px; POSITION: ABSOLUTE">
				<input id="optLifeProductNo" name="LifeProductInd" type="radio" value="0">
				<label for="optLifeProductNo" class="msgLabel">No</label>
			</span> 
		</span>		
	</div>
	<span id="spnLifeProducts" style="TOP: 40px; LEFT: 4px; POSITION: ABSOLUTE">
		<table id="tblLifeProducts" width="596px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
			<tr id="rowTitles"><td width="20%" class="TableHead">Name</td><td width="15%" class="TableHead">Organisation</td><td width="15%" class="TableHead">Policy<br>Type</td><td width="10%" class="TableHead">Monthly<br>Premium</td><td width="10%" class="TableHead">Year of<br>Maturity</td><td width="15%" class="TableHead">Min. Death<br>Benefit</td>	<td class="TableHead">Maturity<br>Value</td></tr>
			<tr id="row01"><td width="20%" class="TableTopLeft">&nbsp</td><td width="15%" class="TableTopCenter">&nbsp</td><td width="15%" class="TableTopCenter">&nbsp</td><td width="10%" class="TableTopCenter">&nbsp</td><td width="10%" class="TableTopCenter">&nbsp</td><td width="15%" class="TableTopCenter">&nbsp</td><td class="TableTopRight">&nbsp</td></tr>
			<tr id="row02"><td width="20%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td class="TableRight">&nbsp</td></tr>
			<tr id="row03"><td width="20%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td class="TableRight">&nbsp</td></tr>
			<tr id="row04"><td width="20%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td class="TableRight">&nbsp</td></tr>
			<tr id="row05"><td width="20%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td class="TableRight">&nbsp</td></tr>
			<tr id="row06"><td width="20%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td class="TableRight">&nbsp</td></tr>
			<tr id="row07"><td width="20%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td class="TableRight">&nbsp</td></tr>
			<tr id="row08"><td width="20%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td class="TableRight">&nbsp</td></tr>
			<tr id="row09"><td width="20%" class="TableLeft">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="10%" class="TableCenter">&nbsp</td><td width="15%" class="TableCenter">&nbsp</td><td class="TableRight">&nbsp</td></tr>
			<tr id="row10"><td width="20%" class="TableBottomLeft">&nbsp</td><td width="15%" class="TableBottomCenter">&nbsp</td><td width="15%" class="TableBottomCenter">&nbsp</td><td width="10%" class="TableBottomCenter">&nbsp</td><td width="10%" class="TableBottomCenter">&nbsp</td><td width="15%" class="TableBottomCenter">&nbsp</td><td class="TableBottomRight">&nbsp</td></tr>
		</table>
	</span>

	<span id="spnButtons" style="TOP: 215px; LEFT: 4px; POSITION: ABSOLUTE">
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
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/dc100Attribs.asp" -->

<script language="JScript">
<!--
var ListXML;
var m_sMetaAction = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
<% /* Radio button value on entry */ %>
var m_sQuestionOnEntry	= null;
<% /* Is there a FinancialSummary record */%>
var m_bIsFinancialSummary = false;
var scScreenFunctions;
var m_sReadOnly="";
var m_blnReadOnly = false;



<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<% /* Make the required buttons available on the bottom of the screen
		(see msgButtons.asp for details) */ %>
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	//next line replaced with line below as per Core V7.0.2 - DPF 20/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	FW030SetTitles("Mortgage Related Life Products","DC100",scScreenFunctions);

	<% /* Initialise the table */ %>
	scTable.initialise(tblLifeProducts, 0, "");

	RetrieveContextData();
			
	<% /* Default Add to disabled. Only enable if question is answered Yes 
		and Default Edit/Delete to disabled*/ %>
	frmScreen.btnAdd.disabled = true;
	frmScreen.btnEdit.disabled = true;
	frmScreen.btnDelete.disabled = true;

	PopulateScreen();
	DoRadioButtonChecks();
	frmScreen.optLifeProductYes.onclick();			
			
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();

	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");

	if (m_sReadOnly=="1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnAdd.disabled =true;
		frmScreen.btnDelete.disabled =true;
	}
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC100");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

<% /* keep the focus within this screen when using the tab key */ %>
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

<% /* keep the focus within this screen when using the tab key */ %>
function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}

<% /* Retrieves all context data required for use within this screen.*/ %>
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
}

<% /* Disables the Add button if the No Radio button is set. */ %>
function frmScreen.optLifeProductNo.onclick()
{
	if(frmScreen.optLifeProductNo.checked == true)
	{
		frmScreen.btnAdd.disabled = true;
	}
}

<% /* Enables the Add button if the Yes Radio button is set. */ %>
function frmScreen.optLifeProductYes.onclick()
{
	if(frmScreen.optLifeProductYes.checked == true)
	{
		frmScreen.btnAdd.disabled = false;
	}
}

<% /* Handles the onclick event from the span surrounding the table. This is 
	done here to handle the enabling of buttons when a row is selected. 
	Using the principle of event bubbling we pick up the onclick event after 
	the table_onclick event in the scTable.htm scriptlet */ %>
function spnLifeProducts.onclick()
{
	if (scTable.getRowSelectedId() != null) 
	{
		frmScreen.btnEdit.disabled = false;
<%				
//				if (m_sReadOnly != "1")
//				{
%>
			frmScreen.btnDelete.disabled = false;					
<%
//				}
%>
	}
}

<% /* Retrieves the data and sets the screen accordingly */ %>
function PopulateScreen()
{
	//next line removed as per Core V7.0.2
	//ListXML = new scXMLFunctions.XMLObject();
	ListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	<% /* <REQUEST USERID=? USERTYPE=? UNIT=?> */ %>
	var TagSEARCH = ListXML.CreateRequestTag(window, "SEARCH");			
	ListXML.CreateActiveTag("FINANCIALSUMMARY");
	ListXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	ListXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
			
	ListXML.ActiveTag = TagSEARCH;
	var TagLIFEPRODUCTLIST = ListXML.CreateActiveTag("POLICYRELATIONSHIPLIST");
			
	<% /* Loop through all customer context entries */ %>
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
		var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		var sCustomerRoleType = scScreenFunctions.GetContextParameter(window,"idCustomerRoleType" + nLoop);

		<% /* If the customer is an applicant, add him/her to the search */ %>
		//SYS1672 get guarantor details too
		if(sCustomerRoleType == "1" || sCustomerRoleType == "2")
		{
			ListXML.CreateActiveTag("LIFEPRODUCT");
			ListXML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
			ListXML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
			ListXML.ActiveTag = TagLIFEPRODUCTLIST;
		}
	}
	ListXML.RunASP(document, "FindLifeProductSummary.asp");
			
	<% /* A record not found error is valid */ %>
	var sErrorArray = new Array("RECORDNOTFOUND");
	var sResponseArray = ListXML.CheckResponse(sErrorArray);
			
	if(sResponseArray[0] == true || sResponseArray[0] == false && sResponseArray[1] == "RECORDNOTFOUND")
	{
		<% /* If Financial Summary details have been returned, 
			set the question and remember that there is a record */ %>
		if(ListXML.SelectTag(null,"FINANCIALSUMMARY") != null)
		{
			m_bIsFinancialSummary = true;
			scScreenFunctions.SetRadioGroupValue(frmScreen, "LifeProductInd", ListXML.GetTagText("LIFEPRODUCTINDICATOR"));
		}
		PopulateTable(0);

		<% /* Initialise the scScrollPlus object */%>
		ListXML.ActiveTag = null;
		//next line replaced by line below as per V7.0.2 - DPF 20/06/02 - BMIDS00077
		//ListXML.CreateTagList("LIFEPRODUCTLIST");
		ListXML.CreateTagList("LIFEPRODUCT");
		scScrollPlus.Initialise(PopulateTable,10,ListXML.ActiveTagList.length);
	}
	<% /* Remember the initial state of the question */ %>
	m_sQuestionOnEntry = scScreenFunctions.GetRadioGroupValue(frmScreen, "LifeProductInd");
}

<% /* Displays a set of records in the table. This function is also used 
	by the scListScroll object. */ %>
function PopulateTable(nStart)
{
	scTable.clear();
			
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("LIFEPRODUCT");
			
	<% /* Populate the table with a set of records, starting with record nStart */ %>
	for (var nLoop=0; ListXML.SelectTagListItem(nStart + nLoop) != false && nLoop < 10; nLoop++)
	{				
		var TagLIFEPRODUCT = ListXML.ActiveTag;
		var sName = "";
		var sOrganisationName = "";
		var sPolicyType = "";
		var sMonthlyPremium	= "";
		var sYearOfMaturity	= "";
		var sMinimumDeathBenefit = "";
		var sMaturityValue = "";

		var sCustomerNumber = ListXML.GetTagText("CUSTOMERNUMBER");
		sName = scScreenFunctions.GetContextCustomerName(window,sCustomerNumber);

		if(ListXML.SelectTag(TagLIFEPRODUCT,"ACCOUNT") != null)
		{
			sOrganisationName = ListXML.GetTagText("COMPANYNAME");
		}
		if(ListXML.SelectTag(TagLIFEPRODUCT,"LIFEPOLICY") != null)
		{
			sPolicyType = ListXML.GetTagAttribute("POLICYTYPE","TEXT");
			sMonthlyPremium	= ListXML.GetTagText("MONTHLYPREMIUM");
			sYearOfMaturity	= ListXML.GetTagText("YEAROFMATURITY");
			sMinimumDeathBenefit = ListXML.GetTagText("MINIMUMDEATHBENEFIT");
			sMaturityValue = ListXML.GetTagText("MATURITYVALUE");
		}
		<% /* Display the details in the appropriate table row */ %>
		var nRow = nLoop + 1;
		scScreenFunctions.SizeTextToField(tblLifeProducts.rows(nRow).cells(0), sName);
		scScreenFunctions.SizeTextToField(tblLifeProducts.rows(nRow).cells(1), sOrganisationName);
		scScreenFunctions.SizeTextToField(tblLifeProducts.rows(nRow).cells(2), sPolicyType);
		scScreenFunctions.SizeTextToField(tblLifeProducts.rows(nRow).cells(3), sMonthlyPremium);
		scScreenFunctions.SizeTextToField(tblLifeProducts.rows(nRow).cells(4), sYearOfMaturity);
		scScreenFunctions.SizeTextToField(tblLifeProducts.rows(nRow).cells(5), sMinimumDeathBenefit);
		scScreenFunctions.SizeTextToField(tblLifeProducts.rows(nRow).cells(6), sMaturityValue);
	}
	// SYS0433 MDC 21/03/2000. Set focus to first item in list if an item exists
	//						   and enable Edit and Delete buttons. 	
	if (nLoop + nStart > 0)
	{
		scTable.setRowSelected(1);
		frmScreen.btnEdit.disabled = false;
		frmScreen.btnDelete.disabled = false;
	}	
}

<% /* Check whether any records exist in the table.
	If they do make sure the Yes button is set.*/ %>
function DoRadioButtonChecks()
{
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("LIFEPRODUCT");

	if(ListXML.ActiveTagList.length > 0)
	{
		<% /* If the Yes button isn't set, set it */%>
		if(frmScreen.optLifeProductYes.checked == false)
		{
			scScreenFunctions.SetRadioGroupValue(frmScreen,"LifeProductInd","1");
			alert("The question has now been set to 'Yes' as new customer data has been retrieved");
		}
		<% /* If there is data the No button can't be accessible, so disable the radio group */ %>
		scScreenFunctions.SetRadioGroupToReadOnly(frmScreen,"LifeProductInd");
	}
}

function frmScreen.btnEdit.onclick()
{
	ListXML.ActiveTag = null;
	ListXML.CreateTagList("LIFEPRODUCT");

	<% /* Get the index of the selected row */ %>
	var nRowSelected = scScrollPlus.getOffset() + scTable.getRowSelected();

	if(ListXML.SelectTagListItem(nRowSelected-1) == true)
	{
		if(CheckForIndicator() == true)
		{
			scScreenFunctions.SetContextParameter(window,"idMetaAction","Edit");
			scScreenFunctions.SetContextParameter(window,"idXML",ListXML.ActiveTag.xml);
			frmToDC101.submit();
		}
	}
}

function frmScreen.btnAdd.onclick()
{
	if(CheckForIndicator() == true)
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
		frmToDC101.submit();
	}
}
		
function frmScreen.btnDelete.onclick()
{	
	var bAllowDelete = confirm("Are you sure?");
				
	if (bAllowDelete == true)
	{
		ListXML.ActiveTag = null;
		ListXML.CreateTagList("LIFEPRODUCT");

		<% /* Get the index of the selected row */ %>
		var nRowSelected = scScrollPlus.getOffset() + scTable.getRowSelected();

		ListXML.SelectTagListItem(nRowSelected-1);
		var TagLIFEPRODUCT = ListXML.ActiveTag;

		<% /* Set up the deletion XML 
		next line replaced with line below as per V7.0.2 - DPF 20/06/02 - BMIDS00077*/ %>
		//var XML = new scXMLFunctions.XMLObject();
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		<% /* <REQUEST USERID=? USERTYPE=? UNIT=?> */ %>
		XML.CreateRequestTag(window, "DELETE");
		XML.CreateActiveTag("LIFEPRODUCT");
		XML.CreateTag("CUSTOMERNUMBER", ListXML.GetTagText("CUSTOMERNUMBER"));
		XML.CreateTag("CUSTOMERVERSIONNUMBER", ListXML.GetTagText("CUSTOMERVERSIONNUMBER"));
		XML.CreateTag("ACCOUNTGUID", ListXML.GetTagText("ACCOUNTGUID"));
<%
//				ListXML.SelectTag(TagPOLICYRELATIONSHIP,"FINANCIALORGANISATION");
//				XML.CreateActiveTag("FINANCIALORGANISATION");
//				XML.CreateTag("ORGANISATIONID",ListXML.GetTagText("ORGANISATIONID"));
%>				
		// 		XML.RunASP(document, "DeleteLifeProduct.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document, "DeleteLifeProduct.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

				
		<% /* If the deletion is successful remove the entry from the list xml and the screen */ %>
		if(XML.IsResponseOK() == true)
		{
			<% /* Remove the entry from the XML and redisplay the list */ %>
			ListXML.ActiveTag = TagLIFEPRODUCT;
			ListXML.RemoveActiveTag();
			scScrollPlus.RowDeleted();
						
			<% /* If the last entry was deleted make sure the selection returns to the new last entry */ %>
			if(ListXML.ActiveTagList.length < nRowSelected)
			{
				nRowSelected = nRowSelected - 1;
			}
						
			<% /* If there is still a line to select, select it. Otherwise, make the 
				No radio button available and disable the Edit/Delete buttons */ %>
			if(nRowSelected > 0)
			{
				scTable.setRowSelected(nRowSelected - scScrollPlus.getOffset());
			}
			else
			{
				scScreenFunctions.SetRadioGroupToWritable(frmScreen,"LifeProductInd");
				frmScreen.btnEdit.disabled = true;
				frmScreen.btnDelete.disabled = true;
				scScreenFunctions.SetFocusToFirstField(frmScreen);
			}
		}
		XML = null;					
	}				
}

function btnSubmit.onclick()
{
	if(scScreenFunctions.GetRadioGroupValue(frmScreen,"LifeProductInd") == null)
	{
		alert("Question must be answered in order to proceed");
		frmScreen.optLifeProductYes.focus();
	}
	else
	{
		if(CheckForIndicator() == true)
		{
			if(scScreenFunctions.CompletenessCheckRouting(window))
				frmToGN300.submit();
			else
				frmToDC080.submit();
		}
	}
}

function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
		frmToDC120.submit();
}

<% /* If the answer to the Life Product question has been changed 
	update the financial summary record */ %>
function CheckForIndicator()
{
	if (m_sReadOnly=="1") return true;
	
	var bOK	= false;
	var sQuestionOnExit	= scScreenFunctions.GetRadioGroupValue(frmScreen, "LifeProductInd");

	<% /* Has the answer changed? */ %>
	if(m_sQuestionOnEntry != sQuestionOnExit)
	{
		//next line replaced by line below as per V7.0.2 - DPF 20/06/02 - BMIDS00077
		//var XML = new scXMLFunctions.XMLObject();
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				
		<% /* <REQUEST USERID=? USERTYPE=? UNIT=?> */ %>
		if(m_bIsFinancialSummary == false)
		{
			XML.CreateRequestTag(window, "CREATE");
		}
		else
		{
			XML.CreateRequestTag(window, "UPDATE");
		}
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("LIFEPRODUCTINDICATOR", sQuestionOnExit);
				
		// If there was no financial summary record on entry create one
		if(m_bIsFinancialSummary == false)
		{
			// 			XML.RunASP(document, "CreateFinancialSummary.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "CreateFinancialSummary.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

		}
		else
		{
			// 			XML.RunASP(document, "UpdateFinancialSummary.asp");
			// Added by automated update TW 09 Oct 2002 SYS5115
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
							XML.RunASP(document, "UpdateFinancialSummary.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

		}
				
		if(XML.IsResponseOK() == true)
		{
			bOK = true;
		}
	}
	else
	{
		bOK = true;
	}
	return bOK;
}
-->
</script>
</body>
</html>




