<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<%
/*
Workfile:      cr010.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Customer Search screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		04/02/00	Optimised version
AY		07/02/00	Paged scroll functionality
AY		11/02/00	Change to msgButtons button types
AY		15/02/00	SYS0176 - Remove Flat No, House Name and House No
					SYS0219 - Remove Customer Number and associated processing
					SYS0010 - Introduced validation for wildcarding
AY		16/02/00	SYS0010 - Validation on surname/postcode not quite correct
AY		14/03/00	SYS0062 - More meaningful message for record not found error
AY		16/03/00	SYS0523 - Reference number was still being checked in IsBlankSearch()
AY		29/03/00	New top menu/scScreenFunctions change
SR		15/05/00	SYS0739 - Changed the width and maxLength of fields SurName and ForeName
							  and 'year4' is added to <FORM> tag.  
							  If forename = null and Surname <> null, display the message 
							  and set focus to Forename
BG		17/05/00	SYS0752 Removed Tooltips
SR		30/05/00	SYS0739 IsForenameRequired - Modified the If Condition
GD		20/02/2000  SYS1752 Child AQR to SYS1748
							Added extra column on table (mother's maiden name)
							Added CustomerNumber (OtherSystemCustomerNumber) Text box.
							Allows search to be done on OtherSystemCustomerNumber.
CL		05/03/01	SYS1920 Read only functionality added
SR		13/06/01	SYS2362 add new column 'Source' to the listbox							
DJP		25/09/01	SYS2564/SYS2735 (child) Add ability for client inheritence for the table and labels
JLD		18/12/01	SYS3097 select correct row after search
JLD		20/05/02	SYS4579 send error message for 0 customers returned from admin search.
LD		23/05/02	SYS4727 Use cached versions of frame functions
JLD		12/06/02	sys4728 use stylesheet at all times
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		22/05/2002				Modified SetContextFormForAddCustomer(), window.onload and Added PopulateScreen();
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
GHun	11/09/2002	BMIDS00429	Display OtherSystemType as Source for admin responses
GHun	13/09/2002	BMIDS00411	Always select 1st row in table when scrolling
GHun	19/09/2002	BMIDS00429	Convert OtherSystemType ValueID to ValueName
PSC		23/10/2002	BMIDS00465	Correct setting of OtherSystemCustomerNumber context
TW      09/10/2002  SYS5115	Modified to incorporate client validation
MV		05/11/2002	BMIDS00723 - Aligned the HTML tags
SA		20/11/2002	BMIDS01021 - Screen Rules now called on Search Button
HMA	    01/07/2004  BMIDS758    CC069 Do not allow the user to add a customer which has previously been deleted
HMA     26/07/2004  BMIDS758    Only check RemovedTOECustomer for Transfer of Equity applications.
HMA     04/08/2004  BMIDS836    Use "CC" as validation type for CC069.
HMA     08/12/2004  BMIDS957    Remove redundant VBScript code.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
PSC		05/10/2005	MAR57		Amend validation of search criteria
GHun	17/03/2006	MAR1453		Display progress window
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom Specific History:

Prog	Date		AQR			Description
PSC		11/01/2007	EP2_741		Replace CUSTOMERSTATUS with MOTHERSMAIDENNAME
SR		04/03/2007	EP2_1644	increased the width of scScrollPlus control
LDM		28/03/2007	EP2_2117	Limit the number of search results
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" viewastext tabIndex="-1"></object>
<% /* Scriptlets */ %>
<% /* CORE UPGRADE 702 <object data="scScreenFunctions.asp" height="1px" id="scScrFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object>
<object data="scXMLFunctions.asp" height="1px" id="scXMLFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabindex="-1"></object> */ %>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" viewastext tabindex="-1"></object>
<span style="LEFT: 306px; POSITION: absolute; TOP: 412px">
	<object data="scPageScroll.htm" id="scScrollPlus" style="LEFT: 0px; TOP: 0px; height:24; width:304" type="text/x-scriptlet" viewastext tabindex="-1"></object>
</span>
	
<% /* Specify Forms Here */ %>
<form id="frmAddCustomerSelect" method="post" action="cr020.asp" style="DISPLAY: none"></form>
<form id="frmNewCustomer" method="post" action="cr020.asp" style="DISPLAY: none"></form>
<form id="frmAddCustomerCancel" method="post" action="cr030.asp" style="DISPLAY: none"></form>
<form id="frmWelcomeMenu" method="post" action="mn010.asp" style="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" mark validate="onchange" year4>
<% /* Search Criteria */ %>
<div style="TOP: 60px; LEFT: 10px; HEIGHT: 122px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Search Criteria...</strong>
	</span>

	<div style="TOP: 24px; LEFT: 13px; HEIGHT: 60px; WIDTH: 578px; POSITION: ABSOLUTE" class="msgGroupLight">
		<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			<LABEL id="idSurname"></LABEL>
			<span style="TOP: -3px; LEFT: 75px; POSITION: ABSOLUTE">
				<input id="txtSurname" maxlength="40" style="WIDTH: 300px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 34px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			<LABEL id="idForename"></LABEL>
			<span style="TOP: -3px; LEFT: 75px; POSITION: ABSOLUTE">
				<input id="txtForenames" maxlength="40" style="WIDTH: 300px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 10px; LEFT: 413px; POSITION: ABSOLUTE" class="msgLabel">
			Date of Birth
			<span style="TOP: -3px; LEFT: 87px; POSITION: ABSOLUTE">
				<input type="text" id="txtDateOfBirth" maxlength="10" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 34px; LEFT: 413px; POSITION: ABSOLUTE" class="msgLabel">
			<LABEL id="idPostCode"></LABEL>
			<span style="TOP: -3px; LEFT: 87px; POSITION: ABSOLUTE">
				<input id="txtPostcode" maxlength="8" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxtUpper">
			</span>
		</span>
		
		<!--GD-->
<!--		<span style="TOP: 58px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Customer No
			<span style="TOP: -3px; LEFT: 75px; POSITION: ABSOLUTE">
				<input id="txtCustomerNo" maxlength="12" style="WIDTH: 300px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>
// -->
		<!--GD-->
	</div>

	<span style="TOP: 92px; LEFT: 522px; POSITION: ABSOLUTE">
		<input id="btnSearch" value="Search" type="button" style="WIDTH: 60px" class="msgButton">
	</span>
</div>

<% /* Search Results */ %>
<div style="TOP: 202px; LEFT: 10px; HEIGHT: 252px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
	<span style="TOP: 4px; LEFT: 4px; POSITION: ABSOLUTE" class="msgLabel">
		<strong>Search Results...</strong>
	</span>
<!--  #include FILE="customise/cr010CustomiseTable.asp" -->
	<span id="spnButtons" style="TOP: 220px; LEFT: 4px; POSITION: ABSOLUTE">
		<span style="TOP: 0px; LEFT: 0px; POSITION: ABSOLUTE">
			<input id="btnSelect" value="Select" type="button" style="WIDTH: 60px" class="msgButton" disabled>
		</span>

		<span style="TOP: 0px; LEFT: 64px; POSITION: ABSOLUTE">
			<input id="btnNewCustomer" value="New Customer" type="button" style="WIDTH: 90px" class="msgButton">
		</span>
	</span>
</div>
</form>

<%/* Main Buttons */ %>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 460px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /*File containing field attributes (remove if not required) - see xx999attribs.asp for example code */ %>
<!--  #include FILE="attribs/cr010attribs.asp" -->
<!--  #include FILE="customise/cr010customise.asp" -->

<% /* BMIDS957 Unused MyTrim VBScript function removed */ %>
	
<script language="JScript" type="text/javascript">
<!--
var XML;		
var scScreenFunctions;
		
<%/* form frmContext information */ %>
var m_sMetaAction = null;
var m_sAmountRequested = null;
var m_sMortgageProductId = null;
var m_sPurchasePrice = null;
var m_sTermYears = null;
var m_sTermMonths = null;
var m_sMortgageProductStartDate = null;
var m_sLendersCode = null;
var m_sApplicationNumber = null;
var DOWNLOADXTABLE = 10;
var m_blnReadOnly = false;
var OtherSystemTypeXML = null;
var m_sTypeOfApplicationValue = null;   // BMIDS758
		
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<%	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
	// Retrieve the context data
	// Set customer lock indicator in context APS 10/09/99 UNIT TEST REF 102
	// Conditionally disable the search criteria fields 
	// Initialise the table
	%>	var sButtonList = new Array("Cancel");
	ShowMainButtons(sButtonList);
	<% //DJP 25/09/2001 SYS2564 %>
	Customise();
	<% /* CORE UPGRADE 702 scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	XML = new scXMLFunctions.XMLObject();*/ %>
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	FW030SetTitles("Find Customer","CR010",scScreenFunctions);
	GetComboList();	<% /* BMIDS00429 */ %>		
	RetrieveContextData();						
	<% /* MV : 22/05/2002 :  BMIDS00005 - BM052 Customer Searches */ %>
	PopulateScreen();
<%	//This function is contained in the field attributes file (remove if not required)
%>	SetMasks();
	Validation_Init();
	scScreenFunctions.SetContextParameter(window,"idCustomerLockIndicator", "0");
	scTable.initialise(tblTable, 0, "");
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
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction", "CreateNewCustomerForNewApplication");
	m_sAmountRequested = scScreenFunctions.GetContextParameter(window,"idAmountRequested", null);
	m_sMortgageProductId = scScreenFunctions.GetContextParameter(window,"idMortgageProductId", null);
	m_sPurchasePrice = scScreenFunctions.GetContextParameter(window,"idPurchasePrice", null);
	m_sTermYears = scScreenFunctions.GetContextParameter(window,"idTermYears", null);
	m_sTermMonths = scScreenFunctions.GetContextParameter(window,"idTermMonths", null);
	m_sMortgageProductStartDate	= scScreenFunctions.GetContextParameter(window,"idMortgageProductStartDate", null);
	m_sLendersCode = scScreenFunctions.GetContextParameter(window,"idLendersCode", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber", null);
	m_sTypeOfApplicationValue = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue",null);    // BMIDS758
}	

<% /* MV : 22/05/2002 :  BMIDS00005 - BM052 Customer Searches */ %>
function PopulateScreen()
{
	m_SearchFlag = scScreenFunctions.GetContextParameter(window,"idSearchFlag", "0");
	
	if ( m_SearchFlag == '1')
	{
		frmScreen.txtForenames.value = scScreenFunctions.GetContextParameter(window,"idSearchForename", "0");
		frmScreen.txtSurname.value = scScreenFunctions.GetContextParameter(window,"idSearchSurname", "0");
		frmScreen.txtDateOfBirth.value = scScreenFunctions.GetContextParameter(window,"idSearchDateOfBirth", "0");
		frmScreen.txtPostcode.value = scScreenFunctions.GetContextParameter(window,"idSearchPostCode", "0");
		scScreenFunctions.SetContextParameter(window,"idSearchFlag", "0");
		<% /* PSC 05/10/2005 MAR57 - Start */ %>
		frmScreen.btnNewCustomer.disabled = false;	
	}
	else
		frmScreen.btnNewCustomer.disabled = true;	
		<% /* PSC 05/10/2005 MAR57 - End */ %>
}
		
function frmScreen.btnNewCustomer.onclick()
{				
	<%	// Do we need to create a new application for this new customer?
	// route to CR020
	%>	var sAction;
			
	/* CORE UPGRADE 702 if (m_sApplicationNumber == "") 
		sAction = "CreateNewCustomerForNewApplication";
	else 
		sAction = "CreateNewCustomerForExistingApplication";*/
	
	if (m_sApplicationNumber == "") sAction = "CreateNewCustomerForNewApplication";
	else sAction = "CreateNewCustomerForExistingApplication";
	
	if (scScreenFunctions.IsOmigaMenuFrame(window))
		window.parent.frames("omigamenu").document.forms("frmContext").idMetaAction.value = sAction;

	frmNewCustomer.submit();
}

function spnTable.ondblclick()
{
	if (scTable.getRowSelectedId() != null) frmScreen.btnSelect.onclick();
}
		
function frmScreen.btnSelect.onclick()
{
	<% /* BMIDS758  
	If this is a Transfer of Equity application stop the user from adding a customer which has previously been removed */ %>
	
	var ValidationList = new Array(1);
	
	ValidationList[0] = "CC";   // Transfer Of Equity   BMIDS836
	
	if ( m_sTypeOfApplicationValue != "" )
	{
		var combXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if(combXML.IsInComboValidationList(document,"TypeOfMortgage", m_sTypeOfApplicationValue , ValidationList)) 
		{				
			var sCustomerNumber;
			var sRemovedCustomerNumber;
			var bRemoved;
			var iNumberOfCustomers;

			XML.ResetXMLDocument();	
			XML.CreateRequestTag(window,null);
			XML.CreateActiveTag("REMOVEDTOECUSTOMER"); 
			XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
								
			switch (ScreenRules())
			{
				case 1: // Warning
				case 0: // OK
					XML.RunASP(document, "GetRemovedTOECustomers.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
			}

			sCustomerNumber = tblTable.rows(scTable.getRowSelectedId()).getAttribute("CustomerNumber");
			bRemoved = false;

			XML.ActiveTag = null;			
			XML.CreateTagList("REMOVEDTOECUSTOMER");
			
			iNumberOfCustomers = XML.ActiveTagList.length;
	
			for (i0 = 0; i0 < iNumberOfCustomers; i0++)
			{				
				XML.ActiveTag = null;
				XML.CreateTagList("REMOVEDTOECUSTOMER");  
				XML.SelectTagListItem(i0);
						
				sRemovedCustomerNumber = XML.GetTagText("OMIGACUSTOMERNUMBER");
		
				if (sCustomerNumber == sRemovedCustomerNumber) bRemoved = true;
			}

			if (bRemoved == true)
			{
				alert("This Customer has been deleted from this application and can not be re-added");
			}
			else
			{
				<% /* Set the data in the frmContext form for CR020. Route to CR020 */ %>
				SetContextFormForAddCustomer();
				frmAddCustomerSelect.submit();
			}		
		}
		else
		{
			SetContextFormForAddCustomer();
			frmAddCustomerSelect.submit();
		}
	}
	else
	{
		SetContextFormForAddCustomer();
		frmAddCustomerSelect.submit();
	}
}
		
function SetContextFormForAddCustomer()
{
<%	// do we need to create a new application for this existing customer?
	// for the selected row get the context information stored as attributes of the 
	// current <tr> tag
%>	var sAction, sCustomerNumber, sOtherSystemCustomerNumber, sCustomerVersionNumber;
			
	/* CORE UPGRADE 702 if (m_sApplicationNumber == "") 
		sAction = "CreateExistingCustomerForNewApplication";
	else 
		sAction = "CreateExistingCustomerForExistingApplication"; */
		
	if (m_sApplicationNumber == "") sAction = "CreateExistingCustomerForNewApplication";
	else sAction = "CreateExistingCustomerForExistingApplication";

	scScreenFunctions.SetContextParameter(window,"idMetaAction", sAction);
	
	<% /* MV : 22/05/2002 :  BMIDS00005 - BM052 Customer Searches 
	Start */ %>
	scScreenFunctions.SetContextParameter(window,"idSearchSurname", frmScreen.txtSurname.value);
	scScreenFunctions.SetContextParameter(window,"idSearchForename", frmScreen.txtForenames.value);
	scScreenFunctions.SetContextParameter(window,"idSearchDateOfBirth", frmScreen.txtDateOfBirth.value);
	scScreenFunctions.SetContextParameter(window,"idSearchPostCode", frmScreen.txtPostcode.value);
	<% /* End */ %>		
	
	sCustomerNumber = tblTable.rows(scTable.getRowSelectedId()).getAttribute("CustomerNumber");
	
	sOtherSystemCustomerNumber = tblTable.rows(scTable.getRowSelectedId()).getAttribute("OtherSystemCustomerNumber");
	sCustomerVersionNumber = tblTable.rows(scTable.getRowSelectedId()).getAttribute("CustomerVersionNumber");			
			
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber", sCustomerNumber);
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber", sCustomerVersionNumber);
	
	<% /* PSC 23/10/2002 BMIDS00465 */ %>
	if (m_sApplicationNumber == "" ||
	   (scScreenFunctions.GetContextParameter(window,"idOtherSystemCustomerNumber", "") == ""))
	scScreenFunctions.SetContextParameter(window,"idOtherSystemCustomerNumber", sOtherSystemCustomerNumber);
}
		
function btnCancel.onclick()
{
<%	// do we need to create a new application for this new customer?
	// route to CR030
%>	if (m_sApplicationNumber == "")
	{
		if (scScreenFunctions.IsOmigaMenuFrame(window))
		{
			window.parent.frames("omigamenu").document.forms("frmContext").idAmountRequested.value = "";
			window.parent.frames("omigamenu").document.forms("frmContext").idMortgageProductId.value = "";
			window.parent.frames("omigamenu").document.forms("frmContext").idPurchasePrice.value = "";
			window.parent.frames("omigamenu").document.forms("frmContext").idTermYears.value = "";
			window.parent.frames("omigamenu").document.forms("frmContext").idTermMonths.value = "";
		}
		frmWelcomeMenu.submit();				
	}
	else frmAddCustomerCancel.submit();
}

function IsSurnameRequired()
{			
	if	((frmScreen.txtSurname.value == "") && ((frmScreen.txtForenames.value != "") || (frmScreen.txtDateOfBirth.value != "")))
	{
		alert ("You must enter " + m_sSurname + " for " + m_sForename + "/Date of Birth");
		frmScreen.txtSurname.focus();
		return true;
	}

	return false;
}

<% /* SR 15/05/00 - SYS0739 - If ForeNams is null and Surname <> null then
	  display a warning message and set frocus to ForeName */
%>
function IsForenameRequired()
{
	if ((frmScreen.txtForenames.value == "") && (frmScreen.txtSurname.value != "" ))
	{
		alert ("You must enter " + m_sForename + " (Name or Initial or Wildcard) with " + m_sSurname);
		frmScreen.txtForenames.focus() ;
		return true ;	
	}
	return false ;
}

<% /* BMIDS957 Removed unused TrimCriteria funtion */ %>

function IsBlankSearch()
{			
	if ((frmScreen.txtSurname.value == "") && (frmScreen.txtForenames.value == "") && (frmScreen.txtDateOfBirth.value == "") &&
		(frmScreen.txtPostcode.value == ""))
	{
		alert ("You must enter search criteria");
		frmScreen.txtSurname.focus();
		return true;
	}			

	return false;
}

function IsCorrectWildcard()
{
<%	/*	Checks the fields which may be wildcarded to ensure that the wildcarding is valid
		The surname and postcode fields must have at least one character preceding the wildcard
		and there must be no characters following it.
		The forenames field does not require a preceding character.  It does, however, allow multiple
		wildcards in order to search on first forename and other forenames.  The rule here is that
		a wildcard may only be immediately followed by a space character.
	*/
	// Surname checks
%>	var sField = frmScreen.txtSurname.value;
	if(sField.indexOf("*") == 0)
	{
		alert("You must enter at least 1 preceding character to wildcard the " + m_sSurname);
		frmScreen.txtSurname.focus();
		return false;
	}

	if(sField.indexOf("*") < sField.length - 1 && sField.indexOf("*") != -1)
	{
		alert("There must be no characters after the wildcard for the " + m_sSurname);
		frmScreen.txtSurname.focus();
		return false;
	}
	
<%	// Forename checks
%>	sField = frmScreen.txtForenames.value;
	var nWildcardIndex = 0;
	do
	{
		nWildcardIndex = sField.indexOf("*",nWildcardIndex);
		if(nWildcardIndex != -1)
		{
			if(nWildcardIndex < sField.length - 1)
				if(sField.charAt(nWildcardIndex + 1) != ' ')
				{
					alert("A wildcard is being followed by a character other than a space in the " + m_sForename);
					frmScreen.txtForenames.focus();
					return false;
				}

			nWildcardIndex++;
		}
	}
	while(nWildcardIndex != -1)

<%	// Postcode checks
%>	sField = frmScreen.txtPostcode.value;
	if(sField.indexOf("*") == 0)
	{
		alert("You must enter at least 1 preceding character to wildcard the " + m_sPostCode);
		frmScreen.txtPostcode.focus();
		return false;
	}

	if(sField.indexOf("*") < sField.length - 1 && sField.indexOf("*") != -1)
	{
		alert("There must be no characters after the wildcard for the " + m_sPostCode);
		frmScreen.txtPostcode.focus();
		return false;
	}

	return true;
}

function frmScreen.btnSearch.onclick()
{						
<%	// Initialise the XML object
	// clear previous results
	// Initialise the scroll scriptlet to zero rows
%>	
	XML.ResetXMLDocument();
	scTable.clear();
	scScrollPlus.Clear();

<%	// remove as no ref in spec
	//TrimCriteria(); 
	// AY 15/02/00 SYS0010 - wildcard validation
%>		
<% /*  SR : 15/05/00  - Include the function IsForenameRequired in the following
						if condition
   */
%>							
<% /* SR : SYS2362	if (frmScreen.txtCustomerNo.value !="") 
	{
		if ((frmScreen.txtSurname.value!="") || (frmScreen.txtDateOfBirth.value !="") || (frmScreen.txtForenames.value !="") || (frmScreen.txtPostcode.value !=""))
		{
			alert("You must enter customer number OR customer details.");
		}
		else
		{

			scScrollPlus.Initialise(FindCustomers,ShowList,10,DOWNLOADXTABLE);
		}	
	}
	else
	{  */
%>
		<% /* PSC 05/10/2005 MAR57 */ %>
		if(!TooFewCriteriaEntered() && IsCorrectWildcard())
			if (!IsForenameRequired())
			{
				scScrollPlus.Initialise(FindCustomers,ShowList,10,DOWNLOADXTABLE);					
				<% /* PSC 05/10/2005 MAR57 */ %>
				frmScreen.btnNewCustomer.disabled = false;
			}
<% /*	} */ %>
}				

function GetForenames()
{
<%	// loop through each character in turn
	// if no white space found then only one forename is found
%>	var i0=0, i1=0;
	var iNumberOfForenames=0;
	var sFirstForename, sSecondForename, sOtherForename;
	var sChar;
	var iForenamesLength=0;			
	var sNames = new Array();
			
	var sForenames = frmScreen.txtForenames.value;			
						
	if (sForenames != "")
	{																			
		iForenamesLength = sForenames.length-1;
				
		while ((i0 <= iForenamesLength) && (iNumberOfForenames <= 2))
		{										
			sChar = sForenames.charAt(i0);					
			
			if (sChar == ' ')
			{
				switch (iNumberOfForenames)
				{
					case 0:	
						sFirstForename = sForenames.substring(0,i0);
						i1 = i0;
						iNumberOfForenames++;
						sSecondForename = sForenames.substring(i1+1);
						break;
					case 1:
						sSecondForename = sForenames.substring(i1+1,i0);
						sOtherForename = sForenames.substring(i0+1);								
						break;
					default:
						break;
				}													
			}
			i0++;														
		}
				
		if (iNumberOfForenames == 0)
		{
			sFirstForename = sForenames;
			sSecondForename = null;
			sOtherForename = null;
		}
				
<%		// where * = one white space then:
		//	"forename*secondname*othername" != "forename*****Secondname*****othername"
		// unless below code is uncommented					
		//else
		//{
		//	sFirstForename = MyTrim(sFirstForename);
		//	sSecondForename = MyTrim(sSecondForename);
		//	sOtherForename = MyTrim(sOtherForename);
		//}
%>																
		sNames[0] = sFirstForename;
		sNames[1] = sSecondForename;
		sNames[2] = sOtherForename;
	}

	return sNames;
}

function FindCustomers(nPageNo)
{						
	
	var	sASPFile;
	
	XML.ResetXMLDocument();
	XML.CreateRequestTag(window,"SEARCH");
	XML.CreateActiveTag("SEARCH");
	XML.SetAttribute("PAGESIZE",DOWNLOADXTABLE * 10);
	XML.SetAttribute("PAGENUMBER",nPageNo);
	XML.CreateActiveTag("CUSTOMER");
	
	scScreenFunctions.progressOn("Searching...", 400);	<% /* MAR1453 GHun */ %>
	
	var sForenames = new Array();
	sForenames = GetForenames();  
			
	XML.CreateTag("SURNAME", frmScreen.txtSurname.value);
	XML.CreateTag("FIRSTFORENAME", sForenames[0]);
				
<%	// APS UNIT TEST REF 53
%>	if (sForenames[1] != null) XML.CreateTag("SECONDFORENAME", sForenames[1]);
	if (sForenames[2] != null) XML.CreateTag("OTHERFORENAMES", sForenames[2]);

	XML.CreateTag("DATEOFBIRTH", frmScreen.txtDateOfBirth.value);
	XML.CreateTag("POSTCODE", frmScreen.txtPostcode.value);
	<%//GD SYS1752 added
	%>
<% /* SR SYS2362 - remove 'CustomerNo' from search
	XML.CreateTag("OTHERSYSTEMCUSTOMERNUMBER", frmScreen.txtCustomerNo.value); */
%>	
	sASPFile = "FindCustomers.asp";
	
	//BMIDS01021 Add in Screen Rules
	//XML.RunASP(document,sASPFile);
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,sASPFile);
			break;
		default: // Error
			XML.SetErrorResponse();
	}
	
	scScreenFunctions.progressOff();	<% /* MAR1453 GHun */ %>
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if(ErrorReturn[0])
	{
		if(XML.SelectTag(null,"CUSTOMERLIST") != null)
		{
			if(XML.GetAttribute("TOTAL") == "0")
			{
				alert("Unable to find any customers for your search criteria. Please amend the search criteria, wildcard your search or create a new customer record");
			}
			<% /* LDM 28/03/2007 EP2_2117 */ %>
			if(XML.GetAttribute("TOTAL") > (DOWNLOADXTABLE * 10))
			{
				alert("To many customers match your search criteria. The first "+(DOWNLOADXTABLE * 10)+" will be shown. Please refine your search criteria.");
				return DOWNLOADXTABLE * 10;
			}
			return XML.GetAttribute("TOTAL");
		}
			/* CORE UPGRADE 702 return XML.GetAttribute("TOTAL");*/
	}
	else if(ErrorReturn[1] == ErrorTypes[0])
		alert("Unable to find any customers for your search criteria. Please amend the search criteria, wildcard your search or create a new customer record");

	return 0;
}
<%
//GD SYS1752  changed from .... to allow sMaidenName
//function ShowRow(nIndex,sSurname,sForename,sDOB,sAddress,sCustomerNumber,sOtherSystemCustomerNumber,sCustomerVersionNumber)

%>
<% /* PSC 05/10/2005 MAR57 */ %>
<% /* PSC 11/01/2007 EP2_741 */ %>
function ShowRow(nIndex,sSurname,sForename,sDOB,sAddress,sMaidenName,sSource,sCustomerNumber,sOtherSystemCustomerNumber,sCustomerVersionNumber)
				
{			 					
<%	// AY 09/09/1999 - Set the table fields
	// Set the invisible context for each row
%>	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(0),sSurname);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(1),sForename);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(2),sDOB);
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(3),sAddress);
<%
//	GD SYS1752 added
%>
	<% /* PSC 05/10/2005 MAR57 */ %>
	<% /* PSC 11/01/2007 EP2_741 */ %>
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(4),sMaidenName);
	<% /* SR 08/06/01 : SYS2362 - add Source to the list box   */%>
	scScreenFunctions.SizeTextToField(tblTable.rows(nIndex).cells(5),sSource);
	
	tblTable.rows(nIndex).setAttribute("CustomerNumber", sCustomerNumber);
	tblTable.rows(nIndex).setAttribute("OtherSystemCustomerNumber", sOtherSystemCustomerNumber);
	tblTable.rows(nIndex).setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
	
}		

function ShowList(nStart)
{

	var i0=1, i1=0, iNumberOfCustomers=0;

	var sSurname, sForenames, sDOB, sMaidenName, sSource, sSourceID;

	var sFirstForename, sSecondForename, sOtherForename;
	var sCustomerNumber, sOtherSystemCustomerNumber, sCustomerVersionNumber;
	var sCustomerStatus;		<% /* PSC 05/10/2005 MAR57 */ %>

		
	scTable.clear();
	XML.ActiveTag = null;			
	XML.CreateTagList("CUSTOMER");
			
	iNumberOfCustomers = XML.ActiveTagList.length;
	
	for (i0 = 0; nStart + i0 < iNumberOfCustomers && i0 < 10; i0++)
	{				
		XML.ActiveTag = null;
		XML.CreateTagList("CUSTOMER");  
		XML.SelectTagListItem(nStart+i0);
						
		sCustomerNumber = XML.GetTagText("CUSTOMERNUMBER");
		sOtherSystemCustomerNumber = XML.GetTagText("OTHERSYSTEMCUSTOMERNUMBER");
		sCustomerVersionNumber = XML.GetTagText("CUSTOMERVERSIONNUMBER");

		sSurname = XML.GetTagText("SURNAME");		
				
<%		// APS 06/09/99 - UNIT TEST REF 33				
%>		sFirstForename = XML.GetTagText("FIRSTFORENAME");
		sSecondForename = XML.GetTagText("SECONDFORENAME");
		sOtherForename = XML.GetTagText("OTHERFORENAMES");		
		sForenames = sFirstForename + " " + sSecondForename + " " + sOtherForename;		
		sDOB = XML.GetTagText("DATEOFBIRTH");
		<%
		//GD SYS1752 Added
		%>
		<% /* PSC 05/10/2005 MAR57 - Start */ %>
		<% /* PSC 11/01/2007 EP2_741 - Start */ %>
		sMaidenName = XML.GetTagText("MOTHERSMAIDENNAME");
		<% /* sCustomerStatus = XML.GetTagText("CUSTOMERSTATUS"); */ %>
		<% /* PSC 11/01/2007 EP2_741 - End */ %> 
		<% /* PSC 05/10/2005 MAR57 - End */ %>
		<%
		//GD SYS1752 changed ShowRow(i0+1.. to ShowRow(i0+2 to account for bigger heading
		//GD SYS1752 plus...below...to account for maiden name
		//ShowRow (i0+2,sSurname,sForenames,sDOB,GetAddress(),sMaidenName,sCustomerNumber,sOtherSystemCustomerNumber,sCustomerVersionNumber); 
		
		//GD SYS1752 Removed: if(XML.SelectTag(XML.ActiveTag,"CUSTOMERADDRESS") != null)
		%>
		
		sSource = XML.GetTagText("SOURCE");
		
		<% /* BMIDS00429 Use OtherSystemType if Source is blank */ %>
		if (sSource.length == 0)
		{	
			sSourceID = XML.GetTagText("OTHERSYSTEMTYPE");
			OtherSystemTypeXML.ActiveTag = null;
			OtherSystemTypeXML.SelectSingleNode("RESPONSE/LIST/LISTNAME/LISTENTRY[VALUEID[.='" + sSourceID + "']]");
			if (OtherSystemTypeXML.ActiveTag != null)
				sSource = OtherSystemTypeXML.GetTagText("VALUENAME");
			else
				sSource = "undefined";
		}
		<% /* BMIDS00429 End */ %>
		
		<% /* PSC 05/10/2005 MAR57 */ %>
		<% /* PSC 11/01/2007 EP2_741 */ %>
		ShowRow (i0+1,sSurname,sForenames,sDOB,GetAddress(),sMaidenName,sSource,sCustomerNumber,sOtherSystemCustomerNumber,sCustomerVersionNumber);
	}
	
	<% /* BMIDS00411 */ %>	
	if(scScrollPlus.GetTotalRows() > 0)
	{
		<% /* select 1st row in the table and enable select button */ %>
		frmScreen.btnSelect.disabled = false;
		scTable.setRowSelected(1);
		frmScreen.btnSelect.focus();
	}	
	else
		frmScreen.btnSelect.disabled = true;
	<% /* BMIDS00411 End */ %>
}	
		
function GetAddress()
{		
	var sAddress = "";
						
<%	// APS UNIT TEST REF 54 - Added PostCode to the beginning of the
	// returned address information
%>	sAddress = AddComma(sAddress, XML.GetTagText("POSTCODE"), false);
	sAddress = AddComma(sAddress, XML.GetTagText("FLATNUMBER"), true);
	sAddress = AddComma(sAddress, XML.GetTagText("BUILDINGORHOUSENAME"), true);
	sAddress = AddComma(sAddress, XML.GetTagText("BUILDINGORHOUSENUMBER"), true); <% /* APS UNIT TEST REF 21 */ %>
	sAddress = AddComma(sAddress, XML.GetTagText("STREET"), true);
	sAddress = AddComma(sAddress, XML.GetTagText("DISTRICT"), true);
	sAddress = AddComma(sAddress, XML.GetTagText("TOWN"), true);
	sAddress = AddComma(sAddress, XML.GetTagText("COUNTY"), true);						
						
	return sAddress;
}
		
function AddComma(sAddress, sTagValue, bAddComma)
{																						
	if ((bAddComma==true) && (sTagValue != "") && (sAddress != "")) sTagValue = ", " + sTagValue;

	sAddress += sTagValue;
		
	return sAddress;
}

function GetComboList()
{
	OtherSystemTypeXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sCombo = new Array("OtherSystemType");
	
	OtherSystemTypeXML.GetComboLists(document,sCombo);
}
<% /* PSC 05/10/2005 MAR57 - Start */ %>
function frmScreen.txtSurname.onchange()
{
	frmScreen.btnNewCustomer.disabled = true;
}

function frmScreen.txtForenames.onchange()
{
	frmScreen.btnNewCustomer.disabled = true;
}

function frmScreen.txtDateOfBirth.onchange()
{
	frmScreen.btnNewCustomer.disabled = true;
}

function frmScreen.txtPostcode.onchange()
{
	frmScreen.btnNewCustomer.disabled = true;
}

function TooFewCriteriaEntered()
{
	var nNoOfFields = 0;
	
	if (frmScreen.txtSurname.value.length)
		nNoOfFields++;
		
	if (frmScreen.txtForenames.value.length)
		nNoOfFields++;
		
	if (frmScreen.txtDateOfBirth.value.length)
		nNoOfFields++;
	
	if (frmScreen.txtPostcode.value.length)
		nNoOfFields++;
		
	if (nNoOfFields < 2)
	{
		alert("At least TWO search criteria must be entered");
		frmScreen.txtSurname.focus();
		return true;
	}
	
	return false;
}
<% /* PSC 05/10/2005 MAR57 - End */ %>

-->
</script>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
</body>
</html>


