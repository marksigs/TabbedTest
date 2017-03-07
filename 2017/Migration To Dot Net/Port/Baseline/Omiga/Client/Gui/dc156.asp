<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      DC156.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Regular Outgoings Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
IW		10/04/2000	Created
BG		17/05/00	SYS0752 Removed Tooltips		
BG		17/05/00	Changed title to Regular Outgoings Details from Lender Details
MC		09/06/00	SYS0866 Standardise format of balance and payment fields
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		02/07/2002  BMIDS000168 -  Core Code Error Ref SYS0972
GHun	16/07/2002	BMIDS00190	DCWP3 BM076 Support linking multiple customers to outgoings
GHun	03/08/2002	BMIDS00401	If there is only one applicant it should always be selected
GHun	30/10/2002	BMIDS00731	Customers with alphas in the customer number are not displayed
SA		07/11/2002	BMIDS00832	Deal with alpha customer numbers
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
SR		12/06/2004  BMIDS772 -  Removed all the commnted code for BMIDS00190
								Update FinancialSummary on Submit (only for create)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
*/%>
</comment>

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
	<title>Regular Outgoings Details</title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>
<object data="scTable.htm" height="1" id="scTable" style="DISPLAY: none; LEFT: 0px; TOP: 0px" type="text/x-scriptlet" VIEWASTEXT tabindex="-1"></object>
	
<form id="frmToDC155" method="post" action="dc155.asp" STYLE="DISPLAY: none"></form>
	
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
	
<!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange">
	<div id="divBackground" style="TOP: 60px; LEFT: 10px; HEIGHT: 220px; WIDTH: 604px; POSITION: ABSOLUTE" class="msgGroup">
		<span style="TOP: 10px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Regular Outgoing
			<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
				<% /* BMIDS00190 Changed from text box to combo box */ %>
				<select id="cboRegularOutgoing" name="RegularOutgoing" style="WIDTH: 200px; POSITION: ABSOLUTE" class="msgCombo">
				</select>
			</span>
		</span>

		<% /* BMIDS00190 */ %>
		<span style="TOP: 36px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Applicant
		</span>

		<span id="spnApplicant" style="TOP: 36px; LEFT: 130px; POSITION: ABSOLUTE">
			<table id="tblApplicant" width="350px" border="0" cellspacing="0"cellpadding="0" class="msgTable">
				<tr id="rowTitles">	<td width="80%" class="TableHead">Name</td>			<td width="20%" class="TableHead">Selected</td></tr>
				<tr id="row01">		<td width="80%" class="TableTopLeft">&nbsp;</td>	<td width="20%" class="TableTopRight">&nbsp</td></tr>
				<tr id="row02">		<td width="80%" class="TableLeft">&nbsp;</td>		<td width="20%" class="TableRight">&nbsp;</td></tr>
				<tr id="row03">		<td width="80%" class="TableLeft">&nbsp;</td>		<td width="20%" class="TableRight">&nbsp;</td></tr>
				<tr id="row04">		<td width="80%" class="TableLeft">&nbsp;</td>		<td width="20%" class="TableRight">&nbsp;</td></tr>
				<tr id="row05">		<td width="80%" class="TableBottomLeft">&nbsp;</td>	<td width="20%" class="TableBottomRight">&nbsp;</td></tr>
			</table>
		</span>
			
		<% /* BMIDS00190 DCWP3 BM076 */ %>
		<span id="spnButtons" style="LEFT: 130px; POSITION: absolute; TOP: 128px">
			<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">
				<input id="btnSelectDeselect" value="Select/De-select" type="button" style="WIDTH: 100px" class="msgButton">
			</span>
		</span>
		<% /* BMIDS00190 End */ %>

		<% /* <span id="spnToPrev" tabindex="0"></span> */ %>

		<span style="TOP: 161px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Amount Paid
			<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
				<input id="txtOutgoingAmount" name="OutgoingAmount" maxlength="6" style="WIDTH: 70px; POSITION: ABSOLUTE" class="msgTxt">
			</span>
		</span>

		<span style="TOP: 190px; LEFT: 10px; POSITION: ABSOLUTE" class="msgLabel">
			Frequency
			<span style="TOP: -3px; LEFT: 120px; POSITION: ABSOLUTE">
				<select id="cboFrequency" name="Frequency" style="WIDTH: 100px" class="msgCombo">
				</select>
			</span>
		</span>

		<% /* <span id="spnToNext" tabindex="0"></span> */ %>
	</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 410px; WIDTH: 408px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>
	
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>
	
<!-- #include FILE="fw030.asp" -->

<!--  #include FILE="attribs/dc156attribs.asp" -->
	
<!-- Specify Code Here -->
<script language="JScript">
<!--
//var m_BaseNonPopupWindow = null;
<% /* BMIDS00190 */ %>
var m_sMetaAction = "";
var InXML = null;
var m_sReadOnly = "";
var m_iNumCustomers = 0;
var m_blnReadOnly = false;
		
<%/* MV - 02/07/2002 - BMIDS000168 - Core Code Error Ref SYS0972 */%>
var scScreenFunctions = null;

<% /* SR 12/06/2004 : BMIDS772 */ %>
var m_sApplicationNumber			= null;
var m_sApplicationFactFindNumber	= null;
<% /* SR 12/06/2004 : BMIDS772 - End */ %>
						
// JScript Code

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{		
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	<%/* MV - 02/07/2002 - BMIDS000168 - Core Code Error Ref SYS0972 */%>
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	RetrieveContextData();
					
	scTable.initialise(tblApplicant, 0, "");
		
	<% /* BMIDS00190 use standard buttons */ %>
	var sButtonList = new Array("Submit","Cancel","Another");

	// If not in add mode then the another button is not required
	if(m_sMetaAction != "Add")
		sButtonList = new Array("Submit","Cancel");

	ShowMainButtons(sButtonList);
	FW030SetTitles("Add/Edit Regular Outgoings","DC156",scScreenFunctions);
	<% /* BMIDS00190 End */ %>
		
	GetComboLists();
	PopulateScreen();

	SetMasks();

	//scScreenFunctions.ShowCollection(frmScreen);
	Validation_Init();
	
	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC156");
	if (m_blnReadOnly==true)
	{
		frmScreen.btnSelectDeselect.disabled = true;
		DisableMainButton("Another");
	}
	<% /* BMIDS00190 End */ %>
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
			
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

<% /* BMIDS00190 */ %>
function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}
<% /* BMIDS00190 End */ %>

/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Function:		GetComboLists()
	Description:	Populates all combos with their options
	Args Passed:	N/a
	Returns:		N/a
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
function GetComboLists()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	<% /* BMIDS00190 Also get regular outgoings type */ %>
	var sGroups = new Array("RegularOutgoingsPaymentFreq", "RegularOutgoingsType");
			
	if(XML.GetComboLists(document,sGroups))
	{
		XML.PopulateCombo(document,frmScreen.cboFrequency,"RegularOutgoingsPaymentFreq",false);
		frmScreen.cboFrequency.value = "12";
				
		<% /* BMIDS00190 Start */ %>
		if (m_sMetaAction == "Add")
		{
			XML.PopulateCombo(document,frmScreen.cboRegularOutgoing,"RegularOutgoingsType", true);
			frmScreen.cboRegularOutgoing.value = "";
		}
		else
			XML.PopulateCombo(document,frmScreen.cboRegularOutgoing,"RegularOutgoingsType", false);
		<% /* BMIDS00190 End */ %>
	}
	<% /* BMIDS00190 */ %>
	PopulateTable();
}

function PopulateScreen()
{		
	<% /* BMIDS00190 */ %>
	var sCustomerNumber;
	
	if (m_sMetaAction == "Edit")
	{
		var sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
		InXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
		InXML.LoadXML(sXML);
		
		if(InXML.SelectTag(null,"REGULAROUTGOINGS") != null)
		{
			if (m_iNumCustomers > 1)
			{
				var CustomersXML = InXML.ActiveTag.cloneNode(true).selectNodes("CUSTOMERVERSIONREGULAROUTGOINGS");
				for(var nCust=0; nCust < CustomersXML.length; nCust++)
				{
					sCustomerNumber = CustomersXML.item(nCust).selectSingleNode("CUSTOMERNUMBER").text;
					for(var nLoop=1; nLoop <= m_iNumCustomers; nLoop++)
					{
						if (tblApplicant.rows(nLoop).getAttribute("CUSTOMERNUMBER") == sCustomerNumber)
						{
							tblApplicant.rows(nLoop).setAttribute("Selected", "Yes");
							scScreenFunctions.SizeTextToField(tblApplicant.rows(nLoop).cells(1),"Yes");	
						}
					}
				}
			}
		
			frmScreen.cboRegularOutgoing.value = InXML.GetTagText("REGULAROUTGOINGSTYPE");
			frmScreen.txtOutgoingAmount.value = InXML.GetTagText("AMOUNT");
			frmScreen.cboFrequency.value = InXML.GetTagText("PAYMENTFREQUENCY");
		}		
	}					
}
	

function spnApplicant.onclick()
{
	sLastField = "";
	PopulateFields();
}
	
function btnSubmit.onclick()
{
	<% /* BMIDS00190 */ %>
	if (frmScreen.onsubmit())
		if (ValidateScreen())
			if (CommitScreen())
				frmToDC155.submit();
	<% /* BMIDS00190 End */ %>
}

function btnCancel.onclick()
{
	<% /* BMIDS00190 
	window.close();
	*/ %>
	frmToDC155.submit();
}

function btnAnother.onclick()
{
	<% /* BMIDS00190 */ %>
	if (frmScreen.onsubmit())
		if (ValidateScreen())
			if (CommitScreen())
			{
				PopulateTable();
				scScreenFunctions.ClearCollection(frmScreen);
				frmScreen.cboFrequency.value = "12";
				scScreenFunctions.SetFocusToFirstField(frmScreen);
			}
	<% /* BMIDS00190 End */ %>
}

<% /* BMIDS00190 */ %>
function CommitScreen()
{
	if (m_sReadOnly=="1") return true;

	var bOK = false;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var TagRequestType = null;
	var sCustomer;
	var sCustomerNumber;
	var blnCustomerPassedIn;
	var DeleteXML = null;

	var TagRequestType = XML.CreateRequestTag(window,null)
	
	XML.CreateActiveTag("REGULAROUTGOINGS");
	XML.CreateTag("REGULAROUTGOINGSTYPE",frmScreen.cboRegularOutgoing.value);
	XML.CreateTag("AMOUNT",frmScreen.txtOutgoingAmount.value);
	XML.CreateTag("PAYMENTFREQUENCY",frmScreen.cboFrequency.value);
	
	if(m_sMetaAction == "Add")
	{
		for(var nLoop=1;nLoop <= m_iNumCustomers; nLoop++)
		{
			if (tblApplicant.rows(nLoop).getAttribute("Selected") == "Yes")
			{
				sCustomerNumber = tblApplicant.rows(nLoop).getAttribute("CustomerNumber");
				sCustomerVersionNumber = tblApplicant.rows(nLoop).getAttribute("CustomerVersionNumber");
				XML.CreateActiveTag("CUSTOMERVERSIONREGULAROUTGOINGS");
				XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
				XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
				XML.ActiveTag = XML.ActiveTag.parentNode;
			}
		}
		
		<% /* SR 12/06/2004 : BMIDS772 - If FinancialSummary(FS) record needs to be updated, pass the 
			 node to CustomerFinancialBO. 
		*/ %>
		XML.ActiveTag = TagRequestType
		XML.CreateActiveTag("FINANCIALSUMMARY");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber );
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("REGULAROUTGOINGSINDICATOR", 1);
		<% /* SR 09/06/2004 : BMIDS772 - End */ %>
		

		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"CreateRegularOutgoings.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	}
	else
	{
		for(var nLoop=1;nLoop <= m_iNumCustomers; nLoop++)
		{
			sCustomerNumber = tblApplicant.rows(nLoop).getAttribute("CustomerNumber");
			sCustomerVersionNumber = tblApplicant.rows(nLoop).getAttribute("CustomerVersionNumber");			
			//BMIDS00832 add in single quotes to deal with alpha customer numbers
			if (InXML.ActiveTag.selectSingleNode(".//CUSTOMERNUMBER[.='" + sCustomerNumber + "']") == null)
				blnCustomerPassedIn = false;
			else
				blnCustomerPassedIn = true;
				
			if (tblApplicant.rows(nLoop).getAttribute("Selected") == "Yes")
			{
				if (!blnCustomerPassedIn) <% /* Add newly selected customers */ %>
				{
					XML.CreateActiveTag("CUSTOMERVERSIONREGULAROUTGOINGS");
					XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
					XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
					XML.ActiveTag = XML.ActiveTag.parentNode;
				}
			}
			else	<% /* Customer was selected, but is no longer, so delete the link */ %>
			{
				if (blnCustomerPassedIn)
				{
					if (DeleteXML == null)
					{	DeleteXML = XML.CreateActiveTag("DELETE");
					}
					
					XML.ActiveTag = DeleteXML;
					XML.CreateActiveTag("CUSTOMERVERSIONREGULAROUTGOINGS");
					XML.CreateTag("CUSTOMERNUMBER",sCustomerNumber);
					XML.CreateTag("CUSTOMERVERSIONNUMBER",sCustomerVersionNumber);
					XML.ActiveTag = XML.ActiveTag.parentNode.parentNode;
				}
			}
		}
		
		XML.CreateTag("REGULAROUTGOINGSGUID", InXML.GetTagText("REGULAROUTGOINGSGUID"));
		// 		XML.RunASP(document,"UpdateRegularOutgoings.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"UpdateRegularOutgoings.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	}

	bOK = XML.IsResponseOK();
	return bOK;
}

<% /* BMIDS00190 DCWP3 BM076 - Start */ %>
function ToggleSelection()
{
	var iRowSelected = scTable.getRowSelected();

	if ((iRowSelected > -1) && (m_iNumCustomers > 1))
	{
		if (tblApplicant.rows(iRowSelected).getAttribute("Selected") == "Yes")
		{
			scScreenFunctions.SizeTextToField(tblApplicant.rows(iRowSelected).cells(1),"No");
			tblApplicant.rows(iRowSelected).setAttribute("Selected", "No");
		}
		else
		{
			scScreenFunctions.SizeTextToField(tblApplicant.rows(iRowSelected).cells(1),"Yes");
			tblApplicant.rows(iRowSelected).setAttribute("Selected", "Yes");
		}
	}
}

function frmScreen.btnSelectDeselect.onclick()
{
	ToggleSelection();
}

function spnApplicant.onclick()
{
	if ((scTable.getRowSelectedId() != null) && (m_iNumCustomers > 1))
		<% /* MO - 18/11/2002 - BMIDS00376 */ %>
		if (m_blnReadOnly==false) {
			frmScreen.btnSelectDeselect.disabled = false;
		}
}

function spnApplicant.ondblclick()
{
	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	if (m_blnReadOnly==false) {
		ToggleSelection();
	}
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	<% /* BMIDS00190 */ %>
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	<% /* SR 12/06/2004 : BMIDS772 */ %>
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber");
	<% /* SR 12/06/2004 : BMIDS772 - End */ %>	
}

function ShowRow(nIndex,sCustomerName,sSelected,sCustomerNumber,sCustomerVersionNumber)
{
	<%	// Set the table fields %>	
	scScreenFunctions.SizeTextToField(tblApplicant.rows(nIndex).cells(0),sCustomerName);
	scScreenFunctions.SizeTextToField(tblApplicant.rows(nIndex).cells(1),sSelected);
	<%	// Set the invisible context for each row %>	
	tblApplicant.rows(nIndex).setAttribute("CustomerNumber", sCustomerNumber);
	tblApplicant.rows(nIndex).setAttribute("CustomerVersionNumber", sCustomerVersionNumber);
	tblApplicant.rows(nIndex).setAttribute("Selected", sSelected);
}

function PopulateTable()
{
	scTable.clear();
	m_iNumCustomers = 0;
		
	// Loop through all customer context entries
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + nLoop);
	
		<% /* BMIDS00731   if (sCustomerNumber > 0) */ %>
		if (sCustomerNumber.length > 0)
		{
			var sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + nLoop);
			var sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + nLoop);
		
			ShowRow(nLoop,sCustomerName,"No",sCustomerNumber,sCustomerVersionNumber);
			m_iNumCustomers++;
		}
	}
		
	<% /* When adding, if there is only one customer then select it by default */ %>
	<% /* BMIDS00401 - if ((m_iNumCustomers == 1) && (m_sMetaAction == "Add")) */ %>
	if (m_iNumCustomers == 1)
	<% /* BMIDS00401 End */ %>
	{
		scScreenFunctions.SizeTextToField(tblApplicant.rows(1).cells(1),"Yes");
		tblApplicant.rows(1).setAttribute("Selected", "Yes");
	}
	
	frmScreen.btnSelectDeselect.disabled = true;
}
	
function ValidateScreen()
{
	if (m_sReadOnly=="1") return true;
	
	var nSelected = 0;
	for(var nLoop=1; nLoop <= m_iNumCustomers; nLoop++)
	{
		if (tblApplicant.rows(nLoop).getAttribute("Selected") == "Yes")
			nSelected++;
	}
	if (nSelected == 0)
	{
		alert("At least one customer must be selected");
		tblApplicant.focus();
		return false;
	}
	
	return true;
}
<% /* BMIDS00190 - End */ %>
-->
</script>
</body>

</html>


