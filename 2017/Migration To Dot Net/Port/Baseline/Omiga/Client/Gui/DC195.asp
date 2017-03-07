<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/*
Workfile:      DC195.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Other Income Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		21/02/2000	Created
AD		02/03/2000	Fixed SYS0319
AY		31/03/00	New top menu/scScreenFunctions change
IVW		06/04/200	SYS0512 Switch off Delete and Edit if no list entries selected.
BG		17/05/00	SYS0752 Removed Tooltips 
BG		25/07/00	SYS0971 - Made customer name field longer to handle max length of customer name
CL		05/03/01	SYS1920 Read only functionality added
SA		30/05/01	SYS0973 When number of entries is greater than 10, alert the user to avoid runtime error.
SA		31/05/01	SYS0933	Edit button enabled in read only mode
JLD		10/12/01	SYS2806 completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		AQR			Description
MV		22/08/2002	BMIDS00355	IE 5.5 Upgrade - Modified the Msgbuttons Position 
GHun	08/11/2002	BMIDS00882	Non numeric customer number fails
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog    Date			Description
MF		22/07/2005		IA_WP01 process flow changes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>
<object data=scTableListScroll.asp name="scTable" id=scScrollTable style="DISPLAY: none; HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>

<%/* FORMS */%>
<form id="frmToDC160" method="post" action="DC160.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC200" method="post" action="DC200.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC196" method="post" action="DC196.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC085" method="post" action="DC085.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
<div id="divBackground" style="HEIGHT: 284px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 14px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Customer Name
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtCustomerName" style="WIDTH: 430px" class="msgReadOnly" readonly tabindex=-1>
			</input>
		</span> 
	</span>

	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 36px">
		<table id="tblTable" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="33%" class="TableHead">Other Income Type&nbsp</td>
								<td width="33%" class="TableHead">Amount&nbsp</td>
								<td width="34%" class="TableHead">Frequency&nbsp</td></tr>
			<tr id="row01">		<td width="33%" class="TableTopLeft">&nbsp;</td>	<td width="33%" class="TableTopCenter">	&nbsp;</td>		<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="33%" class="TableLeft">&nbsp;</td>		<td width="33%" class="TableCenter">&nbsp;</td>			<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="33%" class="TableBottomLeft">&nbsp;</td>	<td width="33%" class="TableBottomCenter">&nbsp;</td>	<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 198px">
		<input id="btnAdd" value="Add" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnAdd.onClick()"> 
	</span>

	<span style="LEFT: 68px; POSITION: absolute; TOP: 198px">
		<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnEdit.onClick()"> 
	</span>

	<span style="LEFT: 132px; POSITION: absolute; TOP: 198px">
		<input id="btnDelete" value="Delete" type="button" style="WIDTH: 60px" class="msgButton" onclick="btnDelete.onClick()"> 
	</span> 

	<span style="LEFT: 14px; POSITION: absolute; TOP: 234px" class="msgLabel">
		Other Income (Annually)
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtTotalOtherIncome" style="WIDTH: 100px" class="msgReadOnly" readonly tabindex=-1>
			</input>
		</span> 
	</span>

	<span style="LEFT: 14px; POSITION: absolute; TOP: 258px" class="msgLabel">
		Total Income (Annually)
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtTotalGrossIncome" style="WIDTH: 100px" class="msgReadOnly" readonly tabindex=-1>
			</input>
		</span> 
	</span>
</div> 
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 370px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/DC195Attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sCustomerNumber = "";
var m_sCustomerVersionNumber = "";
var m_sCustomerName = "";
var m_nCustomerIndex = 0;
var OtherIncomeXML = null;
var scScreenFunctions;
var m_blnReadOnly = false;


/* EVENTS */

function frmScreen.btnAdd.onclick()
{
	<%/*SYS0973 SA 30/5/01 If entries in table has reached 10 - don't allow anymore. */ %>
	OtherIncomeXML.ActiveTag = null;
	OtherIncomeXML.CreateTagList("UNEARNEDINCOME");
	var iNumberOfOtherIncomes = OtherIncomeXML.ActiveTagList.length;
	if ((iNumberOfOtherIncomes >= 10)==true)
	{
		alert("Maximum number of entries reached");
		return false
	}	
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber", m_sCustomerNumber);
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber", m_sCustomerVersionNumber);
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "Add");
	frmToDC196.submit()
}

function btnCancel.onclick()
{
	m_nCustomerIndex--;
	scScreenFunctions.SetContextParameter(window,"idCustomerIndex", m_nCustomerIndex);

	if (RetrieveCustomerData())
	{
		scTable.clear();
		Initialise();
	}
	else
	{
		scScreenFunctions.SetContextParameter(window,"idCustomerIndex", "");
		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else
			frmToDC160.submit();
	}
}

function frmScreen.btnDelete.onclick()
{
	if (!confirm("Are you sure?")) return;

	//Get the XML that just contains the GroupConnection chosen in the listbox
	var XML = GetXMLBlock(true);
	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();

	// 	XML.RunASP(document,"DeleteOtherIncome.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"DeleteOtherIncome.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	if (XML.IsResponseOK())
	{
		// Need to hit the database because the totals need to be re-calculated
		scTable.clear();
		Initialise();
	}
}

function frmScreen.btnEdit.onclick()
{
	var XML = GetXMLBlock(true);

	if (XML != null)
	{
		XML.SelectTag(null,"UNEARNEDINCOME");
		scScreenFunctions.SetContextParameter(window,"idCustomerNumber", m_sCustomerNumber);
		scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber", m_sCustomerVersionNumber);
		scScreenFunctions.SetContextParameter(window,"idXML", XML.XMLDocument.xml);
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "Edit");
		frmToDC196.submit();
	}
}

function btnSubmit.onclick()
{
	var bSuccess = CommitData();
	if(bSuccess)
	{
		m_nCustomerIndex++;
		scScreenFunctions.SetContextParameter(window,"idCustomerIndex", m_nCustomerIndex);

		if (RetrieveCustomerData())
		{
			scTable.clear();
			Initialise();
		}
		else
		{
			scScreenFunctions.SetContextParameter(window,"idCustomerIndex", "");
			if(scScreenFunctions.CompletenessCheckRouting(window))
				frmToGN300.submit();
			else
<% /* MF 22/07/2005 MARS IA_WP01 route to Financial summary
				//frmToDC200.submit(); */ %>
				frmToDC085.submit();
		}
	}
}		

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Other Income Details","DC195",scScreenFunctions);
	
	
	frmScreen.btnEdit.disabled = true;
	frmScreen.btnDelete.disabled = true;

	RetrieveContextData();
	if (m_nCustomerIndex == "") m_nCustomerIndex = 1;
	RetrieveCustomerData();
	Initialise();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC195");
	if (m_blnReadOnly == true) 
	{
		m_sReadOnly = "1";
		<% /* MO - 18/11/2002 - BMIDS00376 */ %>
		frmScreen.btnAdd.disabled = true;
		frmScreen.btnDelete.disabled = true;
	}
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */

function CommitData()
{
	// Nothing to save (as yet)
	return(true);
}

function GetXMLBlock(bForEdit)
{
	OtherIncomeXML.ActiveTag = null;
	OtherIncomeXML.CreateTagList("UNEARNEDINCOME");

	<% /* Get the index of the selected row */ %>
	var nRowSelected = scTable.getOffset() + scTable.getRowSelected();
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	if(OtherIncomeXML.SelectTagListItem(nRowSelected-1) == true)
	{
		XML.CreateRequestTag(window,null);
		XML.ActiveTag.appendChild(OtherIncomeXML.ActiveTag);
	}
	else
		XML = null;

	return(XML);
}

function Initialise()
{
	PopulateScreen();
}

function PopulateScreen()
{
	frmScreen.txtCustomerName.value = m_sCustomerName;
	OtherIncomeXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	OtherIncomeXML.CreateRequestTag(window,null);
	OtherIncomeXML.CreateActiveTag("UNEARNEDINCOME");
	OtherIncomeXML.CreateTag("CUSTOMERNUMBER", m_sCustomerNumber);
	OtherIncomeXML.CreateTag("CUSTOMERVERSIONNUMBER", m_sCustomerVersionNumber);

	// 	OtherIncomeXML.RunASP(document,"FindOtherIncomeList.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			OtherIncomeXML.RunASP(document,"FindOtherIncomeList.asp");
			break;
		default: // Error
			OtherIncomeXML.SetErrorResponse();
		}


	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = OtherIncomeXML.CheckResponse(ErrorTypes);
	var nListLength = 0;
	
	//if ((ErrorReturn[1] == ErrorTypes[0]) | (OtherIncomeXML.XMLDocument.text == ""))
	
	// SYS0512 - The above code to check the text is incorrect, it can bring back
	// two zero amounts giving a string of "00" making the code think it has
	// found income records. The check must be more specific.
	
	OtherIncomeXML.ActiveTag = null;
	OtherIncomeXML.CreateTagList("UNEARNEDINCOME");
	nListLength = OtherIncomeXML.ActiveTagList.length;

	// End SYS0512

	if (ErrorReturn[1] == ErrorTypes[0] || nListLength == 0)
	{
		//Error: record not found
		PopulateTotals();
		if(m_sReadOnly == "1")
			frmScreen.btnAdd.disabled = true;
		else
			frmScreen.btnAdd.disabled = false;
		frmScreen.btnDelete.disabled = true;
		frmScreen.btnEdit.disabled = true;
	}
	else if (ErrorReturn[0] == true)
	{
		// No error
		PopulateTable();
		PopulateTotals();
		if(m_sReadOnly == "1")
		{
			frmScreen.btnAdd.disabled = true;
			frmScreen.btnDelete.disabled = true;
			frmScreen.btnEdit.disabled = false;
			frmScreen.btnEdit.focus();
		}
		else
		{
			frmScreen.btnAdd.disabled = false;
			frmScreen.btnDelete.disabled = false;
			frmScreen.btnEdit.disabled = false;
			frmScreen.btnAdd.focus();
		}
	}

	ErrorTypes = null;
	ErrorReturn = null;

	scTable.initialiseTable(tblTable,0,"",PopulateTable,10,nListLength);
	if (nListLength > 0) scTable.setRowSelected(1);
}

function PopulateTable()
{
	OtherIncomeXML.ActiveTag = null;
	OtherIncomeXML.CreateTagList("UNEARNEDINCOME");
	var iCount;
	var iNumberOfOtherIncomes = OtherIncomeXML.ActiveTagList.length;
	//SYS0973 SA Max number of entries is 10 - so check. User is informed when attempting to add.
	if (iNumberOfOtherIncomes > 10)
	{
		iNumberOfOtherIncomes = 10
	}
	for (iCount = 0; iCount < iNumberOfOtherIncomes ; iCount++)
	{
		OtherIncomeXML.SelectTagListItem(iCount);

		var sOtherIncomeType = OtherIncomeXML.GetTagAttribute("UNEARNEDINCOMETYPE","TEXT");
		var sAmount = OtherIncomeXML.GetTagText("UNEARNEDINCOMEAMOUNT");
		var sFrequency = OtherIncomeXML.GetTagAttribute("PAYMENTFREQUENCY", "TEXT");

		// Add to the search table
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(0),sOtherIncomeType);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(1),sAmount);
		scScreenFunctions.SizeTextToField(tblTable.rows(iCount+1).cells(2),sFrequency);
		tblTable.rows(iCount+1).setAttribute("TagListItemCount", iCount);
	}
}

function PopulateTotals()
{
	var OtherIncomeTag = OtherIncomeXML.SelectTag(null,"UNEARNEDINCOMELIST");
	var sTotalOtherIncome = "0.00";
	var sTotalGrossIncome = "0.00";

	if (OtherIncomeTag != null)
	{
		var sTotalOtherIncome = OtherIncomeXML.GetTagText("TOTALOTHERINCOMEAMOUNT");
		var sTotalGrossIncome = OtherIncomeXML.GetTagText("TOTALGROSSINCOME");
	}

	if (sTotalOtherIncome == "") sTotalOtherIncome = "0.00";
	if (sTotalGrossIncome == "") sTotalGrossIncome = "0.00";

	frmScreen.txtTotalOtherIncome.value = sTotalOtherIncome;
	frmScreen.txtTotalGrossIncome.value = sTotalGrossIncome;
}
		
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_nCustomerIndex = scScreenFunctions.GetContextParameter(window,"idCustomerIndex", "1");
}

function RetrieveCustomerData()
{
	m_sCustomerNumber = "";
	m_sCustomerVersionNumber = "";
	m_sCustomerName = "";

	if ((m_nCustomerIndex > 0) && (m_nCustomerIndex < 6))
	{
		m_sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + m_nCustomerIndex,"1325");
		m_sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + m_nCustomerIndex,"1");
		m_sCustomerName = scScreenFunctions.GetContextParameter(window,"idCustomerName" + m_nCustomerIndex,"Cust");
	}

	<% /* BMIDS00882 Customer numbers can contain alphanumeric characters
	return(!isNaN(parseInt(m_sCustomerNumber)) & !isNaN(parseInt(m_sCustomerVersionNumber)));
	*/ %>
	if ((m_sCustomerNumber.length > 0) && (!isNaN(parseInt(m_sCustomerVersionNumber))))
		return true
	else
		return false;	
	<% /* BMIDS00882 End */ %>
}
-->
</script>
</body>
</html>
