<%@ LANGUAGE="JSCRIPT" %>
<% /*
Workfile:		cm150.asp
Copyright:		Copyright © 1999 Marlborough Stirling

Description:	Incentives Screen.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
RF		08/02/00	Reworked for IE5 and performance.
JLD		07/03/00	SYS0401 - show no decimal places for amount in incentives tables
AY		17/03/00	MetaAction changed to ApplicationMode
AY		29/03/00	scScreenFunctions change
AY		13/04/00	SYS0328 - Dynamic currency display
JLD		17/04/00	Changes to interface with cm110, mc010 and mc030
JLD		20/04/00	Fixes to the listboxes
JLD		03/05/00	If no incentives are chosen, return empty string.
JLD		09/05/00	In SetChosenIncentives, compare AMOUNT as a float not a string ("20" != "20.00")
SR		18/06/01	SYS2377	- Add new columns 'Benefit Type', 'Notional Amount' and 'Add?' to the ListBox
DRC     17/12/01    SYS3275 - Check for non-negative row number in IncentivesTable.ondblclick events
PSC		17/01/02	SYS3801 - Amend CreateIncentivesXML to create correct xml
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
DPF		17/10/2002	BMIDS00046	Takes in extra array element (SubQuoteXML) and calls new methods to get
								incentives
DPF		29/10/2002	BMIDS00757	Have applied fix for when you are searching for 1st loan component created
DPF		30/10/2002	BMIDS00767	Have applied fix to ShowInclusiveIncentives as it had a spelling error.
JD		26/07/04	BMIDS816	ApplicationDate is required in call to FindAvailableIncentives
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
<html>

<%

/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MC		19/04/2004	BMIDS517	White space padded to the title text.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<HEAD>
<META HTTP-EQUIV="expires" CONTENT="Wed, 01 Jan 2003 01:01:00 GMT">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Incentives <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>
	
<% // List Scroll objects %>
<span id=spnInclusiveListScroll>
	<span style="LEFT: 154px; POSITION: absolute; TOP: 150px">
		<object data=scTableListScroll.asp id=scInclusiveTable 
			style="LEFT: 0px; TOP: 0px; height:24; width:304" 
			type=text/x-scriptlet VIEWASTEXT tabindex="-1"></object>
	</span> 
</span>	
<span id=spnExclusiveListScroll>
	<span style="LEFT: 154px; POSITION: absolute; TOP: 322px">
		<object data=scTableListScroll.asp id=scExclusiveTable 
			style="LEFT: 0px; TOP: 0px; height:24; width:304" 
			type=text/x-scriptlet VIEWASTEXT tabindex="-1"></object>
	</span> 
</span>	

<% // Specify Forms Here %>

<% // Specify Screen Layout Here %>
<form id="frmScreen" mark validate="onchange" year4>

	<div id="divBackground" style="TOP: 10px; LEFT: 10px; HEIGHT: 380px; 
		WIDTH: 472px; POSITION: ABSOLUTE" class="msgGroup">
				
		<% // Inclusive Incentives %>
		<div style="HEIGHT: 0px; LEFT: 10px; POSITION: absolute; 
			TOP: 0px; WIDTH: 443px" class="msgGroup">
			<span style="TOP: 8px; LEFT: 6px; POSITION: ABSOLUTE" class="msgLabel">
				<strong>Inclusive Incentives...</strong>
			</span>
			<span id="spnInclusiveIncentivesTable" 
				style="LEFT: 4px; POSITION: absolute; TOP: 30px">
				<table id="tblInclusiveIncentivesTable" width="435" border="0" cellspacing="0" cellpadding="0" class="msgTable">
					<tr id="rowTitles">	<td width="45%" class="TableHead">Description</td>	<td width="30%" class="TableHead">Benefit Type</td>		<td id="tdAmount1" width="10%" class="TableHead">Amount</td>	<td width="10%" class="TableHead">Notional Amount</td>	<td width="5%" class="TableHead">Add?</td>			</tr>
					<tr id="row01">		<td width="45%" class="TableTopLeft">&nbsp</td>		<td width="30%" class="TableTopCenter">&nbsp;</td>		<td width="10%" class="TableTopCenter">&nbsp;</td>				<td width="10%" class="TableTopCenter">&nbsp;</td>		<td width="5%" class="TableTopRight">&nbsp</td>		</tr>
					<tr id="row02">		<td width="45%" class="TableLeft">&nbsp</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp</td>					<td width="10%" class="TableCenter">&nbsp</td>			<td width="5%" class="TableRight">&nbsp</td>		</tr>
					<tr id="row03">		<td width="45%" class="TableLeft">&nbsp</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>					<td width="10%" class="TableCenter">&nbsp;</td>			<td width="5%" class="TableRight">&nbsp</td>		</tr>
					<tr id="row04">		<td width="45%" class="TableLeft">&nbsp</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>					<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableRight">&nbsp</td>		</tr>
					<tr id="row05">		<td width="45%" class="TableBottomLeft">&nbsp</td>	<td width="30%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>			<td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomRight">&nbsp</td>	</tr>
				</table>
			</span>						
		</div> 	

		<% // Exclusive Incentives %>
		<div style="HEIGHT: 0px; LEFT: 10px; POSITION: absolute; 
			TOP: 172px; WIDTH: 443px" class="msgGroup">
			<span style="TOP: 8px; LEFT: 6px; POSITION: ABSOLUTE" class="msgLabel">
				<strong>Exclusive Incentives...</strong>
			</span>
			<span id="spnExclusiveIncentivesTable" 
				style="LEFT: 4px; POSITION: absolute; TOP: 30px">
				<table id="tblExclusiveIncentivesTable" width="435" border="0" cellspacing="0" cellpadding="0" class="msgTable">
					<tr id="rowTitles">	<td width="45%" class="TableHead">Description</td>	<td width="30%" class="TableHead">Benefit Type</td>		<td id="tdAmount2" width="10%" class="TableHead">Amount</td>	<td width="10%" class="TableHead">Notional Amount</td>	<td width="5%" class="TableHead">Add?</td>			</tr>
					<tr id="row01">		<td width="45%" class="TableTopLeft">&nbsp</td>		<td width="30%" class="TableTopCenter">&nbsp;</td>		<td width="10%" class="TableTopCenter">&nbsp;</td>				<td width="10%" class="TableTopCenter">&nbsp;</td>		<td width="5%" class="TableTopRight">&nbsp</td>		</tr>
					<tr id="row02">		<td width="45%" class="TableLeft">&nbsp</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp</td>					<td width="10%" class="TableCenter">&nbsp</td>			<td width="5%" class="TableRight">&nbsp</td>		</tr>
					<tr id="row03">		<td width="45%" class="TableLeft">&nbsp</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>					<td width="10%" class="TableCenter">&nbsp;</td>			<td width="5%" class="TableRight">&nbsp</td>		</tr>
					<tr id="row04">		<td width="45%" class="TableLeft">&nbsp</td>		<td width="30%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableCenter">&nbsp;</td>					<td width="10%" class="TableCenter">&nbsp;</td>			<td width="10%" class="TableRight">&nbsp</td>		</tr>
					<tr id="row05">		<td width="45%" class="TableBottomLeft">&nbsp</td>	<td width="30%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>			<td width="10%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomRight">&nbsp</td>	</tr>
				</table>
			</span>						
		</div>
		
		<% // Buttons %>
		<div style="LEFT: 14px; POSITION: absolute; TOP: 284px"class="msgGroup">
			<span style="TOP: 65px; LEFT: 4px; POSITION: ABSOLUTE">
				<input id="btnOK" value="OK" type="button" 
					style="WIDTH: 60px" class="msgButton">
				<span style="TOP: 0px; LEFT: 64px; POSITION: ABSOLUTE">
					<input id="btnCancel" value="Cancel" type="button" 
						style="WIDTH: 60px" class="msgButton">
				</span>
			</span>					
		</div>
	</div>
	
</form>	

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/cm150Attribs.asp" -->

<% // Specify Code Here %>
<script language="JScript">
<!--	// JScript Code

var m_sReadOnly					= null;

var m_XMLInclusiveIncentives	= null;
var m_XMLExclusiveIncentives	= null;
var m_sApplicationMode			= null;
var	m_sUserId					= null;
var	m_sUserType					= null;
var	m_sUnitId					= null;
var m_sRequestArray				= "";
var m_sInIncentivesXML			= "";
var m_sApplicationNumber		= "";
var m_sApplicationFactFindNumber = "";
var m_sApplicationDate			= ""; // JD BMIDS816
var m_sMortgageSubQuoteNumber	= "";
var m_sLoanComponentSeqNum		= "";
var m_iTableLength				= 5;
var m_ChosenExcIncentive		= null;
var m_ChosenIncArray			= null;
//DPF 17/10/02 - CPWP1 - New variable for SubQuoteXML
var m_sSubQuoteXML				= null;

var scScreenFunctions;
var m_BaseNonPopupWindow = null;
		

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{						
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	RetrieveData();			
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	<% /* scInclusiveTable.EnableMultiSelectTable(); */ %>
	GetIncentiveDetails()									
	SetScreenOnReadOnly();				
			
	if(m_sApplicationMode == "Mortgage Calculator")
		frmScreen.btnCancel.disabled = true;
				
	window.returnValue = null;
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveData()
{
	<% /* Retrieve information from calling screen */ %>
	var sArguments = window.dialogArguments;
	
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];			
	var sParameters		= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];
																	
	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sApplicationMode				= sParameters[0];
	m_sReadOnly						= sParameters[1];
	m_sRequestArray					= sParameters[2];
	m_sMortgageProductCode			= sParameters[3];
	m_sMortgageProductStartDate		= sParameters[4];
	m_sLoanAmount					= sParameters[5];
	if(m_sApplicationMode != "Mortgage Calculator")
	{
		m_sInIncentivesXML = sParameters[7];
		m_sApplicationNumber = sParameters[8];
		m_sApplicationFactFindNumber = sParameters[9];
		m_sMortgageSubQuoteNumber = sParameters[10];
		m_sLoanComponentSeqNum = sParameters[11];
		//DPF 17/10/02 - New incoming parameter (SubQuoteXML)
		m_sSubQuoteXML = sParameters[12];
		m_sApplicationDate = sParameters[13]; // JD BMIDS816
	}

	var sCurrencyText = scScreenFunctions.SetThisCurrency(frmScreen,sParameters[6]);
	tdAmount1.innerText = tdAmount1.innerText + " " + sCurrencyText;
	tdAmount2.innerText = tdAmount2.innerText + " " + sCurrencyText;
				
	m_XMLInclusiveIncentives = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_XMLExclusiveIncentives = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
}

function ShowExclusiveIncentives(iStart)
{			
	var sIncentivesDescription, sIncentivesAmount, sBenefitType, sNotionalAmount;
	var sGuid;	
						
	for (var iLoop = 0; iLoop < m_XMLExclusiveIncentives.ActiveTagList.length && iLoop < m_iTableLength;
		iLoop++)
	{				
		m_XMLExclusiveIncentives.SelectTagListItem(iLoop + iStart);
		sIncentivesDescription		= m_XMLExclusiveIncentives.GetTagText("DESCRIPTION");
		sIncentivesAmount			= m_XMLExclusiveIncentives.GetTagText("AMOUNT");
		sBenefitType				= m_XMLExclusiveIncentives.GetTagAttribute("INCENTIVEBENEFITTYPE", "TEXT");
		sNotionalAmount				= m_XMLExclusiveIncentives.GetTagText("NOTIONALAMOUNT");
		
		ShowExclusiveIncentiveRow(iLoop+1, sIncentivesDescription, sIncentivesAmount, sBenefitType, sNotionalAmount);	
	}
	if(m_sApplicationMode == "Mortgage Calculator")
	{	
		scExclusiveTable.setRowSelected(1);
		scExclusiveTable.DisableTable();
	}
	else setSelectedExclusive();												
}
		
function ShowExclusiveIncentiveRow(nIndex, sIncentivesDescription, sIncentivesAmount, sBenefitType,sNotionalAmount)
{			 					
	scScreenFunctions.SizeTextToField(tblExclusiveIncentivesTable.rows(nIndex).cells(0),
											sIncentivesDescription);
	scScreenFunctions.SizeTextToField(tblExclusiveIncentivesTable.rows(nIndex).cells(1),
											sBenefitType);											
	scScreenFunctions.SizeTextToField(tblExclusiveIncentivesTable.rows(nIndex).cells(2), 
											sIncentivesAmount);
	scScreenFunctions.SizeTextToField(tblExclusiveIncentivesTable.rows(nIndex).cells(3), 
											sNotionalAmount);
}		

function ShowInclusiveIncentives(iStart)
{			
	var sIncentivesDescription, sIncentivesAmount, sBenefitType, sNotionalAmount;
	var sGuid;
						
	for (var iLoop = 0; iLoop < m_XMLInclusiveIncentives.ActiveTagList.length && iLoop < m_iTableLength;
		iLoop++)
	{				
		m_XMLInclusiveIncentives.SelectTagListItem(iLoop + iStart);				
		sIncentivesDescription	= m_XMLInclusiveIncentives.GetTagText("DESCRIPTION");
		sIncentivesAmount		= m_XMLInclusiveIncentives.GetTagText("AMOUNT");
		sBenefitType			= m_XMLInclusiveIncentives.GetTagAttribute("INCENTIVEBENEFITTYPE", "TEXT");
		sNotionalAmount			= m_XMLInclusiveIncentives.GetTagText("NOTIONALAMOUNT");
		
		ShowInclusiveIncentiveRow(iLoop+1, sIncentivesDescription,sIncentivesAmount, sBenefitType, sNotionalAmount);
	}
	if(m_sApplicationMode == "Mortgage Calculator")
	{
		<% /* select All rows in the table */ %>
		scInclusiveTable.setAllRowsSelected();
		scInclusiveTable.DisableTable();
	}
	else setSelectedInclusive();												
}
		
function ShowInclusiveIncentiveRow(nIndex, sIncentivesDescription, sIncentivesAmount, sBenefitType, sNotionalAmount)
{			 					
	scScreenFunctions.SizeTextToField(tblInclusiveIncentivesTable.rows(nIndex).cells(0),
											sIncentivesDescription);
	scScreenFunctions.SizeTextToField(tblInclusiveIncentivesTable.rows(nIndex).cells(1),
											sBenefitType);											
	scScreenFunctions.SizeTextToField(tblInclusiveIncentivesTable.rows(nIndex).cells(2), 
											sIncentivesAmount);
	scScreenFunctions.SizeTextToField(tblInclusiveIncentivesTable.rows(nIndex).cells(3),
											sNotionalAmount);
}

function SetScreenOnReadOnly()
{
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnOK.disabled = true;
		frmScreen.btnCancel.focus();
	}
}		
		
function frmScreen.btnCancel.onclick()
{
	m_XMLInclusiveIncentives = null;
	m_XMLOutgoingTypes	= null;
	window.returnValue	= null;			
	window.close();
}
		
function frmScreen.btnOK.onclick()
{
	var sReturn = new Array();						
	sReturn[0] = IsChanged();
	BuildChosenIncentivesArray();
	sReturn[1] = CalculateTotalIncentives();
	sReturn[2] = CreateIncentivesXML();			
	window.returnValue	= sReturn;			
	window.close();
}

function BuildChosenIncentivesArray()
{
	<% // Build arrays of chosen elements for both Inclusive and exclusive items
	%>
	m_ChosenIncArray = new Array();
	for(var iRow=0 ; iRow<tblInclusiveIncentivesTable.rows.length ; iRow++)
	{
		if(tblInclusiveIncentivesTable.rows(iRow).cells(4).innerText == "Yes")
			m_ChosenIncArray[m_ChosenIncArray.length] = iRow + scInclusiveTable.getOffset() ;
	}

	m_ChosenExcIncentive = null;

	for(var iRow=0 ; iRow<tblExclusiveIncentivesTable.rows.length ; iRow++)
	{
		if(tblExclusiveIncentivesTable.rows(iRow).cells(4).innerText == "Yes")
		{
			m_ChosenExcIncentive = iRow + scExclusiveTable.getOffset() ;
			break ;	
		}
	}
}

function CreateIncentivesXML()
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sXMLString;
	var bHaveIncentives = false;
	
	var ListTag = XML.CreateActiveTag("MORTGAGEINCENTIVELIST");
	
	for (var iLoop = 0; iLoop < m_ChosenIncArray.length ;iLoop++)
	{
		XML.ActiveTag = ListTag;
		XML.CreateActiveTag("MORTGAGEINCENTIVE");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber);
		XML.CreateTag("LOANCOMPONENTSEQUENCENUMBER", m_sLoanComponentSeqNum);
		m_XMLInclusiveIncentives.SelectTagListItem(m_ChosenIncArray[iLoop]-1);
		XML.CreateTag("TYPE", "1");
		//DPF 22/10/02 - CPWP1 - Added in Incentive GUID

		XML.CreateTag("INCENTIVEGUID", m_XMLInclusiveIncentives.GetTagText("INCENTIVEGUID"));
		XML.CreateTag("DESCRIPTION", m_XMLInclusiveIncentives.GetTagText("DESCRIPTION"));
		XML.CreateTag("INCENTIVEAMOUNT", m_XMLInclusiveIncentives.GetTagText("AMOUNT"));
		XML.CreateTag("INCENTIVEBENEFITTYPE", m_XMLInclusiveIncentives.GetTagText("INCENTIVEBENEFITTYPE"));
		XML.CreateTag("NOTIONALAMOUNT", m_XMLInclusiveIncentives.GetTagText("NOTIONALAMOUNT"));
		bHaveIncentives = true;
	}

	if(m_ChosenExcIncentive != null)
	{
		XML.ActiveTag = ListTag;
		XML.CreateActiveTag("MORTGAGEINCENTIVE");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber);
		XML.CreateTag("LOANCOMPONENTSEQUENCENUMBER", m_sLoanComponentSeqNum);
		m_XMLExclusiveIncentives.SelectTagListItem(m_ChosenExcIncentive -1);								
		XML.CreateTag("TYPE", "2");
		//DPF 22/10/02 - CPWP1 - Added in Incentive GUID
		XML.CreateTag("INCENTIVEGUID", m_XMLExclusiveIncentives.GetTagText("INCENTIVEGUID"));
		XML.CreateTag("DESCRIPTION", m_XMLExclusiveIncentives.GetTagText("DESCRIPTION"));
		XML.CreateTag("INCENTIVEAMOUNT", m_XMLExclusiveIncentives.GetTagText("AMOUNT"));
		XML.CreateTag("INCENTIVEBENEFITTYPE", m_XMLExclusiveIncentives.GetTagText("INCENTIVEBENEFITTYPE"));
		XML.CreateTag("NOTIONALAMOUNT", m_XMLExclusiveIncentives.GetTagText("NOTIONALAMOUNT"));

		bHaveIncentives = true;
	}
	if(bHaveIncentives)sXMLString = XML.XMLDocument.xml;				
	else sXMLString = "";
	XML = null;
	return(sXMLString);
}

function CalculateTotalIncentives()
{
	var ExclusiveArray = new Array();
	var InclusiveArray = new Array();

	var sTotal = 0.0;
			
	ExclusiveArray = scExclusiveTable.getArrayofRowsSelected();			
	for (var iLoop = 0; iLoop < ExclusiveArray.length ;iLoop++)
	{								
		m_XMLExclusiveIncentives.SelectTagListItem(ExclusiveArray[iLoop]-1);				
		var sIncentivesAmount = m_XMLExclusiveIncentives.GetTagText("AMOUNT");
		sTotal += parseFloat(sIncentivesAmount);
	}											
			
	InclusiveArray = scInclusiveTable.getArrayofRowsSelected();
	for (var iLoop = 0; iLoop < InclusiveArray.length ;iLoop++)
	{
		m_XMLInclusiveIncentives.SelectTagListItem(InclusiveArray[iLoop]-1);
		var sIncentivesAmount = m_XMLInclusiveIncentives.GetTagText("AMOUNT");
		sTotal += parseFloat(sIncentivesAmount);
	}
	
	return sTotal;
}
function SetChosenIncentives(AllIncentivesXML)
{
<%	/* Set up 2 vars. One holds the row number of the already selected
	   Exlusive incentive and the other holds the row numbers of the 
	   already selected inclusive incentives in an array. These are then used
	   to select the correct rows every time the list box is manipulated
	   by the user - the contents being changed dynamically */
%>	m_ChosenIncArray = new Array();
	
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(m_sInIncentivesXML);
	XML.SelectTag(null, "MORTGAGEINCENTIVELIST");
	XML.CreateTagList("MORTGAGEINCENTIVE");
	for(var iLoop = 0; iLoop < XML.ActiveTagList.length; iLoop++)
	{
		XML.SelectTagListItem(iLoop);
		var sType = XML.GetTagText("TYPE");
		var sDesc = XML.GetTagText("DESCRIPTION");
		var sAmount = XML.GetTagText("INCENTIVEAMOUNT");
		
		if(sType == "1") // INCLUSIVE
		{
			if(AllIncentivesXML.SelectTag(null,"INCLUSIVEINCENTIVELIST") != null)
			{
				AllIncentivesXML.CreateTagList("INCENTIVE");
				var bFound = false;
				for(var iCount = 0; iCount < AllIncentivesXML.ActiveTagList.length && bFound == false; iCount++)
				{
					AllIncentivesXML.SelectTagListItem(iCount);
					if( sDesc == AllIncentivesXML.GetTagText("DESCRIPTION") &&
					    parseFloat(sAmount) == parseFloat(AllIncentivesXML.GetTagText("AMOUNT")) )
					{
						bFound = true;
						m_ChosenIncArray[m_ChosenIncArray.length] = iCount + 1;
					}
				}
			}
		}
		else if(sType == "2")  // EXCLUSIVE. There can be only one.
		{
			
			if(AllIncentivesXML.SelectTag(null,"EXCLUSIVEINCENTIVELIST") != null)
			{
				AllIncentivesXML.CreateTagList("INCENTIVE");
				var bFound = false;
				for(var iCount = 0; iCount < AllIncentivesXML.ActiveTagList.length && bFound == false; iCount++)
				{
					AllIncentivesXML.SelectTagListItem(iCount);
					if( sDesc == AllIncentivesXML.GetTagText("DESCRIPTION") &&
					    sAmount == AllIncentivesXML.GetTagText("AMOUNT") )
					{
						bFound = true;
						m_ChosenExcIncentive = iCount + 1;
					}
				}
			}
		}
	}
	XML = null;
}
function GetIncentiveDetails()
{
	var XMLIncentives = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
			
	if(m_sApplicationMode == "Quick Quote")
	{
		XMLIncentives.CreateRequestTagFromArray(m_sRequestArray, null);
		XMLIncentives.CreateActiveTag("MORTGAGEPRODUCT");

		XMLIncentives.CreateTag("MORTGAGEPRODUCTCODE",m_sMortgageProductCode);
		XMLIncentives.CreateTag("STARTDATE",m_sMortgageProductStartDate);
		XMLIncentives.CreateTag("LOANAMOUNT",m_sLoanAmount);
					
		XMLIncentives.RunASP(document, "GetIncentiveDetails.asp");
	}
	else
	{
		//DPF 17/10/02 - CPWP1 - build new request block for new method FindAvailableIncentives
		var m_SubQuoteXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	
		m_SubQuoteXML.LoadXML(m_sSubQuoteXML);
	
		m_SubQuoteXML.ActiveTag = null;
		m_SubQuoteXML.SelectTag(null, "MORTGAGESUBQUOTE");
	
		m_SubQuoteXML.ActiveTag = null;
		m_SubQuoteXML.SelectTag(null, "VALUATIONFEE");
	
		m_SubQuoteXML.ActiveTag = null;
		m_SubQuoteXML.SelectTag(null, "APPLICATIONFACTFIND");
	
		var m_sPurchasePrice = m_SubQuoteXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE");
		var m_sValuationType = m_SubQuoteXML.GetTagText("VALUATIONTYPE");
		var m_sLocation = m_SubQuoteXML.GetTagText("PROPERTYLOCATION");
		var m_sTypeOfApplication = m_SubQuoteXML.GetTagText("TYPEOFAPPLICATION");

		XMLIncentives.CreateRequestTagFromArray(m_sRequestArray, null);
		XMLIncentives.CreateActiveTag("MORTGAGESUBQUOTE");
		XMLIncentives.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		XMLIncentives.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XMLIncentives.CreateTag("APPLICATIONDATE",m_sApplicationDate); // JD BMIDS816
		XMLIncentives.CreateTag("MORTGAGESUBQUOTENUMBER", m_sMortgageSubQuoteNumber);
		if (m_sLoanComponentSeqNum != "" && m_sLoanComponentSeqNum != null)
		{
			XMLIncentives.CreateTag("LOANCOMPONENTSEQUENCENUMBER", m_sLoanComponentSeqNum);
		}
		XMLIncentives.CreateTag("MORTGAGEPRODUCTCODE",m_sMortgageProductCode);
		XMLIncentives.CreateTag("STARTDATE",m_sMortgageProductStartDate);
		XMLIncentives.CreateTag("LOANAMOUNT",m_sLoanAmount);
		XMLIncentives.CreateTag("LOCATION", m_sLocation);
		XMLIncentives.CreateTag("TYPEOFAPPLICATION", m_sTypeOfApplication);
		XMLIncentives.CreateTag("TYPEOFVALUATION", m_sValuationType);
		XMLIncentives.CreateTag("PURCHASEPRICE", m_sPurchasePrice);
	
		XMLIncentives.RunASP(document, "FindAvailableIncentives.asp");
		
		//END OF CHANGE FOR CPWP1
	}
	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XMLIncentives.CheckResponse(ErrorTypes);			

	if (XMLIncentives.IsResponseOK())
	{
		var InclusiveIncentiveTagList;
		var ExclusiveIncentiveTagList;
		var iInclusiveIncentiveNumRows = 0;
		var iExclusiveIncentiveNumRows = 0;
				
		m_XMLInclusiveIncentives.ActiveTag = null;
		m_XMLExclusiveIncentives.ActiveTag = null;

		SetChosenIncentives(XMLIncentives);
									
		InclusiveIncentiveTagList = XMLIncentives.SelectTag(null,"INCLUSIVEINCENTIVELIST");			

		if (InclusiveIncentiveTagList != null)
		{
			m_XMLInclusiveIncentives.ActiveTagList = XMLIncentives.CreateTagList("INCENTIVE");
			iInclusiveIncentiveNumRows = m_XMLInclusiveIncentives.ActiveTagList.length;			
		}
			
		ExclusiveIncentiveTagList = XMLIncentives.SelectTag(null,"EXCLUSIVEINCENTIVELIST");			
		if (ExclusiveIncentiveTagList != null)
		{
			m_XMLExclusiveIncentives.ActiveTagList = XMLIncentives.CreateTagList("INCENTIVE");
			iExclusiveIncentiveNumRows = m_XMLExclusiveIncentives.ActiveTagList.length;
		}				
		if (iExclusiveIncentiveNumRows > 0) 
		{
			scExclusiveTable.initialiseTable(tblExclusiveIncentivesTable, 0, "",ShowExclusiveIncentives,m_iTableLength,iExclusiveIncentiveNumRows);					
			ShowExclusiveIncentives(0);			
		}
		if (iInclusiveIncentiveNumRows > 0) 
		{
			<% /* enable select button */ %>
			scInclusiveTable.initialiseTable(tblInclusiveIncentivesTable, 0, "",ShowInclusiveIncentives,m_iTableLength,iInclusiveIncentiveNumRows);
			<% /* scInclusiveTable.EnableMultiSelectTable();  */ %>
			ShowInclusiveIncentives(0);
		}
	}
			
	XMLIncentives = null;		
}

function setSelectedInclusive()
{
	if(m_sApplicationMode == "Mortgage Calculator")scInclusiveTable.setAllRowsSelected();
	else
	{
		<%	// highlight all incentives in the listbox that match those in
			// m_sInIncentivesXML - use the array set up previously 
		%>	
		for(var iLoop = 0; iLoop < m_ChosenIncArray.length; iLoop++)
		{
			var bFound = false;
			for(var iRow = 0; iRow < tblInclusiveIncentivesTable.rows.length ; iRow++)
			{
				if(m_ChosenIncArray[iLoop] == iRow + scInclusiveTable.getOffset())
				{
					scScreenFunctions.SizeTextToField(tblInclusiveIncentivesTable.rows(iRow).cells(4),"Yes");
					break ;	
				}
			}
		}
	}
	
	<% //  Set No for the rest 
	%>
	for(var iRow = 0; iRow < tblInclusiveIncentivesTable.rows.length ; iRow++)
	{
		if (tblInclusiveIncentivesTable.rows(iRow).cells(0).innerText != " " && 
			tblInclusiveIncentivesTable.rows(iRow).cells(4).innerText == " ")
			scScreenFunctions.SizeTextToField(tblInclusiveIncentivesTable.rows(iRow).cells(4),"No");
	}
	
}

function setSelectedExclusive()
{
	for(var iRow = 0; iRow < tblExclusiveIncentivesTable.rows.length ; iRow++)
	{
		if(m_ChosenExcIncentive == iRow + scExclusiveTable.getOffset())
			scScreenFunctions.SizeTextToField(tblExclusiveIncentivesTable.rows(iRow).cells(4),"Yes");
		else
		{
			if (iRow > 0 && tblExclusiveIncentivesTable.rows(iRow).cells(0).innerText != " ") 
				scScreenFunctions.SizeTextToField(tblExclusiveIncentivesTable.rows(iRow).cells(4),"No");
		}
	}
}

function spnInclusiveIncentivesTable.ondblclick()
{
	var iRow ;

	iRow = scInclusiveTable.getRowSelected()
	if ((iRow != null) && (iRow >= 0))
	{	   
		if(tblInclusiveIncentivesTable.rows(iRow).cells(4).innerText == "Yes")
			scScreenFunctions.SizeTextToField(tblInclusiveIncentivesTable.rows(iRow).cells(4),"No");
		else
			scScreenFunctions.SizeTextToField(tblInclusiveIncentivesTable.rows(iRow).cells(4),"Yes");
	}
}

function spnExclusiveIncentivesTable.ondblclick()
{
	var iRow ;
	var bOtherRowsAdded = false;

	iRow = scExclusiveTable.getRowSelected()
	if ((iRow != null) && (iRow >= 0))
	{	   
		if(tblExclusiveIncentivesTable.rows(iRow).cells(4).innerText == "Yes")
			scScreenFunctions.SizeTextToField(tblExclusiveIncentivesTable.rows(iRow).cells(4),"No");
		else
		{
			<%// Only one row can have 'Yes' - so check whether any other row 
			  // has got 'Yes' before cahnging the current row   
			%>
			for(var iCurrentRow = 0; iCurrentRow < tblExclusiveIncentivesTable.rows.length ; iCurrentRow++) 
			{
				if(iRow != iCurrentRow && tblExclusiveIncentivesTable.rows(iCurrentRow).cells(4).innerText == "Yes")
					bOtherRowsAdded = true ;
			}
			if(bOtherRowsAdded) alert('Only one exclusive incentive can be added.' );
			else scScreenFunctions.SizeTextToField(tblExclusiveIncentivesTable.rows(iRow).cells(4),"Yes");
		}
	}
}
-->
</script>
</body>
</html>




