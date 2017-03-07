<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
Workfile:      ra040.asp
Copyright:     Copyright © 2000 Marlborough Stirling

Description:   Address Targeting Control screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
INR		12/02/2004	Initial revision BMIDS682
INR		19/03/2004	BMIDS730 Address Targeting Processing
SAB		27/10/2005	MARS245	 Udated for revised call to omTMBO and revised XML response
SAB		31/10/2005	MARS346  Now detects if a repocess or resubmit call was made
GHun	13/11/2005	MAR374   Fixed ActiveTag is null error
INR		27/11/2005	MAR621   Use SeqNumber, instead of Addressindicator
RF		13/02/2006  MAR1243  Address targeting now implemented in DC370
LDM		25/05/2006	EP591	 Epsom has its own version of xmlcreditcheck etc in omtm
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
	<title>Address Targeting Control</title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex="-1" 
type="text/x-scriptlet" width="1" viewastext></object>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<% /* Scriptlets - remove any which are not required */ %>

<% /* BMIDS730 Specify Forms Here */ %>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>

<span id="spnRulesScroll" style="LEFT: 310px; POSITION: absolute; TOP: 448px; VISIBILITY: hidden">
<object data="scTableListScroll.asp" id="scRulesScroll" 
style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex="-1" 
type="text/x-scriptlet" viewastext></OBJECT>
</span>

<% /* Specify Screen Layout Here */ %>
<div id="divStatus" style="HEIGHT: 100px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 614px" class="msgGroup">
	<table id="tblStatus" width="604" height="100" border="0" cellspacing="0" cellpadding="0">
		<tr align= "center"><td id="colStatus" align="middle" class="msgLabelWait" >Please Wait...<br>Running Address Targeting</td></tr>
	</table>
</div>

<div id="divBackground" style="HEIGHT: 434px; LEFT: 10px; POSITION: absolute; TOP: 60px; VISIBILITY: hidden; WIDTH: 604px" class="msgGroup">
	<form id="frmScreen" style="VISIBILITY: hidden">
	<div id="divRules" style="LEFT: 4px; POSITION: absolute; TOP: 4px; WIDTH: 604px">
		<span style="LEFT: 0px; POSITION: absolute; TOP: 0px" class="msgLabel">
			<strong>Address Targeting Control</strong>
		</span>

		<div id="spnTable" style="LEFT: 0px; POSITION: absolute; TOP: 16px">
			<table id="tblTarget" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
			<tr id="rowTitles">	<td width="25%" class="TableHead">Address Type&nbsp;</td>
								<td width="25%" class="TableHead">Targeting Address Type&nbsp;</td>
								<td width="40%" class="TableHead">Customer(s) Name&nbsp;</td>
								<td width="10%" class="TableHead">Block Seq No&nbsp;</td></tr>
			<tr id="row01">		<td width="25%" class="TableTopLeft">&nbsp;</td>			<td width="25%" class="TableTopCenter">	&nbsp;</td>			<td width="40%" class="TableTopCenter">&nbsp;</td>	<td class="TableTopRight">&nbsp;</td></tr>
			<tr id="row02">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>			<td width="40%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row03">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>			<td width="40%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row04">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>			<td width="40%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row05">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>			<td width="40%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row06">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>			<td width="40%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row07">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>			<td width="40%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row08">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>			<td width="40%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row09">		<td width="25%" class="TableLeft">&nbsp;</td>				<td width="25%" class="TableCenter">&nbsp;</td>			<td width="40%" class="TableCenter">&nbsp;</td>		<td class="TableRight">&nbsp;</td></tr>
			<tr id="row10">		<td width="25%" class="TableBottomLeft">&nbsp;</td>		<td width="25%" class="TableBottomCenter">&nbsp;</td>		<td width="40%" class="TableBottomCenter">&nbsp;</td>	<td class="TableBottomRight">&nbsp;</td></tr>
			</table>
		</div>
		
		<span style="LEFT: 0px; POSITION: absolute; TOP: 388px">
			<input id="btnEdit" value="Edit" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		
	</div>
	</div>

<% /* Not used in cost modelling. Not yet supported. Maybe one day
		<span style="TOP: 0px; LEFT: 132px; POSITION: ABSOLUTE">
			<input id="btnPrint" value="Print" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		<span style="TOP: 0px; LEFT: 198px; POSITION: ABSOLUTE">
			<input id="btnDetails" value="Details" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
 */ %>
	</form>

<form id="frmGo" method="post" style="VISIBILITY: hidden"></form>
<form id="frmGoBack" method="post" action="mn060.asp" style="VISIBILITY: hidden"></form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 500px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --></div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/ra040Attribs.asp" -->

<% /* Specify Code Here */ %>
<script language="JScript" type="text/javascript">
<!--
var m_sMetaAction = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sActivityId = null;
var scScreenFunctions;
var AddressTargetSummaryXML = null;
var m_sListXML = null;
var m_sAddressTargetXML = null;
var returnXML = null;
var ccn1XML = null;
//BMIDS730
var m_sDeclareXML = null;
var declareXML = null;
//MAR245
var CCResponseXML = null;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	var sButtonList = new Array("Submit", "Cancel");
	ShowMainButtons(sButtonList);

	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Address Targeting Control","RA040",scScreenFunctions);
	
	window.setTimeout(CompleteInitialisation, 0);
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}	

function CompleteInitialisation()	
{	
	scScreenFunctions.ShowCollection(frmScreen);
	
	RetrieveContextData();	

	AddressTargetSummaryXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	// MAR245
	CCResponseXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	CCResponseXML.LoadXML(m_sAddressTarget);
	CCResponseXML.SelectTag(null, "RESPONSE");
	CCResponseXML.SelectTag(CCResponseXML.ActiveTag, "TARGETINGDATA");
	AddressTargetSummaryXML.LoadXML(CCResponseXML.ActiveTag.xml);
	// MAR245
	returnXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ccn1XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ccn1XML.CreateActiveTag("SAVECCN");
	returnXML.CreateActiveTag("REQUEST");
	returnXML.CreateTag("TARGETREQUEST",null);

	AddressTargetSummaryXML.SelectTag(null,"CCN1LIST");
	ccn1XML.ActiveTag.appendChild(AddressTargetSummaryXML.ActiveTag.cloneNode(true));
	
	//BMIDS730 Only ever want to update this on an initial pass
	declareAddressXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AddressTargetSummaryXML.SelectTag(null,"DECLAREADDRESSLIST");
	declareAddressXML = AddressTargetSummaryXML.ActiveTag.cloneNode(true);
	
	// Get data and populate the tables
	InitTable();
	
	divStatus.style.visibility = "hidden";
	divBackground.style.visibility = "visible";
	frmScreen.style.visibility = "visible";
	spnRulesScroll.style.visibility = "visible";
}
	
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window, "idMetaAction", null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window, "idApplicationNumber", null);
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", null);
	m_sActivityId = scScreenFunctions.GetContextParameter(window, "idStageId", null);
	m_sListXML =  scScreenFunctions.GetContextParameter(window, "idXML", null);
	m_sAddressTarget =  scScreenFunctions.GetContextParameter(window, "idAddressTarget", null);
}

function InitTable()
{
	scRulesScroll.initialiseTable(tblTarget,0,"",PopulateAddTargetTable,24,PopulateAddTargetTable(0));
}

function PopulateAddTargetTable(nStart)
{
	var strAddrIndicator = "";
	var strAddrResolved = "";
	var nRecordCount = 0;
	//BMIDS730
	var strDisplayName = "";
	var strSeqNumber = "";
	var strPattern = "";
	var strAddrType = "";
	
	//want to search it all, so set ActiveTag to null
	AddressTargetSummaryXML.ActiveTag = null;
	AddressTargetSummaryXML.CreateTagList("TARGETDISPLAY");
			
	for (var nLoop=0; AddressTargetSummaryXML.SelectTagListItem(nStart + nLoop); nLoop++)
	{
		strAddrIndicator = 	AddressTargetSummaryXML.GetAttribute("ADDRESSINDICATOR")					
		strAddrResolved = AddressTargetSummaryXML.GetAttribute("ADDRESSRESOLVED");
		//BMIDS730
		strSeqNumber = AddressTargetSummaryXML.GetAttribute("SEQNUMBER");
		strPattern = "CCN1[@BLOCKSEQNUMBER='" + strSeqNumber + "']"
		ccn1XML.SelectTag(null,"CCN1LIST");
		ccn1XML.SelectTag(null,strPattern);
		strDisplayName = ccn1XML.GetAttribute("NAME");

		//Only display it if it hasn't been resolved
		if(strAddrResolved == "N")
		{
			//need a seperate recordcount so it fills the
			//table from the top once some have been resolved
			//MAR621 Use SeqNumber, instead of Addressindicator
			nRecordCount = nRecordCount + 1
			if((strSeqNumber == "01")||(strSeqNumber == "02"))
				strAddrType = "Current"; 
			else if	((strSeqNumber == "11")||(strSeqNumber == "12"))
				strAddrType = "Previous";
			else if	((strSeqNumber == "21")||(strSeqNumber == "22"))
				strAddrType = "Previous, Previous";
			scScreenFunctions.SizeTextToField(tblTarget.rows(nRecordCount).cells(0), strAddrType);
			scScreenFunctions.SizeTextToField(tblTarget.rows(nRecordCount).cells(1), AddressTargetSummaryXML.GetAttribute("ADDRESSTYPE"));
			scScreenFunctions.SizeTextToField(tblTarget.rows(nRecordCount).cells(2), strDisplayName);
			scScreenFunctions.SizeTextToField(tblTarget.rows(nRecordCount).cells(3), strSeqNumber);
		}
	}
	//BMIDS730
	if(nRecordCount == 0)
	{
		frmScreen.btnEdit.disabled = true;
		EnableMainButton("Submit");
	}
	else
	{
		frmScreen.btnEdit.disabled = false;
		DisableMainButton("Submit");
	}
	
	return AddressTargetSummaryXML.ActiveTagList.length;
}

function btnSubmit.onclick()
{
	//BMIDS730 Disable the submit button and Alert the user we are retargeting
	btnSubmit.blur();
	btnSubmit.style.cursor = "wait";
	window.status = "Retargeting...";
	
	<% /*BMIDS730  Call the makeCreditCheckCall after a timeout to allow the 
	cursor time to change.  */ %>
	window.setTimeout(makeCreditCheckCall, 0)
	//var bSuccess = makeCreditCheckCall()
	
	<% /*BMIDS730  Finish off further processing in makeCreditCheckCall 
	if (bSuccess)
	{
		RefreshScreen();
	}
	else
	{
		frmGoBack.action = scScreenFunctions.GetContextParameter(window, "idReturnScreenId", "TM030") + ".asp"
		scScreenFunctions.SetContextParameter(window,"idProcessContext",null);
		scScreenFunctions.SetContextParameter(window,"idXML",null);
		frmGoBack.submit();
	} */ %>
	
}

function btnCancel.onclick()
{
	//BMIDS730 clear down Context's and return to TM030
	scScreenFunctions.SetContextParameter(window,"idXML","");
	scScreenFunctions.SetContextParameter(window,"idAddressTarget","");
	
	<% /* MAR1243 
	frmToTM030.submit(); */ %>
	
	frmGoBack.action = scScreenFunctions.GetContextParameter(window, "idReturnScreenId", "TM030") + ".asp"
	scScreenFunctions.SetContextParameter(window,"idProcessContext",null);
	scScreenFunctions.SetContextParameter(window,"idXML",null);
	frmGoBack.submit();
}
  
function frmScreen.btnEdit.onclick()
{
	var strAddrIndicator = "";
	var strPattern = "";
	var sReturn = null;
	var sPopupScreen = null;
	var ArrayArguments = new Array();
	var sCustomerName;
	
	// close window and return selected screen data
	var iCurrentRow = scRulesScroll.getRowSelected();
	if(iCurrentRow == -1)
	{
		alert("Please select an Ambiguous/Invalid Address");
		return;
	}

	//BMIDS730 Disable the edit button until we repopulate it and decide if
	//it needs to be enabled
	frmScreen.btnEdit.disabled = true;

	// Get the BlockSeqNo corresponding to the currently selected rule
	var sAddressType = tblTarget.rows(iCurrentRow).cells(1).innerText;
	var sBlockSeqNo = tblTarget.rows(iCurrentRow).cells(3).innerText;
	//BMIDS730 Get the name to be displayed
	var strPattern = "CCN1[@BLOCKSEQNUMBER='" + sBlockSeqNo + "']";
	ccn1XML.SelectTag(null,"CCN1LIST");
	ccn1XML.SelectTag(null,strPattern);
	sCustomerName = ccn1XML.GetAttribute("NAME");
	
	if(sAddressType == "Invalid")
		sPopupScreen = "RA042.asp";
	else
		sPopupScreen = "RA041.asp";

	ArrayArguments[0] = sCustomerName;
	ArrayArguments[1] = sBlockSeqNo;
	ArrayArguments[2] = m_sAddressTarget;
	//BMIDS730
	ArrayArguments[3] = declareAddressXML.xml
	//MAR245
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ArrayArguments[4] = XML.GetGlobalParameterBoolean(document,"ExpAddTargettingEnableDeclared");
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, sPopupScreen, ArrayArguments, 635, 400);
	if (sReturn != null)
	{
		//Update the AddressTargeting Control screen.
		var bOK = sReturn[0];
		if (bOK)
		{
			UpdateScreen(sBlockSeqNo);
			var sTargetXML = sReturn[1];
			if(sTargetXML != "")
			{
				AddToRequest(sTargetXML);
			}
			scRulesScroll.clear();
			InitTable();
		}
		else
		{
			//BMIDS730 Address resolution screen has been cancelled, enable edit
			frmScreen.btnEdit.disabled = false;
		}
	}
}

function UpdateScreen(sBlockSeqUpdated)
{
	var strBlockSeqCompare = null;
	//want to search it all, so set ActiveTag to null
	AddressTargetSummaryXML.ActiveTag = null;
	AddressTargetSummaryXML.CreateTagList("TARGETDISPLAY");

	for (var nLoop=0; AddressTargetSummaryXML.SelectTagListItem(nLoop); nLoop++)
	{
		strBlockSeqCompare = AddressTargetSummaryXML.GetAttribute("SEQNUMBER");					
		if(strBlockSeqCompare == sBlockSeqUpdated)
			AddressTargetSummaryXML.SetAttribute("ADDRESSRESOLVED","Y");
	}
}	

function AddToRequest(sTargetXML)
{	
	var acceptXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	acceptXML.LoadXML(sTargetXML);
	acceptXML.SelectTag(null,"ADDRESSTARGET");
	
	returnXML.SelectTag(null,"TARGETREQUEST");
	returnXML.ActiveTag.appendChild(acceptXML.ActiveTag.cloneNode(true));

}  

function makeCreditCheckCall()
{
	var blnSuccess = false;

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	var ReqTag;
	<% /* MAR346 Identify the correct operation to call */ %>
	var strOperation = "";
	var strReprocessOp = "";
	var strRescoreOp = "";
	
	CCResponseXML.SelectTag(null,"RESPONSE");
	strReprocessOp = CCResponseXML.GetAttribute("REPROCESS");
	strRescoreOp = CCResponseXML.GetAttribute("RESCORE");

	if (strReprocessOp == "1") 
		strOperation = "RUNEPSOMREPROCESSCREDITCHECK";<%/* LDM 25/5/6 EP591 */%>
	else if (strRescoreOp == "1")
		strOperation = "RUNEPSOMRESCORECREDITCHECK";<%/* LDM 25/5/6 EP591 */%>
	else
		strOperation = "RUNEPSOMCREDITCHECK";<%/* LDM 25/5/6 EP591 */%>
	
	ReqTag = XML.CreateRequestTag(window, strOperation);
	<% /* MAR346 End */ %>

	XML.SetAttribute("CREDITCHECKINVOKEDFROMACTIONBUTTON", "TRUE");
	XML.SetAttribute("ADDRESSTARGETREQ", "TRUE");

	XML.CreateActiveTag("APPLICATION");
	XML.SetAttribute("APPLICATIONNUMBER", scScreenFunctions.GetContextParameter(window, "idApplicationNumber", ""));
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", scScreenFunctions.GetContextParameter(window, "idApplicationFactFindNumber", ""));
	XML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window, "idApplicationPriority", ""));
	XML.SelectTag(null,"REQUEST");

	returnXML.SelectTag(null,"TARGETREQUEST");
	//XML.ActiveTag.appendChild(returnXML.ActiveTag.cloneNode(true));

	<%/* MAR245 - SAB - The chosen/modified addresses are placed held in the returnXML.  The TARGETTING element
		 therefore needs to be updated before it can be returned to the frontend */%>
	AddressTargetSummaryXML.SelectTag(null,"TARGETINGDATA")
	returnXML.CreateTagList("ADDRESSTARGET");
			
	for (var nLoop=0; returnXML.SelectTagListItem(nLoop); nLoop++)
	{
		var strBlockType = returnXML.GetAttribute("BLOCKTYPE");
		var strBlockSeq = returnXML.GetAttribute("BLOCKSEQNUMBER");
		var strAddrID = returnXML.GetAttribute("ID");
		var sourceNode = AddressTargetSummaryXML.ActiveTag.selectSingleNode("ADDRESSTARGET[@ID='" + strAddrID + "' and @BLOCKSEQNUMBER='" + strBlockSeq + "']");
		if (sourceNode)
		{
			// Only update the AddressResolved flag
			sourceNode.setAttribute("ADDRESSRESOLVED", "Y");
			if (strBlockType == "BUK1")
			{
				sourceNode.setAttribute("FLAT", returnXML.GetAttribute("FLAT"));
				sourceNode.setAttribute("HOUSENAME", returnXML.GetAttribute("HOUSENAME"));
				sourceNode.setAttribute("HOUSENUMBER", returnXML.GetAttribute("HOUSENUMBER"));
				sourceNode.setAttribute("STREET", returnXML.GetAttribute("STREET"));
				sourceNode.setAttribute("DISTRICT", returnXML.GetAttribute("DISTRICT"));
				sourceNode.setAttribute("TOWN", returnXML.GetAttribute("TOWN"));
				sourceNode.setAttribute("COUNTY", returnXML.GetAttribute("COUNTY"));
				sourceNode.setAttribute("POSTCODE", returnXML.GetAttribute("POSTCODE"));
			}
		}
	}
	XML.ActiveTag.appendChild(AddressTargetSummaryXML.ActiveTag.cloneNode(true));

	// Pass the CASETASK tag to get through validation etc.
	CCResponseXML.SelectTag(null,"CASETASK");
	XML.ActiveTag.appendChild(CCResponseXML.ActiveTag.cloneNode(true));

	XML.RunASP(document, "OmigaTMBO.asp");

	if (XML.IsResponseOK()) {
		
		XML.SelectTag(null, "RESPONSE");
		XML.SelectTag(XML.ActiveTag, "TARGETINGDATA"); //MAR245 Revised XML entry point
		var sIsAddrTarget = XML.GetTagText("ADDRESSTARGETING");
		if (sIsAddrTarget == "YES")
		{
			//Need to save the CCN1 info for the next request
			//ccn1XML.SelectTag(null,"CCN1LIST");
			//XML.ActiveTag.appendChild(ccn1XML.ActiveTag.cloneNode(true));

			//Address target info back on an address target request.
			//reset up all the info to be displayed
			m_sAddressTarget = XML.XMLDocument.xml

			blnSuccess = true;				
		}
		else
		{
			//the response was OK, but we have no further address targeting to do
			//return to tm030
			blnSuccess = false;	
		}
	}
	else
	{
		//the response ERRORED, but we have no further address targeting to do
		//return to tm030
		blnSuccess = false;
	}	

	//BMIDS730 Determine next screen routing
	if (blnSuccess)
	{
		RefreshScreen();
	}
	else
	{
		frmGoBack.action = scScreenFunctions.GetContextParameter(window, "idReturnScreenId", "TM030") + ".asp"
		scScreenFunctions.SetContextParameter(window,"idProcessContext",null);
		scScreenFunctions.SetContextParameter(window,"idXML",null);
		frmGoBack.submit();
	}

	//BMIDS730 retargeting is finished
	window.status = "";
	btnSubmit.style.cursor = "hand";

	/*
	AddressTargetSummaryXML.SelectTag(null,"REQHEADER");
	XML.ActiveTag.appendChild(AddressTargetSummaryXML.ActiveTag.cloneNode(true));

	//need a CASETASK tag to get through validation etc.
	XML.SelectTag(null,"REQUEST");
	AddressTargetSummaryXML.SelectTag(null,"CASETASK");
	XML.ActiveTag.appendChild(AddressTargetSummaryXML.ActiveTag.cloneNode(true));

	XML.SelectTag(null,"REQUEST");
	ccn1XML.SelectTag(null,"CCN1LIST");
	XML.ActiveTag.appendChild(ccn1XML.ActiveTag.cloneNode(true));

	XML.RunASP(document, "OmigaTMBO.asp");
	*/
}

function RefreshScreen()	
{	

	AddressTargetSummaryXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	returnXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	returnXML.CreateActiveTag("REQUEST");
	returnXML.SetAttribute("ADDRESSTARGETREQ","TRUE");
	AddressTargetSummaryXML.LoadXML(m_sAddressTarget);
	//MAR374 GHun commented out
	//AddressTargetSummaryXML.SelectTag(null,"REQHEADER");	
	//returnXML.ActiveTag.appendChild(AddressTargetSummaryXML.ActiveTag.cloneNode(true));
	//MAR374 GHun end
	AddressTargetSummaryXML.SelectTag(null,"CASETASK");
	returnXML.ActiveTag.appendChild(AddressTargetSummaryXML.ActiveTag.cloneNode(true));
	ccn1XML.SelectTag(null,"CCN1LIST");
	returnXML.ActiveTag.appendChild(ccn1XML.ActiveTag.cloneNode(true));
	returnXML.CreateTag("TARGETREQUEST",null);

	// Get data and populate the tables
	InitTable();
		
	divStatus.style.visibility = "hidden";
	divBackground.style.visibility = "visible";
	frmScreen.style.visibility = "visible";
	spnRulesScroll.style.visibility = "visible";	
}

//BMIDS730
function spnTable.ondblclick()
{ 
	if (scRulesScroll.getRowSelected() != null)
		frmScreen.btnEdit.onclick();
}

-->
</script>
</body>
</html>
