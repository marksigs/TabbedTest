<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<HTML>
	<HEAD>
		<title></title>
		<% /*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DC015.asp
Copyright:     Copyright © 2006 Vertex Financial Services
Description:   Intermediary Search Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog  Date     Description
JLD		04/11/99	Created
JLD		10/12/1999	DC/013 - Added scrolling to listbox
					DC/015 - Changed the look of the gui
					DC/021 - set address format ok
AD		30/01/2000	Rework
AY		14/02/00	Change to msgButtons button types
AY		29/03/00	New top menu/scScreenFunctions change
BG		17/05/00	SYS0752 Removed Tooltips
CL		05/03/01	SYS1920 Read only functionality added
STB		20/02/02	SYS3777 Results list cleared when no results.
LD		23/05/02	SYS4727 Use cached versions of frame functions
TW      09/10/2002  SYS5115 Modified to incorporate client validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
MO		29/09/2002	BMIDS00502 - INWP1, BM061, Change to introducer search to show level introducer 
						id's and modify search to call BM's system.
MO		14/10/2002	BMIDS00663 - Fixed bug where the value of direct/indirect was passed and not the 
						validation type, put in place the 'proper' introducer services.  Also made
						the change so that omAdmin was called rather than omInt
MO		14/11/2002	BMIDS00936 - Changed the town label to Town / Postcode
HMA     18/09/2003  BM0063     - Amend HTML text for radio buttons
KRW     24/05/2004  BMIDS762   - FSA Validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history:

Prog	Date		Description
GHun	26/07/2005	MAR10 Changed FSA id label to not be customer specific
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific history:

Prog	Date		Description
IK		05/04/2006	EP15 - extensive re-write
IK		03/05/2006	EP447 - alert when no records found 
INR		31/10/2006	EP2_12	 Changes for	New Intermediary structure
INR		05/01/2007	EP2_543 Changes in validation types used and include searching for individual introducers 
IK		29/01/2007	EP2_863 add 'Trading As' to list
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
		<META content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="stylesheet.css" type="text/css" rel="stylesheet">
	</HEAD>
	<BODY>
		<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
		<OBJECT id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px"
			tabIndex="-1" type="text/x-scriptlet" height="1" width="1" data="scClientFunctions.asp"
			viewastext>
		</OBJECT>
		<%/* SCRIPTLETS */%>
		<SCRIPT language="JScript" src="validation.js" type="text/javascript"></SCRIPT>
		<SPAN id="spnListScroll"><SPAN style="LEFT: 315px; POSITION: absolute; TOP: 425px">
				<OBJECT id="scScrollTable" style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" tabIndex="-1"
					type="text/x-scriptlet" data="scTableListScroll.asp" viewastext>
				</OBJECT>
			</SPAN></SPAN>
		<FORM id="frmToDC010" style="DISPLAY: none" action="DC010.asp" method="post">
		</FORM>
		<% /* Span to keep tabbing within this screen */ %>
		<SPAN id="spnToLastField" tabIndex="0"></SPAN><!-- Specify Screen Layout Here -->
		<FORM id="frmScreen" validate="onchange" mark>
			<DIV class="msgGroup" id="divBackground" style="MARGIN-TOP: 60px; MARGIN-LEFT: 10px; WIDTH: 604px; HEIGHT: 360px">
				<DIV class="msgLabel" style="LEFT: 0px; POSITION: relative; TOP: 0px"><SPAN style="MARGIN-LEFT: 8px"><STRONG>Search</STRONG></SPAN>
				</DIV>
				<DIV class="msgLabel" style="LEFT: 0px; PADDING-TOP: 8px; POSITION: relative; TOP: 0px"><SPAN style="MARGIN-LEFT: 8px">FSA 
						Ref #</SPAN> <SPAN style="LEFT: 80px; POSITION: absolute"><INPUT class="msgTxt" id="txtFSAId" style="WIDTH: 100px" type="text" maxLength="12">
					</SPAN><SPAN style="LEFT: 280px; POSITION: absolute">Individual Forename</SPAN> <SPAN style="LEFT: 400px; POSITION: absolute">
						<INPUT  class="msgReadOnly" readonly="true" id="txtForeName" style="WIDTH: 150px" type="text" maxLength="40" name="txtForeName">
					</SPAN>
				</DIV>
				<DIV class="msgLabel" style="MARGIN-TOP: 8px; LEFT: 0px; POSITION: relative; TOP: 0px"><SPAN>&nbsp;</SPAN>
					<SPAN style="LEFT: 280px; POSITION: absolute; TOP: 0px">Individual Surname</SPAN> <SPAN style="LEFT: 400px; POSITION: absolute">
						<INPUT  class="msgReadOnly" readonly="true" id="txtSurname" style="WIDTH: 150px" type="text" maxLength="40" name="txtSurname">
					</SPAN>
				</DIV>
				<DIV class="msgLabel" style="MARGIN-TOP: 8px; LEFT: 0px; POSITION: relative; TOP: 0px"><SPAN>&nbsp;</SPAN>
					<SPAN style="LEFT: 280px; POSITION: absolute; TOP: 0px">Company Name</SPAN> <SPAN style="LEFT: 400px; POSITION: absolute">
						<INPUT class="msgTxt" id="txtName" style="WIDTH: 150px" type="text" maxLength="40" name="txtName">
					</SPAN>
				</DIV>
				<DIV class="msgLabel" style="MARGIN-TOP: 8px; LEFT: 0px; POSITION: relative; TOP: 0px"><SPAN>&nbsp;</SPAN>
					<SPAN style="LEFT: 280px; POSITION: absolute; TOP: 0px">Town</SPAN> <SPAN style="LEFT: 400px; POSITION: absolute">
						<INPUT class="msgTxt" id="txtTown" style="WIDTH: 150px" type="text" maxLength="40" name="txtTown">
					</SPAN>
				</DIV>
				<DIV class="msgLabel" style="MARGIN-TOP: 8px; LEFT: 0px; POSITION: relative; TOP: 0px"><SPAN>&nbsp;</SPAN>
					<SPAN style="LEFT: 280px; POSITION: absolute; TOP: 0px">Post Code</SPAN> <SPAN style="LEFT: 400px; POSITION: absolute">
						<INPUT class="msgTxt" id="txtPostCode" style="TEXT-TRANSFORM: uppercase; WIDTH: 150px"
							type="text" maxLength="10"> </SPAN>
				</DIV>
				<DIV style="MARGIN-TOP: 12px; LEFT: 0px; POSITION: relative; TOP: 0px">
					<SPAN style="LEFT: 400px; WIDTH: 150px; POSITION: absolute; TOP: 0px; TEXT-ALIGN: right">
						<INPUT class="msgButton" id="btnSearch" style="WIDTH: 60px" type="button" value="Search"
							name="btnSearch"></SPAN>
				</DIV>
				<DIV class="msgLabel" style="LEFT: 0px; MARGIN-LEFT: 0px; POSITION: relative; TOP: 0px"><SPAN class="msgLabel" style="MARGIN-LEFT: 8px">Include 
						in search? <SPAN style="LEFT: 126px; POSITION: absolute; TOP: -3px">
							<SELECT class="msgCombo" id="cboIncludeInSearch" style="WIDTH: 180px" name="cboIncludeInSearch" onchange="cboIncludeInSearch_OnChange()"
								menusafe="true">
							</SELECT>
						</SPAN></SPAN>
				</DIV>
				<DIV><SPAN></SPAN>
				</DIV>
				<DIV class="msgLabel" style="MARGIN-TOP: 8px; LEFT: 0px; POSITION: relative; TOP: 0px"><SPAN style="MARGIN-LEFT: 8px"><STRONG>Introducer 
							List</STRONG></SPAN>
				</DIV>
				<DIV id="spnIntermedListTable" style="MARGIN-TOP: 8px; LEFT: 0px; MARGIN-LEFT: 4px; POSITION: relative; TOP: 0px">
					<TABLE class="msgTable" id="tblIntermedList" cellSpacing="0" cellPadding="0" width="596"
						border="0">
						<TR id="rowTitles">
							<TD class="TableHead" style="text-align:center" width="10%">Introducer Id</TD>
							<TD class="TableHead" width="10%">FSA Ref</TD>
							<TD class="TableHead" width="12%">Name</TD>
							<td class="TableHead" style="text-align:center" width="12%">Trading As</td>
							<TD class="TableHead" width="24%">Address</TD>
							<TD class="TableHead" width="9%">Type</TD>
							<TD class="TableHead" width="14%">Listing Status</TD>
							<TD class="TableHead" width="9%">Status</TD>
						</TR>
						<TR id="row01">
							<TD class="TableTopLeft" style="text-align:center" width="10%">&nbsp;</TD>
							<TD class="TableTopCenter" width="10%">&nbsp;</TD>
							<TD class="TableTopCenter" width="12%">&nbsp;</TD>
							<td class="TableTopCenter" width="12%">&nbsp;</td>
							<TD class="TableTopCenter" width="24%">&nbsp;</TD>
							<TD class="TableTopCenter" width="9%">&nbsp;</TD>
							<TD class="TableTopCenter" width="14%">&nbsp;</TD>
							<TD class="TableTopRight">&nbsp;</TD>
						</TR>
						<TR id="row02">
							<TD class="TableLeft" style="text-align:center" width="10%">&nbsp;</TD>
							<TD class="TableCenter" width="10%">&nbsp;</TD>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="24%">&nbsp;</td>
							<TD class="TableCenter" width="9%">&nbsp;</TD>
							<TD class="TableCenter" width="14%">&nbsp;</TD>
							<TD class="TableRight">&nbsp;</TD>
						</TR>
						<TR id="row03">
							<td class="TableLeft" style="text-align:center" width="10%">&nbsp;</td>
							<td class="TableCenter" width="10%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="24%">&nbsp;</td>
							<TD class="TableCenter" width="9%">&nbsp;</TD>
							<TD class="TableCenter" width="14%">&nbsp;</TD>
							<TD class="TableRight">&nbsp;</TD>
						</TR>
						<TR id="row04">
							<td class="TableLeft" style="text-align:center" width="10%">&nbsp;</td>
							<td class="TableCenter" width="10%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="24%">&nbsp;</td>
							<TD class="TableCenter" width="9%">&nbsp;</TD>
							<TD class="TableCenter" width="14%">&nbsp;</TD>
							<TD class="TableRight">&nbsp;</TD>
						</TR>
						<TR id="row05">
							<td class="TableLeft" style="text-align:center" width="10%">&nbsp;</td>
							<td class="TableCenter" width="10%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="24%">&nbsp;</td>
							<TD class="TableCenter" width="9%">&nbsp;</TD>
							<TD class="TableCenter" width="14%">&nbsp;</TD>
							<TD class="TableRight">&nbsp;</TD>
						</TR>
						<TR id="row06">
							<td class="TableLeft" style="text-align:center" width="10%">&nbsp;</td>
							<td class="TableCenter" width="10%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="24%">&nbsp;</td>
							<TD class="TableCenter" width="9%">&nbsp;</TD>
							<TD class="TableCenter" width="14%">&nbsp;</TD>
							<TD class="TableRight">&nbsp;</TD>
						</TR>
						<TR id="row07">
							<td class="TableLeft" style="text-align:center" width="10%">&nbsp;</td>
							<td class="TableCenter" width="10%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="24%">&nbsp;</td>
							<TD class="TableCenter" width="9%">&nbsp;</TD>
							<TD class="TableCenter" width="14%">&nbsp;</TD>
							<TD class="TableRight">&nbsp;</TD>
						</TR>
						<TR id="row08">
							<td class="TableLeft" style="text-align:center" width="10%">&nbsp;</td>
							<td class="TableCenter" width="10%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="24%">&nbsp;</td>
							<TD class="TableCenter" width="9%">&nbsp;</TD>
							<TD class="TableCenter" width="14%">&nbsp;</TD>
							<TD class="TableRight">&nbsp;</TD>
						</TR>
						<tr id="row09">
							<td class="TableLeft" style="text-align:center" width="10%">&nbsp;</td>
							<td class="TableCenter" width="10%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="12%">&nbsp;</td>
							<td class="TableCenter" width="24%">&nbsp;</td>
							<TD class="TableCenter" width="9%">&nbsp;</TD>
							<TD class="TableCenter" width="14%">&nbsp;</TD>
							<TD class="TableRight">&nbsp;</TD>
						</tr>
						<TR id="row10">
							<TD class="TableBottomLeft" style="text-align:center" width="10%">&nbsp;</TD>
							<TD class="TableBottomCenter" width="10%">&nbsp;</TD>
							<td class="TableBottomCenter" width="12%">&nbsp;</td>
							<td class="TableBottomCenter" width="12%">&nbsp;</td>
							<td class="TableBottomCenter" width="24%">&nbsp;</td>
							<TD class="TableBottomCenter" width="9%">&nbsp;</TD>
							<TD class="TableBottomCenter" width="14%">&nbsp;</TD>
							<TD class="TableBottomRight">&nbsp;</TD>
						</TR>
					</TABLE>
				</DIV>
			</DIV>
		</FORM> <!-- Main Buttons -->
		<DIV id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 480px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --></DIV>
		<% /* Span to keep tabbing within this screen */ %>
		<SPAN id="spnToFirstField" tabIndex="0"></SPAN><!-- #include FILE="fw030.asp" -->  <!-- File containing field attributes (remove if not required) --> <!--  #include FILE="attribs/DC015attribs.asp" -->  <!-- Specify Code Here -->
		<SCRIPT language="JScript" type="text/javascript">
<!--
// JScript Code
var m_sMetaAction = null;
var IntermedXML = null;
var m_iTableLength = 10;
var scScreenFunctions;
var m_blnReadOnly = false;
var m_sDC010Data = "";
var m_DC010XML = null;
var strDirectIndirectBusiness = "";
var strDirectIndirectBusinessValType = "";
var XMLIntroducerSearchType = null;
var XMLListingStatus = null;
var XMLFirmStatus = null;
var m_isIndivIntroducer = false;

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	// Make the required buttons available on the bottom of the screen
	// (see msgButtons.asp for details)
	var sButtonList = new Array("Submit","Cancel","Another");
	ShowMainButtons(sButtonList);

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	RetrieveContextData();
	
	//This function is contained in the field attributes file (remove if not required)
	SetMasks();
	Validation_Init();
	
	InitialseScreenValues();
	
	frmScreen.txtFSAId.focus();
	DisableMainButton("Submit");
	DisableMainButton("Another");
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC015");
	
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

<% // MO - 29/09/2002 - BMIDS00502 %>
function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	<% /*get the introducer levels and settings from DC010 */ %>
	m_sDC010Data = scScreenFunctions.GetContextParameter(window,"idXML",null);
}

<% // MO - 29/09/2002 - BMIDS00502 %>
function InitialseScreenValues() {
	// Grab the XML from DC010 and put it on the screen, also default the level
	// TODO : m_DC010XML
	FW030SetTitles("Introducer Search","DC015",scScreenFunctions);
	PopulateCombos();
}

<% /*EP2_12 New combo to populate */ %>
function PopulateCombos()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("IntroducerSearchType","ListingStatus","FirmStatus");
	
	if(XML.GetComboLists(document,sGroupList))
	{
		XMLIntroducerSearchType = XML.GetComboListXML("IntroducerSearchType");
		XMLListingStatus = XML.GetComboListXML("ListingStatus");
		XMLFirmStatus = XML.GetComboListXML("FirmStatus");
		
		var blnSuccess = true;
		//EP2_543 populate all from combo
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboIncludeInSearch,XMLIntroducerSearchType,true);
	}

	XML = null;
}

<% // MO - 29/09/2002 - BMIDS00502 %>
function SetFocusToEmptyField()
{
	if( IsStringEmpty(frmScreen.txtFSAId.value) )
	{
		frmScreen.txtFSAId.value = "";
		frmScreen.txtFSAId.focus();
	}
	else if( IsStringEmpty(frmScreen.txtName.value))
	{
		frmScreen.txtName.value = "";
		frmScreen.txtName.focus();
	}
	else if( IsStringEmpty(frmScreen.txtTown.value))
	{
		frmScreen.txtTown.value = "";
		frmScreen.txtTown.focus();
	}
	else
	{
		<% // default to FSAid %>
		frmScreen.txtFSAId.value = "";
		frmScreen.txtFSAId.focus();
	}	
}

function IsStringEmpty(strString)
{
	var bStringIsEmpty = false;
	var ssArray = strString.split(" ");
	var nLength = ssArray.length;
	var nIndex = 0;
	while(ssArray[nIndex] == "" && nIndex < nLength)
		nIndex++;

	if(ssArray[nIndex] == null || ssArray[nIndex] == "*")
		bStringIsEmpty = true;

	return(bStringIsEmpty);
}

function frmScreen.btnSearch.onclick()
{
	<% //check that all fields are occupied %>
	if( IsStringEmpty(frmScreen.txtFSAId.value) &&
		IsStringEmpty(frmScreen.txtName.value) &&
		IsStringEmpty(frmScreen.txtPostCode.value) &&
		IsStringEmpty(frmScreen.txtTown.value) &&
		!m_isIndivIntroducer)
	{
		alert("You must enter search criteria");
		SetFocusToEmptyField();
	}
	else
	{
		<% //all present and correct, populate the listbox %>
		if(FindIndividualIntermediary())
		{
			scScrollTable.setRowSelected(1);
			EnableMainButton("Submit");
			btnSubmit.focus();
		}
		<% /* MO - 14/10/2002 - BMIDS00663 No longer needed as the error is returned by the BM Introducer system 
		else
		{
			scScrollTable.clear();
			alert("Unable to find any intermediaries for your search criteria. Please amend the search criteria and/or wildcard your search");
			DisableMainButton("Submit");
			DisableMainButton("Another");
			frmScreen.txtFSAId.focus();
		} */ %>
	}	
}

<% // MO - 29/09/2002 - BMIDS00502 %>
function PopulateListBox(nStart)
{
	IntermedXML.ActiveTag = null;
	IntermedXML.CreateTagList("INTERMEDIARY");
	var iNumberOfIntermediaries = IntermedXML.ActiveTagList.length;

	if(iNumberOfIntermediaries > 0)
	{
		scScrollTable.initialiseTable(tblIntermedList, 0, "", ShowList, m_iTableLength, iNumberOfIntermediaries);
		ShowList(nStart);
		return true;
	} 
	else 
	{
		<% // EP2_543 %>
		alert("Unable to find any introducers for your search criteria. Please amend the criteria.");
		return false;
	}
}

<% // MO - 29/09/2002 - BMIDS00502 %>
function ShowList(nStart)
{
	var iCount;
	scScrollTable.clear();
	for (iCount = 0; iCount < IntermedXML.ActiveTagList.length && iCount < m_iTableLength; iCount++)
	{
		IntermedXML.ActiveTag = null;
		IntermedXML.CreateTagList("INTERMEDIARY");  
		IntermedXML.SelectTagListItem(iCount + nStart);

		var sType;
		var listingStatus;
		var introducerId;
		var nodValueID;
		var sStatus; 
		var sPrincipalFirmID = IntermedXML.GetAttribute("PRINCIPALFIRMID");
		var sARFirmID = IntermedXML.GetAttribute("ARFIRMID");
		var sClubAssocID = IntermedXML.GetAttribute("CLUBNETWORKASSOCIATIONID");
		var sIntroducerID = IntermedXML.GetAttribute("INTRODUCERID");
		var fsaRef =  IntermedXML.GetAttribute("FSAREFNUMBER");
		var sName =  IntermedXML.GetAttribute("NAME");
		<% // EP2_863 %>		
		var sTradingAs =  IntermedXML.GetAttribute("TRADINGAS");
		var sValidateType =  IntermedXML.GetAttribute("LISTSTATUSVALIDATION");
		nodValueID=XMLListingStatus.selectSingleNode(".//LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='"+ sValidateType +"']");
		if (nodValueID)
		{
			listingStatus = nodValueID.selectSingleNode(".//VALUENAME").text;
		}
		var sStatusValue =  IntermedXML.GetAttribute("STATUS");
		nodValueID =  XMLFirmStatus.selectSingleNode(".//LISTENTRY[VALUEID='"+ sStatusValue +"']");
		if (nodValueID)
		{
			sStatus = nodValueID.selectSingleNode(".//VALUENAME").text;
		}
		sValidateType =  IntermedXML.GetAttribute("TYPEVALIDATION");
		nodValueID=XMLIntroducerSearchType.selectSingleNode(".//LISTENTRY[VALIDATIONTYPELIST/VALIDATIONTYPE='"+ sValidateType +"']");
		if (nodValueID)
		{
			sType = nodValueID.selectSingleNode(".//VALUENAME").text;
		}
		var sPostcode = IntermedXML.GetAttribute("POSTCODE");
		var sFlatNumber = IntermedXML.GetAttribute("ADDRESSLINE1");
		var sHouseName = IntermedXML.GetAttribute("ADDRESSLINE2");
		var sHouseNumber = IntermedXML.GetAttribute("ADDRESSLINE3");
		var sStreet = IntermedXML.GetAttribute("ADDRESSLINE4");
		var sDistrict = IntermedXML.GetAttribute("ADDRESSLINE5");
		var sTown = IntermedXML.GetAttribute("ADDRESSLINE6");
		var sCounty = IntermedXML.GetAttribute("ADDRESSLINE7");

		<% //create an intelligent address line %>
		var sAddress = "";
		if(sFlatNumber != "") sAddress += sFlatNumber + ",";
		if(sHouseName != "") sAddress += sHouseName + ",";
		if(sHouseNumber != "") sAddress += sHouseNumber + " ";
		if(sStreet != "") sAddress += sStreet + ",";
		if(sDistrict != "") sAddress += sDistrict + ",";
		if(sTown != "") sAddress += sTown + ",";
		if(sCounty != "") sAddress += sCounty + ",";
		if(sPostcode != "") sAddress += sPostcode;

		if(sPrincipalFirmID&&(sPrincipalFirmID.length > 0))
		{
			tblIntermedList.rows(iCount+1).setAttribute("PrincipalFirmID", sPrincipalFirmID);
			introducerId = sPrincipalFirmID;
		}
			
		if(sARFirmID&&(sARFirmID.length > 0))
		{
			tblIntermedList.rows(iCount+1).setAttribute("ARFirmID", sARFirmID);
			introducerId = sARFirmID;
		}
		if(sClubAssocID&&(sClubAssocID.length > 0))
		{
			tblIntermedList.rows(iCount+1).setAttribute("ClubAssocID", sClubAssocID);
			introducerId = sClubAssocID;
		}
		if(sIntroducerID&&(sIntroducerID.length > 0))
		{	
			tblIntermedList.rows(iCount+1).setAttribute("IntroducerID", sIntroducerID);
			introducerId = sIntroducerID;
		}


		<% // Add to the search table %>
		scScreenFunctions.SizeTextToField(tblIntermedList.rows(iCount+1).cells(0),introducerId);
		scScreenFunctions.SizeTextToField(tblIntermedList.rows(iCount+1).cells(1),fsaRef);
		scScreenFunctions.SizeTextToField(tblIntermedList.rows(iCount+1).cells(2),sName);
		<% // EP2_863 %>		
		scScreenFunctions.SizeTextToField(tblIntermedList.rows(iCount+1).cells(3),sTradingAs);
		scScreenFunctions.SizeTextToField(tblIntermedList.rows(iCount+1).cells(4),sAddress);
		scScreenFunctions.SizeTextToField(tblIntermedList.rows(iCount+1).cells(5),sType);
		scScreenFunctions.SizeTextToField(tblIntermedList.rows(iCount+1).cells(6),listingStatus);
		scScreenFunctions.SizeTextToField(tblIntermedList.rows(iCount+1).cells(7),sStatus);

		listingStatus = "";
		sStatus = "";
		sType = "";
	}
}

function FindIndividualIntermediary()
{
	var bSuccess = true;

	IntermedXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	IntermedXML.CreateRequestTag(window);
<%/*Request is done as an operation to allow us to add CRITERIA for the postproc xslt
<REQUEST OPERATION="FindIntroducers" RETURNREQUEST="Y" USERAUTHORITYLEVEL="String" UNITID="String" CHANNELID="String" MACHINEID="String" PASSWORD="String" USERID="String">
	<OPERATION SCHEMA_NAME="FindIntroducersWSSchema" CRUD_OP="READ" ENTITY_REF="PRINCIPALFIRM" COMBOLOOKUP="Y">
		<PRINCIPALFIRM BROKERINDICATOR="1" POSTCODE="GL50*" />
	</OPERATION>
</REQUEST>*/%>
	var entityRef;
	var searchRef;
	var packageInd;
	var brokerType;
	var refFSA;
	var postCode;
	var xe;
	var xs;
	var xe2;
	var xn = IntermedXML.XMLDocument.documentElement;
	xn.setAttribute("OPERATION","FindIntroducers");
	xn.setAttribute("RETURNREQUEST","Y");
	xn.setAttribute("postProcRef","FindIntroducersResponseGUI.xslt" );
	if ((frmScreen.txtFSAId.value.length > 0) && (!m_isIndivIntroducer))
	{
		<% /*  EP2_543 Need to do a search of PrincipalFirm and ArFirm if we have an FSARef No.*/ %>
		<% /*  So make up two operations. However, we will only ever return one of them, */ %>
		<% /*  because FSARef No. is globally unique. */ %>
	
		<% /* Need to search Principal and AR Firms */ %>
		refFSA = frmScreen.txtFSAId.value
	
		<% /*  Broker D/A Firm */ %>
		xe = IntermedXML.XMLDocument.createElement("OPERATION");
		xe.setAttribute("CRUD_OP","READ");
		xe.setAttribute("SCHEMA_NAME","FindIntroducersWSSchema");
		xe.setAttribute("ENTITY_REF","PRINCIPALFIRM");
		xe.setAttribute("COMBOLOOKUP","Y");

		xe2 = IntermedXML.XMLDocument.createElement("PRINCIPALFIRM");
		xe2.setAttribute("FSAREFERENCENUMBER", refFSA);
		<% /* PACKAGERIND always 0 as packager won't have FSARef */ %>
		xe2.setAttribute("PACKAGERIND", "0");

		xe.appendChild(xe2);

		xn.appendChild(xe);	

		<% /*  Broker A/R Firm */ %>
		xe = IntermedXML.XMLDocument.createElement("OPERATION");
		xe.setAttribute("CRUD_OP","READ");
		xe.setAttribute("SCHEMA_NAME","FindIntroducersWSSchema");
		xe.setAttribute("ENTITY_REF","ARFIRM");
		xe.setAttribute("COMBOLOOKUP","Y");

		xe2 = IntermedXML.XMLDocument.createElement("ARFIRM");
		xe2.setAttribute("FSAREFERENCENUMBER", refFSA);

		xe.appendChild(xe2);

		xn.appendChild(xe);	

		xs = IntermedXML.XMLDocument.createElement("SEARCHCRITERIA");
		if((searchRef) && (searchRef.length > 0))
		{
			xs.setAttribute("SEARCHREF", searchRef);
		}
		xs.setAttribute("SCREENREF", "DC015");
		xn.appendChild(xs);

	}
	else
	{
		<% /* Otherwise use entered search criteria */ %>
		<% /*  Needs to cope with multiple validation types */ %>
		if(scScreenFunctions.IsValidationType(frmScreen.cboIncludeInSearch, "PFN"))
		{
			<% /*  Broker D/A Firm */ %>
			entityRef = "PRINCIPALFIRM";
			packageInd = "0";
			searchRef = "0";
			companyRef = "PRINCIPALFIRMNAME";
		}
		else if(scScreenFunctions.IsValidationType(frmScreen.cboIncludeInSearch, "ARF"))
		{
			<% /*  Broker A/R Firm */ %>
			entityRef = "ARFIRM";
			companyRef = "ARFIRMNAME";
		}
		else if(scScreenFunctions.IsValidationType(frmScreen.cboIncludeInSearch, "PKF"))
		{
			<% /*  Packager */ %>
			entityRef = "PRINCIPALFIRM";
			packageInd = "1";
			searchRef = "1";
			companyRef = "PRINCIPALFIRMNAME";
		}
		else if(scScreenFunctions.IsValidationType(frmScreen.cboIncludeInSearch, "PA"))
		{
			<% /*  Packaging Association */ %> 
			entityRef = "MORTGAGECLUBNETWORKASSOCIATION";
			searchRef = "1";
			packageInd = "1";
			companyRef = "COMPANYNAME";
		}
		else if(scScreenFunctions.IsValidationType(frmScreen.cboIncludeInSearch, "MC"))
		{
			<% /*  MortgageClub */ %>
			entityRef = "MORTGAGECLUBNETWORKASSOCIATION";
			searchRef = "0";
			packageInd = "0";
			companyRef = "COMPANYNAME";
		}	<% /* EP2_543 Add search for Individuals */ %>
		else if(scScreenFunctions.IsValidationType(frmScreen.cboIncludeInSearch, "IBAR"))
		{
			<% /*  Individual AR Broker */ %>
			entityRef = "INTRODUCERDETAILS";
			searchRef = "0";
			companyRef = "COMPANYNAME";
			brokerType = "AR";
		}
		else if(scScreenFunctions.IsValidationType(frmScreen.cboIncludeInSearch, "IBDA"))
		{
			<% /*  Individual DA Broker */ %>
			entityRef = "INTRODUCERDETAILS";
			searchRef = "0";
			companyRef = "COMPANYNAME";
			brokerType = "DA";
		}
		else if(scScreenFunctions.IsValidationType(frmScreen.cboIncludeInSearch, "IP"))
		{
			<% /*  Individual Packager */ %>
			entityRef = "INTRODUCERDETAILS";
			searchRef = "1";
			companyRef = "COMPANYNAME";
			brokerType = "PA";
		}
		else
		{
			alert("Please enter search criteria.");
			return(false);	
		}
	
		xe = IntermedXML.XMLDocument.createElement("OPERATION");
		xe.setAttribute("CRUD_OP","READ");
		xe.setAttribute("SCHEMA_NAME","FindIntroducersWSSchema");
		xe.setAttribute("ENTITY_REF",entityRef);
		xe.setAttribute("COMBOLOOKUP","Y");

		xe2 = IntermedXML.XMLDocument.createElement(entityRef);
		<% //EP2_543 Individual search criteria%>
		if( m_isIndivIntroducer &&
			IsStringEmpty(frmScreen.txtName.value) &&
			IsStringEmpty(frmScreen.txtPostCode.value) &&
			IsStringEmpty(frmScreen.txtTown.value))
		{
			
			if( m_isIndivIntroducer &&
				IsStringEmpty(frmScreen.txtFSAId.value) &&
				IsStringEmpty(frmScreen.txtForeName.value) &&
				IsStringEmpty(frmScreen.txtSurname.value))
			{
				alert("You must enter search criteria");
				frmScreen.txtForeName.focus();
				return;
			}
			else
			{
				if (frmScreen.txtForeName.value.length > 0) xe2.setAttribute("FORENAME", frmScreen.txtForeName.value);
				if (frmScreen.txtSurname.value.length > 0) xe2.setAttribute("SURNAME", frmScreen.txtSurname.value);
				if (frmScreen.txtFSAId.value.length > 0)
				{
					xe2.setAttribute("FSAREFERENCENUMBER", frmScreen.txtFSAId.value);
				}
				else
				{
					<% /*  Individual Search with no further refinement than possibly name and brokertype, then we
					need to get everything by all postcodes &  brokertype and then PostProcRef will match name */ %>
					xe2.setAttribute("POSTCODE", "*");
				}
			}
		}
		else
		{
			if (frmScreen.txtForeName.value.length > 0) xe2.setAttribute("FORENAME", frmScreen.txtForeName.value);
			if (frmScreen.txtSurname.value.length > 0) xe2.setAttribute("SURNAME", frmScreen.txtSurname.value);
			if (refFSA && (refFSA.length > 0)) xe2.setAttribute("FSAREFERENCENUMBER", refFSA);
			if (frmScreen.txtName.value.length > 0) xe2.setAttribute(companyRef, frmScreen.txtName.value);
			if (frmScreen.txtTown.value.length > 0) xe2.setAttribute("TOWN", frmScreen.txtTown.value);
			if (frmScreen.txtPostCode.value.length > 0) xe2.setAttribute("POSTCODE", frmScreen.txtPostCode.value);
			if (packageInd && (packageInd.length > 0)) xe2.setAttribute("PACKAGERIND", packageInd);
		}
		<% /*  EP2_543 Need brokerType for individuals*/ %>
		if (brokerType && (brokerType.length > 0)) xe2.setAttribute("BROKERTYPE", brokerType);

		xe.appendChild(xe2);

		xn.appendChild(xe);	

		xs = IntermedXML.XMLDocument.createElement("SEARCHCRITERIA");
		if((searchRef) && (searchRef.length > 0))
		{
			xs.setAttribute("SEARCHREF", searchRef);
		}
		xs.setAttribute("SCREENREF", "DC015");
		xn.appendChild(xs);
	}

	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			IntermedXML.RunASP(document, "omCRUDIf.asp");
			break;
		default: // Error
			SetErrorResponse();
	}

	var bSuccess = IntermedXML.IsResponseOK();

	if (bSuccess == true)
	{
		<% /* Populate Introducer details */ %>
		if (PopulateListBox(0) == false)
		{
			<% /* No Introducers found */ %>
			scScrollTable.clear();
			bSuccess = false;
		}
	}
	else
	{
		scScrollTable.clear();
	}

	return(bSuccess);
}

<% // EP2_21 Save to new ApplicationIntroducer table %>
function btnSubmit.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "LoadIntermediaryFromDC015");
	var nRowSelected =  scScrollTable.getRowSelected();
		
	if (nRowSelected != -1) 
	{
		//	ik_debug
		//	stopIT();
	
		<% //Create to APPLICATIONINTRODUCER %>
		var sPrincipalFirmID = tblIntermedList.rows(nRowSelected).getAttribute("PrincipalFirmID");
		var sARFirmID = tblIntermedList.rows(nRowSelected).getAttribute("ARFirmID");
		var sClubAssocID = tblIntermedList.rows(nRowSelected).getAttribute("ClubAssocID");
		var sIntroducerID = tblIntermedList.rows(nRowSelected).getAttribute("IntroducerID");

		var sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber");
		var sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber");

		IntermedXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		IntermedXML.CreateRequestTag(window);
		var xn = IntermedXML.XMLDocument.documentElement;
		xn.setAttribute("CRUD_OP","CREATE");
		var xe = IntermedXML.XMLDocument.createElement("APPLICATIONINTRODUCER");
		xe.setAttribute("APPLICATIONNUMBER",sApplicationNumber);
		xe.setAttribute("APPLICATIONFACTFINDNUMBER",sApplicationFactFindNumber);
		if(sPrincipalFirmID && (sPrincipalFirmID.length > 0))
			xe.setAttribute("PRINCIPALFIRMID",sPrincipalFirmID);
		if(sClubAssocID && (sClubAssocID.length > 0))
			xe.setAttribute("CLUBNETWORKASSOCID",sClubAssocID);
		if(sARFirmID && (sARFirmID.length > 0))
			xe.setAttribute("ARFIRMID",sARFirmID);
		if(sIntroducerID && (sIntroducerID.length > 0))
			xe.setAttribute("INTRODUCERID",sIntroducerID);
		xn.appendChild(xe);

		IntermedXML.RunASP(document, "omCRUDIf.asp");

		if(!IntermedXML.IsResponseOK()) return;
	}	

	frmToDC010.submit();
}

function btnCancel.onclick()
{
	<% //Set idMetaAction so that DC010 loads any saved settings %>
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "LoadIntermediaryFromDC015");
	frmToDC010.submit();
}

<% //EP2_543 %>
function frmScreen.cboIncludeInSearch.onchange()
{
	var selIndex = frmScreen.cboIncludeInSearch.selectedIndex;
	if(selIndex != -1 && scScreenFunctions.IsOptionValidationType(frmScreen.cboIncludeInSearch,selIndex,"I"))
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtForeName");
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtSurname");
		frmScreen.txtForeName.className = "msgTxt";
		frmScreen.txtSurname.className = "msgTxt";
		m_isIndivIntroducer = true;
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtForeName");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtSurname");
		frmScreen.txtForeName.value = "";
		frmScreen.txtSurname.value = "";
		m_isIndivIntroducer = false;
	}
}

-->
		</SCRIPT>
	</BODY>
</HTML>
