<%@ LANGUAGE="JSCRIPT" %>
<HTML>
<!--

Workfile:      DC013.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
KRW     14/06/2004  New Screen for Regulation Buy to Let Detail BMIDS776 REG025
KRW     23/06/2004  Added Call to CreateNewProperty when none exists BMIDS776 REG025
KRW     25/06/2004  Included m_sFamilyMemberValType in array passed back to DC010 BMIDS776 REG025
INR     08/07/2004  Validate FamilyMemberCombo onOK
KRW     13/07/2004  BMIDS766   Changed Ok and OnCancel processing to return a null value to DC010 if no changes
GHun	11/08/2004	BMIDS840 Fixed violation of primary key
JD		21/09/2004	BMIDS887 Readonly processing
GHun	22/09/2004	BMIDS887 Also save SpecialScheme, RegulationIndicator and LandUsage
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History

Prog    Date		AQR		Description
INR		22/01/2007	EP2_677 E2ECR35 review.
AShaw	20/02/2007	EP2_1447 - Was Not applying stylesheet.
-->

<HEAD>
<LINK href="stylesheet.css" rel="STYLESHEET" type="text/css">

<title>DC013 - Regulation Buy To Let Detail <!-- #include file="includes/TitleWhiteSpace.asp" --></title>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
</HEAD>

<BODY>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>

<% /* Scriptlets */ %>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" year4 validate="onchange" mark>
<div style="HEIGHT: 115px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 400px" class="msgGroup">
	<tr>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<td>If buy to let will you or a relative occupy<br>the property now or in the future?</td>
		<td><span style="LEFT: 270px; POSITION: absolute; TOP: 6px">
			<input id="optOccupyPropertyYes" name="OccupyPropertyGroup" type="radio" value="1"><label for="optOccupyPropertyYes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 330px; POSITION: absolute; TOP: 6px">
			<input id="optOccupyPropertyNo" name="OccupyPropertyGroup" type="radio" value="0"><label for="optOccupyPropertyNo" class="msgLabel">No</label>
		</span></td>
	</span>
	</tr>
	<tr>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 65px" class="msgLabel">
	<td>	Family Member </td>
	<td>	<span style="LEFT: 155px; POSITION: absolute; TOP: -3px">
			<select id="cboFamilyMember" maxlength="30" style="POSITION: absolute; WIDTH: 150px" align="right" class="msgCombo"></select>
		</span></td>
	</span>
	</tr>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 115px">
		<input id="btnOk" value="OK" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
	
	<span style="LEFT: 100px; POSITION: absolute; TOP: 115px">
		<input id="btnCancel" value="Cancel" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
</div>
</form><!-- #include FILE="attribs/DC013attribs.asp" -->

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var scScreenFunctions;
var m_sApplicationNumber;
var m_sOccupyProperty = "";
var m_sFamilyMember = "";
var m_sFamilyMemberValType = "";
var m_sNewPropertyExists;
var m_sLandUsage="";
var m_sIsRegulated;
var m_sNewRegulatedInd;
var m_sRegulatedId = "";
var m_sNonRegulatedId = "";
var m_sOtherMemberId;
var m_aArgArray = null;
var m_BaseNonPopupWindow = null;
var scClientScreenFunctions;
<% /* BMIDS 776 */ %>
var m_bValidComboSelection = true;
<% /* BMIDS887 GHun */ %>
var m_sAppFactFindNumber = "";
var m_sRegulationInd = "";
var m_sSpecialScheme = "";
var m_sLandUsage = "";
<% /* BMIDS887 End */ %>

<% /**** Events *****/ %>
function frmScreen.btnCancel.onclick()
{
	var saReturn = new Array();
	
	//saReturn[0] = "Cancel"
	saReturn = null; // BMIDS766 KW 13/07/2004
	window.returnValue	= saReturn ;
	window.close();
}

function frmScreen.btnOk.onclick()
{
	var saReturn = new Array();
	
	if(IsChanged())
	{
		if (frmScreen.onsubmit())
		{	
			<% /* EP2_677 */ %>
			m_sOccupyProperty = scScreenFunctions.GetRadioGroupValue(frmScreen, "OccupyPropertyGroup") ;
			m_sFamilyMember = frmScreen.cboFamilyMember.value ;
	
		/*	if (m_sRegulatedInd=="1")
			{
				if (m_sFamilyLet == "0")
				{
					m_sNewRegulatedInd = "0" ;
					alert('This Buy To Let application was set as Regulated, this has been changed to Non-Regulated because the let is not to a family let');
				}		
				else
				{
					if (frmScreen.cboFamilyMember.value == m_sOtherMemberId)
					{
						m_sNewRegulatedInd = "0" ;
						alert('This Buy To Let application was set as Regulated, this has been changed to Non-Regulated because the related member of family is set as Other');
					}
					else m_sNewRegulatedInd = "1";
				}
			}
			else
			{
				if(m_sLandUsage == "1" && m_sFamilyLet == "1" && frmScreen.cboFamilyMember.value != m_sOtherMemberId)
				{
					m_sNewRegulatedInd = "1" ;
					alert('This Buy To Let application was set as Non-Regulated, this has been changed to Regulated because this is a family let');
				}
				else m_sNewRegulatedInd = "0" ;
			}
		*/
		
			m_sFamilyMemberValType = scScreenFunctions.GetComboValidationType(frmScreen, frmScreen.cboFamilyMember.id)		

			<% /* BMIDS776 Check that we have a valid combo selection*/ %>
			<% /* EP2_677 */ %>
			if (m_sOccupyProperty == "1")
			{
				if (m_sFamilyMember > 0)
				{
					m_bValidComboSelection = true
				}
				else
				{
					m_bValidComboSelection = false
					alert('A valid Family Member must be selected from the List');
				}
			}

			<% /* BMIDS840 GHun Only SaveNewPropertyData after checking m_bValidComboSelection */ %>
			if (m_bValidComboSelection == true)
			{
				if (SaveNewPropertyData())
				{ <% /* SR 05/04/2004 : BBG121 */ %>
					saReturn[0] = "OK" ;
					<% /* EP2_677 */ %>
					saReturn[1] = m_sOccupyProperty ;
					saReturn[2] = m_sFamilyMember ;
					saReturn[3] = (m_sNewRegulatedInd=='1') ? m_sRegulatedId : m_sNonRegulatedId
					saReturn[4] = m_sFamilyMemberValType;

					window.returnValue	= saReturn ;  <% /* SR 05/04/2004 : BBG121 - End */ %>
					window.close();
				}
			}
		}
	}
	else
	{
		//saReturn[0] = "" ;
		saReturn = null; // BMIDS766 KW 13/07/2004
		window.returnValue	= saReturn ;
		window.close();
	}
}

function window.onload()
{
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();

	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	m_aArgArray = sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sApplicationNumber = m_aArgArray[0];
	m_sNewPropertyExists = m_aArgArray[1];
	<% /* EP2_677 */ %>
	m_sOccupyProperty		 = m_aArgArray[2];
	m_sFamilyMember		 = m_aArgArray[3];
	m_sIsRegulated		 = m_aArgArray[4];
	m_sLandUsage		 = m_aArgArray[5];
	m_sReadOnly			 = m_aArgArray[6];
	<% /* BMIDS887 GHun */ %>
	m_sRegulationInd	 = m_aArgArray[7];
	m_sSpecialScheme	 = m_aArgArray[8];
	m_sAppFactFindNumber = m_aArgArray[9];
	m_sLandUsage		 = m_aArgArray[10];
	<% /* BMIDS887 End */ %>
	
	Validation_Init();
	
	InitialiseScreen();
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	ClientPopulateScreen();
}

<% /* EP2_677 */ %>
function frmScreen.optOccupyPropertyYes.onclick()
{
	frmScreen.cboFamilyMember.disabled = false ;
}

<% /* EP2_677 */ %>
function frmScreen.optOccupyPropertyNo.onclick()
{
	frmScreen.cboFamilyMember.disabled = true ;
	frmScreen.cboFamilyMember.value = "" ;
	<% /* BMIDS776 Combo selection will be valid */ %>
	m_bValidComboSelection = true;
}

<% /**** FUNCTIONS *******/ %>
function InitialiseScreen()
{
	PopulateCombos();
	if(m_sNewPropertyExists == "1")
	{
		<% /* EP2_677 use the new radiogroup */ %>
		scScreenFunctions.SetRadioGroupValue(frmScreen, "OccupyPropertyGroup", m_sOccupyProperty);	
		frmScreen.cboFamilyMember.value = m_sFamilyMember ;
		
		if(m_sOccupyProperty == "") m_sOccupyProperty = "0" <% /* default to NO  */ %>
	}
	else m_sOccupyProperty = "0" <% /* default to NO  */ %>

	if(m_sOccupyProperty == "0") frmScreen.cboFamilyMember.disabled = true ;
	
	if (m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnOk.disabled = true ;
	}
}

function PopulateCombos()
{
	var XMLCombos = null;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	var sGroupList = new Array("FamilyLetMember", "RegulationIndicator");
	if(XML.GetComboLists(document,sGroupList))
	{
		XMLCombos = XML.GetComboListXML("FamilyLetMember");

		var saTemp = new Array();
		saTemp = XML.GetComboIdsForValidation("FamilyLetMember", "O", XMLCombos);
		m_sOtherMemberId = saTemp[0];

		var blnSuccess = true;
		blnSuccess = blnSuccess & XML.PopulateComboFromXML(document,frmScreen.cboFamilyMember,XMLCombos,true);
		
		XMLCombos = XML.GetComboListXML("RegulationIndicator");
		saTemp = XML.GetComboIdsForValidation("RegulationIndicator", "R", XMLCombos);
		m_sRegulatedId = saTemp[0];
		
		saTemp = XML.GetComboIdsForValidation("RegulationIndicator", "N", XMLCombos);
		m_sNonRegulatedId = saTemp[0];
		
		if(!blnSuccess)
		{
			alert('Failed to populate combos');
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			frmScreen.btnOk.disabled = true ;
		}
	}
}

function SaveNewPropertyData()
{
	var bSuccess = true;
	var XMLCombos = null;
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();

	<% /* Check whether to create new record or to update the existing one */ %>
	if(IsChanged())
	{
		<% /* BMIDS887 GHun Update the regulation indicator if necessary */ %>
		if (m_sLandUsage == "1")
		{
			<% /* EP2_677 */ %>
	 		 if(m_sOccupyProperty == "1" )
			 {	  
			     if(m_sFamilyMember != m_sOtherMemberId)
			    	m_sRegulationInd = m_sRegulatedId;
 			     else
		            m_sRegulationInd = m_sNonRegulatedId;
		     }
		     else
				m_sRegulationInd = m_sNonRegulatedId;
		}
		<% /* BMIDS887 End */ %>
			
		var xmlRequest = XML.CreateRequestTagFromArray(m_aArgArray[11], "CreateNewPropertyGeneral");

		XML.CreateActiveTag("NEWPROPERTY");
		<% /* BMIDS840 GHun Compare m_sNewPropertyExists to "1", not 1 as it is a string*/ %>
		if(m_sNewPropertyExists == "1")
		<% /* BMIDS887 GHun */ %>
			XML.SetAttribute("NEWPROPERTYEXISTS", "1");
		else
			XML.SetAttribute("NEWPROPERTYEXISTS", "0");
		<% /* BMIDS887 End */ %>
			
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNumber);
		<% /* EP2_677 */ %>
		XML.CreateTag("OCCUPYPROPERTY", m_sOccupyProperty);
		XML.CreateTag("FAMILYMEMBER", m_sFamilyMember);
		
		<% /* BMIDS887 GHun */ %>
		XML.ActiveTag = xmlRequest;
		XML.CreateActiveTag("APPLICATIONFACTFIND");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNumber);
		XML.CreateTag("REGULATIONINDICATOR", m_sRegulationInd);
		XML.CreateTag("SPECIALSCHEME", m_sSpecialScheme);
		XML.CreateTag("LANDUSAGE", m_sLandUsage);
		<% /* BMIDS887 End */ %>
		
		switch (ScreenRules())
		{
			case 1: // Warning
			case 0: // OK
				<% /* BMIDS887 GHun */ %>
				XML.RunASP(document,"SaveNewPropertyAndUpdateFactFind.asp");
				<% /*				
				if(m_sNewPropertyExists == "1")
					XML.RunASP(document,"UpdateNewPropertyGeneral.asp");
				else
					XML.RunASP(document,"CreateNewPropertyGeneral.asp");
				BMIDS887 End */ %>
				
				break;
			default: // Error
				XML.SetErrorResponse();
		}
		
		bSuccess = XML.IsResponseOK();
		XML = null;

		<% /* BMIDS840 GHun */ %>
		if (bSuccess)
			m_sNewPropertyExists = "1";
	}
	
	return(bSuccess);
}

-->
</script>
</BODY>
</HTML>
