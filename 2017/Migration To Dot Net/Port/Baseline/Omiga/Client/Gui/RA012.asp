<%@ LANGUAGE="JSCRIPT" %>

<HTML>
  <HEAD>
<title>Case Assessment  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
<%
/*
Workfile:      ra012.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Risk Assessment
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:
Prog	Date		Description
SR		20/10/2005	MAR24 
HM		11/11/2005	MAR487 Error in retrieving Application Underwriting Data  
INR		23/11/2005	MAR645 Various
INR		29/11/2005	MAR725 Only save if we have UserID info
HMA     12/12/2005  MAR667 Correct display of Authority Level. Add processing for TASK button.
JD		14/01/2006	MAR1014 Set the underwriter decision if the CA decision ie Refer.
JD		02/02/2006	MAR1179 check the users loanamountmandate
DRC     10/02/2006  MAR1209 - remove stage from getriskassessment calls
HMA	    21/02/2006  MAR1291 Remove stage dependency when viewing.
PE		22/02/2006	MAR1227 SM rules have authority level in RA012 but any user can overide them.
HMA     23/02/2006  MAR1319 Further changes to display of Strategy Manager and Case Assessment.
HMA     27/02/2006  MAR1303 Change to display of Underwriter's Decision.
HMA     07/03/2006  MAR1370 Allow override of any Case Assessment and Strategy Manager.
HMA     08/03/2006  MAR1303 Further change to Underwriter's Decision.
HMA     09/03/2006  MAR1391 Add table offsets.
HMA     14/03/2006  MAR1427 Modify checks on Risk Assessment and Strategy Manager data.
SAB		27/03/2006	EP302 Retrieves the amount requested from the new loan table if a 
						  quotation does not exist.
SAB		13/04/2006	EP386 Include MAR1447 & MAR1563
PB		11/05/2006	EP529 Merged-in MAR1657 - Authority level for Strategy Manager codes not displaying correctly.
PB		24/05/2006	EP604 Fixed syntax error - replaced < with >
AW		14/09/2006	EP1103	CC78 New mode parameter for calling TM031
MAH		02/03/2007	EP2_1580 Effectively set rest of screen items to readonly when readonly flag set
AShaw	19/03/2007	EP2_1015 - Prevent text wrapping in table fields. ALL form positions are absolute.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<META content=http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4 
name=vs_targetSchema>
<META content="Microsoft Visual Studio 6.0" name=GENERATOR><LINK 
href="stylesheet.css" type=text/css rel=STYLESHEET>
  </HEAD>
<BODY>&nbsp;
<OBJECT id=scClientFunctions style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
	type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp VIEWASTEXT>
</OBJECT>

<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<SCRIPT language=JScript src="validation.js"></SCRIPT>

<OBJECT id=scScrFunctions style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
	type=text/x-scriptlet height=1 width=1 data=scScreenFunctions.asp VIEWASTEXT>
</OBJECT>
	
<OBJECT id=scXMLFunctions style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
	type=text/x-scriptlet height=1 width=1 data=scXMLFunctions.asp VIEWASTEXT>
</OBJECT>
	
<span style="LEFT: 300px; POSITION: absolute; TOP: 137px">
<OBJECT data=scTableListScroll.asp id=scScrollTable 
	style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT>
</OBJECT>
</span>
<span style="LEFT: 300px; POSITION: absolute; TOP: 388px">
<OBJECT data=scTableListScroll.asp id="scScrollTable1" 
	style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT>
</OBJECT>
</span>

<% /* Specify Screen Layout Here */ %>
<FORM id=frmScreen style="VISIBILITY: hidden" mark validate="onchange">
	
<DIV class=msgGroup id=divStrategy style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP:5px; HEIGHT: 212px">
	<SPAN class=msgLabel style="LEFT: 4px; POSITION: absolute; TOP: 4px">Strategy Manager</SPAN>
	<SPAN class=msgLabel style="LEFT: 490px; POSITION: absolute; TOP: 4px">Number 
		<SPAN style="LEFT: 50px; POSITION: absolute; TOP: -1px">
			<INPUT class=msgReadOnly id=txtSMPosition style="WIDTH: 50px; POSITION: absolute" tabIndex=-1 readOnly maxLength=10 align=right name=txtSMPosition>
		</SPAN>
	</SPAN>
	<SPAN id=spnStrategyTable style="LEFT: 4px; POSITION: absolute; TOP: 27px">
		<TABLE class=msgTable id=tblStrategy cellSpacing=0 cellPadding=0 width=596 border=0>
			<TR id=SMrowTitles>
				<TD class=TableHead width="5%">Pass</TD>
				<TD class=TableHead width="20%">Reason Code</TD>
				<TD class=TableHead width="33%">Rule</TD>
				<TD class=TableHead width="17%">Authority Level</TD>
				<TD class=TableHead width="15%">Overridden?</TD>
					</TR>
			<TR id=SMrow01>
				<TD class=TableTopLeft width="5%">&nbsp;</TD>
				<TD class=TableTopCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableTopCenter width="33%" nowrap=true>&nbsp;</TD>
				<TD class=TableTopCenter width="17%">&nbsp;</TD>
				<TD class=TableTopRight>&nbsp;</TD></TR>
			<TR id=SMrow02>
				<TD class=TableLeft width="5%">&nbsp;</TD>
				<TD class=TableCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="33%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="17%">&nbsp;</TD>
				<TD class=TableRight>&nbsp;</TD></TR>
			<TR id=SMrow03>
				<TD class=TableLeft width="5%">&nbsp;</TD>
				<TD class=TableCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="33%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="17%">&nbsp;</TD>
				<TD class=TableRight>&nbsp;</TD></TR>
			<TR id=SMrow04>
				<TD class=TableLeft width="5%">&nbsp;</TD>
				<TD class=TableCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="33%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="17%">&nbsp;</TD>
				<TD class=TableRight>&nbsp;</TD></TR>
			<TR id=SMrow05>
				<TD class=TableLeft width="5%">&nbsp;</TD>
				<TD class=TableCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="33%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="17%">&nbsp;</TD>
				<TD class=TableRight>&nbsp;</TD></TR>
			<TR id=SMrow06>
				<TD class=TableBottomLeft width="5%">&nbsp;</TD>
				<TD class=TableBottomCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableBottomCenter width="33%" nowrap=true>&nbsp;</TD>
				<TD class=TableBottomCenter width="17%">&nbsp;</TD>
				<TD class=TableBottomRight>&nbsp;</TD>
			</TR>
		</TABLE>
	</SPAN>
	<DIV style="LEFT: 4px; POSITION: absolute; TOP: 165px">
		<SPAN class=msgLabel style="LEFT: 0px; POSITION: absolute; TOP: 0px">Strategy Manager Decision
			<SPAN style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<INPUT class=msgReadOnly id="txtSMDecision" style="WIDTH: 150px; POSITION: absolute" tabIndex=-1 readOnly name=txtSMDecision>
			</SPAN>
		</SPAN>
		
		<SPAN style="LEFT: 465px; POSITION: absolute; TOP: -5px">
			<INPUT class=msgButton id=btnSMPrevious style="WIDTH: 60px" type=button value=Previous name=btnSMPrevious> 
		</SPAN>
		<SPAN style="LEFT: 530px; POSITION: absolute; TOP: -5px">
			<INPUT class=msgButton id=btnSMNext style="WIDTH: 60px" type=button value=Next name=btnSMNext> 
		</SPAN>
	</DIV>
	<DIV style="LEFT: 4px; POSITION: absolute; TOP: 189px">
		<SPAN class=msgLabel style="LEFT: 0px; POSITION: absolute; TOP: 0px">Date and Time 
			<SPAN style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<INPUT class=msgReadOnly id=txtSMDateTime style="WIDTH: 150px; POSITION: absolute" tabIndex=-1 readOnly name=txtSMDateTime>
			</SPAN>
		</SPAN>
		<SPAN class=msgLabel style="LEFT: 310px; POSITION: absolute; TOP: 0px">User Id
			<SPAN style="LEFT: 45px; POSITION: absolute; TOP: -3px">
				<INPUT class=msgReadOnly id=txtSMUserID style="WIDTH: 80px; POSITION: absolute" tabIndex=-1 readOnly name=txtSMUserID> 
			</SPAN>
		</SPAN>
		<SPAN style="LEFT: 465px; POSITION: absolute; TOP: -5px">
			<INPUT class=msgButton id=btnSMOverRide style="WIDTH: 60px" type=button value=Override name=btnSMOverRide> 
		</SPAN>
		<SPAN style="LEFT: 530px; POSITION: absolute; TOP: -5px">
			<INPUT class=msgButton id=btnSMReview style="WIDTH: 60px" type=button value=Review name=btnSMReview> 
		</SPAN>
	</DIV>
</DIV>

<DIV class=msgGroup id=divTitle style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 227px; HEIGHT: 30px">
	<SPAN class=msgLabel style="LEFT: 4px; POSITION: absolute; TOP: 8px">Stage 
		<SPAN style="LEFT: 34px; POSITION: absolute; TOP: -3px">
			<INPUT class=msgReadOnly id=txtCAStage style="WIDTH: 160px; POSITION: absolute" tabIndex=-1 readOnly name=txtCAStage maxlength="30"> 
		</SPAN>
	</SPAN>
</DIV>

<DIV class=msgGroup id=divCaseAssessment style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 255px; HEIGHT: 305px">
	<SPAN class=msgLabel style="LEFT: 4px; POSITION: absolute; TOP: 4px">Case Assessment</SPAN>
	<SPAN class=msgLabel style="LEFT: 490px; POSITION: absolute; TOP: 7px">Number 
		<SPAN style="LEFT: 50px; POSITION: absolute; TOP: -3px">
			<INPUT class=msgReadOnly id=txtCAPosition style="WIDTH: 50px; POSITION: absolute" tabIndex=-1 readOnly maxLength=10 align=right name=txtCAPosition>
		</SPAN>
	</SPAN>
	<SPAN id=spnCA style="LEFT: 4px; POSITION: absolute; TOP: 27px">
		<TABLE class=msgTable id=tblCA cellSpacing=0 cellPadding=0 width=596 border=0>
			<TR id=CArowTitles>
				<TD class=TableHead width="5%">Pass</TD>
				<TD class=TableHead width="40%">Rule</TD>
				<TD class=TableHead width="20%">Rule Value</TD>
				<TD class=TableHead width="10%">Score</TD>
				<TD class=TableHead width="15%">Overridden?</TD></TR>
			<TR id=CArow01>
				<TD class=TableTopLeft width="5%">&nbsp;</TD>
				<TD class=TableTopCenter width="40%" nowrap=true>&nbsp;</TD>
				<TD class=TableTopCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableTopCenter width="10%">&nbsp;</TD>
				<TD class=TableTopRight>&nbsp;</TD></TR>
			<TR id=CArow02>
				<TD class=TableLeft width="5%">&nbsp;</TD>
				<TD class=TableCenter width="40%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="10%">&nbsp;</TD>
				<TD class=TableRight>&nbsp;</TD></TR>
			<TR id=CArow03>
				<TD class=TableLeft width="5%">&nbsp;</TD>
				<TD class=TableCenter width="40%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="10%">&nbsp;</TD>
				<TD class=TableRight>&nbsp;</TD></TR>
			<TR id=CArow04>
				<TD class=TableLeft width="5%">&nbsp;</TD>
				<TD class=TableCenter width="40%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="10%">&nbsp;</TD>
				<TD class=TableRight>&nbsp;</TD></TR>
			<TR id=CArow05>
				<TD class=TableLeft width="5%">&nbsp;</TD>
				<TD class=TableCenter width="40%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableCenter width="10%">&nbsp;</TD>
				<TD class=TableRight>&nbsp;</TD></TR>
			<TR id=CArow06>
				<TD class=TableBottomLeft width="5%">&nbsp;</TD>
				<TD class=TableBottomCenter width="40%" nowrap=true>&nbsp;</TD>
				<TD class=TableBottomCenter width="20%" nowrap=true>&nbsp;</TD>
				<TD class=TableBottomCenter width="10%">&nbsp;</TD>
				<TD class=TableBottomRight>&nbsp;</TD>
			</TR>
		</TABLE>
	</SPAN>
	<DIV style="LEFT: 4px; POSITION: absolute; TOP: 167px"> 
		<SPAN class=msgLabel style="LEFT: 0px; POSITION: absolute; TOP: 0px">Case Assessment Decision
			<SPAN style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<INPUT class=msgReadOnly id=txtCADecision style="WIDTH: 150px; POSITION: absolute" tabIndex=-1 readOnly name=txtCADecision>
			</SPAN>
		</SPAN>
		<SPAN style="LEFT: 400px; POSITION: absolute; TOP: -5px">
			<INPUT class=msgButton id=btnCAPrevious style="WIDTH: 60px" type=button value=Previous name=btnCAPrevious> 
		</SPAN>
		<SPAN style="LEFT: 465px; POSITION: absolute; TOP: -5px">
			<INPUT class=msgButton id=btnCANext style="WIDTH: 60px" type=button value=Next name=btnCANext> 
		</SPAN>
		<SPAN style="LEFT: 530px; POSITION: absolute; TOP: -5px">
			<INPUT class=msgButton id=btnNew style="WIDTH: 60px" type=button value=New name=btnCANew> 
		</SPAN>
	</DIV>
	<DIV style="LEFT: 4px; POSITION: absolute; TOP: 190px">
		<SPAN class=msgLabel style="LEFT: 0px; POSITION: absolute; TOP: 0px">Date and Time
			<SPAN style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<INPUT class=msgReadOnly id="txtCADateTime" style="WIDTH: 150px; POSITION: absolute" tabIndex=-1 readOnly name=txtCADateTime>
			</SPAN>
		</SPAN>
		<SPAN class=msgLabel style="LEFT: 310px; POSITION: absolute; TOP: 0px">User Id
			<SPAN style="LEFT: 45px; POSITION: absolute; TOP: -3px">
				<INPUT class=msgReadOnly id="txtCAUserID" style="WIDTH: 80px; POSITION: absolute" tabIndex=-1 readOnly name=txtCAUserID>
			</SPAN>
		</SPAN>
		<SPAN style="LEFT: 465px; POSITION: absolute; TOP: -5px">
			<INPUT class=msgButton id=btnCAOverRide style="WIDTH: 60px" type=button value=Override name=btnCAOverRide> 
		</SPAN>
		<SPAN style="LEFT: 530px; POSITION: absolute; TOP: -5px">
			<INPUT class=msgButton id=btnCAReview style="WIDTH: 60px" type=button value=Review name=btnCAReview> 
		</SPAN>
		<SPAN class=msgLabel style="LEFT: 0px; POSITION: absolute; TOP: 21px">Underwriter's Decision
			<SPAN style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<SELECT class=msgCombo id="cboUWDecision" style="WIDTH: 150px; POSITION: absolute" name=cboUWDecision maxlength="30"></SELECT>
			</SPAN>
		</SPAN>
		<SPAN class=msgLabel style="LEFT: 0px; POSITION: absolute; TOP: 44px">Date and Time
			<SPAN style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<INPUT class=msgReadOnly id="txtUWDateTime" style="WIDTH: 150px; POSITION: absolute" tabIndex=-1 readOnly name=txtUWDateTime>
			</SPAN>
		</SPAN>
		<SPAN class=msgLabel style="LEFT: 310px; POSITION: absolute; TOP: 44px">User Id
			<SPAN style="LEFT: 45px; POSITION: absolute; TOP: -3px">
				<INPUT class=msgReadOnly id="txtUWUserID" style="WIDTH: 80px; POSITION: absolute" tabIndex=-1 readOnly name=txtUWUserID>
			</SPAN>
		</SPAN>
		<SPAN class=msgLabel style="LEFT: 0px; POSITION: absolute; TOP: 67px">Counter Offer 
			<SPAN style="LEFT: 140px; POSITION: absolute; TOP: -3px">
				<INPUT class="msgTxt" id="txtUWCounterOffer" style="WIDTH: 150px; POSITION: absolute" tabIndex=-1 name=txtUWCounterOffer>
			</SPAN>
		</SPAN>
		<SPAN style="LEFT: 10px; POSITION: absolute; TOP: 88px">
			<INPUT class=msgButton id="btnOK" style="WIDTH: 60px" type=button value="OK" name="btnOK"> 
		</SPAN>
		
		<SPAN style="LEFT: 75px; POSITION: absolute; TOP: 88px">
			<INPUT class=msgButton id="btnCancel" style="WIDTH: 60px" type=button value="Cancel" name="btnCancel"> 
		</SPAN>
		
		<SPAN style="LEFT: 140px; POSITION: absolute; TOP: 88px">
			<INPUT class=msgButton id="btnTask" style="WIDTH: 60px" type=button value="Task" name="btnTask"> 
		</SPAN>
		<SPAN class=msgLabel style="LEFT: 450px; POSITION: absolute; TOP: 67px">Your Override Level
			<SPAN style="LEFT: 110px; POSITION: absolute; TOP: -3px">
				<INPUT class=msgReadOnly id=txtUWOverrideLevel style="WIDTH: 25px; POSITION: absolute" tabIndex=-1 readOnly name=txtUWOverrideLevel>
			</SPAN>
		</SPAN>

	</DIV>
</DIV>
</FORM>

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/ra012attribs.asp" -->

<% /* Specify Code Here */ %>
<SCRIPT language=JScript>
<!--
var m_sUserId = null;
var m_sUnitId=null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sActiveApplicationFactFindNumber = null;
var m_sUserRole = null ;
var m_sStageNumber = null;
var m_sActiveStageNumber = null;
var m_arrRequestAttributes = null;
var m_SMXML = null;
var m_XMLStage = null;
var m_RAXML = null;
var m_UWXML = null;
var m_iSMRowLength;
var m_iSMSeqNO;
var m_iCurrSMRowIndex=null;
var m_iRASequenceNumber=null;
var m_iRAMaxSequenceNumber=null;
var m_iRowRATot = null ;
var m_blnReadOnly = false;
var m_sDataFreeze = "0";  <% /* MAR1370 */ %>
var scScreenFunctions = null ;
var m_nUserAuthorityLevel;
var m_CCXML = null;
var m_nOrigUnderwriterValue="";

<% /* MAR1319 */ %>
var m_SaveSMXML = null;  
var m_iTotalSMRulesFailed = null;             

<% /* MAR667 */ %>
var m_sApplicationPriority = null;   
var m_sActivityID = null;            
var m_BaseNonPopupWindow = null ;    
var m_CustomerXML = null;            
var m_sCaseStageSequenceNo = null;   

<% /* MAR1291 */ %>
var m_RAStages = new Array();        
var m_sListStage;                    
var m_ArrayIndex;                    
var m_MaxArrayIndex;                 

<% /* MAR1563 */ %>
var m_iRAMaxSequenceNumberDisplay=null;
var m_iRAOffset=null;

<% /* MAR1447 */ %>
var m_bMandateOK;

function window.onload()
{
	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];
	m_BaseNonPopupWindow = sArguments[5];                                   // MAR667

	<% /* Array contains the basic config including userid */ %>
	var sArgArray = sArguments[4];
	m_sUserId=sArgArray[0];
	m_sUnitId=sArgArray[1];
	m_sApplicationNumber = sArgArray[2];
	m_sApplicationFactFindNumber = sArgArray[3];
	m_sActiveApplicationFactFindNumber = m_sApplicationFactFindNumber;
	m_sStageNumber = sArgArray[4];
	m_sActiveStageNumber = m_sStageNumber;
	m_arrRequestAttributes = sArgArray[5];
	m_blnReadOnly = sArgArray[6];
	m_sDataFreeze = sArgArray[7];  // MAR1370
	m_sUserRole = sArgArray[9];    // MAR667

	scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions.ShowCollection(frmScreen);

	<% /* MAR667  Call function to get context data */ %>
	RetrieveContextData();
	
	//MAR645
	GetUserAuthorityLevel();
	GetRiskAssessmentApplicationStages();
	
	PopulateCAStage();

	GetComboLists();
	
	PopulateScreen();
}

function PopulateScreen()
{
	EnableDisableUWFields();
	GetAndShowStrategyManagerData();
	GetAndShowRiskAssessmentData();
	GetAndShowUnderwritingData();
	//MAR645
	InitcboUWDecision();
	SetMasks();
	<%/* MAH EP2_1580 02/03/2007 */%>
	setAppReadOnly(m_blnReadOnly);
}

function GetAndShowStrategyManagerData()
{
	if(getStrategyManagerData())
	{
		if (parseInt(m_iSMRowLength) > 0 )
		{
			<% /* MAR1370 Set row to maximum if not yet set up. Otherwise display the current row */ %>
			if (m_iCurrSMRowIndex == null) 
			{
				m_iCurrSMRowIndex = m_iSMRowLength - 1;
			}
			m_SMXML.SelectTagListItem(m_iCurrSMRowIndex);
			//MAR487 added ActiveTag
			m_SMXML.ActiveTagList = "APPLICATIONCREDITCHECK";
			PopulateStrategyManagerData();
		}
		else
			<% /* MAR645 Still need to configure buttons if we don't have any SM  */ %>
			ConfigureSMButtons();
	}
	else
	{
		<% /* MAR645 Still need to configure buttons if we don't have any SM  */ %>
		ConfigureSMButtons();
	}
}

<% /* MAR667  Add function to retrieve context data */ %>
function RetrieveContextData()
{
	m_sApplicationPriority = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idApplicationPriority", null);
	m_sActivityID = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idActivityId", null);
	m_sCaseStageSequenceNo = scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow, "idCurrentStageOrigSeqNo", null);

	//Get customer data
	
	m_CustomerXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();	
	m_CustomerXML.CreateRequestTag(m_BaseNonPopupWindow, "");
	
	var tagCustomers = m_CustomerXML.CreateActiveTag("CUSTOMERS");
	
	for(var nLoop = 1; nLoop <= 5; nLoop++)
	{
		var sCustomerName			= scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerName" + nLoop, "Jane" + nLoop);
		var sCustomerNumber			= scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerNumber" + nLoop, nLoop);
		var sCustomerVersionNumber	= scScreenFunctions.GetContextParameter(m_BaseNonPopupWindow,"idCustomerVersionNumber" + nLoop, "1");
		if(sCustomerName != "" && sCustomerNumber != "")
		{
			m_CustomerXML.ActiveTag = tagCustomers;
			m_CustomerXML.CreateActiveTag("CUSTOMER");
			m_CustomerXML.SetAttribute("NUMBER", sCustomerNumber);
			m_CustomerXML.SetAttribute("NAME", sCustomerName);
			m_CustomerXML.SetAttribute("VERSION", sCustomerVersionNumber);
		}
	}

}

function GetAndShowRiskAssessmentData()
{
	<% /* MAR1319 Use the latest stage at which Risk Assessment is available */ %>
	if(GetLatestRiskAssessment(m_RAStages[m_MaxArrayIndex]))   
		PopulateRAData();
	else
		<% /* MAR645 Still need to configure buttons if we don't have any RA  */ %>
		ConfigureCAButtons();
}

function PopulateStrategyManagerData()
{
	with (frmScreen)
	{
		<% /* Clear of any existing data  */ %>
		txtSMPosition.value  = "" ;
		txtSMDateTime.value = "";
		txtSMUserID.value = "";
		txtSMPosition.value = (m_iCurrSMRowIndex+1) + ' of ' + m_iSMRowLength ;
		txtSMDateTime.value = m_SMXML.GetTagText("DATETIME");
		txtSMUserID.value = m_SMXML.GetTagText("USERID");
		
		//MAR645 We'll pass this through to RA016
		m_SMXML.CreateTagList("APPLICATIONCREDITCHECKLIST");
		m_CCXML = new scXMLFunctions.XMLObject();
		m_CCXML.CreateActiveTag("CREDITCHECKREASONHISTORY");
		m_CCXML.ActiveTag.appendChild(m_SMXML.ActiveTag.cloneNode(true));

		//MAR487 
		var Node = m_SMXML.SelectTag(m_SMXML.ActiveTag, "CREDITCHECK");
		if (Node != null) frmScreen.txtSMDecision.value = m_SMXML.GetTagText("DECISIONDESCRIPTION");
		
		m_SMXML.CreateTagList("CREDITCHECKREASONCODE");
		m_iTotalSMRulesFailed = m_SMXML.ActiveTagList.length;

		ReadRules(0);
		scScrollTable.initialiseTable(tblStrategy, 1, "", ReadRules, 6, m_iTotalSMRulesFailed);
		
		if(m_iTotalSMRulesFailed > 0) 
		{
			scScrollTable.setRowSelected(1);
		}
		else
		{
			<% /* MAR1319  No rows to display */ %>
			scScrollTable.setRowSelected(0);
		}
	}
	ConfigureSMButtons();
}

function ReadRules(iStart)
{	
	InitialiseRulesTable();   <% /* MAR1319 */ %>
	
	for (var iLoop = 0; iLoop < m_SMXML.ActiveTagList.length && iLoop<6; iLoop++)
	{
		m_SMXML.SelectTagListItem(iLoop+iStart);
		AddRuleToTable(iLoop,m_SMXML.GetTagText("SMOVERRIDEREASONRESULT"), 
					   m_SMXML.GetTagText("REASONCODE"), 
					   m_SMXML.GetTagText("RULEDESCRIPTION"), 
					   <% /* PB EP529 (MAR1657) 11/05/2006 & EP604 24/05/2006 */ %>
					   // sAuthorityLevel,
					   m_SMXML.GetTagText("AUTHORITYLEVEL").substring(2), // EP529
					   m_SMXML.GetTagText("OVERRIDEUSERID"));
	}
}

function AddRuleToTable(nIndex, sOverrideReasonResult ,sReasonCode, sRuleDesc, sAuthorityLevel, sOverriddenBy)
{
	var sIndex, nRowIndex;
	
	nRowIndex=nIndex+1;
	
	<% /* Dynamically load the correct image into the first column */ %>	
	if (sOverrideReasonResult) 
	{	
		if(sOverrideReasonResult == "1")
			tblStrategy.rows(nRowIndex).cells(0).innerHTML = "<img src=\"Images/chk_tick_small_3.gif\">";
		else tblStrategy.rows(nRowIndex).cells(0).innerHTML = "<img src=\"Images/chk_cross_small_3.gif\">";
	}
	else tblStrategy.rows(nRowIndex).cells(0).innerHTML = "<img src=\"Images/chk_cross_small_3.gif\">";
		
	scScreenFunctions.SizeTextToField(tblStrategy.rows(nRowIndex).cells(1), sReasonCode);
	scScreenFunctions.SizeTextToField(tblStrategy.rows(nRowIndex).cells(2), sRuleDesc);
	scScreenFunctions.SizeTextToField(tblStrategy.rows(nRowIndex).cells(3), sAuthorityLevel);
	scScreenFunctions.SizeTextToField(tblStrategy.rows(nRowIndex).cells(4), sOverriddenBy);
}

function InitialiseRulesTable()
{
	for (var iLoop = 1; iLoop <  7; iLoop++)
	{
		<% /* Image is loaded on demand, any existing image is removed */ %>
		scScreenFunctions.SizeTextToField(tblStrategy.rows(iLoop).cells(0)," ");
		scScreenFunctions.SizeTextToField(tblStrategy.rows(iLoop).cells(1),"");
		scScreenFunctions.SizeTextToField(tblStrategy.rows(iLoop).cells(2),"");
		scScreenFunctions.SizeTextToField(tblStrategy.rows(iLoop).cells(3),"");
		scScreenFunctions.SizeTextToField(tblStrategy.rows(iLoop).cells(4),"");
	}
}

function getStrategyManagerData()
{
	m_SMXML = new scXMLFunctions.XMLObject();
	m_SMXML.CreateRequestTagFromArray(m_arrRequestAttributes);
	m_SMXML.CreateActiveTag("APPLICATIONCREDITCHECKDETAILS");
	m_SMXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	m_SMXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			m_SMXML.RunASP(document,"FindApplicationCreditCheckDetailsList.asp");
			break;
		default: // Error
			m_SMXML.SetErrorResponse();
		}

	if(m_SMXML.IsResponseOK()) 
	{	
		m_SMXML.SelectTag(null,"APPLICATIONCREDITCHECKLIST");	
		
		<% /* MAR1319 Save the Strategy Manager data for use in checks */ %>
		m_SaveSMXML = new scXMLFunctions.XMLObject();
		m_SaveSMXML.CreateActiveTag("SAVEDSMDATA");
		m_SaveSMXML.ActiveTag.appendChild(m_SMXML.ActiveTag.cloneNode(true));
			
		m_SMXML.CreateTagList("APPLICATIONCREDITCHECK");
		m_iSMRowLength = m_SMXML.ActiveTagList.length;
	}
	else
	{
		alert("Unable to retrieve Strategy Manager data.");
		return false;
	}
	return true;
}

function ConfigureSMButtons()
{
	if (m_SMXML!=null)
	{
		with (frmScreen)
		{
			<% /* MAR1370  Correct enabling of next and previous buttons */ %>
			<% /* Disable the Previous button if we are at the first Strategy Manager or if none exist */ %>
			if ((m_iCurrSMRowIndex==null) || (m_iCurrSMRowIndex==0) || (m_iSMRowLength==1))
			{
				btnSMPrevious.disabled =true;
			}
			else
			{
				btnSMPrevious.disabled =false;
			}
			
			<% /* Disable the Next button if we are at the final Strategy Manager or if none exist */ %>
			if ((m_iCurrSMRowIndex==null) || (m_iCurrSMRowIndex == m_iSMRowLength-1) || (m_iSMRowLength==1))
			{
				frmScreen.btnSMNext.disabled =true;
			}
			else
			{
				frmScreen.btnSMNext.disabled =false;
			}		
		
			<% /* MAR1370 Correct check of data freeze indicator */ %>
			if((m_blnReadOnly == true) || (m_sDataFreeze == "1"))
			{ 	btnSMOverRide.disabled = true ; btnSMReview.disabled = true ; 	}
			else
			{		
				<% /* MAR1319 Check if any failed reason codes exist */ %>	
				if (m_iTotalSMRulesFailed > 0)
				{	<% /* Override and Review are mutually exclusive. XML 0 based, table 1 based */ %>
				
					<% /* MAR1391  Add table offset */ %>
					m_SMXML.SelectTagListItem(scScrollTable.getOffset() + scScrollTable.getRowSelected()-1);	

					<% /* MAR1370 Allow override of any Strategy Manager */ %>								
					<% /* if(m_iCurrSMRowIndex==m_iSMRowLength-1) */ %>
						if(m_SMXML.GetTagText("OVERRIDEUSERID")=="")
						{	
							<% //MAR1227 %>
							<% //Peter Edney - 22/02/2006 %>						
							var nRuleAuthorityLevel = m_SMXML.GetTagText("AUTHORITYLEVEL").substring(2);							
							if(!isNaN(nRuleAuthorityLevel))
							{
								if(m_nUserAuthorityLevel < nRuleAuthorityLevel)
								{
									btnSMOverRide.disabled = true; 
									btnSMReview.disabled = false ;									
								}
								else 
								{ 
									<% /* MAR1447 Check mandate level */ %>
									if (m_bMandateOK == true)
									{
										btnSMOverRide.disabled = false; 
										btnSMReview.disabled = true; 
									}
									else
									{
										btnSMOverRide.disabled = true; 
										btnSMReview.disabled = false; 
									}
								}
							}		
							else
							{ 
								<% /* MAR1447 Check mandate level */ %>
								if (m_bMandateOK == true)
								{
									btnSMOverRide.disabled = false; 
									btnSMReview.disabled = true; 
								}
								else
								{
									btnSMOverRide.disabled = true; 
									btnSMReview.disabled = false; 
								}
							}											
						}
						else
						{
							btnSMOverRide.disabled = true ; 
							btnSMReview.disabled = false ;	
						} 
					<% /* else { btnSMOverRide.disabled = true ; btnSMReview.disabled = true ; } */ %>
				}
				else 
				{
					btnSMOverRide.disabled = true ; 
					btnSMReview.disabled = true ;
				}
			}
		}
	}
	else
	{
		btnSMNext.disabled = true ; btnSMPrevious.disabled = true ;
		btnSMOverRide.disabled = true ; btnSMReview.disabled = true ; 
	}
}

function frmScreen.btnSMPrevious.onclick()
{
	m_SMXML.ActiveTag = null ;
	m_SMXML.SelectTag(null,"APPLICATIONCREDITCHECKLIST");		
	m_SMXML.CreateTagList("APPLICATIONCREDITCHECK");
	m_iCurrSMRowIndex-=1;
	m_SMXML.SelectTagListItem(m_iCurrSMRowIndex);
	PopulateStrategyManagerData();	
}

function frmScreen.btnSMNext.onclick()
{
	m_SMXML.ActiveTag = null ;
	m_SMXML.SelectTag(null,"APPLICATIONCREDITCHECKLIST");		
	m_SMXML.CreateTagList("APPLICATIONCREDITCHECK");
	m_iCurrSMRowIndex+=1;
	m_SMXML.SelectTagListItem(m_iCurrSMRowIndex);
	PopulateStrategyManagerData();	
}

function spnStrategyTable.onclick()
{
   ConfigureSMButtons();
}

<% /* MAR1291 Change function to get the latest Risk Assessment for the stage passed in */ %>
function GetLatestRiskAssessment(sStage)
{
	var XML = new scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Search");
	XML.CreateActiveTag("RISKASSESSMENT");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sActiveApplicationFactFindNumber);
	XML.CreateTag("STAGENUMBER",sStage);
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"GetLatestRiskAssessment.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{	<% /* It worked */ %>
		m_RAXML = XML;
		m_iRAMaxSequenceNumber=m_RAXML.GetTagInt("RISKASSESSMENTSEQUENCENUMBER");

		<% //MAR1563 %>
		m_iRAMaxSequenceNumberDisplay = m_RAXML.GetTagInt("RACOUNT");
		m_iRAOffset = m_RAXML.GetTagInt("RAOFFSET");

		return true ;
	}
	else 
	{ 	<% /* It failed */ %>
		m_RAXML = null;
		m_iRAMaxSequenceNumber=null;
		m_iRASequenceNumber=null;
		return false ;
	}
}

function PopulateRAData()
{
	with (frmScreen)
	{	<% /* Clear the existing data first */ %>
		txtCADecision.value = "";
		txtCADateTime.value = "";
		txtCAUserID.value= "";
		txtCAPosition.value="";
		InitialiseCATable();
		
		if(m_RAXML!=null)
		{
			m_iRASequenceNumber=m_RAXML.GetTagInt("RISKASSESSMENTSEQUENCENUMBER");
			if(m_iRASequenceNumber !=0)
			{
				txtCADecision.value = m_RAXML.GetTagAttribute("RISKASSESSMENTDECISION","TEXT");;
				
				txtCADateTime.value = m_RAXML.GetTagText("RISKASSESSMENTDATETIME");
				txtCAUserID.value= m_RAXML.GetTagText("USERID");

				m_RAXML.SelectTag(null,"RISKASSESSMENTRULELIST");
				m_RAXML.CreateTagList("RISKASSESSMENTRULE");
				m_iRowRATot = m_RAXML.ActiveTagList.length;
				if (m_iRASequenceNumber>0 && m_iRAMaxSequenceNumber>0)
				{
					<% /*
					Peter Edney - 05/04/06 - MAR1563
					txtCAPosition.value  = m_iRASequenceNumber + " of " + m_iRAMaxSequenceNumber;#
					*/ %>
					txtCAPosition.value  = (m_iRASequenceNumber + m_iRAOffset) + " of " + m_iRAMaxSequenceNumberDisplay;
				}
						
				scScrollTable1.initialiseTable(tblCA, 1, "", ReadRARules, 6, m_iRowRATot);

				if (parseInt(m_iRowRATot) > 0 )
				{
					ReadRARules(0);
					scScrollTable1.setRowSelected(1);
				}
				else
				{
					<% /* MAR1319 No rules to display */ %>
					scScrollTable1.setRowSelected(0)
				}
			}
		}
		else
		{	//Nuke the table
			InitialiseCATable();
			scScrollTable1.initialiseTable(tblCA, 1, "", ReadRARules, 6, 0);
		}
	}
	ConfigureCAButtons();
	return true;
}

function ReadRARules(iStart)
{
	m_RAXML.ActiveTag = null;
	m_RAXML.CreateTagList("RISKASSESSMENTRULE");
	InitialiseCATable(iStart);
	for (var iLoop = 0; iLoop < m_RAXML.ActiveTagList.length && iLoop<6; iLoop++)
	{
		m_RAXML.SelectTagListItem(iLoop+iStart);
		AddRARuleToTable(iLoop,m_RAXML.GetTagText("RARULENAME"), 
					   m_RAXML.GetTagText("RARULEVALUE"), 
					   m_RAXML.GetTagText("RARULESCORE"), 
					   m_RAXML.GetTagText("USERID"));
	}
}

function AddRARuleToTable(nIndex, sName, sValue, sResult, sOverridden)
{
	var sIndex, nRowIndex;
	
	nRowIndex=nIndex+1;
	
	<% /* Dynamically load the correct image into the first column */ %>	
	if (sResult=="0" || sOverridden!="") 
		tblCA.rows(nRowIndex).cells(0).innerHTML = "<img src=\"Images/chk_tick_small_3.gif\">";
	else
		tblCA.rows(nRowIndex).cells(0).innerHTML = "<img src=\"Images/chk_cross_small_3.gif\">";
		
	scScreenFunctions.SizeTextToField(tblCA.rows(nRowIndex).cells(1), sName);
	scScreenFunctions.SizeTextToField(tblCA.rows(nRowIndex).cells(2), sValue);
	scScreenFunctions.SizeTextToField(tblCA.rows(nRowIndex).cells(3), sResult);
	scScreenFunctions.SizeTextToField(tblCA.rows(nRowIndex).cells(4), sOverridden);
}

function GetRiskAssessment()
{
	<% /* Returns a particular risk assessment */ %>
	var XML = new scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Search");
	XML.CreateActiveTag("RISKASSESSMENT");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sActiveApplicationFactFindNumber);
	XML.CreateTag("STAGENUMBER",m_sListStage);                             // MAR1291
	XML.CreateTag("RISKASSESSMENTSEQUENCENUMBER",m_iRASequenceNumber );
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"GetRiskAssessment.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}	
	
	if(XML.IsResponseOK()) m_RAXML = XML;
	else 
	{	
		alert('Unable to retrieve Risk Assessment Data.');
		m_RAXML = null;
	}
}

function RunRiskAssessment()
{
	<% /* Create a new risk assessment */ %>
	var XML = new scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Execute");
	XML.SetAttribute("USERID",m_sUserId );
	XML.CreateActiveTag("RISKASSESSMENT");  //JLD SYS2982
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sActiveApplicationFactFindNumber);
	XML.CreateTag("STAGENUMBER",m_sActiveStageNumber);
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"RunRiskAssessment.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}	
	
	if(XML.IsResponseOK())
	{
		GetLatestRiskAssessment(m_sActiveStageNumber);   // MAR1291
	}
	else m_RAXML = null;
}

function ConfigureCAButtons()
{
	if (m_RAXML!=null)
	{
		with (frmScreen)
		{
			<% /* MAR1291  Disable the Previous button if we are at the first Risk Assessment */ %>
			if ((m_iRASequenceNumber==null) || ((m_iRASequenceNumber == 1) && (m_ArrayIndex == 0)))
			{
				btnCAPrevious.disabled =true;
			}
			else
			{
				btnCAPrevious.disabled =false;
			}
			
			<% /* MAR1291  Disable the Next button if we are at the final Risk Assessment */ %>
			if ((m_iRASequenceNumber==null) || ((m_iRASequenceNumber == m_iRAMaxSequenceNumber) && (m_ArrayIndex == m_MaxArrayIndex)))
			{
				frmScreen.btnCANext.disabled =true;
			}
			else
			{
				frmScreen.btnCANext.disabled =false;
			}		
		
			<% /* MAR1370  Correct check of data freeze indicator */ %>
			if((m_blnReadOnly == true) || (m_sDataFreeze == "1"))
			{ 	
				btnCAOverRide.disabled = true ; 
				btnCAReview.disabled = true ; 	
			}
			else
			{
				<% /* MAR1319 Check that failed rules exist */ %>
				if (m_iRowRATot > 0) 
				{				
					<% /* MAR1391 Add table offset */ %>
					m_RAXML.SelectTagListItem(scScrollTable1.getOffset() + scScrollTable1.getRowSelected()-1);
					
					<% /* MAR1447 Do not allow override of passed rules or rules which have already been overridden */ %>
						if ((m_RAXML.GetTagText("RAOVERRIDEDATETIME")=="") &&
						    (m_RAXML.GetTagText("RARULESCORE") != "0"))
						{
						
<% /* MAR1227 Peter Edney - 22/02/2006 */ %>						
							<% //if (m_RAXML.GetTagText("RARULESCORE")=="0") %>
							var nRuleAuthorityLevel = parseInt(m_RAXML.GetTagText("RARULESCORE"));
							if(!isNaN(nRuleAuthorityLevel))
							{
								if(m_nUserAuthorityLevel < nRuleAuthorityLevel)
								{
									btnCAOverRide.disabled = true; 	
									btnCAReview.disabled = false;
								}
								else 
								{
									<% /* MAR1447  Add check on mandate level */ %>
									if (m_bMandateOK == true)
									{
										btnCAOverRide.disabled = false; 
										btnCAReview.disabled = true; 
									}
									else
									{
										btnCAOverRide.disabled = true; 
										btnCAReview.disabled = false; 
									}
								}
							}
							else 
							{
								<% /* MAR1447  Add check on mandate level */ %>
								if (m_bMandateOK == true)
								{
									btnCAOverRide.disabled = false; 
									btnCAReview.disabled = true;
								}
								else
								{
									btnCAOverRide.disabled = true; 
									btnCAReview.disabled = false;
								}

							}
						} 
						else 
						{	<% /* No logic in allowing a re-override. Cannot review non-overriden rules */ %>
							btnCAOverRide.disabled =true;
							btnCAReview.disabled =(m_RAXML.GetTagText("RARULESCORE")=="0")
						}
					<% /* else { btnCAOverRide.disabled =true; btnCAReview.disabled =true; } */ %>
				}
				else  { btnCAOverRide.disabled =true; btnCAReview.disabled =true; }
			}
		}
	}
	else
	{
		frmScreen.btnCAOverRide.disabled =true;
		frmScreen.btnCAReview.disabled =true;
		frmScreen.btnCANext.disabled = true;
		frmScreen.btnCAPrevious.disabled =true;
	}
	
	<% /* MAR1319 New and Override buttons are not dependant on stage */ %>
	<% /* MAR1370 Correct check of data freeze indicator */ %>
	if ((m_blnReadOnly == false) && (m_sDataFreeze == "0"))	
		frmScreen.btnNew.disabled =false;
}

function spnCA.onclick()
{
	ConfigureCAButtons();
}

function InitialiseCATable()
{
	for (var iLoop = 1; iLoop <  7; iLoop++)
	{	<% /* Image is loaded on demand, any existing image is removed */ %>
		scScreenFunctions.SizeTextToField(tblCA.rows(iLoop).cells(0)," ");
		scScreenFunctions.SizeTextToField(tblCA.rows(iLoop).cells(1),"");
		scScreenFunctions.SizeTextToField(tblCA.rows(iLoop).cells(2),"");
		scScreenFunctions.SizeTextToField(tblCA.rows(iLoop).cells(3),"");
		scScreenFunctions.SizeTextToField(tblCA.rows(iLoop).cells(4),"");
	}
}

function frmScreen.btnNew.onclick ()
{
	frmScreen.style.cursor = "wait";
	RunRiskAssessment();
	
	<% /* MAR1319  Update stage */ %>
	GetRiskAssessmentApplicationStages();
	PopulateCAStage();
	
	PopulateRAData();
	
	frmScreen.style.cursor = "default";
	alert('Case Assessment Complete');
}

function frmScreen.btnCAPrevious.onclick()
{
	<% /* MAR1291 Move through the Risk Assessments regardless of stage */ %>
	m_iRASequenceNumber = m_iRASequenceNumber - 1;
	
	if (m_iRASequenceNumber == 0)
	{
		m_ArrayIndex = m_ArrayIndex - 1;
		m_sListStage = m_RAStages[m_ArrayIndex];
		
		GetLatestRiskAssessment(m_sListStage);
		PopulateCAStage();
	}
	else
	{
		GetRiskAssessment();
	}
	
	PopulateRAData();
}

function frmScreen.btnCANext.onclick()
{
	<% /* MAR1291 Move through the Risk Assessments regardless of stage */ %>
	m_iRASequenceNumber = m_iRASequenceNumber + 1;
	
	if (m_iRASequenceNumber > m_iRAMaxSequenceNumber)
	{
		<% /* Move to the next stage in the list */ %>
		m_ArrayIndex = m_ArrayIndex + 1;
		m_sListStage = m_RAStages[m_ArrayIndex];
		
		m_iRASequenceNumber = 1
	
		<% /* Call GetLatest first to set up the maximum sequence number */ %>
		GetLatestRiskAssessment(m_sListStage);
		GetRiskAssessment();
		PopulateCAStage();
	}
	else
	{
		GetRiskAssessment();
	}

	PopulateRAData();
}

function GetAndShowUnderwritingData()
{
	if(getApplUnderwritingData()) PopulateUnderwritingData();
}

function PopulateUnderwritingData()
{
	m_UWXML.SelectTag(null, "APPLICATIONUNDERWRITING");
	
	<% /* MAR1303 Further change to requirement - leave blank until populated by the underwriter. */ %>
	frmScreen.cboUWDecision.value = m_UWXML.GetTagText("UNDERWRITERSDECISION");
	
	frmScreen.txtUWDateTime.value = m_UWXML.GetTagText("UNDERWRITERDECISIONDATETIME");
	frmScreen.txtUWUserID.value = m_UWXML.GetTagText("USERID");
	frmScreen.txtUWCounterOffer.value = m_UWXML.GetTagText("COUNTEROFFER");
	//MAR645
	m_nOrigUnderwriterValue = frmScreen.cboUWDecision.value;

}

function getApplUnderwritingData()
{ 
	m_UWXML = new scXMLFunctions.XMLObject();
	m_UWXML.CreateRequestTagFromArray(m_arrRequestAttributes);
	m_UWXML.SetAttribute("USERID",m_sUserId );
	m_UWXML.CreateActiveTag("APPLICATIONUNDERWRITING"); 
	m_UWXML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	m_UWXML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sActiveApplicationFactFindNumber);
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			m_UWXML.RunASP(document,"GetApplicationUnderwriting.asp");
			break;
		default: // Error
			m_UWXML.SetErrorResponse();
		}	
	
	//MAR487
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = m_UWXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] || (ErrorReturn[1] == ErrorTypes[0]))
	{
		if(ErrorReturn[1] == ErrorTypes[0])
		{ //Error RECORDNOTFOUND
			return false ;
		}
		else if (ErrorReturn[0] == true)
		{	// No error
			return true ;
		}

	}

}

function EnableDisableUWFields()
{
	<% /* JD MAR1179 set fields as disabled first */ %>
	frmScreen.cboUWDecision.disabled = true;
	frmScreen.txtUWCounterOffer.disabled = true;
	frmScreen.btnNew.disabled = true;

	<% /* MAR1447 */ %>
	m_bMandateOK = false;
	
	var XML = new scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes);
	XML.CreateActiveTag("APPLICATION"); 
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sActiveApplicationFactFindNumber);
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"GetAcceptedOrActiveQuoteData.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}	

	var ErrorRaised = false;

	if(XML.IsResponseOK())
	{
		<% /* MAR1391  Check that Mortgage Sub Quote exists */ %>
		var Node = XML.SelectTag(null,'MORTGAGESUBQUOTE');
		if (Node != null) 
		{
			var dblAmountRequested = XML.GetTagFloat('AMOUNTREQUESTED');
		}
		else
		{
			<% /* EP302 - Call GetNewLoanDetails */ %>
			var XML = new scXMLFunctions.XMLObject();
			XML.CreateRequestTagFromArray(m_arrRequestAttributes);
			XML.CreateActiveTag("NEWLOAN"); 
			XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
			XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sActiveApplicationFactFindNumber);
			switch (ScreenRules())
				{
				case 1: // Warning
				case 0: // OK
					XML.RunASP(document,"GetNewLoanDetails.asp");
					break;
				default: // Error
					XML.SetErrorResponse();
				}

			if(XML.IsResponseOK())
			{
				var Node = XML.SelectTag(null,'NEWLOAN');
				if (Node != null) 
				{
					var dblAmountRequested = XML.GetTagFloat('AMOUNTREQUESTED');
				}
				else
				{
					alert('The amount requested cannot be found');				
					ErrorRaised = true;
				}
			}
			else
			{
				ErrorRaised = true;
			}	
		}
	}		
	else
	{
		ErrorRaised = true;
	}	

	if (ErrorRaised)
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		return false;		
	}
	else
	{
		<% /* Get global parameters UWDecisionAuthorityLevel and CaseAssessmentAuthorityLevel */ %>
		var XMLTemp = null;
		XMLTemp = new scXMLFunctions.XMLObject();
		var iUWDecAuthorityLevel = XMLTemp.GetGlobalParameterAmount(document, "UWDecisionAuthorityLevel");
		var iCAAuthorityLevel = XMLTemp.GetGlobalParameterAmount(document, "CaseAssessmentAuthorityLevel");
		
		<% /* JD MAR1179 Get the users competency info */ %>
		var compXML = new scXMLFunctions.XMLObject();
		compXML.CreateRequestTagFromArray(m_arrRequestAttributes);
		compXML.CreateActiveTag("OMIGAUSER");
		compXML.CreateTag("USERID", m_sUserId);
		compXML.RunASP(document, "GetUserCompetency.asp");
		if(compXML.IsResponseOK())
		{
			compXML.SelectTag(null, "COMPETENCY");
			var sUserMandateLevel = compXML.GetTagText("LOANAMOUNTMANDATE");

			<% /* MAR1447 Set flag true if loan amount mandate level is high enough */ %>
			if(parseInt(sUserMandateLevel,10) >= dblAmountRequested)
			{
				m_bMandateOK = true;
			}

			<% /* Logic for enabling Underwriter's decision combo */ %>
			if(parseInt(m_sUserRole,10) >= iUWDecAuthorityLevel)
			{
				if(parseInt(sUserMandateLevel,10) >= dblAmountRequested)
				{
					frmScreen.cboUWDecision.disabled = false;
					frmScreen.txtUWCounterOffer.disabled = false;
				}
			}
			<% /* Loginc for enabling New button */ %>
			if(parseInt(m_sUserRole,10) >= iCAAuthorityLevel)		
			{
				if(parseInt(sUserMandateLevel,10) >= dblAmountRequested)
				{
					frmScreen.btnNew.disabled = false;
				}
			}
		}		
		return true;
	}
}

function GetRiskAssessmentApplicationStages()
{
	<% /* Identify the user's power. This is for information. The BO's/DO's revalidate */ %>
	var XML = new scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Execute");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("STAGENUMBER",m_sActiveStageNumber);
	switch (ScreenRules())
	{
	case 1: // Warning
	case 0: // OK
		XML.RunASP(document,"GetRiskAssessmentApplicationStages.asp");
		break;
	default: // Error
		XML.SetErrorResponse();
	}	
	
	if (XML.IsResponseOK())
	{
		m_XMLStage = XML;
		
		<% /* MAR1291 Create an array holding the stages */ %>
		var nLoop = null;
		
		m_XMLStage.SelectTag(null,"RISKASSESSMENTAPPLICATIONSTAGELIST");
		
		var TagListLISTENTRY = m_XMLStage.CreateTagList("APPLICATIONSTAGE");
	
		m_MaxArrayIndex = TagListLISTENTRY.length - 1;
		
		for(var nLoop = 0;nLoop <= m_MaxArrayIndex; nLoop++)
		{
			m_XMLStage.ActiveTagList = TagListLISTENTRY;
			m_XMLStage.SelectTagListItem(nLoop);

			m_RAStages[m_MaxArrayIndex - nLoop] = m_XMLStage.GetTagText("STAGENUMBER");
		}
		
		m_ArrayIndex = m_MaxArrayIndex;
		m_sListStage = m_RAStages[m_MaxArrayIndex];
	}
	else
	{
		m_XMLStage=null;
	}
}

<% /* MAR1291 Remove Application Stage combo */ %>

function PopulateCAStage()
{<% /* Populate ApplicationStage text field */ %>
	var TagList = null;
	var nLoop = null;
	var TagOPTION = null;
	m_XMLStage.SelectTag(null,"RISKASSESSMENTAPPLICATIONSTAGELIST");
	var TagListLISTENTRY = m_XMLStage.CreateTagList("APPLICATIONSTAGE");
	
	for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
	{
		m_XMLStage.ActiveTagList = TagListLISTENTRY;
		m_XMLStage.SelectTagListItem(nLoop);

		if (m_XMLStage.GetTagText("STAGENUMBER") == m_sListStage)
		{
			frmScreen.txtCAStage.value = m_XMLStage.GetTagText("STAGENAME");
		}
	}
}

function frmScreen.btnOK.onclick()
{
	if(SaveApplUnderwritingData())
		window.close ();
	else
		alert("Failed to save Underwriting data.");
}

function frmScreen.btnCancel.onclick()
{
	window.close ();
}

function GetComboLists()
{
	var XML = new scXMLFunctions.XMLObject();
	var sGroups = new Array("UnderwritersDecision", "SMReasonCode", "SMDecisionCode");
	var bSuccess = false;

	if(XML.GetComboLists(document,sGroups))
	{
		bSuccess = true;
		bSuccess = bSuccess & XML.PopulateCombo(document,frmScreen.cboUWDecision,"UnderwritersDecision",true);
	}

	if(!bSuccess) scScreenFunctions.SetScreenToReadOnly(frmScreen);
}
//MAR645 New Function
function frmScreen.btnSMOverRide.onclick()
{
	<% /* MAR1391  Add table offset */ %>
	m_SMXML.SelectTagListItem(scScrollTable.getOffset() + scScrollTable.getRowSelected()-1);
	
	<% //MARMAR1227 %>
	<% //Peter Edney - 22/02/2006 %>
	<% //var nRuleAuthorityLevel = parseInt(m_SMXML.GetTagText("RARULESCORE")); %>
	var nRuleAuthorityLevel = m_SMXML.GetTagText("AUTHORITYLEVEL").substring(2);
	
	if(!isNaN(nRuleAuthorityLevel))
	{
		if(m_nUserAuthorityLevel < nRuleAuthorityLevel)
		{
			alert("You do not have the authority to override this rule.");
			return false;
		}
	}

	<% /* Do the override and then get the changes */ %>
	SMDisplayRA016(false);
	GetAndShowStrategyManagerData();	
}
//MAR645 New Function
function frmScreen.btnCAOverRide.onclick()
{
	<% /* MAR1391  Add table offset */ %>
	m_RAXML.SelectTagListItem(scScrollTable1.getOffset() + scScrollTable1.getRowSelected()-1);
	
	var nRuleAuthorityLevel = parseInt(m_RAXML.GetTagText("RARULESCORE"));
	if(!isNaN(nRuleAuthorityLevel))
	{
		if(m_nUserAuthorityLevel < nRuleAuthorityLevel)
		{
			alert("You do not have the authority to override this rule.");
			return false;
		}
	}

	<% /* Do the override and then get the changes */ %>
	CADisplayRA016(false);
	GetRiskAssessment();
	PopulateRAData();
}
//MAR645 New Function
function frmScreen.btnCAReview.onclick()
{
	CADisplayRA016(true);
}
//MAR645 New Function
function frmScreen.btnSMReview.onclick()
{
	SMDisplayRA016(true);
}
//MAR645 New Function
function SMDisplayRA016(bReadOnly)
{
	var sReturn = null;
	var ArrayArguments = new Array(11);

	<% /* MAR1391  Add table offset */ %>
	m_SMXML.SelectTagListItem(scScrollTable.getOffset() + scScrollTable.getRowSelected()-1);
	
	ArrayArguments[0] = m_sUserId ;	
	ArrayArguments[1] = m_sUnitId ;	
	ArrayArguments[2] = m_sApplicationNumber ;
	ArrayArguments[3] = m_sActiveApplicationFactFindNumber ;
	ArrayArguments[4] = m_sActiveStageNumber ;
	ArrayArguments[5] = m_iCurrSMRowIndex ;
	ArrayArguments[6] = m_SMXML.GetTagInt("SEQUENCENUMBER");
	ArrayArguments[7] = bReadOnly;
	ArrayArguments[8] = m_arrRequestAttributes;
	ArrayArguments[9] = m_SMXML.ActiveTag.xml;
	ArrayArguments[10] = "SM";
	ArrayArguments[11] = m_CCXML.ActiveTag.xml;
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "ra016.asp", ArrayArguments, 420, 495);
}
//MAR645 New Function
function CADisplayRA016(bReadOnly)
{
	var sReturn = null;
	var ArrayArguments = new Array(10);

	<% /* MAR1391  Add table offset */ %>
	m_RAXML.SelectTagListItem(scScrollTable1.getOffset() + scScrollTable1.getRowSelected()-1);
	
	ArrayArguments[0] = m_sUserId ;	
	ArrayArguments[1] = m_sUnitId ;	
	ArrayArguments[2] = m_sApplicationNumber ;
	ArrayArguments[3] = m_sActiveApplicationFactFindNumber ;
	ArrayArguments[4] = m_sListStage ;            <% /* MAR1370 Currently display stage */ %>
	ArrayArguments[5] = m_iRASequenceNumber ;
	ArrayArguments[6] = m_RAXML.GetTagInt("RARULENUMBER");
	ArrayArguments[7] = bReadOnly;
	ArrayArguments[8] = m_arrRequestAttributes;
	ArrayArguments[9] = m_RAXML.ActiveTag.xml;
	ArrayArguments[10] = "CA";

	sReturn = scScreenFunctions.DisplayPopup(window, document, "ra016.asp", ArrayArguments, 420, 495);
}
//MAR645 New Function
function GetUserAuthorityLevel()
{
	<% /* Identify the user's power.  */ %>

	var XML = new scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Search");
	XML.CreateTag("USERID",m_sUserId);
	switch (ScreenRules())
	{
	case 1: // Warning
	case 0: // OK
		XML.RunASP(document,"GetUserRiskAssessmentAuthority.asp");
		break;
	default: // Error
		XML.SetErrorResponse();
	}	

	
	if(XML.IsResponseOK())
	{
		m_nUserAuthorityLevel=XML.GetTagInt("RISKASSESSMENTMANDATE");
	}
	else
	{
		m_nUserAuthorityLevel=0;
	}
	frmScreen.txtUWOverrideLevel.value = m_nUserAuthorityLevel;
}
function InitcboUWDecision()
{
	var sUWDecision = 	frmScreen.cboUWDecision.value;
	var TempXML = new scXMLFunctions.XMLObject();
	if ((sUWDecision !="") && (TempXML.IsInComboValidationList(document,"UnderwritersDecision",sUWDecision,["CO"])))
	{
		scScreenFunctions.SetFieldState(frmScreen,"txtUWCounterOffer", "W");    <% /* MAR1319 */ %>
	}
	else
	{
		scScreenFunctions.SetFieldState(frmScreen,"txtUWCounterOffer", "R");    <% /* MAR1319 */ %>

	}
}
//MAR645 New Function
function frmScreen.cboUWDecision.onchange ()
{
	var sUWDecision = 	frmScreen.cboUWDecision.value;
	if(sUWDecision != "")  <% /* JD MAR1014 special case <select> option */ %>
	{
		var TempXML = new scXMLFunctions.XMLObject();
		if (TempXML.IsInComboValidationList(document,"UnderwritersDecision",sUWDecision,["CO"]))
		{	
			if((!CheckSMData()) || (!CheckRAData()))   <% /* MAR1303 */ %>
			{
				<% /* ReSet Combo.  */ %>
				frmScreen.cboUWDecision.value = m_nOrigUnderwriterValue;
				scScreenFunctions.SetFieldState(frmScreen,"txtUWCounterOffer", "R");   <% /* MAR1319 */ %>

				return;
			}
			else
				scScreenFunctions.SetFieldState(frmScreen,"txtUWCounterOffer", "W");   <% /* MAR1319 */ %>


		}
		else
		{
			if (TempXML.IsInComboValidationList(document,"UnderwritersDecision",sUWDecision,["A"]))
			{	
				if((!CheckSMData()) || (!CheckRAData()))   <% /* MAR1303 */ %>
				{
					<% /* ReSet Combo.  */ %>
					frmScreen.cboUWDecision.value = m_nOrigUnderwriterValue;
					return;
				}
			}
			scScreenFunctions.SetFieldState(frmScreen,"txtUWCounterOffer", "R");       <% /* MAR1319 */ %>
		}
	}
	else
		scScreenFunctions.SetFieldState(frmScreen,"txtUWCounterOffer", "R");           <% /* MAR1319 */ %>
		
		
	var dtTimeChanged =  scScreenFunctions.DateTimeToString(scScreenFunctions.GetAppServerDate(true));
	frmScreen.txtUWDateTime.value = dtTimeChanged;
	frmScreen.txtUWUserID.value = m_sUserId;
	
	<% /* MAR1303 Save the new Underwriter's Decision value */ %>
	m_nOrigUnderwriterValue = frmScreen.cboUWDecision.value;
	
}

<% /* MAR1319  Rewrite function to use the latest Risk Assessment for this application */ %>
function CheckRAData()
{
	var reasonResult
	var ScoreCardInd
	
	<% /* Get the latest Risk Assessment for this application */ %>
	if (GetLatestRiskAssessment(m_RAStages[m_MaxArrayIndex]))   
	{
		PopulateRAData();

		if(m_RAXML !=null )
		{
			m_RAXML.SelectTag(null,"RISKASSESSMENTRULELIST");
			m_RAXML.CreateTagList("RISKASSESSMENTRULE");

			for (var iLoop = 0; iLoop < m_RAXML.ActiveTagList.length; iLoop++)
			{
				m_RAXML.SelectTagListItem(iLoop);

				reasonResult = m_RAXML.GetTagText("RAOVERRIDEREASONCODE");
				ScoreCardInd = m_RAXML.GetTagText("SCORECARDIND");
				
				<% /* MAR1427 Check for failed rules that have not been overridden */ %>
				if ((ScoreCardInd != "1") && ((reasonResult == "") || (reasonResult == null)))
				{
					<% /* Return false as outstanding RA rules to be overridden.  */ %>
					alert("Unable to select this decision as there are outstanding Rules Codes that must be overridden.");
					return false;
				}
			}

		}
	}
	return true;
}
//MAR645 New Function
function CheckSMData()
{
	<% /* MAR1319 Use the copy of the Strategy Manager taken on entry. Correct checks for latest Credit Check */ %>
	
	var maxSeq;
	var curSeq;
	var reasonResult;
	var sDecisionClass;

	maxSeq = 0;
	with (frmScreen)
	{
		if(m_SaveSMXML != null)
		{
			m_SaveSMXML.SelectTag(null,"APPLICATIONCREDITCHECKLIST");
			m_SaveSMXML.CreateTagList("APPLICATIONCREDITCHECK");

			<% /* Find the latest Credit Check */ %>
			for (var iLoop = 0; iLoop < m_SaveSMXML.ActiveTagList.length; iLoop++)
			{
				m_SaveSMXML.SelectTagListItem(iLoop);
				curSeq = m_SaveSMXML.GetTagText("SEQUENCENUMBER");
				if (parseInt(curSeq) > parseInt(maxSeq))
					maxSeq = curSeq;
			}

			for (var iLoop = 0; iLoop < m_SaveSMXML.ActiveTagList.length; iLoop++)
			{
				m_SaveSMXML.SelectTagListItem(iLoop);
				curSeq = m_SaveSMXML.GetTagText("SEQUENCENUMBER");
				if (parseInt(curSeq) == parseInt(maxSeq))
				{
					sDecisionClass = m_SaveSMXML.GetTagText("DECISIONCLASS");
				
					<% /* MAR1427  Check that an override is required */ %>
					if (sDecisionClass != "RA")
					{
						m_SaveSMXML.SelectTag(m_SaveSMXML.ActiveTag, "CREDITCHECK")
						m_SaveSMXML.CreateTagList("CREDITCHECKREASONCODE");
					
						<% /* MAR1319 Loop over all the Reason Codes */ %>
						for (var iIndex = 0; iIndex < m_SaveSMXML.ActiveTagList.length; iIndex++)
						{
							m_SaveSMXML.SelectTagListItem(iIndex);
							reasonResult = m_SaveSMXML.GetTagText("SMOVERRIDEREASONRESULT");

							if (reasonResult != 1)
							{
								<% /* Return false - outstanding Reason Codes to be overridden.  */ %>
								alert("Unable to select this decision as there are outstanding Reason Codes that must be overridden.");
							
								return false;
							}
						}
					}
				}
			}
		}
	}
	
	return true;
}

//MAR645 New Function
function SaveApplUnderwritingData()
{ 
	<% /* MAR725 Only save if we have UserID info */ %>
	if (frmScreen.txtUWUserID.value.length > 0)
	{
		var selectedAction;
		var XML = new scXMLFunctions.XMLObject();
		XML.CreateRequestTagFromArray(m_arrRequestAttributes);
		XML.SetAttribute("USERID",m_sUserId );
		XML.CreateActiveTag("APPLICATIONUNDERWRITING"); 
		XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sActiveApplicationFactFindNumber);
		XML.CreateTag("UNDERWRITERSDECISION",frmScreen.cboUWDecision.value);
		XML.CreateTag("UNDERWRITERDECISIONDATETIME",frmScreen.txtUWDateTime.value);
		XML.CreateTag("USERID",frmScreen.txtUWUserID.value);
		XML.CreateTag("COUNTEROFFER",frmScreen.txtUWCounterOffer.value);
		switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document, "SaveApplicationUnderwriting.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}		

		if(XML.IsResponseOK())
			return true;
		else
			return false;
	}
	else
	{
		return true;
	}
}

<% /* MAR667  Add processing for TASK button */ %>

function frmScreen.btnTask.onclick()
{
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK 
		break;
		default: return;
	}

	var sReturn = null;
	
	var XML =  new scXMLFunctions.XMLObject();

	var bUnderwritingTasksOnly = true;
	<% /*AW 14/09/06	EP1103 */ %>
	var bProgressTaskMode = false;
	
	var ArrayArguments = new Array();

	ArrayArguments[0] = m_blnReadOnly;   
	ArrayArguments[1] = m_sActiveStageNumber;  
	ArrayArguments[2] = m_sApplicationNumber;   
	ArrayArguments[3] = m_sApplicationPriority; 
	ArrayArguments[4] = m_sUserRole; 
	ArrayArguments[5] = m_sActivityID;   
	ArrayArguments[6] = m_CustomerXML;  
	ArrayArguments[7] = m_arrRequestAttributes; 
	ArrayArguments[8] = m_sCaseStageSequenceNo;
	ArrayArguments[9] = bUnderwritingTasksOnly;
	<% /*AW 14/09/06	EP1103 */ %>
	ArrayArguments[10] = bProgressTaskMode;  
	sReturn = scScreenFunctions.DisplayPopup(window, document, "TM031.asp", ArrayArguments, 538, 418);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		var sTaskXML = sReturn[1];
		if(sTaskXML != "")
		{
			<% /* we have some tasks to add to the current stage. */ %>
			AddTasksToStage(sTaskXML);
		}
	}
}

function AddTasksToStage(sTaskXML)
{
	var XML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.LoadXML(sTaskXML);
	
	var sTaskId = new Array();
	
	XML.SelectTag(null,"REQUEST");
	XML.CreateTagList("CASETASK");
	
	for (var i=0; i<XML.ActiveTagList.length; i++){
		XML.SelectTagListItem(i);
		sTaskId[i] = XML.GetAttribute("TASKID");
	}

	XML.RunASP(document, "OmigaTMBO.asp");
	
	var ErrorTypes = new Array("NOTASKAUTHORITY");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[1] == ErrorTypes[0])
	{
		alert("User does not have the authority to create this task");
	}
}
<%/* MAH EP2_1580 02/03/2007 */%>
function setAppReadOnly(blnRead_Only)
{
	if(blnRead_Only == true)
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnNew.disabled = true;
		frmScreen.cboUWDecision.disabled = true;
		frmScreen.btnOK.disabled = true;
		frmScreen.btnTask.disabled = true;
		frmScreen.btnCancel.focus();
	}
}

-->
</SCRIPT>
</BODY>
</HTML>
