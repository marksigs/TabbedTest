<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      ra010.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Risk Assessment
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		04/02/00	Optimised version
AY		07/02/00	Forgot to put the new validation script in!
BG		17/05/00	SYS0752 Removed Tooltips
BG		04/09/00	SYS1022 fixed scrolling problem on table.
SA		30/07/01	SYS1019 ConfigureButtons function altered - review button disabled if rule not overridden
SA		31/07/01	SYS1021 Highlight Bar moving when click on cboStage - change to use onchange event.
JLD		14/11/01	SYS2982 Changed request XML for RunRiskAssessment.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		17/06/2002	BMIDS00032 - Modified Populate Screen.
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
SA		07/11/2002	BMIDS00864 Remove Total Failed Rules - and change position of rest of fields (-25)
SA		08/11/2002	BMIDS00893 Reset position of fields moved for above AQR
MDC		18/12/2002	BM0179	Validate override authority level
MV		29/01/2003	BM0302	Amended DisplayRA015()
MV		10/02/2003	BM0337 Amended frmScreen.btnNew.onclick()
BS		19/02/2003	BM0271 Disable New button in ReadOnly mode
GD		22/07/2003	BM0605 Handle ReadOnly correctly
MC		20/04/2004	BMIDS517	ra015 (risk assesment overview) popup dialog height changed.
								white space padded to the title text.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History:

Prog	Date		AQR			Description
MF		01/08/2005	MAR019		Change title text
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
<title>Case Assessment  <!-- #include file="includes/TitleWhiteSpace.asp" --></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<% /* Validation script - Controls Soft Coded Field Attributes */ %>
<script src="validation.js" language="JScript"></script>
<% /* Scriptlets - remove any which are not required */ %>
<% /* Scriptlets - remove any which are not required */ %>
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<span style="LEFT: 310px; POSITION: absolute; TOP: 190px">
<OBJECT data=scTableListScroll.asp id=scScrollTable 
style="HEIGHT: 24px; LEFT: 0px; TOP: 0px; WIDTH: 304px" tabIndex=-1 
type=text/x-scriptlet VIEWASTEXT></OBJECT>
</span>

<% /* Specify Screen Layout Here */ %>
<form id="frmScreen" style ="VISIBILITY: hidden" validate ="onchange" mark>
<div id="divBackground" style="HEIGHT: 320px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 604px" class="msgGroup">

	<span style="LEFT: 4px; POSITION: absolute; TOP: 4px" class="msgLabel">
		Stage
		<span style="LEFT: 34px; POSITION: absolute; TOP: -3px">
			<select id="cboStage" maxlength="30" style="POSITION: absolute; WIDTH: 240px" align="right" class="msgCombo"></select>
		</span>
	</span>
	<span style="LEFT: 420px; POSITION: absolute; TOP: 4px" class="msgLabel">
		Case Assessment Number
		<span style="LEFT: 130px; POSITION: absolute; TOP: -3px">
			<input id="txtPosition" maxlength="10" style="POSITION: absolute; WIDTH: 50px" align="right" class="msgTxt">
		</span>
	</span>

		
	<span id="spnTable" style="LEFT: 4px; POSITION: absolute; TOP: 22px" >
		<table id="tblRules" width="596" border="0" cellspacing="0" cellpadding="0" class="msgTable">
		<tr id="rowTitles">
			<td width="5%" class="TableHead">Pass</td>
			<td width="40%" class="TableHead">Rule</td>	
			<% /* BMIDS00336 MDC 29/08/2002  
			<td width="20%" class="TableHead">Reason for Failure</td> */ %>
			<td width="40%" class="TableHead">Rule Value</td>
			<% /* BMIDS00336 MDC 29/08/2002  */ %>
			<td width="10%" class="TableHead">Score</td>
			<td class="TableHead">Overridden?</td>
		</tr>
		<tr id="row01"><td width="5%" Class="TableTopLeft">&nbsp;</td><td width="40%" class="TableTopCenter">&nbsp;</td><td width="40%" class="TableTopCenter">&nbsp;</td><td width="10%" class="TableTopCenter">&nbsp;</td><td class="TableTopRight">&nbsp;</td></tr>
		<tr id="row02"><td width="5%" Class="TableLeft">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
		<tr id="row03"><td width="5%" Class="TableLeft">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
		<tr id="row04"><td width="5%" Class="TableLeft">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
		<tr id="row05"><td width="5%" Class="TableLeft">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
		<tr id="row06"><td width="5%" Class="TableLeft">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
		<tr id="row07"><td width="5%" Class="TableLeft">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
		<tr id="row08"><td width="5%" Class="TableLeft">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
		<tr id="row09"><td width="5%" Class="TableLeft">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="40%" class="TableCenter">&nbsp;</td><td width="10%" class="TableCenter">&nbsp;</td><td class="TableRight">&nbsp;</td></tr>
		<tr id="row10"><td width="5%" Class="TableBottomLeft">&nbsp;</td><td width="40%" class="TableBottomCenter">&nbsp;</td>	<td width="40%" class="TableBottomCenter">&nbsp;</td>	<td width="10%" class="TableBottomCenter">&nbsp;</td>	<td class="TableBottomRight">&nbsp;</td></tr>
		</table>
	</span>
<%/*BMIDS00864 Remove Total Failed Rules - and change tops of rest of fields (-25)
	<span style="LEFT: 4px; POSITION: absolute; TOP: 190px" class="msgLabel">
		Total Failed Rules
		<span style="LEFT: 135px; POSITION: absolute; TOP: -3px">
			<input id="txtScore" maxlength="5" style="POSITION: absolute; WIDTH: 30px" class="msgTxt">
		</span>
	</span>
*/%>
	<span style="LEFT: 4px; POSITION: absolute; TOP: 215px" class="msgLabel">
		Case Assessment Decision
		<span style="LEFT: 135px; POSITION: absolute; TOP: -3px">
			<input id="txtRiskAssessmentDecision" maxlength="40" style="POSITION: absolute; WIDTH: 240px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 240px" class="msgLabel">
		Underwriter's Decision
		<span style="LEFT: 135px; POSITION: absolute; TOP: -3px">
			<input id="txtUnderwriterDecision" maxlength="40" style="POSITION: absolute; WIDTH: 240px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 265px" class="msgLabel">
		Date and Time
		<span style="LEFT: 135px; POSITION: absolute; TOP: -3px">
			<input id="txtDateandTime" maxlength="25" style="POSITION: absolute; WIDTH: 110px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 254px; POSITION: absolute; TOP: 265px" class="msgLabel">
		UserId
		<span style="LEFT: 36px; POSITION: absolute; TOP: -3px">
			<input id="txtUserID" maxlength="20" style="POSITION: absolute; WIDTH: 90px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 460px; POSITION: absolute; TOP: 265px" class="msgLabel">
		Your Override level
		<span style="LEFT: 100px; POSITION: absolute; TOP: -3px">
			<input id="txtOverride" maxlength="4" style="POSITION: absolute; WIDTH: 36px" class="msgTxt">
		</span>
	</span>


	<span style="LEFT: 10px; POSITION: absolute; TOP: 290px">
		<input id="btnOK" value="OK" type="button" style="WIDTH: 60px" class="msgButton">
		
		<span style="LEFT: 192px; POSITION: absolute; TOP: 0px">
			<input id="btnPrevious" value="Previous" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		
		<span style="LEFT: 256px; POSITION: absolute; TOP: 0px">
			<input id="btnNext" value="Next" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="LEFT: 320px; POSITION: absolute; TOP: 0px">
			<input id="btnNew" value="New" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="LEFT: 462px; POSITION: absolute; TOP: 0px">
			<input id="btnOverride" value="Override" type="button" style="WIDTH: 60px" class="msgButton">
		</span>

		<span style="LEFT: 526px; POSITION: absolute; TOP: 0px">
			<input id="btnReview" value="Review" type="button" style="WIDTH: 60px" class="msgButton">
		</span>
		
	</span>
</div>
</form>

<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/ra010attribs.asp" -->
<% /* Specify Code Here */ %>
<script language="JScript">
<!--

var m_sUserId = null;
var m_sApplicationNumber = null;
var m_sApplicationFactFindNumber = null;
var m_sActiveApplicationFactFindNumber = null;
var m_sStageNumber = null;
var m_sActiveStageNumber = null;
var m_nRASequenceNumber=null;
var m_nRAMaxSequenceNumber=null;
var m_nUserAuthorityLevel=null;
var m_sUnitId=null;
var m_XML=null;
var m_XMLStage=null;
var m_iRowTot = null;
var m_arrRequestAttributes = null;
var scScreenFunctions;
<% /* BS BM0271 19/02/03 */ %>
var m_blnReadOnly = false;
<% //GD BM0605 START %>
var m_blnDataFreeze = false;
<% //GD BM0605 END %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();

	var sArguments = window.dialogArguments;
	window.dialogTop = sArguments[0];
	window.dialogLeft = sArguments[1];
	window.dialogWidth = sArguments[2];
	window.dialogHeight = sArguments[3];


	<% /* Array contains the basic config including userid */ %>
	var sArgArray = sArguments[4];

	m_sUserId=sArgArray[0];
	m_sUnitId=sArgArray[1];
	m_sApplicationNumber = sArgArray[2];
	m_sApplicationFactFindNumber = sArgArray[3];
	m_sActiveApplicationFactFindNumber = m_sApplicationFactFindNumber;
	m_sStageNumber = sArgArray[4];
	m_sActiveStageNumber  = m_sStageNumber;
	m_arrRequestAttributes = sArgArray[5];
	
	//BG SYS1022 instantiate object to fix scrolling problem.
	scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();

	SetMasks();
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPosition");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtRiskAssessmentDecision");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtUnderwriterDecision");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtUserID");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtOverride");
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtDateandTime");
	<%/*BMIDS00864 Remove Total Failed Rules
	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtScore");
	*/%>
	
	GetUserAuthorityLevel();
	GetRiskAssessmentApplicationStages();
	GetLatestRiskAssessment();
	scScreenFunctions.ShowCollection(frmScreen);
	
	
	PopulateCombos(); <% /* Combo's only populated on startup. It's less complicated this way */ %>

	PopulateScreen();
	Validation_Init();
	
	<% /* BS BM0271 19/02/03 
	Put screen in read-only mode if required and if so, disable buttons*/ %>
	m_blnReadOnly = sArgArray[6];
	m_blnDataFreeze = sArgArray[7];
	<% // GD BM0605 START %>	
	if(m_blnReadOnly | m_blnDataFreeze) 
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);

		<% // Disable Override and New %>
		frmScreen.btnOverride.disabled =true;
		frmScreen.btnNew.disabled =true;
		<% //Leave other buttons as decided by screen logic %>
		//frmScreen.btnReview.disabled =false;
		//frmScreen.btnNext.disabled = false;
		//frmScreen.btnPrevious.disabled =false;
		<% //Leave cboStage as writeable %>
		scScreenFunctions.SetFieldState(frmScreen, "cboStage", "W");
		<% // GD BM0605 END %>
	}
	<% /* BS BM0271 End 19/02/03 */ %>
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function spnTable.onclick()
{
   <% /* BS BM0271 19/02/03 */ %>
   //GD BM0605 if(!m_blnReadOnly) 
   ConfigureButtons();
}

function spnTable.ondblclick()
{
	if (!frmScreen.btnOverride.disabled) 
		frmScreen.btnOverride.onclick();
	else if (!frmScreen.btnReview.disabled) 
		frmScreen.btnReview.onclick();
	
}
//SYS1021 Use onchange event not onclick. It only needs to do this if another stage is chosen
//function frmScreen.cboStage.onclick ()
function frmScreen.cboStage.onchange ()
{
	GetRiskAssessmentForStage();
	PopulateScreen();
}

function frmScreen.btnNew.onclick ()
{
	frmScreen.style.cursor = "wait";
	RunRiskAssessment();
	PopulateScreen();
	frmScreen.style.cursor = "default";
	alert('Case Assessment Complete');
}

function frmScreen.btnNext.onclick ()
{
	m_nRASequenceNumber = m_nRASequenceNumber + 1;
	GetRiskAssessment();
	PopulateScreen();
}

function frmScreen.btnOK.onclick ()
{
	window.close ();
}
function frmScreen.btnOverride.onclick ()
{
	<% /* BM0179 MDC 17/12/2002 - Validate authority level */ %>
	m_XML.SelectTagListItem(scScrollTable.getRowSelected()-1);
	var nRuleAuthorityLevel = parseInt(m_XML.GetTagText("RARULESCORE"));
	if(!isNaN(nRuleAuthorityLevel))
	{
		if(m_nUserAuthorityLevel < nRuleAuthorityLevel)
		{
			alert("You do not have the authority to override this rule.");
			return false;
		}
	}
	<% /* BM0179 MDC 17/12/2002 - End */ %>

	<% /* Do the override and then get the changes */ %>
	DisplayRA015(false);
	GetRiskAssessment();
	PopulateScreen();

}

function frmScreen.btnReview.onclick()
{
	DisplayRA015(true);
}

function frmScreen.btnPrevious.onclick()
{

	m_nRASequenceNumber = m_nRASequenceNumber - 1;
	GetRiskAssessment();
	PopulateScreen();
}

function ConfigureButtons()
{
	
	if (m_XML!=null)
		{
		with (frmScreen)
		{
			btnNext.disabled = (m_nRASequenceNumber==m_nRAMaxSequenceNumber);
			btnPrevious.disabled = (m_nRASequenceNumber==1 | m_nRASequenceNumber==null);
			
			if (scScrollTable.getRowSelected() != -1 ) 
			{
				<% /* Override and Review are mutually exclusive. XML 0 based, table 1 based */ %>
				
				m_XML.SelectTagListItem(scScrollTable.getRowSelected()-1);
				if (m_XML.GetTagText("RAOVERRIDEDATETIME")=="")
				{
					 if (m_XML.GetTagText("RARULESCORE")=="0")	
						{
						btnOverride.disabled = true;
						btnReview.disabled =true;
						}
					else 
						{
						 
						<% // GD BM0605 START %>	
						if(!(m_blnReadOnly | m_blnDataFreeze))
						{
							btnOverride.disabled =(m_nRASequenceNumber!=m_nRAMaxSequenceNumber);
						} 
						<% // GD BM0605 END %> 
						 //SYS1019 Review Button shouldn't be enabled when rule hasn't been overwritten
						 //btnReview.disabled = false;
						 btnReview.disabled = true;
						 }
				} 
				else 
				{
					<% /* No logic in allowing a re-override. Cannot review non-overriden rules */ %>
					 btnOverride.disabled =true;
					 btnReview.disabled =(m_XML.GetTagText("RARULESCORE")=="0")
				}
			}
			else 
			{
				 btnOverride.disabled =true;
				 btnReview.disabled =true;
			}
		}
	}
	else
	{
		frmScreen.btnOverride.disabled =true;
		frmScreen.btnReview.disabled =true;
		frmScreen.btnNext.disabled = true;
		frmScreen.btnPrevious.disabled =true;
	}

	<% /* Where the stage and factfind is not the one that started the screen an override will
		not be possible, neither will new risk assessments . */ %>
	if (m_sActiveStageNumber!=m_sStageNumber 
		|| m_sActiveApplicationFactFindNumber!=m_sApplicationFactFindNumber)
	{
		frmScreen.btnNew.disabled =true;
		frmScreen.btnOverride.disabled =true;
	}
	else 
	{
		<% // GD BM0605 START %>	
		if(!(m_blnReadOnly | m_blnDataFreeze) )
		{
			frmScreen.btnNew.disabled =false;
		}
		<% // GD BM0605 END %>
		
	}
}

function PopulateScreen()
{
	var xmlTag;

	<% /* Reset the screen */ %>

	with (frmScreen)
	{
		txtRiskAssessmentDecision.value = "";
		txtUnderwriterDecision.value = "";
		txtDateandTime.value= "";
		txtUserID.value = "";
		<%/*BMIDS00864 Remove Total Failed Rules
		txtScore.value= "";
		*/%>
		txtPosition.value  = "";
		txtOverride.value  = m_nUserAuthorityLevel;
	}
	
	frmScreen.cboStage.value =m_sActiveStageNumber + "|" + m_sActiveApplicationFactFindNumber;
					
	if (m_XML!=null) {
	
		m_XML.SelectTag(null,"RISKASSESSMENT");
		
		m_nRASequenceNumber=m_XML.GetTagInt("RISKASSESSMENTSEQUENCENUMBER");
		
		if (m_nRASequenceNumber !=0) {
			with (frmScreen)
			{
				txtRiskAssessmentDecision.value = m_XML.GetTagAttribute("RISKASSESSMENTDECISION","TEXT");
				txtUnderwriterDecision.value =m_XML.GetTagAttribute("UNDERWRITERDECISION","TEXT")

				txtDateandTime.value=m_XML.GetTagText("RISKASSESSMENTDATETIME");
				txtUserID.value =m_XML.GetTagText("USERID");
				<%/*BMIDS00864 Remove Total Failed Rules
				txtScore.value= m_XML.GetTagText("RISKASSESSMENTSCORE");
				*/%>
					
				m_XML.SelectTag(null,"RISKASSESSMENTRULELIST");
				m_XML.CreateTagList("RISKASSESSMENTRULE");
				m_iRowTot = m_XML.ActiveTagList.length;
				
				<% /* MV - 17/06/2002 - BMIDS00032 - If there are any rule then Initialise the table else no */ %>
				if (parseInt(m_iRowTot) > 0 )
				{
					<% /* do all of the screen values that change */ %>
					if (m_nRASequenceNumber>0 && m_nRAMaxSequenceNumber>0)
						txtPosition.value  = m_nRASequenceNumber + " of " + m_nRAMaxSequenceNumber; 
					
					ReadRules(0);
				
					scScrollTable.initialiseTable(tblRules,1, "", ReadRules, 10, m_iRowTot);
					scScrollTable.setRowSelected(1)
				}
			}
		}
	}
	else
	{
		//Nuke the table
		InitialiseRulesTable();
		scScrollTable.initialiseTable(tblRules,1, "", ReadRules, 10, 0);
	}
    ConfigureButtons();	

	return true;
}

function PopulateCombos()
{
	var TagList = null;
	var nLoop = null;
	var TagOPTION = null;
	m_XMLStage.SelectTag(null,"RISKASSESSMENTAPPLICATIONSTAGELIST");
	var TagListLISTENTRY = m_XMLStage.CreateTagList("APPLICATIONSTAGE");
	
	while(frmScreen.cboStage.options.length > 0) frmScreen.cboStage.options.remove(0);


	for(var nLoop = 0;nLoop < TagListLISTENTRY.length;nLoop++)
	{
		m_XMLStage.ActiveTagList = TagListLISTENTRY;
		m_XMLStage.SelectTagListItem(nLoop);

		TagOPTION = document.createElement("OPTION");
		TagOPTION.value = m_XMLStage.GetTagText("STAGENUMBER") + "|" + m_XMLStage.GetTagText("APPLICATIONFACTFINDNUMBER");
		TagOPTION.text = m_XMLStage.GetTagText("STAGENAME");
		frmScreen.cboStage.options.add(TagOPTION);
	}

}

function InitialiseRulesTable(iStart)
{
	for (var iLoop = 1; iLoop <  11; iLoop++)
	{
		<% /* Image is loaded on demand, any existing image is removed */ %>
		scScreenFunctions.SizeTextToField(tblRules.rows(iLoop).cells(0)," ");
		scScreenFunctions.SizeTextToField(tblRules.rows(iLoop).cells(1),"");
		scScreenFunctions.SizeTextToField(tblRules.rows(iLoop).cells(2),"");
		scScreenFunctions.SizeTextToField(tblRules.rows(iLoop).cells(3),"");
		scScreenFunctions.SizeTextToField(tblRules.rows(iLoop).cells(4),"");
	}
}

function AddRuleToTable(nIndex, sName, sValue, sResult, sOverridden)
{
	var sIndex, nRowIndex;
	
	nRowIndex=nIndex+1;
	
	<% /* Dynamically load the correct image into the first column */ %>	
	if (sResult=="0" || sOverridden!="") 
		tblRules.rows(nRowIndex).cells(0).innerHTML = "<img src=\"Images/chk_tick_small_3.gif\">";
	else
		tblRules.rows(nRowIndex).cells(0).innerHTML = "<img src=\"Images/chk_cross_small_3.gif\">";
		
	scScreenFunctions.SizeTextToField(tblRules.rows(nRowIndex).cells(1), sName);
	scScreenFunctions.SizeTextToField(tblRules.rows(nRowIndex).cells(2), sValue);
	scScreenFunctions.SizeTextToField(tblRules.rows(nRowIndex).cells(3), sResult);
	scScreenFunctions.SizeTextToField(tblRules.rows(nRowIndex).cells(4), sOverridden);
}

function ReadRules(iStart)
{
	m_XML.ActiveTag = null;
	m_XML.CreateTagList("RISKASSESSMENTRULE");
	InitialiseRulesTable(iStart);
	for (var iLoop = 0; iLoop < m_XML.ActiveTagList.length && iLoop<10; iLoop++)
	{
		m_XML.SelectTagListItem(iLoop+iStart);
		AddRuleToTable(iLoop,m_XML.GetTagText("RARULENAME"), 
					   m_XML.GetTagText("RARULEVALUE"), 
					   m_XML.GetTagText("RARULESCORE"), 
					   m_XML.GetTagText("USERID"));
	}
}


function GetRiskAssessment()
{
	<% /* Returns a particular risk assessment */ %>
	var XML = new scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Search");
	XML.CreateActiveTag("RISKASSESSMENT");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sActiveApplicationFactFindNumber);
	XML.CreateTag("STAGENUMBER",m_sActiveStageNumber );
	XML.CreateTag("RISKASSESSMENTSEQUENCENUMBER",m_nRASequenceNumber );
	// 	XML.RunASP(document,"GetRiskAssessment.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"GetRiskAssessment.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}


	if(XML.IsResponseOK())
	{
		m_XML = XML;
	}
	else
	{
		m_XML = null;
	}

}

function GetRiskAssessmentForStage()
{
	<% /* Returns the lastest risk assessment within a particular stage */ %>

	var XML = new scXMLFunctions.XMLObject();
	var arrvalue=null;
	var s = new String(frmScreen.cboStage.value )
	arrvalue=s.split( "|");

	m_nRASequenceNumber=null;
	
	<% /* Only get the data if the combo really has changed */ %>
	if (m_sActiveStageNumber!=arrvalue[0] || m_sActiveApplicationFactFindNumber!=arrvalue[1])
	{
		m_sActiveStageNumber=arrvalue[0];
		m_sActiveFactApplicationFindNumber=arrvalue[1];
		GetLatestRiskAssessment();
	}
}

function GetLatestRiskAssessment()
{
	<% /* Returns the latest risk assessment for the current stage. */ %>

	var XML = new scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Search");
	XML.CreateActiveTag("RISKASSESSMENT");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sActiveApplicationFactFindNumber);
	XML.CreateTag("STAGENUMBER",m_sActiveStageNumber);
	// 	XML.RunASP(document,"GetLatestRiskAssessment.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
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
		{
		/* It worked */
		m_XML = XML;
		m_nRAMaxSequenceNumber=XML.GetTagInt("RISKASSESSMENTSEQUENCENUMBER");
		}
	else 
		{
		/* It failed */
			m_XML = null;
			m_nRAMaxSequenceNumber=null;
			m_nRASequenceNumber=null;
		}

}

function GetUserAuthorityLevel()
{
	<% /* Identify the user's power. This is for information. The BO's/DO's revalidate */ %>

	var XML = new scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Search");
	XML.CreateTag("USERID",m_sUserId);
	XML.RunASP(document,"GetUserRiskAssessmentAuthority.asp");
	if(XML.IsResponseOK())
	{
		m_nUserAuthorityLevel=XML.GetTagInt("RISKASSESSMENTMANDATE");
	}
	else
	{
		m_nUserAuthorityLevel=0;
	}
}

function GetRiskAssessmentApplicationStages()
{
	<% /* Identify the user's power. This is for information. The BO's/DO's revalidate */ %>

	var XML = new scXMLFunctions.XMLObject();
	XML.CreateRequestTagFromArray(m_arrRequestAttributes,"Execute");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("STAGENUMBER",m_sActiveStageNumber);
	XML.RunASP(document,"GetRiskAssessmentApplicationStages.asp");
	if(XML.IsResponseOK())
		m_XMLStage = XML;
	else
		m_XMLStage=null;
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
	// 	XML.RunASP(document,"RunRiskAssessment.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
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
		GetLatestRiskAssessment();
	}
	else
	{
		m_XML = null;
	}

}

function DisplayRA015(bReadOnly)
{
	var sReturn = null;
	var ArrayArguments = new Array(9);

	m_XML.SelectTagListItem(scScrollTable.getRowSelected()-1);
	
	ArrayArguments[0] = m_sUserId ;	
	ArrayArguments[1] = m_sUnitId ;	
	ArrayArguments[2] = m_sApplicationNumber ;
	ArrayArguments[3] = m_sActiveApplicationFactFindNumber ;
	ArrayArguments[4] = m_sActiveStageNumber ;
	ArrayArguments[5] = m_nRASequenceNumber ;
	ArrayArguments[6] = m_XML.GetTagInt("RARULENUMBER");
	ArrayArguments[7] = bReadOnly;
	ArrayArguments[8] = m_arrRequestAttributes;
	ArrayArguments[9] = m_XML.ActiveTag.xml;
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "ra015.asp", ArrayArguments, 420, 495);
}
-->
</script>
</body>
</html>


