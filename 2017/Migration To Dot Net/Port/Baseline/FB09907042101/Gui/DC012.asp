<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /* 
Workfile:      DC012.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Application Verification
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
MV		24/05/2002	BMIDS00013 - Created New 
MV		14/06/2002	BMIDS00013 - Code Review Errors Modified -PopulateScreen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
DF		24/06/2002	BMIDS00077 - Changes made to bring in line with COre V7.0.2
TW      09/10/2002  Modified to incorporate client validation - SYS5115
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
DPF		22/11/2002	BMIDS01040 - If a record has just been saved ensure that the record remains selected
MV		08/01/2003	BM0227 - Amended PopulateScreen()
HMA     17/09/2003  BM0063 - Amended HTML text for radio buttons
MC		22/06/2004	BMIDS772 - Questions TextArea removed, Questions table row length increased from 4 to 8.
MC		24/06/2004	BMIDS772	Details Column removed, QuestionText column display instead
								of QuestionShortText.  ToopTip added
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog	Date		Description
MF		22/07/2005	IA_WP01 change of process flow
MHeys	20/04/2006	MAR1579 Limit additional details to 255 characters
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>

<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp 
VIEWASTEXT></OBJECT>
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>
<% /* removed - DPF 24/06/02 - BMIDS00077 
<OBJECT data="scScreenFunctions.asp" height=1 id=scScrFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data="scXMLFunctions.asp" height=1 id=scXMLFunctions style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scTable.htm height=1 id=scTable style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 type=text/x-scriptlet VIEWASTEXT></OBJECT>
*/ %>
<span style="LEFT: 295px; POSITION: absolute; TOP: 255px">
<OBJECT id=scQuestionListTable 
style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" tabIndex=-1 
type=text/x-scriptlet data=scTableListScroll.asp VIEWASTEXT></OBJECT>
</span>
<% /* FORMS */ %>

<form id="frmToDC010" method="post" action="DC010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC020" method="post" action="DC020.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate ="onchange">

<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 230px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<font face="MS Sans Serif" size="1"><strong>Application Verification Questions</strong></font>
	</span>
	<span id="spnQuestionList" style="LEFT: 10px; WIDTH: 576px; POSITION: absolute; TOP: 30px; HEIGHT: 183px">
		<%/*BMIDS772 - Table Rows increased from 4 to 8*/%>
		<table id="tblQuestionList" width="575" border="0" cellspacing="0" cellpadding="0" class="msgTable" style="WIDTH: 575px; HEIGHT: 160px">
			<tr id="rowTitles">	
				<td width="85%" class="TableHead">Question&nbsp;</td>
				<td width="10%" class="TableHead">Answer&nbsp;</td><!--<td width="10%" class="TableHead">Details&nbsp;</td>-->
				<td class="TableHead"></td>
			</tr>
			<tr id="row01">		
				<td width="85%" class="TableTopLeft">	&nbsp;</td>		
				<td width="10%" class="TableTopCenter">&nbsp;</td><!--<td width="10%" class="TableTopCenter">&nbsp;</td>-->
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">		
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td><!--<td width="10%" class="TableCenter">&nbsp;</td>	-->	
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">		
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td><!--<td width="10%" class="TableCenter">&nbsp;</td>	-->	
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">		
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td><!--<td width="10%" class="TableCenter">&nbsp;</td>-->		
				<td class="TableRight">&nbsp;</td>
			</tr>			
			<tr id="row5">		
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td><!--<td width="10%" class="TableCenter">&nbsp;</td>-->
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row06">		
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td><!--<td width="10%" class="TableCenter">&nbsp;</td>	-->	
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row07">		
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td><!--<td width="10%" class="TableCenter">&nbsp;</td>	-->	
				<td class="TableRight">&nbsp;</td>
			</tr>			
			<tr id="row8">		
				<td width="85%" class="TableBottomLeft">	&nbsp;</td>		
				<td width="10%" class="TableBottomCenter">&nbsp;</td><!--<td width="10%" class="TableBottomCenter">&nbsp;</td>-->
				<td class="TableBottomRight">&nbsp;</td>
			</tr>
		</table>
	</span>
</div>
<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 295px; HEIGHT: 165px" class="msgGroup"><!--<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Question
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 30px">
		<TEXTAREA id=txtQuestion rows=4 style="WIDTH: 575px" maxlength="255" class=msgTxt readonly></TEXTAREA>
	</span>-->
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Answer		
		<span style="LEFT: 50px; POSITION: absolute; TOP: -3px">
			<input id="optAnswer_Yes" name="Answer" type="radio" value="1"  checked><label for="optAnswer_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 105px; POSITION: absolute; TOP: -3px">
			<input id="optAnswer_No" name="Answer" type="radio" value="0"><label for="optAnswer_No" class="msgLabel">No</label>
		</span>	
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Details
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 55px"><TEXTAREA class=msgTxt id=txtDetails style="WIDTH: 575px" rows=4 maxlength="255"></TEXTAREA>
	</span>
	<span style="LEFT: 505px; POSITION: absolute; TOP: 130px">
			<input id="btnSave" value="Save" type="button" style="WIDTH: 75px" class="msgButton">
	</span>
</div>
</form>

<div id="msgButtons" style="LEFT: 8px; WIDTH: 612px; POSITION: absolute; TOP: 505px; HEIGHT: 19px"><!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/DC012Attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--

var scScreenFunctions;
var XML = null ;
var m_sMetaAction ;
var m_bNoFurtherQuestions = false;
var iCurrentRow ;
var iQuestionReference ;
var iCount ; 
<%/*BMIDS772 - Table RowLength increased from 4 to 8*/%>
var m_iTableLength = 8;
var iNumberofElements ;
var bChanged = false ;
var m_blnReadOnly = false;
var m_sReadOnly = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var DeclarationXML = null;
var m_blnGoToDC305 = false;


/** EVENTS **/


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);
	
	//next line raplced by line below - DPF 24/06/02 - BMIDS00077
	//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Application Verification","DC012",scScreenFunctions);
			
	RetrieveContextData();	
	try
	{
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();	
	//DPF 22/11/2002 - BMIDS01040 - On initial entry ensure row 1 is selected
	PopulateScreen(1) ;
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC012");
	<% /* MO - 18/11/2002 - BMIDS00376 */ %>
	if (m_blnReadOnly==true)
	{
		frmScreen.btnSave.disabled = true;
	}
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
	}
	catch(e)
	{
	alert(e);
	}
	
	
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = 	scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","");
	<% /* m_sMetaAction = "0";
	m_sReadOnly = "0";
	m_sApplicationNumber = "100007732"
	m_sApplicationFactFindNumber = "1"  */ %>
}

//DPF 22/11/2002 - BMIDS01040 - pass in a table row to be selected
function PopulateScreen(sRowSelected)
{
	//next line replaced by line below - DPF 24/06/02 - BMIDS00077
	//XML = new scXMLFunctions.XMLObject();
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	var reqTag = XML.CreateRequestTag(window ,"GETADDITIONALQUESTIONSLIST") ; 
	XML.CreateActiveTag("ADDITIONALQUESTIONSLIST") ;
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber) ;
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber) ;
	XML.SetAttribute("TYPE", 40) ;
	// 	XML.RunASP(document,"omAppProc.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"omAppProc.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{
		XML.ActiveTag = null;
		XML.SelectTag(null, "RESPONSE");				
		XML.CreateTagList("ADDITIONALQUESTIONS") ;
	
		iNumberofElements = XML.ActiveTagList.length ;
		if (iNumberofElements > 0)
		{
			<% /* Clear and initialise scrolling list box */ %> 
			scQuestionListTable.clear() ;
			scQuestionListTable.initialiseTable(tblQuestionList, 0, "", PopulateQuestionList, m_iTableLength, iNumberofElements);
			PopulateQuestionList(0) ;
			//DPF 22/11/2002 - BMIDS01040 -  Specify row to be selected
			scQuestionListTable.setRowSelected(sRowSelected);
			<% /* Populate other details with data related to first row in list box */ %>
			iQuestionReference = tblQuestionList.rows(sRowSelected).getAttribute("QUESTIONREFERENCE") ;		
			PopulateOtherDetails(iQuestionReference) ;
		}
		scScreenFunctions.SetFocusToFirstField(frmScreen);
	}
	else
	{
		if (ErrorReturn[1] == "RECORDNOTFOUND")
		{
			alert("No Further Questions exist.") ;
			return false;
		}	
	}

	//BMIDS772 - QUESTION TEXTAREA REMOVED
	//scScreenFunctions.SetFieldState(frmScreen, "txtQuestion", "R");

	if (m_sMetaAction == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen) ;
		frmScreen.btnSave.disabled = true ;
	}
}

function PopulateQuestionList(nStart)
{	
	var iRowIndex = 0;
	var strAnswer ;
	var sQuestionText = "";
	for (var iCount=0; iCount < iNumberofElements && iRowIndex < m_iTableLength; iCount++)
	{
		XML.SelectTagListItem(iCount + nStart) ;
		sQuestionText = XML.GetAttribute("QUESTIONTEXT");
		scScreenFunctions.SizeTextToField(tblQuestionList.rows(iRowIndex+1).cells(0), XML.GetAttribute("QUESTIONTEXT")) ;
		//scScreenFunctions.SizeTextToField(tblQuestionList.rows(iRowIndex+1).cells(0), sQuestionText) ;
		
		if(XML.GetAttribute("RESPONSE") == "1")
			strAnswer = "Yes";
		else if (XML.GetAttribute("RESPONSE") == "0")
			strAnswer = "No";
		else 
			strAnswer = " ";
		
		scScreenFunctions.SizeTextToField(tblQuestionList.rows(iRowIndex+1).cells(1), strAnswer) ;
		//strAnswer = (XML.GetAttribute("FURTHERDETAILS") != "") ? "Yes" : "No" ;
		//scScreenFunctions.SizeTextToField(tblQuestionList.rows(iRowIndex+1).cells(2), strAnswer) ;
		//tblQuestionList.rows(iRowIndex+1).cells(2).setAttribute("Title",strAnswer);
		tblQuestionList.rows(iRowIndex+1).setAttribute("QUESTIONREFERENCE",XML.GetAttribute("QUESTIONREFERENCE")) ;
		
		tblQuestionList.rows(iRowIndex+1).cells(0).title=sQuestionText;
		tblQuestionList.rows(iRowIndex+1).cells(1).title=sQuestionText;
		
		++iRowIndex ;
	}
}

function PopulateOtherDetails(iQuestionReference)
{
	for (iCount=0; iCount < XML.ActiveTagList.length; iCount++)
	{
		XML.SelectTagListItem(iCount) ;
		if (iQuestionReference == XML.GetAttribute("QUESTIONREFERENCE") )
		{
			//MC-txtQuestion textarea removed.
			//frmScreen.txtQuestion.value = (XML.GetAttribute("QUESTIONTEXT") != null) ? XML.GetAttribute("QUESTIONTEXT") : "" ;
			
			frmScreen.txtDetails.value = (XML.GetAttribute("FURTHERDETAILS") != null) ? XML.GetAttribute("FURTHERDETAILS") : "" ;
			scScreenFunctions.SetRadioGroupValue(frmScreen,"Answer", XML.GetAttribute("RESPONSE")) ;
			break ;
		}	
	}
}

function spnQuestionList.onclick()
{
  if (!m_bNoFurtherQuestions)
  {
	iCurrentRow = scQuestionListTable.getRowSelected() ;
	if (iCurrentRow == -1)
	{
		alert("Select one in the list") ;
		return ;
	}
	<% /* Gets customer reference attribute for selected row before populating other details */ %>
	iQuestionReference = tblQuestionList.rows(iCurrentRow).getAttribute("QUESTIONREFERENCE") ;
	PopulateOtherDetails(iQuestionReference) ;
   }	
}

function frmScreen.btnSave.onclick()
{	
	if (CommitChanges()) 
		//DPF 22/11/2002 - BMIDS01040 - Pass in row currently selected
		PopulateScreen(iCurrentRow) ;
}

function CommitChanges()
{
	var bSuccess = true;
	if(frmScreen.onsubmit())
	{
		if (IsChanged()) 
			bSuccess = SaveAppAdditionalQuestions();
	}
	else
		bSuccess = false;
	return(bSuccess);
}

function SaveAppAdditionalQuestions()
{
	var strRadioBtnChk ;
	
	<% /* Check a row is selected */ %>
	iCurrentRow = scQuestionListTable.getRowSelected() ;	
	
	<% /* check if screen is readonly */ %>
	if (m_sReadOnly == "1") return false ; 
	
	<% /* check a row is selected in list box  */ %>
	if (iCurrentRow == -1) 
	{
		alert("An item has not been selected in the list") ;
		return false ;
	}	
	
	//next line replaced by line below - DPF 24/06/02 - BMIDS00077
	//var updXML = new scXMLFunctions.XMLObject();
	updXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	reqTag = updXML.CreateRequestTag(window ,"UpdateAppAdditionalQuestions"); 
	updXML.CreateActiveTag("APPLNADDITIONALQUESTIONS");
	updXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	updXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);	
	updXML.SetAttribute("QUESTIONREFERENCE", iQuestionReference) ;
	updXML.SetAttribute("FURTHERDETAILS", frmScreen.txtDetails.value) ; 
	strRadioBtnChk = (scScreenFunctions.GetRadioGroupValue(frmScreen,"Answer") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"Answer") : "" ;
	updXML.SetAttribute("RESPONSE", strRadioBtnChk) ;			
	
	// 	updXML.RunASP(document,"omAppProc.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			updXML.RunASP(document,"omAppProc.asp");
			break;
		default: // Error
			updXML.SetErrorResponse();
		}

	
	if(!updXML.IsResponseOK()) 
		return false ;
	
	return true ;
}

function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
<% /* MF 22/07/2005 Change to process flow
		//frmToDC010.submit();
*/ %>
		frmToMN060.submit();
	}
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else
<% //MF 22/07/2005 change to process flow %>
			frmToDC010.submit();
	}
}
<% /* MAR1579 M Heys 20/04/2006 Start */ %>
function frmScreen.txtDetails.onkeyup()
{
	scScreenFunctions.RestrictLength(frmScreen, "txtDetails", 255, true);
}
<% /* MAR1579 M Heys 20/04/2006 End */ %>
-->
</script>
</body>
</HTML>




