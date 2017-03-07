<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/* 
Workfile:      DC300.asp
Copyright:     Copyright © 2002 Marlborough Stirling
Description:   Buildings & Contents Declaration
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Descrip
MV		23/05/2002	BMIDS00013 - Created New 
MV		14/06/2002	BMIDS00013 - Code Review Errors Modified -PopulateScreen
GD		19/06/2002	BMIDS00077 - Upgrade to Core 7.0.2
GD		24/06/2002	BMIDS00077 - Upgrade to Core 7.0.2 - changed name of screenfunctions object
GHun	06/09/2002	BMIDS00412 - Mistyped comment cleared table object definitions

BMIDS History:

Prog	Date		Descrip
TW      09/10/2002  Modified to incorporate client validation - SYS5115
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
MV		08/01/2003	BM0227  - Amended btnSave.Onclick();PopulateScreen();
HMA     18/09/2003  BM0063  - Amend HTML text for radio buttons
MC		22/06/2004	BMIDS772  txtQuestion field removed and BC Questions list table row size
							  increased from 4 to 8.
MC		24/06/2004	BMIDS772	Details Column removed, QuestionText column display instead
								of QuestionShortText.  ToopTip added						
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History

Prog	Date		AQR		Description
MF		04/08/2005	MARS20	Routes back to DC201
MF		08/08/2005	MAR20	Route depending on Global Parameters
GHun	07/11/2005	MAR437  Fixed frmCancel error when there are no questions
Maha T	23/11/2005	MAR543	Fixed routing when there are no further questions.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<HEAD>
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
<%/*
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
*/%>
<OBJECT id=scTable style="DISPLAY: none; LEFT: 0px; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 data=scTable.htm VIEWASTEXT></OBJECT>

<span id="spnQuestionListScroll">
	<span style="LEFT: 295px; POSITION: absolute; TOP: 275px">
<OBJECT id=scQuestionListTable 
style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" tabIndex=-1 
type=text/x-scriptlet data=scTableListScroll.asp 
VIEWASTEXT></OBJECT>
	</span> 
</span>

<% /* FORMS */ %>
<form id="frmToDC240" method="post" action="DC240.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC305" method="post" action="DC305.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC295" method="post" action="DC295.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC201" method="post" action="DC201.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM010" method="post" action="CM010.asp" STYLE="DISPLAY: none"></form>
<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" year4 validate="onchange" mark>

<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 242px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<font face="MS Sans Serif" size="1"><strong>Buildings &amp; Contents Questions</strong></font>
	</span>
	<span id="spnQuestionList" style="LEFT: 10px; WIDTH: 576px; POSITION: absolute; TOP: 30px; HEIGHT: 202px">
		<table id="tblQuestionList" width="575" border="0" cellspacing="0" cellpadding="0" class="msgTable" style="WIDTH: 575px; HEIGHT: 166px">
			<tr id="rowTitles">	
				<td width="85%" class="TableHead">Question&nbsp;</td>
				<td width="10%" class="TableHead">Answer&nbsp;</td>	
				<!--<td width="10%" class="TableHead">Details&nbsp;</td>-->	
				<td class="TableHead"></td>
			</tr>
			<tr id="row01">		
				<td width="85%" class="TableTopLeft">	&nbsp;</td>		
				<td width="10%" class="TableTopCenter">&nbsp;</td>		
				<!--<td width="10%" class="TableTopCenter">&nbsp;</td>-->		
				<td class="TableTopRight">&nbsp;</td>
			</tr>
			<tr id="row02">		
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td>		
				<!--<td width="10%" class="TableCenter">&nbsp;</td>-->		
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row03">		
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td>		
				<!--<td width="10%" class="TableCenter">&nbsp;</td>	-->	
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row04">		
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td>		
				<!--<td width="10%" class="TableCenter">&nbsp;</td>		-->
				<td class="TableRight">&nbsp;</td>
			</tr>			
			<tr id="row5">		
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td>		
				<!--<td width="10%" class="TableCenter">&nbsp;</td>-->
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row06">		
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td>		
				<!--<td width="10%" class="TableCenter">&nbsp;</td>	-->	
				<td class="TableRight">&nbsp;</td>
			</tr>
			<tr id="row07">		
				<td width="85%" class="TableLeft">&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td>		
				<!--<td width="10%" class="TableCenter">&nbsp;</td>	-->	
				<td class="TableRight">&nbsp;</td>
			</tr>			
			<tr id="row08">		
				<td width="85%" class="TableBottomLeft">&nbsp;</td>		
				<td width="10%" class="TableBottomCenter">&nbsp;</td>		
				<!--<td width="10%" class="TableBottomCenter">&nbsp;</td>-->
				<td class="TableBottomRight">&nbsp;</td>
			</tr>

		</table>
	</span>
</div>
<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>

<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 310px; HEIGHT: 172px" class="msgGroup"><!--<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
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
	<span style="LEFT: 10px; POSITION: absolute; TOP: 40px" class="msgLabel">
		Details
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 60px"><TEXTAREA class=msgTxt id=txtDetails style="WIDTH: 575px" rows=4 maxlength="255"></TEXTAREA>
	</span>
	<span style="LEFT: 505px; POSITION: absolute; TOP: 135px">
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
<!-- #include FILE="attribs/DC300Attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--

var scScreenFunctions;
var XML = null ;
var m_sMetaAction ;
var m_bNoFurtherQuestions = false;
var iCurrentRow = 0 ;
var iQuestionReference ;
var iCount ; 
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
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	<%//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();%>
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Buildings and Contents Questions","DC300",scScreenFunctions);
			
	RetrieveContextData();	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();	
	Initialise();
	//PopulateScreen() ;
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC300");
	if (m_blnReadOnly==true)
	{
		if (frmScreen.btnSave.style.visibility != "hidden") frmScreen.btnSave.disabled = true;
	}
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveContextData()
{
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = 	scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","");
}

<% /*  Initialises the screen */ %>
function Initialise()
{
	<% /* MF 04/08/2005 MARS20 temporarily avoid call to CheckForRemodelling
	if (!CheckForRemodelling()) */%>
	if(false)
	{
		if (m_blnGoToDC305) 
		{
			frmToDC305.submit();
			return;
		}
		else
		{
			frmToMN060.submit();
			return;
		}
	}

	PopulateScreen(1);
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function CheckForRemodelling()
{
	var bStayInScreen = true;
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	AQXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	AQXML.CreateRequestTag(window,null);
	AQXML.CreateActiveTag("SEARCH");

	var tagApplication = AQXML.CreateActiveTag("BASICQUOTATIONDETAILS");
	AQXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	AQXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	AQXML.RunASP(document,"GetAppQuoteValidatedQuotation.asp");

	if(AQXML.IsResponseOK())
	{
		var xmlQuotation = null ;
		var xmlBCSubQuote = null ;
		xmlQuotation = AQXML.XMLDocument.selectSingleNode("//QUOTATION") ;
		if(xmlQuotation == null)
		{
			alert('An accepted quote must be present');
			scScreenFunctions.SetScreenToReadOnly(frmScreen);
			DisableMainButton('Submit');
			bStayInScreen = true ;
			return (bStayInScreen);
		}
		
		AQXML.ActiveTag = xmlQuotation ;
		
		<% /* If no b&c quote exists, display a message and set screen to read only */ %>
		xmlBCSubQuote = AQXML.SelectTag(xmlQuotation, 'BUILDINGSANDCONTENTSSUBQUOTE');
		if(xmlBCSubQuote == null)
		{
			alert('No Buildings and Contents insurance has been selected on the accepted quote.');
			scScreenFunctions.SetScreenToReadOnly(frmScreen);			 
			DisableMainButton('Submit');
			bStayInScreen = false ;
			m_blnGoToDC305 = true;
			return (bStayInScreen);		
		}
		
		<% /* if no valid b&c subquote exists, display a message and go to MN060 */ %>
		if(AQXML.GetTagText('VALIDBCSUBQUOTE') != "1")
		{
			alert("Your quotation for Buildings and Contents requires remodelling before the declaration can be completed.");
			bStayInScreen = false ;
			return (bStayInScreen);	
		} 
	}
	
	return(bStayInScreen);
}

function PopulateScreen(sRowSelected)
{
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	var reqTag = XML.CreateRequestTag(window ,"GETADDITIONALQUESTIONSLIST") ; 
	XML.CreateActiveTag("ADDITIONALQUESTIONSLIST") ;
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber) ;
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber) ;
	XML.SetAttribute("TYPE",30) ;
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
			//MAR437 GHun
			//frmCancel.submit();
			<% /* START: MAR543 - Maha T 
						 Rather calling btnCancel call btnSubmit. If not user will end up -
						 in loop (dc300 -> dc295 -> dc300 ......)
			*/ %>
			//btnCancel.onclick();
			btnSubmit.onclick();
			<% /* END: MAR543 */ %>
		}	
	}
	//BMIDS772 - txtQuestion field removed.
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
		
		if(XML.GetAttribute("RESPONSE") == "1")
			strAnswer = "Yes";
		else if (XML.GetAttribute("RESPONSE") == "0")
			strAnswer = "No";
		else 
			strAnswer = " ";
		
		scScreenFunctions.SizeTextToField(tblQuestionList.rows(iRowIndex+1).cells(1), strAnswer) ;
		//strAnswer = (XML.GetAttribute("FURTHERDETAILS") != "") ? "Yes" : "No" ;
		//scScreenFunctions.SizeTextToField(tblQuestionList.rows(iRowIndex+1).cells(2), strAnswer) ;
		
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
			//BMIDS772 - txtQuestion field removed.
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
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	var updXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
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
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
		var bNewPropertySummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");			
		if(bNewPropertySummary)	
			frmToDC201.submit();
		else
			frmToDC295.submit();		
	}
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else {						
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
			var bNewPropertySummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");			
			if(bNewPropertySummary)	
				frmToDC201.submit();
			else
				frmToCM010.submit();
		}
	}
}
-->
</script>
</body>
</HTML>





