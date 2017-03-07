<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP450.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Further Questions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Descrip
JR		1/2/01		Screen Design
JR		9/2/01		Started to include functionality
CL		12/03/01	SYS1920 Read only functionality added
JR		28/03/01	SYS2048 Complete further questions functionality
JR		10/04/01	SYS2251 Displayed meaningful error message when no record found
JLD		04/02/02	SYS3113 Show if the question hasn't been answered.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History : 

Prog	Date		AQR			Description
MV		07/08/2002	BMIDS0302	Core Ref : SYS4728 remove non-style sheet styles
KRW     25/09/03    BM0063      Sorted alignment and Made Save button read only when required
MC		22/06/2004	BMIDS772	txtQuestion textarea removed and Further Questions list size 
								increased from 4 to 8 rows
MC		24/06/2004	BMIDS772	Details Column removed, QuestionText column display instead
								of QuestionShortText.  ToopTip added								
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
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
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; VISIBILITY: hidden; TOP: 0px" tabIndex=-1 
type=text/x-scriptlet height=1 width=1 data=scClientFunctions.asp 
VIEWASTEXT></OBJECT>
<span id="spnQuestionListScroll">
	<span style="LEFT: 295px; POSITION: absolute; TOP: 263px">
<OBJECT id=scQuestionListTable 
style="LEFT: 0px; WIDTH: 304px; TOP: 0px; HEIGHT: 24px" tabIndex=-1 
type=text/x-scriptlet data=scTableListScroll.asp 
VIEWASTEXT></OBJECT>
	</span> 
</span>
<script src="validation.js" language="JScript"></script>

<% /* Specify Forms Here */ %>
<form id="frmSubmit" method="post" action="AP400.asp" STYLE="DISPLAY: none"></form>
<form id="frmCancel" method="post" action="AP400.asp" STYLE="DISPLAY: none"></form>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span><!-- Specify Screen Layout Here -->
<form id="frmScreen" mark validate="onchange" year4>

<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 60px; HEIGHT: 244px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Further Questions</strong>
	</span>
	<span id="spnQuestionList" style="LEFT: 10px; WIDTH: 576px; POSITION: absolute; TOP: 38px; HEIGHT: 175px">
		<table id="tblQuestionList" width="575" border="0" cellspacing="0" cellpadding="0" class="msgTable" style="WIDTH: 575px; HEIGHT: 157px">
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
				<!--<td width="10%" class="TableCenter">&nbsp;</td>		-->
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
			<tr id="row05">		
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
				<td width="85%" class="TableLeft">	&nbsp;</td>		
				<td width="10%" class="TableCenter">&nbsp;</td>		
				<!--<td width="10%" class="TableCenter">&nbsp;</td>	-->	
				<td class="TableRight">&nbsp;</td>
			</tr>			
			<tr id="row08">		
				<td width="85%" class="TableBottomLeft">	&nbsp;</td>		
				<td width="10%" class="TableBottomCenter">&nbsp;</td>		
				<!--<td width="10%" class="TableBottomCenter">&nbsp;</td>-->
				<td class="TableBottomRight">&nbsp;</td>
			</tr>
		</table>
	</span>
</div>
<div id="divBackground" style="LEFT: 10px; WIDTH: 604px; POSITION: absolute; TOP: 310px; HEIGHT: 170px" class="msgGroup"><!--<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Question
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 30px"><TEXTAREA class=msgTxt id=txtQuestion style="WIDTH: 575px" rows=4 readOnly maxlength="255"></TEXTAREA>
	</span>-->
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Answer		
		<span style="LEFT: 50px; WIDTH: 44px; POSITION: absolute; TOP: -2px; HEIGHT: 18px">
			<input id="optAnswer_Yes" name="Answer" type="radio" value="1"  CHECKED>
			<label for="optAnswer_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 131px; WIDTH: 42px; POSITION: absolute; TOP: -1px; HEIGHT: 22px">&nbsp;<INPUT id=optAnswer_No type=radio value =0 
name=Answer style="LEFT: 3px; TOP: -1px">
			<label for="optAnswer_No" class="msgLabel">No</label>
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
<!-- #include FILE="attribs/AP450Attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--

var scScreenFunctions;
var XML = null ;
var taskXML = null;

var m_sReadOnly ;
var m_sMetaAction ;
var m_sTaskXML ;
var m_sApplicationNumber, m_sApplicationFactFindNumber ;
var m_bNoFurtherQuestions = false;

var iCurrentRow ;
var iQuestionReference ;
var iCount ; //used for loop counts
var m_iTableLength = 8;
var iNumberofElements ;
var bChanged = false ;
var m_blnReadOnly = false;

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
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Further Questions","AP450",scScreenFunctions);
			
	RetrieveContextData();	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();	
	Initialisation() ;
//	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP450");
   m_sReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP450");


	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveContextData()
{
	/*TEST
	
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","00000280");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
		
	//END TEST */ 
		
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = 	scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","");
			
//DEBUG
//m_sReadOnly = "0" ;
//m_sMetaAction = "0" ;
//m_sTaskXML = "<CASETASK TASKID=\"22\" CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"2\" TASKSTATUS=\"40\"/>";	
}

function Initialisation()
{
	if(m_sTaskXML.length > 0) 
	{
		taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject() ;
		taskXML.LoadXML(m_sTaskXML) ;
		taskXML.ActiveTag = null;
		taskXML.SelectTag(null, "CASETASK");
		PopulateScreen() ;
	}	
}

function PopulateScreen()
{
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = XML.CreateRequestTag(window , "FindROTFurtherQuestionsList") ; 
	XML.CreateActiveTag("APPLICATION") ;
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber) ;
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber) ;
	XML.SetAttribute("TYPE", 20) ;

	//Pass XML to omRot.omRotBO
	// 	XML.RunASP(document,"ReportOnTitle.asp") ;	
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"ReportOnTitle.asp") ;	
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{
		// Process the returning XML
		XML.ActiveTag = null;
		XML.SelectTag(null, "RESPONSE");				
		XML.CreateTagList("ADDITIONALQUESTIONS") ;
	
		iNumberofElements = XML.ActiveTagList.length ;
		if (iNumberofElements > 0)
		{
			//Clear and initialise scrolling list box
			scQuestionListTable.clear() ;
			scQuestionListTable.initialiseTable(tblQuestionList, 0, "", PopulateQuestionList, m_iTableLength, iNumberofElements);
			PopulateQuestionList(0) ;
			//Populate other details with data related to first row in list box
			iQuestionReference = tblQuestionList.rows(1).getAttribute("QUESTIONREFERENCE") ;		
			PopulateOtherDetails(iQuestionReference) ;
		}
	}
	else
	{
		if (ErrorReturn[1] == "RECORDNOTFOUND")
		{
			alert("No Further Questions exist.") ;
			frmCancel.submit() ;
		}	
		
		        
	}
	//Set Question text box to read-only
	//BMIDS772 - txtQuestion removed
	//scScreenFunctions.SetFieldState(frmScreen, "txtQuestion", "R");
	
	if (m_sMetaAction == "1" || m_sReadOnly == 1) // BM0063 KRW 25/09/03
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
<%		//strAnswer = (XML.GetAttribute("RESPONSE") == "1") ? "Yes" : "No" ; JLD Show no answer if that's the case 
%>		if(XML.GetAttribute("RESPONSE") == "1")
			strAnswer = "Yes";
		else if (XML.GetAttribute("RESPONSE") == "0")
			strAnswer = "No";
		else strAnswer = " ";
		scScreenFunctions.SizeTextToField(tblQuestionList.rows(iRowIndex+1).cells(1), strAnswer) ;
		//strAnswer = (XML.GetAttribute("FURTHERDETAILS") != "") ? "Yes" : "No" ;
		//scScreenFunctions.SizeTextToField(tblQuestionList.rows(iRowIndex+1).cells(2), strAnswer) ;
		
		//Set table attributes
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
			//[MC]BMIDS772 - txtQuestion field removed
			//frmScreen.txtQuestion.value = (XML.GetAttribute("QUESTIONTEXT") != null) ? XML.GetAttribute("QUESTIONTEXT") : "" ;
			//SECTION END
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
	//Gets customer reference attribute for selected row before populating other details
	iQuestionReference = tblQuestionList.rows(iCurrentRow).getAttribute("QUESTIONREFERENCE") ;
	PopulateOtherDetails(iQuestionReference) ;
   }	
}

function UpdateAppAdditionalQuestions()
{
	var RotType = 20 ; //indicates a record exists
	var strRadioBtnChk ;
	iCurrentRow = scQuestionListTable.getRowSelected() ;	//Check a row is selected
	
	if (m_sReadOnly == "1") return false ; //check if screen is readonly
	
	if (iCurrentRow == -1) //check a row is selected in list box
	{
		alert("An item has not been selected in the list") ;
		return false ;
	}	
	var updXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	reqTag = updXML.CreateRequestTag(window , "UpdateAppAdditionalQuestions");
	updXML.CreateActiveTag("APPLNADDITIONALQUESTIONS");
	updXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	updXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);	
	updXML.SetAttribute("QUESTIONREFERENCE", iQuestionReference) ;
	//updXML.SetAttribute("TYPE", RotType); no longer required as question ref is PK	
	updXML.SetAttribute("FURTHERDETAILS", frmScreen.txtDetails.value) ; 
			
	strRadioBtnChk = (scScreenFunctions.GetRadioGroupValue(frmScreen,"Answer") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"Answer") : "" ;
	updXML.SetAttribute("RESPONSE", strRadioBtnChk) ;			
	
	//Pass XML to omAppProcBO
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

	
	//Verify XML response
	if(!updXML.IsResponseOK()) return false ;
	
	//Update ok, if we get to here 
	return true ;
}

function CommitChanges()
{
	var bSuccess = true;
	//if (m_sReadOnly != "1")
	if(frmScreen.onsubmit())
	{
		if (IsChanged()) 
			bSuccess = UpdateAppAdditionalQuestions();
	}
	else
		bSuccess = false;
	return(bSuccess);
}

function frmScreen.btnSave.onclick()
{	
	if (CommitChanges()) PopulateScreen() ;
}

function btnSubmit.onclick()
{
	if (CommitChanges()) frmSubmit.submit() ;
}

function btnCancel.onclick()
{
	frmCancel.submit() ;
}

-->
</script>
</body>
</HTML>





