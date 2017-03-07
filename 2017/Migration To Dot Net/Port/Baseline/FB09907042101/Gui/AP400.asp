<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.01 transitional//en">
<html>
<%
/*
Workfile:      AP400.asp
Copyright:     Copyright © 2002 Marlborough Stirling
Description:   Certificate Of Title Menu Screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JR		30/1/01		Screen Design
JR		06/2/01		Added functionality
CL		12/03/01	SYS2034 Read only functionality added
JR		19/03/01	SYS2089 Disable "confirm" and not OK button for specified criteria & remove 
					complete task button
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
JR		28/03/01	SYS2048 Re-enabled Further Questions button
JR		23/04/01	SYS2048 Completed ValidateROT function
BG		26/11/01	SYS3107 m_sProcessInd was being tested for the wrong thing in
					the Initialise function.
BG		26/11/01	SYS3107 removed the outer if - only want code to fire if it
					isn't pending in initialise method.  Added code within the if statement 
					to disable confirm button if it isn't pending.
DB		20/12/01	Added mandatory node CASESTAGESEQUENCENO to ValidateROT.
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BM Specific History:

Prog	Date		Description
DPF		25/06/02	BMIDS00057	-	Changed 'Report On Title' to 'Certificate Of Title'
TW      09/10/02	Modified to incorporate client validation - SYS5115
MV		05/11/2002	BMIDS00745 - Added Validation.js
DB		11/11/2002	BMIDS00862 - Remove any reference/route to ap440. Not required for bmids.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Epsom History:

Prog	Date		Description
GHun	04/04/2007	EP2_2262 - Changed ValidateROT to pass ACTIVITYINSTANCE
GHun	09/04/2007	EP2_2335 - prevent confirm from being clicked multiple times
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"> 
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript" type="text/javascript"></script>
<% /* Specify Forms Here */ %>
<form id="frmApplicants" method="post" action="AP410.asp" STYLE="DISPLAY: none"></form>
<form id="frmSearch" method="post" action="AP420.asp" STYLE="DISPLAY: none"></form>
<form id="frmProperty" method="post" action="AP430.asp" STYLE="DISPLAY: none"></form>
<!-- <form id="frmSolicitors" method="post" action="AP440.asp" STYLE="DISPLAY: none"></form> -->
<form id="frmQuestions" method="post" action="AP450.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>


<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<form id="frmScreen" year4 mark validate="onchange">

<div id="divBackground" style="HEIGHT: 330px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 190px; POSITION: absolute; TOP: 40px">
		<input id="btnAppAndProp" value="Applicants and Property Address" type="button" style="WIDTH:250px; HEIGHT:40px" class="msgButton">
	</span>
	<span style="LEFT: 190px; POSITION: absolute; TOP: 90px">
		<input id="btnSearchAndTitle" value="Details of Search and Title" type="button" style="WIDTH:250px; HEIGHT:40px" class="msgButton">
	</span>
	<span style="LEFT: 190px; POSITION: absolute; TOP: 140px">
		<input id="btnPropAndInsurance" value="Details of Property and Insurance" type="button" style="WIDTH:250px; HEIGHT:40px" class="msgButton">
	</span>
<!-- DB BMIDS00862
	<span style="LEFT: 190px; POSITION: absolute; TOP: 190px">
		<input id="btnBankDetails" value="Solicitors bank details" type="button" style="WIDTH:250px; HEIGHT:40px" class="msgButton">
	</span> 
DB END -->
	<span style="LEFT: 190px; POSITION: absolute; TOP: 190px">
		<input id="btnFurtherQuestions" value="Further Questions" type="button" style="WIDTH:250px; HEIGHT:40px" class="msgButton">
	</span>
</div>
</form>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 395px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/AP400Attribs.asp" -->

<%/* CODE */%>
<script language="JScript" type="text/javascript">
<!--
var m_sApplicationNumber, m_sApplicationFactFindNumber ;
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sUnderReview = "";
var m_sProcessInd = "";
var m_sUserId = "" ;
var m_sUnitId = "" ;
var m_sRole = "" ;
var m_sTaskXML = "";

var m_taskXML = null;
var XML = null;
var scScreenFunctions;
var m_blnReadOnly = false;

/** EVENTS **/


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	var sButtonList = new Array("Submit","Confirm");
	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	//DPF 25/06/02 - have changed this line of code from 'Report On Title' to 'Certificate Of Title' as per BMIDS00057
	FW030SetTitles("Certificate Of Title Summary","AP400",scScreenFunctions);
	
	RetrieveContextData();
	Initialise();
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP400");	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveContextData()
{
	/* TEST
	
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","C00078069");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber","00000094");
	scScreenFunctions.SetContextParameter(window,"idMortgageApplicationValue","10");
	scScreenFunctions.SetContextParameter(window,"idProcessingIndicator","1");
	//scScreenFunctions.SetContextParameter(window,"idTaskXML","<CASETASK CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"2\" TASKSTATUS=\"20\"/>");
	scScreenFunctions.SetContextParameter(window,"idTaskXML","<CASETASK TASKID=\"22\" CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"2\" TASKSTATUS=\"20\"/>");		
	//END TEST */
	
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sUnderReview = scScreenFunctions.GetContextParameter(window,"idAppUnderReview",null);
	m_sProcessInd = scScreenFunctions.GetContextParameter(window,"idProcessingIndicator",null);
	m_sUserId = scScreenFunctions.GetContextParameter(window,"idUserId",null); 
	m_sUnitId = scScreenFunctions.GetContextParameter(window,"idUnitId",null);
	m_sRole = scScreenFunctions.GetContextParameter(window,"idRole",null)
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = 	scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","");

	
//DEBUG
//m_sTaskXML = "<CASETASK TASKID=\"22\" CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"2\" TASKSTATUS=\"30\"/>";	
}

function Initialise()
{	
	var strTaskId = "" ;
	
	if(m_sTaskXML.length > 0)
	{
		m_taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_taskXML.LoadXML(m_sTaskXML);
		m_taskXML.SelectTag(null, "CASETASK");
		
		strTaskId = m_taskXML.GetAttribute("TASKID") ;
		
		//BG 26/11/01 SYS3107 removed the outer if - only want code to fire if it
		//isn't pending.  Added code within the if statement to disable confirm button 
		//if it isn't pending.
		
		//if (m_taskXML.GetAttribute("TASKSTATUS") != "10")
		//{
			if (m_taskXML.GetAttribute("TASKSTATUS") !="20")
			{
				m_sMetaAction = "1";
				DisableButtons() ;
				DisableMainButton("Confirm");			
			}
		//}	
	}
	//BG 26/11/01 SYS3107 m_sProcessInd was being tested for the wrong thing. 
	if ( (strTaskId == "") || (m_sReadOnly == "1") || (m_sUnderReview == "1") || (m_sProcessInd == "0") )
	//if ( (strTaskId == "") || (m_sReadOnly == "1") || (m_sUnderReview == "1") || (m_sProcessInd == "1") )
	{
		m_sMetaAction = "1" ;
		DisableMainButton("Confirm");
	}
// SYS2345 DRC Line removed to preserve Read Only status
	//set context
//	scScreenFunctions.SetContextParameter(window,"idReadOnly","0") ;
}

function ValidateROT()
{
	XML = null ;
	var reqTag ;

	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject() ;
	reqTag = XML.CreateRequestTag(window , "ValidateReportOnTitle");
		
	XML.ActiveTag = reqTag ;	
	XML.CreateActiveTag("APPLICATION");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	
	XML.ActiveTag = reqTag ;
	XML.CreateActiveTag("CASETASK");
	XML.SetAttribute("SOURCEAPPLICATION", m_taskXML.GetAttribute("SOURCEAPPLICATION"));
	XML.SetAttribute("CASEID", m_taskXML.GetAttribute("CASEID")) ;
	XML.SetAttribute("ACTIVITYID", m_taskXML.GetAttribute("ACTIVITYID")) ;
	XML.SetAttribute("ACTIVITYINSTANCE", m_taskXML.GetAttribute("ACTIVITYINSTANCE"));	<% /* GHun EP2_2262 */ %>
	XML.SetAttribute("STAGEID", m_taskXML.GetAttribute("STAGEID")) ;
	XML.SetAttribute("TASKID", m_taskXML.GetAttribute("TASKID")) ;
	XML.SetAttribute("TASKINSTANCE", m_taskXML.GetAttribute("TASKINSTANCE")) ;
	// DB SYS3520 20/12/01 - Added mandatory node CASESTAGESEQUENCENO.
	XML.SetAttribute("CASESTAGESEQUENCENO", m_taskXML.GetAttribute("CASESTAGESEQUENCENO")) ;
	//DB End
					
	//Pass XML to omTmBO
	// 	XML.RunASP(document,"OmigaTMBO.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"OmigaTMBO.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	
	//Verify XML response
	if(XML.IsResponseOK())
	{	
		CheckAppStatus() ;
		CheckROT() ;	
	}
	else 
	{
		return false;
	}
	return true;
}

function CheckAppStatus()
{
	XML.ActiveTag = null;
	XML.SelectTag(null, "APPSTATUS");
	if (XML.ActiveTag != null)
	{
		if (XML.GetAttribute("UNDERREVIEWIND") == "1")
		{
			XML.CreateTagList("REVIEWMESSAGE") ; //there maybe more than 1 reviewmessage element
			var strReviewMsgText = "" ;
			for (var iCount=0; iCount < XML.ActiveTagList.length; iCount++)
			{
				XML.SelectTagListItem(iCount) ;
				if (XML.GetAttribute("REVIEWMESSAGETEXT") != "") strReviewMsgText += XML.GetAttribute("REVIEWMESSAGETEXT") + '\n' ;
			}
			if (strReviewMsgText != "")
			{
				btnConfirm.style.cursor = "hand"; <% /* EP2_2335  GHun */ %>
				alert("The application has been placed under review. The Following rules failed:" + '\n' + strReviewMsgText) ;
			}
			else
			{
				btnConfirm.style.cursor = "hand"; <% /* EP2_2335  GHun */ %>
				alert("The application has been placed under review") ;
			}
			scScreenFunctions.SetContextParameter(window,"idAppUnderReview","1") ;
		}
	}		
}

function CheckROT()
{
	XML.ActiveTag = null;
	XML.SelectTag(null, "ROT");
	/* DB BMIDS00862 - Not required for bmids
	if (XML.GetAttribute("SOLICITORBANKACCMATCH") == "0")
	{
		alert(XML.GetAttribute("SOLICITORBANKACCMESSAGE")) ;
	}
	DB END */
	if (XML.GetAttribute("REISSUEOFFERIND") == "1")
	{
		btnConfirm.style.cursor = "hand";	<% /* EP2_2335  GHun */ %>
		alert(XML.GetAttribute("REISSUEOFFERMESSAGE")) ;
	}
}

function DisableButtons ()
{
	frmScreen.btnFurtherQuestions.disabled = true ;
}

function frmScreen.btnAppAndProp.onclick ()
{	
	frmApplicants.submit() ;
}

function frmScreen.btnSearchAndTitle.onclick()
{
	frmSearch.submit() ;
}

function frmScreen.btnPropAndInsurance.onclick()
{
	frmProperty.submit() ;
}
/* DB BMIDS00862 - Removed for bmids.
function frmScreen.btnBankDetails.onclick ()
{
	frmSolicitors.submit() ;
}
DB END */
function frmScreen.btnFurtherQuestions.onclick ()
{
	frmQuestions.submit() ;
}

function btnConfirm.onclick() 
{
	<% /* EP2_2335 GHun */ %>
	btnConfirm.style.cursor = "wait";
	//btnConfirm.disabled = true;
	btnConfirm.blur();

	window.setTimeout(confirmProcessing, 0);
	window.status = "Validating details...";
	<% /* EP2_2335 End */ %>
}

<% /* EP2_2335 GHun */ %>
function confirmProcessing()
{
	if (ValidateROT())
	{
		window.status = "";
		frmToTM030.submit();
	}
	else
		btnConfirm.disabled = false;
		
	btnConfirm.style.cursor = "hand";
	window.status = "";
}
<% /* EP2_2335 End */ %>

function btnSubmit.onclick() 
{
	//route to calling screen	
	frmToTM030.submit();
}


-->
</script>
</body>
</html>
