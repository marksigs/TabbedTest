<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      ra039.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Run Full Bureau screen
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		26/01/01	Screen Design
SR		30/01/01	included functionality
CL		05/03/01	SYS1920 Read only functionality added
SR		20/04/01	SYS2267
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMids History:

Prog	Date		Description
MDC		09/09/2002	BMIDS00336 Credit Check & Bureau Download
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		09/10/2002	BMIDS00556	Removed Validation_Init();
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"> 
	<title>Run Full Bureau</title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<form id="frmScreen" year4 mark validate="onchange">
<div id="divBackground" style="HEIGHT: 100px; LEFT: 10px; POSITION: absolute; TOP: 10px; WIDTH: 425px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Last Credit Check Run On 
		<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
			<input id="txtLastCCDate" maxlength="20" style="POSITION: absolute; WIDTH: 200px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 44px" class="msgLabel">
		Last Credit Check Successful ?
		<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
			<input id="optLastCCSuccessYes" name="LastCCSuccessfull" type="radio" value="1" checked>
			<label for="optLastCCSuccessYes" class="msgLabel">Yes</label>
		</span> 
		<span style="LEFT: 230px; POSITION: absolute; TOP: -3px">
			<input id="optLastCCSuccessNo" name="LastCCSuccessfull" type="radio" value="0">
			<label for="optLastCCSuccessNo" class="msgLabel">No</label> 
		</span> 
	</span>
</div>

<div id="divSecond" style="HEIGHT: 34px; LEFT: 10px; POSITION: absolute; TOP: 78px; WIDTH: 425px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Last Successful Credit Check Run On 
		<span style="LEFT: 190px; POSITION: absolute; TOP: -3px">
			<input id="txtLastSuccessfulCCDate" maxlength="20" style="POSITION: absolute; WIDTH: 200px" class="msgReadOnly" READONLY tabindex=-1>
		</span>
	</span>
</div>

<div id="divComments" style="HEIGHT: 40px; LEFT: 10px; POSITION: absolute; TOP: 110px; WIDTH: 425px" class="msgGroup">
	<% /* BMIDS00336 MDC 09/09/2002 */ %>
	<% /*
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px">
		<input id="btnRunFullBureau" value="Run Full Bureau" type="button" style="WIDTH: 100px" class="msgButton">
	</span> */ %>
	<% /* BMIDS00336 MDC 09/09/2002 - End */ %>
	
	<span style="LEFT: 110px; POSITION: absolute; TOP: 10px">
		<input id="btnViewResults" value="View Results" type="button" style="WIDTH: 100px" class="msgButton">
	</span>
</div>

</form>
<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 180px; WIDTH: 425px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/RA039Attribs.asp" -->


<%/* CODE */%>
<script language="JScript">
<!--

var scScreenFunctions;

var m_sApplicationNumber, m_sApplicationFactFindNumber, m_sCCSequenceNumber ;
var m_sLastCCRunOn, m_iLastCCSuccessful ;
var m_sLastSuccefullCCRunOn, m_bIsInitialCCOK, m_bWasFBDone ;

var FBRunStatusXML = null;

var m_aArgArray = null ;
var m_aRequestAttribs = null ;
var m_BaseNonPopupWindow = null;


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	
	var sButtonList = new Array("Submit");
	ShowMainButtons(sButtonList);
	
	var sArguments		= window.dialogArguments ;
	window.dialogTop	= sArguments[0];
	window.dialogLeft	= sArguments[1];
	window.dialogWidth	= sArguments[2];
	window.dialogHeight = sArguments[3];	
	m_aArgArray			= sArguments[4];
	m_BaseNonPopupWindow = sArguments[5];

	scScreenFunctions = new m_BaseNonPopupWindow.top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	RetrieveContextData();
	Initialise();
	SetAfterInitialise();
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveContextData()
{
<%	/* TEST	*/
	/*
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","C00077585");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
	
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber","00073261");
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber","1");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber2","00073962");
	scScreenFunctions.SetContextParameter(window,"idCustomerVersionNumber2","1");
	
	scScreenFunctions.SetContextParameter(window,"idCustomerName1","Customer Name 1");
	scScreenFunctions.SetContextParameter(window,"idCustomerName2","Customer Name 2");
	
	scScreenFunctions.SetContextParameter(window,"idUserId","USER");
	scScreenFunctions.SetContextParameter(window,"idUnitId", "5678");
	scScreenFunctions.SetContextParameter(window,"idMachineId", "MACH1");
	scScreenFunctions.SetContextParameter(window,"idDistributionChannelId", "CHAN1");
	*/
	/* END TEST */
%>	
	m_sApplicationNumber		 = m_aArgArray[0];
	m_sApplicationFactFindNumber = m_aArgArray[1];
	m_aRequestAttribs			 = m_aArgArray[10];
}

function Initialise()
{	
	scScreenFunctions.SetRadioGroupToDisabled(frmScreen, "LastCCSuccessfull");
	m_sLastSuccefullCCRunOn = "" ;
	m_bIsInitialCCOK = false ;
	m_bWasFBDone = false ;
	
	FBRunStatusXML = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	FBRunStatusXML.CreateRequestTagFromArray(m_aRequestAttribs, "SEARCH");	
	FBRunStatusXML.CreateActiveTag("APPLICATIONSTATUS");
	FBRunStatusXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	FBRunStatusXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	
	FBRunStatusXML.RunASP(document,"GetCCAndFBStatus.asp");
	
	if(!FBRunStatusXML.IsResponseOK())
	{
		alert('Error retreiving Full Bureau Download Status');
		return ;
	}
	
	<%/* if initial credit check was not run, raise error */ %>
	FBRunStatusXML.CreateTagList("LASTCREDITCHECK");
	
	if(FBRunStatusXML.ActiveTagList.length == 0)
	{
		alert('Initial Credit Check was not run for this application');
		return ;
	}
	else  <%/* if last credit check run failed, check whether there was one prior to that which was succesful */ %>
	{	
		FBRunStatusXML.SelectTagListItem(0) ;
		
		m_sLastCCRunOn = FBRunStatusXML.GetTagText("DATETIME");
		m_iLastCCSuccessful = FBRunStatusXML.GetTagText("SUCCESSINDICATOR");
		
		if(m_iLastCCSuccessful == "0")
		{	
			FBRunStatusXML.ActiveTag = null ;
			FBRunStatusXML.CreateTagList("LASTSUCCESSFULCREDITCHECK");
			if(FBRunStatusXML.ActiveTagList.length == 0)
			{
				m_bIsInitialCCOK = false
				alert('Initial Credit Check run was not successful for this application');
				return ;
			}
			else 
			{	
				FBRunStatusXML.SelectTagListItem(0) ;
				m_sLastSuccefullCCRunOn = FBRunStatusXML.GetTagText("DATETIME");
				m_sCCSequenceNumber = FBRunStatusXML.GetTagText("SEQUENCENUMBER");
				m_bIsInitialCCOK = true ;
			}
		}
		else
		{
			m_sCCSequenceNumber = FBRunStatusXML.GetTagText("SEQUENCENUMBER");
			m_bIsInitialCCOK = true ;
		}
	}
	
	<%/* Check whether Full Bureau Download was already done for this case */ %>
	FBRunStatusXML.ActiveTag = null ;
	FBRunStatusXML.CreateTagList("FULLBUREAUSTANDARDHEADER");
	if(FBRunStatusXML.ActiveTagList.length == 0) m_bWasFBDone = false
	else m_bWasFBDone = true ;
	
	<% /* Populate details  */ %>
	frmScreen.txtLastCCDate.value = m_sLastCCRunOn ;
	frmScreen.txtLastSuccessfulCCDate.value = m_sLastSuccefullCCRunOn ;	
	scScreenFunctions.SetRadioGroupValue(frmScreen, "LastCCSuccessfull", m_iLastCCSuccessful);	
}

<% /* BMIDS00336 MDC 09/09/2002
function frmScreen.btnRunFullBureau.onclick()
{
	var xmlRunFB ;
	xmlRunFB = new m_BaseNonPopupWindow.top.frames[1].document.all.scXMLFunctions.XMLObject();
	xmlRunFB.CreateRequestTagFromArray(m_aRequestAttribs, "SEARCH");
	xmlRunFB.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	xmlRunFB.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	xmlRunFB.CreateTag("SEQUENCENUMBER", m_sCCSequenceNumber);
	
	// 	xmlRunFB.RunASP(document,"RunFullBureau.asp");
	// Added by automated update TW 01 Oct 2002 SYS5115
	switch (ScreenRules())
	{
		case 1: // Warning
		case 0: // OK
			xmlRunFB.RunASP(document,"RunFullBureau.asp");
			break;
		default: // Error
			xmlRunFB.SetErrorResponse();
	}
	
	if(!xmlRunFB.IsResponseOK())
	{
		alert('Error in Full Bureau Download');
		return ;
	}
	else
	{
		frmScreen.btnRunFullBureau.disabled = true ;
		frmScreen.btnViewResults.disabled = false ;
	}	
}
*/ %>

function frmScreen.btnViewResults.onclick()
{
	scScreenFunctions.DisplayPopup(window, document, "RA030.asp", m_aArgArray, 630, 440) ;
}

function SetAfterInitialise()
{
	<% /* If last credit check was successful, Hide the field 'Last Successful CC Run on' else display it */ %>
	if(m_iLastCCSuccessful == "1") frmScreen.all("divSecond").style.visibility = "hidden" ;
	<%
	/* Enable/Disable buttons 'Run Full Bureau' and 'View Results' 
	 if initial credit check was not successful then disable both the buttons */
	%>
	if(!m_bIsInitialCCOK)
	{
		// frmScreen.btnRunFullBureau.disabled = true ; <% /* BMIDS00336 MDC 09/09/2002 */ %>
		frmScreen.btnViewResults.disabled = true ;
	}
	else
	{
		if(m_bWasFBDone)
		{
			// frmScreen.btnRunFullBureau.disabled = true ;	<% /* BMIDS00336 MDC 09/09/2002 */ %>
			frmScreen.btnViewResults.disabled = false ;
		}
		else
		{
			// frmScreen.btnRunFullBureau.disabled =  false ;	<% /* BMIDS00336 MDC 09/09/2002 */ %>
			frmScreen.btnViewResults.disabled = true ;
		}
	}
}

function btnSubmit.onclick()
{
	window.close();
}
-->
</script>
</BODY>
</HTML>



