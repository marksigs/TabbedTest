<%@ LANGUAGE="JSCRIPT" %>
<html>
<%/*
Workfile:      DC330.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Attitude to Borrowing
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AY		05/10/1999	Created
AD		04/02/2000	Rework
AY		14/02/00	Change to msgButtons button types
JLD		01/03/00	SYS0331 Added RestrictLength() and an onkeypress() 
					event handler for each notes field. 
					SYS0330 Removed an unused scTable.asp.
JLD		07/03/00	SYS0331 - made Restrct length stop at 255, not 256 !!
AY		03/04/00	New top menu/scScreenFunctions change
SR		15/05/00	SYS0404 Using ApplicationMode instead of metaAction.
BG		17/05/00	SYS0752 Removed Tooltips
MC		07/06/00	SYS0837 Standardise approach to restricting length
					of text in textarea fields.
AS		07/06/00	SYS0152 In Cost Modelling mode "Cancel" -> MN060
					In QuickQuote mode "Cancel" -> QQ010
MC		22/06/00	SYS0837 Use OnKeyUp event to restrict length of textarea
CL		05/03/01	SYS1920 Read only functionality added
JLD		4/12/01		SYS2806 completeness check routing
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
TW      09/10/2002  SYS5115		Modified to incorporate client validation - 
MV		25/10/2002	BMIDS00714	Change Attitude to Borowing to be Type of Advice. Change question 1 to "Type of Advice Given" 
MV		31/10/2002	BMIDS00719	Amended SaveAttitudeToBorrowingDetails()
MV		11/11/2002	BMIDS00906	AMended the HTML text
MO		20/11/2002	BMIDS00376	Data freezing errors
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<%/* SCRIPTLETS */%>
<script src="validation.js" language="JScript"></script>

<%/* FORMS */%>
<form id="frmToMN060" method="post" action="mn060.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM010" method="post" action="cm010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToQQ010" method="post" action="qq010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToGN300" method="post" action="GN300.asp" STYLE="DISPLAY: none"></form>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark validate="onchange">
	<div id="divBackground" style="HEIGHT: 150px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
		
		<span style="LEFT: 4px; POSITION: absolute; TOP: 22px" class="msgLabel">
			<div style="HEIGHT: 30px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 268px">
				Type Of Advice Given
			</div>

			<span style="LEFT: 150px; POSITION: absolute; TOP: 0px">
				<select id="cboResponse1" name="Response1" style="WIDTH: 160px" class="msgCombo">
				</select>
			</span>
			<span style="LEFT: 4px; POSITION: absolute; TOP: 40px" class="msgLabel">
				Notes
			</SPAN>
			 <span style="LEFT: 150px; POSITION: absolute; TOP: 40px" class="msgLabel">
				<TEXTAREA class=msgTxt id=txtNotes1 name=Notes1 rows=4 style="POSITION: absolute; WIDTH: 220px"></TEXTAREA> 
			</span>  
		</span>
		
		<% /* <span style="LEFT: 276px; POSITION: absolute; TOP: 3px" class="msgLabel">
			<strong>Response</strong> 
			<span style="LEFT: 104px; POSITION: absolute; TOP: 0px" class="msgLabel">
				<strong>Notes</strong> 
			</span> 
		</span> 

		 <span style="LEFT: 4px; POSITION: absolute; TOP: 22px" class="msgLabel">
			<div style="HEIGHT: 30px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 268px">
				 How important is it that the loan interest rate remains fixed for the first few years of the mortgage? 
				Type Of Advice Given
			</div>

			<span style="LEFT: 150px; POSITION: absolute; TOP: 0px">
				<select id="cboResponse1" name="Response1" style="WIDTH: 160px" class="msgCombo">
				</select>
			</span>

			 <span style="LEFT: 376px; POSITION: absolute; TOP: 0px" class="msgLabel"><TEXTAREA class=msgTxt id=txtNotes1 name=Notes1 rows=4 style="POSITION: absolute; WIDTH: 220px"></TEXTAREA> 
			</span>  
		</span>

		  <span style="LEFT: 4px; POSITION: absolute; TOP: 90px" class="msgLabel">
			<div style="HEIGHT: 30px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 268px">
				How important is it that any initial fixed rate period lasts for more than just the first few years?
			</div>

			<span style="LEFT: 272px; POSITION: absolute; TOP: 0px">
				Information on single Product
				 <select id="cboResponse2" name="Response2" style="WIDTH: 100px" class="msgCombo"> 
				</select>
			</span>

			<% /* <span style="LEFT: 376px; POSITION: absolute; TOP: 0px" class="msgLabel"><TEXTAREA class=msgTxt id=txtNotes2 name=Notes2 rows=4 style="POSITION: absolute; WIDTH: 220px"></TEXTAREA> 
			</span>  
		</span> */ %>

		<% /* <span style="LEFT: 4px; POSITION: absolute; TOP: 158px" class="msgLabel">
			<div style="HEIGHT: 30px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 268px">
				How important is it that the loan interest rate is variable but discounted during the first few years?
			</div>
				
			<span style="LEFT: 272px; POSITION: absolute; TOP: 0px">
				<select id="cboResponse3" name="Response3" style="WIDTH: 100px" class="msgCombo">
				</select>
 			</span>

			<span style="LEFT: 376px; POSITION: absolute; TOP: 0px" class="msgLabel"><TEXTAREA class=msgTxt id=txtNotes3 name=Notes3 rows=4 style="POSITION: absolute; WIDTH: 220px"></TEXTAREA> 
			</span> 
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 226px" class="msgLabel">
			<div style="HEIGHT: 30px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 268px">
				How important is it that any discount lasts for the whole duration of the loan?
			</div>
				
			<span style="LEFT: 272px; POSITION: absolute; TOP: 0px">
				<select id="cboResponse4" name="Response4" style="WIDTH: 100px" class="msgCombo">
				</select>
			</span>

			<span style="LEFT: 376px; POSITION: absolute; TOP: 0px" class="msgLabel"><TEXTAREA class=msgTxt id=txtNotes4 name=Notes4 rows=4 style="POSITION: absolute; WIDTH: 220px"></TEXTAREA> 
			</span> 
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 292px" class="msgLabel">
			<div style="HEIGHT: 30px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 268px">
				How important is it that you receive a cash lump sum following completion of the loan?
			</div>

			<span style="LEFT: 272px; POSITION: absolute; TOP: 0px">
				<select id="cboResponse5" name="Response5" style="WIDTH: 100px" class="msgCombo">
				</select>
			</span>

			<span style="LEFT: 376px; POSITION: absolute; TOP: 0px" class="msgLabel"><TEXTAREA class=msgTxt id=txtNotes5 name=Notes5 rows=4 style="POSITION: absolute; WIDTH: 220px"></TEXTAREA> 
			</span> 
		</span>

		<span style="LEFT: 4px; POSITION: absolute; TOP: 360px" class="msgLabel">
			<div style="HEIGHT: 30px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 268px">
				How important is it that you can partially or totally repay the loan without penalty during the first few years?
			</div>

			<span style="LEFT: 272px; POSITION: absolute; TOP: 0px">
				<select id="cboResponse6" name="Response6" style="WIDTH: 100px" class="msgCombo">
				</select>
			</span>

			<span style="LEFT: 376px; POSITION: absolute; TOP: 0px" class="msgLabel"><TEXTAREA class=msgTxt id=txtNotes6 name=Notes6 rows=4 style="POSITION: absolute; WIDTH: 220px"></TEXTAREA> 
			</span> 
		</span>  */ %>
	</div> 
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 220px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="attribs/dc330attribs.asp" -->

<%/* CODE */%>
<script language="JScript">
<!--
<% //  var m_sMetaAction = null;  SR 15/05/00  %>
var m_sApplicationMode = null ;
var m_sReadOnly = null;
var m_sApplicationNumber = null
var m_sApplicationFactFindNumber = null
var scScreenFunctions;
var m_blnReadOnly = false;


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
	FW030SetTitles("Attitude To Borrowing","DC330",scScreenFunctions);

	RetrieveContextData();

	//This function is contained in the field attributes file (remove if not required)
	SetMasks();
	Validation_Init();
	GetComboLists();
	PopulateScreen();
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC330");
	if (m_blnReadOnly == true) m_sReadOnly = "1";
	
	<% /* MO 20/11/2002 BMIDS00376 SetScreenToReadOnly();  */ %>
	
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

function RetrieveContextData()
{
<% /* 	SR/15/05/00 - SYS0404 - Use Aplication Mode instead of meta action
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction", null);	
*/ %>
	m_sApplicationMode = scScreenFunctions.GetContextParameter(window,"idApplicationMode","Quick Quote");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B3069");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window, "idReadOnly", "0");
}

function GetComboLists()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("AttitudeToBorrowing_1") ; 
	<% /* var sGroups = new Array("AttitudeToBorrowing_1",
	                        "AttitudeToBorrowing_2",
	                        "AttitudeToBorrowing_3",
	                        "AttitudeToBorrowing_4",
	                        "AttitudeToBorrowing_5",
	                        "AttitudeToBorrowing_6"); */ %>

	if(XML.GetComboLists(document,sGroups))
	{
		XML.PopulateCombo(document,frmScreen.cboResponse1,"AttitudeToBorrowing_1",true);
		<% /* XML.PopulateCombo(document,frmScreen.cboResponse2,"AttitudeToBorrowing_2",true);
		XML.PopulateCombo(document,frmScreen.cboResponse3,"AttitudeToBorrowing_3",true);
		XML.PopulateCombo(document,frmScreen.cboResponse4,"AttitudeToBorrowing_4",true);
		XML.PopulateCombo(document,frmScreen.cboResponse5,"AttitudeToBorrowing_5",true);
		XML.PopulateCombo(document,frmScreen.cboResponse6,"AttitudeToBorrowing_6",true); */ %>
	}
}

<% /* MO 20/11/2002 BMIDS00376 
function SetScreenToReadOnly()
{
	if (m_sReadOnly == "1")
	{
		DisableMainButton("Submit");
	}
}
 */ %>

function PopulateScreen()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("SEARCH");

	XML.CreateActiveTag("ATTITUDETOBORROWING");
	XML.CreateTag("APPLICATIONNUMBER",m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber);

	XML.RunASP(document,"GetAttitudeToBorrowing.asp");

	// Allow Record Not Found error JLD aqr ref 12
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);

	if ((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[0]))
	{
		XML.CreateTagList("ATTITUDETOBORROWING");
		for(var nLoop = 0;XML.SelectTagListItem(nLoop) != false;nLoop++)
		{
			var sQuestionNumber	= XML.GetTagText("QUESTIONNUMBER");
			var sResponse = XML.GetTagText("RESPONSETOQUESTION");
			var sNotes = XML.GetTagText("NOTES");

			switch(sQuestionNumber)
			{
				case "1":
					frmScreen.cboResponse1.value = sResponse;
					frmScreen.txtNotes1.value = sNotes;
				break;

				<% /* case "2":
					frmScreen.cboResponse2.value = sResponse;
					frmScreen.txtNotes2.value = sNotes;
				break;

				case "3":
					frmScreen.cboResponse3.value = sResponse;
					frmScreen.txtNotes3.value = sNotes;
				break;

				case "4":
					frmScreen.cboResponse4.value = sResponse;
					frmScreen.txtNotes4.value = sNotes;
				break;

				case "5":
					frmScreen.cboResponse5.value = sResponse;
					frmScreen.txtNotes5.value = sNotes;
				break;

				case "6":
					frmScreen.cboResponse6.value = sResponse;
					frmScreen.txtNotes6.value = sNotes;
				break; */ %>

				default:
				break;
			}
		}
	}
	else
	{
		alert(ErrorReturn[2]);
	}
}

function frmScreen.txtNotes1.onkeyup()
{
	scScreenFunctions.RestrictLength(frmScreen, "txtNotes1", 255, true);
}
<% /* function frmScreen.txtNotes2.onkeyup()
{
	scScreenFunctions.RestrictLength(frmScreen, "txtNotes2", 255, true);
}
function frmScreen.txtNotes3.onkeyup()
{
	scScreenFunctions.RestrictLength(frmScreen, "txtNotes3", 255, true);
}
function frmScreen.txtNotes4.onkeyup()
{
	scScreenFunctions.RestrictLength(frmScreen, "txtNotes4", 255, true);
}
function frmScreen.txtNotes5.onkeyup()
{
	scScreenFunctions.RestrictLength(frmScreen, "txtNotes5", 255, true);
}
function frmScreen.txtNotes6.onkeyup()
{
	scScreenFunctions.RestrictLength(frmScreen, "txtNotes6", 255, true);
} */%>
function btnCancel.onclick()
{
<%  /* SR - 15/05/00 SYS0404 - Use Applicaiton Mode instead of Meta Action   
	   APS - 07/06/00 SYS0152 - Route to MN060 on CANCEL from "Cost Modelling"*/
%>
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
	{
		if(m_sApplicationMode == "Cost Modelling")		
			frmToMN060.submit();
		else
			frmToQQ010.submit();
	}
}

function btnSubmit.onclick()
{
	if (CommitChanges())
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
			frmToGN300.submit();
		else
			frmToCM010.submit();
	}
}

function CommitChanges()
{
	var bSuccess = true;
	if(CheckTextAreaLength())
	{
		if (m_sReadOnly != "1")
			if(frmScreen.onsubmit())
			{
				if (IsChanged())
					bSuccess = SaveAttitudeToBorrowingDetails();
			}
			else
				bSuccess = false;
	}
	else
		bSuccess = false;
		
	return(bSuccess);
}

function CheckTextAreaLength()
{
	if(scScreenFunctions.RestrictLength(frmScreen, "txtNotes1", 255, true))
		return false;
	<% /* if(scScreenFunctions.RestrictLength(frmScreen, "txtNotes2", 255, true))
		return false;
	if(scScreenFunctions.RestrictLength(frmScreen, "txtNotes3", 255, true))
		return false;
	if(scScreenFunctions.RestrictLength(frmScreen, "txtNotes4", 255, true))
		return false;
	if(scScreenFunctions.RestrictLength(frmScreen, "txtNotes5", 255, true))
		return false;
	if(scScreenFunctions.RestrictLength(frmScreen, "txtNotes6", 255, true))
		return false; */ %>
	
	return true;
}

function CommitScreen()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	TagREQUEST = XML.CreateRequestTag(window,null);
	var TagRequestType = null;

	TagRequestType = XML.CreateActiveTag("CREATE");

	if(TagRequestType != null)
	{
		var TagATTITUDETOBORROWINGLIST = XML.CreateActiveTag("ATTITUDETOBORROWINGLIST");

		WriteAttitudeToBorrowingDetails(XML,TagATTITUDETOBORROWINGLIST);

		var thisFrame = document.frames("idXmlFrame")
		var thisForm = 	thisFrame.document.forms("idXmlForm");

		thisForm.idXmlReq.value = XML.XMLDocument.xml;

		frmXMLState.btnResponse.focus();
		frmXMLState.txtXMLState.value = "Request";

		divScreenCover.style.visibility = "visible";
		thisForm.submit();

		XML = null;
	}
}

function SaveAttitudeToBorrowingDetails()
{
	bReturn = false;
	nQuestions = 1 ;

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("CREATE");
	var TagParent = XML.CreateActiveTag("ATTITUDETOBORROWINGLIST");

	for(var nLoop = 0;nLoop < nQuestions;nLoop++)
	{
		XML.ActiveTag = TagParent;
		XML.CreateActiveTag("ATTITUDETOBORROWING");

		switch(nLoop)
		{
			case 0:
				sResponse = frmScreen.cboResponse1.value;
				sNotes = frmScreen.txtNotes1.value;
			break;

			<% /* case 1:
				sResponse = frmScreen.cboResponse2.value;
				sNotes = frmScreen.txtNotes2.value;
			break;

			case 2:
				sResponse = frmScreen.cboResponse3.value;
				sNotes = frmScreen.txtNotes3.value;
			break;

			case 3:
				sResponse = frmScreen.cboResponse4.value;
				sNotes = frmScreen.txtNotes4.value;
			break;

			case 4:
				sResponse = frmScreen.cboResponse5.value;
				sNotes = frmScreen.txtNotes5.value;
			break;

			case 5:
				sResponse = frmScreen.cboResponse6.value;
				sNotes = frmScreen.txtNotes6.value;
			break; */ %>

			default:
			break;
		}
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		XML.CreateTag("QUESTIONNUMBER", nLoop+1);
		XML.CreateTag("RESPONSETOQUESTION", sResponse);
		XML.CreateTag("NOTES", sNotes);
		
	}

	// 	XML.RunASP(document,"WriteAttitudeToBorrowingDetails.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			XML.RunASP(document,"WriteAttitudeToBorrowingDetails.asp");
			break;
		default: // Error
			XML.SetErrorResponse();
		}

	bReturn = XML.IsResponseOK();

	XML = null;
	return(bReturn);
}

-->
</script>
</body>
</html>

