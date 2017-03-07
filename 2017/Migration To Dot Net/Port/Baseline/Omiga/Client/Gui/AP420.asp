<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP420.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Search and Title
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JR		31/1/01		Screen Design
JR		7/2/01		Added functionality
CL		12/03/01	SYS2034 Read only functionality added
JR		15/03/01	SYS1878 Change create report on title to call omigatmbo.asp
JR		16/03/01	SYS2089 Added an attribs file for mandatory fields (completion date)
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
JR		22/03/01	SYS2048 Call to ValuationBO changed & correct setting radio btn attribute
JR		10/04/01	SYS2251 Translate middle-tier error msg to a more appropriate msg
JR		18/05/01	SYS2048	Added STAGEID and CASESTAGESEQUENCENO attributes to CreateReportonTitle request. 
BG		26/11/01	SYS3107 need to change the TaskXML held in context to reflect the task 
					has been moved to pending from not started in ReportOnTitleAction function.
SG		14/12/01	SYS2376 Check for Not A Number before calculating Retention Amount.
JLD		04/02/02	SYS3090 make purchase price and retention amount editable
GD		09/04/02	SYS4374 Fix PopulateRetentionAmount() typo.
LD		23/05/02	SYS4727 Use cached versions of frame functions
TW      09/11/02	Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		05/11/02	BMIDS00745 - Added Validation.js
MV		11/11/02	BMIDS00810	- Amended PopulateScreen();
DPF		20/11/02	BMIDS00810	- Amended PopulateScreen() to grab Purchase Price & Retention Amount on 1st
								  entry to screen and then pull them back with ROT data thereafter
DPF		21/11/02	BMIDS01049	- Catered for "" as well as Null in if condition for above change
HMA     18/09/03    BM0063      - Amend HTML text for radio buttons
MC		20/04/04	BMIDS517	- Pound sign display problem fixed £ char replaced with &pound;
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History : 

Prog	Date		AQR			Description
LDM		27/04/2006	MAR1624		Only create 1 report on title record
LDM		31/08/2006	MAR1947		Use the latest valuation report for the retention amt 
*/
%>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"> 
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>

<% /* Specify Forms Here */ %>
<form id="frmToAP400" method="post" action="AP400.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP430" method="post" action="AP430.asp" STYLE="DISPLAY: none"></form>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<form id="frmScreen" validate  ="onchange" mark year4>

<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>
<div id="divBackground" style="HEIGHT: 105px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Is the mortgage to be registered at H.M. Land Registry?
		<span style="LEFT: 370px; POSITION: absolute; TOP: -3px">
			<input id="optMortgageRegistered_Yes" name="MortgageRegistered" type="radio" value="1"><label for="optMortgageRegistered_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 430px; POSITION: absolute; TOP: -3px">
			<input id="optMortgageRegistered_No" name="MortgageRegistered" type="radio" value="0"><label for="optMortgageRegistered_No" class="msgLabel">No</label>
		</span>	
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Subject to first registration?
		<span style="LEFT: 370px; POSITION: absolute; TOP: -3px">
			<input id="optFirstRegistration_Yes" name="FirstRegistration" type="radio" value="1"><label for="optFirstRegistration_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 430px; POSITION: absolute; TOP: -3px">
			<input id="optFirstRegistration_No" name="FirstRegistration" type="radio" value="0"><label for="optFirstRegistration_No" class="msgLabel">No</label>
		</span>	
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Unregistered and recorded at Registry of Sasines/Deed in Northern Ireland?
		<span style="LEFT: 370px; POSITION: absolute; TOP: -3px">
			<input id="optSasines_Yes" name="Sasines" type="radio" value="1"><label for="optSasines_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 430px; POSITION: absolute; TOP: -3px">
			<input id="optSasines_No" name="Sasines" type="radio" value="0"><label for="optSasines_No" class="msgLabel">No</label>
		</span>	
	</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 85px" class="msgLabel">
		Unregistered and subject to registration?
		<span style="LEFT: 370px; POSITION: absolute; TOP: -3px">
			<input id="optUnregistered_Yes" name="Unregistered" type="radio" value="1"><label for="optUnregistered_Yes" class="msgLabel">Yes</label>
		</span>
		<span style="LEFT: 430px; POSITION: absolute; TOP: -3px">
			<input id="optUnregistered_No" name="Unregistered" type="radio" value="0" checked><label for="optUnregistered_No" class="msgLabel">No</label>
		</span>	
	</span>
</div>
<div id="divBackground" style="HEIGHT: 135px; LEFT: 10px; POSITION: absolute; TOP: 170px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Title No.
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="txtTitleNumber" maxlength="10" style="POSITION: absolute; WIDTH: 90px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 280px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Completion Date
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="txtCompletionDate" maxlength="10" style="POSITION: absolute; WIDTH: 90px" class="msgTxt">
		</span>
	</span>
	<!--  WP13 MAR49 - add Title 2, Title 3  -->
	<span style="LEFT: 10px; POSITION: absolute; TOP: 35px" class="msgLabel">
		Title No.
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="txtTitleNumber2" maxlength="10" style="POSITION: absolute; WIDTH: 90px" class="msgTxt">
		</span>
	</span>
		<span style="LEFT: 10px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Title No.
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="txtTitleNumber3" maxlength="10" style="POSITION: absolute; WIDTH: 90px" class="msgTxt">
		</span>
	</span>

	<span style="LEFT: 280px; POSITION: absolute; TOP: 85px" class="msgLabel">
		Reference
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="txtReference" maxlength="10" style="POSITION: absolute; WIDTH: 90px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 280px; POSITION: absolute; TOP: 110px" class="msgLabel">
		Solicitor's Ref Number
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="txtSolicitorRefNum" maxlength="10" style="POSITION: absolute; WIDTH: 90px" class="msgTxt">
		</span>
	</span>
	
<!--  DPF 20/11/2002 - BMIDS00810 - make Purchase Price & Retention Amount read only -->
	<span style="LEFT: 10px; POSITION: absolute; TOP: 85px" class="msgLabel">
		Purchase Price
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="txtPurchasePrice" maxlength="10" style="POSITION: absolute; WIDTH: 90px" class="msgTxt" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>
	<span style="LEFT: 110px; POSITION: absolute; TOP: 85px" class="msgLabel">&pound;</span>
	<span style="LEFT: 10px; POSITION: absolute; TOP: 110px" class="msgLabel">
		Retention Amount	
		<span style="LEFT: 110px; POSITION: absolute; TOP: -3px">
			<input id="txtRetentionAmount" maxlength="10" style="POSITION: absolute; WIDTH: 90px" class="msgTxt" class="msgReadOnly" readonly tabindex=-1>
		</span>
	</span>
		<span style="LEFT: 110px; POSITION: absolute; TOP: 110px" class="msgLabel">&pound;</span>
</div>	
<div id="divBackground" style="HEIGHT: 100px; LEFT: 10px; POSITION: absolute; TOP: 310px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 10px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Except for the following we have nothing to report:
		<span style="LEFT: 0px; POSITION: absolute; TOP: 16px">
			<TEXTAREA id=txtFurtherNotes rows=4 style="WIDTH: 470px" maxlength="255" class=msgTxt ></TEXTAREA>
		</span>
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 440px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->


<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP420attribs.asp" -->
<%/* CODE */%>
<script language="JScript">
<!--

var m_sApplicationNumber, m_sApplicationFactFindNumber ;
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sUserRole = "" ;
var m_sUserId = "" ;
var m_sUnitId = "" ;
var m_sTaskXML = "";

var scScreenFunctions;
var XML = null;
var taskXML = null ;
var rotXML = null ;
var xmlresponseNode ;
var xmlDataHeaderKeys ;

var m_blnReadOnly = false;
var m_EntryStatus = null;


/** EVENTS **/


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel","Next");
	ShowMainButtons(sButtonList);
	
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("Details of Search & Title","AP420",scScreenFunctions);
	
	RetrieveContextData();			
	SetMasks();
	Validation_Init();
	Initialise();
			
	scScreenFunctions.SetFocusToFirstField(frmScreen);
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP420");
	
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
	/* TEST
	
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","C00059072");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber","00000094");
	scScreenFunctions.SetContextParameter(window,"idMortgageApplicationValue","10");
	
	END TEST */
		
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	
	m_sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	m_sUserId = scScreenFunctions.GetContextParameter(window,"idUserId","");
	m_sUnitId =	scScreenFunctions.GetContextParameter(window,"idUnitId","");
		
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = 	scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","");
		
//DEBUG
//m_sMetaAction = "0" ;
//m_sTaskXML = "<CASETASK TASKID=\"22\" CUSTOMERIDENTIFIER=\"1\" CONTEXT=\"2\" TASKSTATUS=\"20\"/>";	
}

function Initialise()
{
	if(m_sTaskXML.length > 0) 
	{
		taskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject() ;
		taskXML.LoadXML(m_sTaskXML) ;
		taskXML.ActiveTag = null;
		taskXML.SelectTag(null, "CASETASK");
		m_EntryStatus = taskXML.GetAttribute ("TASKSTATUS");<%/* LDM 27/04/2006 MAR1624 need to know what the task status on entry to the screen */%> 
		PopulateScreen();
	}
}

function PopulateScreen()
{
	<%/* LDM 27/04/2006 MAR1624 */%> 
	rotXML = null ;
	rotXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	rotXML.CreateRequestTag(window , "GetReportOnTitleData");
	rotXML.CreateActiveTag("REPORTONTITLE");
	rotXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	rotXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);

	//Pass XML to ApplicationManager BO
	rotXML.RunASP(document,"ReportOnTitle.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = rotXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{
		taskXML.SetAttribute("TASKSTATUS",20);	<%/* LDM 27/04/2006 MAR1624 force to be a status of update if ROT rec is there already */%> 
		scScreenFunctions.SetContextParameter(window,"idTaskXML",taskXML.XMLDocument.xml);
	}
	<%/* LDM 27/04/2006 MAR1624 */%> 

	if (taskXML.GetAttribute("TASKSTATUS") != "10")
	{
		if (ErrorReturn[0] == true)
		{
			// Process the returning XML
			rotXML.ActiveTag = null;
			rotXML.SelectTag(null, "REPORTONTITLE");
						
			//populate radio/text fields
			scScreenFunctions.SetRadioGroupValue(frmScreen, "MortgageRegistered", rotXML.GetAttribute("HMLREGISTERED")) ;
			scScreenFunctions.SetRadioGroupValue(frmScreen, "FirstRegistration", rotXML.GetAttribute("FIRSTREGISTRATION")) ;
			scScreenFunctions.SetRadioGroupValue(frmScreen, "Sasines", rotXML.GetAttribute("UNREGISTEREDRECORDED")) ;
			scScreenFunctions.SetRadioGroupValue(frmScreen, "Unregistered", rotXML.GetAttribute("UNREGISTEREDNOREG")) ;
			frmScreen.txtTitleNumber.value		= (rotXML.GetAttribute("TITLENUMBER") != null) ? rotXML.GetAttribute("TITLENUMBER") : "" ;
			<% /* WP13 MAR49 added TITLENUMBER2, TITLENUMBER3 */ %>
			frmScreen.txtTitleNumber2.value		= (rotXML.GetAttribute("TITLENUMBER2") != null) ? rotXML.GetAttribute("TITLENUMBER2") : "" ;
			frmScreen.txtTitleNumber3.value		= (rotXML.GetAttribute("TITLENUMBER3") != null) ? rotXML.GetAttribute("TITLENUMBER3") : "" ;
			frmScreen.txtCompletionDate.value	= (rotXML.GetAttribute("COMPLETIONDATE") != null) ? rotXML.GetAttribute("COMPLETIONDATE") : "" ;
			frmScreen.txtReference.value		= (rotXML.GetAttribute("REFERENCE") != null) ? rotXML.GetAttribute("REFERENCE") : "" ;
			frmScreen.txtSolicitorRefNum.value	= (rotXML.GetAttribute("SOLICITORSREFNUMBER") != null) ? rotXML.GetAttribute("SOLICITORSREFNUMBER") : "" ;
			frmScreen.txtFurtherNotes.value		= (rotXML.GetAttribute("REPORTNOTES") != null) ? rotXML.GetAttribute("REPORTNOTES") : "" ;
		
			//DPF - BMIDS00810
			//populate puchase price textbox - value should be in ROT xml but if not call PurchasePrice function
			if (rotXML.GetAttribute("PURCHASEPRICE") != null && rotXML.GetAttribute("PURCHASEPRICE") != "" )
				frmScreen.txtPurchasePrice.value = rotXML.GetAttribute("PURCHASEPRICE");
			else
				PopulatePurchasePrice() ;
			
			
			<%/* LDM 31/8/06 MAR1947 Always get the retention amount from the getvaluationreport table rather than the reportontitle table.
			   A new valuation may have come in since the reportontitle was updated.
		
			//Populate security address textbox -  - value should be in ROT xml but if not call PurchasePrice function
			if (rotXML.GetAttribute("RETENTIONAMOUNT") != null && rotXML.GetAttribute("RETENTIONAMOUNT") != "")
				frmScreen.txtRetentionAmount.value =  rotXML.GetAttribute("RETENTIONAMOUNT");
			else	*/%>
			PopulateRetentionAmount() ;				
		}
	}
	else
	{
		//DPF - BMIDS00810 - grab the values on 1st call in (STATUS = InComplete) for Purchase Price and Retention Amount
		PopulatePurchasePrice() ;
		PopulateRetentionAmount() ;	
	}

	if (m_sReadOnly == "1") scScreenFunctions.SetScreenToReadOnly(frmScreen) ;
}

function PopulatePurchasePrice()
{
	var appXML = null ;
	appXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	appXML.CreateRequestTag (window, "GetApplicationData") ;
	appXML.CreateActiveTag("APPLICATION") ;
	appXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber) ;

	//Pass XML to ApplicationManager BO
	appXML.RunASP(document,"GetApplicationData.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = appXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{		
		if (appXML.SelectTag(null, "APPLICATIONFACTFIND") != null)
			frmScreen.txtPurchasePrice.value  = appXML.GetTagText ("PURCHASEPRICEORESTIMATEDVALUE") ;
	}
}

function PopulateRetentionAmount()
{
	var valuationXML = null ;
	valuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	valuationXML.CreateRequestTag (window, "GetValuationReport") ;
	valuationXML.CreateActiveTag("VALUATION") ;
	valuationXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber) ;
	valuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	
	//Pass XML to ApplicationProcessing
	valuationXML.RunASP(document,"omAppProc.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = valuationXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{
		var iRetWorksVal = 0;
		var iRetRoadsVal = 0;
		if (valuationXML.SelectTag(null, "GETVALUATIONREPORT") != null)
		{
			iRetWorksVal = parseInt(valuationXML.GetAttribute("RETENTIONWORKS")) ;
			iRetRoadsVal = parseInt(valuationXML.GetAttribute("RETENTIONSROADS")) ;
			
		<%	//SG 14/12/01 SYS2376 Check for Not A Number.
		%>	if (isNaN(iRetWorksVal)) iRetWorksVal = 0;
			if (isNaN(iRetRoadsVal)) iRetRoadsVal = 0;	
			
			frmScreen.txtRetentionAmount.value = iRetWorksVal + iRetRoadsVal ;

			<%/* LDM 31/8/06 MAR1947 Force the saving of details. A new valuation 
				may have come in since the reportontitle was last created/amended. */%>
			FlagChange(true);
		}
	}
}

function ReportOnTitleAction()
{
	XML = null ;
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag ;

	if ( (taskXML.GetAttribute("TASKID") == null) || (m_sReadOnly == "1") ) return true;

	if ( (scScreenFunctions.GetRadioGroupValue(frmScreen,"MortgageRegistered") == "1") && (frmScreen.txtTitleNumber.value == "") )
	{
		alert("HML Registered has been selected but no Title Number is entered") ;
		frmScreen.txtTitleNumber.select() ;
		return false ;
	} 
	
	//Check completion date valid for processing
	if (!CheckCompletionDate()) return false ;
	 		
	if (taskXML.GetAttribute("TASKSTATUS") == "20")
	{
		reqTag = XML.CreateRequestTag(window , "UpdateReportOnTitle");
									
		//Complete the XML with ROT Attributes 
		SetROTAttributes(reqTag, "Update") ; 
				
		//Pass XML to omRotBO
		// 		XML.RunASP(document,"ReportOnTitle.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"ReportOnTitle.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	
		//Verify XML response
		<% /* LDM 27/04/2006 MAR1624 */ %>
		if(XML.IsResponseOK())
		{
			if (m_EntryStatus == "20")
			{
				return true;
			}
			else (m_EntryStatus == "10")
			{
				<% /* LDM 27/04/2006 MAR1624 - If everything worked ok, need to set the task to pending.
					If this is the first time into the cot screens by "any" action; then the status will be automatically moved
					to pending, when the createreportontitle is run.
					If we got here by the details btn in the application stage. Then the status is already pending, so do not reset to pending
					If we got here by creating another cot action and this is the first time in need to reset the status to pending  */ %>
				SetToPendingXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
				SetToPendingXML.CreateRequestTag(window , "UpdateCaseTask");
				SetToPendingXML.CreateActiveTag("CASETASK");
				SetToPendingXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
				SetToPendingXML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
				SetToPendingXML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
				SetToPendingXML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
				SetToPendingXML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
				SetToPendingXML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
				SetToPendingXML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
				SetToPendingXML.SetAttribute("TASKSTATUS", "20"); <% /* Pending */ %>
				
				SetToPendingXML.RunASP(document, "msgTMBO.asp") ;
				
				if (SetToPendingXML.IsResponseOK())
				{
					return true;
				}
				else
				{
					return false;
				}
			}
		}
		else
		{
			alert("Error updating ROT") ;
			return false;
		}
		<% /* LDM 27/04/2006 MAR1624 */ %>
	}
	else
	{
		if (taskXML.GetAttribute("TASKSTATUS") == "10")
		{					
			reqTag = XML.CreateRequestTag(window , "CreateReportOnTitle");
			XML.SetAttribute("USERID", m_sUserId) ;
			XML.SetAttribute("UNITID", m_sUnitId) ;
			XML.SetAttribute("USERAUTHORITYLEVEL", m_sUserRole) ;
		
			XML.ActiveTag = reqTag ;
			XML.CreateActiveTag("CASETASK");
			XML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION"));
			XML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID"));
			XML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID"));
			XML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID"));
			XML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE"));
			// JR - SYS2048, add the following attributes to Request
			XML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
			XML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
									
			//Complete the XML with ROT Attributes
			SetROTAttributes(reqTag, "Create") ;
		
			//Pass XML to omRotBO
			// 			XML.RunASP(document,"OmigaTMBO.asp");
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
			//BG 26/11/01 SYS3107 need to change the TaskXML held in context to reflect
			//the task has been moved to pending from not started.
			//if(!XML.IsResponseOK())	return false ;
		
			if(!XML.IsResponseOK())	
			{
				return false ;
			}
			else
			{
				taskXML.SetAttribute("TASKSTATUS",20);
				scScreenFunctions.SetContextParameter(window,"idTaskXML",taskXML.XMLDocument.xml);
			}		
		}
	}	
	return true ;
}

function SetROTAttributes(reqTag, sOperation) 
{
	var strTest ;
	
	//build ReportOnTitle attributes		 
	XML.ActiveTag = reqTag ;
	XML.CreateActiveTag("REPORTONTITLE");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber) ;
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber) ;
	if (sOperation == "Update") XML.SetAttribute("ROTGUID", rotXML.GetAttribute("ROTGUID")) ;
	
	//SYS2048 Prevent error message if radio buttons not set
	strTest	= (scScreenFunctions.GetRadioGroupValue(frmScreen,"MortgageRegistered") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"MortgageRegistered") : "" ;
	XML.SetAttribute("HMLREGISTERED", strTest) ;
	strTest = (scScreenFunctions.GetRadioGroupValue(frmScreen,"FirstRegistration") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"FirstRegistration") : "" ; 
	XML.SetAttribute("FIRSTREGISTRATION", strTest) ;
	strTest = (scScreenFunctions.GetRadioGroupValue(frmScreen,"Sasines") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"Sasines") : "" ;
	XML.SetAttribute("UNREGISTEREDRECORDED", strTest) ;
	strTest = (scScreenFunctions.GetRadioGroupValue(frmScreen,"Unregistered") != null) ? scScreenFunctions.GetRadioGroupValue(frmScreen,"Unregistered") : "" ;
	XML.SetAttribute("UNREGISTEREDNOREG", strTest) ;
	
	XML.SetAttribute("TITLENUMBER", frmScreen.txtTitleNumber.value) ;
	<% /* WP13 MAR49 added TITLENUMBER2 TITLENUMBER3*/ %>
	XML.SetAttribute("TITLENUMBER2", frmScreen.txtTitleNumber2.value) ;
	XML.SetAttribute("TITLENUMBER3", frmScreen.txtTitleNumber3.value) ;
	XML.SetAttribute("PURCHASEPRICE", frmScreen.txtPurchasePrice.value) ;
	XML.SetAttribute("RETENTIONAMOUNT", frmScreen.txtRetentionAmount.value) ;
	XML.SetAttribute("COMPLETIONDATE", frmScreen.txtCompletionDate.value) ;
	XML.SetAttribute("REFERENCE", frmScreen.txtReference.value) ;
	XML.SetAttribute("SOLICITORSREFNUMBER", frmScreen.txtSolicitorRefNum.value) ;
	XML.SetAttribute("REPORTNOTES", frmScreen.txtFurtherNotes.value) ;
}

function CheckCompletionDate()
{
	var dateXML = new top.frames[1].document.all.scXMLFunctions.XMLObject() ;
	dateXML.CreateRequestTag(window , null);
	dateXML.CreateActiveTag("SYSTEMDATE");
	dateXML.CreateTag("DATE", frmScreen.txtCompletionDate.value);
	dateXML.CreateTag("CHANNELID", "SD1");
		
	//Pass XML to omBase.SystemDatesBO
	// 	dateXML.RunASP(document,"CheckNonWorkingOccurence.asp");
	// Added by automated update TW 09 Oct 2002 SYS5115
	switch (ScreenRules())
		{
		case 1: // Warning
		case 0: // OK
			dateXML.RunASP(document,"CheckNonWorkingOccurence.asp");
			break;
		default: // Error
			dateXML.SetErrorResponse();
		}

	
	//Verify XML response
	if(dateXML.IsResponseOK())
	{	
		dateXML.SelectTag(null, "SYSTEMDATE") ;
		if (dateXML.GetTagText("NONWORKINGIND") == "1") 
		{
			alert("The completion date does not fall on a working day") ;
			frmScreen.txtCompletionDate.select() ;
			return false ;
		}
		return true ;
	}
	return false ;
}

function CommitChanges()
{
	var bSuccess = true;
	//if (m_sReadOnly != "1")
	if(frmScreen.onsubmit())
	{
		if (IsChanged()) bSuccess = ReportOnTitleAction();
	}
	else
		bSuccess = false;
	return(bSuccess);
}

function btnSubmit.onclick()
{
	if (CommitChanges()) frmToAP400.submit() ;
}

function btnNext.onclick()
{
	if (CommitChanges()) frmToAP430.submit() ;
}

function btnCancel.onclick()
{
	taskXML.SetAttribute("TASKSTATUS",m_EntryStatus); <%/* LDM 02/04/2006 MAR1624 reset to initial status */%> 
	scScreenFunctions.SetContextParameter(window,"idTaskXML",taskXML.XMLDocument.xml); <%/* LDM 02/05/2006 MAR1624 */%>
	frmToAP400.submit() ;
}
-->
</script>
</body>
</HTML>


