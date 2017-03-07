<%@ LANGUAGE="JSCRIPT" %>
<html>
<%
/*
Workfile:      AP440.asp
Copyright:     Copyright © 2000 Marlborough Stirling
Description:   Solicitors Bank Account Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Descrip
JR		1/2/01		Screen Design
JR		9/2/01		Added functionlaity
CL		12/03/01	SYS1920 Read only functionality added
JR		15/03/01	SYS1878 Change create report on title to call omigatmbo.asp
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
JR		28/03/01	SYS2048 Re-enable the main "Next" navigation button
JR		18/05/01	SYS2048	Added STAGEID and CASESTAGESEQUENCENO attributes to CreateReportonTitle request
					Add ROTGUID on Updates and not create
					Add AP440attribs.asp file
DRC     18/07/01    SYS2251 Input field widths altered to comply with database
BG		26/11/01	SYS3107 need to change the TaskXML held in context to reflect the task 
					has been moved to pending from not started in ReportOnTitleAction function.
MEVA	22/04/02	SYS4440 Solicitors Bank Details currency field is using incorrect combogroup.				
MEVA	08/04/02	SYS4541 Population correction
LD		23/05/02	SYS4727 Use cached versions of frame functions
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History : 

Prog	Date		AQR			Description
LDM		27/04/2006	MAR1624		Only create 1 report on title record. screen not used in Mars. Put code in anyway, as did 410,420 & 430
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
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<script src="validation.js" language="JScript"></script>

<% /* Specify Forms Here */ %>
<form id="frmToAP400" method="post" action="AP400.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP450" method="post" action="AP450.asp" STYLE="DISPLAY: none"></form>

<% /* span field to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>
<form id="frmScreen" validate      ="onchange" mark year4>
<div id="divBackground" style="HEIGHT: 180px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 120px; POSITION: absolute; TOP: 30px" class="msgLabel">
		Bank Name
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtBankName" maxlength="45" style="HEIGHT: 20px; LEFT: 0px; POSITION: absolute; TOP: -3px; WIDTH: 224px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 120px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Bank Account Name
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountName" maxlength="55" style="HEIGHT: 20px; LEFT: 2px; POSITION: absolute; TOP: 0px; WIDTH: 292px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 120px; POSITION: absolute; TOP: 90px" class="msgLabel">
		Bank Sort Code
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtSortCode" maxlength="8" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 50px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 120px; POSITION: absolute; TOP: 120px" class="msgLabel">
		Bank Account Number
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountNumber" maxlength="20" style="HEIGHT: 20px; POSITION: absolute; WIDTH: 120px" class="msgTxt">
		</span>
	</span>
	<span style="LEFT: 120px; POSITION: absolute;  TOP: 150px" class="msgLabel">
		Account Currency
		<span style="LEFT: 120px; POSITION: absolute; TOP: -3px">
			<select id="cboAccountCurrency" style="HEIGHT: 20px; WIDTH: 78px" class="msgCombo"></select>
		</span>
	</span>
</div>
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 260px; WIDTH: 612px"> 
<!-- #include FILE="msgButtons.asp" --> 
</div>

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<% /* File containing field attributes (remove if not required) */ %>
<!-- #include FILE="attribs/AP440attribs.asp" -->
<%/* CODE */%>
<script language="JScript">
<!--

var RotXML = null ;
var taskXML = null;
var XML = null ;
var m_sUnitId = "" ;
var m_sUserId = "" ;
var m_sUserRole = "" ;
var m_sMetaAction = "" ;
var m_sReadOnly = "";
var m_sTaskXML = "";

var scScreenFunctions;
var strRotGuid ;
var m_sApplicationNumber, m_sApplicationFactFindNumber ;
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
	FW030SetTitles("Solicitor's Bank Account Details","AP440",scScreenFunctions);
		
	RetrieveContextData();	
	SetMasks();
	PopulateCurrencyCombo() ;		
	Validation_Init();	
	Initialisation() ;	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP440");
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

function RetrieveContextData()
{
	 /*TEST
	
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","C00059072");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");
	scScreenFunctions.SetContextParameter(window,"idCustomerNumber","00000094");
	
	//END TEST */
		
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly",null);
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	
	m_sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0")
	m_sUserId = scScreenFunctions.GetContextParameter(window,"idUserId","");
	m_sUnitId =	scScreenFunctions.GetContextParameter(window,"idUnitId","");
		
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
	m_sApplicationFactFindNumber = 	scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","");
		
//DEBUG
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
		m_EntryStatus = taskXML.GetAttribute ("TASKSTATUS");<%/* LDM 27/04/2006 MAR1624 need to know what the task status on entry to the screen */%> 		
		PopulateScreen() ;
	}	
}

function PopulateScreen()
{
	<%/* LDM 27/04/2006 MAR1624 */%> 
	RotXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	RotXML.CreateRequestTag(window , "GetReportOnTitleData");
	RotXML.CreateActiveTag("REPORTONTITLE");
	RotXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
	RotXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	
	//Pass XML to ReportOnTitle BO
	RotXML.RunASP(document,"ReportOnTitle.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = RotXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true)
	{
		taskXML.SetAttribute("TASKSTATUS",20);	<%/* LDM 27/04/2006 MAR1624 force to be a status of update if ROT rec is there already */%> 
		scScreenFunctions.SetContextParameter(window,"idTaskXML",taskXML.XMLDocument.xml);
	}
	<%/* LDM 27/04/2006 MAR1624 */%> 

	if (taskXML.GetAttribute ("TASKSTATUS") != "10")
	{
		//Verify XML response
		if(RotXML.IsResponseOK())
		{
			// Process the returning XML
			RotXML.SelectTag(null, "REPORTONTITLE");
			strRotGuid = RotXML.GetAttribute("ROTGUID") ;
			
			RotXML.ActiveTag = null;
			RotXML.SelectTag(null, "SOLBANKACCT");
			frmScreen.txtBankName.value			= (RotXML.GetAttribute("BANKNAME") != null) ? RotXML.GetAttribute("BANKNAME") : "" ;
			frmScreen.txtAccountName.value		= (RotXML.GetAttribute("BANKACCOUNTNAME") != null) ? RotXML.GetAttribute("BANKACCOUNTNAME") : "" ;
			frmScreen.txtSortCode.value			= (RotXML.GetAttribute("BANKSORTCODE") != null) ? RotXML.GetAttribute("BANKSORTCODE") : "" ;
			frmScreen.txtAccountNumber.value	= (RotXML.GetAttribute("BANKACCOUNTNUMBER") != null) ? RotXML.GetAttribute("BANKACCOUNTNUMBER") : "" ;
			frmScreen.cboAccountCurrency.value	= RotXML.GetAttribute("ACCOUNTCURRENCY") ;
		}
		else
		{
			alert('Error retreiving GetReportOnTitleData results');
			return ;
		}
	}
	if (m_sMetaAction == "1") scScreenFunctions.SetScreenToReadOnly(frmScreen) ;
}

function PopulateCurrencyCombo()
{
	var currencyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroups = new Array("AccountCurrency");
	var bSuccess = false;
	if (currencyXML.GetComboLists(document, sGroups) == true)
	{
		bSuccess = true;
		bSuccess = bSuccess & currencyXML.PopulateCombo(document,frmScreen.cboAccountCurrency,"AccountCurrency",true);
	}
}

function ReportOnTitleAction()
{
	XML = null ;
	var reqTag ;

	if ( (taskXML.GetAttribute ("TASKID") == null) || (m_sMetaAction == "1") ) return true ;

	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject()
	if (taskXML.GetAttribute ("TASKSTATUS") == "20")
	{
		reqTag = XML.CreateRequestTag(window , "UpdateReportOnTitle");
											
		//Complete the XML with ROT Attributes 
		SetROTAttributes(reqTag, "Update") ; 
					
		//Pass XML to ROTBO
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
		if (taskXML.GetAttribute ("TASKSTATUS") == "10")
		{					
			reqTag = XML.CreateRequestTag(window , "CreateReportOnTitle");
			XML.SetAttribute("USERID", m_sUserId) ;
			XML.SetAttribute("UNITID", m_sUnitId) ;
			XML.SetAttribute("USERAUTHORITYLEVEL", m_sUserRole) ;
		
			XML.ActiveTag = reqTag ;
			XML.CreateActiveTag("CASETASK");
			XML.SetAttribute("SOURCEAPPLICATION", taskXML.GetAttribute("SOURCEAPPLICATION")) ;
			XML.SetAttribute("ACTIVITYID", taskXML.GetAttribute("ACTIVITYID")) ;
			XML.SetAttribute("CASEID", taskXML.GetAttribute("CASEID")) ;
			XML.SetAttribute("TASKID", taskXML.GetAttribute("TASKID")) ;
			XML.SetAttribute("TASKINSTANCE", taskXML.GetAttribute("TASKINSTANCE")) ;
			// JR - SYS2048, add the following attributes to Request
			XML.SetAttribute("STAGEID", taskXML.GetAttribute("STAGEID"));
			XML.SetAttribute("CASESTAGESEQUENCENO", taskXML.GetAttribute("CASESTAGESEQUENCENO"));
														
			//Complete the XML with ROT Attributes
			SetROTAttributes(reqTag, "Create") ;
						
			//Pass XML to OmTmBO
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
	XML.ActiveTag = reqTag ;
	XML.CreateActiveTag("REPORTONTITLE");
	XML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber) ;
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber) ;
	// JR - SYS2048, Add RotGuid only on Updates
	if (sOperation == "Update") XML.SetAttribute("ROTGUID", strRotGuid) ;
		
	XML.CreateActiveTag("SOLBANKACCT");
	XML.SetAttribute("BANKNAME", frmScreen.txtBankName.value) ;
	XML.SetAttribute("BANKACCOUNTNAME", frmScreen.txtAccountName.value) ;
	XML.SetAttribute("BANKACCOUNTNUMBER", frmScreen.txtAccountNumber.value) ;
	XML.SetAttribute("BANKSORTCODE", frmScreen.txtSortCode.value) ;
	XML.SetAttribute("ACCOUNTCURRENCY", frmScreen.cboAccountCurrency.value) ;
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

function btnSubmit.onclick ()
{
	if (CommitChanges()) frmToAP400.submit() ;
}

function btnCancel.onclick ()
{
	taskXML.SetAttribute("TASKSTATUS",m_EntryStatus); <%/* LDM 02/04/2006 MAR1624 reset to initial status */%> 
	scScreenFunctions.SetContextParameter(window,"idTaskXML",taskXML.XMLDocument.xml); <%/* LDM 02/05/2006 MAR1624 */%>
	frmToAP400.submit() ;
}

function btnNext.onclick ()
{
	if (CommitChanges()) frmToAP450.submit() ;	
}

-->
</script>
</body>
</HTML>

