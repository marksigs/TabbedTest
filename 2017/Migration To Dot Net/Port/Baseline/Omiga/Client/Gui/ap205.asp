<%@ LANGUAGE="JSCRIPT" %>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en">
<html>
<%
/*
Workfile:      AP205.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description: 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		01/02/01	SYS1839 Created
DJP		13/03/01	SYS1839 Always enable OK button
ADP		16/03/01	SYS2091	Standardise size of menu buttons
DJP		19/03/01	SYS2114 Move Submit button onto Main list of buttons
IK		21/03/01	SYS1924 use idTaskXML for context not idXML 
					(idXML gets over-written in task processing screens)
DRC     06/06/01    SYS2271 Added customer list to xml for Validating the valuation 
                    report. Also added the facility to add a Remodel Quote task in
                    the event of the LTV changing as a result of the valuation					
DRC     21/11/01    SYS3048 Get Instruction Sequence No if not present in the Task Context and put it
                    into the omiga menu context    
PSC		12/12/01	SYS3388 Prompt before running confirm process and always remodel if
					LTV changed                                    
AT		30/04/02	SYS4359 Enable client versions to override labels
LD		23/05/02	SYS4727 Use cached versions of frame functions
TW      09/10/02    SYS5115 Modified to incorporate client validation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specifc History:

Prog	Date		AQR			Description
SA		17/11/2002	BMIDS00473	Put application under review if rules fail.
BS		12/02/2003	BM0291		Display LTV if it has changed
GD		01/07/2003	BM0582		Check for presence of  'Present Valuation' before confirming from screen.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
INGDUK Specific History:
Prog    Date		Description
JD		22/09/2005	MAR40 Changed on confirm processing
JD		14/02/2006	MAR1241 add check for valuationtype and appointmentdate on OK and Confirm.
JD		21/03/2006	MAR1434 check which rules are called in validateValuationReport
GHun	31/02/2006	MAR1300 Change of property changes
GHun	11/04/2006	MAR1607 Change ConfirmChangeOfProperty to complete ChgOfProperty Remodel task
GHun	19/04/2006	MAR1607 Back out previous change as it should have been on PP010
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/

%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css"/>
	<title></title>
</head>
<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" viewastext tabIndex="-1"></object>
<!-- Form Manager Object - Controls Soft Coded Field Attributes -->
<!-- Screen Functions Object - see comments within scScreenFunctions.htm for details -->

<!-- Specify Forms Here -->
<form id="frmToAP200" method="post" action="AP200.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP210" method="post" action="AP210.asp" STYLE="DISPLAY: none"></form>
<form id="frmToTM030" method="post" action="TM030.asp" STYLE="DISPLAY: none"></form>
<form id="frmToAP030" method="post" action="AP030.asp" STYLE="DISPLAY: none"></form>

<span id="spnToLastField" tabindex="0"></span>
<!-- Specify Screen Layout Here -->
<form id="frmScreen" validate ="onchange" mark>
<div id="divBackground" style="HEIGHT: 230px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<br/>
	<br/>
	<table border="0" width="100%" cellpadding="8">
		<tr>
			<td align="center">
				<input id="btnValuation" value="" type="button" style="HEIGHT: 40px; WIDTH: 250px" class="msgButton">
			</td>
		<tr>
			<td align="center">
				<input id="btnPropertyDetails" value="Property Details" type="button" style="HEIGHT: 40px; WIDTH: 250px" class="msgButton">
			</td>
		</tr>
		<tr>
			<td align="right">
				<br/>
				<input id="btnReturnReassign" value="Return/Reassign" type="button" style="WIDTH: 100px" class="msgButton">
			</td>
		</tr>
	</table>
</div>				
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute;  TOP: 295px; WIDTH: 612px">
	<!-- #include FILE="msgButtons.asp" -->
</div>

<span id="spnToFirstField" tabindex="0"></span>

<!-- #include FILE="fw030.asp" -->
<!-- #include FILE="Customise/AP205Customise.asp" -->

<% /* File containing field attributes */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<!-- #include FILE="attribs/ap205Attribs.asp" -->

<script language="JSCRIPT" type="text/javascript">
var scScreenFunctions;
var m_TaskXML = null;
var m_sTaskXML = "";
var m_sReadOnly = null;
var m_sAppNo = "";
var m_sAppFactFindNo = "";
var m_sTaskID = "";
var m_sInsSeqNo = "";
var csTaskPending = "20";
var csTaskComplete = "40";
var csValuerStatusComplete = "30";
var csValuerTypeStaff = "10";
var m_sCallingScreen = "";
var m_sContextParams = "";
var m_sValuationType = "";
var m_sAppointmentDate = "";

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	RetrieveContextData();
	var sButtonList = new Array("Submit","Confirm");

	ShowMainButtons(sButtonList);
	FW030SetTitles("Valuation Report Summary","AP205",scScreenFunctions);
<%	// This function sets the correct currency character on any currency fields
	// Remove if not required
%>	scScreenFunctions.SetCurrency(window,frmScreen);
	Customise();

<%	//This function is contained in the field attributes file (remove if not required)
%>	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	SetMasks();
	Validation_Init();
	PopulateScreen();
 	scScreenFunctions.SetFocusToFirstField(frmScreen);
	frmScreen.btnReturnReassign.disabled =true;

	if(m_sReadOnly != "1")
	{
		var bReadOnly = false;
		bReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "AP205");	
		if(bReadOnly)
		{
			m_sReadOnly = "1";		
		}
	}
	
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
	var sParameters;
	var sCustomerNo;
	var sCustVerNo;
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	m_sTaskXML = scScreenFunctions.GetContextParameter(window,"idTaskXML",null);
	if(m_sTaskXML != "")
	{
		m_TaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		m_TaskXML.LoadXML(m_sTaskXML);
		m_TaskXML.SelectTag(null, "CASETASK");
		m_sInsSeqNo = m_TaskXML.GetAttribute("CONTEXT");
		m_sTaskID = m_TaskXML.GetAttribute("TASKID");
	}

	m_sAppNo = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
	m_sAppFactFindNo = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null);
	m_sCallingScreen = scScreenFunctions.GetContextParameter(window,"idREturnScreenId",null);
}

function PopulateScreen()
{
	var sTaskStatus;
	var bDisableConfirm = true;

	if(m_sReadOnly != "1")
	{
		if(m_sTaskID != null)
		{
			sTaskStatus = m_TaskXML.GetAttribute("TASKSTATUS");
			if(csTaskComplete == sTaskStatus  && m_sCallingScreen == "TM030")
			{
				m_sReadOnly = "1";
				scScreenFunctions.SetContextParameter(window,"idMetaAction", m_sReadOnly);
			}
			else if( sTaskStatus == csTaskPending)
			{
				bDisableConfirm = false;
			}
		}
	}

	if(bDisableConfirm)
	{
		DisableMainButton("Confirm");
	}
	else
	{
		EnableMainButton("Confirm");
	}

	frmScreen.btnReturnReassign.disabled = bDisableConfirm;
	
	GetValuerInstructions();
	
	if(m_sReadOnly == "1")
	{
		scScreenFunctions.SetScreenToReadOnly(frmScreen);
		frmScreen.btnReturnReassign.disabled = true;
		DisableMainButton("Confirm");
	}
}

function GetValuerInstructions()
{
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = XML.CreateRequestTag(window , "GetValuerInstructions");

	XML.CreateActiveTag("VALUERINSTRUCTIONS");
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
	XML.SetAttribute("INSTRUCTIONSEQUENCENO",m_sInsSeqNo);

	XML.RunASP(document, "omAppProc.asp");

	if(XML.IsResponseOK())
	{
		XML.SelectTag(null, "VALUERINSTRUCTIONS");
		var sValuerType;
		
		if(m_sReadOnly != "1")
		{		
			sValuerType = XML.GetAttribute("VALUERTYPE");
			if( sValuerType != csValuerTypeStaff)
			{
				frmScreen.btnReturnReassign.disabled = true;			
			}
			<%//AQR SYS 3048 get the Instruction sequence number and put it into context%>
			if (m_sInsSeqNo == "")
			{
				m_sInsSeqNo = XML.GetAttribute("INSTRUCTIONSEQUENCENO");
				scScreenFunctions.SetContextParameter(window,"idInstructionSequenceNo",m_sInsSeqNo);
			}
		}
		<%//MAR1241 check we have a valuationtype and appointmentdate%>
		m_sValuationType = XML.GetAttribute("VALUATIONTYPE");
		m_sAppointmentDate = XML.GetAttribute("APPOINTMENTDATE");
	}
}

function frmScreen.btnPropertyDetails.onclick()
{
	frmToAP210.submit(); 	
}

function frmScreen.btnValuation.onclick()
{
	frmToAP200.submit(); 	
}

function btnConfirm.onclick()
{
	<% /* PSC 12/12/01 SYS3388 */ %>
	if(CheckValuerInstruction())  <%//MAR1241%>
	{
		if (confirm("Please ensure all data has been entered correctly before continuing"))
		{
			<% // GD BM0582 %>
			if(HasPresentValuationBeenEntered())
			{
				if (ConfirmChangeOfProperty()) <% /* MAR1300 GHun */ %>
				{
					if(SetValuationTask())
					{
						if (SaveValuerInstructions())
						{
							RouteScreen();
						}
					}
				}
			}
			
		}
	}
}

<%// GD BM0582 START %>
function HasPresentValuationBeenEntered()
{
	var ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ValuationXML.CreateRequestTag(window , "GetValuationReport");
	var bSuccess = false;
	ValuationXML.CreateActiveTag("VALUATION");
	ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
	ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
	ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO",m_sInsSeqNo);
	ValuationXML.RunASP(document, "omAppProc.asp");
	if (ValuationXML.IsResponseOK())
	{
		<%//Extract PresentValuation amount, if present%>
		ValuationXML.SelectTag(null,"GETVALUATIONREPORT");
		var sPresentValuation = ValuationXML.GetAttribute("PRESENTVALUATION");
		if (sPresentValuation != '') <%//There is some value for PRESENTVALUATION%>
		{
			bSuccess = true;
		} else
		{
			alert("Please enter a Present Valuation Amount before confirming.");		
		}
	} <%//else //An error occured%>


	return(bSuccess);

}

<%// GD BM0582 END %>
function SetValuationTask()
{
	var sInsSeqNo;
	var sAppFactFindNo;
	var sInstructionNo;
	var sUserRole;
	var sCustomerNo;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = XML.CreateRequestTag(window , "ValidateValuationReport");
	var bSuccess = false;
	var valTag;
	var sLTVChanged;
	var OldLTV, NewLTV;

	<%//MAR1434 if we are using valuationRulesBO we need to remember the current LTV,
	//calculate the new one based on presentvaluation, save the new LTV to the
	//mortgagesubquote and set the local flags before calling the rules.%>
	var gXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if(gXML.GetGlobalParameterBoolean(document, "ValuationUseAppProcRules") == 0)
	{
		var MsqXML = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
		MsqXML.CreateRequestTag(window, "GetAcceptedQuoteData");
		MsqXML.CreateActiveTag("APPLICATION");
		MsqXML.CreateTag("APPLICATIONNUMBER", m_sAppNo);
		MsqXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
		MsqXML.RunASP(document,"GetAcceptedOrActiveQuoteData.asp");
	
		if(MsqXML.IsResponseOK())
		{
			MsqXML.SelectTag(null, "MORTGAGESUBQUOTE");
			OldLTV = MsqXML.GetTagText("LTV");
			var sAmtReq = MsqXML.GetTagText("AMOUNTREQUESTED");
			
			var ltvXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			ltvXML.CreateRequestTag(window,null);
			ltvXML.CreateActiveTag("LTV");
			ltvXML.CreateTag("APPLICATIONNUMBER", m_sAppNo);
			ltvXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
			ltvXML.CreateTag("AMOUNTREQUESTED", sAmtReq);
			ltvXML.RunASP(document, "AQCalcCostModelLTV.asp");
			if(ltvXML.IsResponseOK())
			{
				ltvXML.SelectTag(null, "RESPONSE");
				NewLTV = ltvXML.GetTagText("LTV");
				if(parseFloat(NewLTV) != parseFloat(OldLTV))
				{
					sLTVChanged = true;
					<%//update mortgagesubquote with new ltv%>
					var updatemsqXML = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
					updatemsqXML.CreateRequestTag(window, "UpdateMortgageSubquote");
					updatemsqXML.CreateActiveTag("MORTGAGESUBQUOTE");
					updatemsqXML.AddXMLBlock(MsqXML.CreateFragment());
					updatemsqXML.SetTagText("LTV", NewLTV);
				
					updatemsqXML.RunASP(document, "AQUpdateMortgageSubQuote.asp");
					updatemsqXML.IsResponseOK();
				}
				else
					sLTVChanged = false;
			}
		}
	}
		
    XML.SetAttribute("USERID", scScreenFunctions.GetContextParameter(window,"idUserID",null));
	XML.SetAttribute("UNITID", scScreenFunctions.GetContextParameter(window,"idUnitID",null));		
	sUserRole = scScreenFunctions.GetContextParameter(window,"idRole",null);
	XML.SetAttribute("USERAUTHORITYLEVEL", sUserRole);
	valTag = XML.CreateActiveTag("VALUATION");		
	XML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
	XML.SetAttribute("TYPEOFMORTGAGE",scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue",null));
	XML.SetAttribute("INSTRUCTIONSEQUENCENO",m_sInsSeqNo);
	<%
	// build a customer tag for each non-empty customernumber#
	%>
	for (var idx=1; idx<=5; idx++)
	{
		sCustomerNo = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + idx, null);
	
		if (sCustomerNo != "")
		{
		    XML.ActiveTag = valTag;
			XML.CreateActiveTag("CUSTOMER");		
			XML.SetAttribute("CUSTOMERNUMBER", sCustomerNo);     
			XML.SetAttribute("CUSTOMERVERSIONNUMBER",scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + idx, null));
		}
    }
	XML.ActiveTag = reqTag;
	XML.CreateActiveTag("CASETASK");
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("ACTIVITYID", m_TaskXML.GetAttribute("ACTIVITYID"));
	XML.SetAttribute("STAGEID", m_TaskXML.GetAttribute("STAGEID"));
	XML.SetAttribute("CASESTAGESEQUENCENO", m_TaskXML.GetAttribute("CASESTAGESEQUENCENO"));
	XML.SetAttribute("CASEID", m_TaskXML.GetAttribute("CASEID"));
	XML.SetAttribute("TASKID", m_TaskXML.GetAttribute("TASKID"));
	XML.SetAttribute("TASKINSTANCE", m_TaskXML.GetAttribute("TASKINSTANCE"));
	XML.ActiveTag = reqTag;
	XML.CreateTag("OLDLTV", OldLTV);

	XML.RunASP(document, "OmigaTmBO.asp");

	if(XML.IsResponseOK())
	{
		bSuccess = true;
		
		<%//MAR1434 set LTV flags if we ran AppProcRules%>
		if(gXML.GetGlobalParameterBoolean(document, "ValuationUseAppProcRules") == 1)
		{
			XML.SelectTag(null, "LTV");
			<%//var sLTVChanged;%>
			sLTVChanged = XML.GetAttribute("LTVCHANGED");
			OldLTV = XML.GetAttribute("OLDLTV");
			NewLTV = XML.GetAttribute("NEWLTV");
		}
		
		<% /* PSC 12/12/01 SYS3388 */ %>
		if( sLTVChanged == "1")
		{
			<% /* BS BM0291 12/02/03 */ %>
			<%//var OldLTV, NewLTV;
			//OldLTV = XML.GetAttribute("OLDLTV");
			//NewLTV = XML.GetAttribute("NEWLTV");%>
			alert("The LTV has changed from " + OldLTV + " to " + NewLTV);
			<% /* BS BM0291 End 12/02/03 */ %>
			<%//JD MAR40 check global parameter before remodelling quote%>
			gXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			if(gXML.GetGlobalParameterBoolean(document, "ValuationLTVChgRemodelTask") == 1)
				RemodelQuote();
		}
		
		<%//BMIDS00473 Put application under review if we need to.%>
		gXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if(gXML.GetGlobalParameterBoolean(document, "ValuationUseAppProcRules") == 1) <%//JD MAR40%>
		{
			XML.SelectTag(null, "APPSTATUS");
			if(XML.GetAttribute("UNDERREVIEWIND") == "1")
			{
				alert("The Application has been placed under review");
				scScreenFunctions.SetContextParameter(window,"idAppUnderReview","1");
			}
		}
		else
		{
			<%//create the tasks returned from the valuation rules and update the report status%>
			ValuationRulesPostProcess(XML);
		}
		
<% /*		
//		XML.SelectTag(null, "APPSTATUS");
//		var sUnderReview;
//		sUnderReview = XML.GetAttribute("UnderReviewInd")
//		if( sUnderReview == "1")
//		{
//			frmScreen.btnReturnReassign.disabled = true;			
//		}
*/ %>	}
	return( bSuccess );
}

function ValuationRulesPostProcess(XML)
{
	<%//Update the status on the valuationreport%>
	var ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ValuationXML.CreateRequestTag(window, "UpdateValuationReport");
	ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	sASPFile = "omAppProc.asp";
	ValuationXML.CreateActiveTag("VALUATION");
	ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
	ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
	ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInsSeqNo);
	ValuationXML.SetAttribute("VALUATIONSTATUS", csValuerStatusComplete);
	switch (ScreenRules())
	{
		case 1: <%// Warning%>
		case 0: <%// OK%>
				ValuationXML.RunASP(document, sASPFile);
				break;
		default: <%// Error%>
				ValuationXML.SetErrorResponse();
	}
	ValuationXML.IsResponseOK();
	
	<%//Create any tasks that were returned.%>
	XML.SelectTag(null,"TASKS");
	XML.CreateTagList("TASK");
	for(var nTask = 0; nTask < XML.ActiveTagList.length; nTask++)
	{
		XML.SelectTagListItem(nTask);
		sTaskId = XML.GetAttribute("TASKID");

		CreateAdhocCaseTask(sTaskId, "40"); <% //taskStatus = 40 complete %>
		if(XML.GetAttribute("TASKNOTES") != "")
		{
			<% //Create a tasknote too %>
			CreateTaskNote(sTaskId, XML.GetAttribute("TASKNOTES"));
		}
	}
}

function SaveValuerInstructions()
{
	var bSuccess = false;
	var ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sTaskStatus;
	var sASPFile;
			
	if(m_sInsSeqNo.length > 0)
	{
		var reqTag = ValuationXML.CreateRequestTag(window, "UpdateValuerInstructions");

		ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");

		sASPFile = "omAppProc.asp";
		ValuationXML.CreateActiveTag("VALUERINSTRUCTION");
		ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sAppNo);
		ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
		ValuationXML.SetAttribute("INSTRUCTIONSEQUENCENO", m_sInsSeqNo);
		ValuationXML.SetAttribute("VALUATIONSTATUS", csValuerStatusComplete);
		bSuccess = true;
	}			

	if(bSuccess)
	{
		switch (ScreenRules())
			{
			case 1: <%// Warning %>
			case 0: <%// OK %>
					ValuationXML.RunASP(document, sASPFile);
				break;
			default: <%// Error %>
				ValuationXML.SetErrorResponse();
			}

		if(ValuationXML.IsResponseOK())
		{
			<%// Update the context %>
			bSuccess = true;
		}
		else
		{
			bSuccess = false;
		}
	}	

	return(bSuccess);
}

function CheckValuerInstruction()  <%// MAR1241 %>
{
	var bOK = true;
	
	if( m_sValuationType == "" || m_sAppointmentDate == "")
	{
		alert("Please add Valuation Type and Instruction Date by pressing valuer details button on the Valuation screen.");
		bOK = false;
	}
	
	return bOK;
}

function btnSubmit.onclick()
{
	if(CheckValuerInstruction()) <%// MAR1241 %>
		RouteScreen();
}

function RouteScreen()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
	if(m_sCallingScreen == "AP030")
	{
		frmToAP030.submit();
	}
	else if(m_sCallingScreen == "TM030")
	{
		frmToTM030.submit();
	}
	else
	{
		alert("Unknown calling screen ID: " + m_sCallingScreen);
	}
}

function GetMaxTaskInstance(sTaskId)
{
	var sTaskInst = "";
	
	var CaseTaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = CaseTaskXML.CreateRequestTag(window, "GetMaxCaseTaskInstance");
	
	CaseTaskXML.SetAttribute("USERID", scScreenFunctions.GetContextParameter(window,"idUserID",null));
	CaseTaskXML.SetAttribute("UNITID", scScreenFunctions.GetContextParameter(window,"idUnitID",null));		
	sUserRole = scScreenFunctions.GetContextParameter(window,"idRole",null);
	CaseTaskXML.SetAttribute("USERAUTHORITYLEVEL", sUserRole);				
	CaseTaskXML.CreateActiveTag("APPLICATION");		
	CaseTaskXML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
	CaseTaskXML.ActiveTag = reqTag;	
	CaseTaskXML.CreateActiveTag("CASETASK");
	CaseTaskXML.SetAttribute("CASEACTIVITYGUID", m_TaskXML.GetAttribute("CASEACTIVITYGUID"));
	CaseTaskXML.SetAttribute("CASESTAGESEQUENCENO", m_TaskXML.GetAttribute("CASESTAGESEQUENCENO"));
	CaseTaskXML.SetAttribute("STAGEID", m_TaskXML.GetAttribute("STAGEID"));
	CaseTaskXML.SetAttribute("TASKID", sTaskId);
    CaseTaskXML.RunASP(document, "OmigaTmBO.asp");
	if(CaseTaskXML.IsResponseOK())
	{
		sTaskInst = CaseTaskXML.GetTagText("TASKINSTANCE");
	}
	
	return sTaskInst;
}

function CreateTaskNote(sTaskId, sTaskNote)
{
	<%//First find the right instance of the task to add the note to %>
	var sTaskInstance = GetMaxTaskInstance(sTaskId);
	if(sTaskInstance != "")
	{
		var CaseTaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var reqTag = CaseTaskXML.CreateRequestTag(window, "CreateTaskNote");

		CaseTaskXML.SetAttribute("USERID", scScreenFunctions.GetContextParameter(window,"idUserID",null));
		CaseTaskXML.SetAttribute("UNITID", scScreenFunctions.GetContextParameter(window,"idUnitID",null));		
		sUserRole = scScreenFunctions.GetContextParameter(window,"idRole",null);
		CaseTaskXML.SetAttribute("USERAUTHORITYLEVEL", sUserRole);				
		CaseTaskXML.CreateActiveTag("APPLICATION");		
		CaseTaskXML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
		CaseTaskXML.ActiveTag = reqTag;	
		CaseTaskXML.CreateActiveTag("TASKNOTE");
		CaseTaskXML.SetAttribute("CASEACTIVITYGUID", m_TaskXML.GetAttribute("CASEACTIVITYGUID"));
		CaseTaskXML.SetAttribute("CASESTAGESEQUENCENO", m_TaskXML.GetAttribute("CASESTAGESEQUENCENO"));
		CaseTaskXML.SetAttribute("STAGEID", m_TaskXML.GetAttribute("STAGEID"));
		CaseTaskXML.SetAttribute("TASKID", sTaskId);
		CaseTaskXML.SetAttribute("TASKINSTANCE", sTaskInstance);
		CaseTaskXML.SetAttribute("NOTEENTRY", sTaskNote);
		CaseTaskXML.RunASP(document, "MsgTMBO.asp");
		CaseTaskXML.IsResponseOK();
	}
	else
		alert("Could not create task note " + sTaskNote + " for task " + sTaskId);
}

function CreateAdhocCaseTask(sTaskId, sTaskStatus)
{
	var bSuccess = false;
	var sUserRole;
	var CaseTaskXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var reqTag = CaseTaskXML.CreateRequestTag(window, "CreateAdhocCaseTask");

	CaseTaskXML.SetAttribute("USERID", scScreenFunctions.GetContextParameter(window,"idUserID",null));
	CaseTaskXML.SetAttribute("UNITID", scScreenFunctions.GetContextParameter(window,"idUnitID",null));		
	sUserRole = scScreenFunctions.GetContextParameter(window,"idRole",null);
	CaseTaskXML.SetAttribute("USERAUTHORITYLEVEL", sUserRole);				
	CaseTaskXML.CreateActiveTag("APPLICATION");		
	CaseTaskXML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
	CaseTaskXML.ActiveTag = reqTag;	
	CaseTaskXML.CreateActiveTag("CASETASK");
	CaseTaskXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	CaseTaskXML.SetAttribute("CASEID", m_TaskXML.GetAttribute("CASEID"));
	CaseTaskXML.SetAttribute("ACTIVITYID", m_TaskXML.GetAttribute("ACTIVITYID"));
	CaseTaskXML.SetAttribute("ACTIVITYINSTANCE", m_TaskXML.GetAttribute("ACTIVITYINSTANCE"));
	CaseTaskXML.SetAttribute("CASESTAGESEQUENCENO", m_TaskXML.GetAttribute("CASESTAGESEQUENCENO"));
	CaseTaskXML.SetAttribute("STAGEID", m_TaskXML.GetAttribute("STAGEID"));
	CaseTaskXML.SetAttribute("TASKID", sTaskId);
	if(sTaskStatus != "")
		CaseTaskXML.SetAttribute("TASKSTATUS", sTaskStatus);
    CaseTaskXML.RunASP(document, "OmigaTmBO.asp");
	if(CaseTaskXML.IsResponseOK())
		bSuccess = true;
	else
		bSuccess = false;

	return(bSuccess);
}

function RemodelQuote()
{
	var bSuccess = false;
	if(CreateAdhocCaseTask("Remodel_Quote", ""))
	{
		<%// Update the context %>
		bSuccess = true;
	}
	else
	{
		bSuccess = false;
	}	

	return(bSuccess);
}

<% /* MAR1300 GHun */ %>
function ConfirmChangeOfProperty()
{
	if (scScreenFunctions.GetContextParameter(window, "idPropertyChange", "") == "1")
	{
		if (confirm("A change of property is in-progress, by confirming this valuation you will end the change of property process."))
		{
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTag(window, null);
			XML.CreateActiveTag("NEWPROPERTY");
			XML.CreateTag("APPLICATIONNUMBER", m_sAppNo);
			XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sAppFactFindNo);
			XML.CreateTag("CHANGEOFPROPERTY", "0");
			
			XML.RunASP(document, "UpdateNewPropertyGeneral.asp");

			XML.IsResponseOK();
			scScreenFunctions.SetContextParameter(window, "idPropertyChange", "");
			
			return(true);
		}
		else
			return(false);
	}
	else
		return(true);
}
<% /* MAR1300 End */ %>
</script>
<script src="validation.js" type="text/javascript"></script>
</body>
</html>
