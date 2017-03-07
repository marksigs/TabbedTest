
<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /*
Workfile:      DC270.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Bank/Building Society Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		22/02/2000	Created
AD		13/03/2000	Fixed SYS0477
AD		15/03/2000	Incorporated third party include files.
IW		23/03/2000	Added SalaryFedIndicator.
AY		03/04/00	New top menu/scScreenFunctions change
MC		28/04/2000	Validate only one Repayment Account
MH		03/05/00	SYS0618 Postcode validation
SR		19/05/00    SYS0689 [5],[7],[8],[9],[10],[15],[16][17],[18],[19],[21]
BG		19/05/00	SYS0752 Removed Tooltips
MC		23/05/00	Add Bank Wizard functionality
SR		01/06/00	SYS0689 [2],[15]. AccountNumber should allow 20 characters. 
									  TimeatBank should not allow values greater than 100
BG		08/09/00	SYS1277  Validation on Account number
BG		12/09/00	SYS1277 only display validated message when appropriate. 
BG		09/10/00	SYS1559 Added code to handle new contact title combo in ThirdPartyDetails.asp
CL		05/03/01	SYS1920 Read only functionality added
CL		22/06/01	SYS2340 Adjustments for FirstPaymentDate
SR		06/09/01	SYS2412
MDC		01/10/01	SYS2785 Enable client versions to override labels
SG		22/11/01	SYS3050 Remove additional message box. 
SR		19/12/2001	SYS2547
GF		11/01/2002	SYS3733	Add comment notation to critical data check comment.
JLD		14/01/02	SYS3734 on retrieval of data pass contact details to DC241.
JLD		29/01/02	SYS3734 pass correct details to omTMBO
JLD		12/02/02	SYS4054 On adding to directory, force the user to input a branchname.
SA		24/02/02	SYS3938 Proposed Repayment Date not set correctly.
LD		23/05/02	SYS4727 Use cached versions of frame functions
SG		29/05/02	SYS4767 MSMS to Core integration
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History :

Prog	Date		AQR			Description
MV		08/08/2002	BMIDS00302	Core Ref : SYS4728 remove non-style sheet styles
ASu		04/09/2002	BMIDS00136	Make screen title shorter to prevent overlap.
ASu		10/09/2002	BMIDS00418	Fix IsInDirectory & RepaymentAccount radio buttons to display correctly
DRC     25/09/2002  BMIDS00183  Core Ref : AQR SYS4847 -call to bank wizard include file
MO		01/11/2002	BMIDS00725	Made the changes to save BankWizardIndicator
MO		18/11/2002	BMIDS00376	Sort out setting the screen to readonly
HMA     07/07/2003  BMIDS599    Disable Clear and PAF enabled buttons when screen is read only.
HMA     17/09/2003  BM0063      Amend HTML text for radio buttons

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History

Prog	Date		AQR		Description
MF		04/08/2005	MARS0	Routes back to DC280
MF		08/08/2005	MAR20	Parameterised routing depending on global parameters:
							"ThirdPartySummary" & "NewPropertySummary"
TL		06/09/2005	MAR39	Added "TRANSPOSEDINDICATOR"
HM      10/11/2005  MAR334  still have an echo from blocked ThirdParty code
HMA     11/11/2005  MAR507  Moved Direct Debit text box.
Maha T	01/12/2005	MAR333	If BankName entered manually, Don't display warning message on OK button,
							as we are not getting BankName in BankWizard XSDS file.
PE		20/02/2006	MAR1288	Don't let the user out of the screen if the details have not been authenticated by bank wizard.
PE		25/03/2006	MAR1278	Not all screens are frozen when the case is in the decline stage
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM 2 Specific History

Prog	Date		AQR		Description
HMA		20/09/2006	EP2_3	Add Roll Number. Save Bank Address details.
AShaw	27/10/2006	EP2_8	Disable All fields if Additional Borrowing.
PSC		31/01/2007	EP2_1114 Correct disabling of fields on Further Borrowing
INR		21/02/2007	EP2_1402 Multiple entries can be added by multiple clicks
PSC		23/03/2007	EP2_2087 Change length of roll number to 20 characters
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Automated Update History:

Prog    Date           Description
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<HEAD>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data="scClientFunctions.asp" height="1px" id="scClientFunctions" style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" type="text/x-scriptlet" width="1" VIEWASTEXT tabIndex="-1"></object>
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>

<% /* FORMS */ %>
<form id="frmToDC240" method="post" action="DC240.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC280" method="post" action="DC280.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM010" method="post" action="CM010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC295" method="post" action="DC295.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" validate  ="onchange" mark>
<div style="HEIGHT: 336px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		<strong>Bank/Building Society Details</strong> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Bank/Building Society Name
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtCompanyName" maxlength="45" style="POSITION: absolute; WIDTH: 200px" class="msgTxt" onchange="ContactDetailsChanged()">
		</span> 
	</span>

	<span style="LEFT: 470px; POSITION: absolute; TOP: 31px" class="msgLabel">
		<input id="btnDirectorySearch" value="Directory Search" type="button" style="WIDTH: 100px" class="msgButton" onclick="btnDirectorySearch.onclick()">
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Branch Name
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtBranchName" maxlength="50" style="POSITION: absolute; WIDTH: 300px" class="msgTxt" onchange="ContactDetailsChanged()">
		</span> 
	</span>

	<span style="LEFT: 470px; POSITION: absolute; TOP: 55px" class="msgLabel">
		<input id="btnBankWizard" value="Bank Wizard" type="button" style="WIDTH: 100px" class="msgButton" onclick="btnBankWizard.onclick()">
	</span>

	<% /* BM0063 - Put <input> and <label> on the same line for radio buttons to fix Read Only problem in IE5.5 */%>

	<% /* MAR20 Hide section */ %>
	<span style="display:none; LEFT: 4px; POSITION: absolute; TOP: 84px" class="msgLabel">
		Is in the directory?
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px; WIDTH: 60px">
			<input id="optInDirectoryYes" name="InDirectoryGroup" type="radio" value="1" onclick="InDirectoryChanged()"><label for="optInDirectoryYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px; WIDTH: 60px">
			<input id="optInDirectoryNo" name="InDirectoryGroup" type="radio" value="0" onclick="InDirectoryChanged()"><label for="optInDirectoryNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<% /* MAR20 Hide section */ %>
	<span style="display:none; LEFT: 4px; POSITION: absolute; TOP: 108px" class="msgLabel">
		Account Salary Paid Into?
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px; WIDTH: 60px">
			<input id="optSalaryFedIndicatorYes" name="SalaryFedIndicatorGroup" type="radio" value="1"><label for="optSalaryFedIndicatorYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px; WIDTH: 60px">
			<input id="optSalaryFedIndicatorNo" name="SalaryFedIndicatorGroup" type="radio" value="0"><label for="optSalaryFedIndicatorNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<% /* MAR39 Hide section */ %>
	<span style="display:none; LEFT: 4px; POSITION: absolute; TOP: 108px" class="msgLabel">
		Transposed Indicator
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px; WIDTH: 60px">
			<input id="optTransposedIndicatorYes" name="TransposedIndicatorGroup" type="radio" value="1"><label for="optTransposedIndicatorYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px; WIDTH: 60px">
			<input id="optTransposedIndicatorNo" name="TransposedIndicatorGroup" type="radio" value="0"><label for="optTransposedIndicatorNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 84px" class="msgLabel">
		Sort Code
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtSortCode" maxlength="8" style="POSITION: absolute; WIDTH: 50px" class="msgTxt" onkeyup="BankWizardDetailsChanged()">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 108px" class="msgLabel">
		Account Number
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountNumber" maxlength="8" style="POSITION: absolute; WIDTH: 130px" class="msgTxt" onkeyup="BankWizardDetailsChanged()">
		</span> 
	</span>

	<span style="LEFT: 300px; POSITION: absolute; TOP: 108px" class="msgLabel">
		Roll Number
		<span style="LEFT: 64px; POSITION: absolute; TOP: -3px">
			<input id="txtRollNumber" maxlength="20" style="POSITION: absolute; WIDTH: 130px" class="msgTxt" onkeyup="BankWizardDetailsChanged()">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 132px" class="msgLabel">
		Account Name
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtAccountName" maxlength="50" style="POSITION: absolute; WIDTH: 330px" class="msgTxt">
		</span> 
	</span>
	
	<% /* MAR20 Hide section */ %>
	<span style="display:none; LEFT: 4px; POSITION: absolute; TOP: 204px" class="msgLabel">
		Time at Bank/Building Society
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtTimeAtBank" maxlength="3" style="POSITION: absolute; WIDTH: 50px" class="msgTxt">
			<span style="LEFT: 55px; POSITION: absolute; TOP: 3px" class="msgLabel">
				(Years)
			</span>
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 156px" class="msgLabel">
		Repayment Account?
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="optRepaymentBankAccountYes" name="RepaymentBankAccountGroup" type="radio" value="1" onclick="RepaymentBankAccountChanged()"><label for="optRepaymentBankAccountYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px">
			<input id="optRepaymentBankAccountNo" name="RepaymentBankAccountGroup" type="radio" value="0" onclick="RepaymentBankAccountChanged()"><label for="optRepaymentBankAccountNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 180px" class="msgLabel">
		Proposed Method of Repayment
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<select id="cboProposedMethodOfRepayment" style="WIDTH: 150px" class="msgCombo"></select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 204px" class="msgLabel">
		DD Explained
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="optDDExplainedYes" name="DDExplainedGroup" type="radio" value="1"><label for="optDDExplainedYes" class="msgLabel">Yes</label> 
		</span> 
		<span style="LEFT: 214px; POSITION: absolute; TOP: -3px">
			<input id="optDDExplainedNo" name="DDExplainedGroup" type="radio" value="0"><label for="optDDExplainedNo" class="msgLabel">No</label> 
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 228px" class="msgLabel">
		<span style="LEFT: 2px; POSITION: absolute; TOP: -3px">
			<TEXTAREA class="msgTxt" id=txtDDExplanation readonly rows=13 style="POSITION: absolute; WIDTH: 565px" NAME="txtDDExplanation"></TEXTAREA>
		</span>
	</span>	

	<span style="LEFT: 4px; POSITION: absolute; TOP: 426px" class="msgLabel">
		DD Reference
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtDDReference" maxlength="20" style="POSITION: absolute; WIDTH: 70px" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 450px" class="msgLabel">
		Proposed Repayment Day
		<span style="LEFT: 160px; POSITION: absolute; TOP: -3px">
			<input id="txtProposedRepaymentDate" maxlength="2" style="POSITION: absolute; WIDTH: 30px" class="msgTxt">
		</span> 
	</span>
</div>

<% /* MAR20 hide section */ %>
<div style="display:none; HEIGHT: 232px; LEFT: 10px; POSITION: absolute; TOP: 394px; WIDTH: 604px" class="msgGroup">
<% /* MAR334 Following the changes include ThirdPartyDetails should be blocked */ %>
<% /* <!-- #include FILE="includes/thirdpartydetails.htm" -->*/ %></div> 
</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 536px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- #include FILE="attribs/DC270attribs.asp" -->

<% /* MAR334 Following the changes include ThirdPartyDetails should be blocked */ %>
<% /* <!-- #include FILE="attribs/ThirdPartyDetailsAttribs.asp" -->*/ %>

<% /* SG 29/05/02 SYS4767 */ %>
<!-- #include FILE="includes/BankWizard.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sXML = "";
var m_strDayOfFirstPaymentDate = "";
var m_PaidFound = "";
var m_strFirstPaymentDate = "";
var m_RepaymentAccountIndicator = "";

var scScreenFunctions;
	
var BankXML = null;
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber  = "";
var m_sBankAccountSequenceNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sPreferedPaymentDay = "";
var m_txtProposedRepaymentDate_First = "";
var m_txtProposedRepaymentDate_Last = "";
var XMLTempNode = "";
var LCPaymentListXML = null ;

var m_bBankAccountValid = false;
var TPBankXML = null;
var m_blnReadOnly = false;
var dateObj; 
var strPaymentsFound = false;
var strFirstPaymentDate = "";
var m_sCancelledStatusId = "" ;
var m_sInitialFirstPaymentDate = "";
var m_bRepaymentIndEditable = true ;

var m_sThirdPartyGUID = "";			//MAR334

//SG 29/05/02 SYS4767
var m_sScreenId = "DC270";

<% /* EP2_8 - New variable */ %>
var m_blnAddtnBorrowing = false;

<% /* EVENTS */ %>
function FindLoanComponentPaymentList()
{
	if(m_RepaymentAccountIndicator = "1")
	<% //if (frmScreen.optRepaymentBankAccountYes.checked)
	%>
	{
		LCPaymentListXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		LCPaymentListXML.CreateRequestTag(window, "FindLoanComponentPaymentList");
		LCPaymentListXML.CreateActiveTag("LOANCOMPONENTPAYMENT");
		m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
		LCPaymentListXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
		LCPaymentListXML.RunASP(document,"PaymentProcessingRequest.asp");
	
		<% // if(!LCPaymentListXML.CheckResponse())
		%>
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = LCPaymentListXML.CheckResponse(ErrorTypes);
		
		if (ErrorReturn[1] != ErrorTypes[0])
		{
			strPaymentsFound = true;
			<% //else if(ErrorReturn[0] == true)
			%>
			{
				var xslPattern = "RESPONSE/LOANCOMPONENTPAYMENT[$not$ PAYMENTSTATUS='U']";  // +S
				<% //Check returned XML for any occurences of xslPattern
				%>
				m_PaidFound =	LCPaymentListXML.SelectSingleNode(xslPattern);
				XMLTempNode = LCPaymentListXML.SelectSingleNode("//LOANCOMPONENTPAYMENT");			
								
				<% //Get the DAY OF FIRSTPAYMENTDATE
				%>
				m_strFirstPaymentDate = LCPaymentListXML.GetAttribute("FIRSTPAYMENTDATE");
			
				m_strDayOfFirstPaymentDate = LCPaymentListXML.GetAttribute("FIRSTPAYMENTDAY");

				if (m_PaidFound != "")
				{
					<% //Set the txtProposedRepaymentDate to the "FIRSTREPAYMENTDATE"
					%>
					//frmScreen.txtProposedRepaymentDate.value = m_strDayOfFirstPaymentDate		//Disable fields
					scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtProposedRepaymentDate");
				}	
				else 
				{
					<% //Set the txtProposedRepaymentDate to the "FIRSTREPAYMENTDATE"
					%>
					//frmScreen.txtProposedRepaymentDate.value = m_strDayOfFirstPaymentDate
					//Enable fields
					scScreenFunctions.SetFieldToWritable(frmScreen,"txtProposedRepaymentDate");
				}	

			}
		}
	}
	else
	{
		m_strDayOfFirstPaymentDate = "";
	}	
}

<% /* SR 28/08/01 : SYS2412  */ %>
function IsProposedDateEditable()
{
	<% /* Check whether all the loan component payments have status less than PAID */ %>
	var xmlNode, xmlROFNodeList, xmlROFNode ;
	var lngActualPayment, lngTotalROFPayment ; 
	var sPaymentSeqNo ;

	var blnDateEditable = false ; 
	var xmlNodeList = LCPaymentListXML.XMLDocument.selectNodes("//LOANCOMPONENTPAYMENT");
	var lngTotalNoOfLCPayments = xmlNodeList.length ;
	
	var sCondition = "//LOANCOMPONENTPAYMENT[@PAYMENTSTATUS $eq$ 'U' $or$ @PAYMENTSTATUS $eq$ 'S' " +
					 " $or$ @PAYMENTSTATUS $eq$ 'C' $or$ @PAYMENTSTATUS $eq$ 'E']" ;
	
	xmlNodeList = LCPaymentListXML.XMLDocument.selectNodes(sCondition);
	var lngUnPaidLCPayments = xmlNodeList.length ;
	
	blnDateEditable = false ;
	
	if(lngTotalNoOfLCPayments == lngUnPaidLCPayments)
	{
		blnDateEditable = true ;
		return true ;
	}
	
	if(!blnDateEditable)
	{
		<% /* if any of the records have status more than PAID (Awaiting interface response, interfaced,
           interfaced not paid), do not allow to edit the payment date. */
        %>        
        sCondition = "//LOANCOMPONENTPAYMENT[@PAYMENTSTATUS $eq$ 'R' $or$ @PAYMENTSTATUS $eq$ 'I' $or$ " +
                                         " @PAYMENTSTATUS $eq$ 'INP']"
		xmlNodeList = LCPaymentListXML.XMLDocument.selectNodes(sCondition)
		
		if(xmlNodeList.length > 0) return false ;
		else
		{
			<%	/* For each PAID record, find the corresponding total ROF amount. If ROF amount is less
			       than the actual payment amount, the date cannot be changed. */
			%>      
			sCondition = "//LOANCOMPONENTPAYMENT[@PAYMENTSTATUS $eq$ 'P' ]" ;
			xmlNodeList = LCPaymentListXML.XMLDocument.selectNodes(sCondition) ;
			
			blnDateEditable = true ;
			
			for(var iCount = 0 ; iCount < xmlNodeList.length ; iCount++)
			{
				xmlNode = xmlNodeList.item(iCount) ;				
				lngActualPayment = 0 ;
				lngTotalROFPayment = 0 ;
				sPaymentSeqNo = LCPaymentListXML.GetNodeAttribute(xmlNode, "PAYMENTSEQUENCENUMBER") ;
				
				sCondition = "//LOANCOMPONENTPAYMENT[ASSOCPAYSEQNUMBER $eq$ " + sPaymentSeqNo + "]" ;
				xmlROFNodeList = LCPaymentListXML.XMLDocument.selectNodes(sCondition) ;
				
				if(xmlROFNodeList.length > 0)
				{ <% /* if the total ROF amount is less than the acual payment amount, payment date cannot be edited */ %>
					for(var iRofCount = 0 ; iRofCount < xmlROFNodeList.length ; iRofCount++)
					{
						xmlROFNode = xmlROFNodeList.item(iRofCount) ;
						sROFAmount = LCPaymentListXML.GetNodeAttribute(xmlROFNode, "AMOUNT") ;
						if(! isNaN(sROFAmount)) lngTotalROFPayment = lngTotalROFPayment + parseFloat(sROFAmount);
						
						if(lngActualPayment > lngTotalROFPayment) return false ;
					}
				}
				else return false;
			}
		}	
	}	
	return blnDateEditable ;
}

function BankWizardDetailsChanged()
{
	frmScreen.btnBankWizard.disabled = ((frmScreen.txtAccountNumber.value == "") |
			(frmScreen.txtSortCode.value == ""));
	m_bBankAccountValid = false;
}

function btnCancel.onclick()
{
	scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
	scScreenFunctions.SetContextParameter(window,"idXML", "");	
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
	var bThirdPartySummary = XML.GetGlobalParameterBoolean(document,"ThirdPartySummary");			
	if(bThirdPartySummary)
		frmToDC240.submit();
	else
		frmToDC280.submit();
}

<%/* EP2_1402 reworked */%>
function btnSubmit.onclick()
{
	if( (m_sReadOnly != "1") && IsChanged())
	{
		<% /* EP2_1402 Disable Submit and display hourglass until processing complete */ %>
		btnSubmit.disabled = true;
		btnSubmit.blur();
		btnSubmit.style.cursor = "wait";
		<% /* Call CommitChanges after timeout, so cursor can change */ %>
		if (frmScreen.onsubmit()&& ValidateScreen())
		{
			window.setTimeout("CommitChanges()",0);
		}
		else
		{
			btnSubmit.style.cursor = "hand";
			btnSubmit.disabled = false;
		}
	}
	else
	{
		//Nothing changed, sort out where to route to
		doRouting(true)
	}
}

function doRouting(changesOK)
{
	if (changesOK)
	{
		//clear the contexts
		<% /* CheckDates(); */ %>
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
		scScreenFunctions.SetContextParameter(window,"idXML", "");
		<% /* MF MAR20 route depending on Global parameters */ %>
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
		var bThirdPartySummary = XML.GetGlobalParameterBoolean(document,"ThirdPartySummary");			
		if(bThirdPartySummary)
			frmToDC240.submit();
		else{
			var bNewPropertySummary = XML.GetGlobalParameterBoolean(document,"NewPropertySummary");			
			if(bNewPropertySummary){	
				frmToCM010.submit();			
			}else{				
				frmToDC295.submit();
			}
		}
	}
	btnSubmit.style.cursor = "hand";
	btnSubmit.disabled = false;
}

function frmScreen.btnBankWizard.onclick()
{
    with (frmScreen)
		m_bBankAccountValid = BankWizard(txtSortCode,txtAccountNumber,txtRollNumber);  <% /*EP2_3*/ %>
	
}

<% /* SR 19/12/2001 : SYS2547 - do not have to call this method. All taken care in the manage method
function CheckDates()
{
	m_txtProposedRepaymentDate_Last = frmScreen.txtProposedRepaymentDate.value;
	
	if(m_bBankAccountValid ==  false)
	{
		// SG 22/11/01 SYS3050
		//if(confirm("The bank details have not been validated using Bank Wizard. Continue anyway ?"))
		//{
			//if metaaction = "EDIT"
			m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction",null);
			if(m_sMetaAction == "Edit")
			{
				//If any data has been ammended 
				if(frmScreen.onsubmit())
				{
					if (IsChanged())
					{
						if(m_txtProposedRepaymentDate_First != m_txtProposedRepaymentDate_Last)
						{
							//If ProposedPaymentDate is null
							if(m_txtProposedRepaymentDate_Last == "")
							{
								//Make XML and run UpdateFirstPaymentDate with GLOBAL PaymentDay
								UpdatePaymentDateXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
								UpdatePaymentDateXML.CreateRequestTag(window, "UpdateFirstPaymentDate");
								UpdatePaymentDateXML.CreateActiveTag("LOANCOMPONENTPAYMENT");
								m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
								//m_sPreferedPaymentDay = scScreenFunctions.GetContextParameter(window,"m_sPreferedPaymentDay",null);
								UpdatePaymentDateXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
								//UpdatePaymentDateXML.SetAttribute("PREFEREDPAYMENTDAY",m_sPreferedPaymentDay);
								UpdatePaymentDateXML.RunASP(document,"PaymentProcessingRequest.asp");
							}
							else
							{
								//Make XML and run UpdateFirstPaymentDate with SCREEN PaymentDay
								UpdatePaymentDateXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
								UpdatePaymentDateXML.CreateRequestTag(window, "UpdateFirstPaymentDate");
								UpdatePaymentDateXML.CreateActiveTag("LOANCOMPONENTPAYMENT");
								m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
								m_sPreferedPaymentDay = frmScreen.txtProposedRepaymentDate.value;
								UpdatePaymentDateXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
								UpdatePaymentDateXML.SetAttribute("PREFEREDPAYMENTDAY",m_sPreferedPaymentDay);
								UpdatePaymentDateXML.RunASP(document,"PaymentProcessingRequest.asp");
							}											
							
							//Call ThirdParty.UpdateBankBuildingSociety 
																		
							var BankBuildingSocietyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

							BankBuildingSocietyXML.CreateRequestTag(window,null);
							BankBuildingSocietyXML.CreateActiveTag("APPLICATIONBANKBUILDINGSOCIETY");
							BankBuildingSocietyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
							BankBuildingSocietyXML.CreateTag("PREFEREDPAYMENTDAY", m_sPreferedPaymentDay);
							BankBuildingSocietyXML.RunASP(document,"UpdateBankBuildingSociety.asp");										
						}
					}
				else
					{		
					//Route to DC240
					frmToDC240.submit();
					}			
				}
			}
			else if(m_sMetaAction == "Add")// MetaAction is ADD
			{
				//If ProposedPaymentDate <> Null
				if(m_txtProposedRepaymentDate_Last != "")
				{
					if(m_txtProposedRepaymentDate_First != m_txtProposedRepaymentDate_Last)
			 		{
			 			//Make XML and run UpdateFirstPaymentDate with SCREEN PaymentDay
						UpdatePaymentDateXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
						UpdatePaymentDateXML.CreateRequestTag(window, "UpdateFirstPaymentDate");
						UpdatePaymentDateXML.CreateActiveTag("LOANCOMPONENTPAYMENT");
						m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null);
						m_sPreferedPaymentDay = frmScreen.txtProposedRepaymentDate;
						UpdatePaymentDateXML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
						m_sPreferedPaymentDay = frmScreen.txtProposedRepaymentDate.value;
						UpdatePaymentDateXML.SetAttribute("PREFEREDPAYMENTDAY",m_sPreferedPaymentDay);
						UpdatePaymentDateXML.RunASP(document,"PaymentProcessingRequest.asp");
					}		
			 	//Call ThirdParty.CreateBankBuildingSociety 
			 	
			 	var BankBuildingSocietyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			 	
			 	BankBuildingSocietyXML.CreateRequestTag(window,null);
				BankBuildingSocietyXML.CreateActiveTag("APPLICATIONBANKBUILDINGSOCIETY");
				BankBuildingSocietyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
				BankBuildingSocietyXML.CreateTag("PREFEREDPAYMENTDAY", m_sPreferedPaymentDay);
				BankBuildingSocietyXML.RunASP(document,"CreateBankBuildingSociety.asp");		
			 				 	
			 	}		
			 	//Route to DC240
			 	frmToDC240.submit();		
			}
		//}
		//else
		//{
		//	//Remain on DC270
		//}
	}	
}			
*/
%>

function frmScreen.cboProposedMethodOfRepayment.onchange()
{
	var sRepaymentType = scScreenFunctions.GetComboValidationType(frmScreen, "cboProposedMethodOfRepayment");
	if (sRepaymentType == "D")
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtDDReference");
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtDDReference");
		frmScreen.txtDDReference.value = "";
	}
}
<% /* SR 18/05/00 SYS0689 : In Edit Mode, if the values of comany name or sort code are changed 
	set default values to bank account related fields  */ 
%>	
function frmScreen.txtSortCode.onchange()
{
	if (m_sMetaAction == "Edit") DefaultBankAccountFields() ; 
}

function frmScreen.txtCompanyName.onchange()
{
	if (m_sMetaAction == "Edit") DefaultBankAccountFields() ;
}

<% /* SR 01/06/00-SYS0689 : Value entered in the field 'Time at Bank/Building Society cannot be greater than 100 */ %>
function frmScreen.txtTimeAtBank.onblur()
{
	if (this.value > 100)
	{
		alert ('Time at Bank/Building Society cannot be greater than 100');
		this.focus() ;
	}	
}

function RepaymentBankAccountChanged()  //OnChangeRepaymentIndicator
{
	var UpdatePaymentXML = null ;
	
 	<% /* Repayment indicator changed from NO to YES  */ %>
 	if (frmScreen.optRepaymentBankAccountYes.checked) 
	{
		if(LCPaymentListXML.XMLDocument.selectNodes("//LOANCOMPONENTPAYMENT").length > 0)
		{
			//SYS3938 Spec says this should be set to day of first payment date, whether it's readonly or not.
			frmScreen.txtProposedRepaymentDate.value =  m_strDayOfFirstPaymentDate ;
			<%/*ASu - BMIDS00418 - Start  
			if(!IsProposedDateEditable()) scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtProposedRepaymentDate");
			else 
			<%/*ASu - BMIDS00418 - End */ %>
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtProposedRepaymentDate");
		}
	}
	else // Repayment indicator changed from YES to NO
	{
		frmScreen.txtProposedRepaymentDate.value = "";  
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtProposedRepaymentDate"); 
	}
		
	if (scScreenFunctions.GetRadioGroupValue(frmScreen,"RepaymentBankAccountGroup") == "1")
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboProposedMethodOfRepayment");
		if (scScreenFunctions.GetComboValidationType(frmScreen, "cboProposedMethodOfRepayment") == "D")
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtDDReference");
		else
		{
			scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtDDReference");
			frmScreen.txtDDReference.value = "";
		}
		<%//scScreenFunctions.SetFieldToWritable(frmScreen,"txtProposedRepaymentDate");
		%>
		<%/*ASu - BMIDS00418 - Start */ %>
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtProposedRepaymentDate");
		<%/*ASu - BMIDS00418 - Start */ %>
	}
	else
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboProposedMethodOfRepayment");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtDDReference");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtProposedRepaymentDate");
		frmScreen.cboProposedMethodOfRepayment.selectedIndex = 0;
		frmScreen.txtDDReference.value = "";
		frmScreen.txtProposedRepaymentDate.value = "";
	}
}

function spnToFirstField.onfocus()
{
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function spnToLastField.onfocus()
{
	scScreenFunctions.SetFocusToLastField(frmScreen);
}


<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
var scClientScreenFunctions;
function window.onload()
{
	<% /* MF 05/08/2005 MARS20 hide directory search button */ %>
	document.all.item("btnDirectorySearch").style.visibility = "hidden";
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	<%/* Parameters for thirdpartydetails.asp */%>
	m_sThirdPartyType = "3";
	m_ctrBranchName = frmScreen.txtBranchName;
	m_ctrSortCode = frmScreen.txtSortCode;
	m_fValidateScreen = ValidateScreen;	

	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	
	<% /* ASu BMIDS00136 */ %>	
	FW030SetTitles("Add/Edit Bank Details","DC270",scScreenFunctions);

	//MAR334 Block ThirdParty
	// MDC SYS2564 / SYS2785 Client Customisation
	//var sGroups = new Array("Country", "ProposedRepayMethod");
	//objDerivedOperations = new DerivedScreen(sGroups);

	RetrieveContextData();
	//MAR334 Block ThirdParty
	//SetThirdPartyDetailsMasks();
	SetMasks();

	//MAR334 Block ThirdParty
	// MC SYS2564/SYS2785 for client customisation
	//ThirdPartyCustomise();

	Validation_Init();	
	Initialise(true);

	// Peter Edney - 25/03/2006
	// MAR1278 - Not all screens are frozen when the case is in the decline stage
	// scScreenFunctions.SetFocusToFirstField(frmScreen);
	
	m_sInitialFirstPaymentDate = frmScreen.txtProposedRepaymentDate.value ;
	
	<% /* SR 19/05/00 - SYS0689 : In read-only mode,disable all fields on the screen except OK and Cancel button */ %>
	<%/*ASu - BMIDS00418 - Start */ %>
	//SetAvailableFunctionality(); MAR334
	RepaymentBankAccountChanged();
	<%/*ASu - BMIDS00418 - End */ %>

	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC270");

	if (m_blnReadOnly == true) 
	{
		m_sReadOnly = "1";
		<% /* MO - 18/11/2002 - BMIDS00376 */ %>
		if (frmScreen.btnDirectorySearch.style.visibility != "hidden") frmScreen.btnDirectorySearch.disabled = true;
		if (frmScreen.btnBankWizard.style.visibility != "hidden") frmScreen.btnBankWizard.disabled = true;
		<% /* MAR334 code below blocked */ %>
		<% /* HMA - 07/07/2003 - BMIDS599 */ %>
		//if (frmScreen.btnClear.style.visibility != "hidden") frmScreen.btnClear.disabled = true;
		//if (frmScreen.btnPAFSearch.style.visibility != "hidden") frmScreen.btnPAFSearch.disabled = true;
		//if (frmScreen.btnAddToDirectory.style.visibility != "hidden") frmScreen.btnAddToDirectory.disabled = true;
	}
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();

	// Peter Edney - 25/03/2006
	// MAR1278 - Not all screens are frozen when the case is in the decline stage
	if(!m_blnReadOnly && !m_blnAddtnBorrowing)
	{
		scScreenFunctions.SetFocusToFirstField(frmScreen);
	}
	
	<% /* Add extra condition if Additional Borrowing*/ %>
	<% /* PSC 31/01/2007 EP2_1114 - Start */ %>
	if ( m_blnAddtnBorrowing == true && m_RepaymentAccountIndicator == "1")
	{ 
		SetScreenToReadOnly();
		frmScreen.txtCompanyName.removeAttribute("required");
		frmScreen.txtBranchName.removeAttribute("required");
		frmScreen.txtSortCode.removeAttribute("required");
		frmScreen.txtAccountNumber.removeAttribute("required");
		frmScreen.txtAccountName.removeAttribute("required");
		m_bBankAccountValid = true;
	}
	<% /* PSC 31/01/2007 EP2_1114 - End */ %>
}

<% /* FUNCTIONS */ %>

function SetScreenToReadOnly()
{
<%	// Set all controls on the screen to read only
	// The buttons are treated separately
%>	
	scScreenFunctions.SetScreenToReadOnly(frmScreen);

	frmScreen.btnBankWizard.disabled = true;
	frmScreen.btnDirectorySearch.disabled = true;
	<% /* MAR334 code below blocked */ %>
	//frmScreen.btnClear.disabled = true;
	//frmScreen.btnPAFSearch.disabled = true;
	//frmScreen.btnAddToDirectory.disabled = true;
}
function ValidateScreen()
{
	<% /* PSC 31/01/2007 EP2_1114 - Start */ %>
	if (m_blnAddtnBorrowing == true && m_RepaymentAccountIndicator == "1")
	{
		return (true);
	}
	<% /* PSC 31/01/2007 EP2_1114 - End */ %>

	if(frmScreen.txtBranchName.value.length == 0)
	{
		alert("Please enter the name of the branch");
		frmScreen.txtBranchName.focus();
		return(false);
	}
	if (frmScreen.txtAccountNumber.value.length < 8)
	{
		alert("The Account Number must be 8 digits in length, including leading zeroes.");
		frmScreen.txtAccountNumber.focus();
		return(false);
	}

	if (((frmScreen.txtProposedRepaymentDate.value == 0) && (frmScreen.txtProposedRepaymentDate.value != "")) ||
	    (frmScreen.txtProposedRepaymentDate.value > 28))
	{
		alert("The Proposed Repayment Date must be in the range 1-28.");
		frmScreen.txtProposedRepaymentDate.focus();
		return(false);
	}
	<%	/* SR 17/05/00-SYS0689 : If RepaymentBankAccountInd is TRUE, then Proposed method of 	
								 repayment should not be empty. Address should not be empty */
	%>		
	if (frmScreen.optRepaymentBankAccountYes.checked)
		if (frmScreen.cboProposedMethodOfRepayment.value == "")
		{
			alert("The proposed method of repayment should not be left empty for a Repayment Account.") ;
			return(false);
		}
	<%	/*EP2_1402 moved this so all validation in one place */%>
	if (!m_bBankAccountValid)
	{
		alert("The bank account details have not been validated using the Bank Wizard.\n\n" +
			"Bank account details must be validated before continuing." );
		return(false);							
	}	
	<% /* MAR20 Section is hidden don't validate
	if (IsAddressEmpty())
	{
		alert ("Address should not be empty") ;
		frmScreen.txtPostcode.focus();
		return (false) ;
	}
	*/ %>	
			
	return(true);
}

<%/* EP2_1402 reworked */%>
function CommitChanges()
{
	var bSuccess = true;
	
	bSuccess = UpdatePaymentDateAndSaveBankBuildingSociety();

	doRouting(bSuccess);
}

function DefaultFields()
// Inserts default values into all fields
{
	//ClearFields(true, true, true); MAR334

	with (frmScreen)
	{
		txtSortCode.value = "";
		txtAccountNumber.value = "";
		txtRollNumber.value = "";      <%/* EP2_3 */%>
		txtAccountName.value = "";
		txtTimeAtBank.value = "";
		scScreenFunctions.SetRadioGroupValue(frmScreen,"RepaymentBankAccountGroup","1");
		scScreenFunctions.SetRadioGroupValue(frmScreen,"SalaryFedIndicatorGroup","0");				
		scScreenFunctions.SetComboOnValidationType(frmScreen,"cboProposedMethodOfRepayment","D");				
		txtDDReference.value = "";
		txtProposedRepaymentDate.value = "01";
	}

	m_bBankAccountValid = false;
	BankWizardDetailsChanged();
	RepaymentBankAccountChanged();
	m_sDirectoryGUID = "";
	m_bDirectoryAddress = false;
}

<% /* Set default values to fields related to Bank Account  */ %>
function DefaultBankAccountFields()
{
	with (frmScreen)
	{
		txtAccountNumber.value = "";
		txtRollNumber.value = "";    <%/* EP2_3 */%>
		txtAccountName.value = "";
		txtTimeAtBank.value = "";
		scScreenFunctions.SetRadioGroupValue(frmScreen,"RepaymentBankAccountGroup","1");
		scScreenFunctions.SetRadioGroupValue(frmScreen,"SalaryFedIndicatorGroup","0");		
		scScreenFunctions.SetComboOnValidationType(frmScreen,"cboProposedMethodOfRepayment","D");				
		txtDDReference.value = "";
		txtProposedRepaymentDate.value = "01";
	}
	m_bBankAccountValid = false;
	RepaymentBankAccountChanged();
	BankWizardDetailsChanged()
}

function ValidateRepaymentAccount()
{
var BankList = null;
var BankAccount = null;
var sSequenceNumber = "";
var ErrorTypes = new Array("RECORDNOTFOUND");
var ErrorReturn;

	<% /* Check for another account. If one exists and is a repayment account, disable
		  the repayment account option buttons	*/ %>
	if (m_sBankAccountSequenceNumber == "")
		sSequenceNumber = "0";
	else
		sSequenceNumber = m_sBankAccountSequenceNumber;
		
	TPBankXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	TPBankXML.CreateRequestTag(window,null);

	var tagApplication = TPBankXML.CreateActiveTag("APPLICATIONBANKBUILDINGSOC");
	TPBankXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	TPBankXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	
	TPBankXML.RunASP(document,"FindBankBuildingSocietyList.asp");
	ErrorReturn = TPBankXML.CheckResponse(ErrorTypes);
	if (ErrorReturn[0] == true) 
	{
		BankList = TPBankXML.CreateTagList("APPLICATIONBANKBUILDINGSOC")
		if(BankList.length > 0)
		{
			for(var nLoop = 0; nLoop < BankList.length; nLoop++)
			{
				BankAccount = TPBankXML.SelectTagListItem(nLoop);
				if((TPBankXML.GetTagText("BANKACCOUNTSEQUENCENUMBER") != sSequenceNumber) && (TPBankXML.GetTagText("REPAYMENTBANKACCOUNTINDICATOR") != "0"))
				{
				 	scScreenFunctions.SetRadioGroupValue(frmScreen,"RepaymentBankAccountGroup", "0");
				 	scScreenFunctions.SetRadioGroupState(frmScreen,"RepaymentBankAccountGroup", "R");
				 	scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtDDReference");
				 	
				 	m_bRepaymentIndEditable = false ;
				}
			}
		}
	}
}

function Initialise()
<% /* Initialises the screen */ %>
{
	PopulateCombos();
	
	FindLoanComponentPaymentList();

	if(m_sMetaAction == "Edit")
	{	
		PopulateScreen();
		
		<%/*AShaw - EP2_8 - Disable if Additional Borrowing */ %>
		var m_sidMortgageApplicationValue = scScreenFunctions.GetContextParameter(window,"MortgageApplicationValue","");
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		if(XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("F")) && XML.IsInComboValidationList(document,"TypeOfMortgage", m_sidMortgageApplicationValue, Array("M")))
			m_blnAddtnBorrowing = true;

	}
	else
	{
		DefaultFields();
		//SetAvailableFunctionality(); MAR334
		<% //ValidateRepaymentAccount();%>
		
		// Peter Edney - 25/03/2006
		// MAR1278 - Not all screens are frozen when the case is in the decline stage
		//scScreenFunctions.SetFocusToFirstField(frmScreen);
	}
	
	ValidateRepaymentAccount();
	
	if(LCPaymentListXML.XMLDocument.selectNodes("//LOANCOMPONENTPAYMENT").length > 0 && m_bRepaymentIndEditable)
	{
		//if(m_sMetaAction != "Edit") frmScreen.txtProposedRepaymentDate.value =  m_strDayOfFirstPaymentDate ;
		if(!IsProposedDateEditable()) scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtProposedRepaymentDate");
	}
}

function PopulateCombos()
<% /* Populates all combos on the screen */ %>
{
	//PopulateTPTitleCombo(); MAR334

	// MDC SYS2564 / SYS2785 Client Customisation
	//var sControlList = new Array(frmScreen.cboCountry, frmScreen.cboProposedMethodOfRepayment);
	//objDerivedOperations.GetComboLists(sControlList); MAR334

	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	//var sGroupList = new Array("PaymentStatus");		MAR334
	var sGroupList = new Array("PaymentStatus", "ProposedRepayMethod");

	if(XML.GetComboLists(document,sGroupList))
	{
//		blnSuccess = XML.PopulateCombo(document,frmScreen.cboCountry,"Country",false);
//		blnSuccess = blnSuccess & XML.PopulateCombo(document,frmScreen.cboProposedMethodOfRepayment,"ProposedRepayMethod",true);
		//MAR334
		XML.PopulateCombo(document,frmScreen.cboProposedMethodOfRepayment,"ProposedRepayMethod",true);
		
		var XMLCombos = XML.GetComboListXML("PaymentStatus");
		m_sCancelledStatusId = XML.GetComboIdForValidation("PaymentStatus", "C", XMLCombos)  ;

//		if(blnSuccess == false)
//			scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}

	XML = null;

}

function PopulateScreen()
// Populates the screen with details of the item selected in 
{
	//we are in edit mode. Load the dependant XML from context
	BankXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	BankXML.LoadXML(m_sXML);
	// SYS3734 JLD set up contactdetails tag for display in DC241
	BankXML.SelectTag(null, "CONTACTDETAILS");
	if(BankXML.ActiveTag != null)
	{
		m_sXMLContact = BankXML.ActiveTag.xml;
	}
	BankXML.SelectTag(null,"APPLICATIONTHIRDPARTY");

	m_sDirectoryGUID = BankXML.GetTagText("DIRECTORYGUID");
	m_sThirdPartyGUID = BankXML.GetTagText("THIRDPARTYGUID");
	m_bDirectoryAddress = (m_sDirectoryGUID != "");
	m_RepaymentAccountIndicator = BankXML.GetTagText("REPAYMENTBANKACCOUNTINDICATOR");

	//with (frmScreen)	MAR234
	//{					MAR234
		m_sBankAccountSequenceNumber = BankXML.GetTagText("BANKACCOUNTSEQUENCENUMBER");
		frmScreen.txtCompanyName.value = BankXML.GetTagText("COMPANYNAME");
		frmScreen.txtAccountName.value = BankXML.GetTagText("ACCOUNTNAME");
		frmScreen.txtAccountNumber.value = BankXML.GetTagText("ACCOUNTNUMBER");
		frmScreen.txtRollNumber.value = BankXML.GetTagText("ROLLNUMBER");    <% /* EP2_3 */ %>
		frmScreen.txtSortCode.value = (m_bDirectoryAddress ? BankXML.GetTagText("NAMEANDADDRESSBANKSORTCODE") : BankXML.GetTagText("THIRDPARTYBANKSORTCODE"));
		frmScreen.txtBranchName.value = BankXML.GetTagText("BRANCHNAME");
		scScreenFunctions.SetRadioGroupValue(frmScreen,"SalaryFedIndicatorGroup",BankXML.GetTagText("SALARYFEDINDICATOR"));
		frmScreen.txtTimeAtBank.value = BankXML.GetTagText("TIMEATBANK");
		scScreenFunctions.SetRadioGroupValue(frmScreen,"RepaymentBankAccountGroup",BankXML.GetTagText("REPAYMENTBANKACCOUNTINDICATOR"));
		frmScreen.cboProposedMethodOfRepayment.value = BankXML.GetTagText("PROPOSEDPAYMENTMETHOD");
		frmScreen.txtProposedRepaymentDate.value = BankXML.GetTagText("PREFEREDPAYMENTDAY");
		m_txtProposedRepaymentDate_First = BankXML.GetTagText("PREFEREDPAYMENTDAY");
		frmScreen.txtDDReference.value = BankXML.GetTagText("DDREFERENCE");
		<% /* MAR20 New field: DDExplained radio */ %>
		scScreenFunctions.SetRadioGroupValue(frmScreen,"DDExplainedGroup",BankXML.GetTagText("DDEXPLAINEDIND"));
		//txtContactForename.value = BankXML.GetTagText("CONTACTFORENAME");	MAR234
		//txtContactSurname.value = BankXML.GetTagText("CONTACTSURNAME");	MAR234
		//cboTitle.value = BankXML.GetTagText("CONTACTTITLE");				MAR234
		
		// MDC SYS2564/SYS2785 
		//objDerivedOperations.LoadCounty(BankXML);							MAR234
		// txtCounty.value = BankXML.GetTagText("COUNTY");
		
		//txtDistrict.value = BankXML.GetTagText("DISTRICT");				MAR234
	<% /* 	txtEMailAddress.value = BankXML.GetTagText("EMAILADDRESS"); 
		txtFaxNo.value = BankXML.GetTagText("FAXNUMBER"); 
		txtTelephoneNo.value = BankXML.GetTagText("TELEPHONENUMBER");*/ %>
		//txtFlatNumber.value = BankXML.GetTagText("FLATNUMBER");			MAR234
		//txtHouseName.value = BankXML.GetTagText("BUILDINGORHOUSENAME");	MAR234
		//txtHouseNumber.value = BankXML.GetTagText("BUILDINGORHOUSENUMBER");MAR234
		//txtPostcode.value = BankXML.GetTagText("POSTCODE");				MAR234
		//frmScreen.txtStreet.value = BankXML.GetTagText("STREET");			MAR234
		//txtTown.value = BankXML.GetTagText("TOWN");						MAR234
		//cboCountry.value = BankXML.GetTagText("COUNTRY");					MAR234
	//}
	
	<%/* Call event handlers to set the status of certain controls */%>
	BankWizardDetailsChanged();
	RepaymentBankAccountChanged();
	frmScreen.cboProposedMethodOfRepayment.onchange()

	<% /* MAR333: Maha T (Moved this condition from above Call Event handlers, as BankWizardDetailsChanged()
				  will always set m_bBankAccountValid to false.
	*/ %>
 	if (BankXML.GetTagText("BANKWIZARDINDICATOR") == "1") {
		m_bBankAccountValid = true;
	} else {
		m_bBankAccountValid = false;
	}

}


<% /* MF MAR20 09/08/2005 Created to load XML for third party data
		where screen DC240 is not present in screen flow. DC240 normally 
		loads third party data and stores in the Context object.
*/ %>
function LoadThirdPartyData(v_sTPType)
{
	var sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B00003743");
	var sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");  

	// Retrieve main data
	ThirdPartyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	ThirdPartyXML.CreateRequestTag(window,null);
	var tagApplication = ThirdPartyXML.CreateActiveTag("APPLICATION");
	ThirdPartyXML.CreateTag("APPLICATIONNUMBER", sApplicationNumber);
	ThirdPartyXML.CreateTag("APPLICATIONFACTFINDNUMBER", sApplicationFactFindNumber);
	ThirdPartyXML.RunASP(document,"FindApplicationThirdPartyList.asp");
	ThirdPartyXML.ActiveTag = null;
	ThirdPartyXML.CreateTagList("APPLICATIONTHIRDPARTY");
	<% /* Loop through third parties and select v_sTPType */ %>
	var iCount=0;
	var bFound=false;	
	while (iCount < ThirdPartyXML.ActiveTagList.length && !bFound)
	{
		ThirdPartyXML.SelectTagListItem(iCount);
		var sThirdPartyType = ThirdPartyXML.GetTagText("THIRDPARTYTYPE");
		if(sThirdPartyType == v_sTPType){
			bFound=true;
		}
		iCount++;
	}
	return bFound==true ? ThirdPartyXML.ActiveTag.xml : "";	
}

function RetrieveContextData()
{
<%
/*
	// TEST
	//scScreenFunctions.SetContextParameter(window,"idApplicationNumber","C00078387");
	scScreenFunctions.SetContextParameter(window,"idApplicationNumber","A100002048");
	scScreenFunctions.SetContextParameter(window,"idApplicationFactFindNumber","1");  
	
	scScreenFunctions.SetContextParameter(window,"idMetaAction","Add");
	scScreenFunctions.SetContextParameter(window,"idReadOnly","0");
	scScreenFunctions.SetContextParameter(window,"idProcessingIndicator","1");
	
	//END TEST
*/
%>
	m_sMetaAction = scScreenFunctions.GetContextParameter(window,"idMetaAction","Edit");
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");

	<% /* MF If no third party summary screen (DC240) present then read in the Third Party 
		data normally loaded inside DC240 */ %> 
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();		
	var bThirdPartySummary = XML.GetGlobalParameterBoolean(document,"ThirdPartySummary");			
	if(!bThirdPartySummary){
		m_sXML = LoadThirdPartyData("3");
		if(m_sXML=="")
			m_sMetaAction = "Add";
		else	 
			m_sMetaAction = "Edit";	
	}else if(m_sMetaAction == "Edit")
		m_sXML = scScreenFunctions.GetContextParameter(window,"idXML",null);
	
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","B00003743");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");  

	var sUserRole = scScreenFunctions.GetContextParameter(window,"idRole","0");
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	m_bCanAddToDirectory = (sUserRole >= XML.GetGlobalParameterAmount(document,"AddToDirectoryRole"));

	XML = null;
}

//MAR334 added here from ThirdParty.asp
function ContactDetailsChanged()
{
	// Do nothing - can be added to later
}

function UpdatePaymentDateAndSaveBankBuildingSociety()
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	
	var xmlRequest = XML.CreateRequestTag(window,null)
	XML.CreateTag("DIRECTORYADDRESSINDICATOR", m_bDirectoryAddress ? "1" : "0");

	<% /* start OPERATION  CP_28/04/2001 AQR SYS2050 */ %>
	if (m_sMetaAction == "Edit")
		XML.SetAttribute("OPERATION","CriticalDataCheck"); 
	<% /* end OPERATION  CP_28/04/2001 AQR SYS2050 */ %>

	XML.CreateTag("PREFEREDPAYMENTDAY", frmScreen.txtProposedRepaymentDate.value);
	
	var xmlApplBBS = XML.CreateActiveTag("APPLICATIONBANKBUILDINGSOC");
	// Create tags for the key
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("ACCOUNTNAME", frmScreen.txtAccountName.value);
	XML.CreateTag("ACCOUNTNUMBER", frmScreen.txtAccountNumber.value);
	XML.CreateTag("TIMEATBANK", frmScreen.txtTimeAtBank.value);
	XML.CreateTag("REPAYMENTBANKACCOUNTINDICATOR", scScreenFunctions.GetRadioGroupValue(frmScreen,"RepaymentBankAccountGroup"));
	XML.CreateTag("SALARYFEDINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"SalaryFedIndicatorGroup"));
	XML.CreateTag("PROPOSEDPAYMENTMETHOD", frmScreen.cboProposedMethodOfRepayment.value);
	XML.CreateTag("PREFEREDPAYMENTDAY", frmScreen.txtProposedRepaymentDate.value);
	XML.CreateTag("DDREFERENCE", frmScreen.txtDDReference.value);
	<% /* MAR20 New field: DDExplained radio */ %>
	XML.CreateTag("DDEXPLAINEDIND", scScreenFunctions.GetRadioGroupValue(frmScreen,"DDExplainedGroup"));	
	if (m_bBankAccountValid == true) {
		XML.CreateTag("BANKWIZARDINDICATOR", "1");
	} else {
		XML.CreateTag("BANKWIZARDINDICATOR", "0");
	}

	<% /* MAR39 New field: TransposedIndicator radio */ %>
	if (frmScreen.optTransposedIndicatorYes.checked == true) {
		XML.CreateTag("TRANSPOSEDINDICATOR", "1");
	} else {
		XML.CreateTag("TRANSPOSEDINDICATOR", "0");
	}

	XML.CreateTag("ROLLNUMBER", frmScreen.txtRollNumber.value);	

	if (m_sMetaAction == "Edit")
	{		
		XML.CreateTag("BANKACCOUNTSEQUENCENUMBER", m_sBankAccountSequenceNumber);
		// Retrieve the original third party/directory GUIDs
		var sOriginalThirdPartyGUID = BankXML.GetTagText("THIRDPARTYGUID");
		var sOriginalDirectoryGUID = BankXML.GetTagText("DIRECTORYGUID"); 
		// Only retrieve the address/contact details GUID if we are updating an existing third party record
		var sAddressGUID = (sOriginalThirdPartyGUID != "") ? BankXML.GetTagText("ADDRESSGUID") : "";
		var sContactDetailsGUID = (sOriginalThirdPartyGUID != "") ? BankXML.GetTagText("CONTACTDETAILSGUID") : "";
	}	

	<% // If changing from a directory to a third party, or vice-versa, the old directory/thirdparty GUID
	   // should still be specified to alert the middler tier to the fact that the old link needs deleting
	%>
	XML.CreateTag("DIRECTORYGUID", m_sDirectoryGUID != "" ? m_sDirectoryGUID : sOriginalDirectoryGUID);
	XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
	
	if (!m_bDirectoryAddress)
	{
		// Store the third party details
		XML.CreateActiveTag("THIRDPARTY");
		XML.CreateTag("THIRDPARTYGUID", m_sThirdPartyGUID != "" ? m_sThirdPartyGUID : sOriginalThirdPartyGUID);
		XML.CreateTag("THIRDPARTYTYPE", 3); // Bank/Building society
		XML.CreateTag("COMPANYNAME", frmScreen.txtCompanyName.value);
		XML.CreateTag("THIRDPARTYBANKSORTCODE", frmScreen.txtSortCode.value);
		XML.CreateTag("BRANCHNAME", frmScreen.txtBranchName.value);
		XML.CreateTag("PREFEREDPAYMENTDAY", frmScreen.txtProposedRepaymentDate.value);

		<% /* MAR334 block SaveAddress & ContactDetails leaving ADDRESS tag in */ %>
		<% /* Address */ %>
		XML.CreateTag("ADDRESSGUID", sAddressGUID);
		//SaveAddress(XML, sAddressGUID);
		XML.CreateActiveTag("ADDRESS");
		XML.CreateTag("ADDRESSGUID", sAddressGUID);
		<% /* Contact Details  */ %>	
		<% /* MAR334 Contact Details are not relevan due to GUI changes */ %>
		//XML.SelectTag(null, "THIRDPARTY");
		<% /* SR 18/05/00 SYS0689: Append Contact Details only when required */ %>
		//if (m_sMetaAction == "Edit")
		//	if (sContactDetailsGUID != "")  <% /*Contact details were existing previously */%> 
		//	{
		//		XML.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
		//		SaveContactDetails(XML, sContactDetailsGUID);
		//	}
		//	else
		//	{
		//		if (!IsContactEmpty())
		//		{
		//			XML.CreateTag("CONTACTDETAILSGUID", sContactDetailsGUID);
		//			SaveContactDetails(XML, sContactDetailsGUID);
		//		}
		//	}
		//else     <% /* New Contact to be saved */%> 
		//	if (!IsContactEmpty())SaveContactDetails(XML, sContactDetailsGUID); 
	}

	if(m_sInitialFirstPaymentDate != frmScreen.txtProposedRepaymentDate.value || scScreenFunctions.GetRadioGroupValue(frmScreen,"RepaymentBankAccountGroup") == "1")
	{			
		XML.ActiveTag = xmlRequest ;
		XML.CreateActiveTag("LOANCOMPONENTPAYMENT");
		XML.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	}
	
	if(m_sMetaAction == "Edit")
	{
		XML.SelectTag(null,"REQUEST");
		XML.CreateActiveTag("CRITICALDATACONTEXT");
		XML.SetAttribute("APPLICATIONNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null));
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null));
		XML.SetAttribute("SOURCEAPPLICATION","Omiga");
		XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
		XML.SetAttribute("ACTIVITYINSTANCE","1");
		XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
		XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
		XML.SetAttribute("COMPONENT","omPayProc.PaymentProcessingBO");
		XML.SetAttribute("METHOD","omPayProcRequest");
		XML.SetAttribute("OPERATION","UpdateFirstPaymentDateAndBankBuildingSoc");
					
		window.status = "Critical Data Check - please wait";

		// 		XML.RunASP(document,"OmigaTMBO.asp");
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

		window.status = "";		
	}
	else
	{
		XML.SelectTag(null,"REQUEST");
		XML.SetAttribute("OPERATION", "UPDATEFIRSTPAYMENTDATEANDCREATEBANKBUILDINGSOC");
		
		// 		XML.RunASP(document,"PaymentProcessingRequest.asp");
		// Added by automated update TW 09 Oct 2002 SYS5115
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
					XML.RunASP(document,"PaymentProcessingRequest.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	}
	
	bSuccess = XML.IsResponseOK();
	XML = null;
	return(bSuccess);
}
-->
</script>
<% /* MAR334 Following the changes include ThirdPartyDetails should be blocked */ %>
<% /*<!-- #include FILE="includes/thirdpartydetails.asp" --> */ %>
</body>
</html>


