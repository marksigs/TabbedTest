<%@ LANGUAGE="JSCRIPT" %>
<html>
<% /* 
Workfile:      DC200.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   New Loan & Property
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AD		28/02/2000	Created
AD		22/03/2000	SYS0395 - Removed ownership percentage
IW		29/03/2000	SYS0546 - Removed SameMortgageType,SameLoanFinish & Term etc
AY		31/03/00	New top menu/scScreenFunctions change
IW		02/05/00	SYS0446 - See Deuce for 9 changes made.
BG		17/05/00	SYS0752 Removed Tooltips
MH      21/06/00    SYS0792 SharedProperty/Owernship changes
BG		12/09/00	SYS1040	SharedOwnership changes.
CL		05/03/01	SYS1920 Read only functionality added
TJ		30/03/01	SYS2050 Critical Data functionality added
GD		11/05/01	SYS2050 Critical Data functionality ROLLED BACK
SA		30/05/01	SYS1267 Changed Maxlength of txtDiscount from 10 to 6.
							Added validation: Discount amount must be less than
							or equal to Purchase price. If Discount radion button set to yes,
							discount amount is mandatory.
DC      20/07/01    SYS2038 Critical Data functionality ROLLED FORWARD AGAIN							
SA		25/07/01	SYS2378 FamilySale/Right to Buy Discount radio button not working properly with keyboard
JLD		10/12/01	SYS2806 Use screen functions for completeness check routing
STB		17/04/02	SYS3481 Ensure shared amount is always less than the property amount.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		10/06/2002	BMIDS00032 - Modified PopulateScreen(); SaveLoanPropertyDetails(); ClearFields();
					Amended New Fields
GD		19/06/2002	BMIDS00077 - Upgrade to Core 7.0.2
ASu		06/09/2002	BMIDS00385 - Increse Valuation/Purchase price amount field to 7 Characters
MV		02/10/2002	BMIDS00536	Amended SaveLoanPropertyDetails()
TW      09 Oct 2002    Modified to incorporate client validation - SYS5115
MV		11/10/2002	BMIDS00590	Modified SaveLoanPropertyDetails() and PopulateScreen()
MO		07/11/2002	BMIDS00815  Made change so thatb if PropertyType has more than 1 combo validation type 
									in it, it doesnt make a complete mess of the property descriptions
SA		19/11/2002	BMIDS01002	Calls to Screen Rules added
AW		02/12/2002	BMIDS01120	Do conversion for integer comparisons
AW		02/12/2002	BM0116		Remove discount box and re-word question
AW		06/12/2002	BM0164		Extend Purchase Price length to 9
BS		17/02/2003	BM0187		Allow valuation types with validation type 'N' for all application types
MV		25/03/2003	BM0063      Amended the HTML Option buttons 
HMA     24/11/2003  BMIDS673    Enable Shared Ownership field.
MC		06/05/2004	bmids468	Radio Buttons Alignment problem Fixed
MC		22/06/2004	BMIDS772	DEFECT FIXES FOR BMIDS772 - ROUTING ETC (DC110,DC200,CM010)
JD		06/07/2004	BMIDS772	Add global parameter changes and stage checking before moving to CM010
								(functions MoveToStage, CheckStage, GetLocalPrinter, GetStageName copied from FW020)
SR		14/07/2004  BMIDS796
 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
INGDUK Specific History:
JD		23/09/2005	MAR40		changes for ESurv/Hometrack valuation  
PE		14/12/2005	MAR868		If the global parameter "NewPropertySummary" is 0, the screen flow is changed.
PE		22/02/2006	MAR1311		Need to make Property description mandatory
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
EPSOM Specific History:

SAB		03/04/2006	EP8			Updated routing
MAH		24/11/2006	E2CR35		Added BUSINESSCONNECTION, FULLMARKETVALUE,FINANCIALINCENTIVES from NEWPROPERTY
AShaw	28/11/2006	EP2_2		Valuation Report Alignment.
AShaw	15/12/2006	EP2_18		Add Pre-emption End date and Discount for Family Sales.
								New method to set default values for Family Sales.
DS      17/01/2007  EP2_611     Fixed Defect EP2_611 - Now checking that cboValuationType.value 
								should not be null before calling IsInComboValidationList method
DS      24/01/2007  EP2_786     Fixed Defect EP2_786 - Commented NewProperty select tag 
AShaw	01/02/2007	EP2_1002	Move question to DC210
INR		13/02/2007	EP2_780 / EP2_788
PE		12/02/2007  EP2_1287    Bind txtSharedAmount to SHAREDAMOUNT. SHAREDPERCENTAGE is calculated in ApplicationBO.cls.
LDM		15/02/2007	EP2_1165	Added Direct financial benefit ind
INR		11/03/2007	EP2_1731 	Don't disable fields if additional borrowing
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<HEAD>
	<meta name=vs_targetSchema content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<OBJECT data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<% /* SCRIPTLETS */ %>
<script src="validation.js" language="JScript"></script>
<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2
/*
<OBJECT data=scScreenFunctions.asp height=1 id=scScrFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scMathFunctions.asp height=1 id=scMathFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
<OBJECT data=scXMLFunctions.asp height=1 id=scXMLFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 VIEWASTEXT></OBJECT>
*/
%>
<% /* FORMS */ %>
<form id="frmToDC195" method="post" action="DC195.asp" STYLE="DISPLAY: none"></form>
<form id="frmToMN060" method="post" action="MN060.asp" STYLE="DISPLAY: none"></form>
<% /* SAB - 03/04/2006 - EP8 */ %>
<form id="frmToDC085" method="post" action="DC085.asp" STYLE="DISPLAY: none"></form>

<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM010" method="post" action="CM010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC210" method="post" action="DC210.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC202" method="post" action="DC202.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark   validate ="onchange">
<div style="HEIGHT: 345px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">
		Type of Application
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtTypeOfApplication" maxlength="45" style="POSITION: absolute; WIDTH: 230px" class="msgReadOnly" readonly tabindex=-1>
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Type of Property
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<select id="cboTypeOfProperty" style="WIDTH: 230px" class="msgCombo"></select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Property Description
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<select id="cboPropertyDescription" style="WIDTH: 230px" class="msgCombo"></select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 84px" class="msgLabel">
		Property Location
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<select id="cboPropertyLocation" style="WIDTH: 230px" class="msgCombo"></select>
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 108px" class="msgLabel">
		Purchase Price/Estimated Value
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtPurchasePrice" maxlength="9" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
		</span> 
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 132px" class="msgLabel">
		Valuation Type
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<select id="cboValuationType" style="WIDTH: 230px" class="msgCombo"></select>
		</span>
		<span style="TOP: -3px; LEFT: 450px; POSITION: ABSOLUTE">
			<input id="btnRetype" value="Retype" type="button" style="WIDTH: 60px" class="msgButton" NAME="btnRetype">
		</span>
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 156px" class="msgLabel">
		Present Valuation
		<span style="LEFT: 180px; POSITION: absolute; TOP: -3px">
			<input id="txtPresentValuation" maxlength="9" style="POSITION: absolute; WIDTH: 100px" readonly class="msgReadOnly">
		</span> 
		<span style="LEFT: 313px; POSITION: absolute; TOP: 0px" class="msgLabel">
		Pre-emption End Date
			<span style="LEFT: 120px; POSITION: absolute; TOP: 0px">
				<input id="txtPreEmptionDate" maxlength="10" style="POSITION: absolute; WIDTH: 100px" class="msgTxt" NAME="txtPreEmptionDate">
			</span> 
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 180px" class="msgLabel">
		Family Sale / Right To Buy<br>Discount?
		<span style="LEFT: 178px; POSITION: absolute; TOP: 5px">
			<input id="optFamilySaleYes" name="FamilySaleGroup" type="radio" value="1"><label for="optFamilySaleYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 240px; POSITION: absolute;  TOP: 5px">
			<input id="optFamilySaleNo" name="FamilySaleGroup" type="radio" value="0"><label for="optFamilySaleNo" class="msgLabel">No</label> 
		</span> 
		<span style="LEFT: 313px; POSITION: absolute; TOP: 5px" class="msgLabel">
		Discount
			<span style="LEFT: 60px; POSITION: absolute; TOP: -3px">
				<input id="txtDiscount" maxlength="6" style="POSITION: absolute; WIDTH: 100px" class="msgTxt" NAME="txtDiscount">
			</span> 
		</span>
	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 214px" class="msgLabel">
		Is this a Shared Ownership Case?
		<span style="LEFT: 178px; POSITION: absolute; TOP: -3px">
			<input id="optSharedOwnershipCaseYes" name="SharedOwnershipCaseGroup" type="radio" value="1" onclick="SharedOwnershipCaseChanged()"><label for="optSharedOwnershipCaseYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 240px; POSITION: absolute; TOP: -3px">
			<input id="optSharedOwnershipCaseNo" name="SharedOwnershipCaseGroup" type="radio" value="0" onclick="SharedOwnershipCaseChanged()"><label for="optSharedOwnershipCaseNo" class="msgLabel">No</label> 
		</span>
		
		<span style="LEFT: 143px; POSITION: relative; TOP: 0px" class="msgLabel">
		Amount
			<span style="LEFT:60px; POSITION: absolute; TOP: -3px">
				<input id="txtSharedAmount" maxlength="6" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
			</span> 
		</span>
		 
	</span>
		
	<span style="LEFT: 4px; POSITION: absolute; TOP: 245px" class="msgLabel">
		Is the property being purchased at full market value?
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="optFullMarketValueYes" name="FullMarketValueGroup" type="radio" value="1"><label for="optFullMarketValueYes" class="msgLabel">Yes</label> 
		</span> 
		<span style="LEFT: 350px; POSITION: absolute; TOP: -3px">
			<input id="optFullMarketValueNo" name="FullMarketValueGroup" type="radio" value="0"><label for="optFullMarketValueNo" class="msgLabel">No</label> 
		</span>		 
	</span>		
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 274px" class="msgLabel">
		Are there any other financial incentives being offered?
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="optFinancialIncentivesYes" name="FinancialIncentivesGroup" type="radio" value="1"><label for="optFinancialIncentivesYes" class="msgLabel">Yes</label> 
		</span> 
		<span style="LEFT: 350px; POSITION: absolute; TOP: -3px">
			<input id="optFinancialIncentivesNo" name="FinancialIncentivesGroup" type="radio" value="0"><label for="optFinancialIncentivesNo" class="msgLabel">No</label> 
		</span>		 
	</span>			
	<span style="LEFT: 4px; POSITION: absolute; TOP: 303px" class="msgLabel">
		Is the loan for the direct financial benefit and advantage<BR> of all applicants?
		<span style="LEFT: 300px; POSITION: absolute; TOP: -3px">
			<input id="optDirectFinancialBenefitYes" name="DirectFinancialBenefitGroup" type="radio" value="1"><label for="optDirectFinancialBenefitYes" class="msgLabel">Yes</label> 
		</span> 
		<span style="LEFT: 350px; POSITION: absolute; TOP: -3px">
			<input id="optDirectFinancialBenefitNo" name="DirectFinancialBenefitGroup" type="radio" value="0"><label for="optDirectFinancialBenefitNo" class="msgLabel">No</label> 
		</span>		 
	</span>			
</div>

<div id="divFurtherAdvance" style="HEIGHT: 136px; LEFT: 10px; POSITION: absolute; TOP: 408px; WIDTH: 604px" class="msgGroup">
	<span style="LEFT: 4px; POSITION: absolute; TOP: 10px" class="msgLabel">

			<strong>Further Advance</strong>

	</span>

	<span style="LEFT: 4px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Is a Solicitor required?
		<span style="LEFT: 230px; POSITION: absolute; TOP: -3px">
			<input id="optSolicitorYes" name="SolicitorGroup" type="radio" value="1"><label for="optSolicitorYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 290px; POSITION: absolute; TOP: -3px">
			<input id="optSolicitorNo" name="SolicitorGroup" type="radio" value="0"><label for="optSolicitorNo" class="msgLabel">No</label> 
		</span> 
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Does Futher Advance fit within exisiting term?
		<span style="LEFT: 230px; POSITION: absolute; TOP: -3px">
			<input id="optFurtherAdvanceTermYes" name="FurtherAdvanceTermGroup" type="radio" value="1"><label for="optFurtherAdvanceTermYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 290px; POSITION: absolute; TOP: -3px">
			<input id="optFurtherAdvanceTermNo" name="FurtherAdvanceTermGroup" type="radio" value="0"><label for="optFurtherAdvanceTermNo" class="msgLabel">No</label> 
		</span> 
	</span>
	
	<span style="LEFT: 4px; POSITION: absolute; TOP: 84px" class="msgLabel">
		Further Advance Repayment Type to be<BR>the same as the exisiting loan?
		<span style="LEFT: 230px; POSITION: absolute; TOP: -3px">
			<input id="optFurtherAdvanceRepTypeYes" name="FurtherAdvanceRepTypeGroup" type="radio" value="1"><label for="optFurtherAdvanceRepTypeYes" class="msgLabel">Yes</label> 
		</span> 

		<span style="LEFT: 290px; POSITION: absolute; TOP: -3px">
			<input id="optFurtherAdvanceRepTypeNo" name="FurtherAdvanceRepTypeGroup" type="radio" value="0"><label for="optFurtherAdvanceRepTypeNo" class="msgLabel">No</label> 
		</span> 
	</span>
	
	
</div> 

</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 640px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- #include FILE="attribs/DC200attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sTypeOfApplicationDesc = "";
var m_sTypeOfApplication = "";
var LoanPropertyXML = null;
var ComboXML = null;
var scScreenFunctions;
var bShared = false;
var m_blnReadOnly = false;
var mbRetypeValuationDetails = false;
var XMLRetype = new top.frames[1].document.all.scXMLFunctions.XMLObject(); // EP2_2 AShaw 28/11/2006
var m_sRetypeXML = "";  //EP2_2 Retype Valuation XML


/* EVENTS - BUTTONS */

function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
	{
		// MAR868
		// Peter Edney - 14/12/2005
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var bNewPropertySummary = XML.GetGlobalParameterBoolean(document, "NewPropertySummary");
					
		if (bNewPropertySummary)	
		{
			scScreenFunctions.SetContextParameter(window,"idMetaAction", "");
			frmToDC195.submit();
		}
		else
		{
			<% /* SAB - 03/04/2006 - EP8 */ %>
			if (XML.GetGlobalParameterBoolean(document, "HideDecision"))
				frmToDC085.submit();
			else
				frmToMN060.submit();
		}
	}
}

function btnSubmit.onclick()
{
	<% /* SYS3481 - Ensure shared amount is always less than the property amount. */ %>
	if(!(ValidateSharedAmount()))
	{
		frmScreen.txtSharedAmount.focus();
		return;
	}
	
	//	AW	02/12/2002	BM0116 - Start
	//<%/* SYS1267 Extra validation addeed */%>
	// EP2_18 - All re-enabled.
	if((frmScreen.optFamilySaleYes.checked & (frmScreen.txtDiscount.value == "0" | frmScreen.txtDiscount.value == "" )))
	{	
		alert("Discount amount cannot be zero")
		frmScreen.txtDiscount.focus ();
		return false
	}
	//AW	02/12/2002	BMIDS01120
	if (parseInt(frmScreen.txtDiscount.value, 10) > parseInt(frmScreen.txtPurchasePrice.value, 10))
	{
		alert("Discount amount must be less than or equal to purchase price")
		frmScreen.txtDiscount.focus ();
		return false
	}
	//<%/* SYS1267 Extra validation addeed - end */%>
	//	AW	02/12/2002	BM0116 - End
	// EP2_18 - END All re-enabled.
	
	// EP2_18 - Add Date test.
	if (frmScreen.txtPreEmptionDate.value != "")
	{
		var PreemptiveDate = frmScreen.txtPreEmptionDate.value
		if (scScreenFunctions.CompareDateStringToSystemDate(PreemptiveDate , "<=") == true)
		{
			alert("Pre-emption End Date must be in the future.")
			frmScreen.txtPreEmptionDate.focus ();
			return false
		}
	}
	
	// EP2_2 - Warn if losing Retype data.
	var ValTXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	//DS - EP2_611
	if (frmScreen.cboValuationType.value.length != 0)
		var bIsRetype = ValTXML.IsInComboValidationList(document,"ValuationType",frmScreen.cboValuationType.value,Array("RT"));
	//Begin: DS - EP2_786
	//LoanPropertyXML.SelectTag(null,"NEWPROPERTY");
	//End: DS - EP2_786
	var bHasRetypeData = (LoanPropertyXML.GetTagText("DATEOFORIGINALVALUATION") != "");
	if (bIsRetype == false && bHasRetypeData == true)
	{	
		if (!confirm("Changing the Valuation Type from Retype will result in the deletion of Retype Valuation Details. Are you sure?"))
		{	
			frmScreen.cboValuationType.focus();
			return;
		}
	}
	// EP2_2 - Warn if Retype selected but no data (Naughty user!)
	if (bIsRetype == true && bHasRetypeData == false)
	{	
		alert("A Valuation Type of Retype requires the entry of Retype Valuation Details.")
		frmScreen.cboValuationType.focus();
		return;
	}
	
	if (frmScreen.txtSharedAmount.value == "0")
	{	
		alert("Shared amount cannot be zero")
		frmScreen.txtSharedAmount.focus();
	}
	else if (CommitChanges())
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
		{
			scScreenFunctions.SetContextParameter(window,"idReturnScreenId","MN060");
			frmToGN300.submit();
		}
		else
		{
			//frmToMN060.submit();
			//*[MC]BMIDS772 - Routing return screen set to DC200
			//scScreenFunctions.SetContextParameter(window,"idReturnScreenId","DC200");
			//frmToCM010.submit();
			
			// MAR868
			// Peter Edney - 14/12/2005
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var bNewPropertySummary = XML.GetGlobalParameterBoolean(document, "NewPropertySummary");
						
			if (bNewPropertySummary)
			{
				GoToCostModelling(); //JD BMIDS772 Carry out stage checks and global parameter changes
			}
			else
			{
				frmToDC210.submit();
			}

		}
	}
}

function GoToCostModelling()
{
	// Carry out stage checks and global parameter changes before going to CM010
	if(CheckStage())
	{
		scScreenFunctions.SetContextParameter(window,"idApplicationMode","Cost Modelling");
		scScreenFunctions.SetContextParameter(window,"idReturnScreenId","DC200");
		frmToCM010.submit();
	}
	else
	{
		// we have already committed the data, so make sure data isn't created again
		m_sMetaAction = "Edit";
	}
}
function GetStageName(sStageId)
{
<%	/* returns the stage name for the stage id sent in */
%>	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var saReturnArray = new Array();
	XML.CreateRequestTag(window, "GetStageDetail");
	XML.CreateActiveTag("STAGE");
	XML.SetAttribute("STAGEID", sStageId);
	XML.RunASP(document, "MsgTMBO.asp");
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null, "STAGE");
		sStageName = XML.GetAttribute("STAGENAME");
	}
	XML = null;
	return(sStageName);
}
function CheckStage()
{
	var sStageName = "";
	var sParamStageID = "";
	var sParamStageSeq = "";
	var sNewStageSeq = "";
	var sNewStageId = "";
	var nCurrentStageSeq;
	var nNewStageSeq;
	var nNewStageID;
	
	var GlobalParamXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	sNewStageId = GlobalParamXML.GetGlobalParameterString(document,"TMCMStageID");
	sStageName = GetStageName(sNewStageId);

	//Errors?
	if (GlobalParamXML.IsResponseOK == false)
		return false;
			
	//Blank parameter?
	if (sNewStageId == "")
	{
		//Yes - route to screen
		return true;		
	}
	else			
	{
		//No - get the stage sequence no
		sNewStageSeq = GlobalParamXML.GetGlobalParameterAmount(document,"TMCostModellingStage");
		
		//Error?
		if (GlobalParamXML.IsResponseOK == false)
			return false;
		
		//Blank parameter?
		if (sNewStageSeq == "")
		{
			alert("The parameter TMCostModellingStage has not been defined within Supervisor. Please refer to your System Administrator");
			return false;
		}
	}
	
	//Get current stage from context
	nCurrentStageSeq = parseInt(scScreenFunctions.GetContextParameter(window,"idCurrentStageOrigSeqNo",null));
	nNewStageSeq = parseInt(sNewStageSeq);
	nNewStageID = parseInt(sNewStageId);
	
	if (nNewStageSeq == nCurrentStageSeq + 1)
	{
		if (MoveToStage(sStageName, sNewStageId, sNewStageSeq))
			return true;
		else
			return false;
	}
	else
	{
		// may already be at CM stage, just set the parameters
		if (nNewStageSeq == nCurrentStageSeq)
		{
			scScreenFunctions.SetContextParameter(window,"idStageName",sStageName);
			scScreenFunctions.SetContextParameter(window,"idCurrentStageOrigSeqNo",sNewStageSeq);
			scScreenFunctions.SetContextParameter(window,"idStageId",sNewStageId);
			
			// AQR SYS4530 - Check whether a new AdminSys account number was added in an Automatic task
			var sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","");
			scScreenFunctions.SetOtherSystemACNoInContext(sApplicationNumber);
			return true
		}
		return true // will be at > CM stage so global params will be ok.
	}
}
function MoveToStage(sStageName, sNewStageId, sNewStageSeq)
{
	if (scScreenFunctions.IsMainScreenReadOnly(window, "") == true) 
	{
		alert("You cannot advance to the next stage of a Read Only application.");
		return false;
	}

	if (!confirm("Do you wish to progress the application to the " + sStageName + " stage?"))
		return false;

	var sLocalPrinters = GetLocalPrinters();
	LocalPrintersXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	LocalPrintersXML.LoadXML(sLocalPrinters);
	LocalPrintersXML.SelectTag(null, "RESPONSE");	
		
	var sPrinter = LocalPrintersXML.GetTagText("PRINTER[DEFAULTINDICATOR='1']/PRINTERNAME");
								
	if(sPrinter == "")
	{					
		alert("You do not have a default printer set on your PC.");	
		return false;		
	}		
		
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		
	ReqTag = XML.CreateRequestTag(window, "MoveToStage");
	XML.CreateActiveTag("CASESTAGE");
	XML.SetAttribute("STAGEID", sNewStageId);
	XML.SetAttribute("SOURCEAPPLICATION", "Omiga");
	XML.SetAttribute("CASEID", scScreenFunctions.GetContextParameter(window,"idApplicationNumber",""));
	XML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window,"idActivityId",""));
	XML.SetAttribute("ACTIVITYINSTANCE", "1");
	XML.SetAttribute("CASESTAGESEQUENCENUMBER", scScreenFunctions.GetContextParameter(window,"idCurrentStageOrigSeqNo",""));
			
	XML.ActiveTag = ReqTag;
	XML.CreateActiveTag("APPLICATION");
	var sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","")
	XML.SetAttribute("APPLICATIONNUMBER", sApplicationNumber);
	XML.SetAttribute("APPLICATIONFACTFINDNUMBER", scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",""));
	XML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window,"idApplicationPriority",""));
		
	var iCount = 0;
	var sCustomerVersionNumber = "";
	var sCustomerNumber = "";
	for (iCount = 1; iCount <= 5; iCount++)
	{								
		sCustomerNumber = scScreenFunctions.GetContextParameter(window,"idCustomerNumber" + iCount,null);
		if (sCustomerNumber != "")
		{	
			sCustomerVersionNumber = scScreenFunctions.GetContextParameter(window,"idCustomerVersionNumber" + iCount,null);	
			XML.SelectTag(null, "APPLICATION");
			XML.CreateActiveTag("CUSTOMER");				
			XML.SetAttribute("CUSTOMERNUMBER", sCustomerNumber);
			XML.SetAttribute("CUSTOMERVERSIONNUMBER", sCustomerVersionNumber);																		
		}
	}
			
	XML.ActiveTag = ReqTag;
	XML.CreateActiveTag("PRINTER");
	XML.SetAttribute("PRINTERNAME", sPrinter);
	XML.SetAttribute("DEFAULTIND", "1");	
	
	XML.RunASP(document, "omTmNoTxBO.asp");
	
	<% /* MO - 25/11/2002 - BMIDS01076 - Capture Automatted task errors */ %> 
	var ErrorTypes = new Array("MANDTASKSOUTSTANDING", "AUTOTASKERROR");
	var ErrorReturn = XML.CheckResponse(ErrorTypes);
	if(ErrorReturn[1] == ErrorTypes[0])
		alert("There are mandatory tasks still outstanding on the current stage.  You must complete these before moving to another stage");
	else if((ErrorReturn[0] == true) || (ErrorReturn[1] == ErrorTypes[1]))
	{
		<% /* MO - 25/11/2002 - BMIDS01076 - An automatted task errored */ %>
		if (ErrorReturn[1] == ErrorTypes[1]) {
			<% /* Show the error message */ %>
			alert(ErrorReturn[2]);
		}
			
		scScreenFunctions.SetContextParameter(window,"idStageName",sStageName);
		scScreenFunctions.SetContextParameter(window,"idCurrentStageOrigSeqNo",sNewStageSeq);
		scScreenFunctions.SetContextParameter(window,"idStageId",sNewStageId);
			
		// AQR SYS4530 - Check whether a new AdminSys account number was added in an Automatic task
		var sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","")
		scScreenFunctions.SetOtherSystemACNoInContext(sApplicationNumber);
			
		return true;
	}
	return false;
}
function GetLocalPrinters()
{
	var strOut = "";
	var objOmPC = new ActiveXObject("omPC.PCAttributesBO");
	if (objOmPC != null)
	{
		var strXML = "<?xml version='1.0'?><REQUEST ACTION='CREATE'></REQUEST>";
		strOut = objOmPC.FindLocalPrinterList(strXML);
		objOmPC = null;
	}
	return strOut;
}

<% /* EVENTS - CLICKS */ %>

function frmScreen.optSharedOwnershipCaseYes.onclick()
{
	scScreenFunctions.SetFieldToWritable(frmScreen,"txtSharedAmount");
	frmScreen.txtSharedAmount.setAttribute("required", "true");
	bShared = true;
}

function frmScreen.optSharedOwnershipCaseNo.onclick()
{
	frmScreen.txtSharedAmount.setAttribute("required", "false");
	frmScreen.txtSharedAmount.value = "";
	scScreenFunctions.SetFieldToDisabled(frmScreen,"txtSharedAmount");
	bShared = false;
	
}
<% /* SR 14/07/2004 : BMIDS796  */ %>
function frmScreen.cboTypeOfProperty.onchange()
{
	if (frmScreen.cboTypeOfProperty.selectedIndex > 0)
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboPropertyDescription");
		PopulatePropertyDescriptionCombo();

		<% //MAR1311 %>
		<% //Peter Edney - 22/02/2006 %>
		var sPropertyTypeIndex = frmScreen.cboTypeOfProperty.selectedIndex;	
		if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfProperty, sPropertyTypeIndex, "MPD"))
		{
			frmScreen.cboPropertyDescription.setAttribute("required","true");	    
			frmScreen.cboPropertyDescription.parentElement.parentElement.style.color = "red";	
		}
		else   
		{
			frmScreen.cboPropertyDescription.removeAttribute("required");	    	    
			frmScreen.cboPropertyDescription.parentElement.parentElement.style.color = "";	
		}	

	}
	else
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboPropertyDescription");
}
<% /* SR 14/07/2004 : BMIDS796  - End */ %>

function frmScreen.cboTypeOfProperty.onclick() 
{
	if (frmScreen.cboTypeOfProperty.selectedIndex > 0)
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"cboPropertyDescription");
		PopulatePropertyDescriptionCombo();
	}
	else
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"cboPropertyDescription");
		
}

function frmScreen.optFamilySaleNo.onclick()
{
	//SYS2378 Use the checked property not the value.
	//if (frmScreen.optFamilySaleNo.value == "0")	scScreenFunctions.SetFieldToDisabled(frmScreen,"txtDiscount");
	//	AW	02/12/2002	BM0116
	// EP2_18 - Re-enabled.
	if (frmScreen.optFamilySaleNo.checked) 
	{
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtDiscount");
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtPreEmptionDate");
	}
}	

function frmScreen.optFamilySaleYes.onclick()
{
	//SYS2378 Use the checked property not the value.
	//if (frmScreen.optFamilySaleYes.value == "1") scScreenFunctions.SetFieldToWritable(frmScreen,"txtDiscount");
	//	AW	02/12/2002	BM0116
	// EP2_18 - Re-enabled.
	if (frmScreen.optFamilySaleYes.checked) 
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtDiscount");
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtPreEmptionDate");
	}
}

<% /* EVENTS - BLURS */ %>

function frmScreen.txtSharedAmount.onblur()
{	
	<% /* SYS3481 - Ensure shared amount is always less than property value. */ %>
	ValidateSharedAmount();
}

function frmScreen.optFamilySaleNo.onblur()
{
	frmScreen.optFamilySaleNo.onclick();
}

function frmScreen.optFamilySaleYes.onblur()
{
	frmScreen.optFamilySaleYes.onclick();
}


<% /* EVENTS - OTHERS */ %>

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
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	scClientScreenFunctions = new scClientFunctions.ClientScreenFunctionsObject();
	GetRulesData();
	var sButtonList = new Array("Submit","Cancel");
	ShowMainButtons(sButtonList);

	<%//scScreenFunctions = new scScrFunctions.ScreenFunctionsObject();%>
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	scScreenFunctions = new top.frames[1].document.all.scScrFunctions.ScreenFunctionsObject();
	FW030SetTitles("New Loan & Property","DC200",scScreenFunctions);
	
	RetrieveContextData();
	SetMasks();

	Validation_Init();	
	Initialise();	

	scScreenFunctions.SetFocusToFirstField(frmScreen);
		
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	
	EnableRadios();	
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC200");	
	if (m_blnReadOnly == true) m_sReadOnly = "1";	
	
	// Added by automated update TW 09 Oct 2002 SYS5115
	ClientPopulateScreen();
}

/* FUNCTIONS */
<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
function EnableRadios()
{

// AQR SYS4882&4   - DRC 19/06/02 - changed the logic here so that the Radios are disabled for
//                               Further Advances and also the amount fields are only writable
//                               if the radio buttons are on 

	if(scScreenFunctions.GetRadioGroupValue(frmScreen,"SharedOwnershipCaseGroup") != "1")
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtSharedAmount");
		bShared = false;
	}
	else
	{
		<% /*BMIDS673 Always enable Shared Ownership field 
	    if (m_sTypeOfApplication != "F")                        */ %>
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtSharedAmount");
		bShared = true;
	}
	
	//	AW	02/12/2002	BM0116 - Start	
	// EP2_18 - re-enabled code and add date
	if(scScreenFunctions.GetRadioGroupValue(frmScreen,"FamilySaleGroup") != "1")
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtDiscount");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPreEmptionDate");
	}
	//else
	// if (m_sTypeOfApplication != "F")
	//	scScreenFunctions.SetFieldToWritable(frmScreen,"txtDiscount");
	//	AW	02/12/2002	BM0116 - End
//EP2_1731 		
//	if (m_sTypeOfApplication == "F")
//	{	
		//	AW	02/12/2002	BM0116
		// EP2_18 - re-enabled and add date
//		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtDiscount");
//		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtPreEmptionDate");		
//		scScreenFunctions.SetFieldToDisabled(frmScreen,"optFamilySaleYes");
//		scScreenFunctions.SetFieldToDisabled(frmScreen,"optFamilySaleNo");
//        scScreenFunctions.SetRadioGroupToDisabled(frmScreen,"FamilySaleGroup");
        
        <% /*BMIDS673 Always enable Shared Ownership field
		scScreenFunctions.SetRadioGroupToDisabled(frmScreen,"SharedOwnershipCaseGroup"); 
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtSharedAmount"); */ %>
//	}		
	
}


function ClearFields(bFurtherAdvance,bOther)
{
	with (frmScreen)
	{
		if (bFurtherAdvance)
		{
			scScreenFunctions.SetRadioGroupValue(frmScreen,"SameMortgageGroup","0");
			scScreenFunctions.SetRadioGroupValue(frmScreen,"SolicitorGroup","0");
			<% /* MV - 06/06/2002 - BMIDS00032 - Amended New Columns */ %>
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FurtherAdvanceTermGroup","0");
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FurtherAdvanceRepTypeGroup","0");
		}

		if (bOther)
		{
			cboPropertyDescription.selectedIndex = 0;
			cboPropertyLocation.selectedIndex = 0;
			cboTypeOfProperty.selectedIndex = 0;
			//JD MAR40 default valuation type
			//cboValuationType.selectedIndex = 0;
			DefaultValuationType();
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FamilySaleGroup","0");
			//	AW	02/12/2002	BM0116
			//EP12_18 - re-enabled, and add txtPreEmptionDate.
			txtDiscount.value = "";
			txtPreEmptionDate.value = "";
			
			txtPurchasePrice.value = "";
		}
	}
}

function DefaultValuationType()
{
	//choose valuation type based on type of mortgage
	if(m_sTypeOfApplicationDesc == "R")
		scScreenFunctions.SetComboOnValidationType(frmScreen, "cboValuationType", "AU");
	else
		scScreenFunctions.SetComboOnValidationType(frmScreen, "cboValuationType", "SC1");
}
function CommitChanges()
{
	var bSuccess = true;
	var bSaveLoanPropertyDetails = true;

	if (m_sReadOnly != "1")
		if(frmScreen.onsubmit())
		{
			if (IsChanged())
				bSuccess = SaveLoanPropertyDetails();
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function DefaultFields()
// Inserts default values into all fields
{
	ClearFields(true,true);
	frmScreen.optFamilySaleNo.onclick();
	InitialiseRightToBuy();	
}

function Initialise()
// Initialises the screen
{
	PopulateCombos();
	PopulateScreen();

	// If Retype button Visible AND Retype then enable button, else Disabled.
	// EP2_2 - Retrieve GParam and set VisCond on button.
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	mbRetypeValuationDetails = XML.GetGlobalParameterBoolean(document, "RetypeValuationDetails");
	if (mbRetypeValuationDetails == true)
		frmScreen.btnRetype.style.visibility="visible";
	else
		frmScreen.btnRetype.style.visibility="hidden";
	// Now set the enabled status
	frmScreen.cboValuationType.onclick()


	frmScreen.txtTypeOfApplication.value = m_sTypeOfApplicationDesc;

	if (m_sMetaAction == "Add")
		DefaultFields();

	if (m_sTypeOfApplication == "F")
		scScreenFunctions.ShowCollection(divFurtherAdvance);
	else
		scScreenFunctions.HideCollection(divFurtherAdvance);

	//JD MAR40 set present valuation based on valuationtype
	PopulatePresentValuation();
	
	<%/* Fire off events to hide/show controls */%>
	frmScreen.cboTypeOfProperty.onclick();

	<% //MAR1311 %>
	<% //Peter Edney - 22/02/2006 %>
	frmScreen.cboTypeOfProperty.onchange();
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}

function PopulateCombos()
// Populates all combos on the screen
{
	var blnSuccess = true;
	if (m_sTypeOfApplication == "F")
	{	
		//	AW	02/12/2002	BM0116
		// EP2_18 re-enabled and add Date
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtDiscount");
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtPreEmptionDate");
		scScreenFunctions.SetFieldToDisabled(frmScreen,"optFamilySaleYes");
		scScreenFunctions.SetFieldToDisabled(frmScreen,"optFamilySaleNo");
	}	

	<%/* Retrieve the combo lists */%>
	<%//ComboXML = new scXMLFunctions.XMLObject();%>
	ComboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("PropertyType","PropertyDescription","PropertyLocation",
			"ValuationType");

	if(ComboXML.GetComboLists(document,sGroupList))
	{
		// Populate Country combo
		blnSuccess = blnSuccess & ComboXML.PopulateCombo(document,frmScreen.cboTypeOfProperty,"PropertyType",true);
		blnSuccess = blnSuccess & ComboXML.PopulateCombo(document,frmScreen.cboPropertyDescription,"PropertyDescription",true);
		blnSuccess = blnSuccess & ComboXML.PopulateCombo(document,frmScreen.cboPropertyLocation,"PropertyLocation",true);
		blnSuccess = blnSuccess & ComboXML.PopulateCombo(document,frmScreen.cboValuationType,"ValuationType",true);
		<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
		// AQR SYS4737 DRC - Remove 'No Valuation' Valuation type
		<% /* BS BM0187 17/02/03 
        if (m_sTypeOfApplication != "F")
        { 
			var iCount = 0;
			for (iCount = frmScreen.cboValuationType.length - 1; iCount > 0; iCount--)
			{
				if (scScreenFunctions.IsOptionValidationType(frmScreen.cboValuationType, iCount, "N") )
					frmScreen.cboValuationType.remove(iCount);
			}		
		}
		BS BM0187 End */ %>
		if(!blnSuccess) scScreenFunctions.SetScreenToReadOnly(frmScreen);
	}
}

function PopulatePropertyDescriptionCombo()
{
	var sCurrentValue = frmScreen.cboPropertyDescription.value;
	
	<% /* MO - 07/11/2002 - BMIDS00815 - Made change so that the correct property type is found - Start */ %>
	var sPropertyType;
	<% /* var sPropertyType = scScreenFunctions.GetComboValidationType(frmScreen,"cboTypeOfProperty"); */ %>
	var sPropertyTypeIndex = frmScreen.cboTypeOfProperty.selectedIndex;
	
	<% /* Find out whether the property type is a H,B,F or M */ %>
	if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfProperty, sPropertyTypeIndex, "H")) {
		sPropertyType = "H";
	} else {
		if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfProperty, sPropertyTypeIndex, "B")) {
			sPropertyType = "B";
		} else {
			if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfProperty, sPropertyTypeIndex, "F")) {
				sPropertyType = "F";
			} else {
				if (scScreenFunctions.IsOptionValidationType(frmScreen.cboTypeOfProperty, sPropertyTypeIndex, "M")) {
					sPropertyType = "M";
				}
			}
		}
	}
	
	<% /* var bRemoveNonHAndBs = ((sPropertyType == "H") | (sPropertyType == "B"));
	var bRemoveHAndBs = !bRemoveNonHAndBs; */ %>

	<% /* SR 14/07/2004 : BMIDS796  */ %>
	<%/* Populate the combo initially with all items */%>
	for (iCount = frmScreen.cboPropertyDescription.length - 1; iCount >= 0; iCount--)
	{
		frmScreen.cboPropertyDescription.remove(iCount);
	}
	<% /* SR 14/07/2004 : BMIDS796 - End  */ %>
	ComboXML.PopulateCombo(document,frmScreen.cboPropertyDescription,"PropertyDescription",true);		

	<%/* MO - THIS HAS BEEN REMOVED!  Either remove all house/bungalow items from the combo or all NON-house/bungalow items, depending
		 upon the value of bRemoveHAndBs */%>
	<% /* optOption = null; */ %>
	var iCount = 0;
	<% /* var bIsHouseOrBungalow = false; */ %>
	for (iCount = frmScreen.cboPropertyDescription.length - 1; iCount > 0; iCount--)
	{
		<% /* optOption = frmScreen.cboPropertyDescription.item(iCount); 
		bIsHouseOrBungalow = (scScreenFunctions.IsOptionValidationType(frmScreen.cboPropertyDescription, iCount, "H") |
							  scScreenFunctions.IsOptionValidationType(frmScreen.cboPropertyDescription, iCount, "B"))

		if ((bRemoveHAndBs & bIsHouseOrBungalow) |
		    (bRemoveNonHAndBs & !bIsHouseOrBungalow))
			frmScreen.cboPropertyDescription.remove(iCount); */ %>
		
		<% /* MO - if this validation type isnt in this property description remove it */ %>
		if (scScreenFunctions.IsOptionValidationType(frmScreen.cboPropertyDescription, iCount, sPropertyType) == false) {
			frmScreen.cboPropertyDescription.remove(iCount);
		}
		
	}
	
	<% /* MO - 07/11/2002 - BMIDS00815 - End */ %>
	
	frmScreen.cboPropertyDescription.value = sCurrentValue;
	if (frmScreen.cboPropertyDescription.value != sCurrentValue)
		frmScreen.cboPropertyDescription.selectedIndex = 0;
	
}

function PopulateScreen()
// Populates the screen with details of the item selected in DC195
{
	<%//LoanPropertyXML = new scXMLFunctions.XMLObject();%>
	LoanPropertyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();

	LoanPropertyXML.CreateRequestTag(window,null);
	LoanPropertyXML.CreateActiveTag("LOANPROPERTYDETAILS");
	LoanPropertyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	LoanPropertyXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	LoanPropertyXML.RunASP(document,"GetLoanPropertyDetails.asp");

	var ErrorTypes = new Array("RECORDNOTFOUND");
	var LoanPropertyErrorReturn = LoanPropertyXML.CheckResponse(ErrorTypes);

	if(LoanPropertyErrorReturn[1] == ErrorTypes[0])
	{
		<%/* Record not found */%>
		m_sMetaAction = "Add";
		scScreenFunctions.SetRadioGroupValue(frmScreen,"SharedOwnershipCaseGroup","0");
	}
	else if (LoanPropertyErrorReturn[0] == true)
		<%/* Record found */%>
	{
		m_sMetaAction = "Edit";

		<%/* Populate the screen with the details held in the XML */%>
		LoanPropertyXML.SelectTag(null, "LOANPROPERTYDETAILS")
		with (frmScreen)
		{
			cboPropertyDescription.value = LoanPropertyXML.GetTagText("DESCRIPTIONOFPROPERTY");
			cboPropertyLocation.value = LoanPropertyXML.GetTagText("PROPERTYLOCATION");
			cboTypeOfProperty.value = LoanPropertyXML.GetTagText("TYPEOFPROPERTY");
			cboValuationType.value = LoanPropertyXML.GetTagText("VALUATIONTYPE");
			//	AW	02/12/2002	BM0116
			// EP2_18 Re-enable Discount and add date.
			txtDiscount.value = LoanPropertyXML.GetTagText("DISCOUNTAMOUNT");
			txtPreEmptionDate.value = LoanPropertyXML.GetTagText("PREEMPTIONENDDATE");
			txtSharedAmount.value = LoanPropertyXML.GetTagText("SHAREDAMOUNT");
			//txtSharedAmount.value = LoanPropertyXML.GetTagText("SHAREDPERCENTAGE");
			txtPurchasePrice.value = LoanPropertyXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE");
			txtTypeOfApplication.value = m_sTypeOfApplicationDesc;
			scScreenFunctions.SetRadioGroupValue(frmScreen,"SolicitorGroup",LoanPropertyXML.GetTagText("SOLICITORINDICATOR"));
			// EP2_18
			var bFamilyOwnership = LoanPropertyXML.GetTagText("FAMILYSALEINDICATOR");
			if (bFamilyOwnership == "") 
				InitialiseRightToBuy();
			else
				scScreenFunctions.SetRadioGroupValue(frmScreen,"FamilySaleGroup",bFamilyOwnership);
			scScreenFunctions.SetRadioGroupValue(frmScreen,"SharedOwnershipCaseGroup",LoanPropertyXML.GetTagText("SHAREDOWNERSHIPINDICATOR"));
			<% /* MV - 06/06/2002 - BMIDS00032 - Amended New Columns */ %>
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FurtherAdvanceTermGroup",LoanPropertyXML.GetTagText("FURTHERADVANCETERMIND"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FurtherAdvanceRepTypeGroup",LoanPropertyXML.GetTagText("FURTHERADVANCEREPTYPEIND"));
			<%/*MAH 24/11/2006 E2CR35*/%>
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FullMarketValueGroup",LoanPropertyXML.GetTagText("FULLMARKETVALUE"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FinancialIncentivesGroup",LoanPropertyXML.GetTagText("FINANCIALINCENTIVES"));
			<%/*LDM 15/02/2007 EP2_1165*/%>
			scScreenFunctions.SetRadioGroupValue(frmScreen,"DirectFinancialBenefitGroup",LoanPropertyXML.GetTagText("DIRECTFINANCIALBENEFITIND"));
		}
	}
	
	ErrorTypes = null;
	LoanPropertyErrorReturn = null;
}
function PopulatePresentValuation()
{
	var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if(m_sTypeOfApplicationDesc == "R" && 
	   TempXML.IsInComboValidationList(document,"ValuationType",frmScreen.cboValuationType.value,["AU"]))
	{
		//HomeTrack valuation. Get the present valuation from the HomeTrack response
		/*** middle tier not implemented yet
		var ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var reqTag = ValuationXML.CreateRequestTag(window , "GetPresentValuation");
	
		ValuationXML.CreateActiveTag("HOMETRACKVALUATION");
		ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		ValuationXML.RunASP(document, "omHT.asp");
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = ValuationXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{
			//record not found - valuation not carried out
		}
		else if(ErrorReturn[0] == true)
		{
			ValuationXML.SelectTag(null, "HOMETRACKVALUATION");
			frmScreen.txtPresentValuation.value = ValuationXML.GetAttribute("VALUATIONRESULTSVALUATIONAMOUNT");
		}
		***/
	}
	else
	{
		//ESurv valuation. Get the present valuation from the latest valuationreport
		var ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var reqTag = ValuationXML.CreateRequestTag(window , "GetValuationReport");
	
		ValuationXML.CreateActiveTag("VALUATION");
		ValuationXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
		ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		ValuationXML.RunASP(document, "omAppProc.asp");
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = ValuationXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{
			//record not found - valuation not carried out
		}
		else if(ErrorReturn[0] == true)
		{
			ValuationXML.SelectTag(null, "GETVALUATIONREPORT");
			frmScreen.txtPresentValuation.value = ValuationXML.GetAttribute("PRESENTVALUATION");
		}		
	}
	frmScreen.txtPresentValuation.disabled = true;
}
function RetrieveContextData()
{
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","1325");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	m_sTypeOfApplicationDesc = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationDescription","");
	var sTypeOfApplication = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","");

	<%//var TempXML = new scXMLFunctions.XMLObject();%>
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if (TempXML.IsInComboValidationList(document,"TypeOfMortgage",sTypeOfApplication,["N"]))
		m_sTypeOfApplication = "N";
	else if (TempXML.IsInComboValidationXML(["R"]))
		m_sTypeOfApplication = "R";
	else if (TempXML.IsInComboValidationXML(["F"]))
		m_sTypeOfApplication = "F";
	else if (TempXML.IsInComboValidationXML(["T"]))
		m_sTypeOfApplication = "T";
	TempXML = null;
		
}

function SaveLoanPropertyDetails()
{
	var bSuccess = true;
	<%//var XML = new scXMLFunctions.XMLObject();%>
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null)
	//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD ROLLBACK 11/05/01 SYS2050
	// Add OPERATION TJ_03/04/2001 SYS2050
	XML.SetAttribute("OPERATION","CriticalDataCheck");
	XML.CreateActiveTag("CUSTOMER");
	// End OPERATION
	// END ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308 GD ROLLBACK 11/05/01 SYS2050 
	var tagLoanPropertyDetails = XML.CreateActiveTag("LOANPROPERTYDETAILS");
	<%/* NEWPROPERTY */%>
	XML.CreateActiveTag("NEWPROPERTY");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("DESCRIPTIONOFPROPERTY", frmScreen.cboPropertyDescription.value);
	XML.CreateTag("PROPERTYLOCATION", frmScreen.cboPropertyLocation.value);
	XML.CreateTag("TYPEOFPROPERTY", frmScreen.cboTypeOfProperty.value);
	XML.CreateTag("VALUATIONTYPE", frmScreen.cboValuationType.value);
	//	AW	02/12/2002	BM0116
	// EP2_18- re-enable Amount and add date.
	XML.CreateTag("DISCOUNTAMOUNT", frmScreen.txtDiscount.value);
	XML.CreateTag("PREEMPTIONENDDATE", frmScreen.txtPreEmptionDate.value);

	XML.CreateTag("FAMILYSALEINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"FamilySaleGroup"));
	
	<%/*MAH 24/11/2006 E2CR35*/%>
	XML.CreateTag("FULLMARKETVALUE",scScreenFunctions.GetRadioGroupValue(frmScreen,"FullMarketValueGroup"));
	XML.CreateTag("FINANCIALINCENTIVES",scScreenFunctions.GetRadioGroupValue(frmScreen,"FinancialIncentivesGroup"));
	
	// Save the DC202 New property stuff if relevant.
	// Only save our Retype data if still Retype?
	var ValTXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bIsRetype = ValTXML.IsInComboValidationList(document,"ValuationType",frmScreen.cboValuationType.value,Array("RT"));
	if (bIsRetype == true)
	{
		XML.CreateTag("DATEOFORIGINALVALUATION", LoanPropertyXML.GetTagText("DATEOFORIGINALVALUATION"));
		XML.CreateTag("ORIGINALVALUERCOMPANYNAME", LoanPropertyXML.GetTagText("ORIGINALVALUERCOMPANYNAME"));
		XML.CreateTag("ORIGINALVALUERCOMPANYPOSTCODE", LoanPropertyXML.GetTagText("ORIGINALVALUERCOMPANYPOSTCODE"));
		XML.CreateTag("ORIGINALVALUERCOMPANYFLATNONAME", LoanPropertyXML.GetTagText("ORIGINALVALUERCOMPANYFLATNONAME"));
		XML.CreateTag("ORIGINALVALUERCOMPANYBUILDINGORHOUSENUMBER", LoanPropertyXML.GetTagText("ORIGINALVALUERCOMPANYBUILDINGORHOUSENUMBER"));
		XML.CreateTag("ORIGINALVALUERCOMPANYBUILDINGORHOUSENAME", LoanPropertyXML.GetTagText("ORIGINALVALUERCOMPANYBUILDINGORHOUSENAME"));
		XML.CreateTag("ORIGINALVALUERCOMPANYSTREET", LoanPropertyXML.GetTagText("ORIGINALVALUERCOMPANYSTREET"));
		XML.CreateTag("ORIGINALVALUERCOMPANYDISTRICT", LoanPropertyXML.GetTagText("ORIGINALVALUERCOMPANYDISTRICT"));
		XML.CreateTag("ORIGINALVALUERCOMPANYTOWN", LoanPropertyXML.GetTagText("ORIGINALVALUERCOMPANYTOWN"));
		XML.CreateTag("ORIGINALVALUERCOMPANYCOUNTY", LoanPropertyXML.GetTagText("ORIGINALVALUERCOMPANYCOUNTY"));
		XML.CreateTag("ORIGINALVALUERCOMPANYCOUNTRY", LoanPropertyXML.GetTagText("ORIGINALVALUERCOMPANYCOUNTRY"));
	}
	else
	{
		XML.CreateTag("DATEOFORIGINALVALUATION", "");
		XML.CreateTag("ORIGINALVALUERCOMPANYNAME", "");
		XML.CreateTag("ORIGINALVALUERCOMPANYPOSTCODE", "");
		XML.CreateTag("ORIGINALVALUERCOMPANYFLATNONAME", "");
		XML.CreateTag("ORIGINALVALUERCOMPANYBUILDINGORHOUSENUMBER", "");
		XML.CreateTag("ORIGINALVALUERCOMPANYBUILDINGORHOUSENAME", "");
		XML.CreateTag("ORIGINALVALUERCOMPANYSTREET", "");
		XML.CreateTag("ORIGINALVALUERCOMPANYDISTRICT", "");
		XML.CreateTag("ORIGINALVALUERCOMPANYTOWN", "");
		XML.CreateTag("ORIGINALVALUERCOMPANYCOUNTY", "");
		XML.CreateTag("ORIGINALVALUERCOMPANYCOUNTRY", "");
	}
	
	<%/* NEWLOAN */%>
	XML.ActiveTag = tagLoanPropertyDetails;
	XML.CreateActiveTag("NEWLOAN");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("SOLICITORINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"SolicitorGroup"));
	<% /* LDM - 19/02/2006 - EP2_1165 - Amended New Column */ %>
	XML.CreateTag("DIRECTFINANCIALBENEFITIND", scScreenFunctions.GetRadioGroupValue(frmScreen,"DirectFinancialBenefitGroup"));

	<%/* APPLICATIONFACTFIND */%>
	XML.ActiveTag = tagLoanPropertyDetails;
	XML.CreateActiveTag("APPLICATIONFACTFIND");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value);
	<% /* MV - 06/06/2002 - BMIDS00032 - Amended New Columns */ %>
	XML.CreateTag("FURTHERADVANCETERMIND", scScreenFunctions.GetRadioGroupValue(frmScreen,"FurtherAdvanceTermGroup"));
	XML.CreateTag("FURTHERADVANCEREPTYPEIND", scScreenFunctions.GetRadioGroupValue(frmScreen,"FurtherAdvanceRepTypeGroup"));
	

	<%/* SHAREOWNERSHIPDETAILS */%>
	XML.ActiveTag = tagLoanPropertyDetails;
	XML.CreateActiveTag("SHAREDOWNERSHIPDETAILS");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("SHAREDAMOUNT", frmScreen.txtSharedAmount.value);
	XML.CreateTag("SHAREDOWNERSHIPINDICATOR", scScreenFunctions.GetRadioGroupValue(frmScreen,"SharedOwnershipCaseGroup"));
	//XML.CreateTag("SHAREDPERCENTAGE",frmScreen.txtSharedAmount.value);
	
	if (bShared)
		XML.CreateTag("SHAREDOWNERSHIPTYPE", "2");
	else
		XML.CreateTag("SHAREDOWNERSHIPTYPE", "");
		
	// Save the details
	if (m_sMetaAction == "Add")
		//BMIDS01002 Add in call to screen rules
		//XML.RunASP(document,"CreateLoanProperty.asp");
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
				XML.RunASP(document,"CreateLoanProperty.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	else
	{
		//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD ROLLBACK 11/05/01 SYS2050
		// Add CRITICALDATACONTEXT TJ_30/03/2001 AQR SYS2050
		XML.SelectTag(null,"REQUEST");
		XML.CreateActiveTag("CRITICALDATACONTEXT");
		XML.SetAttribute("APPLICATIONNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationNumber",null));
		XML.SetAttribute("APPLICATIONFACTFINDNUMBER",scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber",null));
		XML.SetAttribute("SOURCEAPPLICATION","Omiga");
		XML.SetAttribute("ACTIVITYID",scScreenFunctions.GetContextParameter(window,"idActivityId",null));
		XML.SetAttribute("ACTIVITYINSTANCE","1");
		XML.SetAttribute("STAGEID",scScreenFunctions.GetContextParameter(window,"idStageId",null));
		XML.SetAttribute("APPLICATIONPRIORITY",scScreenFunctions.GetContextParameter(window,"idApplicationPriority",null));
		XML.SetAttribute("COMPONENT","omApp.ApplicationBO");
		XML.SetAttribute("METHOD","UpdateLoanProperty");
		
		window.status = "Critical Data Check - please wait";
		//BMIDS01002 Add in call to screen rules
		//XML.RunASP(document,"OmigaTMBO.asp");
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
		//* XML.RunASP(document,"UpdateLoanProperty.asp");
		//end CRITICALDATACONTEXT  TJ_29/03/2001 
		//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  END GD ROLLBACK 11/05/01 SYS2050
	}	
	bSuccess = XML.IsResponseOK();

	XML = null;
	return(bSuccess);
}


<% /* SYS3481 - Ensure shared amount is always less than the property amount. */ %>
function ValidateSharedAmount()
{
	var bValid = true;
	
	if ((parseInt(frmScreen.txtSharedAmount.value,10) >= parseInt(frmScreen.txtPurchasePrice.value,10)) ||
		((frmScreen.txtPurchasePrice.value == "") && (frmScreen.txtSharedAmount.value != "")))
	{
		alert("Shared amount must be less than Purchase price");		
		bValid = false;
	}
	return(bValid);
}


// EP2_2 - New Methods
function frmScreen.cboValuationType.onclick()
{
	// Retype, or not retype?
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	//DS - EP2_611
	if (frmScreen.cboValuationType.value.length != 0)
		var isRetype = XML.IsInComboValidationList(document,"ValuationType",frmScreen.cboValuationType.value,Array("RT"));
	

	// If Retype button Visible AND Retype then enable button, else Disabled.
	if ( mbRetypeValuationDetails == true && isRetype == true)
		frmScreen.btnRetype.disabled = false;
	else
		frmScreen.btnRetype.disabled = true;
}	

function frmScreen.btnRetype.onclick()
{
	// Passes Valuation data (if present) to DC202.
	ShowRetypeDetailsPopUp();
}

function ShowRetypeDetailsPopUp()
// This function calls pop window for Buy to Let details.
{ 
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = LoanPropertyXML.XMLDocument.xml;
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = m_sApplicationNumber;
	ArrayArguments[3] = m_sApplicationFactFindNumber;
	
	sReturn = scScreenFunctions.DisplayPopup(window, document, "DC202.asp", ArrayArguments, 480, 430);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sRetypeXML = sReturn[1];
		SaveRetypeDetails(LoanPropertyXML);
	}
}

function SaveRetypeDetails(XML)
{
	// Load the DC202 New property stuff if relevant.
	XMLRetype.LoadXML(m_sRetypeXML);
	if(XMLRetype.SelectTag(null, "RETYPE") != null)
	{
		XML.SelectTag(null,"NEWPROPERTY");
		XML.SetTagText("DATEOFORIGINALVALUATION", XMLRetype.GetAttribute("DATEOFORIGINALVALUATION"));
		XML.SetTagText("ORIGINALVALUERCOMPANYNAME", XMLRetype.GetAttribute("ORIGINALVALUERCOMPANYNAME"));
		XML.SetTagText("ORIGINALVALUERCOMPANYPOSTCODE", XMLRetype.GetAttribute("ORIGINALVALUERCOMPANYPOSTCODE"));
		XML.SetTagText("ORIGINALVALUERCOMPANYFLATNONAME", XMLRetype.GetAttribute("ORIGINALVALUERCOMPANYFLATNONAME"));
		XML.SetTagText("ORIGINALVALUERCOMPANYBUILDINGORHOUSENUMBER", XMLRetype.GetAttribute("ORIGINALVALUERCOMPANYBUILDINGORHOUSENUMBER"));
		XML.SetTagText("ORIGINALVALUERCOMPANYBUILDINGORHOUSENAME", XMLRetype.GetAttribute("ORIGINALVALUERCOMPANYBUILDINGORHOUSENAME"));
		XML.SetTagText("ORIGINALVALUERCOMPANYSTREET", XMLRetype.GetAttribute("ORIGINALVALUERCOMPANYSTREET"));
		XML.SetTagText("ORIGINALVALUERCOMPANYDISTRICT", XMLRetype.GetAttribute("ORIGINALVALUERCOMPANYDISTRICT"));
		XML.SetTagText("ORIGINALVALUERCOMPANYTOWN", XMLRetype.GetAttribute("ORIGINALVALUERCOMPANYTOWN"));
		XML.SetTagText("ORIGINALVALUERCOMPANYCOUNTY", XMLRetype.GetAttribute("ORIGINALVALUERCOMPANYCOUNTY"));
		XML.SetTagText("ORIGINALVALUERCOMPANYCOUNTRY", XMLRetype.GetAttribute("ORIGINALVALUERCOMPANYCOUNTRY"));
	}
}


//Changed for EP2_780/EP2_788
// New logic to set values and enabling on new fields.
function InitialiseRightToBuy()
{
	//Get the NatureOfLoan and TypeOfApplication values from the ApplicationFactFind table.
	var iNatureOfLoan; // Clue in the name?
	var iTypeOfApp;    // The TypeofApplication amazingly.
	var iSpecialScheme;
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	XML.CreateActiveTag("APPLICATIONFACTFIND");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.RunASP(document,"GetApplicationFFData.asp");
	if(XML.IsResponseOK())
	{
		XML.SelectTag(null,"APPLICATIONFACTFIND");
		iNatureOfLoan = XML.GetTagText("NATUREOFLOAN");
		iTypeOfApp = XML.GetTagText("TYPEOFAPPLICATION");
		iSpecialScheme = XML.GetTagText("SPECIALSCHEME");
	}
	// Now check whether either is a Right to Buy.
	XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var isRTB = false;
	if (iNatureOfLoan != "")
		isRTB = XML.IsInComboValidationList(document,"NatureOfLoan",iNatureOfLoan,Array("RTB"));
	if ((isRTB == false) && (iTypeOfApp != ""))
		isRTB = XML.IsInComboValidationList(document,"TypeOfMortgage",iTypeOfApp,Array("RTB"));
	if ((isRTB == false) && (iSpecialScheme != ""))
		isRTB = XML.IsInComboValidationList(document,"SpecialSchemes",iSpecialScheme,Array("RTB"));
	
	//If NewProperty.FamilySaleIndicator is not set to anything
	if((scScreenFunctions.GetRadioGroupValue(frmScreen,"FamilySaleGroup") != "1") && (scScreenFunctions.GetRadioGroupValue(frmScreen,"FamilySaleGroup") != "0"))
	{
		// Obtain AccountGUID from Application Table 
		var AccountGUID;
		var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		XML.CreateRequestTag(window,null);
		XML.CreateActiveTag("APPLICATION");
		XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
		XML.RunASP(document,"GetApplicationData.asp");
		if(XML.IsResponseOK())
		{
			XML.SelectTag(null,"APPLICATION");
			AccountGUID = XML.GetTagText("ACCOUNTGUID");
		}
		
  
		// Now for the really cool bit (he lied).
		if(AccountGUID != "")  // AccountGUID found.
		{
			// Get the PreemptionEndDate
			XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			XML.CreateRequestTag(window,null);
			XML.CreateActiveTag("MORTGAGEACCOUNT");
			XML.CreateTag("ACCOUNTGUID", AccountGUID);
			XML.RunASP(document, "GetMortgageAccountDetails.asp");
			if(XML.IsResponseOK())
			{
				XML.SelectTag(null,"MORTGAGEACCOUNT");
				dPreEmptionEndDate = XML.GetTagText("PREEMPTIONENDDATE");
			}
			//If PreemptionEndDate > SystemDate
			if((dPreEmptionEndDate != "") && (scScreenFunctions.CompareDateStringToSystemDate(dPreEmptionEndDate , ">") == true))
			{
				// Set to Yes and disable Options.
				scScreenFunctions.SetRadioGroupValue(frmScreen,"FamilySaleGroup", 1);
				scScreenFunctions.SetRadioGroupToReadOnly(frmScreen,"FamilySaleGroup");
				// Set the Pre-emptive Enddate
				frmScreen.txtPreEmptionDate.value = dPreEmptionEndDate;
				// Enable the discount and PreEmptionEnddate fields.		
				scScreenFunctions.SetFieldToWritable(frmScreen,"txtDiscount");
				scScreenFunctions.SetFieldToWritable(frmScreen,"txtPreEmptionDate");
			}
			else // PreemptionEndDate <= SystemDate or not present.
			{
				// Set to No but DONT disable Options.
				scScreenFunctions.SetRadioGroupValue(frmScreen,"FamilySaleGroup", 0);
				scScreenFunctions.SetRadioGroupToWritable(frmScreen,"FamilySaleGroup");
				// Enable the discount and PreEmptionEnddate fields.		
				scScreenFunctions.SetFieldToDisabled(frmScreen,"txtDiscount");
				scScreenFunctions.SetFieldToDisabled(frmScreen,"txtPreEmptionDate");		
			}
		
		}
		else      // AccountGUID not found.
		{

			// Now, work our magic on the enabling and defaults.
			if (isRTB == true) // Is a Right to Buy
			{
				// Set to Yes and disable Options.
				scScreenFunctions.SetRadioGroupValue(frmScreen,"FamilySaleGroup", 1);
				scScreenFunctions.SetRadioGroupToReadOnly(frmScreen,"FamilySaleGroup");
				// Enable the discount and PreEmptionEnddate fields.		
				scScreenFunctions.SetFieldToWritable(frmScreen,"txtDiscount");
				scScreenFunctions.SetFieldToWritable(frmScreen,"txtPreEmptionDate");
			}
			else // Not a Right to Buy
			{
				// Set to No but DONT disable Options.
				scScreenFunctions.SetRadioGroupValue(frmScreen,"FamilySaleGroup", 0);
				scScreenFunctions.SetRadioGroupToWritable(frmScreen,"FamilySaleGroup");
				// Enable the discount and PreEmptionEnddate fields.		
				scScreenFunctions.SetFieldToDisabled(frmScreen,"txtDiscount");
				scScreenFunctions.SetFieldToDisabled(frmScreen,"txtPreEmptionDate");		
			}
		
		} // AccountGUID not found.
	}
	else
	{
		//If NewProperty.FamilySaleIndicator is set to yes and a RTB
		if((scScreenFunctions.GetRadioGroupValue(frmScreen,"FamilySaleGroup") == "1") && isRTB)
		{
			scScreenFunctions.SetRadioGroupToReadOnly(frmScreen,"FamilySaleRadioGroup");
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtPreEmptionDate");	
			scScreenFunctions.SetFieldToWritable(frmScreen,"txtDiscount");	
		}
	}
}


// EP2_2 - END New Methods

-->
</script>
</body>
</html>


