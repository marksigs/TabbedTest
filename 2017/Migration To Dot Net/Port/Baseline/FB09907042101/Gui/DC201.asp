<%@ LANGUAGE="JSCRIPT" %>
<html>
	<% /*  
Workfile:      DC201.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Property Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		AQR		Description
MF		03/08/2005	MAR20	Created from combination of DC200 & DC220
MF		08/08/2005	MAR20	Parameterised routing for "NewPartySummary" & "ThirdPartySummary" 
JD		23/09/2005	MAR40	default valuation type and set present valuation. Add Leasehold button. Save screen data on btnResidents and btnInsurance
MV		12/10/2005	MAR22	Amended RetreiveContextData() and SaveLoanPropertyDetails() and added frmScreen.txtEarliestRemortgage.onchange()
Maha T	10/10/2005	MAR127  Hide Other Residents button if  answer is not Yes.
MV		18/10/2005	MAR197	Amended retrieveContextData() and SaveLoanPropertyDetails()
MV		19/10/2005	MAR176  Amended ProcessPropertyLocationSelection()
JD		20/10/2005	MAR41	include call to hometrackbo
JD		20/10/2005	MAR263	Get Tag names correct for HometrackBO
JD		08/11/05	MAR299	for remortgage, default valuationtype.
JD		30/11/2005	MAR709  call hometrack through an aspx page to make sure it is called with the correct database logon info.
GHun	01/12/2005	MAR537	Changed SaveLoanPropertyDetails to fix task creation
JD		14/12/2005	MAR783	Only do the call for the Earliest completion date change if we already have an ACA message back from FirstTitle.
PE		22/02/2006	MAR1311 Need to make Property description mandatory
PE		10/03/2006	MAR1061	Changed txtPurchasePrice to read only.
AShaw	28/11/2006	EP2_2		Valuation Report Alignment.
INR		16/01/2007	EP2_677 Spec review changes
INR		13/02/2007	EP2_780 / EP2_788 New Fields
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="STYLESHEET" type="text/css">
	<title></title>
</head>

<body>

<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
<object data=scClientFunctions.asp height=1 id=scClientFunctions 
style="DISPLAY: none; LEFT: 0px; TOP: 0px; VISIBILITY: hidden" tabIndex=-1 
type=text/x-scriptlet width=1 viewastext></object>
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
<form id="frmToGN300" method="post" action="gn300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToCM010" method="post" action="CM010.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC230" method="post" action="DC230.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC295" method="post" action="DC295.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC300" method="post" action="DC300.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC280" method="post" action="DC280.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC210" method="post" action="DC210.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC240" method="post" action="DC240.asp" STYLE="DISPLAY: none"></form>
<form id="frmToDC202" method="post" action="DC202.asp" STYLE="DISPLAY: none"></form>
<% /* Span to keep tabbing within this screen */ %>
<span id="spnToLastField" tabindex="0"></span>

<form id="frmScreen" mark   validate ="onchange">
			<div style="HEIGHT: 400px; LEFT: 10px; POSITION: absolute; TOP: 60px; WIDTH: 604px" class="msgGroup">
				<table style="MARGIN-LEFT: 10px; MARGIN-TOP: 10px;" width="80%" class="msgLabel" border="0">
					<tr>
						<td colspan="4">Type of Application</td>
						<td colspan="7">
							<input id="txtTypeOfApplication" maxlength="45" style="WIDTH: 230px" class="msgReadOnly"
								readonly tabindex="-1">
						</td>
						<td colspan="4">&nbsp;</td>
					</tr>
					<tr height="22">
						<td colspan="4">Type of Property</td>
						<td colspan="7">
							<select id="cboTypeOfProperty" style="WIDTH: 230px" class="msgCombo">
							</select>
						</td>
						<td colspan="4">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="5">Number of flats in the block?</td>
						<td colspan="2">
							<input id="txtNumberOfFlats" maxlength="2" style="WIDTH: 70px;" class="msgTxt" NAME="txtNumberOfFlats">
						</td>
						<td colspan="7">&nbsp;</td>
					</tr>
				</table>
				<table style="MARGIN-LEFT: 10px;" width="80%" class="msgLabel" border="0" ID="Table3">
					<tr>
						<td>Is the flat above commercial premises?</td>
						<td>
							<input id="optFlatAboveCommercialYes" name="FlatAboveCommercialGroup" type="radio" value="1">
						</td>
						<td>
							<label for="optFlatAboveCommercialYes" class="msgLabel">Yes</label>
						</td>
						<td>
							<input id="optFlatAboveCommercialNo" name="FlatAboveCommercialGroup" type="radio" value="0">
						</td>
						<td>
							<label for="optFlatAboveCommercialNo" class="msgLabel">No</label>
						</td>
						<td colspan="9">&nbsp;</td>
						<td colspan="9">&nbsp;</td>
						<td colspan="9">&nbsp;</td>
						<td colspan="9">&nbsp;</td>
						<td colspan="9">&nbsp;</td>
					</tr>
				</table>
				<table style="MARGIN-LEFT: 10px;" width="80%" class="msgLabel" border="0" ID="Table4">
					<tr>
						<td colspan="4">Property Description</td>
						<td colspan="7">
							<select id="cboPropertyDescription" menusafe="true" style="WIDTH: 230px" class="msgCombo">
							</select>
						</td>
						<td colspan="4">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="4">Property Location</td>
						<td colspan="7">
							<select id="cboPropertyLocation" menusafe="true" style="WIDTH: 230px" class="msgCombo">
							</select>
						</td>
						<td colspan="4">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="4">Purchase Price/ Estimated Value</td>
						<td colspan="3">
							<input id="txtPurchasePrice" maxlength="9" style="WIDTH: 100px" class="msgReadOnly">
						</td>
						<td colspan="8">&nbsp;</td>
					</tr>
					<tr id="valuationType" style="display:none">
						<td colspan="4">Valuation Type</td>
						<td colspan="7">
							<select id="cboValuationType" style="WIDTH: 230px" menusafe="true" class="msgCombo">
							</select>
						</td>
						<td colspan="2">
							<input id="btnRetype" value="Retype" type="button" style="WIDTH: 60px" class="msgButton"
								NAME="btnRetype">
						</td>
						<td colspan="1">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="4">Present Valuation</td>
						<td colspan="3">
							<input id="txtPresentValuation" maxlength="9" style="WIDTH: 100px" class="msgReadOnly"
								readonly>
						</td>
						<td colspan="8">&nbsp;</td>
					</tr>
				</table>
				<table style="MARGIN-LEFT: 10px;" width="80%" class="msgLabel" border="0" ID="Table5">
					<tr>
						<td colspan="6">Family Sale/Right To Buy Discount?</td>
						<td colspan="1">&nbsp;</td>
						<td>
							<input id="optFamilySaleYes" name="FamilySaleGroup" type="radio" value="1">
						</td>
						<td>
							<label for="optFamilySaleYes" class="msgLabel">Yes</label>
						</td>
						<td>
							<input id="optFamilySaleNo" name="FamilySaleGroup" type="radio" value="0">
						</td>
						<td>
							<label for="optFamilySaleNo" class="msgLabel">No</label>
						</td>
						<td colspan="1">&nbsp;</td>
						<td>Discount</td>
						<td colspan="2">
							<input id="txtDiscount" maxlength="6" style="WIDTH: 70px;" class="msgTxt" NAME="txtDiscount">
						</td>
						<td colspan="5">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="5">Pre-emption End Date</td>
						<td colspan="2">
							<input id="txtPreEmptionDate" maxlength="10" style="WIDTH: 70px;" class="msgTxt" NAME="txtPreEmptionDate">
						</td>
						<td colspan="8">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="6">Is this a Shared Ownership Case?</td>
						<td colspan="1">&nbsp;</td>
						<td>
							<input id="optSharedOwnershipCaseYes" name="SharedOwnershipCaseGroup" type="radio" value="1"">
						</td>
						<td>
							<label for="optSharedOwnershipCaseYes" class="msgLabel">Yes</label>
						</td>
						<td>
							<input id="optSharedOwnershipCaseNo" name="SharedOwnershipCaseGroup" type="radio" value="0"">
						</td>
						<td>
							<label for="optSharedOwnershipCaseNo" class="msgLabel">No</label>
						</td>
						<td colspan="1">&nbsp;</td>
						<td>Amount</td>
						<td colspan="2">
							<input id="txtSharedAmount" maxlength="6" style="WIDTH: 70px;" class="msgTxt" NAME="txtSharedAmount">
						</td>
						<td colspan="5">&nbsp;</td>
					</tr>
					<% /* MF MARS20 Hide
		<span style="LEFT: 4px; POSITION: absolute; TOP: 156px" class="msgLabel">
			Private Sale?
			<span style="LEFT: 110px; POSITION: relative; TOP: -5px">
				<input id="optFamilySaleYes" name="FamilySaleGroup" type="radio" value="1"><label for="optFamilySaleYes" class="msgLabel">Yes</label> 
			</span> 

			<span style="LEFT: 240px; POSITION: absolute;  TOP: -5px">
				<input id="optFamilySaleNo" name="FamilySaleGroup" type="radio" value="0"><label for="optFamilySaleNo" class="msgLabel">No</label> 
			</span> 	
		</span>
		<span style="LEFT: 4px; POSITION: absolute; TOP: 180px" class="msgLabel">
			Is this a Shared Ownership Case?
			<span style="LEFT: 7px; POSITION: relative; TOP: 0px">
				<input id="optSharedOwnershipCaseYes" name="SharedOwnershipCaseGroup" type="radio" value="1" onclick="SharedOwnershipCaseChanged()"><label for="optSharedOwnershipCaseYes" class="msgLabel">Yes</label> 
			</span> 

			<span style="LEFT: 240px; POSITION: absolute; TOP: 0px">
				<input id="optSharedOwnershipCaseNo" name="SharedOwnershipCaseGroup" type="radio" value="0" onclick="SharedOwnershipCaseChanged()"><label for="optSharedOwnershipCaseNo" class="msgLabel">No</label> 
			</span>		
			<span style="LEFT: 100px; POSITION: relative; TOP: 0px" class="msgLabel">
			Amount
				<span style="LEFT: 60px; POSITION: absolute; TOP: 0px">
					<input id="txtSharedAmount" maxlength="6" style="POSITION: absolute; WIDTH: 100px" class="msgTxt">
				</span> 
			</span>		 
		</span>
		*/ %>
					<tr>
						<td colspan="4">Tenure of Property</td>
						<td colspan="7">
							<select id="cboTenureOfProperty" style="WIDTH: 230px" menusafe="true" class="msgCombo" NAME="cboTenureOfProperty">
							</select>
						</td>
						<td colspan="4">
							<input id="btnLeaseholdDetails" value="Leasehold Details" type="button" style="WIDTH: 105px"
								class="msgButton" NAME="btnLeaseholdDetails">
						</td>
					</tr>
					<tr>
						<td colspan="6">New Property?</td>
						<td colspan="1">&nbsp;</td>
						<td>
							<input id="optNewPropertyYes" name="NewPropertyRadioGroup" type="radio" value="1">
						</td>
						<td>
							<label for="optNewPropertyYes" class="msgLabel">Yes</label>
						</td>
						<td>
							<input id="optNewPropertyNo" name="NewPropertyRadioGroup" type="radio" value="0">
						</td>
						<td>
							<label for="optNewPropertyNo" class="msgLabel">No</label>
						</td>
						<td colspan="8">&nbsp;</td>
					</tr>
				</table>
				<table style="MARGIN-LEFT: 10px;" width="80%" class="msgLabel" border="0" ID="Table2">

					<tr>
						<td colspan="4">House Builder's Guarantee</td>
						<td colspan="5">
							<select id="cboHouseBuildersGuarantee" style="WIDTH: 146px" class="msgCombo" NAME="cboHouseBuildersGuarantee"
								menusafe="true">
							</select>
						</td>
						<td colspan="7">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="4">Date Of Entry / Target Date for Exchange</td>
						<td colspan="2">
							<input id="txtDateOfEntry" maxlength="10" style="WIDTH: 70px;" class="msgTxt" NAME="txtDateOfEntry">
						</td>
						<!--<td colspan="1">&nbsp;</td>-->

					</tr>
					<tr>
						<td colspan="4">Number of Bedrooms</td>
						<td colspan="2">
							<input id="txtNumberOfBedrooms" maxlength="2" style="WIDTH: 70px;" class="msgTxt" NAME="txtNumberOfBedrooms">
						</td>
						<td colspan="8">&nbsp;</td>
					</tr>
				</table>
				<table style="MARGIN-LEFT: 10px;" width="80%" class="msgLabel" border="0" ID="Table1">
					<tr id="otherResidents" style="display:none">
						<td colspan="6">Are there any Other Residents?</td>
						<td>
							<input id="optOtherResidentsYes" name="OtherResidentsGroup" type="radio" value="1">
						</td>
						<td>
							<label for="optOtherResidentsYes" class="msgLabel">Yes</label>
						</td>
						<td>
							<input id="optOtherResidentsNo" name="OtherResidentsGroup" type="radio" value="0">
						</td>
						<td>
							<label for="optOtherResidentsNo" class="msgLabel">No</label>
						</td>
						<td colspan="1">&nbsp;</td>
						<td colspan="3">
							<input id="btnOtherResidents" value="Other Residents" type="button" style="WIDTH: 105px"
								class="msgButton" NAME="btnOtherResidents">
						</td>
						<td colspan="3">&nbsp;</td>
					</tr>
					<tr id="earliestCompletion" style="display:none">
						<td colspan="2">Re-mortgage earliest completion date</td>
						<td colspan="2">
							<input id="txtEarliestRemortgage" maxlength="10" style="WIDTH: 70px; " class="msgTxt" NAME="txtEarliestRemortgage">
						</td>
						<td colspan="9">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="6">Is the customer arranging their own buildings and contents 
							insurance with another organisation?</td>
						<td>
							<input id="optOwnBuilConYes" name="OwnBuilConRadioGroup" type="radio" value="1">
						</td>
						<td>
							<label for="optOwnBuilConYes" class="msgLabel">Yes</label>
						</td>
						<td>
							<input id="optOwnBuilConNo" name="OwnBuilConRadioGroup" type="radio" value="0">
						</td>
						<td>
							<label for="optOwnBuilConNo" class="msgLabel">No</label>
						</td>
						<td colspan="1">&nbsp;</td>
						<td colspan="3">
							<input id="btnCoverDetails" value="Cover" type="button" style="WIDTH: 105px" class="msgButton"
								NAME="btnCoverDetails">
						</td>
						<td colspan="3">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="6">Is the property being purchased at full market value?</td>
						<td>
							<input id="optFullMarketValueYes" name="FullMarketValueGroup" type="radio" value="1">
						</td>
						<td>
							<label for="optFullMarketValueYes" class="msgLabel">Yes</label>
						</td>
						<td>
							<input id="optFullMarketValueNo" name="FullMarketValueGroup" type="radio" value="0">
						</td>
						<td>
							<label for="optFullMarketValueNo" class="msgLabel">No</label>
						</td>
						<td colspan="1">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="6">Are there any other financial incentives being offered?</td>
						<td>
							<input id="optFinancialIncentivesYes" name="FinancialIncentivesGroup" type="radio" value="1">
						</td>
						<td>
							<label for="optFinancialIncentivesYes" class="msgLabel">Yes</label>
						</td>
						<td>
							<input id="optFinancialIncentivesNo" name="FinancialIncentivesGroup" type="radio" value="0">
						</td>
						<td>
							<label for="optFinancialIncentivesNo" class="msgLabel">No</label>
						</td>
						<td colspan="1">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="6">Is the property affected by an agricultural restriction/covenant?
						</td>
						<td>
							<input id="optAgriculturalRestrictionYes" name="AgriculturalRestrictionGroup" type="radio"
								value="1">
						</td>
						<td>
							<label for="optAgriculturalRestrictionYes" class="msgLabel">Yes</label>
						</td>
						<td>
							<input id="optAgriculturalRestrictionNo" name="AgriculturalRestrictionGroup" type="radio"
								value="0">
						</td>
						<td>
							<label for="optAgriculturalRestrictionNo" class="msgLabel">No</label>
						</td>
					</tr>
				</table>
			</div>
<!--			<div id="divFurtherAdvance" style="HEIGHT: 136px; LEFT: 10px; POSITION: absolute; TOP: 564px; WIDTH: 604px"-->

			<div id="divFurtherAdvance" style="HEIGHT: 136px; LEFT: 10px; POSITION: absolute; TOP: 600px; WIDTH: 604px"
				class="msgGroup">
				<span style="LEFT: 12px; POSITION: absolute; TOP: 10px" class="msgLabel">
					<strong>Further Advance</strong>
				</span>
				<span style="LEFT: 12px; POSITION: absolute; TOP: 36px" class="msgLabel">
		Is a Solicitor required?
		<span style="LEFT: 270px; POSITION: absolute; TOP: -3px">
						<input id="optSolicitorYes" name="SolicitorGroup" type="radio" value="1"><label for="optSolicitorYes" class="msgLabel">Yes</label>
					</span> 

		<span style="LEFT: 320px; POSITION: absolute; TOP: -3px">
						<input id="optSolicitorNo" name="SolicitorGroup" type="radio" value="0"><label for="optSolicitorNo" class="msgLabel">No</label>
					</span> 
	</span>
				<span style="LEFT: 12px; POSITION: absolute; TOP: 60px" class="msgLabel">
		Does Futher Advance fit within exisiting term?
		<span style="LEFT: 270px; POSITION: absolute; TOP: -3px">
						<input id="optFurtherAdvanceTermYes" name="FurtherAdvanceTermGroup" type="radio" value="1"><label for="optFurtherAdvanceTermYes" class="msgLabel">Yes</label>
					</span> 

		<span style="LEFT: 320px; POSITION: absolute; TOP: -3px">
						<input id="optFurtherAdvanceTermNo" name="FurtherAdvanceTermGroup" type="radio" value="0"><label for="optFurtherAdvanceTermNo" class="msgLabel">No</label>
					</span> 
	</span>
				<span style="LEFT: 12px; POSITION: absolute; TOP: 84px" class="msgLabel">
		Further Advance Repayment Type to be<BR>the same as the exisiting loan?
		<span style="LEFT: 270px; POSITION: absolute; TOP: -3px">
						<input id="optFurtherAdvanceRepTypeYes" name="FurtherAdvanceRepTypeGroup" type="radio"
							value="1"><label for="optFurtherAdvanceRepTypeYes" class="msgLabel">Yes</label>
					</span> 

		<span style="LEFT: 320px; POSITION: absolute; TOP: -3px">
						<input id="optFurtherAdvanceRepTypeNo" name="FurtherAdvanceRepTypeGroup" type="radio" value="0"><label for="optFurtherAdvanceRepTypeNo" class="msgLabel">No</label>
					</span> 
	</span>
</div>

</form>

<div id="msgButtons" style="HEIGHT: 19px; LEFT: 8px; POSITION: absolute; TOP: 720px; WIDTH: 612px"><!-- #include FILE="msgButtons.asp" --> 
</div> 

<% /* Span to keep tabbing within this screen */ %>
<span id="spnToFirstField" tabindex="0"></span><!-- #include FILE="fw030.asp" --><!-- #include FILE="attribs/DC201attribs.asp" -->

<script language="JScript">
<!--
var m_sMetaAction = "";
var m_sReadOnly = "";
var m_sApplicationNumber = "";
var m_sApplicationFactFindNumber = "";
var m_sTypeOfApplicationDesc = "";
var m_sTypeOfApplication = "";
var m_sLeaseholdXML = null;
var m_sLeaseholdHoldingXML = null;
var LoanPropertyXML = null;
var ComboXML = null;
var scScreenFunctions;
var bShared = false;
var m_blnReadOnly = false;
var m_sCurrency = "";
var bEarlierRemortgageDateChanged = false;
var m_sTypeOfApplicationValue ="";
var mbRetypeValuationDetails = false;
var XMLRetype = new top.frames[1].document.all.scXMLFunctions.XMLObject(); // EP2_2 AShaw 28/11/2006
var m_sRetypeXML = "";  //EP2_2 Retype Valuation XML


/* EVENTS - BUTTONS */

function ProcessPropertyLocationSelection()
{
	if (m_sTypeOfApplication == "R")
		document.all.item("earliestCompletion").style.display="block";
	else
		document.all.item("earliestCompletion").style.display="none";
		
	<% /* var bShowCompletionDate = (m_sTypeOfApplication == "R");
	if (!bShowCompletionDate){
		var sPropertyLocationIndex = frmScreen.cboPropertyLocation.selectedIndex;			
		// Find out whether the property type is Scotland 'S'
		if (scScreenFunctions.IsOptionValidationType(frmScreen.cboPropertyLocation, sPropertyLocationIndex, "S")) 
			bShowCompletionDate = true;
		else			
			bShowCompletionDate = false;
	}
	if(bShowCompletionDate)
		document.all.item("earliestCompletion").style.display="block";
	else
		document.all.item("earliestCompletion").style.display="none"; */ %>
}

function frmScreen.cboPropertyLocation.onclick()
{
	ProcessPropertyLocationSelection();
}

function btnCancel.onclick()
{
	if(scScreenFunctions.CompletenessCheckRouting(window))
		frmToGN300.submit();
	else
	{
		scScreenFunctions.SetContextParameter(window,"idMetaAction", "");		
		frmToDC210.submit();
	}
}

function btnSubmit.onclick()
{	
	// EP2_2 - Warn if losing Retype data.
	var ValTXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bIsRetype = ValTXML.IsInComboValidationList(document,"ValuationType",frmScreen.cboValuationType.value,Array("RT"));
	LoanPropertyXML.SelectTag(null,"NEWPROPERTY");
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


	if (CommitChanges())
	{
		if(scScreenFunctions.CompletenessCheckRouting(window))
		{
			scScreenFunctions.SetContextParameter(window,"idReturnScreenId","MN060");
			frmToGN300.submit();			
		}
		else
		{
			<% /*
			//frmToMN060.submit();
			//*[MC]BMIDS772 - Routing return screen set to DC200
			//scScreenFunctions.SetContextParameter(window,"idReturnScreenId","DC200");
			//frmToCM010.submit();
			//GoToCostModelling(); //JD BMIDS772 Carry out stage checks and global parameter changes
			*/%>
			<% /* read */ %>
			<% /* MF Read Global Parameter ThirdPartySummary to decide route */ %>
			var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
			var bThirdPartySummary = XML.GetGlobalParameterBoolean(document,"ThirdPartySummary");	
			if(bThirdPartySummary)			
				frmToDC240.submit();
			else
				frmToDC280.submit();
		}
	}
}
function frmScreen.btnLeaseholdDetails.onclick()
{
	var sReturn = null;
	var ArrayArguments = new Array();
	ArrayArguments[0] = m_sLeaseholdXML;
	ArrayArguments[1] = m_sReadOnly;
	ArrayArguments[2] = m_sApplicationNumber;
	ArrayArguments[3] = m_sApplicationFactFindNumber;
	ArrayArguments[4] = m_sCurrency;

	sReturn = scScreenFunctions.DisplayPopup(window, document, "DC222.asp", ArrayArguments, 340, 369);
	if (sReturn != null)
	{
		FlagChange(sReturn[0]);
		m_sLeaseholdXML = sReturn[1];
	}
}
function frmScreen.cboTenureOfProperty.onchange()
{
	DoButtonEnabling();
}
function DoButtonEnabling()
{
	if (scScreenFunctions.IsValidationType(frmScreen.cboTenureOfProperty, "L"))
		frmScreen.btnLeaseholdDetails.disabled = false;
	else
		frmScreen.btnLeaseholdDetails.disabled = true;	
}
function GoToCostModelling()
{
	// Carry out stage checks and global parameter changes before going to CM010
	if(CheckStage())
	{
		scScreenFunctions.SetContextParameter(window,"idApplicationMode","Cost Modelling");
		scScreenFunctions.SetContextParameter(window,"idReturnScreenId","DC201");
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

<% /* EVENTS - CLICKS

<% /* EP2_780 */ %>
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
function frmScreen.optOtherResidentsNo.onclick()
{
	frmScreen.btnOtherResidents.disabled=true;
}

function frmScreen.optOtherResidentsYes.onclick()
{
	frmScreen.btnOtherResidents.disabled=false;
}

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

function frmScreen.btnOtherResidents.onclick()
{
	//JD MAR40 save screen data before leaving the screen
	SaveLoanPropertyDetails(false);
	frmToDC230.submit();
}

<% /* Start: MAR127	Maha T - Show/Hide other residents button */ %>
function ChkOtherResidentsGroup()
{
	if (frmScreen.optOtherResidentsYes.checked)
	{
		frmScreen.optOtherResidentsYes.onclick()
	}
	else
	{
		frmScreen.optOtherResidentsNo.onclick()
	}
}
<% /* End: MAR127 */ %>

function frmScreen.btnCoverDetails.onclick()
{
	//JD MAR40 save screen data before leaving the screen
	SaveLoanPropertyDetails(false);
		
	if(frmScreen.optOwnBuilConNo.checked){		
		frmToDC300.submit();
	} else {						
		frmToDC295.submit();
	}
}
<% /* EP2_780 */ %>
function frmScreen.optFamilySaleNo.onclick()
{
	if (frmScreen.optFamilySaleNo.checked) 
	{
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtDiscount");
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtPreEmptionDate");
	}
}	

function frmScreen.optFamilySaleYes.onclick()
{
	if (frmScreen.optFamilySaleYes.checked) 
	{
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtDiscount");
		scScreenFunctions.SetFieldToWritable(frmScreen,"txtPreEmptionDate");
	}
}

<% /* EVENTS - BLURS */ %>

function frmScreen.txtSharedAmount.onblur()
{	
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
	scScreenFunctions.btnsubmit
	SetFocusToLastField(frmScreen);
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
	FW030SetTitles("Property Details","DC201",scScreenFunctions);
	
	RetrieveContextData();
	
	SetMasks();
	
	Initialise();
	Validation_Init();	
	

	scScreenFunctions.SetFocusToFirstField(frmScreen);
		
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	
	EnableRadios();	
	
	m_blnReadOnly = scScreenFunctions.SetMainScreenToReadOnly(window, frmScreen, "DC201");	
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
	
	<% /* EP2_780 */ %>
	if(scScreenFunctions.GetRadioGroupValue(frmScreen,"FamilySaleGroup") != "1")
	{
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtDiscount");
		scScreenFunctions.SetFieldToReadOnly(frmScreen,"txtPreEmptionDate");
	}
	else
	// if (m_sTypeOfApplication != "F")
	//	scScreenFunctions.SetFieldToWritable(frmScreen,"txtDiscount");

	if (m_sTypeOfApplication == "F")
	{	
		//	AW	02/12/2002	BM0116
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtDiscount");
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtPreEmptionDate");		
		//scScreenFunctions.SetFieldToDisabled(frmScreen,"optFamilySaleYes");
		//scScreenFunctions.SetFieldToDisabled(frmScreen,"optFamilySaleNo");
        scScreenFunctions.SetRadioGroupToDisabled(frmScreen,"FamilySaleGroup");
        
        <% /*BMIDS673 Always enable Shared Ownership field
		scScreenFunctions.SetRadioGroupToDisabled(frmScreen,"SharedOwnershipCaseGroup"); 
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtSharedAmount"); */ %>
	}		
	
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
			<% /* EP2_780 */ %>
			txtDiscount.value = "";
			txtPreEmptionDate.value = "";
	
			txtPurchasePrice.value = "";
		}
	}
}
function DefaultValuationType()
{
	//choose valuation type based on type of mortgage
	if(m_sTypeOfApplication == "R")
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
				bSuccess = SaveLoanPropertyDetails(true);
		}
		else
			bSuccess = false;
	return(bSuccess);
}

function DefaultFields()
// Inserts default values into all fields
{
	ClearFields(true,true);
	//frmScreen.optFamilySaleNo.onclick();
	<% /* MF 08/09/2005 MAR20 */ %>
	frmScreen.optOwnBuilConYes.checked = true;
	
	<% /* Maha T	MAR127  Hide other residents button by default */ %>
	frmScreen.btnOtherResidents.disabled=true;
	<% /* EP2_780 */ %>
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
	else {
		scScreenFunctions.HideCollection(divFurtherAdvance);
		<% /* MF MARS20 move up buttons */ %>		
		msgButtons.style.top = "670px";
	}
	<% /* MF MARS20 hide Other residents radios & button hidden
			 when not a new loan */ %>			
	if (m_sTypeOfApplication == "N"){		
		document.all.item("otherResidents").style.display="block";		
	}
	if (m_sTypeOfApplication != "R"){							
		document.all.item("valuationType").style.display="block";		
		document.all.item("cboValuationType").setAttribute("required", "true");	
	}
	<% /* MF Hide Earliest completion date when new loan and property location 
			is not scotland */ %>
	ProcessPropertyLocationSelection();
	
	<% /* JD MAR40 set present valuation based on valuationtype*/%>
	PopulatePresentValuation();
	DoButtonEnabling();
	
	<% /* Maha T	MAR127	Show/Hide other residents button */ %>
	ChkOtherResidentsGroup();
	
	<%/* Fire off events to hide/show controls */%>
	frmScreen.cboTypeOfProperty.onclick();

	<% //MAR1311 %>
	<% //Peter Edney - 22/02/2006 %>
	frmScreen.cboTypeOfProperty.onchange();
	
	scScreenFunctions.SetFocusToFirstField(frmScreen);
}
function PopulatePresentValuation()
{
	var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if(m_sTypeOfApplication == "R" && 
	   TempXML.IsInComboValidationList(document,"ValuationType",frmScreen.cboValuationType.value,["AU"]))
	{
		//HomeTrack valuation. Get the present valuation from the HomeTrack response
		var ValuationXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
		var reqTag = ValuationXML.CreateRequestTag(window , "GetPresentValuation");
	
		ValuationXML.CreateActiveTag("HOMETRACKVALUATION");
		ValuationXML.SetAttribute("APPLICATIONNUMBER", m_sApplicationNumber);
		ValuationXML.SetAttribute("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
		//ValuationXML.RunASP(document, "omHT.asp"); //MAR709
		ValuationXML.RunASP(document, "GetPresentValuation.aspx");  //MAR709
		var ErrorTypes = new Array("RECORDNOTFOUND");
		var ErrorReturn = ValuationXML.CheckResponse(ErrorTypes);
		if (ErrorReturn[1] == ErrorTypes[0])
		{
			//record not found - valuation not carried out
		}
		else if(ErrorReturn[0] == true)
		{
			ValuationXML.SelectTag(null, "HOMETRACKVALUATION");
			frmScreen.txtPresentValuation.value = ValuationXML.GetAttribute("VALUATIONAMOUNT");
		}
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
function PopulateCombos()
// Populates all combos on the screen
{
	var blnSuccess = true;
	if (m_sTypeOfApplication == "F")
	{	
		//	EP2_780/EP2_788
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtDiscount");
		scScreenFunctions.SetFieldToDisabled(frmScreen,"txtPreEmptionDate");
		scScreenFunctions.SetFieldToDisabled(frmScreen,"optFamilySaleYes");
		scScreenFunctions.SetFieldToDisabled(frmScreen,"optFamilySaleNo");
	}	

	<%/* Retrieve the combo lists */%>
	<%//ComboXML = new scXMLFunctions.XMLObject();%>
	ComboXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var sGroupList = new Array("PropertyType","PropertyDescription","PropertyLocation",
			"ValuationType","PropertyTenure","HouseBuildersGuarantee");

	if(ComboXML.GetComboLists(document,sGroupList))
	{
		// Populate Country combo
		blnSuccess = blnSuccess & ComboXML.PopulateCombo(document,frmScreen.cboTypeOfProperty,"PropertyType",true);
		blnSuccess = blnSuccess & ComboXML.PopulateCombo(document,frmScreen.cboPropertyDescription,"PropertyDescription",true);
		blnSuccess = blnSuccess & ComboXML.PopulateCombo(document,frmScreen.cboPropertyLocation,"PropertyLocation",true);
		blnSuccess = blnSuccess & ComboXML.PopulateCombo(document,frmScreen.cboValuationType,"ValuationType",true);
		blnSuccess = blnSuccess & ComboXML.PopulateCombo(document,frmScreen.cboTenureOfProperty,"PropertyTenure",true);
		blnSuccess = blnSuccess & ComboXML.PopulateCombo(document,frmScreen.cboHouseBuildersGuarantee,"HouseBuildersGuarantee",true);
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
	var ErrorReturn = LoanPropertyXML.CheckResponse(ErrorTypes);

	if(ErrorReturn[1] == ErrorTypes[0])
	{
		<%/* Record not found */%>
		m_sMetaAction = "Add";
		scScreenFunctions.SetRadioGroupValue(frmScreen,"SharedOwnershipCaseGroup","0");
	}
	else if (ErrorReturn[0] == true)
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
			if(LoanPropertyXML.GetTagText("VALUATIONTYPE") != "")
				cboValuationType.value = LoanPropertyXML.GetTagText("VALUATIONTYPE");
			else
			{
				//JD MAR299 if we have no valuation type then if we are a remortgage default to 'AU'
				if(m_sTypeOfApplication == "R")
				{
					scScreenFunctions.SetComboOnValidationType(frmScreen, "cboValuationType", "AU");
				}
			}
			<% /* EP2_780 */ %>
			txtDiscount.value = LoanPropertyXML.GetTagText("DISCOUNTAMOUNT");
			txtSharedAmount.value = LoanPropertyXML.GetTagText("SHAREDAMOUNT");
			txtSharedAmount.value = LoanPropertyXML.GetTagText("SHAREDPERCENTAGE");
			txtPurchasePrice.value = LoanPropertyXML.GetTagText("PURCHASEPRICEORESTIMATEDVALUE");
			txtEarliestRemortgage.value = LoanPropertyXML.GetTagText("EARLIESTCOMPLETIONDATE");
			txtTypeOfApplication.value = m_sTypeOfApplicationDesc;
			scScreenFunctions.SetRadioGroupValue(frmScreen,"SolicitorGroup",LoanPropertyXML.GetTagText("SOLICITORINDICATOR"));
			// EP2_780/EP2_788
			var bFamilyOwnership = LoanPropertyXML.GetTagText("FAMILYSALEINDICATOR");
			if (bFamilyOwnership == "") 
				InitialiseRightToBuy();
			else
			{
				scScreenFunctions.SetRadioGroupValue(frmScreen,"FamilySaleGroup",bFamilyOwnership);
				txtPreEmptionDate.value = LoanPropertyXML.GetTagText("PREEMPTIONENDDATE");
			}
			scScreenFunctions.SetRadioGroupValue(frmScreen,"SharedOwnershipCaseGroup",LoanPropertyXML.GetTagText("SHAREDOWNERSHIPINDICATOR"));
			<% /* MV - 06/06/2002 - BMIDS00032 - Amended New Columns */ %>
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FurtherAdvanceTermGroup",LoanPropertyXML.GetTagText("FURTHERADVANCETERMIND"));
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FurtherAdvanceRepTypeGroup",LoanPropertyXML.GetTagText("FURTHERADVANCEREPTYPEIND"));
			<% /* MF 04/08/2005 MARS20 Read data for new/reinstated controls default to yes */%>
			var sOwnCoverText=LoanPropertyXML.GetTagText("ARRANGEOWNCOVERINDICATOR");			
			if(sOwnCoverText=="")
				frmScreen.optOwnBuilConYes.checked=true;
			else
				scScreenFunctions.SetRadioGroupValue(frmScreen,"OwnBuilConRadioGroup",sOwnCoverText);			
			
			LoadNewPropertyData();			
			
			cboHouseBuildersGuarantee.value = LoanPropertyXML.GetTagText("HOUSEBUILDERSGUARANTEE");
			cboTenureOfProperty.value= LoanPropertyXML.GetTagText("TENURETYPE");
			scScreenFunctions.SetRadioGroupValue(frmScreen,"NewPropertyRadioGroup",LoanPropertyXML.GetTagText("NEWPROPERTYINDICATOR"));	
			scScreenFunctions.SetRadioGroupValue(frmScreen,"OtherResidentsGroup",LoanPropertyXML.GetTagText("ANYOTHERRESIDENTSINDICATOR"));	
			txtDateOfEntry.value = LoanPropertyXML.GetTagText("DATEOFENTRY");
			<% /* EP2_677 */%>
			txtNumberOfFlats.value = LoanPropertyXML.GetTagText("NOFLATSINBLOCK");
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FlatAboveCommercialGroup",LoanPropertyXML.GetTagText("FLATABOVECOMMERCIAL"));	
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FullMarketValueGroup",LoanPropertyXML.GetTagText("FULLMARKETVALUE"));	
			scScreenFunctions.SetRadioGroupValue(frmScreen,"FinancialIncentivesGroup",LoanPropertyXML.GetTagText("FINANCIALINCENTIVES"));	
			scScreenFunctions.SetRadioGroupValue(frmScreen,"AgriculturalRestrictionGroup",LoanPropertyXML.GetTagText("AGRICULTURALRESTRICTIONS"));	
			
		}
	}

	ErrorTypes = null;
	ErrorReturn = null;
}

<% /* MF Load the XML for NewProperty data which contains data for the number of rooms */ %>
function LoadNewPropertyData()
{
	var NewPropertyXML = null;
	NewPropertyXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	NewPropertyXML.CreateRequestTag(window, "SEARCH");

	NewPropertyXML.CreateActiveTag("NEWPROPERTY");
	NewPropertyXML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	NewPropertyXML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);

	NewPropertyXML.RunASP(document,"GetNewPropertyDescription.asp");
	var ErrorTypes = new Array("RECORDNOTFOUND");
	var ErrorReturn = NewPropertyXML.CheckResponse(ErrorTypes);

	if (ErrorReturn[0] == true)
	{			
		if (NewPropertyXML.SelectTag(null,"NEWPROPERTYROOMTYPELIST") != null){			
			var tagList = NewPropertyXML.CreateTagList("NEWPROPERTYROOMTYPE");			
			for (var nItem = 0; nItem < tagList.length; nItem++)
			{
				NewPropertyXML.SelectTagListItem(nItem);
				if(NewPropertyXML.GetTagText("ROOMTYPE")=="2"){
					frmScreen.txtNumberOfBedrooms.value=NewPropertyXML.GetTagText("NUMBEROFROOMS");
				}				
			}	
		}
		if (NewPropertyXML.SelectTag(null,"NEWPROPERTYLEASEHOLD") != null)
		{	
			m_sLeaseholdHoldingXML = NewPropertyXML.ActiveTag.xml; // BMIDS00140 - assign to 2nd variable to fix bug
			m_sLeaseholdXML = NewPropertyXML.ActiveTag.xml;
		}
	}
}

function RetrieveContextData()
{
	m_sReadOnly = scScreenFunctions.GetContextParameter(window,"idReadOnly","0");
	m_sApplicationNumber = scScreenFunctions.GetContextParameter(window,"idApplicationNumber","1325");
	m_sApplicationFactFindNumber = scScreenFunctions.GetContextParameter(window,"idApplicationFactFindNumber","1");
	m_sTypeOfApplicationDesc = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationDescription","");
	m_sCurrency = scScreenFunctions.GetContextParameter(window,"idCurrency","");
	m_sTypeOfApplication = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","");
	m_sTypeOfApplicationValue = scScreenFunctions.GetContextParameter(window,"idMortgageApplicationValue","");
	<%//var TempXML = new scXMLFunctions.XMLObject();%>
	<%//GD 19/06/2002 BMIDS00077 - Upgrade to Core 7.0.2%>
	var TempXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	if (TempXML.IsInComboValidationList(document,"TypeOfMortgage",m_sTypeOfApplication,["N"]))
		m_sTypeOfApplication = "N";
	else if (TempXML.IsInComboValidationXML(["R"]))
		m_sTypeOfApplication = "R";
	else if (TempXML.IsInComboValidationXML(["F"]))
		m_sTypeOfApplication = "F";
	else if (TempXML.IsInComboValidationXML(["T"]))
		m_sTypeOfApplication = "T";
	TempXML = null;
		
}

function SaveLoanPropertyDetails(bOnCommit)
{
	var bSuccess = true;
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XML.CreateRequestTag(window,null);
	<% /* EP2_780 */ %>
	if(!(ValidateSharedAmount()))
	{
		frmScreen.txtSharedAmount.focus();
		return;
	}
	if((frmScreen.optFamilySaleYes.checked & (frmScreen.txtDiscount.value == "0" | frmScreen.txtDiscount.value == "" )))
	{	
		alert("Discount amount cannot be zero");
		frmScreen.txtDiscount.focus ();
		return false;
	}
	if (parseInt(frmScreen.txtDiscount.value, 10) > parseInt(frmScreen.txtPurchasePrice.value, 10))
	{
		alert("Discount amount must be less than or equal to purchase price");
		frmScreen.txtDiscount.focus ();
		return false;
	}
	if (frmScreen.txtPreEmptionDate.value != "")
	{
		var PreemptiveDate = frmScreen.txtPreEmptionDate.value;
		if (scScreenFunctions.CompareDateStringToSystemDate(PreemptiveDate , "<=") == true)
		{
			alert("Pre-emption End Date must be in the future.");
			frmScreen.txtPreEmptionDate.focus ();
			return false;
		}
	}
	
	if(bOnCommit)
	{
		// called from OK button so do critical data check.
		XML.SetAttribute("OPERATION","CriticalDataCheck");
		XML.CreateActiveTag("CUSTOMER");
	}
	
	var tagLoanPropertyDetails = XML.CreateActiveTag("LOANPROPERTYDETAILS");
	
	<%/* NEWPROPERTY */%>
	var tagNewProp = XML.CreateActiveTag("NEWPROPERTY");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("DESCRIPTIONOFPROPERTY", frmScreen.cboPropertyDescription.value);
	XML.CreateTag("PROPERTYLOCATION", frmScreen.cboPropertyLocation.value);
	XML.CreateTag("TYPEOFPROPERTY", frmScreen.cboTypeOfProperty.value);
	XML.CreateTag("VALUATIONTYPE", frmScreen.cboValuationType.value);
	<% /* EP2_780/788 */ %>
	XML.CreateTag("DISCOUNTAMOUNT", frmScreen.txtDiscount.value);
	XML.CreateTag("PREEMPTIONENDDATE", frmScreen.txtPreEmptionDate.value);
	XML.CreateTag("FAMILYSALEINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"FamilySaleGroup"));
	<% /* MF 04/08/2005 MARS20 New fields/fields from DC220 */ %>
	XML.CreateTag("ANYOTHERRESIDENTSINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"OtherResidentsGroup"));	
	XML.CreateTag("DATEOFENTRY",frmScreen.txtDateOfEntry.value);	
	XML.CreateTag("TENURETYPE",frmScreen.cboTenureOfProperty.value);
	XML.CreateTag("NEWPROPERTYINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"NewPropertyRadioGroup"));	
	XML.CreateTag("HOUSEBUILDERSGUARANTEE",frmScreen.cboHouseBuildersGuarantee.value);
	<% /* EP2_677 */ %>
	XML.CreateTag("FLATABOVECOMMERCIAL",scScreenFunctions.GetRadioGroupValue(frmScreen,"FlatAboveCommercialGroup"));	
	XML.CreateTag("FULLMARKETVALUE",scScreenFunctions.GetRadioGroupValue(frmScreen,"FullMarketValueGroup"));	
	XML.CreateTag("FINANCIALINCENTIVES",scScreenFunctions.GetRadioGroupValue(frmScreen,"FinancialIncentivesGroup"));	
	XML.CreateTag("AGRICULTURALRESTRICTIONS",scScreenFunctions.GetRadioGroupValue(frmScreen,"AgriculturalRestrictionGroup"));	
	XML.CreateTag("NOFLATSINBLOCK",frmScreen.txtNumberOfFlats.value);	
	
	// Load the DC202 New property stuff if relevant.
	// Only save our Retype data if still Retype?
	var ValTXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var isRetype = ValTXML.IsInComboValidationList(document,"ValuationType",frmScreen.cboValuationType.value,Array("RT"));
	if (isRetype == true)
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

	var XMLLeasehold = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	//check if Leasehold to ensure we don't pass it data that is not required  JD MAR40
	if ((frmScreen.cboTenureOfProperty.value != 2) && (frmScreen.cboTenureOfProperty.value != 3))
	{
		XMLLeasehold.LoadXML(m_sLeaseholdHoldingXML);
	}
	else
	{
		XMLLeasehold.LoadXML(m_sLeaseholdXML);
	}
	XML.AddXMLBlock(XMLLeasehold.XMLDocument);
	
	<% /* MF MAR20 Add no. of rooms */ %>
	XML.CreateActiveTag("NEWPROPERTYROOMTYPELIST");
	XML.CreateActiveTag("NEWPROPERTYROOMTYPE");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("ROOMTYPE","2");
	XML.CreateTag("NUMBEROFROOMS",frmScreen.txtNumberOfBedrooms.value);	
		
	<%/* NEWLOAN */%>
	XML.ActiveTag = tagLoanPropertyDetails;
	XML.CreateActiveTag("NEWLOAN");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("SOLICITORINDICATOR",scScreenFunctions.GetRadioGroupValue(frmScreen,"SolicitorGroup"));

	<%/* APPLICATIONFACTFIND */%>
	XML.ActiveTag = tagLoanPropertyDetails;
	XML.CreateActiveTag("APPLICATIONFACTFIND");
	XML.CreateTag("APPLICATIONNUMBER", m_sApplicationNumber);
	XML.CreateTag("APPLICATIONFACTFINDNUMBER", m_sApplicationFactFindNumber);
	XML.CreateTag("PURCHASEPRICEORESTIMATEDVALUE", frmScreen.txtPurchasePrice.value);
	XML.CreateTag("EARLIESTCOMPLETIONDATE", frmScreen.txtEarliestRemortgage.value);	 
	XML.CreateTag("ARRANGEOWNCOVERINDICATOR", scScreenFunctions.GetRadioGroupValue(frmScreen,"OwnBuilConRadioGroup"));
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
	//EP2_780/788
	XML.CreateTag("SHAREDPERCENTAGE",frmScreen.txtSharedAmount.value);
	
	if (bShared)
		XML.CreateTag("SHAREDOWNERSHIPTYPE", "2");
	else
		XML.CreateTag("SHAREDOWNERSHIPTYPE", "");
	
	<% /* MF MAR20 Add tag to signify this call will be coming from
		DC201 rather than the original DC200. DC200 can still 
		call "UpdateLoanProperty" */ %>
	XML.ActiveTag = tagNewProp;
	XML.CreateTag("NEWPROPERTYSUMMARY", "");
	
	// Save the details
	if (m_sMetaAction == "Add")
		//BMIDS01002 Add in call to screen rules
		//XML.RunASP(document,"CreateLoanProperty.asp");
		switch (ScreenRules())
			{
			case 1: // Warning
			case 0: // OK
				<% /* MF Created new method in order to save number of bedrooms data 
				XML.RunASP(document,"CreateLoanProperty.asp"); */ %>
				XML.RunASP(document,"CreatePropertyDetails.asp");
				break;
			default: // Error
				XML.SetErrorResponse();
			}

	else
	{		
		//ROLLED FORWARD AGAIN DRC - 20/07/01 SYS2308  GD ROLLBACK 11/05/01 SYS2050
		// Add CRITICALDATACONTEXT TJ_30/03/2001 AQR SYS2050
		if(bOnCommit)
		{
			// JD MAR40
			// called from OK button. Do critical data check
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
			// JD MAR40
			//called from btnOtherResidents or btnOwnInsurance so don't do critical data check
			XML.RunASP(document,"UpdateLoanProperty.asp");
		}
		
	}	

	bSuccess = XML.IsResponseOK();
	
	if (bSuccess)
	{
		if ( XML.IsInComboValidationList(document,"TypeOfMortgage",m_sTypeOfApplicationValue,["R"]) && bEarlierRemortgageDateChanged)
		{
			//MAR783 JD Check for an ACA message first
			if(ACAMessageReceived())
			{
				//MAR537 GHun Get current stage details
				var stageXML = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
				stageXML.CreateRequestTag(window, "GetCurrentStage");
				stageXML.CreateActiveTag("CASEACTIVITY");
				stageXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
				stageXML.SetAttribute("CASEID", m_sApplicationNumber);
				stageXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window, "idActivityId", null));
				stageXML.SetAttribute("ACTIVITYINSTANCE", "1");
				stageXML.RunASP(document, "MsgTMBO.asp");
			
				if(!stageXML.IsResponseOK())
					return false;
				//MAR537 End
			
				var sTMFTEarliestCompDateTaskID = XML.GetGlobalParameterString(document,"TMFTEarliestCompDateTaskID");	
			
				//MAR537 GHun Fix request XML and asp file name
				var CaseTaskXML = new window.top.frames[1].document.all.scXMLFunctions.XMLObject();
				var reqTag = CaseTaskXML.CreateRequestTag(window, "CreateAdhocCaseTask");
				CaseTaskXML.SetAttribute("USERID", scScreenFunctions.GetContextParameter(window, "idUserID", null));
				CaseTaskXML.SetAttribute("UNITID", scScreenFunctions.GetContextParameter(window, "idUnitID", null));		
				CaseTaskXML.SetAttribute("USERAUTHORITYLEVEL", scScreenFunctions.GetContextParameter(window, "idRole", null));
				CaseTaskXML.CreateActiveTag("APPLICATION");		
				CaseTaskXML.SetAttribute("APPLICATIONPRIORITY", scScreenFunctions.GetContextParameter(window, "idApplicationPriority", null));
				CaseTaskXML.ActiveTag = reqTag;	
				CaseTaskXML.CreateActiveTag("CASETASK");
				CaseTaskXML.SetAttribute("SOURCEAPPLICATION", "Omiga");
				CaseTaskXML.SetAttribute("CASEID", m_sApplicationNumber);	
				CaseTaskXML.SetAttribute("ACTIVITYID", scScreenFunctions.GetContextParameter(window, "idActivityId", null));
				CaseTaskXML.SetAttribute("ACTIVITYINSTANCE", "1");
				CaseTaskXML.SetAttribute("CASESTAGESEQUENCENO", stageXML.GetTagAttribute("CASESTAGE","CASESTAGESEQUENCENO"));
				CaseTaskXML.SetAttribute("STAGEID", scScreenFunctions.GetContextParameter(window, "idStageId", null));
				CaseTaskXML.SetAttribute("TASKID", sTMFTEarliestCompDateTaskID);

				CaseTaskXML.RunASP(document, "OmigaTmBO.asp");
				//MAR537 End

				if(CaseTaskXML.IsResponseOK())
				{
					bSuccess = true;
				}
				else
				{
					bSuccess = false;
				}
			}	
		}		
	}
	
	XML = null;
	return(bSuccess);
}
function ACAMessageReceived()
{
	var bACAMessageRec = true;
	
	var XMLAppFirstTitle = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	XMLAppFirstTitle.CreateRequestTag(window, "GetApplicationFirstTitle");
	XMLAppFirstTitle.CreateActiveTag("APPLICATION");
	XMLAppFirstTitle.SetAttribute("APPLICATIONNUMBER",m_sApplicationNumber);
	XMLAppFirstTitle.SetAttribute("APPLICATIONFACTFINDNUMBER",m_sApplicationFactFindNumber );
	XMLAppFirstTitle.SetAttribute("MESSAGETYPE","ACA");
		
	XMLAppFirstTitle.RunASP(document, "omFirstTitle.aspx");		

	<% /*  Check for a record returned. Method does not return RECORDNOTFOUND omiga error. */ %>
	XMLAppFirstTitle.SelectTag(null, "APPLICATIONFIRSTTITLE");
	if(XMLAppFirstTitle.ActiveTag == null)
		bACAMessageRec = false;  //record not found
	
	return(bACAMessageRec);
}
function frmScreen.txtEarliestRemortgage.onchange()
{
	bEarlierRemortgageDateChanged = true;
}

// EP2_2 - New Methods
function frmScreen.cboValuationType.onclick()
{
	// Retype, or not retype?
	var XML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
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

//EP2_780 New logic to set values and enabling on new fields.
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
<% /* EP2_780/788 */ %>
function ValidateSharedAmount()
{
	var bValid = true;
	
	if ((parseInt(frmScreen.txtSharedAmount.value,10) >= parseInt(frmScreen.txtPurchasePrice.value,10)) ||
		((frmScreen.txtPurchasePrice.value == "") && (frmScreen.txtSharedAmount.value != "")))
	{
		alert("Shared amount cannot be greater than the Purchase Price/Estimated");		
		bValid = false;
	}
	return(bValid);
} 

// EP2_2 - END New Methods


-->
</script>
</body>
</html>
