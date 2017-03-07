<script language="JScript">
function SetMasks()
{
	frmScreen.txtLoanAccountNumber.setAttribute("required","true");

	<% /* Peter Edney - 09/01/2007  EP2_717 */ %>
	<% /*  frmScreen.cboPurposeOfLoan.setAttribute("required", "true"); */ %>
	
	frmScreen.txtMortgageProductDescription.setAttribute("filter", "[-.,0-9a-zA-Z'/&% ]");
	frmScreen.cboRepaymentType.setAttribute("required", "true");
	frmScreen.txtOriginalPartAndPartIntOnlyAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtOriginalLoanAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtInterestRate.setAttribute("filter", "[0-9.]");
	frmScreen.txtOutstandingBalance.setAttribute("filter", "[0-9]");
	frmScreen.txtMonthlyRepayment.setAttribute("required", "true");
	frmScreen.txtMonthlyRepayment.setAttribute("filter", "[0-9.]");
	frmScreen.txtMonthlyRepayment.setAttribute("amount", ".");
	frmScreen.txtMonthlyRepayment.setAttribute("min", "0");
	frmScreen.txtMonthlyRepayment.setAttribute("max", "999999.99");
	frmScreen.txtStartDate.setAttribute("filter", "[0-9/]");
	frmScreen.txtStartDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtOriginalTermYears.setAttribute("filter", "[0-9]");
	frmScreen.txtOriginalTermMonths.setAttribute("filter", "[0-9]");
	frmScreen.txtOriginalTermMonths.setAttribute("integer", "true");
	frmScreen.txtOriginalTermMonths.setAttribute("min", "0");
	frmScreen.txtOriginalTermMonths.setAttribute("max", "11");
	frmScreen.cboRedemptionStatus.setAttribute("required", "true");
	frmScreen.txtRedemptionDate.setAttribute("filter", "[0-9/]");
	frmScreen.txtRedemptionDate.setAttribute("date", "DD/MM/YYYY");

	<% /* PSC 09/08/2002 BMIDS00006 - Start */ %>
	frmScreen.txtDisbursedAmount.setAttribute("filter", "[0-9.]");
	frmScreen.txtDisbursedAmount.setAttribute("min","0");
	frmScreen.txtDisbursedAmount.setAttribute("max","999999.99");
	<% /* MAR10 GHun */ %>
	frmScreen.txtAdminOutstandingTerm.setAttribute("filter", "[0-9]");
	frmScreen.txtAdminOutstandingTerm.setAttribute("min","0");
	frmScreen.txtAdminOutstandingTerm.setAttribute("max","99999");
	<% /* MAR10 End */ %>
	frmScreen.txtAvailableForDisbursement.setAttribute("filter", "[0-9.]");
	frmScreen.txtAvailableForDisbursement.setAttribute("min","0");
	frmScreen.txtAvailableForDisbursement.setAttribute("max","999999.99");
	frmScreen.txtCCIIndicator.setAttribute("filter", "[0-9a-zA-Z]");
	frmScreen.txtCCAIndicator.setAttribute("filter", "[0-9a-zA-Z]");
	frmScreen.txtPenaltyPlanDesc.setAttribute("filter", "[.,0-9a-zA-Z' ]");
	frmScreen.txtLoanClass.setAttribute("filter", "[0-9]");
	frmScreen.txtLoanClass.setAttribute("min","0");
	frmScreen.txtLoanClass.setAttribute("max","99999");
	frmScreen.txtPenaltyPlanCode.setAttribute("filter", "[0-9a-zA-Z]");
	frmScreen.txtPenaltyPlanEndDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtOverpayments.setAttribute("filter", "[0-9.]");
	frmScreen.txtOverpayments.setAttribute("min","0");
	frmScreen.txtOverpayments.setAttribute("max","999999.99");
	frmScreen.txtLoanType.setAttribute("filter", "[0-9]");
	frmScreen.txtLoanClass.setAttribute("min","0");
	frmScreen.txtLoanClass.setAttribute("max","99999");
	frmScreen.txtExistingRetentions.setAttribute("filter", "[0-9.]");
	frmScreen.txtExistingRetentions.setAttribute("min","0");
	frmScreen.txtExistingRetentions.setAttribute("max","999999.99");
	frmScreen.txtRedemptionPenalty.setAttribute("filter", "[0-9.]");
	frmScreen.txtRedemptionPenalty.setAttribute("min","0");
	frmScreen.txtRedemptionPenalty.setAttribute("max","999999.99");
	<% /* PSC 09/08/2002 BMIDS00006 - End */ %>

	<% /* BMIDS00336 MDC 30/08/2002 */ %>
	frmScreen.txtLoanEndDate.setAttribute("date", "DD/MM/YYYY");
	<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
	
	<%/*=[MC]BMIDS756 Regulation change Start*/%>
	frmScreen.txtProductStartDate.setAttribute("filter", "[0-9/]");
	frmScreen.txtProductStartDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtIndexCode.setAttribute("filter", "[0-9/]");
	frmScreen.txtIndexCode.setAttribute("max","99");
	//frmScreen.txtVariance.setAttribute("filter", "[0-9/]");
	frmScreen.txtVariance.setAttribute("filter", "[0-9.]");
	frmScreen.txtVariance.setAttribute("max","9.999999");
	<%/*=[MC]BMIDS756 Regulation change End*/%>
	
	<% /* BMIDS907 GHun Validation of new porting fields */ %>
	<% /* MAR46 Remove Remaining Duration and Current Step 
	frmScreen.txtCurrentInterestRateStep.setAttribute("filter", "[0-9]");
	frmScreen.txtRemainingStepDuration.setAttribute("filter", "[0-9]");    */ %>
	
	<% /* MAR46 Add Current Rate Expiry Date */ %>
	frmScreen.txtCurrentRateExpiryDate.setAttribute("filter", "[0-9/]");
	frmScreen.txtCurrentRateExpiryDate.setAttribute("date", "DD/MM/YYYY");

	frmScreen.txtRemainingIntOnlyBalance.setAttribute("filter", "[0-9]");
	frmScreen.txtRemainingCapAndIntAmount.setAttribute("filter", "[0-9]");
	<% /* BMIDS907 End */ %>
	
	<% /* EP2_55 Original LTV */ %>
	<% /* PSC 19/02/2007 EP2_1488 - Start */ %>
	frmScreen.txtOriginalLTV.setAttribute("filter", "[0-9.]");
	frmScreen.txtOriginalLTV.setAttribute("min","0");
	<% /* PSC 19/02/2007 EP2_1488 - End */ %>	
}

<% /* Get data required for client validation here */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function GetRulesData()
{
}

<% /* Specify client validation here */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function ScreenRules()
{
        return 0;
}

<% /* Specify client code for populating screen */ %>
<% /* Added by automated update TW 09 Oct 2002 SYS5115 */ %>
function ClientPopulateScreen() 
{
}
</script>
