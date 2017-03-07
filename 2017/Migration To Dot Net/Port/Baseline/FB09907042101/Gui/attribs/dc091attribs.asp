<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	<% /* BMIDS00190 DCWP3 BM076 cboApplicant has been replaced by a table
	frmScreen.cboApplicant.setAttribute("required", "true");
	*/ %>
	<% /* BM0147 MDC 04/12/2002
	frmScreen.txtCompanyName.removeAttribute("required"); 
	*/ %>
	frmScreen.cboOrganisationType.removeAttribute("required");
	frmScreen.txtTotalOutstandingBalance.setAttribute("filter", "[0-9]");
	frmScreen.txtMonthlyRepayment.setAttribute("amount",".");
	frmScreen.txtMonthlyRepayment.setAttribute("required","true");
	frmScreen.txtMonthlyRepayment.setAttribute("min","0");
	frmScreen.txtMonthlyRepayment.setAttribute("max","999999.99");
	frmScreen.txtMonthlyRepayment.setAttribute("filter","[0-9.]");
	frmScreen.txtEndDate.setAttribute("date","DD/MM/YYYY");
	frmScreen.txtAdditionalDetails.setAttribute("required", "true");
	<% /* BMIDS00336 MDC 30/08/2002 */ %>
	frmScreen.txtOriginalLoanAmount.setAttribute("filter", "[0-9]");
	<% /* BMIDS00336 MDC 30/08/2002 - End */ %>
	
	//START: MAR36 - New code added by Joyce Joseph on 10-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtCompanyName.style.textTransform = "capitalize";
	//END: MAR36
	//MAR250 
	frmScreen.txtLoanAccountNumber.maxLength = 20;
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
</SCRIPT>