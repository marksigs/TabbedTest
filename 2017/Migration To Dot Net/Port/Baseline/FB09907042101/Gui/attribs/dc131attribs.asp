<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	<% /* BMIDS00190 DCWP3 BM076 cboApplicant has been replaced by a table
	frmScreen.cboApplicant.setAttribute("required", "true");
	*/ %>
	frmScreen.txtMaximumBalance.setAttribute("amount",".");
	frmScreen.txtMaximumBalance.setAttribute("filter","[0-9]");
	frmScreen.txtMonthlyRepayment.setAttribute("amount",".");
	frmScreen.txtMonthlyRepayment.setAttribute("filter","[0-9.]");
	frmScreen.txtMonthlyRepayment.setAttribute("min","0");
	frmScreen.txtMonthlyRepayment.setAttribute("max","999999.99");
	frmScreen.cboDescriptionOfLoan.setAttribute("required","true");
	frmScreen.txtMaximumNumberOfMonths.setAttribute("integer","true");
	frmScreen.txtMaximumNumberOfMonths.setAttribute("filter", "[0-9]");
	frmScreen.txtMaximumNumberOfMonths.setAttribute("min","1");
	frmScreen.txtMaximumNumberOfMonths.setAttribute("max","9999");
	frmScreen.txtDateCleared.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtAdditionalDetails.setAttribute("required", "true");
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