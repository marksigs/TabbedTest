<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	<% /* BMIDS00190 DCWP3 BM076 cboApplicant has been replaced by a table
	frmScreen.cboApplicant.setAttribute("required","true");
	*/ %>
	frmScreen.txtAmountOfDebt.setAttribute("required","true");
	frmScreen.txtAmountOfDebt.setAttribute("amount",".");
	frmScreen.txtAmountOfDebt.setAttribute("filter","[0-9]");
	frmScreen.txtMonthlyRepayment.setAttribute("required","true");
	frmScreen.txtMonthlyRepayment.setAttribute("amount",".");
	frmScreen.txtMonthlyRepayment.setAttribute("filter","[0-9.]");
	frmScreen.txtMonthlyRepayment.setAttribute("min","0");
	frmScreen.txtMonthlyRepayment.setAttribute("max","999999.99");
	frmScreen.txtDateDeclared.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDateOfDischarge.setAttribute("date", "DD/MM/YYYY");
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