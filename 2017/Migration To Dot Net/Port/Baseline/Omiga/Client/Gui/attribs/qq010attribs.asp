<SCRIPT LANGUAGE="JScript">

function SetMasks()
{	
	//DB SYS4767 - MSMS Integration
<%	// Application Source Masks
%>	<% /* MO 4/10/2002 BMIDS00578 frmScreen.txtIndividualName.setAttribute("required", "y"); */ %>
	frmScreen.cboDirectIndirect.setAttribute("required", "y");
	//DB End
<%	// Applicant Masks
%>	frmScreen.cboApplicant1EmploymentStatus.setAttribute("required", "y");
	frmScreen.cboApplicant2EmploymentStatus.setAttribute("required", "y");
	frmScreen.txtApplicant1GrossIncOrNetProfit.setAttribute("filter", "[0-9]");	
	frmScreen.txtApplicant2GrossIncOrNetProfit.setAttribute("filter", "[0-9]");
	frmScreen.txtApplicant1GrossIncOrNetProfit.setAttribute("required", "y");
	frmScreen.txtApplicant2GrossIncOrNetProfit.setAttribute("required", "y");
	frmScreen.txtApplicant1TotalMonthlyOutgoings.setAttribute("filter", "[0-9]");
	frmScreen.txtApplicant2TotalMonthlyOutgoings.setAttribute("filter", "[0-9]");
	frmScreen.txtApplicant1TotalMonthlyOutgoings.setAttribute("required", "y");
	frmScreen.txtApplicant2TotalMonthlyOutgoings.setAttribute("required", "y");
	frmScreen.txtApplicant1TotalMonthlyDebtRepayments.setAttribute("filter", "[0-9]");
	frmScreen.txtApplicant2TotalMonthlyDebtRepayments.setAttribute("filter", "[0-9]");
	frmScreen.txtApplicant1TotalMonthlyDebtRepayments.setAttribute("required", "y");
	frmScreen.txtApplicant2TotalMonthlyDebtRepayments.setAttribute("required", "y");
	frmScreen.txtApplicant1TotalOSBalance.setAttribute("filter", "[0-9]");
	frmScreen.txtApplicant2TotalOSBalance.setAttribute("filter", "[0-9]");
	frmScreen.txtApplicant1TotalOSBalance.setAttribute("required", "y");
	frmScreen.txtApplicant2TotalOSBalance.setAttribute("required", "y");
	frmScreen.cboApplicant1CCJHistory.setAttribute("required", "y");
	frmScreen.cboApplicant2CCJHistory.setAttribute("required", "y");
<%	// Loan Detail masks
%>	frmScreen.txtPurchasePrice.setAttribute("filter", "[0-9]");
	frmScreen.txtOSLoanAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtTotalMonthlyRepayments.setAttribute("amount", ".");
	frmScreen.txtPurchasePrice.setAttribute("required", "y");
	frmScreen.txtOSLoanAmount.setAttribute("required", "y");
	frmScreen.txtTotalMonthlyRepayments.setAttribute("required", "y");
	// BMIDS00654 MDC 04/11/2002
	frmScreen.txtAmountRequested.setAttribute("required", "y");
	frmScreen.txtAmountRequested.setAttribute("filter", "[0-9]");
	// BMIDS00654 MDC 04/11/2002 - End
	
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