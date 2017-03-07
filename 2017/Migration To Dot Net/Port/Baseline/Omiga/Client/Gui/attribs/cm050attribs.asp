<SCRIPT LANGUAGE="JScript">
function SetMasks() 
{
	frmScreen.txtMonthMortRepayments.setAttribute("mask", "99999");
	frmScreen.txtMonthMortRelatedInsurance.setAttribute("mask", "99999");
	frmScreen.txtMonthLoansAndLiabilities.setAttribute("mask", "99999");
	frmScreen.txtMonthOtherOutgoings.setAttribute("mask", "99999");
	frmScreen.txtTotalMonthOutgoings.setAttribute("mask", "99999");
	frmScreen.txtTotalMonthIncome.setAttribute("mask", "999999");
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