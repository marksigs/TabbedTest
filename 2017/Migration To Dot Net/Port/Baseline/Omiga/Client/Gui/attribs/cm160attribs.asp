<SCRIPT LANGUAGE="JScript">

function SetMasks()
{	
	frmScreen.txtInterestOnlyAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtInterestOnlyAmount.setAttribute("mustenter","true");
	frmScreen.txtCapitalAndInterestAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtTotalLoanAmount.setAttribute("filter", "[0-9]");
}
</SCRIPT>