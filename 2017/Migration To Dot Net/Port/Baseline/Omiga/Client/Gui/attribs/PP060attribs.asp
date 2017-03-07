<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtAdvanceAmount.setAttribute("required", "true");
	frmScreen.txtAdvanceAmount.setAttribute("filter", "[0-9-]");
	
	frmScreen.cboPaymentType.setAttribute("required", "true");
	frmScreen.cboPaymentMethod.setAttribute("required", "true");
	frmScreen.cboPayee.setAttribute("required", "true");
	
	frmScreen.txtIssueDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtIssueDate.setAttribute("filter", "[0-9/]");
	frmScreen.txtIssueDate.setAttribute("required", "true");

	frmScreen.txtCompDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtCompDate.setAttribute("filter", "[0-9/]");
}
</SCRIPT>