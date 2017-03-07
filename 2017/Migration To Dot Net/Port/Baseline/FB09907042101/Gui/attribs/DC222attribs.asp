<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtDateLeaseStarted.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtGroundRent.setAttribute("amount", ".");
	frmScreen.txtGroundRent.setAttribute("required","true")
	frmScreen.txtOriginalTermOfLeaseYears.setAttribute("filter", "[0-9]");
	frmScreen.txtServiceCharges.setAttribute("amount", ".");
	frmScreen.txtServiceCharges.setAttribute("required", "true");
	frmScreen.txtUnexpiredTermOfLeaseYears.setAttribute("filter", "[0-9]");
	frmScreen.txtUnexpiredTermOfLeaseYears.setAttribute("required", "true");
}
</SCRIPT>