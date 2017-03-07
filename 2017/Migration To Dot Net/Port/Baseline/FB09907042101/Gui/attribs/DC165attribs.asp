<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtTaxOffice.setAttribute("filter", "[-A-Za-z0-9.&()/\ ]");
	frmScreen.txtTaxReferenceNumber.setAttribute("filter", "[0-9a-zA-Z /\]");
	frmScreen.txtNationalInsuranceNumber.setAttribute("mask","AA999999A");
	frmScreen.txtNationalInsuranceNumber.setAttribute("upper", "true");
	frmScreen.txtNonUKTaxArrangements.setAttribute("filter", "[-A-Za-z0-9.&()/\ ]");
}
</SCRIPT>
