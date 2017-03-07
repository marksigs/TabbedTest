<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtLastValuerName.setAttribute("required", "true");
	frmScreen.txtLastValuationDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtLastValuationAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtReinstatementAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtEstimatedCurrentValue.setAttribute("filter", "[0-9]");
	frmScreen.txtBuildingsSumInsured.setAttribute("filter", "[0-9]");
	frmScreen.cboHomeInsuranceType.setAttribute("required", "true");
}
</SCRIPT>


