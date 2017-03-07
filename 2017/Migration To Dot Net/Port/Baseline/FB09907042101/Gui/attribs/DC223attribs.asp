<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtAmount1.setAttribute("amount", ".");
	frmScreen.txtAmount1.setAttribute("min", "0");
	frmScreen.txtAmount1.setAttribute("max", "999999.99");
	frmScreen.txtAmount1.setAttribute("filter", "[0-9.]");
	
	frmScreen.txtAmount2.setAttribute("amount", ".");
	frmScreen.txtAmount2.setAttribute("min", "0");
	frmScreen.txtAmount2.setAttribute("max", "999999.99");
	frmScreen.txtAmount2.setAttribute("filter", "[0-9.]");
	
	frmScreen.txtAmount3.setAttribute("amount", ".");
	frmScreen.txtAmount3.setAttribute("min", "0");
	frmScreen.txtAmount3.setAttribute("max", "999999.99");
	frmScreen.txtAmount3.setAttribute("filter", "[0-9.]");
	
	frmScreen.txtAmount4.setAttribute("amount", ".");
	frmScreen.txtAmount4.setAttribute("min", "0");
	frmScreen.txtAmount4.setAttribute("max", "999999.99");
	frmScreen.txtAmount4.setAttribute("filter", "[0-9.]");
	
	frmScreen.txtAmount5.setAttribute("amount", ".");
	frmScreen.txtAmount5.setAttribute("min", "0");
	frmScreen.txtAmount5.setAttribute("max", "999999.99");
	frmScreen.txtAmount5.setAttribute("filter", "[0-9.]");
	
	frmScreen.txtAmount6.setAttribute("amount", ".");
	frmScreen.txtAmount6.setAttribute("min", "0");
	frmScreen.txtAmount6.setAttribute("max", "999999.99");
	frmScreen.txtAmount6.setAttribute("filter", "[0-9.]");
	
	frmScreen.txtAmount2.setAttribute("required", "true");
	frmScreen.txtAmount3.setAttribute("required", "true");
	frmScreen.txtAmount4.setAttribute("required", "true");
	frmScreen.txtAmount5.setAttribute("required", "true");
	frmScreen.txtAmount6.setAttribute("required", "true");
	
	frmScreen.txtTotalDeposit.setAttribute("amount", ".");
	frmScreen.txtTotalDeposit.setAttribute("filter", "[0-9.]");
	
}
</SCRIPT>