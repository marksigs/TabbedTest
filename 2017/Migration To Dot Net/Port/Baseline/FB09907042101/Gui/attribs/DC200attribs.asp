<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtPurchasePrice.setAttribute("filter", "[0-9]");
	//	AW	06/12/02	BM0116
	//frmScreen.txtDiscount.setAttribute("filter", "[0-9]");
	frmScreen.txtSharedAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtPurchasePrice.setAttribute("required", "true");
	<%/* SA 30/5/01 SYS1835 ValuationType & PropertyLocation now mandatory */%>
	frmScreen.cboValuationType.setAttribute("required", "true");
	frmScreen.cboPropertyLocation.setAttribute("required", "true");
	<% /* MV - 11/10/2002 - BMIDS00590	- Amended */ %>
	frmScreen.txtSharedAmount.setAttribute("min","0");
	frmScreen.txtSharedAmount.setAttribute("max","99.99");
	// Ep2_18 - New fields	
	frmScreen.txtPreEmptionDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDiscount.setAttribute("filter", "[0-9]");

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
