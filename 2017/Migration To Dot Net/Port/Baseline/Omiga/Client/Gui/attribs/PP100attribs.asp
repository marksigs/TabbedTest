<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtPostcode.setAttribute("filter", "[0-9a-zA-Z ]");
	frmScreen.txtPostcode.setAttribute("upper", "true");
	frmScreen.txtBankSortCode.setAttribute("mask","99-99-99");
	frmScreen.txtBankSortCode.setAttribute("filter","[0-9- ]");
	frmScreen.txtPostcode.setAttribute("required", "true");
	frmScreen.cboPayeeType.setAttribute("required", "true");
	frmScreen.txtPayeeName.setAttribute("required", "true");
	<% /* MV - 05/11/2002 - BMIDS00819 - Bank Details are not mandatory  
	frmScreen.txtBankSortCode.setAttribute("required", "true");
	frmScreen.txtAccountNumber.setAttribute("required", "true");
	frmScreen.txtBankName.setAttribute("required", "true");
	*/ %>
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