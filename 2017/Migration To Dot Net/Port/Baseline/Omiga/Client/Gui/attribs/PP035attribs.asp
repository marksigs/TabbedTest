<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtTotalAmountDue.setAttribute("integer", "true");
	frmScreen.cboFeeType.setAttribute("required", "true");
	frmScreen.txtTotalAmountDue.setAttribute("required", "true");
	frmScreen.txtTotalAmountDue.setAttribute("min", "1");
	frmScreen.txtTotalAmountDue.setAttribute("max", "99999999");
	frmScreen.txtTotalAmountDue.setAttribute("filter", "[0-9]");
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
