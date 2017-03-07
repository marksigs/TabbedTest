<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtPercentSharesHeld.setAttribute("filter", "[0-9]");
	frmScreen.txtDateFinancialInterestHeld.setAttribute("date", "DD/MM/YYYY");
	
	//START: MAR36 - New code added by Joyce Joseph on 10-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtEmployerName.style.textTransform = "capitalize";
	//END: MAR36
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
