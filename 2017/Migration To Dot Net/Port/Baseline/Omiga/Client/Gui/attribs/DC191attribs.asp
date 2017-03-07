<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtYearsActingForCustomer.setAttribute("filter", "[0-9]");
	
	//START: MAR36 - New code added by Joyce Joseph on 11-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtCompanyName.style.textTransform = "capitalize";	
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
