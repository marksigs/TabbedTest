<SCRIPT LANGUAGE="JavaScript">

function SetMasks() 
{
	frmScreen.txtDateReceived.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDateReceived.setAttribute("required", "y");
	frmScreen.txtReinstatementValue.setAttribute("filter", "[0-9]");
	frmScreen.txtPostWorksValuation.setAttribute("filter", "[0-9]");
	frmScreen.txtPresentValuation.setAttribute("filter", "[0-9]");
	frmScreen.txtRetentionAmount.setAttribute("filter", "[0-9]");
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