<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtBuildingsSIAmt.setAttribute("amount",".");
	frmScreen.txtBuildingsSIAmt.setAttribute("filter","[0-9]");
	frmScreen.txtRenewalDate.setAttribute("date", "DD/MM/YYYY");
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