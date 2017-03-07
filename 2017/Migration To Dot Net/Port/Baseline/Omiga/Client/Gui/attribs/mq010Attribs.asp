<SCRIPT LANGUAGE="JScript">

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
	<% /* START: MAR552 Maha T */ %>
	frmScreen.txtPurchasePrice.setAttribute("integer", "true");
	frmScreen.txtPurchasePrice.setAttribute("filter", "[0-9]");
	
	frmScreen.txtAmtRequested.setAttribute("integer", "true");
	frmScreen.txtAmtRequested.setAttribute("filter", "[0-9]");
	<% /* END: MAR552 */ %>
}

<% /* Get data required for client validation here */ %>
function GetRulesData()
{
}

<% /* Specify client validation here */ %>
function ScreenRules()
{
        return 0;
}

<% /* Specify client code for populating screen */ %>
function ClientPopulateScreen() 
{
}

</SCRIPT>
