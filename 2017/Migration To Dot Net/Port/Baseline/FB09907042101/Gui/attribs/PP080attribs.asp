<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtAmountReturned.setAttribute("required", "true");
	frmScreen.txtAmountReturned.setAttribute("filter", "[0-9]");
}
<% /* BMIDS00977 - Screen Rules Added */ %>
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