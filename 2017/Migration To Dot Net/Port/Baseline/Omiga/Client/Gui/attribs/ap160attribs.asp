<SCRIPT LANGUAGE="JScript">

function SetMasks() 
{
	frmScreen.txtCustStartDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtStartDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtCustEndDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtEndDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtCustRent.setAttribute("filter", "[0-9]");
	frmScreen.txtRent.setAttribute("filter", "[0-9]");
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