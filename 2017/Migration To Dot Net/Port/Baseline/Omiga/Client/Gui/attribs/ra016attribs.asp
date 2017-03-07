<SCRIPT LANGUAGE="JScript">
<% /* INR 25/11/2005 MAR645 Created */ %>

function SetMasks() 
{
	with (frmScreen) {
		txtUser.setAttribute("required", "y");
		txtPassword.setAttribute("required", "y");
		txtPassword.setAttribute("filter", ".");
		cboOverrideReason.setAttribute("required", "y");
	}
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