<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtCurPwd.setAttribute("required", "true");
	frmScreen.txtNewPwd.setAttribute("required", "true");
	frmScreen.txtConfirmedNewPwd.setAttribute("required", "true");

	frmScreen.txtCurPwd.setAttribute("filter", ".");
	frmScreen.txtNewPwd.setAttribute("filter", ".");
	frmScreen.txtConfirmedNewPwd.setAttribute("filter", ".");
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