<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtChequeNum.setAttribute("filter", "[0-9]");
	frmScreen.txtChequeNum.setAttribute("required", "true");
	frmScreen.txtPassword.setAttribute("required", "true");
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
</SCRIPT><% /* OMIGA BUILD VERSION 045.02.05.20.00 */ %>
