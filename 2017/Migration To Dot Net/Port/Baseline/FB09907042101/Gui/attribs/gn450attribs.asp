<SCRIPT LANGUAGE="JScript">
function SetMasks()
{	
	frmScreen.txtDetails.setAttribute("required", true);	
	frmScreen.cboReason.setAttribute("required", true);	
	<% /* ASu - BMIDS00273 - Increase character set */%>
	frmScreen.txtDetails.setAttribute("filter", "[0-9a-zA-Z .!\"£$%^&*()-_=+\[{\\]};:'@#~\\\<,>.\?\/]");
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


	