<SCRIPT LANGUAGE="JScript">
function SetMasks()
{	
	frmScreen.txtOutgoingAmount.setAttribute("filter", "[0-9]");
	<% /* BMIDS00190 DCWP3 BM076 */ %>
	frmScreen.cboRegularOutgoing.setAttribute("required", "true");
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