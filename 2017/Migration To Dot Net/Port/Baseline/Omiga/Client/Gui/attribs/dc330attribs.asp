<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.cboResponse1.setAttribute("required", "true");
	<% /* MV 25/10/2002	BMIDS00714	*/%>  
	<% /* frmScreen.cboResponse2.setAttribute("required", "true");
	frmScreen.cboResponse3.setAttribute("required", "true");
	frmScreen.cboResponse4.setAttribute("required", "true");
	frmScreen.cboResponse5.setAttribute("required", "true");
	frmScreen.cboResponse6.setAttribute("required", "true"); */ %>
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