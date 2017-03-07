<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtCompanyName.setAttribute("wildcard", "true");
	frmScreen.txtTown.setAttribute("wildcard", "true");
	frmScreen.txtSortCode.setAttribute("mask","99-99-99");
	frmScreen.txtSortCode.setAttribute("filter","[0-9- ]");
	<%/* ASu BMIDS00152 26/06/02 */%>
	frmScreen.txtPanelId.setAttribute("wildcard", "true");
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