<SCRIPT LANGUAGE="JavaScript">
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Prog	Date		Defect		Description
INR		20/12/2006 EP2_522 - Need max & min for FFPercentage.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
function SetMasks() 
{
	//EP2_2 New Field
	frmScreen.txtFFPercentage.setAttribute("filter", "[0-9]");
	// EP2_522
	frmScreen.txtFFPercentage.setAttribute("integer", "true");
	frmScreen.txtFFPercentage.setAttribute("min", "0");
	frmScreen.txtFFPercentage.setAttribute("max", "100");
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