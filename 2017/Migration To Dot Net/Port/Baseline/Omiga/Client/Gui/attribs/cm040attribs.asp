<SCRIPT LANGUAGE="JScript">

function SetMasks()
{
	frmScreen.txtTotalCost.setAttribute("amount", ".");
	frmScreen.txtTotalCost.setAttribute("min", "0");
	frmScreen.txtTotalCost.setAttribute("max", "999999.99");
	frmScreen.txtTotalCost.setAttribute("filter", "[0-9.]");
	<% /* ASu BMIDS00597 */ %>	
	frmScreen.cboCoverType.setAttribute("required", "y");		
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