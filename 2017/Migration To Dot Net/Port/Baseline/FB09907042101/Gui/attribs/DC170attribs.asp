<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.cboEmploymentStatus.setAttribute("required", "true");
	frmScreen.txtDateStartedOrEstablished.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDateLeftOrCeasedTrading.setAttribute("date", "DD/MM/YYYY");
	<% /* MV - 15/10/2002 - BMIDS00272 
	frmScreen.txtDateStartedOrEstablished.setAttribute("required", "true"); */%>
	
	//START: MAR36 - New code added by Joyce Joseph on 11-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtCompanyName.style.textTransform = "capitalize";
	frmScreen.txtJobTitle.style.textTransform = "capitalize";
	//END: MAR36
	
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
