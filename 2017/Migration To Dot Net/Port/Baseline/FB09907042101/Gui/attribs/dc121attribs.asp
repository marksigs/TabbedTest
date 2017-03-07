<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.cboApplicant.setAttribute("required", "true");
	frmScreen.txtDateDeclined.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtMortgageDeclinedDetails.setAttribute("required", "true");
	frmScreen.txtAdditionalDetails.setAttribute("required", "true");
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