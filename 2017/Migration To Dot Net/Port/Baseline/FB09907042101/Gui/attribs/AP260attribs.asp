<SCRIPT LANGUAGE="JScript">

function SetMasks()
{
	frmScreen.cboValuationType.setAttribute("required", "true");
	frmScreen.txtAppointmentDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDateOfInstruction.setAttribute("date","DD/MM/YYYY");
	frmScreen.txtDateOfInstruction.setAttribute("required","true");
	frmScreen.txtValuationFee.setAttribute("amount", ".");
	frmScreen.txtCompanyName.setAttribute("wildcard", "true");
	frmScreen.txtPanelNo.setAttribute("wildcard", "true");
	//BMIDS00658/660
	frmScreen.txtSearchTown.setAttribute("wildcard", "true");
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