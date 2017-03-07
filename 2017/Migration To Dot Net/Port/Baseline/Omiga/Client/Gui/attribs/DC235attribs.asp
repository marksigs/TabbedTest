<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtFirstForename.setAttribute("required", "true");
	frmScreen.txtSurname.setAttribute("required", "true");
	frmScreen.txtDateOfBirth.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDateOfBirth.setAttribute("required", "true");
	frmScreen.cboGender.setAttribute("required", "true");
	frmScreen.cboRelationship.setAttribute("required", "true"); //EP2_745
	
	//START: MAR36 - New code added by Joyce Joseph on 10-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtFirstForename.style.textTransform = "capitalize";
	frmScreen.txtSecondForename.style.textTransform = "capitalize";
	frmScreen.txtOtherForenames.style.textTransform = "capitalize";
	frmScreen.txtSurname.style.textTransform = "capitalize";	
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
