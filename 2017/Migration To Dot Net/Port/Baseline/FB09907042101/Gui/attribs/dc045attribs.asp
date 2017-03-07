<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtDateOfBirth.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDateOfBirth.setAttribute("filter", "[0-9/]");
	frmScreen.cboApplicantName.setAttribute("required", "y");
	frmScreen.txtFirstForename.setAttribute("required", "y");
	frmScreen.txtSurname.setAttribute("required", "y");
	frmScreen.cboRelationship.setAttribute("required", "y");
	frmScreen.txtAdditionalRelatives.setAttribute("required", "y");
	
	<% /*SYS0994 Make Sex a mandatory field */ %>
	frmScreen.cboGender.setAttribute("required", "y");
	
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