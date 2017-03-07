<SCRIPT LANGUAGE="JScript">

<% /* Created by automated update TW 09 Oct 2002 SYS5115 */ %>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
	<% /* BM0035 GHun 05/12/2002 - CC011 */ %>
	frmScreen.txtSurname.setAttribute("wildcard", "true");
	frmScreen.txtForenames.setAttribute("wildcard", "true");
	frmScreen.txtDateOfBirth.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtPostcode.setAttribute("upper", "true");
	frmScreen.txtPostcode.setAttribute("wildcard", "true");
	frmScreen.txtPostcode.setAttribute("filter", "[0-9a-zA-Z ]");
	<% /* BM0035 End */ %>
}

<% /* Get data required for client validation here */ %>
function GetRulesData()
{
}

<% /* Specify client validation here */ %>
function ScreenRules()
{
        return 0;
}

<% /* Specify client code for populating screen */ %>
function ClientPopulateScreen() 
{
}

</SCRIPT>
