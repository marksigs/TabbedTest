<SCRIPT LANGUAGE="JScript">

<% /* Created by automated update TW 09 Oct 2002 SYS5115 */ %>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
	//START: MAR36 - New code added by Joyce Joseph on 11-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtCompanyName.style.textTransform = "capitalize";
	//END: MAR36
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
