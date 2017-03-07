<SCRIPT LANGUAGE="JScript">

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history
PE		06/03/2006	MAR1331	Omiga - Error message for an invalid email did not appear
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

function SetMasks()
{
	frmScreen.cboLenderType.setAttribute("required","true");
	frmScreen.txtCompanyName.setAttribute("required","true");
	frmScreen.txtCompanyName.setAttribute("filter","[-0-9a-zA-Z',./&* ]");
	frmScreen.txtTelephone.setAttribute("filter", "[0-9 ]");
	frmScreen.txtFax.setAttribute("filter", "[0-9 ]");	
	
	<% // MAR1331 - Omiga - Error message for an invalid email did not appear %>
	<% // Peter Edney %>	
	frmScreen.txtContactEmailAddress.setAttribute("filter", "[-A-Za-z0-9@._/]");
	frmScreen.txtContactEmailAddress.setAttribute("msg", "Please enter a valid email address.");
	frmScreen.txtContactEmailAddress.setAttribute("regexp", "^([a-zA-Z0-9_\\-\\.]+)@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.)|(([a-zA-Z0-9\\-]+\\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\\]?)$");

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