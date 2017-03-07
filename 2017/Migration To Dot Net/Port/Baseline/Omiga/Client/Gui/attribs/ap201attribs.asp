<SCRIPT LANGUAGE="JavaScript">

/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history
PE		06/03/2006	MAR1331	Omiga - Error message for an invalid email did not appear
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

function SetMasks() 
{
	frmScreen.txtCompany.setAttribute("wildcard", "true");
	//BMIDS00657 Company should be mandatory
	frmScreen.txtCompany.setAttribute("required", "true");
	//BMIDS00658/660 Allow wildcard in Panel Id and Town
	frmScreen.txtSearchPanelNo.setAttribute("wildcard", "true");
	frmScreen.txtSearchTown.setAttribute("wildcard", "true");
	frmScreen.txtDateOfInspection.setAttribute("date", "DD/MM/YYYY"); //JD MAR40

	<% // MAR1331 - Omiga - Error message for an invalid email did not appear %>
	<% // Peter Edney %>	
	frmScreen.txtEmail.setAttribute("filter", "[-A-Za-z0-9@._/]");
	frmScreen.txtEmail.setAttribute("msg", "Please enter a valid email address.");
	frmScreen.txtEmail.setAttribute("regexp", "^([a-zA-Z0-9_\\-\\.]+)@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.)|(([a-zA-Z0-9\\-]+\\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\\]?)$");

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





	
	