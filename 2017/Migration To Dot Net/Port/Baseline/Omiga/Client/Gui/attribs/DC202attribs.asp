<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtOVDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtOVDate.setAttribute("required", "true");
	frmScreen.txtOVCoName.setAttribute("required", "true");
	frmScreen.txtOVCoPostCode.setAttribute("required", "true");
	frmScreen.txtOVCoFlatNo.setAttribute("required", "true");
	frmScreen.txtOVCoHouseNo.setAttribute("required", "true");
	frmScreen.txtOVCoHouseName.setAttribute("required", "true");
  
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
