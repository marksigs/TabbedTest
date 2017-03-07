<SCRIPT LANGUAGE="JScript">

function SetMasks() 
{
	frmScreen.txtCurrCompDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtRevCompDate.setAttribute("date", "DD/MM/YYYY");
}	

<% /* Get data required for client validation here  */ %>
function GetRulesData()
{
}

<% /* Specify client validation here */ %>
function ScreenRules()
{
        return  0;
}

<% /* Specify client code for populating screen */ %>
function ClientPopulateScreen() 
{
}
</SCRIPT>