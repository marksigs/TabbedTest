<script type="text/javascript">

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
	<% /* MAR1332 GHun */%>
	frmScreen.cboDocumentPurpose.setAttribute("required", "true");
	frmScreen.cboDocumentName.setAttribute("required", "true");
	frmScreen.cboDocumentGroup.setAttribute("required", "true");
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

</script>
