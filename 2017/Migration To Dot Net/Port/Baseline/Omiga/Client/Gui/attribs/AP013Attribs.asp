<SCRIPT LANGUAGE="JScript">

<% /* Created by automated update TW 09 Oct 2002 SYS5115 */ %>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
	<% /* ASu BMIDS00253 - Start. Please note that the following fields should only be checked whilst in 'Override' Mode
	and should the 'Set Masks' function be required in any other mode the calling .asp will need further intelligence */ %>
	frmScreen.cboOverrideReason.setAttribute("required", "y");
	frmScreen.txtOverrideComments.setAttribute("required", "y");
	<% /* ASu - End */ %>
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
