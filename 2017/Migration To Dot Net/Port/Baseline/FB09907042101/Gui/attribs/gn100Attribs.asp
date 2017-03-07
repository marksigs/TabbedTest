<SCRIPT LANGUAGE="JScript">

<% /* Created by automated update TW 09 Oct 2002 SYS5115 */ %>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
	<% /* ASu - BMIDS00554 - Increase character set */%>
	frmScreen.txtMemoEntry.setAttribute("filter", "[0-9a-zA-Z .!\"£$%^&*()-_=+\[{\\]};:'@#~\\\<,>.\?\/]");
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
