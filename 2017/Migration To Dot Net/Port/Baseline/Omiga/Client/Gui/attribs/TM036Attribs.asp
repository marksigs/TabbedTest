<SCRIPT LANGUAGE="JScript">

<% /* Created by automated update TW 09 Oct 2002 SYS5115 */ %>
<% /* AW EP2_1945 13/03/07	Extended filter for txtNote*/ %>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
	<% /* AW - EP1264 - Increase character set */%>
	frmScreen.txtNote.setAttribute("filter", "[0-9a-zA-Z .!\"£$%^&*()-_=+\[{\\]};:'@#~\\\<,>.\?\/]");
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
