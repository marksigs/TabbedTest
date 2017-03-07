<SCRIPT LANGUAGE="JScript">

<% /* Created by automated update TW 09 Oct 2002 SYS5115 */ %>
<%
/*
Description:   Extended Search Results
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SD		19/12/2005	Make Print button invisible
*/
%>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
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
	//Hide the print button
	frmScreen.btnPrint.style.visibility = "hidden";
}

</SCRIPT>
