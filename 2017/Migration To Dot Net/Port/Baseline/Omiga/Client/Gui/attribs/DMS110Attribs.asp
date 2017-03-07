<script language="JScript">
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DMS110Attribs.asp
Copyright:     Copyright © 2007 Vertex

Description:   Attributes for DMS110.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
AS		17/01/2007	EP1299 DMS110 - '&' not allowed in search criteria and failing to locate Global Parameter
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
	frmScreen.txtAccountNameSearch.setAttribute("wildcard", "true");
	frmScreen.txtAccountNameSearch.setAttribute("filter", "[a-zA-Z '-/&]");
	frmScreen.txtUserName.maxLength = 64;
	frmScreen.txtAppNumberSearch.maxLength = 12;
	frmScreen.txtAccountNameSearch.maxLength = 200;
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
