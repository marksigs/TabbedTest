<SCRIPT LANGUAGE="JScript">
<%
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      TM020attribs.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Task Search Screen attributes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
PSC		22/08/2005	MAR32 set attributes for txtDueDateStart and txtDueDateEnd
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
%>

function SetMasks()
{
	<% /* PSC 22/08/2005 MAR32 - Start */ %>
	frmScreen.txtDueDateStart.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDueDateEnd.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtSLAExpiryWithin.setAttribute("filter", "[0-9-]");
	<% /* PSC 22/08/2005 MAR32 - End */ %>
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