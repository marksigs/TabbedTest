<SCRIPT LANGUAGE="JScript">
<%
/*
History:

Prog	Date		Description
SG		04/03/02	Created
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
%>
function SetMasks()
{
	frmScreen.txtNumberOfOccupants.setAttribute("integer","true");
	frmScreen.txtNumberOfOccupants.setAttribute("filter", "[0-9]");	
	frmScreen.txtNumberOfOccupants.setAttribute("max", "99");	

	frmScreen.txtMonthlyRentalIncome.setAttribute("amount",".");
	frmScreen.txtMonthlyRentalIncome.setAttribute("filter","[0-9.]");
	frmScreen.txtMonthlyRentalIncome.setAttribute("min","0");
	frmScreen.txtMonthlyRentalIncome.setAttribute("max","999999.99");
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