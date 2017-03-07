<%
/*
Workfile:      dc140Customise.asp
Copyright:     Copyright © 2001 Marlborough Stirling

Description:   Cosmetic customisation file for dc140
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		25/09/01	First version, SYS2564/SYS2787 (child)  Client specific routing
MF		26/07/2005	MARS019 Change of routing back to DC085
*/
%>

<% /* SR 14/06/2004 : BMIDS772 - route to DC110 on submit instead of DC150 */  %>
<% /* MF 26/07/2005 : MAR019 - route back to DC085 */ %>
<form id="frmToOnSubmit" method="post" action="dc085.asp" STYLE="DISPLAY: none"></form>

<script language="JScript">
function RouteNext()
{
	frmToOnSubmit.submit();
}

</script>
