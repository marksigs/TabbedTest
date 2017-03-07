<%
/*
Workfile:      dc155Customise.asp
Copyright:     Copyright © 2001 Marlborough Stirling

Description:   Cosmetic customisation file for dc155
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		25/09/01	First version, SYS2564/SYS2789 (child)  Client specific routing
MV		15/05/02	BMIDS00008 Modified the Routing Screen.
*/
%>

<form id="frmToDC070" method="post" action="dc070.asp" STYLE="DISPLAY: none"></form>

<script language="JScript">
function RoutePrevious()
{
	frmToDC070.submit();
}

</script>
