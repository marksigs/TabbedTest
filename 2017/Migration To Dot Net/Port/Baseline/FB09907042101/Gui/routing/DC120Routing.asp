<%
/*
Workfile:      dc120Customise.asp
Copyright:     Copyright © 2001 Marlborough Stirling

Description:   Cosmetic customisation file for dc120
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		25/09/01	First version, SYS2564/SYS2786 (child)  Client specific routing
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		AQR			Description
MV		17/05/2002	BMIDS00008	Modified Routing Screen 
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<form id="frmToDC150" method="post" action="DC150.asp" STYLE="DISPLAY: none">

<script language="JScript">
function RoutePrevious()
{
	frmToDC150.submit();
}

</script>
