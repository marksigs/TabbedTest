<%
/*
Workfile:      dc090Customise.asp
Copyright:     Copyright © 2001 Marlborough Stirling

Description:   Cosmetic customisation file for dc090
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		25/09/01	First version, SYS2564/SYS2784 (child)  Client specific routing
MV		17/05/02	BMIDS00008 - BM046 - Modified Routing Screen
SR		01/06/2004	BMIDS772 - Modified routing : Now to DC085
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>

<form id="frmToDC085" method="post" action="DC085.asp" STYLE="DISPLAY: none">

<script language="JScript">
function RouteNext()
{
	frmToDC085.submit();
}

</script>
