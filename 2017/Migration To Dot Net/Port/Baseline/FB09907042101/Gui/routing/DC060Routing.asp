<%
/*
Workfile:      dc015Customise.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Routing customisation file for dc060
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
DJP		25/09/01	SYS2564/SYS2751 (child) First version, Client specific routing
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS Specific History:

Prog	Date		Description
MV		07/06/2002	BMIDS00013 - Changed Routing Previous Screen.
SR		19/06/2004  BMIDS772 - Changed RouteNext (DC085 instead of DC070).
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history

Prog    Date			Description
MF		22/07/2005		IA_WP01 process flow changes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<form id="frmToDC030" method="post" action="DC030.asp" STYLE="DISPLAY: none"></form>
<%/* SR 19/06/2004 : BMIDS772 - route to DC085 instead of DC070 */%>
<%/* MF 22/07/2005 : MARS IA_WP01 - route to DC160 */%>
<form id="frmToNext" method="post" action="dc160.asp" STYLE="DISPLAY: none"></form>
<script language="JScript">
function RoutePrevious()
{
	frmToDC030.submit();
}
function RouteNext()
{
	frmToNext.submit();
}

</script>
