<%
/*
Workfile:      modal_omigamenuCustomise.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Customise Menu screen. Client specific version.
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date			Description
DP		19/01/2002		SYS2564 (parent), SYS3829(child) allow client projects to set
						their own currency.
MV		31/01/2002		SYS3961 - allow Client projects to Disable MortgageCalculator Menu option
MO		02/07/2002		BMIDS00090 - removed the mortgage calculator from the menu
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<SCRIPT LANGUAGE="JavaScript">
function Customise()
{
	// MO - BMIDS00090 - Removed the mortgage calculator from the system menu, to put it back in uncomment this line
	// document.all("aMortgageCalculator").innerHTML = "<A id=aMortgageCalculator href=# tabindex=-1>Mortgage Calculator</A>"
}
</SCRIPT>
