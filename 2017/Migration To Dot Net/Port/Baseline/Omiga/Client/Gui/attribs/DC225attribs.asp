<SCRIPT LANGUAGE="JScript">
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DC225attribs.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Application source Screen attributes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
IW		20/04/00	Created
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
GD		10/07/02	Mask Grant Amount correctly
*/

function SetMasks()
{
	frmScreen.txtGRANTAMOUNT.setAttribute("amount",".");
	frmScreen.txtGRANTAMOUNT.setAttribute("filter","[0-9.]");
	frmScreen.txtGRANTAMOUNT.setAttribute("min","0");
	frmScreen.txtGRANTAMOUNT.setAttribute("max","999999.99");
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