<SCRIPT LANGUAGE="JScript">
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DC015attribs.asp
Copyright:     Copyright © 2006 Vertex Financial Services

Description:   Intermediary search Screen attributes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
JLD		04/11/99	Created
AD		30/01/00	Rework
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMIDS History:

Prog	Date		Description
MO		26/09/2002	BMIDS00502 Changes
MO		24/10/2002	BMDIS00713 Fixed bug

EPSOM History:

Prog	Date		Description
IK		05/04/2006	EP15 Changes
IK		18/04/2006	EP394 fix error caused by EP15 comment
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

function SetMasks()
{
	frmScreen.txtName.setAttribute("wildcard", "true");
	frmScreen.txtTown.setAttribute("wildcard", "true");
	frmScreen.txtPostCode.setAttribute("wildcard", "true");
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