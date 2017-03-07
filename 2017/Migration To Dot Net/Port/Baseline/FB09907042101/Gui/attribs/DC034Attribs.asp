<SCRIPT LANGUAGE="JScript">
/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DC034attribs.asp
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   Alias/Association Details Screen attributes

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BMids History:

Prog	Date		Description
MDC		11/11/2002	BMIDS00911 Created
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

function SetMasks()
{
	frmScreen.txtFirstForename1.setAttribute("required","true");
	frmScreen.txtSurname1.setAttribute("required","true");
	frmScreen.cboTitle1.setAttribute("required","true");
	frmScreen.cboAliasType1.setAttribute("required","true");
	frmScreen.txtDateOfChange1.setAttribute("date", "DD/MM/YYYY");
	
	//START: MAR36 - New code added by Joyce Joseph on 10-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtFirstForename1.style.textTransform = "capitalize";
	frmScreen.txtSecondForename1.style.textTransform = "capitalize";
	frmScreen.txtOtherForenames1.style.textTransform = "capitalize";
	frmScreen.txtSurname1.style.textTransform = "capitalize";
	frmScreen.txtOtherTitle1.style.textTransform = "capitalize";
	//END: MAR36
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