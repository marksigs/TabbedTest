<SCRIPT LANGUAGE="JScript">

function SetMasks() 
{
	frmScreen.txtAmtRequested.setAttribute("required", "y");
	frmScreen.txtAmtRequested.setAttribute("filter", "[0-9]");
	frmScreen.txtTotalLoanAmt.setAttribute("filter", "[0-9]");
	frmScreen.txtTotalMnthCost.setAttribute("filter", "[0-9]");
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
	if (scScreenFunctions.GetContextParameter(window,"idFreezeOveride_Shortfall")=="1") 
		{//temporarily overide the main context freezedata to disable read only status for the whole of this screen  ;
			m_sReadOnly = "0";
			
		}
}
</SCRIPT>