<SCRIPT LANGUAGE="JScript">

<% /* Created by automated update TW 09 Oct 2002 SYS5115 */ %>

<% /* Specify screen attributes here */ %>
function SetMasks() 
{
}

<% /* Get data required for client validation here */ %>
function GetRulesData()
{
}

<% /* Specify client validation here */ %>
function ScreenRules()
{
        return 0;
}

<% /* Specify client code for populating screen */ %>
function ClientPopulateScreen() 
{
	if (scScreenFunctions.GetContextParameter(window,"idFreezeOveride_Shortfall")=="1") 
		{
		<% /*  MAR1560 - only allow Create new Qoute and calculate  */ %>
			//frmScreen.btnMortgage.disabled = false;
			//EnableSubQuoteButton(frmScreen.btnMortgage, "msgMortgage");
			//frmScreen.btnAccept.disabled = false;
			//frmScreen.btnRecommend.disabled = false;
			//frmScreen.btnSummary.disabled = false;
			frmScreen.btnCreateNewQuote.disabled = false;
			//frmScreen.btnStoredQuotes.disabled = false;
			frmScreen.btnCalculate.disabled = false;
		}
}

</SCRIPT>

