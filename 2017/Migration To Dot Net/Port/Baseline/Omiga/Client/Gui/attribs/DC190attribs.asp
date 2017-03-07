<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtNetIncome.setAttribute("filter", "[0-9]");
	/*
	MO - 2/7/2002 - BMIDS00146 - This attribute has been put 'back'
	into code as the field is currently needed for costmodelling, when the new
	income calculations are put into BMIDS this attribute can be removed.
	*/
	/*
	ASu 05/09/02 - BMIDS00268 - No Longer mandatory data
	frmScreen.txtNetIncome.setAttribute("required","true");
	*/
	<% /* MAR30 Masks for percentage field */ %>
	frmScreen.txtOtherIncomePercentage.setAttribute("filter", "[0-9]");
	frmScreen.txtOtherIncomePercentage.setAttribute("integer", "true"); //JD MAR239
	frmScreen.txtOtherIncomePercentage.setAttribute("min", "0");
	frmScreen.txtOtherIncomePercentage.setAttribute("max", "100");
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
