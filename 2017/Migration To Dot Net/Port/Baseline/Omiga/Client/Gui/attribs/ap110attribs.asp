<SCRIPT LANGUAGE="JScript">

function SetMasks() 
{
	frmScreen.txtCustDateEstablished.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtCustDateFinInterest.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtDateEstablished.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtCustYearsActing.setAttribute("filter", "[0-9]");
	frmScreen.txtYearsActing.setAttribute("filter", "[0-9]");
	frmScreen.txtAnnualTurnover1.setAttribute("filter", "[0-9]");
	frmScreen.txtAnnualTurnover2.setAttribute("filter", "[0-9]");
	frmScreen.txtAnnualTurnover3.setAttribute("filter", "[0-9]");
	frmScreen.txtGrossProfit1.setAttribute("filter", "[0-9]");
	frmScreen.txtGrossProfit2.setAttribute("filter", "[0-9]");
	frmScreen.txtGrossProfit3.setAttribute("filter", "[0-9]");
	frmScreen.txtNetProfit1.setAttribute("filter", "[0-9]");
	frmScreen.txtNetProfit2.setAttribute("filter", "[0-9]");
	frmScreen.txtNetProfit3.setAttribute("filter", "[0-9]");
	frmScreen.txtSalary1.setAttribute("filter", "[0-9]");
	frmScreen.txtSalary2.setAttribute("filter", "[0-9]");
	frmScreen.txtSalary3.setAttribute("filter", "[0-9]");
	frmScreen.txtDividend1.setAttribute("filter", "[0-9]");
	frmScreen.txtDividend2.setAttribute("filter", "[0-9]");
	frmScreen.txtDividend3.setAttribute("filter", "[0-9]");
	frmScreen.txtCapital1.setAttribute("filter", "[0-9]");
	frmScreen.txtCapital2.setAttribute("filter", "[0-9]");
	frmScreen.txtCapital3.setAttribute("filter", "[0-9]");
	frmScreen.txtRatio1.setAttribute("filter", "[0-9]");
	frmScreen.txtRatio2.setAttribute("filter", "[0-9]");
	frmScreen.txtRatio3.setAttribute("filter", "[0-9]");
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