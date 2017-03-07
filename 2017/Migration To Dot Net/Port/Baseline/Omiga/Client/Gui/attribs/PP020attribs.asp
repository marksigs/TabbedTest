<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.cboFeePaymentMethod.setAttribute("required", "true");
	frmScreen.cboPaymentEvent.setAttribute("required", "true");
	frmScreen.txtIssueDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtIssueDate.setAttribute("filter", "[0-9/]");
	frmScreen.txtPaymentAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtPaymentAmount.setAttribute("required", "true");
	
	<%/* SR 11-12-01 : SYS2555 */ %>
	frmScreen.txtRefundAmount.setAttribute("filter", "[0-9]");
	frmScreen.txtRefundDate.setAttribute("date", "DD/MM/YYYY");
	frmScreen.txtRefundDate.setAttribute("filter", "[0-9/]");
	<%/* SR 11-12-01 : SYS2555  END */ %>
	
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