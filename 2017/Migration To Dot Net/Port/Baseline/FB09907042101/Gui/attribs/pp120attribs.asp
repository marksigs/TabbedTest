<SCRIPT LANGUAGE="JScript">

function SetMasks() 
{
	frmScreen.txtPaymentAmount.setAttribute("required", "true");
	frmScreen.cboCardType.setAttribute("required", "true");
	frmScreen.txtCardAccountNumber.setAttribute("required", "true");
	frmScreen.txtCardVerNumber.setAttribute("required", "true");
	frmScreen.cboCardVerificationMethod.setAttribute("required", "true");
	frmScreen.txtCardExpiryDate.setAttribute("required", "true");
	
	frmScreen.txtCardAccountNumber.setAttribute("filter", "[0-9]");
	frmScreen.txtCardVerNumber.setAttribute("filter", "[0-9]");
	frmScreen.txtCardStartDate.setAttribute("filter","[0-9/]");
	frmScreen.txtCardExpiryDate.setAttribute("filter", "[0-9/]");
	frmScreen.txtCardIssue.setAttribute("filter", "[0-9]");
	frmScreen.txtCardBillAddress.setAttribute("filter", "[0-9]");
	frmScreen.txtCardBillPostCode.setAttribute("filter", "[0-9]");
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
}
</SCRIPT>