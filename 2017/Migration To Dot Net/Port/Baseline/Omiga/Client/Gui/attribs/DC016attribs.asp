<SCRIPT LANGUAGE="JScript">
<%/* ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      DC016attribs.asp
Copyright:     Copyright © 

Description:   Introducer Fees Screen attributes
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
SR		22/02/2007	Created
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/ %>
function SetMasks()
{
	frmScreen.txtFeeAmount.setAttribute("filter","[0-9]");
	frmScreen.txtFeeAmount.setAttribute("min", "0");
	frmScreen.txtFeeAmount.setAttribute("max", "99999");
	frmScreen.txtFeeAmount.setAttribute("required", "true");
	frmScreen.txtFeeAmount.setAttribute("integer");
	frmScreen.txtFeeAmount.setAttribute("nym", "Fee Amount");	

	frmScreen.txtRebateAmount.setAttribute("filter","[0-9]");
	frmScreen.txtRebateAmount.setAttribute("min", "0");
	frmScreen.txtRebateAmount.setAttribute("max", "99999");
	frmScreen.txtRebateAmount.setAttribute("integer");
	frmScreen.txtRebateAmount.setAttribute("nym", "Rebate Amount");

	frmScreen.txtRefundAmount.setAttribute("filter","[0-9]");
	frmScreen.txtRefundAmount.setAttribute("min", "0");
	frmScreen.txtRefundAmount.setAttribute("max", "99999");
	frmScreen.txtRefundAmount.setAttribute("integer");
	frmScreen.txtRefundAmount.setAttribute("nym", "Refund Amount");
}

</SCRIPT>