<SCRIPT LANGUAGE="JScript">
/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific history
Prog	Date		AQR		Description
JJ      10/10/2005  MAR119  Fields disabled. 
MV		04/10/2005  MAR252	Amended PaymentDueDate
PE		07/02/2006	MAR1189	Make various fields mandatory 
SC		14/03/2006	UAT1608 Make redemption status field mandatory
PSC		25/01/2007	EP2_928	Make Bank Name read only
PSC		30/01/2007	EP2_1108 Remove validation on Bank Details
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/
function SetMasks()
{
	//SYS4436 - Lender Name is required
	//frmScreen.txtCompanyName.removeAttribute("required");
	
	<% /* PSC 02/08/2002 BMIDS00006 - Start */ %>
	frmScreen.cboOrganisationType.removeAttribute("required");
	frmScreen.txtRentalIncome.setAttribute("filter","[0-9.]");
	frmScreen.txtRentalIncome.setAttribute("amount", ".");
	frmScreen.txtRentalIncome.setAttribute("min", "0");
	frmScreen.txtRentalIncome.setAttribute("max", "999999.99");
	<% /* PSC 02/08/2002 BMIDS00006 - End */ %>
	
	frmScreen.txtIndemnityAmount.setAttribute("filter","[0-9.]");
	frmScreen.txtIndemnityAmount.setAttribute("amount", ".");
	frmScreen.txtIndemnityAmount.setAttribute("min", "0");
	frmScreen.txtIndemnityAmount.setAttribute("max", "999999.99");
	frmScreen.txtIndemnityMortgageAmount.setAttribute("filter","[0-9]");
	frmScreen.txtIndemnityMortgageAmount.setAttribute("amount", ".");
	frmScreen.txtIndemnityDatePaid.setAttribute("filter","[0-9/]");
	frmScreen.txtIndemnityDatePaid.setAttribute("date","DD/MM/YYYY");
	
	<%/* MV - MAR252 - 27/10/2005 
	frmScreen.txtPaymentDueDate.setAttribute("date","DD/MM/YYYY"); */%>
	frmScreen.txtPaymentDueDate.setAttribute("min", "0");
	frmScreen.txtPaymentDueDate.setAttribute("max", "99");
	<%/*[MC]BMIDS756 Regulation Changes Start*/%>
	frmScreen.txtCallateralID.setAttribute("filter","[0-9]");
	frmScreen.txtCallateralID.setAttribute("MaxLength","12");
	frmScreen.txtCollateralBalance.setAttribute("filter","[0-9.]");
	frmScreen.txtCollateralBalance.setAttribute("MaxLength","14");
	frmScreen.txtSortCode.setAttribute("filter","[-0-9]");
	frmScreen.txtSortCode.setAttribute("MaxLength","8");
	frmScreen.txtSortCode.setAttribute("mask","99-99-99");
	frmScreen.txtSortCode.setAttribute("filter","[0-9- ]");
	
	<% /* MAR20 allow monetary type only for Total Monthly Cost */ %>
	frmScreen.txtTotalMonthly.setAttribute("filter","[0-9.]");
	
	//frmScreen.txtInterestRate.setAttribute("filter","[0-9a-zA-Z]");
	frmScreen.txtBankAccountNumber.setAttribute("MaxLength","18");
	frmScreen.txtBankAccountName.setAttribute("MaxLength","18");
	frmScreen.txtBusinessChannel.setAttribute("MaxLength","40");
	frmScreen.txtPaymentDueDate.setAttribute("filter","[0-9]");
	//frmScreen.txtPaymentDueDate.setAttribute("date","DD/MM/YYYY");
	
	<%/*[MC]BMIDS756 Regulation Changes END*/%>
	
	//START: MAR36 - New code added by Joyce Joseph on 11-Aug-2005
	// Capitlise text fields in Screens
	frmScreen.txtCompanyName.style.textTransform = "capitalize";
	frmScreen.txtIndemnityCompanyName.style.textTransform = "capitalize";
	//END: MAR36

	<% //MAR1189 %>
	frmScreen.txtCollateralBalance.setAttribute("required","true");
	frmScreen.txtTotalMonthly.setAttribute("required","true");
	
	//START:UAT1608
	frmScreen.cboRedemptionStatus.setAttribute("required","true");
	//END:UAT1608

	//EP2_18
	frmScreen.txtPreEmptionEndDate.setAttribute("date","DD/MM/YYYY");

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
	//START: MAR119 - New code added by Joyce Joseph on 10-Oct-2005
	//Fields disabled.
	frmScreen.txtInterestRate.className = 'msgReadOnly';
	frmScreen.txtInterestRate.disabled=true;
	
	frmScreen.txtSortCode.className = 'msgReadOnly';
	frmScreen.txtSortCode.disabled=true;
	<% /* PSC 30/01/2007 EP2_1108 - Start */ %>
	frmScreen.txtSortCode.removeAttribute("MaxLength");
	frmScreen.txtSortCode.removeAttribute("mask");
	frmScreen.txtSortCode.removeAttribute("filter");
	<% /* PSC 30/01/2007 EP2_1108 - End */ %>

	
	frmScreen.txtBankAccountNumber.className = 'msgReadOnly';
	frmScreen.txtBankAccountNumber.disabled=true;
	<% /* PSC 30/01/2007 EP2_1108 */ %>
	frmScreen.txtBankAccountNumber.removeAttribute("MaxLength");

	frmScreen.txtBankAccountName.className = 'msgReadOnly';
	frmScreen.txtBankAccountName.disabled=true;
	<% /* PSC 30/01/2007 EP2_1108 */ %>
	frmScreen.txtBankAccountName.removeAttribute("MaxLength");

	frmScreen.txtBusinessChannel.className = 'msgReadOnly';
	frmScreen.txtBusinessChannel.disabled=true;
	//END: MAR119.
	
	<% /* PSC 25/01/2007 EP2_928 - Start */ %>
	frmScreen.txtBankName.className = 'msgReadOnly';
	frmScreen.txtBankName.disabled=true;
	<% /* PSC 25/01/2007 EP2_928 - End */ %>

}
</SCRIPT>