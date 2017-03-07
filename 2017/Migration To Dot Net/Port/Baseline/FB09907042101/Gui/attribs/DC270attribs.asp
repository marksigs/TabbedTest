<SCRIPT LANGUAGE="JScript">
function SetMasks()
{
	frmScreen.txtAccountNumber.setAttribute("filter", "[0-9]");
	frmScreen.txtAccountNumber.setAttribute("max", "8");
	frmScreen.txtTimeAtBank.setAttribute("filter", "[0-9]");
	frmScreen.txtProposedRepaymentDate.setAttribute("filter", "[0-9]");
	frmScreen.txtAccountNumber.setAttribute("required", "true");
	frmScreen.txtAccountName.setAttribute("required", "true");
<% /* ASu - BMIDS00259 - Branch name should also be mandatory */ %>	
	frmScreen.txtBranchName.setAttribute("required", "true");
	
<% /* SR 17/05/00-SYS0689:Proposed method of Repayment is required only when 
						  RepaymentBankAccount is set to True	
	frmScreen.cboProposedMethodOfRepayment.setAttribute("required", "true");
   */
%>
	frmScreen.txtAccountName.setAttribute("upper","true");
	
	//START: MAR36 - New code added by Joyce Joseph on 15-Aug-2005
	// Capitlise text fields in Screens	
	frmScreen.txtCompanyName.style.textTransform = "capitalize";
	frmScreen.txtBranchName.style.textTransform = "capitalize";
	//END: MAR36
		
	//MAR334 as a result of exluding ThirdPart include
	frmScreen.txtCompanyName.setAttribute("wildcard", "true");
	frmScreen.txtCompanyName.setAttribute("required", "true");
	frmScreen.txtSortCode.setAttribute("mask","99-99-99");
	frmScreen.txtSortCode.setAttribute("required", "true");
	frmScreen.txtSortCode.setAttribute("filter","[0-9- ]");

	<% /* MF 01/09/2005 MAR20 Set text for Direct Debit Explained text area */ %>
	<% /* MAR507  Add new Direct Debit text */ %>
	var sText = "";
	sText += "Mr/Mrs xxxx I now need to take your bank details in order to set up a direct debit, so that you can pay your\r\n";
	sText += "monthly mortgage repayment direct from your current account.\r\n\r\n";
	
	sText += "Can you please confirm your account name?\r\n";
	sText += "If the account is not in the payers name, we are not able to send a paper DDI "
	sText += "so confirm with the customer that we need an account in their name.\r\n\r\n";

	sText += "Can you confirm that this is a personal account?\r\n";
	sText += "Confirm that the person entering into the transaction is the only ";
	sText += "person required to authorise debits from the account. If not, we need to speak ";
	sText += "to the relevant person.\r\n\r\n";
	
	sText += "May I take your bank sort code and account number please?\r\n";
	sText += "(use bank wizard to find the details)\r\n";
	sText += "Can I confirm that is…….. and the address is……….\r\n\r\n";
	
	sText += "Which day of the month would you like to make your payments?\r\n";
	sText += "This is the day on which your Mortgage payments will be collected ";
	sText += "each month and interest added.\r\n\r\n";
	
	sText += "If applicant wants to make their payments from the 28 to the 31 of the month:\r\n";
	sText += "As you have requested to make your payment 28-31 of the month, ";
	sText += "the payment will be collected on the last working day of the month. ";
	sText += "We will submit the request to debit your ";
	sText += "payment account two business days before your mortgage repayments are due.\r\n\r\n";
	
	sText += "We will send the direct debit instruction to your bank once your mortgage ";
	sText += "has completed and the new account is active.\r\n";
	sText += "If we are unable to lodge the DDI within 5 working days, we will advise you.\r\n";
	sText += "When money is transferred from your bank account the payment will appear on your ";
	sText += "statement as a direct debit to DB UK Bank Ltd with your mortgage account number.\r\n\r\n";
	
	sText += "The arrangement is covered by the direct debit scheme and if there ";
	sText += "was an error you are entitled to an immediate and full refund from your bank or ";
	sText += "building society.\r\n";
	sText += "If, in the future, there is a change to the date, amount or frequency of the ";
	sText += "direct debit, we will provide advance notice of 5 working days of the account being debited.\r\n\r\n";
	
	sText += "You have the right to cancel at any time, provided you maintain your ";
	sText += "payments and this guarantee is offered by all banks and building societies that ";
	sText += "take part in the direct debit scheme.\r\n\r\n";
	
	sText += "You will be sent confirmation of this direct debit within the next ";
	sText += "three days so you can check the details are correct. ";
	sText += "A copy of the direct debit guarantee will also be included in the ";
	sText += "letter. \r\n";
	frmScreen.txtDDExplanation.value = sText;	
	frmScreen.txtDDExplanation.style.color = "DarkOrange";
	
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
<% /* BMIDS692  */ %>
  
  /* START: CODE COMMENTED BY Joyce Joseph AQR-MAR59 
  if (scScreenFunctions.GetContextParameter(window,"idFreezeOveride_Rev_and_Decis")=="1") 
  {
  
  if(m_sMetaAction == "Edit")
	{	
		frmScreen.cboProposedMethodOfRepayment.disabled = false;
		frmScreen.cboProposedMethodOfRepayment.style.backgroundColor = "white";
		frmScreen.txtDDReference.disabled = false;
 
    // call to onchange triggers enabling DD reference text box
		frmScreen.cboProposedMethodOfRepayment.onchange();
	}
	else
	{
		DefaultFields();
		SetAvailableFunctionality();
		frmScreen.txtAccountNumber.readOnly = false;
		frmScreen.txtAccountNumber.disabled = false;
		frmScreen.txtAccountNumber.style.backgroundColor ="white";
		frmScreen.txtAccountName.readOnly = false;
		frmScreen.txtAccountName.style.backgroundColor ="white";
		frmScreen.txtTimeAtBank.readOnly = false;
		frmScreen.txtTimeAtBank.style.backgroundColor ="white";
		frmScreen.optRepaymentBankAccountYes.disabled = false;
		frmScreen.optRepaymentBankAccountNo.disabled = false;
		scScreenFunctions.SetFocusToFirstField(frmScreen);
	}
	   // set screen level read only to false so any change will be recorded
    m_sReadOnly = "0"
  } 
  END: CODE COMMENTED BY Joyce Joseph AQR-MAR59 */ 
   
}
</SCRIPT>
