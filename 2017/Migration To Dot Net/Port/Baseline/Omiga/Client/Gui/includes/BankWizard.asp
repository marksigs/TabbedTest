<%
/*
Workfile:      BankWizard.asp
Copyright:     Copyright © 2002 Marlborough Stirling

Description:   Contains shared functions for calling BankWizard
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description
INR		11/03/02	Created for SYS4192.
DRC     24/09/02   BMIDS00183 see Core AQR SYS4847
MO		08/11/2002 BMIDS00726, fixed the return value
AM		15/08/2005 MARS39	Adding the transposeindicator and redirectindicator
TL		06/09/2005	MAR39 Changed to use C# component
Maha T	21/11/2005	MAR334 Comment Address details as its no more used

HMA     20/09/2006  EP2_3 Apply Epsom changes. Add roll number checking.
*/
%>
<script language="JScript">
function BankWizard(ctrSortcode,ctrAccountNumber,ctrRollNumber)
{
	function IsStringEmpty(strString)
	<%/* Ascertains whether the passed string is empty */%>
	{
		var bStringIsEmpty = false;
		var ssArray = strString.split(" ");
		var nLength = ssArray.length;
		var nIndex = 0;

		while(ssArray[nIndex] == "" && nIndex < nLength)
			nIndex++;

		if (ssArray[nIndex] == null)
			bStringIsEmpty = true;

		return(bStringIsEmpty);
	}

	<%/*BMIDS0183 DRC */%>
	var BankSearchXML = new top.frames[1].document.all.scXMLFunctions.XMLObject();
	var bHasData = false;

	BankSearchXML.CreateActiveTag("ValidateBankDetailsRequestType");
	
	<%/* Record the search criteria */%>
	if (ctrSortcode != null) {
		if (!IsStringEmpty(ctrSortcode.value)) 
			BankSearchXML.CreateTag("NewSortCode",ctrSortcode.value.replace(/-/g,""));
	}
	if (ctrAccountNumber != null) {
		if (!IsStringEmpty(ctrAccountNumber.value)) 
			BankSearchXML.CreateTag("NewAccountNumber",ctrAccountNumber.value);
	}
	<%/* EP2_3 Add Roll Number */%>
	if (ctrRollNumber != null) {
		if (!IsStringEmpty(ctrRollNumber.value)) 
			BankSearchXML.CreateTag("NewRollNumber",ctrRollNumber.value);
	}
	
	<% /* MO - 08/11/2002 - BMIDS00726 */ %>
	m_bBankAccountValid = false;

	BankSearchXML.RunASP(document,"GetBankDetails.aspx");
		
	if(BankSearchXML.IsResponseOK())
	{
		var XMLActiveTag = BankSearchXML.SelectSingleNode("//RESPONSE/AccountDetails");

		if (XMLActiveTag == null) 
		{
			alert("BANKDATA data not returned.");
		}
		else if (BankSearchXML.GetTagText("BankWizardError") == "")
		{
			alert("The Bank Sort Code and Account Number have been successfully validated.");
		}
		else
		{
			alert(BankSearchXML.GetTagText("BankWizardError"));
		}

		<%/*Populate Screen*/%>
		if(m_sScreenId == "DC270")
		{
			if (XMLActiveTag != null)
			{
				if (frmScreen.txtBranchName) frmScreen.txtBranchName.value = BankSearchXML.GetTagText("BranchName");

				if (BankSearchXML.GetTagText("BankName") == "")
				{
					if (frmScreen.txtCompanyName) frmScreen.txtCompanyName.value = "";
					alert("Bank Name data not returned.");
				}
				else
				{
					if (frmScreen.txtCompanyName) frmScreen.txtCompanyName.value = BankSearchXML.GetTagText("BankName");
				}

				if (BankSearchXML.GetTagText("TransposedIndicator") == "1")
				{
					alert("Account details are not in standard form and were transposed!");
					if (frmScreen.optTransposedIndicatorYes) frmScreen.optTransposedIndicatorYes.checked = true;
				}
				else
				{
					if (frmScreen.optTransposedIndicatorYes) frmScreen.optTransposedIndicatorYes.checked = false;
				}
				
				if (BankSearchXML.GetTagText("RedirectedIndicator") == "1")
				{
					alert("Account details provided relate to a branch that was closed.  We have given you the new account details.");
				}
			
				XMLActiveTag = BankSearchXML.SelectSingleNode("//RESPONSE/AccountDetails/Address");
				
				<% /* EP2_3 (EP272) Only populate those fields that are available on the target screen */%>
				if (XMLActiveTag == null)
				{
					alert("Bank Address data not returned.");
				}
				else
				{
					if (frmScreen.txtStreet) frmScreen.txtStreet.value = BankSearchXML.GetTagText("Address1");
					if (frmScreen.txtDistrict) frmScreen.txtDistrict.value = BankSearchXML.GetTagText("Address2");
					if (frmScreen.txtTown) frmScreen.txtTown.value = BankSearchXML.GetTagText("Address3");
					if (frmScreen.txtCounty) frmScreen.txtCounty.value = BankSearchXML.GetTagText("Address4");
					if (frmScreen.txtPostcode) frmScreen.txtPostcode.value = BankSearchXML.GetTagText("PostCode");
				}
			}
		}
		else
		{
			if(m_sScreenId == "PP100")
			{
				<%/*Maintain Payee details*/%>
				if (XMLActiveTag != null)
				{
					if (BankSearchXML.GetTagText("BankName") == "")
					{
						if (frmScreen.txtBankName) frmScreen.txtBankName.value = "";
						alert("Bank Name data not returned.");
					}
					else
					{
						if (frmScreen.txtBankName) frmScreen.txtBankName.value = BankSearchXML.GetTagText("BankName");
					}
				}
			}
		}

		m_bBankAccountValid = true;
	}
	else
	{
		m_bBankAccountValid = false;
	}
		 
	BankSearchXML = null;
	
	<% /* MO - 08/11/2002 - BMIDS00726 
	return(true); */ %>
	return(m_bBankAccountValid);
	
}


</script>
