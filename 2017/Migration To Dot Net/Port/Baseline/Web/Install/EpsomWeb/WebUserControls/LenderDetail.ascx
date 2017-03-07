<%@ Control Language="c#" AutoEventWireup="false" Codebehind="LenderDetail.ascx.cs" Inherits="Epsom.Web.WebUserControls.LenderDetail" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="TelephoneNumbers.ascx" TagName="TelephoneNumbers" TagPrefix="Epsom" %>
<%@ Register Src="Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/ApplicantType.ascx" TagName="applicantType" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>

<epsom:applicanttype id="ctlApplicantType" runat="server" />
<table  class="formtable">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Name of lender</td>
    <td>
      <asp:TextBox ID="txtLenderName" Runat="server"/>
      <asp:requiredfieldvalidator controltovalidate="txtLenderName" runat="server" text="***" 
      errormessage="Please enter lender name" id="Requiredfieldvalidator3" enableclientscript="false"/>
  </td>
  </tr>
</table>
  
  <epsom:address id="ctlLenderAddress" runat="server"  isresidential="false" />

  <epsom:TelephoneNumbers id="ctlLenderTelephoneNumbers" runat="server" usageComboGroup="ContactTelephoneUsage"/>
  
<table  class="formtable">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Applicant Account/roll number</td>
    <td>
      <asp:TextBox ID="txtAccountNumber" Runat="server" />
      <asp:requiredfieldvalidator controltovalidate="txtAccountNumber" runat="server" text="***" 
      errormessage="Please enter account number" id="Requiredfieldvalidator1" enableclientscript="false"/>
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Select the address this mortgage relates to or enter a new address</td>
    <td><asp:ListBox ID="lstSelectSecurityAddress" Runat="server" />
      <asp:customvalidator id="Customvalidator1" runat="server" enableclientscript="false"
        onservervalidate="ValidatorAddressMortgageRelatesTo"
        errormessage="Address mortgage relates to is required" text="***" />
    </td>
  </tr>
  <tr>
    <td class="prompt"></td>
    <td>
    <asp:Button ID="cmdAddSecurityAddress" Runat="server" Text="Add security address"/>
    <asp:Button ID="cmdResetSecurityAddress" Runat="server" Text="Reset"/>
    </td>
  </tr>
</table>

<asp:Panel ID="pnlAddedAddress" Runat="server">
  <epsom:address id="ctlSecurityAddress" runat="server"  isresidential="false" />
</asp:Panel>

<asp:Panel ID="pnlIsRemortgage" Runat="server">

  <table  class="formtable">
    <tr>
      <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Is this the account you wish to remortgage</td>
      <td>
            <epsom:nullableyesno id="nynRemortgageAccount" runat="server" required="true" autopostback="false" enabled="true"
            errormessage="Specify if the account is a remortgage" >
            </epsom:nullableyesno>
      </td>
    </tr>
  </table>
</asp:Panel>

<table  class="formtable">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Monthly payment</td>
    <td><asp:TextBox ID="txtMonthlyPayment" Runat="server" />
    <asp:requiredfieldvalidator controltovalidate="txtMonthlyPayment" runat="server" text="***" 
    errormessage="Please enter monthly payment" id="Requiredfieldvalidator2" enableclientscript="false"/>
    <asp:comparevalidator operator=DataTypeCheck type="Currency" id="CurrencyValidator1" runat="server" 
         controltovalidate="txtMonthlyPayment"
         text="***" errormessage="Monthly payment is invalid" enableclientscript="false" />
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Repayment Type</td>
    <td>
      <asp:DropDownList ID="cmbRepaymentType" Runat="server" />
      <validators:requiredselecteditemvalidator runat="server" enableclientscript="false"
          text="***" errormessage="Repayment Type is required" controltovalidate="cmbRepaymentType" 
          id="Requiredselecteditemvalidator2"/> 
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Balance outstanding</td>
    <td><asp:TextBox ID="txtBalanceOutstanding" Runat="server" />
    <asp:requiredfieldvalidator controltovalidate="txtBalanceOutstanding" runat="server" text="***" 
    errormessage="Please enter balance outstanding" id="Requiredfieldvalidator4" enableclientscript="false"/>
    <asp:comparevalidator operator=DataTypeCheck type="Currency" id="Comparevalidator1" runat="server" 
         controltovalidate="txtBalanceOutstanding"
         text="***" errormessage="Balance outstanding is invalid" enableclientscript="false" />
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Date mortgage started</td>
    <td><asp:TextBox ID="txtDateMortgageStarted" Runat="server" />
    <asp:requiredfieldvalidator controltovalidate="txtDateMortgageStarted" runat="server" text="***" 
    errormessage="Please enter date mortgage started" id="Requiredfieldvalidator5" enableclientscript="false"/>
    <validators:datevalidator id="DateValidator1" runat="server" controltovalidate="txtDateMortgageStarted" 
            text="***" errormessage="Invalid date mortgage started"
          enableclientscript="false" /> dd/mm/yyyy
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Is this a second charge</td>
    <td>
      <epsom:nullableyesno id="nynSecondCharge" runat="server" required="true" autopostback="true" Enabled="true"
            errormessage="Specify if the account is a second charge" >
            </epsom:nullableyesno>
    </td>
  </tr>
  <tr>
        <td  class="prompt">Number of months currently in arrears</td>
        <td>
          <asp:dropdownlist id="cmbMonthsInArrears" runat="server" height="13">
            <asp:listitem runat="server" value="0" id="Listitem22">-Select-</asp:listitem>
            <asp:listitem runat="server" value="1" id="Listitem23">1</asp:listitem>
            <asp:listitem runat="server" value="2" id="Listitem24">2</asp:listitem>
            <asp:listitem runat="server" value="3" id="Listitem25">3</asp:listitem>
            <asp:listitem runat="server" value="4" id="Listitem26">4</asp:listitem>
            <asp:listitem runat="server" value="5" id="Listitem27">5</asp:listitem>
            <asp:listitem runat="server" value="6" id="Listitem28">6</asp:listitem>
            <asp:listitem runat="server" value="7" id="Listitem29">7</asp:listitem>
            <asp:listitem runat="server" value="8" id="Listitem30">8</asp:listitem>
            <asp:listitem runat="server" value="9" id="Listitem31">9</asp:listitem>
            <asp:listitem runat="server" value="10" id="Listitem32">10</asp:listitem>
            <asp:listitem runat="server" value="11" id="Listitem33">11</asp:listitem>
            <asp:listitem runat="server" value="12" id="Listitem34">12</asp:listitem>
          </asp:dropdownlist>
        </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Original loan amount</td>
    <td>
      <asp:TextBox ID="txtOriginalLoanAmount" Runat="server" />
      <asp:requiredfieldvalidator controltovalidate="txtOriginalLoanAmount" runat="server" text="***" 
      errormessage="Please enter Original Loan Amount" id="Requiredfieldvalidator6" enableclientscript="false"/>
          <asp:comparevalidator operator=DataTypeCheck type="Currency" id="Comparevalidator2" runat="server" 
         controltovalidate="txtOriginalLoanAmount"
         text="***" errormessage="Original Loan Amount is invalid" enableclientscript="false" />
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Will this mortgage be redeemed on completion of the new mortgage?</td>
    <td>
      <asp:DropDownList ID="cmbRedemptionStatus" Runat="server" />
      <validators:requiredselecteditemvalidator runat="server" enableclientscript="false"
          text="***" errormessage="Redemption status is required" controltovalidate="cmbRedemptionStatus" id="Requiredselecteditemvalidator3"/> 
    </td>
  </tr>
</table>

<asp:Panel ID="pnlNotRedeemedDetails" Runat="server">

  <table  class="formtable">
    <tr>
      <td class="prompt">If to be rented out, what will the monthly rental income be</td>
      <td><asp:TextBox ID="txtRentalIncome" Runat="server" />
          <asp:comparevalidator operator=DataTypeCheck type="Currency" id="Comparevalidator3" runat="server" 
         controltovalidate="txtRentalIncome"
         text="***" errormessage="Rental income is invalid" enableclientscript="false" />
      </td>
    </tr>
  </table>

</asp:Panel>
  
<table  class="formtable">
  <tr>
    <td class="prompt">Do you own/partly own any other property or are you party to any other mortgage</td>
    <td>
      <asp:RadioButton ID="rdoOwnOtherPropertyYes" AutoPostBack="True" GroupName="grpOwnOtherProperty" Runat="server" Text="Yes" />
      <asp:RadioButton ID="rdoOwnOtherPropertyNo"  AutoPostBack="True" GroupName="grpOwnOtherProperty" Runat="server" Text="No" />
              <asp:customvalidator id="Customvalidator2" runat="server" enableclientscript="false"
        onservervalidate="ValidatorOwnOtherPropertyRequired"
        errormessage="You must say if you own any other property" text="***" />

    </td>
  </tr>
</table>

<asp:Panel ID="pnlPartlyOwnDetails" Runat="server">

  <table  class="formtable">
    <tr>
      <td class="prompt"></td>
      <td><asp:TextBox id="txtOwnOtherPropertyDetails" size=30"" Runat="server" TextMode="MultiLine"/></td>
    </tr>
  </table>

</asp:Panel>