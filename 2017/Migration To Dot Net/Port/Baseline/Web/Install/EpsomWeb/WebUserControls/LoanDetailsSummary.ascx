<%@ Control Language="c#" AutoEventWireup="false" Codebehind="LoanDetailsSummary.ascx.cs" Inherits="Epsom.Web.WebUserControls.LoanDetailsSummary" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<fieldset>

  <legend>Loan details</legend>
  
  <table class="displaytable">
    <tr>
      <td class="prompt">Category</td>
      <td><asp:label id="lblCategory" runat="server" /></td>
      <td class="prompt">Credit limit</td>
      <td><asp:label id="lblCreditLimit" runat="server" /></td>
    </tr>
    
    <tr>
      <td class="prompt">Type of application</td>
      <td><asp:label id="lblTypeOfApplication" runat="server" /></td>
      <td class="prompt">Amount requested</td>
      <td><asp:label id="lblLoanAmountRequested" runat="server" /></td>
    </tr>
    
    <tr>
      <td class="prompt">Type of buyer</td>
      <td><asp:label id="lblTypeOfBuyer" runat="server" /></td>
      <td class="prompt" rowspan="2">Purchase price/estimated value</td>
      <td><asp:label id="lblPurchasePriceOrEstimatedValue" runat="server" /></td>
    </tr>
    
    <tr>
      <td class="prompt">Property Location</td>
      <td><asp:label id="lblPropertyLocation" runat="server" /></td>
    </tr>
   
    <asp:Repeater id="rptDepositSource" runat="server">
      <ItemTemplate>
      <tr>
        <td class="prompt">Deposit Source</td>
        <td><asp:label id="lblDepositSource" runat="server" /></td>
      </tr>
      
      <tr>
        <td class="prompt">Amount</td>
        <td><asp:label id="lblDepositAmount" runat="server" /></td>
      </tr>
      </ItemTemplate>
    </asp:Repeater>
          
    <tr>
      <td class="prompt">Antipated Rental Income</td>
      <td><asp:label id="lblAntipatedRentalIncome" runat="server" /></td>
    </tr>
    
    <tr>
      <td class="prompt">Scheme</td>
      <td><asp:label id="lblScheme" runat="server" /></td>
    </tr>
    
    <tr>
      <td class="prompt">Status</td>
      <td><asp:label id="lblStatus" runat="server" /></td>
      <td class="prompt"><asp:label id="lblDiscountCaption" runat="server" /></td>
      <td><asp:label id="lblDiscount" runat="server" /></td>
    </tr>     
    
    <asp:Panel id="pnlFlexiSurpress" runat="server" Visible="True">
    <tr>
      <td class="prompt">Flexi features</td>
      <td><asp:label id="lblFlexiFeatures" runat="server" /></td>
      <td class="prompt"><asp:label id="lblCalculatedDiscountCaption" runat="server" /></td>
      <td><asp:label id="lblCalculatedDiscount" runat="server" /></td>

    </tr>
    
    <tr>
      <td class="prompt">Product features</td>
      <td><asp:label id="lblProductClassFeatures" runat="server" /></td>
      <td class="prompt"><asp:label id="lblAllocatedToComponentsCaption" runat="server" /></td>
      <td><asp:label id="lblAllocatedToComponents" runat="server" /></td>
    </tr>
    </asp:Panel>
  </table>
<!--
  <hr />

  <table class="displaytable">
    <tr>
      <td style="text-align:right"><asp:Button ID="cmdEditLoan" commandName="cmdEditLoan" runat="server" text="Edit" CssClass="button" /></td>
    </tr>
  </table>
-->
</fieldset>

