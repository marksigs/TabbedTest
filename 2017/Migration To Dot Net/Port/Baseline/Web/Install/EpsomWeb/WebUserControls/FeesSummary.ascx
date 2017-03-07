<%@ Control Language="c#" AutoEventWireup="false" Codebehind="FeesSummary.ascx.cs" Inherits="Epsom.Web.WebUserControls.FeesSummary" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<h2 class="form_heading">Fees</h2>

<table width="100%">
    <tr>
      <td class="prompt"><B>Identifier</B></td>
      <td class="prompt"><B>Amount</B></td>
      <td class="prompt"><B>When Payable</B></td>
      <td class="prompt"><B>Refundable</B></td>
    </tr>
      
      
     <tr>
     &nbsp; 
    
    </tr>
      
      
<asp:repeater id="repeaterFees" runat="server">
  <itemtemplate>
  <asp:panel id="pnlFee">
    <tr>
      <td><asp:label id="labelIdentifier" runat="server" /></td>
      <td><asp:label id="labelAmount" runat="server" /></td>
      <td><asp:label id="labelWhenPayable" runat="server" /></td>
      <td><asp:label id="labelRefundable" runat="server" /></td>
    </tr>
  </asp:panel>
  </itemtemplate>
</asp:repeater>
</table>
