<%@ Control Language="c#" AutoEventWireup="false" Codebehind="PropertyDetails.ascx.cs" Inherits="Epsom.Web.WebUserControls.PropertyDetails" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<fieldset>
      <legend>Property Details</legend>
      <table class=formtable>
        
        <tr>
          <td class=prompt><span class=mandatoryfieldasterisk>*</span>Property Location</td>
          <td>
            <asp:dropdownlist id="ddlCategory" runat="server" CausesValidation="no" autopostback="false"></asp:dropdownlist></td>
          <td rowspan="3"> 
          </td>
        </tr>
      
        <asp:Repeater id="rptDepositSource" runat="server" >
        <ItemTemplate>
          <tr>
            <td class=prompt><span class=mandatoryfieldasterisk>*</span>Deposit source type</td>
            <td><asp:dropdownlist  id="ddlDepositSourceType"  runat="server" CausesValidation="no" autopostback="false"  DataMember="HttpSession.CurrentApplication.DepositSources.DepositSourceType"></asp:dropdownlist></td>
            <td class=prompt>Amount &nbsp;&nbsp;<asp:textbox id="txtAmount"  runat="server" DataMember="HttpSession.CurrentApplication.DepositSources.Amount"></asp:textbox></td>
            <td class=prompt></td>
          </tr>
         </ItemTemplate>
        </asp:Repeater>
    
        <tr>
          <td>&nbsp;</td>
          <td class="button_orphan"><asp:button id="btnAddDepositSource" runat="server" enabled="True" text="Add another deposit source"></asp:button></td>
          <td class="button_orphan"><asp:button id="btnRemoveDepositSource" runat="server" enabled="True" text="Remove a deposit source"></asp:button></td>  
        </tr>
    
    <asp:Panel id="pnlRentalIncome"  runat="server"></asp:Panel>
        <tr>
          <td class=prompt></span>Anticipated rental income</td>
          <td><asp:textbox id="txtRentalIncome" runat="server"></asp:textbox></td>
        </tr>   
      </table>
    </fieldset> 
