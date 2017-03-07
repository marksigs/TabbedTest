<%@ Control Language="c#" AutoEventWireup="false" Codebehind="ResidentAddress.ascx.cs" Inherits="Epsom.Web.WebUserControls.ResidentAddressDeprecated" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<%@ Register Src="Address.ascx" TagName="Address" TagPrefix="Epsom" %>

<!-- break the table into several sections so we can show-hide appropriate sections -->
<asp:panel runat="server" id="pnlCountryOrBFPO">
 <table class="formtable">
  <tr>
    <td class="prompt"><asp:label id="lblBritishForcesPostOffice" runat="server" text="BFPO" /></td>
    <td><asp:checkbox id="chkBritishForcesPostOffice" runat="server" AutoPostBack="True"/></td>
  </tr>
  <tr>
    <td class="prompt"><asp:label id="lblCountryType" runat="server" text="Country"  /></td>
    <td><asp:dropdownlist id="cmbCountryType" runat="server" AutoPostBack="True" /></td>
  </tr>
 </table>
</asp:Panel> 
<epsom:Address id="ctlAddressUK" runat="server"></epsom:Address>
<asp:panel runat="server" id="pnlResidence">
  <table class="formtable">
  <tr>
    <td class="prompt">Resident from * mm/yyyy [date moved in]</td>
    <td>
      <asp:DropDownList id="cmbResidentFromMonth" runat="server" AutoPostBack="True"/>
      <asp:DropDownList id="cmbResidentFromYear" runat="server" AutoPostBack="True"/>
      <asp:CustomValidator ID="cmbResidentFromValidator" Runat="server" ErrorMessage="Invalid resident from date"
        EnableClientScript="false" OnServerValidate="ResidentFromValidator" />
    </td>
  </tr>
  </table>
</asp:Panel> 
