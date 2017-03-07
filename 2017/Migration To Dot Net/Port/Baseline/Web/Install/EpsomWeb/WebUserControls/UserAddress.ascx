<%@ Control Language="c#" AutoEventWireup="false" Codebehind="UserAddress.ascx.cs" Inherits="Epsom.Web.WebUserControls.UserAddressDeprecated" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<H2 class=form_heading>Correspondence address &amp; contact numbers</H2>
<!-- break the table into several sections so we can show-hide appropriate sections -->
<asp:panel runat="server" id="pnlCountry">
 <table class="formtable">
  <tr>
    <td class="prompt"><asp:label id="lblCountryType" runat="server" text="Country"  /></td>
    <td><asp:dropdownlist id="cmbCountryType" runat="server" AutoPostBack="True" /></td>
  </tr>
 </table>
</asp:Panel> 

<epsom:Address id="ctlAddress" runat="server"></epsom:Address>
