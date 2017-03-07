<%@ Control Language="c#" AutoEventWireup="false" Codebehind="LenderDetails.ascx.cs" Inherits="Epsom.Web.WebUserControls.LenderDetails" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/LenderDetail.ascx" TagName="LenderDetail" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<asp:Repeater id="rptLenderDetails" Runat="server">
  <ItemTemplate>
    <h2 class="form_heading">Mortgage commitment <asp:Literal ID="lblMortgageCommitmentIndex" Runat="server" /></h2>
    <epsom:LenderDetail id="ctlLenderDetail" runat="server"/>
  </ItemTemplate>
</asp:Repeater>
  <table>
  <tr><td class="displaytable"><asp:label class="prompt" ID="lblMortgageText" Runat="server"><cms:cmsBoilerplate cmsref="280" runat="server" /></asp:label></td><tr>
  </table>
  </BR>
  <p class="button_orphan"> 
    <asp:Button ID="cmdRemoveCommitment" Runat="server" Text="Remove"></asp:Button>
    <asp:Button ID="cmdAddCommitment" Runat="server" Text="Add mortgage commitment"></asp:Button>
  </p>
