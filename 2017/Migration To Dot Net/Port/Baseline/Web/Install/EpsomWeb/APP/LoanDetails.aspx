<%@ Page language="c#" Codebehind="LoanDetails.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.LoanDetails" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">
<mp:content id="region1" runat="server">
<epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>
</mp:content>
<mp:content id="region2" runat="server">
<h1>Loan details</h1>
<epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>
<asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />
<table class="formtable">
<tr>
  <td class=prompt>Loan type:</td>
  <td><asp:DropDownList id="cmbLoanType" Runat="server"></asp:DropDownList></td>
</tr>
</table>
<asp:button id=cmdSave runat="server" text="Save"></asp:button>
<asp:button id=cmdSaveAndClose runat="server" text="Save and Close"></asp:button>
<asp:button id=cmdCancel runat="server" text="Cancel"></asp:button>
<epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" runat="server"></epsom:dipnavigationbuttons>
</mp:content>
</mp:contentcontainer>

