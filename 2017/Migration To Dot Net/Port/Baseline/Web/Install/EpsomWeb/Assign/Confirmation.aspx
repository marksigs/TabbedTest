<%@ Page language="c#" Codebehind="Confirmation.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Assign.Confirmation" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">
<mp:content id="region1" runat="server">
<epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>
</mp:content>
<mp:content id="region2" runat="server">
<h1>Assigned Work Confirmation</h1>
<epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>
<asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />

<p>
Your application has been assigned to <asp:Label id="lblPackagerName" runat="server"></asp:label>.
</p>
<asp:button id="cmdHome" runat="server" text="Home"></asp:button>
<table class="formtable">
<tr>
</tr>
</table>
<epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" runat="server"></epsom:dipnavigationbuttons>
</mp:content>
</mp:contentcontainer>
