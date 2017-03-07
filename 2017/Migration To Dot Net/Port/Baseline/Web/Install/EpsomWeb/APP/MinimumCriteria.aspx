<%@ Page language="c#" Codebehind="MinimumCriteria.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.MinimumCriteria" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/MinimumCriteria.ascx" TagName="MinimumCriteria" TagPrefix="Epsom" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
	<cms:helplink id="helplink" class="helplink" runat="server" helpref="2020" />
    <h1><cms:cmsBoilerplate cmsref="602" runat="server" /></h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />

    <asp:validationsummary id=Validationsummary1 runat="server" displaymode="BulletList" showsummary="true" headertext="Input errors:" CssClass="validationsummary"></asp:validationsummary>

    <epsom:minimumcriteria id="ctlMinimumCriteria" runat="server"></epsom:minimumcriteria>

    <epsom:dipnavigationbuttons buttonbar="true" id="ctlDIPNavigationButtons2" runat="server" />

    <sstchur:smartscroller id="ctlSmartScroller" runat="server" />

  </mp:content>

</mp:contentcontainer>
