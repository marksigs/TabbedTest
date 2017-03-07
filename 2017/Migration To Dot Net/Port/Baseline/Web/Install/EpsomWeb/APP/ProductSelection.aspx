<%@ Page language="c#" Codebehind="ProductSelection.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.ProductSelection" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/ProductSelection.ascx" TagName="ProductSelection" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
	<cms:helplink id="helplink" class="helplink" runat="server" helpref="2040" />
    <h1><cms:cmsBoilerplate cmsref="604" runat="server" /></h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />

    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" cssclass="validationsummary" showsummary="true" id="Validationsummary1" />
       
    <epsom:productselection id="ctlProductSelection" runat="server" DisplayOnly="true"/>
       
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" runat="server"></epsom:dipnavigationbuttons>

    <sstchur:smartscroller id="ctlSmartScroller" runat="server" />


  </mp:content>

</mp:contentcontainer>

