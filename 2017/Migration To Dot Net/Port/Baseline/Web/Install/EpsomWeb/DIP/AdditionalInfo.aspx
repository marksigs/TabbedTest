<%@ Page language="c#" Codebehind="AdditionalInfo.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Dip.AdditionalInfo" %>
<%@ Register Src="~/WebUserControls/AdditionalInfo.ascx" TagName="AdditionalInfo" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>


<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server" />

  </mp:content>

  <mp:content id="region2" runat="server">
  <cms:helplink id="helplink" class="helplink" runat="server" helpref="2070" />
  
    <h1><cms:cmsBoilerplate cmsref="607" runat="server" /></h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />

    <epsom:additionalinfo id="ctlAdditionalInfo" runat="server" />

    <epsom:dipnavigationbuttons buttonbar="true" id="ctlDIPNavigationButtons2" runat="server" />

  </mp:content>

</mp:contentcontainer>


