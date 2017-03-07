<%@ Page language="c#" Codebehind="AdviserDetails.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Dip.AdviserDetails" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/Adviser.ascx" TagName="Adviser" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>


<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server" />

  </mp:content>

  <mp:content id="region2" runat="server">
    <cms:helplink id="helplink" class="helplink" runat="server" helpref="2010" />
    
    <h1><cms:cmsBoilerplate cmsref="601" runat="server" /></h1>
		<cms:cmsBoilerplate cmsref="10" runat="server" />
		
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />

    <epsom:adviser id="ctlAdviser" runat="server" />

    <epsom:dipnavigationbuttons buttonbar="true" id="ctlDIPNavigationButtons2" runat="server" />

  </mp:content>

</mp:contentcontainer>
