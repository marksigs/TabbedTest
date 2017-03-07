;<%@ Page language="c#" Codebehind="AdditionalInfo.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.AdditionalInfo" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/AdditionalInfoSummary.ascx" TagName="AdditionalInfoSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/AdditionalInfo.ascx" TagName="AdditionalInfo" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
	<cms:helplink id="helplink" class="helplink" runat="server" helpref="2070" />
	
    <h1><cms:cmsBoilerplate cmsref="607" runat="server" /></h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>

    <epsom:additionalinfo id="Additionalinfo1" runat="server"></epsom:additionalinfo>
  
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" buttonbar="true" savepage="true" saveandclosepage="true" runat="server"></epsom:dipnavigationbuttons>

  </mp:content>

</mp:contentcontainer>

