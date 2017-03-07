<%@ Page language="c#" Codebehind="LandlordDetails.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.LandlordDetails" %>
<%@ Register Src="~/WebUserControls/LandlordDetails.ascx" TagName="LandLordDetails" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
<cms:helplink id="helplink" class="helplink" runat="server" helpref="41000" />
    <h1>Landlord details</h1>
    
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>

    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />
    
    <cms:cmsBoilerplate cmsref="100" runat="server" />
    
    <epsom:LandlordDetails id="ctlLandlordDetails" runat="server"/>
 
    <epsom:dipnavigationbuttons buttonbar="true" savepage="true" saveandclosepage="true" id="ctlDIPNavigationButtons2" runat="server" />
    
    <sstchur:SmartScroller id="ctlSmartScroller" runat="server" />
    
  </mp:content>

</mp:contentcontainer>
