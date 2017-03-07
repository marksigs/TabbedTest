<%@ Page language="c#" Codebehind="Commitments.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Dip.Commitments" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/Commitments.ascx" TagName="commitments" TagPrefix="Epsom" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>


<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server" />

  </mp:content>

  <mp:content id="region2" runat="server">
  <cms:helplink id="helplink" class="helplink" runat="server" helpref="2160" />

    <h1><asp:Label id="lblHeading" runat="server" Text="Commitments" /></h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />
    
    <asp:validationsummary id="Validationsummary1" runat="server" showsummary="true" CssClass="validationsummary"
      headertext="Input errors:" displaymode="BulletList" />
    
    <epsom:commitments id="ctlCommitments" runat="server" ShowMortgageCommitments="True"/>
 
    <epsom:dipnavigationbuttons buttonbar="true" id="ctlDIPNavigationButtons2" runat="server" />
    
    <sstchur:SmartScroller id="ctlSmartScroller" runat="server" />

  </mp:content>

</mp:contentcontainer>


