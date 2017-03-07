<%@ Page language="c#" Codebehind="FeesAndProcFees.aspx.cs" AutoEventWireup="false" SmartNavigation="false"  
Inherits="Epsom.Web.Kfi.FeesAndProcFees" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/FeesProcFees.ascx" TagName="FeesProcFees"   TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">
    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>
  </mp:content>
  
  <mp:content id="region2" runat="server">

    <cms:helplink id="helplink" class="helplink" runat="server" helpref="3500" />

    <h1>Fees and Proc Fees</h1>
    
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>
    
    <epsom:feesprocfees id="ctlFeesProcFees" runat="server"></epsom:feesprocfees>
    
    <epsom:dipnavigationbuttons buttonbar="true" id="ctlDIPNavigationButtons2" runat="server"></epsom:dipnavigationbuttons>
    
  </mp:content>
  
</mp:contentcontainer>
