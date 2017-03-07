<%@ Page language="c#" Codebehind="LoanRequirements.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Dip.LoanRequirements" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/LoanRequirements.ascx" TagName="LoanRequirements" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>


<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server" />

  </mp:content>
  
  <mp:content id="region2" runat="server">
  <cms:helplink id="helplink" class="helplink" runat="server" helpref="40000" />
 
    <h1><cms:cmsBoilerplate cmsref="603" runat="server" /></h1>
    
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />
    
    <asp:validationsummary id="Validationsummary1" runat="server" displaymode="BulletList" showsummary="true" headertext="Input errors:" CssClass="validationsummary" />
    
    <epsom:LoanRequirements id="ctlLoanRequirements" runat="server" />
    
    <epsom:dipnavigationbuttons buttonbar="true" id="ctlDIPNavigationButtons2" runat="server" />

    <sstchur:SmartScroller id="ctlSmartScroller" runat="server" />
    
  </mp:content>

</mp:contentcontainer>
