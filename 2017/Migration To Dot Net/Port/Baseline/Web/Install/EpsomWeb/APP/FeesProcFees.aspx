<%@ Page language="c#" Codebehind="FeesProcFees.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.FeesProcFees" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/FeesProcFees.ascx" TagName="FeesProcFees"   TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
	 <cms:helplink id="helplink" class="helplink" runat="server" helpref="3500" />
    <h1><cms:cmsBoilerplate cmsref="635" runat="server" /></h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>

    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />
    
    <epsom:feesprocfees id="ctlFeesProcFees" runat="server"></epsom:feesprocfees>
      
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" buttonbar="true" savepage="true" saveandclosepage="true" runat="server"></epsom:dipnavigationbuttons>
   
  </mp:content>

</mp:contentcontainer>

