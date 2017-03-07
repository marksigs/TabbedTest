<%@ Page language="c#" Codebehind="FeesProcFees2.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.FeesProcFees2" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">

    <h1>Fees and proc fees 2</h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>

    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />
    
    Fees and proc fees: <asp:TextBox ID="txtFees" Runat="server" /> 
      
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" buttonbar="true" savepage="true" saveandclosepage="true" runat="server"></epsom:dipnavigationbuttons>
   
  </mp:content>

</mp:contentcontainer>

