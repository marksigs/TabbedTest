<%@ Page language="c#" Codebehind="Default.aspx.cs" AutoEventWireup="false" 
  SmartNavigation="false" Inherits="Epsom.Web.Kfi.Default" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">

    <h1>DIP</h1>

    Welcome to the DIP intro page.

  </mp:content>

</mp:contentcontainer>
