<%@ Page language="c#" Codebehind="ProductSummary.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Dip.ProductSummary" %>
<%@ Register Src="~/WebUserControls/ProductSummary.ascx" TagName="ProductSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">

    <h1>Product summary</h1>
    <p>
      <CMS:Literal runat="server" cmsref="50">******** DEFAULT NON-CMS TEXT *********</CMS:Literal>
    </p>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>

    <epsom:productsummary id="ctlProductSummary" runat="server" />
    <p>
		<asp:Button ID="cmdAddComponent" runat="server" text="Add another product" />    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" runat="server"></epsom:dipnavigationbuttons>
    </p>
  </mp:content>

</mp:contentcontainer>

