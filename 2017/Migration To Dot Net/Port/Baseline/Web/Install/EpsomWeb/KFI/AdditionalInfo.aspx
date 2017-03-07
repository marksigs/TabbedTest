<%@ Page language="c#" Codebehind="AdditionalInfo.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Kfi.AdditionalInfo" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
  <cms:helplink id="helplink" class="helplink" runat="server" helpref="2070" />


    <h1><cms:cmsBoilerplate cmsref="607" runat="server" /></h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>

    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />

    <cms:cmsBoilerplate cmsref="75" runat="server" />

    <table>
      <tr>
		    <td class="prompt"><cms:cmsBoilerplate cmsref="76" runat="server" /></td>
		    <td><asp:Textbox id="txtAdditionalInformation" runat="server" Width="272px" Height="176px" TextMode="MultiLine"/></td>
	    </tr>
	  </table>
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" runat="server"></epsom:dipnavigationbuttons>

  </mp:content>

</mp:contentcontainer>

