<%@ Control Language="c#" AutoEventWireup="false" Codebehind="LoginAndRegistrationMenu.ascx.cs" Inherits="Epsom.Web.WebUserControls.LoginAndRegistrationMenu" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/LogOnDetails.ascx" TagName="LogOnDetails" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<epsom:LogOnDetails id="ctlLogOnDetails" runat="server"></epsom:LogOnDetails>

<asp:Panel id="pnlUnrestrictedItems" Runat="server">

  <h2>Brokers</h2>

  <p>
    <cms:cmsBoilerplate cmsref="4013" runat="server" ID="Cmsboilerplate1"/>
  </p>
  <p>
    <a href="<%=ResolveUrl("~/Profile/Register.aspx")%>" title="" class="fancylink">Register with us</a>
  <p/>
	
  <BR>

  <p>
    <cms:cmsBoilerplate cmsref="4014" runat="server" ID="Cmsboilerplate2"/>
  </p>
  <p>
    <a href="<%=ResolveUrl("~/Profile/Register.aspx")%>" title="" class="fancylink">Register with us</a>
  </p>

  <h2>Packagers</h2>

  <p>
    <cms:cmsBoilerplate cmsref="4015" runat="server" ID="Cmsboilerplate3"/>
  </p>
  <p>
      <a href="<%=ResolveUrl("~/Profile/RegisterPackager.aspx")%>" title="" class="fancylink">Register with us</a>
  </p>

</asp:Panel>
