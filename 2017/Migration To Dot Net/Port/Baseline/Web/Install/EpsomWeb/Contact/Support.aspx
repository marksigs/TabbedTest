<%@ Page language="c#" Codebehind="Support.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Contact.Support" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Register Src="~/WebUserControls/Documents.ascx" TagName="Documents" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">
<mp:content id="region1" runat="server">

<h1 class="form_heading">Support</h1>

<p><cms:cmsBoilerplate cmsref="130" runat="server" /></p>
<p><asp:LinkButton id="cmdBdmFinder" runat="server" CssClass="fancylink" text="BDM Finder" /></p>
<epsom:documents id="Documents" documentcategory="SupportForm" runat="server" />
<br>
<hr>
<br>
<asp:panel id="pnlLoggedOut" runat="server">
  <cms:cmsBoilerplate cmsref="132" runat="server" />
</asp:panel>


<asp:panel id="pnlLoggedIn" runat="server">
 <cms:cmsBoilerplate cmsref="133" runat="server"/>
</asp:panel>


</mp:content>
</mp:contentcontainer>
