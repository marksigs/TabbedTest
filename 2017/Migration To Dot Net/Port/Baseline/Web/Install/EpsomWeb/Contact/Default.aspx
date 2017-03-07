<%@ Page language="c#" Codebehind="Default.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Contact.Default" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">
    <h1>Contact Us</h1>

    <p>
      <CMS:Literal runat="server" cmsref="110" />
    </p>
    
    <table>
    <tr>
    <td>

    <h2 class="form_heading">Do you need support?</h2>
      <p>
        <CMS:Literal runat="server" cmsref="111" />
        <asp:LinkButton id="cmdSupport" text="Contact us for Support" runat="server"></asp:LinkButton>
      </p>
    </td>
    <td>&nbsp;</td>
    <td>
      <h2 class="form_heading">Are you looking for a BDM?</h2>
      <p>
        <CMS:Literal runat="server" cmsref="112" />
        <asp:HyperLink id="lnkItem" runat="server" NavigateUrl="BusinessDevelopmentManagerFinder.aspx" text="Contact our BDMs"></asp:HyperLink>
      </td>
    </tr>
  </table>
  </mp:content>

</mp:contentcontainer>
