<%@ Page language="c#" Codebehind="ProductUpdates.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.ProductUpdates" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>Product Updates Archive</h1>
    <h2 class="form_heading">&nbsp;</h2>
      
<asp:repeater id="repeaterArchive" runat="server">
  <itemtemplate>
      <p><asp:label id="labelDate" runat="server" /></p>
      <p><asp:hyperlink id="linkNews" cssclass="fancylink" navigateurl="article.aspx" runat="server" /></p>
  </itemtemplate>
</asp:repeater>
<p>Go to page <asp:repeater id="repeaterPageNumber" runat="server">
<itemtemplate><asp:hyperlink id="linkPageNumber" navigateurl="article.aspx" runat="server" />&nbsp;</itemtemplate>
</asp:repeater>
</p>
<p><asp:hyperlink id="linkBack" cssclass="fancylink" runat="server" text="Back to headlines" navigateurl="~/news/news.aspx"></asp:hyperlink></p>
  
  </mp:content>

</mp:contentcontainer>
