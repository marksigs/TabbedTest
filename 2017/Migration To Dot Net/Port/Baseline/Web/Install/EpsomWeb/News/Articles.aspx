<%@ Page language="c#" Codebehind="Articles.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Articles" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>News Archive</h1>
      
    <asp:repeater id="repeaterArchive" runat="server">
      <itemtemplate>
        <p><asp:label id="labelDate" runat="server" /></p>
        <p><asp:hyperlink id="linkNews" cssclass="fancylink" navigateurl="article.aspx" runat="server" /></p>
      </itemtemplate>
    </asp:repeater>

    <br />

    <p>
      Go to page
      <asp:repeater id="repeaterPageNumber" runat="server">
        <itemtemplate><asp:hyperlink id="linkPageNumber" navigateurl="article.aspx" runat="server" />&nbsp;</itemtemplate>
      </asp:repeater>
    </p>

    <br />

    <p>
      <asp:hyperlink cssclass="fancylink" id="linkBack" runat="server" text="Back to headlines" navigateurl="~/news/news.aspx"></asp:hyperlink>
    </p>
  
  </mp:content>

</mp:contentcontainer>
