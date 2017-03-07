<%@ Page language="c#" Codebehind="Article.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.article" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>News</h1>
    
    <table width="100%" border="0">
    
    <tr>
      <th style="padding: 0.2em 0.4em"> <asp:label id="labelHeadline" runat="server" /></th>
    </tr>

    <tr>
      <td>&nbsp;</td>
    </tr>
    
    <tr>
      <td><asp:label id="labelBodyText" runat="server" /></td>
      <td valign="top" rowspan="3"><asp:hyperlink id="linkNewsImage" valign="top" imageurl="~/_gfx/arrow.gif" runat="server" style="padding: 0.5em" /></td>
    </tr>
    <tr>
      <td><asp:hyperlink cssclass="fancylink" id="linkBack" runat="server" text="Back" ></asp:hyperlink></td>
    </tr>
    </table>
    
		
  </mp:content>

</mp:contentcontainer>