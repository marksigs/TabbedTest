<%@ Page language="c#" Codebehind="News.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.News" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>News and product updates</h1>
   
    <h2 class="form_heading">Product Updates</h2>
    <table>
    <tr>
      <td rowspan="7" valign="top">
      <!--<asp:hyperlink id="linkProductImage" valign="top" imageurl="~/_gfx/arrow.gif" runat="server"  />-->
      </td>
      <td class="newsheading"><asp:label id="lblProductHeading" runat="server" /></td>
      </tr>
      <tr><td>&nbsp;</td></tr>
      <tr>
          <td>
        <asp:label id="lblProductDescription" runat="server" />
        <asp:hyperlink id="linkProductMore" cssclass="fancylink" navigateurl="article.aspx" text="> More" runat="server" />
        </td>
      </tr>
    
     <tr><td>&nbsp;</td></tr>
    
     <tr><td><asp:hyperlink id="linkProductNext1" cssclass="fancylink" navigateurl="articles.aspx" runat="server" /></td>
     </tr>
     <tr><td><asp:hyperlink id="linkProductNext2" cssclass="fancylink"  navigateurl="articles.aspx" runat="server" /></td>
     </tr>
     <tr><td><asp:hyperlink id="linkProductArchive" cssclass="fancylink" navigateurl="productUpdates.aspx" runat="server" text="Archived product updates"/></td>
     </tr>

    </table>


    <!--
    
    <table>
    
    <asp:Repeater id="rprProductUpdates" Runat="server">
      <ItemTemplate>
        <tr>
          <td colspan="2" class="newsheading">
            <asp:label id="lblProductHeading" runat="server"></asp:label>
          </td>
        </tr>
        <tr>
          <td>IMAGE</td>
          <td>
            <asp:label id="lblProductDescription" runat="server"></asp:label>
          </td>
        </tr>
      </ItemTemplate>
    </asp:Repeater>
    
    </table-->
    
    <h2 class="form_heading">News</h2>
    
    <table>
    <tr>
      <td rowspan="7" valign="top"><asp:hyperlink id="linkNewsImage" valign="top" imageurl="~/_gfx/arrow.gif" runat="server"  /></td>
      <td class="newsheading"><asp:label id="lblNewsHeading" runat="server" /></td>
      </tr>
      <tr><td>&nbsp;</td></tr>
      <tr>
        <td>
        <asp:label id="lblNewsDescription" runat="server" />
        <asp:hyperlink id="linkNewsMore" cssclass="fancylink" navigateurl="article.aspx" text="> More" runat="server" />
        </td>
      </tr>
    
     <tr><td>&nbsp;</td></tr>
    
     <tr><td><asp:hyperlink id="linkNewsNext1" cssclass="fancylink" navigateurl="article.aspx" runat="server" /></td>
     </tr>
     <tr><td><asp:hyperlink id="linkNewsNext2" cssclass="fancylink" navigateurl="article.aspx" runat="server" /></td>
     </tr>
     <tr><td><asp:hyperlink id="linkNewsArchive" cssclass="fancylink" navigateurl="articles.aspx" runat="server" text="Archived news"/></td>
     </tr>

    </table>
    
  </mp:content>

</mp:contentcontainer>
