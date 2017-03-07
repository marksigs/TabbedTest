<%@ Page language="c#" Codebehind="ProductUpdate.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.ProductUpdate" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>News</h1>
    
    <h2 class="form_heading"></h2>

        
    <table width="100%" border="0">
    
    <tr><th> <asp:label id="labelHeadline" runat="server" /></th></tr>
    
    <tr>
    <td><asp:label id="labelBodyText" runat="server" /></td>
    <td valign="top" rowspan="3">
    <!--<asp:hyperlink id="linkProductImage" valign="top" imageurl="~/_gfx/arrow.gif" runat="server"  />-->
    </td>
    </tr>
    <tr><td><asp:hyperlink id="linkBack" cssclass="fancylink" runat="server" text="Back" ></asp:hyperlink></td></tr>
    </table>
    
		
  </mp:content>

</mp:contentcontainer>