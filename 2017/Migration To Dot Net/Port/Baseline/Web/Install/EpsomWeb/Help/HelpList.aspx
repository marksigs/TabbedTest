<%@ Page language="c#" Codebehind="HelpList.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.HelpList" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>Help Topics</h1>
    
    <ul>
    <asp:repeater id="repeaterHelp" runat="server">
    <itemtemplate>
   
      <li><cms:helplink id="helplink" style="font-weight:bold" runat="server" /></li>
    
    </itemtemplate>
    </asp:repeater>
    </ul>

		
  </mp:content>

</mp:contentcontainer>
