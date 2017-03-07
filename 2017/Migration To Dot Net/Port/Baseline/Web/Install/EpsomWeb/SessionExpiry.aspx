<%@ Page language="c#" Codebehind="SessionExpiry.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.SessionExpiry" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>

<epsom:dipmenu id="DipMenu" runat="server" />

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="./masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">
  
    <h1>Session expired</h1>

    <p>Your session has expired. Please <a href=<%=ResolveUrl(logonURL)%>>Log In</a> to continue.</p>

  </mp:content>
 
</mp:contentcontainer>
  
