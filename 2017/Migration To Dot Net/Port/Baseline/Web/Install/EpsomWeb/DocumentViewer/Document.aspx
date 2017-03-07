<%@ Page language="c#" Codebehind="Document.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.DocumentViewer.Document" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>

<epsom:dipmenu id="DipMenu" runat="server" />

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">
  
    <h1>View document</h1>

    <p class="error">
		  <asp:Literal ID="ltlError" runat="server" />
    </p>

		<br />

		<asp:Label ID="DebugLabel" runat="server" Visible="False"></asp:Label>

  </mp:content>
 
</mp:contentcontainer>
  
