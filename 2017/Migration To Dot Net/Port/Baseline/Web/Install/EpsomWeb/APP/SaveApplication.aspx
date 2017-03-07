<%@ Page language="c#" Codebehind="SaveApplication.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.SaveApplication" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>

<epsom:dipmenu id="DipMenu" runat="server" />

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">
  
    <h1>Application stopped</h1>

    <p class="dipnavigationbuttons">
		  <asp:Button ID="cmdContinue" runat="server" text="Continue" />
    </p>

    <p>
		  <asp:Label ID="lblMessage" runat="server" />
    </p>

  </mp:content>
 
</mp:contentcontainer>
  
