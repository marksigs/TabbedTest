<%@ Page language="c#" Codebehind="MyProfile.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.MyProfile" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/UserProfile.ascx" TagName="UserProfile" TagPrefix="Epsom" %>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">
<mp:content id="region1" runat="server">
<h1>My Profile</h1>
<epsom:UserProfile id="ctlUserProfile" runat="server" IsRegistration="false"></epsom:UserProfile>
</mp:content>
</mp:contentcontainer>
