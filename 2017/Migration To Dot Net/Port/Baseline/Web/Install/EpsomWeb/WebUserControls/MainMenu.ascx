<%@ Control Language="c#" AutoEventWireup="false" Codebehind="MainMenu.ascx.cs" Inherits="Epsom.Web.WebUserControls.MainMenu" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>

<asp:hyperlink id="lnkLogOn" navigateurl="~/LogOn.aspx" runat="server" text="Log in" />
|
<asp:hyperlink id="lnkRegister" navigateurl="~/Profile/Register.aspx" runat="server" text="Register" />
|
<asp:hyperlink id="lnkAboutUs" navigateurl="~/Contact/AboutUs.aspx" runat="server" text="About us" />
|
<asp:hyperlink id="lnkContactUs" navigateurl="~/Contact/Support.aspx" runat="server" text="Contact us" />
|
<asp:hyperlink id="lnkHelp" navigateurl="~/Help/HelpList.aspx" runat="server" text="Help" />
