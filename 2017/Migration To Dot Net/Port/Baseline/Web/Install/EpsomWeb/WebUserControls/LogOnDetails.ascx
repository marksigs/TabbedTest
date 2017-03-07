<%@ Control Language="c#" AutoEventWireup="false" Codebehind="LogOnDetails.ascx.cs" Inherits="Epsom.Web.WebUserControls.LogOnDetails" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<asp:panel runat="server" id="pnlLoggedIn">

  <h2>Welcome <asp:Label ID="lblFullName" Runat="server" /></h2>

  <h3>Your firm</h3>

  <p>
    <asp:Label ID="lblFirmName" Runat="server" />
  </p>

  <h3>FSA permissions</h3>
      <p>
       <asp:label id="lblPermissions" runat="server" />
      </p>


</asp:panel>

<asp:panel runat="server" id="pnlLoggedOut"> 

  <h2>Login to my homepage</h2>

  <p>
    If you are a registered user of this site, you can login here.
  <p/>
  <p>
    <a href="<%=ResolveUrl("~/Logon.aspx")%>" title="" class="fancylink">Login here</a>
  </p>

</asp:panel>
