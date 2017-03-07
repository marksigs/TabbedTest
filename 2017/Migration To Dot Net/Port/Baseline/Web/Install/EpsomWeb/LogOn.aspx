<%@ Page language="c#" Codebehind="LogOn.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.LogOn" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register Src="~/WebUserControls/Messages.ascx" TagName="Messages" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>Login</h1>

    <asp:Panel id="pnlLoginDetails" runat="server">

      <asp:validationsummary id="Validationsummary1" runat="server" headertext="Input errors:" CssClass="validationsummary" showsummary="true" displaymode="BulletList" />

      <p><cms:cmsBoilerplate cmsref="4070" runat="server" /></p>
      
      <epsom:Messages id="ctlMessages" runat="server" />
      
      <asp:panel id="pnlError" Runat="server">
        <p class="error">
          Error: the supplied username and/or password did not match any recognised combination.
          <br />
          You may try again, but the system will only permit two retries before locking the user account.
        </p>
      </asp:panel>
      
      <table class="formtable">
      
      <asp:panel id="pnlLogOn" Runat="server">
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> User name</td>
          <td>
            <asp:TextBox ID="txtUserName" Runat="server" CssClass="text" style="width: 16em" />
            <asp:RequiredFieldValidator id=RequiredfieldvalidatorUserName Runat="Server"
            ErrorMessage="Please enter your user name" ControlToValidate="txtUserName" text="***" />
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Password</td>
          <td>
            <asp:TextBox ID="txtPassword" Runat="server" Textmode="Password" CssClass="text" style="width: 16em" />
            <asp:RequiredFieldValidator Runat="Server"
            ErrorMessage="Please enter your password" ControlToValidate="txtPassword" text="***" />
          </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>
            <asp:Button id="cmdLogOn" Runat="server" text="Login" CssClass="button" />
            <asp:button id=cmdCancel runat="server" text="Cancel" CssClass="button" CausesValidation="false"/>
          </td>
        </tr>
      </asp:panel>
      
      <asp:panel id="pnlFirm" Runat="server">
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Firm</td>
          <td><asp:dropdownlist id="cmbFirm" runat="server" />
              <validators:RequiredSelectedItemValidator Runat="server" EnableClientScript="false"
                Text="***" ErrorMessage="Firm is required" ControlToValidate="cmbFirm"/>
          </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td><asp:Button id="cmdConfirm" Runat="server" text="Confirm selection" CssClass="button" /></td>
        </tr>
      </asp:panel>
      
      </table>

      <p>
        Forgotten your password?
        &nbsp;
        <asp:hyperlink id="lnkContact" navigateurl="~/Contact/Support.aspx" runat="server" text="Contact Us" /></td>
      </p>

    </asp:Panel>

    <asp:panel id="pnlOmigaAvailability" Runat="server">
      <p class="error">
        The system is currently unavailable
      </p>
    </asp:panel>
   
  </mp:content>

</mp:contentcontainer>
