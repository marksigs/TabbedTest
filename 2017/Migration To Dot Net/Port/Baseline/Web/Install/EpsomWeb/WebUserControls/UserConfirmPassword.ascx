<%@ Control Language="c#" AutoEventWireup="false" Codebehind="UserConfirmPassword.ascx.cs" Inherits="Epsom.Web.WebUserControls.UserConfirmPassword" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/Messages.ascx" TagName="Messages" TagPrefix="Epsom" %>

<h2 class="form_heading">Confirm Password</h2>
<p>
  Please enter your password to confirm your identity before editing your profile.
</p>
<table class="formtable">
	<tr>
		<td class="prompt"><span class="mandatoryfieldasterisk">*</span> Current password</td>
		<td>
			<asp:TextBox id="txtPassword" Runat="server" Textmode="Password"></asp:TextBox>
			<asp:RequiredFieldValidator id="RequiredfieldPassword" Runat="Server" EnableClientScript="false"
          ControlToValidate="txtPassword" ErrorMessage="Please enter your password" text="***">
      </asp:RequiredFieldValidator>
    </td>
  <tr>
		<td class="prompt"></td>
		<td>
    <asp:button id="cmdContinue" runat="server" text="Continue" CssClass="button"></asp:button>
    </td>
	</tr>
</table>
    
<epsom:Messages id="ctlMessages" runat="server"></epsom:Messages>
  
<table class="formtable">
	<tr>
		<td>&nbsp;</td>
		<td><asp:button id="cmdCancel" runat="server" text="Cancel" CausesValidation="False" CssClass="button"></asp:button></td>
	</tr>
</table>
