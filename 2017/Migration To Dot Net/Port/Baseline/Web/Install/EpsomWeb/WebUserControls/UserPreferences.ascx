<%@ Control Language="c#" AutoEventWireup="false" Codebehind="UserPreferences.ascx.cs" Inherits="Epsom.Web.WebUserControls.UserPreferences" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<h2 class="form_heading">Preferences</h2>
<table class="formtable">
  <tr>
	<TD><asp:Label id="lblErrorMessage" runat="server"></asp:Label></TD>
  </tr>
  <tr>
	<td><asp:checkbox id="chkNotificationFlag" runat="server" text="Case tracking notifications?"/></td>
  </tr>
  <tr>
	<td>
      <table>
		<tr>
			<td>How</td>
		</tr>
		<tr>
			<td><asp:checkbox id="chkNotificationEmail" runat="server" text="Email? (default and mandatory)"/></td>
		</tr>
	  </table>
	</td>
  </tr>
  <tr>
	<td><asp:checkbox id="chkMarketingFlag" text="Marketing data. Please tick the forms of marketing communication." runat="server"/></td>
	<TD><asp:Label id="lblMarketingFlag" runat="server"></asp:Label></TD>
  </tr>
  <tr>
	<td>
      <table>
		<tr>
	      <td><asp:checkbox id="chkMarketingPost" text="Post?" runat="server"/></td>
		</tr>
		<tr>
	      <td><asp:checkbox id="chkMarketingFax" text="Fax?" runat="server"/></td>
		</tr>
		<tr>
	      <td><asp:checkbox id="chkMarketingEmail" text="Email?" runat="server"/></td>
		</tr>
	  </table>
 	</td>
  </tr>
</table>
