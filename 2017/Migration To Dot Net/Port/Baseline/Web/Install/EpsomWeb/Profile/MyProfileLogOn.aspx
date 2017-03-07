<%@ Page language="c#" Codebehind="MyProfileLogOn.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.MyProfileLogOn" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">
	<mp:content id="region1" runat="server">

			<h1>My Profile</h1>

 		  <asp:panel id="pnlError" Runat="server">
        <p class="error">
          <asp:Label ID="lblMessage" Runat="server" Visible="false" />
        </p>
      </asp:panel>

			<asp:validationsummary id="Validationsummary1" runat="server" headertext="Input errors:" CssClass="validationsummary" showsummary="true" displaymode="BulletList"></asp:validationsummary>

			<h2 class="form_heading">Edit Login Details</h2>
			<table class="formtable">
 		  <asp:panel id="pnlBroker" Runat="server">
				<tr>
					<td class="prompt">Current email address</td>
					<td>
						<asp:TextBox ID="txtCurrentEmailAddress" Runat="server"></asp:TextBox>
					</td>
				</tr>
				<tr>
					<td class="prompt">New email address</td>
					<td>
						<asp:TextBox ID="txtEmailAddress" Runat="server" maxlength="350"></asp:TextBox>
						<asp:RegularExpressionValidator Runat="server" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
							ErrorMessage="Please enter a valid email address" EnableClientScript="false" ControlToValidate="txtEmailAddress" text="***">
          </asp:RegularExpressionValidator>
					</td>
				</tr>
				<tr>
					<td class="prompt">Confirm new email address</td>
					<td>
						<asp:TextBox ID="txtConfirmEmailAddress" Runat="server" maxlength="350"></asp:TextBox>
						<asp:CompareValidator Runat="server" EnableClientScript="false"
						 ErrorMessage="Email addresses must match" text="***"
							ControlToValidate="txtEmailAddress" ControlToCompare="txtConfirmEmailAddress" Operator="Equal" Type="String">
           </asp:CompareValidator>
					</td>
				</tr>
				<tr>
					<td><p></p>
					</td>
				</tr>
				</asp:panel>
				<tr>
					<td class="prompt"><span class="mandatoryfieldasterisk">*</span> Current Password</td>
					<td>
						<asp:TextBox ID="txtCurrentPassword" Runat="server" Textmode="Password"></asp:TextBox>
            <asp:RequiredFieldValidator Runat="Server"
              ErrorMessage="Please enter your current password" ControlToValidate="txtCurrentPassword" text="***" />
					</td>
				</tr>
				<tr>
					<td class="prompt">New Password</td>
					<td>
						<asp:TextBox ID="txtPasswordNew" Runat="server" Textmode="Password" MaxLength="15"></asp:TextBox>
					</td>
				</tr>
				<tr>
					<td class="prompt">Confirm password</td>
					<td>
						<asp:TextBox ID="txtConfirmPassword" Runat="server" Textmode="Password" MaxLength="15"></asp:TextBox>
          <asp:CompareValidator Runat="server" EnableClientScript="false" ErrorMessage="Passwords must match" text="***"
           ControlToValidate="txtPasswordNew" ControlToCompare="txtConfirmPassword" Operator="Equal" Type="String">
          </asp:CompareValidator>
					</td>
				</tr>
			</table>
			<asp:button id="cmdSubmit" runat="server" text="Submit Changes" CssClass="button" />
			<asp:button id="cmdCancel" runat="server" text="Cancel" CssClass="button" CausesValidation="False" />
	</mp:content>
</mp:contentcontainer>
