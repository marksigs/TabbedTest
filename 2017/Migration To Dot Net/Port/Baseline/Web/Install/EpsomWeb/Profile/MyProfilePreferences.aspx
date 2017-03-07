<%@ Register Src="~/WebUserControls/UserPreferences.ascx" TagName="UserPreferences" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Page language="c#" Codebehind="MyProfilePreferences.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.MyProfilePreferences" %>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">
	<mp:content id="region1" runat="server">
		<asp:panel id="pnlEdit" Runat="server">
			<H1>My Profile - Edit Preferences</H1>
			<epsom:UserPreferences id="ctlUserPreferences" runat="server"></epsom:UserPreferences>
			<BR>
			<asp:button id="cmdCancel" runat="server" text="Cancel" CausesValidation="False"></asp:button>
			<asp:button id="cmdSubmit" runat="server" text="Submit"></asp:button>
		</asp:panel>
		<asp:panel id="pnlPassword" Runat="server">
			<H1>My Profile - Confirm Edit</H1>
			<TABLE class="formtable">
				<TR>
					<TD>Password  *</TD>
					<TD><asp:TextBox id="txtPassword" Runat="server" Textmode="Password"></asp:TextBox>
					  <asp:RequiredFieldValidator id="RequiredfieldPassword" Runat="Server"
                        ControlToValidate="txtPassword" EnableClientScript="false">Missing password
            </asp:RequiredFieldValidator>
					  <asp:CustomValidator id="CustomValidatorPassword" Runat="Server"
                           OnServerValidate="CustomValidatorPassword_ServerValidate"
                           ControlToValidate="txtPassword">Invalid password       
            </asp:CustomValidator>
          </TD>
				</TR>
				<TR>
					<TD>
						<asp:Label id="lblMessage" runat="server"></asp:Label></TD>
				</TR>
			</TABLE>
			<asp:button id="cmdCancelPassword" runat="server" text="Cancel" CausesValidation="False"></asp:button>
			<asp:button id="cmdSubmitPassword" runat="server" text="Submit Changes"></asp:button>
		</asp:panel>
	</mp:content>
</mp:contentcontainer>
