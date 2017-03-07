<%@ Page language="c#" Codebehind="MyProfilePersonal.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Profile.MyProfilePersonal" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/UserConfirmPassword.ascx" TagName="UserConfirmPassword" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">
	<mp:content id="region1" runat="server">
		<h1>My Profile</h1>
 		<asp:panel id="pnlError" Runat="server">
      <p class="error">
          <asp:Label ID="lblMessage" Runat="server" Visible="false" />
        </p>
    </asp:panel>
		<asp:panel id="pnlEdit" Runat="server">
			<asp:validationsummary runat="server" headertext="Input errors:" CssClass="validationsummary" showsummary="true" displaymode="BulletList"></asp:validationsummary>
			<h2 class="form_heading">Edit personal details</h2>
      <p>Update the personal details we hold for your account here.</p>
			<table class="formtable">
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Title</td>
          <td>
            <asp:dropdownlist id="cmbTitle" runat="server" /> </td>
          <td>  
            <validators:RequiredSelectedItemValidator Runat="server" EnableClientScript="false"
              Text="***" ErrorMessage="Please enter a title" ControlToValidate="cmbTitle" ID="Requiredselecteditemvalidator1" />
          </td>
        </tr>
				<tr>
					<td class="prompt"><span class="mandatoryfieldasterisk">*</span> First name</td>
					<td><asp:TextBox ID="txtFirstName" Runat="server"></asp:TextBox>
            <asp:RequiredFieldValidator Runat="Server" Text="***" ErrorMessage="Please enter your first name" ControlToValidate="txtFirstName" EnableClientScript="false" ID="Requiredfieldvalidator1"/>
          </td>
				</tr>
				<tr>
					<td class="prompt"><span class="mandatoryfieldasterisk">*</span> Last name</td>
					<td><asp:TextBox ID="txtLastName" Runat="server"></asp:TextBox>
            <asp:RequiredFieldValidator Runat="Server" Text="***" ErrorMessage="Please enter your last name" ControlToValidate="txtLastName" EnableClientScript="false" ID="Requiredfieldvalidator2" NAME="Requiredfieldvalidator1" />
          </td>
				</tr>
        <tr>
          <td class="prompt">Date of birth</td>
          <td><asp:TextBox id="txtBirth" Runat="server" CssClass="text" Enabled="false" /></td>
          <td>If your date of birth has been input incorrectly, please call our helpline.</td>
        </tr>
      </table>
      <table class="formtable">
        <tr>
          <td>&nbsp;</td>
          <td><asp:button id="cmdSubmit" runat="server" text="Submit changes" CssClass="button" /></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td><asp:button id="cmdCancel" runat="server" text="Cancel" CausesValidation="False" CssClass="button" /></td>
        </tr>
			</table>
		</asp:panel>
		<asp:panel id="pnlPassword" Runat="server" Visible="false">
			<epsom:UserConfirmPassword id="ctlUserConfirmPassword" runat="server"></epsom:UserConfirmPassword>
		</asp:panel>
	</mp:content>
</mp:contentcontainer>
