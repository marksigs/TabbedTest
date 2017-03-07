<%@ Page language="c#" Codebehind="MyProfileAddress.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.MyProfileAddress" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/UserConfirmPassword.ascx" TagName="UserConfirmPassword" TagPrefix="Epsom" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/TelephoneNumbers.ascx" TagName="TelephoneNumbers" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">
	<mp:content id="region1" runat="server">
		<h1>My Profile</h1>
		<asp:validationsummary id="Validationsummary1" runat="server" headertext="Input errors:" CssClass="validationsummary" showsummary="true" displaymode="BulletList"></asp:validationsummary>
		<asp:panel id="pnlEdit" Runat="server">
    <h2 class="form_heading">Edit Contact details</h2>
    <p>
      You can change the address to which we send case documents and marketing information here.
    </p>
			<epsom:Address id="ctlAddress" runat="server" UnitedKingdomOnly="True"></epsom:Address>
			<br />
    <h2 class="form_heading">Update your telephone numbers</h2>
    <epsom:telephonenumbers id="ctlTelephoneNumbers" runat="server" />


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
    <sstchur:SmartScroller id="ctlSmartScroller" runat="server" />
	</mp:content>
</mp:contentcontainer>
