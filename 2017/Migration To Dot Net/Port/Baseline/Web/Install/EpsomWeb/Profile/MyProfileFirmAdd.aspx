<%@ Page language="c#" Codebehind="MyProfileFirmAdd.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.MyProfileFirmAdd" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/UserConfirmPassword.ascx" TagName="UserConfirmPassword" TagPrefix="Epsom" %>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">
	<mp:content id="region1" runat="server">
			<h1>My Profile</h1>
			<asp:validationsummary id="Validationsummary1" runat="server" headertext="Input errors:" CssClass="validationsummary" showsummary="true" displaymode="BulletList"></asp:validationsummary>
      <asp:panel id="pnlEdit" runat="server">
      <h2 class="form_heading">Add Firms</h2>
      <p>You can add additional firms for which you are a representative here.</p>
      <table class=formtable>
        <tr>
          <td class=prompt>Are you an appointed representative?</td>
          <td>
            <asp:radiobutton id=rbAr runat="server" GroupName="Representative" text="Yes"></asp:radiobutton>
            <asp:radiobutton id=rbDa runat="server" GroupName="Representative" text="No" checked="True"></asp:radiobutton>
          </td>
        </tr>
        
        <tr>
          <td class=prompt><span class="mandatoryfieldasterisk">*</span> Firm FSA ref. number</td>
          <td>
            <asp:TextBox id="txtFsaNumber" Runat="server"></asp:TextBox>
            <asp:RequiredFieldValidator Runat="Server" Text="***" ErrorMessage="Please enter FSA number and validate"
             ControlToValidate="txtFsaNumber" EnableClientScript="false" ID="Requiredfieldvalidator1"></asp:RequiredFieldValidator>
            <asp:CustomValidator id="validatorFsaNumber" Runat="server" EnableClientScript="false" ControlToValidate="txtFsaNumber"
              Errormessage="Please validate FSA number" Text="***"></asp:CustomValidator>
            <asp:button id="cmdValidate" Runat="server" Text="Validate" CssClass="button" CausesValidation="false"></asp:button>
            <asp:button id="cmdReset" runat="server" text="Reset" CssClass="button" causesvalidation="false"></asp:button>
          </td>
        </tr>
        
        <asp:panel id="pnlFindFirmMessage" runat="server" Visible="false">
          <tr>
            <td class=prompt></td>
            <td>
              <asp:label id="lblFindFirmMessage" runat="server"></asp:label>              
            </td>
          </tr>
        </asp:panel>
       
        <asp:panel id="pnlFirmDetails" runat="server" Visible="false">
          <tr>
            <TD class=prompt>Company Name</TD>
            <TD><asp:Label id=lblFsaCompanyName Runat="server"></asp:Label></TD>
          </TR>
          <tr>
            <TD class=prompt>FSA Permissions</TD>
            <td><asp:Label id="lblPermissions" Runat="server" /></td>
          </TR>
         
        </asp:panel>
        
        </table class=formtable> 
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
