<%@ Page language="c#" Codebehind="Register.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Register" %>
<%@ Register Src="~/WebUserControls/Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register Src="~/WebUserControls/TelephoneNumbers.ascx" TagName="TelephoneNumbers" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Register TagPrefix="CustomControl" Namespace="MetaBuilders.WebControls" Assembly="MetaBuilders.WebControls.GlobalRadioButton" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id=region1 runat="server">

    <h1>Register Broker</h1>
	<cms:cmsBoilerplate cmsref="4060" runat="server" ID="Cmsboilerplate4"/><br/>
    <asp:validationsummary id="Validationsummary1" runat="server" headertext="Input errors:" CssClass="validationsummary" showsummary="true" displaymode="BulletList" />
    
    <asp:panel id="pnlFormError" runat="server" Visible="False">
      <p class="error">
        <asp:label id="lblFormErrorMessage" runat="server" /> 
      </p>
    </asp:panel>

    <h2 class="form_heading">FSA details and personal details</h2>
    
    <fieldset>
    
      <legend>FSA details</legend>
    
      <table class="formtable">

        <tr>
          <td class="prompt">Are you an appointed representative?</td>
          <td>
            <asp:radiobutton id="rbAr" runat="server" GroupName="Representative" text="Yes" />
            <asp:radiobutton id="rbDa" runat="server" GroupName="Representative" text="No" checked="True" />
          </td>
        </tr>

        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Firm FSA ref. number</td>
          <td>
            <asp:TextBox id="txtFsaNumber" Runat="server" CssClass="text" />
            <asp:button id="cmdFindFirm" Runat="server" Text="Validate" CssClass="button" CausesValidation="false" />
            <asp:button id="cmdResetFirm" runat="server" text="Reset" CssClass="button" causesvalidation="false" />
            <asp:RequiredFieldValidator Runat="Server" Text="***" ErrorMessage="Please enter FSA number and validate"
              ControlToValidate="txtFsaNumber" EnableClientScript="false" ID="Requiredfieldvalidator4"/>
            <asp:CustomValidator id="validatorFsaNumber" Runat="server" EnableClientScript="false" ControlToValidate="txtFsaNumber"
              Errormessage="Please validate FSA number" Text="***" OnServerValidate="ValidateFirm"/>
          </td>
        </tr>

      </table> 
        
      <asp:panel id="pnlFindFirmMessage" runat="server" Visible="False">

        <p class="error">
          <asp:label id="lblFindFirmMessage" runat="server" /> 
        </p>  

      </asp:panel>
      
      <asp:panel id="pnlNetworkFirms" runat="server" Visible="false">
        <table class="formtable">
          <tr>
            <th>Fsa Ref</th>
            <th>CompanyName</th> 
            <th>Permissions</th>
            <th>
              <asp:customvalidator id="validateNetworkSelected" runat="server" 
              errormessage="Please Select a Network"
              enableclientscript="false" onservervalidate="IsNetworkSelected" text="***" >
            </asp:customvalidator> 
            </th>
          </tr>
          
          <asp:repeater runat="server" id="rptNetworkFirms">
            <ItemTemplate>
              <tr>
                <td><asp:label id="lblNetworkFsaNumber" runat="server"></asp:label></td>
                <td><asp:label id="lblNetworkCompanyName" runat="server"></asp:label></td>
                <td><asp:label id="lblNetworkPermissions" runat="server"></asp:label></td>
                <td><CustomControl:GlobalRadioButton runat="server" id="rdoNetworkSelected" GroupName="grpNetworkSelected" /></td> 
            </ItemTemplate>
          </asp:repeater>
        </p>  

      </asp:panel>
              
      <asp:panel id="pnlFirmDetails" runat="server" Visible="False">

        <table class="formtable" style="border:1px solid #CCC;background-color:#EEE">

          <tr>
            <td class="prompt">Company Name</td>
            <td><asp:Label id="lblFsaCompanyName" Runat="server" /></td>
          </tr>
          <tr>
            <td class="prompt">FSA Permissions</td>
            <td><asp:Label id="lblPermissions" Runat="server" /></td>
          </tr>
        </table> 
<!- removed following email from Brett -->
        <table class="formtable">
          <tr>
            <td class="prompt">Trading as</td>
            <td><asp:TextBox id="txtTradingAs" Runat="server" CssClass="text" /></td>
          </tr>  

        </table> 
      
      </asp:panel>
         
    </fieldset>
      
    <fieldset>
    
      <legend>Personal details</legend>
        
      <table class="formtable">
        
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Title</td>
          <td>
            <asp:dropdownlist id="cmbTitle" runat="server" />
            <validators:RequiredSelectedItemValidator Runat="server" EnableClientScript="false"
              Text="***" ErrorMessage="Please enter a title" ControlToValidate="cmbTitle" ID="Requiredselecteditemvalidator1" />
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> First name</td>
          <td>
            <asp:TextBox id="txtFirstName" Runat="server" CssClass="text" Maxlength="30" />
            <asp:RequiredFieldValidator Runat="Server" Text="***" ErrorMessage="Please enter your first name" ControlToValidate="txtFirstName" EnableClientScript="false" />
          </td>
        </tr>  
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Last name</td>
          <td>
            <asp:TextBox id="txtLastName" Runat="server" CssClass="text" MaxLength="35" />
            <asp:RequiredFieldValidator Runat="Server" Text="***" ErrorMessage="Please enter your last name" ControlToValidate="txtLastName" EnableClientScript="false" ID="Requiredfieldvalidator1" NAME="Requiredfieldvalidator1" />
          </td>
        </tr>
       
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Date of birth</td>
          <td>
            <asp:TextBox id="txtBirth" Runat="server" CssClass="text" /> &nbsp;(dd/mm/yyyy)
            <asp:RequiredFieldValidator Runat="Server" Text="***" ErrorMessage="Please enter your date of birth" ControlToValidate="txtBirth" EnableClientScript="false" ID="Requiredfieldvalidator2" NAME="Requiredfieldvalidator1" />
            <asp:CompareValidator Operator="DataTypeCheck" Type="Date" Runat="server" ControlToValidate="txtBirth" Text="***" ErrorMessage="Date is invalid" EnableClientScript="false" />
          </td>
        </tr>
      </table>
      
    </fieldset>

    <h2 class="form_heading">Correspondence address and contact numbers</h2>
    
    <fieldset>
    
      <legend>Address</legend>
    
      <epsom:Address id="ctlAddress" runat="server" UnitedKingdomOnly="True"/>
    
    </fieldset>
    
    <fieldset>
    <legend>Telephone Numbers</legend>
    <epsom:telephonenumbers id="ctlTelephoneNumbers" runat="server" />
    </fieldset>

    <h2 class="form_heading">Login details</h2>

    <table class="formtable">

      <tr>
        <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Personal email address</td>
        <td>
          <asp:TextBox id="txtEmailAddress" Runat="server" CssClass="text" style="width: 30em" maxlength="350" />
          <asp:RequiredFieldValidator Runat="Server" ErrorMessage="Please enter your email" ControlToValidate="txtEmailAddress" EnableClientScript="false" text="***" />
          <asp:RegularExpressionValidator ID="RegularExpressionValidator1" Runat="server" 
              ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*" 
              ErrorMessage="Please enter a valid E-mail ID" ControlToValidate="txtEmailAddress" EnableClientScript="false" text="*" />
        </td>
      </tr>
      <tr>
        <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Confirm email address</td>
        <td>
          <asp:TextBox id="txtConfirmEmailAddress" Runat="server" CssClass="text" style="width: 30em" maxlength="350"/>
          <asp:RequiredFieldValidator Runat="Server" EnableClientScript="false" ErrorMessage="Please confirm your email" ControlToValidate="txtConfirmEmailAddress" text="***" />
          <asp:CompareValidator id="CompareValidatorEmailAddress" Runat="server" EnableClientScript="false" ErrorMessage="Email addresses must match" ControlToValidate="txtEmailAddress" ControlToCompare="txtConfirmEmailAddress" Operator="Equal" Type="String" Text="*" />
        </td>
      </tr>
      <tr>
        <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Password</td>
        <td>
          <asp:TextBox id="txtPassword" Runat="server" Textmode="Password" CssClass="text" maxlength="15" />
          <asp:RequiredFieldValidator Runat="Server" EnableClientScript="false" ErrorMessage="Please enter your password" ControlToValidate="txtPassword" text="***" />
        </td>
      </tr>
      <tr>
        <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Confirm password</td>
        <td>
          <asp:TextBox id="txtConfirmPassword" Runat="server" Textmode="Password" CssClass="text" maxlength="15" />
          <asp:RequiredFieldValidator Runat="Server" EnableClientScript="false" ErrorMessage="Please confirm your password" ControlToValidate="txtConfirmPassword" text="***" ID="Requiredfieldvalidator3" NAME="Requiredfieldvalidator3" />
          <asp:CompareValidator Runat="server" EnableClientScript="false" ErrorMessage="Passwords must match" ControlToValidate="txtPassword" ControlToCompare="txtConfirmPassword" Operator="Equal" Type="String" Text="*" />
        </td>
      </tr>
    </table>

    <p style="text-align:right">
      <asp:button id=cmdCancel runat="server" text="Cancel" CssClass="button" />
      &nbsp;
      <asp:button id=cmdRegister runat="server" text="Register" CssClass="button" />
    </p>
    
    <sstchur:SmartScroller id="ctlSmartScroller" runat="server" />

  </mp:content>

</mp:contentcontainer>
