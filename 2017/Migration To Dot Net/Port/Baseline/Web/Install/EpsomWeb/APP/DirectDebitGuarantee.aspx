<%@ Register Src="~/WebUserControls/Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Page language="c#" Codebehind="DirectDebitGuarantee.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.DirectDebitGuarantee" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
<cms:helplink id="helplink" class="helplink" runat="server" helpref="2220" />
    <h1><cms:cmsBoilerplate cmsref="622" runat="server" /></h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />

    <asp:validationsummary id="Validationsummary1" runat="server" showsummary="true" CssClass="validationsummary"
      headertext="Input errors:" displaymode="BulletList" />
	<cms:cmsBoilerplate cmsref="220" runat="server" />
	
    <fieldset label="Bank/Building Society Details">

      <legend>Bank/Building Society Details</legend>

      <table class="formtable">
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">* </span>Bank/Building Society name</td>
          <td><asp:TextBox id="txtBankName" Runat="server" cssclass="text" />
              <asp:requiredfieldvalidator controltovalidate="txtBankName" runat="server" text="***" 
             errormessage="Please enter bank name" id="Requiredfieldvalidator1" enableclientscript="false"/>
          </td>
        </tr>
      </table>

      <epsom:address id="ctlAddress" runat="server" showresidencydates="false" shownatureofoccupancy="false"
        unitedkingdomonly="true" isresidential="false" />

      <table class="formtable">
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">* </span>Name of Account holder</td>
          <td><asp:TextBox id="txtAccountHolder" runat="server" cssclass="text" />
              <asp:requiredfieldvalidator controltovalidate="txtAccountHolder" runat="server" text="***" 
             errormessage="Please enter account holder" id="Requiredfieldvalidator2" enableclientscript="false"/>
         </td>
        </tr>
        <tr>
          <td class="prompt">2nd account holder (if applicable)</td>
          <td><asp:TextBox id="txtAccountHolder2" runat="server" cssclass="text" /></td>
        </tr>
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">* </span>Sort code</td>
          <td>
            <asp:TextBox id="txtSortCode1" runat="server" cssclass="text" width="20px" maxlength="2"/> -
            <asp:TextBox id="txtSortCode2" runat="server" cssclass="text" width="20px" maxlength="2"/> -
            <asp:TextBox id="txtSortCode3" runat="server" cssclass="text" width="20px" maxlength="2"/>
            <asp:customvalidator runat="server" id="SortCodeValidator" onservervalidate="ValidateSortCode" 
              validateemptytext="True"  text="***" errormessage="Invalid Sort Code" enableclientscript="false" />
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Account Number</td>
          <td>
            <asp:TextBox id="txtAccountNumber" runat="server" cssclass="text" />
            <asp:customvalidator runat="server" id="AccountCodeValidator" onservervalidate="ValidateAccountCode" 
             validateemptytext="True"  text="***" errormessage="Invalid Account Code" enableclientscript="false" />
          </td>
        </tr>
        <tr>
          <td class="prompt">Building Society Roll Number</td>
          <td><asp:TextBox id="txtRollNumber" runat="server" cssclass="text" /></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>
            <asp:button id="btnValidate" runat="server" cssclass="button" text="Validate account" causesValidation="false" />
            <asp:button id="btnClearDetails" runat="server" cssclass="button" text="Clear details" />
            <asp:customvalidator runat="server" id="BankWizardValidatePerformed" onservervalidate="BankWizardValidateHasBeenPerformed" 
             validateemptytext="True"  text="***" errormessage="You must have successfully validated account details" enableclientscript="false" />
          </td>
        </tr>
      </table>

    </fieldset>

    <b><asp:label id="lblValidateMsg" runat="server" /></b>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" buttonbar="true" savepage="true" saveandclosepage="true" runat="server" />

  </mp:content>

</mp:contentcontainer>
