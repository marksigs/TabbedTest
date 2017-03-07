<%@ Control Language="c#" AutoEventWireup="false" Codebehind="Adviser.ascx.cs" Inherits="Epsom.Web.WebUserControls.Adviser" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

    <asp:validationsummary id="Validationsummary1" runat="server" showsummary="true" cssclass="validationsummary" headertext="Input errors:" displaymode="BulletList" />

    <h2 class="form_heading">Adviser Details</h2>

    <br />

    <p><cms:cmsBoilerplate cmsref="11" runat="server" /></p>
    
    <table class="formtable">

      <asp:panel id="pnlSearchCriteria" runat="server">
        <tr>
          <td class="prompt">&nbsp;</td>
          <td>
            <asp:radiobuttonlist tabindex="1" id="rblType" runat="Server" repeatlayout="flow" repeatdirection="horizontal">
              <asp:listitem value="0" selected="true">Email address&nbsp;&nbsp; or &nbsp;&nbsp;</asp:listitem>
              <asp:listitem value="1">FSA number</asp:listitem>
            </asp:radiobuttonlist>
          </td>
        </tr>
      </asp:Panel>

      <tr>
        <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Broker identification</td>
        <td>
          <asp:textbox id="txtIdentification" runat="server" cssclass="text" style="width: 25em" autopostback="false" tabindex="2"/>
          <asp:button id="CmdValidateBroker" runat="server" text="Find broker" cssclass="button" causesvalidation="false" tabindex="3"/>
          <asp:button id="CmdResetBroker" runat="server" text="Reset" cssclass="button" causesvalidation="false" tabindex="4"/>
          <asp:requiredfieldvalidator runat="server" controltovalidate="txtIdentification" text="***" errormessage="Please enter an FSA reference number or email address"
          id="requiredValidatorFsaNumber" enableclientscript="false" />
          <asp:customvalidator id="validatorFsaNumber" runat="server" enableclientscript="false" controltovalidate="txtIdentification" 
            errormessage="Please validate FSA number or email address" text="***" onservervalidate="ValidateFirm"/>
        </td>
      </tr>
    </table>

    <asp:panel id="pnlValidateBrokerMessage" runat="server" visible="False">

      <p class="error">
        <asp:label id="lblValidateBrokerMessage" runat="server" /> 
      </p>

    </asp:panel>

    <asp:panel id="pnlFirms" runat="server" visible="False">
      <p>
        This broker operates on behalf of the following firms, please select the 
        appropriate option for this transaction from the list below.
      </p>
      <table class="formtable">
        <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Firm</td>
          <td>
            <asp:dropdownlist id="cmbFirms" runat="server" autopostback="True" style="width:30em" />
          </td>
        </tr>
      </table>
    </asp:panel>

    <asp:panel id="pnlFirm" runat="server" visible="False">

      <h2 class="form_heading">Registered Details</h2>

      <table class="formtable">
        <tr>
          <td class="prompt">FSA number</td>
          <td>
            <asp:label id="lblFsaNumber" runat="server" />
          </td>
        </tr>
        <tr>
          <td class="prompt">Company Name</td>
          <td>
            <asp:label id="lblCompanyName" runat="server" />
          </td>
        </tr>
        <tr>
          <td class="prompt">Address</td>
          <td>
            <asp:label id="lblAddress" runat="server" />
          </td>
        </tr>
      </table>

    </asp:panel>
