<%@ Control Language="c#" AutoEventWireup="false" Codebehind="Income.ascx.cs" Inherits="Epsom.Web.WebUserControls.Income" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<%@ Register Src="EmploymentLength.ascx" TagName="EmploymentLength" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>

<asp:repeater id="rptEmploymentLength" runat="server">
  <itemtemplate>
    <fieldset>
      <legend><asp:label id="lblEmploymentLength" runat="server" /></legend>
      <epsom:Employmentlength id="ctlEmploymentLength" runat="server" /> 
    </fieldset> 
  </itemtemplate>
</asp:repeater>
<p class="button_orphan">
  <br />
  <asp:button id="cmdRemoveEmploymentLength" runat="server" causesvalidation="False" text="Remove this employment" cssclass="button" />
  <asp:button id="cmdAddEmploymentLength" runat="server" causesvalidation="False" text="Add another employment" cssclass="button" />
  <asp:customvalidator runat="server" id="rptEnoughEmployments" onservervalidate="EnoughEmployments" 
    validateemptytext=True  text="*" errormessage="12 months employment history needed" enableclientscript="false" />
</p>

<asp:panel id="pnlNotSelfEmployedIncome" runat="server">
<table class="formtable">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Annual guaranteed regular income</td>
    <td>
      <asp:textbox id="txtGuaranteedIncome" runat="server" columns="10" cssclass="text" />
      <asp:requiredfieldvalidator runat="server" controltovalidate="txtGuaranteedIncome" text="***" errormessage="Guaranteed Income cannot be blank"
        id="Requiredfieldvalidator2" enableclientscript="false" />
      <asp:comparevalidator id="CompareValidator1" controltovalidate="txtGuaranteedIncome" errormessage="Guaranteed Income is not numeric"
              text="***" enableclientscript="false" runat="server" type="Currency" operator="DataTypeCheck" />
    </td>
    <td>
      Non-guaranteed / irregular sources of income include Overtime, Bonus, Commission etc.
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Non guaranteed</td>
    <td>
      <asp:textbox id="txtUnguaranteedIncome" runat="server" columns="10" cssclass="text" />
      <asp:requiredfieldvalidator runat="server" controltovalidate="txtUnguaranteedIncome" text="***" errormessage="Unguaranteed Income cannot be blank"
             id="Requiredfieldvalidator3" enableclientscript="false" />
      <asp:comparevalidator id="Comparevalidator2" controltovalidate="txtUnguaranteedIncome" errormessage="Unguaranteed Income is not numeric"
              text="***" enableclientscript="false" runat="server" type="Currency" operator="DataTypeCheck" />
    </td>
    <td>Only 50% of this income will be used for the affordability calculation.</td>
    </tr>
    </table>
    </asp:panel>
    <asp:panel id="pnlSelfEmployedIncome" runat="server">
<table class="formtable">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>Annual self-employed income</td>
    <td>
      <asp:textbox id="txtSelfEmployedIncome" runat="server" columns="10" cssclass="text" />
      <asp:requiredfieldvalidator runat="server" controltovalidate="txtSelfEmployedIncome" text="***" errormessage="Self Employed Income cannot be blank"
        id="Requiredfieldvalidator1" enableclientscript="false" />
      <asp:comparevalidator id="Comparevalidator3" controltovalidate="txtSelfEmployedIncome" errormessage="Self Employed Income is not numeric"
              text="***" enableclientscript="false" runat="server" type="Currency" operator="DataTypeCheck" />
    </td>
    <td>
    </td>
  </tr>
</table> 
    </asp:panel> 
    <table class="formtable">
    <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> UK Taxpayer
          <td>
            <epsom:nullableyesno id="nynTaxpayer" runat="server" required="true"
            errormessage="Specify if you are a UK Taxpayer" >
            </epsom:nullableyesno>
          </td>
   </tr> 
  </table> 
