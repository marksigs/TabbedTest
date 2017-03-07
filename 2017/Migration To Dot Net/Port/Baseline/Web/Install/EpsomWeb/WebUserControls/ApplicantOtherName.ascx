<%@ Control Language="c#" AutoEventWireup="false" Codebehind="ApplicantOtherName.ascx.cs" Inherits="Epsom.Web.WebUserControls.ApplicantOtherName" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<table class="formtable">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Title</td>
    <td><asp:dropdownlist id="cmbTitle" runat="server" />
        <validators:requiredselecteditemvalidator runat="server" enableclientscript="false"
          text="***" errormessage="Title is required" controltovalidate="cmbTitle" id="Requiredselecteditemvalidator1"/>
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> First name</td>
    <td>
      <asp:textbox id="txtApplicantFirstName" runat="server" columns="40" cssclass="text" />
      <asp:requiredfieldvalidator runat="server" controltovalidate="txtApplicantFirstName" text="***" errormessage="First name cannot be blank"
        id="txtApplicantFirstName_validator" enableclientscript="false" />
    </td>
  </tr>
  <tr>
    <td class="prompt">Middle names</td>
    <td><asp:textbox id="txtApplicantMiddleNames" runat="server" columns="40" cssclass="text" /></td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Last name</td>
    <td>
      <asp:textbox id="txtApplicantLastName" runat="server" columns="40" cssclass="text" />
      <asp:requiredfieldvalidator runat="server" controltovalidate="txtApplicantLastName" text="***" errormessage="Last name cannot be blank"
        id="txtApplicantLastName_validator" enableclientscript="false" />
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span>When did you stop using this name?</td>
    <td>
      <asp:textbox id="txtDateStopped" runat="server" columns="10" cssclass="text" />
      <validators:datevalidator id="DateValidator1" runat="server" controltovalidate="txtDateStopped" text="***" errormessage="Invalid date stopped using name"
        enableclientscript="false" /> dd/mm/yyyy
    </td>
  </tr>
</table>
