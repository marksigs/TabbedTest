<%@ Control Language="c#" AutoEventWireup="false" Codebehind="ContractLength.ascx.cs" Inherits="Epsom.Web.WebUserControls.EmploymentLength" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>

<table class="formtable">
	<tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Job title</td>
    <td>
      <asp:dropdownlist id="cmbJobTitle" runat="server" />
      <validators:requiredselecteditemvalidator runat="server" enableclientscript="false"
          text="***" errormessage="Job Title is required" controltovalidate="cmbJobTitle" id="Requiredselecteditemvalidator1"/> 
    </td>
    </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Employment status</td>
    <td>
      <asp:dropdownlist id="cmbEmploymentStatus" runat="server" autopostback="True"/>
      <validators:requiredselecteditemvalidator runat="server" enableclientscript="false"
          text="***" errormessage="Employment Status is required" controltovalidate="cmbEmploymentStatus" id="Requiredselecteditemvalidator2"/> 
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Date Employment Started</td>
    <td>
      <asp:textbox id="txtDateEmploymentStarted" runat="server" columns="10" cssclass="text" autopostback="True"/>
      <validators:datevalidator id="DateStartedValidator" runat="server" controltovalidate="txtDateEmploymentStarted" text="***" errormessage="Invalid date Employment Started"
        enableclientscript="false" />
      <asp:label id="lblDateEmploymentFinished" runat="server"></asp:label>
      <asp:panel runat="server" id="pnlDateEmploymentFinished">
        <asp:textbox id="txtDateEmploymentFinished" runat="server" columns="10" cssclass="text" />
        <validators:datevalidator id="DateFinishedValidator" runat="server" controltovalidate="txtDateEmploymentFinished" text="***" errormessage="Invalid date Employment Finished"
          enableclientscript="false" />
      </asp:panel>
      dd/mm/yyyy. 
      </td>
      <td>You must enter 1 year's worth of employment dates.  
    </td>
  </tr>

<asp:panel runat="server" id ="pnlContractLength">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Length of contract</td>
    <td>
      <asp:textbox id="txtContractMonths" runat="server" columns="10" cssclass="text" />months.&nbsp;
      <asp:requiredfieldvalidator runat="server" controltovalidate="txtContractMonths" text="***" errormessage="Contract Months cannot be blank"
        id="txtContractMonths_validator" enableclientscript="false" />
      <asp:comparevalidator id="CompareValidator1" controltovalidate="txtContractMonths" errormessage="Contract Months is not numeric"
              text="***" enableclientscript="false" runat="server" type="Integer" operator="DataTypeCheck" />
      </td>
    </tr>
    <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Renewable?
          <td>
            <asp:radiobuttonlist id="rblRenewable" runat="Server" autopostback="true" repeatdirection="Horizontal" repeatlayout="Flow">
              <asp:listitem value="True">Yes</asp:listitem>
              <asp:listitem value="False">No</asp:listitem>
            </asp:radiobuttonlist>
            <asp:requiredfieldvalidator runat="server" id="rblRenewableValidator" controltovalidate="rblRenewable" text="***" errormessage="Specify if your contract is renewable" enableclientscript=false ></asp:requiredfieldvalidator> 
          </td>
        </tr> 
    </td>
  </tr>
</asp:panel> 

  </table>