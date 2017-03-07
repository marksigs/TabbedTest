<%@ Control Language="c#" AutoEventWireup="false" Codebehind="MinimumCriteria.ascx.cs" Inherits="Epsom.Web.WebUserControls.MinimumCriteria" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Register Src="~/WebUserControls/Documents.ascx" TagName="Documents" TagPrefix="Epsom" %>

<h2 class="form_heading">Adviser's declaration</h2>
<table class=formtable>
  <tr>
    <td class="wideprompt"><span class=mandatoryfieldasterisk>* </span>
    <cms:cmsBoilerplate cmsref="21" runat="server" />
    </td>
    <td>
      <asp:radiobuttonlist id="rblRegulatedLoan" runat="Server" repeatdirection="Horizontal" repeatlayout="Flow">
        <asp:listitem value="True">Yes</asp:listitem>
        <asp:listitem value="False">No</asp:listitem>
      </asp:radiobuttonlist>
      <asp:requiredfieldvalidator runat="server" enableclientscript="false"
       id="rblRegulatedLoanRequired" controltovalidate="rblRegulatedLoan" text="***"
       errormessage="Please specify if application for a regulated loan" />
    </td>
  
  </tr> 
  <tr>
    <td class="wideprompt""><span class=mandatoryfieldasterisk>* </span>
    <cms:cmsBoilerplate cmsref="22" runat="server" />
    </td>
    <td>
        <asp:dropdownlist id="cmbLevelOfService" runat="server" ></asp:dropdownlist>
        <validators:RequiredSelectedItemValidator id="Requiredselecteditemvalidator7" Runat="server" EnableClientScript="false" Text="***"
          ErrorMessage="Please specify level of service" ControlToValidate="cmbLevelOfService">
          </validators:RequiredSelectedItemValidator>
    </td>
  </tr> 
  <tr>
    <td class="wideprompt"><span class=mandatoryfieldasterisk>* </span>
    <cms:cmsBoilerplate cmsref="23" runat="server" />
    </td>
     <td>
        <asp:dropdownlist id="cmbProcFeeToCust" runat="server" AutoPostBack="True"></asp:dropdownlist>
        <validators:requiredselecteditemvalidator id="Requiredselecteditemvalidator1" runat="server" enableclientscript="false" text="***"
          errormessage="Please specify procuration fee/commission" controltovalidate="cmbProcFeeToCust">
          </validators:requiredselecteditemvalidator>
    </td>
  </tr> 
  <tr>
     <td class="wideprompt" style="text-align:right"><cms:cmsBoilerplate cmsref="24" runat="server" /> </td>
     <td><asp:textbox runat="server" CssClass="text" id="txtFeeToCustomer"  columns="10"></asp:textbox>
         <asp:CompareValidator id=CompareValidator4 Runat="server" EnableClientScript="false"
          Text="***" ErrorMessage="Fee to customer is not numeric" ControlToValidate="txtFeeToCustomer"
          Operator="DataTypeCheck" Type="Currency"></asp:CompareValidator>
        <asp:customvalidator id="validatorFeeToCustomer" runat="server" enableclientscript="false"
        controltovalidate="cmbProcFeeToCust"
        errormessage="Fee to customer is required" text="***" />
     </td>
  </tr>
</table>
<h2 class="form_heading">Minimum Criteria</h2>
<br />
<p>
  <cms:cmsBoilerplate cmsref="20" runat="server" />
</p>
<table class="minimumcriteriatable">
  <asp:repeater id="rptMinimumCriteria" runat="server">
    <itemtemplate>
      <tr>
        <td>
     	    <asp:button runat="server" id="cmdExpand" commandname="cmdExpand" cssclass="expandbutton" text="+" CausesValidation="False" />
        </td>
        <td>
          <asp:label id="lblSummary" runat="server" />
          <asp:panel id="pnlDetail" runat="server" visible="False" cssclass="minimumcriteriondetail">
            <asp:label id="lblDetail" runat="server" />
          </asp:panel>
        </td>
        <td class="date">
          <asp:label id="lblLastUpdated" runat="server" />
        </td>
      </tr>
	  </itemtemplate>
  </asp:repeater>
</table>

<br /><br />

<epsom:documents id="Documents" documentcategory="LendingPolicies" runat="server" />



