<%@ Control Language="c#" AutoEventWireup="false" Codebehind="ApplicantType.ascx.cs" Inherits="Epsom.Web.WebUserControls.ApplicantType" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<table class="formtable" >

  <tr>
		<td class="prompt"><span class="mandatoryfieldasterisk">*</span> Customers</td>
    <td>
      <asp:checkbox id="chkApplicant1" runat="server" text="Applicant 1"/>
      <asp:checkbox id="chkApplicant2" runat="server" text="Applicant 2"/>
      <asp:checkbox id="chkGuarantor" runat="server" text="Guarantor"/>
    
      <asp:customvalidator id="CustomValidatorCustomerSelected" runat="Server"
        onservervalidate="CustomValidatorCustomerSelected_ServerValidate"
        text="***" errormessage="Please select one or more customers" >      
        </asp:customvalidator>
    </td>
	</tr>
</table>
