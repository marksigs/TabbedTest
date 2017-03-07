<%@ Control Language="c#" AutoEventWireup="false" Codebehind="Commitment.ascx.cs" Inherits="Epsom.Web.WebUserControls.Commitment" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/ApplicantType.ascx" TagName="applicantType" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<epsom:applicanttype id="ctlApplicantType" runat="server" />
<table class="formtable">
	<tr>
		<td class="prompt"><span class="mandatoryfieldasterisk">*</span> Commitment type</td>
		<td><asp:dropdownlist cssclass="text" id="cmbCommitmentType" runat="server" autopostback="True"></asp:dropdownlist>
			<asp:requiredfieldvalidator controltovalidate="cmbCommitmentType" runat="server" text="***" errormessage="Please select a commitment type"
				enableclientscript="false"/>
		</td>
	</tr>
	<tr>
		<td class="prompt"><span class="mandatoryfieldasterisk">*</span> Name of Lender</td>
		<td><asp:textbox cssclass="text" id="txtLenderName" runat="server"></asp:textbox>
			<asp:requiredfieldvalidator controltovalidate="txtLenderName" runat="server" text="***" errormessage="Please enter name of lender"
				enableclientscript="false"/>
		</td>
	</tr>
	<tr>
		<td class="prompt">Account Number</td>
		<td><asp:textbox cssclass="text" id="txtAccountNumber" runat="server"></asp:textbox>
		</td>
	</tr>
	<asp:panel id="pnlMonthly" runat="server">
		<tr>
			<td class="prompt"><span class="mandatoryfieldasterisk">*</span> Monthly payment</td>
			<td>
				<asp:textbox id="txtPayment" runat="server" cssclass="text"></asp:textbox>
				<asp:requiredfieldvalidator id="txtPayment_validator" runat="server" enableclientscript="false" errormessage="Please enter a monthly payment"
					text="***" controltovalidate="txtPayment"></asp:requiredfieldvalidator>
				<asp:comparevalidator runat="server" enableclientscript="false" errormessage="Monthly payment is invalid"
					text="***" controltovalidate="txtPayment" operator="DataTypeCheck" type="Currency"></asp:comparevalidator>
				<asp:label id="lblPayment" runat="server"></asp:label></td>
		</tr>
	</asp:panel>
	<asp:panel id="pnlBalance" runat="server">
		<tr>
			<td class="prompt"><span class="mandatoryfieldasterisk">*</span> Balance 
				outstanding</td>
			<td>
				<asp:textbox id="txtOutstanding" runat="server" cssclass="text"></asp:textbox>
				<asp:requiredfieldvalidator id="txtOutstanding_validator" runat="server" enableclientscript="false" errormessage="Please enter the outstanding balance"
					text="***" controltovalidate="txtOutstanding"></asp:requiredfieldvalidator>
				<asp:comparevalidator runat="server" enableclientscript="false" errormessage="Balance outstanding is invalid"
					text="***" controltovalidate="txtOutstanding" operator="DataTypeCheck" type="Currency"></asp:comparevalidator>
				<asp:label id="lblOutstanding" runat="server"></asp:label></td>
		</tr>
	</asp:panel>
	
	<asp:Panel Runat="server" ID="pnlAppOnly">
		<tr>
			<td class="prompt">Date of final payment</td>
			<td>
				<asp:textbox id="txtFinalPayment" runat="server" cssclass="text"></asp:textbox>
				<validators:emptydatevalidator runat="server" enableclientscript="false" errormessage="Invalid date of final payment"
					text="***" controltovalidate="txtFinalPayment"></validators:emptydatevalidator>dd/mm/yyyy
			</td>
		</tr>
		<tr>
			<td class="prompt">If you are in arrears, state by how many months</td>
			<td>
				<asp:textbox id="txtArrears" runat="server" cssclass="text"></asp:textbox></td>
		</tr>
  </asp:Panel>
	<tr>
		<td class="prompt"></td>
		<td>
			<asp:checkbox id="chkPaidOnCompletion" runat="server" text="To be repaid on or before completion"></asp:checkbox>
		</td>
	</tr>
	<tr>
		<td class="prompt"></td>
		<td>
			<asp:checkbox id="chkPaidByBusiness" runat="server" text="Paid by applicant's business"></asp:checkbox>
		</td>
	</tr>
</table>
