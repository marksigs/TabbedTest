<%@ Control Language="c#" AutoEventWireup="false" Codebehind="InsuranceSummary.ascx.cs" Inherits="Epsom.Web.WebUserControls.InsuranceSummary" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<h2 class="form_heading">Insurance</h2>
<p>
<asp:label runat="server" id="lblNotEntered"></asp:label></p>
<table class="displaytable">
<asp:panel runat="server" id="pnlHeadings">
 <tr>
   <td>Insurance Type</td>
   <td>Company</td>
   <td>Monthly Payment</td>
 </tr>
</asp:panel> 
<asp:panel runat="server" id="panelInsurance1">
  <tr>
    <td><asp:label runat="server" id="labelInsuranceType1" /></td>
    <td><asp:label runat="server" id="labelInsuranceCompany1" /></td>
    <td><asp:label runat="server" id="labelInsurancePayment1" /></td>
  </tr>
</asp:panel>
  
<asp:panel runat="server" id="panelInsurance2">
  <tr>
    <td><asp:label runat="server" id="labelInsuranceType2" /></td>
    <td><asp:label runat="server" id="labelInsuranceCompany2" /></td>
    <td><asp:label runat="server" id="labelInsurancePayment2" /></td>
  </tr>
</asp:panel>
  
<asp:panel runat="server" id="panelInsurance3">
  <tr>
    <td><asp:label runat="server" id="labelInsuranceType3" /></td>
    <td><asp:label runat="server" id="labelInsuranceCompany3" /></td>
    <td><asp:label runat="server" id="labelInsurancePayment3" /></td>
  </tr>
</asp:panel>

</table>
  
