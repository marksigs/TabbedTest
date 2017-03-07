<%@ Control Language="c#" AutoEventWireup="false" Codebehind="GuarantorSummary.ascx.cs" Inherits="Epsom.Web.WebUserControls.GuarantorSummary" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<table class="formtable">
  <tr>
		<td><h3 class="form_subheading">Guarantor</h3></td>
		<td>&nbsp</td>
		<td><h3 class="form_subheading">Address</h3></td>
  </tr>
  <tr>
    <td class="prompt">First name</td>
    <td class="promptblack"><asp:label id="lblFirstName" runat="server" /></td>
    <td class="promptblack"><asp:label id="lblAddress1" runat="server" /></td>
		<td><asp:Button id="cmdEdit" runat="server" text="Edit"/></td>
  </tr>
  <tr>
    <td class="prompt">Middle names</td>
    <td class="promptblack"><asp:label id="lblSecondName" runat="server" /></td>
    <td class="promptblack"><asp:label id="lblAddress2" runat="server" /></td>
  </tr>
  <tr>
    <td class="prompt">Surname</td>
    <td class="promptblack"><asp:label id="lblLastName" runat="server" /></td>
    <td class="promptblack"><asp:label id="lblAddress3" runat="server" /></td>
  </tr>
  <tr>
    <td class="prompt">Date of birth</td>
    <td class="promptblack"><asp:label id="lblDateOfBirth" runat="server" /></td>
    <td class="promptblack"><asp:label id="lblAddress4" runat="server" /></td>
  </tr>
  <tr>
    <td class="prompt">Gender</td>
    <td class="promptblack"><asp:label id="lblGender" runat="server" /></td>
    <td class="prompt">From:</td>
    <td class="promptblack"><asp:label id="lblAddressFrom" runat="server" /></td>
  </tr>
  <tr>
    <td class="prompt">Marital status</td>
    <td class="promptblack"><asp:label id="lblMaritalStatus" runat="server" /></td>
  </tr>
</table>
<h3 class="form_subheading">Employment</h3>
<table class="formtable">
  <tr>
    <td class="prompt">Employment status</td>
    <td class="promptblack"><asp:label id="lblEmploymentStatus" runat="server" /></td>
    <td class="prompt">&nbsp</td>
    <td class="promptblack">&nbsp</td>
  </tr>
</table>
