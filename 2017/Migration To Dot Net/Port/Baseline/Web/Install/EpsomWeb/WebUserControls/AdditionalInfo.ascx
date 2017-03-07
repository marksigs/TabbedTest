<%@ Control Language="c#" AutoEventWireup="false" Codebehind="AdditionalInfo.ascx.cs" Inherits="Epsom.Web.WebUserControls.AdditionalInfo" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>

<asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary"
  showsummary="true" id="Validationsummary1" />

<h2 class="form_heading">Additional Information</h2>

<p>
  <cms:cmsBoilerplate cmsref="70" runat="server" />
</p>

<table class="formtable">
  <tr>
    <td class="prompt">
      Your reference
    </td>
    <td>
      <asp:textbox id="txtIntroducerReference" runat="server" cssclass="text" />
    </td>
  </tr>
</table>

<table class="formtable">
  <tr>
    <td class="prompt">
      Total number of dependants
      <br>
     (for all applicants)
    </td>
    <td><asp:textbox id="txtNumberOfDependants" runat="server" columns="20" cssclass="text" /></td>
  </tr>
  <asp:panel id="pnlFinancialBenefitOfAll" runat="server" >
  <tr>
    <td class="prompt">
     Is the purpose of the proposed loan for the direct financial benefit of all applicants?
    </td>
    <td>      
    <epsom:nullableyesno  id="nynFinancialBenefitOfAll" runat="server" required="true" 
    errormessage="Specify if the loan has direct financial benefit" enableclientscript="false" />
    </td>
  </tr>
  </asp:panel>
</table>

<table class="formtable">
  <tr>
    <td class="prompt"><cms:cmsBoilerplate cmsref="71" runat="server" ID="lblDipInfoLabel"/><p><asp:label id="lblCharactersRemaining" Runat="server">Characters remaining : </asp:label><asp:textbox columns="4" enabled="False" ID="txtCharactersRemaining" runat="server"></asp:textbox></p></td>
    <td><asp:textbox id="txtDipAdditionalInformation" rows="15" columns="25" runat="server" textmode="MultiLine" /></td>
  </tr>
  <asp:Panel ID="pnlAppAdditionalInfo" Runat="server">
    <tr>
      <td class="prompt"><cms:cmsBoilerplate cmsref="71" runat="server" /><p><asp:label id="lblCharactersRemaining2" Runat="server">Characters remaining : </asp:label><asp:textbox columns="4" border="0" enabled="False"  ID="txtCharactersRemaining2" runat="server"></asp:textbox></p></td>
      <td><asp:textbox id="txtAppAdditionalInformation" rows="15" columns="25" runat="server" textmode="MultiLine" /></td>
    </tr>
  </asp:Panel>
</table>
