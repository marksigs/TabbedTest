<%@ Control Language="c#" AutoEventWireup="false" Codebehind="ApplicantSummary.ascx.cs" Inherits="Epsom.Web.WebUserControls.ApplicantSummary" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/AddressDisplay.ascx" TagName="AddressDisplay" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/ApplicantOtherNameDisplay.ascx" TagName="ApplicantOtherNameDisplay" TagPrefix="Epsom" %>

  <h3 class="form_heading">Personal details</h3>

  <table class="displaytable">
    <tr>
      <td class="prompt">Title</td>
      <td><asp:label id="lblTitle" runat="server" /></td>
      <td class="prompt">Date of birth</td>
      <td><asp:label id="lblDateOfBirth" runat="server" /></td>
    </tr>
    <tr>
      <td class="prompt">First name</td>
      <td><asp:label id="lblFirstName" runat="server" /></td>
      <td class="prompt">Gender</td>
      <td><asp:label id="lblGender" runat="server" /></td>
    </tr>
    <tr>
      <td class="prompt">Middle names</td>
      <td><asp:label id="lblSecondName" runat="server" /></td>
      <td class="prompt">Marital status</td>
      <td><asp:label id="lblMaritalStatus" runat="server" /></td>
    </tr>
    <tr>
      <td class="prompt">Surname</td>
      <td><asp:label id="lblLastName" runat="server" /></td>
      <td class="prompt">Nationality</td>
      <td><asp:label id="lblNationality" runat="server" /></td>
    </tr>
  </table>

  <asp:repeater id="rptApplicantOtherNames" runat="server">

    <itemtemplate>

      <h3 class="form_heading"><asp:Label ID="lblApplicantOtherNameHeading" Runat="server" /></h3>

      <epsom:applicantOtherNameDisplay id="ctlApplicantOtherNameDisplay" runat="server" />

    </itemtemplate>

  </asp:repeater>

  <asp:repeater id="rptApplicantAddresses" runat="server">

    <itemtemplate>

      <h3 class="form_heading"><asp:Label ID="lblApplicantAddressHeading" Runat="server" /></h3>

      <epsom:addressdisplay id="ctlAddressDisplay" runat="server" IsResidential="true" UnitedKingdomOnly="false" />

    </itemtemplate>

  </asp:repeater>
  <asp:panel id="pnlIncome" runat ="server">
  <h3 class="form_heading">Income details</h3>

  <table class="displaytable">
  <asp:panel id="pnlEmployedIncome" runat ="server">
    <tr>
      <td class="prompt">Guaranteed/regular income</td>
      <td><asp:label id="lblGuaranteedOrRegularIncome" runat="server" /></td>
    </tr>
    <tr>
      <td class="prompt">Non-guaranteed/irregular income</td>
      <td><asp:label id="lblNonGuaranteedOrIrregularIncome" runat="server" /></td>
    </tr>
    </asp:panel> 
    <asp:panel id="pnlSelfEmployedIncome" runat ="server">
    <tr>
      <td class="prompt">Self-employed income</td>
      <td><asp:label id="lblSelfEmployedIncome" runat="server" /></td>
    </tr>
    </asp:panel> 
    <tr>
      <td class="prompt">UK taxpayer?</td>
      <td><asp:label id="lblUKTaxPayer" runat="server" /></td>
    </tr>
  </table>
</asp:panel>

	<asp:panel runat="server" id="pnlEmployments">
	<asp:Repeater ID="rptApplicantEmployments" Runat="server">
	  <itemtemplate>
      <table class="displaytable" style="margin: 1em 0 0 0">
        <tr>
          <td class="prompt">Job title</td>
          <td><asp:label id="lblJobTitle" runat="server" /></td>
        </tr>
        <tr>
          <td class="prompt">Employment status</td>
          <td><asp:label id="lblEmploymentStatus" runat="server" /></td>
        </tr>

        <asp:panel runat="server" id="pnlEmploymentStart">
        <tr>
          <td class="prompt">Date employment started</td>
          <td><asp:label id="lblDateEmploymentStarted" runat="server" /></td>
        </tr>
        </asp:panel> 
        <asp:panel runat="server" id="pnlEmploymentDuration">
        <tr>
          <td class="prompt">Time in this employment</td>
          <td><asp:label id="lblTimeInThisEmployment" runat="server" /></td>
        </tr>
        </asp:Panel>
        <asp:panel runat="server" id="pnlContractDetails">
          <tr>
            <td class="prompt">Length of contract</td>
            <td><asp:label id="lblLengthOfContract" runat="server" /></td>
          </tr>
          <tr>
            <td class="prompt">Renewable</td>
            <td><asp:label id="lblRenewable" runat="server" /></td>
          </tr>
        </asp:Panel>
      </table>
	  </itemtemplate>
	</asp:Repeater>
  </asp:panel>
  
  
