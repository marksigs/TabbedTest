<%@ Control Language="c#" AutoEventWireup="false" Codebehind="ApplicantAndGuarantorSummary.ascx.cs" Inherits="Epsom.Web.WebUserControls.ApplicantAndGuarantorSummary" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/ApplicantSummary.ascx" TagName="ApplicantSummary" TagPrefix="Epsom" %>

<h2 class="form_heading">Customer Details</h2>
      
<asp:repeater id="rptApplicants" runat="server">
  <itemtemplate>
    <fieldset>
      <legend><asp:label id="lblApplicantTitle" runat="server" /></legend>
      <epsom:applicantSummary id="ctlApplicantSummary" runat="server" />
      <table class="displaytable">
        <tr>
          <td>&nbsp;</td>
          <td style="text-align:right"><asp:Button ID="cmdEditApplicant" commandName="cmdEditApplicant" runat="server"
            text="Edit" CssClass="button" /></td>
        </tr>
      </table>
    </fieldset> 
  </itemtemplate> 
</asp:repeater> 

<asp:Panel id="pnlGuarantorDetails" runat="Server" visible="false">
    <fieldset>
      <legend><asp:label id="Label1" runat="server" text="Guarantor"/></legend>
      <epsom:applicantSummary id="ctlGuarantorSummary" runat="server" />
    </fieldset> 
    <table class="displaytable">
      <tr>
        <td>&nbsp;</td>
        <td style="text-align:right"><asp:Button ID="cmdEditGuarantor" commandName="cmdEditGuarantor" runat="server"
          text="Edit" CssClass="button" /></td>
      </tr>
    </table>
</asp:Panel>
